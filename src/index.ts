import * as xlsx from 'xlsx';
import * as fs from 'fs-extra-promise';
import * as path from 'path';

import * as utils from './utils'
import gCfg from "./config.json";
import {CTypeChecker,ETypeNames} from "./TypeChecker";

utils.SetEnableDebugOutput(gCfg.EnableDebugOutput);
utils.SetLineBreaker(gCfg.LineBreak);
const gExportWrapper: utils.IExportWrapper = <utils.IExportWrapper>utils.ExportWrapperMap.get(gCfg.Export.type);
if (gExportWrapper == undefined) {
	utils.exception(utils.red(`Export is not currently supported for the current type "${utils.yellow_ul(gCfg.Export.type)}"!`));
}

const gRootDir = process.cwd();
CTypeChecker.DateFmt = gCfg.DateFmt;
CTypeChecker.TinyDateFmt = gCfg.TinyDateFmt;
CTypeChecker.FractionDigitsFMT = gCfg.FractionDigitsFMT;

async function HandleDir(dirName: string): Promise<void> {
	let pa = await fs.readdirAsync(dirName);
	pa.forEach(async function(fileName, index){
		const filePath = path.join(dirName, fileName);
		let info = await fs.statAsync(filePath);
		if(!info.isFile()) {
			return;
		}
		await HandleExcelFile(filePath);
	});
}

class XlsColumnIter
{
	public constructor(max: string) {
		for (let c of max) {
			if (XlsColumnIter.VAILDWORD.indexOf(c)<0) {
				throw 'Xls ColumnIter Max Format Error!!!';
			}
		}
		this._max = XlsColumnIter.S26ToNum(max);
	}
	public seekToBegin() { this._curr = 1; }
	public get next(): string|undefined {
		if (this.end) {
			return undefined;
		}
		return XlsColumnIter.NumToS26(++this._curr);
	}
	public get curr26(): string { return XlsColumnIter.NumToS26(this._curr); }
	public get curr10(): number { return this._curr; }
	public get end(): boolean { return this._curr > this._max; }

	public static S26ToNum(str: string) {
		let result = 0;
		let ss = str.toUpperCase();
		for (let i = str.length - 1, j = 1; i >= 0; i--, j *= 26) {
			let c = ss[i];
			if (c < 'A' || c > 'Z') return 0;
			result += (c.charCodeAt(0) - 64) * j;
		}
		return result;
	}
	public static NumToS26(num: number): string{
		let result="";
		while (num > 0){
			let m = num % 26;
			if (m == 0) m = 26;
			result = String.fromCharCode(m + 64) + result;
			num = (num - m) / 26;
		}
		return result;
	}
	private _curr = 1;
	private _max: number;
	private static readonly VAILDWORD = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
}

function GetCellData(worksheet: xlsx.WorkSheet, column: string, row: number): xlsx.CellObject|undefined {
	return worksheet[column + row.toString()];
}

function HandleWorkSheet(fileName: string, sheetName: string, worksheet: xlsx.WorkSheet): utils.DataTable|undefined {
	// find csv name
	if (worksheet[gCfg.CSVNameCellID] == undefined || utils.NullStr(worksheet[gCfg.CSVNameCellID].w)) {
		utils.logger(false, `excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" CSV name not defined. Ignore it!`);
		return;
	}
	const CSVName = worksheet[gCfg.CSVNameCellID].w;
	if (gCfg.ExcludeSheetNames.indexOf(CSVName) >= 0) {
		utils.logger(true, `- Pass CSV "${CSVName}"`);
		return;
	}

	let ColumnMax = 'A';
	let RowMax = 0;
	let ColumnArry = new Array<{sid:string, id:number, name:string, checker:CTypeChecker}>();
	// find max column and rows
	{
		const REF = worksheet["!ref"];
		if (!REF) {
			utils.logger(false, `excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" IS EMPTY. Ignore it!`);
			return;
		}
		let SPREF = REF.split(":");
		if (SPREF.length != 2) {
			utils.logger(false, `excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" [!ref] = "${REF}" format error!!`);
			return;
		}
		ColumnMax = SPREF[1].toUpperCase().replace(/([0-9]*)/g, '');
		RowMax = parseInt(SPREF[1].toUpperCase().replace(/([A-Z]*)/g, ''));
	}
	let rowIdx = 2;
	let columnIdx = new XlsColumnIter(ColumnMax);
	let DataTable = { name:CSVName, datas:new Array<Array<string>>() };
	// find column name
	for (; rowIdx <= RowMax; ++rowIdx) {
		const firstCell = GetCellData(worksheet, 'A', rowIdx);
		if (firstCell == undefined || firstCell.w == undefined || utils.NullStr(firstCell.w)) {
			continue;
		}
		if (firstCell.w[0] == '#') {
			if (gCfg.EnableExportCommentRows) {
				columnIdx.seekToBegin();
				const tmpArry = [];
				do {
					const cell = GetCellData(worksheet, columnIdx.curr26, rowIdx);
					tmpArry.push((cell && cell.w)?cell.w:'');
				}while(columnIdx.next);
				DataTable.datas.push(tmpArry);
			}
			continue;
		}
		columnIdx.seekToBegin();
		const tmpArry = [];
		do {
			const colName = columnIdx.curr26;
			const cell = GetCellData(worksheet, colName, rowIdx);
			if (cell == undefined || cell.w == undefined || utils.NullStr(cell.w) || (gCfg.EnableExportCommentColumns == false && cell.w[0] == '#')) {
				continue;
			}
			ColumnArry.push({id:columnIdx.curr10, sid:colName, name:cell.w, checker:(cell.w[0] == '#')?new CTypeChecker(ETypeNames.string):<any>undefined});
			tmpArry.push(cell.w);
		}while(columnIdx.next);
		DataTable.datas.push(tmpArry);
		++rowIdx;
		break;
	}
	// find type
	for (; rowIdx <= RowMax; ++rowIdx) {
		const firstCell = GetCellData(worksheet, ColumnArry[0].sid, rowIdx);
		if (firstCell == undefined || firstCell.w == undefined || utils.NullStr(firstCell.w)) {
			continue;
		}
		if (firstCell.w[0] == '#') {
			if (gCfg.EnableExportCommentRows) {
				const tmpArry = [];
				for (let col of ColumnArry) {
					const cell = GetCellData(worksheet, col.sid, rowIdx);
					tmpArry.push((cell && cell.w)?cell.w:'');
				}
				DataTable.datas.push(tmpArry);
			}
			continue;
		}

		if (firstCell.w[0] != '*') {
			utils.exception(`excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" CSV Type Column not found!`);
		}
		const tmpArry = [];
		for (const col of ColumnArry) {
			// continue...
			const cell = GetCellData(worksheet, col.sid, rowIdx);
			if (col.checker != undefined) {
				if (gCfg.EnableExportCommentColumns) {
					tmpArry.push((cell && cell.w)?cell.w:'');
				}
				continue;
			}
			if (cell == undefined || cell.w == undefined) {
				utils.exception(`excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" CSV Type Column "${utils.yellow_ul(col.name)}" not found!`);
				return;
			}
			const typeStr = col.id <= 1 ? cell.w.substr(1):cell.w;
			try {
				col.checker = new CTypeChecker(typeStr);
				tmpArry.push(cell.w);
			} catch (ex) {
				new CTypeChecker(typeStr);
				utils.exception(`excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" CSV Type Column`
						+ ` "${utils.yellow_ul(col.name)}" format error "${utils.yellow_ul(cell.w)}". expect is "${utils.yellow_ul(typeStr)}"!`, ex);
			}
		}
		DataTable.datas.push(tmpArry);
		++rowIdx;
		break;
	}

	// handle datas
	for (; rowIdx <= RowMax; ++rowIdx) {
		let firstCol = true;
		const tmpArry = [];
		for (let col of ColumnArry) {
			const cell = GetCellData(worksheet, col.sid, rowIdx);
			if (firstCol) {
				if (cell == undefined || cell.w == undefined || utils.NullStr(cell.w)) {
					break;
				}
				else if (cell.w[0] == '#') {
					if (gCfg.EnableExportCommentRows) {
						const tmpArry = [];
						for (let col of ColumnArry) {
							const cell = GetCellData(worksheet, col.sid, rowIdx);
							tmpArry.push((cell && cell.w)?cell.w:'');
						}
						DataTable.datas.push(tmpArry);
					}
					break;
				}
				firstCol = false;
			}
			const value = cell && cell.w ? cell.w : '';
			if (gCfg.EnableTypeCheck) {
				if (!col.checker.CheckDataVaildate(cell)) {
					col.checker.CheckDataVaildate(cell);
					utils.exception(`excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" CSV Cell `
							+ `"${utils.yellow_ul(col.sid+(rowIdx).toString())}" format not match "${utils.yellow_ul(value)}" with ${utils.yellow_ul(col.checker.s)}!`);
					return;
				}
			}
			try {
				tmpArry.push(col.checker.ParseDataStr(cell));
			} catch (ex) {
				// col.checker.ParseDataStr(cell);
				utils.exception(`excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" CSV Cell "${utils.yellow_ul(col.sid+(rowIdx).toString())}" `
						+ `Parse Data "${utils.yellow_ul(value)}" With ${utils.yellow_ul(col.checker.s)} Cause utils.exception "${utils.red(ex)}"!`);
				return;
			}
		}
		if (!firstCol) {
			DataTable.datas.push(tmpArry);
		}
	}
	return DataTable;
}

async function HandleExcelFile(fileName: string) {
	const extname = path.extname(fileName);
	if (extname != '.xls' && extname != '.xlsx') {
		return;
	}
	if (gCfg.ExcludeFileNames.indexOf(path.basename(fileName)) >= 0) {
		utils.logger(true, `- Pass File "${fileName}"`);
		return;
	}
	if (path.basename(fileName).indexOf(`~$`) == 0) {
		utils.logger(true, `- Pass File "${fileName}"`);
		return;
	}
	let opt:xlsx.ParsingOptions = {
		type: "buffer",
		// codepage: 0,//If specified, use code page when appropriate **
		cellFormula: false,//Save formulae to the .f field
		cellHTML: false,//Parse rich text and save HTML to the .h field
		cellText: true,//Generated formatted text to the .w field
		cellDates: true,//Store dates as type d (default is n)
		/**
		 * If specified, use the string for date code 14 **
		 * https://github.com/SheetJS/js-xlsx#parsing-options
		 *		Format 14 (m/d/yy) is localized by Excel: even though the file specifies that number format,
		 *		it will be drawn differently based on system settings. It makes sense when the producer and
		 *		consumer of files are in the same locale, but that is not always the case over the Internet.
		 *		To get around this ambiguity, parse functions accept the dateNF option to override the interpretation of that specific format string.
		 */
		dateNF: 'yyyy/mm/dd',
		WTF: true,//If true, throw errors on unexpected file features **
	};
	const filebuffer = await fs.readFileAsync(fileName);
	const excel = xlsx.read(filebuffer, opt);
	if (excel == null) {
		utils.exception(`excel ${utils.yellow_ul(fileName)} open failure.`);
	}
	if (excel.Sheets == null) {
		return;
	}
	for (let sheetName of excel.SheetNames) {
		utils.logger(true, `handle excel "${utils.brightWhite(fileName)}" sheet "${utils.yellow_ul(sheetName)}"`);
		const worksheet = excel.Sheets[sheetName];
		const datatable = HandleWorkSheet(fileName, sheetName, worksheet);
		if (datatable) {
			utils.ExportExcelDataMap.set(datatable.name, datatable);
			gExportWrapper.exportTo(datatable, gCfg.Export.OutputDir);
		}
	}
}

async function execute() {
	if (!fs.existsSync(gCfg.Export.OutputDir)) {
		fs.mkdirSync(gCfg.Export.OutputDir);
	}
	for (let fileOrPath of gCfg.IncludeFilesAndPath) {
		if (!path.isAbsolute(fileOrPath)) {
			fileOrPath = path.join(gRootDir, fileOrPath);
		}
		if (!fs.existsSync(fileOrPath)) {
			utils.logger(false, `file or directory "${utils.yellow_ul(fileOrPath)}" not found!`);
			continue;
		}
		if (fs.statSync(fileOrPath).isDirectory()) {
			await HandleDir(fileOrPath);
		} else if (fs.statSync(fileOrPath).isFile()) {
			await HandleExcelFile(fileOrPath);
		} else {
			utils.exception(`UnHandle file or directory type : "${utils.yellow_ul(fileOrPath)}"`);
		}
	}
}

////////////////////////////////////////////////////////////////////////////////
function main() {
	try {
		execute()
	} catch (ex) {
		utils.exception(ex);
	}
}

// main entry
main();
