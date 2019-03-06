import * as xlsx from 'xlsx';
import * as fs from 'fs-extra-promise';
import * as path from 'path';

import * as utils from './utils'
import ConfTpl from "./config_tpl.json";
let gCfg: typeof ConfTpl = ConfTpl; // default config
if (process.argv.length >= 3 && fs.existsSync(process.argv[2])) {
	gCfg = JSON.parse(<string>fs.readFileSync(process.argv[2], { encoding: 'utf8' }));
	function check(cfg: any, tpl: any):void{
		for (let key in tpl) {
			if (tpl[key] != null && typeof cfg[key] !== typeof tpl[key]) {
				throw utils.red(`configure format error. key "${utils.yellow(key)}" not found!`);
			}
			if (utils.isObject(typeof tpl[key])) {
				check(cfg[key], tpl[key]);
			}
		}
	};
	check(gCfg, ConfTpl);
}
import {CTypeChecker,ETypeNames} from "./TypeChecker";

utils.SetEnableDebugOutput(gCfg.EnableDebugOutput);
utils.SetLineBreaker(gCfg.LineBreak);

const gExportWrapperLst = new Array<utils.IExportWrapper>();
for (const exportCfg of gCfg.Export) {
	const Constructor = utils.ExportWrapperMap.get(exportCfg.type);
	if (Constructor == undefined) {
		utils.exception(utils.red(`Export is not currently supported for the current type "${utils.yellow_ul(exportCfg.type)}"!`));
		throw `ERROR : Export constructor not found. initialize failure!`;
	}
	const Exportor = Constructor.call(Constructor, exportCfg);
	if (Exportor) {
		if ((<any>exportCfg).ExtName == undefined) {
			(<any>exportCfg).ExtName = Exportor.DefaultExtName;
		}
		gExportWrapperLst.push(Exportor);
	}
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

function GetCellData(worksheet: xlsx.WorkSheet, c: number, r: number): xlsx.CellObject|undefined {
	const cell = xlsx.utils.encode_cell({c, r});
	return worksheet[cell];
}

function HandleWorkSheet(fileName: string, sheetName: string, worksheet: xlsx.WorkSheet): utils.SheetDataTable|undefined {
	// find csv name
	if (worksheet[gCfg.CSVNameCellID] == undefined || utils.NullStr(worksheet[gCfg.CSVNameCellID].w)) {
		utils.logger(false, `excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" CSV name not defined. Ignore it!`);
		return;
	}
	const TableName = worksheet[gCfg.CSVNameCellID].w;
	if (gCfg.ExcludeSheetNames.indexOf(TableName) >= 0) {
		utils.logger(true, `- Pass CSV "${TableName}"`);
		return;
	}

	const Range = xlsx.utils.decode_range(<string>worksheet["!ref"]);
	const ColumnMax = Range.e.c;
	const RowMax = Range.e.r;
	const ColumnArry = new Array<{cIdx:number, name:string, checker:CTypeChecker}>();
	// find max column and rows
	let rIdx = 1;
	const DataTable = new utils.SheetDataTable(TableName);
	// find column name
	for (; rIdx <= RowMax; ++rIdx) {
		const firstCell = GetCellData(worksheet, 0, rIdx);
		if (firstCell == undefined || firstCell.w == undefined || utils.NullStr(firstCell.w)) {
			continue;
		}
		if (firstCell.w[0] == '#') {
			if (gCfg.EnableExportCommentRows) {
				const tmpArry = [];
				for (let cIdx = 0; cIdx <= ColumnMax; ++cIdx) {
					const cell = GetCellData(worksheet, cIdx, rIdx);
					tmpArry.push((cell && cell.w)?cell.w:'');
				}
				DataTable.values.push({type:utils.ESheetRowType.comment, values: tmpArry})
			}
			continue;
		}
		const tmpArry = [];
		for (let cIdx = 0; cIdx <= ColumnMax; ++cIdx) {
			const cell = GetCellData(worksheet, cIdx, rIdx);
			if (cell == undefined || cell.w == undefined || utils.NullStr(cell.w) || (gCfg.EnableExportCommentColumns == false && cell.w[0] == '#')) {
				continue;
			}
			ColumnArry.push({cIdx, name:cell.w, checker:(cell.w[0] == '#')?new CTypeChecker(ETypeNames.string):<any>undefined});
			tmpArry.push(cell.w);
		}
		DataTable.values.push({type:utils.ESheetRowType.header, values: tmpArry});
		++rIdx;
		break;
	}
	// find type
	for (; rIdx <= RowMax; ++rIdx) {
		const firstCell = GetCellData(worksheet, ColumnArry[0].cIdx, rIdx);
		if (firstCell == undefined || firstCell.w == undefined || utils.NullStr(firstCell.w)) {
			continue;
		}
		if (firstCell.w[0] == '#') {
			if (gCfg.EnableExportCommentRows) {
				const tmpArry = [];
				for (let col of ColumnArry) {
					const cell = GetCellData(worksheet, col.cIdx, rIdx);
					tmpArry.push((cell && cell.w)?cell.w:'');
				}
				DataTable.values.push({type:utils.ESheetRowType.comment, values: tmpArry});
			}
			continue;
		}

		if (firstCell.w[0] != '*') {
			utils.exception(`excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" CSV Type Column not found!`);
		}
		const tmpArry = [];
		let typeHeader = new Array<utils.SheetHeader>();
		for (const col of ColumnArry) {
			const cell = GetCellData(worksheet, col.cIdx, rIdx);
			if (col.checker != undefined) {
				if (gCfg.EnableExportCommentColumns) {
					const stype = (cell && cell.w)?cell.w:'';
					tmpArry.push(stype);
					typeHeader.push({name:col.name, typeChecker:col.checker, stype, comment:true});
				}
				continue;
			}
			if (cell == undefined || cell.w == undefined) {
				utils.exception(`excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" CSV Type Column "${utils.yellow_ul(col.name)}" not found!`);
				return;
			}
			const typeStr = col.cIdx == 0 ? cell.w.substr(1):cell.w;
			try {
				col.checker = new CTypeChecker(typeStr);
				tmpArry.push(cell.w);
				typeHeader.push({name:col.name, typeChecker:col.checker, stype:cell.w, comment:false});
			} catch (ex) {
				// new CTypeChecker(typeStr);
				utils.exception(`excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" CSV Type Column`
						+ ` "${utils.yellow_ul(col.name)}" format error "${utils.yellow_ul(cell.w)}". expect is "${utils.yellow_ul(typeStr)}"!`, ex);
			}
		}
		DataTable.headerLst = typeHeader;
		DataTable.values.push({type:utils.ESheetRowType.type, values: tmpArry});
		++rIdx;
		break;
	}

	// handle datas
	for (; rIdx <= RowMax; ++rIdx) {
		let firstCol = true;
		const tmpArry = [];
		for (let col of ColumnArry) {
			const cell = GetCellData(worksheet, col.cIdx, rIdx);
			if (firstCol) {
				if (cell == undefined || cell.w == undefined || utils.NullStr(cell.w)) {
					break;
				}
				else if (cell.w[0] == '#') {
					if (gCfg.EnableExportCommentRows) {
						const tmpArry = [];
						for (let col of ColumnArry) {
							const cell = GetCellData(worksheet, col.cIdx, rIdx);
							tmpArry.push((cell && cell.w)?cell.w:'');
						}
						DataTable.values.push({type:utils.ESheetRowType.comment, values: tmpArry});
					}
					break;
				}
				firstCol = false;
			}
			const value = cell && cell.w ? cell.w : '';
			// if (cell) {
			// 	cell.w = utils.StringTranslate.ReplaceNewLineToLashN(cell.w||'');
			// }
			if (gCfg.EnableTypeCheck) {
				if (!col.checker.CheckDataVaildate(cell)) {
					col.checker.CheckDataVaildate(cell); // for debug used
					utils.exception(`excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" `
								  + `CSV Cell "${utils.yellow_ul(utils.FMT26.NumToS26(col.cIdx)+(rIdx).toString())}" `
								  + `format not match "${utils.yellow_ul(value)}" with ${utils.yellow_ul(col.checker.s)}!`);
					return;
				}
			}
			try {
				tmpArry.push(col.checker.ParseDataByType(cell));
			} catch (ex) {
				// col.checker.ParseDataStr(cell);
				utils.exception(`excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" `
							  + `CSV Cell "${utils.yellow_ul(utils.FMT26.NumToS26(col.cIdx)+(rIdx).toString())}" `
							  + `Parse Data "${utils.yellow_ul(value)}" With ${utils.yellow_ul(col.checker.s)} `
							  + `Cause utils.exception "${utils.red(ex)}"!`);
				return;
			}
		}
		if (!firstCol) {
			DataTable.values.push({type:utils.ESheetRowType.data, values: tmpArry});
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
		utils.logger(true, `- Handle excel "${utils.brightWhite(fileName)}" sheet "${utils.yellow_ul(sheetName)}"`);
		const worksheet = excel.Sheets[sheetName];
		const datatable = HandleWorkSheet(fileName, sheetName, worksheet);
		if (datatable) {
			utils.ExportExcelDataMap.set(datatable.name, datatable);
			for (const handler of gExportWrapperLst) {
				await handler.ExportTo(datatable, gCfg);
			}
		}
	}
}

async function execute() {
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
async function main() {
	try {
		utils.SetBeforeExistHandler(()=>{
			for (let handler of gExportWrapperLst) {
				handler.ExportEnd(gCfg);
			}
		});
		await execute()
		console.log('--------------------------------------------------------------------');
	} catch (ex) {
		utils.exception(ex);
	}
}

// main entry
main();
