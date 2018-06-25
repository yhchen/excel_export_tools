import * as xlsx from 'xlsx';

import gCfg from "./config.json";
import * as fs from 'fs';
import * as path from 'path';
import * as chalk from 'chalk';

import {CTypeChecker,ETypeNameMap} from "./TypeChecker";

/*************** console color ***************/
const yellow_ul = chalk.default.yellow.underline;	//yellow under line
const yellow = chalk.default.yellow;
const red = chalk.default.redBright;
const green = chalk.default.greenBright;
const brightWhite = chalk.default.whiteBright.bold
function logger(debugMode: boolean, ...args: any[]) {
	if (!gCfg.EnableDebugOutput && debugMode) {
		return;
	}
	console.log(...args);
}
function trace(...args: any[]) { logger(true, ...args); }
function exception(txt: string, ex?:any) {
	logger(false, red(`[ERROR] `) + txt);
	if (ex) { logger(false, red(ex)); }
	throw txt;
}

function NullStr(s: string) {
	if (typeof s === "string") {
		return s.trim().length <= 0;
	}
	return true;
}

/************ console color end*************/

const gRootDir = process.cwd();
let gExcelFileList = new Array<string>();

function HandleDir(dirName: string): void {
	var pa = fs.readdirSync(dirName);
	pa.forEach(function(fileName,index){
		const filePath = path.join(dirName, fileName);
		var info = fs.statSync(filePath);
		if(!info.isFile()) {
			return;
		}
		HandleExcelFile(filePath);
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

function HandleWorkSheet(fileName: string, sheetName: string, worksheet: xlsx.WorkSheet): void{
	const StartTick = Date.now();

	// find csv name
	if (worksheet[gCfg.CSVNameCellID] == undefined || NullStr(worksheet[gCfg.CSVNameCellID].w)) {
		logger(false, `excel file "${yellow_ul(fileName)}" sheet "${yellow_ul(sheetName)}" CSV name not defined. Ignore it!`);
		return;
	}
	const CSVName = worksheet[gCfg.CSVNameCellID].w;
	if (gCfg.ExcludeCsvTableNames.indexOf(CSVName) >= 0) {
		logger(true, `- Pass CSV "${CSVName}"`);
	}

	let ColumnMax = 'A';
	let RowMax = 0;
	let ColumnArry = new Array<{sid:string, id:number, name:string, checker:CTypeChecker}>();
	// find max column and rows
	{
		const REF = worksheet["!ref"];
		if (!REF) {
			logger(false, `excel file "${yellow_ul(fileName)}" sheet "${yellow_ul(sheetName)}" IS EMPTY. Ignore it!`);
			return;
		}
		let SPREF = REF.split(":");
		if (SPREF.length != 2) {
			logger(false, `excel file "${yellow_ul(fileName)}" sheet "${yellow_ul(sheetName)}" [!ref] = "${REF}" format error!!`);
			return;
		}
		ColumnMax = SPREF[1].toUpperCase().replace(/([0-9]*)/g, '');
		RowMax = parseInt(SPREF[1].toUpperCase().replace(/([A-Z]*)/g, ''));
	}
	let rowIdx = 2;
	let csvcontent = '';
	let columnIdx = new XlsColumnIter(ColumnMax);
	let tmpArry: string[] = [];
	// find column name
	for (; rowIdx <= RowMax; ++rowIdx) {
		const firstCell = GetCellData(worksheet, 'A', rowIdx);
		if (firstCell == undefined || firstCell.w == undefined || NullStr(firstCell.w)) {
			continue;
		}
		if (firstCell.w[0] == '#') {
			if (gCfg.EnableExportCommentRows) {
				columnIdx.seekToBegin();
				tmpArry = [];
				do {
					const cell = GetCellData(worksheet, columnIdx.curr26, rowIdx);
					tmpArry.push((cell && cell.w)?cell.w:'');
				}while(columnIdx.next);
				csvcontent += tmpArry.join(',') + gCfg.LineBreak;
			}
			continue;
		}
		columnIdx.seekToBegin();
		tmpArry = [];
		do {
			const colName = columnIdx.curr26;
			const cell = GetCellData(worksheet, colName, rowIdx);
			if (cell == undefined || cell.w == undefined || NullStr(cell.w) || (gCfg.EnableExportCommentColumns == false && cell.w[0] == '#')) {
				continue;
			}
			ColumnArry.push({id:columnIdx.curr10, sid:colName, name:cell.w, checker:(cell.w[0] == '#')?new CTypeChecker(ETypeNameMap.string):<any>undefined});
			tmpArry.push(cell.w);
		}while(columnIdx.next);
		++rowIdx;
		break;
	}
	csvcontent += tmpArry.join(',') + gCfg.LineBreak;
	// find type
	for (; rowIdx <= RowMax; ++rowIdx) {
		const firstCell = GetCellData(worksheet, ColumnArry[0].sid, rowIdx);
		if (firstCell == undefined || firstCell.w == undefined || NullStr(firstCell.w)) {
			continue;
		}
		if (firstCell.w[0] == '#') {
			if (gCfg.EnableExportCommentRows) {
				columnIdx.seekToBegin();
				tmpArry = [];
				do {
					const cell = GetCellData(worksheet, columnIdx.curr26, rowIdx);
					tmpArry.push((cell && cell.w)?cell.w:'');
				}while(columnIdx.next);
				csvcontent += tmpArry.join(',') + gCfg.LineBreak;
			}
			continue;
		}

		if (firstCell.w[0] != '*') {
			exception(`excel file "${yellow_ul(fileName)}" sheet "${yellow_ul(sheetName)}" CSV Type Column not found!`);
		}
		tmpArry = [];
		for (const col of ColumnArry) {
			// continue...
			if (col.checker != undefined) {
				if (gCfg.EnableExportCommentColumns) {
					tmpArry.push('');
				}
				continue;
			}
			const cell = GetCellData(worksheet, col.sid, rowIdx);
			if (cell == undefined || cell.w == undefined) {
				exception(`excel file "${yellow_ul(fileName)}" sheet "${yellow_ul(sheetName)}" CSV Type Column "${yellow_ul(col.name)}" not found!`);
				return;
			}
			const typeStr = col.id <= 1 ? cell.w.substr(1):cell.w;
			try {
				col.checker = new CTypeChecker(typeStr);
				let v = cell.w.replace(/"/g, `""`);
				if (v.indexOf(',') >= 0) {
					v = '"' + v + '"';
				}
				tmpArry.push(`${v}`);
			} catch (ex) {
				exception(`excel file "${yellow_ul(fileName)}" sheet "${yellow_ul(sheetName)}" CSV Type Column`
						+ ` "${yellow_ul(col.name)}" format error "${yellow_ul(cell.w)}". expect is "${yellow_ul(typeStr)}"!`, ex);
			}
		}
		++rowIdx;
		break;
	}
	csvcontent += `${tmpArry.join(',')}${gCfg.LineBreak}`;

	// handle datas
	for (; rowIdx <= RowMax; ++rowIdx) {
		let firstCol = true;
		tmpArry = [];
		for (let col of ColumnArry) {
			const cell = GetCellData(worksheet, col.sid, rowIdx);
			if (firstCol) {
				if (cell == undefined || cell.w == undefined || NullStr(cell.w)) {
					break;
				}
				else if (cell.w[0] == '#') {
					if (gCfg.EnableExportCommentRows) {
						columnIdx.seekToBegin();
						tmpArry = [];
						do {
							const cell = GetCellData(worksheet, columnIdx.curr26, rowIdx);
							tmpArry.push((cell && cell.w)?cell.w:'');
						}while(columnIdx.next);
						csvcontent += tmpArry.join(',') + gCfg.LineBreak;
					}
					break;
				}
				firstCol = false;
			}
			const value = cell && cell.w ? cell.w : '';
			if (gCfg.EnableTypeCheck) {
				if (!col.checker.CheckValue(cell)) {
					col.checker.CheckValue(cell);
					exception(`excel file "${yellow_ul(fileName)}" sheet "${yellow_ul(sheetName)}" CSV Cell "${yellow_ul(col.sid+(rowIdx).toString())}" format not match "${yellow_ul(value)}" with ${yellow_ul(col.checker.s)}!`);
					return;
				}
			}
			tmpArry.push(col.checker.GetCsvData(cell));
		}
		if (!firstCol) {
			csvcontent += tmpArry.join(',').replace(/\n/g, '\\n').replace(/\r/g, '') + gCfg.LineBreak;
		}
	}
	fs.writeFileSync(path.join(gCfg.Export.OutputDir, CSVName+'.csv'), csvcontent, {encoding:'utf8', flag:'w+'});
	logger(false, `${green('[SUCCESS]')} Output file "${yellow_ul(path.join(gCfg.Export.OutputDir, CSVName+'.csv'))}". Total use tick:${green((Date.now() - StartTick).toString())}`);
}

async function HandleExcelFile(fileName: string) {
	const extname = path.extname(fileName);
	if (extname != '.xls' && extname != '.xlsx') {
		return;
	}
	if (gCfg.ExcludeFileNames.indexOf(path.basename(fileName)) >= 0) {
		logger(true, `- Pass File "${fileName}"`);
		return;
	}
	if (path.basename(fileName).indexOf(`~$`) == 0) {
		logger(true, `- Pass File "${fileName}"`);
		return;
	}
	const excel = xlsx.readFile(fileName);
	if (excel == null) {
		exception(`excel ${yellow_ul(fileName)} open failure.`);
	}
	if (excel.Sheets == null) {
		return;
	}
	for (let sheetName of excel.SheetNames) {
		logger(true, `handle excel "${brightWhite(fileName)}" sheet "${yellow_ul(sheetName)}"`);
		const worksheet = excel.Sheets[sheetName];
		HandleWorkSheet(fileName, sheetName, worksheet);
	}
}

function main() {
	const StartTimer = Date.now();
	if (!fs.existsSync(gCfg.Export.OutputDir)) {
		fs.mkdirSync(gCfg.Export.OutputDir);
	}
	for (let fileOrPath of gCfg.IncludeFilesAndPath) {
		if (!path.isAbsolute(fileOrPath)) {
			fileOrPath = path.join(gRootDir, fileOrPath);
		}
		if (!fs.existsSync(fileOrPath)) {
			logger(false, `file or directory "${yellow_ul(fileOrPath)}" not found!`);
			continue;
		}
		if (fs.statSync(fileOrPath).isDirectory()) {
			HandleDir(fileOrPath);
		} else if (fs.statSync(fileOrPath).isFile()) {
			HandleExcelFile(fileOrPath);
		} else {
			exception(`UnHandle file or directory type : "${yellow_ul(fileOrPath)}"`);
		}
	}

	logger(false, `Total Use Tick : "${yellow_ul((Date.now() - StartTimer).toString())}"`);
	logger(false, yellow("----------------------------------------"));
	logger(false, yellow("-            DONE WITH ALL             -"));
	logger(false, yellow("----------------------------------------"));
}

////////////////////////////////////////////////////////////////////////////////
main()