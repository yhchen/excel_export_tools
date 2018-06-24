import * as xlsx from 'xlsx';

import gCfg from "./config.json";
import * as fs from 'fs';
import * as path from 'path';
import * as chalk from 'chalk';

import {CTypeChecker} from "./TypeChecker";
import { start } from 'repl';

/*************** console color ***************/
const yellow = chalk.default.yellowBright;
const red = chalk.default.redBright;
const green = chalk.default.greenBright;
const brightWhite = chalk.default.whiteBright.bold
function logger(debugMode: boolean, ...args: any[]) {
	if (!gCfg.enableDebugOutput && debugMode) {
		return;
	}
	console.log(...args);
}
function trace(...args: any[]) { logger(true, ...args); }
function exception(txt: string, ex?:any) {
	logger(false, brightWhite(txt));
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
	public seekToBegin() { this._curr = 0; }
	public get next(): string|undefined {
		if (this.end) {
			return undefined;
		}
		return XlsColumnIter.NumToS26(++this._curr);
	}
	public get curr26(): string { return XlsColumnIter.NumToS26(this._curr); }
	public get curr10(): number { return this._curr; }
	public get end(): boolean { return this._curr >= this._max; }

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
			result += String.fromCharCode(m + 64);
			num = (num - m) / 26;
		}
		return result;
	}
	private _curr = 0;
	private _max: number;
	private static readonly VAILDWORD = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
}

function GetCellData(worksheet: xlsx.WorkSheet, column: string, row: number): xlsx.CellObject|undefined {
	return worksheet[column + row.toString()];
}

function HandleWorkSheet(fileName: string, sheetName: string, worksheet: xlsx.WorkSheet): void{
	const StartTick = Date.now();
	let ColumnMax = 'A';
	let RowMax = 0;
	let ColumnArry = new Array<{sid:string, id:number, name:string, checker:CTypeChecker}>();
	// find max column and rows
	{
		const REF = worksheet["!ref"];
		if (!REF) {
			logger(false, `excel file [${yellow(fileName)}] sheet [${yellow(sheetName)}] IS EMPTY. Ignore it!`);
			return;
		}
		let SPREF = REF.split(":");
		if (SPREF.length != 2) {
			logger(false, `excel file [${yellow(fileName)}] sheet [${yellow(sheetName)}] [!ref] = [${REF}] format error!!`);
			return;
		}
		ColumnMax = SPREF[1].toUpperCase().replace(/([0-9]*)/g, '');
		RowMax = parseInt(SPREF[1].toUpperCase().replace(/([A-Z]*)/g, ''));
	}
	// find csv name
	if (worksheet[gCfg.CSVNameCellID] == undefined || NullStr(worksheet[gCfg.CSVNameCellID].w)) {
		logger(false, `excel file [${yellow(fileName)}] sheet [${yellow(sheetName)}] CSV name not defined. Ignore it!`);
		return;
	}
	const CSVName = worksheet[gCfg.CSVNameCellID].w;
	if (gCfg.ExcludeCsvTableNames.indexOf(CSVName) >= 0) {
		logger(true, `- Pass CSV [${CSVName}]`);
	}

	let rowIdx = 2;
	let csvcontent = '';
	let columnIdx = new XlsColumnIter(ColumnMax);
	let tmpArry: string[] = [];
	// find column name
	for (; rowIdx <= RowMax; ++rowIdx) {
		const firstCell = GetCellData(worksheet, 'A', rowIdx);
		if (firstCell == undefined || firstCell.w == undefined || NullStr(firstCell.w) || firstCell.w[0] == '#') {
			continue;
		}
		columnIdx.seekToBegin();
		tmpArry = [];
		do {
			const colName = columnIdx.curr26;
			let cell = GetCellData(worksheet, colName, rowIdx);
			if (cell == undefined || cell.w == undefined || NullStr(cell.w) || cell.w[0] == '#') {
				continue;
			}
			ColumnArry.push({id:columnIdx.curr10, sid:colName, name:cell.w, checker:<any>undefined});
			tmpArry.push(cell.w);
		}while(columnIdx.next);
		++rowIdx;
		break;
	}
	csvcontent += tmpArry.join(',') + gCfg.LineBreak;
	// find type
	for (; rowIdx <= RowMax; ++rowIdx) {
		const firstCell = GetCellData(worksheet, ColumnArry[0].sid, rowIdx);
		if (firstCell == undefined || firstCell.w == undefined || NullStr(firstCell.w) || firstCell.w[0] == '#') {
			continue;
		}
		if (firstCell.w[0] != '*') {
			exception(`excel file [${yellow(fileName)}] sheet [${yellow(sheetName)}] CSV Type Column not found!`);
		}
		tmpArry = [];
		for (let col of ColumnArry) {
			let cell = GetCellData(worksheet, col.sid, rowIdx);
			if (cell == undefined || cell.w == undefined) {
				exception(`excel file [${yellow(fileName)}] sheet [${yellow(sheetName)}] CSV Type Column [${yellow(col.name)}] not found!`);
				return;
			}
			try {
				col.checker = new CTypeChecker(col.id <= 1 ? cell.w.substr(1):cell.w);
				let v = cell.w.replace(/"/g, `""`);
				if (v.indexOf(',') >= 0) {
					v = '"' + v + '"';
				}
				tmpArry.push(`${v}`);
			} catch (ex) {
				exception(`excel file [${yellow(fileName)}] sheet [${yellow(sheetName)}] CSV Type Column [${yellow(col.name)}] format error [${yellow(cell.w)}]!`, ex);
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
			let cell = GetCellData(worksheet, col.sid, rowIdx);
			if (firstCol) {
				if (cell == undefined || cell.w == undefined || NullStr(cell.w) || cell.w[0] == '#') {
					break;
				}
				firstCol = false;
			}
			let value = cell && cell.w ? cell.w : '';
			if (gCfg.enableTypeCheck) {
				if (!col.checker.CheckValue(cell)) {
					// col.checker.CheckValue(cell);
					exception(`excel file [${yellow(fileName)}] sheet [${yellow(sheetName)}] CSV Cell [${yellow(col.sid+(rowIdx+1).toString())}] format not match [${yellow(value)}]!`);
					return;
				}
			}
			tmpArry.push(col.checker.GetValue(cell));
		}
		if (!firstCol) {
			csvcontent += tmpArry.join(',').replace(/\n/g, '\\n').replace(/\r/g, '') + gCfg.LineBreak;
		}
	}
	fs.writeFileSync(path.join(gCfg.OutputDir, CSVName+'.csv'), csvcontent, {encoding:'utf8', flag:'w+'});
	logger(false, `${green('[SUCCESS]')} Output file [${yellow(path.join(gCfg.OutputDir, CSVName+'.csv'))}]. Total use tick:${green((Date.now() - StartTick).toString())}`);
}

async function HandleExcelFile(fileName: string) {
	const extname = path.extname(fileName);
	if (extname != '.xls' && extname != '.xlsx') {
		return;
	}
	if (gCfg.ExcludeFileNames.indexOf(path.basename(fileName)) >= 0) {
		logger(true, `- Pass File [${fileName}]`);
		return;
	}
	const excel = xlsx.readFile(fileName);
	if (excel == null) {
		exception(`excel ${yellow(fileName)} open failure.`);
	}
	if (excel.Sheets == null) {
		return;
	}
	for (let sheetName of excel.SheetNames) {
		logger(true, `handle excel [${brightWhite(fileName)}] sheet [${yellow(sheetName)}]`);
		const worksheet = excel.Sheets[sheetName];
		HandleWorkSheet(fileName, sheetName, worksheet);
	}
}

function main() {
	const StartTimer = Date.now();
	if (!fs.existsSync(gCfg.OutputDir)) {
		fs.mkdirSync(gCfg.OutputDir);
	}
	for (let fileOrPath of gCfg.IncludeFilesAndPath) {
		if (!path.isAbsolute(fileOrPath)) {
			fileOrPath = path.join(gRootDir, fileOrPath);
		}
		if (!fs.existsSync(fileOrPath)) {
			logger(false, `file or directory [${yellow(fileOrPath)}] not found!`);
			continue;
		}
		if (fs.statSync(fileOrPath).isDirectory()) {
			HandleDir(fileOrPath);
		} else if (fs.statSync(fileOrPath).isFile()) {
			HandleExcelFile(fileOrPath);
		} else {
			exception(`UnHandle file or directory type : [${yellow(fileOrPath)}]`);
		}
	}

	logger(false, `Total Use Tick : [${yellow((Date.now() - StartTimer).toString())}]`);
	logger(false, yellow("----------------------------------------"));
	logger(false, yellow("-            DONE WITH ALL             -"));
	logger(false, yellow("----------------------------------------"));
}

////////////////////////////////////////////////////////////////////////////////
main()