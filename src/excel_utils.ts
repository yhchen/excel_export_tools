import * as path from 'path';
import * as fs from 'fs-extra-promise';
import * as xlsx from 'xlsx';
import * as utils from './utils';
import { ETypeNames, CTypeParser } from './CTypeParser';
import { gCfg } from './config'
import { CHightTypeChecker } from './CHighTypeChecker';

export async function HandleExcelFile(fileName: string): Promise<boolean> {
	try {
		const extname = path.extname(fileName);
		if (extname != '.xls' && extname != '.xlsx') {
			return true;
		}
		if (path.basename(fileName)[0] == '!') {
			utils.debug(`- Pass File "${fileName}"`);
			return true;
		}
		if (path.basename(fileName).indexOf(`~$`) == 0) {
			utils.debug(`- Pass File "${fileName}"`);
			return true;
		}
		let opt: xlsx.ParsingOptions = {
			type: "buffer",
			// codepage: 0,//If specified, use code page when appropriate **
			cellFormula: false,//Save formulae to the .f field
			cellHTML: false,//Parse rich text and save HTML to the .h field
			cellText: true,//Generated formatted text to the .w field
			cellDates: true,//Store dates as type d (default is n)
			cellStyles: true,//Store style/theme info to the .s field
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
			return true;
		}
		for (let sheetName of excel.SheetNames) {
			utils.debug(`- Handle excel "${utils.brightWhite(fileName)}" Sheet "${utils.yellow_ul(sheetName)}"`);
			const worksheet = excel.Sheets[sheetName];
			const datatable = HandleWorkSheet(fileName, sheetName, worksheet);
			if (datatable) {
				const oldDataTable = utils.ExportExcelDataMap.get(datatable.name);
				if (oldDataTable) {
					utils.exception(`found duplicate file name : ${utils.yellow_ul(datatable.name)} \n`
						+ `at excel ${utils.yellow_ul(fileName)} \n`
						+ `and excel ${utils.yellow_ul(oldDataTable.filename)}`);
				}
				utils.ExportExcelDataMap.set(datatable.name, datatable);
			}
		}
	} catch (ex) {
		return false;
	}
	return true;
}

////////////////////////////////////////////////////////////////////////////////
//#region private side
function GetCellData(worksheet: xlsx.WorkSheet, c: number, r: number): xlsx.CellObject | undefined {
	const cell = xlsx.utils.encode_cell({ c, r });
	return worksheet[cell];
}

// get cell front groud color
function GetCellFrontGroudColor(cell: xlsx.CellObject): string {
	if (!cell.s || !cell.s.fgColor || !cell.s.fgColor.rgb)
		return '000000'; // default return white
	return cell.s.fgColor.rgb;
}

function HandleWorkSheetTypeColumn(worksheet: xlsx.WorkSheet,
	rIdx: number, RowMax: number, ColumnMax: number,
	fileName: string, sheetName: string, DataTable: utils.SheetDataTable,
	arrHeaderName: Array<{ cIdx: number, name: string, parser: CTypeParser, color: string }>): number {

	rIdx = HandleWorkSheetHighTypeColumn(worksheet, rIdx, RowMax, ColumnMax, fileName, sheetName, DataTable, arrHeaderName);
	for (; rIdx <= RowMax; ++rIdx) {
		const firstCell = GetCellData(worksheet, arrHeaderName[0].cIdx, rIdx);
		if (firstCell == undefined || !utils.StrNotEmpty(firstCell.w)) {
			continue;
		}
		if (firstCell.w[0] == '#') {
			continue;
		}

		if (firstCell.w[0] != '*') {
			utils.exception(`Excel "${utils.yellow_ul(fileName)}" Sheet "${utils.yellow_ul(sheetName)}" Sheet Type Row not found!`);
		}
		firstCell.w = firstCell.w.substr(1); // skip '*'
		const tmpArry = [];
		let typeHeader = new Array<utils.SheetHeader>();
		for (const col of arrHeaderName) {
			const cell = GetCellData(worksheet, col.cIdx, rIdx);
			if (cell == undefined || cell.w == undefined) {
				utils.exception(`Excel "${utils.yellow_ul(fileName)}" ` +
					`Sheet "${utils.yellow_ul(sheetName)}"  Type Row "${utils.yellow_ul(col.name)}" not found!`);
				return -1;
			}
			try {
				col.parser = new CTypeParser(cell.w);
				tmpArry.push(cell.w);
				typeHeader.push({ name: col.name, cIdx: col.cIdx, typeChecker: col.parser, stype: cell.w, comment: false, color: col.color });
			} catch (ex) {
				// new CTypeParser(cell.w); // for debug used
				utils.exception(`Excel "${utils.yellow_ul(fileName)}" Sheet "${utils.yellow_ul(sheetName)}" Sheet Type Row`
					+ ` "${utils.yellow_ul(col.name)}" format error "${utils.yellow_ul(cell.w)}"!`, ex);
			}
		}
		DataTable.arrTypeHeader = typeHeader;
		DataTable.arrValues.push({ type: utils.ESheetRowType.type, values: tmpArry, cIdx: rIdx });
		++rIdx;
		break;
	}
	rIdx = HandleWorkSheetHighTypeColumn(worksheet, rIdx, RowMax, ColumnMax, fileName, sheetName, DataTable, arrHeaderName);
	return rIdx;
}

function HandleWorkSheetHighTypeColumn(worksheet: xlsx.WorkSheet,
	rIdx: number, RowMax: number, ColumnMax: number,
	fileName: string, sheetName: string, DataTable: utils.SheetDataTable,
	arrHeaderName: Array<{ cIdx: number, name: string, parser: CTypeParser, color: string }>): number {
	for (; rIdx <= RowMax; ++rIdx) {
		const firstCell = GetCellData(worksheet, arrHeaderName[0].cIdx, rIdx);
		if (firstCell == undefined || !utils.StrNotEmpty(firstCell.w)) {
			continue;
		}
		if (firstCell.w[0] == '#') {
			continue;
		}
		// found '*' or not not '@'. return rIdx for continue
		if (firstCell.w[0] == '*' || firstCell.w[0] != '@') {
			return rIdx;
		}
		firstCell.w = firstCell.w.substr(1); // skip '@'
		for (let i = 0; i < arrHeaderName.length; ++i) {
			const cell = GetCellData(worksheet, arrHeaderName[i].cIdx, rIdx);
			if (cell != undefined && cell.w != undefined && cell.w != '') {
				const header = DataTable.arrTypeHeader[i];
				header.highCheck = new CHightTypeChecker(cell.w);
			}
		}
		++rIdx;
		break;
	}
	return rIdx;
}

function HandleWorkSheetNameColumn(worksheet: xlsx.WorkSheet,
	rIdx: number, RowMax: number, ColumnMax: number,
	fileName: string, sheetName: string, DataTable: utils.SheetDataTable,
	arrHeaderName: Array<{ cIdx: number, name: string, parser: CTypeParser, color: string }>): number {
	// find column name
	for (; rIdx <= RowMax; ++rIdx) {
		const firstCell = GetCellData(worksheet, 0, rIdx);
		if (firstCell == undefined || !utils.StrNotEmpty(firstCell.w)) {
			continue;
		}
		if (firstCell.w[0] == '#') {
			continue;
		}
		const tmpArry:any[] = [];
		for (let cIdx = 0; cIdx <= ColumnMax; ++cIdx) {
			const cell = GetCellData(worksheet, cIdx, rIdx);
			if (cell == undefined || !utils.StrNotEmpty(cell.w) || cell.w[0] == '#') {
				continue;
			}
			const colGrp = GetCellFrontGroudColor(cell);
			const NamedGrp = (<any>gCfg.ColorToGroupMap)[colGrp];
			if (NamedGrp == undefined) {
				utils.warn(`Excel "${utils.yellow_ul(fileName)}" Sheet "${utils.yellow_ul(sheetName)}" `
					+ `Cell "${utils.yellow_ul(utils.FMT26.NumToS26(cIdx) + (rIdx).toString())}" `
					+ `Name Group ${utils.yellow_ul(colGrp)} Invalid"!`);
			}
			arrHeaderName.push({ cIdx, name: cell.w, parser: new CTypeParser(ETypeNames.string), color: colGrp });
			tmpArry.push(cell.w);
		}
		DataTable.arrValues.push({ type: utils.ESheetRowType.header, values: tmpArry, cIdx: rIdx });
		++rIdx;
		break;
	}
	return rIdx;
}

function HandleWorkSheetDataColumn(worksheet: xlsx.WorkSheet,
	rIdx: number, RowMax: number, ColumnMax: number,
	fileName: string, sheetName: string, DataTable: utils.SheetDataTable,
	arrHeaderName: Array<{ cIdx: number, name: string, parser: CTypeParser, color: string }>): number {
	for (; rIdx <= RowMax; ++rIdx) {
		let firstCol = true;
		const tmpArry = [];
		for (let col of arrHeaderName) {
			const cell = GetCellData(worksheet, col.cIdx, rIdx);
			if (firstCol) {
				if (cell == undefined || !utils.StrNotEmpty(cell.w)) {
					break;
				}
				else if (cell.w[0] == '#') {
					break;
				}
				firstCol = false;
			}
			const value = cell && cell.w ? cell.w : '';
			let colObj;
			try {
				colObj = col.parser.ParseContent(cell);
				tmpArry.push(colObj);
			} catch (ex) {
				// col.checker.ParseDataStr(cell);
				utils.exceptionRecord(`Excel "${utils.yellow_ul(fileName)}" Sheet "${utils.yellow_ul(sheetName)}" `
					+ `Cell "${utils.yellow_ul(utils.FMT26.NumToS26(col.cIdx) + (rIdx + 1).toString())}" `
					+ `Parse Data "${utils.yellow_ul(value)}" With ${utils.yellow_ul(col.parser.s)} `
					+ `Cause utils.exception "${utils.red(ex)}"!`);
				return -1;
			}
			if (gCfg.EnableTypeCheck) {
				if (!col.parser.CheckContentVaild(colObj)) {
					col.parser.CheckContentVaild(colObj); // for debug used
					utils.exceptionRecord(`Excel "${utils.yellow_ul(fileName)}" Sheet "${utils.yellow_ul(sheetName)}" `
						+ `Cell "${utils.yellow_ul(utils.FMT26.NumToS26(col.cIdx) + (rIdx + 1).toString())}" `
						+ `format not match "${utils.yellow_ul(value)}" with ${utils.yellow_ul(col.parser.s)}!`);
					return -1;
				}
			}
		}
		if (!firstCol) {
			DataTable.arrValues.push({ type: utils.ESheetRowType.data, values: tmpArry, cIdx: rIdx });
		}
	}
	return rIdx;
}

function GetCellIndex(cellName: string): number {
	const idx = cellName.search(/[0-9]/);
	if (idx < 0) return 0;
	return parseInt(cellName.substr(idx)) - 1;
}

function HandleWorkSheet(fileName: string, sheetName: string, worksheet: xlsx.WorkSheet): utils.SheetDataTable | undefined {
	if (worksheet['!ref'] == undefined) {
		utils.debug(`- Pass Sheet "${sheetName}" : Sheet is empty`);
		return;
	}
	// find csv name
	const TableCell: xlsx.CellObject = worksheet[gCfg.CSVNameCellID];
	if (TableCell == undefined) {
		utils.logger(false, `excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" Sheet name not defined. Ignore it!`);
		return;
	}
	const TableName = TableCell.w;
	if (!utils.StrNotEmpty(TableName)) {
		utils.logger(false, `excel file "${utils.yellow_ul(fileName)}" sheet "${utils.yellow_ul(sheetName)}" Sheet name not defined. Ignore it!`);
		return;
	}

	if (TableName[0] == "!") {
		utils.debug(`- Pass Sheet "${TableName}" : Sheet Name start with "!"`);
		return;
	}

	const Range = xlsx.utils.decode_range(<string>worksheet['!ref']);
	const ColumnMax = Range.e.c;
	const RowMax = Range.e.r;
	const arrHeaderName = new Array<{ cIdx: number, name: string, parser: CTypeParser, color: string }>();
	// find max column and rows
	let rIdx = GetCellIndex(gCfg.CSVNameCellID) + 1;
	const DataTable = new utils.SheetDataTable(TableName, fileName);
	// find column name
	rIdx = HandleWorkSheetNameColumn(worksheet, rIdx, RowMax, ColumnMax, fileName, TableName, DataTable, arrHeaderName);
	if (rIdx < 0)
		return;
	// find type
	rIdx = HandleWorkSheetTypeColumn(worksheet, rIdx, RowMax, ColumnMax, fileName, TableName, DataTable, arrHeaderName);
	if (rIdx < 0)
		return;
	// handle datas
	rIdx = HandleWorkSheetDataColumn(worksheet, rIdx, RowMax, ColumnMax, fileName, TableName, DataTable, arrHeaderName);
	if (rIdx < 0)
		return;
	return DataTable;
}

//#endregion
