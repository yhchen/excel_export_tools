import * as fs from 'fs-extra';
export { isString, isNumber, isArray, isObject, isBoolean } from 'lodash';
export { isDate } from 'moment';
import { isString } from 'util';
import { gCfg } from './config';
import { CTypeParser } from './CTypeParser';
import { CHightTypeChecker } from './CHighTypeChecker';

////////////////////////////////////////////////////////////////////////////////
//#region console color
import chalk from 'chalk';
export const yellow_ul = chalk.yellow.underline;	//yellow under line
export const orange_ul = chalk.magentaBright.underline.bold;	//orange under line
export const yellow = chalk.yellow;
export const red = chalk.redBright;
export const green = chalk.greenBright;
export const brightWhite = chalk.whiteBright.bold
//#endregion

////////////////////////////////////////////////////////////////////////////////
//#region Logger
export const enum E_ERROR_LEVEL {
	EXECUTE_FAILURE = -1,
	INIT_EXTENDS = -1001,
}
export function logger(...args: any[]) {
	console.log(...args);
}
export function debug(...args: any[]) {
	if (!gCfg.EnableDebugOutput)
		return;
	logger(...args);
}
export function warn(txt: string): void {
	logger(`${orange_ul(`+ [WARN] `)} ${txt}\n`);
}
let ExceptionLogLst = new Array<string>();
export function exception(txt: string, ex?: any): never {
	exceptionRecord(txt, ex);
	throw txt;
}
// record exception not throw.
export function exceptionRecord(txt: string, ex?: any): void {
	const LOG_CTX = `${red(`+ [ERROR] `)} ${txt}\n${red(ex ? ex : '')}\n`;
	ExceptionLogLst.push(LOG_CTX);
	logger(LOG_CTX);
}
//#endregion

////////////////////////////////////////////////////////////////////////////////
//#region base function
export function StrNotEmpty(s: any): s is string {
	if (isString(s)) {
		return s.trim().length > 0;
	}
	return false;
}
//#endregion

////////////////////////////////////////////////////////////////////////////////
//#region Time Profile
/************* total use tick ****************/
let BeforeExistHandler: () => void;
export function SetBeforeExistHandler(handler: () => void) {
	BeforeExistHandler = handler;
}
export module TimeUsed {
	export function LastElapse(): string {
		const Now = Date.now();
		const elpase = Now - _LastAccess;
		_LastAccess = Now;
		return elpase.toString() + 'ms';
	}

	export function TotalElapse(): string {
		return ((Date.now() - _StartTime) / 1000).toString() + 's';
	}

	const _StartTime = Date.now();
	let _LastAccess = _StartTime;

	process.addListener('beforeExit', () => {
		process.removeAllListeners('beforeExit');
		const HasException = ExceptionLogLst.length > 0;
		if (BeforeExistHandler && !HasException) {
			BeforeExistHandler();
		}
		const color = HasException ? red : green;
		logger(color(`----------------------------------------`));
		logger(color(`-          ${HasException ? 'Got Exception !!!' : '    Well Done    '}           -`));
		logger(color(`----------------------------------------`));
		logger(`Total Use Tick : "${yellow_ul(TotalElapse())}"`);

		if (HasException) {
			logger(red("Exception Logs:"));
			logger(ExceptionLogLst.join('\n'));
			process.exit(-1);
		} else {
			process.exit(0);
		}
	});
}
//#endregion

////////////////////////////////////////////////////////////////////////////////
//#region async Worker Monitor
export class AsyncWorkMonitor {
	public addWork(cnt: number = 1) {
		this._leftCnt += cnt;
	}
	public decWork(cnt: number = 1) {
		this._leftCnt -= cnt;
	}
	public async WaitAllWorkDone(): Promise<boolean> {
		if (this._leftCnt <= 0)
			return true;
		while (true) {
			if (this._leftCnt <= 0) {
				return true;
			}
			await this.delay()
		}
	}
	public async delay(ms: number = 0) {
		return new Promise(resolve => setTimeout(resolve, ms));
	}
	private _leftCnt = 0;
}
//#endregion

////////////////////////////////////////////////////////////////////////////////
//#region Datas
// excel gen data table
export enum ESheetRowType {
	header = 1, // name row
	type = 2, // type
	data = 3, // data
	comment = 4, // comment
}
export type SheetRow = {
	type: ESheetRowType,
	values: Array<any>,
	cIdx: number;
}
export type SheetHeader = {
	name: string, // name
	stype: string, // type string
	cIdx: number, // header idx
	typeChecker: CTypeParser, // type checker
	comment: boolean; // is comment line?
	color: string; // group
	highCheck?: CHightTypeChecker;
}
export class SheetDataTable {
	public constructor(name: string, filename: string) {
		this.name = name;
		this.filename = filename;
	}
	// check sheet column contains key
	public checkColumnContainsValue(columnName: string, value: any): boolean {
		if (!this.columnKeysMap.has(columnName)) {
			if (!this.makeColumnKeyMap(columnName)) {
				exception(`CALL [checkColumnContainsValue] failure: sheet column name ${yellow_ul(this.name + '.' + columnName)}`);
				return false;
			}
		}
		let s = this.columnKeysMap.get(columnName);
		return s != undefined && s.has(value);
	}
	public containsColumName(name: string): boolean {
		for (const header of this.arrTypeHeader) {
			if (header.name == name)
				return true;
		}
		return false;
	}
	public name: string;
	public filename: string;
	public arrTypeHeader = new Array<SheetHeader>();
	public arrValues = new Array<SheetRow>();

	private makeColumnKeyMap(columnName: string): boolean {
		for (let i = 0; i < this.arrTypeHeader.length; ++i) {
			if (this.arrTypeHeader[i].name == columnName) {
				const s = new Set<any>();
				for (const row of this.arrValues) {
					if (row.type != ESheetRowType.data) continue;
					s.add(row.values[i]);
				}
				this.columnKeysMap.set(columnName, s);
				return true;
			}
		}
		return false;
	}
	private columnKeysMap = new Map<string, Set<any>>();
}
// all export data here
export const ExportExcelDataMap = new Map<string, SheetDataTable>();

// line breaker
export let LineBreaker = '\n';
export function SetLineBreaker(v: string) {
	LineBreaker = v;
}

//#endregion

////////////////////////////////////////////////////////////////////////////////
//#region export config

export type ExportCfg = {
	type: string;
	OutputDir: string;
	UseDefaultValueIfEmpty: boolean;
	GroupFilter: Array<string>;
	ExportTemple?: string;
	ExtName?: string;
}
// export template
export abstract class IExportWrapper {
	public constructor(exportCfg: ExportCfg) {
		this._exportCfg = exportCfg;
	}
	public abstract get DefaultExtName(): string;
	public async ExportToAsync(dt: SheetDataTable, endCallBack: (ok: boolean) => void): Promise<boolean> {
		let ok = false;
		try {
			ok = await this.ExportTo(dt);
		} catch (ex) {
			// do nothing...
		}
		if (endCallBack) {
			endCallBack(ok);
		}
		return ok;
	}
	public async ExportToGlobalAsync(endCallBack: (ok: boolean) => void): Promise<boolean> {
		let ok = false;
		try {
			ok = await this.ExportGlobal();
		} catch (ex) {
			// do nothing...
		}
		if (endCallBack) {
			endCallBack(ok);
		}
		return ok;
	}
	protected abstract async ExportTo(dt: SheetDataTable): Promise<boolean>;
	protected abstract async ExportGlobal(): Promise<boolean>;
	protected CreateDir(outdir: string): boolean {
		if (!fs.existsSync(outdir)) {
			fs.ensureDirSync(outdir);
			return fs.existsSync(outdir);
		}
		return true;
	}

	protected IsFile(s: string): boolean {
		const ext = this._exportCfg.ExtName || this.DefaultExtName;
		const idx = s.lastIndexOf(ext);
		if (idx < 0)
			return false;
		return (idx + ext.length == s.length);
	}

	protected _exportCfg: ExportCfg;
}


export function ExecGroupFilter(arrGrpFilters: Array<string>, arrHeader: Array<SheetHeader>): Array<SheetHeader> {
	let result = new Array<SheetHeader>();
	if (arrGrpFilters.length <= 0)
		return result;
	// translate
	const RealFilter = new Array<string>();
	for (const ele of arrGrpFilters) {
		let found = false;
		for (const name in gCfg.ColorToGroupMap) {
			if ((<any>gCfg.ColorToGroupMap)[name] == ele) {
				RealFilter.push(name);
				found = true;
				break;
			}
		}
		if (!found) {
			logger(`Filter Name ${yellow_ul(ele)} Not foud In ${yellow_ul('ColorToGroupMap')}. Ignore It!`);
		}
	}
	// filter
	if (RealFilter.includes('*')) {
		return arrHeader;
	}
	for (const header of arrHeader) {
		if (RealFilter.includes(header.color)) {
			result.push(header);
		}
	}

	return result;
}


export type ExportWrapperFactory = (cfg: ExportCfg) => IExportWrapper;
export const ExportWrapperMap = new Map<string, ExportWrapperFactory>([
	['csv', require('./export/export_to_csv')],
	['json', require('./export/export_to_json')],
	['js', require('./export/export_to_js')],
	['tsd', require('./export/export_to_tsd')],
	['lua', require('./export/export_to_lua')],
]);

//#endregion

////////////////////////////////////////////////////////////////////////////////
//#region Format Converter
export module FMT26 {
	export function NumToS26(num: number): string {
		let result = "";
		++num;
		while (num > 0) {
			let m = num % 26;
			if (m == 0) m = 26;
			result = String.fromCharCode(m + 64) + result;
			num = (num - m) / 26;
		}
		return result;
	}

	export function S26ToNum(str: string): number {
		let result = 0;
		let ss = str.toUpperCase();
		for (let i = str.length - 1, j = 1; i >= 0; i-- , j *= 26) {
			let c = ss[i];
			if (c < 'A' || c > 'Z')
				return 0;
			result += (c.charCodeAt(0) - 64) * j;
		}
		return --result;
	}
}

//#endregion

