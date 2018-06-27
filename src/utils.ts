export { isString, isNumber, isArray, isObject, isBoolean, isDate } from 'util';
import { isString, isNumber, isArray, isObject, isBoolean, isDate } from 'util';

////////////////////////////////////////////////////////////////////////////////
/*************** console color ***************/
import * as chalk from 'chalk';
export const yellow_ul = chalk.default.yellow.underline;	//yellow under line
export const yellow = chalk.default.yellow;
export const red = chalk.default.redBright;
export const green = chalk.default.greenBright;
export const brightWhite = chalk.default.whiteBright.bold

////////////////////////////////////////////////////////////////////////////////
/************ logger function **************/
export let EnableDebugOutput: boolean = true;
export function SetEnableDebugOutput(b: boolean) { EnableDebugOutput = b; }
export function logger(debugMode: boolean, ...args: any[]) {
	if (!EnableDebugOutput && debugMode) {
		return;
	}
	console.log(...args);
}
function trace(...args: any[]) { logger(true, ...args); }
let ExceptionLog = '';
export function exception(txt: string, ex?:any) {
	ExceptionLog += `${red(`+ [ERROR] `)} ${txt}\n`
	logger(false, red(`[ERROR] `) + txt);
	if (ex) { logger(false, red(ex)); }
	throw txt;
}

////////////////////////////////////////////////////////////////////////////////
/************* base function ***************/
export function NullStr(s: string) {
	if (isString(s)) {
		return s.trim().length <= 0;
	}
	return true;
}


////////////////////////////////////////////////////////////////////////////////
/************* total use tick ****************/
// timer calc
export module TimeUsed
{
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

	process.addListener('beforeExit', ()=>{
		process.removeAllListeners('beforeExit');
		const color = NullStr(ExceptionLog) ? green : yellow;
		logger(false, color("----------------------------------------"));
		logger(false, color("-            DONE WITH ALL             -"));
		logger(false, color("----------------------------------------"));
		logger(false, `Total Use Tick : "${yellow_ul(TotalElapse())}"`);

		if (!NullStr(ExceptionLog)) {
			logger(false, red("Exception Logs:"));
			logger(false, ExceptionLog);
			process.exit(-1);
		} else {
			process.exit(0);
		}
	});
}

// excel gen data table
export type DataTable = {
	name: string,
	datas: Array<Array<string>>;
}
// all export data here
export const ExportExcelDataMap = new Map<string, DataTable>();

// line breaker
export let LineBreaker = '\n';
export function SetLineBreaker(v: string) { LineBreaker = v; }
// export template
export abstract class IExportWrapper {
	public abstract async exportTo(dt: DataTable, outdir: string): Promise<boolean>;
}
