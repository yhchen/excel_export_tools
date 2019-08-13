import * as path from 'path';
import * as fs from 'fs-extra-promise';
import * as utils from './utils';
import { gCfg, gRootDir, gGlobalIgnoreDirName } from './config'
import { HandleExcelFile } from './excel_utils'

const gExportWrapperLst = new Array<utils.IExportWrapper>();
function InitEnv(): boolean {
	for (const exportCfg of gCfg.Export) {
		const Constructor = utils.ExportWrapperMap.get(exportCfg.type);
		if (Constructor == undefined) {
			utils.exceptionRecord(utils.red(`Export is not currently supported for the current type "${utils.yellow_ul(exportCfg.type)}"!\n` +
				`ERROR : Export constructor not found. initialize failure!`));
			return false;
		}
		const Exportor = Constructor.call(Constructor, exportCfg);
		if (Exportor) {
			if ((<any>exportCfg).ExtName == undefined) {
				(<any>exportCfg).ExtName = Exportor.DefaultExtName;
			}
			gExportWrapperLst.push(Exportor);
		}
	}
	return true;
}

export async function execute(): Promise<boolean> {
	if (!InitEnv()) {
		throw `InitEnv failure!`;
	}
	if (!await HandleReadData()) {
		throw `handle read excel data failure.`;
	}
	if (!HandleHighLevelTypeCheck()) {
		throw `handle check hight level type failure.`;
	}
	if (!await HandleExportAll()) {
		throw `handle export failure.`;
	}
	return true;
}

////////////////////////////////////////////////////////////////////////////////
//#region private side
const WorkerMonitor = new utils.AsyncWorkMonitor();
async function HandleExcelFileWork(fileName: string, cb: (ret:boolean)=>void): Promise<void> {
	WorkerMonitor.addWork();
	cb(await HandleExcelFile(fileName));
	WorkerMonitor.decWork();
}

async function HandleDir(dirName: string, cb: (ret:boolean)=>void): Promise<void> {
	if (gGlobalIgnoreDirName.has(path.basename(dirName))) {
		utils.logger(`ignore folder ${dirName}`);
		return;
	}
	WorkerMonitor.addWork();
	const pa = await fs.readdirAsync(dirName);
	pa.forEach(function (fileName) {
		const filePath = path.join(dirName, fileName);
		let info = fs.statSync(filePath);
		if (!info.isFile()) {
			return;
		}
		HandleExcelFileWork(filePath, cb);
	});
	WorkerMonitor.decWork();
}

async function HandleReadData(): Promise<boolean> {
	let ret = true;
	const cb = (v: boolean) => {
		ret = ret && v;
	}
	for (let fileOrPath of gCfg.IncludeFilesAndPath) {
		if (!path.isAbsolute(fileOrPath)) {
			fileOrPath = path.join(gRootDir, fileOrPath);
		}
		if (!fs.existsSync(fileOrPath)) {
			utils.exception(`file or directory "${utils.yellow_ul(fileOrPath)}" not found!`);
			break;
		}
		if (fs.statSync(fileOrPath).isDirectory()) {
			HandleDir(fileOrPath, cb);
		} else if (fs.statSync(fileOrPath).isFile()) {
			HandleExcelFileWork(fileOrPath, cb);
		} else {
			utils.exception(`UnHandle file or directory type : "${utils.yellow_ul(fileOrPath)}"`);
		}
	}
	await WorkerMonitor.delay();
	await WorkerMonitor.WaitAllWorkDone();
	utils.logger(`${utils.green('[SUCCESS]')} READ ALL SHEET DONE. Total Use Tick : ${utils.green(utils.TimeUsed.LastElapse())}`);
	return ret;
}

function HandleHighLevelTypeCheck(): boolean {
	if (!gCfg.EnableTypeCheck) {
		return true;
	}
	let foundError = false;
	for (const kv of utils.ExportExcelDataMap) {
		const database = kv[1];
		for (let colIdx = 0; colIdx < database.arrTypeHeader.length; ++colIdx) {
			let header = database.arrTypeHeader[colIdx];
			if (!header.highCheck) continue;
			try {
				header.highCheck.init();
			} catch (ex) {
				utils.exception(`Excel "${utils.yellow_ul(database.filename)}" Sheet "${utils.yellow_ul(database.name)}" High Type`
					+ ` "${utils.yellow_ul(header.name)}" format error "${utils.yellow_ul(header.highCheck.s)}"!`);
			}
			for (let rowIdx = 0; rowIdx < database.arrValues.length; ++rowIdx) {
				const row = database.arrValues[rowIdx];
				if (row.type != utils.ESheetRowType.data) continue;
				const data = row.values[colIdx];

				if (!data) continue;
				try {
					if (!header.highCheck.checkType(data)) {
						throw '';
					}
				} catch (ex) {
					foundError = true;
					// header.highCheck.checkType(data); // for debug
					utils.exceptionRecord(`Excel "${utils.yellow_ul(database.filename)}" `
						+ `Sheet Row "${utils.yellow_ul(database.name + '.' + utils.yellow_ul(header.name))}" High Type format error`
						+ `Cell "${utils.yellow_ul(utils.FMT26.NumToS26(header.cIdx) + (row.cIdx + 1).toString())}" `
						+ ` "${utils.yellow_ul(data)}"!`, ex);
				}
			}
		}
	}
	utils.logger(`${foundError ? utils.red('[FAILURE]') : utils.green('[SUCCESS]')} `
		+ `CHECK ALL HIGH TYPE DONE. Total Use Tick : ${utils.green(utils.TimeUsed.LastElapse())}`);
	return !foundError;
}

async function HandleExportAll(): Promise<boolean> {
	const monitor = new utils.AsyncWorkMonitor();
	let allOK = true;
	for (const kv of utils.ExportExcelDataMap) {
		for (const handler of gExportWrapperLst) {
			monitor.addWork();
			handler.ExportToAsync(kv[1], (ok) => {
				allOK = allOK && ok;
				monitor.decWork();
			});
		}
	}
	for (const handler of gExportWrapperLst) {
		monitor.addWork();
		handler.ExportToGlobalAsync((ok) => {
			allOK = allOK && ok;
			monitor.decWork();
		});
	}
	await monitor.WaitAllWorkDone();
	return allOK;
}

//#endregion
