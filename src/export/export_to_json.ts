import * as path from 'path';
import * as fs from "fs-extra-promise";
import * as utils from "../utils";
import * as TC from "../TypeChecker";

function ParseJsonLine(header: Array<utils.SheetHeader>, sheetRow: utils.SheetRow, rootNode: any, cfg: utils.ExportCfg) {
	if (sheetRow.type != utils.ESheetRowType.data) return;
	let item: any = {};
	for (let i = 0; i < header.length && i < sheetRow.values.length; ++i) {
		if (!header[i] || header[i].comment) continue;
		if (sheetRow.values[i] != null) {
			item[header[i].name] = sheetRow.values[i];
		} else {
			item[header[i].name] = header[i].typeChecker.DefaultValue;
		}
	}
	rootNode[sheetRow.values[0]] = item;
}


class JSONExport implements utils.IExportWrapper {
	public async ExportTo(dt: utils.SheetDataTable, outdir: string, cfg: utils.ExportCfg): Promise<boolean> {
		let jsonObj = {};
		for (let row of dt.values) {
			ParseJsonLine(dt.headerLst, row, jsonObj, cfg);
		}
		if (JSONExport.IsFile(outdir)) {
			this._globalObj[dt.name] = jsonObj;
		} else {
			if (!fs.existsSync(outdir)) {
				utils.exception(`output path "${utils.yellow_ul(outdir)}" not exists!`);
				return false;
			}
			const jsoncontent = JSON.stringify(jsonObj);
			const outfile = path.join(outdir, dt.name+'.json');
			await fs.writeFileAsync(outfile, jsoncontent, {encoding:'utf8', flag:'w+'});
			utils.logger(true, `${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(outfile)}". `
							 + `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);
		}
		return true;
	}

	public ExportEnd(outdir: string, cfg: utils.ExportCfg): void {
		if (!JSONExport.IsFile(outdir)) return;
		if (!JSONExport.CreateFileDir(outdir)) {
			utils.exception(`create output path "${utils.yellow_ul(path.basename(outdir))}" failure!`);
			return;
		}
		const jsoncontent = JSON.stringify(this._globalObj);
		fs.writeFileSync(outdir, jsoncontent, {encoding:'utf8', flag:'w+'});
		utils.logger(true, `${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(outdir)}". `
						 + `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);
	}

	private static IsFile(s: string): boolean { return (path.extname(s) == '.json'); }
	private static CreateFileDir(outfile: string): boolean {
		const basepath = path.dirname(outfile);
		if (!fs.existsSync(basepath)) {
			fs.ensureDirSync(basepath);
			return fs.existsSync(basepath);
		}
		return true;
	}

	private _globalObj: any = {};
}

module.exports = function():utils.IExportWrapper { return new JSONExport(); };
