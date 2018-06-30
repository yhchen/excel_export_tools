import * as path from 'path';
import * as fs from "fs-extra-promise";
import * as utils from "../utils";

function ParseJsonLine(header: Array<utils.SheetHeader>, sheetRow: utils.SheetRow, rootNode: any, cfg: utils.GlobalCfg, exportCfg: utils.ExportCfg) {
	if (sheetRow.type != utils.ESheetRowType.data) return;
	let item: any = {};
	for (let i = 0; i < header.length && i < sheetRow.values.length; ++i) {
		if (!header[i] || header[i].comment) continue;
		if (sheetRow.values[i] != null) {
			item[header[i].name] = sheetRow.values[i];
		} else if (exportCfg.UseDefaultValueIfEmpty) {
			item[header[i].name] = header[i].typeChecker.DefaultValue;
		}
	}
	rootNode[sheetRow.values[0]] = item;
}


class JSONExport extends utils.IExportWrapper {
	constructor(exportCfg: utils.ExportCfg) { super(exportCfg); }

	public async ExportTo(dt: utils.SheetDataTable, cfg: utils.GlobalCfg): Promise<boolean> {
		const outdir = this._exportCfg.OutputDir;
		let jsonObj = {};
		for (let row of dt.values) {
			ParseJsonLine(dt.headerLst, row, jsonObj, cfg, this._exportCfg);
		}
		if (JSONExport.IsFile(outdir)) {
			this._globalObj[dt.name] = jsonObj;
		} else {
			if (!this.CreateDir(outdir)) {
				utils.exception(`create output path "${utils.yellow_ul(outdir)}" failure!`);
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

	public ExportEnd(cfg: utils.GlobalCfg): void {
		const outdir = this._exportCfg.OutputDir;
		if (!JSONExport.IsFile(outdir)) return;
		if (!this.CreateDir(path.basename(outdir))) {
			utils.exception(`create output path "${utils.yellow_ul(path.basename(outdir))}" failure!`);
			return;
		}
		const jsoncontent = JSON.stringify(this._globalObj);
		fs.writeFileSync(outdir, jsoncontent, {encoding:'utf8', flag:'w+'});
		utils.logger(true, `${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(outdir)}". `
						 + `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);
	}

	private static IsFile(s: string): boolean { return (path.extname(s) == '.json'); }

	private _globalObj: any = {};
}

module.exports = function(exportCfg: utils.ExportCfg):utils.IExportWrapper { return new JSONExport(exportCfg); };
