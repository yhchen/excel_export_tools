import * as utils from "../utils";
import * as fs from "fs-extra-promise";
import * as path from 'path';
const json2lua = require('json2lua')

function ParseJsonObject(header: Array<utils.SheetHeader>, sheetRow: utils.SheetRow, rootNode: any, cfg: utils.GlobalCfg, exportCfg: utils.ExportCfg) {
	if (sheetRow.type != utils.ESheetRowType.data) return;
	let item: any = {};
	let ids = new Array<any>();
	for (let i = 0; i < header.length && i < sheetRow.values.length; ++i) {
		if (!header[i] || header[i].comment) continue;
		if (sheetRow.values[i] != null) {
			item[header[i].name] = sheetRow.values[i];
		} else if (exportCfg.UseDefaultValueIfEmpty) {
			if (header[i].typeChecker.DefaultValue != undefined) {
				item[header[i].name] = header[i].typeChecker.DefaultValue;
			}
		}
		if (i == 0) {
			ids.push(item[header[0].name])
		}
	}
	rootNode[sheetRow.values[0]] = item;
	rootNode["ids"] = ids;
}


class LuaExport extends utils.IExportWrapper {
	constructor(exportCfg: utils.ExportCfg) { super(exportCfg); }

	public get DefaultExtName(): string { return '.lua'; }
	public async ExportTo(dt: utils.SheetDataTable, cfg: utils.GlobalCfg): Promise<boolean> {
		const outdir = this._exportCfg.OutputDir;
		let jsonObj = {};
		for (let row of dt.values) {
			ParseJsonObject(dt.headerLst, row, jsonObj, cfg, this._exportCfg);
		}
		if (this.IsFile(outdir)) {
			this._globalObj[dt.name] = jsonObj;
		} else {
			if (!this.CreateDir(outdir)) {
				utils.exception(`create output path "${utils.yellow_ul(outdir)}" failure!`);
				return false;
			}

			let FMT: string|undefined = this._exportCfg.ExportTemple;
			if (FMT == undefined) {
				utils.exception(`[Config Error] ${utils.yellow_ul("Export.ExportTemple")} not defined!`);
				return false;
			}
			if (FMT.indexOf('{data}') < 0) {
				utils.exception(`[Config Error] ${utils.yellow_ul("Export.ExportTemple")} not found Keyword ${utils.yellow_ul("{data}")}!`);
				return false;
			}
			if (FMT.indexOf('{name}') < 0) {
				utils.exception(`[Config Error] ${utils.yellow_ul("Export.ExportTemple")} not found Keyword ${utils.yellow_ul("{name}")}!`);
				return false;
			}
			const jscontent = FMT.replace("{name}", dt.name).replace("{data}", json2lua.fromObject(jsonObj));
			const outfile = path.join(outdir, dt.name+this._exportCfg.ExtName);
			await fs.writeFileAsync(outfile, jscontent, {encoding:'utf8', flag:'w+'});
			utils.logger(true, `${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(outfile)}". `
							 + `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);		}
		return true;
	}

	public ExportEnd(cfg: utils.GlobalCfg): void {
		const outdir = this._exportCfg.OutputDir;
		if (!this.IsFile(outdir)) return;
		if (!this.CreateDir(path.dirname(outdir))) {
			utils.exception(`create output path "${utils.yellow_ul(path.dirname(outdir))}" failure!`);
			return;
		}
		let FMT: string|undefined = this._exportCfg.ExportTemple;
		if (FMT == undefined) {
			utils.exception(`[Config Error] ${utils.yellow_ul("Export.ExportTemple")} not defined!`);
			return;
		}
		if (FMT.indexOf('{data}') < 0) {
			utils.exception(`[Config Error] ${utils.yellow_ul("Export.ExportTemple")} not found Keyword ${utils.yellow_ul("{data}")}!`);
			return;
		}
		const jscontent = FMT.replace("{data}", json2lua.fromObject(this._globalObj));
		fs.writeFileSync(outdir, jscontent, {encoding:'utf8', flag:'w+'});
		utils.logger(true, `${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(outdir)}". `
						 + `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);
	}

	private _globalObj: any = {};
}

module.exports = function(exportCfg: utils.ExportCfg):utils.IExportWrapper { return new LuaExport(exportCfg); };
