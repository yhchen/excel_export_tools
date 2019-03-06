import * as utils from "../utils";
import * as fs from "fs-extra-promise";
import * as path from 'path';
import { EType } from "../TypeChecker";
import * as json_to_lua from 'json_to_lua';

function ParseJsonObject(header: Array<utils.SheetHeader>, sheetRow: utils.SheetRow, rootNode: any, cfg: utils.GlobalCfg, exportCfg: utils.ExportCfg) {
	if (sheetRow.type != utils.ESheetRowType.data) return;
	let item: any = {};
	if (rootNode["ids"] == undefined) {
		rootNode["ids"] = new Array<any>();
	}
	for (let i = 0; i < header.length && i < sheetRow.values.length; ++i) {
		let hdr = header[i];
		if (!hdr || hdr.comment) continue;
		let val = sheetRow.values[i];
		if (val != null) {
			// FIXME : 特殊处理，json被用做Base类型(string)处理了，所以此处需要做一次特殊处理
			// 后续可能升级json类型的处理方法
			if (hdr.typeChecker.type.typename == "json") {
				item[hdr.name] = JSON.parse(val);
			} else {
				item[hdr.name] = val;
			}
		} else if (exportCfg.UseDefaultValueIfEmpty) {
			if (hdr.typeChecker.DefaultValue != undefined) {
				item[hdr.name] = hdr.typeChecker.DefaultValue;
			}
		}
		if (i == 0) {
			rootNode["ids"].push(item[header[0].name])
		}
	}
	rootNode[sheetRow.values[0]] = item;
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
			const jscontent = FMT.replace("{name}", dt.name).replace("{data}", json_to_lua.jsObjectToLuaPretty(jsonObj, 2));
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
		const jscontent = FMT.replace("{data}", json_to_lua.jsObjectToLuaPretty(this._globalObj, 3));
		fs.writeFileSync(outdir, jscontent, {encoding:'utf8', flag:'w+'});
		utils.logger(true, `${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(outdir)}". `
						 + `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);
	}

	private _globalObj: any = {};
}

module.exports = function(exportCfg: utils.ExportCfg):utils.IExportWrapper { return new LuaExport(exportCfg); };
