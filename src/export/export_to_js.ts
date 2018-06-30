import * as path from 'path';
import * as fs from "fs-extra-promise";
import * as utils from "../utils";

function ParseJSLine(header: Array<utils.SheetHeader>, sheetRow: utils.SheetRow,
					 rootNode: any, cfg: utils.GlobalCfg, exportCfg: utils.ExportCfg): string|undefined {
	if (sheetRow.type != utils.ESheetRowType.data) return undefined;
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

function DumpToString(data: any) {
	if (utils.isString(data)) {
		return `'${data}'`;
	} else if (utils.isNumber(data)) {
		return data.toString();
	} else if (utils.isArray(data)) {
		let s = '';
		for (let i = 0; i < data.length; ++i) {
			s += (i === 0?'':',') + DumpToString(data[i]);
		}
		return `[${s}]`;
	} else if (utils.isObject(data)) {
		let s = '';
		let first = true;
		for (let name in data) {
			s += `${first?'':','}${name}:${DumpToString(data[name])}`;
			first = false;
		}
		return `{${s}}`;
	} else if (utils.isBoolean(data)) {
		return data ? 'true' : 'false';
	} else {
		utils.exception(`Internal ERROR! type not Handle${data}!`);
	}
	return '';
}

class JSExport extends utils.IExportWrapper {
	constructor(exportCfg: utils.ExportCfg) { super(exportCfg); }

	public async ExportTo(dt: utils.SheetDataTable, cfg: utils.GlobalCfg): Promise<boolean> {
		const outdir = this._exportCfg.OutputDir;
		let jsObj = {};
		for (let row of dt.values) {
			ParseJSLine(dt.headerLst, row, jsObj, cfg, this._exportCfg);
		}
		if (JSExport.IsFile(outdir)) {
			this._globalObj[dt.name] = jsObj;
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
			const jscontent = FMT.replace("{name}", dt.name).replace("{data}", DumpToString(jsObj));
			const outfile = path.join(outdir, dt.name+'.js');
			await fs.writeFileAsync(outfile, jscontent, {encoding:'utf8', flag:'w+'});
			utils.logger(true, `${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(outfile)}". `
							 + `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);
		}
		return true;
	}

	public ExportEnd(cfg: utils.GlobalCfg): void {
		const outdir = this._exportCfg.OutputDir;
		if (!JSExport.IsFile(outdir)) return;
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
		const jscontent = FMT.replace("{data}", DumpToString(this._globalObj));
		fs.writeFileSync(outdir, jscontent, {encoding:'utf8', flag:'w+'});
		utils.logger(true, `${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(outdir)}". `
						 + `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);
	}

	private static IsFile(s: string): boolean { return (path.extname(s) == '.js'); }

	private _globalObj: any = {};
}

module.exports = function(exportCfg: utils.ExportCfg):utils.IExportWrapper { return new JSExport(exportCfg); };
