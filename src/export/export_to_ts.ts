import * as path from 'path';
import * as fs from "fs-extra-promise";
import * as utils from "../utils";

class TSExport extends utils.IExportWrapper {
	constructor(exportCfg: utils.ExportCfg) {
		super(exportCfg);
		this._jsExportor = (<utils.ExportWrapperFactory>utils.ExportWrapperMap.get('js'))(exportCfg);
	}

	public async ExportTo(dt: utils.SheetDataTable, cfg: utils.GlobalCfg): Promise<boolean> {
		const outdir = this._exportCfg.OutputDir;
		if (!this._jsExportor.ExportTo(dt, cfg)) {
			return false;
		}
		if (this.IsFile(outdir)) {
			return true;
		}
		let FMT: string = <string>this._exportCfg.ExportTemple;
		// const jscontent = FMT.replace("{name}", dt.name).replace("{data}", DumpToString(jsObj));
		// const outfile = path.join(outdir, dt.name+'.d.ts');
		// await fs.writeFileAsync(outfile, jscontent, {encoding:'utf8', flag:'w+'});
		// utils.logger(true, `${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(outfile)}". `
		// 				 + `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);
		return true;
	}

	public ExportEnd(cfg: utils.GlobalCfg): void {
		const outdir = this._exportCfg.OutputDir;
		this._jsExportor.ExportEnd(cfg);
		if (!this.IsFile(outdir)) return;
		let FMT: string = <any>this._exportCfg.ExportTemple;

		// const jscontent = FMT.replace("{data}", DumpToString(this._globalObj));
		// fs.writeFileSync(outdir, jscontent, {encoding:'utf8', flag:'w+'});
		// utils.logger(true, `${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(outdir)}". `
		// 				 + `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);
	}

	private _jsExportor:utils.IExportWrapper;
}

module.exports = function(exportCfg: utils.ExportCfg):utils.IExportWrapper { return new TSExport(exportCfg); };
