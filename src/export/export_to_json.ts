import * as utils from "../utils";
import * as fs from "fs-extra-promise";
import * as path from 'path';
import * as TC from "../TypeChecker";

function ParseJsonLine(header: Array<utils.SheetHeader>, sheetRow: utils.SheetRow, rootNode: any, cfg: utils.ExportCfg) {
	if (sheetRow.type != utils.ESheetRowType.data) return;
	let item: any = {};
	for (let i = 0; i < header.length && i < sheetRow.values.length; ++i) {
		for (let col of header) {
			if (col.comment) continue;
		}
		let value = sheetRow.values[i];
		switch (header[i].typeChecker.type) {
			case TC.EType.array:	item[header[i].name] = value || [];		break;
			case TC.EType.object:	item[header[i].name] = value || {};		break;
			default:				item[header[i].name] = value;			break;
		}
	}
	rootNode[sheetRow.values[0]] = item;
}


class JSONExport implements utils.IExportWrapper {
	public async exportTo(dt: utils.SheetDataTable, outdir: string, cfg: utils.ExportCfg): Promise<boolean> {
		if (!fs.existsSync(outdir)) {
			utils.exception(`output path "${utils.yellow_ul(outdir)}" not exists!`);
			return false;
		}
		let jsonObj = {};
		for (let row of dt.values) {
			ParseJsonLine(dt.headerLst, row, jsonObj, cfg);
		}
		const jsoncontent = JSON.stringify(jsonObj);
		const outfile = path.join(outdir, dt.name+'.json');
		await fs.writeFileAsync(outfile, jsoncontent, {encoding:'utf8', flag:'w+'});

		utils.logger(true, `${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(outfile)}". `
							+ `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);

		return true;
	}
}

module.exports = function():utils.IExportWrapper { return new JSONExport(); };
