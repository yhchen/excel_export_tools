import * as utils from "../utils";
import * as fs from "fs-extra-promise";
import * as path from 'path';

function ParseCSVLine(header: Array<utils.SheetHeader>, sheetRow: utils.SheetRow, exportCfg: utils.ExportCfg): string {
	let tmpArry = new Array<string>();
	for (let i = 0; i < sheetRow.values.length; ++i) {
		let value = sheetRow.values[i];
		let tmpValue = '';
		if (value == null) {
			if (exportCfg.UseDefaultValueIfEmpty) {
				tmpValue = header[i].typeChecker.SDefaultValue;
				if (tmpValue === undefined) {
					tmpValue = '';
				}
			}
		} else {
			if (utils.isString(value)) {
				tmpValue = value;
			} else if (utils.isObject(value) || utils.isArray(value)) {
				tmpValue = JSON.stringify(value);
			} else if (utils.isNumber(value)) {
				if (value < 1 && value > 0) {
					tmpValue = value.toString().replace(/0\./g, '.');
				} else {
					tmpValue = value.toString();
				}
			} else if (utils.isBoolean(value)) {
				tmpValue = value ? 'true' : 'false';
			} else {
				utils.exception(`export INNER ERROR`);
			}
			if (tmpValue.indexOf(',') < 0 && tmpValue.indexOf('"') < 0) {
				tmpValue = tmpValue.replace(/"/g, `""`);
			} else {
				tmpValue = `"${tmpValue.replace(/"/g, `""`)}"`;
			}
		}
		tmpArry.push(tmpValue);
	}
	return tmpArry.join(',').replace(/\n/g, '\\n').replace(/\r/g, '');
}


class CSVExport extends utils.IExportWrapper {
	constructor(exportCfg: utils.ExportCfg) { super(exportCfg); }

	public get DefaultExtName(): string { return '.csv'; }
	protected async ExportTo(dt: utils.SheetDataTable): Promise<boolean> {
		const outdir = this._exportCfg.OutputDir;
		if (!this.CreateDir(outdir)) {
			utils.exception(`output path "${utils.yellow_ul(outdir)}" not exists!`);
			return false;
		}
		let arrTmp = new Array<string>();
		const arrExportHeader = utils.ExecGroupFilter(this._exportCfg.GroupFilter, dt.arrTypeHeader)
		if (arrExportHeader.length <= 0) {
			utils.debug(`Pass Sheet ${utils.yellow_ul(dt.name)} : No Column To Export.`);
			return true;
		}
		for (let row of dt.arrValues) {
			if (row.type != utils.ESheetRowType.data && row.type != utils.ESheetRowType.header) continue;
			arrTmp.push(ParseCSVLine(arrExportHeader, row, this._exportCfg));
		}
		const csvcontent = arrTmp.join(utils.LineBreaker) + utils.LineBreaker;
		await fs.writeFileAsync(path.join(outdir, dt.name + this._exportCfg.ExtName), csvcontent, { encoding: 'utf8', flag: 'w+' });

		utils.debug(`${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(path.join(outdir, dt.name + this._exportCfg.ExtName))}". `
			+ `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);

		return true;
	}
	protected async ExportGlobal(): Promise<boolean> { return true; }
}

module.exports = function (exportCfg: utils.ExportCfg): utils.IExportWrapper { return new CSVExport(exportCfg); };
