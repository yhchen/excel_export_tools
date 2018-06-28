import * as utils from "../utils";
import * as fs from "fs-extra-promise";
import * as path from 'path';

function ParseCSVLine(sheetRow: utils.SheetRow, cfg: utils.ExportCfg): string {
	let tmpArry = new Array<string>();
	for (let i = 0; i < sheetRow.values.length; ++i) {
		let value = sheetRow.values[i];
		let tmpValue = '';
		if (value != null) {
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
			} else {
				throw `export INNER ERROR`;
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


class CSVExport implements utils.IExportWrapper {
	public async exportTo(dt: utils.SheetDataTable, outdir: string, cfg: utils.ExportCfg): Promise<boolean> {
		if (!fs.existsSync(outdir)) {
			utils.exception(`output path "${utils.yellow_ul(outdir)}" not exists!`);
			return false;
		}
		let tmpArr = new Array<string>();
		for (let row of dt.values) {
			tmpArr.push(ParseCSVLine(row, cfg));
		}
		const csvcontent = tmpArr.join(utils.LineBreaker) + utils.LineBreaker;
		await fs.writeFileAsync(path.join(outdir, dt.name+'.csv'), csvcontent, {encoding:'utf8', flag:'w+'});

		utils.logger(true, `${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(path.join(outdir, dt.name+'.csv'))}". `
							+ `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);

		return true;
	}
}

module.exports = function():utils.IExportWrapper { return new CSVExport(); };
