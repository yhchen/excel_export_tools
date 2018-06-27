import * as utils from "../utils";
import * as fs from "fs-extra-promise";
import * as path from 'path';

function ParseCSVLine(arry: Array<string>): string {
	for (let i = 0; i < arry.length; ++i) {
		let value = arry[i];
		if (value == null) {
			value = '';
		} else {
			if (value.indexOf(',') < 0 && value.indexOf('"') < 0) {
				value = value.replace(/"/g, `""`);
			} else {
				value = `"${value.replace(/"/g, `""`)}"`;
			}
		}
		arry[i] = value;
	}
	return arry.join(',').replace(/\n/g, '\\n').replace(/\r/g, '');
}


class CSVExport implements utils.IExportWrapper {
	public async exportTo(dt: utils.DataTable, outdir: string): Promise<boolean> {
		if (!fs.existsSync(outdir)) {
			utils.exception(`output path "${utils.yellow_ul(outdir)}" not exists!`);
			return false;
		}
		let tmpArr = new Array<string>();
		for (let row of dt.datas) {
			tmpArr.push(ParseCSVLine(row));
		}
		const csvcontent = tmpArr.join(utils.LineBreaker) + utils.LineBreaker;
		await fs.writeFileAsync(path.join(outdir, dt.name+'.csv'), csvcontent, {encoding:'utf8', flag:'w+'});

		utils.logger(false, `${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(path.join(outdir, dt.name+'.csv'))}". `
							+ `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);

		return true;
	}
}

module.exports = function():utils.IExportWrapper { return new CSVExport(); };
