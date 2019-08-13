import * as fs from 'fs-extra-promise';
import * as utils from './utils'
import { CTypeParser } from './CTypeParser';

import ConfTpl from "./config_tpl.json";

// Work Root Dir
export const gRootDir = process.cwd();
export const gGlobalIgnoreDirName = new Set([
	".svn",
	".git"
]);

// Global Export Config
export let gCfg: typeof ConfTpl = ConfTpl; // default config

export function InitGlobalConfig(fpath: string = ''): boolean {
	if (fpath != '') {
		gCfg = JSON.parse(<string>fs.readFileSync(fpath, { encoding: 'utf8' }));
		function check(gCfg: any, ConfTpl: any): boolean {
			for (let key in ConfTpl) {
				if (ConfTpl[key] != null && typeof gCfg[key] !== typeof ConfTpl[key]) {
					utils.exception(utils.red(`configure format error. key "${utils.yellow(key)}" not found!`));
					return false;
				}
				if (utils.isObject(typeof ConfTpl[key])) {
					check(gCfg[key], ConfTpl[key]);
				}
			}
			return true;
		};
		if (!check(gCfg, ConfTpl)) {
			return false;
		}
	}
	utils.SetLineBreaker(gCfg.LineBreak);

	CTypeParser.DateFmt = gCfg.DateFmt;
	CTypeParser.TinyDateFmt = gCfg.TinyDateFmt;
	CTypeParser.FractionDigitsFMT = gCfg.FractionDigitsFMT;

	return true;
}
