import * as fs from "fs-extra-promise";
import * as path from "path";
import * as utils from "../utils";
import { ETypeNames, CType, EType } from "../CTypeParser";

const TSTypeTranslateMap = new Map<ETypeNames, {s:string, opt:boolean}>([
	[ETypeNames.char,		{s:'number', opt:false}],
	[ETypeNames.uchar,		{s:'number', opt:false}],
	[ETypeNames.short,		{s:'number', opt:false}],
	[ETypeNames.ushort,		{s:'number', opt:false}],
	[ETypeNames.int,		{s:'number', opt:false}],
	[ETypeNames.uint,		{s:'number', opt:false}],
	[ETypeNames.int64,		{s:'number', opt:false}],
	[ETypeNames.uint64,		{s:'number', opt:false}],
	[ETypeNames.string,		{s:'string', opt:false}],
	[ETypeNames.double,		{s:'number', opt:false}],
	[ETypeNames.float,		{s:'number', opt:false}],
	[ETypeNames.bool,		{s:'boolean', opt:false}],
	[ETypeNames.json,		{s:'any', opt:true}],
	[ETypeNames.date,		{s:'string', opt:true}],
	[ETypeNames.tinydate,	{s:'string', opt:true}],
	[ETypeNames.timestamp,	{s:'number', opt:true}],
	[ETypeNames.utctime,	{s:'number', opt:true}],
]);

////////////////////////////////////////////////////////////////////////////////
class TSDExport extends utils.IExportWrapper {
	constructor(exportCfg: utils.ExportCfg) {
		super(exportCfg);
	}

	public get DefaultExtName(): string { return '.d.ts'; }

	protected async ExportTo(dt: utils.SheetDataTable): Promise<boolean> {
		let outdir = this._exportCfg.OutputDir;

		if (this.IsFile(outdir)) {
			return true;
		}

		if (!this.CreateDir(outdir)) {
			utils.exception(`create output path "${utils.yellow_ul(outdir)}" failure!`);
			return false;
		}
		let FMT: string | undefined = this._exportCfg.ExportTemple;
		if (FMT == undefined) {
			utils.exception(`[Config Error] ${utils.yellow_ul("Export.ExportTemple")} not defined!`);
			return false;
		}
		if (FMT.indexOf('{data}') < 0) {
			utils.exception(`[Config Error] ${utils.yellow_ul("Export.ExportTemple")} not found Keyword ${utils.yellow_ul("{data}")}!`);
			return false;
		}
		if (FMT.indexOf('{type}') < 0) {
			utils.exception(`[Config Error] ${utils.yellow_ul("Export.ExportTemple")} not found Keyword ${utils.yellow_ul("{type}")}!`);
			return false;
		}
		let ctx = this.GenSheetType(dt.name, dt.arrTypeHeader);
		if (!ctx)
			return true;
		let interfaceContent = FMT.replace('{data}', ctx.type).replace('{type}', ctx.tbtype);
		const outfile = outdir + dt.name + this._exportCfg.ExtName;
		await fs.writeFileAsync(outfile, interfaceContent, { encoding: 'utf8', flag: 'w' });
		utils.debug(`${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(outfile)}". `
			+ `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);
		return true;
	}

	protected async ExportGlobal(): Promise<boolean> {
		const outdir = this._exportCfg.OutputDir;
		if (!this.IsFile(outdir))
			return true;
		if (!this.CreateDir(path.dirname(outdir))) {
			utils.exception(`create output path "${utils.yellow_ul(path.dirname(outdir))}" failure!`);
			return false;
		}
		let FMT: string | undefined = this._exportCfg.ExportTemple;
		if (FMT == undefined) {
			utils.exception(`[Config Error] ${utils.yellow_ul("Export.ExportTemple")} not defined!`);
			return false;
		}
		if (FMT.indexOf('{data}') < 0) {
			utils.exception(`[Config Error] ${utils.yellow_ul("Export.ExportTemple")} not found Keyword ${utils.yellow_ul("{data}")}!`);
			return false;
		}
		if (FMT.indexOf('{type}') < 0) {
			utils.exception(`[Config Error] ${utils.yellow_ul("Export.ExportTemple")} not found Keyword ${utils.yellow_ul("{type}")}!`);
			return false;
		}

		let data = '';
		let type = '{\n';
		const exportexp = FMT.indexOf('export') >= 0 ? 'export ' : '';
		for (let iter of utils.ExportExcelDataMap) {
			const name = iter[1].name;
			let ctx = this.GenSheetType(name, iter[1].arrTypeHeader);
			if (ctx) {
				data += `${exportexp}${ctx.type}${exportexp}${ctx.tbtype}\n\n`;
				type += `\t${name}:T${name};\n`;
			}
		}
		type += `}\n`;
		FMT.indexOf('{data}');
		data = FMT.replace('{data}', data).replace('{type}', type);
		await fs.writeFileAsync(outdir, data, { encoding: 'utf8', flag: 'w' });
		utils.debug(`${utils.green('[SUCCESS]')} Output file "${utils.yellow_ul(outdir)}". `
			+ `Total use tick:${utils.green(utils.TimeUsed.LastElapse())}`);
		return true;
	}

	private GenSheetType(sheetName: string, arrHeader: utils.SheetHeader[]): { type: string, tbtype: string } | undefined {
		const arrExportHeader = utils.ExecGroupFilter(this._exportCfg.GroupFilter, arrHeader)
		if (arrExportHeader.length <= 0) {
			utils.debug(`Pass Sheet ${utils.yellow_ul(sheetName)} : No Column To Export.`);
			return;
		}

		let type = `type ${sheetName} = {\n`;
		for (let header of arrExportHeader) {
			if (header.comment) continue;
			type += `\t${header.name}${this.GenTypeName(header.typeChecker.type, false)};\n`;
		}
		type += '}\n';
		let tbtype = `type T${sheetName} = {[Key in number|string]?: ${sheetName}};\n`
		return { type, tbtype };
	}

	private GenTypeName(type: CType | undefined, opt: boolean = false): string {
		const defaultval = `?: any`;
		if (type == undefined) {
			return defaultval;
		}
		switch (type.type) {
			case EType.base:
			case EType.date:
				if (type.typename) {
					let tdesc = TSTypeTranslateMap.get(type.typename);
					if (tdesc) {
						return `${opt || tdesc.opt || !this._exportCfg.UseDefaultValueIfEmpty ? '?' : ''}: ${tdesc.s}`;
					}
				} else {
					return defaultval;
				}
				break;
			case EType.array:
				{
					let tname = `[]`;
					for (; type != undefined; type = type.next) {
						if (type.type == EType.array) {
							tname += `[]`;
						} else {
							tname = this.GenTypeName(type, true) + tname;
						}
					}
					return tname;
				}
				break;
			case EType.object:
				if (type.obj)
				{
					let tname = `?: { `;
					for (let tt in type.obj) {
						tname += tt + this.GenTypeName(type.obj[tt], true) + `; `;
					}
					tname += `}`;
				}
				break;
			default:
				utils.exception(`call "${utils.yellow_ul('GenTypeName')}" failure`);
				return defaultval;
		}
		return defaultval;
	}

}

module.exports = function (exportCfg: utils.ExportCfg): utils.IExportWrapper { return new TSDExport(exportCfg); };
