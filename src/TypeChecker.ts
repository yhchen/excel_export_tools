/**
 * Support Format Check:
 * 		Using typescript type define link grammar.
 * 		Extends number type for digital size limit check.
 * 		Add common types of game development(like vector2 vector3...);
 *
 * Base Type:
 * 		-----------------------------------------------------------------------------
 * 		|		type		|						desc							|
 * 		-----------------------------------------------------------------------------
 * 		|		char		| min:-127					max:127						|
 * 		|		uchar		| min:0						max:255						|
 * 		|		short		| min:-32768				max:32767					|
 * 		|		ushort		| min:0						max:65535					|
 * 		|		int			| min:-2147483648			max:2147483647				|
 * 		|		uint		| min:0						max:4294967295				|
 * 		|		int64		| min:-9223372036854775808	max:9223372036854775807		|
 * 		|		uint64		| min:0						max:18446744073709551615	|
 * 		|		string		| auto change 'line break' to '\n'						|
 * 		|		double		| ...													|
 * 		|		float		| ...													|
 * 		|		date		| yyyy/mm/dd HH:MM:ss									|
 * 		|		tinydate	| yyyy/mm/dd											|
 * 		|		timestamp	| Linux time stamp										|
 * 		|		utctime		| UTC time stamp										|
 * 		-----------------------------------------------------------------------------
 *
 *
 * Combination Type:
 *
 * 		-----------------------------------------------------------------------------
 * 		| {<name>:<type>}	| start with '{' and end with '}' with json format.		|
 * 		|					| <type> is one of "Base Type" or "Combination Type".	|
 * 		-----------------------------------------------------------------------------
 * 		| <type>[<N>|null]	| <type> is one of "Base Type" or "Combination Type".	|
 * 		|					| <N> is empty(variable-length) or number.				|
 * 		-----------------------------------------------------------------------------
 * 		| vector2			| float[2]												|
 * 		-----------------------------------------------------------------------------
 * 		| vector3			| float[3]												|
 * 		-----------------------------------------------------------------------------
 * 		| json				| JSON.parse() is vaild									|
 * 		-----------------------------------------------------------------------------
 */
import * as xlsx from 'xlsx';
import { isArray, isObject, isNumber } from 'util';
import * as dateformat from 'dateformat';

function NullStr(s: string) {
	if (typeof s === "string") {
		return s.trim().length <= 0;
	}
	return true;
}

function VaildNumber(s: string): boolean {
	return (+s).toString() === s;
}

const WORDSCHAR = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_';
const NUMCHAR = '0123456789';
function FindWord(s: string, idx?: number): {start:number, end:number, len:number}|undefined {
	let first = true;
	let start = 0, end = s.length-1;
	for (let i = idx?idx:0; i < s.length; ++i) {
		if (first) {
			if (s[i] == ' ' || s[i] == '	') continue;
			if (WORDSCHAR.indexOf(s[i]) < 0) return undefined;
			first = false;
			start = i;
			continue;
		}
		if (WORDSCHAR.indexOf(s[i]) < 0) {
			end = i-1;
			break;
		}
	}
	return (end>=start)&&!first?{start,end,len:end-start+1}:undefined;
}

function FindNum(s: string, idx?: number): {start:number, end:number, len:number}|undefined {
	let first = true;
	let start = 0, end = s.length-1;
	for (let i = idx?idx:0; i < s.length; ++i) {
		if (first) {
			if (s[i] == ' ' || s[i] == '	') continue;
			if (NUMCHAR.indexOf(s[i]) < 0) return undefined;
			first = false;
			start = i;
			continue;
		}
		if (NUMCHAR.indexOf(s[i]) < 0) {
			end = i-1;
			break;
		}
	}
	return (end>=start)&&!first?{start,end,len:end-start+1}:undefined;
}

// all base type
const BaseTypeSet = new Set<string>([ 'char', 'uchar', 'short', 'ushort', 'int', 'uint', 'int64', 'uint64', 'string', 'double', 'float',
									  'vector2', 'vector3', 'json', 'date', 'tinydate', 'timestamp', 'utctime', ]);
// number type
const BaseNumberTypeSet = new Set<string>([ 'char', 'uchar', 'short', 'ushort', 'int', 'uint', 'int64', 'uint64', 'double', 'float', ]);
// date type
const BaseDateTypeSet = new Set<string>([ 'date', 'tinydate', 'timestamp', 'utctime', ]);
let DateFmt: string = 'yyyy/mm/dd HH:MM:ss';
console.log(`[TypeCheck] : Default Date format is "${DateFmt}"`)
let TinyDateFMT: string = 'yy/mm/dd';
console.log(`[TypeCheck] : Default Tiny Date format is "${TinyDateFMT}"`)
const TimeZoneOffset = new Date().getTimezoneOffset() * 60;
// number type range
const NumberRangeMap = new Map<string, {min:number, max:number}>([
		['char',	{ min:-127,			max:127 }],
		['uchar',	{ min:0,			max:255 }],
		['short',	{ min:-32768,		max:32767 }],
		['ushort',	{ min:0,			max:65535 }],
		['int',		{ min:-2147483648,	max:2147483647 }],
		['uint',	{ min:0,			max:4294967295 }],
	]);

function CheckNumberInRange(n: number, type: CType): boolean {
	if (type.typename == undefined) return true;
	const range = NumberRangeMap.get(type.typename);
	if (range == undefined) return true;
	return n >= range.min && n <= range.max;
}

// type name enum
export enum ETypeNames {
	char		=	'char',
	uchar		=	'uchar',
	short		=	'short',
	ushort		=	'ushort',
	int			=	'int',
	uint		=	'uint',
	int64		=	'int64',
	uint64		=	'uint64',
	string		=	'string',
	double		=	'double',
	float		=	'float',
	vector2		=	'vector2',
	vector3		=	'vector3',
	json		=	'json',
	date		=	'date',
	tinydate	=	'tinydate',
	timestamp	=	'timestamp',
	utctime		=	'utctime',
};

enum EType {
	object,
	array,
	base,
}

const RangeMap = {
	"char":		{ min:-127,						max:127 },
	"uchar":	{ min:0,						max:255 },
	"short":	{ min:-32768,					max:32767 },
	"ushort":	{ min:0,						max:65535 },
	"int":		{ min:-2147483648,				max:2147483647 },
	"uint":		{ min:0,						max:4294967295 },
	"int64":	{ min:-9223372036854775808,		max:9223372036854775807 },
	"uint64":	{ min:0,						max:18446744073709551615 },
};

export interface CType
{
	type: EType;
	is_number: boolean;
	typename?: string;
	num?: number;
	next?: CType;
	obj?: {[name:string]: CType};
}

export class CTypeChecker
{
	public constructor(typeString: string) {
		let s = typeString.replace(/ /g, '').replace(/\t/g, '').replace(/\n/g, '').replace(/\r/g, '');
		this.__s = s;
		if (NullStr(s)) {
			this._type = {type:EType.base, typename:ETypeNames.string, is_number:false};
			return;
		}
		let tt = this.initType({s, i:0});
		if (tt == undefined) throw `gen type check error: no data`;
		this._type = tt;
	}

	public get s(): string { return this.__s; }
	// get and set Date Format
	public static set DateFmt(s: string) { DateFmt = s; console.log(`[TypeCheck] : Change Date format to "${DateFmt}"`) }
	public static get DateFmt(): string { return DateFmt; }
	// get and set Tiny Date Format
	public static set TinyDateFmt(s: string) { TinyDateFMT = s; console.log(`[TypeCheck] : Change Tiny Date format to "${TinyDateFMT}"`); }
	public static get TinyDateFmt(): string { return TinyDateFMT; }

	public CheckValue(value: xlsx.CellObject|undefined): boolean {
		if (value == undefined || value.w == undefined || NullStr(value.w)) {
			return true;
		}
		if (this._type.is_number) {
			if (typeof value.v !== 'number') return false;
			return CheckNumberInRange(value.v, this._type);
		} else if (BaseDateTypeSet.has(this._type.typename||'')) {
			return value.t === 'd';
		} else {
			if (this._type.typename == ETypeNames.string) {
				return true;
			}
			let tmpObj:any = undefined;
			try {
				tmpObj = JSON.parse(value.w);
			} catch (ex) {
				console.log(false, `Check Object Type JSON Parse ERROR. ${ex}`);
				return false;
			}
			return this.CheckJsonType(tmpObj, this._type);
		}
		return true;
	}

	public ParseCellStr(value: xlsx.CellObject|undefined): string {
		if (this._type.is_number) {
			if (value == undefined) return '';
			if (typeof value.v === 'number') {
				if (value.v < 1) {
					return value.v.toString().replace(/0./g, '.');
				}
				return value.v.toString();
			}
			return value.w ? value.w : '';
		} else if (BaseDateTypeSet.has(this._type.typename||'')) {
			if (value == undefined) return '';
			if (value.t === 'd') {
				const date:Date = <Date>value.v;
				if (date == undefined) {
					return '';
				}
				if (this._type.typename === ETypeNames.utctime) {
					return (Math.round(date.getTime() / 1000 + TimeZoneOffset)).toString();
				} else if (this._type.typename === ETypeNames.timestamp) {
					return (Math.round(date.getTime() / 1000)).toString();
				} else if (this._type.typename === ETypeNames.date) {
					return dateformat.default(date, DateFmt);
				} else if (this._type.typename === ETypeNames.tinydate) {
					return dateformat.default(date, TinyDateFMT);
				}
			}
			return '';
		} else {
			if (value && value.w) {
				return value.w;
			}
			return '';
		}
	}

	private CheckJsonType(tmpObj:any, type: CType|undefined) : boolean {
		if (tmpObj == undefined || type == undefined)	return true;
		switch (type.type) {
		case EType.array:
			if (!isArray(tmpObj)) return false;
			if (type.num !== undefined && type.num != tmpObj.length) return false;
			for (let i = 0; i < tmpObj.length; ++i) {
				if (type.next == undefined || !this.CheckJsonType(tmpObj[i], type.next)) {
					return false;
				}
			}
			break;
		case EType.base:
			if (type.is_number) {
				if (typeof tmpObj === 'number') {
					return CheckNumberInRange(tmpObj, type);
				}
				else if (typeof tmpObj === 'string' && VaildNumber(tmpObj)) {
					return CheckNumberInRange(+tmpObj, type);
				}
				return false;
			}
			else if (type.typename == ETypeNames.json) {
				return true;
			}
			return typeof tmpObj === 'string';
			break;
		case EType.object:
			if (!isObject(tmpObj)) return false;
			if (!type.obj) return false;
			for (let key in tmpObj) {
				if (type.obj[key] == undefined) return false;
				if (!this.CheckJsonType(tmpObj[key], type.obj[key])) {
					return false;
				}
			}
			break;
		}
		return true;
	}

	private initType(p:{s:string, i:number}): CType|undefined {
		let thisNode:CType|undefined = undefined;
		// skip write space
		if (p.i >= p.s.length) undefined;
		let result:CType;
		if (p.s[p.i] == '{') {
			thisNode = {type:EType.object, is_number:false};
			++p.i;
			while (true)
			{
				// find name:
				const namescope = FindWord(p.s, p.i);
				if (!namescope) throw `gen type check error: object name not found!`;
				const name = p.s.substr(namescope.start, namescope.len);
				p.i = namescope.end + 1;
				if (p.s[p.i++] != ':') throw `gen type check error: object name key not join with ':'`;
				let tt = this.initType(p);
				if (!tt) throw `gen type check error: object name ${name} not found value!`;
				if (p.i >= p.s.length)	throw `gen type check error: '}' not found!`;
				if (!thisNode.obj) thisNode.obj = {};
				thisNode.obj[name] = tt;
				if (p.s[p.i] == ',') {
					++p.i;
					if (p.i >= p.s.length)	throw `gen type check error: '}' not found!`;
				}
				if (p.s[p.i] == '}') {
					++p.i;
					break;
				}
				if (p.i >= p.s.length)	throw `gen type check error: '}' not found!`;
			}
		} else if (p.s[p.i] == '[') {
			++p.i;
			let num:number|undefined = undefined;
			if (p.i >= p.s.length)	throw `gen type check error: '}' not found!`;
			if (p.s[p.i] != ']') {
				let numscope = FindNum(p.s, p.i);
				if (numscope == undefined) throw `gen type check error: array [<NUM>] format error!`;
				p.i = numscope.end+1;
				num = parseInt(p.s.substr(numscope.start, numscope.len));
			}
			if (p.i >= p.s.length)	throw `gen type check error: ']' not found!`;
			if (p.s[p.i] != ']') {
				throw `gen type check error: array [<NUM>] ']' not found!`;
			}
			++p.i;
			let arrNode:CType = {type:EType.array, num, is_number:false};
			if (p.i < p.s.length && p.s[p.i] == '[') {
				let nextArrNode = this.initType(p);
				if (!nextArrNode) throw `gen type check error: multi array error!`;
				arrNode.next = nextArrNode;
			}
			return arrNode;
		} else {
			const typescope = FindWord(p.s, p.i);
			if (!typescope) {
				throw `gen type check error: base type not found!`;
			}
			const typename = p.s.substr(typescope.start, typescope.len);
			if (!BaseTypeSet.has(typename)) throw `gen type check error: base type = ${this._type.typename} not exist!`;
			if (typename == ETypeNames.vector2 || typename == ETypeNames.vector3) {
				thisNode = {type:EType.base, typename:ETypeNames.float, is_number:true};
				const prevNode: CType = { type:EType.array, num:((typename==ETypeNames.vector2)?2:3), is_number:false };
				prevNode.next = thisNode;
				thisNode = prevNode;
			} else {
				thisNode = {type:EType.base, typename:typename, is_number:BaseNumberTypeSet.has(typename)};
			}
			p.i = typescope.end+1;
		}

		if (p.s[p.i] == '[') {
			let tt = this.initType(p);
			if (tt == undefined) throw `gen type check error: [] type error`;
			const typeNode = thisNode;
			thisNode = tt;
			while (tt.next != undefined) {
				tt = tt.next;
			}
			tt.next = typeNode;
		}
		return thisNode;
	}

	private _type:CType;
	private __s: string;
}

export function TestTypeChecker() {
	console.log(new CTypeChecker('int'));
	console.log(new CTypeChecker('string'));
	console.log(new CTypeChecker('int[]'));
	console.log(new CTypeChecker('int[2]'));
	console.log(new CTypeChecker('int[][2]'));
	console.log(new CTypeChecker('int[][]'));
	console.log(new CTypeChecker('{t:string}'));
	console.log(new CTypeChecker('{t:string, t1:string}'));
	console.log(new CTypeChecker('{t:string, t1:string}[]'));
	console.log(new CTypeChecker('{t:string, t1:{ut1:string}}[]'));
	console.log(new CTypeChecker('{t:string, t1:{ut1:string}[]}[]'));
}
