import { types } from "util";

/**
 * Support Format Check:
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
 * 		|		vector2		| [float, float]										|
 * 		|		vector3		| [float, float, float]									|
 * 		|		double		| ...													|
 * 		|		float		| ...													|
 * 		-----------------------------------------------------------------------------
 *
 *
 * Combination Type:
 *
 * 		-----------------------------------------------------------------------------
 * 		| {"<name>":<type>}	| start with '{' and end with '}' with json format.		|
 * 		|					| <type> is one of "Base Type" or "Combination Type".	|
 * 		-----------------------------------------------------------------------------
 * 		|	<type>[<N>]		| <type> is one of "Base Type" or "Combination Type".	|
 * 		|					| <N> is empty(variable-length) or number.				|
 * 		-----------------------------------------------------------------------------
 */

function NullStr(s: string) {
	if (typeof s === "string") {
		return s.trim().length <= 0;
	}
	return true;
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


const BaseTypeSet = new Set<string>([ 'char','uchar','short','ushort','int','uint','int64','uint64','string','double','float','vector2','vector3' ]);

enum ETypeRelation {
	char	=	"char",
	uchar	=	"uchar",
	short	=	"short",
	ushort	=	"ushort",
	int		=	"int",
	uint	=	"uint",
	int64	=	"int64",
	uint64	=	"uint64",
	string	=	"string",
	double	=	"double",
	float	=	"float",
	vector2	=	"vector2",
	vector3	=	"vector3",
};

enum EType {
	undefined,
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
	typename?: string;
	num?: number;
	next?: CType;
	obj?: {[name:string]: CType};
}

export class CTypeChecker
{
	public constructor(typeString: string) {
		this.__s = typeString.replace(/ /g, '').replace(/\t/g, '').replace(/\n/g, '').replace(/\r/g, '');
		if (NullStr(this.__s)) {
			this._type = {type:EType.base, typename:ETypeRelation.string};
			return;
		}
		let tt = this.initType();
		if (tt == undefined) throw `gen type check error: no data`;
		this._type = tt;
	}

	public CheckValue(value: string): boolean {
		return true;
	}
	public GetValue(value: string|undefined): string {
		return value ? value : '';
	}

	private initType(): CType|undefined {
		let thisNode:CType|undefined = undefined;
		// skip write space
		if (this.__idx >= this.__s.length) undefined;
		let result:CType;
		if (this.__s[this.__idx] == '{') {
			thisNode = {type:EType.object};
			++this.__idx;
			while (true)
			{
				// find name:
				const namescope = FindWord(this.__s, this.__idx);
				if (!namescope) throw `gen type check error: object name not found!`;
				const name = this.__s.substr(namescope.start, namescope.len);
				this.__idx = namescope.end + 1;
				if (this.__s[this.__idx++] != ':') throw `gen type check error: object name key not join with ':'`;
				let tt = this.initType();
				if (!tt) throw `gen type check error: object name ${name} not found value!`;
				if (this.__idx >= this.__s.length)	throw `gen type check error: '}' not found!`;
				if (!thisNode.obj) thisNode.obj = {};
				thisNode.obj[name] = tt;
				if (this.__s[this.__idx] == ',') {
					++this.__idx;
					if (this.__idx >= this.__s.length)	throw `gen type check error: '}' not found!`;
				}
				if (this.__s[this.__idx] == '}') {
					++this.__idx;
					break;
				}
				if (this.__idx >= this.__s.length)	throw `gen type check error: '}' not found!`;
			}
		} else if (this.__s[this.__idx] == '[') {
			++this.__idx;
			let num:number|undefined = undefined;
			if (this.__idx >= this.__s.length)	throw `gen type check error: '}' not found!`;
			if (this.__s[this.__idx] != ']') {
				let numscope = FindNum(this.__s, this.__idx);
				if (numscope == undefined) throw `gen type check error: array [<NUM>] format error!`;
				this.__idx = numscope.end+1;
				num = parseInt(this.__s.substr(numscope.start, numscope.len));
			}
			if (this.__idx >= this.__s.length)	throw `gen type check error: ']' not found!`;
			if (this.__s[this.__idx] != ']') {
				throw `gen type check error: array [<NUM>] ']' not found!`;
			}
			++this.__idx;
			let arrNode:CType = {type:EType.array, num};
			if (this.__idx < this.__s.length && this.__s[this.__idx] == '[') {
				let nextArrNode = this.initType();
				if (!nextArrNode) throw `gen type check error: multi array error!`;
				arrNode.next = nextArrNode;
			}
			return arrNode;
		} else {
			const typescope = FindWord(this.__s, this.__idx);
			if (!typescope) {
				throw `gen type check error: base type not found!`;
			}
			const typename = this.__s.substr(typescope.start, typescope.len);
			if (!BaseTypeSet.has(typename)) throw `gen type check error: base type = ${this._type.typename} not exist!`;
			thisNode = {type:EType.base, typename:typename};
			this.__idx = typescope.end+1;
		}

		if (this.__s[this.__idx] == '[') {
			let tt = this.initType();
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
	private __idx = 0;
	private __s: string;
}

// console.log(new CTypeChecker('int'));
// console.log(new CTypeChecker('string'));
// console.log(new CTypeChecker('int[]'));
// console.log(new CTypeChecker('int[2]'));
// console.log(new CTypeChecker('int[][2]'));
// console.log(new CTypeChecker('int[][]'));
// console.log(new CTypeChecker('{t:string}'));
// console.log(new CTypeChecker('{t:string, t1:string}'));
// console.log(new CTypeChecker('{t:string, t1:string}[]'));
// console.log(new CTypeChecker('{t:string, t1:{ut1:string}}[]'));
// console.log(new CTypeChecker('{t:string, t1:{ut1:string}[]}[]'));
