#!/usr/bin/env node

import * as lodash from 'lodash';

let MaxPreetyDepth = 0;

export function toLuaString(obj: any, maxPreetyDepth: number = 0): string {
	MaxPreetyDepth = maxPreetyDepth
	return toLua(obj, 0);
}

const NumberCharSet = new Set<string>(['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']);
const FirstInvalidCharSet = new Set<string>(['a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z'
											,'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'
											,'_']);



function toLua(obj: any, currDepth: number, CurrEntry?: string): string {
	const preety = MaxPreetyDepth > currDepth;
	const NextDepth = currDepth + 1;
	CurrEntry = (CurrEntry != undefined && preety) ? CurrEntry + '\t' : '';
	const ObjectEntry = (CurrEntry != undefined && preety) ? CurrEntry + '\t' : '';
	const EndLine = preety ? '\n' : '';
	const WriteSpace = preety ? ' ' : '';
    if (obj === null || obj === undefined) {
        return 'nil';
    }
    if (!lodash.isObject(obj)) {
        if (typeof obj === 'string') {
            return '"' + obj + '"';
        }
        return obj.toString();
	}
    let result = `{` + EndLine,
        isArray = obj instanceof Array,
        len = lodash.size(obj),
		i = 0;
	lodash.forEach(obj, function (v, k) {
		let objStr = '';
        if (isArray) {
            objStr = toLua(v, NextDepth, CurrEntry);
        } else {
			if (parseInt(k).toString() == k || (k.length > 0 && !FirstInvalidCharSet.has(k[0]))) {
				objStr = '[' + k + `]${WriteSpace}=${WriteSpace}` + toLua(v, NextDepth, CurrEntry);
			} else if (NumberCharSet.has(k[0])) {
				objStr = '["' + k + `"]${WriteSpace}=${WriteSpace}` + toLua(v, NextDepth, CurrEntry);
			} else {
				objStr = k + `${WriteSpace}=${WriteSpace}` + toLua(v, NextDepth, CurrEntry);
			}
        }
        if (i < len - 1) {
            objStr += ',';
        }
		i += 1;
		result += ObjectEntry + objStr + EndLine;
    });
    result += CurrEntry + '}';
    return result;
}
