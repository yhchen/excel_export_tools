#!/usr/bin/env node

import * as fs from 'fs';
import { execute } from './works'
import * as utils from './utils'
import * as config from './config'


////////////////////////////////////////////////////////////////////////////////
function printHelp() {
	console.log(false, `${process.argv[0]} ${process.argv[1]} <config file path(optional)>`);
}

async function main() {
	try {
		const configPath = process.argv.length >= 3 ? process.argv[2] : undefined;
		if (configPath == '-h' || configPath == '--h' || configPath == '/h' || configPath == '/?') {
			printHelp();
			return;
		} else if (!fs.existsSync(process.argv[2])) {
			printHelp();
			return;
		}
		if (!config.InitGlobalConfig(configPath)) {
			utils.exception(`Init Global Config "${configPath}" Failure.`);
			return;
		}
		await execute();
		console.log('--------------------------------------------------------------------');
	} catch (ex) {
		utils.exception(ex);
		process.exit(utils.E_ERROR_LEVEL.EXECUTE_FAILURE);
	}
}

// main entry
main();
