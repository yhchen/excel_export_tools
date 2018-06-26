#!/bin/bash

echo ======================================================
echo = check nodejs
type node &> /dev/null;
if [ ! $? -eq 0 ]; then
	echo [ERROR] node.js not found!;
	echo [ERROR] please install node.js first!;
	exit -1;
fi
echo = node version is :`node -v`

if [ ! -d "./node_modules" ]; then
	echo = package.json not init! run npm install...
	ECHO_SPEED_UP_HELP

	npm install
fi

echo ======================================================
echo = check command typescript
type tsc &> /dev/null
if [ ! $? -eq 0 ]; then
	echo "= typescript not found! run install last version of typescript..."
	ECHO_SPEED_UP_HELP

	npm install -g typescript
fi
echo = typescript version is :`tsc -v`


if [ ! "$1" == "0" ]; then
	echo ======================================================
	echo = tsc compile
	tsc -p ./tsconfig.json

	echo ======================================================
	echo = execute build config files
	node ./bin/index.js
fi

ECHO_SPEED_UP_HELP() {
	echo "="
	echo "= If you want to speed up download time in china area."
	echo "= use this command below:"
	echo "= npm config set registry https://registry.npm.taobao.org"
}
