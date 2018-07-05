excel_export_tools
==================


excel export with type check using nodejs in super fast speed!
---

How To Use
---

* Configure file is `config.json`. Add files and directory to `[IncludeFilesAndPath]` to dealing with. Add files to `[ExcludeFileNames]` to disable processing. Add table name to `[ExcludeCsvTableNames]` to disable processing.
* Set output `table name` at cell `B1`(Default value). It can be change at config Key `[CSVNameCellID]`.
* Add output table `Column Name` Row under `table name` Row. If a Column Name Start with `#` means is was a comment Column. It can be disable output at config `EnableExportCommentColumns`.
* Add `Column Format Type` line under `Column Name` Row.
* `[Export]`  is temporary not supported.
* A comment Row was start with word `#` at `A Column`. Is can be disable output at config `EnableExportCommentRows`.
* Cell `[N]A([N] is Row index)` start with `#` is a comment line.


Support Format Export
---

type|desc
---|---
csv|csv format
json|Export as `json` format. If `Export.OutputDir` is a directory. Export as a separate file per sheet, otherwise export as one file.
js|Export as `js` format. Set `Export.ExportTemple` for export template. The tools will replace `{name}` with sheet name, and replace `{data}` with sheet data. If `Export.OutputDir` is a directory. Export as a separate file per sheet, otherwise export as one file.
ts|Support in future releases
lua|Support in future releases


Support Format Check
---


* Declare type definitions using grammar rules like typescript interface.
* Support numeric type size overflow validation.
* More convenient type definitions for game developers(like vector2 vector3 etc...).

### Base Type:

type|desc
---|---
char|min:-127                  max:127
uchar|min:0                     max:255
short|min:-32768                max:32767
ushort|min:0                     max:65535
int|min:-2147483648           max:2147483647
uint|min:0                     max:4294967295
int64|min:-9223372036854775808  max:9223372036854775807
uint64|min:0                     max:18446744073709551615
string|auto change 'line break' to '\n
double|...
float|...
bool|true: 'true' or '1'       false: 'false' empty or '0
date|YYYY/MM/DD HH:mm:ss
tinydate|YYYY/MM/DD
timestamp|Linux time stamp
utctime|UTC time stamp

### Combination Type:

type|desc
---|---
`{<name>:<type>}`   | start with '{' and end with '}' with json format. `<type>` is one of "Base Type" or "Combination Type".
`<type>[<N>|null]`  | `<type>` is one of "Base Type" or "Combination Type". `<N>` is empty(variable-length) or number.
`vector2`           | float\[2\]
`vector3`           | float\[3\]
`json`              | JSON.parse() is vaild
