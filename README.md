yhc_excel_export
================


excel export with type check using nodejs in super fast speed!
---

How To Use:
---
* Configure file is `config.json`. Add files and directory at `[IncludeFilesAndPath]` to export. Add files at `[ExcludeFileNames]` to disable file export. Add table name at `[ExcludeCsvTableNames]` to disable table output.
* Set output `table name` at cell `B1`(Default value). It can be change at config Key `[CSVNameCellID]`.
* Add output table `Column Name` Row under `table name` Row. If a Column Name Start with `#` means is was a comment Column. It can be disable output at config `EnableExportCommentColumns`.
* Add `Column Format Type` line under `Column Name` Row.
* `[Export]`  is temporary not supported.
* A comment Row was start with word `#` at `A Column`. Is can be disable output at config `EnableExportCommentRows`.
* Cell `[N]A([N] is Row index)` start with `#` is a comment line.



Support Format Check:
```
Support Format Check:
     Using typescript type define link grammar.
     Extends number type for digital size limit check.
     Add common types of game development(like vector2 vector3...);

Base Type:
     -----------------------------------------------------------------------------
     |        type       |                        desc                           |
     -----------------------------------------------------------------------------
     |      char         | min:-127                  max:127                     |
     |      uchar        | min:0                     max:255                     |
     |      short        | min:-32768                max:32767                   |
     |      ushort       | min:0                     max:65535                   |
     |      int          | min:-2147483648           max:2147483647              |
     |      uint         | min:0                     max:4294967295              |
     |      int64        | min:-9223372036854775808  max:9223372036854775807     |
     |      uint64       | min:0                     max:18446744073709551615    |
     |      string       | auto change 'line break' to '\n'                      |
     |      double       | ...                                                   |
     |      float        | ...                                                   |
     |      date         | YYYY/MM/DD HH:mm:ss                                   |
     |      tinydate     | YYYY/MM/DD                                            |
     |      timestamp    | Linux time stamp                                      |
     |      utctime      | UTC time stamp                                        |
     -----------------------------------------------------------------------------

Combination Type:
     -----------------------------------------------------------------------------
     | {<name>:<type>}   | start with '{' and end with '}' with json format.     |
     |                   | <type> is one of "Base Type" or "Combination Type".   |
     -----------------------------------------------------------------------------
     | <type>[<N>|null]  | <type> is one of "Base Type" or "Combination Type".   |
     |                   | <N> is empty(variable-length) or number.              |
     -----------------------------------------------------------------------------
     | vector2           | float[2]                                              |
     -----------------------------------------------------------------------------
     | vector3           | float[3]                                              |
     -----------------------------------------------------------------------------
     | json              | JSON.parse() is vaild                                 |
     -----------------------------------------------------------------------------

```
