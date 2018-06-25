yhc_excel_export
================


excel export with type check using nodejs
---


Support Format Check:
```
Base Type:
    -----------------------------------------------------------------------------
    |       type        |                        desc                           |
    -----------------------------------------------------------------------------
    |        char       | min:-127                    max:127                   |
    |        uchar      | min:0                        max:255                  |
    |        short      | min:-32768                max:32767                   |
    |        ushort     | min:0                        max:65535                |
    |        int        | min:-2147483648            max:2147483647             |
    |        uint       | min:0                        max:4294967295           |
    |        int64      | min:-9223372036854775808    max:9223372036854775807   |
    |        uint64     | min:0                        max:18446744073709551615 |
    |        string     | auto change 'line break' to '\n'                      |
    |        double     | ...                                                   |
    |        float      | ...                                                   |
    -----------------------------------------------------------------------------
Combination Type:
    -----------------------------------------------------------------------------
    | {"<name>":<type>}  | start with '{' and end with '}' with json format.    |
    |                    | <type> is one of "Base Type" or "Combination Type".  |
    -----------------------------------------------------------------------------
    | <type>[<N>]        | <type> is one of "Base Type" or "Combination Type".  |
    |                    | <N> is empty(variable-length) or number.             |
    -----------------------------------------------------------------------------
    | vector2            | float[2]                                             |
    -----------------------------------------------------------------------------
    | vector3            | float[3]                                             |
    -----------------------------------------------------------------------------
```
