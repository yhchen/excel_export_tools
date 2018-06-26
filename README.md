yhc_excel_export
================


excel export with type check using nodejs
---


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
     |      char         | min:-127                    max:127                   |
     |      uchar        | min:0                        max:255                  |
     |      short        | min:-32768                max:32767                   |
     |      ushort       | min:0                        max:65535                |
     |      int          | min:-2147483648            max:2147483647             |
     |      uint         | min:0                        max:4294967295           |
     |      int64        | min:-9223372036854775808    max:9223372036854775807   |
     |      uint64       | min:0                        max:18446744073709551615 |
     |      string       | auto change 'line break' to '\n'                      |
     |      double       | ...                                                   |
     |      float        | ...                                                   |
     |      date         | yyyy/mm/dd HH:MM:ss    not support "Combination Type" |
     |      tinydate     | yyyy/mm/dd            not support "Combination Type"  |
     |      timestamp    | Linux time stamp        not support "Combination Type"|
     |      utctime      | UTC time stamp        not support "Combination Type"  |
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
