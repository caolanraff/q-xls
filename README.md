# xls.q

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A kdb+/q library for writing tables to Excel (`.xls`, SpreadsheetML XML
format). It wraps the built-in `.h` namespace to support per-cell styling —
number formats, font styles, alignment, borders, and conditional cell
colouring — and appending extra tabs to a workbook that's already been
written.

## Requirements

- [kdb+](https://kx.com/) (uses the built-in `.h` namespace, no external
  dependencies)

## Usage

Load the library:

```q
q)\l xls.q
```

Available styles:

```q
q).xls.style.help
s62| "comma separators (e.g. 100,000)"
s63| "currency (e.g. $1,234.56)"
s64| "percentage (must be 0.## format)"
s65| "bold"
s66| "italic"
s67| "underline"
s68| "date (dd/mm/yyyy)"
s69| "center-aligned"
s70| "right-aligned"
s71| "thin border"
s73| "red cell"
s74| "yellow cell"
s75| "green cell"
s76| "blue cell"
s77| "orange cell"
s78| "gray cell"
```

Applying some conditional colour formatting, number formatting, changing
column names to bold and saving a table per tab:

```q
q) tabOrig:tabNew:([]c1:10 20;c2:100000 1000;c3:0.1 0.2;c4:`a`b);
q) update c1:.xls.as'[?[c1>10;`s73;`s75];c1] from `tabNew;
q) update c2:.xls.as'[`s62;c2] from `tabNew;
q) update c3:.xls.as'[`s64;c3] from `tabNew;
q) tabNew:(`$.xls.as[`s65;]each cols tabNew) xcol tabNew;
q) .xls.write[`:file.xls;`tabOrig`tabNew]
```

Unformatted tab:

![Alt text](examples/tabOrig.png?raw=true "tabOrig")

Formatted tab:

![Alt text](examples/tabNew.png?raw=true "tabNew")

## API

| Function | Description |
| --- | --- |
| `.xls.write[file;tables]` | Write one or more tables to a new workbook, one tab per table. |
| `.xls.append[file;tables]` | Append one or more tables as new tabs to an existing workbook. |
| `.xls.as[style;data]` | Tag a column/value with a style so it renders formatted when written. |
| `.xls.style.help` | Dictionary describing the available built-in styles. |

## License

Released under the [MIT License](LICENSE).
