xls.q
=====

Wrapper on .h namespace to allow for formatting of cells in an excel document.


Examples
-------

Available styles:
```q
q).xls.style.help
s62| "comma separators (e.g. 100,000)"
s64| "percentage (must be 0.## format)"
s65| "bold"
s73| "red cell"
s74| "yellow cell"
s75| "green cell"
```

Applying some conditional colour formatting, number formatting, changing column names to bold and saving a table per tab:
```q
q) tabOrig:tabNew:([]c1:10 20;c2:100000 1000;c3:0.1 0.2;c4:`a`b);
q) update c1:.xls.as'[?[c1>10;`s73;`s75];string c1] from `tabNew;
q) update c2:.xls.as'[`s62;c2] from `tabNew;
q) update c3:.xls.as'[`s64;c3] from `tabNew;
q) tabNew:(`$.xls.as[`s65;]each cols tabNew) xcol tabNew;
q) .xls.write[`:file.xls;`tabOrig`tabNew]
```
Unformatted tab:

![Alt text](examples/tabOrig.png?raw=true "tabOrig")

Formatted tab:

![Alt text](examples/tabNew.png?raw=true "tabNew")