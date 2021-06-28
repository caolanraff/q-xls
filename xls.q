\d .xls

style.default:.h.htac[`Style;(`ss:ID`ss:Name)!("Default";"Normal");"<Alignment ss:Vertical=\"Bottom\"/><Borders/><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\"/><Interior/><NumberFormat/><Protection/>"];
style.s62:.h.htac[`Style;(enlist`ss:ID)!enlist"s62";"<NumberFormat ss:Format=\"#,##0\"/>"];
style.s64:.h.htac[`Style;(enlist`ss:ID)!enlist"s64";"<NumberFormat ss:Format=\"0%\"/>"];
style.s65:.h.htac[`Style;(enlist`ss:ID)!enlist"s65";"<Font ss:FontName=\"Calibri\" ss:Size=\"11\" ss:Color=\"#000000\" ss:Bold=\"1\"/>"];
style.s73:.h.htac[`Style;(enlist`ss:ID)!enlist"s73";"<Interior ss:Color=\"#FF0000\" ss:Pattern=\"Solid\"/>"];
style.s74:.h.htac[`Style;(enlist`ss:ID)!enlist"s74";"<Interior ss:Color=\"#FFFF00\" ss:Pattern=\"Solid\"/>"];
style.s75:.h.htac[`Style;(enlist`ss:ID)!enlist"s75";"<Interior ss:Color=\"#92D050\" ss:Pattern=\"Solid\"/>"];
styles:.h.htc[`Styles;raze 1_value .xls.style];

style.help:()!();
style.help[`s62]:"comma separators (e.g. 100,000)";
style.help[`s64]:"percentage (must be 0.## format)";
style.help[`s65]:"bold";
style.help[`s73]:"red cell";
style.help[`s74]:"yellow cell";
style.help[`s75]:"green cell";

k)ec0:{.h.htac[`Data;(,`ss:Type)!,$`String`Number`String`DateTime`DateTime`String i](.h.xs;$:;.h.xs@$:;.h.iso8601;.h.iso8601 1899.12.31+"n"$;.h.xs@$:)[i:-10 1 10 12 16 20h bin -@x]x};

ec:{$["<Cell"~5#v:$[-11h=type x;string x;x];v;.h.htc[`Cell;ec0 x]]};

k)es:{.h.htac[`Worksheet;(,`ss:Name)!,$x].h.htc[`Table]@,/(.h.htc[`Row]@,/ec')'(,!+y),+.+y:0!y};

k)edsn0:{"\r\n"/:(!x)es'. x};
k)edsn:{.h.ex .h.eb@styles,edsn0 x};

write:{[f;t]
  d:{enlist[x]!enlist value x}each t;
  f 0:edsn $[1<count t;raze d;d]
  };

append:{[f;t]
  o:read0 f;
  o:@[o;count[o]-1;-11_];
  d:{enlist[x]!enlist value x}each t;
  n:edsn0 $[1<count t;raze d;d];
  n,:"</Workbook>";
  f 0:o,enlist n
  }

as:{[s;d]
  .h.htac[`Cell;$[null s;()!();(enlist`ss:StyleID)!enlist string s];ec0 d]
  };

\d .