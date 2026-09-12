/ Write kdb+ tables to Excel SpreadsheetML (.xls), with per-cell styling.
/ Cell/row/worksheet encoding (ec0, es, edsn0, edsn) is written in k for
/ terseness and speed on large tables; everything else here is plain q.
\d .xls

style.default:.h.htac[`Style;(`ss:ID`ss:Name)!("Default";"Normal");"<Alignment ss:Vertical=\"Bottom\"/><Borders/><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\"/><Interior/><NumberFormat/><Protection/>"];
style.s62:.h.htac[`Style;(enlist`ss:ID)!enlist"s62";"<NumberFormat ss:Format=\"#,##0\"/>"];
style.s63:.h.htac[`Style;(enlist`ss:ID)!enlist"s63";"<NumberFormat ss:Format=\"$#,##0.00\"/>"];
style.s64:.h.htac[`Style;(enlist`ss:ID)!enlist"s64";"<NumberFormat ss:Format=\"0%\"/>"];
style.s65:.h.htac[`Style;(enlist`ss:ID)!enlist"s65";"<Font ss:FontName=\"Calibri\" ss:Size=\"11\" ss:Color=\"#000000\" ss:Bold=\"1\"/>"];
style.s66:.h.htac[`Style;(enlist`ss:ID)!enlist"s66";"<Font ss:FontName=\"Calibri\" ss:Size=\"11\" ss:Color=\"#000000\" ss:Italic=\"1\"/>"];
style.s67:.h.htac[`Style;(enlist`ss:ID)!enlist"s67";"<Font ss:FontName=\"Calibri\" ss:Size=\"11\" ss:Color=\"#000000\" ss:Underline=\"Single\"/>"];
style.s68:.h.htac[`Style;(enlist`ss:ID)!enlist"s68";"<NumberFormat ss:Format=\"dd/mm/yyyy\"/>"];
style.s69:.h.htac[`Style;(enlist`ss:ID)!enlist"s69";"<Alignment ss:Horizontal=\"Center\"/>"];
style.s70:.h.htac[`Style;(enlist`ss:ID)!enlist"s70";"<Alignment ss:Horizontal=\"Right\"/>"];
style.s71:.h.htac[`Style;(enlist`ss:ID)!enlist"s71";"<Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders>"];
style.s73:.h.htac[`Style;(enlist`ss:ID)!enlist"s73";"<Interior ss:Color=\"#FF0000\" ss:Pattern=\"Solid\"/>"];
style.s74:.h.htac[`Style;(enlist`ss:ID)!enlist"s74";"<Interior ss:Color=\"#FFFF00\" ss:Pattern=\"Solid\"/>"];
style.s75:.h.htac[`Style;(enlist`ss:ID)!enlist"s75";"<Interior ss:Color=\"#92D050\" ss:Pattern=\"Solid\"/>"];
style.s76:.h.htac[`Style;(enlist`ss:ID)!enlist"s76";"<Interior ss:Color=\"#4472C4\" ss:Pattern=\"Solid\"/>"];
style.s77:.h.htac[`Style;(enlist`ss:ID)!enlist"s77";"<Interior ss:Color=\"#FFA500\" ss:Pattern=\"Solid\"/>"];
style.s78:.h.htac[`Style;(enlist`ss:ID)!enlist"s78";"<Interior ss:Color=\"#D9D9D9\" ss:Pattern=\"Solid\"/>"];
styles:.h.htc[`Styles;raze 1_value .xls.style];

style.help:()!();
style.help[`s62]:"comma separators (e.g. 100,000)";
style.help[`s63]:"currency (e.g. $1,234.56)";
style.help[`s64]:"percentage (must be 0.## format)";
style.help[`s65]:"bold";
style.help[`s66]:"italic";
style.help[`s67]:"underline";
style.help[`s68]:"date (dd/mm/yyyy)";
style.help[`s69]:"center-aligned";
style.help[`s70]:"right-aligned";
style.help[`s71]:"thin border";
style.help[`s73]:"red cell";
style.help[`s74]:"yellow cell";
style.help[`s75]:"green cell";
style.help[`s76]:"blue cell";
style.help[`s77]:"orange cell";
style.help[`s78]:"gray cell";

k)ec0:{.h.htac[`Data;(,`ss:Type)!,$`String`Number`String`DateTime`DateTime`String i](.h.xs;$:;.h.xs@$:;.h.iso8601;.h.iso8601 1899.12.31+"n"$;.h.xs@$:)[i:-10 1 10 12 16 20h bin -@x]x};

ec:{$["<Cell"~5#v:$[-11h=type x;string x;x];v;.h.htc[`Cell;ec0 x]]};

k)es:{.h.htac[`Worksheet;(,`ss:Name)!,$x].h.htc[`Table]@,/(.h.htc[`Row]@,/ec')'(,!+y),+.+y:0!y};

k)edsn0:{"\r\n"/:(!x)es'. x};
k)edsn:{.h.ex .h.eb@styles,edsn0 x};

fmt:{
  d:{enlist[x]!enlist value x}each x;
  $[1<count x;raze d;d]
  };

write:{[f;t] f 0:edsn fmt t};

append:{[f;t]
  o:read0 f;
  o:@[o;count[o]-1;-11_];
  n:edsn0[fmt t],"</Workbook>";
  f 0:o,enlist n
  };

as:{[s;d]
  .h.htac[`Cell;$[null s;()!();(enlist`ss:StyleID)!enlist string s];ec0 d]
  };

\d .