
/* SummeWenn über einen Bereich, der mehrere Spalten umfasst */

=SUMPRODUCT((B5:B10="red")*(C5:E10))

/* Bsp. mit Tabellen- & Spaltennamen */
=SUMPRODUCT((Table1[Wer]=[@Wer])*(Table1[[Essen]:[Bier]]))

/* Bsp. mit Zellenangaben */
=SUMPRODUCT((B5:B10=F5)*(C5:E10))

/* =SUMPRODUCT((Spalte mit den Kriterien=Kriterium[in "=fester Wert" hier kann aber auch eine Zelle stehen])*(Bereich mit den Werten, die Summiert werden sollen)) */
