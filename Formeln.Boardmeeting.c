<>
/*
.........................................Muss das denn immer sein?...............................................................................

*/

/*
Zeigt alle Wagen an, bei denen die Materialnummer (Zelle C4) fehlt und gibt diese in der Zelle aus, in der die Formel steht.
Textjoin verbindet dabei die einzelnen Werte innerhalb einer Zelle.
Ohne Textjoin werden die einzelnen Werte jeweils in einer neuen Spalte ausgegeben.
Ohne TRANSPOSE werden die Werte Zeilenweise ausgegeben..
*/

=TEXTJOIN("; ";TRUE;
  TRANSPOSE(
    UNIQUE(
      FILTER(tbArbeitsliste[Seriennr];Auswertung!C4=tbArbeitsliste[Komponente])
    )
  )
)

/*
Unique + Spaltenfilter nach Headernamen und Filter der Seriennummer über ein dynamisches Dropdown in Zelle A2
*/

=FILTER(
  FILTER(UNIQUE(tbArbeitsliste);
      ISNUMBER(XMATCH(tbArbeitsliste[#Headers];Sheet1!C3:J3)));
          ISNUMBER(SEARCH(A2;tbArbeitsliste[Seriennr])
  )
)

/*
Textfilter über zwei Spalten, in einer Zelle, der Treffer muss nur in einer erfolgen.
*/

=FILTER(UNIQUE(tbArbeitsliste);
(
ISNUMBER(SEARCH(E1;tbArbeitsliste[Seriennr]))
)+
(
ISNUMBER(SEARCH(E1;tbArbeitsliste[Lieferantenname]))
  )
)


/*
Filtert zuerst die Spalten anhand der Überschriften und dann..
Textfilter über zwei Spalten, in einer Zelle, der Treffer muss nur in einer erfolgen.
*/

=FILTER(
FILTER(tbArbeitsliste;ISNUMBER(XMATCH(tbArbeitsliste[#Headers];Sheet1!C3:J3)));

(
ISNUMBER(SEARCH(E1;tbArbeitsliste[Seriennr])))+
(
ISNUMBER(SEARCH(E1;tbArbeitsliste[Lieferantenname]))
  )
)


=INDEX(tbArbeitsliste;SEQUENCE(ROWS(tbArbeitsliste));{9\7\6})

XMATCH(B3:C3;F3:G3)

/*
Gibt in einem dynamischen Array die Spalten aus, von denen der Tabellenkopf zur Tabellenkopf des Arrays passt
Tabelle hat z.B. 20 Spalten, im dynamischen array will ich aber nur 3, 8 und 15 haben und dann auch noch durch einander.
XMATCH(B3:D3 = die neuen Überschriften ( 15, 3, 8)
tbArbeitsliste[#Headers] = der Tabellenkop der Ausgangstabelle.
tbArbeitsliste = name der Ausgangstabelle
*/



=INDEX(tbArbeitsliste;

SEQUENCE(ROWS(tbArbeitsliste));XMATCH(B3:D3;tbArbeitsliste[#Headers])


)
