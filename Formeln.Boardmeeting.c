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
	(ISNUMBER(SEARCH(E1;tbArbeitsliste[Seriennr]))
	)+
	(ISNUMBER(SEARCH(E1;tbArbeitsliste[Lieferantenname]))
	)
)


/*
Filtert zuerst die Spalten anhand der Überschriften und dann..
Textfilter über zwei Spalten, in einer Zelle, der Treffer muss nur in einer erfolgen.
*/

=FILTER(
	FILTER(tbArbeitsliste;ISNUMBER(XMATCH(tbArbeitsliste[#Headers];Sheet1!C3:J3)));

		(ISNUMBER(SEARCH(E1;tbArbeitsliste[Seriennr])))+
		(ISNUMBER(SEARCH(E1;tbArbeitsliste[Lieferantenname]))
  )
)
/*
Gibt Zeile 2,4 & 6 aus Spalte 9,7 & 6 wider.
ACHTUNG: Bei Zeilen muss es ein Semikolon sein und bei Spalten ein Backslash
Die Ausgabe erfolgt in der angegebenen Reihenfolge
*/

=INDEX(AusgangstabelleODERBereich;

{2;4;6};{9\7\6})

/*
Gibt 5 Zeilen und 4 Spalten der AusgangstabelleODERBereich wider.
ACHTUNG das Semikolon in der zweiten SEQUENCE sagt Excel, dass es sich bei der Angabe um Spalten handelt.
*/

=INDEX(AusgangstabelleODERBereich;

	SEQUENCE(5);SEQUENCE(;4)
      )


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
  
  
  
  
/* Gibt eine Matrix von 7 Spalten und (Enddatum-Startdatum) aus.
    C1 = Startdatum
    C2 = Enddatum
*/

=(C1+SEQUENCE(((C2-C1)/7)+1;7))



/* Gibt die Wochennummer der Zellerechts neben der Zelle mit der Formel wider
    Adress gibt die aktuelle position im Blatt aus z.B. $D$8  diese Ausgabe muss man über INDIRECT wieder in einen Zellbezug bzw. auf den Inhalt der Zelle umwandeln.
    ;1 bei Adress und TRUE bei INDIRECT stehen für einen absoluten Zellbezug ( Alles geht von Zelle $A$1 aus)
*/

=WEEKNUM(
	INDIRECT(ADDRESS(ROW();COLUMN()+1;1);TRUE);2
)

  
  
  
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

/*
Gibt eine Art Kalender zurück
C1 = Startdatum
C2 = Enddatum
Das Array hat 8 Spalten
in der ersten Spalte steht die Kalenderwoche
in den anderen 7 Spalten steht das Enddatum
*/


=CHOOSE({1\2\2\2\2\2\2\2};

	WEEKNUM((C1+SEQUENCE(((C2-C1)/7)+1;;0;7));21);

	INDEX(C1-1+SEQUENCE(((C2-C1)/7)+1;7);SEQUENCE((C2-C1)/7+1);SEQUENCE(;8;0))

)

/* 
gibt die einzigartigen Werte einer Tabelle wider und filtert dabei in der zweiten Spüalte die werte aus, die " x" enthalten. 
(Die Werte, die " x" enthalten werden also NICHT angezeigt.)
 - UNIQUE(tbArbeitsliste[[Lieferantenname]:[Materialkurztext]]); => welcher Bereich soll gefiltert werden 
 - ISNUMBER(SEARCH(" x";INDEX(UNIQUE(tbArbeitsliste[[Lieferantenname]:[Materialkurztext]]);;2))  => Sucht in der zweiten Spalte nach Werten, die den Wert " x" enthalten
 - NOT( => dreht die Suche um, da wir ja Werte suchen, die nicht " x" enthalten und eine Suche nach ungleich " x" führt nicht zum erfolg, da die FILTER-Funktion bis jetzt keine Wildcard erlaubt
 - ;;2))  => in welcher Spalte des Bereichs soll gesucht werden
Komplette Formel siehe unten.
*/
	
=FILTER(
	UNIQUE(tbTabellenname[[SpalteVon]:[SpalteBis]]);
	NOT(ISNUMBER(SEARCH(" x";INDEX(UNIQUE(tbTabellenname[[SpalteVon]:[SpalteBis]]);;2)
			   )
		    )
	   );""
	)
  

  
  
  
  
  
  
  
  
  
