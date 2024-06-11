\\/////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\				     
/////////////////////////// MCRO SKRIP /////////////////
\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
///////////////////////////////////////////////////////////////////////////////

////// Modul 1\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


'*****************************************************************************
'
'               Programm fuer die Abrechung der Einkaeufe
'       ======================================================
'
'   Es wird eine leere Liste Erzeugt.
'   Zuerst muss der Name  eingegeben werden; anschließend der Betrag der fuer den Einkauf zur
'   Verfuehgung steht.
'   Das Programm sucht die naechste leere Zeile. Dann muss die Artikel Menge
'   eingegben werden. Dann der Stueckpreis des Artikels. Danach wird
'   abgefragt ob noch weitere Artikel eingegeben werden sollen. Es folgt die
'   Abfrage nach dem Druckbefehl.
'
'   Konstanten in allen Modulen verfuehgbar:
'   NAME        String      gibt die Zelle fuer den Namen an
'   BETRAG      String      gibt die Zelle fuer den Einkaufbetrag an
'   RESTGELD    String      gibt die Zelle fuer den Restgeldbetrag an
'   SUMME       String      gibt die Zelle an fuer die Summe
'                           aller Artikelpreise
'
'   Variablen:
'   Pname   String  ein Name des Einkäufers
'   Szahl   Integer ein Menge der gekauften Artikel
'   Apreis  Double  ein Stueckpreis des Artikels
'   Gbetrag Double  ein Betrag der fuer den Einkauf gegeben wurde
'   Rest    Double  aus Restbetrag nach abzug aller Artikel
'                   kann positiv oder negativ sein
'
'*****************************************************************************

'' Konstanten Definieren
Public Const BETRAG = "$D$2"
Public Const RESTGELD = "$E$2"
Public Const NAME = "$F$2"
Public Const SUMME = "$C$31"

'' Variablen definieren
Dim NameP, NameE, NameT As String
Dim BetragE, BetragT, MengeE, MengeT, PreisE, PreisT, NextE, NextT As String
Dim Gbetrag, Apreis, Rest As Double
Sub ZeigerA1()
'''''''''''''''''''''' Zeiger auf A1 setzen ''''''''''''''''''''''''''''''''''
        Worksheets("Abrechnung").Activate
        Worksheets("Abrechnung").Range("A1").Activate
        ActiveCell.Select
        
''''''' solange aktive Zelle nicht leer  ist '''''''''''''''''''''''''''''''''
            While (Not (ActiveCell.VALUE = ""))
                 ActiveCell.Offset(rowOffset:=1).Activate
            Wend
End Sub
Sub AnzPos()
'''''' Anzahl der eingetragen Artikel / Positionen ausgeben ''''''''''''''''''
'''''''''''''''''''''' Zeiger auf A1 setzen ''''''''''''''''''''''''''''''''''
        ZeigerA1
        
''''''''' Blattschutz aufheben '''''''''''''''''''''''''''''''''''''''''''''''
        ActiveSheet.Unprotect
        
''''' Zeiger auf listen ende setzen''''''''''''''''''''''''''''''''''''''''''
                cmax = ZeilenSuche()
                cord = "E" & cmax
                cordl = cmax - 2
                cordt = "D" & cmax
                
'''' Beschriftung "Anzahl Positionen :" einfuegen und Zelle formatieren ''''''
                Worksheets("Abrechnung").Range(cordt).Activate
                ActiveCell.Select
                Selection.Locked = False
                ActiveCell.VALUE = "Anzahl Positionen: "
                Selection.Font.Bold = True
                
'''' Zelle formatiern ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            Selection.Borders(xlEdgeLeft).LineStyle = xlNone
                With Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
            Selection.Borders(xlEdgeBottom).LineStyle = xlNone
                With Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
            Selection.Borders(xlInsideVertical).LineStyle = xlNone
            Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.349986266670736
                    .PatternTintAndShade = 0
                End With
            Selection.Locked = True
            
'''''' Wert Anzahl Positioen in Zelle einfuegen ''''''''''''''''''''''''''''''
                Worksheets("Abrechnung").Range(cord).Activate
                ActiveCell.Select
                
''''' Zellen Format loeschen '''''''''''''''''''''''''''''''''''''''''''''''''
                ActiveCell.ClearFormats
                
'''' Zelle entsperren ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Selection.Locked = False
                
'''' Wert in Zelle einsetzen '''''''''''''''''''''''''''''''''''''''''''''''''
                ActiveCell.Select
                ActiveCell.FormulaR1C1 = cordl
                Selection.Font.Bold = True
                Selection.NumberFormat = "#0"
                
'''' Zelle formatiern ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            Selection.Borders(xlEdgeLeft).LineStyle = xlNone
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            Selection.Borders(xlEdgeBottom).LineStyle = xlNone
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            Selection.Borders(xlInsideVertical).LineStyle = xlNone
            Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.349986266670736
                .PatternTintAndShade = 0
            End With
            
End Sub
Sub Einf()
'''''''''''''''''''''' Zeiger auf A1 setzen ''''''''''''''''''''''''''''''''''
        ZeigerA1
        
'''''''''''''''''''''''' Artikel Menge und Preis eingeben ''''''''''''''''''''
'''''''''''''''' MengeE Eingabe Meldung: Artikel Menge '''''''''''''''''''''''
        MengeE = "Artikelmenge eingeben"
        
'''''''''''''''''' MsgBox Titel definieren '''''''''''''''''''''''''''''''''''
        MengeT = "Stueckzahl"
        
'''' Zeilenmarke 106 festlegen '''''''''''''''''''''''''''''''''''''''''''''''
106:

''''''''' Stueckzahl in InputBox eingeben und in Variable  Szahl speichern '''
        Szahl = Application.InputBox(prompt:=MengeE, Title:=MengeT, _
                Default:=0#, Type:=1 + 2)
                
'''''''' wenn Abbrechen gedrueckt wurde ''''''''''''''''''''''''''''''''''''''
                If Szahl = FALSCH Then
                    Szahl = 0#
                    Selection.NumberFormat = "#0"
                    ActiveCell.FormulaR1C1 = Szahl
                    Modul2.CellFormat1
                Else
                
'''Pruefen ob eine Zahl eingegeben wurde und Wert in aktive Zelle einsetzen'''
                    If (IsNumeric(Szahl)) Then
                        Selection.NumberFormat = "#0"
                        ActiveCell.FormulaR1C1 = Szahl
                        Modul2.CellFormat1
                    Else
                        Meldung = MsgBox("Bitte eine Zahl eingeben!", _
                        vbOKOnly + 16, "FEHLER!")
                        GoTo 106
                    End If
                End If
'''''''''''''' Zeiger eine Zeile weiter nach rechts setzen '''''''''''''''''''
        ActiveCell.Offset(columnOffset:=1).Activate
''''''''''''''' PreisE Eingabe Meldung: Artikelpreis '''''''''''''''''''''''''
        PreisE = "Einzelpreis eingeben"
'''''''''''''''' MsgBox Titel definieren '''''''''''''''''''''''''''''''''''''
        PreisT = "Artikelpreis"
'''' Zeilenmarke 107 festlegen '''''''''''''''''''''''''''''''''''''''''''''''
107:
''''''' Stueckpreis in InputBox eingeben und in Variable Szahl speichern '''''
        Apreis = Application.InputBox(prompt:=PreisE, Title:=PreisT, _
                    Default:=0#, Type:=1 + 2)
            If Apreis = FALSCH Then
                Apreis = 0#
                Selection.NumberFormat = "#,##0.00 $"
                ActiveCell.FormulaR1C1 = Apreis
                Modul2.CellFormat1
                Modul2.CellTotal
            Else
'''Pruefen ob eine Zahl eingegeben wurde und Wert in aktive Zelle einsetzen ''
                    If (IsNumeric(Apreis)) Then
                            Selection.NumberFormat = "#,##0.00 $"
                            ActiveCell.FormulaR1C1 = Apreis
                            Modul2.CellFormat1
                    Else
                        Meldung = MsgBox("Bitte eine Zahl eingeben!", _
                        vbOKOnly + 16, "FEHLER!")
                    GoTo 107
                    End If
'''''' nach Eingabe Produkt aus Menge und Einzelpreis bilden '''''''''''''''''
        CellTotal
            End If
'''' Zeiger zuruecksetzen: eine Spalte nach links... '''''''''''''''''''''''''
        ActiveCell.Offset(columnOffset:=-1).Activate
''' ...und eine Zeile weiter runter ''''''''''''''''''''''''''''''''''''''''''
        ActiveCell.Offset(rowOffset:=1).Activate
'''''''''' Anzahl der eingetragen Artikel ausgeben '''''''''''''''''''''''''''
        AnzPos
'''' Artikel aufsummieren und Gesamtsumme errechnen ''''''''''''''''''''''''''
        Sum
''''''''' Blatt schuetzen ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, _
        Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, _
        AllowInsertingRows _
        :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, _
        AllowFiltering:=True, AllowUsingPivotTables:=True
''' Restgeld berechnen '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Restbetrag
End Sub
Sub Sum()
'''' Artikel aufsummieren und Gesamtsumme errechnen ''''''''''''''''''''''''''
'''''''''''''''''''''' Zeiger auf A1 setzen ''''''''''''''''''''''''''''''''''
        ZeigerA1
''''''''' Blattschutz aufheben '''''''''''''''''''''''''''''''''''''''''''''''
        ActiveSheet.Unprotect
''''' Zeiger auf listen ende setzen''''''''''''''''''''''''''''''''''''''''''
                gmax = ZeilenSuche()
                ges = "C" & gmax
                gesl = gmax - 2
                gest = "=SUM(R[-" & gesl & "]C:R[-1]C)"
                Worksheets("Abrechnung").Range(ges).Activate
                ActiveCell.Select
''''' Zellen Format loeschen '''''''''''''''''''''''''''''''''''''''''''''''''
                ActiveCell.ClearFormats
'''' Zelle entsperren ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Selection.Locked = False
'''' Wert in Zelle einsetzen''''''''''''''''''''''''''''''''''''''''''''''''''
        ActiveCell.Select
        ActiveCell.FormulaR1C1 = gest
        Selection.Font.Bold = True
'''' Zelle formatiern ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            Selection.Borders(xlEdgeLeft).LineStyle = xlNone
            With Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            Selection.Borders(xlEdgeBottom).LineStyle = xlNone
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            Selection.Borders(xlInsideVertical).LineStyle = xlNone
            Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = -0.349986266670736
                .PatternTintAndShade = 0
            End With
End Sub
Sub Main()
''''''''' Alte Daten aus Arbeitsblatt entfernen ''''''''''''''''''''''''''''''
    Modul2.ClearCells
''''''''' Blattschutz aufheben '''''''''''''''''''''''''''''''''''''''''''''''
    ActiveSheet.Unprotect
''''''''' Neue Zeile Einfuegen '''''''''''''''''''''''''''''''''''''''''''''''
    Modul2.ZeilePlus
'''' Zellen entsperren '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Selection.Locked = False
''''' Zellen sichtbar machen '''''''''''''''''''''''''''''''''''''''''''''''''
    Selection.FormulaHidden = False
''''''''' Blatt schuetzen ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, _
        AllowInsertingRows:=True, AllowInsertingHyperlinks:=True, _
        AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True
'''''''''''''''''''''''''' Name des Patienten eingeben '''''''''''''''''''''''
'''''''''''' Curser auf Zelle F2 setzen und aktive Zelle weiss faerben '''''''
        Worksheets("Abrechnung").Activate
        Worksheets("Abrechnung").Range(NAME).Activate
        ActiveCell.Select
        Modul2.WhiteCell
'''''''''''''''' NameP Eingabe Meldung: Artikel Menge ''''''''''''''''''''''''
        NameE = "Bitte einen Namen eingeben!"
'''''''''''''''''' InputBox Titel definieren '''''''''''''''''''''''''''''''''
        NameT = "Name eingeben!"
'''' Zeilenmarke 100 festlegen '''''''''''''''''''''''''''''''''''''''''''''''
100:
''''''''''''''''' Pruefung ob Zelle 'Name'leer ist '''''''''''''''''''''''''''
    If (Worksheets("Abrechnung").Range(NAME) = "") Then
''''''''''''''''''''''' wenn leer dann... , sonst weiter bei -101- '''''''''''
''''''''' Name in InputBox eingeben und in Variable NameP speichern ''''''''''
        NameP = Application.InputBox(prompt:=NameE, Title:=NameT, _
                    Type:=2)
              ActiveCell.FormulaR1C1 = NameP
                If (Worksheets("Abrechnung").Range(NAME) = "Falsch") Then
                    GoTo 303
                Else
                    If (Not (IsNumeric(NameP))) Then
                        ActiveCell.FormulaR1C1 = NameP
                        Modul2.Frame
                    Else
                        Meldung = MsgBox("Bitte einen Namen eingeben!", _
                        vbOKOnly + 48, "FEHLER!")
                        NameP = ""
                        ActiveCell.FormulaR1C1 = NameP
                        Modul2.Frame
                        GoTo 100
                    End If
                End If
                GoTo 100
    Else
        GoTo 101
    End If
'''' Zeilenmakarke 101 festlegen '''''''''''''''''''''''''''''''''''''''''''''
101:
'''''''''''''' Curser auf Zelle D2 setzen und aktive Zelle weiss faerben '''''
        Worksheets("Abrechnung").Activate
        Worksheets("Abrechnung").Range(BETRAG).Activate
        ActiveCell.Select
        Modul2.WhiteCell
''''''''''''''''' Geldbetrag fuer Einkauf eingeben '''''''''''''''''''''''''''
'''''''''''''''' MengeE Eingabe Meldung: Geld Betrag '''''''''''''''''''''''''
       BetragE = "Geld Betrag fuer Einkauf eingeben"
        
'''''''''''''''''' InputBox Titel definieren '''''''''''''''''''''''''''''''''
        BetragT = "Geldbetrag"
'''' Zeilenmarke 108 festlegen '''''''''''''''''''''''''''''''''''''''''''''''
108:
''''''''' Betrag in InputBox eingeben und in Variable  Gbetrag speichern '''''
        Gbetrag = Application.InputBox(prompt:=BetragE, Title:=BetragT, _
                    Default:=0#, Type:=1 + 2)
        If Gbetrag = FALSCH Then
                Gbetrag = 0#
                Selection.NumberFormat = "#,##0.00 $"
                ActiveCell.FormulaR1C1 = Gbetrag
                Modul2.Frame
        Else
'''Pruefen ob eine Zahl eingegeben wurde und Wert in aktive Zelle einsetzen'''
                If (IsNumeric(Gbetrag)) Then
                    Selection.NumberFormat = "#,##0.00 $"
                    ActiveCell.FormulaR1C1 = Gbetrag
                    Modul2.Frame
                Else
                    Meldung = MsgBox("Bitte eine Zahl eingeben!", _
                    vbOKOnly + 16, "FEHLER!")
                GoTo 108
                End If
        End If
''''''''''''''''' Pruefung ob Zelle 'Betrag fuer Einkauf'leer ist ''''''''''''
    If (Worksheets("Abrechnung").Range(BETRAG) = "" Or 0) Then
''''''''''''''''''''''' wenn leer dann... , sonst weiter bei -102- '''''''''''
'''''''''''''' Curser auf Zelle D2 setzen ''''''''''''''''''''''''''''''''''''
            Worksheets("Abrechnung").Activate
            Worksheets("Abrechnung").Range(BETRAG).Activate
            ActiveCell.Select
''''''''''''''''''''''''Geldbetrag eingeben ''''''''''''''''''''''''''''''''''
'''' Zeilenmarke 109 festlegen '''''''''''''''''''''''''''''''''''''''''''''''
109:
'''' Betrag in InputBox eingeben und in Variable  Gbetrag speichern ''''''''''
            Gbetrag = Application.InputBox(prompt:=BetragE, Title:=BetragT, _
            Default:=0#, Type:=1 + 2)
                If Gbetrag = FALSCH Then
                    Gbetrag = 0#
                    Selection.NumberFormat = "#,##0.00 $"
                    ActiveCell.FormulaR1C1 = Gbetrag
                    Modul2.Frame
                Else
'''Pruefen ob eine Zahl eingegeben wurde und Wert in aktive Zelle einsetzen'''
                    If (IsNumeric(Gbetrag)) Then
                        Selection.NumberFormat = "#,##0.00 $"
                        ActiveCell.FormulaR1C1 = Gbetrag
                        Modul2.Frame
                    Else
                        Meldung = MsgBox("Bitte eine Zahl eingeben!", _
                        vbOKOnly + 16, "FEHLER!")
                        GoTo 109
                    End If
                End If
'''''''''' Zeiger auf Feld A2 setzen''''''''''''''''''''''''''''''''''''''''''
        ActiveCell.Offset(columnOffset:=-3).Activate
        GoTo 102:
    Else
'''' Zeilenmarke 102 festlegen '''''''''''''''''''''''''''''''''''''''''''''''
102:
''''''''Zeilenweise Menge und Einzelpreis einfuegen und Produkt bilden '''''''
        Einf
''''''''''''''''''''''' Abfrage weitere Artikel eingeben '''''''''''''''''''''
        NextE = "Weiteren Artikel eingeben?"
        NextT = "Weitere Eingabe?"
        AntwortN = MsgBox(NextE, 4, NextT)
            If AntwortN = 6 Then
                Modul2.ZeilePlus
                GoTo 102
            Else
''''''' Betrag aufsummieren, Duck abfrage & Programm beenden  ''''''''''''''''
                Modul2.Drucken
            End If
'''' Zeilenmarke 303 festlegen '''''''''''''''''''''''''''''''''''''''''''''''
303:
    End If
End Sub









////// Modul 2 \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'*****************************************************************************
'
'               Modul mit Funktionen fuer EinkaufAbrechnung
'       =======================================================
'       - sDebug Code fuer die Debugg-Meldungen.
'       - ZeilenSuche(): Code fuer die Funktion fuer die Suche der
'           naechsten freien Zeile in der Splate "Menge".
'       - BetragSuche(): Funktion fuer Suche der naechsten freien Zeile in
'           der Splate "Gesamt". Von dem Wert wird 1 subtrahiert.
'       - CellTotal(): Funktion fuer Produkt aus Menge und Einzelpreis.
'        Code Zellenformatierungen:
'       - WhiteCell(): Zellen weiss faerben
'       - GreyCell(): Zellen grau faerben
'       - BedingtesFromat(): Zellen rot bei Werten <0, grau bei Werten =0
'           gruen bei Werten >0
'       - ZeilePlus():  Funktion zum hinzufuegen von Zeilen.
'       - ZeileMinus(): Funktion zum loeschen von Zeilen.
'       - ClearCells(): Funktion zum loeschen der Inhalte am Anfang
'           des Programms.
'       - CellFormat1(): Rahmen oben, unten, rechts
'       - Frame(): Rahmen oben, unten, rechts, links
'       - Plus(): Pruefen ob Felder "Name" & "Betrag fuer Einkauf" ausgefuellt
'            sind und Zeilen einfuegen wenn Pruefung erfolgreich war.
'       - Restbetrag(): Summe vom Betrag fuer Einkkauf abziehen
'           und Restgeld berechnen
'       - Drucken(): Funktion fuer die Druckbefehlabfrage.
'       - Die Funktionen fuer die Geldschein-Icons(addTen, addTwenty...).
'
'*****************************************************************************
'' Use for Debuging
Sub sDebug()
        Dim Test As String
        
''''''''''''' Debugg Ausgabe in msgBox '''''''''''''''''''''''''''''''''''''''
'Der Wert der Laufzeitvariablen I und der
'Visual Basic-Objekttyp der Markierung angezeigt.
       MsgBox "DER WERT VON ZeilenSuche (I) ist: " & ZeilenSuche & Chr$(13) & _
        "The selection object type is: " & TypeName(Selection) & Chr$(13) & _
       "DER WERT VON BetragSuche (N) ist: " & BetragSuche & Chr$(13) & _
        "Der Wert von CORD ist: " & ZeilenSuche + 1 & Chr$(13)
        'Zeile
        'Test = ZeilenSuche() & ":" & ZeilenSuche() + 3
End Sub
Public Function ZeilenSuche() As Integer
'''''''' leere Zellen in Splate Menge suchen '''''''''''''''''''''''''''''''''
'''' Laufzeit Variable deklarieren '''''''''''''''''''''''''''''''''''''''''''
                Dim i As Integer
                
'''' Laufzeit Variable initialisieren ''''''''''''''''''''''''''''''''''''''''
                i = 1
                Worksheets("Abrechnung").Activate
                Worksheets("Abrechnung").Range("A1").Activate
                ActiveCell.Select
                    While (Not (ActiveCell.VALUE = "Summe:"))
                        i = i + 1
                        ActiveCell.Offset(rowOffset:=1).Activate
                    Wend
                ZeilenSuche = i
End Function
Public Function BetragSuche() As Integer
'' Betraege in Einzelpreis suchen um die Anzahl der Positionen zu ermitteln ''
'''' Laufzeit Variable deklarieren '''''''''''''''''''''''''''''''''''''''''''
                Dim n As Integer
                
'''' Laufzeit Variable initialisieren ''''''''''''''''''''''''''''''''''''''''
                n = 0
                Worksheets("Abrechnung").Activate
                Worksheets("Abrechnung").Range("B1").Activate
                ActiveCell.Select
                    While (Not (ActiveCell.VALUE = ""))
                        n = n + 1
                        ActiveCell.Offset(rowOffset:=1).Activate
                    Wend
                BetragSuche = n - 1
End Function
 Sub CellTotal()
'''''''' Produkt aus Menge und Einzelpreis bilden ''''''''''''''''''''''''''''
        Worksheets("Abrechnung").Activate
        ActiveCell.Offset(columnOffset:=1).Activate
        ActiveCell.Select
        ActiveCell.FormulaR1C1 = "=SUM(RC[-2]*RC[-1])"
        Selection.NumberFormat = "#,##0.00 $"
        CellFormat1
End Sub
Sub WhiteCell()
''''' Zellen weiss faerben '''''''''''''''''''''''''''''''''''''''''''''''''''
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub GreyCell()
''''' Zellen grau faerben ''''''''''''''''''''''''''''''''''''''''''''''''''''
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
        .PatternTintAndShade = 0
    End With
End Sub
Sub BedingtesFormat()
''''''''''''''''''''Bedingte Formatierung: '''''''''''''''''''''''''''''''''''
''''''''''''' rot bei Werten <0, grau Werten = 0, gruen bei Werten >0 ''''''''
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
    Selection.FormatConditions(Selection.FormatConditions.Count) _
        .SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(1).VALUE = -0.1
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 5263615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(2).VALUE = 0
        With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.249977111117893
        End With
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(3).VALUE = 0.1
    With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 6750105
        .TintAndShade = 0
    End With
End Sub
Sub ZeilePlus()
''''''' Zeile einfuegen ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim location As String
    Dim cord As String
    cord = ZeilenSuche()
    location = cord & ":" & cord
    Rows(location).Select
    
    ''''''''' Blattschutz aufheben '''''''''''''''''''''''''''''''''''''''''''
    ActiveSheet.Unprotect
    
''''''''' Neue Zeile Einfuegen '''''''''''''''''''''''''''''''''''''''''''''''
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromAbove
    
'''' Zellen entsperren '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Selection.Locked = False
    
''''' Zellen sichtbar machen '''''''''''''''''''''''''''''''''''''''''''''''''
    Selection.FormulaHidden = False
    
''''' Zellen Formate loeschen ''''''''''''''''''''''''''''''''''''''''''''''''
    Selection.ClearFormats
    
'''' Zellen entsperren '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Selection.Locked = False
    GreyCell
    
End Sub
Sub ZeileMinus()
''''''' Zeile loeschen '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim location As String
    Dim cord As String
        cord = ZeilenSuche() + 1
        location = cord & ":" & cord
        Rows(location).Select
        Selection.Delete Shift:=xlUp
End Sub
Sub ClearCells()
'''''''' Alle Zeilen bis auf erste entfernen und letzte ''''''''''''''''''''''
       If (ZeilenSuche() > 2) Then
            lst = ZeilenSuche() - 1
            zone = "2:" & lst
            Rows(zone).Select
            Selection.Delete Shift:=xlUp
        Else
        End If
End Sub
Sub CellFormat1()
''''''''''''''''' Zelle nach Eingabe Fromatieren '''''''''''''''''''''''''''''
''''' Zellen weiss faerben '''''''''''''''''''''''''''''''''''''''''''''''''''
        WhiteCell
        
'''' Rahmen : oben, unten, rechts ''''''''''''''''''''''''''''''''''''''''''''
                Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                    With Selection.Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlEdgeBottom)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
                    With Selection.Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = xlThin
                    End With
End Sub
Sub Frame()
'''''' Zelle Rahmenlinen zeichnen ''''''''''''''''''''''''''''''''''''''''''''
                    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                        With Selection.Borders(xlEdgeLeft)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeTop)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                        With Selection.Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThin
                        End With
                    Selection.Borders(xlInsideVertical).LineStyle = xlNone
                    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
Sub Plus()
''''' Pruefen ob Felder "Name" & "Betrag fuer Einkauf" aus gefuellt sind '''''
        If ((Worksheets("Abrechnung").Range(NAME) = "") Or _
        (Worksheets("Abrechnung").Range(NAME) = FALSCH) Or _
        (Worksheets("Abrechnung").Range(BETRAG) = "")) Then
            Meldung = MsgBox _
            ("Bitte zuerst ""Namen"" und ""Betrag fuer Einkauf"" eingeben! ", _
            vbOKOnly + 16, "FEHLER!")
        Else
        
''''' Zeile einfuegen ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Modul2.ZeilePlus
            Modul1.Einf
        End If
        
End Sub
Sub Restbetrag()
'''' Gesamtsumme vom Einkaufsgeld abziehen und Restgeld ausrechen ''''''''''''
        res = "E2"
        resc = ZeilenSuche() - 2
        Worksheets("Abrechnung").Range(res).Activate
        ActiveCell.Select
        Modul2.BedingtesFormat
        ActiveCell.Select
        resg = "=SUM(RC[-1]-R[" & resc & "]C[-2])"
        ActiveCell.FormulaR1C1 = resg
        
''''' Zellen erfassen und Restsumme ermitteln ''''''''''''''''''''''''''''''''
        Selection.NumberFormat = "#,##0.00 $"
        Modul2.Frame
        Modul1.ZeigerA1
        
End Sub
Sub Drucken()
'''' Drucken Funktion ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim pr As String
        pr = "$A$1:$F$" & ZeilenSuche() + 1
        PrintE = "Drucken?"
        PrintT = "Abrechung Drucken?"
        StileP = vbYesNo + vbQuestion + vbDefaultButton1
        AntwortP = MsgBox(PrintE, StileP, PrintT)
            If AntwortP = 6 Then
                ActiveSheet.PageSetup.PrintArea = pr
                ActiveWindow.SelectedSheets.PrintOut Copies:=1
            Else
            End If
End Sub
Sub addTen()
''''' Pruefen ob Felder "Name" & "Betrag fuer Einkauf" ausgefuellt sind ''''''
        If (Worksheets("Abrechnung").Range(NAME) = "") Then
            Meldung = MsgBox _
            ("Bitte zuerst ""Abrechnung Starten"" und ""Name""" & _
            "eingeben! ", vbOKOnly + 16, "FEHLER!")
        Else
        
'''''''''''''' Curser auf Zelle D2 setzen ''''''''''''''''''''''''''''''''''''
            Worksheets("Abrechnung").Activate
            Worksheets("Abrechnung").Range(BETRAG).Activate
            ActiveCell.Select
            
''''''' Betrag auf 10 setzen und einfuegen ''''''''''''''''''''''''''''''''''''
            Gbetrag = 10
            Selection.NumberFormat = "#,##0.00 $"
            ActiveCell.FormulaR1C1 = Gbetrag
            
''''' Danach weiter bei Artikel Einfuegen ''''''''''''''''''''''''''''''''''''
            Restbetrag
        End If
        
End Sub
Sub addTwenty()
''''' Pruefen ob Felder "Name" & "Betrag fuer Einkauf" ausgefuellt sind ''''''
        If (Worksheets("Abrechnung").Range(NAME) = "") Then
            Meldung = MsgBox _
            ("Bitte zuerst ""Abrechnung Starten"" und ""Name""" & _
            "eingeben! ", vbOKOnly + 16, "FEHLER!")
        Else
        
'''''''''''''' Curser auf Zelle D2 setzen ''''''''''''''''''''''''''''''''''''
                Worksheets("Abrechnung").Activate
                Worksheets("Abrechnung").Range(BETRAG).Activate
                ActiveCell.Select
                
''''''' Betrag auf 20 setzenund einfuegen ''''''''''''''''''''''''''''''''''''
                Gbetrag = 20
                Selection.NumberFormat = "#,##0.00 $"
                ActiveCell.FormulaR1C1 = Gbetrag
                Frame
                
''''' Danach weiter bei Artikel Einfuegen ''''''''''''''''''''''''''''''''''''
                Restbetrag
        End If
        
End Sub
Sub addFifty()
''''' Pruefen ob Felder "Name" & "Betrag fuer Einkauf" ausgefuellt sind ''''''
        If (Worksheets("Abrechnung").Range(NAME) = "") Then
            Meldung = MsgBox _
            ("Bitte zuerst ""Abrechnung Starten"" und ""Name""" & _
            "eingeben! ", vbOKOnly + 16, "FEHLER!")
        Else
        
'''''''''''''' Curser auf Zelle D2 setzen ''''''''''''''''''''''''''''''''''''
            Worksheets("Abrechnung").Activate
            Worksheets("Abrechnung").Range(BETRAG).Activate
            ActiveCell.Select
            
''''''' Betrag auf 50 setzenund einfuegen ''''''''''''''''''''''''''''''''''''
            Gbetrag = 50
            Selection.NumberFormat = "#,##0.00 $"
            ActiveCell.FormulaR1C1 = Gbetrag
            Frame
            
''''' Danach weiter bei Artikel Einfuegen ''''''''''''''''''''''''''''''''''''
            Restbetrag
        End If
        
End Sub
Sub addHundred()
''''' Pruefen ob Felder "Name" & "Betrag fuer Einkauf" ausgefuellt sind ''''''
        If (Worksheets("Abrechnung").Range(NAME) = "") Then
            Meldung = MsgBox _
            ("Bitte zuerst ""Abrechnung Starten"" und ""Name""" & _
            "eingeben! ", vbOKOnly + 16, "FEHLER!")
        Else
        
'''''''''''''' Curser auf Zelle D2 setzen ''''''''''''''''''''''''''''''''''''
            Worksheets("Abrechnung").Activate
            Worksheets("Abrechnung").Range(BETRAG).Activate
            ActiveCell.Select
            
''''''' Betrag auf 100 setzenund einfuegen '''''''''''''''''''''''''''''''''''
            Gbetrag = 100
            Selection.NumberFormat = "#,##0.00 $"
            ActiveCell.FormulaR1C1 = Gbetrag
            Frame
            
''''' Danach weiter bei Artikel Einfuegen ''''''''''''''''''''''''''''''''''''
            Restbetrag
        End If
        
End Sub
Sub addtwohundred()
''''' Pruefen ob Felder "Name" & "Betrag fuer Einkauf" ausgefuellt sind ''''''
        If (Worksheets("Abrechnung").Range(NAME) = "") Then
            Meldung = MsgBox _
            ("Bitte zuerst ""Abrechnung Starten"" und ""Name""" & _
            "eingeben! ", vbOKOnly + 16, "FEHLER!")
        Else
        
'''''''''''''' Curser auf Zelle D2 setzen ''''''''''''''''''''''''''''''''''''
            Worksheets("Abrechnung").Activate
            Worksheets("Abrechnung").Range(BETRAG).Activate
            ActiveCell.Select
            
''''''' Betrag auf 200 setzenund einfuegen '''''''''''''''''''''''''''''''''''
            Gbetrag = 200
            Selection.NumberFormat = "#,##0.00 $"
            ActiveCell.FormulaR1C1 = Gbetrag
            Frame
            
''''' Danach weiter bei Artikel Einfuegen ''''''''''''''''''''''''''''''''''''
            Restbetrag
        End If
        
End Sub
Sub addfivehundred()
''''' Pruefen ob Felder "Name" & "Betrag fuer Einkauf" ausgefuellt sind ''''''
        If (Worksheets("Abrechnung").Range(NAME) = "") Then
            Meldung = MsgBox _
            ("Bitte zuerst ""Abrechnung Starten"" und ""Name""" & _
            "eingeben! ", vbOKOnly + 16, "FEHLER!")
        Else
        
'''''''''''''' Curser auf Zelle D2 setzen ''''''''''''''''''''''''''''''''''''
            Worksheets("Abrechnung").Activate
            Worksheets("Abrechnung").Range(BETRAG).Activate
            ActiveCell.Select
            
''''''' Betrag auf 500 setzen und einfuegen ''''''''''''''''''''''''''''''''''
            Gbetrag = 500
            Selection.NumberFormat = "#,##0.00 $"
            ActiveCell.FormulaR1C1 = Gbetrag
            Frame
            
''''' Danach weiter bei Artikel Einfuegen ''''''''''''''''''''''''''''''''''''
            Restbetrag
        End If
        
End Sub




///////////////////////////////////////////////////////////////////////////////




''''''''''''''''''' Extra Funktionen ''''''''''''''''''''''''''''''''''''''''


Public Function Pfand()

'''''' Abfrage ob Pfand abgegeben wurde oder nicht '''''''''''''''''''''''''''
'''''' Frage nach Pfandbetrag pro Flasche '''''''''''''''''''''''''''''''''''
''''''' Frage nach Flaschen Menge '''''''''''''''''''''''''''''''''''''''''''
''''''' Frage nach weiter Eingabe ''''''''''''''''''''''''''''''''''''''''''

End Function



////// Modul 3 \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


'*****************************************************************************
'
'              Modul mit Funktioen zum Ausfüllen der Übersicht
'             ==================================================
'
'       - CursorA1(): setzt den Corsor auf Zelle A1 und sucht die nächste frei Zelle
'           in der Spalte A
'       - ZeilenSuche(): Code für die Funktion für die Suche der
'           nächsten freien Zeile in der Splate "Name".
'       - UebersichtDrucken(): Abfrage ob Arbeitsblatt "Übersicht" gedruckt werden soll.
'           wird "nein" ausgewählt, dann spring das Programm zurück zum Arbeitsblatt "Abrechnung"
'
'
'
'
'
'
'
'
'*****************************************************************************

Sub CursorA1()
'''''''''''''''''''''' Cursor auf leere Zelle setzen '''''''''''''''''''''''''
        Worksheets("Übersicht").Activate
        Worksheets("Übersicht").Range("A1").Activate
        ActiveCell.Select
''''''' solange aktive Zelle nicht leer ist '''''''''''''''''''''''''''''''''
            While (Not (ActiveCell.VALUE = ""))
                 ActiveCell.Offset(rowOffset:=1).Activate
            Wend
End Sub


Public Function ZeilenSuche() As Integer
''''''''' Cursor auf A1 setzten ''''''''''''''''''''''''''''''''''''''''''''
            CursorA1
'''''''' leere Zellen in Splate Menge suchen ''''''''''''''''''''''''''''''''
'''' Laufzeit Variable deklarieren ''''''''''''''''''''''''''''''''''''''''''
                Dim i As Integer
'''' Laufzeit Variable initialisieren '''''''''''''''''''''''''''''''''''''''
                i = 1
                Worksheets("Übersicht").Activate
                Worksheets("Übersicht").Range("A1").Activate
                ActiveCell.Select
                    While (Not (ActiveCell.VALUE = "Flaschen abgegeben:"))
                        i = i + 1
                        ActiveCell.Offset(rowOffset:=1).Activate
                    Wend
                ZeilenSuche = i
End Function

Sub UebersichtDrucken()
'''' Funktion Übersicht Drucken  '''''''''''''''''''''''''''''''''''''''''''''''
        Dim pr As String
        Sheets("Übersicht").Select
        pr = "$A$1:$H$" & Modul3.ZeilenSuche() + 1
        PrintE = "Übersicht Drucken?"
        PrintT = "Übersicht Drucken"
        StileP = vbYesNo + vbQuestion + vbDefaultButton1
        AntwortP = MsgBox(PrintE, StileP, PrintT)
        
            If AntwortP = 6 Then
                ActiveSheet.PageSetup.PrintArea = pr
                ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
                IgnorePrintAreas:=False
                Worksheets("Abrechnung").Activate
            Else
                Worksheets("Abrechnung").Activate
        End If
End Sub


