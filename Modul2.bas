Attribute VB_Name = "Modul2"
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


