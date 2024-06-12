Attribute VB_Name = "Modul1"
'*****************************************************************************
' Filename: EinkaufAbrechung.xlsm
' Author: Marcello Jorizzo
' Language: German
' Created:        August 07, 2021
' Last Modified:  July 12, 2024
'******************************************************************************
'
'               Programm fuer die Abrechung vom Einkaeufen
'       ======================================================
'
'   Es wird eine leere Liste Erzeugt.
'   Zuerst muss der Name eingegeben werden;
'   anschlieﬂend der Betrag der fuer den Einkauf zur
'   Verfuehgung steht.
'   Das Programm sucht die naechste leere Zeile. Dann muss die Artikel Menge
'   eingegben werden. Dann der Stueck preis des Artikels. Danach wird
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
'   Pname   String  ein Name
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
        
'''''''''''''''''''''''''' Name  eingeben ''''''''''''''''''''''''''''''''''''
'''''''''''' Curser auf Zelle F2 setzen und aktive Zelle weiss faerben '''''''
        Worksheets("Abrechnung").Activate
        Worksheets("Abrechnung").Range(NAME).Activate
        ActiveCell.Select
        Modul2.WhiteCell
        
'''''''''''''''' NameP Eingabe Meldung: Name eingeben ''''''''''''''''''''''''
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
       BetragE = "Geldbetrag fuer Einkauf eingeben"
        
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





