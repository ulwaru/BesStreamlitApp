Attribute VB_Name = "T5"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit 'T5

'Beep
    Private Declare Function Beep Lib "kernel32" (Optional ByVal dwFreq As Long = 900, Optional ByVal dwDuration As Long = 100) As Long

Sub zzz_T5()
    'showProcs "char"
    
    WriteToCell_StrichVertical "T5", 3, 28
    EE 1: Beep
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Called from    [UserClick to select another T5-Cell]
        
    'Vorab (bei jedem Klick auf eine andere T5-Zelle)
        Dim s%, z%, sLast&
        z = Target.Row: s = Target.Column: DoArr: sLast = Get_T5_sLast(): Beep1
        
'        T5_Kill_EmptyEndLines sLast
'        beep 666, 99 'Kontroll-Beep für SelectionChangeEvent-Aufruf
'    'Eintrag Jhg. falls GebDatum vorhanden
'        If s = 16 And z > 7 Then T5_AddJhgFromGebDat z, s
'    'CtrlKeyWasDown
'        If CtrlKeyWasDown Then
'            T5_CtrlKeyWasDownAt Target.Address: Exit Sub
'        End If
        
    'Open Folder
        If s = 18 And z > 7 And Cells(z, s) = "1" Then OpenFolder ArrC(4) + "\" + Cells(z, 10).Value
    'Change Name (global)
        Dim Symb$: Symb = Application.WorksheetFunction.Unichar(128260) 'Synchro-Symbol
        If s = 19 And z > 7 And Cells(z, s) = Symb Then Jump_FromT5_toT2_ChangeNameGlobal z
    'Search for
        If z = 4 And s = 10 Then Exit Sub 'Input "Suche" ermöglichen
        If z = 5 And s = 10 Then T5_SearchFor Range("J4").Value: Exit Sub
    'Select A1
        If z < 6 And s < 23 Then EE 0: [A1].Select: EE 1
    'Sort
        If (z = 6 And s > 2 And s < 18) Or (z = 6 And s = 22) Then T5_Sort s: Exit Sub
End Sub

Private Sub Worksheet_Activate() 'Sobald dieses Blatt aktiviert/angezeigt wird
    'ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
    T5_PepUp
    T5_Simulate_2ArrowsDown_WithCells
End Sub

Sub T5_PepUp()
    'Called from    Worksheet_Activate
    
    'Vorbereitung
        Dim i%, zLast%, r As Range, ArrT5(), Symb$
        zLast = Get_NrOfLastRowInColumnNr(3, "T5"): With Sheets("T5")
    'Fill ArrT5
        ArrT5 = Range(Cells(8, 3), Cells(zLast, 15))
    'Hintergrund einfärben
        For i = 8 To zLast
            If CStr(ArrT5(i - 7, 3)) = "w" Then .Range(.Cells(i, 3), .Cells(i, 20)).Interior.Color = 14083324: .Range(.Cells(i, 13), .Cells(i, 15)).Interior.Color = 14083324
            If CStr(ArrT5(i - 7, 3)) = "m" Then .Range(.Cells(i, 3), .Cells(i, 20)).Interior.Color = 15652797: .Range(.Cells(i, 13), .Cells(i, 15)).Interior.Color = 15652797
            .Cells(i, 2) = i - 7
        Next
    'Rahmen - each CellBorder: white
        With .Range(.Cells(8, 2), .Cells(zLast, 15)).Borders
            .LineStyle = xlContinuous: .Color = RGB(250, 250, 250): .Weight = xlThin
        End With
    'Spalte 18 (Personenordner öffnen)
        Set r = Range(Cells(8, 18), Cells(zLast, 18))
        r.Font.NAME = "Wingdings": r.Font.size = 11: r.HorizontalAlignment = xlCenter:
        r.VerticalAlignment = xlCenter: r.Value = 1
        Columns(18).ColumnWidth = 3.7
    'Spalte 19 (Namen ändern; global)
        Set r = Range(Cells(8, 19), Cells(zLast, 19))
        Symb = Application.WorksheetFunction.Unichar(128260) 'Synchro-Symbol
        'With r: .Value = Symb: .Font.NAME = "Segoe UI Emoji": .Font.size = 12: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: End With
        r.Value = Symb: r.Font.NAME = "Segoe UI Emoji": r.Font.size = 12: r.HorizontalAlignment = xlCenter: r.VerticalAlignment = xlCenter
        Columns(19).ColumnWidth = 2.86
    'Finals
        End With
End Sub

Sub T5_AddJhgFromGebDat(z%, s%)
    'Called from    Worksheet_SelectionChange
    's              = 16
    
    'Vorbereitung
        Dim G$, j$
        With Sheets("T5")
    'Abfragen
        G = .Cells(16, z - 1): j = .Cells(6, z - 1)
        If G Like "########" And j = "" Then .Cells(6, z - 1) = Left(G, 4)
    
    
    
    'Finals
        End With
End Sub


Sub Write_PdSmall()
    'Called from    [none]
    'Action         schreibt PdSmall mit LeadingText in PersonData.txt
    
    'Vorbereitung
        Dim AnzNamen$, PD$: DoArrc
    'Write
        PD = Get_PdSmall_FromT5 'Sort_TextLines is done
        PD = Add_LeadingText_ToPdSmall(PD)
        AnzNamen = CStr(anzAinB(vbCrLf, PD) - 4)
        PD = Replace(PD, "repräsentierten Personen", "repräsentierten Personen (z. Zt. " + AnzNamen + ")")
        writeStringToFile ArrC(2) + "\PersonData.txt", PD
    'Log
        LogBuch "PersonData.txt wurde mit den T5-Daten neu geschrieben"
End Sub

Sub Write_PersonDataTxt_Give_PdFull(Pd9$)
    'Status     Einzige Prozedur, welche die Datei 'PersonData.txt' schreibt
    
    Dim AnzNamen$, p1$, p2$, Pd2$
    Pd2 = Pd9
    AnzNamen = CStr(anzAinB(vbCrLf, Pd2) - 6)
    Pd2 = Replace(Pd2, "repräsentierten Personen", "repräsentierten Personen (z. Zt. " + AnzNamen + ")")

    'Erstellung der neuen Datei 'PersonData.txt'
         p1 = ArrC(2) + "\z PersonData.txt"
         writeStringToFile p1, Pd2
    'Erstellung einer SicherungsDatei 'PersonData.txt'
         p2 = ArrC(1) + "\prog\old\PersonData\PersonData " + Format(Now(), "yyyymmdd_hhmmss") + ".txt"
         writeStringToFile p2, Pd2
End Sub

Function Add_LeadingText_ToPdSmall(PD$) As String
    'Called from    Write_PdSmall
    
    'Vorbereitung
        Dim Header$, s$, v$, C1%, C2%
        Header = "|Nachname|Vorname|m/w|Jhg.|Verein|Landesturnverband|Nat|Ordnername in 'Leute'|"
        v = vbCrLf: s = Header + v + PD: s = TAB_Simulation(s)
    'Linie oberhalb Header
        C1 = InStr(1, s, v) - 1         'c1 = letztes Zeichen des Headers = Breite Tabelle
        s = "PersonData.txt  -  Daten zu den im Archiv repräsentierten Personen" + v + v _
            + String(C1, "-") + v + s
    'Linie unterhalb Headers            'c2 = Position des letzten Header-Zeichens
        C2 = InStr(1, s, "m/w"): C2 = InStr(C2, s, v) - 1
        s = Left(s, C2) + v + String(C1, "-") + Mid(s, C2 + 1)
    'Finals
        Add_LeadingText_ToPdSmall = s
End Function

Function Get_PdSmall_FromT5() As String
    'Called from    Write_PdSmall
    'Scope          Spalten Nachname bis Personenordner
    'Action         Erstellung des Strings Pd (aus T5-Tabelle)

    'Vorbereitung
        Dim PD$, v$, i%, j%, sLast%, zLast%, ArrT5()
        zLast = Get_NrOfLastRowInColumnNr(3, "T5"):  v = vbCrLf
        sLast = 10
    'Fill ArrT5
        ArrT5 = Range(Cells(8, 3), Cells(zLast, sLast))
    'Action
        For i = 1 To UBound(ArrT5, 1)
            For j = 1 To UBound(ArrT5, 2)
                PD = PD + "|" + Trim(ArrT5(i, j))
            Next
            PD = PD + "|" + v
        Next
        PD = Sort_TextLines(PD)
        PD = Delete_EmptyRowsInString(PD)
    'Finals
        'show Pd
        Get_PdSmall_FromT5 = PD
End Function

Function Get_PdBig_FromT5() As String
    'Called from    [none]
    'Scope          Spalten Nachname bis Label
    'Action         Erstellung des Strings Pd (aus T5-Tabelle)

    'Vorbereitung
        Dim PD$, v$, i%, j%, ArrT5()
        v = vbCrLf
    'Fill ArrT5
        ArrT5 = Range(Cells(Get_T5_zFirst, Get_T5_sFirst), Cells(Get_T5_zLast, Get_T5_sLast))
    'Action
        For i = 1 To UBound(ArrT5, 1)
            For j = 1 To UBound(ArrT5, 2)
                PD = PD + "|" + Trim(ArrT5(i, j))
            Next
            PD = PD + "|" + v
        Next
        PD = Sort_TextLines(PD)
        PD = Delete_EmptyRowsInString(PD)
    'Finals
        'show Pd
        Get_PdBig_FromT5 = PD
End Function

Sub T5_SearchFor(s$)
    Dim i%, zLast%, Arr1()
    [A1].Select
    zLast = Get_T5_zLast
    'Spalte Nachname
        Arr1 = Range(Cells(1, 3), Cells(zLast, 3))
    For i = Get_T5_zFirst To zLast
        If LCase(Arr1(i, 1)) Like LCase(s) + "*" Then Range(Cells(i, 2), Cells(i, 23)).Select: Exit For
    Next
End Sub

Sub T5_Kill_EmptyEndLines(sLast&)
    'Called from    Worksheet_SelectionChange[T5]
    'Action         löscht Zeilen am Ende der T5-PersonData-Liste
    
    'Vorbereitung
        Dim zLast&
        zLast = Get_NrOfLastRowInColumnNr(3, "T5")
    'CellColor
        If Cells(zLast + 1, 3).Interior.ColorIndex <> xlNone Then
            Range(Cells(zLast + 1, 2), Cells(zLast + 10, sLast + 10)).Clear
        End If
End Sub

Function CtrlKeyWasDown() As Boolean
    Dim iResult%
    iResult = GetKeyState(VK_CONTROL) And &H80
    If iResult = 128 Then CtrlKeyWasDown = True
End Function

Sub Fill_T5EmptyCellsInBigPart()
    'Called from    [none]
    
    Dim s%, z%
    With Sheets("T5")
    For z = 8 To Get_T5_zLast
    For s = 12 To 21
        'If .Cells(z, s) = "" Then .Cells(z, s) = "-"
        If .Cells(z, s) = "-" Then .Cells(z, s).HorizontalAlignment = xlCenter Else .Cells(z, s).HorizontalAlignment = xlLeft
    Next
    Next
    End With
    Beep
End Sub

Sub T5_Simulate_2ArrowsDown_WithCells()
    WriteToCell_StrichVertical "T5", 3, 18: WriteToCell_StrichVertical "T5", 4, 18
    WriteToCell_StrichVertical "T5", 5, 18: WriteToCell_StrichVertical "T5", 6, 18
    WriteToCell_PfeilNachUnten "T5", 7, 18
    WriteToCell_StrichVertical "T5", 4, 19
    WriteToCell_StrichVertical "T5", 5, 19
    WriteToCell_StrichVertical "T5", 6, 19
    WriteToCell_PfeilNachUnten "T5", 7, 19
End Sub

Sub WriteToCell_PfeilNachUnten(NameOfSheet$, z&, s&)
    Dim r As Range
    Set r = Sheets(NameOfSheet).Cells(z, s): EE 0
    r.Value = "â": r.HorizontalAlignment = xlCenter: r.VerticalAlignment = xlBottom
    With r.Font: .NAME = "Wingdings": .size = 14: .Color = RGB(155, 155, 155): End With: EE 1
End Sub

Sub WriteToCell_StrichVertical(NameOfSheet$, z&, s&)
    Dim r As Range
    Set r = Sheets(NameOfSheet).Cells(z, s): EE 0
    r.Value = "|": r.HorizontalAlignment = xlCenter: r.VerticalAlignment = xlCenter
    With r.Font: .NAME = "Calibri": .size = 20: .Color = RGB(155, 155, 155): End With: EE 1
End Sub



