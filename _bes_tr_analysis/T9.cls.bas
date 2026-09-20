Attribute VB_Name = "T9"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Sub zzz_M_Check()
    showProcs "textpart"
    
    '[Q9] = Chr(158)
    'RenameModule "M19", "M_T9"
    '[c10].Interior.Color = Green3
    
    EE 1: Beep
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
        Dim s%, z%, zLast&, r As Range
        z = Target.Row: s = Target.Column: zLast = Get_NrOfLastRowInColumnNr(2, "T9")
        With Sheets("T9"): Beep1
    'Farbige Leerzeilen am Tabellenende löschen
        Set r = .Range(.Cells(zLast + 1, 2), .Cells(zLast + 10, 16)): r.Clear
    'YearRahmen zurücksetzen
        Set r = .Range(.Cells(9, 2), .Cells(zLast, 3)): End With
        With r.Borders: .Color = vbWhite: .LineStyle = xlContinuous: .Weight = xlThin
        End With
    'YearRahmen um Date-Zellen gleicher Jahreszahl
        If z > 8 And s > 1 And s < 17 Then T9_RahmenUmDateZellen z 'kein Exit Sub
    's3Date ergänzen
        T9_LookFor_s3Date_toFill zLast ': Exit Sub
        If ArrC(79) = "1" Then FillArrC 79, "0": Exit Sub
    'Sort
        If z = 7 And s > 1 And s < 16 Then T9_Sort s: Exit Sub
    'CellButton OpenFolder
        If z > 8 And s = 13 Then T9_OpenOneFolder z, s: Exit Sub
    'Knollen in Spalte 1 setzen (bei Klick auf Spalte 2 oder 3
        [A:A] = "": If z > 8 And z <= zLast And (s = 2 Or s = 3) Then T9_Write_KnollenInSpalte1 (z): Exit Sub
    'Knollen-Klick auswerten
        If z > 8 And s = 1 Then T9_KnollenKlick_Auswerten (z): Exit Sub
    'Knollen in Spalte 1 löschen
        If Not (z > 8 And (s = 2 Or s = 3)) Then [A:A] = "": T9_Write_B4: Exit Sub
End Sub

Private Sub Worksheet_Activate() 'Sobald dieses Blatt aktiviert/angezeigt wird
    T9_FolderIcon_InSpalteExist_OnlyIfFolderExists
    T9_Write_NewEventFolders_ToT9
    T9_PepUp_T9
    'beep
End Sub

Sub T9_PepUp_T9()
    'Called from    T9_Add_EventFoldersToTerminalender
    
    'Vorbereitung
        Dim c%, zLast&, T(), OneCol(), r As Range
        With Sheets("T9"): EE 0
    'sort
        'T9_Sort 4: T9_Sort 2
    'Load
        zLast = Get_NrOfLastRowInColumnNr(3, "T9")
'        T9_Load_Terminkalender T
    'Green1 (Spalten 2,4,6,8,10,12,14,16), Green2 (Spalten 3,5,7,9,11,13,15)
        Set r = Union(.Range(Cells(13, 2), Cells(zLast, 2)), .Range(Cells(13, 4), Cells(zLast, 4)), .Range(Cells(13, 6), Cells(zLast, 6)), .Range(Cells(13, 8), Cells(zLast, 8)), .Range(Cells(13, 10), Cells(zLast, 10)), .Range(Cells(13, 12), Cells(zLast, 12)), .Range(Cells(13, 14), Cells(zLast, 14)), .Range(Cells(13, 16), Cells(zLast, 16)))
        r.Interior.Color = Green1
        Set r = Union(.Range(Cells(13, 3), Cells(zLast, 3)), .Range(Cells(13, 5), Cells(zLast, 5)), .Range(Cells(13, 7), Cells(zLast, 7)), .Range(Cells(13, 9), Cells(zLast, 9)), .Range(Cells(13, 11), Cells(zLast, 11)), .Range(Cells(13, 13), Cells(zLast, 13)), .Range(Cells(13, 15), Cells(zLast, 15)))
        r.Interior.Color = Green2
    'Green3
        T9_Set_s2Date_to_Green3_ifFolderExists
    'All
        Set r = .Range(Cells(9, 2), Cells(zLast, 16))
        r.Borders.Color = vbWhite
    'T9-Spalte 2
        Set r = .Range(Cells(9, 2), Cells(zLast, 2))
        r.Font.size = 8: r.HorizontalAlignment = xlCenter
    'T9-Spalte 13  'exist-Spalte auf Wingdings setzen (Ordnersymbol)
        Set r = .Range(.Cells(9, 13), .Cells(zLast, 13))
        r.Font.NAME = "Wingdings": r.Font.size = 12: r.HorizontalAlignment = xlCenter
    'T9-Spalte 16  'SchlussBalken "|"
        Set r = .Range(Cells(9, 16), Cells(zLast, 16))
        r.Value = "|": r.HorizontalAlignment = xlLeft
    'Center     'Date10 'Nation
        Set r = Union(.Range(Cells(9, 3), Cells(zLast, 3)), .Range(Cells(9, 7), Cells(zLast, 7)))
        r.HorizontalAlignment = xlCenter: r.Font.size = 11
    'RowHeight
        .Rows("1:" + CStr(zLast)).RowHeight = 15
    'Below Last
        Set r = .Range(Cells(zLast + 1, 2), Cells(zLast + 300, 12))
        r.Value = "": r.Interior.ColorIndex = 0: r.Borders.ColorIndex = xlNone
    'T9-[B4] Anzeige
        T9_Write_B4
    'Finals
        End With: EE 1
End Sub

Sub T9_UpdateTerminalender()
    'Vorbereitung
        Dim i%, j%, t1(), T2(), T3()
        T9_Load_Terminkalender t1
        T9_Load_DetailsOfSomeDateEventFolders T2
    'Add "fromEventFolder" to "Bemerkung"
        For i = 1 To UBound(T2, 1)
            T2(i, 7) = "fromEventFolder"
        Next
    'T3()
        ReDim T3(1 To UBound(t1, 1) + UBound(T2, 1), 1 To 10)
        For i = 1 To UBound(t1, 1)
            For j = 1 To UBound(t1, 2)
                T3(i, j) = CStr(t1(i, j))
            Next
        Next
        For i = 1 To UBound(T2, 1)
            For j = 1 To UBound(T2, 2)
                T3(UBound(t1, 1) + i, j) = CStr(T2(i, j))
            Next
        Next
    'sort
        QuickSort2D T3, , , 3
        QuickSort2D T3, , , 1
    'Paste
        xxx
    showArray2D T3
End Sub

Sub T9_Sort(s%)
    'Called from    Worksheet_SelectionChange[T9]
    's              SpaltenNr, nach der sortiert werden soll 's = 2 (Datum8), ...
    'Pfeil Up       "ñ" (Zeichensatz Wingdings)
    'Pfeil Down     "ò" (Zeichensatz Wingdings)
    
    'Vorbereitung
        Dim sL$, Sortierbereich$, SortSpNew$, SortStOld$, SortSpOld%, zLast%
        zLast = Get_NrOfLastRowInColumnNr(3): EE 0
    'SpaltenNrLastSort 'StatusLastUpDown
        SortSpOld = Get_ColumnNr_HoldingMyTextPartInRowX("T9", 8, 1, "ñ") 'Pfeil Up
        If SortSpOld = 0 Then
            SortSpOld = Get_ColumnNr_HoldingMyTextPartInRowX("T9", 8, 1, "ò") 'Pfeil Down
            SortStOld = "D"                        'T9 Status LastSort, D=Down
        Else
            SortStOld = "U"                        'T9 Status LastSort, U=Up
        End If
    'SortierSpalteNew festlegen mittels oberster Zelle
        sL = Get_ColumnLetter(s)
        SortSpNew = sL + "9"
    'Sortierbereich
        Sortierbereich = "B9:P" + CStr(zLast)
    'Zeile Up-/DownPfeil leeren (unterhalb des Tabellenkopfes)
        Range("B8:P8") = ""
    'sort
        If CInt(SortSpOld) = s Then
            'Gleiche Sortierspalte wurde erneut gedrückt
            If SortStOld = "U" Then
                T9_SortDown Sortierbereich, SortSpNew:  Cells(8, s) = "ò"  'Down
            Else
                T9_SortUp Sortierbereich, SortSpNew:    Cells(8, s) = "ñ"  'Up
            End If
        Else
            T9_SortUp Sortierbereich, SortSpNew:        Cells(8, s) = "ñ"  'Up
        End If
    'Finals
        [a6].Select: T9_Write_B4: EE 1
End Sub

Sub T9_Write_B4()
    'Vorbereitung
        Dim c%, zLast&, OneCol(), r As Range
        zLast = Get_NrOfLastRowInColumnNr(3, "T9"): With ThisWorkbook.Worksheets("T9")
    'B4 Anzeige
        Set r = .Range(.Cells(9, 13), .Cells(zLast, 13)): OneCol = r.Value
        c = Count_CellsLikeMyStringInArray2D(OneCol, "*1*")
        [B4] = CStr(zLast - 8) + " Termine (zu " + CStr(c) + " Terminen existiert ein Event-Ordner)"
    End With
End Sub

Sub T9_SortUp(Sortierbereich$, Sortierspalte$)
    With ThisWorkbook.Worksheets("T9")
        .Range(Sortierbereich).Sort Key1:=.Range(Sortierspalte), Order1:=xlAscending, _
        Header:=xlNo, orderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    End With
End Sub

Sub T9_SortDown(Sortierbereich$, Sortierspalte$)
    With ThisWorkbook.Worksheets("T9")
        .Range(Sortierbereich).Sort Key1:=.Range(Sortierspalte), Order1:=xlDescending, _
        Header:=xlNo, orderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    End With
End Sub

Sub T9_Load_Terminkalender(ByRef T)
    'Vorbereitung
        Dim zLast&
        zLast = Get_NrOfLastRowInColumnNr(3, "T9"): With Sheets("T9"): EE 0
        T = .Range(.Cells(13, 2), .Cells(zLast, 11)).Value
    'Finals
        EE 1: End With
End Sub

Sub T9_repairBuli()
    'ist erledigt
    Dim i%, T()
    T9_Load_Terminkalender T
    With Sheets("T9"): EE 0
    For i = 1 To UBound(T, 1)
        If T(i, 3) Like "*Buli*" Then
            If T(i, 7) Like "*-*" Then
                T(i, 3) = T(i, 3) + " " + T(i, 7): T(i, 7) = ""
                'Cells(i + 12, 4).Select
                Cells(i + 12, 4) = T(i, 3)
                Cells(i + 12, 8) = ""
            End If
        End If
    Next
    EE 1: End With
    showArray2D T
End Sub

Sub T9_Replace_String1_String2_InAllTkCellsOfTkColumnX(s1$, s2$, TkColumn%)
    'Called from    [None]
    'Aufruf: z.B.   T9_Replace_String1_String2_InAllTkCellsOfTkColumnX "fromEventFolder", "", 7
    
    'Vorbereitung
        Dim c%, i%, zLast&, OneCol(), r As Range, OK As Boolean
        zLast = Get_NrOfLastRowInColumnNr(3, "T9"): With Sheets("T9"): EE 0
    'Load OneCol
        Set r = .Range(.Cells(13, TkColumn + 1), .Cells(zLast, TkColumn + 1))
        OneCol = r.Value
    'Replace
        For i = 1 To UBound(OneCol, 1)
            c = InStr(1, OneCol(i, 1), s1)
            If c > 0 Then OneCol(i, 1) = Replace(OneCol(i, 1), s1, s2): OK = True
        Next
    'Paste
        If OK Then Paste_2DArrayToCell_z_s "T9", 13, TkColumn + 1, OneCol
    'Finals
        End With: EE 1
End Sub

Sub T9_Add_EventFoldersToTerminalender()
    'Called from    [None]
    'Status         Sub wurde bei Ersteinrichtung des Terminkalenders verwendet
    'Action         fügt pro EventOrdner eine Zeile zum Terminkalender hinzu,
    '               auch, wenn der Termin im Terminkalender bereits enthalten ist
    
    'Vorbereitung
        Dim i%, zLast&, T()
        zLast = Get_NrOfLastRowInColumnNr(3, "T9"): EE 0
    'Load T(); im T9-Raster
        T9_Load_DetailsOfSomeDateEventFolders T
    'Add "fromEventFolder" to "Bemerkung"
        For i = 1 To UBound(T, 1)
            T(i, 7) = "fromEventFolder"
        Next
    'Paste
        Paste_2DArrayToCell_z_s "T9", zLast + 5, 2, T
    'PepUp
        T9_PepUp_T9
    EE 1: Beep
End Sub

Sub T9_OpenOneFolder(z%, s%)
    'Called from    [UserClick onto FolderSymbol], Worksheet_SelectionChange[T9]
    
    With Sheets("T9"): EE 0
    If .Cells(z, s) <> "1" Then Exit Sub
    OpenFolder .Cells(z, s + 2): End With: EE 1
End Sub

Sub T9_FolderIcon_InSpalteExist_OnlyIfFolderExists()
    'Called from    Worksheet_Activate

    Dim i&, zLast&
    With Sheets("T9"): EE 0: Screen 0
    zLast = Get_NrOfLastRowInColumnNr(3, "T9")
    'Aktion auf Blatt T9, Zeile für Zeile
        For i = 9 To zLast
            '.Cells(i, 13).Select
            If .Cells(i, 15) = "" Then
                .Cells(i, 14) = ""
                .Cells(i, 13) = Chr(158)        'kleiner Knollen
            Else
                If FolderExists(.Cells(i, 15)) Then
                    .Cells(i, 13) = "1"         'OrdnerSymbol
                Else
                    .Cells(i, 13) = Chr(158)    'kleiner Knollen
                    .Cells(i, 14) = ""
                    .Cells(i, 15) = ""
                End If
            End If
        Next
    'Set r = Range(Cells(9, s), Cells(zLast, s))
    End With: EE 1: Screen 1
End Sub

Function T9_Get_PathsOfEventFolders_NewToT9() As String
    'Called from    T9_Write_NewEventFolders_ToT9
    'Status         alle T9-Ordner (T9-Spalte 15) sind SubFolder von 'Events'
    
    'Vorbereitung
        Dim D$, EvNotInT9$, L$, T$, T9NotInEv$, v$
        Dim i%, zLast&, ArrD() As String, ArrT() As String
        zLast = Get_NrOfLastRowInColumnNr(3, "T9"): v = vbCrLf
    'D = Pfade aller Ordner in 'Events'
        D = v + Get_PathsOfAllDateEventFolders
        'show D 'F:\Archiv Trampolin 1900-1999\Events\1959-07-09 Turnfest Basel_CH
        ArrD = Split(D, v): D = D + v
    'T = Pfade aller Ordner in T9
        Load_Array1D_OfColumnPart_NoEmptyRows ArrT, "T9", 9, CInt(zLast), 15
        'showArray ArrT '1. T9-Pfad in ArrT(1)
        T = Join(ArrT, v) + v 'show T
    'D-Search
        For i = 1 To UBound(ArrD)
            'Schleife über alle Pfade in 'Events'
            L = ArrD(i)          'L = one Event-Subfolder-Path
            If Not T Like "*" + v + L + v + "*" Then
                'Der Event-Ordner L ist kein T9-Folder
                EvNotInT9 = EvNotInT9 + v + L
            End If
        Next
        If Left(EvNotInT9, 2) = v Then EvNotInT9 = Mid(EvNotInT9, 3)
        T9_Get_PathsOfEventFolders_NewToT9 = EvNotInT9
End Function

Function T9_Get_NationFromEventName(N$) As String
    T9_Get_NationFromEventName = "D"
    If Not N Like "*_*" Then Exit Function
    If N Like "*_[A-Z]" Then T9_Get_NationFromEventName = Right(N, 1): Exit Function
    If N Like "*_[A-Z][A-Z]" Then T9_Get_NationFromEventName = Right(N, 2): Exit Function
    If N Like "*_[A-Z][A-Z][A-Z]" Then T9_Get_NationFromEventName = Right(N, 3): Exit Function
End Function

Function T9_Get_EventNameNoNation(N$, Nat$) As String
    Dim c%: T9_Get_EventNameNoNation = N: c = Len(Nat) + 1
    If Right(N, c) = "_" + Nat Then T9_Get_EventNameNoNation = Left(N, Len(N) - c)
End Function

Sub T9_Write_NewEventFolders_ToT9()
    'Called From    Worksheet_Activate
    
    'Vorbereitung
        Dim AllPaths$, Nat$, N1$, N2$, p$, r1 As Range, r2 As Range
        Dim s2Date$, s3Date$, s4Event$, s5Ort$, i&, zLast&, A() As String
        With Sheets("T9"): zLast = Get_NrOfLastRowInColumnNr(3, "T9")
    'Range Zeile 9 (für Format übertragen)
        Set r1 = .Range(.Cells(9, 2), .Cells(9, 16))
    'Alle Pfade von EventOrdnern NewToT9 laden
        AllPaths = T9_Get_PathsOfEventFolders_NewToT9
        'show AllPaths 'F:\Archiv Trampolin 1900-1999\Events\1959-07-14 Intern. Lehrgang Freiburg
    'Jede AllPaths-Zeile soll in T9-Spalte 15 eingetragen werden
        'Die Werte der übrigen T9-Spalten werden ermittelt und ebenfalls in T9 eigetragen
    A = Split(AllPaths, vbCrLf)
    For i = 0 To UBound(A)
        p = A(i)    'one path, new to T9
        N1 = getNameOfPath(p)                   'N1  = Name EventFolder
        Nat = T9_Get_NationFromEventName(N1)    'Nat = Nation 'GB
        N2 = T9_Get_EventNameNoNation(N1, Nat)  'N2  = N1 ohne "_GB"
        T9_Fill_s2Date_s3Date_s4Event_s5Ort N2, s2Date, s3Date, s4Event, s5Ort
        'Format in aktuelle Zeile übertragen
            Set r2 = .Range(.Cells(zLast + 1 + i, 2), .Cells(zLast + 1 + i, 16))
            r1.Copy: r2.PasteSpecial xlPasteFormats: Application.CutCopyMode = False
        'Spalte 2
            .Cells(zLast + 1 + i, 2) = s2Date
        'Spalte 3
            .Cells(zLast + 1 + i, 3) = s3Date
        'Spalte 4
            .Cells(zLast + 1 + i, 4) = s4Event
        'Spalte 5
            .Cells(zLast + 1 + i, 5) = s5Ort
        'Spalte 7
            .Cells(zLast + 1 + i, 7) = Nat
        'Spalte 9
            If p Like "* Liga\*" Then .Cells(zLast + 1 + i, 9) = "Liga"
            If p Like "*Buli*" Then .Cells(zLast + 1 + i, 9) = "Liga"
        'Spalte 13
            .Cells(zLast + 1 + i, 13) = "1" 'Ordner-Sybol (Wingdings)
        'Spalte 14
            .Cells(zLast + 1 + i, 14) = N1   'Name EventOrdner
        'Spalte 15
            .Cells(zLast + 1 + i, 15) = p
        'Spalte 16
            .Cells(zLast + 1 + i, 16) = "|"
    Next
        
    'Finals
        'T9_PepUp_T9
        End With
End Sub

Sub T9_Fill_s2Date_s3Date_s4Event_s5Ort(N$, s2Date$, s3Date$, s4Event$, s5Ort$)
    'Called from    T9_Write_NewEventFolders_ToT9
    'N              = 1959-09-05 Nissen-Cup2 Wasen '1968-08 Nissen-Cup '1999 WM21 Sun City
    
    'Vorbereitung
        Dim D$, r$, W1$, W2$, W3$, W4$, E%, i%, w%, A() As String
    'Datum - vor dem 1. Leerzeichen
        D = Left(N, InStr(N, " ") - 1)  '1959-09-05 '1968-08 '1999
        If Len(D) = 4 Then s2Date = D + "0000"
        If Len(D) = 7 Then s2Date = Left(D, 4) + Mid(D, 6, 2) + "00"
        If Len(D) = 10 Then s2Date = Replace(D, "-", "")
        s3Date = Right(s2Date, 2) + "." + Mid(s2Date, 5, 2) + "." + Left(s2Date, 4)
    'Rest - N ohne Datum; Event oder Event+Ort
        r = Mid(N, Len(D) + 2) 'Nissen-Cup2 Wasen 'Nissen-Cup 'LK SA-D Sun City
        A = Split(r, " ")
        
    'Letztes EventWord suchen
        w = UBound(A) + 1      'w = Anzahl Worte
        For i = UBound(A) To 0 Step -1
            If IsEventWord(A(i)) Then E = i + 1: Exit For
        Next
        'e-tes Wort ist letztes EventWord   'w = Anzahl Worte
        If E = 1 And w = 1 Then s4Event = A(0): s5Ort = ""
        If E = 1 And w = 2 Then s4Event = A(0): s5Ort = A(1)
        If E = 1 And w = 3 Then s4Event = A(0): s5Ort = A(1) + " " + A(2)
        If E = 2 And w = 2 Then s4Event = A(0) + " " + A(1): s5Ort = ""
        If E = 2 And w = 3 Then s4Event = A(0) + " " + A(1): s5Ort = A(2)
        If E = 2 And w = 4 Then s4Event = A(0) + " " + A(1): s5Ort = A(2) + " " + A(3)
        If E = 3 And w = 3 Then s4Event = A(0) + " " + A(1) + " " + A(2): s5Ort = ""
        If E = 3 And w = 4 Then s4Event = A(0) + " " + A(1) + " " + A(2): s5Ort = A(3)
        If E = 3 And w = 5 Then s4Event = A(0) + " " + A(1) + " " + A(2): s5Ort = A(3) + " " + A(4)
        If s4Event = "" Then Stop 'keiner der vorangegangenen Fälle trifft zu
End Sub

Function IsEventWord(w$) As Boolean
    'w      = Wortfolge aus NameOfEventFolder (ohne Datum, ohne "_GB")
    If w Like "*[A-Z][A-Z]*" Then IsEventWord = True:       Exit Function
    If w Like "*[a-z][A-Z]*" Then IsEventWord = True:       Exit Function
    If w Like "*#*" Then IsEventWord = True:                Exit Function
    If w Like "*Cup*" Then IsEventWord = True:              Exit Function
    If w Like "*Buli*" Then IsEventWord = True:             Exit Function
    If LCase(w) Like "*fest*" Then IsEventWord = True:      Exit Function
    If LCase(w) Like "*sport*" Then IsEventWord = True:     Exit Function
    If LCase(w) Like "*treff*" Then IsEventWord = True:     Exit Function
    If LCase(w) Like "*games*" Then IsEventWord = True:     Exit Function
    If LCase(w) Like "*pokal*" Then IsEventWord = True:     Exit Function
    If LCase(w) Like "*turnier*" Then IsEventWord = True:   Exit Function
    If LCase(w) Like "*lehrgang*" Then IsEventWord = True:  Exit Function
    If LCase(w) Like "* open*" Then IsEventWord = True:     Exit Function
    If LCase(w) Like "* offen*" Then IsEventWord = True:    Exit Function
End Function

Sub T9_Write_KnollenInSpalte1(z%)
    Dim r As Range
    With Sheets("T9")
        With .Cells(z, 1)
            .Font.NAME = "Wingdings": .Font.size = 26:  .Value = Chr(158): .Font.Color = vbRed
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        End With
        If .Cells(z, 2) = "" Then
            .[B4] = "Klick auf roten Knollen: Zeile " + CStr(z) + " wird gelöscht, da Zelle(2, " + CStr(z) + ") leer ist."
        Else
            .[B4] = "Klick auf roten Knollen: Eine neue Zeile wird direkt unterhalb des Knollens angelegt, da Zelle(2, " + CStr(z) + ") nicht leer ist."
        End If
    End With
End Sub

Sub T9_Set_s2Date_to_Green3_ifFolderExists()
    Dim i&, zLast&, r As Range, A() 'As String
    With Sheets("T9")
    zLast = Get_NrOfLastRowInColumnNr(3, "T9")
    Set r = .Range(.Cells(9, 13), .Cells(zLast, 13))
    A = r.Value
    For i = 1 To UBound(A, 1)
        If A(i, 1) = "1" Then .Cells(i + 8, 3).Interior.Color = Green3
    Next
    End With
End Sub

Sub T9_KnollenKlick_Auswerten(z%)
    Dim r As Range
    With Sheets("T9")
    If .Cells(z, 2) = "" Then
        'Zeile r löschen und Zellen von unten nach oben nachrücken lassen
        Set r = .Range(.Cells(z, 2), .Cells(z, 16))
        r.Delete Shift:=xlUp
    Else
        'Zeile einfügen und die darunter liegenden Zellen nach unten verschieben
        Set r = Range(Cells(z + 1, 2), Cells(z + 1, 16))
        r.Insert Shift:=xlDown 'inkl. komplett leeren farbigen Zeilen
        CursorInZelleErzwingen z + 1, 2
    End If
    'Finals
        End With: T9_Write_B4
End Sub

Sub T9_RahmenUmDateZellen(z%)
    'Called from    Worksheet_SelectionChange
    
    'Vorbereitung
        Dim y$, i&, z1&, z2&, zLast&, A(), r As Range
        With Sheets("T9"): If .Cells(z, 2) = "" Then Exit Sub
        zLast = Get_NrOfLastRowInColumnNr(3, "T9")
    'z1, z2 ermitteln (gleiche JahresZahl bei Zeilen z1 bis z2)
        y = Left(.Cells(z, 2), 4)
        Set r = .Range(Cells(9, 2), Cells(zLast, 2)): A = r.Value
        For i = 1 To UBound(A, 1)
            If z1 = 0 And CStr(A(i, 1)) Like y + "*" Then z1 = i + 8
            If CStr(A(i, 1)) Like y + "*" Then z2 = i + 8
            If z2 > 0 And Not CStr(A(i, 1)) Like y + "*" Then Exit For
        Next
    'YearRahmen setzen
        Set r = .Range(.Cells(z1, 2), .Cells(z2, 3))
        r.BorderAround LineStyle:=xlContinuous, Weight:=xlThick, ColorIndex:=3
        End With
End Sub

Sub T9_LookFor_s3Date_toFill(zLast&)
    Dim D$, s4$, i&, zD&, zT9&, r As Range, A()
    With Sheets("T9"): DoArrc
    Set r = .Range(.Cells(9, 2), .Cells(zLast, 3)): A = r.Value
    'Datum eintragen
        For i = 1 To UBound(A, 1)
            If CStr(A(i, 2)) = "" Then
                If CStr(A(i, 1)) Like "########" Then
                    D = CStr(A(i, 1)): zD = i + 8
                    .Cells(zD, 3) = Right(D, 2) + "." + Mid(D, 5, 2) + "." + Left(D, 4)
                End If
            End If
        Next
    'neues Datum einordnen
        If D = "" Then Exit Sub
        s4 = .Cells(zD, 4)
        .Cells(zD, 4) = "°°°" + s4
        T9_Sort 4: T9_Sort 2
        '"°°°" finden
        zT9 = Get_RowNr_HoldingMyTextPartInColumnX("T9", 4, 8, "°°°")
        .Cells(zT9, 4) = s4
        .Cells(zT9, 4).Select
        FillArrC 79, "1"
        CursorInZelleErzwingen zT9, 4
        
    End With
End Sub

Sub CursorInZelleErzwingen(z&, s&)
    'Das Excel-Fenster sichtbar machen und in den Vordergrund zwingen
        Application.Visible = True
    'Fokus auf das Excel-Hauptfenster setzen
        On Error Resume Next: AppActivate Application.Caption: On Error GoTo 0
    'Die Zelle auswählen
        'EE 0: Cells(z, s).Select: EE 1
        Cells(z, s).Select
    'Den Bearbeitungsmodus intern starten
        On Error Resume Next: Application.CommandBars.FindControl(Id:=1).Execute: On Error GoTo 0
    DoEvents: Application.SendKeys "{F2}"
End Sub


