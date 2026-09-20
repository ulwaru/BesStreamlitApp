Attribute VB_Name = "M31"
Option Explicit 'M_T3

Sub zzz_M31()
    
    'showProcs "hide"
    Sheets("T3").Cells(1, 1).Select
    'T3_Sort_Nn
    EE 1: Beep
End Sub

Sub T3_Move_RedRect_ToName(Nn$, Vn$)
    'Called from    xx
    
    'Vorbereitung
        Dim z%, r As Range
        
    'Get RowNr
        z = Get_RowNr_HoldingMyTextWholeInColumnX("T3", 2, 7, Nn)
        z = Get_RowNr_HoldingMyTextWholeInColumnX("T3", 3, z - 1, Vn)
        Set r = Sheets("T3").Cells(z, 2)
    'Move
        With Sheets("T3").Shapes("T3RedRect")
            .Left = r.Left - 3: .Top = r.Top - 3: .Height = r.Height + 6
            .Width = r.Width + Sheets("T3").Cells(z, 3).Width + 6: .Visible = True
        End With
    RefreshScreen
    EE 0: Sheets("T3").Cells(z, 5).Select: EE 1
    'Stop
End Sub

Sub T3_Create_ResultsTextFile_ForEachCompetitor()
    'Called from    xx
    
    'Vorbereitung
        Dim L$, LL$, N$, Nn$, PF$, v$, Vn$, C1%, C2%, i&, j&, T3(), ArrLL() As String
        v = vbCrLf: DoArrc: EE 0
    'Sort T3 nach Nn Vn EventTitle (zeigt nach Nn, Vn auch Alter aufsteigend)
        Call T3_Sort_mw: T3_Sort_Title: T3_Sort_Vn: T3_Sort_Nn 'sort ohne Select
    'Get Data
        T3_Load_DataArea T3 '8629 Zeilen, 10 Spalten
    'Action
        For i = 1 To UBound(T3, 1)
            'Schleife über alle T3-Zeilen
            N = "|" + T3(i, 1) + "|" + T3(i, 2) + "|"       'N    = "|May|Lea|"
            L = N                                           'L    = one Line
                For j = 4 To 10
                    If CStr(T3(i, j)) <> "" Then L = L + CStr(T3(i, j)) + "|"
                Next
            If LL Like N + "*" Then
                LL = LL + v + L                     'LL = some Lines, same Name
            Else
                'LL ist vollständig; enthält alle Zeilen einer Person
                If LL Like "*" + v + "*" + v + "*" Then 'mindestens 3 Zeilen
                    LL = Replace(LL, " (Einzel)", "")
                    LL = Replace(LL, " (Synchron)", "")
                    LL = Replace(LL, " (Mannschaft)", "")
                    LL = Replace(LL, "|Einzel", "|°Einzel    ")
                    LL = Replace(LL, "|Synchron", "|°Synchron  ")
                    LL = Replace(LL, "|Mannschaft", "|°Mannschaft")
                    ArrLL = Split(LL, v)
                    C2 = 0
                    'c2 ermitteln   (c2 = max CountOfChars to "|°")
                        For j = 0 To UBound(ArrLL)
                            C1 = InStr(1, ArrLL(j), "|°")
                            If C1 > C2 Then C2 = C1
                        Next
                    'Tab-Simulation (Einfügen von Leerzeichen)
                        For j = 0 To UBound(ArrLL)
                            C1 = InStr(1, ArrLL(j), "|°")
                            ArrLL(j) = Replace(ArrLL(j), "|°", String(C2 - C1, " ") + "|°")
                        Next
                    'Rang formatieren
                        For j = 0 To UBound(ArrLL)
                            C1 = Len(ArrLL(j))
                            If C1 = C2 + 14 Then
                                ArrLL(j) = Left(ArrLL(j), C2 + 12) + "Rang:  " + Mid(ArrLL(j), C2 + 13)
                            Else
                                ArrLL(j) = Left(ArrLL(j), C2 + 12) + "Rang: " + Mid(ArrLL(j), C2 + 13)
                            End If
                        Next
                    'Vn Nn
                        C1 = InStr(1, LL, "|19"): Vn = Mid(LL, 2, C1 - 1) 'May|Lea|SU
                        C1 = InStr(1, Vn, "|"): C2 = InStr(C1 + 1, Vn, "|")
                        Nn = Left(Vn, C1 - 1)
                        Vn = Mid(Vn, C1 + 1, C2 - C1 - 1)
                    'EndText erstellen
                        LL = Join(ArrLL, v): LL = Replace(LL, "°", "")
                        LL = CStr(anzAinB("Rang:", LL)) _
                        + " Wettkampfergebnisse aus den z.Zt. verarbeiteten Ergebnislisten (" _
                        + Format(Now(), "dd.mm.yyyy") + ")" + v + v + LL
                    'RedRect
                        T3_Move_RedRect_ToName Nn, Vn
                    'Path of PersonFolder
                        PF = T5_Get_PathOfPersonfolder_Give_VnNn(Vn + " " + Nn)
                        'OpenFolder PF
                    'Write
                        writeStringToFile PF + "\[results] " + Vn + " " + Nn + ".txt", LL
                End If
                LL = L
            End If
        Next
        Hide_Shape "T3", "T3RedRect"
        EE 0: Sheets("T3").Cells(1, 1).Select: EE 1
End Sub

Sub T3_Create_TextFile_ListOfCompetitors()
    'Called from    [T3-CellBtn "Create Textfile of this list"] Worksheet_SelectionChange[T3]
    'Action         erstellt TextDatei "...\Archiv Trampolin\Competitors.txt";
    'DatenTabelle   wird als TabSimulation dargestellt (Auffüllen einer Spalte
    '               mit Leerzeichen bis feste Spaltenbreite erreicht ist);
    
    'Vorbereitung
        Dim s$: DoArrc
    'Get Data
        s = T3_Get_Text_CompetitorsList
    'Show
        'writeStringToFile ArrC(2) + "\Competitors.txt", s
        show s
End Sub

Function T3_Get_Text_CompetitorsList() As String
    'Called from    T3_Create_TextFile_ListOfCompetitors
    'Action         liefert String der T3-DataArea CompetitorsList;
    'DatenTabelle   wird als TabSimulation dargestellt (Auffüllen einer Spalte
    '               mit Leerzeichen bis Spaltenbreite erreicht ist);
    'Spaltenbreite  ist je Spalte fest vorgegeben
    
    'Vorbereitung
        Dim s$, v$, i&, T3()
        v = vbCrLf
        T3_Load_DataArea T3 '8629 Zeilen, 10 Spalten
    'Data
    For i = 1 To UBound(T3, 1)
        'Schleife über alle T3-Zeilen
        T3(i, 1) = "|" + Left(T3(i, 1) + String(44, " "), 19)   '19 = feste Vorgabe Len(Spalte1)
        T3(i, 2) = Left(T3(i, 2) + String(44, " "), 13)         '13 = feste Vorgabe Len(Spalte2)
        T3(i, 3) = " " + T3(i, 3) + " "                                         'w/m
        If Len(T3(i, 4)) = 0 Then T3(i, 4) = "    "                             'Alter
        If Len(T3(i, 4)) = 1 Then T3(i, 4) = "  " + CStr(T3(i, 4)) + " "        'Alter
        If Len(T3(i, 4)) = 2 Then T3(i, 4) = " " + CStr(T3(i, 4)) + " "         'Alter
        T3(i, 5) = Left(T3(i, 5) + String(44, " "), 23)                         'Verein
        T3(i, 6) = Left(T3(i, 6) + String(44, " "), 18)                         'LTV
        If Len(T3(i, 7)) = 1 Then T3(i, 7) = " " + T3(i, 7) + "   "             'Nation
        If Len(T3(i, 7)) = 2 Then T3(i, 7) = " " + T3(i, 7) + "  "              'Nation
        If Len(T3(i, 7)) = 3 Then T3(i, 7) = " " + T3(i, 7) + " "               'Nation
        If T3(i, 8) Like "*" + T3(i, 9) + "*" Then T3(i, 8) = Left(T3(i, 8), Len(T3(i, 8)) - Len(T3(i, 9)) - 2)
        T3(i, 8) = Left(T3(i, 8) + String(44, " "), 44)                         'Event
        T3(i, 9) = T3(i, 9) + String(10 - Len(T3(i, 9)), " ")                   'Typ
        If Len(T3(i, 10)) = 1 Then T3(i, 10) = "  " + CStr(T3(i, 10)) + " |"    'Rang
        If Len(T3(i, 10)) = 2 Then T3(i, 10) = " " + CStr(T3(i, 10)) + " |"     'Rang
        If Len(T3(i, 10)) = 3 Then T3(i, 10) = CStr(T3(i, 10)) + " |"           'Rang
    Next
    'Kopfzeile + Data
        s = "Competitors" + "  -  Aktive aus den z.Zt. verarbeiteten Ergebnislisten (" _
        + Format(Now(), "dd.mm.yyyy") + ")            " _
        + Sheets("T3").[B4].Value + v + v + String(154, "-") + v _
        + "   Nachname            Vorname     w/m Alter    Verein/Ort                 " _
        + "LTV        Nation            Event                               " _
        + "Typ      Rang" + v + String(154, "-") + v + Join2d(T3, v, "|") + v + String(154, "-")
    'Finals
        T3_Get_Text_CompetitorsList = s
End Function

Sub T3_FarbWechsel_CellBtn_CreateTxtFile()
    'Called from    xxx
    'Action         CellButton 'Create textfile of this list' blinken lassen
    
    Dim r As Range: Set r = Sheets("T3").Cells(6, 13)
    If CInt(ArrC(25)) Mod 2 = 0 Then r.Interior.Color = Green3 Else r.Interior.Color = Green1
    'If r.Interior.Color = Green1 Then r.Interior.Color = Green3 Else r.Interior.Color = Green1
End Sub

Sub T3_Create_CompetitorsList()
    'Called from    [T3-CellButton 'Update this list"]
    '               [T6-CellButton 'Update CompetitorsList']
    'Action         T4SomeDgData DgTitle-Spalte --> D1
    '               T4SomeDgData DgNames-Spalte --> D8
    
    'Vorbereitung
        Dim Club$, clubs$, D$, N$, Nation$, Nations$, s$, T5Stats$, Title$, v$
        Dim AnzC%, AnzClubs%, AnzM%, AnzNations%, AnzOld%, AnzT%, AnzW%
        Dim i%, j%, Jhg%, y%, zLast%, zT5%, zT6%
        Dim T3(), T5(), U1()
        Dim D1() As String, D8() As String, u() As String
        Dim N1() As String, N2() As String, N3() As String
        v = vbCrLf: Ticks1
    'AnzOld merken  (bisherige Anzahl der CompetitorList-Zeilen) (wegen CountOfChanges)
        s = Sheets("T3").[B4]: i = InStr(1, s, " "): AnzOld = CInt(Left(s, i - 1))
    'MyCountDown Init
        zT6 = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 2, 4, " Update CompetitorList")
        '   = T6-ZeilenNr, die den Cell-Button enthält (für T6_DoDoneRemarks-Einträge)
        MyCountDown_Init "|T6|" + CStr(zT6) + "|19|50|" '= |SheetName|z|s|StartNr| of CountDownShowCell
        T3_Update_CountDownNr
    'D1() D8()
        T4_Load_T4SomeDgData_OneCol 1, D1   'all DgTitles  '1967-06-17 WM04 London (Synchron)
        T4_Load_T4SomeDgData_OneCol 8, D8   'all DgNames   '002 Names: |#|Jo Do|6|3|Einz|1|#|...|9|
        Call MyCountDown_OneMoreMainStep: T3_Update_CountDownNr
    's = Sammlung von           "|Vn Nn|6|3|Einz|1|DgTitle"-Zeilen
        MyCountDown_MainStepsAllowed 30: MyCountDown_SubStepsMax UBound(D1)
        For i = 1 To UBound(D1)
            N = D8(i)              'N = all Names of one Dg '002 Names: |#|Jo Do|6|3|Einz|1|#|...|9|
            If N Like "*|*" Then
                'Die Names-Zelle von T4SomeDgData enthält nicht nur "000 Names"
                N1 = Split(N, "|#")
                For j = 1 To UBound(N1)
                    s = s + N1(j) + "|" + D1(i) + v ' "|Lea May|6|3|Einz|1|1966 DM Bonn"
                Next
            End If
            Call MyCountDown_OneMoreSubStep: If i Mod 5 = 0 Then T3_Update_CountDownNr
        Next
        Call MyCountDown_OneMoreMainStep: If i Mod 5 = 0 Then T3_Update_CountDownNr
        'show s: Stop
        'Repair     falls Rangzahl im Dg fehlte
            s = Replace(s, "|||", "|-|")
            s = Replace(s, "|Einz||19", "|Einz|-|19")
            s = Replace(s, "|Synchron|", "|Sync|")      'sollte unnötig sein
            s = Replace(s, "|Sync||19", "|Sync|-|19")
            s = Replace(s, "|Team||19", "|Team|-|19")
            s = Replace(s, "||", "|")
            Call MyCountDown_OneMoreMainStep: T3_Update_CountDownNr
            s = Delete_EmptyRowsInString(s)
    'N2() 1D                                    "|Vn Nn|6|3|Einz|1|DgTitle"
        N2 = Split(v + s, v)
        QuickSort N2                            'N2(0)=""
        AnzC = UBound(N2)                       'AnzC = Anzahl CompetitorsZeilen
        FillArrC 24, CStr(AnzC)                 'MyCountDown CountOfChanges
        'showArray N2: Stop
    'T5()  2D     '[T5]                         "|Nn|Vn|mw|Jhg|Club|LTV|Nation|"
        With Sheets("T5")
        Call T5_Sort_Nn: T5_Sort_VnNn           'SortUp nach VnNn ohne Select
        T5 = .Range(.Cells(8, 3), .Cells(Get_T5_zLast, 9)).Value     ':showArray2D T5
        AnzT = UBound(T5, 1)                    'AnzT = Anzahl Namen in T5
        'showArray2D T5, "T5() [2D-Array; T5 PersonData]"
    'Stats       '[T5]
        T5Stats = .Cells(4, 2): N3 = Split(T5Stats, " ")
        AnzW = CInt(Mid(N3(2), 2)): AnzM = AnzT - AnzW: AnzNations = CInt(N3(6))
    'U()    1D-Array;   U(1) = "A. Nishikawa";  U(1965) = "Zuzana Svebisova"
        U1 = .Range(.Cells(8, 22), .Cells(Get_T5_zLast, 22)).Value  'U1(x, 1) = [Vn] [Nn] '2D
        Load_StringArray1D_from_1ColumnArray2D u, U1                'U(x)     = [Vn] [Nn] '1D
        End With 'T5
        'showArray U: Stop
        Call MyCountDown_OneMoreMainStep: T3_Update_CountDownNr
    'T3()                                       'T3: "|Nn|Vn|mw|Age|Club|LTV|Nation|DgTitle|ESyM|Rg"
        'Leeres 2D-Array bereitstellen
            ReDim T3(1 To AnzC, 1 To 10)        'AnzC = Anzahl CompetitorsZeilen
        MyCountDown_MainStepsAllowed 5:         MyCountDown_SubStepsMax AnzC
        For i = 1 To AnzC
            N3 = Split(N2(i), "|")              'N2(i) = "|Vn Nn|6|3|Einz|1|DgTitle"
            'showArray N3: Stop
            
            
            'zT5        U durchsuchen   'U(1) = "A. Nishikawa";  U(1965) = "Zuzana Svebisova"
                For j = 1 To AnzT               'AnzT  = Anzahl Namen in T5, U()
                    zT5 = 0
                    If u(j) = N3(1) Then        'N3(1) = Vn Nn 'aus N2(i) = "|Vn Nn|6|3|Einz|1|DgTitle"
                        'N3(1) wurde in U() beim Index j gefunden
                        zT5 = j
                        j = AnzT '= Exit For j
                    End If
                Next
                If zT5 = 0 Then
                    'N3(1) wurde nicht in U() gefunden
                    show "T3_Create_CompetitorsList" + v + v _
                    + "Der Name '" + N3(1) + "' wurde nicht in T5 gefunden;" + v _
                    + "ggf. die Schreibweise in T4 und T5 angleichen"
                    Stop: Exit Sub
                End If
                'T5(zT5, 1) = Nn     'T5 trägt 7 Items in Zeile zT5: |Nn|Vn|mw|Jhg|Club|LTV|Nation|
            '1  Nn
                T3(i, 1) = T5(zT5, 1)
            '2  Vn
                T3(i, 2) = T5(zT5, 2)
            '3  m/w
                T3(i, 3) = T5(zT5, 3)
            '4  Alter
                If Trim(T5(zT5, 4)) = "" Then
                    'Jhg ist nicht in T5 angegeben
                    T3(i, 4) = ""
                Else
                    Title = N3(6): Jhg = CInt(T5(zT5, 4))
                    y = CInt(Left(Title, 4)) 'Y = Year of Event
                    T3(i, 4) = CStr(y - Jhg)
                End If
            '5  Club
                Club = CStr(T5(zT5, 5)): T3(i, 5) = Club
                If Len(Club) > 1 Then
                    If Not clubs Like "*|" + Club + "|*" Then
                        clubs = clubs + "|" + Club + "|": AnzClubs = AnzClubs + 1 '226?
                    End If
                End If
            '6  LTV
                T3(i, 6) = T5(zT5, 6)
            '7  Nation
                T3(i, 7) = T5(zT5, 7)
            '8  DgTitle
                T3(i, 8) = N3(6)
            '9  ESyM
                If N3(4) = "Einz" Then T3(i, 9) = "Einzel"
                If N3(4) = "Sync" Then T3(i, 9) = "Synchron"
                If N3(4) = "Team" Then T3(i, 9) = "Mannschaft"
            '10 Rang
                T3(i, 10) = N3(5)
            'CountDown
                Call MyCountDown_OneMoreSubStep: If i Mod 500 = 0 Then T3_Update_CountDownNr
        Next
        Call MyCountDown_OneMoreMainStep: If i Mod 500 = 0 Then T3_Update_CountDownNr
    'Stats      '[T3]
        With Sheets("T3")
        .[B4] = CStr(AnzC) + " Zeilen, " + CStr(AnzT) _
        + " Aktive (" + CStr(AnzW) + " w, " + CStr(AnzM) + " m), " _
        + CStr(AnzNations) + " Nations, " + CStr(AnzClubs) + " Clubs"
    'Clear old  '[T3]
        zLast = Get_NrOfLastRowInColumnNr(2, "T3")
        .Range(.Cells(8, 2), .Cells(zLast, 11)).ClearContents
        .Range(.Cells(8, 2), .Cells(zLast + 10, 11)).Borders.LineStyle = xlNone
        Call MyCountDown_OneMoreMainStep:  T3_Update_CountDownNr
    'Paste      '[T3]
        Paste_2DArrayToCell_z_s "T3", 8, 2, T3
        Call MyCountDown_OneMoreMainStep: T3_Update_CountDownNr
    'Hintergrund einfärben  '[T3]
        'show CStr(getColor([e8], 0, "T4")) '--> 14083324 (w-Farbe, rosa)
        'show CStr(getColor([e9], 0, "T4")) '--> 15652797 (m-Farbe, hellblau)
        For i = 8 To zLast
            If CStr(T3(i - 7, 3)) = "w" Then .Range(.Cells(i, 2), .Cells(i, 11)).Interior.Color = 14083324
            If CStr(T3(i - 7, 3)) = "m" Then .Range(.Cells(i, 2), .Cells(i, 11)).Interior.Color = 15652797
        Next
        Call MyCountDown_OneMoreMainStep: T3_Update_CountDownNr
        For i = zLast To zLast + 100
            .Range(.Cells(i, 2), .Cells(i, 11)).Interior.ColorIndex = xlNone
        Next
        Call MyCountDown_OneMoreMainStep: T3_Update_CountDownNr
    'Rahmen - each CellBorder: white    '[T3]
        With .Range(.Cells(8, 2), .Cells(zLast, 11)).Borders
            .LineStyle = xlContinuous: .Color = RGB(250, 250, 250): .Weight = xlThin
        End With
        Call MyCountDown_OneMoreMainStep: T3_Update_CountDownNr
    'Alphabetisch sortieren     '[T3]
        'Call Screen0: .Activate: .[i6].Select: .[c6].Select: .[B6].Select '[T3]
        Call T3_Sort_mw: T3_Sort_Vn: T3_Sort_Nn 'sort ohne Select
        'Sheets("T6").Activate
        Call MyCountDown_OneMoreMainStep: T3_Update_CountDownNr
    'DoneRemarks to T6          '[T3]
        FillArrC 24, CStr(AnzC - AnzOld)                 'MyCountDown CountOfChanges
        T6_DoDoneRemarks
    'T3-Countdown               '[T3]
        .[M15] = ""
    'Write and Show
        writeStringToFile ArrC(2) + "\Competitors.txt", T3_Get_Text_CompetitorsList
        openFile ArrC(2) + "\Competitors.txt"
'    'Finals
        End With
        Beep
End Sub

Sub T3_Load_DataArea(ByRef Arr2D())
    With Sheets("T3")
        Arr2D = .Range(.Cells(8, 2), .Cells(Get_NrOfLastRowInColumnNr(2, "T3"), 11)).Value
    End With
End Sub

Sub T3_Update_CountDownNr()
    Sheets("T3").Cells(15, 13) = ArrC(25)
End Sub

Sub T3_Count_Clubs(AnzClubs%, AllClubs$, Club$)
    If Club <> "" Then
        If InStr(1, AllClubs, "|" + Club + "|") = 0 Then
            AllClubs = AllClubs + "|" + Club + "|" + vbCrLf
            AnzClubs = AnzClubs + 1
        End If
    End If
End Sub

Sub T3_Count_Competitors(AnzW%, AnzM%, AnzAktive%, Nn$, Vn$, mw$, AllNnVnMW$)
    Dim NnVnMW$
    NnVnMW = "|" + Nn + "|" + Vn + "|" + mw + "|"
    If InStr(1, AllNnVnMW, NnVnMW) = 0 Then
        AllNnVnMW = AllNnVnMW + NnVnMW + vbCrLf
        If mw = "w" Then AnzW = AnzW + 1
        If mw = "m" Then AnzM = AnzM + 1
        AnzAktive = AnzAktive + 1
    End If
End Sub

Sub T3_Count_Nations(AnzNationen%, AllNations$, Nation$)
    If Nation <> "" Then
        If InStr(1, AllNations, "|" + Nation + "|") = 0 Then
            AllNations = AllNations + "|" + Nation + "|" + vbCrLf
            AnzNationen = AnzNationen + 1
        End If
    End If
End Sub

Sub T3_Sort(s%)
    'Called from    Worksheet_SelectionChange
    's              = 5 (Nachname), = 6 (Vorname), ...
    'Pfeil Up       "ñ" (Zeichensatz Wingdings)
    'Pfeil Down     "ò" (Zeichensatz Wingdings)
    
    'Vorbereitung
        Dim sL$, Sortierbereich$, Sortierspalte$, SortStOld$, SortSpOld%, zLast%
        With Sheets("T3"): DoArr
        zLast = Get_NrOfLastRowInColumnNr(2, "T3")
        If ArrC(55) = "" Then ArrC(55) = CStr(s)    'ArrC(55) = Alte Sortierspalte
        SortSpOld = CInt(ArrC(55))                  'T3 Nr der Sortierspalte LastSort
        SortStOld = ArrC(56)                        'T3 Status LastSort, U=Up, D=Down
        sL = Get_ColumnLetter(s)
        Sortierspalte = sL + "8"                    '"F8"
        Sortierbereich = "B8:K" + CStr(zLast)
    'Zeile unterhalb des Tabellenkopfes (grau) leeren
        .Range("B7:K7") = ""
    'ErledigtHaken bei CellButton 'Create textfile' entfernen
        .Range("r7") = ""
    'sort
        If CInt(SortSpOld) = s Then
            If SortStOld = "U" Then
                T3_SortDown Sortierbereich, Sortierspalte: ArrC(56) = "D": .Cells(7, s) = "ò"
            Else
                T3_SortUp Sortierbereich, Sortierspalte: ArrC(56) = "U": .Cells(7, s) = "ñ"
            End If
        Else
            T3_SortUp Sortierbereich, Sortierspalte: ArrC(56) = "U": .Cells(7, s) = "ñ"
        End If
        ArrC(55) = CStr(s)
        If ActiveSheet.NAME = "T3" Then EE 0: .[a6].Select: EE 1
    'Finals
        End With
End Sub

Sub T3_Sort_Nn(): T3_Sort 2: End Sub    'sort ohne Select
Sub T3_Sort_Vn(): T3_Sort 3: End Sub
Sub T3_Sort_mw(): T3_Sort 4: End Sub
Sub T3_Sort_Title(): T3_Sort 9: End Sub

Sub T3_SortUp(Sortierbereich$, Sortierspalte$)
    With ThisWorkbook.Worksheets("T3")
        .Range(Sortierbereich).Sort Key1:=.Range(Sortierspalte), Order1:=xlAscending, _
        Header:=xlNo, orderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    End With
End Sub

Sub T3_SortDown(Sortierbereich$, Sortierspalte$)
    With ThisWorkbook.Worksheets("T3")
        .Range(Sortierbereich).Sort Key1:=.Range(Sortierspalte), Order1:=xlDescending, _
        Header:=xlNo, orderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    End With
End Sub

Sub T3_CountDown1(i%, max%)
    'Called from    T3_Create_TextFile_ListOfCompetitors
    
    'Vorbereitung
        Dim c%, D%
        't = ChrW(159) '189 = Zeichencode für mittelgroßen Knollen; Wingdings
        With Sheets("T3")
    'Action
        If i = 1 Then .[M15] = 33
        If i = max \ 2 Then .[M15] = 32
        If i = 2 * max \ 3 Then .[M15] = 31
        If i = max Then .[M15] = 30
        DoEvents
    'Finals
        End With
End Sub

Sub T3_CountDown2(Counter%, i%, iMax%)
    'Called from    TAB_Simulation
    
    'Vorbereitung
        Dim ContDownFrom%
        ContDownFrom = 30
    'Action
        If Counter = 0 Then Counter = ContDownFrom
        'c soll zu je gleichen i-Abständen von max bis 1 herunterzählen
        If i Mod iMax \ (ContDownFrom - 2) = 0 Then
            Counter = Counter - 1: Sheets("T3").[M15] = Counter
        End If
End Sub

Sub T3_CountDown3(Counter%, i%, iMax%)
    'Called from    TAB_Simulation
    
    'Vorbereitung
        Dim ContDownFrom%
        ContDownFrom = 9
    'Action
        If Counter = 0 Then Counter = ContDownFrom
        'c soll zu je gleichen i-Abständen von max bis 1 herunterzählen
        If i Mod iMax \ (ContDownFrom - 1) = 0 Then
            Counter = Counter - 1: Sheets("T3").[Q3] = Counter
        End If
End Sub

Sub T3_CountDown4(tic1&, CdMax%, CdNr%)
    'Called from    T3_Create_CompetitorsList
    
    'Vorbereitung
        Dim s$, TimeNeeded$, zNr%, t1#
    'Action
        If tic1 = 0 Then
            tic1 = GetTickCount()
            zNr = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 2, 4, " Update_CompetitorList")
            TimeNeeded = Sheets("T6").Cells(zNr, 21) '3,6 sec '39,9 min
            If Right(TimeNeeded, 3) = "sec" Then
                s = Replace(TimeNeeded, " sec", "")
                s = Replace(s, ",", "")
                CdMax = CInt(2 * CInt(s) \ 10)
            End If
            If Right(TimeNeeded, 3) = "min" Then
                s = Replace(TimeNeeded, " min", "")
                s = Replace(s, ",", "")
                CdMax = CInt(2 * 60 * CSng(s) \ 10)
            End If
            CdNr = CdMax
        Else
            t1 = (GetTickCount() - tic1) \ 500
            CdNr = CdMax - CInt((GetTickCount() - tic1) \ 500)
        End If
        Sheets("T3").[M15] = "'" + CStr(CdNr)
        DoEvents
End Sub

Sub CD2_CountDown2(iStart%, iEnd%, i%, CDNrMax%, CDNrMin%, CDNrShow%, t1&, Go As Boolean)
    'i              Laufvariable im Caller; läuft von iStart bis iEnd
    'CDNr           CountDown-Nr; soll (entsprechend i) von CDNrMax bis CDNrMin% laufen
    't1             xxx
    'Go             xxx
    'Action         Liefert eine CountDown-Nr
    'Usage          Im Caller: Dim CDNrShow%, t1&, Go As Boolean
    '               CD2_CountDown2 ...
    '               If Go Then Sheets("T6").[a1] = "[" + CStr(CDNrShow) + "]"
    
    'Vorbereitung
        Dim c As Double
    'Exit
        If t1 = 0 Then t1 = GetTickCount(): Go = True: CDNrShow = CDNrMax: Exit Sub
        If GetTickCount() - t1 < 500 Then Go = False: Exit Sub
    'CDNrShow wird berechnet
        c = (i - iStart) / (iEnd - iStart) 'Anteil Laufweg von i
        
        CDNrShow = CInt(CDNrMax - c * (CDNrMax - CDNrMin))
        
        Go = True: t1 = GetTickCount()
End Sub




