Attribute VB_Name = "M51"
Option Explicit 'M51

Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Const VK_CONTROL As Integer = &H11 'Ctrl-Taste
 
Sub zzz_M_T5()
    'showProcs "nation"
    
    Zeilenlöschen "T2", 9, 13
    EE 1: Beep
End Sub

Sub T5_Sort_Nn(): T5_Sort 3:    End Sub    'sort ohne Select
Sub T5_Sort_Vn(): T5_Sort 4:    End Sub    'sort ohne Select
Sub T5_Sort_VnNn(): T5_Sort 22: End Sub    'sort ohne Select

Sub T5_Sort(s%)
    'Called from    Worksheet_SelectionChange
    's              = 3 (Nachname), = 4 (Vorname), ...
    'Pfeil Up       "ñ" (Zeichensatz Wingdings)
    'Pfeil Down     "ò" (Zeichensatz Wingdings)
    
    'Vorbereitung
        Dim sL$, Sortierbereich$, Sortierspalte$, SortStOld$
        Dim i%, SortSpOld%, m%, w%, sLast%, zLast%, Arr1(), r As Range
        With Sheets("T5"): DoArr
        zLast = Get_T5_zLast
        sLast = Get_T5_sLast
        If ArrC(81) = "" Then ArrC(81) = CStr(s)    'ArrC(81) = Alte Sortierspalte
        SortSpOld = CInt(ArrC(81))                  'T5 Nr der Sortierspalte LastSort
        SortStOld = ArrC(82)                        'T5 Status LastSort, U=Up, D=Down
        sL = Get_ColumnLetter(s)
        Sortierspalte = sL + "8" '"F8"; Spalte, nach der der User sortieren möchte
        Sortierbereich = "C8:" + Get_ColumnLetter(sLast) + CStr(zLast)     'Nachname' bis 'VnNn'
    'Zeile unterhalb des Tabellenkopfes leeren
        .Range("C7:" + Get_ColumnLetter(sLast) + "7") = ""
    'sort
        If CInt(SortSpOld) = s Then
            'Die gleiche Spalte soll erneut sortiert werden
            If SortStOld = "U" Then 'falls zuletzt 'Up' sortiert wurde ...
                If s > 11 Then
                    T5_Fill_SortHelperColumn s, "Down"
                    Sortierspalte = "K8"
                End If
                T5_SortDown Sortierbereich, Sortierspalte
                ArrC(82) = "D"      'T5 Status LastSort, U=Up, D=Down
                .Cells(7, s) = "ò"   'Down-Pfeil
            Else
                If s > 11 Then T5_Fill_SortHelperColumn s, "Up": Sortierspalte = "K8"
                T5_SortUp Sortierbereich, Sortierspalte: ArrC(82) = "U": .Cells(7, s) = "ñ"
            End If
        Else
                If s > 11 Then T5_Fill_SortHelperColumn s, "Up": Sortierspalte = "K8"
            T5_SortUp Sortierbereich, Sortierspalte: ArrC(82) = "U": .Cells(7, s) = "ñ"
        End If
        ArrC(81) = CStr(s)          'Nr der Sortierspalte LastSort
    'PepUp
        'Spalte 2
            Set r = .Range(.Cells(8, 2), .Cells(zLast, 2)): r.HorizontalAlignment = xlCenter: r.Font.size = 6
            r.Interior.ColorIndex = xlNone: r.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
        'Spalte 23
            Set r = .Range(.Cells(8, 23), .Cells(zLast, 23)): r.Value = "|": r.HorizontalAlignment = xlLeft
        'Spalte UserSort
            T5_Format_ColumnUserSort s, zLast
    'Finals
        End With: T5_Update_CountInfos
End Sub

Sub T5_Format_ColumnUserSort(s%, zLast%)
    'Called from    T5_Sort
    's              = Spalte > 11, nach der der User sortieren möchte
    'Vorbereitung
        If s < 12 Or s > 21 Then Exit Sub
        Dim i%, r As Range, ArrU()
        With Sheets("T5")
    'r, Spalte UserSort
        Set r = .Range(.Cells(8, s), .Cells(zLast, s))
        ArrU = r.Value
        For i = 1 To UBound(ArrU, 1)
            If ArrU(i, 1) = "" Then .Cells(i + 7, s) = "-": ArrU(i, 1) = "-"
            If ArrU(i, 1) <> "-" Then
                .Cells(i + 7, s).HorizontalAlignment = xlLeft
            End If
        Next
    'Finals
        End With
End Sub

Sub T5_Fill_SortHelperColumn(sU%, SortUpOrDown$)
    'Called from    T5_Sort
    'sU             = Spalte > 11, nach der der User sortieren möchte
    '               Spalte sU enthält Zellen mit "-" oder "", die bei jeder
    '                Sortierung nach unten sortiert werden sollen
    'sH             = Spalte 11, HelperSpalte
    '                wird entsprechend gefüllt und zur Sortierspalte
    
    'Vorbereitung
        Dim T$, i%, zLast%, ArrH(), rH As Range, rU As Range
        With Sheets("T5"): zLast = Get_T5_zLast
        Set rU = .Range(.Cells(8, sU), .Cells(zLast, sU))
        Set rH = .Range(.Cells(8, 11), .Cells(zLast, 11))
        ArrH = rU.Value
        
    'Fill Arr
        If SortUpOrDown = "Down" Then
            For i = 1 To zLast - 7
                T = ArrH(i, 1)
                If T = "-" Or T = "" Then ArrH(i, 1) = "A" Else ArrH(i, 1) = "B" + T
            Next
        Else
            For i = 1 To zLast - 7
                T = ArrH(i, 1)
                If T = "-" Or T = "" Then ArrH(i, 1) = "B" Else ArrH(i, 1) = "A" + T
            Next
        End If
    'Paste ArrH
        Paste_2DArrayToCell_z_s "T5", 8, 11, ArrH
    'format sU
        rU.Font.size = 8: rU.HorizontalAlignment = xlCenter
        rU.Font.Color = vbBlack
    'format sH
        rH.Font.size = 6: rH.HorizontalAlignment = xlLeft
        rH.Font.Color = RGB(200, 200, 200) 'Grau
    'Finals
        End With
End Sub

Sub T5_Update_CountInfos()
    'Called from    T5_Sort
    
    'Vorbereitung
        Dim A$, N$, s$, T$, u$, c%, i%, m%, w%, r%, zLast%, Arr1(), Arr2() As String
        With Sheets("T5"): zLast = Get_NrOfLastRowInColumnNr(3, "T5"): s = "xxx"
    'Update Anzahl Personen/m/w
        Arr1 = .Range(.Cells(1, 5), .Cells(zLast, 9)).Value
        '    = nur die Spalten |m/w|Jhg|...|Nation|
        For i = 8 To zLast
            'm/w zählen
                If UCase(Arr1(i, 1)) = "W" Then w = w + 1
                If UCase(Arr1(i, 1)) = "M" Then m = m + 1
            'Nation zählen
                N = UCase(Arr1(i, 5))
                c = InStr(1, s, "|" + N + "|")
                If c = 0 Then
                    s = s + "0001|" + N + "|" + vbCrLf
                Else
                    r = CInt(Mid(s, c - 4, 4)) '0012
                    s = Left(s, c - 5) + Format(r + 1, "0000") + Mid(s, c)
                End If
        Next
    'Nation-Liste
        s = Mid(s, 4): s = Left(s, Len(s) - 2)
        Arr2 = Split(s, vbCrLf)
        QuickSort Arr2 ': s = Join(arr2, vbCrLf): show s
        '4 Nationen, highest score
            For i = UBound(Arr2) To UBound(Arr2) - 3 Step -1
                A = Arr2(i)
                u = u + CStr(CInt(Left(A, 4))) + " " + Mid(A, 6, Len(A) - 6) + ", "
            Next
        u = u + "...)"  'show u
        ' = 605 D, 61 USA, 61 GB, 40 SU, ...), 29 Nationen
        T = CStr(zLast - 7) + " Namen (" + CStr(w) + " w, " + CStr(m) + " m), " _
            + CStr(UBound(Arr2) + 1) + " Nationen ("
        ' = 1144 Namen (538 w, 606 m |
        .Cells(4, 2) = T + u
    'Finals
        End With
End Sub

Sub T5_SortUp(Sortierbereich$, Sortierspalte$)
    With ThisWorkbook.Worksheets("T5")
        .Range(Sortierbereich).Sort Key1:=.Range(Sortierspalte), Order1:=xlAscending, _
        Header:=xlNo, orderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    End With
End Sub

Sub T5_SortDown(Sortierbereich$, Sortierspalte$)
    With ThisWorkbook.Worksheets("T5")
        .Range(Sortierbereich).Sort Key1:=.Range(Sortierspalte), Order1:=xlDescending, _
        Header:=xlNo, orderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    End With
End Sub

Sub TEST_Create_AllVhFoldersOfOnePerson()
    DoArrc
    'Create_AllVhFoldersOfOnePerson "Mai, Ute (vh Jun Jul Aug, TV Bonn)"
    Create_AllVhFoldersOfOnePerson "Ehlig, Heidrun (vh Flöß Schrage, TSG Bruchsal)"
    'Create_AllVhFoldersOfOnePerson "Harz, Ute (vh Oder, TSG Wiesloch)"
    'Create_AllVhFoldersOfOnePerson "Rother, Christiane (vh Schnierda, rn Goldie, TV Unterbach)"
    'Create_AllVhFoldersOfOnePerson "Czech, Ute (vh Latton Luxon Pitkamin, TGJ Salzgitter)"
End Sub

Sub Create_AllVhFoldersOfOnePerson(NameOfOnePersonalFolder$)
    Dim i%, ArrVh() As String
    'Anzahl VhFolderNames = UBound(ArrVh)
    Load_vhFolderNames ArrVh, NameOfOnePersonalFolder
    'showArray ArrVh
    For i = 1 To UBound(ArrVh)
        Create_vhFolderAndLink ArrVh(i), NameOfOnePersonalFolder
    Next
End Sub

Function T5_Get_AllClubs()
    'Called from    Change_AllLinksOfOneFolder_ToVbsLinks
    
    'Vorbereitung
        Dim B$, s$, i%, ArrClubs() As String
    'Action
        T5_Load_Club ArrClubs
        s = "|"
        For i = LBound(ArrClubs) To UBound(ArrClubs)
            B = Trim(ArrClubs(i))
            If B Like "[A-Z]*" And Len(B) > 3 Then
                If Not s Like "*|" + B + "|*" Then s = s + B + "|"
            End If
        Next
        If Left(s, 1) = "|" Then s = Mid(s, 2)
        If Right(s, 1) = "|" Then s = Left(s, Len(s) - 1)
        ArrClubs = Split(s, "|")
        QuickSort ArrClubs ': showArray ArrClubs
        s = "|" + Join(ArrClubs, "|") + "|" ': show s
    'Finals
        T5_Get_AllClubs = s
End Function

Sub Load_vhFolderNames(ArrVh, NameOfOnePersonalFolder$)
    'Called from:   Create_AllVhFoldersOfOnePerson
    'ArrVh          = 1D-StringArray; zunächst leer; wird hier gefüllt
    'NameOf...      = "Mai, Ute (vh Jun Jul Aug, TV Bonn)" oder "Mai, Ute (vh Jun)"
    'gebOrdner      = PersonenOrdner, dessen Namen mit dem Geburtsnamen beginnt
    '               = NameOfOnePersonalFolder = HauptOrdner; enthält alle jpgs, ...
    'vhOrdner       = Ordner in 'Leute', dessen Namen mit einem vh-Namen beginnt
    'Action         lädt alle vHFolder-Namen in das Array ArrVh
    
    'Vorbereitung
        Dim Gn$, N$, N1$, N2$, N3$, p$, p1$, p2$, p3$, VnNn$, v$, vh$, Vn$
        Dim A%, C1%, C2%, Arr1() As String
        N = NameOfOnePersonalFolder: v = vbCrLf: p = ArrC(4) + "\": DoArrc
        If Not N Like "*(vh *" Then Exit Sub
        If Not FolderExists(p + N) Then
            show "Create_AllVhFoldersOfOnePerson" + v + v _
            + "Der Ordner '" + p + N + "' existiert nicht.": Exit Sub
            End If
    'vh ermitteln ("vh Jun", "vh Jun Jul Aug", ...)
        C1 = InStr(1, N, ",")
        C2 = InStr(1, N, "(")
        Gn = Left(N, C1 - 1)                'Geburtsname
        Vn = Mid(N, C1 + 2, C2 - C1 - 3)    'Vorname
        vh = Mid(N, C2 + 1): C1 = InStr(1, vh, ",")
        If C1 > 0 Then VnNn = Mid(vh, C1): VnNn = Left(VnNn, Len(VnNn) - 1) '", TV Bonn"
        If C1 = 0 Then vh = Left(vh, Len(vh) - 1) Else vh = Left(vh, C1 - 1)
    'Namen weiterer PersonenOrdner ermitteln
        Arr1 = Split(vh, " ")
        A = UBound(Arr1)        'a = Anzahl zusätzlicher Nachnamen
        N1 = Arr1(1)
    'a = 1      (1 zusätzlicher Nachname)
        If A = 1 Then       '"Mai, Ute (vh Jun, TV Bonn)"
            p1 = N1 + ", " + Vn + " (geb " + Gn + VnNn + ")"
            '  = "Jun, Ute (geb Mai, TV Bonn)"              = Name of PersonalFolder1
            ArrVh = Split("|" + p1, "|")
        End If
    'a = 2      (2 zusätzliche Nachnamen)
        If A = 2 Then       '"Mai, Ute (vh Jun Jul, TV Bonn)"
            N2 = Arr1(2)
            p1 = N1 + ", " + Vn + " (geb " + Gn + ", vh " + N2 + VnNn + ")"
            '  = "Jun, Ute (geb Mai, vh Jul, TV Bonn)"      = Name of PersonalFolder1
            p2 = N2 + ", " + Vn + " (geb " + Gn + ", vh " + N1 + VnNn + ")"
            '  = "Jul, Ute (geb Mai, vh Jun, TV Bonn)"      = Name of PersonalFolder2
            ArrVh = Split("|" + p1 + "|" + p2, "|")
        End If
    'a = 3      (3 zusätzliche Nachnamen)
        If A = 3 Then       '"Mai, Ute (vh Jun Jul Aug, TV Bonn)"
            N2 = Arr1(2): N3 = Arr1(3)
            p1 = N1 + ", " + Vn + " (geb " + Gn + ", vh " + N2 + " " + N3 + VnNn + ")"
            '  = "Jun, Ute (geb Mai, vh Jul Aug, TV Bonn)"      = Name of PersonalFolder1
            p2 = N2 + ", " + Vn + " (geb " + Gn + ", vh " + N1 + " " + N3 + VnNn + ")"
            '  = "Jul, Ute (geb Mai, vh Jun Aug, TV Bonn)"      = Name of PersonalFolder2
            p3 = N3 + ", " + Vn + " (geb " + Gn + ", vh " + N1 + " " + N2 + VnNn + ")"
            '  = "Jul, Ute (geb Mai, vh Jun Aug, TV Bonn)"      = Name of PersonalFolder2
            ArrVh = Split("|" + p1 + "|" + p2 + "|" + p3, "|")
        End If
End Sub

Sub Create_vhFolderAndLink(p1$, N$)
    Dim p$: p = ArrC(4) + "\"
    If Not FolderExists(p + p1) Then CreateFolder p + p1
    Create_OneLinkFile p + p1, p + N, N
End Sub

Sub Save_T5CopyOfPersonData()
    'Called from:   [none]
    
    'Vorbereitung
        Dim F$, p$, T$, ArrT5()
        Call DoArrc: F = "PersonData " + Format(Now(), "yyyymmdd_hhmmss") + ".txt"
        p = ArrC(1) + "\prog\old\PersonData\" + F
    'Fill ArrT5
        ArrT5 = Range(Cells(Get_T5_zFirst, Get_T5_sFirst), Cells(Get_T5_zLast, Get_T5_sLast))
    'Write
        T = getTextFrom2DArray(ArrT5)
        writeStringToFile p, T
    'Log
        LogBuch "T5CopyOfPersonData was saved: " + F
    Beep '800, 100
End Sub

Sub Once_UpDate_T5_L1()
    'Called from    [None] 1x-Anwendung
    
    'Vorbereitung
        Dim Club$, L1$, LTV$, m$, Nation$, rn$, vh$, sL1%, sRuf%, sVh%, z%, zFirst%, ArrT5()
        zFirst = Get_T5_zFirst: sL1 = Get_T5_sL1: sRuf = Get_T5_sRuf: sVh = Get_T5_sVh
        With Sheets("T5")
    'Fill ArrT5
        ArrT5 = Range(Cells(zFirst, Get_T5_sFirst), Cells(Get_T5_zLast, Get_T5_sLast))
        'showArray2D ArrT5
    'Search in ArrT5
        For z = 1 To UBound(ArrT5, 1)
            L1 = ArrT5(z, sL1 - 2)
            If L1 = "-" Then
                '[vh Mai/Juni, 'Spatz'] vh-, rn-Eintrag in Klammern
            'vh
                vh = ArrT5(z, sVh - 2)
                If vh <> "-" Then L1 = "[vh " + vh + "]"
            'rn
                rn = ArrT5(z, sRuf - 2)
                If rn <> "-" Then
                    If InStr(1, L1, "[") > 0 Then
                        L1 = Replace(L1, "]", ", '" + rn + "']")
                    Else
                        L1 = "['" + rn + "']"
                    End If
                End If
                'L1 = "-", "[vh Meier]", "['Spatz']" oder "[vh Meier, 'Spatz']"
            'Club LTV Nation
                Club = ArrT5(z, 5): LTV = ArrT5(z, 6): Nation = ArrT5(z, 7)
                If Club = "" Then m = "0" Else m = "1" 'm Marker
                If LTV = "" Then m = m + "0" Else m = m + "1"
                If Nation = "D" Then m = m + "D" Else m = m + "N"
                'm= 00D 00N 01D 01N 10D 10N 11D 11N
                If m = "00D" Or m = "00N" Then L1 = L1 + " " + Nation
                If m = "01D" Then L1 = L1 + " " + LTV
                If m = "01N" Then L1 = L1 + " " + LTV + "/" + Nation
                If m = "10D" Then L1 = L1 + " " + Club
                If m = "10N" Then L1 = L1 + " " + Club + "/" + Nation
                If m = "11D" Then L1 = L1 + " " + Club + " (" + LTV + ")"
                If m = "11N" Then L1 = L1 + " " + Club + "/" + LTV + "/" + Nation
                If Left(L1, 2) = "- " Then L1 = Mid(L1, 3)
                'L = L + L1 + vbCrLf
            'write
                .Cells(z + 7, sL1).Select
                .Cells(z + 7, sL1).HorizontalAlignment = xlLeft
                .Cells(z + 7, sL1) = L1
            End If
        Next
    'Finals
        End With ': show L
        Beep
End Sub

Sub Once_UpDate_T5_rn_gn_vh()
    'Called from    [None] 1x-Anwendung
    
    'Vorbereitung
        Dim FN$, GebN$, rn$, s$, vh$, i%, sGebN%, sRuf%, sVh%, zFirst%, ArrT5()
        zFirst = Get_T5_zFirst
        sGebN = Get_T5_sGebN
        sVh = Get_T5_sVh
        sRuf = Get_T5_sRuf
        With Sheets("T5")
    'Fill ArrT5
        ArrT5 = Range(Cells(zFirst, Get_T5_sFirst), Cells(Get_T5_zLast, Get_T5_sLast))
    'Search in ArrT5
        For i = 1 To UBound(ArrT5, 1)
            FN = CStr(ArrT5(i, 8)) 'FolderName
            If FN Like "*(*vh *" Then
                'Put gn
                    GebN = Left(FN, InStr(1, FN, ",") - 1)
                    .Cells(i + zFirst - 1, sGebN) = GebN
                'Put vh
                    vh = Get_T5_vhFromFolderName(FN)
                    .Cells(i + zFirst - 1, sVh) = vh
                'If Not s Like "*" + FN + "*" Then s = s + FN + vbCrLf + "   " + vh + vbCrLf
            End If
            If FN Like "*(*rn *" Then
                'Put rn
                    rn = Get_T5_rnFromFolderName(FN)
                    .Cells(i + zFirst - 1, sRuf) = rn
                'If Not s Like "*" + FN + "*" Then s = s + FN + vbCrLf + "   " + vh + vbCrLf
            End If
        Next
    'Finals
        End With
End Sub

Function Get_T5_rnFromFolderName(folderName$) As String
    'Called from    [None] 1x-Anwendung
    
    Dim FN$, rn$, i%, Arr1() As String
    FN = folderName: FN = Mid(FN, InStr(1, FN, "(") - 1): FN = Replace(FN, "(", ",")
    FN = Replace(FN, ")", ","): FN = Replace(FN, " rn", "rn")
    Arr1 = Split(FN, ",")
    For i = 1 To UBound(Arr1)
        If Left(Arr1(i), 3) = "rn " Then rn = rn + "/" + Mid(Arr1(i), 4)
    Next
    rn = Mid(rn, 2)
    Get_T5_rnFromFolderName = rn
End Function

Function Get_T5_vhFromFolderName(folderName$) As String
    'Called from    [None] 1x-Anwendung

    Dim FN$, vh$, i%, Arr1() As String
    FN = folderName: FN = Replace(FN, "(", ",")
    FN = Replace(FN, ")", ","): FN = Replace(FN, " vh", "vh")
    Arr1 = Split(FN, ",")
    For i = 1 To UBound(Arr1)
        If Left(Arr1(i), 3) = "vh " Then vh = vh + "/" + Mid(Arr1(i), 4)
    Next
    vh = Mid(vh, 2)
    Get_T5_vhFromFolderName = vh
End Function

Function Get_T5_zHeader() As Integer
    Get_T5_zHeader = Get_RowNr_HoldingMyTextWhole("T5", "Nachname")
End Function

Function Get_T5_zFirst() As Integer
    Get_T5_zFirst = Get_T5_zHeader + 2
End Function

Function Get_T5_sFirst() As Integer
    Get_T5_sFirst = Get_ColumnNr_HoldingMyTextWhole("T5", "Nachname")
End Function

Function Get_T5_zLast() As Long
    Dim sNr&
    sNr = Get_ColumnNr_HoldingMyTextWhole("T5", "Nachname")
    Get_T5_zLast = Get_NrOfLastRowInColumnNr(sNr, "T5")
End Function

Function Get_T5_sLast() As Long
    Dim RowNr%, i&, L&, sNr&, ws As Worksheet
    Set ws = Sheets("T5"): RowNr = Get_T5_zHeader
    L = ws.Cells(RowNr, ws.Columns.count).End(xlToLeft).Column
    For i = L To 1 Step -1
        If Trim(ws.Cells(RowNr, i)) <> "" Then sNr = i: Exit For
    Next
    Get_T5_sLast = sNr
End Function

Function Get_T5_sRuf() As Integer
    Get_T5_sRuf = Get_ColumnNr_HoldingMyTextPartInRowX("T5", Get_T5_zHeader, 1, "rn  Rufname")
End Function

Function Get_T5_sGebN() As Integer
    Get_T5_sGebN = Get_ColumnNr_HoldingMyTextPartInRowX("T5", Get_T5_zHeader, 1, "gn  Geburtsname")
End Function

Function Get_T5_sVh() As Integer
    Get_T5_sVh = Get_ColumnNr_HoldingMyTextPartInRowX("T5", Get_T5_zHeader, 1, "vh  verheiratet")
End Function

Function Get_T5_sL1() As Integer
    Get_T5_sL1 = Get_ColumnNr_HoldingMyTextPartInRowX("T5", Get_T5_zHeader, 1, "L1 Label")
End Function


