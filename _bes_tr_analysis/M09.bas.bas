Attribute VB_Name = "M09"
Option Explicit 'M09

Sub zzz_M09()
    'showProcs "t9"
    'RenameModule "Modul1", "M08"
    'T1_ShowLogBuch
    'T1_ShowPersonDataTxt
    EE 1: Beep
End Sub

Function Count_CellsLikeMyStringInArray2D(Arr, MyString$) As Integer
    Dim c%, i%, j%
    For i = 1 To UBound(Arr, 1)
        For j = 1 To UBound(Arr, 2)
            If Arr(i, j) Like MyString Then c = c + 1
        Next
    Next
    Count_CellsLikeMyStringInArray2D = c
End Function

Function Get_NnVn_LeuteFolderName(ArrPD1)
    'Called from    Update_T4CompetitorsList_LLinksInEFolder
    'Action         liefert Zeilen wie "Juni, Lea|Mai, Lea (vh Juni, TV Bonn)"
    
    Dim s$, i%
    For i = 1 To UBound(ArrPD1, 1)
        s = s + ArrPD1(i, 1) + ", " + ArrPD1(i, 2) + "|" + ArrPD1(i, 9) + vbCrLf
    Next
    Get_NnVn_LeuteFolderName = s
    'show s
End Function

Sub Update_LeuteLinks_inOneEventFolder(NameOfEventFolder$, PathOfEventFolder$, AllNnVn_OfOneDesign$, NnVn_LeuteFolderName$)
    'Called from    Update_T4CompetitorsList_LLinksInEFolder
    'AllNnVn_Of...  = Zeilen wie "Juni, Lea"   (Competitor of one Design)
    '               = Alle "Nachname, Vorname", aber nur dieses einzelnen Designs
    'NnVn_LeuteF..  = Zeilen wie "Juni, Lea|Mai, Lea (vh Juni, TV Bonn)" = Name + zugeh. LeuteOrdner
    'Action         xxx
    
    'Vorbereitung
        Dim L$, N$, NameOfNewLink$, NameOfOldLink$, PathOfLeuteFolder$, PL$
        Dim PointsTo$, r$, Report$, T$, v$
        Dim c%, i%, OK%, Arr1() As String
        v = vbCrLf: PL = ArrC(4) + "\"  'pL = Path of Folder "Leute"
                



        Report = "NameOfEventFolder = " + NameOfEventFolder + v + v
    'Action
        Arr1 = Split(v + AllNnVn_OfOneDesign, v)
        For i = 1 To UBound(Arr1)
            'Schleife über alle Competitors des aktuellen Designs
            N = Arr1(i)  'one CompetitorName
            ' = "Juni, Lea"
                If N = "" Then GoTo jump1
            L = Get_FirstLine_ContainingMyString(NnVn_LeuteFolderName, N + "|")
            ' = "Juni, Lea|Mai, Lea (vh Juni, TV Bonn)"
            If L <> "" Then
                T = ""
                PathOfLeuteFolder = PL + Replace(L, N + "|", "")
                'Get OneOldLink to delete in EventFolder
                    NameOfOldLink = Get_AllFileNames_Like_OfOneFolder(PathOfEventFolder, N + "*.lnk")
                    If NameOfOldLink Like "*" + v + "*" Then Stop 'mehrere OldLinks
                    Report = Report + "OldLink = '" + NameOfOldLink + "'" + v
                'Get OneNewLink to delete in EventFolder
                    c = InStr(L, "(")
                    If c = 0 Then
                        NameOfNewLink = N + ".lnk"
                        PointsTo = " '" + N + "'"
                    Else
                        T = Mid(L, c + 1)           'vh Juni, TV Bonn) 'USA)
                        T = Left(T, Len(T) - 1)     'vh Juni, TV Bonn  'USA
                        c = InStrRev(T, ", ")
                        If c > 0 Then               'T = Verein oder Nation
                            T = Mid(T, c + 2)       'TV Bonn
                        End If
                        NameOfNewLink = N + " (" + T + ").lnk"
                        PointsTo = " '" + Replace(L, N + "|", "") + "'"
                    End If
                    If NameOfNewLink Like "*.lnk.lnk" Then NameOfNewLink = Left(NameOfNewLink, Len(NameOfNewLink) - 4)
                'Delete/Create
                    If NameOfOldLink <> NameOfNewLink Then
                        r = Left("NewLink = '" + NameOfNewLink _
                            + "'" + String(55, " "), 55) + "points to" + PointsTo + v + v
                        Report = Report + r
'                        'To see Changes:
'                            OK = OK + 1
'                            If OK = 1 Then
'                                OpenFolder PathOfEventFolder
'                                Stop
'                            End If
                        'Delete
                            If NameOfOldLink <> "" Then DeleteFile PathOfEventFolder + "\" + NameOfOldLink
                        'Create
                            Create_OneLinkFile PathOfEventFolder, PathOfLeuteFolder, NameOfNewLink
                    End If
            End If
jump1:
        Next
    If OK > 0 Then show Report: Stop: CloseFolder PathOfEventFolder
End Sub

Function T4_GetNameOfResultJpg() As String
    'Called from    CmdCreate_Click
    'Status         Es soll jedenfalls ein neues result.jpg geschrieben werden,
    '               dessen FileName hier bereitgestellt werden soll
    'FileName       des result.jpg wird nach fixiertem Muster erzeugt:
    '               DgTitle + " #result " + Pid + " ub.jpg"
    '               z. B. "19590905 .Nissen-Cup2 Wasen (Einzel) #result p3394-00 ub.jpg"
    'DgTitle        ist bekannt; z. B. "1959-09-05 Nissen-Cup2 Wasen/CH (Einzel)"
    'Pid            wird übernommen, falls ein result.jpg bereits existiert,
    '               ansonsten neu erzeugt
    

    
    'Vorbereitung
        Dim N1$, N2$, N3$, p$, Pid$, s$, s2$, c%, C2%, i%, Arr1() As String
        N1 = ArrC(41) 'DgTitle
        N1 = Replace(N1, "/", "_")
'    'N2 = "[Title] #result ub.jpg" ohne Zusatzangabe Land
'        c = InStr(1, N1, "_")
'        If c > 0 Then        'Aus "Wasen_CH 4" wird "Wasen"
'            'alle Zeichen ab dem ersten "_" werden weggelassen
'            N2 = Left(N1, c - 1) + " #result "
'            'ggf. (Einzel) (Synchron) (Mannschaft) hinzu
'                C2 = InStr(1, N1, "(")
'                If C2 > 0 Then N2 = Replace(N2, "#", Mid(N1, C2) + " #")
'        Else
'            N2 = N1 + " #result "
'        End If
        N2 = N1 + " #result "
        'N2 = "1959-09-05 Nissen-Cup11 Grenchen_CH 5 (Einzel) #result "
    'N2 - Datum anpassen
        If N2 Like "#### *" Then N2 = Left(N2, 4) + "0000" + Mid(N2, 5)
        If N2 Like "####-## *" Then N2 = Left(N2, 4) + Mid(N2, 6, 2) + "00" + Mid(N2, 8)
        If N2 Like "####-##-## *" Then N2 = Left(N2, 4) + Mid(N2, 6, 2) + Mid(N2, 9, 2) + Mid(N2, 11)
        'N2 = "19590905 Nissen-Cup2 Wasen (Einzel) #result "
    'Add point (wegen Sortierung)
        N2 = Left(N2, 9) + "." + Mid(N2, 10)
        'N2 = "19590905 .Nissen-Cup2 Wasen (Einzel) #result "; Pid + " ub.jpg" fehlen noch
    'Pid
        'Falls ein result.jpg bereits existiert, beginnt dessen FileName mit N2
        p = ArrC(57) 'Path of Folder of actual Event
        If p = "" Then Stop
        N3 = Replace(N2, "#", "[#]") '# = Like-Operator-Ziffer-Zeichen
        s = Get_AllFileNames_Like_OfOneFolder(p, N3 + "*")
        If s = "" Then
            'ein result.jpg existiert noch nicht
            Pid = Get_NextFreePID
        Else
            'ein result.jpg existiert bereits
            Pid = Mid(s, Len(N2) + 1, 8)
            If Not Pid Like "p####-00" Then Stop
        End If
    'N2 - kompletter FileName
        N2 = N2 + Pid + " ub.jpg"
    'Finals
        T4_GetNameOfResultJpg = N2
End Function

Sub AddSomeValuesToArrDesign(ByRef ArrDesign())
    'Called from    Fill_ListOfCompetitors
    'Status         Ein bestimmtes Design ist ausgewählt, AnalyzeDesign-Daten vorhanden

    'Vorbereitung
        Dim Kopf$, Sp1$, Sp2$, SpNation$, SpRang$, SpVerein$, Titel$
        Dim i%, j%, r%, Arr1() As String
        SpRang = ArrC(47)           'T4 SpaltenNrn der Rang  -Spalten ',3,8,
        SpVerein = ArrC(51)         'T4 SpaltenNrn der Verein-Spalten ',3,8,
        SpNation = ArrC(50)         'T4 SpaltenNrn der Nation-Spalten

    'arrDesign - add in Sp1: 1 Buchstabe in Sp1 der jew. Zeile
        '(E|Y|M)=(Einzel|Synchron|Mannschaft)
        Titel = ArrDesign(2, 2)
        'Falls sich das Design auf einen reinen Mannschafts- oder SynchronWK bezieht:
            If InStr(1, Titel, " LK") > 0 Then Sp1 = "M"
            If InStr(1, Titel, "JLK") > 0 Then Sp1 = "M"
            If InStr(1, Titel, "Sy") > 0 Then Sp1 = "Y"
    'arrDesign - add in Sp1
    
    xxx
            If CStr(ArrC(52)) = "0" Then
                For i = 6 To UBound(ArrDesign, 1) - 1
                    'im aktuellen Design existiert keine Spalte mit "Einzel" oder "Synchron"
                    'also:  Sorte Rang-Name-Verein/Nation (Ergebnisse Einzel)
                    'oder:  Sorte Rang-Verein/Nation-Name (Ergebnisse Mannschaft)
                    If Sp1 = "" Then
                        Sp1 = "E"
                        If InStr(1, SpVerein, ",3,") > 0 Then Sp1 = "M"
                        If InStr(1, SpNation, ",3,") > 0 Then Sp1 = "M"
                        If InStr(1, Titel, "LK") > 0 Then Sp1 = "M"     'bei LK oder JLK
                        ArrDesign(i, 1) = Sp1
                    Else
                        ArrDesign(i, 1) = Sp1
                    End If
                Next
            End If
    'arrDesign - add in Sp1
            'Spalte mit "Einzel" oder "Synchron" existiert
            If CStr(ArrC(52)) = "2" Then 'SpTyp in Sp2 (Einz Syn)
                For i = 6 To UBound(ArrDesign, 1) - 1
                    Kopf = CStr(ArrDesign(4, 4))         'KopfZeile
                    Sp2 = CStr(ArrDesign(i, 2))         'Einz/Sy oder 1/2/3/...
                    If Sp2 <> "" Then
                        'Synchron (Mixed) --> Synchron
                            If Left(Sp2, 2) = "Sy" Then Sp2 = "Synchron": ArrDesign(i, 2) = Sp2
                        If Left(Sp2, 2) = "Ei" Then Sp1 = "E"
                        If Left(Sp2, 2) = "Sy" Then Sp1 = "Y"
                    End If
                    ArrDesign(i, 1) = Sp1
                Next
            End If
            
            
    'arrDesign - add in Sp Rang/Verein/Nation bei Synchron und Mannschaft
        If Len(SpRang) > 1 Then                 'SpRang = ",3," oder ",3,8,"
            Arr1 = Split(SpRang, ",")
            For j = 1 To UBound(Arr1) - 1
                If IsNumeric(Arr1(j)) Then
                
                
                    r = CInt(Arr1(j))                      'R = SpaltenNr einer Rang-Spalte
                    For i = 6 To UBound(ArrDesign, 1) - 1
                        'Synchron
                        If Left(ArrDesign(i, 1), 1) = "Y" Then
                            'Rang auffüllen in Zelle(i,R)
                            'Falls rechts neben (i,R) ein Name/... steht (keine leere Zelle),
                            'dann bekommt (i,R), wenn leer, den selben Rang der Zelle darüber
                            If ArrDesign(i, r + 1) <> "" Then
                                If ArrDesign(i, r) = "" Then ArrDesign(i, r) = ArrDesign(i - 1, r)
                            End If
                            'Nation auffüllen
                            If InStr(1, SpNation, r + 2) > 0 Then
                                If ArrDesign(i, r + 1) <> "" Then 'Zelle rechts daneben nicht leer
                                    If ArrDesign(i, r + 2) = "" Then ArrDesign(i, r + 2) = ArrDesign(i - 1, r + 2)
                                End If
                            End If
                        End If
                        'Mannschaft
                        If Left(ArrDesign(i, 1), 1) = "M" Then
                            'Rang auffüllen
                            If ArrDesign(i, r + 2) <> "" Then
                                If ArrDesign(i, r) = "" Then ArrDesign(i, r) = ArrDesign(i - 1, r)
                            End If
                            'Verein/Nation auffüllen
                            If InStr(1, SpVerein, r + 1) > 0 Or InStr(1, SpNation, r + 1) > 0 Then
                                If ArrDesign(i, r + 2) <> "" Then 'Zelle rechts daneben nicht leer
                                    If ArrDesign(i, r + 1) = "" Then ArrDesign(i, r + 1) = ArrDesign(i - 1, r + 1)
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        End If
End Sub

Sub SearchForDesignBordersQuick(z%, s%)
    'Called from    xxx
    
    'Vorbereitung
        Dim s1%, s2%, z1%, z2%
    'Search for borders
        z1 = Get_RowNr_OfNextBorder_Top(z, s)
        z2 = Get_RowNr_OfNextBorder_Bottom(z, s)
        s1 = Get_ColNr_OfNextBorder_Left(z, s)
        s2 = Get_ColNr_OfNextBorder_Right(z, s)
        ArrC(41) = Cells(z1 + 1, s1 + 1) 'Title
        ArrC(36) = Get_ColumnLetter(s1) + CStr(z1) + ":" + Get_ColumnLetter(s2) + CStr(z2) 'Range
        ArrC(37) = CStr(z1): ArrC(38) = CStr(z2): ArrC(39) = CStr(s1): ArrC(40) = CStr(s2) 'Nrs
End Sub

Sub xxx()
    Application.EnableEvents = True: DoArrc
End Sub

Sub T4_AutoFit()
    'Called from    T4_ActionsOnDgClickNext, HelperBox_Click
    
    'Vorbereitung
        Dim B$, i%, s1%, s2%, ArrBreiten(), Arr1() As String, ws As Worksheet, rg As Range
        s1 = CInt(ArrC(39)): s2 = CInt(ArrC(40))
        Set ws = Sheets("T4")
    'ArrBreiten laden
        Set rg = ws.Range(ws.Cells(3, s1), ws.Cells(3, s2))
        ArrBreiten = rg.Value   '|2|3|0|0|4|5|5|...
    'SpaltenBreiten der Dg-Spalten 1,2,3,4,...
        For i = 1 To UBound(ArrBreiten, 2)
            T4_AutoFitCalc i, CSng(ArrBreiten(1, i))
        Next
End Sub

Sub T4_AutoFitCalc(SpNrToFit%, Optional FixWidth! = 0)
    'Called from    T4_AutoFit
    'SpNrToFit      = 6     'geeignete Breite ermitteln für DgSpalte 6
    'FixWidth       = 5     'Breite ist bereits festgelegt; soll 5 sein
    '               = 0     'Breite wird berechnet gem. Spalteninhalte
    
    'Vorbereitung
        Dim s1%, w!
        s1 = CInt(ArrC(39))
    'Action
        If FixWidth = 0 Then 'kein fester Wert angegeben
            'Texte der aktuellen Spalte holen
            w = Get_WidthOfOneColumn_forAutoFit(ArrDg, SpNrToFit)
        Else
            w = FixWidth
        End If
    'Ermittelte Breite auf SpalteToFit anwenden
        Columns(s1 - 1 + SpNrToFit).ColumnWidth = w
End Sub

Function Get_WidthOfOneColumn_forAutoFit(ArrDesign, SpNrToFit%) As Single
        'Called from    T4_AutoFitCalc
        'AFS            AutoFitSpalte = Hilfsspalte in T8
        
        'Vorbereitung
            Dim arrTxt() As String, Arr1() As String
            Dim A$, s$, T$, AFS%, i%, j%, w!
            AFS = Get_ColumnNr_HoldingMyTextWhole("T8", "AutoFitSpalte")
        'Text to ignore
            A = ArrC(34) + ",Schüler,klasse,Juti,Jutu,Damen,Herren,Jugend"
            Arr1 = Split(A, ",")
        'ZellTexte der aktuellen Spalte sammeln
            For i = 6 To UBound(ArrDesign, 1) - 1
                T = CStr(ArrDesign(i, SpNrToFit)) + "|" 'Text einer Zelle
                T = Replace(T, vbLf, "|")   'Return innerhalb einer Zelle
                For j = 0 To UBound(Arr1)
                    If T Like "*" + Arr1(j) + "*" Then T = "": j = UBound(Arr1)
                Next
                s = s + T
            Next
            arrTxt = Split(s, "|")
        'ZellTexte in die AutoFitSpalte schreiben
            Paste_1DArrayToCol "T8", 5, AFS, arrTxt
            With Sheets("T8")
                .Columns(AFS).EntireColumn.AutoFit 'TestSpalte
                w = .Columns(AFS).ColumnWidth
                .Range(.Cells(5, AFS), .Cells(5 + UBound(arrTxt), AFS)).ClearContents
                .Columns(AFS).EntireColumn.AutoFit 'TestSpalte
            End With
        'Finals
            Get_WidthOfOneColumn_forAutoFit = w
End Function

Function Get_ZeSp_OfDesignTitleNotFolderName(Titel$) As String
        Dim z%, s%, z2%, s2%
        z = Get_RowNr_HoldingMyTextWhole("T4", Titel)
        s = Get_ColumnNr_HoldingMyTextWhole("T4", Titel)
        If z = 0 Then Stop 'Titel existiert nicht
        EE 0
        If Trim(Cells(z + 1, s)) <> "" Then
            'Steht eine Zelle tiefer ein Text, so ist man bei einem FolderName gelandet,
            ' nicht bei einem DesignTitel;
            ' deshalb muss kurzzeitig der FolderName entfernt und neu gesucht werden
            Cells(z, s) = ""
            z2 = Get_RowNr_HoldingMyTextWhole("T4", Titel)
            s2 = Get_ColumnNr_HoldingMyTextWhole("T4", Titel)
            Cells(z, s) = Titel
            If z2 > z Then z = z2: s = s2
        End If
        Get_ZeSp_OfDesignTitleNotFolderName = Format(z, "000") + Format(s, "000")
End Function

Function Get_ColNr_OfNextBorder_Right(z%, s%) As Integer
    Dim i%, s2%, ws As Worksheet
    Set ws = Sheets("T4")
    s2 = lastCol(ws) + 10
    With Sheets("T4")
    For i = s To s2
        If .Cells(z, i).Borders(xlEdgeRight).LineStyle <> xlLineStyleNone _
            And .Cells(z, i).Borders(xlEdgeRight).Color = vbBlack Then
            Get_ColNr_OfNextBorder_Right = i
            Exit For
        End If
    Next
    End With
End Function

Function Get_ColNr_OfNextBorder_Left(z%, s%) As Integer
    Dim i%
    With Sheets("T4")
    For i = s To 1 Step -1
        If .Cells(z, i).Borders(xlEdgeLeft).LineStyle <> xlLineStyleNone _
            And .Cells(z, i).Borders(xlEdgeLeft).Color = vbBlack Then
            Get_ColNr_OfNextBorder_Left = i
            Exit For
        End If
    Next
    End With
End Function

Function Get_RowNr_OfNextBorder_Top(z%, s%) As Integer
    Dim i%
    With Sheets("T4")
    For i = z To 1 Step -1
        If .Cells(i, s).Borders(xlEdgeTop).Color = vbBlack _
            And .Cells(i, s).Borders(xlEdgeTop).LineStyle <> xlLineStyleNone Then
            Get_RowNr_OfNextBorder_Top = i
            Exit For
        End If
    Next
    End With
End Function

Function Get_RowNr_OfNextBorder_Bottom(z%, s%) As Integer
    Dim i%, z2%, ws As Worksheet
    Set ws = Sheets("T4")
    z2 = LastRow(ws) + 10
    With Sheets("T4")
    'LastRow
    For i = z To z2
        If .Cells(i, s).Borders(xlEdgeBottom).Color = vbBlack _
            And .Cells(i, s).Borders(xlEdgeBottom).LineStyle <> xlLineStyleNone Then
            Get_RowNr_OfNextBorder_Bottom = i
            Exit For
        End If
    Next
    End With
End Function




