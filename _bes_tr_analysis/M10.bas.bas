Attribute VB_Name = "M10"
Option Explicit 'M10

Sub zzz_M10()
    'showProcs "event"
    Application.EnableEvents = True
    'RenameModule "Tabelle1", "T6"
    'T1_ShowLogBuch
    'showArray2D ArrPd
    Beep
End Sub

Sub TEST_SortUp_Range_Column()
    SortUp_Range_Column "T8", 7, 0, 10, 13, 10
End Sub

Sub SortUp_Range_Column(NameOfSheet$, z1%, z2_Or_0_means_zlast%, s1%, s2%, SortColumn%)
    With ThisWorkbook.Worksheets(NameOfSheet)
        Dim z2: z2 = z2_Or_0_means_zlast
        If z2 = 0 Then z2 = Get_NrOfLastRowInColumnNr(CLng(s1), NameOfSheet)
        .Range(.Cells(z1, s1), .Cells(z2, s2)).Sort Key1:=.Cells(z1, SortColumn), _
        Order1:=xlAscending, Header:=xlNo, orderCustom:=1, MatchCase:=False, _
        Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    End With
End Sub

Sub SortDown_Range_Column(NameOfSheet$, z1%, z2_Or_0_means_zlast%, s1%, s2%, SortColumn%)
    With ThisWorkbook.Worksheets(NameOfSheet)
        Dim z2: z2 = z2_Or_0_means_zlast
        If z2 = 0 Then z2 = Get_NrOfLastRowInColumnNr(CLng(s1), NameOfSheet)
        .Range(.Cells(z1, s1), .Cells(z2, s2)).Sort Key1:=.Cells(z1, SortColumn), _
        Order1:=xlDescending, Header:=xlNo, orderCustom:=1, MatchCase:=False, _
        Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    End With
End Sub

Sub T6_Write_T6_DoDoneRemarks(NameOfButton$, CountOfChanges%, Optional Extras$ = "")
    'Called from    xxx
    
    'Vorbereitung
        Dim v$, zT6%, r As Range
        v = vbCrLf: With Sheets("T6")
    'Action
        zT6 = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 2, 7, NameOfButton)
        'LastDone
            Set r = .Cells(zT6, 18)
            r.Value = Format(Now(), "yyyymmdd_hhmmss")
            r.Font.size = 6: r.HorizontalAlignment = xlLeft
        'Days ago
            Set r = .Cells(zT6, 19)
            r.Value = 0: r.Font.size = 8: r.HorizontalAlignment = xlCenter
        'Count of changes
            Set r = .Cells(zT6, 20)
            r.Value = CountOfChanges: r.Font.size = 8: r.HorizontalAlignment = xlCenter
        'Erledigt-Häkchen setzen
            .Cells(zT6, 16) = "ü"
        'Duration
            Set r = .Cells(zT6, 21)
            r.Value = T6_GetDuration(): r.Font.size = 8: r.HorizontalAlignment = xlCenter
    'Extras
        If NameOfButton = " Update DgPositions in T8SomeDgData" Then
            'Extras = "|"+CStr(ShowList)+"|"+CStr(c1)+"|"+CStr(c2)+"|"+Pos1+"|"+Pos2+"|"
            Dim sList$, Arr1() As String: Arr1 = Split(Extras, "|")
            'showList
                If Arr1(1) = "1" Then
                    If Arr1(5) = "" Then Arr1(5) = "-/-"
                    sList = "Info zu 'Update DgPositions in T8SomeDgData'" + v _
                          + "Titles von ListOfDesignTitles und T8SomeDgData stimmen überein." + v + v _
                          + Arr1(2) + " Positionsdaten korrekt:" + v + v + Arr1(4) + v _
                          + Arr1(3) + " Positionsdaten erneuert:" + v + v + Arr1(5)
                    show sList
                End If
            End If
    'Finals
        .Activate
        End With
End Sub


Sub T4_Show_DataLineOfDgCompetitors_MissingAge()
    'Called from    [None]
    'Status         Alter eines Aktiven wird bei Dg-Aktivierung ergänzt;
    '               Positionsdaten in T8SomeDgData müssen aktuell sein
    'Action         Ohne Dg-Aktivierung:
    '               Listet DataZeilen (Name, Event, ...) zu jedem Dg-Namen,
    '               bei dem trotz bekanntem Jhg kein Alter vermerkt ist
    
    'Vorbereitung
        Dim DgTitle$, folder$, iJhg$, L1$, Nn$, s1$, s2$, s3$, s4$
        Dim searchT8$, sJhg$, sNn$, sVn$, v$, Vn$, VnNn$
        Dim c%, i%, Jhg%, s%, s1Sh%, sDg%, z%, z1Sh%, zDg%
        Dim T5(), T8(), Arr1() As String, L() As String
        Dim ArrJhg() As String, ArrNn() As String, ArrVn() As String
        v = vbCrLf
    '2D-Array T8 (T8SomeDgData)
        T8_Load_ArrT8SomeDgData T8
    '2D-Array T5
        T5_Load_Nn_to_Jhg T5
        For i = 1 To UBound(T5, 1)
            If T5(i, 4) Like "####" Then
                sNn = sNn + v + T5(i, 1)        'Sammlung Nachnamen
                sVn = sVn + v + T5(i, 2)
                sJhg = sJhg + v + CStr(T5(i, 4))
            End If
        Next
    '1D-Arrays ArrNn, ArrVn, ArrJhg
        ArrNn = Split(sNn, v): ArrVn = Split(sVn, v): ArrJhg = Split(sJhg, v)
        'showArray ArrNn: showArray ArrVn
    'Suche
        For i = 1 To UBound(ArrNn)
            'Schleife über alle NamenMitJhg aus T5
            searchT8 = "#" + ArrNn(i) + "|" + ArrVn(i) + "|" 'SearchString to search in T8
            iJhg = CStr(ArrJhg(i))
            s1 = T6_Find_MissingAgeData_InT8SomeDgData(T8, searchT8, iJhg)
            If s1 <> "" Then s2 = s2 + s1 + v
            'CountDown
                Sheets("T8").[M4] = UBound(ArrNn) - i
            DoEvents
        Next
    'Ergebnisliste s2
        s2 = Delete_EmptyRowsInString(s2)
        s2 = Replace(s2, "#", "")
        'PepUp-Anzeige s3
            If s2 = "" Then
                show "Update_Age_inDgs" + v + v + "Nothing to update."
            Else
                c = InStr(1, s2, v)
                If c = 0 Then
                    s3 = s2 's2 trägt nur 1 Zeile
                Else
                    s3 = Left(s2, InStr(1, s2, v) - 1) '1. Zeile von s2
                End If
                For i = 1 To anzAinB("|", s3): s4 = s4 + CStr(i) + "|": Next
                s3 = s4 + v + s2
                show TAB_Simulation(s3)
            End If
'    'Add Age
'        Arr1 = Split(v + s2, v)
'        With Sheets("T4")
'        For i = 1 To UBound(Arr1)
'            L1 = Arr1(i) 'one Line
'            L = Split("|" + L1, "|")
'            Nn = L(1): Vn = L(2): zDg = CInt(L(3)): sDg = CInt(L(4)): Jhg = CInt(L(14))
'            DgTitle = L(16): Folder = L(18): z1Sh = CInt(L(21)): s1Sh = CInt(L(25))
'            VnNn = Vn + " " + Nn: z = z1Sh + zDg - 1: s = s1Sh + sDg - 1
'
'
'            .Select: .Cells(z, s).Select
'            If .Cells(z, s) = VnNn Then
'                Stop
'            Else
'                Stop
'            End If
'
'
'        Next
'        End With
End Sub

Function T6_Find_MissingAgeData_InT8SomeDgData(T8, searchT8$, Jhg$) As String
    'Called from    Update_Age_inDgs
    'searchT8       = "#Nn|Vn|" 'SearchString to search in T8 'Namen mit JhgExists
    'Action         Listet eine DataZeile (Name, Event, ...) zu jedem Dg-Namen,
    '               bei dem trotz bekanntem Jhg kein Alter vermerkt ist
    
    'Vorbereitung
        Dim s1$, s2$, s3$, s4$, AfterRowNr%, zT8%, C1&, C2&, C3&
    'Action
        AfterRowNr = 6: zT8 = 1
        Do While zT8 > 0
            C3 = C3 + 1
            zT8 = Get_RowNr_HoldingMyTextPartInColumnX("T8", 13, AfterRowNr, searchT8)
            If s4 Like "*|" + CStr(zT8) + "|*" Then Exit Do Else s4 = s4 + "|" + CStr(zT8) + "|"
            If zT8 > 0 Then
                'Zeile mit gesuchtem Namen in T8SomeDgData gefunden
                s1 = T8(zT8 - 6, 4) 'Zelle mit allen Namen
                C1 = InStr(1, s1, searchT8): C2 = InStr(C1 + 1, s1, "#")
                s2 = CutVonBis(s1, C1, C2) 'Daten 1 Person
                'Nur übernehmen, falls kein Alter vermerkt ist
                If s2 Like "*|####||*" Or s2 Like "*|w|||*" Or s2 Like "*|m|||*" Then
                    'Jhg existiert in T5, ist aber nicht in T8 berücksichtigt
                    s3 = s3 + s2 + "|Jhg=|" + Jhg + "|DgTitle=|" + Trim(T8(zT8 - 6, 1)) _
                       + "|Folder=|" + T8(zT8 - 6, 2) + "|POS|" + T8(zT8 - 6, 3) + "|" + vbCrLf
                    DoEvents
                End If
            End If
            AfterRowNr = zT8 + 1 ': zT8 = 0
            DoEvents
        Loop
        T6_Find_MissingAgeData_InT8SomeDgData = s3
End Function

Sub T6_Update_CompetitorsList(z%)
    'Called from    T6_UserClickOnColumn2CellButton
    
    T3_Create_CompetitorsList
    'Nach Erledigung wieder zurück
        With Sheets("T6"): .Activate: .Cells(z, 1).Select: End With
    Beep
End Sub

Sub T6_Remove_PrivatTags()
    'Vorbereitung
        Dim E1$, E2$, E3$, PathOfFolder$, v$
        Dim c%, i%, RowNr%, u%, t1&, E() As String, r As Range
        RowNr = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 2, 7, " Remove privat tags in JPGs")
        With Sheets("T6"): v = vbCrLf: DoArrc
    'StartTime
        t1 = GetTickCount()
    'Erledigt-Häkchen löschen
        .Cells(RowNr, 16) = ""
    'CountDown
        Set r = .Cells(RowNr, 18): r.Value = "[999]": r.Font.size = 8: r.HorizontalAlignment = xlCenter
    'Days ago
        Set r = .Cells(RowNr, 19): r.Value = 0: r.Font.size = 8: r.HorizontalAlignment = xlCenter
    'Count of changes (Anzahl geänderter T5-Zeilen)
        Set r = .Cells(RowNr, 20): r.Value = c: r.Font.size = 8: r.HorizontalAlignment = xlCenter
    'Time needed
        Set r = .Cells(RowNr, 21): r.Value = "0,0 sec": r.Font.size = 8: r.HorizontalAlignment = xlCenter
    'NameOfFolder
        Set r = .Cells(RowNr, 23): r.Value = "Folder": r.Font.size = 8: r.HorizontalAlignment = xlLeft
    'FolderPaths
        E1 = Get_Paths_ofAllSubfoldersAllLevelsAsStringUseGlobalVar(ArrC(3))   'Folder "Events"
        'show E1
        .Cells(RowNr, 18) = "[998]"
        E2 = Get_Paths_ofAllSubfoldersAllLevelsAsStringUseGlobalVar(ArrC(11))  'Folder "ClubsNations"
        .Cells(RowNr, 18) = "[997]"
        E3 = Get_Paths_ofAllSubfoldersAllLevelsAsStringUseGlobalVar(ArrC(4))   'Folder "Leute"
        .Cells(RowNr, 18) = "[996]"
    'Array E
        E = Split(v + E1 + v + E2 + v + E3, v)
        'E = Split(v + "F:\Archiv TR\Archiv Trampolin\Events\1973 Liga\1973-12 DMM BuliE Bruchsal")
        u = UBound(E)
        For i = 1 To u
            'CountDown
                .Cells(RowNr, 18) = "[" + CStr(u - i) + "]"
            'Folder
                PathOfFolder = E(i)
                .Cells(RowNr, 23) = Get_NameFromPath(PathOfFolder)
                If T6_FolderHasNonResultJpgs(PathOfFolder) Then
                    'Remove (if necessary)
                        T6_Remove_PrivatTagsFromJPGsOfOneFolder PathOfFolder, c, t1
                End If
            DoEvents
        Next
    'LastDone
        Set r = .Cells(RowNr, 18): r.Value = Format(Now(), "yyyymmdd_hhmmss"): r.Font.size = 6: r.HorizontalAlignment = xlLeft
    'Erledigt-Häkchen setzen
        .Cells(RowNr, 16) = "ü"
    'Finals
        .Cells(RowNr, 23) = ""
        End With
        Beep
End Sub

Function T6_FolderHasNonResultJpgs(PathOfFolder$) As Boolean
    'Called from    T6_Remove_PrivatTags
    Dim s$
    s = Get_AllFileNames_Like_OfOneFolder(PathOfFolder, "*.jpg")
    If Not s Like "*.jpg*" Then Exit Function
    If s Like "*#result*" Then Exit Function
    T6_FolderHasNonResultJpgs = True
End Function

Sub Ticks1()
    DoArrc
    FillArrC 35, CStr(GetTickCount())
End Sub

Function T6_GetDuration() As String
    Dim T$, t1&, tDiff&
    t1 = CLng(ArrC(35))
    tDiff = GetTickCount() - t1
    If tDiff < 60000 Then
        T = Format(CDbl(tDiff / 1000), "0.0") & " sec"
    Else
        T = Format(CDbl(tDiff / 60000), "0.0") & " min"
    End If
    T6_GetDuration = T
End Function

Sub T6_Remove_PrivatTagsFromJPGsOfOneFolder(PathOfSourceFolder$, c%)
    'Vorbereitung
        Dim s1$, s2$, T$, C1%, C3%, C5%, RowNr%, tDiff&
        RowNr = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 2, 7, " Remove privat tags in JPGs")
        With Sheets("T6")
    'Text1
        s1 = T6_Get_TextWithCertainTagOfJpgs(PathOfSourceFolder, 1)
        'show s1
        C1 = anzAinB(":/", s1)
        C3 = anzAinB("|-", s1)
        C5 = C1 - C3            'Anzahl SourceFolder-JPGs mit PrivatTags
    'Duration
        .Cells(RowNr, 21) = T6_GetDuration()
    'NoRemove
        If C5 = 0 Then
            'keine JPGs mit PrivatTags 'Remove nicht nötig
            Exit Sub
        End If
    'Remove
        T6_Remove_PrivatTagsFromJPGs PathOfSourceFolder
    'Text2
        s2 = T6_Get_TextWithCertainTagOfJpgs(PathOfSourceFolder, 2)
        'show s2
    'Compare 'Verify s1, s2
        C5 = T6_Get_CountOfJpgsWithPrivatTagsRemoved(s1, s2)
        c = c + C5
        .Cells(RowNr, 20) = c 'Count of changes
        .Cells(RowNr, 20).Font.size = 8
    'Duration
        tDiff = GetTickCount() - t1
        If tDiff < 60000 Then T = Format(CDbl(tDiff / 1000), "0.0") & " sec" Else T = Format(CDbl(tDiff / 60000), "0.0") & " min"
        .Cells(RowNr, 21) = T: .Cells(RowNr, 21).Font.size = 8
    'Finals
        End With
        'Beep
End Sub

Function T6_Get_CountOfJpgsWithPrivatTagsRemoved(s1$, s2$) As Integer
    'Called from    T6_Remove_PrivatTagsFromJPGsOfOneFolder
    's1             = Liste JpgPfade, je mit RegionAreaX-Tag  vor   der Remove-Aktion
    '               Shell/ExifTool schreibt die TextDatei "TagsInFolder_..._old.txt"
    '               in den Folder "Tags", die danach als s1 gelesen wird
    's2             = Liste JpgPfade, je mit RegionAreaX-Tag  nach  der Remove-Aktion
    '               dito "TagsInFolder_..._new.txt" (existiert ggf. nicht)
    
    Dim C1%, C2%, C3%, C4%, C5%
    C1 = anzAinB(":/", s1)
    C2 = anzAinB(":/", s2)
    If C1 <> C2 Then Stop
    C3 = anzAinB("|-", s1)
    C4 = anzAinB("|-", s2)
    
    
    If C2 <> C4 Then
        'c2 Pfade (:/), c4 leere PrivatTags (|-)
        'Folder...new.txt = Folder...old.txt
        'trat auf bei
        '  F:\Archiv TR\Archiv Trampolin\Events\1973 Liga
        '  F:\Archiv TR\Archiv Trampolin\Events\1973 Liga\1973-12 DMM BuliE Bruchsal
        show s1
        show s2
        Stop
    End If
    
    
    C5 = C1 - C3
    T6_Get_CountOfJpgsWithPrivatTagsRemoved = C5
End Function

Function T6_Get_TextWithCertainTagOfJpgs(PathOfSourceFolder$, OldNew%) As String
    'Vorbereitung
        Dim N$, p$, PathOfFileToWrite$, s$
    'PathOfFileToWrite
        N = Get_NameFromPath(PathOfSourceFolder)
        N = Replace(N, "-", ""): N = Replace(N, " ", "")
            '1978Liga '196011DMFrankfurt '19641112DMBerlin
        If Not Mid(N, 5, 1) Like "#" Then N = Left(N, 4) + "0000" + Mid(N, 5)
        If Not Mid(N, 7, 1) Like "#" Then N = Left(N, 6) + "00" + Mid(N, 7)
        p = ArrC(1) + "\Tags\TagsInFolder_" + N
        If OldNew = 1 Then PathOfFileToWrite = p + "_old.txt"
        If OldNew = 2 Then PathOfFileToWrite = p + "_new.txt"
    'Write textfile
        T6_Write_TextFile_MyTags PathOfSourceFolder, PathOfFileToWrite
        If FileExists(PathOfFileToWrite) Then s = ReadFile(PathOfFileToWrite)
        T6_Get_TextWithCertainTagOfJpgs = s
End Function

Sub TEST_T6_Remove_PrivatTagsFromJPGs()
    'Vorbereitung
        Dim p1$, t1&, tDiff&
        DoArrc
    'PathOfSourceFolder
        p1 = "F:\Archiv TR\Archiv Trampolin\Events\1973 Liga\1973-12 DMM BuliE Bruchsal"
    'StartTime
        't1 = GetTickCount()
    'Write textfile
        T6_Remove_PrivatTagsFromJPGs p1
    'Duration
        'tDiff = GetTickCount() - t1
        'show CStr(CDbl(tDiff / 1000)) & " seconds"
    'Finals
        Beep
End Sub

Sub T6_Remove_PrivatTagsFromJPGs(PathOfSourceFolder$)
    'Called from    T6
    'Action         Entfernt Namen von Personen aus einigen Tags
    '               bei allen JPGs des Ordners PathOfSourceFolder:
    'Tags           [XMP-dc]Subject [XMP-digiKam]TagsList [XMP-mwg-rs]Region*
    '               [XMP-microsoft]LastKeywordXMP [XMP-lr]HierarchicalSubject
    '               [XMP-mediapro]CatalogSets [IPTC]Keywords
    
    
    'Vorbereitung
        Dim myOptions$, IgnoreSomeJpg$, pExif$, pSource$, qq$, shellCmd$, shResult%
        qq = Chr(34)
        pExif = qq + ArrC(9) + qq + " "
        pSource = qq + PathOfSourceFolder + qq + " "
    'myOptions
        IgnoreSomeJpg = " -if " + qq + "$filename !~ /result/ and not $RegionAreaX eq ''" + qq + " "
        myOptions = IgnoreSomeJpg + "-L -p -ext jpg " _
                  + " -Region*= -*Subject= -CatalogSets= -*Keyword*= " _
                  + "-TagsList= -Categories= -overwrite_original"
    'ShellCmd erstellen
        shellCmd = pExif + pSource + myOptions
        '-XMP-acdsee-rs:all= -XMP-mwg-rs:all= -XMP-MP:all= -PersonInImage=
        'show ShellCmd
    'ShellCmd senden
        shResult = ShellAndWait(shellCmd, 0, vbHide, PromptUser)
        If shResult > 0 Then show "[ShellAndWait] shResult = " + CStr(shResult)
    Beep
End Sub

Sub TEST_T6_Write_TextFile_MyTags()
    'Vorbereitung
        Dim N$, p1$, p2$, T$, t1&, tDiff&
        DoArrc
    'PathOfSourceFolder
        p1 = "F:\Archiv TR\Archiv Trampolin\Events\1973 Liga\1973-12 DMM BuliE Bruchsal"
    'PathOfFileToWrite
        N = Get_NameFromPath(p1) + ".txt": N = Replace(N, "-", ""): N = Replace(N, " ", "")
        p2 = ArrC(1) + "\Tags\TagsInFolder_" + N
    'StartTime
        t1 = GetTickCount()
    'Write textfile
        T6_Write_TextFile_MyTags p1, p2
    'Duration
        tDiff = GetTickCount() - t1
        'show CStr(CDbl(tDiff / 1000)) & " seconds"
    'Finals
        Beep
        'PepupTags
End Sub

Sub T6_Write_TextFile_MyTags(PathOfSourceFolder$, PathOfFileToWrite$)
    'Called from    T6
    'Action         erzeugt 1 TextDatei (001_ExifToolTags.txt) mit den angegebenen
    '               tags aller JPGs in SourceFolder+Subfolder
    'Vorbereitung
        Dim myCharset$, myOptions$, myTags$, IgnoreSomeJpg$
        Dim p1$, p2$, pE$, qq$, shellCmd$, shResult%
        qq = Chr(34)
        pE = qq + ArrC(9) + qq              'PathOfExiftool"
        p1 = qq + PathOfSourceFolder + qq   'directory
        p2 = qq + PathOfFileToWrite + qq    '.txt
    'myCharset
        myCharset = " -charset filename=latin "
    'myTags
        myTags = " " + qq + "$Directory/$FileName|$RegionAreaX" + qq + " "
    'MyOptions
        'IgnoreSomeJpg = " -if " + qq + "$filename !~ /result/ and not $RegionAreaX eq ''" + qq + " "
        'myOptions = IgnoreSomeJpg + " -ext jpg -L -f -m -r -sep °°° -p "
        myOptions = " -ext jpg -L -f -m  -sep °°° -p "
    'ShellCmd erstellen
        shellCmd = pE + myOptions + myTags + p1 + " -W+! " + p2
            'show ShellCmd
            'Bei Eingabe des ShellCmd in das CmdWindow wird die Meldung
            ' FileName encoding not specified.  Use "-charset FileName=CHARSET"
            'ausgegeben, das Cmd aber abgearbeitet
    'ShellCmd senden
        shResult = ShellAndWait(shellCmd, 0, vbHide, PromptUser)
        If shResult > 0 Then show "[ShellAndWait] shResult = " + CStr(shResult)
End Sub

Sub T6_Notiz()

    'myTags
        myTags = " " + qq + "[Directory:] $Directory [FileName:] $FileName [ImageWidth:] $ImageWidth [ImageHeight:] $ImageHeight "
        myTags = myTags + "[RegionName:] $RegionName [RegionAreaH:] $RegionAreaH [RegionAreaW:] $RegionAreaW "
        myTags = myTags + "[RegionAreaX:] $RegionAreaX [RegionAreaY:] $RegionAreaY [END]" + qq + " "
    'MyOptions
        '-L       (Use Windows Latin1 encoding; für äöüß)
        '-k       (cmdFenster bleibt; für Kontrolle/Fehlermeldungen)
        '-f       (angegebener Tag wird immer mitgelistet - auch wenn leer)
        '-m       (Ignore minor errors and warnings)
        '-progress[:[TITLE]]     (Show file progress count; Anzeige Dateien im cmdFenster)
        '-sep [|] (Seperator "|" innerhalb eines tags - anstelle ",")
        '-s       (tag names statt Descriptions)
        '-s2      (tag names statt Descriptions; ohne zusätzliche Blanks)
        '-r       (causes subdirectories to be processed recursively)
        '-P       (Preserve file modification date/time)
        '-w+!     (-w = Textdatei erzeugen;
        '           + = in existierende Datei dazuschreiben;
        '           ! = Datei überschreiben, falls existent)
End Sub

Sub PepupTags()
    Dim p1$, p2$, s$
    
    p1 = "F:\Archiv TR\Archiv Trampolin\Leute\Besenfelder, Ulrich (rn Betz, TSG Bruchsal)"
    p2 = ArrC(1) + "\Tags\TagsIn_BesenfelderUlrich.txt"
    s = ReadFile(p2)
    s = Replace(s, "[", vbCrLf + "   [")
    s = Replace(s, "[Directory:] ", vbCrLf + "[Directory:]      ")
    s = Replace(s, "[FileName:] ", "[FileName:]    ")
    s = Replace(s, "[ImageWidth:] ", "[ImageWidth:]  ")
    s = Replace(s, "[ImageHeight:] ", "[ImageHeight:] ")
    s = Replace(s, "[RegionName:] ", "[RegionName:]  ")
    show s 'TAB_Simulation(s)
End Sub

Sub TEST_T6_Write_AllTagsOfOneJpg()
    Dim p1$
    p1 = "F:\Archiv TR\Archiv Trampolin\Events\1966-05-08 DJM Nürnberg\19660508 DJM Nürnberg a03 (Kurt Flöß) p1747-04 ub.jpg"
    T6_Write_AllTagsOfOneJpg p1
    Beep
End Sub
   
Sub T6_Write_AllTagsOfOneJpg(PathOfJpg$)
    'Called from    T6
        Dim myOptions$, pExif$, pJpg$, pOUT$, qq$, shellCmd$, shResult%
        qq = Chr(34): DoArrc
    'Vorbereitung
        pExif = qq + ArrC(9) + qq + " "
        pJpg = " " + qq + PathOfJpg + qq + " "
        'pOUT = Replace(pJPG, ".jpg", ".jpg -G1 .txt") '-G
        pOUT = " " + qq + ArrC(1) + "\Tags\test.txt" + qq + " "
    'Options
        myOptions = "-G1 -a -s"     '[XMP-mwg-rs] RegionName ...
        '-G     include group names
        '-g1    sort tags by group (for family 1)
        '-GROUP:TAG=        Deletes TAG only in specified group
        '-GROUP:all=        Deletes all information in specified group
        '-TAG=VALUE         Sets value of TAG (in all GROUPS)
        '-[GROUP:]TAG=VALUE Sets value of TAG (only in GROUP if specified)
        '-a                 -a, --a (-duplicates, --duplicates)
        '                   Allow (-a) or suppress (--a) duplicate tag names to be extracted.
        '                   By default, duplicate tags are suppressed when reading unless the -ee or -X options are used or the Duplicates option is enabled in the configuration file. When writing, this option allows multiple Warning messages to be shown. Duplicate tags are always extracted when copying.
        '-s                 Short output format
        '-s1 or -s          print tag names instead of descriptions
        '-s2 or -s -s or -S no extra spaces to column-align values
        '-s3 or -s -s -s    print values only (no tag names)
    'ShellCmd erstellen
        shellCmd = pExif + myOptions + pJpg + " -W+! " + pOUT
        show shellCmd
    'ShellCmd senden
        shResult = ShellAndWait(shellCmd, 0, vbHide, PromptUser)
        If shResult > 0 Then show "[ShellAndWait] shResult = " + CStr(shResult)
End Sub












