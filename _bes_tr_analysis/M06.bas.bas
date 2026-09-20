Attribute VB_Name = "M06"
Option Explicit 'M06
''WaitForFileToExist
'    Const TimeOutSeconds As Integer = 20
''myFSO
'    Private localFSO As Object

Sub zzz_M06()
    Application.EnableEvents = True
    
    showProcs "load"
    
    'RenameModule "Modul1", "M05"
    'T1_ShowLogBuch
    'Sheets("T4").shapes("RE1").Visible = True
    Beep
End Sub

Sub AddToReport02(Report$)
    If ArrC(60) = "0" Then     '0 = don't write ResultJPGs
        Report = Report + "    No JPG:       [" + ArrC(43) + "] was not produced in EventFolder [" + ArrC(42) + "]" + vbCrLf
    Else                       '1 = write ResultJPGs
        Report = Report + "    Original-JPG: [" + ArrC(43) + "] was produced in EventFolder [" + ArrC(42) + "]" + vbCrLf
    End If
End Sub

Sub AddToReport01(Report$, AllNamesOneDg$)
    If AllNamesOneDg = "" Then
        Report = Report + "    Competitors:  00" + vbCrLf
    Else
        Report = Report + "    Competitors:  " + Format(anzAinB(vbCrLf, AllNamesOneDg) + 1, "00") + vbCrLf
    End If
End Sub

Sub ShowInfo01NoFolder(Title$)
    show "Ein Ordner zum Design '" + Title + "' existiert nicht in 'Events'."
End Sub

Sub Update_NamesOfPersonFolders(PathsOfAllLeuteFolder$, AllNamesOfOneDesign$, AllNameVerein$, AllNameNation$, Report$)
    'Called from            Update_Jpg_Links_Comp
    'PathsOfAllLeuteFolder  = "F:\Archiv...\Leute\Aaron, Syd (GB)"  + ähnl. Zeilen
    'AllNamesOfOneDesign    = "Agathe Jarosch"                      + ähnl. Zeilen
    'AllNameVerein          = "Agathe Jarosch|TSG Steglitz"         + ähnl. Zeilen
    'AllNameNation          = "Karl Marx|SU"                        + ähnl. Zeilen
    'Action         For each Name in one Design: Change his FolderName inside Folder "Leute"
    '               "Marx, Karl" --> "Marx, Karl (SU)"
    '               "Jarosch, Agathe (Berlin)" --> "Jarosch, Agathe (TSG Berlin-Steglitz)"
    '               "Harz, Ute (vh Oder, Wiesloch)" --> "Harz, Ute (vh Oder, TSG Wiesloch)"
    
    'Vorbereitung
        Dim N1$, N2$, N3$, N4$, Nation$, Nn$, p$, v$, Verein$, Vn$, zU$, i%, c&, Arr1() As String
        v = vbCrLf
        p = ArrC(4) + "\"   'Path of Folder "Leute"
    'Action
        Arr1 = Split(AllNamesOfOneDesign, v)
        For i = 0 To UBound(Arr1)
            N1 = "": N2 = "": N3 = "": N4 = "": Verein = "": Nation = "": zU = "": Vn = "": Nn = ""
            'N1 (Name im Design)
                N1 = Arr1(i)                    'N1 = "Agathe Jarosch"
            'N2 (Name für die Suche)            'N2 = "Jarosch, Agathe"
                c = InStr(1, N1, " "): N2 = ""        'Nn, Vn (Nachname, Vorname)
                If c > 0 Then Vn = Left(N1, c - 1):    Nn = Mid(N1, c + 1): N2 = Nn + ", " + Vn
            'N3 (aktueller OrdnerName)          'N3 = "Jarosch, Agathe (Berlin)"
                c = InStr(1, PathsOfAllLeuteFolder, N2)
                If c > 0 Then N3 = CutVonBis(PathsOfAllLeuteFolder + v, c, InStr(c, PathsOfAllLeuteFolder + v, v) - 1)
            'N3 = ""    'Ggf. enthält N1 nicht den GeburtsNamen (also kein LeuteOrdner unter diesem Namen);
                        'dann besteht kein Bedarf, den zugehörigen PersonenOrdner zu finden/ändern
                If c = 0 Then GoTo NextCompetitorOfThisDesign
            'Existiert ein Zusatz in N3?
                c = InStr(1, N3, "(")
                If c = 0 Then
                    'N3 enthält nur den Namen+Vornamen
                    'Existiert die Angabe eines zugehörigen Vereins?
                        c = InStr(1, AllNameVerein, N1 + "|")
                        If c > 0 Then
                            'N4 = N3 + Zusatz Verein 'Zum Namen N3 existiert die Angabe eines zugehörigen Vereins
                                Verein = CutVonBis(AllNameVerein + v, c + Len(N1) + 1, InStr(c, AllNameVerein + v, v) - 1)
                                If Verein = "" Then GoTo NextCompetitorOfThisDesign
                                N4 = N3 + " (" + Verein + ")"
                            'Rename
                                If Not FolderExists(p + N3) Then Stop
                                If Verein <> "" Then RenameFolder p + N3, p + N4
                        Else
                    'Existiert die Angabe einer zugehörigen Nation?
                            c = InStr(1, AllNameNation, N1 + "|")
                            If c > 0 Then
                                'N4 = N3 + Zusatz Nation 'Zum Namen N3 existiert die Angabe einer zugehörigen Nation
                                    Nation = CutVonBis(AllNameNation + v, c + Len(N1) + 1, InStr(c, AllNameNation + v, v) - 1)
                                    If Nation = "" Then GoTo NextCompetitorOfThisDesign
                                    N4 = N3 + " (" + Nation + ")"
                                'Rename
                                    If Not FolderExists(p + N3) Then Stop
                                    If Nation <> "" Then RenameFolder p + N3, p + N4
                            End If
                        End If
                    'New FolderName --> PathsOfAllLeuteFolder
                        If N4 <> "" Then
                            PathsOfAllLeuteFolder = Replace(PathsOfAllLeuteFolder, p + N3, p + N4)
                            If N4 Like "*()*" Then Stop
                            Report = Report + "    Umbenennung:  Ordner [" + N3 + "] --> Ordner [" + N4 + "]" + v
                        End If
                Else
                    zU = Mid(N3, c, Len(N3) - c - 1)    'Zu = "Berlin"
                End If
NextCompetitorOfThisDesign:
        Next
        Report = Report + v
End Sub

Sub UpdateOnce_OrtToVerein_InsideNameOfPersonFolders()
    'Called from    xxx
    
    'Vorbereitung
        Dim p1$, p2$, PathsOfAllLeuteFolder$, s$, v$, i%, j%, N%
        Dim Arr1() As String, Arr2() As String, Arr3() As String, Arr4() As String
        v = vbCrLf: DoArr
    'Pfade aller Ordner in "Leute"
        PathsOfAllLeuteFolder = Get_Paths_ofAllSubfolders_OneLevel(ArrC(4))
        'show PathsOfAllLeuteFolder
    'SOLL-Änderungen
        s = "Bruchsal,TSG Bruchsal|Villingen,TV Villingen|Duisburg,Tura Duisburg" _
            + "|Nürnberg,ASV Nürnberg|Freiburg,FT Freiburg|Leimen,KuSG Leimen" _
            + "|Bad Kreuznach,MTV Bad Kreuznach|Elze,MTV Elze|Hamburg,PolSV Hamburg" _
            + "|Witten-Annen,SU Witten-Annen|Andernach,TB Andernach|Düsseldorf,TV Düsseldorf" _
            + "|Osnabrück,TB Osnabrück|TG Süchteln,TG Süchteln|Salzgitter,TGJ Salzgitter" _
            + "|Mutterstadt,TSG Mutterstadt|TSG Steglitz,TSG Berlin-Steglitz|Wiesloch,TSG Wiesloch" _
            + "|TSV Spandau,TSV Berlin-Spandau|Erlangen,TV Erlangen|Gernsbach,TV Gernsbach" _
            + "|Obing,TV Obing|Unterbach,TV Unterbach|Bremerhaven,TVL Bremerhaven"
        s = "Ichenheim,TV Ichenheim|Hemsbach,SuT Hemsbach|Linkenheim,TV Linkenheim" _
            + "|Rinteln,VT Rinteln|Frankfurt-Nied,SG Frankfurt-Nied|Spandau,TSV Berlin-Spandau" ' _
            + "|xxx,xxx|xxx,xxx|xxx,xxx"
        Arr2 = Split(s, "|"): N = UBound(Arr2)
        ReDim Arr3(0 To N): ReDim Arr4(0 To N)
        For i = 0 To N
            Arr3(i) = Left(Arr2(i), InStr(1, Arr2(i), ",") - 1) 'Bruchsal
            Arr4(i) = Mid(Arr2(i), InStr(1, Arr2(i), ",") + 1)  'TSG Bruchsal
        Next
    'Action
        Arr1 = Split(PathsOfAllLeuteFolder, v)
        For i = 0 To UBound(Arr1)
            p1 = Arr1(i)         'one path 'F:\Archiv...\Leute\Mai, Karl (Bruchsal)
            
            For j = 0 To N
                If p1 Like "*" + Arr3(j) + ")" And Not p1 Like "*" + Arr4(j) + ")" Then
                    p2 = Replace(p1, Arr3(j), Arr4(j))
                    RenameFolder p1, p2
                    'Stop
                End If
            Next
        Next
End Sub

Sub Create_LinkOfDgLeute(ArrPathsDg, Report$)
    'Called from    Update_Jpg_Links_Comp
    'ArrPathsDg     1D-Array; Pfade zu Ordnern, in die das JPG kopiert werden soll
    'Action         Erstellt für jeden DgName einen LinkInsideDg zu seinem LeuteOrdner
    
    'Vorbereitung
        Dim NameOfPersonFold$, PathOfPersonFold$, pEventFold$, i%
        pEventFold = ArrC(57)                       'Path of EventFolder
        
        
        'OpenFolder pEventFold: Stop
        
        
    'Copy
        For i = 1 To UBound(ArrPathsDg)
            PathOfPersonFold = ArrPathsDg(i)               'Path of one CompetitorFolder
            NameOfPersonFold = Get_NameFromPath(PathOfPersonFold)   'Name of one CompetitorFolder
            'Create_OneLinkFile pEventFold, PathOfPersonFold, NameOfPersonFold
            Stop
            Create_OneVbsLink_Person pEventFold + "\" + NameOfPersonFold
        Next
End Sub

Sub CopyOneJpgToFolders(ArrPathsDg, Report$)
    'Called from    Update_Jpg_Links_Comp
    'ArrPathsDg     1D-Array; ArrPathsDg() As String
    '               Pfade zu Ordnern, in die das JPG kopiert werden soll
    'Action         Kopiert das ResultJpg eines Designs
    '               von seinem EventOrdner zu allen LeuteOrdnern der CompetitorsInsideDesign
    
    'Vorab
        If ArrPathsDg(1) = "" Then
            Report = Report + "    JpgCopy to:   -" + vbCrLf
        Else
            Report = Report + "    JpgCopy to:   "
        End If
        'If ArrC(60) = "0" Then Exit Sub     'T4Dg Write ResultJpg? 0 = No, 1 = Yes
    'Vorbereitung
        Dim fSizeOld$, fSizeOriginal$, NameOfPath$, NameJpg$, pFOLD$, p1$, p2$, v$, i%
        v = vbCrLf
        NameJpg = ArrC(43)                  'Name of result.jpg
        p1 = ArrC(57) + "\" + NameJpg       'Path of original result.jpg
        'show Join(ArrPathsDg, v)
    'Copy
        If FileExists(p1) Then
            fSizeOriginal = CStr(FileSize(p1))
            For i = 1 To UBound(ArrPathsDg)
                pFOLD = ArrPathsDg(i)           'Path of one LeuteFolder
                If Not FolderExists(pFOLD) Then Stop
                NameOfPath = Get_NameFromPath(pFOLD)
                
                p2 = pFOLD + "\" + NameJpg   'Path of CopyFile to write
                If FileExists(p2) Then fSizeOld = CStr(FileSize(p2)) Else fSizeOld = "0"
                'Copy
                    If ArrC(60) = "1" Then CopyFile p1, p2
                If i = 1 Then
                    Report = Report + Format(i, "00") _
                        + " [" + Left(NameOfPath + String(40, " "), 40) + "] " _
                        + "SizeOld " + fSizeOld + ", SizeOriginal " + fSizeOriginal + v
                Else
                    Report = Report + "                  " + Format(i, "00") + " [" + Left(NameOfPath + String(40, " "), 40) + "] " _
                        + "SizeOld " + fSizeOld + ", SizeOriginal " + fSizeOriginal + v
                End If
            Next
        Else
            Report = Report + "                ERROR: File [" + p1 + "] not existing" + v
        End If
End Sub

Sub RedRectToNextDgListTitle(Title$)
    'Called from    Update_Jpg_Links_Comp
    
    Dim s%, z%
    If Right(Title, 1) <> " " Then Title = Title + " "
    '(z,s)=(ZeilenNr,SpaltenNr) des Titles in der T4-Spalte B (List of DesignTitles)
        z = Get_RowNr_HoldingMyTextWhole("T4", Title): s = Get_ColumnNr_HoldingMyTextWhole("T4", Title)
        If z = 0 Or s = 0 Then
            show "'List of DesignTitles' enthält nicht den Eintrag '" + Title + "'"
            Stop
        End If
    With Sheets("T4").Shapes("T4RedRect"): .Visible = True
        .Left = Sheets("T4").Cells(z, s).Left - 3:    .Top = Sheets("T4").Cells(z, s).Top - 3
        .Width = Sheets("T4").Cells(z, s).Width + 6:  .Height = Sheets("T4").Cells(z, s).Height + 6
    End With
    EE 0: Application.ScreenUpdating = True
    Application.GoTo Reference:=Worksheets("T4").Cells(z - 7, s - 1), Scroll:=True
    DoEvents
    Application.ScreenUpdating = False: EE 1
End Sub

Function CutVonBis(MyString, VON&, BIS&) As String
    CutVonBis = Mid(MyString, VON, BIS - VON + 1)
    'Stop
End Function

Function Get_vhRenamings(PathsLeute$) As String
    Dim Nn$, Nn2$, s$, vh$, Vn$, c%, C2%, i%, Arr1() As String
    Arr1 = Split(PathsLeute, vbCrLf)
    For i = 0 To UBound(Arr1)
        If InStr(1, Arr1(i), "vh ") > 0 Then
            s = Arr1(i) '...\Leute\Czech, Ute (vh Latton, vh Luxon, vh Pitkamin, Salzgitter)
            s = Mid(s, InStr(1, s, "\Leute\") + 7) 'Czech, Ute (vh ...
            'soll liefern: "Ute Latton|Ute Czech" v "Ute Luxon|Ute Czech" v "Ute Pitkamin|Ute Czech"
            c = InStr(1, s, ", ")
            Nn = Left(s, c - 1)         'Czech
            s = Mid(s, c + 2)           'Ute (vh Latton, ...
            c = InStr(1, s, " (")
            If c > 0 Then
                Vn = Left(s, c - 1)     'Ute
                s = Mid(s, c + 2)       'vh Latton, vh Luxon, ...   'vh Maier)
                s = Replace(s, ")", ",") 'vh Latton, vh Luxon, vh Pitkamin, Salzgitter,  'vh Maier,
                Do While s Like "*vh *"
                    c = InStr(1, s, "vh ")
                    C2 = InStr(2, s, ",")
                    Nn2 = Mid(s, c + 3, C2 - c - 3)
                    s = Mid(s, C2)
                    vh = vh + Vn + " " + Nn2 + "|" + Vn + " " + Nn + vbCrLf
                Loop
            End If
        End If
    Next
    vh = Delete_EmptyEndRowsInString(vh)
    Get_vhRenamings = vh
    'show vh
End Function

Sub Check_Format(PathsOfAllJPGs$)
    'Vorbereitung
        Dim D$, F$, Id$, NAME$, p$, Q$, s$, v$
        Dim AnzD%, AnzF%, AnzQ%, i%, N%, Arr1() As String
        v = vbCrLf: With Sheets("T1")
        .Cells(16, 19) = "Fehlende Quellenkürzel: " ':  .Cells(16, 31) = 0
        .Cells(17, 19) = "Bad format 'ID...jpg': " ':   .Cells(17, 31) = 0
        .Cells(18, 19) = "Bad Date: " ':                .Cells(18, 31) = 0
    'Action
        Arr1 = Split(PathsOfAllJPGs, v)
            N = UBound(Arr1) + 1                'n = Anzahl JPGs
        For i = 0 To UBound(Arr1)
            p = Arr1(i)                         'p = 1 JPG-Pfad
            NAME = Get_NameFromPath(p)
            If Not Right(p, 7) Like " [a-z][a-z].jpg" Then
                Q = Q + p + v                   'Q = JPG-Pfade ohne Quellenkürzel
                AnzQ = AnzQ + 1
                .Cells(16, 31) = "'" + CStr(AnzQ)
            ElseIf Not Right(p, 16) Like " [p|v]####-## [a-z][a-z].jpg" Then
                F = F + p + v                   'F = JPG-Pfade in unkorrektem ID-Format
                AnzF = AnzF + 1
                .Cells(17, 31) = "'" + CStr(AnzF)
            ElseIf Not Left(NAME, 9) Like "######## " And Not Left(NAME, 16) Like "########_###### " Then
                D = D + p + v                   'D = JPG-Pfade mit unkorrektem Datum
                AnzD = AnzD + 1
                .Cells(18, 31) = "'" + CStr(AnzD)
            End If
            If (N - i) Mod 200 = 0 Then
                .Cells(18, 19) = "Bad Date:  " + CStr((N - i) \ 200) 'CountDown
                DoEvents
            End If
        Next
        .Cells(16, 31) = "'" + CStr(AnzQ)
        .Cells(17, 31) = "'" + CStr(AnzF)
        .Cells(18, 19) = "Bad Date:": .Cells(18, 31) = "'" + CStr(AnzD)
        If Q <> "" Then s = CStr(AnzQ) + " JPG-Pfade ohne Quellenkürzel:" + v + Q + v
        If F <> "" Then s = s + CStr(AnzF) + " JPG-Pfade in unkorrektem ID-Format:" + v + F + v
        If D <> "" Then s = s + CStr(AnzD) + " JPG-Pfade mit unkorrektem Datum:" + v + D + v
        If s <> "" Then show "T6_Konsistenzprüfung" + v + v + s
        End With
End Sub

Sub Check_DoubleIDsInEvents(PathsOfAllJPGs$)
    'Vorbereitung
        Dim E$, G$, Id$, p$, v$, AnzG%, i%, N%, Arr1() As String
        v = vbCrLf: With Sheets("T1")
        .Cells(19, 19) = "ID-Doubles in Events: " ':              .Cells(19, 31) = 0

    'Action
        Arr1 = Split(PathsOfAllJPGs, v)
            N = UBound(Arr1) + 1                'n = Anzahl JPGs in Events
        For i = 0 To UBound(Arr1)
            p = Arr1(i)                         'p = 1 JPG-Pfad
            If p Like "*########* [p|v]####-## [a-z][a-z].jpg" Then
                'JPG-Pfad ist in korrektem Format
                Id = Mid(p, Len(p) - 14, 8) 'p1234-00 oder v1234-56
                If anzAinB(Id, PathsOfAllJPGs) > 1 Then
                    If InStr(1, G, Id) = 0 Then G = G + Id + v  'G = IDs, die mehrfach vorkommen
                    If InStr(1, E, p) = 0 Then E = E + p + v    'E = JPG-Pfade mit ID-Double
                    AnzG = anzAinB(v, G)
                    .Cells(19, 31) = "'" + CStr(AnzG)
                End If
            End If
            If (N - i) Mod 200 = 0 Then
                .Cells(19, 19) = "ID-Doubles in Events:  " + CStr((N - i) \ 200) 'CountDown
            End If
        Next
        .Cells(19, 19) = "ID-Doubles in Events: ": .Cells(19, 31) = "'" + CStr(AnzG)
        If G <> "" Then show "T6_Konsistenzprüfung" + v + v + CStr(AnzG) + " ID-Doubles" + v + G + v _
            + " JPG-Pfade mit ID-Double:" + v + E + v
        End With
End Sub

Function Get_Paths_OfAllJPGs_InFolderAndSubfolders(PathOfSourceFolder$) As String
    'Vorbereitung
        Dim F$, Fi$, Fo$, ID1$, N$, p$, p2$, Ren$, v$, Z2L$, Arr1() As String
        Dim c%, CountRen%, i%, nFi%, gFi%, nFo%, E As Boolean, KP As Boolean
        With Sheets("T1"): v = vbCrLf
        If .Cells(13, 19) = "- T6_Konsistenzprüfung -" Then KP = True
        N = Get_NameFromPath(PathOfSourceFolder)
        If N = "Events" Then E = True
        If KP Then Z2L = .Cells(14, 19): Z2L = Left(Z2L, InStr(1, Z2L, ";") + 1) 'Z2L = Zeile 2 in T1-Info, linker Teil
    'SubFolders in 'Events'
        F = PathOfSourceFolder + v + Get_Paths_ofAllSubfoldersAllLevelsAsStringUseGlobalVar(PathOfSourceFolder)
        Arr1 = Split(F, vbCrLf)
            nFo = UBound(Arr1) + 1  'Anzahl Ordner
    'Pfade aller JPGs in SourceFolder
        For i = 0 To UBound(Arr1)
            'Schleife über alle SubFolder von SourceFolder
            Fo = Arr1(i)    'Path of one Folder
            c = nFo - i     'Countdown Folder
            Fi = Get_AllFilePaths_WithMyStringInFileName_OfOneFolder(Fo, ".jpg")  'mit .JPG
            nFi = anzAinB(".jpg", Fi)   'nFi = Anzahl JPGs im Ordner Fo
            gFi = gFi + nFi             'gFi = Anzahl JPGs gesamt
            p = p + Fi + vbCrLf         'P   = Alle JPG-Pfade
            If KP Then
                If E Then
                    If c Mod 10 = 0 Then .Cells(14, 19) = "JPGs: " + CStr(gFi) + " in 'Events'; 0 in 'Leute'"
                Else
                    If c Mod 10 = 0 Then .Cells(14, 19) = Z2L + CStr(gFi) + " in 'Leute'"
                End If
            End If
        Next
        p = Delete_EmptyEndRowsInString(p)
    'jpg only
        Arr1 = Split(p, vbCrLf): c = 0
        If KP Then
            .Cells(15, 19) = "Rename '.JPG' --> '.jpg'"
            Ren = CStr(.Cells(15, 31))
            If Ren = "" Then
                .Cells(15, 31) = 0: CountRen = 0
            Else: CountRen = CInt(Ren)
            End If
        End If
        For i = 0 To UBound(Arr1)
            p = Arr1(i)
            If KP Then
                If (UBound(Arr1) - i) Mod 200 = 0 Then .Cells(15, 19) = "Rename '.JPG' --> '.jpg' " + CStr((UBound(Arr1) - i) \ 200)
            End If
            If p Like "*.jpg" Then
                p2 = p2 + p + vbCrLf
            Else
    'JPG --> jpg
                If LCase(p) Like "*jpg" Then
                    '"... .jpg.lnk" bleiben weg, also Ext = "JPG", "jPg", ...
                    RenameFile p, Left(p, Len(p) - 3) + "jpg"
                    p2 = p2 + Left(p, Len(p) - 3) + "jpg"
                    CountRen = CountRen + 1
                End If
            End If
        Next
        If KP Then .Cells(15, 19) = "Rename '.JPG' --> '.jpg'": .Cells(15, 31) = CountRen
        p2 = Delete_EmptyRowsInString(p2)
    'Return the result
        Get_Paths_OfAllJPGs_InFolderAndSubfolders = p2
    End With
End Function

Sub getJpgFiles()
    'sehr schnell, aber CmdWindow flashed kurz auf
    Dim i&, fName$, fFull$, fPath$, outLines As Variant, output As Object
    Dim inLines() As String, sLine$, lines&, ShCmd$, SourceFolder$, qq$
    'Set output = ShellOutput("Dir D:\BYoung\*.pdf /s /b /a:-d")
    Call DoArr: qq = Chr(34) 'quote, "-Zeichen
    'SourceFolder = ArrC(3)
    SourceFolder = "F:\Archiv TR\Lieferungen von\20240922 von ik Ivonne Kraft\Ordner 01 Bilderbuch1"
    
    'shCmd = "chcp 1250|Dir " + qq + SourceFolder + qq + "\*.jpg /s /b /a:-d"
    ShCmd = "chcp 1250|cmd.exe /c Dir " + qq + SourceFolder + qq + "\*b.jpg /s /b /a:-d"
        'chcp 1250  aktiviert CodePage für Umlaute
        '/s         auch Unterverzeichnisse durchsuchen
        '/b         Verzeichnissen/Dateien ohne zusätzliche Informationen
        '/a:-d      keine Verzeichnisse mit auflisten
    Set output = ShellOutput(ShCmd)
    
    Do While Not output.AtEndOfStream
        sLine = output.ReadLine
        'If right(sLine, 4) = ".jpg" Then
            'If sLine Like "*ö*" Then Stop
            i = i + 1
            ReDim Preserve inLines(1 To i)
            inLines(i) = sLine
        'End If
        'If i Mod 100 = 0 Then DoEvents
    Loop
    
    showArray inLines

End Sub

Function Get_PathOfMyFiles(PathOfSourceFolder$, SubDirectories01%, FileNameLike$) As String
    'Status     sehr schnell, aber CmdWindow flashed kurz auf
    
    'Vorbereitung
        Dim A$, L$, s$, output As Object
        Dim sLine$, ShCmd$, p$, p2$, qq$
        qq = Chr(34) 'quote, "-Zeichen
        p = PathOfSourceFolder
            If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
        L = "*" 'FileNameLike
        A = "chcp 1250|cmd.exe /c Dir "
    'ShellCommand
        If SubDirectories01 = 1 Then
            ShCmd = A + qq + p + "\" + L + qq + " /s /b /a:-d"
        Else
            'shCmd = a + qq + p + qq + "\" + L + " /b /a:-d": p2 = p + "\"
            ShCmd = A + qq + p + "\" + L + qq + " /b /a:-d": p2 = p + "\"
            'chcp 1250  aktiviert CodePage für Umlaute
            '/s         auch Unterverzeichnisse durchsuchen; returns full path
            '/b         Verzeichnissen/Dateien ohne zusätzliche Informationen
            '/a:-d      keine Verzeichnisse mit auflisten
        End If
        'show shCmd
    'Suche
        Set output = ShellOutput(ShCmd)
        Do While Not output.AtEndOfStream
            sLine = output.ReadLine
            If sLine Like FileNameLike Then s = s + p2 + sLine + vbCrLf
        Loop
    'Finals
        s = Delete_EndReturnsInString(s)
        Get_PathOfMyFiles = s
End Function

Public Function ShellOutput(sCmd As String) As Object
    'Create a Shell, executes a command, and returns the output stream
    Dim oShell As Object, oExec As Object, oOutput As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command
    Set oExec = oShell.Exec("cmd.exe /c " & sCmd)
    Set oOutput = oExec.StdOut
    Set ShellOutput = oOutput
End Function

Function Kill_LastChars(StrA$, StringToKill$) As String
    Dim L%
    L = Len(StringToKill)
    If Right(StrA, L) = StringToKill Then
        Kill_LastChars = Left(StrA, Len(StrA) - L)
    Else
        Kill_LastChars = StrA
    End If
End Function

Sub Get_ColorGreen1()
    Dim s As String, r As Range
    Set r = Range("B2")
    s = CStr(getColor(r))
    show s '14348258
End Sub

Public Function WaitForFileToExist(ByVal PathOfFile As String) As Boolean
    'Check      jede Sekunde
    Dim timeElapsed As Single
    Dim startTime As Single
    startTime = Timer
    Do
        If myFSO.FileExists(PathOfFile) Then
            WaitForFileToExist = True
            Exit Do
        End If
        'File existiert noch nicht
        DoEvents
        Application.Wait Now + TimeValue("0:00:01")
        DoEvents
        timeElapsed = Timer - startTime
    Loop Until timeElapsed > TimeOutSeconds 'TimeOutSeconds vordefiniert = 2
    DoEvents
End Function

Public Function myFSO() As Object
    If localFSO Is Nothing Then Set localFSO = CreateObject("Scripting.FileSystemObject")
    Set myFSO = localFSO
End Function

Sub TEST_GetOutputOfPowerShellCmd()
    show GetOutputOfPowerShellCmd("Get-ComputerInfo -Property 'OsName'")
End Sub

Public Function GetOutputOfPowerShellCmd(ByVal sPSCmd As String) As String
    'Setup the powershell command properly
    sPSCmd = "powershell -command " & sPSCmd & "|clip"
    'Execute the command which is being pushed to  the clipboard
    CreateObject("WScript.Shell").Run sPSCmd, 0, True
    'Get an instance of the clipboard to capture the save value
    With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .GetFromClipboard
        GetOutputOfPowerShellCmd = .GetText(1)
    End With
End Function

Sub DrawW1Rect(NameOfSheet$, NameOfRect$, Left!, Top!, Width!, Height!)
    Dim sH As Object
    Delete_W1Shape NameOfRect
    Set sH = Sheets(NameOfSheet).Shapes.AddShape(msoShapeRectangle, Left, Top, Width, Height)
    sH.NAME = NameOfRect
    sH.Fill.Transparency = 1
    sH.Line.Weight = 1.5
    sH.Line.ForeColor.RGB = RGB(255, 255, 0)
    sH.ZOrder msoBringToFront
End Sub

Sub DrawW1RectBlack(NameOfSheet$, NameOfRect$, Left!, Top!, Width!, Height!)
    Dim sH As Object
    Delete_W1Shape NameOfRect
    Set sH = Sheets(NameOfSheet).Shapes.AddShape(msoShapeRectangle, Left, Top, Width, Height)
    sH.NAME = NameOfRect
    sH.Fill.Transparency = 1
    sH.Line.Weight = 0.5
    sH.Line.ForeColor.RGB = RGB(0, 0, 0)
    sH.ZOrder msoBringToFront
End Sub

Sub DrawLine(NameOfSheet$, NameOfLine$, x1!, y1!, X2!, y2!)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(NameOfSheet)
    Delete_W1Shape NameOfLine
    ws.Shapes.AddLine(x1, y1, X2, y2).NAME = NameOfLine
    ws.Shapes(NameOfLine).Line.ForeColor.RGB = RGB(255, 0, 0)
    ws.Shapes(NameOfLine).Line.Weight = 2
End Sub

Sub ClickOnGreen1(z%, s%)
    'Status     User klickte auf eine hellgrüne (Green1) Zelle
    'Action     Ermittlung der Zelle li ob des hellgrünen Bereiches
    
    Dim i%, s1%, s2%, z1%, z2%
    Application.EnableEvents = False
    With Sheets("T1")
    'z1 - oberen Rand suchen
        For i = 0 To 99
            If .Cells(z - i, s).Borders.LineStyle = xlNone And getColor(.Cells(z - i, s)) <> 14348258 Then z1 = z - i + 1: Exit For
        Next
    's1 - linken Rand suchen
        For i = 0 To 99
            If .Cells(z, s - i).Borders.LineStyle = xlNone And getColor(.Cells(z, s - i)) <> 14348258 Then s1 = s - i + 1: Exit For
        Next
    'select
        .Cells(z1, s1).Select
    'Scroll if Klick on SourceImage-Green1
        If .Cells(z1, s1 + 1) = "SourceImage" Then
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = s1 - 1 'Spalte, die ganz links zu sehen sein soll
            ActiveWindow.Zoom = 100
        End If
    End With
    Application.EnableEvents = True
End Sub

Sub xx()
    show Get_AllDgFolderNames_FromDgs
End Sub

Function Get_AllDgFolderNames_FromDgs() As String
    'Called from    xxx
    
    'Vorbereitung
        Dim A$, B$, s$, v$, i&, j&, w As Worksheet, Arr(), Arr1() As String
        v = vbCrLf: Set w = Sheets("T4")
    'Gesamten Dg-Bereich in Array nehmen
        Arr = w.Range(w.Cells(8, 19), w.Cells(LastRow(Sheets("T4")), lastCol(Sheets("T4"))))
    For i = 1 To UBound(Arr, 1) - 2
        For j = 1 To UBound(Arr, 2)
            A = Replace(Arr(i, j), "/", "_")
            If A Like "19##*" Then
                '(i,j)="19.."
                If Arr(i + 1, j) Like "19##*" Then
                    '(i,j)="19.." (i+1,j)="19.."            --> a = FolderName from above DgTitle
                    If InStr(1, B, A + "|") = 0 Then B = B + A + "|"
                Else
                    '(i,j)="19.." und (i+1,j)=""
                    If Not Arr(i - 1, j) Like "19##*" Then
                        '(i,j)="19.." (i+1,j)="" (i-1,j)="" --> a = FolderName from DgTitle
                        '(i,j)=DgTitle
                        If InStr(1, B, A + "|") = 0 Then B = B + A + "|"
                    End If
                End If
            End If
        Next
    Next
    Arr1 = Split(B, "|")
    QuickSort Arr1
    B = Join(Arr1, "|")
    Get_AllDgFolderNames_FromDgs = B
    'show B
End Function

Sub RGB_TEST()
    Dim D$, L$, s$, v$, i%, j%, Sum%, c&, r&, G&, B&, Arr1() As String
    v = vbCrLf
    
    For i = 0 To 999
        D = Format(i, "000")
        r = 255 - CInt(Mid(D, 1, 1))
        G = 255 - CInt(Mid(D, 2, 1))
        B = 255 - CInt(Mid(D, 3, 1))
        Sum = r + G + B
        c = r + G * 256 + B * 256 * 256
        s = s + "sum" + CStr(Sum) + " RGB(" + CStr(r) + "," _
              + CStr(G) + "," + CStr(B) + ") = " + CStr(c) + v
    Next
    s = Delete_EmptyEndRowsInString(s)
    Arr1 = Split(s, v)
    QuickSort Arr1
    For i = UBound(Arr1) To 700 Step -1
        L = Arr1(i)
        r = CInt(Mid(L, 12, 3))
        G = CInt(Mid(L, 16, 3))
        B = CInt(Mid(L, 20, 3))
        For j = 1 To 50
            [do2].Interior.Color = RGB(r, G, B)
        Next
    Next
    's = Join(arr1, v)
    Beep
    'show s
End Sub












