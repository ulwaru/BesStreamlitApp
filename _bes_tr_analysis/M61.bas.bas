Attribute VB_Name = "M61"
Option Explicit 'M61

Sub zzz_M61()

    showProcs "like"
    
    'RenameModule "Modul1", "M62"
    'T1_ShowLogBuch
    EE 1: Beep
    'If MyCountDown Then
    'MyCountDown_MainStepsAllowed 20: MyCountDown_SubStepsMax AnzFold
    'MyCountDown_OneMoreSubStep
    'MyCountDown_OneMoreMainStep
End Sub

Sub Update_FreePids_UpTo_p5000()
    'Called from    xxx
    'Pid            Photo-ID; p1234-00; "...\19631019 DM Bonn d05 p1234-00 aw.jpg"
    'Action         prüft vorhandene Pids; erstellt 500er Liste
    
    'Vorbereitung
        Dim p$, p1$, p2$, p3$, PidNr$, s$, s1$, T$, u$, v$, Nr$, i&, j&
        Dim B() As String, c() As String
        v = vbCrLf: DoArrc
    'Pfade Ordner
        p1 = ArrC(3): p2 = ArrC(11): p3 = ArrC(4) 'Events ClubsNations Leute
        s = Get_Paths_OfAllSubfolders_AllLevels_Collection(p1)
        s = s + v + Get_Paths_OfAllSubfolders_AllLevels_Collection(p2)
        s = s + v + Get_Paths_OfAllSubfolders_AllLevels_Collection(p3)
        'show s
    's1 'string mit 5000 Nrn '0001 0002 ...
        For i = 1 To 5000
            s1 = s1 + "p" + Format(i, "0000") + "-00 "
        Next
        'show s1' 0001 0002 ...
    'B() - Array OrdnerPfade
        B = Split(s, v)
        For i = 0 To UBound(B)
            [AM2] = UBound(B) - i 'Anzeige
            s = Get_AllFileNames_Like_OfOneFolder(B(i), "* p####-## *.jpg")
            c = Split(s, v)
            For j = 0 To UBound(c)
                PidNr = Right(c(j), 15)
                If Left(PidNr, 6) Like "p####-" Then
                    Nr = Mid(PidNr, 2, 4)          'Nr wurde verwendet
                    s1 = Replace(s1, "p" + Nr + "-00 ", "") 'Nr aus s1 löschen
                Else
                    show PidNr
                    Stop
                End If
            Next
            DoEvents
        Next
    'Zeige alle Pids unbenutzt in Events ClubsNations
        p = ArrC(6) + "\FreePIDs.txt"
        'show CStr(anzAinB("p", s1)) _
            + v + s1 'alle Pids unbenutzt in Events ClubsNations
        'Stop: show ReadFile(p):Stop
        writeStringToFile p, s1
        Beep
End Sub

Function Get_AllFilePathsOfFolderList_PidVid8size10Date8Folder6Path(FolderList$) As String
    Dim s$, v$, i&, Arr1() As String
    v = vbCrLf
    Arr1 = Split(FolderList, v)
    MyCountDown_MainStepsAllowed 30: MyCountDown_SubStepsMax UBound(Arr1)
    For i = 0 To UBound(Arr1)
        MyCountDown_OneMoreSubStep
        s = s + Get_AllFilePathsOfOneFolder_PidVid8Size10Date8Folder6Path(Arr1(i)) + v
    Next
    s = Delete_EmptyRowsInString(s)
    Arr1 = Split(s, v)
    QuickSort Arr1
    MyCountDown_OneMoreMainStep
    s = Join(Arr1, v)
    Get_AllFilePathsOfFolderList_PidVid8size10Date8Folder6Path = s
End Function

Function Add_FrontBlanks5(MyString$)
    'Called from    T6_Check_SamePid_SameJpg
    'Action         Stellt dem String s Leerzeichen voran, bis s die Länge 5 hat
    
    Dim s$: s = MyString
    If Len(s) = 4 Then
        s = " " + s
    ElseIf Len(s) = 3 Then s = "  " + s
    ElseIf Len(s) = 2 Then s = "   " + s
    ElseIf Len(s) = 1 Then s = "    " + s
    ElseIf Len(s) = 0 Then s = "     "
    End If
    Add_FrontBlanks5 = s
End Function

Function T6_PidsEqualGroupContainsDifferentSizes(B$) As Boolean
    'Called from    T6_Check_SamePid_SameJpg
    
    'Vorbereitung
        Dim L$, Siz$, v$, i&, Arr1() As String
        v = vbCrLf
        If anzAinB(":", B) < 2 Then Exit Function
    'Action
        B = Delete_EmptyEndRowsInString(B)
        Arr1 = Split(B, v)
        For i = 0 To UBound(Arr1)
            L = Arr1(i) 'one Line 'p0001-00|0000087544|19850000|Inside|F:\Archiv...jpg
            If i = 0 Then Siz = Mid(L, 10, 10)
            If Mid(L, 10, 10) <> Siz Then
                T6_PidsEqualGroupContainsDifferentSizes = True
                QuickSort Arr1
                B = Join(Arr1, v)
                Exit Function
            End If
        Next
End Function

Sub T6_Repair_PidsNuE2(NuE2$)
    'Called from    T6_Check_PIDs_VIDs
    'NuE2           Liste mit FilePaths, deren pId geändert werden soll (neue pId)
    
    'Vorbereitung
        If NuE2 = "" Then Exit Sub
        Dim L$, p$, PidLastNew$, PidLastOld$, PidNew$, PidOld$, v$, i%, Arr1() As String
        v = vbCrLf
    'Action
        Arr1 = Split(NuE2, v)
        MyCountDown_MainStepsAllowed 5: MyCountDown_SubStepsMax UBound(Arr1)
        For i = 0 To UBound(Arr1)
            MyCountDown_OneMoreSubStep
            L = Arr1(i)     'p1234-|F:\...
            If InStr(1, L, "-|") <> 6 Or InStr(1, L, ":") <> 9 Then Stop
            PidOld = Left(L, 6)         '"p0123-"
            
            'nur der erste wird geändert
            
            
            If PidOld = PidLastOld Then
                PidNew = PidLastNew
            Else
                PidNew = Get_NextFreePID: If Not PidNew Like "p####-00" Then Stop
                PidNew = Left(PidNew, 6) '"p1234-"
                'Logbuch
                    LogBuch "pIdOld=" + Left(PidOld, 5) + " was changed to PidNew=" + Left(PidNew, 5)
            End If
            p = Mid(L, 8)               '"F:\Archiv.."
            'Change
                RenameFile p, Replace(p, PidOld, PidNew)
            'Remember
                PidLastOld = PidOld
                PidLastNew = PidNew
        Next
End Sub

Function T6_Get_PidsNuE2_ToRepair(Liste2$, PidsL$, NuE2$) As String
    'Called from    T6_Check_PIDs_VIDs
    'Liste2          = Alle jpg/lnk-FilePaths in 'Events' und 'ClubsNations' mit PID
    '               = Zeilen wie: "p2567-00|19590709|Events|jpg|F:\..\Events\..\19590709..p2567-00 ww.jpg"
    'PidsL          = Alle jpg/lnk-FilePaths in 'Leute' mit PID
    'Nu             = pIds, not unique
    'NuE2            = pIds, die genau 2 mal vorkommen (gleiche pId für 2 verschiedene JPG);
    '               je 2 NuE2-Zeilen tragen die selbe pId, aber unterschiedliche Pfade;
    '               NuE2 beginnt mit "--------[v]p0001-00|Events|jpg|F:..."
    'Status         pIds (in 'Events' und 'ClubsNations') sollten Unikate sein;
    '               gleiche pId: gleiches Image (kommt in 'Leute' vor);
    'NuE2           Liste mit FilePaths, deren pId geändert werden soll (neue pId)
    'Action         Bereitstellung von NuE2:
    '               FilePath-Pärchen (pId gleich, Datum verschieden) werden aus NuE2
    '               herausgefiltert; Grund: pId (p2567-00) des 2. FilePath
    '               soll ersetzt werden (neue pId aus FreePIDs)
    '               und auch alle anderen Pfade mit "p2567" in 'Events', 'Inside..', 'Leute';
    '               alle diese Pfade werden in NuE2 gelistet
    
    'Vorbereitung
        Dim A$, D1$, D2$, D3$, L1$, L2$, N1$, N2$, NuE2$, p1$, Pid$, v$
        Dim i%, j%, C1&, C2&, C3&, Arr1() As String
        v = vbCrLf: MyCountDown_OneMoreMainStep
        A = v + Liste2 + v + PidsL + v 'Alle FilePaths mit pIds
    'Action
        Arr1 = Split(NuE2, v)
        MyCountDown_MainStepsAllowed 5: MyCountDown_SubStepsMax UBound(Arr1) \ 3
        For i = 1 To UBound(Arr1) - 2 Step 3
            'Schleife über alle FilePath-Pärchen aus NuE2
            MyCountDown_OneMoreSubStep
            L1 = Arr1(i): L2 = Arr1(i + 1) 'Line 1, Line 2
            Pid = Left(L1, 6)
            If Pid Like "p####-" And Left(L2, 6) = Pid Then
                'L1, L2  sind eine Pärchen (je FileName trägt gleiche pId)
                N1 = getNameOfPath(Mid(L1, 21))
                N2 = getNameOfPath(Mid(L2, 21))
                D1 = Left(N1, 8): D2 = Left(N2, 8) 'Date8
                If D1 <> D2 Then
                    'pId gleich, Datum verschieden
                    'A durchsuchen  'Alle Zeilen mit pId und D2 finden
                        C3 = 1
                        For j = 1 To 99
                            'Schleife für das Weitersuchen
                            C1 = InStr(C3, A, v + Pid)
                            If C1 > 0 Then
                                D3 = Mid(A, C1 + 11, 8)
                                If D3 = D2 Then
                                    'Zeile gefunden
                                    C2 = InStr(C1 + 10, A, v): If C2 = 0 Then Stop
                                    p1 = CutVonBis(A, C1 + 31, C2 - 1)  'one path
                                    If Not NuE2 Like "*" + p1 + "*" Then
                                        NuE2 = NuE2 + Pid + "|" + p1 + v
                                        ' = all FilePaths to change pId
                                    End If
                                End If
                            Else
                                j = 99
                            End If
                            C3 = C1 + 10
                        Next
                End If
            Else
                Stop
            End If
        Next
        NuE2 = Sort_TextLines(NuE2)
        NuE2 = Delete_EmptyRowsInString(NuE2)
    'Finals
        T6_Get_PidsNuE2_ToRepair = NuE2
        'show "NuE2: Alle " + CStr(anzAinB(":", NuE2)) + " Filepaths mit PIDs, die geändert werden sollen" + v + v + NuE2
End Function

Function T6_Get_PidsNuG2(NuGE2$, NuE2$) As String
    'Called from    T6_Check_PIDs_VIDs
    'NuGE2            = PIDs (not unique), die mehrfach vorkommen (für verschiedene JPGs)
    'NuE2            = PIDs, die genau 2 mal vorkommen (gleiche PID für 2 verschiedene JPG)
    '               je 2 Zeilen tragen die selbe PID, aber unterschiedliche Pfade
    'NuG2            = PIDs, die 3 mal oder öfter vorkommen (gleiche PID für >2 verschiedene JPG)
    '               je >2 Zeilen tragen die selbe PID, aber unterschiedliche Pfade
    
    'Vorbereitung
        Dim L$, NuG2$, v$, i%, Arr1() As String
        v = vbCrLf: NuGE2 = "--------" + v + NuGE2
        NuG2 = NuGE2: MyCountDown_OneMoreMainStep
    'Action
        Arr1 = Split(NuE2, v)
        For i = 0 To UBound(Arr1)
            L = Arr1(i) 'one Line
            If Left(L, 1) = "p" Then NuG2 = Replace(NuG2, L, "")
        Next
        NuG2 = Delete_EmptyRowsInString(NuG2)
        NuG2 = Replace99(NuG2, "--------" + v + "--------", "--------")
    'Finals
        T6_Get_PidsNuG2 = NuG2
        'show "NuG2: Alle " + CStr(anzAinB(":", NuG2)) + " Filepaths mit PIDs, die mehr als 2 mal vorkommen (gleiche PID für >2 verschiedene JPG)" + v + v + NuG2
End Function

Function T6_Get_PidsNuE2(NuGE2$) As String
    'Called from    T6_Check_PIDs_VIDs
    'NuGE2             = PIDs, die mehrfach vorkommen (für verschiedene JPGs)
    'NuE2             = PIDs, die genau 2x vorkommen (gleiche PID für 2 verschiedene JPG)
    '               je 2 Zeilen tragen die selbe PID, aber unterschiedliche Pfade
    
    'Vorbereitung
        Dim NuE2$, v$, i%, Arr1() As String
        v = vbCrLf: NuGE2 = "--------" + v + NuGE2
        NuE2 = "--------" + v
    'Action
        MyCountDown_OneMoreMainStep
        Arr1 = Split(NuGE2, v)
        For i = 0 To UBound(Arr1) - 3
            If Arr1(i) = "--------" And Arr1(i + 3) = "--------" Then
                NuE2 = NuE2 + Arr1(i + 1) + v + Arr1(i + 2) + v + "--------" + v
            End If
        Next
        NuE2 = Delete_EmptyEndRowsInString(NuE2)
    'Finals
        T6_Get_PidsNuE2 = NuE2
        'show "NuE2: Alle " + CStr(anzAinB(":", NuE2) ) + " Filepaths mit PIDs, die genau 2x vorkommen (gleiche PID für 2 verschiedene JPG)" + v + v + NuE2
End Function

Function T6_Get_PIDsVIDsNotUnique(Liste1$) As String
    'Called from    T6_Check_PIDs_VIDs
    'Liste1         = Alle FilePaths mit PID (aus 'Events' und 'Inside..')
    '               = PID|Folder|FilePath aus E+I
    'NU          = PIDs, die mehrfach vorkommen (GE2, GreaterOrEqual 2, >=2)
    'Action         Lieferung von NU
    
    'Vorbereitung
        Dim ID1$, ID2$, ID3$, NU$, v$, i%, Arr1() As String
        v = vbCrLf
    'Action
        Arr1 = Split(Liste1, v)
        QuickSort Arr1
        MyCountDown_MainStepsAllowed 2: MyCountDown_SubStepsMax UBound(Arr1)
        For i = 0 To UBound(Arr1) - 1
            'Schleife über alle PID|Folder|FilePath aus E+I
            MyCountDown_OneMoreSubStep
            ID1 = Left(Arr1(i), 8)
            ID2 = Left(Arr1(i + 1), 8)
            If ID1 = ID2 Then
               If Not NU Like "*" + Arr1(i) + v + "*" Then NU = NU + Arr1(i) + v
               If Not NU Like "*" + Arr1(i + 1) + v + "*" Then NU = NU + Arr1(i + 1) + v
               ID3 = ID1
            End If
            If ID1 = ID3 And ID2 <> ID3 Then NU = NU + "--------" + v: ID3 = ""
        Next
        NU = Delete_EmptyEndRowsInString(NU)
    'Finals
        T6_Get_PIDsVIDsNotUnique = NU
        'show "NU: Alle " + CStr(anzAinB(":", NU)) + " Filepaths mit nicht eindeutiger PID" + v + v + NU
End Function

Function T6_Check_FreePIDs(Liste1$) As String
    'Called from    T6_Check_PIDs_VIDs
    
    'Vorbereitung
        Dim fp$, FreePIDs$, Id$, p$, t1$, T2$, AnzFP%, c%, i%
        p = ArrC(6) + "\FreePIDs.txt"
    'Check
        fp = ReadFile(p): AnzFP = anzAinB("p", fp)
        MyCountDown_OneMoreMainStep
        If AnzFP < 100 Then
            t1 = Add_FrontBlanks5(500) + " neue FreePIDs wurden erzeugt"
            T6_Check_FreePIDs = t1
        Else
            T2 = Add_FrontBlanks5(CStr(AnzFP)) + " FreePIDs sind noch vorhanden; Neuproduktion nicht erforderlich"
            For i = 1 To 4: MyCountDown_OneMoreMainStep: Next
            T6_Check_FreePIDs = T2: Exit Function
        End If
    'Search
        MyCountDown_MainStepsAllowed 3: MyCountDown_SubStepsMax 500
        For i = 1 To 9999
            Id = "p" + Format(i, "0000")
            If Not Liste1 Like "*" + Id + "*" Then
                FreePIDs = FreePIDs + Id + "-00 "
                c = c + 1: MyCountDown_OneMoreSubStep
            End If
            If c = 500 Then Exit For
        Next
    'Write FreePIDs
        writeStringToFile p, FreePIDs
        LogBuch "500 free PIDs have been written to FreePIDs.txt"
End Function

Function T6_Check_FreeVIDs(Liste1$) As String
    'Called from    T6_Check_PIDs_VIDs
    
    'Vorbereitung
        Dim FV$, FreeVIDs$, Id$, p$, t1$, T2$, AnzFV%, c%, i%
        p = ArrC(6) + "\FreeVIDs.txt"
    'Check
        FV = ReadFile(p): AnzFV = anzAinB("v", FV)
        MyCountDown_OneMoreMainStep
        If AnzFV < 100 Then
            t1 = Add_FrontBlanks5(500) + " neue FreeVIDs wurden erzeugt"
            T6_Check_FreeVIDs = t1
        Else
            T2 = Add_FrontBlanks5(CStr(AnzFV)) + " FreeVIDs sind noch vorhanden; Neuproduktion nicht erforderlich"
            For i = 1 To 4: MyCountDown_OneMoreMainStep: Next
            T6_Check_FreeVIDs = T2: Exit Function
        End If
    'Search
        MyCountDown_MainStepsAllowed 3: MyCountDown_SubStepsMax 500
        For i = 1 To 9999
            Id = "v" + Format(i, "0000")
            If Not Liste1 Like "*" + Id + "*" Then
                FreeVIDs = FreeVIDs + Id + "-00 "
                c = c + 1: MyCountDown_OneMoreSubStep
            End If
            If c = 500 Then Exit For
        Next
    'Write FreeVIDs
        writeStringToFile p, FreeVIDs
        LogBuch "500 free VIDs have been written to FreeVIDs.txt"
End Function

Function Get_AllFilePathsOfOneFolder_PidVid8Size10Date8Folder6Path(PathOfOneFolder$) As String
    'Called from    xxx

    'Vorbereitung
        Dim fso As Object, objVerzeichnis As Object, objDateienliste As Object
        Dim objDatei As Object, Date8$, Folder6$, L$, Last4$, N$, p$, PidVid8$, s$, size10$
        Last4 = ".lnk.url.pdf.srt.txt.dng.ini"
    'FSO
        Set fso = CreateObject("scripting.FileSystemObject")
        Set objVerzeichnis = fso.GetFolder(PathOfOneFolder)
        Set objDateienliste = objVerzeichnis.Files
    'All FilePaths of one folder
        For Each objDatei In objDateienliste
             If Not objDatei Is Nothing Then
                p = CStr(objDatei.path) 'Path of file
                N = CStr(objDatei.NAME) 'Name of file
                If N = "a Leute.jpg" Then GoTo jump1
                L = Right(N, 4): If Last4 Like "*" + L + "*" Then GoTo jump1
                PidVid8 = Get_PidOrVid_FromFileName(N)
                If PidVid8 = "" Then GoTo jump1
                size10 = Get_size10_FromFilePath(p)
                
                'GoTo jump1
                
                Date8 = Get_Date8_FromFileName(N)
                Folder6 = Get_Folder6_FromFilePath(p)
                s = s + PidVid8 + "|" + size10 + "|" + Date8 + "|" + Folder6 + "|" + p + vbCrLf
jump1:
             End If
        Next objDatei
        s = Delete_EmptyRowsInString(s)
    'Finals
        Get_AllFilePathsOfOneFolder_PidVid8Size10Date8Folder6Path = s
End Function

Function Get_size10_FromFilePath(p$) As String
    Get_size10_FromFilePath = Format(FileSize(p), "0000000000")
End Function

Function Get_Folder6_FromFilePath(p$) As String
    If p Like "*\Events\*" Then
        Get_Folder6_FromFilePath = "Events"
    ElseIf p Like "*\ClubsNations\*" Then Get_Folder6_FromFilePath = "Inside"
    ElseIf p Like "*\Leute\*" Then Get_Folder6_FromFilePath = "Leute_"
    Else
        Get_Folder6_FromFilePath = "eeeeee"
    End If
End Function

Function Get_Date8_FromFileName(N$) As String
    If Left(N, 8) Like "########" Then Get_Date8_FromFileName = Left(N, 8) Else Get_Date8_FromFileName = "dddddddd"
End Function

Function T6_Get_PidPath_OfEachFileInOneFolder(s1$, Optional LikeString$ = "* p####-## *") As String
    'Called from    T6_Check_PIDs_VIDs
    's1             Liste aller Folder-/SubfolderPaths in 'Events'
    's2             Liste aller Folder-/SubfolderPaths in 'Inside..'
    'Action         Erstellung Liste, je Zeile mit pId|Ev|Path, z. B.:
    '               "p2549-00|Inside|F:\Archiv..\ClubsNations\..\..\19591000..p2549-00 ww.jpg"
    
    'Vorbereitung
        Dim Date8$, Ev$, Id$, jpg$, L$, N$, p$, s3$, s4$, v$, i%, Arr1() As String
        v = vbCrLf
    'Action
        Arr1 = Split(v + s1, v)
        'showArray Arr1
        'LikeString
            L = LikeString
        MyCountDown_MainStepsAllowed 30: MyCountDown_SubStepsMax UBound(Arr1)
        For i = 1 To UBound(Arr1)
            p = Arr1(i)         'OneFolderPath
            'L = "* p####-## *"  'Like-String
            MyCountDown_OneMoreSubStep
            s3 = s3 + Get_AllFilePaths_LikeMyStringInFileName_OfOneFolder(p, L) + v
            
            
            'Get_AllFilePaths_LikeX1_NotLikeX2_InFileName_OfOneFolder
            
            
        Next
        s3 = Delete_EmptyRowsInString(s3)
        'show s3
    'Add Id and Event|Inside leftmost to each FilePath
        Arr1 = Split(s3, v)
        MyCountDown_MainStepsAllowed 2: MyCountDown_SubStepsMax UBound(Arr1)
        For i = 0 To UBound(Arr1)
            MyCountDown_OneMoreSubStep
            p = Arr1(i) 'one path of a file containing "p####-##"
            N = Get_NameFromPath(p)
            Date8 = Left(N, 8): If Not Date8 Like "########" Then Stop: Date8 = "xxxxxxxx"
            Id = Get_Pid_FromFileName(N)
            If p Like "*\Events\*" Then Ev = "Events"               'Len(Ev) = 6
            If p Like "*\ClubsNations\*" Then Ev = "Inside"  'Len(Ev) = 6
            If p Like "*\Leute\*" Then Ev = "Leute_"                'Len(Ev) = 6
            jpg = Right(N, 3)
            Arr1(i) = Id + "|" + Date8 + "|" + Ev + "|" + jpg + "|" + Arr1(i)
        Next
        s4 = Join(Arr1, v)
    'Finals
        T6_Get_PidPath_OfEachFileInOneFolder = s4
End Function

Function Get_PidOrVid_FromFileName(fileName$) As String
    Dim ext$, s$, c%, i%
    ext = Get_FileExtension(fileName)
    s = Replace(fileName, "." + ext, "")
    s = Right(s, 11): s = Left(s, 8)
    If s Like "[p|v]####-##" Then Get_PidOrVid_FromFileName = s Else Get_PidOrVid_FromFileName = ""
End Function

Function Get_Vid_FromFileName(fileName$) As String
    Dim s$
    s = Get_PidOrVid_FromFileName(fileName$)
    If s Like "v####-##" Then Get_Vid_FromFileName = s
End Function

Function Get_Pid_FromFileName(fileName$) As String
    Dim s$
    s = Get_PidOrVid_FromFileName(fileName$)
    If s Like "p####-##" Then Get_Pid_FromFileName = s
End Function

Sub Show_NextFreeID_TrampolinPhotos()
    'Called from    [none]

    Dim p1$, p2$, s$, s1$, i&
    p1 = "G:\Archiv Photos\- Photos ubes\Photos Trampolin\Events"
    p2 = "G:\Archiv Photos\- Photos ubes\Photos Trampolin\Events\0000 Info\show last ID used"
    
    s = Get_Paths_ofAllFiles_withID_inFolderAndSubFolders(p1)
    For i = 1 To 999999
        If InStr(1, s, "id" + Format(i, "000000")) = 0 Then Exit For
    Next
    show "id" + Format(i, "000000")
End Sub

Sub TEST_Get_Paths_ofAllFiles_withID_inFolderAndSubFolders()
    Dim p1$, p2$, p3$, s$
    DoArrc
    p1 = ArrC(3)                'T0 Path of Folder "Events"
    p2 = ArrC(4)                'T0 Path of Folder "Leute"
    p3 = ArrC(11)               'T0 Path of Folder "ClubsNations"
    s = Get_Paths_ofAllFiles_withID_inFolderAndSubFolders(p1)
    show s
End Sub

Function Get_Paths_ofAllFiles_withID_inFolderAndSubFolders(PathOfSourceFolder$) As String
    'Called from    Show_NextFreeID_TrampolinPhotos

    Dim oFSO As Object, oFolder As Object, oFile As Object, sF
    Dim s$, i%, colFolders As New Collection, C1%, C2%, N$, Arr1() As String
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(PathOfSourceFolder)
    colFolders.Add oFolder          'start with this folder
    
    Do While colFolders.count > 0      'process all folders
        Set oFolder = colFolders(1)    'get a folder to process
        colFolders.Remove 1            'remove item at index 1
        For Each oFile In oFolder.Files
            N = oFile.NAME
            
            C1 = InStr(1, N, " id")
            If C1 > 0 Then
                C2 = InStr(1, N, "-00")
                If C2 > 0 Then
                    s = s + Mid(N, C1 + 1, 8) + vbCrLf
                    i = i + 1
                End If
            End If
        Next oFile
        'add any subfolders to the collection for processing
        For Each sF In oFolder.subfolders
            colFolders.Add sF
        Next sF
    Loop
    s = Delete_EmptyEndRowsInString(s)
    Arr1 = Split(s, vbCrLf)
    QuickSort Arr1
    s = Join(Arr1, vbCrLf)
    Get_Paths_ofAllFiles_withID_inFolderAndSubFolders = s
End Function

Function Get_Paths_OfAllFiles_WithDateIdJpg_InFolderAndSubFolders(PathOfSourceFolder$) As String
    'Called from    xxx

    Dim oFSO As Object, oFolder As Object, oFile As Object, sF
    Dim colFolders As New Collection
    Dim s$, N$, Arr1() As String
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(PathOfSourceFolder)
    colFolders.Add oFolder          'start with this folder
    
    Do While colFolders.count > 0      'process all folders
        Set oFolder = colFolders(1)    'get a folder to process
        colFolders.Remove 1            'remove item at index 1
        For Each oFile In oFolder.Files
            N = oFile.NAME
            If N Like "######## * p####-## [a-z][a-z].jpg" Then
                'File besitzt korrekten Anfang + Ende
                    
                s = s + oFolder.path + "\" + N + vbCrLf
                
            End If
        Next oFile
        'add any subfolders to the collection for processing
        For Each sF In oFolder.subfolders
            colFolders.Add sF
        Next sF
    Loop
    s = Delete_EmptyEndRowsInString(s)
    'arr1 = Split(s, vbCrLf)
    'QuickSort arr1
    's = Join(arr1, vbCrLf)
    's = "Anzahl JPGs = " + CStr(c1) + vbCrLf + vbCrLf + s
    Get_Paths_OfAllFiles_WithDateIdJpg_InFolderAndSubFolders = s
End Function

Function Get_NextFreePID() As String
    'Called from    T4_GetNameOfResultJpg, xxx
        Dim p$, Pid$, s$, Anz%
        Call DoArr: p = ArrC(6) + "\FreePIDs.txt"
    'Die Datei 'FreePIDs.txt' wird in s eingelesen
        s = ReadFile(p) '"p2563-00 p2593-00 ... p3396-00 "
        Anz = anzAinB("-", s)
        If Anz < 100 Then MsgBox "Es stehen nur noch " + CStr(Anz) + " FreePIDs zu Verfügung." + vbCrLf + "Bitte neue FreePIDs erzeugen."
        If Anz < 9 Then Stop
    'Die erste pID wird aus s entnommen
        Pid = Left(s, 8): If Not Pid Like "p####-00" Then Stop
        s = Mid(s, 10)
    'Reduzierte FreePIDs-Liste zurückschreiben
        writeStringToFile p, s
    Get_NextFreePID = Pid
End Function

Sub Get_NextFreeVID_TEST()
    show Get_NextFreeVID
End Sub

Function Get_NextFreeVID() As String
    'Called from    T4_GetNameOfResultJpg, xxx
        Dim p$, s$, vID$, Anz%
        Call DoArr: p = ArrC(6) + "\FreeVIDs.txt"
    'Die Datei 'FreeVIDs.txt' wird in s eingelesen
        s = ReadFile(p) '"v0164-00 v0165-00 ... v0663-00 "
        Anz = anzAinB("-", s)
        If Anz < 100 Then MsgBox "Es stehen nur noch " + CStr(Anz) + " FreeVIDs zu Verfügung." + vbCrLf + "Bitte neue FreeVIDs erzeugen."
        If Anz < 9 Then Stop
    'Die erste VID wird aus s entnommen
        vID = Left(s, 8): If Not vID Like "v####-00" Then Stop
        s = Mid(s, 10)
    'Reduzierte FreeVIDs-Liste zurückschreiben
        writeStringToFile p, s
    Get_NextFreeVID = vID
End Function

Sub Add_pID_toFileName_inEvents(z%, s%)
    'Called from    [Klick auf T1-CellButton 'Add ID toFileName of Photos']
    'pID            = Photo-ID = "p1234-56" oder andere Ziffern
    'vID            = Video-ID = "v1234-56" oder andere Ziffern
    'Action         sucht nach FileNames ohne ID;
    '               setzt eine ID vor das Quellenkürzel;
    '               sucht nur in 'Events', nicht in 'Leute'
        Dim p1$, p2$, r$, s1$, s2$, s3$, s4$, s5$, s6$, T$, SourceFolder$
        Dim ANZAHL%, c%, i%, Arr1() As String, Arr2() As String
    'Einstellungen
        ANZAHL = 300 'neue Ids müssen vorher ermittelt werden
    'T1-Anzeige Date
        EE 0
        With ActiveWorkbook.Sheets("T1")
        .Cells(z - 1, s - 1) = "." 'Marker "Add_pID is running"
        T = Format(Now(), "dd.mm.yyyy hh:mm:ss") + " Uhr"
        .Cells(z + 6, s) = T        'Neues Datum schreiben
        .Cells(z + 7, s) = ""       'Text entfernen
        .Cells(z + 8, s) = ""       'Text entfernen
        .Cells(z + 7, s + 11) = ""  'Zahl entfernen
        .Cells(z + 8, s + 11) = ""  'Zahl entfernen
        .Cells(z + 7, s + 12) = ""  'Erledigt-Haken entfernen
        .Cells(z + 8, s + 12) = ""  'Erledigt-Haken entfernen
        .Cells(z + 10, s).Font.Color = RGB(150, 150, 150)
    'Alle File-Pfade holen
        SourceFolder = CStr(ArrC(3))  'Path of Folder "Events"
        s1 = Get_FilePaths_Like_insideSourceFolderAndSubFolders(SourceFolder)
        'show s1
    'Nur .jpg, .mp4, .avi, .mov, .wmv, .mpg, ...
        .Cells(z + 8, s) = "Nur Photos und Videos"  'xxx
        Arr1 = Split(s1, vbCrLf)
        s1 = ""
        For i = 0 To UBound(Arr1)
            r = LCase(Right(Arr1(i), 4))
            If r <> ".txt" And r <> ".lnk" And r <> ".pdf" And r <> ".ini" Then
                s1 = s1 + Arr1(i) + vbCrLf
            End If
        Next
        s1 = Delete_EmptyEndRowsInString(s1)
        'show  s1
        
    'Get new pIDs
        .Cells(z + 8, s) = "Neue IDs"  'T1-Anzeige
        For i = 1 To 9999
            s2 = "p" + Format(i, "0000") + "-00"    's2 = "p1234-00"
            s3 = " " + s2 + " "                     's3 = " p1234-00 "
            If InStr(1, s1, s3) = 0 Then
                's3 (eine pID) kommt in s1 (alle FilePfade) nicht vor
                c = c + 1
                s4 = s4 + s2 + vbCrLf               's2 (neue pID) --> s4-Sammlung
                If c = ANZAHL Then Exit For
            End If
        Next
        Arr2 = Split(s4, vbCrLf)
        'show s4 '= p0144-00, p0145-00, ...
        c = 0
   'QuellenKennung hinzufügen (falls nicht vorhanden)
        .Cells(z + 8, s) = "QuellenKennung hinzufügen"  'T1-Anzeige
        Arr1 = Split(s1, vbCrLf)
        For i = 0 To UBound(Arr1)
            p1 = Arr1(i)         '1 FilePath
            If Right(p1, 4) = ".jpg" Then
                If Not (Right(p1, 7) Like " [a-z][a-z].jpg") Then
                    'p1 = alter Pfad ohne QuellenKennung
                    'p2 = neuer Pfad mit  QuellenKennung:
                    p2 = Replace(p1, ".jpg", " nn.jpg")
                    RenameFile p1, p2
                    Arr1(i) = p2
                    's6 = s6 + p1 + vbCrLf + p2 + vbCrLf + vbCrLf
                End If
            End If
        Next
        'show s6
    'pID hinzufügen (falls nicht vorhanden)
        .Cells(z + 8, s) = "Hinzugefügte IDs:"  'T1-Anzeige
        .Cells(z + 8, s + 11) = "'0"            'T1-Anzeige
        'arr1 = Split(s1, vbCrLf)
        For i = 0 To UBound(Arr1)
            p1 = Arr1(i)         '1 FilePath
            r = Right(p1, 7)        'r = " nn.jpg"
            If r Like " [a-z][a-z].jpg" Then
                'ID kann pID oder vID sein (vID bei Video-snapshots)
                If Not (p1 Like "* p####-## *" Or p1 Like "* v####-## *") Then
                    If Not (p1 Like "* p####-## [a-z][a-z].jpg" Or p1 Like "* v####-## [a-z][a-z].jpg") Then
                        'FilePath p1 hat noch keine ID
                        c = c + 1
                        's5 = s5 + p1 + vbCrLf
                        'p1 = alter Pfad ohne ID
                        'p2 = neuer Pfad mit  ID:
                        p2 = Replace(p1, r, " " + Arr2(c) + r)
                        s5 = s5 + p2 + vbCrLf + vbCrLf
                        RenameFile p1, p2
                        .Cells(z + 8, s + 11) = "'" + CStr(c)   'T1-Anzeige
                        If c = ANZAHL Then Exit For
                    End If
                End If
            End If
        Next
    'finals
        .Cells(z + 8, s + 12) = "ü" 'Erledigt-Haken setzen
        .Cells(z - 1, s - 1).Select 'SelectCell Green1 li ob
        .Cells(z - 1, s - 1) = ""   'Marker "Add_pID is not running"
        EE 1
        p1 = ArrC(6) + "\ImageFilesWithNewIDs.txt" 'ArrC(6) = Path of Folder "helpers"
        writeStringToFile p1, s5
        If c = 0 Then
            .Cells(z + 10, s).Font.Color = RGB(150, 150, 150)
        Else
            .Cells(z + 10, s).Font.Color = RGB(0, 0, 0)
        End If
        End With
        Beep
        'show "Umbenennungen:" + vbCrLf + s5
End Sub

Sub ShowMissingVIDs_toFileName()
    'Called from    [none]
    'pID            = Photo-ID = "p1234-56" oder andere Ziffern
    'vID            = Video-ID = "v1234-56" oder andere Ziffern
    'Action         sucht nach FileNames ohne ID;
    '               setzt eine ID vor das Quellenkürzel;
    '               sucht nur in 'Events', nicht in 'Leute'
        Dim p1$, p2$, r$, s1$, s2$, s3$, s4$, s5$, s6$, SourceFolder$
        Dim ANZAHL%, c%, i%, Arr1() As String, Arr2() As String
    'Einstellungen
        SourceFolder = "F:\Archiv Trampo\Archiv Trampolin\Events"
        ANZAHL = 50 'neue Ids müssen vorher ermittelt werden
    'Alle File-Pfade holen
        s1 = Get_FilePaths_Like_insideSourceFolderAndSubFolders(SourceFolder)
        'show s1
    'Get pIDs
        'arr1 = Split(s1, vbCrLf)
        For i = 1 To 9999
            s2 = "v" + Format(i, "0000") + "-00"
            s3 = " " + s2 + " "
            If InStr(1, s1, s3) = 0 Then
                c = c + 1
                s4 = s4 + s2 + vbCrLf
                If c = ANZAHL Then Exit For
            End If
        Next
        Arr2 = Split(s4, vbCrLf)
        'show s4 '= p0144-00, p0145-00, ...
        c = 0
    'Who has no pID?
        Arr1 = Split(s1, vbCrLf)
        For i = 0 To UBound(Arr1)
            p1 = Arr1(i)         '1 FilePath
            r = Right(p1, 7)
            If r Like " [a-z][a-z].mp4" Or r Like " [a-z][a-z].mpg" _
                Or r Like " [a-z][a-z].avi" Or r Like " [a-z][a-z].wmv" Then
                'FileName hat eine Quellen-Kennung
                If Not p1 Like "* v####-## *" Then
                    If Not (p1 Like "* v####-## [a-z][a-z].mp4" _
                        Or p1 Like "* v####-## [a-z][a-z].avi") Then
                        c = c + 1
                        s5 = s5 + p1 + vbCrLf
                        'Neuer Pfad:
                        p2 = Replace(p1, r, " " + Arr2(c) + r)
                        s5 = s5 + p2 + vbCrLf + vbCrLf
                        
                        RenameFile p1, p2
                        
                        If c = ANZAHL Then Exit For
                    End If
                End If
            Else
                If r Like "*.mp4" Or r Like "*.avi" Then
                    'FileName hat keine Quellen-Kennung
                    
                    s6 = s6 + p1 + vbCrLf
                End If
            End If
        Next
        
        show "Umbenennungen:" + vbCrLf + s5
        show "FileName hat keine Quellen-Kennung:" + vbCrLf + s6
End Sub

Sub ShowListOfImageFilesWithNewIDs()
    'Vorbereitung
        Dim p1$, s$, T$, i%, Arr1() As String
        DoArrc
        p1 = ArrC(6) + "\ImageFilesWithNewIDs.txt"  'Path of Folder "helpers"
        T = "Liste von Photo-Dateipfaden, zu deren Dateinamen eine ID hinzugefügt wurde:" + vbCrLf + vbCrLf
        s = ReadFile(p1): s = Delete_EmptyRowsInString(s)
        If s = "" Then show T + "- keine -": Exit Sub
    'Action
      Arr1 = Split(vbCrLf + s, vbCrLf)
      For i = 1 To UBound(Arr1)
        Arr1(i) = "[" + Format(i, "000") + "] " + Arr1(i)
      Next
      show T + Join(Arr1, vbCrLf)
End Sub

Sub Add_IDs_toFilesInOneEventFolder(PathOfOneEventFolder$, C2%)
    'Called from    T1_UpdateStatistics, T6_Add_IDs_toFilesInOpenEventFolders
    'Scope          erreicht Events/ClubsNations mit FolderName = "####..." (1985...)
    '               FileName ohne ####-## .lnk .url .txt !
    
    'Vorbereitung
        If Not (PathOfOneEventFolder Like "*\Events\*" _
            Or PathOfOneEventFolder Like "*\ClubsNations\*") Then Exit Sub
        Dim f1$, f2$, N1$, p$, p3$, s$, v$, c%, i%
        Dim Arr1() As String
        v = vbCrLf: DoArr: p3 = ArrC(3)              'p3 = Path of Folder "Events"
        p = PathOfOneEventFolder
    's = FilePaths of folder p
        s = Get_AllFilePaths_WithMyStringInFileName_OfOneFolder(p, ".")
        ' = Paths of all files inside one EventFolder
        ' = [Zeilen wie:]
        '    F:\...\Events\...\19641112 DM Berlin a00.url
        '    F:\...\Events\...\19641112 DM Berlin a01.txt
        '    F:\...\Events\...\19641112 DM Berlin a02.mp4
        '    F:\...\Events\...\19641112 DM Berlin a03.jpg
        '    F:\...\Events\...\19641112 DM Berlin a06 (Klaus Förster).jpg
        '    F:\...\Events\...\19641112 DM Berlin ... sf.jpg
        '    F:\...\ClubsNations\D\ ....jpg
        'show s: Stop
    's ohne ####-## .lnk .url .txt !
        Arr1 = Split(s, v): s = ""
        For i = 0 To UBound(Arr1)
            f1 = Arr1(i) 'PathOfOneFile
            N1 = Get_NameFromPath(f1)
            If Not (N1 Like "*!*" Or N1 Like "*####-##*" Or N1 Like "*.lnk" Or N1 Like "*.url" Or N1 Like "*.txt") Then
                s = s + v + Arr1(i)
            End If
        Next
        s = Mid(s, 3)
        
    'VideoHandling
        'zuerst muss ermittelt werden, ob ein VideoFile ohne vID existiert
        '(pIDs werden dann noch nicht vergeben, da JPGs ggf. VideoSnaps sind)
        Arr1 = Split(s, v)
        For i = 0 To UBound(Arr1)
            'Schleife über alle Files im EventFolder p
            f1 = Arr1(i)                            'f1 = Path of one File (.JPG?)
            If isVideo(f1) Then
                If Not f1 Like "* v####-## [a-z][a-z].*" Then
                    'Video f1 besitzt keine VideoID
                    If f1 Like "* [a-z][a-z].*" Then
                        'Video f1 besitzt ein Quellenkürzel, kann also eine ID erhalten
                        c = InStrRev(f1, ".")
                        f2 = Left(f1, c - 3) + Get_NextFreeVID + Mid(f1, c - 3)
                        RenameFile f1, f2: C2 = C2 + 1
                        FillArrC 77, "1": LogBuch "vID was added: " + f2
                        Exit Sub 'OpenFolder p
                    Else
                        'Video f1 besitzt kein Quellenkürzel
                        show "Proc: Add_IDs_toFilesInOneEventFolder - Reperaturbedarf" + v + v _
                           + "Das Video " + v + "   " + f1 + v + "besitzt kein Quellenkürzel." + v _
                           + "Bitte fügen sie ein solches in den Dateinamen ein." + v _
                           + "Erst danach kann das Video eine ID erhalten."
                        Exit Sub 'OpenFolder p
                    End If
                End If
            End If
        Next i
        'Ein VideoFile ohne vID existiert nicht.
        'VideoSnaps sind (hoffentlich) bereits mit vIDs versehen (*v###-##*.jpg)
        'pIDs können also vergeben werden (an JPGs ohne pID, mit Quellenkürzel)
        
        'pID für JPGs
        For i = 0 To UBound(Arr1)
            f1 = Arr1(i)                            'f1 = Path of one File (.JPG?)
            If LCase(Right(f1, 4)) = ".jpg" Then
                f2 = Left(f1, Len(f1) - 4) + ".jpg" 'f2 = Path of one File (.jpg)
                'f2 ist ein jpg-File, 'jpg' klein geschrieben (nur in f2, nicht im File selbst)
                If f2 Like "*\####* [a-z][a-z].jpg" Then
                    'f2 = FilePath (hat QuellenKürzel)
                    If Not f2 Like "* [p|v]####-## *" Then
                        'f2 hat noch keine pID; pID hinzufügen
                        f2 = Left(f2, Len(f2) - 6) + Get_NextFreePID + Mid(f2, Len(f2) - 6)
                        RenameFile f1, f2: C2 = C2 + 1
                        FillArrC 77, "1": LogBuch "pID was added: " + f2
                    End If
                Else
                    'Photo f1 besitzt kein Quellenkürzel
                    If Not f1 Like "*\a Leute.jpg" Then 'kein Quellenkürzel nötig
                        show "Proc: Add_IDs_toFilesInOneEventFolder - Reperaturbedarf" + v + v _
                           + "Das Photo " + v + "   " + f1 + v + "besitzt kein Quellenkürzel." + v _
                           + "Bitte fügen sie ein solches in den Dateinamen ein." + v _
                           + "Erst danach kann das Photo eine ID erhalten."
                        OpenFolder p: Exit Sub
                    End If
                End If
            End If
        Next i
    Beep
End Sub

Sub MyCountDown_Init(InitStr$)
    'Called from    xxx
    'InitStr        = "|T6|24|21|99|" = |SheetName|z|s|StartNr| of CountDownShowCell
    '               aktuelle CountDownNr wird in Sheets(SheetName).Cells(z, s) geschrieben
    'Action         Ermöglicht eine CountDown-Anzeige in einer CountDownShowCell
    '               (einer bestimmten Zelle eines bestimmten Blattes)
    '               über mehrere Subs/Functions hinweg
    'Usage          (1) In einer Prozedur wird MyCountDown_Init aufgerufen
    '                   (MyCountDown_Init "|T6|24|21|99|")
    '               (2) Vor einer Schleife (in Prozedur oder Sub, in welcher der CountDown
    '                   weitergezählt werden soll) wird vermerkt, wie viele CountDown-Schritte
    '                   diese Schleife verbrauchen soll (MyCountDown_MainStepsAllowed 9)
    '                   und wieviel SchleifenDurchläufe anstehen (MyCountDown_SubStepsMax 199)
    '               (3) In der Schleife werden SubSteps mitgezählt (MyCountDown_OneMoreSubStep)

    'Vorbereitung
        Dim r As Range, Arr1() As String
        Arr1 = Split(InitStr, "|"): DoArrc
    'ArrC
        FillArrC 24, "0"                'T0 MyCountDown CountOfChanges
        FillArrC 25, CStr(Arr1(4))      'T0 MyCountDown MainSteps RunDownNr
        FillArrC 26, "0"                'T0 MyCountDown MainSteps AllowedToRunDown
        FillArrC 27, "0"                'T0 MyCountDown SubSteps  Max
        FillArrC 28, "0"                'T0 MyCountDown SubSteps  RunUpNr
        FillArrC 29, CStr(Arr1(1))      'T0 MyCountDown ShowCell  NameOfSheet
        FillArrC 30, CStr(Arr1(2))      'T0 MyCountDown ShowCell  NrOfRow    (z)
        FillArrC 31, CStr(Arr1(3))      'T0 MyCountDown ShowCell  NrOfColumn (s)
    'CountDownStartNr in CountDownShowCell setzen
        Sheets(Arr1(1)).Cells(CInt(Arr1(2)), CInt(Arr1(3))) = CInt(Arr1(4))
    'CellButton (Green3) auf Green4 setzen
        If ArrC(29) = "T6" Then 'nur, falls CountDownShowCell in T6
            With Sheets("T6")
            Set r = .Range(.Cells(CInt(ArrC(30)), 2), .Cells(CInt(ArrC(30)), 15)): r.Interior.Color = Green4
            T6_CleanOldDoneRemarks
            End With
        End If
End Sub

Sub T6_CleanOldDoneRemarks()
    'Called from    T6_Update_AgeOfAllCompetitorsInsideAllDgs, ...
    
    'Vorbereitung
        Dim z%, r As Range
        z = CInt(ArrC(30))          'T6-CountDownShowCell-NrOfRow
        With Sheets("T6")
    'LastDone   'neues Datum setzen
        Set r = .Cells(z, 18): r.Value = Format(Now(), "yyyymmdd_hhmmss")
        r.Font.size = 6: r.HorizontalAlignment = xlLeft
    'Erledigt-Häkchen, Duration löschen
        .Cells(z, 16) = "": .Cells(z, 21) = ""
    'Finals
        End With
End Sub

Sub MyCountDown_MainStepsAllowed(MainStepsAllowed%)
    FillArrC 26, CStr(MainStepsAllowed)
End Sub

Sub MyCountDown_OneMoreMainStep()
    'Called from    xxx
    'Action         Neue CountDownNr wird angezeigt
    
    'Vorbereitung
        Dim MainStepNr%, z%, r As Range
    'Action
        MainStepNr = CInt(ArrC(25)): MainStepNr = MainStepNr - 1: ArrC(25) = CStr(MainStepNr)
        'DoArrC
        Set r = Sheets(ArrC(29)).Cells(CInt(ArrC(30)), CInt(ArrC(31))) 'CountDownShowCell
        r.Value = MainStepNr
    'Farbwechsel CountDown-Zelle
        If r.Interior.Color = Green1 Then r.Interior.Color = Green3 Else r.Interior.Color = Green1
        RefreshScreen
End Sub

Sub MyCountDown_OneMoreSubStep()
    'Called from    xxx
    'SubStepsMax          = 1234 = Schleifendurchläufe
    
    'Vorbereitung
        Dim C1&, MainStepsAllowed&, SubStepsMax&, TpNrHighest&, SubStepNr&, r As Range
    'Action
        'One more SubStep (SubStepNr wird von 1 zu SubStepsMax\MainStepsAllowed hochgezählt)
        SubStepNr = ArrC(28): SubStepNr = SubStepNr + 1: ArrC(28) = SubStepNr
        'T6-Anzeige
            SubStepsMax = CLng(ArrC(27))      'fester Wert; Schleifendurchläufe TrackedProcedure
            MainStepsAllowed = CLng(ArrC(26))       'CountDown RunDownSteps
            TpNrHighest = SubStepsMax \ MainStepsAllowed
            If SubStepNr = TpNrHighest Then
                'Neue CountDownNr anzeigen
                    MyCountDown_OneMoreMainStep
                SubStepNr = 0: ArrC(28) = "0"
            End If
End Sub

Sub MyCountDown_SubStepsMax(SubStepsMax&)
    'A new procedure was contacted
    FillArrC 27, CStr(SubStepsMax)
    FillArrC 28, "0"        'SubStepNr
End Sub

Sub MyCountDown_TestMain()
    DoArrc
    MyCountDown_Init "|T6|24|21|99|" '= |SheetName|z|s|StartNr| of CountDownShowCell
    MyCountDown_TestSub
End Sub

Sub MyCountDown_TestSub()
    Dim s$, i%
    MyCountDown_MainStepsAllowed 19: MyCountDown_SubStepsMax 199
    For i = 1 To 199
        MyCountDown_OneMoreSubStep
        s = s + CStr(i) + "|": Wait_MilliSec 10
    Next
    'show s
End Sub

Sub MyCountDown_End(NameOfButton$)
    'Called from    xxx
    
    'Vorbereitung
        Dim N$, z%, s%
        N = ArrC(29)            'T0 MyCountDown ShowCell-NameOfSheet
        z = CInt(ArrC(30))      'T0 MyCountDown ButtonLine-NrOfRow
        s = CInt(ArrC(31))      'T0 MyCountDown ShowCell-NrOfColumn
        Sheets(N).Cells(z, 1).Select
    'Clear CountDownShowCell
        Sheets(N).Cells(z - 1, s) = "" 'clear CountDownShowCell
    'Clear ArrC/MyCountDownItems
        FillArrC 25, "0": FillArrC 26, "0": FillArrC 27, "0": FillArrC 28, "0": FillArrC 29, "0": FillArrC 30, "0": FillArrC 31, "0"
End Sub




