Attribute VB_Name = "M_Brac"
Option Explicit 'M_Brac

Sub zzz_M_Brac()
    
    showProcs "square"
        
    EE 1: Beep
End Sub

Sub Check_Bracs_OfSomeFolders()
    Dim p$, v$, i%, max%, Arr1() As String
    v = vbCrLf: DoArrc
    p = Get_Paths_OfAllSubfolders_AllLevels_Collection(ArrC(3)) 'ArrC(3): Events
    FillArrC 87, "0"
    Arr1 = Split(v + p, v): max = UBound(Arr1)
    UF3_CountDown 2, "Check " + CStr(max) + " Folders"
        UF3.B1 = "Folder": UF3.B2 = "BracName"
    For i = 1 To 100 'UBound(Arr1)
        Check_Bracs_OfOneFolder Arr1(i), max, i
        'Check_...
        
        
        'DoEvents
        'UF3_ShowLines x    'zeigt midestens x Zeilen
    Next
    UF3.B2 = "BracNames checked total: " + ArrC(87)
    UF3.Hide
    Beep
End Sub

Sub xx()
    Dim p$: p = "F:\Archiv TR\Archiv Trampolin\Events\1959-12-14 Übungsleiter-Lehrgang Frankfurt"
    Check_Bracs_OfOneFolder p, 1, 1
End Sub

Sub Check_Bracs_OfOneFolder(PathOfActualFolder$, max%, c%)
    'Called from    Check_Bracs_OfSomeFolders
    
    'Vorbereitung
        Dim FiNaBr$, FiNaId$, NaFo$, VnNns$, i%, A() As String
        If Not (PathOfActualFolder Like "*\Events\*" Or PathOfActualFolder Like "*\ClubsNations\*") Then Exit Sub
        NaFo = getNameOfPath(PathOfActualFolder)
        With UF3: DoArrc
        .B1 = "Folder " + CStr(c) + " of " + CStr(max) + ":   '" + NaFo + "'": .C1 = max - c
    'Get BracFiles
        FiNaBr = Get_AllNamesOfBracFiles_insideOneFolder(PathOfActualFolder)
        If Not FiNaBr Like "*(*" Then Exit Sub 'No BracFiles in actual folder
        'All Brac-VnNn are T5-VnNn
        'show FiNaBr: Stop '19610000 LK D-GB Kiel - a03 (Peter Quinney) RAF p3355-01 ww.jpg
    'Get VnNns
        VnNns = Get_VnNns_OfSomeBracFileNames_WithT5VnNn(FiNaBr)
        'show VnNns '|Peter Quinney|...|Roland Schillinger|
    'Get all FilesOfActualFolder with Id
        FiNaId = Get_AllFileNames_Like_OfOneFolder(PathOfActualFolder, "* [pv]####-## *")
        'show FiNaId '"... - a06 (Jo May) p1234-05 ww.jpg", "... p1234-00 ww.jpg", "..."
        'Stop
    'Check each BracVnNn
        A = Split(VnNns, "|")
        .B2 = "BracName 1 of " + CStr(UBound(A) - 1)
        For i = 1 To UBound(A) - 1
            ArrC(87) = CStr(CInt(ArrC(87) + 1))
            UF3.B2 = "BracNames checked total: " + ArrC(87) _
                + "; BracName " + CStr(i) + " of " + CStr(UBound(A) - 1) _
                + " inside folder " + CStr(c) + ":    '" + A(i) + "'"
                'If CInt(ArrC(87)) Mod 5 = 0 Then DoEvents
                'If UBound(A) > 5 And UBound(A) Mod 5 = 0 Then DoEvents
                'If UBound(A) > 10 Then DoEvents
                DoEvents
            Check_OneBracVnNn PathOfActualFolder, FiNaId, A(i)
        Next
        End With
        Beep
End Sub

Sub Check_OneBracVnNn(PathOfActualFolder$, FiNaId$, VnNn$)
    'Called from    Check_Bracs_OfOneFolder
    'FiNaId         = "... - a06 (Jo May) p1234-05 ww.jpg", "... p1234-00 ww.jpg", "..."
    
    'Vorbereitung
        Dim Fi00$, FiNa$, FiNas$, Id$, Id00$, mw$, NameOfIcon$
        Dim NameOfPersonfolder$, NnVn$, OneFileName$, p1$, p2$, PathOfIcon$
        Dim PathOfPersonfolder$, v$, c%, i%, Arr1() As String
        v = vbCrLf
    'Get NameOfPersonfolder of VnNn
        NameOfPersonfolder = T5_Get_NameOfPersonfolder_Give_VnNn(VnNn)
        PathOfPersonfolder = ArrC(4) + "\" + NameOfPersonfolder
    'Get PathOfIcon
        NameOfIcon = Replace(NameOfPersonfolder, ",", "§") + ".ico"
        PathOfIcon = ArrC(4) + "\zzico\" + NameOfIcon
        If Not FileExists(PathOfIcon) Then
            'Es existiert noch kein icon zu VnNn    'm/w-Icon
            mw = T5_Get_mw_Give_VnNn(VnNn)
            If mw = "w" Then CopyFile ArrC(4) + "\zzico\zzw.ico", PathOfIcon Else CopyFile ArrC(4) + "\zzico\zzm.ico", PathOfIcon
        End If
    'Check NnVn-Links in actual Folder (sichert Existenz korrekter VbsLinks)
        NnVn = Get_NnVn_Give_VnNn(VnNn)
        FiNas = Get_AllFileNames_Like_OfOneFolder(PathOfActualFolder, NnVn + "*.lnk")
        Check_NnVnLinks_InsideActualFolder PathOfActualFolder, PathOfPersonfolder, PathOfIcon, FiNas
        'Zum aktuellen BracName existiert jetzt genau 1 Vbs-Link im ActualFolder
    'Copy VnNn-Files from ActualFolder to PersonalFolder
        Arr1 = Split(FiNaId, v)
        For i = 0 To UBound(Arr1)
            OneFileName = Arr1(i) '
            If OneFileName Like "*(" + VnNn + ")*.jpg" Then
                p1 = PathOfActualFolder + "\" + OneFileName
                p2 = PathOfPersonfolder + "\" + OneFileName
                Id = Get_PidOrVid_FromFileName(OneFileName)
                Id00 = Left(Id, 5) + "-00"
                'copy jpg
                    'Ggf. existiert ein jpg mit gleicher Id, ungleich OneFileName
                    c = Reduce_JpgsInPersonfolder(PathOfPersonfolder, OneFileName, Id)
                    If c = 0 Then 'c = 0 oder 1
                        'Es existiert kein Id-jpg im PersonFolder
                        OpenFolder PathOfPersonfolder: Stop
                        CopyFile p1, p2 '5: UF3.c2 = UF3.c2 + 1
                        'Stop 'hit F5 in Personf5older to see new jpg
                    End If
                    If c = 1 Then
                        'Es existiert genau 1 Id-jpg im PersonFolder
                        FiNa = Get_FirstFileName_Like_OfOneFolder(PathOfPersonfolder, "*" + Id + "*")
                        If FiNa <> OneFileName Then
                            Stop
                            DeleteFile PathOfPersonfolder + "\" + FiNa
                            Stop
                            CopyFile p1, p2: UF3.C2 = UF3.C2 + 1
                            Stop
                        End If
                    End If
                    
                        
                'copy p-00 jpg to Personfolder
                    If Left(Id, 1) = "p" Then
                        Fi00 = Get_FirstFileName_Like_OfOneFolder(PathOfActualFolder, "*" + Id00 + "*")
                        p1 = PathOfActualFolder + "\" + Fi00
                        p2 = PathOfPersonfolder + "\" + Fi00
                        FiNa = Get_FirstFileName_Like_OfOneFolder(PathOfPersonfolder, "*" + Id00 + "*")
                        If FiNa <> Fi00 Then
                            'Stop
                            DeleteFile PathOfPersonfolder + "\" + FiNa
                            'Stop
                        End If
                        If Not FileExists(p2) Then
                            OpenFolder PathOfPersonfolder
                            'Stop
                            CopyFile p1, p2
                            'UF3.c2 = UF3.c2 + 1
                            Stop
                            CloseFolder PathOfPersonfolder
                        End If
                    End If
                    
                    
                'Create v-00 Link inside Personfolder if necessary
                    If Left(Id, 1) = "v" Then
                        Check_Link_FromPersonfolder_ToV00_OfOneBracNameInActualFolder PathOfActualFolder, PathOfPersonfolder, Id00
                    End If
            End If
        Next
End Sub

Function Check_Link_FromPersonfolder_ToV00_OfOneBracNameInActualFolder(PathOfActualFolder$, PathOfPersonfolder$, Id00$)
    'Called from     Check_OneBracVnNn
   
   'pVideo (ohne ".lnk")
   
    'Vorbereitung
        Dim LinkNameIST$, LinkNameSOLL$, NameOfVideo$, p1$, p2$, PathOfVideoIcon$, c%
        With UF3
    'Check v00-Link in Personfolder  '19650000 LK ... v0118-00 ww.mp4.lnk
        PathOfVideoIcon = ArrC(4) + "\zzico\zzy_Video.ico"
        NameOfVideo = Get_FirstFileName_Like_OfOneFolder(PathOfActualFolder, "*" + Id00 + "*")
        LinkNameSOLL = NameOfVideo + ".lnk"
        p1 = PathOfPersonfolder + "\" + LinkNameSOLL    'PathOfLink
        p2 = PathOfActualFolder + "\" + NameOfVideo     'PathOfVideo
        c = Reduce_v00LinksInPersonfolder(PathOfPersonfolder, LinkNameSOLL, Id00)
        If c = 0 Then 'c = 0 oder 1
            'Es existiert noch kein v00-Link im PersonFolder
                'OpenFolder PathOfPersonfolder: Stop
            .B2 = .B2 + " [Add Video-Link]": .C2 = .C2 + 1
            Create_OneVbsLink PathOfPersonfolder, NameOfVideo, p2, PathOfVideoIcon
                'Stop 'hit F5 in Personfolder to see new VideoIcon
        End If
        If c = 1 Then
            'Es existiert genau 1 Link im PersonFolder
            LinkNameIST = Get_FirstFileName_Like_OfOneFolder(PathOfPersonfolder, "*" + Id00 + "*.lnk")
            If LinkNameIST <> LinkNameSOLL Then
                DeleteFile ArrC(4) + "\" + LinkNameIST
                Create_OneVbsLink PathOfPersonfolder, NameOfVideo, p2, PathOfVideoIcon
            Else
                'Der Link im PersonFolder könnte ein korrekter VbsLink sein
                If Not IsVbsLink(p1) Then
                    .B2 = .B2 + " [Add Video-Link]": .C2 = .C2 + 1
                    Create_OneVbsLink PathOfPersonfolder, NameOfVideo, p2, PathOfVideoIcon
                End If
            End If
        End If
        End With
End Function

Function Reduce_v00LinksInPersonfolder(PathOfPersonfolder$, LinkNameSOLL$, Id00$)
    Dim B$, c%, i%
    B = Get_AllFileNames_Like_OfOneFolder(PathOfPersonfolder, "*" + Id00 + "*")
    c = anzAinB(Id00, B)
    If c = 1 Then
        If B <> LinkNameSOLL Then
            'OpenFolder PathOfPersonfolder
            'Stop
            DeleteFile PathOfPersonfolder + "\" + B: c = 0
            'Stop
        End If
    ElseIf c > 1 Then
        Dim Arr1() As String: Arr1 = Split(B, vbCrLf)
        For i = 0 To UBound(Arr1)
            If Arr1(i) <> LinkNameSOLL Then
                Stop
                DeleteFile PathOfPersonfolder + "\" + Arr1(i): c = c - 1
                Stop
            End If
        Next
    End If
    Reduce_v00LinksInPersonfolder = c
End Function

Function Reduce_JpgsInPersonfolder(PathOfPersonfolder$, OneFileName$, Id$)
    Dim B$, c%, i%
    B = Get_AllFileNames_Like_OfOneFolder(PathOfPersonfolder, "*" + Id + "*")
    c = anzAinB(Id, B)
    If c = 1 Then
        If B <> OneFileName Then
            'OpenFolder PathOfPersonfolder
            'Stop
            DeleteFile PathOfPersonfolder + "\" + B: c = 0
            'Stop
        End If
    ElseIf c > 1 Then
        Dim Arr1() As String: Arr1 = Split(B, vbCrLf)
        For i = 0 To UBound(Arr1)
            If Arr1(i) <> OneFileName Then
                DeleteFile PathOfPersonfolder + "\" + Arr1(i): c = c - 1
            End If
        Next
    End If
    Reduce_JpgsInPersonfolder = c
End Function

Function Get_VnNns_OfSomeBracFileNames_WithT5VnNn(FileNames$)
   'Called from Check_Bracs_OfOneFolder
   
   'Vorbereitung
        Dim s$, VnNn$, i%, j%, A() As String, B() As String
   'Action
        A = Split(FileNames, vbCrLf): s = "|"
        For i = 0 To UBound(A)                              'A(0) = "1987*DM*(Vn Nn)*.jpg"
            B = Split(A(i), "(")
            For j = 1 To UBound(B)
                VnNn = Left(B(j), InStr(1, B(j), ")") - 1)  'VnNn = "Lea Hipp"
                If Not s Like "*|" + VnNn + "|*" Then s = s + VnNn + "|"
            Next                                            's    = "|Jo Hipp|Lea Hopp|...|"
        Next
    'Finals
        Get_VnNns_OfSomeBracFileNames_WithT5VnNn = s
End Function

Sub AllAbout_Brac()
    'Brac           = VnNn with   Brackets      = "(Lea Boll)"; inside a FileName
    'BracName       = VnNn inside Brackets      = "Lea Boll"
    'BracFile       = FileName containing Brac  = "19870605 DM (Lea Boll)...jpg"
    'BracFolder     = PersonFolder of BracName  = "...\Leute\Boll, Lea (TV Bonn)"
    'BracLink       = Link from EventFolder to BracFolder
    'Scope          CheckBrackLinks in Events, ClubsNations
    '                         nicht in Leute, Faces, Register
    'Procs          FolderHasBracNoResult
    '               Get_AllPathsOfBracFiles_OfAllFolders
    '               Get_AllPathsOfBracFiles_insideSourceFolderAndSubFolders
    '               +Get_AllPathsOfBracFiles_insideOneFolder
    '               -Load_ArrPathsOfFoldersWithBracFilesButNoResultFiles
    '               Copy_BracFiles_OfOneFolder_ToPersonFolders
    '               T4_Create_AllBracLinks_OfOneFolder
    '               T4_Create_MissingBracLinks_OfOneEvent
    '               T4_Get_Id_Folders_OfOneNameInBracketsJpg
    '               Get_VnNn_OfBracFiles_InsideOneFolder
    '               -T4_ManageBracs_EventsAndClubsNations_NoCopResults
    
    'FolderHasBracNoResult 'True if HasBrac but NoResulzJpgs
    'Load_ArrPathsOfFoldersWithBracFilesButNoResultFiles
    'Load_ArrPathsOfFoldersWithBracFiles
    'Copy_BracFiles_OfOneFolder_ToPersonFolders
        'Called from    ClickButton_CreateJpg_InActualDg
        '               T4_ManageBracs_EventsAndClubsNations_NoCopResults
        'Status         CompetitorLinks to their PersonFolders exist
        'Brac           = VnNn in Brackets = "(Lea Boll)"; inside a FileName
        'BracFile       = FileName with VnNn inside Brackets = "19870605 DM (Lea Boll)...jpg"
        'g              = |FileName####-##|Id|Vn Nn|Folder|Folder|...|Folder|   [x Zeilen]
        '               = "|19800913 DM - b02 (Lea Mai) (Kai Do) p1234-56 nn.jpg|p1234-56|Mai, Lea (Bonn)|Do, Kai (TV Abc)|"
        'o              = FileName####-00                                       [x Zeilen]
        '               = "19800913 DM - b02 (Lea Mai) (A. Bcd) (Kai Do) p1234-00 nn.jpg"
        'Action         1) Get g; Get o
        '               2) Copy JPGs ####-## and his ####-00 to PersonFolders;
        '                  falls jpg mit gleicher Id bereits im PersonFolder liegt, wird jpg ersetzt
        '               3) Create MissingBracLinks InEventFolder to Personfolder
        '               4) Create Link InPersonfolderOfNameInBracket to ActualEventFolder (if necessary)
        '               5) Create DateSquare InPersonfolderOfNameInBracket (if necessary)
    'Create_AllBracLinks_InsideOneFolder
    'T4_Create_MissingBracLinks_OfOneEvent
    'T4_Get_Id_Folders_OfOneNameInBracketsJpg
    'Get_VnNn_OfBracFiles_InsideOneFolder
    'T4_ManageBracs_EventsAndClubsNations_NoCopResults
        'Load_ArrPathsOfFoldersWithBracFilesButNoResultFiles
            'FolderHasBracNoResult
    'ClickButton_CreateJpg_InActualDg '_Create_Jpg_EventLink_InDgPersonFolders_
        'Called from    CmdCreate_Click [UserClick on 'Create jpg'-Button inside a Design]
        'Action         (1) Erzeugt ein #result.jpg des aktuellen Dg im DgEventOrdner;
        '               (2) kopiert #result jpg     in alle PersonFolder der DgCompetitors
        '               (3) kopiert °Files          in alle PersonFolder der DgCompetitors
        '               (4) kopiert BracFiles       in ihre PersonFolder
        '               (5) legt BracLinks          in den EventFolder
        '               (6) legt ein EventLink      in alle PersonFolder der DgCompetitors
        '               (7) legt ein DateSquare     in alle PersonFolder der DgCompetitors
End Function

Sub Create_AllBracLinks_InsideOneFolder(PathOfFolder$)
    'Called from    T4_Create_AllBracLinks_OfSomeFolders
    'BracFolder     = PersonFolder of BracName  = "...\Leute\Boll, Lea (TV Bonn)"
    'BracLink       = BracName-Link from EventFolder to BracFolder
    'Action         (1) Alle VnNn der BracFiles ermitteln
    '               (2) Alle Links des aktuellen Ordners einlesen
    '               (3) Existenz des BracName   in T5      sicherstellen
    '               (4) Existenz des BracFolder in 'Leute' sicherstellen
    '               (5) Pro BracName BracLink erstellen, wenn nötig; Varianten löschen
    
    
    'Link mit icon relativ
    
    
    'Vorbereitung
        Dim BracName$, JmpToFo$, LnkName$, NnVn$, s$, j%
        Dim B() As String, ArrNnVn() As String
   '(1) Alle VnNn der BracFiles ermitteln
        s = Get_VnNn_OfBracFiles_InsideOneFolder(PathOfFolder) '"|Jo Hipp|Lea Hopp|...|"
        B = Split(s, "|") 'B(1) = 1. Name
        'show s
   '(2) Alle Links des aktuellen Ordners einlesen; "|Aa, Lea (vh Bb, SU)|...|Zz, Jo (TV Xx)|"
            Load_ListOfLinkNames_OfOneFolder PathOfFolder, ArrNnVn
        For j = 1 To UBound(B) - 1
            'Schleife über alle VnNn der BracFiles des aktuellen Ordners
            BracName = B(j)         '"Jo Hipp"
            JmpToFo = T5_Get_PathOfPersonfolder_Give_VnNn(BracName) '...\Leute\Hipp, Jo (vh ...)
            UF3.B2 = "Create BracLink inside folder '" _
                     + getNameOfPath(PathOfFolder) + "': " + BracName
            If Not FolderExists(JmpToFo) Then
                'BracName not in T5, BracFolder does not exist
   '(3)        'Existenz des BracName in T5 sicherstellen
                UF3.B2 = "Create BracLink inside folder '" + getNameOfPath(PathOfFolder) + "': " + BracName
                
                Call Beep: OpenFolder PathOfFolder: Stop
                
                T4_Add_NewLineInT5_GiveVnNn BracName
                JmpToFo = T5_Get_PathOfPersonfolder_Give_VnNn(BracName)
                
                Stop
                
   '(4)        'Existenz des BracFolder in 'Leute' sicherstellen
                CreateFolder JmpToFo
            End If
   '(5)     Pro BracName BracLink erstellen, wenn nötig; Varianten löschen
            Verify_OnePersonLinkOnly PathOfFolder, JmpToFo, ArrNnVn
            DoEvents
        Next
End Sub

Sub Load_ArrPathsOfFoldersWithBracFilesButNoResultFiles(ByRef Arr() As String)
    'Called from    T4_ManageBracs_EventsAndClubsNations_NoCopResults
    
    'Vorbereitung
        Dim PathOfFolder$, t1$, T2$, c%, i%, j%
        Dim BF() As String, CN() As String, Ev() As String
        With UF3: ReDim BF(1 To 222) 'EV71+CN15
   '(1) Load all Paths of folders in Events and ClubsNations
            Load_ArrPathsOfAllFolders_Events Ev
            Load_ArrPathsOfAllFolders_ClubsNations CN
   '(2) BF() = PathsOfFolders with BracFiles but without #ResultFiles
        For i = 1 To UBound(Ev)
            PathOfFolder = Ev(i)
            'UF3CountDown 1, UBound(EV), i, 9
            UF3CountDown 1, UBound(Ev) + UBound(CN), i, 9
            If FolderHasBracNoResult(PathOfFolder) Then c = c + 1: BF(c) = PathOfFolder
        Next
        For j = 1 To UBound(CN)
            PathOfFolder = CN(j): UF3CountDown 1, UBound(Ev) + UBound(CN), i + j, 9
            If FolderHasBracNoResult(PathOfFolder) Then c = c + 1: BF(c) = PathOfFolder
        Next
            .C1.Visible = False: .D1.Visible = True
   '(3) Arr() = Reduce BF from 222 down to c (no empty entries)
        ReDim Arr(1 To c)
        For i = 1 To c
            Arr(i) = BF(i)
        Next
    'Finals
        End With
End Sub

Function Get_AllNamesOfBracFiles_insideOneFolder(PathOfFolder$) As String
    'Called from    xxx
    
    Dim s$
    s = Get_FileNames_Like_NotLike_insideOneFolder(PathOfFolder, "* (* *)*[pv]####-##*", "*(*,*)*")
    s = Delete_EmptyRowsInString(s)
    'show CStr(anzAinB(":", s)) + vbCrLf + s
    Get_AllNamesOfBracFiles_insideOneFolder = s
End Function

Function Get_AllPathsOfBracFiles_insideSourceFolderAndSubFolders(PathOfFolder$) As String
    'Called from    xxx
    
    Dim s$
    s = Get_FilePaths_Like_NotLike_insideSourceFolderAndSubFolders(PathOfFolder, "* (* *)*[pv]####-##*", "*(*,*)*")
    s = Delete_EmptyRowsInString(s)
    'show CStr(anzAinB(":", s)) + vbCrLf + s
    Get_AllPathsOfBracFiles_insideSourceFolderAndSubFolders = s
End Function

Function Get_AllPathsOfBracFiles_OfAllFolders() As String
    'Called from    [once]
    
    'Vorbereitung
        Dim N$, p$, s$, T$, v$, VnNn$, VNs$, i%, j%, Arr1() As String, Arr2() As String
        v = vbCrLf: DoArrc: Ticks1
    'FilePaths
        s = Get_FilePaths_Like_NotLike_insideSourceFolderAndSubFolders(ArrC(3), "* (* *)*[pv]####-##*", "*(*,*)*")
        s = s + vbCrLf + Get_FilePaths_Like_NotLike_insideSourceFolderAndSubFolders(ArrC(11), "* (* *)*[pv]####-##*", "*(*,*)*")
        s = Delete_EmptyRowsInString(s)
    'All VnNn in T5
        VNs = T5_Get_AllVnNn '|Aaron May|...|Zuzu June|
    'Check Name in T5
        Arr1 = Split(s, v)
        For i = 0 To UBound(Arr1)
            p = Arr1(i)         'one FilePath with Brac
            N = getNameOfPath(p)
            Arr2 = Split(N, " (")
            For j = 1 To UBound(Arr2)
                VnNn = Left(Arr2(j), InStr(1, Arr2(j), ")") - 1)
                If VNs Like "*|" + VnNn + "|*" Then
                    'VnNn ist in T5 enthalten
                    
                Else
                    'VnNn ist nicht in T5 enthalten
                    T = T + p + v
                End If
            Next
        Next
    't-Paths
        If T <> "" Then
            T = "Get_AllPathsOfBracFiles_OfAllFolders (" + T6_GetDuration + ")" + v + v + "BracName ist nicht in T5; " _
              + "bitte dort neu anlegen oder Schreibweise ändern (über EVERYTHING suchen)" + v + v + T
            show T: Stop
        End If
        
        show CStr(anzAinB(":", s)) + vbCrLf + s
        Get_AllPathsOfBracFiles_OfAllFolders = s
End Function

Function Get_AllPathsOfBracFiles_OfAllFolders2() As String
    'Called from    [once]
    
    'Vorbereitung
        Dim N$, p$, s$, T$, v$, VnNn$, VNs$, i%, j%, Arr1() As String, Arr2() As String
        v = vbCrLf: DoArrc: Ticks1
    'FilePaths
        s = Get_FilePaths_Like_insideSourceFolderAndSubFolders(ArrC(3), "* (* *)*[pv]####-##*")
        s = s + vbCrLf + Get_FilePaths_Like_insideSourceFolderAndSubFolders(ArrC(11), "* (* *)*[pv]####-##*")
        s = Delete_EmptyRowsInString(s)
    'All VnNn in T5
        VNs = T5_Get_AllVnNn '|Aaron May|...|Zuzu June|
    'Check Name in T5
        Arr1 = Split(s, v)
        For i = 0 To UBound(Arr1)
            p = Arr1(i)         'one FilePath with Brac
            N = getNameOfPath(p)
            Arr2 = Split(N, " (")
            For j = 1 To UBound(Arr2)
                VnNn = Left(Arr2(j), InStr(1, Arr2(j), ")") - 1)
                If VNs Like "*|" + VnNn + "|*" Then
                    'VnNn ist in T5 enthalten
                    
                Else
                    'VnNn ist nicht in T5 enthalten
                    T = T + p + v
                End If
            Next
        Next
    't-Paths
        If T <> "" Then
            T = "Get_AllPathsOfBracFiles_OfAllFolders (" + T6_GetDuration + ")" + v + v + "BracName ist nicht in T5; " _
              + "bitte dort neu anlegen oder Schreibweise ändern (über EVERYTHING suchen)" + v + v + T
            show T: Stop
        End If
        
        show CStr(anzAinB(":", s)) + vbCrLf + s
        Get_AllPathsOfBracFiles_OfAllFolders2 = s
End Function

Function FolderHasBracNoResult(PathOfFolder$) As Boolean
    'Called from    Load_ArrPathsOfFoldersWithBracFilesButNoResultFiles

    'Vorbereitung
        Dim fso As Object, Fold As Object, Fils As Object
        Dim Fil As Object, N$
    'FSO
        Set fso = CreateObject("scripting.FileSystemObject")
        Set Fold = fso.GetFolder(PathOfFolder)
        Set Fils = Fold.Files
    'All FilePaths of one folder
        For Each Fil In Fils
             If Not Fil Is Nothing Then
                N = CStr(Fil.NAME)
                If Not N Like "*#result*" Then
                   If N Like "[12][90]*([A-ZÄÖÜ]* [A-ZÄÖÜ]*)*" Then
                        FolderHasBracNoResult = True: Exit Function
                   End If
                End If
             End If
        Next Fil
End Function

Sub Copy_BracFiles_OfOneFolder_ToPersonFolders_TEST()
    Dim p$: p = "F:\Archiv TR\Archiv Trampolin\Events\1961 LK D-GB Kiel"
    Copy_BracFiles_OfOneFolder_ToPersonFolders p
    '19670617 WM04 London_GB - a62 (David Jacobs) v0017-00 ww.mp4
End Sub

Sub Copy_BracFiles_OfOneFolder_ToPersonFolders(PathOfFolder$)
    'Called from    ClickButton_CreateJpg_InActualDg
    '               T4_ManageBracs_EventsAndClubsNations_NoCopResults
    'Status         CompetitorLinks to their PersonFolders exist
    'Brac           = VnNn in Brackets = "(Lea Boll)"; inside a FileName
    'BracFile       = FileName with VnNn inside Brackets = "19870605 DM (Lea Boll)...jpg"
    'g              = |FileName####-##|Id|Vn Nn|Folder|Folder|...|Folder|   [x Zeilen]
    '               = "|19800913 DM - b02 (Lea Mai) (Kai Do) p1234-56 nn.jpg|p1234-56|Mai, Lea (Bonn)|Do, Kai (TV Abc)|"
    'o              = FileName####-00                                       [x Zeilen]
    '               = "19800913 DM - b02 (Lea Mai) (A. Bcd) (Kai Do) p1234-00 nn.jpg"
    'Action         1) Get g; Get o
    '               2) Copy JPGs ####-## and his ####-00 to PersonFolders;
    '                  falls jpg mit gleicher Id bereits im PersonFolder liegt, wird jpg ersetzt
    '               3) Create MissingBracLinks InEventFolder to Personfolder
    '               4) Create Link InPersonfolderOfNameInBracket to ActualEventFolder (if necessary)
    '               5) Create DateSquare InPersonfolderOfNameInBracket (if necessary)
    
    'Vorbereitung
        Dim FiNa$, FiNa00$, FN$, G$, Id$, Id00$, IdFoFo$
        Dim o$, PathOfPersonfolder$, p1$, p2$, p3$, v$
        Dim c%, i%, j%, zT4%, A() As String, B() As String, A1() As String
        v = vbCrLf: DoArrc
    '1) Get g; Get o
        
        FN = Get_AllFileNames_Like_OfOneFolder(PathOfFolder, "*####-##*")
        'show FN: Stop
        A = Split(FN, v)
        For i = 0 To UBound(A)
            'Schleife über alle FileNamesMitId des aktuellen SourceFolder (Fotos+Videos)
            FiNa = A(i): IdFoFo = ""                                'FiNa = one FileName ####-##
            If FiNa Like "*####-00*" Then o = o + FiNa + v          'o = FileNames ####-00
            IdFoFo = T4_Get_Id_Folders_OfOneNameInBracketsJpg(FiNa)
            '      = "|p1234-01|May, Pia (TV Bonn)|Mau, Lea (Bern)|..."
            If IdFoFo <> "" Then G = G + "|" + FiNa + IdFoFo + v    'g = FileNames ####-##
        Next
        G = Delete_EndReturnsInString(G)
        'show g         'FileNames    '19..(..) (..)..Id##..jpg  inside SourceFolder
        'show o: Stop   'FileNames    '19.............Id00..jpg  inside SourceFolder
        
        
    '2) Copy JPGs ####-## and his ####-00 to PersonFolders
        A = Split(G, v)
        B = Split(o, v)
        For i = 0 To UBound(A)
            A1 = Split(A(i), "|"): c = 3    '    | 2|   3  |
            '  A1(i) = oneLine = |FileName####-##|Id|Folder|...|Folder|
            Id = A1(2)
            Do While c < UBound(A1)
                'Schleife über alle NameInBracket
                PathOfPersonfolder = ArrC(4) + "\" + A1(c)
                p1 = PathOfFolder + "\" + A1(1)         'SourceFolder + actualFileName
                p2 = PathOfPersonfolder + "\" + A1(1)   'PersonFolder + actualFileName
                p3 = PathOfPersonfolder + "\" + Get_AllFileNames_Like_OfOneFolder(PathOfPersonfolder, "*" + Id + "*")
                '  = File mit gleicher Id ####-## bereits im PersonFolder?
                
                'Copy
                    If p2 <> p3 Then DeleteFile p3
                    CopyFile p1, p2
                '-00 jpg
                    If Not Id Like "*####-00" Then
                        'FileName ####-00 in B suchen (soll auch in den PersonFolder kopiert werden)
                        Id00 = Left(Id, 5) + "-00"
                        For j = 1 To UBound(B)
                            If B(j) Like "*" + Id00 + "*" Then
                                'Folder-File ####-00 gefunden
                                FiNa00 = B(j): j = UBound(B) '= Exit For j
                            End If
                        Next
                        p1 = PathOfFolder + "\" + FiNa00  'File ####-00 in Folder
                        p2 = PathOfPersonfolder + "\" + FiNa00 'File ####-00 in PersonFolder
                        p3 = PathOfPersonfolder + "\" + Get_AllFileNames_Like_OfOneFolder(PathOfPersonfolder, "*" + Id00 + "*")
                        '  = File mit gleicher Id ####-00 ggf. bereits im PersonFolder?
                        'Copy
                            If p3 <> "" Then DeleteFile p3
                            CopyFile p1, p2
                        End If
                c = c + 1
            Loop
        Next
    'Finals
        Beep
End Sub

Sub Check_NnVnLinks_InsideActualFolder(PathOfActualFolder$, PathOfPersonfolder$, PathOfIcon$, FiN$)
    'Called from    Check_OneBracVnNn
    'FiN            = AllFileNames of one NnVn  'May, Lea   'May,Lea (TV Bonn)   '...
    'Action         Stellt die Existenz von BracLinks der Bracfiles eines Ordners sicher;
    '               erzeugt ggf. korrekte VbsLinks, löscht unkorrekte
    
    'Vorbereitung
        Dim FolderToHoldLink$, FolderToJumpTo$, LinkNameToDisplay$
        Dim NameOfLink$, NameOfPersonfolder$, PathOfLink$
        Dim AnzLnk%, i%, Arr1() As String
        NameOfPersonfolder = getNameOfPath(PathOfPersonfolder)
        FolderToHoldLink = PathOfActualFolder
        FolderToJumpTo = PathOfPersonfolder
        LinkNameToDisplay = NameOfPersonfolder
    'Action
        AnzLnk = anzAinB(".lnk", FiN)
        If AnzLnk = 0 Then
            'Es existiert noch kein NnVn-Link in actual Folder
            Create_OneVbsLink FolderToHoldLink, LinkNameToDisplay, FolderToJumpTo, PathOfIcon
        ElseIf AnzLnk = 1 Then
            'Es existiert bereits ein NnVn-Link
            NameOfLink = FiN
            PathOfLink = PathOfActualFolder + "\" + FiN
            If Not (NameOfLink = NameOfPersonfolder + ".lnk" And IsVbsLink(PathOfLink)) Then
                'Link löschen, neuen VbsLink erzeugen
                DeleteFile PathOfLink
                Create_OneVbsLink FolderToHoldLink, LinkNameToDisplay, FolderToJumpTo, PathOfIcon
            End If
        Else
            'Es existieren >1 NnVn-Links 'May, Lea 'May,Lea (TV Bonn)
            Arr1 = Split(FiN, vbCrLf)
            For i = 0 To UBound(Arr1)
                NameOfLink = Arr1(i)
                PathOfLink = PathOfActualFolder + "\" + Arr1(i)
                If Not (NameOfLink = NameOfPersonfolder + ".lnk" And IsVbsLink(PathOfLink)) Then
                    'Link löschen, neuen VbsLink erzeugen
                    DeleteFile PathOfLink
                    Create_OneVbsLink FolderToHoldLink, LinkNameToDisplay, FolderToJumpTo, PathOfIcon
                End If
            Next
        End If
End Sub



