Attribute VB_Name = "M_Vbs"
Option Explicit
Private ArrVid() As String

Sub zzz_M_VBS()
    showProcs "report"
    
    'show Get_TextOfAllOpenNotepadWindows
    EE 1: Beep
End Sub

'ToUse -------------------------------------------------------------

'Sub CopyToFolder_WheatSquareJpg_Leute         (PathOfFolder$)
'Sub Create_AllPersonLinks_Register            ()
'Sub Create_OneVbsLink_Person                  (pLinkWithPersName$)
'Sub Create_VbsLink_all_Club_to_D              ()
'Sub Create_VbsLink_all_Club_to_LTV            ()
'Sub Create_VbsLink_all_LTV_to_D               ()
'Sub Create_WheatSquareJpgAndIcon_Nation       (NationShorty$)
'Sub Fill_L1_L2_L3_IfLinkIsPersLink            (PathOfLink$, L1_pLink$, L2_pJumpTo$, L3_pIcon$)
'Sub Load_List_ClubLtvD_FromT5                 (ByRef Arr   () As String)
'Sub Load_List_ClubLtvNation_FromT5            (ByRef Arr() As String)
'Sub Load_List_LTV_FromT5                      (ByRef Arr() As String)
'Sub Move_FolderWindow_XYWH                    (PathOfFolder, PosX, PosY, Breite, Hoehe)
'Sub Repair_VbsLinks_OneFolder                 (PathOfFolder$)
'Sub Repair_VbsLinks_AllFolders                ()
'Sub StartVBS_Delete_OlderVersionsOfLink       (PathOfFolder$)
'Sub Separate_SomeLinkNames_Leute              (PathOfFolder$)
'Sub Show_BadLinks_GiveSourceFolder            ()
'Sub Show_NonVbsLinks_in_Events                ()
'Sub Show_NonVbsLinks_in_Leute                 ()
'Sub Show_NonVbsLinks_in_ClubsNations          ()
'Sub Show_NonVbsLinks_in_Archiv                ()
'Sub T4_ActivateEachPersNameOfEachDg           ()

'Function CountFolders                         (ByVal PathOfRootFolder$) As Long
'Function Get_InfoOfOneLink_inOneLine          (PathOfLink$) As String
'Function Get_NonVbsLinks_insideFolder         (PathOfRootFolder$) As String
'Function HasVbsLinkDescription                (PathOfLink$) As Boolean

'ToDo -------------------------------------------------------------
'EventLinks (in PersFolders) wo werden die erstellt?
'Sub Create_VbsLink_all_Nation_to_PersFold     ()
'Sub Create_VbsLink_all_Club_to_PersFold       ()
'Sub Create_VbsLink_all_Pers_to_ClubNation     ()
'-----------------------------------------------------------------------

Function T2_GetPathsA(SourceFolder$) As String
    'Called from    xxx
    'Status         No global variables, no Collection
    'Action         sammelt OrdnerPfade, all levels
    
   'Vorbereitung
        Dim F$, s$, c&, FileSystem As Object
        F = SourceFolder
        If Right(F, 1) <> "/" Then F = F + "/"
    'FileSystemObject
        Set FileSystem = CreateObject("Scripting.FileSystemObject")
    'Alle Pfade in s schreiben
        T2_GetPathsB FileSystem.GetFolder(F), s, c
    'Finals
        If Right(s, 2) = vbCrLf Then s = Left(s, Len(s) - 2)
        Get_Paths_OfAllSubfolders_AllLevels = s
End Function

Sub T2_GetPathsB(folder, s$, c&)
    'Called from    T2_GetPathsA
    'Status         recursiv
    Dim subFolder As Object
    For Each subFolder In folder.subfolders
        s = s + CStr(subFolder) + vbCrLf
        c = c + 1: Sheets("T2").Cells() = c
        T2_GetPathsB subFolder, s
    Next
End Sub

'Sub x()
'    Dim p$: p = "F:\Archiv Trampolin prog\prog\old - tmpShow-Files\4.txt"
'    PepUp_BadLinksReport ReadFile(p)
'End Sub

Sub PepUp_BadLinksReport(s$)
    Dim L$, T$, v$, i&, A() As String, B() As String
    v = vbCrLf
    s = Delete_EmptyRowsInString(s)
    A = Split(s, v): QuickSort A ': s = Join(A, v)
    For i = 0 To UBound(A)
        L = A(i) 'one Line
        B = Split(Mid(L, 9), "|")
        T = T + Left(L, 7) + "| |" + "1 --> " + Mid(L, 2, 1) + "|" + B(1) + v _
                    + "         |2 --> " + Mid(L, 3, 1) + "|" + B(2) + v _
                    + "         |3 --> " + Mid(L, 4, 1) + "|" + B(3) + v _
                    + "         |4 --> " + Mid(L, 5, 1) + "|" + B(4) + v _
                    + "         |5 --> " + Mid(L, 6, 1) + "|" + B(5) + v _
                    + "         |6 --> " + Mid(L, 7, 1) + "|" + B(6) + v + v
    Next
    show "BadLinksReport: " + CStr(UBound(A) + 1) + vbCrLf + vbCrLf + T
End Sub

Sub Repair_VbsLinks_AllFolders()
    'Vorbereitung
        Dim Report$, v$, i&, A() As String, B() As String
        With Sheets("W1"): .Activate: DoArrc: v = vbCrLf
    'Load all paths of Folders
        'Load_ArrPathsOfAllFolders_Archiv A
        Load_ArrPathsOfAllFolders_Leute A
        'Load 1
            'ReDim A(0 To 0): A(0) = "F:\Archiv Trampolin 1900-1999\Leute\Bevan, Randall (Cyncoed, GB)"
        [E11] = "Load_ArrPathsOfAllFolders_...": [E12] = "": [E13] = ""
    'Load all paths of original Videos
        Load_AllVideos_Events_ClubsNations ArrVid
        [E11] = "Load_AllVideos_Events_ClubsNations"

    'Action
        For i = 0 To UBound(A)
            'Anzeige CountDown auf sheet W1
                [B11] = UBound(A) - i + 1:            [E11] = "Folder " + CStr(i + 1) + " of " + CStr(UBound(A) + 1)
                [E12] = Get_NameOfParentFolder(A(i)): [E13] = Get_NameOfFolder(A(i))
                DoEvents
            'Repair
                Repair_VbsLinks_OneFolder A(i), Report
        Next
    'Finals
        show "Repair_VbsLinks_AllFolders" + v + v + Report
        [B11] = 0: Beep: End With
End Sub

Function Get_LinkTypNr_1to9(PathOfOneLink$) As Integer
    'Called from    Change_AllLinksOfOneFolder_ToVbsLinks
    'LinkName       = "Froehlich, Ron (SA).lnk"
    'Typ-Ziffer     = 1 Person, 2 Event, 3 Club, 4 LTV, 5 Nation 6 Video 9 Sonst
    'Action         liefert Typ-Ziffer für PathOfOneLink
    
    'Vorbereitung
        Dim N$, M1$, M2$, vID$
        vID = "|.avi|.divx|.mkv|.mp4|.mpg|.mov|.vob|.wmv|"
        N = NameOfPath(PathOfOneLink)           'one LinkName
        M1 = Replace(N, ".lnk", "")             'LinkName ohne .lnk
        M2 = LCase("*|" + Right(M1, 4) + "|*")  'LikeString "*|.mp4|*"
    'Separate to Typ
        If N Like "-   [A-Z]*" Then Get_LinkTypNr_1to9 = 5:    Exit Function
        If N Like "-  [A-Z]*" Then Get_LinkTypNr_1to9 = 4:     Exit Function
        If vID Like M2 Then Get_LinkTypNr_1to9 = 6:            Exit Function
        If N Like "*, *" Then Get_LinkTypNr_1to9 = 1:          Exit Function
        If N Like "######## ! *" Then Get_LinkTypNr_1to9 = 2:  Exit Function
        If N Like "- [A-Z]*" Then Get_LinkTypNr_1to9 = 3:      Exit Function
        Get_LinkTypNr_1to9 = 9
End Function


Sub Repair_VbsLinks_OneFolder(PathOfFolder$, Optional Report$)
    'Called from    Repair_VbsLinks_AllFolders
    'Status         No SubFolders
    'Action         Alle .lnk-Files des aktuellen Ordners werden gelesen
    '               in Link-Typen separiert, dann Typ für Typ repariert
    
    'Vorbereitung
        Dim clubs$, LinkNames$, LTV$, v$
        Dim LnkPerson$, LnkEvent$, LnkClub$, LnkLTV$, LnkNation$, LnkVideo$   'werden hier gefüllt
        v = vbCrLf
    'LinkNames
        LinkNames = Get_AllFileNames_Like_OfOneFolder(PathOfFolder, "*.lnk")
    'Separate to 5 kinds of Link       'LnkPerson, LnkEvent, ... je als "xxx|xxx|...|xxx"
        Separate_SomeLinkNames clubs, LinkNames, LnkPerson, LnkEvent, LnkClub, LnkLTV, LnkNation, LnkVideo
        
        
        show "LnkPerson: " + LnkPerson + v + "LnkEvent:  " + LnkEvent + v + "LnkClub:   " + LnkClub + v + "LnkLTV:    " + LnkLTV + v + "LnkNation: " + LnkNation
        Stop
        
        
    'LinkTyp1
        Repair_VbsLinks_OneFolder_LinkTyp1_Person PathOfFolder, LnkPerson, Report
    'LinkTyp2
        Repair_VbsLinks_OneFolder_LinkTyp2_Events PathOfFolder, LnkEvent, Report
    'LinkTyp3
        Repair_VbsLinks_OneFolder_LinkTyp3_Club PathOfFolder, LnkClub, Report
    'LinkTyp4
        Repair_VbsLinks_OneFolder_LinkTyp4_LTV PathOfFolder, LnkLTV, Report
    'LinkTyp5
        Repair_VbsLinks_OneFolder_LinkTyp5_Nation PathOfFolder, LnkNation, Report
    'LinkTyp6
        Repair_VbsLinks_OneFolder_LinkTyp6_Video PathOfFolder, LnkVideo, Report
End Sub

Sub Create_OneVbsLink_OnePerson()
    Dim p$
    p = "F:\Archiv Trampolin 1900-1999\ClubsNations\- International\Froehlich, Ron (SA).lnk"
    Create_OneVbsLink_Person p
End Sub

Sub Repair_VbsLinks_OneFolder_LinkTyp1_Person(PathOfFolder$, LnkPerson$, Report$)
    'Called from    Repair_VbsLinks_OneFolder
    'LnkPerson      = 1 oder mehrere NameOfLinkTyp1     'mit Trenner "|"
    '               = "Aaron, Syd (GB).lnk|...|Lea May (geb June, TV Bonn).lnk"
    'Status         No SubFolders
    
    'Vorbereitung
        If LnkPerson = "" Then Exit Sub
        Dim L$, PL$, i&, A() As String
        'Dim L1_pLink$, L2_pJumpTo$, L3_pIcon$                       'werden hier gefüllt
    'LinkTyp1
        If LnkPerson <> "" Then '"Aaron, Syd (GB).lnk|...|Lea May (geb June, TV Bonn).lnk"
            'Links jumping to one PersonFolder
            CopyToFolder_WheatSquareJpg_Leute PathOfFolder
            A = Split(LnkPerson, "|")
            For i = 0 To UBound(A) - 1
                L = A(i)                    'One NameOfLink 'Aaron, Syd (GB).lnk
                PL = PathOfFolder + "\" + L 'One PathOfLink
                
                
                'Fill_L1_L2_L3_IfLinkIsPersLink pL, L1_pLink, L2_pJumpTo, L3_pIcon
                If "*\Register" Like PathOfFolder Then Stop
                Create_OneVbsLink_Person PL
                
                
            Next
        End If
End Sub

Sub Repair_VbsLinks_OneFolder_LinkTyp2_Events(PathOfFolder$, LnkEvent$, Report$)
    'Called from    Repair_VbsLinks_OneFolder
    'Status         No SubFolders
    
    'Vorbereitung
        If LnkEvent = "" Then Exit Sub
        Dim E$, EventFolder$, L$, PL$, v$, y$, i&, A() As String
        Dim L1_pLink$, L2_pJumpTo$, L3_pIcon$                       'werden hier gefüllt
        v = vbCrLf
        If Right(LnkEvent, 1) = "|" Then LnkEvent = Left(LnkEvent, Len(LnkEvent) - 1)
    'Links jumping to one EventFolder
        A = Split(LnkEvent, "|")
        For i = 0 To UBound(A)
            L = A(i)                    'One NameOfLink '19670617 ! WM04 London_GB.lnk
            PL = PathOfFolder + "\" + L 'One PathOfLink
            'EventFolder festellen
                E = Replace(Replace(L, ".lnk", ""), "! ", "") '19670617 WM04 London_GB
                'Datum 19990101 --> 1999-01-01
                E = Left(E, 4) + "-" + Mid(E, 5, 2) + "-" + Mid(E, 7): E = Replace(E, "-00", "")
                EventFolder = ArrC(3) + "\" + E
        
                If Not FolderExists(EventFolder) Then
                    If L Like "*Buli*" Then
                        y = Left(L, 4)
                        EventFolder = ArrC(3) + "\" + y + " Liga\" + E
                        If Not FolderExists(EventFolder) Then
                            Report = Report + "EventLink; from: " + PL + v _
                                            + "           to ?: " + EventFolder + v
                            show Report: Stop
                        End If
                    End If
                End If
            L1_pLink = PathOfFolder + "\" + L
            L2_pJumpTo = EventFolder
            L3_pIcon = ArrC(4) + "\zzico\zzy_Event.ico"
            Create_OneVbsLink L1_pLink, L2_pJumpTo, L3_pIcon
        Next
End Sub

Sub Repair_VbsLinks_OneFolder_LinkTyp3_Club(PathOfFolder$, LnkClub$, Report$)
    'Called from    Repair_VbsLinks_OneFolder
    'Status         No SubFolders
    
    'Vorbereitung
        If LnkClub = "" Then Exit Sub
        Dim Club$, ClubFolder$, clubs$, L$, LTV$, PL$, i&, zT5&, A() As String
        Dim L1_pLink$, L2_pJumpTo$, L3_pIcon$                       'werden hier gefüllt
        With Sheets("T5")
    'Clubs
        If PathOfFolder Like "*Liga*" Then clubs = T5_Get_AllClubs '|ASV Nürnberg|...|Zoetermeer|
    'LinkTyp3
        'Links jumping to one D-ClubFolder (with icon WheatSquareClubArrowUp)
        A = Split(LnkClub, "|")
        For i = 0 To UBound(A) - 1
            L = A(i)                    'One NameOfLink 'SV Blankenloch.lnk
            PL = PathOfFolder + "\" + L 'One PathOfLink
            'ClubLinks in D als "MTV Elze", "- MTV Elze"
                Club = Replace(L, ".lnk", ""): Club = Replace(Club, "- ", "")
            'LTV festellen
                zT5 = Get_RowNr_HoldingMyTextPartInColumnX("T5", 7, 7, Club)
                If zT5 = 0 Then
                    Report = Report + "ClubLink; No zT5; " + PL + vbCrLf
                Else
                    LTV = .Cells(zT5, 8): If LTV = "" Then Stop
                End If
            'ClubFolder festellen
                ClubFolder = ArrC(11) + "\D\" + LTV + "\" + Club
                If Not FolderExists(ClubFolder) Then
                    Report = Report + "ClubLink; " + ClubFolder + vbCrLf
                End If
            L1_pLink = PathOfFolder + "\" + L
            L2_pJumpTo = ClubFolder
            L3_pIcon = ArrC(4) + "\zzico\zzy_Club.ico"
            Create_OneVbsLink L1_pLink, L2_pJumpTo, L3_pIcon
        Next
        End With
End Sub

Sub Repair_VbsLinks_OneFolder_LinkTyp4_LTV(PathOfFolder$, LnkLTV$, Report$)
    'Called from    Repair_VbsLinks_OneFolder
    'Status         No SubFolders
    
    'Vorbereitung
        If LnkLTV = "" Then Exit Sub
        Dim LTVFolder$, L$, LTV$, i&, A() As String
        Dim L1_pLink$, L2_pJumpTo$, L3_pIcon$                       'werden hier gefüllt
    'LinkTyp4
        'Links jumping to one D-LTVFolder (with icon WheatSquareClubArrowUp)
        A = Split(LnkLTV, "|")
        For i = 0 To UBound(A) - 1
            L = A(i)                    'One NameOfLink '"-  Baden"
                LTV = Replace(L, ".lnk", ""): LTV = Replace(LTV, "-  ", "")
            'LTVFolder festellen
                LTVFolder = ArrC(11) + "\D\" + LTV
                If Not FolderExists(LTVFolder) Then Report = Report + "LTVLink; " + LTVFolder + vbCrLf
            L1_pLink = PathOfFolder + "\" + L
            L2_pJumpTo = LTVFolder
            L3_pIcon = ArrC(4) + "\zzico\zzy_LTV.ico"
            Create_OneVbsLink L1_pLink, L2_pJumpTo, L3_pIcon
        Next
End Sub

Sub Repair_VbsLinks_OneFolder_LinkTyp5_Nation(PathOfFolder$, LnkNation$, Report$)
    'Called from    Repair_VbsLinks_OneFolder
    'Status         No SubFolders
    
    'Vorbereitung
        If LnkNation = "" Then Exit Sub
        Dim NationFolder$, L$, Nation$, i&, A() As String
        Dim L1_pLink$, L2_pJumpTo$, L3_pIcon$                       'werden hier gefüllt
    'LinkTyp5
        'Links jumping to one NationFolder (with icon WheatSquareNationArrowUp)
        A = Split(LnkNation, "|")
        For i = 0 To UBound(A) - 1
            L = A(i)                    'One NameOfLink '"-   GB"
                Nation = Replace(L, ".lnk", ""): Nation = Replace(Nation, "-   ", "")
            'NationFolder festellen
                NationFolder = ArrC(11) + "\" + Nation
                If Not FolderExists(NationFolder) Then Report = Report + "NationLink; " + NationFolder + vbCrLf
            L1_pLink = PathOfFolder + "\" + L
            L2_pJumpTo = NationFolder
            L3_pIcon = ArrC(4) + "\zzico\zzz_" + Nation + ".ico"
            If Not FileExists(L3_pIcon) Then Create_WheatSquareJpgAndIcon_Nation Nation
            Create_OneVbsLink L1_pLink, L2_pJumpTo, L3_pIcon
        Next
End Sub

Sub Repair_VbsLinks_OneFolder_Start()
    Dim p$, Report$: DoArrc 'ArrC(): 2 Archiv, 3 Events, 4 Leute, 5 Register, 11 ClubsNations
    'p = "F:\Archiv Trampolin 1900-1999\Leute\Budenberg, Michael (PSV Berlin)"
    p = "F:\Archiv Trampolin 1900-1999\Leute\Abele, Xxx (Schwaben)"
    Repair_VbsLinks_OneFolder p, Report$
    Beep
End Sub

Sub Repair_VbsLinks_OneFolder_LinkTyp6_Video(PathOfFolder$, LnkVideo$, Report$)
    'Called from    Repair_VbsLinks_OneFolder
    'Status         No SubFolders
    'LnkVideo       Die Typ6-Links des aktuellen Ordners werden zu den bisher gesammelten
    '               Video-Links (der vorher durchsuchten Ordner) in LnkVideo hinzugefügt
    'Beispiel       pL = "F:\..\Leute\Budenberg, Michael (PSV Berlin) _
    '                    \19641112 DM Berlin - a02 v0140-00 ww.mp4.lnk"
    '               LinkOrt:    PersonFolder "...\Leute\Budenberg, Michael (PSV Berlin)"
    '               Sprungziel: EventFolder  "...\Events\19641112 DM Berlin"
    '               Der Link pL wurde beim Untersuchen des Ordners PathOfFolder gefunden;
    '               Bei Doppelklick auf diesen Link soll das Video geöffnet werden
    
    'Vorbereitung
        If LnkVideo = "" Then Exit Sub
        Dim L$, E$, pOrig$, vID$, VideoFolder$, c&, i&, j&, A() As String
        Dim L1_pLink$, L2_pJumpTo$, L3_pIcon$                       'werden hier gefüllt
    'LinkTyp6
        'Links jumping to one Video-File (with icon WheatSquareVideo)
        A = Split(LnkVideo, "|")
        For i = 0 To UBound(A) - 1
            L = A(i) 'One NameOfLink    '"19641112 DM Berlin - a02 v0140-00 ww.mp4.lnk"
            vID = Get_Vid_FromFileName(L)
            'Pfad des OriginalVideos ermitteln 'ArrVid() enthält alle Pfade von Originalvideos
                pOrig = ""
                For j = 0 To UBound(ArrVid)
                    If ArrVid(j) Like "*" + vID + "*" Then pOrig = ArrVid(j): Exit For
                Next
                If pOrig = "" Then Stop
                
            'OpenFolder PathOfFolder: Stop
            
            'Create
                L1_pLink = PathOfFolder + "\" + L
                L2_pJumpTo = pOrig
                L3_pIcon = ArrC(4) + "\zzico\zzy_Video.ico"
                Create_OneVbsLink L1_pLink, L2_pJumpTo, L3_pIcon
                
            'Stop: CloseFolder PathOfFolder
        Next
    
End Sub











