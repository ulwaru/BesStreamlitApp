Attribute VB_Name = "M48"
Option Explicit 'M48

Sub zzz_M48()
    
    showProcs "load fold"
    
    'RenameModule "Modul1", "M41"
    EE 1: Beep
End Sub

Function Get_TargetOfShortCut(ByVal PathOfLnkFile$)
    On Error GoTo Err
    With CreateObject("Wscript.Shell").CreateShortcut(PathOfLnkFile)
        Get_TargetOfShortCut = .targetPath
    End With
    Exit Function
Err:
    show "Error in Function Get_TargetOfShortCut(ByVal PathOfLnkFile$)" + vbCrLf + vbCrLf _
       + "PathOfLnkFile        = " + PathOfLnkFile + vbCrLf _
       + "Get_TargetOfShortCut = " + Get_TargetOfShortCut
    Stop
End Function

Sub T4_ManageBracs_EventsAndClubsNations_NoCopResults()
    'Called from    xxx
    'Brac           = VnNn with   Brackets      = "(Lea Boll)"; inside a FileName
    'BracName       = VnNn inside Brackets      = "Lea Boll"
    'BracFile       = FileName containing Brac  = "19870605 DM (Lea Boll)...jpg"
    'BracFolder     = PersonFolder of BracName  = "...\Leute\Boll, Lea (TV Bonn)"
    'BracLink       = BracName-Link from SourceFolder to BracFolder
    'BF()           = 1D-Array = ArrPathsOfFolders with BracFiles but no ResultFiles
    
    'see     Sub ClickButton_CreateJpg_InActualDg()
    '        Copy_BracFiles_OfOneFolder_ToPersonFolders PathOfEventFolder      '(4)
    '        T4_Create_MissingBracLinks_OfOneEvent PathOfEventFolder                    '(5)
    '        T4_Create_EventLink_DateSquare_OfOneEventfolder_ToPersonFolders PathOfEventFolder '(6) (7)
    
    'Action         (1) Load ArrPathsOfFolders with BracFiles but no #result.jpg
    '               (2) legt BracLinks          in den  SourceFolder
    
    '               (3) kopiert BracFotos       in ihre PersonFolder
    '               (4) legt LinkOfBracVideos   in ihre PersonFolder
    '               (5) legt ein EventLink      in alle PersonFolder der DgCompetitors
    '               (6) legt ein DateSquare     in alle PersonFolder der DgCompetitors

    'Vorbereitung
        Dim ActualFolder$, i%, BF() As String
        DoArrc
    'UF3
        With UF3: UF3_Format: UF3_CountDown 3, "T4_ManageBracs_EventsAndClubsNations_NoCopResults"
        .B1 = "Search for folders with BracNames/no #result:   Events, ClubsNations"
        .B2 = "Actual Folder: "
        .B3 = "Create BracLink inside actual folder"
   '(1) Load ArrPathsOfFolders with BracFiles but no #result.jpg 'include #resultFolders?
        Load_ArrPathsOfFoldersWithBracFilesButNoResultFiles BF
   '(2) Manage each folder (of some EventFolders/ClubNationFolders)
        For i = 1 To UBound(BF)
            ActualFolder = BF(i)
            .C2 = UBound(BF) - i: .B2 = "Actual Folder:     " + NameOfPath(ActualFolder)
            .C3 = "+ 0"
            Create_AllBracLinks_InsideOneFolder ActualFolder
                
            'T4_Copy_BracFotos_OfOneFolder_ToPersonFolders ActualFolder: DoEvents
            'T4_Create_LinksOfBracVideos_OfOneFolder_ToPersonFolders ActualFolder: DoEvents
            'T4_Create_EventLink_InPersonFolders ActualFolder
            'T4_Create_DateSquare_InPersonFolders ActualFolder
        Next
    'Finals
        End With: Beep: WaitSecs 2: UF3.Hide
End Sub

Sub Verify_OnePersonLinkOnly(PathOfFolder$, JmpToFo$, ByRef ArrNnVn() As String)
    'Called from    Create_AllBracLinks_InsideOneFolder
    'ArrNnVn()      = Array1D; alle im Folder existierenden Links
    '               = |Aa, Lea (vh Bb, SU)|...|Zz, Jo (TV Xx)|
    'Action         Pro Name einen Link erstellen, wenn nötig; Varianten löschen
    
    'Vorbereitung
        Dim ExistingLinkName$, LnkName$, NnVn$, c%, i%
        'showArray ArrNnVn: Stop '|Aaa, Lea (CH)|Aaa, Lea (TV Xx, CH)|...
    'Action
        LnkName = NameOfPath(JmpToFo) 'Hipp, Jo (vh Hopp, TV ..)
        If LnkName = "" Then Stop
        c = InStr(1, LnkName, " (")
        If c = 0 Then NnVn = LnkName Else: NnVn = Left(LnkName, c - 1) 'Hipp, Jo
    'Link löschen, falls nicht exakt LnkName
        For i = 1 To UBound(ArrNnVn)
            ExistingLinkName = ArrNnVn(i)
            If ExistingLinkName Like NnVn + "*" Then
                'Ein Link, beginnend mit NnVn, existiert in Folder
                'könnte SOLL-LnkName sein   [Nn, Vn (Club, Nation)]
                'oder Variante              [Nn, Vn (Nation)]
                If ExistingLinkName <> LnkName Then
                    'Call beep: OpenFolder PathOfFolder: Stop
                    DeleteFile PathOfFolder + "\" + ExistingLinkName + ".lnk"
                    'OpenFolder PathOfFolder: Stop
                End If
            End If
        Next
    'Falls SollLink bereits existiert: DoNothing; sonst: Create
        If Not FileExists(PathOfFolder + "\" + LnkName + ".lnk") Then
            'BracLink mit Namen LnkName existiert noch nicht
            UF3.B2 = "Create BracLink inside folder '" _
                     + getNameOfPath(PathOfFolder) + "': " + LnkName
            Call Beep: OpenFolder PathOfFolder: Stop
            Create_OneLinkFile PathOfFolder, JmpToFo, LnkName 'CrInFo, JmpToFo, LnkName
        End If
End Sub

Function Load_ListOfLinkNames_OfOneFolder(PathOfFolder$, ByRef ArrNnVn() As String)
    'Called from    Create_AllBracLinks_InsideOneFolder
    'ArrNnVn()      = 1DArray; wird hier gefüllt
    '               = |Aa, Lea (vh Bb, SU)|...|Zz, Jo (TV Xx)
    
    'Vorbereitung
        Dim s$
    'Action
        s = Get_AllFileNames_Like_OfOneFolder(PathOfFolder, "*.lnk")
        s = Replace(s, ".lnk", "")
    'Finals
        ArrNnVn = Split(vbCrLf + s, vbCrLf)
End Function

Sub T4_Add_NewLineInT5_GiveVnNn(VnNn$)
    'Called from    Create_AllBracLinks_InsideOneFolder
    'VnNn           = Vorname Nachname = "Jo Hipp"
    'Action         stellt Existenz einer T5-Zeile sicher (zu VnNn)
    
    'Vorbereitung
        If VnNn = "" Then Stop
        Dim mw$, Nation$, Nn$, PF$, Vn$, c%, i%, zT5%, r As Range
        zT5 = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, VnNn)
        If zT5 > 0 Then Exit Sub 'VnNn existiert bereits in T5
    'Vn, Nn
        '|Nachname|Vorname| ist nicht in T5 enthalten
        c = InStr(1, VnNn, " "): Vn = Left(VnNn, c - 1): Nn = Mid(VnNn, c + 1)
        zT5 = Get_NrOfLastRowInColumnNr(3, "T5") + 1
    'w/m? Nation?
        ShowUF1 VnNn
        mw = ArrC(49): Nation = ArrC(50)
    'Neue Zeile in T5 anlegen und formatieren
        With Sheets("T5")
            'Hintergrundfarbe der T5-Zeile
                Set r = .Range(.Cells(zT5, 3), .Cells(zT5, 10))
                If mw = "w" Then r.Interior.Color = 14083324 Else r.Interior.Color = 15652797 'hellblau/Blau1 (m-Farbe)
                Set r = .Range(.Cells(zT5, 12), .Cells(zT5, 22))
                If mw = "w" Then r.Interior.Color = 14083324 Else r.Interior.Color = 15652797 'hellblau/Blau1 (m-Farbe)
            'BorderAround
                For i = 3 To 22
                    Set r = .Cells(zT5, i)
                    r.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, ColorIndex:=2
                Next
            'ZeilenNr
                Set r = .Cells(zT5, 2): r.Value = zT5 - 7
                r.HorizontalAlignment = xlCenter: r.Font.size = 6
            'Nachname, Vorname
                .Cells(zT5, 3) = Nn:      .Cells(zT5, 4) = Vn
            'm/w
                Set r = .Cells(zT5, 5): r.Value = mw: r.HorizontalAlignment = xlCenter
            'Jahrgang, Verein
                Set r = .Cells(zT5, 6): r.HorizontalAlignment = xlCenter
'            'Club/Verein
'                .Cells(zT5, 7) = VerDg
            'Nation
                Set r = .Cells(zT5, 9): r.Value = Trim(Left(Nation, 3))
                r.HorizontalAlignment = xlCenter
            'PersonenOrdner
                PF = Nn + ", " + Vn: If Nation <> "D" Then PF = PF + " (" + Nation + ")"
                .Cells(zT5, 10) = PF
            'ZusatzSpalte 11    'Hilfsspalte, hellgraue Schrift
                Set r = .Cells(zT5, 1): r.Font.Color = RGB(155, 155, 155): r.HorizontalAlignment = xlLeft: r.Font.size = 6
            'ZusatzSpalten 12-21
                Set r = .Range(.Cells(zT5, 12), .Cells(zT5, 21)): r.Value = "-": r.HorizontalAlignment = xlCenter
            'ZusatzSpalte 22: Vn Nn
                Set r = .Cells(zT5, 22): r.Value = VnNn: r.Font.size = 8
            'Zusatzspalte 23
                .Cells(zT5, 23) = "|"
        End With
End Sub

Function Get_VnNn_OfBracFiles_InsideOneFolder(PathOfFolder$)
    'Called from    T4_Create_MissingBracLinks_OfOneEvent
    'Action         xxx
    
    'Vorbereitung
        Dim FN$, VnNn$, s$, i%, j%, A() As String, B() As String
    'FN = FileNames in EventFolder mit "19xx*(Vn Nn)*"
        FN = Get_AllFileNames_Like_OfOneFolder(PathOfFolder, "[12][90]*([A-ZÄÖÜ]* *[A-ZÄÖÜ]*)*")
        '  = "1987*DM*(Vn Nn)*.jpg[v]..."
    's = alle VnNn aus den JpgFiles (im EventFolder) (auch NoCops = NonCompetitors)
        A = Split(FN, vbCrLf): s = "|"
        For i = 0 To UBound(A)                              'A(0) = "1987*DM*(Vn Nn)*.jpg"
            B = Split(A(i), "(")
            For j = 1 To UBound(B)
                VnNn = Left(B(j), InStr(1, B(j), ")") - 1)  'VnNn = "Lea Hipp"
                If Not VnNn Like "*,*" Then
                    If Not s Like "*|" + VnNn + "|*" Then s = s + VnNn + "|"
                End If
            Next                                            's    = "|Jo Hipp|Lea Hopp|...|"
        Next
    'Finals
    Get_VnNn_OfBracFiles_InsideOneFolder = s
    'show s
End Function

Sub T4_Create_MissingBracLinks_OfOneEvent(PathOfEventFolder$)
    'Called from    Copy_BracFiles_OfOneFolder_ToPersonFolders
    'Action         Alle VnNn aus den JpgFiles ermitteln; s = "|Jo Hipp|Lea Hopp|...|";
    '               alle NnVn aus den Links ermitteln;    L = "|Hopp, Lea|...|Hipp, Kai|"
    '               fehlende Links erzeugen (für s-Namen nicht in L)
    
    'Vorbereitung
        Dim JmpToFo$, L$, LnkName$, NnVn$, s$, ToDo$
        Dim c%, i%, A() As String
        'PathOfEventFolder = "F:\Archiv TR\Archiv Trampolin\Events\1980-09-13 DJEMM Stadtallendorf"
        DoArrc
        s = Get_VnNn_OfBracFiles_InsideOneFolder(PathOfEventFolder) '"|Jo Hipp|Lea Hopp|...|"
    'L = alle Links in EventFolder
        L = Get_AllFileNames_Like_OfOneFolder(PathOfEventFolder, "*.lnk")
        ' = "Hipp, Lea (TV Bonn).lnk[v]...[v]Hopp, Kai (Abc).lnk"
    'L = alle NnVn aus den Links (im EventFolder)
        A = Split(L, vbCrLf)
        For i = 0 To UBound(A)
            A(i) = Replace(A(i), ".lnk", "")        'Hipp, Lea (TV Bonn)
            c = InStr(1, A(i), " (")
            If c > 0 Then A(i) = Left(A(i), c - 1)  'Hipp, Lea
        Next
        L = "|" + Join(A, "|") + "|"
        ' = "|Hipp, Lea|...|Hopp, Kai|"     alle Namen existierender PersonLinks (im EventFolder)
    'Welche der s-Namen haben bereits einen Link (zu ihrem PersonOrdner)?
        A = Split(s, "|"): ToDo = "|"
        For i = 1 To UBound(A) - 1
            c = InStr(1, A(i), " ")
            NnVn = Mid(A(i), c + 1) + ", " + Left(A(i), c - 1)
            If Not L Like "*|" + NnVn + "|*" Then ToDo = ToDo + A(i) + "|"
        Next
        If ToDo = "|" Then Exit Sub
    'ToDo                'Namen, die einen Link bekommen sollen
        'show ToDo: Stop 'ToDo = "|Jo Hipp|Lea Hopp|...|"
        A = Split(ToDo, "|")
        For i = 1 To UBound(A) - 1
            JmpToFo = T5_Get_PathOfPersonfolder_Give_VnNn(A(i))
            LnkName = NameOfPath(JmpToFo)
            Create_OneLinkFile PathOfEventFolder, JmpToFo, LnkName 'CrInFo, JmpToFo, LnkName
        Next
End Sub


