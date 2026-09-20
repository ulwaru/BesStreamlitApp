Attribute VB_Name = "M43"
Option Explicit 'M_Dg

Sub zzz_M_Dg()
    showProcs "t4some"
    'RenameModule "MT7", "M_T7"
    'T1_ShowLogBuch
    EE 1: Beep
    'T4_RoRe 0
End Sub

Function T4_Get_DgEventFolderName_by_zSh_sSh(z&, s&)
    'Vorbereitung
        Dim F$, i&, s1&, s2&, z1&, z2&, A(), r As Range
    'A() '5 Spalten aus T4SomeDgData
        T4_Load_T4SomeDgData_2to6 A
    'Dg-Ranges durchlaufen
        For i = 1 To UBound(A, 1)
            z1 = A(i, 1): s1 = A(i, 2): z2 = A(i, 3): s2 = A(i, 4)
            Set r = Range(Cells(z1, s1), Cells(z2, s2))
            If Cell_is_in_Range(z, s, r) Then F = A(i, 5): Exit For
        Next
        'Call Beep: EE 0: Cells(z, s).Select: EE 1: show F
        T4_Get_DgEventFolderName_by_zSh_sSh = F
End Function

Function Cell_is_in_Range(z As Long, s As Long, r As Range) As Boolean
    Dim targetCell As Range
    Set targetCell = r.Worksheet.Cells(z, s)
    
    Cell_is_in_Range = Not Intersect(targetCell, r) Is Nothing
End Function





Sub ClickButton_CreateJpg_InActualDg() '_Create_Jpg_EventLink_InDgPersonFolders_
    'Called from    CmdCreate_Click [UserClick on 'Create jpg'-Button inside a Design]
    'Action         (1) Erzeugt ein #result.jpg des aktuellen Dg im DgEventOrdner;
    '               (2) kopiert #result jpg     in alle PersonFolder der DgCompetitors
    '               (3) kopiert °Files          in alle PersonFolder der DgCompetitors
    '               (4) kopiert BracFiles       in ihre PersonFolder
    '               (5) legt BracLinks          in den EventFolder
    '               (6) legt ein EventLink      in alle PersonFolder der DgCompetitors
    '               (7) legt ein DateSquare     in alle PersonFolder der DgCompetitors
    'Brac           = VnNn with   Brackets      = "(Lea Boll)"; inside a FileName
    'BracName       = VnNn inside Brackets      = "Lea Boll"
    'BracFile       = FileName containing Brac  = "19870605 DM (Lea Boll)...jpg"
    'BracFolder     = PersonFolder of BracName  = "...\Leute\Boll, Lea (TV Bonn)"
    'BracLink       = Link from EventFolder to BracFolder
'Vorbereitung
        Dim A$, N2$, p$, PathOfEventFolder$, T$, T2$, s1%, s2%, z1%, z2%, ZeTrenn%
        Dim r As Range
        Call DoArr: EE 0: With Sheets("T4"): Screen0
        Set r = .Range(ArrC(36))    'Range of choosen design
        z1 = CInt(ArrC(37)): z2 = CInt(ArrC(38)): s1 = CInt(ArrC(39)): s2 = CInt(ArrC(40))
        PathOfEventFolder = ArrC(57)
    '(1) CreateResultsJpg
        'N2 = Name of ResultJpg (übernimmt Pid)
        '   pID übernehmen, falls eine solche in der Vorversion des aktuellen ResultJpg enthalten war
            T = .Cells(z1 + 1, s1 + 1)       'TitleInsideJpg
            'T wird nicht benutzt; weg?     'zur Kontrolle? T sollte einen DgTitle tragen
            N2 = T4_GetNameOfResultJpg()    '"19660508 .DJMM Nürnberg #result ..."
            FillArrC 43, N2
        'p  = Path of #resultJpg
            p = PathOfEventFolder + "\" + N2 'ArrC(3) = Path of Folder "Events"
        'CellButtons kurzzeitig ausblenden
            .CmdCreate.Visible = False: .CmdOpen.Visible = False
            .CmdBack.Visible = False:   .Cmdinfo.Visible = False: .CmdSingleRanking.Visible = False
            T4_LabelForCountDown_Hide
        CreateResultsJpg p, r
        'CellButtons wieder einblenden
            .CmdCreate.Visible = True: .CmdOpen.Visible = True: .CmdBack.Visible = True
            .Cmdinfo.Visible = True: T4_LabelForCountDown_Show
            If ArrDg(2, 1) = "M2V8" Then .CmdSingleRanking.Visible = True
            ActiveWindow.Zoom = 100
            LogBuch "ResultJpg was created: " + N2
    '(2)-(7)
        T4_Copy_ResultJpgToAllCompetitors p                                 '(2)
        T4_Copy_°FilesOfOneEvent_ToPersonFolders PathOfEventFolder          '(3)
        Copy_BracFiles_OfOneFolder_ToPersonFolders PathOfEventFolder      '(4)
        T4_Create_MissingBracLinks_OfOneEvent PathOfEventFolder                    '(5)
        T4_Create_EventLink_DateSquare_OfOneEventfolder_ToPersonFolders PathOfEventFolder '(6) (7)
    'Finals
        EE 1: screen1: End With: Beep
End Sub

Sub T4_Copy_ResultJpgToAllCompetitors(PathOfResultJpg$)
    'Called from    ClickButton_CreateJpg_InActualDg
    'Action         Get all CompetitorNames of actual Dg (T4GetSomeDgData-NamesSpalte)
    '               Copy ResultJpg to all Competitor-PersonFolder
    '               (competitors only; no other LinkFiles)
    
    'Vorbereitung
        Dim DgTitle$, L$, NameOfPersonalFolder$, names$, Nn$, PathOfEventFolder$
        Dim PathOfPersonalFolder$, PathOfPersonalJpg$, Report$, v$, Vn$, VnNn$
        Dim i%, zT4%, zT5%, Arr1() As String, Arr2() As String
        Dim FolderInWhitchTheLinkShouldBeCreated$, FileOrFolderToJumpTo$, LinkNameToDisplay$
        Call DoArrc: DoArrDg: PathOfEventFolder = ArrC(3) + "\" + CStr(ArrDg(1, 2))
        DgTitle = CStr(ArrDg(2, 2)): v = vbCrLf
    'Get all CompetitorNames of actual Dg
        zT4 = Get_RowNr_HoldingMyTextWholeInColumnX("T4", 5, 7, DgTitle)
        '   = ZeilenNr T4GetSomeDgData
        names = Sheets("T4").Cells(zT4, 12)
        '     = Names-Zelle der Zeile zT4 in der T4GetSomeDgData-NamesSpalte
        Arr1 = Split(names, "#")
    'Copy to all
        For i = 1 To UBound(Arr1)
            L = Arr1(i)         '|Nicole Mai|6|4|Mannschaft|1   'Competitor in actual event
            If Left(L, 1) <> "|" Then L = "|" + L
            Arr2 = Split(L, "|")
            'Name in T5 suchen
                VnNn = Arr2(1)
                zT5 = Get_RowNr_HoldingMyTextPartInColumnX("T5", 22, 7, VnNn)
                ' = ZeilenNr der T4GetSomeDgData-NamesZelle, die 'Vn Nn' enthält
                
                If zT5 = 0 Then
                    Report = Report + VnNn + " is not in T5" + v
                Else
                    NameOfPersonalFolder = Sheets("T5").Cells(zT5, 10) 'T5
                    PathOfPersonalFolder = ArrC(4) + "\" + NameOfPersonalFolder
                    If Not FileExists(PathOfResultJpg) Then
                        Report = Report + "File does not exist: " + PathOfResultJpg + v
                    End If
                    If Not FolderExists(PathOfPersonalFolder) Then
                        Report = Report + "Folder does not exist: " + PathOfPersonalFolder + v
                    End If
                    'Write Jpg
                    If FileExists(PathOfResultJpg) And FolderExists(PathOfPersonalFolder) Then
                        PathOfPersonalJpg = PathOfPersonalFolder + "\" + NameOfPath(PathOfResultJpg)
                        CopyFile PathOfResultJpg, PathOfPersonalJpg
                    End If
                End If
        Next
        If Report <> "" Then show Report
End Sub

Function T4_ChangeNameGlobal_Get_zT5_NewName(VnOldNnOld$) As Integer
    'Called from    T4_Copy_ResultJpgToAllCompetitors
    
    'Vorbereitung
        Dim NnNew$, NnOld$, Ns$, VnNew$, VnOld$, c%, zT5%, Arr1() As String
        Ns = ArrC(78)               'T4 ChangeNameGlobal:|VnOld|NnOld|VnNew|NnNew|
        c = InStr(1, VnOldNnOld, " ")
        VnOld = Left(VnOldNnOld$, c - 1): NnOld = Mid(VnOldNnOld$, c + 1)
        Arr1 = Split(Ns, "|")
        If Arr1(1) = VnOld And Arr1(2) = NnOld Then
            VnNew = Arr1(3): NnNew = Arr1(4)
            zT5 = Get_RowNr_HoldingMyTextPartInColumnX("T5", 22, 7, VnNew + " " + NnNew)
        End If
    'Finals
        T4_ChangeNameGlobal_Get_zT5_NewName = zT5
End Function

Sub CreateResultsJpg(p$, r As Range)
    Dim H&, w&, ws As Worksheet, c As ChartObject
    Set ws = Sheets("T4")
    
'    Stop
'    r.Select
'    Stop
    
    ActiveWindow.Zoom = 400
    DoEvents
    r.CopyPicture xlScreen, xlPicture
    w = r.Width: H = r.Height
    Set c = ws.ChartObjects.Add(Left:=0, Top:=0, Width:=w, Height:=H)
    c.Activate
    With c.Chart
        .Paste: .Export fileName:=p, Filtername:="JPG"
    End With
    c.Delete
End Sub

Sub T4_ActionsOnDgClickNone()
    'Called from    T4_Handle_UserClick, T8_Jump_ToDg
    'Status         User klickte außerhalb eines Design-Rahmens
    
    'Hide RoundRect
        T4_RoRe 0
    'Change some ArrC-Items
        ArrC(61) = ""      'ActiveCellContent (Keylogger) = "-"
        ArrC(36) = "-"     'Range      of choosen design
        ArrC(37) = "-"     'ZeilenNr   of DesignBorder-Top
        ArrC(38) = "-"     'ZeilenNr   of DesignBorder-Bottom
        ArrC(39) = "-"     'SpaltenNr  of DesignBorder-Left
        ArrC(40) = "-"     'SpaltenNr  of DesignBorder-Right
        ArrC(41) = "-"     'Titel      to see inside design and jpg
        ArrC(42) = "-"     'FolderName to write jpg into
        ArrC(43) = "-"     'Name       of #result-jpg
        ArrC(44) = "-"     'Name       of DgGroup
        ArrC(45) = "-"     'zClubs
        ArrC(46) = "-"     'DgId
        ArrC(48) = "0"     'DgClickNr
        ArrC(57) = "-"     'EventFolderPath to write jpg into
        ArrC(60) = "-"     'Write ResultJpg? 0 = No, 1 = Yes
        ArrC(65) = "-"     'Dg ActiveCell = "W", "M" oder "X" (Damen, Herren, Aktive)
        ArrC(67) = "-"     'Design: Nr of actual DesignGroup
        ArrC(68) = "-"     'Design: AllExisting Calc DesignGroups
    'sDgNameW sDgNameM sDgNameX sDgNationW ...:
        ArrC(90) = "0":   ArrC(91) = "0":  ArrC(92) = "0": ArrC(93) = "0": ArrC(94) = "0"
        ArrC(95) = "0":   ArrC(96) = "0":  ArrC(97) = "0": ArrC(98) = "0": ArrC(99) = "0"
        ArrC(100) = "0": ArrC(101) = "0": ArrC(102) = "0"

    'Paste
        PasteArrC
        Format_T8_ArrC_Blue
        T4_Check_CountOfDgs
        T4_Check_DgPositions_ViaT4SomeDgData
End Sub

Sub Hide_HelperBox()
    With Sheets("T4").Shapes("HelperBox")
        .Left = 0: .Top = 0: .Visible = False
    End With
End Sub

Sub Hide_CellBox()
    With Sheets("T4").Shapes("CellBox")
        .Left = 0: .Top = 0: .Visible = False
    End With
End Sub

Function InRange(Range1 As Range, Range2 As Range) As Boolean
    ' returns True if Range1 is within Range2
    InRange = Not (Application.Intersect(Range1, Range2) Is Nothing)
End Function

Sub TEST_InRange()
    If InRange(ActiveCell, Range("A1:D100")) Then
        ' code to handle that the active cell is within the right range
        MsgBox "Active Cell In Range!"
    Else
        ' code to handle that the active cell is not within the right range
        MsgBox "Active Cell NOT In Range!"
    End If
End Sub

Function T4_CellIsInDg(z%, s%) As Boolean
    If ArrC(36) = "-" Then Exit Function
    With Sheets("T4")
        T4_CellIsInDg = InRange(.Cells(z, s), .Range(ArrC(36)))
    End With
End Function

Sub HideDesignButtons()
    With Sheets("T4")
    .CmdCreate.Visible = False: .CmdOpen.Visible = False:         .CmdBack.Visible = False
    .Cmdinfo.Visible = False:   .CmdSingleRanking.Visible = False
    End With
End Sub

Sub ShowDesignButtons()
    T4_ReplaceDesignButtons
    Sheets("T4").CmdCreate.Visible = True
    Sheets("T4").CmdOpen.Visible = True
    Sheets("T4").CmdBack.Visible = True
    Sheets("T4").Cmdinfo.Visible = True
    If ArrDg(2, 1) = "M2V8" Then Sheets("T4").CmdSingleRanking.Visible = True
    screen1
End Sub





