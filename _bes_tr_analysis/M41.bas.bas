Attribute VB_Name = "M41"
Option Explicit 'M41

Sub zzz_M41()
    
    'showProcs "check name"
    
    
    'RenameModule "M_Dg", "M43"
    'RenameModule "Modul1", "M41"
    'T1_ShowLogBuch
    'show CStr(ArrDg(1, 1))
    'ScrollRightSlow 5
    'show CStr(zScroll)
    EE 1: Beep
End Sub

Sub T4_Check_ActualDgFolder_forNonVbsLinks()
    'Called from    T4_ActionsOnDgClickNext
    'Action         prüft, ob lnkFile.Description = "VbsLink";
    '               falls nein: neuen VBSLink erstellen
    'Vorbereitung
        Dim D$, F$, L$, NameOfLinkFile$, NnVn$, PathOfLinkFile$, s$, T$, v$, VnNn$
        Dim c&, i&, A() As String, lnkFile As Object, w As Object
        v = vbCrLf: DoArrc
    'Actual EventFolder
        F = ArrC(3) + "\" + ArrC(42)   'T4 Dg: Path of actual EventFolder
        L = Get_AllFileNames_Like_OfOneFolder(F, "*.lnk") 'Baechler, Urs (TV Bern-Berna, CH).lnk
        A = Split(L, v)
        Set w = CreateObject("WScript.Shell")
    For i = 0 To UBound(A)
        NameOfLinkFile = A(i)
        PathOfLinkFile = F + "\" + NameOfLinkFile 'Path of one LinkFile in actual EventFolder
        Set lnkFile = w.CreateShortcut(PathOfLinkFile)
        D = lnkFile.Description
        If D <> "VbsLink" Then
            'falls ", " in NameOfLinkFile enthalten ist, lohnt sich eine NamenPrüfung
            If NameOfLinkFile Like "*, *" Then
                'NameOfLinkFile enthält ein ", "
                T = Replace(NameOfLinkFile, ".lnk", ""): c = InStr(1, T, " (")
                If c > 0 Then T = Left(T, c - 1)            'Mai, Lea
                't trägt jetzt den Text des NameOfLink ohne Klammer
                Extract_VnNn_NnVn_FromTxt VnNn, NnVn, T
                If VnNn = "" Then
                    'NameOfLinkFile enthält keinen Namen
                    s = s + "NoName | " + PathOfLinkFile + v
                Else
                    'NameOfLinkFile enthält einen Namen

                    Create_OneVbsLink_Person PathOfLinkFile
                    s = s + "NewVbs | " + PathOfLinkFile + v
                End If
            Else
                'NameOfLinkFile enthält kein ", "
                s = s + PathOfLinkFile + v
            End If
        End If
    Next i
    
    If s <> "" Then show "T4_Check_ActualDgFolder_forNonVbsLinks" + v + v + s
        
End Sub

Sub T4_Add_NewAgesToDg(z1%, z2%, s1%, s2%)
    'Called from    T4_ActionsOnDgClickFirst
    
    'Vorbereitung
        Dim A$, B$, N$, sN$, i&, j&, sN1&, sN2&, Ye&, Yn&, r As Range
        Dim N1() As String, N2() As String, T5j() As String, T5v() As String
        Ye = CLng(Left(ArrDg(2, 2), 4)) 'YearOfEvent
    'Namenspalten ermitteln
        For i = 2 To UBound(ArrDg, 2) - 1
            If "WMX" Like "*" + ArrDg(5, i) + "*" Then sN = sN + Format(i, "00") '0319
        Next
        If sN = "" Then Exit Sub
        sN1 = CLng(Left(sN, 2)): If Len(sN) = 4 Then sN2 = CLng(Right(sN, 2))
    'N1(), N2() NamenArray; Dg-Namen ohne Altersangabe " (17)"
        'getrennte Arrays für späteres Paste
        ReDim N1(1 To UBound(ArrDg, 1) - 6): ReDim N2(1 To UBound(ArrDg, 1) - 6)
        For i = 6 To UBound(ArrDg, 1) - 1
            A = ArrDg(i, sN1)
            If A Like "* (*" Then A = Left(A, InStr(1, A, " (") - 1)
            N1(i - 5) = A
            If sN2 > 0 Then
                B = ArrDg(i, sN2)
                If B Like "* (*" Then B = Left(B, InStr(1, B, " (") - 1)
                N2(i - 5) = B
            End If
        Next ': showArray N2
    'Nach Alter in T5 suchen
        T5_Load_VnNn T5v: T5_Load_Jhg T5j
    'N1
        For i = 1 To UBound(N1)
            N = N1(i) 'one DgName, VnNn
            For j = 1 To UBound(T5v)
                If T5v(j) = N Then
                    If T5j(j) <> "" Then
                        'N besitzt Jhg in T5
                        Yn = CLng(T5j(j)) 'Year of Name-N-Competitor
                        N1(i) = N1(i) + " (" + CStr(Ye - Yn) + ")"
                    End If
                End If
            Next
        Next ': showArray N1
        'Paste
        Paste_1DArrayToCol "T4", z1 + 5, s1 + sN1 - 1, N1
    'N2
        If sN2 > 0 Then
            For i = 1 To UBound(N2)
                N = N2(i) 'one DgName, VnNn
                For j = 1 To UBound(T5v)
                    If T5v(j) = N Then
                        If T5j(j) <> "" Then
                            'N besitzt Jhg in T5
                            Yn = CLng(T5j(j)) 'Year of Name-N-Competitor
                            N2(i) = N2(i) + " (" + CStr(Ye - Yn) + ")"
                        End If
                    End If
                Next
            Next ': showArray N2
        End If
        'Paste
        If sN2 > 0 Then Paste_1DArrayToCol "T4", z1 + 5, s1 + sN2 - 1, N2
        DoEvents
End Sub

Sub T4_ActionsOnDgClickFirst(St$)
    'Called from    T4_Handle_UserClick
    'st             = "1|z1|s1|z2|s2|Title" '1|0000|0000|0000|0000|Title
    '               1 = Status (FirstClick); z1, s1, z2, s2: je 4 Ziffern
    'Status         User klickte in ein bisher nicht aktives Dg
    '               ArrC 36-40 sind gefüllt (DgRange,z1,z2,s1,s2)
    'Action         new: PersonData.txt existiert nicht mehr; jetzt alles in T5
    '               old: Eventuell neue ArrPd-Werte werden in PersonData.txt gesichert
    '                    ArrDg wird geladen; mit DgZeile5 als T4Zeile2 (MWXVLTYAEGR)
    
    'Vorbereitung
        Dim DgFolderName$, DgFolderPath$, Title$, s1%, s2%, z1%, z2%, rg As Range
        If Left(St, 1) <> "1" Then Stop
        With Sheets("T4")
    'FillArrC 36-41 48 - z1,z2,s1,s2,Range,Title,FirstClick
        z1 = CInt(Mid(St, 3, 4)): s1 = CInt(Mid(St, 8, 4))
        z2 = CInt(Mid(St, 13, 4)): s2 = CInt(Mid(St, 18, 4)): Title = Mid(St, 23)
        FillArrC 36, Get_ColumnLetter(s1) + CStr(z1) + ":" + Get_ColumnLetter(s2) + CStr(z2)
        FillArrC 37, CStr(z1): FillArrC 38, CStr(z2): FillArrC 41, Title
        FillArrC 39, CStr(s1): FillArrC 40, CStr(s2): FillArrC 48, "1"
    'ArrDg neu belegen
        FillArrDg 'auch FillArrC 42, 44, 46-47, 49-54, 57
    'DgRowHeight 15
        Set rg = Range(ArrC(36)): rg.RowHeight = 15 'rg = Range of choosen design
    'Look for new T5-Ages
       T4_Add_NewAgesToDg z1, z2, s1, s2
    'Finals
        T4_ActionsOnDgClickNext
        End With
End Sub

Sub T4_ActionsOnDgClickNext()
    'Called from    T4_Handle_UserClick
    'Status         Ein Dg ist bereits aktiv; kein DgFirstClick;
    '               der eben ausgeführte Klick ist ein weiterer Klick im gleichen Dg;
    '               T4_ActionsOnDgClickFirst sind bereits durchgeführt
    '               (neues ArrDg, neue Werte in ArrC, ...)
    
    'Vorbereitung
        Dim sSh%, zSh%
        zSh = CInt(ArrC(104))          'T4   SelectionChange to SheetZeilenNr  zSh
        sSh = CInt(ArrC(105))          'T4   SelectionChange to SheetSpaltenNr sSh
        
    'ArrDg neu belegen (soll bei jedem Klick geschehen)
        FillArrDg
    'Add 'z1s1Aktuell-->ArrC(46) 'z1s1Last-->ArrC(47) 'Dg-Wechsel?
        T4_Write_DgLO_Last1Last2_ToArrC
    'T4_DgClick CheckOneName_LastSelected
        'CompetitorName (der gerade verlassene Zelle) komplett versorgen;
        'Zusatzangaben in NameCell herausfiltern und abarbeiten (j= c= ... n= global)
        'ggf. +PersonFolder; T5 ergänzen; ggf. +Age +Club +Nation in Dg
        '+PersonLink in den EventFolder; update Names in T4SomeDgData
        T4_DgClick_CheckOneName_LastSelected 'Jhg, m/w, Verein/Ort, Nation; Liga: ZellFarbe in Dg
    'Alle Links des aktuellen EventFolders prüfen ob es VBSLinks sind
        'T4_Check_ActualDgFolder_forNonVbsLinks 'dauert zu lange
    'FontBlack if ClickedCell = ""
        If Sheets("T4").Cells(zSh, sSh) = "" Then Sheets("T4").Cells(zSh, sSh).Font.Color = vbBlack
    'UCase Nation
       T4_Check_OneDgNation_UCase_ValidNation 'T4_Change_AllDgNation_ToUCase
    'Write Nation of previous competitor to T5
        T4_Write_NationToT5
    'Write Club of actual competitor to T5
        T4_Write_ClubToT5
    'AutoFit
        T4_AutoFit
    'AutoCalc
        T4_AutoCalc  'berechnet Ergebnisse
    'DesignButtons Show/Hide
        If ArrC(57) <> "-" Then ShowDesignButtons
    'HelperBox
        HelperBoxInit '--> CellBox, HelperBox, CellBox_KeyDown
    'Finale
        EE 1: Application.ScreenUpdating = True
End Sub

Sub T4_Write_DgLO_Last1Last2_ToArrC()
    'Called from    T4_ActionsOnDgClickNext
    'Status         FillArrDg wurde gerade durchgeführt
    'Action         ermöglicht die Feststellung eines Dg-Wechsels;
    '               registriert in ArrC, in welche Dgs die letzten beiden Klicks erfolgten
    
    'Vorbereitung
        Dim s1%, z1%
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39))    'T4(z1,s1) = Dg-Ecke li ob
    'Action
        FillArrC 47, ArrC(46)                   'Dg-LO-Cell-Coord Last2 (previous selected)
        FillArrC 46, CStr(z1) + "-" + CStr(s1)  'Dg-LO-Cell-Coord Last1 (aktuell) '2150-165
End Sub

Sub T4_ReplaceDesignButtons()
    Dim r!, T!, w!, s1%, s2%, z1%
    z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39)): s2 = CInt(ArrC(40))

    With Sheets("T4")
    '.CmdSingleRanking.Visible = True
    r = .Cells(z1, s2 + 1).Left: T = .Cells(z1, s2).Top + 3
    .Cmdinfo.Top = T: .CmdCreate.Top = T: .CmdOpen.Top = T: .CmdBack.Top = T: .CmdSingleRanking.Top = T
    .Cmdinfo.Height = 20: .CmdCreate.Height = 20: .CmdOpen.Height = 20: .CmdBack.Height = 20: .CmdSingleRanking.Height = 20
    
    .Cmdinfo.Width = 20:            .Cmdinfo.Left = r - 23
    .CmdCreate.Width = 46:          .CmdCreate.Left = r - 71
    .CmdOpen.Width = 52:            .CmdOpen.Left = r - 126
    .CmdBack.Width = 20:            .CmdBack.Left = r - 149
    .CmdSingleRanking.Width = 54:   .CmdSingleRanking.Left = r - 206

    End With
End Sub

Sub FillArrDg()
    'Called from    T4_ActionsOnDgClickFirst, T4_ActionsOnDgClickNext
    'Dg             = Design; Sheet-T4-Rechteck; trägt Ergebnisliste eines Events
    'T4Row2_DgRow5  = Sheet-Zeile 2 oberhalb des Designs WillBeCopiedTo ArrDg-Zeile 5
    'ArrDg          = globales 2D-Array; alle T4-Zellen des aktuell aktiven Dg
    '                 + Zusatzeinträge (Rang, Verein, T4Row2_DgRow5)
    '               wird bei jedem Klick in ein Dg neu geladen
    'Zusatzeinträge Dg(1,1)=LO; Dg(1,2)=DgGroup; Dg(1,3)=Best4; Dg(1,4)=FixValues
    '               Dg(2,1)=Folder; Dg(3,1)=SingleRanking
    'Action         Lädt ArrDg
    '               FillArrC 42, 44, 46-47, 49-54, 57
    
    'Vorbereitung
        'Ticks "FillArrDg Start"
        Call DoArrc: If ArrC(37) = "-" Then Exit Sub 'z1
        Dim A$, ClubName$, DgFolderName$, DgFolderPath$, zClubs$, zsCompetitors$
        Dim i%, j%, s1%, s2%, z1%, z2%
        z1 = CInt(ArrC(37)): z2 = CInt(ArrC(38)): s1 = CInt(ArrC(39)): s2 = CInt(ArrC(40))
        With Sheets("T4")
    'ArrDg                         'wird bei jedem Klick in ein Dg neu geladen
        ArrDg = .Range(.Cells(z1, s1), .Cells(z2, s2)).Value
    'FillArrC 90-102
        T4_Add_T4Row2_ToArrDg
        FillArrC_DgColumnNrs_FromDgRow5 'Auswertung BuchstabenKette_T4Zeile2
    'DgFolderName
        DgFolderName = CStr(ArrDg(1, 2)): FillArrC 42, DgFolderName
    'DgFolderPath
        DgFolderPath = ArrC(3) + "\" + DgFolderName: FillArrC 57, DgFolderPath
       If FolderExists(DgFolderPath) Then
            ShowDesignButtons
        Else
            'DgFolderPath existiert nicht
            CreateFolder DgFolderPath
            'showArray2D ArrDg: Stop
       End If
    'Check
        T4_Check_Best4
        T4_Add_FixValues_ToDg_1_4
        T4_Check_DgGroup    'FillArrC 44
        'T4_Check_SingleRanking
        T4_Change_IsNull_ToEmptyString
    'Add
        T4_Add_Rank_ToAllArrDg_SyPartner
        T4_Add_Nation_ToAllArrDg_SyPartner
        T4_Add_Nation_ToAllArrDg_TeamMembers
        T4_Add_ClubNameAndRank_ToMembersInArrDg_M2V8
        T4_Add_Rank_ToAllArrDg_TeamMembers_NotM2V8
    'zClubs M2V8
        zClubs = T4_Get_AllClubNameRowNrs_M2V8: If zClubs = "" Then zClubs = "-"
        FillArrC 45, zClubs
        End With ': showArray2D ArrDg: Stop
        'Ticks "FillArrDg End"
End Sub

Function T4_Get_ESyM() As String
    'Called from    T4_Change_T4SomeDgData_OneNameData
    
    'Vorbereitung
        Dim DgGroup$, ESyM$, s$
        DgGroup = ArrDg(2, 1): s = Left(DgGroup, 1)
    'Action
        If Not "ESML" Like "*" + s + "*" Then Stop
        If s = "E" Then
            ESyM = "Einzel"
        ElseIf s = "S" Then ESyM = "Synchron"
        ElseIf s = "M" Then ESyM = "Mannschaft"
        ElseIf s = "L" Then ESyM = "Mannschaft" 'LIGA
        End If
    'Finals
        T4_Get_ESyM = ESyM
End Function

Sub T4_Check_SingleRanking()
    'Called from    FillArrDg
        
    'Vorbereitung
        Dim s1%, z1%, r As Range
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39)): Set r = Sheets("T4").Cells(z1 + 2, s1)
    'Check
        If Not Left(ArrDg(3, 1), 13) = "SingleRanking" Then
            ArrDg(3, 1) = "SingleRanking": r.Value = "SingleRanking": r.Font.Color = vbWhite
        End If
End Sub

Sub T4_Check_DgGroup()
    'Called from    FillArrDg
        
    'Vorbereitung
        Dim DgGroup$, s1%, z1%, r As Range, w As Worksheet
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39))
        Set w = Sheets("T4"): Set r = w.Cells(5, s1): DgGroup = r.Value
        FillArrC 44, DgGroup: Set r = w.Cells(z1 + 1, s1)
    'Check
        If ArrDg(2, 1) = "" Then 'DgGroup steht noch nicht im Dg
            ArrDg(2, 1) = DgGroup: r.Value = DgGroup: r.Font.Color = vbWhite
        End If
        
End Sub

Sub T4_Add_FixValues_ToDg_1_4()
    'Called from    FillArrDg
    
    'Vorbereitung
        Dim s1%, z1%, r As Range
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39)): Set r = Sheets("T4").Cells(z1, s1 + 3)
    'Action
        If Not ArrDg(1, 4) Like "FixValues: *" Then
            ArrDg(1, 4) = "FixValues: " + ArrDg(1, 4): r.Value = ArrDg(1, 4): r.Font.Color = vbWhite
        End If
End Sub

Sub T4_Check_Best4()
    'Called from    FillArrDg
        
    'Vorbereitung
        Dim s1%, z1%, r As Range
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39)): Set r = Sheets("T4").Cells(z1, s1 + 2)
    'Check
        If Not Left(ArrDg(1, 3), 4) = "Best" Then
            ArrDg(1, 3) = "Best": r.Value = "Best": r.Font.Color = vbWhite
        End If
End Sub

Sub T4_Change_IsNull_ToEmptyString()
    'Called from    FillArrDg
        
    'Vorbereitung
        Dim z%, s%
    'Change
        For z = 1 To UBound(ArrDg, 1)
            For s = 1 To UBound(ArrDg, 2)
                If IsNull(ArrDg(z, s)) Then ArrDg(z, s) = ""
            Next
        Next
End Sub

Sub T4_Add_Rank_ToAllArrDg_TeamMembers_NotM2V8()
    'Called from    FillArrDg
    
    'Exit
        If ArrDg(2, 1) = "M2V8" Then Exit Sub
    'Vorbereitung
        Dim i%, s%
    'Add
        For s = 2 To UBound(ArrDg, 2) - 1
            'Schleife über die Spalten der DgZeile5
            If "IJK" Like "*" + ArrDg(5, s) + "*" Then
                'Es wurde eine RangSpalte gefunden
                'LeerZellen unterhalb einer Zelle mit RangZahl sollen gefüllt werden
                For i = 6 To UBound(ArrDg, 1) - 2
                    If ArrDg(i, s) <> "" And ArrDg(i + 1, s) = "" Then
                        ArrDg(i + 1, s) = CStr(ArrDg(i, s))
                    End If
                Next
            End If
        Next
        'showArray2D ArrDg
End Sub

Function T4_Get_AllClubNameRowNrs_M2V8()
    'Called from    FillArrDg
    'Team only      nur für Mannschaftswettkämpfe (DgGroupName = "M2V8")
    'Action         ermittelt zClubs = "|6|13|20|27|34|41|48|55|62|69"

    'Exit
        If Not ArrDg(2, 1) Like "M2V8" Then Exit Function
    'Vorbereitung
        Dim zClubs$, i%, s1%, z1%
        Dim r1 As Range, r2 As Range, r3 As Range, r4 As Range
        With Sheets("T4"): z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39))
    'Vereinszeilen ermitteln (oberhalb m/w-Färbung der TeamMembers-Zellen)
        For i = 6 To UBound(ArrDg, 1) - 1
            'In Spalte 4 und 17 stehen die TeamMemberNames, farbunterlegt
            Set r1 = .Cells(z1 - 1 + i, s1 + 3)     'Zelle mit  TeamMemberName w
            Set r2 = .Cells(z1 - 1 + i, s1 + 16)    'Zelle mit  TeamMemberName m
            Set r3 = .Cells(z1 - 2 + i, s1 + 3)     'Zelle über TeamMemberName w
            Set r4 = .Cells(z1 - 2 + i, s1 + 16)    'Zelle über TeamMemberName m
            If r1.Interior.Color = 14083324 Or r2.Interior.Color = 15652797 Then
                If r3.Interior.ColorIndex = 2 And r4.Interior.ColorIndex = 2 Then
                    'Zeile i-1 = weiße Zeile oberhalb gefärbter Zeile
                    
                    If i = 29 Then Stop
                    
                    zClubs = zClubs + "|" + CStr(i - 1)
                End If
            End If
        Next
    'Finals
        T4_Get_AllClubNameRowNrs_M2V8 = zClubs
        End With
End Function

Sub T4_Add_ClubNameAndRank_ToMembersInArrDg_M2V8()
    'Called from    FillArrDg
    'Team only      nur für Mannschaftswettkämpfe (DgGroupName = "M2V8")
    'Action         ermittelt cMembers = "7|8|9|10|11@14|15|16|17|18@...@70|71|72|73|74"

    'Exit
        If Not ArrDg(2, 1) Like "M2V8" Then Exit Sub
    'Vorbereitung
        Dim ClubName$, zClubs$
        Dim i%, j%, z1%, z2%, zClub%, cMembers%, Arr_zClubs() As String
        With Sheets("T4"): z1 = CInt(ArrC(37)): z2 = CInt(ArrC(38))
    'xxx
        zClubs = ArrC(45)                       '= "|6|13|20|27|34|41|48|55|62|69"
        Arr_zClubs = Split(zClubs, "|")
        For i = 1 To UBound(Arr_zClubs)
            'ZeilenNr der i. ClubZeile
                zClub = CInt(Arr_zClubs(i))                         '6
            'cMembers (Anzahl der Competitors des i. Clubs)
                If i < UBound(Arr_zClubs) Then
                    cMembers = CInt(Arr_zClubs(i + 1)) - zClub - 2  '13-6-2=5
                Else
                    cMembers = z2 - z1 - zClub
                End If
            'Damen ClubName
                ClubName = ArrDg(zClub, 3)
                If ClubName <> "" Then
                    'ClubName in die MemberZeilen schreiben (ArrDg)
                    For j = 0 To cMembers: ArrDg(zClub + j, 3) = ClubName: Next
                End If
            'Damen Rang
                For j = 1 To cMembers: ArrDg(zClub + j, 2) = ArrDg(zClub, 2): Next
            'Herren ClubName
                ClubName = ArrDg(zClub, 16)
                If ClubName <> "" Then
                    'ClubName in die MemberZeilen schreiben (ArrDg)
                    For j = 0 To cMembers: ArrDg(zClub + j, 16) = ClubName: Next
                End If
            'Herren Rang
                For j = 1 To cMembers: ArrDg(zClub + j, 15) = ArrDg(zClub, 15): Next
        Next
    'Finals
        End With

End Sub

Sub CreateRoundRect_Cell_TEST()
    CreateRoundRect_Cell "T4", "hallo", 148, 372
End Sub

Sub CreateRoundRect_Cell(NameOfSheet$, NameOfShape$, z%, s%)
    Dim ws As Worksheet, r As Range, shp As shape
    Set ws = Sheets(NameOfSheet)
    Set r = ws.Cells(z, s)
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, r.Left, r.Top, r.Width + 1, r.Height + 1)
    With shp: .NAME = NameOfShape: .Fill.Visible = msoFalse
    With .Line: .Visible = msoTrue: .ForeColor.RGB = RGB(255, 0, 0): .Weight = 3: End With: End With
End Sub

Sub T4_RoRe(Visible_0_OR_1%, Optional zDg% = 0, Optional sDg% = 0)
    'Called from    T4_ActionsOnDgClickNone
    'RoRe           RoundRect; zur Markierung eines CompetitorNamens
    
    'Vorbereitung
        Dim s1%, sSh%, z1%, zSh%, r As Range
        DoArrc
        If ArrC(37) = "-" Then Hide_Shape "T4", "RoRe": Exit Sub
        z1 = CInt(ArrC(37))         'T4 Dg ZeilenNr   of DesignBorder-Top
        s1 = CInt(ArrC(39))         'T4 Dg SpaltenNr  of DesignBorder-Left
        zSh = z1 - 1 + zDg: sSh = s1 - 1 + sDg
        
        'Set shp = ws.Shapes("RoRe")
        If ShapeExists("T4", "RoRe") Then
            With ThisWorkbook.Sheets("T4").Shapes("RoRe")
        'Action
            If Visible_0_OR_1 = 0 Then Hide_Shape "T4", "RoRe" Else Show_Shape "T4", "RoRe"
            If zDg > 0 And sDg > 0 Then
                Set r = Sheets("T4").Cells(zSh, sSh)
                .Left = r.Left: .Top = r.Top: .Width = r.Width + 1: .Height = 17 - 1
            End If
            End With
        End If
End Sub

Sub T4_Scroll_WhileSelectingEachDgCompetitorName(zDg%, sDg%, C1%)
    'Called from    T4_Update_T8SomeDgData_ActualDgNames_viaT4ActivateEachName
    '(zDg, sDg)     Cell to be selected; Zelle soll gut zu sehen sein
    
    'Vorbereitung
        Dim Li%, Re%, Ob%, Un%, s1%, s2%, z1%, z2%, sSh%, zSh%, sDgW%, sDgM%
        z1 = CInt(ArrC(37)): z2 = CInt(ArrC(38)): s1 = CInt(ArrC(39)): s2 = CInt(ArrC(40))
        zSh = z1 - 1 + zDg: sSh = s1 - 1 + sDg
        With Sheets("T4") ': Screen 0
    'VisibleRange
        Li = Get_FirstVisibleColumnLeft:    Re = Get_LastVisibleColumnRight
        Ob = Get_FirstVisibleRowTop:        Un = Get_LastVisibleRowBottom
    'No scroll?
        If Re > s2 And Un > z2 Then
            Exit Sub
        End If
    'SpaltenNr NamenWeiblich, NamenMännlich
        'showArray2D ArrDg
        If IsNumeric(ArrC(52)) Then sDgW = ArrC(52)
        If IsNumeric(ArrC(53)) Then sDgM = ArrC(53)
    'Scroll
        If sDg = sDgW Then 'Damen
            If zSh > Un - 4 Then
                T4_Scroll_ToCell zScroll + 1, sSh - 3
            End If
        End If
        If sDg = sDgM Then 'Herren
            'Obere Zeilen
                If zDg < 13 Then
                    'Fixiere ScrollLeft für alle Select
                        If C1 = 0 Then
                            C1 = s1 + sDgM + 2
                            T4_Scroll_ToSeeColumnY_OnTheRight C1
                        End If
                    'Zeige Dg bei SelectBeginn oben
                        T4_Scroll_ToCell z1 - 1, 0
                End If
            'Untere Zeilen
                If zSh > Un - 4 Then
                    T4_Scroll_ToCell zScroll + 1, 0 'Herren
                End If
        End If
    'Finals
        End With
End Sub

Sub T4_Scroll_ToSeeColumnY_OnTheRight(y%)
    Dim Li%, LiNew%, Ob%, Re%
    Li = Get_FirstVisibleColumnLeft
    Ob = Get_FirstVisibleRowTop
    Re = Get_LastVisibleColumnRight
    LiNew = y + Li - Re
    
    T4_Scroll_ToCell Ob, LiNew
    
End Sub

Function zScroll() As Integer
    zScroll = ActiveWindow.ScrollRow
End Function

Function sScroll() As Integer
    sScroll = ActiveWindow.ScrollColumn
End Function

Sub T4_Scroll_ToCell(zSh%, sSh%)
    EE 0
    If zSh = 0 Then zSh = ActiveWindow.ScrollRow        'topmost  sichtbare Zeile  bleibt
    If sSh = 0 Then sSh = ActiveWindow.ScrollColumn     'leftmost sichtbare Spalte bleibt
    Application.GoTo Reference:=Worksheets("T4").Cells(zSh, sSh), Scroll:=True
    EE 1
End Sub

Function T4_RoRe_IsVisible() As Boolean
    On Error Resume Next
    If Sheets("T4").Shapes("RoRe").Visible = msoTrue Then T4_RoRe_IsVisible = True
    On Error GoTo 0
End Function

Sub Scroll_ToSeeCell_z_s(z&, s&, zFromTop&, sFromLeft&)
    'Called from    T4_Update_T8SomeDgData_ActualDgNames_viaT4ActivateEachName
    Dim ss&, zz&
    EE 0: zz = z - zFromTop: ss = s - sFromLeft
    If zz < 1 Then Exit Sub
    If ss < 1 Then Exit Sub
    Application.GoTo Reference:=Worksheets("T4").Cells(zz, ss), Scroll:=True
    EE 1
End Sub

Sub ScrollRightSlow(CountOfCellsToGo%)
    Dim c%, i%, s%, z%
    c = CountOfCellsToGo: EE 0
    z = ActiveWindow.ScrollRow
    s = ActiveWindow.ScrollColumn
    For i = 1 To c
        Application.GoTo Reference:=Worksheets("T4").Cells(z, s + i), Scroll:=True
        Sleep 200: DoEvents
    Next: EE 1
End Sub

Sub Scroll_ToSeeRowNrX_OnTheTOP_Step(x%)
    Dim c%, i%, Li%, Re%, s%, z%
    Ob = Get_FirstVisibleRowTop
    Un = Get_LastVisibleRowBottom
    
    If x + 2 < Re And x - 2 > Li Then Exit Sub 'Y gut sichtbar
    'Spalte Y ist nicht sichtbar
    z = ActiveWindow.ScrollRow
    s = ActiveWindow.ScrollColumn
    EE 0
    If x > Re Then
        For i = 1 To x - Re + 2
            Application.GoTo Reference:=Worksheets("T4").Cells(z, s + i), Scroll:=True
            Sleep 100: DoEvents
        Next: EE 1
    End If
    If x < Li Then
        For i = 1 To Li - x + 2
            Application.GoTo Reference:=Worksheets("T4").Cells(z, s - i), Scroll:=True
            Sleep 100: DoEvents
        Next
    End If
    EE 1
End Sub

Sub Scroll_ToSeeColumnNrY_OnTheLeft_Step(y%)
    Dim c%, i%, Li%, Re%, s%, z%
    Li = Get_FirstVisibleColumnLeft
    Re = Get_LastVisibleColumnRight
    If y + 2 < Re And y - 2 > Li Then Exit Sub 'Y gut sichtbar
    'Spalte Y ist nicht sichtbar
    z = ActiveWindow.ScrollRow
    s = ActiveWindow.ScrollColumn
    EE 0
    If y > Re Then
        For i = 1 To y - Re + 2
            Application.GoTo Reference:=Worksheets("T4").Cells(z, s + i), Scroll:=True
            Sleep 100: DoEvents
        Next: EE 1
    End If
    If y < Li Then
        For i = 1 To Li - y + 2
            Application.GoTo Reference:=Worksheets("T4").Cells(z, s - i), Scroll:=True
            Sleep 100: DoEvents
        Next
    End If
    EE 1
End Sub

Function Get_FirstVisibleColumnLeft() As Integer
    Get_FirstVisibleColumnLeft = ActiveWindow.VisibleRange.Column
End Function

Function Get_LastVisibleColumnRight() As Integer
    Get_LastVisibleColumnRight = ActiveWindow.VisibleRange.Column + ActiveWindow.VisibleRange.Columns.count - 2
End Function

Function Get_LastVisibleRowBottom() As Integer
    Get_LastVisibleRowBottom = ActiveWindow.VisibleRange.Row + ActiveWindow.VisibleRange.Rows.count - 2
End Function

Function Get_FirstVisibleRowTop() As Integer
    Get_FirstVisibleRowTop = ActiveWindow.VisibleRange.Row
End Function

Sub LeuteFolder()
    Dim s$
    s = Get_AllSubfolderPaths_LikeMyString_OneLevel(ArrC(4), "*(#*")
    's = Get_AllSubfolderPaths_LikeMyString_OneLevel(ArrC(4), "*# *")
    show CStr(anzAinB(vbCrLf, s) + 1) + vbCrLf + vbCrLf + s
End Sub


