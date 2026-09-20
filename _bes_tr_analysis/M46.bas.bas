Attribute VB_Name = "M46"
Option Explicit 'M46

Sub zzz_M46()

    'showProcs "next"
        
    EE 1: Beep
End Sub

Sub xxxT4_Add_EventLink_ToPersonFolders_TEST_NEW()
    'Called from    [T6-CellButton ' Add EventLink to PersonFolders']
    'EventFolder    = EventOrdner; Ordner im Folder "Events"
    'EventLink      = Link-File, der zu einem EventFolder springt
    'PersonFolder   = PersonenOrdner; Ordner im Folder "Leute"
    'PersonLink     = Link-File, der zu einem PersonFolder springt
    'Status         In jedem EventFolder befinden sich bereits PersonLinks
    '               (zu den PersonFolder der am Event beteiligten Leute);
    '               ein PersonLink in einem EventFolder wurde automatisch oder von Hand hinzugefügt
    'Action         Jeder solcher PersonFolder soll einen EventLink erhalten;
    '               dazu werden alle EventFolders nach existierenden PersonLinks durchsucht;
    '               der zugehörige PersonFolder eines solchen PersonLinks erhält den EventLink
    
    'Vorbereitung
        Dim A$, Club$, E1$, E2$, L$, N$, NameOfEventFolder$, Nation$, NewLnkName$, Nn$
        Dim OldLnkName$, p$, PathOfEventFolder$, PathOfPersonalFolder$, Report$, v$, Vn$
        Dim C1%, C2%, c8%, c9%, i%, zT5%, zT6%, ArrEventFolderPaths() As String, E() As String
        v = vbCrLf: Ticks1: LogBuch "T4_Add_EventLink_ToPersonFolders was started"
    'zT6            T6-ZeilenNr, die den Cell-Button enthält (für T6_DoDoneRemarks-Einträge)
        zT6 = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 2, 4, " Add EventLink to PersonFolders")
    'MyCountDown Init
        MyCountDown_Init "|T6|" + CStr(zT6) + "|19|140|" '= |SheetName|z|s|StartNr| (CountDownShowCell)
    'Load ArrEventFolderPaths
        Load_ArrPathsOfAllFolders_Events ArrEventFolderPaths
        'showArray ArrEventFolderPaths:Stop
        c9 = UBound(ArrEventFolderPaths)
        MyCountDown_MainStepsAllowed 20: MyCountDown_SubStepsMax c9
    'Add EventLink for each EventOrdner to his PersonFolders
        For i = 1 To c9
            'Schleife über alle EventFolder
            MyCountDown_OneMoreSubStep
            p = ArrEventFolderPaths(i)                              'p = one EventFolderPath
            T4_Create_EventLink_DateSquare_OfOneEventfolder_ToPersonFolders (p)
            DoEvents
        Next
    'T6_DoDoneRemarks
        'FillArrC 24, CStr(c8) 'CountOfChanges
        T6_DoDoneRemarks
    'Finals
        If Report <> "" Then
            Report = "PROCEDURE:   T4_Add_EventLink_ToPersonFolders(zT6%)" + v _
              + "ERROR:       Format 'Nachname, Vorname' nicht erkennbar" + v + v + Report
            show Report
        End If
        End With: Beep
        LogBuch "T4_Add_EventLink_ToPersonFolders has ended"
End Sub

Sub T4_Create_EventLink_DateSquare_OfOneEventfolder_ToPersonFolders(PathOfOneEventFolder$)
    'Called from    ClickButton_CreateJpg_InActualDg
    'EventFolder    = EventOrdner; Ordner im Folder "Events"
    'EventLink      = Link-File, der zu einem EventFolder springt = "19800621 ! DSJEMSy Neustadt"
    'PersonFolder   = PersonenOrdner; Ordner im Folder "Leute"
    'PersonLink     = Link-File, der zu einem PersonFolder springt
    'Status         In jedem EventFolder befinden sich bereits PersonLinks
    '               (zu den PersonFolder der am Event beteiligten Leute);
    '               ein PersonLink in einem EventFolder wurde automatisch oder von Hand hinzugefügt
    'Action         Jeder solcher PersonFolder soll ein DateSquare und einen EventLink erhalten;
    '               dazu wird der aktuelle EventFolder nach existierenden PersonLinks durchsucht;
    '               der zugehörige PersonFolder eines solchen PersonLinks erhält die beiden Files
    
    'Vorbereitung
        Dim A$, Club$, E1$, E2$, L$, N$, NameOfEventFolder$, Nation$, NewLnkName$, Nn$
        Dim OldLnkName$, pDate8File$, pE$, PathOfPersonalFolder$, Report$, v$, Vn$
        Dim C1%, C2%, c8%, i%, zT5%, E() As String
        pE = PathOfOneEventFolder: v = vbCrLf
        LogBuch "T4_Create_EventLink_DateSquare_OfOneEventfolder_ToPersonFolders was started: " + getNameOfPath(pE)
    'A = alle Links innerhalb des aktuellen EventOrdners
        A = Get_AllFilePaths_LikeMyStringInFileName_OfOneFolder(pE, "[A-Z]*.lnk") + v
        A = Delete_EmptyRowsInString(A)
        E = Split(v + A, v): With Sheets("T5")
        c8 = UBound(E)
    'xxx
        For i = 1 To c8
            'Schleife über alle PersonLinks innerhalb des aktuellen EventFolders
            E1 = E(i) 'path of one PersonLink inside EventFolder
            'NameOfEventfolder
                NameOfEventFolder = NameOfPath(pE)             '1986 DJM '1984-09-15 LK D-GB
            'Name of PersonLink inside EventFolder
                N = NameOfPath(E1)                  'May, Lisa (DK).lnk
                N = Replace(N, ".lnk", "")          'May, Lisa (DK)         'May (23), Lisa (3)
                'man kommt so an mehr Personen (in deren PersonFolder ein
                'EventLink sinnvoll wäre) als die DgCompetitors liefern würden
            'pure Name of PersonLink
                C1 = InStr(1, N, " ("): C2 = InStr(1, N, ")")
                If C1 * C2 > 0 Then N = Left(N, C1 - 1) + Mid(N, C2 + 1)    'May, Lisa (3)
                C1 = InStr(1, N, " ("): C2 = InStr(1, N, ")")
                If C1 * C2 > 0 Then N = Left(N, C1 - 1) + Mid(N, C2 + 1)    'May, Lisa
                C1 = InStr(1, N, ",")
            'pure Name not complete
                If C1 = 0 Then
                    'REPORT
                    Report = Report + "NameOfEventfolder: " + NameOfEventFolder + " (FolderToJumpTo)" _
                    + v + "   Im EventFolder liegt ein Link '" + N _
                          + "'; Format 'Nachname, Vorname' nicht erkennbar" _
                    + v + "   dieser Link sieht vor, zum 'Leute'-Ordner '" _
                          + NameOfPath(Get_LinkTargetPath(E1)) + "' zu springen" + v
                    GoTo Jump
                End If
            'pure Name is OK
                Nn = Left(N, C1 - 1)                 'May
                Vn = Mid(N, C1 + 2)                  'Lisa
            'Name in T5 bekannt?
                zT5 = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, Vn + " " + Nn)
                If zT5 > 7 Then
                    PathOfPersonalFolder = ArrC(4) + "\" + .Cells(zT5, 10) 'FolderInWhitchTheLinkShouldBeCreated
                    L = T4_Get_LinkNameToDisplay_ofEvFoldName(NameOfEventFolder)   'LinkNameToDisplay
                    ' = "19840000 ..LK D-GB Bonn" (also nicht "1984 LK" etc.)
                    'Write EventLink to PersonFolder
                        'OpenFolder PathOfPersonalFolder: Stop 'pE = PathOfOneEventFolder
                        Create_OneLinkFile PathOfPersonalFolder, pE, L 'FolderInWhitchTheLinkShouldBeCreated, FolderToJumpTo, LinkNameToDisplay
                        'Stop
                    'Write OneDateSquare into PersonFolder
                        pDate8File = PathOfPersonalFolder + "\" + T9_Get_Date8FromEventFolderName(NameOfEventFolder)
                        'OpenFolder PathOfPersonalFolder: Stop
                        Create_OneDateSquare_Give_Date8File pDate8File
                        'Stop
                    'Change actual EventFolder-PersonLink to "May, Lisa (TV Bonn)"
                        Club = Sheets("T5").Cells(zT5, 7)
                        Nation = Sheets("T5").Cells(zT5, 9)
                        If Nation = "D" Then
                            If Club = "" Then
                                NewLnkName = Nn + ", " + Vn + ".lnk"
                            Else
                                NewLnkName = Nn + ", " + Vn + " (" + Club + ")" + ".lnk"
                            End If
                        Else
                            NewLnkName = Nn + ", " + Vn + " (" + Nation + ")" + ".lnk"
                        End If
                        OldLnkName = NameOfPath(E1)
                        'e1 = path of actual PersonFolderLink inside EventFolder
                        If OldLnkName <> NewLnkName Then
                            'OpenFolder pe: Stop
                            E2 = Replace(E1, OldLnkName, NewLnkName)
                            RenameFile E1, E2
                        End If
                Else
                    Report = Report + NameOfEventFolder + "   -->   '" + NameOfPath(E1) _
                          + "' existiert als Link, aber Name nicht in T5" + v
                End If
            DoEvents
Jump:
        Next
    'Finals
        If Report <> "" Then
            Report = "PROCEDURE:   T4_Create_EventLink_DateSquare_OfOneEventfolder_ToPersonFolders" + v _
              + "             OneEventfolder: " + getNameOfPath(pE) + v _
              + "ERROR:       Format 'Nachname, Vorname' nicht erkennbar" + v + v + Report
            show Report
        End If
        End With: Beep
        LogBuch "T4_Create_EventLink_DateSquare_OfOneEventfolder_ToPersonFolders has ended"
End Sub

Sub T4_Add_EventLink_ToPersonFolders()
    'Called from    [T6-CellButton ' Add EventLink to PersonFolders']
    'EventFolder    = EventOrdner; Ordner im Folder "Events"
    'EventLink      = Link-File, der zu einem EventFolder springt
    'PersonFolder   = PersonenOrdner; Ordner im Folder "Leute"
    'PersonLink     = Link-File, der zu einem PersonFolder springt
    'Status         In jedem EventFolder befinden sich bereits PersonLinks
    '               (zu den PersonFolder der am Event beteiligten Leute);
    '               ein PersonLink in einem EventFolder wurde automatisch oder von Hand hinzugefügt
    'Action         Jeder solcher PersonFolder soll einen EventLink erhalten;
    '               dazu werden alle EventFolders nach existierenden PersonLinks durchsucht;
    '               der zugehörige PersonFolder eines solchen PersonLinks erhält den EventLink
    'Steps          (1) Load all EventFolderPaths [ArrEventFolderPaths]
    '               (2) Load all LinkPaths of allEventFolders [E]
    '               (3) Write LinkOfEventFolder to each EventPersonFolder
    'Vorbereitung
        Dim A$, Club$, E1$, E2$, L$, N$, NameOfEventFolder$, Nation$, NewLnkName$, Nn$
        Dim OldLnkName$, p$, PathOfEventFolder$, PathOfPersonfolder$, r$, v$, Vn$
        Dim C1%, C2%, c8%, c9%, i%, zT5%, zT6%, ArrEventFolderPaths() As String, E() As String
        v = vbCrLf: Ticks1: LogBuch "T4_Add_EventLink_ToPersonFolders was started"
    'zT6            T6-ZeilenNr, die den Cell-Button enthält (für T6_DoDoneRemarks-Einträge)
        zT6 = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 2, 4, " Add EventLink to PersonFolders")
    'MyCountDown Init
        MyCountDown_Init "|T6|" + CStr(zT6) + "|19|140|" '= |SheetName|z|s|StartNr| (CountDownShowCell)
    'Load ArrEventFolderPaths
        Load_ArrPathsOfAllFolders_Events ArrEventFolderPaths
        'showArray ArrEventFolderPaths:Stop
        c9 = UBound(ArrEventFolderPaths)
        MyCountDown_MainStepsAllowed 20: MyCountDown_SubStepsMax c9
    'A = alle Links aller EventOrdner
        For i = 1 To c9
            'Schleife über alle EventFolder
            MyCountDown_OneMoreSubStep
            p = ArrEventFolderPaths(i)                              'p = one EventFolderPath
            'Get all LinkFiles inside one EventFolderPath
                A = A + Get_AllFilePaths_LikeMyStringInFileName_OfOneFolder(p, "[A-Z]*.lnk") + v
                DoEvents
        Next
        A = Delete_EmptyRowsInString(A)
        E = Split(v + A, v): With Sheets("T5")
        c8 = UBound(E): FillArrC 24, CStr(c8) 'CountOfChanges
        MyCountDown_MainStepsAllowed 120: MyCountDown_SubStepsMax c8
    'xxx
        For i = 1 To c8
            'Schleife über alle Pfade von PersonLinks aller EventFolder
            MyCountDown_OneMoreSubStep
            E1 = E(i) 'one path of PersonLink inside EventFolder
            'PathOfEventfolder
                PathOfEventFolder = Get_ParentFolderPath_OfFolderPath(E1) 'FolderToJumpTo
            'NameOfEventfolder
                NameOfEventFolder = NameOfPath(PathOfEventFolder)             '1986 DJM '1984-09-15 LK D-GB
            'Name of PersonLink inside EventFolder
                N = NameOfPath(E1)                  'May, Lisa (DK).lnk
                N = Replace(N, ".lnk", "")          'May, Lisa (DK)         'May (23), Lisa (3)
                'man kommt so an mehr Personen (in deren PersonFolder ein
                'EventLink sinnvoll wäre) als die DgCompetitors liefern würden
            'pure Name of PersonLink
                C1 = InStr(1, N, " ("): C2 = InStr(1, N, ")")
                If C1 * C2 > 0 Then N = Left(N, C1 - 1) + Mid(N, C2 + 1)    'May, Lisa (3)
                C1 = InStr(1, N, " ("): C2 = InStr(1, N, ")")
                If C1 * C2 > 0 Then N = Left(N, C1 - 1) + Mid(N, C2 + 1)    'May, Lisa
                C1 = InStr(1, N, ",")
            'pure Name not complete
                If C1 = 0 Then
                    'REPORT
                    r = r + "NameOfEventfolder: " + NameOfEventFolder + " (FolderToJumpTo)" _
                    + v + "   Im EventFolder liegt ein Link '" + N _
                          + "'; Format 'Nachname, Vorname' nicht erkennbar" _
                    + v + "   dieser Link sieht vor, zum 'Leute'-Ordner '" _
                          + NameOfPath(Get_LinkTargetPath(E1)) + "' zu springen" + v
                    GoTo Jump
                End If
            'pure Name is OK
                Nn = Left(N, C1 - 1)                 'May
                Vn = Mid(N, C1 + 2)                  'Lisa
            'Name in T5 bekannt?
                zT5 = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, Vn + " " + Nn)
                If zT5 > 7 Then
                    PathOfPersonfolder = ArrC(4) + "\" + .Cells(zT5, 10) 'FolderInWhitchTheLinkShouldBeCreated
                    L = T4_Get_LinkNameToDisplay_ofEvFoldName(NameOfEventFolder)   'LinkNameToDisplay
                    ' = "19840000 ..LK D-GB Bonn" (also nicht "1984 LK" etc.)
                    'Write EventLink to PersonFolder
                    
                    
                        'OpenFolder PathOfPersonFolder: Stop
                        
                        
                        Create_OneLinkFile PathOfPersonfolder, PathOfEventFolder, L
                        'Create.. FolderInWhitchTheLinkShouldBeCreated, FolderToJumpTo, LinkNameToDisplay
                        'OpenFolder PathOfPersonFolder: Stop
                        
                    'Change actual EventFolder-PersonLink to "May, Lisa (TV Bonn)"
                        zT5 = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, Vn + " " + Nn)
                        Club = Sheets("T5").Cells(zT5, 7)
                        Nation = Sheets("T5").Cells(zT5, 9)
                        If Nation = "D" Then
                            If Club = "" Then
                                NewLnkName = Nn + ", " + Vn + ".lnk"
                            Else
                                NewLnkName = Nn + ", " + Vn + " (" + Club + ")" + ".lnk"
                            End If
                        Else
                            NewLnkName = Nn + ", " + Vn + " (" + Nation + ")" + ".lnk"
                        End If
                        OldLnkName = NameOfPath(E1)
                        'e1 = path of actual PersonFolderLink inside EventFolder
                        If OldLnkName <> NewLnkName Then
                            'OpenFolder PathOfEventFolder: Stop
                            E2 = Replace(E1, OldLnkName, NewLnkName)
                            RenameFile E1, E2
                        End If
                Else
                    r = r + NameOfEventFolder + "   -->   '" + NameOfPath(E1) _
                          + "' existiert als Link, aber Name nicht in T5" + v
                End If
            DoEvents
Jump:
        Next
    'T6_DoDoneRemarks
        T6_DoDoneRemarks
    'Finals
        If r <> "" Then
            r = "PROCEDURE:   T4_Add_EventLink_ToPersonFolders(zT6%)" + v _
              + "ERROR:       Format 'Nachname, Vorname' nicht erkennbar" + v + v + r
            show r
        End If
        End With: Beep
        LogBuch "T4_Add_EventLink_ToPersonFolders has ended"
End Sub

Sub T4_Copy_°FilesOfOneEvent_ToPersonFolders(PathOfEventFolder$)
    'Called from    ClickButton_CreateJpg_InActualDg
    '°File          jpg, das in den jew. PersonenOrdner der beteiligten Personen
    '               z. B. Plakat des Events
    'Action         xxx
    
    'Vorbereitung
        Dim Fs$, Id$, Ls$, p1$, p2$, p3$, PF$, PL$, v$, i%, j%, L() As String, F() As String
        v = vbCrLf: DoArrc
    'All °Files of actual EventFolder
        Fs = Get_AllFileNames_Like_OfOneFolder(PathOfEventFolder, "*°*")
        If Fs = "" Then Exit Sub
    'All Links of actual EventFolder
        Ls = Get_AllFileNames_Like_OfOneFolder(PathOfEventFolder, "*.lnk")
    'Action
        L = Split(Ls, v)
        F = Split(Fs, v)
        For i = 0 To UBound(L)
            'Schleife über alle Links of actual EventFolder
            For j = 0 To UBound(F)
                'Schleife über alle °Files of one EventFolder
                p1 = PathOfEventFolder + "\" + F(j)     'Path of File to copy
                    PL = PathOfEventFolder + "\" + L(i) 'Path of Link in EventFolder
                    
                    
                    PF = Get_LinkTargetPath(PL)         'Path of PersonFolder
                    
                    
                    If Not FolderExists(PF) Then
                        Stop
                        'Link ist veraltet; PersonFolder inzwischen umbenannt
                        PF = T5_Get_PathOfPersonfolder_Give_NnVn(L(i))
                        If PF = "" Then                 'Name nicht in T5
                            DeleteFile PL               'Link in EventFolder
                        Else
                            If Not FolderExists(PF) Then Stop
                            DeleteFile PL               'old Link in EventFolder
                            Create_OneLinkFile PathOfEventFolder, PF, getNameOfPath(PF)
                        End If
                    End If
                If PF <> "" Then p2 = PF + "\" + F(j): T4_Copy_p1p2_OnlyIfNew p1, p2
            Next
        Next
        Beep
End Sub

Sub T4_Copy_p1p2_OnlyIfNew(p1$, p2$)
    'Called from    T4_Copy_°FilesOfOneEvent_ToPersonFolders
    'p1             File to copy
    'p2             PathOfFileInPersonFolder (p1-Copy)
    
    'Vorbereitung
        Dim Id$, p3$, PF$
    'File mit gleicher Id ####-## bereits im PersonFolder?
        PF = Get_PathOfParentFolder(p2) 'PersonFolder
        Id = Get_IdFromPath(p1)
        p3 = PF + "\" + Get_AllFileNames_Like_OfOneFolder(PF, "*" + Id + "*")
        If Replace(p3, "\Leute\", "") Like "*\Leute\*" Then Stop 'Id existiert 2x
    'Copy
        If p3 = "" Then
            CopyFile p1, p2
        Else
            If p2 <> p3 Then DeleteFile p3: CopyFile p1, p2
        End If
End Sub

Function T4_Get_AllVnNnInDgsOfOneEvent_FromT4SomeDgData(PathOrNameOfEventFolder$) As String
    'Called from    xxx
    'AllVnNn        = "|...|Lea May|Kai Müller|...|"
    
    'Vorbereitung
        Dim NE$, RowNr$, RowNrs$, VnNn$, i%, A() As String, B() As String, r() As String
        NE = PathOrNameOfEventFolder: If NE Like "*\*" Then NE = NameOfPath(NE) 'NameOfEventFolder
        RowNrs = "|"
    'RowNrs T4SomeDgData
        T4_Load_T4SomeDgData_OneCol 6, A 'A() = 1D-Array-1Based = AllEventFolderNames
        T4_Load_T4SomeDgData_OneCol 8, B 'B() = 1D-Array-1Based = AllDgNames
        For i = 1 To UBound(A)
            If A(i) = NE Then RowNr = CStr(i) + "|": RowNrs = RowNrs + RowNr
            '       = all RowNrs holding NameOfEventFolder (ggf. für Einzel, Synchron, Mannschaft)
        Next
        'show RowNrs: Stop 'Sheet-RowNrs = "|107|108|"; T4(107,10) = NE; T4(108,10) = NE
    'VnNn
        r = Split(RowNrs, "|")
        For i = 1 To UBound(r) - 1
            VnNn = VnNn + Mid(B(r(i)), 12)
        Next
        'show VnNn: Stop 'VnNn = "|#|Lea May|7|4|Einz|1|...|#|Kai Abc|8|4|Team|2|"
        r = Split(VnNn, "|"): VnNn = "|"
        For i = 2 To UBound(r)
                If r(i - 1) = "#" Then
                    If Not VnNn Like "*|" + r(i) + "|*" Then VnNn = VnNn + r(i) + "|"
                End If
        Next
        'show "VnNn" + vbCrLf + vbCrLf + VnNn: Stop 'VnNn = "|Lea May|...|Kai Abc|"
    'Finals
        T4_Get_AllVnNnInDgsOfOneEvent_FromT4SomeDgData = VnNn
End Function

Function Get_VnNn_From_NameOfPersFold(NameOfPersonfolder$) As String
    'Called from    Copy_BracFiles_OfOneFolder_ToPersonFolders
    'PersFol        = "Mai, Lea (Bonn)", "Mai, Lea"
    
    Dim F$, VnNn$, c%
    F = NameOfPersonfolder
    c = InStr(1, F, " ("): If c > 0 Then F = Left(F, c - 1) '"Mai, Lea"
    c = InStr(1, F, ", ")
    If c = 0 Then Exit Function
    VnNn = Mid(F, c + 2) + " " + Left(F, c - 1)
    Get_VnNn_From_NameOfPersFold = VnNn
End Function

Function T4_Get_Id_Folders_OfOneNameInBracketsJpg(OneFileNameWithId$)
    'Called from    Copy_BracFiles_OfOneFolder_ToPersonFolders
    '..Brackets..   NameOfJpg enthält VnNn in ()
    'Id             = "p1234-05" oder "v0123-00"
    'Folders        = PersonFolder of Pia May, Lea Mau, ...
    'OneNamedJpg    = 19800621 DM Bonn - a21 (Pia May) (Lea Mau) p1234-05 ub.jpg
    
    'Vorbereitung
        If Not OneFileNameWithId Like "*(* *[A-ZÄÖÜ]*" Then Exit Function
        Dim F$, Fo$, Id$, N$, NF$, RET$, v$, c%, i%, zT5%, A() As String
        v = vbCrLf: DoArrc: F = OneFileNameWithId: c = InStr(1, F, "(")
        F = Mid(F, c)              'F = "(Pia May) (Lea Mau) v1234-05 nn.jpg"
    'Id
'        Stop 'Allg. prüfen
        'Id ermitteln
            Id = Get_PidOrVid_FromFileName(F)
        'BracketNames ermitteln
            c = InStr(1, F, Id)        'f reduzieren
            'Stop
            
            F = Left(F, c - 2)         'f = "(Pia May) (Lea Mau)"
    '(Vn Nn)
            A = Split(F, "(")
            For i = 1 To UBound(A)
                N = A(i)            '"Pia May) "
                c = InStr(1, N, ")"): N = Left(N, c - 1)
                'N nur in RET aufnehmen, wenn N ein T5-Name ist
                If N Like "* *" Then
                    zT5 = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, N)
                    If zT5 > 0 Then
                        'FileName f enthält den T5-Namen N
                        Fo = Sheets("T5").Cells(zT5, 10)        'fo = NameOfPersonFolder
                        RET = RET + Fo + "|"
                    End If
                End If
            Next
        If RET <> "" Then
            RET = "|" + Id + "|" + RET
            T4_Get_Id_Folders_OfOneNameInBracketsJpg = RET
        End If
End Function

Sub T4_CellBtn_ub(z2%, s2%)
    'Called from    [UserClick on "LU"] Worksheet_SelectionChange[T4]
    'ub             Dg-EckZelle rechts unten
    'use?           Get_RowNr_HoldingMyTextWholeInColumnXSearchUp
    '               Get_ColumnNr_HoldingMyTextWholeInRowXSearchLeft
    '               Get_ColumnNr_HoldingMyTextWholeInRowXSearchRight
    
    'Vorbereitung
        Dim DgTitle$, i%, s1%, z1%, zSome%, A()
        With Sheets("T4"): EE 0: ReDim A(1 To 4)
        .Cells(z2, s2).Interior.ColorIndex = 3
    'RU-->RO-->LO-->Title
    'Get z1     'Nächstes "RO" in T4-Salte s2 von z2 aus nach oben suchen
        z1 = Get_RowNr_HoldingMyTextWholeInColumnXSearchUp("T4", s2, z2, "RO"): If z1 = 0 Then Stop
    'Get s1     'Nächstes "LO" in T4-Zeile z1 von s2 aus nach links suchen
        s1 = Get_ColumnNr_HoldingMyTextWholeInRowXSearchLeft("T4", z1, s2, "LO")
    'Paste z1,s1,z2,s2 to T4SomeDgData
        DgTitle = .Cells(z1 + 1, s1 + 1): A(1) = z1: A(2) = s1: A(3) = z2: A(4) = s2
        zSome = Get_RowNr_HoldingMyTextWholeInColumnX("T4", 5, 7, DgTitle) 'zT4SomeDgData
        Paste_1DArray_ToRow "T4", zSome, 6, A
    'Check one Dg   (FillArrC, FillArrDg, Fill T4SomeDgData NamesCell, Club/Nation in T5)
        T4_Check_OneDg z1, s1, z2, s2, DgTitle, zSome
    'Finals
        EE 0: .Cells(z2 - 1, s2).Select: .Cells(z2, s2).Interior.Color = vbWhite: End With: EE 1
End Sub

Sub T4_CellBtn_LU(z2%, s1%)
    'Called from    [UserClick on "LU"] Worksheet_SelectionChange[T4]
    'LU             Dg-EckZelle links unten
    
    'Vorbereitung
        Dim DgTitle$, i%, s2%, z1%, zSome%, A()
        With Sheets("T4"): EE 0: ReDim A(1 To 4)
        .Cells(z2, s1).Interior.ColorIndex = 3
    'LU-->LO-->RO-->Title
    'Nächstes "LO" oben suchen 'z1
    'Get z1     'Nächstes "LO" in T4-Salte s1 von z2 aus nach oben suchen
        z1 = Get_RowNr_HoldingMyTextWholeInColumnXSearchUp("T4", s1, z2, "LO"): If z1 = 0 Then Stop
    'Get s2     'Nächstes "RO" in T4-Zeile z1 von s1 aus nach rechts suchen
        s2 = Get_ColumnNr_HoldingMyTextWholeInRowXSearchRight("T4", z1, s1, "RO")
    'zSome      = zNr of DgTitle in zT4SomeDgData
        DgTitle = .Cells(z1 + 1, s1 + 1): A(1) = z1: A(2) = s1: A(3) = z2: A(4) = s2
        zSome = Get_RowNr_HoldingMyTextWholeInColumnX("T4", 5, 7, DgTitle) 'zT4SomeDgData
        If zSome = 0 Then
            'DgTitle steht noch nicht in T4SomeDgData
            T4_Write_T4SomeDgData
            zSome = Get_RowNr_HoldingMyTextWholeInColumnX("T4", 5, 7, DgTitle) 'zT4SomeDgData
            If zSome = 0 Then Stop
        End If
    'Paste z1,s1,z2,s2 to T4SomeDgData
        Paste_1DArray_ToRow "T4", zSome, 6, A
    'Check one Dg   (FillArrC, FillArrDg, Fill T4SomeDgData NamesCell, Club/Nation in T5)
        T4_Check_OneDg z1, s1, z2, s2, DgTitle, zSome
    'Finals
        EE 0: .Cells(z2 - 1, s1).Select: .Cells(z2, s1).Interior.Color = vbWhite: End With: EE 1
End Sub

Sub T4_Check_OneDg(z1%, s1%, z2%, s2%, Title$, zSome%)
    'Called from    T4_CellBtn_LU
    
    'Vorbereitung
        Dim Club$, DgNames$, Nation$, VnNn$, c%, i%, sDg%, zDg%, zT5%, A() As String
    'Some T4_ActionsOnDgClickFirst actions ohne T4_ActionsOnDgClickNext
        FillArrC 36, Get_ColumnLetter(s1) + CStr(z1) + ":" + Get_ColumnLetter(s2) + CStr(z2)
        FillArrC 37, CStr(z1): FillArrC 38, CStr(z2): FillArrC 41, Title
        FillArrC 39, CStr(s1): FillArrC 40, CStr(s2): FillArrC 48, "1"
        FillArrDg 'auch FillArrC 42, 44, 46-47, 49-54, 57
    'Fill T4SomeDgData NamesCell
        DgNames = T4_Get_DataFromOneDg_ToFillOneCell_Names_OfT4SomeDgData(0, z1, s1, z2, s2)
        '       = "004 Names: |#|Kurt Bächler|6|4|Team|1|#|...|#|Rudi Stengel|13|4|Team|2|
        Sheets("T4").Cells(zSome, 12) = DgNames
        'T4SomeDgData-Zeile des aktuellen Dg ist jetzt komplett (FolderCell unverändert)
    'DgNames liefert alle DgNamen; Prüfung Club/Nation-Einträge in T5
        A = Split(DgNames, "|")
        For i = 0 To CInt(Left(DgNames, 3)) - 1
            VnNn = A(2 + 6 * i) 'ContentOfNameCell
                c = InStr(1, VnNn, " ("): If c > 0 Then VnNn = Left(VnNn, c - 1)
            zDg = A(3 + 6 * i)
            sDg = A(4 + 6 * i)
            Club = T4_Get_Club_FromDgClubColumn(zDg, sDg)
            Nation = T4_Get_Nation_FromDgNationColumn(zDg, sDg)
            'Club/Nation ggf. in T5 ergänzen
                zT5 = T5_Get_T5RowNrOfOneVnNn(VnNn)
                With Sheets("T5")
                If Club <> "" Then
                    If .Cells(zT5, 7) = "" Then .Cells(zT5, 7) = Club
                End If
                If Nation <> "" Then
                    If .Cells(zT5, 9) = "" Then .Cells(zT5, 7) = Nation
                End If
                End With
        Next
End Sub

Sub T4_DgHandleNameCell_KnownName(zPd%, zDg%, sDg%, z1%, s1%, Vn$, Nn$)
    'Called from    T4_DgClick_CheckNames
    'Action         Dg-NamenZelle neu füllen
    'Status         User könnte Zusatzangaben gemacht haben, die nicht mehr sichtbar sein sollen;
    '               alle Zusatzangaben wurden bereits ausgewertet;
    '               die Zelle soll letztlich nur Nn, Vn, ggf. Alter tragen, sonst nichts
    
    Dim Alter$, Jhg$, y%: With Sheets("T4")
    y = CInt(Left(ArrDg(2, 2), 4))      'Year of event
    Jhg = CStr(Sheets("T5").Cells(zPd, 6))
    If Jhg = "" Then
        'Zum Aktiven in T5-Zeile zPd ist kein Jahrgang vermerkt
        .Cells(z1 - 1 + zDg, s1 - 1 + sDg) = Vn + " " + Nn 'Dg-NamenZelle ohne Jhg
    Else
        Alter = CStr(y - CInt(Jhg))
        .Cells(z1 - 1 + zDg, s1 - 1 + sDg) = Vn + " " + Nn + " (" + Alter + ")"
    End If
    End With
End Sub

Sub T4_DgHandleLiga(zPd%, zDg%, z1%, s1%)
    'Called from    T4_DgClick_CheckNames
    'Status         Nn, Vn existiert in T5
    'Liga           m/w ist in einem Liga-Dg nicht vermerkt
    '               m/w-Eintrag in Pd (Sheet T5) suchen, dann in Liga-Dg (Sheet T4) anwenden
    '               Verein ist in Liga-Dg immer vermerkt; ggf. in T5 eintragen
    'zPd            = ZeilenNr des aktuellen Namens in T5
    
    Dim DgFoldName$, j%, r As Range: With Sheets("T4")
    DgFoldName = ArrC(42)       'T4 Dg FolderName to write jpg into

    If DgFoldName Like "* Liga\*" Then
        'Verein     (steht bei LigaDg immer in Spalte 3)
            For j = zDg To 6 Step -1
                'NameCompetitor in Zeile zDg steht unterhalb des Vereinsnamens
                If ArrDg(j, 3) <> "" Then
                    If Sheets("T5").Cells(zPd, 7) = "" Then Sheets("T5").Cells(zPd, 7) = ArrDg(j, 3)
                    Exit For
                End If
            Next
    End If
    End With
End Sub

Sub T4_DgHandleLigaFarben(zPd%, zDg%, z1%, s1%)
    'Called from    T4_DgClick_CheckNames
    
    Dim DgFoldName$, r As Range: With Sheets("T4")
    DgFoldName = ArrC(42)       'T4 Dg FolderName to write jpg into

    If DgFoldName Like "* Liga\*" Then
        'm/w-Farben
            Set r = .Range(.Cells(z1 - 1 + zDg, s1 + 3), .Cells(z1 - 1 + zDg, s1 + 11))
            '     = Range der ganzen Zeile, die eingefärbt werden soll
            If Sheets("T5").Cells(zPd, 5) = "m" Then
                r.Interior.Color = 15652797 'hellblau/Blau1 (m-Farbe)
            ElseIf Sheets("T5").Cells(zPd, 5) = "w" Then
                r.Interior.Color = 14083324 'hellrot        (w-Farbe)
            Else
                r.Interior.Color = RGB(222, 222, 222) 'grau (?-Farbe)
            End If
    End If
    End With
End Sub

Sub T4_Change_AllDgNation_ToUCase()
    'Called from    T4_ActionsOnDgClickNext
    'Action         stellt sicher, dass in den Nation-Spalten nur Großbuchstaben vorkommen
    
    'Exit       falls das Dg keine Nation-Spalte besitzt
        If Not ArrDg(2, 1) Like "*N*" Then Exit Sub 'DgGroup: keine Nation-Spalte
    'Vorbereitung
        Dim s%, s1%, z%, z1%
        z1 = CInt(ArrC(37)):  s1 = CInt(ArrC(39))
    'Action
        With Sheets("T4")
        For s = 2 To UBound(ArrDg, 2) - 1
            If "STU" Like "*" + ArrDg(5, s) + "*" Then
                For z = 6 To UBound(ArrDg, 1) - 1
                    If ArrDg(z, s) <> "" Then
                        If ArrDg(z, s) <> UCase(ArrDg(z, s)) Then
                            ArrDg(z, s) = UCase(ArrDg(z, s))
                            .Cells(z1 - 1 + z, s1 - 1 + s) = UCase(ArrDg(z, s))
                        End If
                    End If
                Next
            End If
        Next
        End With
End Sub

Sub T4_Check_OneDgNation_UCase_ValidNation()
    'Called from    T4_ActionsOnDgClickNext
    'Action         Ändert ggf. 1 Eintrag in NationZelle in Versalien
    
    'Exit
        If ArrDg(2, 1) = "M0V1" Then Exit Sub
        If ArrDg(2, 1) = "M0V3" Then Exit Sub
    'Vorbereitung
        Dim NatDg$, s1%, sDg%, sShLast2%, z1%, zDg%, zShLast2%
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39))
    'Zelle, die gerade verlassen wurde
        zShLast2 = CInt(ArrC(106))  'T4 Sh ZeilenNr  Selection Last2
        sShLast2 = CInt(ArrC(107))  'T4 Sh SpaltenNr Selection Last2
        'Exit if FirstClick
            If ArrC(48) = "1" Then Exit Sub
    'Nation-Spalte
        If sShLast2 < 20 Then Exit Sub
        zDg = zShLast2 - z1 + 1: sDg = sShLast2 - s1 + 1
    'Dg(zDg, sDg) = NationCell?
        If Not "STU" Like "*" + ArrDg(5, sDg) + "*" Then Exit Sub
    'Check UCase
        NatDg = ArrDg(zDg, sDg)
        If NatDg = "" Then
            Exit Sub
        Else
            If NatDg <> UCase(NatDg) Then
                NatDg = UCase(NatDg): ArrDg(zDg, sDg) = NatDg
                Sheets("T4").Cells(z1 - 1 + zDg, s1 - 1 + sDg) = NatDg
            End If
        End If
    'Check ValidNation
        Dim A$, NatsT8$, antwort%, i%, zLastT8%, ArrT8()
        zLastT8 = Get_NrOfLastRowInColumnNr(9, "T8"):               With Sheets("T8")
        ArrT8 = .Range(.Cells(7, 9), .Cells(zLastT8, 11)).Value: End With 'Spalte 9
        For i = 1 To UBound(ArrT8, 1)
            NatsT8 = NatsT8 + "|" + ArrT8(i, 1)
        Next
        NatsT8 = NatsT8 + "|" '|ARG|AUS|...|WAL|
        If NatsT8 Like "*|" + NatDg + "|*" Then Exit Sub 'alles OK
    'NatDg ist keine ValidNation
        'Ist NatDg in der "auch verwendete Kennungen"-Spalte enthalten?
            For i = 1 To UBound(ArrT8, 1)
                A = "|" + ArrT8(i, 3) + "|" 'a = "|ABC DE|" 'oneCellContent OtherNationShorts-Column
                A = Replace(A, " ", "|")    'a = "|ABC|DE|"
                If A Like "*|" + NatDg + "|*" Then
                    'Natdg ist in a enthalten (in Zeile i der "auch verwendete Kennungen"-Spalte)
                    NatDg = ArrT8(i, 1) 'NatDg = ValidNation
                    'NatDg-Eingabe wird durch ValidNation ersetzt
                        ArrDg(zDg, sDg) = NatDg
                        Sheets("T4").Cells(z1 - 1 + zDg, s1 - 1 + sDg) = NatDg: Exit Sub
                End If
            Next
    'Die eingegebenen Buchstaben sind eine neue NationKennung
        If Len(NatDg) > 3 Then Exit Sub '"FixValues: " steht ggf. in einer Nation-Spalte
        antwort = MsgBox(Prompt:="Soll       " + NatDg + vbCrLf + "als neue Nation-Kennung " _
        + "registriert werden?", Buttons:=vbYesNo + vbQuestion, Title:="Nation-Kennung")
        If antwort = vbYes Then
            'Neuer Eintrag in T8
            Sheets("T8").Cells(zLastT8 + 1, 9) = NatDg
        Else
            Sheets("T4").Cells(z1 - 1 + zDg, s1 - 1 + sDg) = ""
        End If
End Sub









