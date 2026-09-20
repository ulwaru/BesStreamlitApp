Attribute VB_Name = "T6"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Vorbereitung
        Dim z As Integer, s As Integer                  'z = Zeile, s = Spalte
        z = Target.Row: s = Target.Column               'momentan ausgewählte Zelle ermitteln
        Beep1
    'Einer der CellButtons (T6 - Spalte 2)
        If s = 2 Then T6_UserClickOnColumn2CellButton z, s: Exit Sub
    'OK-Button bei eingeblendeter InfoListe 'Change ... to ... in EventFolderNames'
        If s = 25 Then T6_CellButton_Column25 z, s: Exit Sub
    'Abbrechen-Button bei eingeblendeter InfoListe 'Change ... to ... in EventFolderNames'
        If s = 28 Then T6_CellButton_Column28 z, s: Exit Sub
    'Ja-Button bei 'Format Event-FileNames'
        If z = 8 And s = 37 Then T6_BMB_CellButton_Ja z, s: Exit Sub
    'Nein-Button bei 'Format Event-FileNames'
        If z = 8 And s = 39 Then T6_BMB_CellButton_Nein z, s: Exit Sub
    'Abbrechen-Button bei 'Format Event-FileNames'
        If z = 8 And s = 41 Then T6_BMB_CellButton_Abbrechen z, s: Exit Sub
    'OpenEventFolder-Button bei 'Format Event-FileNames'
        If z = 8 And s = 43 Then T6_BMB_CellButton_OpenEventFolder z, s: Exit Sub
End Sub

Sub Add_pID_to_oneJpgFile(PathOfOneJpgFile$)
    'Procedure 1 for VBS
    'Vorbereitung
        Dim f1$, f2$, N1$, v$
        f1 = PathOfOneJpgFile: v = vbCrLf
        N1 = Get_NameFromPath(f1)
    'f1 bitte ohne ####-## .lnk .url .txt !
        If N1 Like "*!*" Or N1 Like "*####-##*" Or N1 Like "*.lnk" Or N1 Like "*.url" Or N1 Like "*.txt" Then Exit Sub
    'f1 = Path of one File (.JPG?)
    If LCase(Right(f1, 4)) = ".jpg" Then
        'f2 = jpg-File-Pfad, 'jpg' klein geschrieben (nur in f2, nicht in f1)
        f2 = Left(f1, Len(f1) - 4) + ".jpg"
        If f2 Like "*\####* [a-z][a-z].jpg" Then
            'f2 = FilePath (hat QuellenKürzel)
            If Not f2 Like "* p####-## *" Then
                'f2 hat noch keine pID; pID hinzufügen
                f2 = Left(f2, Len(f2) - 6) + GetPID + Mid(f2, Len(f2) - 6)
                RenameFile f1, f2
            End If
        Else
            'Photo f1 besitzt kein Quellenkürzel
            If Not f1 Like "*\a Leute.jpg" Then 'kein Quellenkürzel nötig
                MsgBox "Das Photo " + v + "   " + f1 + v + "besitzt kein Quellenkürzel." + v _
                   + "Bitte fügen sie ein solches in den Dateinamen ein." + v _
                   + "Erst danach kann das Photo eine ID erhalten."
                Exit Sub
            End If
        End If
    End If
    Beep
End Sub

Function GetPID() As String
    'Procedure 2 for VBS
    Dim p$, Pid$, s$
    p = "F:\Archiv Trampolin prog\prog\helpers\FreePIDs.txt"
    'Die Datei 'FreePIDs.txt' wird in s eingelesen
        s = ReadFile(p) '"p2563-00 p2593-00 ... p3396-00 "
    'Die erste pID wird aus s entnommen
        Pid = Left(s, 8)
        s = Mid(s, 10)
    'Reduzierte FreePIDs-Liste zurückschreiben
        writeStringToFile p, s
    Get_NextFreePID = Pid
End Function

Sub writeStringToFile(pathForFileToWrite$, StringToWrite$)
    'Procedure 3 for VBS
    'schreibt Ansi-TextFile
    Dim i As Long: i = FreeFile
    Open pathForFileToWrite For Output As i: Print #i, StringToWrite: Close i
End Sub

Sub zzz_T6()
    showProcs "add club"
    
    'ActiveWindow.SmallScroll Up:=1, ToRight:=1
    'show [z78].Value
    'RenameModule "ArrHelpers2", "T6"
    'T1_ShowLogBuch
    'Sheets("T6").Range("W5:ai6").Interior.Color = Green4
    EE 1: Beep
End Sub

Sub T6_UserClickOnColumn2CellButton(z%, s%)
    'Called from    Worksheet_SelectionChange[T6]
    's              = 2 'SpaltenNr des UserClick
    
    'CellButtons
        Select Case Cells(z, s).Value
            Case "": Exit Sub
            Case " Add EventLink to PersonFolders":             T4_Add_EventLink_ToPersonFolders: Exit Sub
            Case " Add IDs in open EventFolders":               T6_Add_IDs_toFilesInOpenEventFolders z: Exit Sub
            Case " Add Links: Person-Club-LTV-Nation":          T6_Add_Links_PersonLTVClubNation: Exit Sub
            Case " Add LTV/Nation in T5":                       T6_Add_T5LTVNation_LikeExistingCombinations: Exit Sub
            Case " Change String   xxx", "xxx    to    xxx":    T6_ChangeEFN_Start z: Exit Sub
                   Case " inside 'Events'-FolderNames …      ": T6_ChangeEFN_Start z - 1: Exit Sub
            Case " Check folder paths":                         T6_Check_FolderPaths z: Exit Sub
            Case " Check PIDs VIDs":                            T6_Check_PIDs_VIDs: Exit Sub
            Case " Remove privat tags in JPGs":                 T6_Remove_PrivatTags: Exit Sub
            Case " Show missing persons - in 'Leute' / in T5":  T6_ShowMissingPersonsInLeuteInT5: Exit Sub
            Case " Show all Paths of Events-Result-JPGs":       T6_Show_AllPathsOfEventResultJpgs z: Exit Sub
            Case " Update CompetitorList":                      T6_Update_CompetitorsList z: Exit Sub
            Case " Update Age of all Competitors in all Dgs":   T6_Update_AgeOfAllCompetitorsInsideAllDgs z: Exit Sub
            Case " Update each Result.jpg":                     T6_Update_EachResultJpg z: Exit Sub
            Case " EventFileNames: Add ' - ' ...":              T6_BMB_EventFileNames: Exit Sub
            Case " Create DateSquares":                         T6_Create_DateSquares: Exit Sub
            Case " Create ico of zzJpg in open PersonFolder":   T6_Create_PersonLinkIcon_Of_zzJpg_in_OpenPersonFolder z: Exit Sub
        End Select
End Sub

Sub T6_()
    Dim L1$, L2$, M1$, M2$, p$, s$, T$, v$, i%, Arr1() As String
    v = vbCrLf
    p = ArrC(4)                'T0 Path of Folder "Leute"

    s = Get_FilePaths_Like_insideSourceFolderAndSubFolders(p, "*[#]result*")
'    arr1 = Split(s, v)
'    For i = 0 To UBound(arr1) - 1
'        L1 = arr1(i):       M1 = Left(L1, InStr(1, L1, "#result"))
'        L2 = arr1(i + 1):   M2 = Left(L2, InStr(1, L2, "#result"))
'        If L1 = L2 Then t = t + L1
'    Next
    show s
End Sub

Sub T6_Update_NationenKennungen()
    'Called from    xxx

    'Vorbereitung
        Dim L$, L1$, Nat$, NatsT5$, NatsT8$, OneNat$, p$, s$, s2$, s3$, v$
        Dim c%, i%, zLastT5%, zLastT8%, Arr5() As String, ArrT5(), ArrT8()
        v = vbCrLf: DoArrc
        p = ArrC(2) + "\Nationen-Kennungen.txt"
'    'Read
'        s = ReadFile(p)
'        s = Replace(s, "Nationen-Kennungen" + v + v, "")
'        'show s: Stop
    'NatsT5: NationenListe aus T5 erstellen
        With Sheets("T5"): zLastT5 = Get_NrOfLastRowInColumnNr(9, "T5")
        ArrT5 = .Range(.Cells(8, 9), .Cells(zLastT5, 9)).Value: End With
        For i = 1 To UBound(ArrT5, 1)
            OneNat = ArrT5(i, 1)
            If Not NatsT5 Like "*|" + OneNat + "|*" Then NatsT5 = NatsT5 + "|" + OneNat + "|" + v
        Next
        NatsT5 = Delete_EmptyRowsInString(NatsT5)
        NatsT5 = Replace(NatsT5, "|", "")      'GB D SA ...
        Arr5 = Split(NatsT5, v)
    'NatsT8: NationenListe aus T8 erstellen
        zLastT8 = Get_NrOfLastRowInColumnNr(9, "T8")
        With Sheets("T8")
        ArrT8 = .Range(.Cells(7, 9), .Cells(zLastT8, 9)).Value 'Spalte 9
        For i = 1 To UBound(ArrT8, 1)
            NatsT8 = NatsT8 + "|" + ArrT8(i, 1)
        Next
        NatsT8 = NatsT8 + "|"
    'NationenListe in T8 ggf. ergänzen
        For i = 0 To UBound(Arr5)
            If Not NatsT8 Like "*|" + Arr5(i) + "|*" Then
                c = c + 1
                .Cells(zLastT8 + c, 9) = Arr5(i)
            End If
        Next
    'Finals
        End With
End Sub

Sub T4_CheckAllNames_ActualDg()
    'Called from    T6_Update_EachResultJpg
    'ActualDg       One Dg is active; ArrC-Entries done, ArrDg exists
    
    'Vorbereitung
        Dim i%, s%, sDgNameW%, sDgNameM%, sDgNameX%, z%, s1&, z1&, ArrCol(1 To 3) As Integer
        With Sheets("T4")
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39))

    'NameColumns
        sDgNameW = CInt(ArrC(90))    'T4 Dg SpaltenNr Namen  weibl.
        sDgNameM = CInt(ArrC(91))    'T4 Dg SpaltenNr Namen  männl.
        sDgNameX = CInt(ArrC(92))    'T4 Dg SpaltenNr Namen  mix
        ArrCol(1) = sDgNameW: ArrCol(2) = sDgNameM: ArrCol(3) = sDgNameX
    'Select all competitors of actual Dg
        For i = 1 To 3
            s = ArrCol(i)   's = Dg-SpaltenNr einer NamenSpalte
            If s > 0 Then
                For z = 6 To UBound(ArrDg, 1) - 1
                    If ArrDg(z, s) <> "" Then .Cells(z1 - 1 + z, s1 - 1 + s).Select: RefreshScreen
                    'Stop
                Next
                .Cells(z1, s1).Select 'wegen LastSelection
            End If
        Next
    'Finals
        End With
End Sub

Public Sub Worksheet_Activate()
    'Zeilenhöhe
    
    'Neuen Wert je DaysAgo setzen
        Dim D$, c%, i%, zLast&, A()
        zLast = Get_NrOfLastRowInColumnNr(18, "T6")
        A = Range(Cells(1, 18), Cells(zLast, 18)).Value
        For i = 1 To UBound(A, 1)
            D = A(i, 1)
            If D Like "########*" Then
                D = Left(D, 8)
                c = DaysAgo_GiveDate8(D)
                Cells(i, 19) = c
            End If
        Next
End Sub

Sub T6_Check_FolderPaths(zT6%)

    Stop 'alt; arbeitet mit T8SomeDgData; ggf. überarbeiten
    Exit Sub
    
    'Called from    T6_UserClickOnColumn2CellButton
    'zT6            ZeilenNr des CellButtons ' Check folder paths'
    'Check          (1) E1 = Pfade EFN, EventFolderNames
    '               (2) P1 = Pfade Termine; Übereinstimmung mit (1) herstellen
    '               (3) P2 = Pfade T8SomeDgData; Spalte 11: "Folder"
    '               (4) P3 = Pfade OriginalFoto zu T7 Labels
    
    'Vorbereitung
         Dim L$, E1$, E2$, p1$, p2$, p3$, p4$, Q2$, s$, s1$, v$, C2%, i%, zLast&
         Dim A0() As String, A2() As String, a3() As String, T(), r As Range
         v = vbCrLf: DoArrc: Ticks1: s1 = v + String(60, "-") + v
         s = "T6_Check_FolderPaths" + v + s1
    '(0) E1 = Pfade aller Folder in 'Events', die die MainEventFiles tragen ("...#result..", ...)
         E1 = Get_PathsOfAllDateEventFolders + v 'nur existierende Ordner
         '  = ReferenzNamen, SOLL-Schreibweisen;
         E2 = Get_Paths_ofAllSubfoldersAllLevelsAsStringUseGlobalVar(ArrC(2) + "\ClubsNations\") + v
    '(1) P1 = Pfade, wie sie in T9 (Terminkalender) vermerkt sind
         p1 = Get_ListOfNonEmptyEntriesOfOneColumn("T9", 13, 11) + v
         'Alle E1-Pfade sollten auch in P1 vorkommen
         A0 = Split(E1, v)
         For i = 0 To UBound(A0)
            L = A0(i) + v       'one DateEventFolder
            If p1 Like "*" + L + "*" Then
                p1 = Replace(p1, L, v) 'L aus P1 herauslöschen
            Else
                Q2 = Q2 + L
            End If
         Next
         p1 = Delete_EmptyRowsInString(p1) 'Pfade, in T9, nicht in Events
         Q2 = Delete_EmptyRowsInString(Q2) 'Pfade, in Events, nicht in T9
         C2 = anzAinB(v, p1) + anzAinB(v, Q2)
         'Show
            s = s + "     CHECK 1: Terminkalender" + v
            If p1 <> "" Then s = s + "P1   Pfade, die im T9-Terminkalender vermerkt sind," + v + "     aber nicht als Ordner in 'Events' existieren." + v + v + p1 + v + v
            If Q2 <> "" Then s = s + "Q2   Pfade, die als Ordner in 'Events' existieren," + v + "     aber nicht im T9-Terminkalender vermerkt sind." + v + v + Q2 + v + v
            If p1 + Q2 <> "" Then s = s + "Bitte korrigieren Sie dies in T9." + s1
            If p1 + Q2 = "" Then s = s + "Alle Events-Pfade existieren auch in T9." + s1
    '(2) P2 = Pfade aus T8SomeDgData, die nicht mehr aktuell sind (nicht in E1)
         zLast = Get_NrOfLastRowInColumnNr(11, "T8")
         Load_Array1D_OfColumnPart A2, "T8", 7, CInt(zLast), 11
         For i = 1 To UBound(A2)
            L = ArrC(3) + "\" + A2(i)
            If Not E1 Like "*" + L + "*" + v Then p2 = p2 + L + v: C2 = C2 + 1
         Next
         s = s + "     CHECK 2: T4-Designs" + v
         If p2 = "" Then s = s + "Alle in T8SomeDgData vermerkten Event-Pfade sind aktuell," _
                             + v + "d. h. auch alle Pfade in den T4-Designs sind aktuell." + s1
         If p2 <> "" Then s = s + "Die folgenden Event-Pfade in T8SomeDgData sind nicht mehr aktuell." _
                             + v + "Bitte korrigieren Sie dies in den T4-Designs." + s1
    '(3) P3 = nicht mehr aktuelle Pfade aus T7 Labels
         zLast = Get_NrOfLastRowInColumnNr(52, "T7")
         Load_Array1D_OfColumnPart a3, "T7", 5, CInt(zLast), 52
         'showArray A3
         For i = 1 To UBound(a3)
            If Not FileExists(a3(i)) Then p3 = p3 + a3(i) + v: C2 = C2 + 1
         Next
         s = s + "     CHECK 3: Labels" + v
         If p3 = "" Then s = s + "Alle in T7 (Labels) vermerkten Pfade sind aktuell." + s1
         If p3 <> "" Then s = s + "Die folgenden Pfade sind nicht mehr aktuell:" _
                             + v + v + p3 + v + "Bitte korrigieren Sie dies in T7." + s1
    'DoneRemarks to T6
        'T6_DoDoneRemarks c2
    'Finals
        show s
End Sub

Sub T6_Add_SomeFolderNames_T5_Leute()
    'Vorbereitung
        Dim s$, L$, Leute_NotT5$, m$, NameOld$, Nn$, p1$, p2$, SOLL$, T5_NotLeute$, v$, Vn$
        Dim C1%, C2%, C3&, i%, z%, A() As String, B() As String, N() As String
        v = vbCrLf
    'Load
        T6_Load_inconsistencies_inPdFolderNames C1, C2, T5_NotLeute, Leute_NotT5
    'NamesLeuteOrdner
        m = Get_Names_OfAllSubfolders_OneLevel(ArrC(4))
    'Format s (= T5_NotLeute)
        s = T5_NotLeute
        s = Replace(s, v + "   FolderName SOLL:       ", "|")
        s = Replace(s, v + "   FolderName in T5:      ", "|")
        s = Replace(s, v + "   FolderName in 'Leute': ", "|")
        s = Replace(s, v, "|" + v)
        For i = 1 To C2
            s = Replace(s, "(" + CStr(i) + ")  ", "")
        Next
        s = Delete_EmptyEndRowsInString(s)
        'show s
    'SOLL-FolderName to T5, to Leute
        A = Split(s, v): With Sheets("T5")
        For i = 0 To UBound(A)          'L = Nn, Vn  |SOLL           |in T5|in 'Leute'|
            L = A(i)        'one Line   'L = May, Lea|May, Lea (Bonn)|xxxxx|-| 'FolderNames
            B = Split(L, "|")
            Nn = Left(B(0), InStr(1, B(0), ",") - 1)
            Vn = Replace(B(0), Nn + ", ", "")
            SOLL = B(1)
            'SOLL-FolderName in T5 schreiben
                'Nn, Vn existieren in T5
                z = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, Vn + " " + Nn)
                If z > 0 Then .Cells(z, 10) = SOLL Else Stop
            'Nn, Vn in M suchen
                C3 = InStr(1, m + v, Nn + ", " + Vn + " ")
                p2 = ArrC(4) + "\" + SOLL
            'FolderName in 'Leute' in SOLL-FolderName ändern
                If C3 > 0 Then
                    NameOld = CutVonBis(m, C3, InStr(C3, m + v, v) - 1)
                    p1 = ArrC(4) + "\" + NameOld
                    RenameFolder p1, p2
                End If
            'SOLL-FolderName neu anlegen
                If C3 = 0 Then CreateFolder p2
        Next
    'New Names inside Leute to T5
        Leute_NotT5 = Delete_EmptyEndRowsInString(Leute_NotT5)
        N = Split(Leute_NotT5, v): With Sheets("T5")
        For i = 0 To UBound(A)          'L = Nn, Vn  |SOLL           |in T5|in 'Leute'|
            
            
            
            
        Next
        
        
    'Finals
        End With
        Beep
End Sub

Sub T6_ShowMissingPersonsInLeuteInT5()
    'Vorbereitung
        Dim Leute_NotT5$, s$, T5_NotLeute$, v$, C1%, C2%
        v = vbCrLf
    'Load
        T6_Load_inconsistencies_inPdFolderNames C1, C2, T5_NotLeute, Leute_NotT5
    'Show
        s = CStr(C1) + " Personenordner in 'Leute', aber Person nicht in T5:" _
        + v + v + Leute_NotT5
        If C2 > 0 Then s = s + v + CStr(C2) _
            + " Personen in T5, aber OrdnerName nicht/nicht identisch in 'Leute':" _
            + v + v + T5_NotLeute
        show s
End Sub

Sub T6_Load_inconsistencies_inPdFolderNames(C1%, C2%, T5_NotLeute$, Leute_NotT5$)
    'Vorbereitung
        Dim Club$, Gn$, L$, lösch$, m$, N$, NaFoLe$, NaFoT5$, Nation$, Nn$
        Dim VnNnM$, rn$, s$, v$, vh$, Vn$, i%, zNr%, C3&, C4&, A(), r As Range
        v = vbCrLf: DoArrc: Ticks1
    'Load T5
        T5_Load_Nn_to_Vh A
    'NamesLeuteOrdner
        m = Get_Names_OfAllSubfolders_OneLevel(ArrC(4)) + v
        VnNnM = m 'Zeilen wie "May, Lea (Bonn)"
        'show VnNnM
        
    'Action
        For i = 1 To UBound(A, 1)
            'Schleife über alle T5-Zeilen
            'items of one T5Line
                Nn = A(i, 1):     Vn = A(i, 2): Club = A(i, 5): Nation = A(i, 7)
                rn = A(i, 10):    Gn = A(i, 11)
                NaFoT5 = A(i, 8)  'NameOfFolder AsWrittenInT5 'May, Lea (vh Juni, rn Susi, Bonn)
                vh = A(i, 12)     '- 'Latton/Luxon/Pitkamin
            'calc
                If Club = "" And A(i, 6) <> "" Then Club = A(i, 6)
                If vh = "-" Then vh = "" Else vh = "vh " + Replace(vh, "/", " ")
                If Gn = "-" Then Gn = Nn
                If rn = "-" Then rn = "" Else rn = "rn " + rn
                If NaFoT5 = "" Then NaFoT5 = "-"
            'VnNnM NamesLeuteOrdner
                C3 = InStr(1, VnNnM, Nn + ", " + Vn)
                If C3 > 0 Then
                    'Zeile mit Nn, Vn aus VnNnM herauslöschen
                    C4 = InStr(C3 + 1, VnNnM, v)
                    lösch = CutVonBis(VnNnM, C3, C4 + 1)
                    VnNnM = Replace(VnNnM, lösch, "")
                End If
            'NameOfFolder SOLL
                N = Gn + ", " + Vn + " (" + vh + ", " + rn + ", " + Club + ", " + Nation + ")"
                N = Replace99(N, "(, ", "("):   N = Replace99(N, ", , ", ", ")
                N = Replace(N, ", )", ")"):     N = Replace(N, ", D)", ")")
            'NameInFolderLeute
                C3 = InStr(1, m, Nn + ", " + Vn + " ")
                If C3 > 0 Then
                    NaFoLe = CutVonBis(m, C3, InStr(C3 + 1, m, v) - 1)
                Else
                    NaFoLe = "-"
                End If
            'List
                If Not (N = NaFoT5 And N = NaFoLe) And Not vh Like "*" + Nn + "*" Then
                    C2 = C2 + 1
                    L = L + "(" + CStr(C2) + ")  " + Nn + ", " + Vn + v _
                        + "   FolderName SOLL:       " + N + v _
                        + "   FolderName in T5:      " + NaFoT5 + v _
                        + "   FolderName in 'Leute': " + NaFoLe + v
                End If
        Next
    'DoneRemarks to T6
        'T6_DoDoneRemarks c1
        zNr = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 2, 4, " Show missing persons - in 'Leute' / in T5")
    'LastDone
        Set r = Sheets("T6").Cells(zNr, 18)
        r.Value = Format(Now(), "yyyymmdd_hhmmss")
        r.Font.size = 6: r.HorizontalAlignment = xlLeft
    'Days ago
        Set r = Sheets("T6").Cells(zNr, 19)
        r.Value = 0: r.Font.size = 8: r.HorizontalAlignment = xlCenter
    'Count of changes (Anzahl geänderter T5-Zeilen)
        Set r = Sheets("T6").Cells(zNr, 20)
        C1 = CInt(anzAinB("(", VnNnM))
        r.Value = C1 + C2: r.Font.size = 8: r.HorizontalAlignment = xlCenter
    'Erledigt-Häkchen setzen
        Sheets("T6").Cells(zNr, 16) = "ü"
    'Duration
        Sheets("T6").Cells(zNr, 21) = T6_GetDuration()
    'Übergabe
        T5_NotLeute = L
        Leute_NotT5 = VnNnM
'CreateFolder
        
'    Update_NameOfLeuteFolders_From_Pd
'    Update_NamesOfPersonFolders
'    T6_CheckFolders
'    RenameFolder
'    RenameFolders_Date10CharsTo8Chars
'    Load_ArrPathsOfAllFolders_Events
'    Load_ArrNamesOfAllFolders_Leute
'    FolderExists
'    CalCnewFolderName
End Sub

Sub T6_CheckFolders()
    'Vorbereitung
        Dim F$, p$, v$, C1%, i%, j%, s1%, z1%, Arr1() As String, Arr2()
        v = vbCrLf: DoArrc
        p = ArrC(2)             'Path of Folder "Archiv Trampolin"
        z1 = 5: s1 = 25         'Zeile/Spalte linke obere Ecke
        'DoEvents

    'OrdnerPfade Level 1
        F = Get_Paths_ofAllSubfolders_OneLevel(p)
        'f = Get_Paths_ofAllSubfoldersAllLevelsAsStringUseGlobalVar(p2)
        'show f
        Arr1 = Split(F, v)
        C1 = UBound(Arr1) + 1 'Anzahl Ordner Level 1
        ReDim Arr2(1 To C1 + 1, 1 To 10)
        For i = 0 To UBound(Arr1)
            Arr2(i + 1, 1) = i + 1
            Arr2(i + 1, 2) = Replace(Arr1(i), p + "\", "")
            For j = 3 To 10
                Arr2(i + 1, j) = 0
            Next
        Next
        Paste_2DArrayToCell_z_s "T6", z1 + 1, s1, Arr2
        
'    'Pfade aller JPGs in 'Leute'
'        pL = Get_Paths_OfAllJPGs_InFolderAndSubfolders(ArrC(4)) 'Leute
'        .Cells(14, 19) = "JPGs: " + CStr(anzAinB(v, pE) + 1) + " in 'Events'; " + CStr(anzAinB(v, pL) + 1) + " in 'Leute'"
'    'Check
'        Check_Format pE + v + pL
'        Check_DoubleIDsInEvents pE
'
'            'Same ID, different Date
'            'Video format doubleIDs
'            'Files with same ID should have the same Shorty (Quellenkürzel)
'
'        LogBuch "CheckConsistence has ended"
'    End With: EE 1: Beep
End Sub

Function Get_T6vh(r$) As String
    'Called from    Write_T6AllFiles_IdTxt
    'R              = "Mai, Lea (vh Juni, ..., vh Juli, TV Bern, CH)" = FolderNameLeute
    
    'Vorbereitung
        If Right(r, 1) <> ")" Then Exit Function
        If Not r Like "*(*vh*" Then Exit Function
        Dim vh$, i%, Arr1() As String
    'Action
        r = CutVonBis(r, InStr(1, r, "(") + 1, Len(r) - 1)   'vh Juni, ..., CH
        Arr1 = Split(r, ", ")
        For i = 0 To UBound(Arr1)
            If Left(Arr1(i), 3) = "vh " Then vh = vh + ", " + Mid(Arr1(i), 4)
        Next
        vh = Mid(vh, 3)
        Get_T6vh = vh
End Function

Sub T6_ScrollSlow(CountOfCells%)
    Dim A%, B%, c%, i%
    'x = Cells(1, 26).EntireColumn.ColumnWidth
    A = ActiveWindow.ScrollColumn: B = A + CountOfCells
    If B < 1 Then B = 1
    If CountOfCells > 0 Then c = 1 Else c = -1
    For i = A To B Step c
        ActiveWindow.ScrollColumn = i  'Spalte, die ganz links zu sehen sein soll
        If Cells(1, i).EntireColumn.Hidden = False Then Sleep 20
    Next

    '[a1].Select
End Sub

Sub T6_ChangeSpaltenAusblenden() 'Sp 22-29
    'Vorbereitung
        Dim i%, w!, W2!
    'Breite Spalte[AB] auf Width=2 verkleinern
        w = [ab1].EntireColumn.ColumnWidth '30
        For i = 1 To 10
            W2 = i * w / 10: [ab1].EntireColumn.ColumnWidth = w - W2: Sleep 20: DoEvents
        Next
    'Spalte AA = " "
        Range(Cells(5, 27), Cells(12, 34)) = " ": DoEvents
    'Breite Spalte[Z] auf Width=2 verkleinern
        w = [z1].EntireColumn.ColumnWidth '30
        For i = 1 To 10
            W2 = i * w / 10: [z1].EntireColumn.ColumnWidth = w - W2
            Call T6_Scroll1Left: Sleep 20: DoEvents
        Next
    'Breite Spalten auf Width=0 verkleinern
        For i = 29 To 22 Step -1
            Cells(1, i).EntireColumn.Hidden = True: T6_Scroll1Left: Sleep 20: DoEvents
        Next
End Sub

Sub T6_ChangeSpaltenEinblenden() 'Sp 22-29
    'Vorbereitung
        Dim i%, w!, W2!
        w = 30 'Soll-Breite
        Range(Cells(5, 34), Cells(99, 34)) = " "
    'SpaltenBreiten auf w=2 vergößern
        For i = 22 To 29
            Cells(1, i).EntireColumn.ColumnWidth = 2: Sleep 20: DoEvents
        Next
        Cells(1, 25).EntireColumn.ColumnWidth = 4: Sleep 20: DoEvents
    'Breite [AB] auf w=30 vergrößern
        For i = 1 To 10
            W2 = i * w / 10: [ab1].EntireColumn.ColumnWidth = W2
            T6_Scroll1Right
            Sleep 20: DoEvents
        Next
    'Spalte AA = ""
        Range(Cells(5, 27), Cells(12, 27)) = "": DoEvents
    'Breite [Z] auf w=30 vergrößern
        For i = 1 To 10
            W2 = i * w / 10: [z1].EntireColumn.ColumnWidth = W2: Sleep 20: DoEvents
        Next
End Sub

Sub T6_Scroll1Left()
    Dim sc%
    sc = ActiveWindow.ScrollColumn
    If sc > 1 Then
        ActiveWindow.ScrollColumn = sc - 1 'Spalte, die ganz links zu sehen sein soll
    End If
End Sub

Sub T6_Scroll1Right()
    Dim sc%
    sc = ActiveWindow.ScrollColumn
    ActiveWindow.ScrollColumn = sc + 1 'Spalte, die ganz links zu sehen sein soll
End Sub

Sub T6_CellButton_Column25(z%, s%)
    'Called from    [UserClick on OK-Button] OK-Button bei eingeblendeter InfoListe
    '               'Change ... to ... in EventFolderNames'
    'Status         In Spalte 25 befindet sich ein CellButton "OK";
    '               wurde dieser gedrückt, erfolgen Änderungen von EventFolderNames
    
    'Nur OK-Btn wird berücksichtigt
        If Cells(z, s) <> "OK" Then Exit Sub
    'Vorbereitung
        Dim s1$, C1%, i%, z1%, zNr%, A(), r As Range
    'OK für Änderung einiger EventFolderNames
        Range(Cells(z, s), Cells(z + 1, s)).Interior.Color = Green1
        Wait_MilliSec 200
        Range(Cells(z, s), Cells(z + 1, s)).Interior.Color = Green3
     'Get saved T6RangeOfNames
        s1 = [AF2]
        If Not s1 Like "|###|###|" Then [AF2] = "": GoTo jump1
     'A() = 2D-Array = Liste PfadAlt, PfadNeu
        z1 = CInt(Mid(s1, 2, 3)) 'ZeilenNr Start Liste OrdnerNamen  (z.B. 13 aus "|013|007|")
        C1 = CInt(Mid(s1, 6, 3)) 'Anzahl OrdnerNamen                (z.B.  7 aus "|013|007|")
        FillArrC 24, CStr(C1) 'CountOfChanges
        A = Range(Cells(z1, 31), Cells(z1 + C1 - 1, 32)).Value
    'Änderung der FolderNames
        For i = 1 To UBound(A)
            RenameFolder CStr(A(i, 1)), CStr(A(i, 2))
        Next
    'DoneRemarks to T6
        'CountDownShowCell-ZeilenNr wurde in T6_ChangeEFN_Start gesetzt
        T6_DoDoneRemarks
jump1:
    'T6-Spalten ausblenden
        T6_ChangeSpaltenAusblenden
        [A1].Select
End Sub

Sub T6_CellButton_Column28(z%, s%)
    'Status         In Spalte 28 könnte sich entweder ein CellButton "OK"
    '               oder ein CellButton "Abbrechen" befinden;
    '               in keinem Fall wird hier ein EventFolderName geändert
    
    If Cells(z, s) = "" Then Exit Sub
    If Cells(z, s) = "OK" And Cells(z, s - 2) = "" Then
        'Kein Änderung von EventFolderNames (nur LnNo enthält Zeilen)
        Range(Cells(z, s), Cells(z + 1, s)).Interior.Color = Green1
        Wait_MilliSec 200
        Range(Cells(z, s), Cells(z + 1, s)).Interior.Color = Green3
        'T6-Spalten ausblenden
            T6_ChangeSpaltenAusblenden
        [A1].Select
    End If
    If Cells(z, s) = "Abbrechen" And Cells(z, s - 3) = "OK" Then
        'Einige EventFolderNames könnten geändert werden,
        'allerdings wurde der CellButton "Abbrechen" gedrückt
        Range(Cells(z, s), Cells(z + 1, s)).Interior.Color = Green1
        Wait_MilliSec 200
        Range(Cells(z, s), Cells(z + 1, s)).Interior.Color = Green3
        'T6-Spalten ausblenden
            T6_ChangeSpaltenAusblenden
        [A1].Select
    End If
End Sub

Sub T6_Analyze_ChangeEventFolderName(p$, N$, t1$, T2$, LnNo$, LnFrom$, LnTo$, LpFrom$, LpTo$)
    'Called from    T6_ChangeEFN_Start
    'P, N           Path, Name of one EventFolder
    'T1, T2         String T1 soll (innerhalb N) durch T2 ersetzt werden;
    '               T1 ist in EventFolderName N enthalten, ggf. mehrfach;
    '               keine erneute Ersetzungen, falls T1 in T2 vorkommt
    'LnNo           Liste FolderNames, NoChange                              (anfangs leer)
    'LnFrom, LnTo   Liste FolderNames, Change from/to   (OldName to NewName) (anfangs leer)
    'LpFrom, LpTo   Liste FolderPaths, Change from/to   (OldPath to NewPath) (anfangs leer)
    
    'Vorbereitung
        Dim Nnew$, v$, C1%, C2%, C3%, C4%, T1inT2 As Boolean
        v = vbCrLf
    'c1 (Position T1inN, 1. Vorkommen), c2 (Position T1inN, 2. Vorkommen)
        C1 = InStr(1, N, t1): C2 = InStr(C1 + 1, N, t1)
        C3 = InStr(1, N, T2): C4 = InStr(C3 + 1, N, T2)
    'T1inT2 (T1 ist in T2 enthalten)
        If InStr(1, T2, t1) > 0 Then T1inT2 = True Else T1inT2 = False
    'Typ1:  T1 nicht in T2 oder (T1inT2, T2 nicht in N)
            If Not T1inT2 Or (T1inT2 And C3 = 0) Then
                LnFrom = LnFrom + N + v: LpFrom = LpFrom + p + v: Nnew = Replace(N, t1, T2)
                LnTo = LnTo + Nnew + v: LpTo = LpTo + Replace(p, N, Nnew) + v: Exit Sub
            End If
    'Typ2:  T1inT2, T1 1x in N, T2 1x in N
            If T1inT2 And C2 = 0 And C3 > 0 Then LnNo = LnNo + N + v: Exit Sub
    'Typ3:  T1inT2, T1 2x in N, T2 1x in N
            If T1inT2 And C2 > 0 And C4 = 0 Then
                '1 der 2 Vorkommen von T1 in N kann durch T2 ersetzt werden
                Nnew = Replace(N, T2, t1) 'Rückersetzung T2 --> T1
                Nnew = Replace(Nnew, t1, T2) 'beide T1 werden je zu T2
                LnFrom = LnFrom + N + v: LpFrom = LpFrom + p + v
                LnTo = LnTo + Nnew + v: LpTo = LpTo + Replace(p, N, Nnew) + v: Exit Sub
            End If
    'Typ4:  T1inT2, T1 2x in N, T2 2x in N
            If T1inT2 And C2 > 0 And C4 > 0 Then LnNo = LnNo + N + v
End Sub

Sub T6_ChangeEFN_Show(LnNo$, LnFrom$, LnTo$, LpFrom$, LpTo$, String1$, String2$)
    'Called from    T6_ChangeEFN_Start
    'LnNo           Liste FolderNames, NoChange
    'LnFrom, LnTo   Liste FolderNames, Change from/to   (OldName to NewName)
    'Status         LnNo oder LnFrom können je leer sein, nicht beide
    
    'Vorbereitung
        Dim g2$, g4$, g5$, g6$, g8$, N$, qq$, s$, s3$, v$
        Dim C1%, C2%, CharsMax%, cRows%, i%, zEndGreen%, zStartGreen%, zString2InputField%
        Dim zLast&, A() As String, B() As String, c() As String, D() As String, r As Range
        v = vbCrLf: qq = Chr(34) 'qq="=DoubleQuotes
        zLast = LastRow(Sheets("T6"))
        LnNo = Delete_EmptyEndRowsInString(LnNo)
        LnFrom = Delete_EmptyEndRowsInString(LnFrom)
        LnTo = Delete_EmptyEndRowsInString(LnTo)
    'Find RowNr of CellBtn " Change this string to:"
        zString2InputField = Get_RowNr_HoldingMyTextWhole("T6", " Change String   xxx")
    'c1, c2 (Anzahl Zeilen LnNo, LnFrom)
        If LnNo = "" Then C1 = 0 Else C1 = anzAinB(v, LnNo) + 1
        If LnFrom = "" Then C2 = 0 Else C2 = anzAinB(v, LnFrom) + 1
    'Old GreenShowNamesArea löschen
        Range(Cells(1, 22), Cells(1, 33)).EntireColumn.Hidden = True
        Set r = Range(Cells(5, 24), Cells(zLast, 33)): r.MergeCells = False: r.Clear
    'Old SaveT6RangeOfNames löschen
        [AF2] = ""
    '(1)    c1>0, c2=0 (nur LnNo enthält Zeilen)
        If C2 = 0 Then
            zStartGreen = zString2InputField - C1 - 13 '(T6-ZeilenNr, bei der GreenShowNamesArea startet)
            If zStartGreen < 5 Then zStartGreen = 5
            zEndGreen = zStartGreen + C1 + 14
            'Green Area
                Range(Cells(zStartGreen, 24), Cells(zEndGreen, 29)).Interior.Color = Green1
            'T6-ZusatzSpalten einblenden, slow
                T6_ChangeSpaltenEinblenden
'                For i = 22 To 29
'                    Cells(1, i).EntireColumn.Hidden = False: Sleep 50: DoEvents
'                Next
            'GreenRow 2 (Title)
                Set r = Range(Cells(zStartGreen + 1, 25), Cells(zStartGreen + 1, 28))
                r.MergeCells = True: r.HorizontalAlignment = xlCenter: r.Font.Bold = True
                r.Value = "Change  " + qq + String1 + qq + "  to  " + qq + String2 + qq + "  in EventFolderNames" 'Title
            'GreenRow 4
                Set r = Range(Cells(zStartGreen + 3, 25), Cells(zStartGreen + 3, 28))
                r.MergeCells = True: r.HorizontalAlignment = xlCenter
                r.Value = CStr(C1) + " Namen von Event-Ordnern, enthalten den String " _
                 + qq + String1 + qq + "." 'GreenShowNamesArea-Zeile Nr 4
            'GreenRow 5
                Set r = Range(Cells(zStartGreen + 4, 25), Cells(zStartGreen + 4, 28))
                r.MergeCells = True: r.HorizontalAlignment = xlCenter
                r.Value = qq + String1 + qq + "  wurde jedoch bereits durch  " + qq + String2 + qq + "  ersetzt."
            'GreenRow 6
                Set r = Range(Cells(zStartGreen + 5, 25), Cells(zStartGreen + 5, 28))
                r.MergeCells = True: r.HorizontalAlignment = xlCenter
                r.Value = "Es gibt nichts zu tun."
            'GreenRow 8 border top
                Set r = Range(Cells(zStartGreen + 7, 25), Cells(zStartGreen + 7, 28))
                r.Borders(xlEdgeTop).Weight = xlThin
            'GreenRow 9
                Cells(zStartGreen + 8, 25).Value = CStr(C1) + " Ordnernamen, die den String " + qq + String1 + qq + " enthalten:"
            'GreenRow 11ff (Liste LnNo)
                A = Split(v + LnNo, v)
                For i = 1 To UBound(A)
                    Cells(zStartGreen + 9 + i, 25).Value = "[" + Format(i, "00") + "]"
                    Cells(zStartGreen + 9 + i, 26).Value = A(i)
                Next
            'Save T6RangeOfNames
                [AF2] = "|" + Format(zStartGreen + 10, "000") + "|" + Format(C1, "000") + "|"
            'GreenRow (Last - 3) border top
                Set r = Range(Cells(zEndGreen - 3, 25), Cells(zEndGreen - 3, 28))
                r.Borders(xlEdgeTop).Weight = xlThin
            'Button2:
                Set r = Range(Cells(zEndGreen - 2, 28), Cells(zEndGreen - 1, 28))
                r.Interior.Color = Green3: r.MergeCells = True: r.Value = "OK"
                r.HorizontalAlignment = xlCenter:   r.Font.size = 24
                r.VerticalAlignment = xlCenter:     r.Font.Bold = True
            'Zeilenhöhe
                Sheets("T6").UsedRange.RowHeight = 12.75
            'Spaltenbreite Sp26, Sp28
                Cells(1, 26).EntireColumn.ColumnWidth = 23
                Cells(1, 28).EntireColumn.ColumnWidth = 23
            'T6-Spalten einblenden
                Range(Cells(1, 22), Cells(1, 29)).EntireColumn.Hidden = False
                ActiveWindow.ScrollRow = 1
                ActiveWindow.ScrollColumn = 3 'Spalte, die ganz links zu sehen sein soll
            Exit Sub
        End If
        
    '(2)    c1=0, c2>0   (nur LnFrom enthält Zeilen)
        If C1 = 0 And C2 > 0 Then
            zStartGreen = zString2InputField - C2 - 13 '(T6-ZeilenNr, bei der GreenShowNamesArea startet)
            If zStartGreen < 5 Then zStartGreen = 5
            zEndGreen = zStartGreen + C2 + 12
            'Green Area
                Range(Cells(zStartGreen, 24), Cells(zEndGreen, 29)).Interior.Color = Green1
            'T6-ZusatzSpalten einblenden, slow
                T6_ChangeSpaltenEinblenden
'                For i = 22 To 29
'                    Cells(1, i).EntireColumn.Hidden = False: Sleep 50: DoEvents
'                Next
            'GreenRow 2 (Title)
                Set r = Range(Cells(zStartGreen + 1, 25), Cells(zStartGreen + 1, 28))
                r.MergeCells = True: r.HorizontalAlignment = xlCenter: r.Font.Bold = True
                r.Value = "Change  " + qq + String1 + qq + "  to  " + qq + String2 + qq + "  in EventFolderNames" 'Title
            'GreenRow 4
                Set r = Range(Cells(zStartGreen + 3, 25), Cells(zStartGreen + 3, 28))
                r.MergeCells = True: r.HorizontalAlignment = xlCenter
                r.Value = CStr(C2) + " Namen von Event-Ordnern, enthalten den String " _
                 + qq + String1 + qq + "." 'GreenShowNamesArea-Zeile Nr 4
            'GreenRow 6 border top
                Set r = Range(Cells(zStartGreen + 5, 25), Cells(zStartGreen + 5, 28))
                r.Borders(xlEdgeTop).Weight = xlThin
            'GreenRow 7
                Cells(zStartGreen + 6, 25).Value = CStr(C2) + " Ordnernamen werden wie folgt geändert:"
            'GreenRow 9ff (Listen LnFrom, LnTo, LpFrom, LpTo)
                A = Split(v + LnFrom, v): B = Split(v + LnTo, v)
                c = Split(v + LpFrom, v): D = Split(v + LpTo, v)
                For i = 1 To UBound(A)
                    Cells(zStartGreen + 7 + i, 26).Value = A(i) 'LnFrom
                    Cells(zStartGreen + 7 + i, 28).Value = B(i) 'LnTo
                    Cells(zStartGreen + 7 + i, 31).Value = c(i) 'LpFrom
                    Cells(zStartGreen + 7 + i, 32).Value = D(i) 'LpTo
                    Cells(zStartGreen + 7 + i, 25).Value = "[" + Format(i, "00") + "]"
                    Cells(zStartGreen + 7 + i, 29).Value = " "
                    Cells(zStartGreen + 7 + i, 33).Value = "-"
                    Set r = Cells(zStartGreen + 7 + i, 27) 'kleiner RechtsPfeil
                    r.Value = qq: r.Font.NAME = "Wingdings 3": r.Font.size = 12
                    r.HorizontalAlignment = xlCenter: r.VerticalAlignment = xlCenter
                Next
            'Save T6RangeOfNames
                [AF2] = "|" + Format(zStartGreen + 8, "000") + "|" + Format(C2, "000") + "|"
            'GreenRow (Last - 3) border top
                Set r = Range(Cells(zEndGreen - 3, 25), Cells(zEndGreen - 3, 28))
                r.Borders(xlEdgeTop).Weight = xlThin
            'Button1:
                Set r = Range(Cells(zEndGreen - 2, 25), Cells(zEndGreen - 1, 26))
                r.Interior.Color = Green3: r.MergeCells = True: r.Value = "OK"
                r.HorizontalAlignment = xlCenter:   r.Font.size = 24
                r.VerticalAlignment = xlCenter:     r.Font.Bold = True
            'Button2:
                Set r = Range(Cells(zEndGreen - 2, 28), Cells(zEndGreen - 1, 28))
                r.Interior.Color = Green3: r.MergeCells = True: r.Value = "Abbrechen"
                r.HorizontalAlignment = xlCenter:   r.Font.size = 24
                r.VerticalAlignment = xlCenter:     r.Font.Bold = True
            'Zeilenhöhe
                Sheets("T6").UsedRange.RowHeight = 12.75
            'Spaltenbreite Sp26, Sp28
                Cells(1, 26).EntireColumn.ColumnWidth = 30
                Cells(1, 28).EntireColumn.ColumnWidth = 30
            'T6-Spalten einblenden
                ActiveWindow.ScrollColumn = 9 'Spalte, die ganz links zu sehen sein soll
            Exit Sub
        End If

    '(3)    c1>0, c2>0 (beide Listen, LnNo und LnFrom, enthalten Zeilen)
        If C1 > 0 And C2 > 0 Then
            zStartGreen = zString2InputField - C1 - C2 - 13 '(T6-ZeilenNr, bei der GreenShowNamesArea startet)
            If zStartGreen < 5 Then zStartGreen = 5
            zEndGreen = zStartGreen + C1 + C2 + 18
            'Green Area
                Range(Cells(zStartGreen, 24), Cells(zEndGreen, 29)).Interior.Color = Green1
            'T6-ZusatzSpalten einblenden, slow
                T6_ChangeSpaltenEinblenden
            'GreenRow 2 (Title)
                Set r = Range(Cells(zStartGreen + 1, 25), Cells(zStartGreen + 1, 28))
                r.MergeCells = True: r.HorizontalAlignment = xlCenter: r.Font.Bold = True
                r.Value = "Change  " + qq + String1 + qq + "  to  " + qq + String2 + qq + "  in EventFolderNames" 'Title
            'GreenRow 4
                Set r = Range(Cells(zStartGreen + 3, 25), Cells(zStartGreen + 3, 28))
                r.MergeCells = True: r.HorizontalAlignment = xlCenter
                r.Value = CStr(C1 + C2) + " Namen von Event-Ordnern, enthalten den String " _
                 + qq + String1 + qq + ";" 'GreenShowNamesArea-Zeile Nr 4
            'GreenRow 5
                Set r = Range(Cells(zStartGreen + 4, 25), Cells(zStartGreen + 4, 28))
                r.MergeCells = True: r.HorizontalAlignment = xlCenter
                r.Value = "davon werden " + CStr(C1) + " Ordnernamen nicht geändert;"
            'GreenRow 6
                Set r = Range(Cells(zStartGreen + 5, 25), Cells(zStartGreen + 5, 28))
                r.MergeCells = True: r.HorizontalAlignment = xlCenter
                r.Value = "die verbleibenden " + CStr(C2) + " Ordnernamen können geändert werden."
            'GreenRow 8 border top
                Set r = Range(Cells(zStartGreen + 7, 25), Cells(zStartGreen + 7, 28))
                r.Borders(xlEdgeTop).Weight = xlThin
            'GreenRow 9
                Cells(zStartGreen + 8, 25).Value = "Die " + CStr(C1) + " Ordnernamen, " _
                + "die nicht geändert werden (enthalten bereits " + qq + String2 + qq + "):"
            'GreenRow 11ff (Liste LnNo)
                A = Split(v + LnNo, v)
                For i = 1 To UBound(A)
                    Cells(zStartGreen + 9 + i, 25).Value = "[" + Format(i, "00") + "]"
                    Cells(zStartGreen + 9 + i, 26).Value = A(i)
                Next
            'GreenRow 12+c1 border top
                Set r = Range(Cells(zStartGreen + 11 + C1, 25), Cells(zStartGreen + 11 + C1, 28))
                r.Borders(xlEdgeTop).Weight = xlThin
            'GreenRow 13+c1
                Cells(zStartGreen + 12 + C1, 25).Value = "Die " + CStr(C2) + " Ordnernamen, " _
                + "die wie folgt geändert werden:"
            'GreenRow(15+c1)ff (Listen LnFrom, LnTo, LpFrom, LpTo)
                A = Split(v + LnFrom, v): B = Split(v + LnTo, v)
                c = Split(v + LpFrom, v): D = Split(v + LpTo, v)
                For i = 1 To UBound(A)
                    Cells(zStartGreen + 13 + C1 + i, 26).Value = A(i) 'LnFrom
                    Cells(zStartGreen + 13 + C1 + i, 28).Value = B(i) 'LnTo
                    Cells(zStartGreen + 13 + C1 + i, 31).Value = c(i) 'LpFrom
                    Cells(zStartGreen + 13 + C1 + i, 32).Value = D(i) 'LpTo
                    Cells(zStartGreen + 13 + C1 + i, 25).Value = "[" + Format(i, "00") + "]"
                    Cells(zStartGreen + 13 + C1 + i, 29).Value = " "
                    Cells(zStartGreen + 13 + C1 + i, 33).Value = "-"
                    Set r = Cells(zStartGreen + 13 + C1 + i, 27) 'kleiner RechtsPfeil
                    r.Value = qq: r.Font.NAME = "Wingdings 3": r.Font.size = 12
                    r.HorizontalAlignment = xlCenter: r.VerticalAlignment = xlCenter
                Next
            'Save T6RangeOfNames
                [AF2] = "|" + Format(zStartGreen + 14 + C1, "000") + "|" + Format(C2, "000") + "|"
            'GreenRow(Last - 3) border top
                Set r = Range(Cells(zEndGreen - 3, 25), Cells(zEndGreen - 3, 28))
                r.Borders(xlEdgeTop).Weight = xlThin
            'Button1:
                Set r = Range(Cells(zEndGreen - 2, 25), Cells(zEndGreen - 1, 26))
                r.Interior.Color = Green3: r.MergeCells = True: r.Value = "OK"
                r.HorizontalAlignment = xlCenter:   r.Font.size = 24
                r.VerticalAlignment = xlCenter:     r.Font.Bold = True
            'Button2:
                Set r = Range(Cells(zEndGreen - 2, 28), Cells(zEndGreen - 1, 28))
                r.Interior.Color = Green3: r.MergeCells = True: r.Value = "Abbrechen"
                r.HorizontalAlignment = xlCenter:   r.Font.size = 24
                r.VerticalAlignment = xlCenter:     r.Font.Bold = True
            'Zeilenhöhe
                Sheets("T6").UsedRange.RowHeight = 12.75
            'Spaltenbreite Sp26, Sp28
                Cells(1, 26).EntireColumn.ColumnWidth = 30
                Cells(1, 28).EntireColumn.ColumnWidth = 30
            'T6-Spalten einblenden
                ActiveWindow.ScrollColumn = 9 'Spalte, die ganz links zu sehen sein soll
            Exit Sub
        End If
End Sub

Sub T6_ChangeEFN_Start(zT6%)
    'Called from    [UserClick on CellButton "Change String ..."] T6_UserClickOnColumn2CellButton
    'zT6            ZeilenNr "Change String xxx to ..."
    'Änderungen     werden nur auf dem Speichermedium vorgenommen
    'Action         ...
    
    'Vorbereitung
        Dim LnNo$, LnFrom$, LnTo$, LpFrom$, LpTo$, N$, p$, s$, s3$, String1$, String2$, Title$, v$
        Dim i%, iMsgBox%, Arr1() As String, r As Range
        v = vbCrLf: With Sheets("T6"): EE 0: Ticks1
        FillArrC 30, zT6 'CountDownShowCell-ZeilenNr
    'CellButton Green4
        .Cells(zT6, 1).Select 'Selection wechseln
        Set r = Union(.Cells(zT6, 2), .Cells(zT6 + 1, 2), .Cells(zT6, 10), .Cells(zT6, 15)): r.Interior.Color = Green4
    'Get String1, String2
        String1 = .Cells(zT6, 7): String2 = .Cells(zT6, 12)
        Title = "Change '" + String1 + "' to '" + String2 + "' in EventFolderNames"
    'Exit
        If String1 = "" Then show "Change '" + String1 + "' to '" + String2 _
            + "' in EventFolderNames:" + v + v + "Leerer SuchString." + v _
            + "Es gibt nichts zu tun.": GoTo jump1
        If String1 = String2 Then show "Change '" + String1 + "' to '" + String2 _
            + "' in EventFolderNames:" + v + v + "SuchString = ErsetzungsString." + v _
            + "Es gibt nichts zu tun.": GoTo jump1
    'DateEventFolders
        'EventFolder, die das EventDatum tragen; Folder mit resultJPGs, JPGs, ...
        s = Get_PathsOfAllDateEventFolders 'nur existierende Ordner
    'Exit
        If Not s Like "*" + String1 + "*" Then show _
            "Change '" + String1 + "' to '" + String2 + "' in EventFolderNames:" + v + v _
            + "Der String '" + String1 + "' ist in keinem der" + v + "Ordnernamen in 'Events' enthalten." + v _
            + "Es gibt nichts zu tun.": GoTo jump1
    'Arr1
        Arr1 = Split(v + s, v)
    'LnNo, LnFrom, LnTo, LpFrom, LpTo
        For i = 1 To UBound(Arr1)
            p = Arr1(i)          'path of one DateEventFolder
            N = getNameOfPath(p)             'FolderName
            If N Like "*" + String1 + "*" Then
                T6_Analyze_ChangeEventFolderName p, N, String1, String2, LnNo, LnFrom, LnTo, LpFrom, LpTo
            End If
        Next
    'Show in Sheet T6
        T6_ChangeEFN_Show LnNo, LnFrom, LnTo, LpFrom, LpTo, String1, String2
        [AF3] = T6_GetDuration()
jump1:
    'CellButton Green3
        r.Interior.Color = Green3
    'Finals
        'show pList1
        End With: EE 1
End Sub

Function Get_ListOfNonEmptyEntriesOfOneColumn(NameOfSheet$, z1%, s%) As String
    Dim L$, i%, zLast%, A() As String
    zLast = Get_NrOfLastRowInColumnNr(CLng(s), NameOfSheet)
    Load_Array1D_OfColumnPart A, NameOfSheet, z1, zLast, s
    For i = LBound(A) To UBound(A)
        If Trim(A(i)) <> "" Then L = L + A(i) + vbCrLf
    Next
    L = Delete_EmptyEndRowsInString(L)
    Get_ListOfNonEmptyEntriesOfOneColumn = L
End Function

Sub T6_RenameFolders(T, txt1$, txt2$, Nr%, Rep$, DoAllChanges As Boolean)
    'Called from    T6_Check_EventFolderNames
    'T              2D-Array im T9-Raster, anfangs leer
    'Rename         efolgt in SheetT9(Name,Path), in T [T(i,9),T(i,10)], in Volume
    
    'Vorbereitung
        Dim N1$, N2$, p1$, p2$, Report$, v$, i%, zT9%
        v = vbCrLf: Rep = "": With Sheets("T9")
    'Action
        Report = "(" + CStr(Nr) + ") Schreibweise von Ordnernamen: '" + txt1 + "' --> " + "'" + txt2 + "'" + v
        For i = 1 To UBound(T, 1)
            N1 = CStr(T(i, 9))      'T9-FolderName
            p1 = CStr(T(i, 10))     'T9-FolderPath
            If N1 Like "*" + txt1 + "*" Then
                '"." --> ". " soll nicht erneut stattfinden (txt1 immernoch vorhanden)
                If InStr(1, txt2, txt1) = 0 Then
                    N2 = Replace(N1, txt1, txt2)
                    p2 = Replace(p1, txt1, txt2)
                    If DoAllChanges Then
                        'Rename in T9
                            zT9 = Get_RowNr_HoldingMyTextWholeInColumnX("T9", 10, 12, N1)
                            .Cells(zT9, 10).Select
                            .Cells(zT9, 10) = N2: .Cells(zT9, 11) = p2
                        'Rename in T()
                            T(i, 9) = N2: T(i, 10) = p2
                        'Rename in Volume
                            RenameFolder p1, p2
                    End If
                    Rep = Rep + "    " + N1 + " --> " + N2 + v
                End If
            End If
        Next
        If Rep = "" Then Report = Report + "    Keine Änderung nötig." + v Else Report = Report + "    Renamings:" + v + Rep
    'Finals
        Rep = Report: End With
End Sub




