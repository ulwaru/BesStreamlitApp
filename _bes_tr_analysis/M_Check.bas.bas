Attribute VB_Name = "M_Check"
Option Explicit 'M_Check

Sub zzz_M_Check()
    
    'showProcs "check fold"

    'RenameModule "Tabelle1", "T6"
    'RenameModule "Modul1", "M_Check2"
    'T1_ShowLogBuch
    'showArray2D ArrPd
    'T1_ShowPersonDataTxt
    EE 1: Beep
End Sub

Sub OpenFolderGrosseSymbole_TEST()
    Dim p$
    p = "F:\Archiv Trampolin 1900-1999\Events\1959-09-05 Nissen-Cup2 Wasen_CH"
    OpenFolderGrosseSymbole p
End Sub

Sub OpenFolderGrosseSymbole(ByVal ordnerPfad$)
    Dim oShell As Object, oFolder As Object, oWindows As Object, oWin As Object
    Set oShell = CreateObject("Shell.Application")
    Set oWindows = oShell.Windows
    'Öffnen Sie den Ordner im Explorer, falls er noch nicht offen ist
        'Dies stellt sicher, dass wir eine Window-Instanz zum Manipulieren haben
        OpenFolder ordnerPfad
    'Durchsuche alle offenen Explorer-Fenster
        For Each oWin In oWindows
            On Error Resume Next ' Fehler bei nicht-Explorer-Fenstern ignorieren
            Set oFolder = oWin.Document
            On Error GoTo 0
            'Überprüfen, ob das aktuelle Fenster der gesuchte Ordner ist
                If Not oFolder Is Nothing Then
                    If LCase(oFolder.folder.Self.path) = LCase(ordnerPfad) Then
                        'Setze den ViewMode auf 'Große Symbole' (Wert 5)
                        oFolder.CurrentViewMode = 5 ' Oder 8 für 'Sehr Große Symbole'
                        oFolder.IconSize = 96 ' Beispiel für eine gängige Große-Symbole-Größe
                        Set oFolder = Nothing: Set oWin = Nothing
                        Exit For ' Fenster gefunden und geändert, Schleife beenden
                    End If
                End If
            Set oFolder = Nothing
        Next oWin
    Set oWindows = Nothing: Set oShell = Nothing
    With Application.VBE.MainWindow: .SetFocus: .Visible = True: End With
End Sub


Function T4_Get_DgFolderPath(DgFolderName$) As String
    Dim s1%, z1%
    If DgFolderName = "-" Then
        T4_Get_DgFolderPath = "-"
        With Sheets("T4").Cells(z1, s1 + 1) 'Hinweis oberhalb des DgTitle
            .Font.size = 8: .Font.NAME = "Arial": .HorizontalAlignment = xlLeft: .Value = "-"
        End With
    Else
        T4_Get_DgFolderPath = ArrC(3) + "\" + DgFolderName
    End If
End Function

Sub Update_DgData()
    'Called from    [none]
    'Darstellung    auf Blatt T5
    'Duration       0,3 sec
    'Action

    'Vorbereitung
        Dim AllTitles$, DgChars$, DgFolder$, DgId$, DgTitle$, F$, L$
        Dim p$, t1$, T2$, T3$, ToDoNoFold$, v$
        Dim c%, C1%, C2%, i%, j%, s0%, s1%, s1Arr%, s2%, z0%, z1%, z1Arr%, z2%
        Dim SpNr&, tic&, ZeNr&, w As Worksheet, Arr(), ArrDg(), ArrDgData(), Arr1() As String
        v = vbCrLf: DoArrc: p = ArrC(3) + "\": Set w = Sheets("T4")
        ReDim ArrDgData(1 To 80, 1 To 12): tic = GetTickCount()
    'Einstellungen
        z1Arr = 8: s1Arr = 25 'StartZelle Arr (= T4-Bereich mit allen Designs)
    'Gesamten T4-Dg-Bereich in Array nehmen
        Arr = w.Range(w.Cells(z1Arr, s1Arr), w.Cells(LastRow(Sheets("T4")) + 2, lastCol(Sheets("T4"))))
        z0 = z1Arr - 1: s0 = s1Arr - 1
        
    For ZeNr = 1 To UBound(Arr, 1) - 2
        'Schleife über alle Zeilen des Arrays      (des T4-Tabellenteils ab T4-ArrStartZelle)
        For SpNr = 1 To UBound(Arr, 2)
            'Schleife über alle Spalten des Arrays (des T4-Tabellenteils ab T4-ArrStartZelle)
            If Arr(ZeNr, SpNr) Like "19##*" And Not Arr(ZeNr + 1, SpNr) Like "19##*" Then
                'Ein Design wurde gefunden; Title steht in Arr(ZeNr, SpNr)
                c = c + 1:  ArrDgData(c, 4) = "-":  ArrDgData(c, 12) = "-"
                'DgTitle
                    DgTitle = Arr(ZeNr, SpNr)
                    ArrDgData(c, 2) = DgTitle
                    AllTitles = AllTitles + DgTitle + v
                'Dg-Positions z1, z2, s1, s2
                    If Arr(ZeNr - 1, SpNr - 1) = "LO" Then z1 = z0 + ZeNr - 1: s1 = s0 + SpNr - 1 Else Stop
                    For i = 1 To 99
                        If Arr(ZeNr - 1, SpNr + i) = "RO" Then s2 = s0 + SpNr + i: Exit For
                    Next
                    For i = 1 To 999
                        If Arr(ZeNr + i, SpNr - 1) = "LU" Then z2 = z0 + ZeNr + i: Exit For
                    Next
                    ArrDgData(c, 6) = z1: ArrDgData(c, 7) = z2
                    ArrDgData(c, 8) = s1: ArrDgData(c, 9) = s2
                'DgRange
                    ArrDgData(c, 5) = Get_ColumnLetter(s1) + CStr(z1) + ":" + Get_ColumnLetter(s2) + CStr(z2)
                'Dg-Columns
                    ArrDgData(c, 4) = s2 - s1 + 1
                'Dg-ID
                    If Arr(ZeNr - 1, s2 - s0 - 1) Like "Dg####" Then
                        'Im T4-Design trägt die DgId-Zelle bereits eine korrekte DgId
                        DgId = Arr(ZeNr - 1, s2 - s0 - 1)
                        ArrDgData(c, 1) = DgId
                    End If
                'DgFolder
                    DgFolder = Arr(ZeNr - 1, SpNr)
                    ArrDgData(c, 3) = DgFolder
                    'DgFolder soll nur in ArrDgData landen, falls Pfad existiert
                    If DgFolder = "" Then
                        'Folder=Title?
                        F = Replace(DgTitle, "/", "_")
                        If FolderExists(p + F) Then
                            Put_EventFolderNameToDg z1, s1 + 1, F       'Write to Sheet
                            ArrDgData(c, 3) = F
                        Else 'Liga-Folder?
                            L = Left(F, 4) + " Liga\" + F
                            If FolderExists(p + L) Then
                                ArrDgData(c, 3) = L
                                Put_EventFolderNameToDg z1, s1 + 1, L   'Write to Sheet
                            Else
                                ToDoNoFold = ToDoNoFold + "DgTitle=[" + DgTitle + "]" + v
                            End If
                        End If
                    Else
                        If Not FolderExists(p + DgFolder) Then
                            ToDoNoFold = ToDoNoFold + "DgTitle=[" + DgTitle + "]; DgFolder=[" + DgFolder + "]" + v
                            ArrDgData(c, 3) = ""
                        End If
                    End If

            End If
        Next
    Next
    'Row8-Dg-IDs - ggf. neue Dg-IDs für Row8Dgs vergeben
        For i = 1 To UBound(ArrDgData, 1)
            If ArrDgData(i, 6) = 8 Then
                C1 = C1 + 1         'c1 = SOLL-Anzahl DgGroups
                If Len(ArrDgData(i, 1)) = 6 Then
                    C2 = C2 + 1     'c2 =  IST-Anzahl DgGroups
                    T2 = T2 + Mid(ArrDgData(i, 1), 3, 1) + "|"
                End If
            End If
        Next
        't2 = alle existierenden GroupNrn = "01|02|...|" (aus Dg-IDs '"Dg0101|Dg0201|...|" oder "")
        If C1 <> C2 Then
            'Es gibt Designs, die in Zeile 8 beginnen und keine ID haben
            't1 bereitstellen; Kandidaten für GroupNr
                For i = 1 To C1
                    t1 = t1 + Format(i, "00") + "," 't1 = "01,02,03,...,"
                Next
            't2 (bereits vergebene GroupNrn) aus t1 herausnehmen
                If T2 <> "" Then                    't2 = "01|02|03|...|"
                    Arr1 = Split(T2, "|")
                    For i = 0 To UBound(Arr1) - 1
                        t1 = Replace(t1, Arr1(i) + ",", "")
                    Next
                End If
            'Neue Dg-IDs vergeben
                For i = 1 To UBound(ArrDgData, 1)
                    If ArrDgData(i, 6) = 8 Then
                        If Len(ArrDgData(i, 1)) <> 6 Then
                            ArrDgData(i, 1) = "Dg" + Left(t1, 2) + "01"
                            t1 = Mid(t1, 4)
                        End If
                    End If
                Next
        End If
    'VnNn Dg-IDs - ggf. neue Dg-IDs vergeben
        t1 = "": T2 = "": C1 = 0: C2 = 0
        'Alle Gruppenmitglieder haben den gleichen s1-Wert, verschiedene z1-Werte
        't1-Liste erzeugen (Zeilen wie "|s1=|0025|z1=|0008|ID=|Dg0101")
            For i = 1 To UBound(ArrDgData, 1)
                If ArrDgData(i, 8) > 0 Then 'keine Leerzeile
                    t1 = t1 + "|s1=|" + Format(ArrDgData(i, 8), "0000") + "|z1=|" _
                        + Format(ArrDgData(i, 6), "0000") + "|ID=|" + ArrDgData(i, 1) + v
                End If
            Next
            t1 = Delete_EmptyEndRowsInString(t1)            't1 = "|s1=|0025|z1=|0008|ID=|Dg0101" ...
        'Arr1 (t1-Liste sortieren; GruppenMitglieder stehen dann untereinender)
            Arr1 = Split(t1, v)
            QuickSort Arr1 'show Join(Arr1, v)
        'Arr1 (Zeilen ohne Dg-ID erhalten eine Dg-ID)
            For i = 0 To UBound(Arr1)
                If Len(Arr1(i)) = 29 Then
                    T2 = Mid(Arr1(i), 24, 4)                't2 = "Dg01"
                Else
                    'Dg-ID in Arr1 eintragen
                    For j = 2 To 99
                        DgId = T2 + Format(j, "00")         'DgId = Dg0102 = Test-ID
                        If Not t1 Like "*" + DgId + "*" Then
                            Arr1(i) = Arr1(i) + DgId: t1 = t1 + v + DgId: j = 99
                        End If
                    Next
                End If
            Next
            t1 = Join(Arr1, v) 'In t1 haben jetzt alle Zeilen eine Dg-ID
        'Dg-IDs in ArrDgData eintragen
            For i = 1 To UBound(ArrDgData, 1)
                If ArrDgData(i, 2) <> "" Then
                    If Len(ArrDgData(i, 1)) <> 6 Then
                        T2 = "|s1=|" + Format(ArrDgData(i, 8), "0000") _
                           + "|z1=|" + Format(ArrDgData(i, 6), "0000") + "|"
                        C1 = InStr(1, t1, T2)
                        DgId = Mid(t1, C1 + 23, 6)
                        ArrDgData(i, 1) = DgId
                    End If
                End If
            Next
        'Dg-IDs in T4 eintragen
            For i = 1 To UBound(ArrDgData, 1)
                If ArrDgData(i, 2) <> "" Then   'keine Leerzeile
                    'Write once to T4       (z1,s1-1)=ID
                    Put_DgIdToDg CInt(ArrDgData(i, 6)), CInt(ArrDgData(i, 8)) - 1, CStr(ArrDgData(i, 1))
                End If
            Next
    'DgChars
        For i = 1 To UBound(ArrDgData, 1)
            If ArrDgData(i, 5) <> "" Then
                ArrDg = Sheets("T4").Range(ArrDgData(i, 5)).Value
                DgChars = Get_DgChars_GiveArrDg(ArrDg)
                ArrDgData(i, 10) = DgChars
                ArrDgData(i, 11) = Len(DgChars)
            End If
        Next
    'Paste
        Paste_2DArrayToCell_z_s "T5", 7, 4, ArrDgData
    'Update ListOfDesignTitles
        T4_Sort_Paste_Format_ListOfDesignTitles AllTitles
    'ToDoNoFold-Meldung
        If ToDoNoFold <> "" Then show "Makro 'ShowSomeItemsOfAllDgs'" + v + v + "EventOrdner wurde nicht gefunden:" + v + v + ToDoNoFold
    'Finals
        'ShowTime tic
        EE 1
        Beep
End Sub

Sub Put_DgIdToDg(z%, s%, DgId$)
    With Sheets("T4").Cells(z, s)
        .Value = DgId: .Font.Color = vbWhite
    End With
End Sub

Sub Put_EventFolderNameToDg(z%, s%, folder$)
    With Sheets("T4").Cells(z, s)
        .Value = folder: .Font.Color = vbWhite
    End With
End Sub

Sub Check_MyDgGroup()
    'Called from    Check_DgId
    'Action         Alle Dg der aktuellen Gruppe sollten eine DgId mit gleicher GroupNr haben
    
    'Vorbereitung
        Dim DgId$, DgNr$, AllDgIDs$, G$, GroupNr$
        Dim i%, j%, s2%, z2%, IdSpalte&, Arr1(), r As Range
        s2 = CInt(ArrC(40))
        IdSpalte = s2 - 1
        With Sheets("T4")
    'T4-IdSpalte + letzte Spalte von Zeile 1 bis Zeile z2 wird in arr1 eingelesen
        z2 = Get_NrOfLastRowInColumnNr(CLng(s2), "T4")
        Arr1 = .Range(.Cells(1, IdSpalte), .Cells(z2, s2)).Value
        'showArray2D arr1
    'Sammeln der AllDgIDs dieser DgGroup
        For i = 1 To UBound(Arr1, 1)
            DgId = Arr1(i, 1)
            If DgId Like "DgID######" Then
                G = Mid(DgId, 5, 3)
                If GroupNr = "" Then GroupNr = G
                If GroupNr <> G Then Beep: Stop
                AllDgIDs = AllDgIDs + DgId + "|"
            End If
        Next
        'AllDgIDs enthält alle bisher vergebenen DgID dieser DgGroup: '"DgID014001|...|"
    'Ggf neue AllDgIDs vergeben
        
        For i = 1 To UBound(Arr1, 1)
            If Arr1(i, 2) = "RO" Then
                Set r = .Cells(i, IdSpalte)
                
                DgId = r.Value
                If Not DgId Like "DgID######" Then
                    'Next free
                    For j = 1 To 99
                        DgNr = Format(j, "000")
                        If InStr(1, AllDgIDs, "DgID" + GroupNr + DgNr) = 0 Then
                            DgId = "DgID" + GroupNr + DgNr
                            'Neue DgID in aktuelles Dg schreiben
                                r.Value = DgId:  r.HorizontalAlignment = xlLeft
                                r.Font.size = 6: r.Font.Color = vbWhite
                                
                                
'                                r.Select
'                                Stop
                                
                                
                            AllDgIDs = AllDgIDs + DgId + "|"
                            j = 99 'Exit For j
                        End If
                    Next
                End If
            End If
        Next
        'Alle Row8-Dgs haben jetzt eine DgID
    


    End With
End Sub

Sub Check_DgId_OfT4Row8Dgs()
    'Called from    Check_DgId
    'Dg-ID          = "Dg0918" --> GroupNr = "09"; DgNr = "18"
    'Row8Dg         Das jeweils oberste Design (DgNr="01") beginnt in Zeile 8
    '               alle darunter platzierten Dgs haben den gleichen Aufbau
    '               und gehören deshalb der gleichen Gruppe an (gleiche GroupNr)
    'Status         Alle Row8-Dgs sind in den 4 Ecken mit LO, RO, LU bzw. RU beschriftet
    'Action         Alle Row8-Dgs bekommen eine Dg-ID ("Dg0101","Dg0201",...);
    '               diese wird in die Zelle links der 'RO'-Zelle geschrieben
    
    'Vorbereitung
        Dim AllGroupNrs$, DgId$, Nr$, i%, j%, s1%, s2%, z%, Arr1(), r As Range
        With Sheets("T4")
    'T4-Zeile z von Spalte s1 bis Spalte s2 wird in arr1 eingelesen
        z = 8    'T4-ZeilenNr der obersten Dg-Zeile des jew. obersten Dgs
        s1 = 17: s2 = lastCol(Sheets("T4"))
        Arr1 = .Range(.Cells(z, 17), .Cells(z, lastCol(Sheets("T4")))).Value
        'showArray2D arr1
    'Sammeln der Nummern aller bisher existierenden DgGroups
        For i = 1 To UBound(Arr1, 2)
            If Arr1(1, i) = "RO" Then
                DgId = CStr(.Cells(z, s1 + i - 2))
                If DgId Like "DgID######" Then
                    AllGroupNrs = AllGroupNrs + Mid(DgId, 5, 3) + "|"
                End If
            End If
        Next
        'AllGroupNrs enthält alle bisher vergebenen Nrn der DgGroups: '"001|002|...|"
    'Ggf neue DgIDs vergeben
        For i = 1 To UBound(Arr1, 2)
            If Arr1(1, i) = "RO" Then
                Set r = .Cells(z, s1 + i - 2)
                DgId = r.Value
                If Not DgId Like "DgID######" Then
                    'Next free
                    For j = 1 To 99
                        Nr = Format(j, "000")
                        If InStr(1, AllGroupNrs, Nr) = 0 Then
                            DgId = "DgID" + Nr + "001"
                            'Neue DgID in aktuelles Dg schreiben
                                r.Value = DgId:  r.HorizontalAlignment = xlLeft
                                r.Font.size = 6: r.Font.Color = vbWhite
                            AllGroupNrs = AllGroupNrs + Nr + "|"
                            j = 99 'Exit For j
                        End If
                    Next
                End If
            End If
        Next
        'Alle Row8-Dgs haben jetzt eine DgID
    End With
End Sub

Sub Check_DgId()
    'Called from    T4_SearchForBorders
    'Status         User klickte in ein Dg; Dg ist aktiviert
    'Dg-ID          = "Dg0918"; GroupNr = "09"; DgNr = "18"
    'Action         ...
    
    'Alle Dg mit z1=8 sollten eine DgId haben
        Check_DgId_OfT4Row8Dgs
    'Alle Dg der aktuellen Gruppe sollten eine DgId mit gleicher GroupNr haben
        Check_MyDgGroup
End Sub

Function Get_NextDgIdGrNr() As String
    'Called from    Check_DgGroupId
    'DgID           = "DgID001001", "DgID002001|...|", ...
    
    'Vorbereitung
        Dim Id$, AllGroupNrs$, Nr$, AfterCol%, i%, s%, sOld%
        AfterCol = 16
    'ID-Group-Nummern sammeln
        For i = 1 To 99
            s = Get_ColumnNr_HoldingMyTextPartInRowX("T4", 8, AfterCol, "DgID")
            If s = sOld Then Exit For
            If sOld = 0 Then sOld = s
            Id = Sheets("T4").Cells(8, s)                       '"DgID001001"
            AllGroupNrs = AllGroupNrs + Mid(Id, 5, 3) + "|"     '"001|002|...|"
            AfterCol = s + 1
        Next
    'Next free
        For i = 1 To 99
            Nr = Format(i, "000")
            If InStr(1, AllGroupNrs, Nr) = 0 Then Exit For
        Next
        show AllGroupNrs + " --> " + Nr
        
    Get_NextDgIdGrNr = Nr
    
End Function

Function Get_ColumnNr_HoldingMyTextPartInRowX(MySheet$, MyRow%, AfterColumnNr%, MyText$) As Integer
    'Called from    Check_DgGroupId
    'Action         ...
    
    Dim z%, r1 As Range
    z = MyRow
    With ThisWorkbook.Sheets(MySheet)
        Set r1 = .Rows(z).Find(What:=MyText, _
            After:=.Cells(z, AfterColumnNr), LookIn:=xlValues, LookAt:=xlPart, _
            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End With
    If r1 Is Nothing Then
        'MyText nicht zu finden; ggf. ausgeblendet
        Get_ColumnNr_HoldingMyTextPartInRowX = 0
    Else
        Get_ColumnNr_HoldingMyTextPartInRowX = CInt(r1.Column)
    End If
End Function

Function Get_RowNr_HoldingMyTextPartInColumnX(MySheet$, MyColumn%, AfterRowNr%, MyText$) As Integer
    'Called from    Check_DgGroupId
    'Action         ...
    
    Dim s%, r1 As Range
    s = MyColumn
    With ThisWorkbook.Sheets(MySheet)
        Set r1 = .Columns(s).Find(What:=MyText, _
            After:=.Cells(AfterRowNr, s), LookIn:=xlValues, LookAt:=xlPart, _
            SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    End With
    If r1 Is Nothing Then
        'MyText nicht zu finden; ggf. ausgeblendet
        Get_RowNr_HoldingMyTextPartInColumnX = 0
    Else
        Get_RowNr_HoldingMyTextPartInColumnX = CInt(r1.Row)
    End If
End Function

Function Get_ColumnNr_HoldingMyTextWholeInRowXSearchRight(MySheet$, MyRow%, AfterColumnNr%, MyText$) As Integer
    'Called from    Check_DgGroupId
    'Action         ...
    
    Dim z%, r1 As Range
    z = MyRow
    With ThisWorkbook.Sheets(MySheet)
        Set r1 = .Rows(z).Find(What:=MyText, _
            After:=.Cells(z, AfterColumnNr), LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End With
    If r1 Is Nothing Then
        'MyText nicht zu finden; ggf. ausgeblendet
        Get_ColumnNr_HoldingMyTextWholeInRowXSearchRight = 0
    Else
        Get_ColumnNr_HoldingMyTextWholeInRowXSearchRight = CInt(r1.Column)
    End If
End Function

Function Get_ColumnNr_HoldingMyTextWholeInRowXSearchLeft(MySheet$, RowX%, AfterColumnNr%, MyText$) As Integer
    'Called from    Check_DgGroupId
    'Action         ...
    
    Dim r1 As Range
    With ThisWorkbook.Sheets(MySheet)
        Set r1 = .Rows(RowX).Find(What:=MyText, _
            After:=.Cells(RowX, AfterColumnNr), LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    End With
    If r1 Is Nothing Then
        'MyText nicht zu finden; ggf. ausgeblendet
        Get_ColumnNr_HoldingMyTextWholeInRowXSearchLeft = 0
    Else
        Get_ColumnNr_HoldingMyTextWholeInRowXSearchLeft = CInt(r1.Column)
    End If
End Function

Function Get_RowNr_HoldingMyTextWholeInColumnXSearchUp(MySheet$, MyColumn%, AfterRowNr%, MyText$) As Integer
    'Called from    Check_DgGroupId
    'Action         ...
    
    Dim s%, r1 As Range
    s = MyColumn
    With ThisWorkbook.Sheets(MySheet)
        Set r1 = .Columns(s).Find(What:=MyText, _
            After:=.Cells(AfterRowNr, s), LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    End With
    If r1 Is Nothing Then
        'MyText nicht zu finden; ggf. ausgeblendet
        Get_RowNr_HoldingMyTextWholeInColumnXSearchUp = 0
    Else
        Get_RowNr_HoldingMyTextWholeInColumnXSearchUp = CInt(r1.Row)
    End If
End Function

Function Get_RowNr_HoldingMyTextWholeInColumnXSearchDown(MySheet$, ColumnX%, AfterRowNr%, MyText$) As Integer
    Get_RowNr_HoldingMyTextWholeInColumnXSearchDown = Get_RowNr_HoldingMyTextWholeInColumnX(MySheet, ColumnX, AfterRowNr, MyText)
End Function

Function Get_RowNr_HoldingMyTextWholeInColumnX(MySheet$, ColumnX%, AfterRowNr%, MyText$) As Integer
    'Called from    Check_DgGroupId
    'Action         ...
    
    Dim r1 As Range
    With ThisWorkbook.Sheets(MySheet)
        Set r1 = .Columns(ColumnX).Find(What:=MyText, _
            After:=.Cells(AfterRowNr, ColumnX), LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    End With
    If r1 Is Nothing Then
        'MyText nicht zu finden; ggf. ausgeblendet
        Get_RowNr_HoldingMyTextWholeInColumnX = 0
    Else
        Get_RowNr_HoldingMyTextWholeInColumnX = CInt(r1.Row)
    End If
End Function

Sub Check_DgFolder_forDgEdge()
    'Status     User klickte gerade in ein Design
    
    'Vorbereitung
        Dim p1$, p2$, p3$, Title$, i%, s1%, s2%, z1%, z2%
        Dim Arr1(), TitleIsInDgList As Boolean, Ecke As Range
        With ThisWorkbook.Sheets("T4"): DoArrc
        p3 = ArrC(3)               'p3 = Path of Folder "Events"
        z1 = CInt(ArrC(37)): z2 = CInt(ArrC(38)): s1 = CInt(ArrC(39)): s2 = CInt(ArrC(40))
        Set Ecke = ThisWorkbook.Sheets("T4").Cells(z1, s2)
        Ecke.Font.Color = vbWhite
    'PathOfDgFolder may be already in DgEdge
            p1 = .Cells(z1, s2)
            If p1 <> "" Then
                If FolderExists(p1) Then Exit Sub
            End If
    'Title
        Title = .Cells(z1 + 1, s1 + 1)
    'Format Edge
        .Cells(z1, s2).HorizontalAlignment = xlRight: .Cells(z1, s2).Font.size = 6
        .Cells(z1, s2).Font.Color = vbWhite
    'Path
        p1 = p3 + "\" + Title                               'p1 könnte PathOfDgFolder sein
        p2 = p3 + "\" + Left(Title, 4) + " Liga\" + Title   'p2 könnte PathOfDgFolder sein
        If FolderExists(p1) Then
            'DgFolder ist direkter   Unterordner von 'Events'
                                     Ecke = p1     'p1 wird in DgEckZelleReOb eingetragen
        ElseIf FolderExists(p2) Then Ecke = p2
        Else
            'Dg besitzt keinen DgFolder
            'Ggf. neues Design
                'ListOfDesignTitels einlesen
                Arr1 = .Range(.Cells(8, 2), .Cells(Get_NrOfLastRowInColumnNr(2, "T4"), 2)).Value
                    'showArray2D arr1
                'TitleIsInDgList?
                    TitleIsInDgList = False
                    For i = 1 To UBound(Arr1, 1)
                        If InStr(1, Arr1(i, 1), Title) > 0 Then TitleIsInDgList = True: Exit For
                    Next
                If TitleIsInDgList Then
                    'DgTitle ist bereits in DgList
                    'Ggf. Umbenennung eines EventFolders
                    Check_ChangeOfNameOfDgFolder
                Else
                    If MsgBox("Ein Ordner '" + Title + "' existiert nicht (im Ordner 'Events')." _
                        + vbCrLf + "Soll er angelegt werden?", vbYesNo) = 6 Then
                        MkDir p1: .Cells(z1, s2) = p1
                    End If
                    T4_Write_ListOfDesignTitles_FromDgs
                End If
        End If
    'Finals
        End With
End Sub

Sub Check_ChangeOfNameOfDgFolder()
    'Called from    Check_DgFolder_forDgEdge
    'DgFolder       = ein Ordner in 'Events' der zu einem Design gehört
    'Action         Schnellsuche aller DgFolder
    '               Check, ob zu irgendeinem Dg kein DgFolder existiert; Info falls ja

    'Vorbereitung
        Dim DgN$, Info$, p3$, s2$, v$, i%
        Dim Arr1() As String, Arr2() As String
        v = vbCrLf: DoArrc 'FillArrC 76, "-"

    's2 = new list of PartOfEventSubfolderPath '...|1986 Liga|1986 Liga\1986 Badenliga|...
        p3 = ArrC(3)                        'Path of Folder "Events"
        s2 = Get_Paths_OfAllSubfolders_AllLevels_Collection(p3)   'show s2
        s2 = Replace(s2, p3 + "\", ""): s2 = Replace(s2, v, "|"): s2 = "|" + s2 + "|"
        'show "s2" + v + v + s2
    '
        DgN = Get_AllDgFolderNames_FromDgs ': show "DgN" + v + v + DgN
        Info = Get_MissingDgFolders(s2, DgN)
        If Info <> "" Then OpenFolder ArrC(3): show Info
End Sub

Function Get_MissingDgFolders(s2$, DgN$) As String
    's2         |...|1990 DM Bonn|...|1990 Liga|1990 Liga\1990-11...Bonn-Köln"|...|
    'DgN        |...|1990 DM Bonn|...                    |1990-11...Bonn-Köln"|...|; LastLevel only
    'Part1      = Path of EventFolder       '"c1:\...\Events\"
    'Part2      = Name of EventSubFolder(s) '"1990 DM Bonn", "1990 Liga\1990-11...Bonn-Köln"
    'Action     gets all Part2
    
    'Vorbereitung
        Dim N$, N2$, N3$, N4$, s$, T$, v$, A%, i%, j%, k%, C1&, C2&, C3&, C4&
        Dim ArrDgN() As String, Arr1() As String
        v = vbCrLf
    'arrDgN
        ArrDgN = Split(DgN, "|")
    't
        T = v + "ÄNDERUNG ERFORDERLICH" + v + v + String(84, "-") + v + v _
            + "In der Excel-Datei 'Bes TR.xlsm', Tabellenblatt 'T4', befinden sich mehrere Designs." + v _
            + "Jedes Design dient der Erstellung einer bestimmten Ergebnisliste ('#result'-JPG)." + v _
            + "Ein Design enthält den Namen des Event-Ordners, in den das JPG gelegt werden soll." + v + v _
            + "Die Designs beziehen sich derzeit auf insgesamt " _
            + CStr(UBound(ArrDgN)) + " solcher Event-Ordner." + v _
            + "Von diesen " + CStr(UBound(ArrDgN)) + " Ordnern "
    'Search for missing EventFolders
        'Die einzelnen Dgn-Ordnernamen sollten in s2 zu finden sein
        For i = 1 To UBound(ArrDgN) - 1
            N = ArrDgN(i)
            C1 = InStr(1, s2, N + "|")
            If C1 = 0 Then
                s = s + N + "|"
            End If
        Next
        If s = "" Then Get_MissingDgFolders = "": Exit Function
    'xx
        Arr1 = Split("|" + s, "|")
        A = UBound(Arr1) - 1    'Anzahl nicht gefundener EventFolder
        If A = 1 Then T = T + "wurde einer nicht gefunden:" + v + v Else T = T + "wurden " + CStr(A) + " nicht gefunden:" + v + v
        N4 = ""
        For i = 1 To A
            N = Arr1(i)
            T = T + "[" + CStr(i) + "]" + "  " + N + v
            'Suche ähnlicher Ordnernamen
                For j = 1 To Len(N) - 4
                    N2 = Left(N, Len(N) - j): C2 = 1
                    For k = 1 To 5
                        C1 = InStr(C2, s2, N2) 'c=902
                        If C1 > 0 Then
                            'N2 (linker Teil des nicht existierenden Namens N)
                            '   wurde in s2 an der Stelle c1 gefunden
                            C4 = InStr(C1, s2, "|")
                            N3 = Mid(s2, C1, C4 - C1) '= OrdnerName in s2 (ähnlich N)
                            'Ist N3 ein SubSubFolder von 'Events'?
                                If Mid(s2, C1 - 1, 1) = "\" Then
                                    C3 = InStrRev(s2, "|", C1)
                                    N3 = Mid(s2, C3 + 1, C4 - C3 - 1)
                                End If
                            'show Mid(s2, c1 - 10, InStr(c1, s2, "|") - c1)
                            
                            If InStr(1, N4, N3) = 0 Then
                                If N4 <> "" Then N4 = N4 + ", "
                                N4 = N4 + N3
                            Else
                                k = 5
                            End If
                            C2 = C1 + 5
                        Else
                            k = 5 '= Exit k
                        End If
                    Next
                    If N4 <> "" Then j = Len(N) - 3 '= Exit j
                Next
                If N4 <> "" Then T = T + "     ähnlich: " + N4 + v: N4 = ""
                T = T + v
        Next
        T = T + "Bitte ändern sie den Namen des Ordners in 'Events' oder im Design so," + v _
            + "dass beide Namen übereinstimmen." + v + v + String(84, "-") + v
    Get_MissingDgFolders = T
End Function

'---------------------

Function Get_DgChars_GiveArrDg(ArrDg) As String
    Dim s$, i%, j%
    For i = 1 To UBound(ArrDg, 1)
        For j = 1 To UBound(ArrDg, 2)
            s = s + CStr(ArrDg(i, j)) + "|"
        Next
    Next
    Get_DgChars_GiveArrDg = s
End Function

Function Get_CharsOfOneActivatedDesign() As String
    'Called from    TEST_ZeitStop
    'Action         Inhalt aller Zellen des aktuellen Designs wird aneinandergereiht
    '                (ermöglicht CharCheck/DesignÄnderung seit letzter Aktivierung)
    
    'Vorbereitung
        Dim s$, i%, j%, ArrDesign()
        Load_ArrDesign ArrDesign
    For i = 1 To UBound(ArrDesign, 1)
        For j = 1 To UBound(ArrDesign, 2): s = s + CStr(ArrDesign(i, j)) + "|": Next
    Next
    Get_CharsOfOneActivatedDesign = s
End Function


