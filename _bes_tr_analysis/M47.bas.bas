Attribute VB_Name = "M47"
Option Explicit 'M47

Sub zzz_M47()
    
    'showProcs "names"
    
    'RenameModule "Modul1", "M41"
    EE 1: Beep
End Sub

Sub T4_Sort_Paste_Format_ListOfDesignTitles(AllDgTitles$)
    'Called from    T4_Write_ListOfDesignTitles_FromDgs, Update_DgData
    
    'Vorbereitung
        Dim Anz$, T$, C1%, C2%, C3%, i%, s9%, z9%
        Dim Arr1() As String, Wechsel As Boolean
        With ThisWorkbook.Sheets("T4")
        T = Delete_EmptyEndRowsInString(AllDgTitles)
    'Alte Einträge löschen  'nur T4-Spalten 2 und 3
        EE 0: z9 = LastRow(Sheets("T4")): s9 = lastCol(Sheets("T4"))
        .Range(.Cells(8, 2), .Cells(z9, 3)).Interior.Color = xlNone
        .Range(.Cells(8, 2), .Cells(z9, 3)).ClearContents
    'Arr1
        Arr1 = Split(T, vbCrLf)
        Anz = CStr(UBound(Arr1) + 1)
        QuickSort Arr1
        For i = 0 To UBound(Arr1)
            'In ListDgTitle: DgTitle --> DgTitle + " "
            Arr1(i) = Arr1(i) + " " 'zur Unterscheidung der Titel
        Next
        Paste_1DArrayToCol "T4", 8, 2, Arr1
        .Range(.Cells(8, 3), .Cells(UBound(Arr1) + 8, 3)) = "ü"
        't = Join(arr1, vbCrLf)
    'Einfärben
        For i = 0 To UBound(Arr1)
            C2 = CInt(Left(Arr1(i), 4))
            If C2 > C1 Then Wechsel = True Else Wechsel = False
            C1 = C2
            If Wechsel Then
                If C3 = 237 Then C3 = 205 Else C3 = 237 'hellgrau, dunkelgrau
            End If
            .Range(.Cells(i + 8, 2), .Cells(i + 8, 3)).Interior.Color = RGB(C3, C3, C3)
        Next
    'Nummerierung
        For i = 0 To UBound(Arr1)
            Arr1(i) = i + 1
        Next
    'Paste
        Paste_1DArrayToCol "T4", 8, 1, Arr1
        .[B6] = "(von " + CStr(i) + " bislang verfügbaren Ergebnislisten/Quellen)"
    'Finals
        End With: EE 1
        LogBuch "Up02 ListOfDesignTitles was updated (" + Anz + " titles)"
End Sub

Sub T4_Check_CountOfDgs()
    'Called from    T4_ActionsOnDgClickNone
    
    'Vorbereitung
        Dim qq$, T$, txt$, v$, C1%, C2%, i%, TitNew() As String, TitOld() As String
        v = vbCrLf: qq = Chr(34)
    'Count
        C1 = T4_Get_CountOfDgs_ViaDgLO
        C2 = T4_Get_CountOfDgs_ViaCountOfLinesOfT4SomeDgData
        If C1 = C2 Then Exit Sub
    'Compare
        T4_Load_T4SomeDgData_OneCol 1, TitOld 'T4SomeDgData_Col1_Titles

        T4_Write_T4SomeDgData
        T4_Load_T4SomeDgData_OneCol 1, TitNew 'T4SomeDgData_Col1_Titles
        'Dg hinzugefügt
            If C1 > C2 Then
                'TitNew trägt mehr Zeilen als TitOld
                T = Join(TitOld, v)
                For i = 1 To UBound(TitNew)
                    If Not T Like "*" + TitNew(i) + "*" Then Exit For
                Next
                If i < UBound(TitNew) Then
                    txt = "Ein Dg wurde hinzugefügt:" + v + qq + TitNew(i) + qq
                    UF3_Show_Text_ForXSeconds txt, 3
                End If
            End If
            T4_Write_ListOfDesignTitles_FromT4SomeDgData
        'Dg gelöscht
            If C2 > C1 Then
                'TitOld (vor T4_Write_T4SomeDgData) trägt mehr Zeilen als TitNew
                'Stop
                T4_Write_T4SomeDgData
                T4_Write_ListOfDesignTitles_FromT4SomeDgData
            End If
End Sub

Sub T4_Handle_UserClick(z%, s%)
    'Called from    Worksheet_SelectionChange[T4]
    'Action         Ermittlung DgStatus (Wo erfolgte der UserClick?) + Handling
    
    'Vorbereitung
        Dim St$
    'Ermittlung DgStatus
        St = T4_Get_WhereUserClicksInT4_0_1_2(z, s)
    'Status 0       UserClick außerhalb eines Dg
        If St = "0" Then T4_ActionsOnDgClickNone:                   Exit Sub
    'Status 1       UserClick innerhalb eines Dg, kein Dg aktiv
        If Left(St, 1) = "1" Then T4_ActionsOnDgClickFirst St:      Exit Sub
        '           sT = "1|z1|s1|z2|s2|Title" [z1, s1, ... je 4 Ziffern]
    'Status 2       weiterer UserClick innerhalb des bereits aktiven Dg
        If St = "2" Then FillArrC 48, "2": T4_ActionsOnDgClickNext: Exit Sub
End Sub

Function T4_Get_WhereUserClicksInT4_0_1_2(z%, s%) As String
    'Called from    T4_Handle_UserClick
    'Action         liefert                 [z1, s1, ... je 4 Ziffern]
    '               "0",                    falls UserClick außerhalb eines Dg
    '               "1|z1|s1|z2|s2|Title",  falls UserClick innerhalb eines Dg, kein Dg aktiv
    '               "2",                    falls UserClick innerhalb eines bereits aktiven Dg
    
    'Vorbereitung
        Dim Status$, TitleOfActiveDg$, TitleOfDgClicked$: DoArrc
        Status = T4_Get_DgPosAndTitle_IfUserClickWasInDg(z, s)
    'Action
        If Status = "" Then
            'UserClick was not inside a Dg
            Status = "0"
        Else
            'UserClick was inside a Dg
            If ArrC(37) = "-" Then              'z1
                'no Dg is activ
                'UserClick was FirstClick inside a Dg
                Status = "1" + Status           '
            Else
                'one Dg is activ
                TitleOfActiveDg = ArrC(41)
                TitleOfDgClicked = Mid(Status, 22)
                If TitleOfActiveDg = TitleOfDgClicked Then
                    'UserClick was another click inside same Dg
                    Status = "2"
                Else
                    'UserClick was FirstClick inside a Dg
                    Status = "1" + Status
                End If
            End If
        End If
    'Finals
        T4_Get_WhereUserClicksInT4_0_1_2 = Status
End Function

Function T4_Get_CountOfDgs_ViaCountOfLinesOfT4SomeDgData() As Integer
    'Called from    T4_Check_CountOfDgs
    
    'Vorbereitung
        Dim c%, zLast&
        zLast = Get_NrOfLastRowInColumnNr(5, "T4")
        c = zLast - 7
    'Finals
        T4_Get_CountOfDgs_ViaCountOfLinesOfT4SomeDgData = c '141
End Function

Function T4_Get_CountOfDgs_ViaDgLO() As Integer
    'Called from    T4_Check_CountOfDgs
    
    'Vorbereitung
        Dim ColNrs$, w As Worksheet, r As Range
        Dim c%, i%, s&, z&, zLast&, Arr1() As String
        Set w = Sheets("T4")
    'LO_inRowNr8
        ColNrs = T4_Get_ColNrs_ofAllLO_inRowNr8 '"20|31|42|...|521|537|"
    'Count
        Arr1 = Split(ColNrs, "|")
        For i = 0 To UBound(Arr1) - 1
            s = CLng(Arr1(i))       'one ColumnNr of "LO"  in Row8
            zLast = Get_NrOfLastRowInColumnNr(s, "T4")
            Set r = w.Range(w.Cells(8, s), w.Cells(zLast, s))
            c = c + Application.WorksheetFunction.CountIf(r, "LO")
        Next
    'Finals
        T4_Get_CountOfDgs_ViaDgLO = c '141
End Function


Function T4_Get_DataFromOneDg_ToFillOneCell_Names_OfT4SomeDgData(i%, z1%, s1%, z2%, s2%)
    'Called from    T4_Write_T4SomeDgData
    'Dg             wird anstelle ArrDg verwendet (FillArrDg dauert zu lange)
    'Action         liefert "053 Names: |#|Irena Tolmacheva|6|4|Team|1|#|Marina ...|"
    '               für Inhalt Zelle in Spalte 8 (Names) von T4SomeDgData
    
    'Vorbereitung
        Dim E$, N$, Ns$, r$, s%, z%, Arr1(), Dg()
    'Dg
        With Sheets("T4"): Dg = .Range(.Cells(z1, s1), .Cells(z2, s2)).Value: End With
        T4_Add_T4Row2_ToDg Dg, s1, s2
        'showArray2D Dg: Stop
    'Einzel/Synchron/Mannschaft
        If Dg(2, 2) Like "*(Einzel)" Then E = "Einz"
        If Dg(2, 2) Like "*(Synchron)" Then E = "Sync"
        If Dg(2, 2) Like "*(Mannschaft)" Then E = "Team"
        If UCase(Dg(1, 2)) Like "*LIGA*" Then E = "Team"
        If E = "" Then Stop
        If E = "Sync" Then T4_Add_Rank_ToAllDg_SyPartner Dg, z1, s1
        If E = "Team" Then T4_Add_Rank_ToAllDg_TeamMembers Dg
    'Get Names
        For s = 1 To UBound(Dg, 2)
            If CStr(Dg(5, s)) = "W" Or CStr(Dg(5, s)) = "M" Or CStr(Dg(5, s)) = "X" Then
                'Spalte s in Dg ist eine NamenSpalte
                For z = 6 To z2 - z1
                    N = Dg(z, s)        'one Name
                    If InStr(1, N, " (") > 0 Then N = Left(N, InStr(1, N, " (") - 1)
                    If IsValidName(N) Then
                        'Es existiert ein Name in Dg(z, s)
                        r = "-"
                        If E = "Einz" Then r = T4_Get_SomeDgData8_RankEinz(Dg, z, s)
                        If E = "Sync" Then r = T4_Get_SomeDgData8_RankSync(Dg, z, s, z1, s1)
                        If E = "Team" Then r = T4_Get_SomeDgData8_RankTeam(Dg, z, s)
'                        'for testing only
'                            If E = "Team" Then
'                                r = T4_Get_SomeDgData8_RankTeam(Dg, z, s)
'                                Sheets("T4").Activate
'                                Sheets("T4").Cells(z1 - 1 + z, s1 - 1 + s).Select: Stop
'                                'r = T4_Get_SomeDgData8_RankTeam(Dg,  z, s)
'                            End If
'                        'for testing only
'                            If r = "-" And i > 171 Then
'                                Sheets("T4").Activate
'                                Sheets("T4").Cells(z1 - 1 + z, s1 - 1 + s).Select: Stop
'                            End If
                        Ns = Ns + "#|" + N + "|" + CStr(z) + "|" + CStr(s) + "|" + E + "|" + r + "|"
                    End If
                Next
            End If
        Next
        If Ns = "" Then Ns = "000 Names" Else Ns = Format(anzAinB("#", Ns), "000") + " Names: |" + Ns
    'Finals
        T4_Get_DataFromOneDg_ToFillOneCell_Names_OfT4SomeDgData = Ns
End Function



Function T4_Get_SomeDgData8_RankTeam(Dg, z%, s%) As String
    'Called from    T4_ActionsOnDgClickNone
    'Scope          Dg eines Mannschafts-Wettkampfes
    'Dg(z, s)       = ValidNameCell
    'Action         liefert RankOfActualCompetitor
    
    'Vorbereitung
        Dim r$
        T4_Get_SomeDgData8_RankTeam = "-"
    'Exit
        If Not "IJK" Like "*" + Dg(5, s - 2) + "*" Then Exit Function
    'Mannschafts-Wettkämpfe:   '2. Zelle links des Namens steht der Rang
        r = Trim(CStr(Dg(z, s - 2)))
        If Not r Like "*#*" Then Exit Function
        T4_Get_SomeDgData8_RankTeam = r
End Function

Sub T4_Add_Rank_ToAllDg_SyPartner(Dg, z1%, s1%)
    'Called from    T4_Get_DataFromOneDg_ToFillOneCell_Names_OfT4SomeDgData
    
    'Vorbereitung
        Dim r%, s%, z%
    'Action
        'Rang-Spalte suchen
        For s = 2 To UBound(Dg, 2) - 1
            If "IJK" Like "*" + Dg(5, s) + "*" Then
                'RangSpalte gefunden: s
                
                
                'Leere Rang-Zelle bei SyPartner suchen
                    For z = 6 To UBound(Dg, 1) - 2
                    
                        If Dg(z + 1, s) Like "*,*" Then
                            'In der RangZelle ist ein Wert 9,88... fest eingetragen
                            EE 0: Sheets("T4").Cells(z1 + z, s1 - 1 + s).Select: EE 1
                            'Stop
                            Sheets("T4").Cells(z1 + z, s1 - 1 + s) = ""
                            Dg(z + 1, s) = ""
                        End If
                    
                        If CStr(Dg(z, s)) Like "*#*" _
                          And Dg(z + 1, s + 1) <> "" _
                          And Dg(z + 1, s) = "" Then
                          'RangZelle enthält eine Ziffer
                          'AND Zelle darunter, rechts daneben enthält einen Namen
                          'AND Zelle darunter ist leer
                            r = CInt(Dg(z, s)): Dg(z + 1, s) = r
                            'Dg(z + 1, s) = Dg(z, s) wird nicht verwendet; speichert 8,899..
                            'Kontrolle
                                If Len(r) > 2 Then Stop
                                If Len(Dg(z + 1, s)) > 2 Then Stop
                                If r Like "*,*" Then '9,88.. statt 10
                                    'Dg anzeigen
                                        With Sheets("T4"): EE 0: .Activate
                                        .Cells(z1 - 1 + z, s1 - 1 + s).Select: EE 1: End With
                                    showArray2D Dg: Stop
                                End If
                        End If
                    Next
            End If
        Next
        'showArray2D Dg: Stop
End Sub

Sub T4_Add_Rank_ToAllDg_TeamMembers(Dg)
    'Called from    T4_Get_DataFromOneDg_ToFillOneCell_Names_OfT4SomeDgData
    
    'Vorbereitung
        Dim i%, s%
    'Add
        For s = 2 To UBound(Dg, 2) - 1
            'Schleife über die Spalten der DgZeile5
            If "IJK" Like "*" + Dg(5, s) + "*" Then
                'Es wurde eine RangSpalte gefunden
                'LeerZellen unterhalb einer Zelle mit RangZahl sollen gefüllt werden
                For i = 6 To UBound(Dg, 1) - 2
                    If Dg(i, s) <> "" And Dg(i + 1, s) = "" Then
                        Dg(i + 1, s) = CStr(Dg(i, s))
                    End If
                Next
            End If
        Next
        'showArray2D Dg: Stop
End Sub


Sub T4_Add_T4Row2_ToDg(Dg, s1, s2)
    Dim i%, D2()
    With Sheets("T4")
    D2 = .Range(.Cells(2, s1), .Cells(2, s2)).Value   'Zeile2 oberhalb Dg
    'Add values to DgRow 5
        For i = 1 To UBound(D2, 2)
            Dg(5, i) = D2(1, i)
            If Dg(5, i) = "" Then Dg(5, i) = "-"
        Next
    End With
End Sub

Function T4_Get_SomeDgData8_RankSync(Dg, z%, s%, z1%, s1%) As String
    'Called from    T4_ActionsOnDgClickNone
    'Scope          Dg eines Synchron-Wettkampfes
    'Dg(z, s)       = ValidNameCell
    'Action         liefert RankOfActualCompetitor
    
    'Vorbereitung
        Dim r$
        T4_Get_SomeDgData8_RankSync = "-"
    'Exit
        If Not "IJK" Like "*" + Dg(5, s - 1) + "*" Then Exit Function
    'Synchron-Wettkämpfe:   Links des Namens/Partnernamens steht der Rang
            'falls Dg(x,y)=Leer zeigt IsNumeric True
        r = CStr(Dg(z, s - 1))
        If Not r Like "*#*" Then r = ""
        If r = "" Then
            'Zelle links des Namens hat noch keine Rangzahl;
            'Zelle soll die Rangzahl der Zelle darüber erhalten
            r = CStr(Dg(z - 1, s - 1))
        End If
        If r = "" Then Exit Function
        If r Like "*,*" Then '9,88.. statt 10
            'Dg anzeigen
                With Sheets("T4"): EE 0: .Activate
                .Cells(z1 - 1 + z, s1 - 1 + s).Select: EE 1: End With
            showArray2D Dg: Stop
        End If
        T4_Get_SomeDgData8_RankSync = r
End Function

Function T4_Get_SomeDgData8_RankEinz(Dg, z%, s%) As String
    'Called from    T4_ActionsOnDgClickNone
    'Scope          Dg eines Einzel-Wettkampfes
    'Dg(z, s)       = ValidNameCell
    'Action         liefert RankOfActualCompetitor
    
    'Vorbereitung
        Dim A$
        T4_Get_SomeDgData8_RankEinz = "-"
    'Exit
        If Not "IJK" Like "*" + Dg(5, s - 1) + "*" Then Exit Function
    'Einzel-Wettkämpfe:     Links des Namens steht der Rang
        A = CStr(Dg(z, s - 1)) 'falls Dg(z, s - 1) = leer zeigt IsNumeric True
        If A Like "*#*" Then
                If CInt(A) > 0 Then T4_Get_SomeDgData8_RankEinz = A
        End If
End Function

Sub T4_Check_DgPositions_ViaT4SomeDgData()
    'Called from    T4_ActionsOnDgClickNone
    'Action         Checkt alle T4SomeDgData-Zeilen, ob Dg(z1, s1) = "LO";
    '               schreibt ggf. T4SomeDgData neu
    
    Exit Sub
    
    'Vorbereitung
        Dim L$, Title$, v$, i%, s1%, z1%, DgData5(), w As Worksheet
        EE 0: v = vbCrLf: Set w = Sheets("T4")
    'DgData5
        T4_Load_T4SomeDgData_1to5 DgData5
    'Action
        For i = 1 To UBound(DgData5, 1)
            z1 = DgData5(i, 2): s1 = DgData5(i, 3)
            If w.Cells(z1, s1) <> "LO" Then
                Title = DgData5(i, 1)
                L = L + "Dg '" + Title + "' was moved from (" + CStr(z1) + ", " + CStr(s1) + ")" + v
            End If
        Next
    'Finals
        EE 1:  If L <> "" Then show L: T4_Write_T4SomeDgData
End Sub

Function T4_Get_ColNrs_ofAllLO_inRowNr8() As String
    'Called from    T4_Write_T4SomeDgData
    'RowNr8         Alle ersten Dgs einer DgGroup beginnen in T4-Zeile 8
    'Action         liefert die SpaltenNrn aller "LO" in T4-Zeile 8: "20|31|42|...|521|537|"
    
    'Vorbereitung
        Dim ColNrs$, AfterColNr%, i%, s%, sFirst%
        AfterColNr = 19
    'SpaltenNrn aller "LO" in T4-Zeile 8
        For i = 1 To 999
            s = Get_ColumnNr_HoldingMyTextWholeInRowXSearchRight("T4", 8, AfterColNr, "LO")
            'Nach dem letzten "LO"-Fund wird wieder der erste gefunden
            If s = sFirst Then Exit For Else ColNrs = ColNrs + CStr(s) + "|"
            If sFirst = 0 Then sFirst = s
            AfterColNr = s
        Next
    'Finals
        T4_Get_ColNrs_ofAllLO_inRowNr8 = ColNrs
        'show ColNrs
End Function

Function T4_Get_DgPosAndTitle_IfUserClickWasInDg(z%, s%) As String
    'Called from    xxx
    'Action         xxx
    
    'Vorbereitung
        Dim s1$, s2$, T$, Title$, z1$, z2$, i%, DgData5()
    'DgData5
        T4_Load_T4SomeDgData_1to5 DgData5
        'showArray2D DgData5
    'Ist ein UserClick auf (z, s) innnerhalb eines Dgs?
        For i = 1 To UBound(DgData5, 1)
            If z >= DgData5(i, 2) Then
                If z <= DgData5(i, 4) Then
                    If s >= DgData5(i, 3) Then
                        If s <= DgData5(i, 5) Then
                            'Dg ist ermittelt
                            z1 = "|" + Format(DgData5(i, 2), "0000")
                            s1 = "|" + Format(DgData5(i, 3), "0000")
                            z2 = "|" + Format(DgData5(i, 4), "0000")
                            s2 = "|" + Format(DgData5(i, 5), "0000")
                            Title = "|" + DgData5(i, 1)
                            T = z1 + s1 + z2 + s2 + Title
                        End If
                    End If
                End If
            End If
        Next
    'Finals
        T4_Get_DgPosAndTitle_IfUserClickWasInDg = T
End Function

Sub T4_Load_AllDgCorners_ViaDgAreaSearch(ByRef Co())
    'Called from    xxx
    'Action         Füllt 2D-Array Co() mit 1 Zeile je Dg: |z1|s1|z2|s2| '|0008|0020|0018|0028|
    
    'Vorbereitung
        Dim A$, v$, c%, i%, j%, s1%, z1%, s&, z&, DgArea(), ArrA() As String
        v = vbCrLf
    'Gesamten Dg-Bereich in Array nehmen
        T4_Load_DgArea DgArea
    'Get list of z1, s1 --> A
        For s = 1 To UBound(DgArea, 2)
            If DgArea(1, s) = "LO" Then
                For z = 1 To UBound(DgArea, 1)
                    If DgArea(z, s) = "LO" Then
                        A = A + "|" + Format(z + 7, "0000") + "|" + Format(s + 7, "0000") + "|" + v
                    End If
                Next
            End If
        Next
        If Right(A, 2) = v Then A = Left(A, Len(A) - 2)
        ArrA = Split(v + A, v)
    'ArrA --> 2D-Array Co()
        ReDim Co(1 To UBound(ArrA), 1 To 4)
        For i = 1 To UBound(ArrA)
            Co(i, 1) = Mid(ArrA(i), 2, 4)
            Co(i, 2) = Mid(ArrA(i), 7, 4)
        Next
        'showArray2D Co
    'Add z2, s2 to Co
        For i = 1 To UBound(Co, 1)
            z1 = CInt(Co(i, 1)): s1 = CInt(Co(i, 2))
            For j = 1 To 999
                If DgArea(z1 - 7, s1 - 7 + j) = "RO" Then Co(i, 4) = Format(s1 + j, "0000"): j = 999
            Next
            For j = 1 To 999
                If DgArea(z1 - 7 + j, s1 - 7) = "LU" Then Co(i, 3) = Format(z1 + j, "0000"): j = 999
            Next
        Next
        'showArray2D Co
    'Finals
        'T4_Get_AllDgCorners = A
End Sub

Sub T4_Get_DgTitles_Fast_TEST()
    Dim s$, v$, c&: v = vbCrLf
    s = T4_Get_DgTitles_Fast
    c = anzAinB(v, s) + 1
    show CStr(c) + v + v + s
    'show Get_AllDgTitles_FromDgs
End Sub

Function T4_Get_DgTitles_Fast()
    'Called from    xxx
    
    'Vorbereitung
        Dim A$, v$, s&, z&, DgArea(): v = vbCrLf
    'Gesamten Dg-Bereich in Array nehmen
        T4_Load_DgArea DgArea
    'Get list of titles --> A
        For s = 1 To UBound(DgArea, 2)
            If DgArea(1, s) = "LO" Then
                For z = 1 To UBound(DgArea, 1)
                    If DgArea(z, s) = "LO" Then A = A + DgArea(z + 1, s + 1) + v
                Next
            End If
        Next
        If Right(A, 2) = v Then A = Left(A, Len(A) - 2)
    'Finals
        T4_Get_DgTitles_Fast = A
End Function

Sub T4_Load_DgArea(ByRef DgArea())
    Dim w As Worksheet
    Set w = Sheets("T4")
    'Gesamten Dg-Bereich in Array nehmen
        DgArea = w.Range(w.Cells(8, 8), w.Cells(LastRow(Sheets("T4")), lastCol(Sheets("T4"))))
End Sub

Sub T4_Load_T4SomeDgData_1to5(ByRef D1to5())
    'Called from    T4_Get_DgPosAndTitle_IfUserClickWasInDg
    'Action         T4SomeDgData-Spalten 1 bis 5 in 2D-Array nehmen
    
    'Vorbereitung
        Dim zLast%, w As Worksheet: Set w = Sheets("T4")
        zLast = Get_NrOfLastRowInColumnNr(5, "T4")
    'Action
        D1to5 = w.Range(w.Cells(8, 5), w.Cells(zLast, 9))
End Sub

Sub T4_Load_T4SomeDgData_2to6(ByRef D2to6())
    'Called from    xxx
    'Action         T4SomeDgData-Spalten 2 bis 6 in 2D-Array nehmen
    
    'Vorbereitung
        Dim zLast%, w As Worksheet: Set w = Sheets("T4")
        zLast = Get_NrOfLastRowInColumnNr(5, "T4")
    'Action
        D2to6 = w.Range(w.Cells(8, 6), w.Cells(zLast, 10))
End Sub

Sub T4_Load_T4SomeDgData_1to8(ByRef D1to8())
    'Called from    xxx
    'Action         T4SomeDgData-Spalten 1 bis 8 in 2D-Array nehmen
    
    'Vorbereitung
        Dim zLast%, w As Worksheet: Set w = Sheets("T4")
        zLast = Get_NrOfLastRowInColumnNr(5, "T4")
    'Action
        D1to8 = w.Range(w.Cells(8, 5), w.Cells(zLast, 12))
End Sub


Sub T4_Load_T4SomeDgData_OneCol(MyColNr, ByRef StringArray1D() As String)
    'Called from    xxx
    'Action         T4SomeDgData-Spalte MyColNr in 1D-Array nehmen
    
    'Vorbereitung
        Dim zLast%, w As Worksheet, Array2D()
        Set w = Sheets("T4")
        zLast = Get_NrOfLastRowInColumnNr(5, "T4")
    'Action
        Array2D = w.Range(w.Cells(8, MyColNr + 4), w.Cells(zLast, MyColNr + 4))
        Load_StringArray1D_from_1ColumnArray2D StringArray1D, Array2D
End Sub

Sub T4_Load_ListOfDesignTitles(ByRef StringArray1D() As String)
    'Called from    T4_Write_ListOfDesignTitles_FromT4SomeDgData
    'Action         ListOfDesignTitles [ab T4(8, 2)] in 1D-Array nehmen
    
    'Vorbereitung
        Dim zLast%, w As Worksheet, Array2D()
        Set w = Sheets("T4")
        zLast = Get_NrOfLastRowInColumnNr(2, "T4")
    'Action
        Array2D = w.Range(w.Cells(8, 2), w.Cells(zLast, 2))
        Load_StringArray1D_from_1ColumnArray2D StringArray1D, Array2D
End Sub

Sub Load_StringArray1D_from_1ColumnArray2D(ByRef Arr1D() As String, ByRef Arr2D())
    Dim c%, i%
    c = UBound(Arr2D, 1): ReDim Arr1D(1 To c)
    For i = 1 To c: Arr1D(i) = CStr(Arr2D(i, 1)): Next
End Sub

Function Get_AllDgTitles_FromDgs() As String
    'Called from    T4_Write_ListOfDesignTitles_FromDgs, T4_Load_ArrDgTitles_FromDgs
    
    Dim A$, i&, j&, sT4Last&, zT4Last&, w As Worksheet, Arr()
    Set w = Sheets("T4")
    'Gesamten Dg-Bereich in Array nehmen
        zT4Last = LastRow(Sheets("T4"))
        sT4Last = lastCol(Sheets("T4"))
        Arr = w.Range(w.Cells(8, 19), w.Cells(zT4Last, sT4Last))
    For i = 1 To UBound(Arr, 1) - 2
        'Schleife über alle T4-Zeilen
        For j = 1 To UBound(Arr, 2)
            'Schleife über alle T4-Spalten von Zeile i
            If Arr(i, j) Like "19##*" And Not Arr(i + 1, j) Like "19##*" Then A = A + Arr(i, j) + vbCrLf
        Next
    Next
    If Right(A, 2) = vbCrLf Then A = Left(A, Len(A) - 2)
    Get_AllDgTitles_FromDgs = A
End Function

Function T4_Get_DgTxt(z1%, s1%, z2%, s2%)
    Dim s$, i%, j%, Arr1()
    With Sheets("T4")
        Arr1 = .Range(.Cells(z1, s1), .Cells(z2, s2)).Value
        For i = 1 To UBound(Arr1, 1)
            For j = 1 To UBound(Arr1, 2)
                s = s + "|" + CStr(Arr1(i, j))
            Next
        Next
    End With
    T4_Get_DgTxt = Format(Len(s), "0000") + "|" + s
End Function




