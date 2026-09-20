Attribute VB_Name = "M_T4DgData"
Option Explicit

Sub zzz_M_T4DgData()
    'showProcs "paste"
    
        
    EE 1: Beep
End Sub

Sub zzz_Procs_ThisModule_M_T4DgData()
'here
'+   T4_Write_T4SomeDgData()
'+   T4_Write_ListOfDesignTitles_FromT4SomeDgData
'    T4_Write_ListOfDesignTitles_FromDgs
'
'not yet here
'    T4_Change_T4SomeDgData_OneNameData
'    T4_Check_DgPositions_ViaT4SomeDgData
'    T4_Get_AllVnNnInDgsOfOneEvent_FromT4SomeDgData
'    T4_Get_CountOfDgs_ViaCountOfLinesOfT4SomeDgData
'    T4_Get_DataFromOneDg_ToFillOneCell_Names_OfT4SomeDgData
'    T4_Get_SomeDgData8_RankEinz
'    T4_Get_SomeDgData8_RankSync
'    T4_Get_SomeDgData8_RankTeam
'    T4_Load_ArrDgTitles_FromT4SomeDgData
'    T4_Load_T4SomeDgData_1to5
'    T4_Load_T4SomeDgData_1to8
'    T4_Load_T4SomeDgData_OneCol
'    T4_Set_WidthOf_T4SomeDgData
'    T6_Find_MissingAgeData_InT8SomeDgData
'    Update_DgData
'
'    CheckCells_ListOfDesignTitles_ListOfCompetitors
'    T4_Load_ArrDgTitles_FromT4ListOfDesignTitles
'    T4_Load_ListOfDesignTitles
'    T4_Save_PositionOfHeaderCell_OfListOfDesignTitles_ToArrC
'    T4_Sort_Paste_Format_ListOfDesignTitles
'
'    Get_AllDgTitles_FromDgs
'    ShowAllDgTitlesAndFolders
'    T4_Get_DgTitles_Fast
'    T4_Get_DgTitles_Fast_TEST
'    T4_Load_ArrDgTitles_FromDgs
'    T4_Load_ArrDgTitles_FromT4ListOfDesignTitles
'    T4_Load_ArrDgTitles_FromT4SomeDgData

End Sub

Sub CompareTime_twoProcs()
    Dim s1$, s2$
    'Proc1
        Ticks1
        T4_Write_TitlesAndData_FromDgs
        s1 = T6_GetDuration()
    'Proc2
        Ticks1
        T4_Write_T4SomeDgData
        T4_Write_ListOfDesignTitles_FromT4SomeDgData
        s2 = T6_GetDuration()
        show s1 + vbCrLf + s2
End Sub

Sub T4_Write_ListOfDesignTitles_FromT4SomeDgData()
    'Called from    T4_Check_CountOfDgs, Worksheet_SelectionChange[T4]
    
    'Vorbereitung
        Dim D$, L$, v$, c%, C1%, C2%, C3%, i%, ArrD() As String, ArrL()
        Dim r As Range, w As Worksheet, Wechsel As Boolean
        v = vbCrLf: EE 0: Set w = Sheets("T4")
    'ArrD       'T4SomeDgData_Col1_Titles   'ohne Zusatz-Leerzeichen
        T4_Load_T4SomeDgData_OneCol 1, ArrD '1. Spalte von T4SomeDgData
        c = UBound(ArrD)
    'Alte Einträge löschen 'List of DesignTitles
        Set r = w.Range(w.Cells(8, 1), w.Cells(c + 9, 4))
        r.Interior.Color = xlNone: r.ClearContents
    'ArrL - Array für neue 'List of Design Titles'
        ReDim ArrL(1 To c + 9, 1 To 3)
        For i = 1 To c
            ArrL(i, 1) = i          '1. Spalte: Nummerierung
            ArrL(i, 2) = ArrD(i)    '2. Spalte: Titel
                'If ArrD(i) = "" Then Stop
            ArrL(i, 3) = "ü"        '3. Spalte: Erledigt-Haken
        Next
    'Einfärben
        For i = 1 To c
            'showArray ArrD 'Liste titles 'ohne zusätzliches " "
            C2 = CInt(Left(ArrD(i), 4)) '1960
            If C2 > C1 Then Wechsel = True Else Wechsel = False
            C1 = C2
            If Wechsel Then
                If C3 = 237 Then C3 = 205 Else C3 = 237 'hellgrau, dunkelgrau
            End If
            w.Range(w.Cells(i + 7, 2), w.Cells(i + 7, 3)).Interior.Color = RGB(C3, C3, C3)
        Next
    'Paste
        Paste_2DArrayToCell_z_s "T4", 8, 1, ArrL
        w.[B6] = "(von " + CStr(c) + " bislang verfügbaren Ergebnislisten/Quellen)"
    'Finals
        EE 1
End Sub

Sub T4_Write_T4SomeDgData()
    'Called from    xxx
    
    'Vorbereitung
        Dim DgNames$, L$, ss1$, ss2$, v$
        Dim i%, j%, k%, s1%, s2%, z1%, z2%, sT4Last&, zT4Last&
        Dim A(), B() As String, c() As String, SomeDgData()
        With Sheets("T4"): v = vbCrLf
    'Gesamten T4-Bereich in Array A nehmen
        zT4Last = LastRow(Sheets("T4"))
        sT4Last = lastCol(Sheets("T4"))
        A = .Range(.Cells(1, 1), .Cells(zT4Last, sT4Last)).Value
    'Alle "i " und " i" in Zeile 2 suchen (alle s1 und s2 der Dgs)
         For j = 1 To UBound(A, 2)
            If A(2, j) = "i " Then ss1 = ss1 + "|" + CStr(j) '|20|29|40|53|72|107|...
            If A(2, j) = " i" Then ss2 = ss2 + "|" + CStr(j) '|26|37|50|69|104|117|...
         Next j
        'show ss1 + v + ss2
    'L = Liste für T4SomeDgData
        B = Split(ss1, "|"): c = Split(ss2, "|")
        If UBound(B) <> UBound(c) Then Stop
        For j = 1 To UBound(B)
            'B(j) und C(j) ist der s1- und s2-Wert aller Dgs der DgGroup j
            For i = 1 To zT4Last
                s1 = B(j)
                If A(i, s1) = "LO" Then z1 = i: k = k + 1
                If A(i, s1) = "LU" Then
                    z2 = i: s2 = c(j)       'Jetzt sind z1,z2,s1,s2 bekannt
                    'L = Title°z1°s1°z2°s2°Folder°-°Names°-
                    L = L + v + A(z1 + 1, s1 + 1) + "°" + CStr(z1) + "°" + CStr(s1) + "°" _
                        + CStr(z2) + "°" + CStr(s2) + "°" + A(z1, s1 + 1) + "°-°"
                        DgNames = T4_Get_DataFromOneDg_ToFillOneCell_Names_OfT4SomeDgData(k, z1, s1, z2, s2)
                    L = L + DgNames + "°-"
                End If
            Next
        Next
        L = Mid(L, 3) ': show L
    'SomeDgData() aufbauen  (= 2DArray)
        B = Split(L, v): QuickSort B
        ReDim DgTitles(1 To UBound(B) + 1)
        ReDim SomeDgData(1 To UBound(B) + 1, 1 To 9)
        For i = 1 To UBound(B) + 1
            'Schleife über alle L-Zeilen    'Title°z1°s1°z2°s2°Folder°-°Names°-
            c = Split(B(i - 1), "°")
            For j = 1 To UBound(c) + 1
                'Schleife über alle 9 Items der L-Zeile i
                SomeDgData(i, j) = c(j - 1)
                'DgTitles(i) = SomeDgData(i, 1) + " "
            Next
        Next
    'Alte T4SomeDgData-Zellen löschen
        .Range(.Cells(8, 5), .Cells(zT4Last, 13)).ClearContents
    'Schriftgröße
        .Range(.Cells(8, 2), .Cells(zT4Last, 2)).Font.size = 11
        .Range(.Cells(8, 5), .Cells(zT4Last, 13)).Font.size = 6
    'Paste
        Paste_2DArrayToCell_z_s "T4", 8, 5, SomeDgData
        'Paste_1DArrayToCol "T4", 8, 2, DgTitles
    'Finals
        End With: Beep
End Sub

Sub T4_Write_ListOfDesignTitles_FromDgs()
    Dim T$
    Sheets("T4").Activate
    T = Get_AllDgTitles_FromDgs
    T4_Sort_Paste_Format_ListOfDesignTitles T
End Sub



