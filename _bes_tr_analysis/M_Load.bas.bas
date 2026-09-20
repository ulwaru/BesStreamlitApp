Attribute VB_Name = "M_Load"
Option Explicit 'M_Load
'

Sub zzz_M_Load()
    showProcs "LOAD LIKE"
    
    'RenameModule "Modul1", "Load"
    EE 1: Beep
    'Dim A() As String: T5_Load_Club A: showArray A
End Sub

Sub T5_Load_Nn(Arr1D() As String):     T5_Load_OneColumn Arr1D, 3:  End Sub
Sub T5_Load_Vn(Arr1D() As String):     T5_Load_OneColumn Arr1D, 4:  End Sub
Sub T5_Load_Jhg(Arr1D() As String):    T5_Load_OneColumn Arr1D, 6:  End Sub
Sub T5_Load_Club(Arr1D() As String):   T5_Load_OneColumn Arr1D, 7:  End Sub
Sub T5_Load_LTV(Arr1D() As String):    T5_Load_OneColumn Arr1D, 8:  End Sub
Sub T5_Load_Nation(Arr1D() As String): T5_Load_OneColumn Arr1D, 9:  End Sub
Sub T5_Load_Folder(Arr1D() As String): T5_Load_OneColumn Arr1D, 10: End Sub
Sub T5_Load_VnNn(Arr1D() As String):   T5_Load_OneColumn Arr1D, 22: End Sub

Sub T5_Load14_TEST()
    Dim A() As String
    T5_Load_1_to_17 A
    showArray2D A
End Sub


Sub T5_Load_1_to_17(A() As String)
    '[1]Vn    [2]Nn    [3]mw    [4]Club  [5]LTV   [6]Nation   [7]Folder   [8]rn    [9]gn
    '[10]VnNn [11]NnVn [12]vh1  [13]vh2  [14]vh3  [15]LiVnNn  [16]LiNnVn  [17]LiRn
    Dim s$, vh$, C1%, C2%, i%, j%, zLast%, Arr2D()
    zLast = Get_NrOfLastRowInColumnNr(3, "T5")
    With Sheets("T5")
        Arr2D = .Range(.Cells(8, 1), .Cells(zLast, 22)).Value
        ReDim A(1 To UBound(Arr2D, 1), 1 To 17)
        For i = 1 To UBound(A, 1)
            'Column 1-11
                A(i, 1) = Arr2D(i, 3)               'Nn
                A(i, 2) = Arr2D(i, 4)               'Vn
                A(i, 3) = Arr2D(i, 5)               'm/w
                A(i, 4) = Arr2D(i, 7)               'Club
                A(i, 5) = Arr2D(i, 8)               'LTV
                A(i, 6) = Arr2D(i, 9)               'Nation
                A(i, 7) = Arr2D(i, 10)              'Name of PersonFolder
                A(i, 8) = Arr2D(i, 12)              'rn
                A(i, 9) = Arr2D(i, 13)              'gn
                A(i, 10) = Arr2D(i, 22)             'VnNn
                A(i, 11) = A(i, 1) + ", " + A(i, 2) 'NnVn
                If A(i, 8) = "-" Then A(i, 8) = ""
                If A(i, 9) = "-" Then A(i, 9) = ""
            'Column 12-14   'vh1 vh2 vh3
                vh = Arr2D(i, 14) 'Jun/Jul/Aug
                If vh = "-" Then
                    A(i, 12) = "": A(i, 13) = "": A(i, 14) = ""
                ElseIf Not vh Like "*/*" Then A(i, 12) = Arr2D(i, 14): A(i, 13) = "": A(i, 14) = ""
                ElseIf vh Like "*/*/*" Then
                    C1 = InStr(1, vh, "/"): C2 = InStr(C1 + 1, vh, "/")
                    A(i, 12) = Left(vh, C1 - 1)
                    A(i, 13) = CutVonBis(vh, C1 + 1, C2 - 1)
                    A(i, 14) = Mid(vh, C2 + 1)
                ElseIf vh Like "*/*" Then
                    C1 = InStr(1, vh, "/")
                    A(i, 12) = Left(vh, C1 - 1)
                    A(i, 13) = Mid(vh, C1 + 1)
                    A(i, 14) = ""
                Else: Stop
                End If
            'A(i, 15)       'LiVnNn 'LinkName VnNn,     also nach A(i, 10)
                If vh = "-" Then
                    'weiterer Nachname existiert nicht
                    s = A(i, 10)                                        'Lea May
                    If A(i, 4) <> "" And A(i, 6) = "D" Then _
                        s = s + " (" + A(i, 4) + ")"                    'Lea May (Bonn)
                    If A(i, 4) <> "" And A(i, 6) <> "D" Then _
                        s = s + " (" + A(i, 4) + ", " + A(i, 6) + ")"   'Lea May (Bern, CH)
                    If A(i, 4) = "" And A(i, 6) <> "D" Then _
                        s = s + " (" + A(i, 6) + ")"                    'Lea May (CH)
                    A(i, 15) = s
                Else
                    'weiterer Nachname existiert
                    s = A(i, 10)                                        'Lea May  oder 'Lea Jun
                    If A(i, 1) = A(i, 9) Then
                        'Nn ist auch gn
                        If A(i, 4) <> "" And A(i, 6) = "D" Then _
                            s = s + " (" + A(i, 4) + ")"                    'Lea May (Bonn)
                        If A(i, 4) <> "" And A(i, 6) <> "D" Then _
                            s = s + " (" + A(i, 4) + ", " + A(i, 6) + ")"   'Lea May (Bern, CH)
                        If A(i, 4) = "" And A(i, 6) <> "D" Then _
                            s = s + " (" + A(i, 6) + ")"                    'Lea May (CH)
                        A(i, 15) = s
                    Else
                        'Nn ist ein vh-Name
                        s = s + " (geb " + A(i, 9)                          'Lea Jun (geb May
                        If A(i, 4) <> "" And A(i, 6) = "D" Then _
                            s = s + ", " + A(i, 4) + ")"                    'Lea Jun (geb May, Bonn)
                        If A(i, 4) <> "" And A(i, 6) <> "D" Then _
                            s = s + ", " + A(i, 4) + ", " + A(i, 6) + ")"   'Lea Jun (geb May, Bern, CH)
                        If A(i, 4) = "" And A(i, 6) <> "D" Then _
                            s = s + ", " + A(i, 6) + ")"                    'Lea Jun (geb May, CH)
                        A(i, 15) = s
                        
                    End If
                End If
            'A(i, 16)       'LiNnVn 'LinkName NnVn,     also nach A(i, 11)
                If vh = "-" Then
                    'weiterer Nachname existiert nicht
                    s = A(i, 11)                                        'Lea May
                    If A(i, 4) <> "" And A(i, 6) = "D" Then _
                        s = s + " (" + A(i, 4) + ")"                    'Lea May (Bonn)
                    If A(i, 4) <> "" And A(i, 6) <> "D" Then _
                        s = s + " (" + A(i, 4) + ", " + A(i, 6) + ")"   'Lea May (Bern, CH)
                    If A(i, 4) = "" And A(i, 6) <> "D" Then _
                        s = s + " (" + A(i, 6) + ")"                    'Lea May (CH)
                    A(i, 16) = s
                Else
                    'weiterer Nachname existiert
                    s = A(i, 11)                                        'Lea May  oder 'Lea Jun
                    If A(i, 1) = A(i, 9) Then
                        'Nn ist auch gn
                        If A(i, 4) <> "" And A(i, 6) = "D" Then _
                            s = s + " (" + A(i, 4) + ")"                    'Lea May (Bonn)
                        If A(i, 4) <> "" And A(i, 6) <> "D" Then _
                            s = s + " (" + A(i, 4) + ", " + A(i, 6) + ")"   'Lea May (Bern, CH)
                        If A(i, 4) = "" And A(i, 6) <> "D" Then _
                            s = s + " (" + A(i, 6) + ")"                    'Lea May (CH)
                        A(i, 16) = s
                    Else
                        'Nn ist ein vh-Name
                        s = s + " (geb " + A(i, 9)                          'Lea Jun (geb May
                        If A(i, 4) <> "" And A(i, 6) = "D" Then _
                            s = s + ", " + A(i, 4) + ")"                    'Lea Jun (geb May, Bonn)
                        If A(i, 4) <> "" And A(i, 6) <> "D" Then _
                            s = s + ", " + A(i, 4) + ", " + A(i, 6) + ")"   'Lea Jun (geb May, Bern, CH)
                        If A(i, 4) = "" And A(i, 6) <> "D" Then _
                            s = s + ", " + A(i, 6) + ")"                    'Lea Jun (geb May, CH)
                        A(i, 16) = s
                        
                    End If
                End If
            'A(i, 17)       'LiRn   'LinkName Rufname,  also nach A(i, 8)
                If A(i, 8) <> "" Then
                    'Rufname existiert
                    If vh = "-" Then
                        'weiterer Nachname existiert nicht
                        s = A(i, 8) + " (" + A(i, 10)                   'Fifi (Lea May
                        If A(i, 4) <> "" Then s = s + ", " + A(i, 4)    'Fifi (Lea May, Bern
                        If A(i, 6) <> "D" Then s = s + ", " + A(i, 6)   'Fifi (Lea May, Bern, CH
                        A(i, 17) = s + ")"                              'Fifi (Lea May, Bern, CH)
                    Else
                        'weiterer Nachname existiert
                        s = A(i, 8) + " (" + A(i, 2) + " " + A(i, 9)   'Fifi (Lea May
                        If A(i, 4) <> "" Then s = s + ", " + A(i, 4)    'Fifi (Lea May, Bern
                        If A(i, 6) <> "D" Then s = s + ", " + A(i, 6)   'Fifi (Lea May, Bern, CH
                        A(i, 17) = s + ")"                              'Fifi (Lea May, Bern, CH)
                    End If
                End If
        Next
    End With
End Sub

Function T5_Get_AllVnNn()
    Dim s$, i%, zLast&, Arr2D()
    zLast = Get_NrOfLastRowInColumnNr(3, "T5"): s = "|"
    With Sheets("T5")
        Arr2D = .Range(.Cells(8, 22), .Cells(zLast, 22)).Value
    End With
    For i = 1 To UBound(Arr2D, 1)
        s = s + Arr2D(i, 1) + "|"
    Next
    T5_Get_AllVnNn = s
    
    'show s
    
End Function

Sub T5_Load_OneColumn(Arr1D() As String, Col%)
    Dim zLast&, Arr2D()
    zLast = Get_NrOfLastRowInColumnNr(3, "T5")
    With Sheets("T5")
        Arr2D = .Range(.Cells(8, Col), .Cells(zLast, Col)).Value
    End With
    Load_StringArray1D_from_1ColumnArray2D Arr1D, Arr2D
End Sub

Sub T5_Load_Nn_to_Vh(Arr())
    Dim zLast&
    zLast = Get_NrOfLastRowInColumnNr(3, "T5")
    With Sheets("T5")
        Arr = .Range(.Cells(8, 3), .Cells(zLast, 14)).Value
    End With
End Sub

Sub Load_Array1D_OfColumnPart(ByRef Arr() As String, SheetName$, z1%, z2%, s%)
    'Action     belegt (1-basiert) das bereits definierte 1D-Array Arr
    '           mit Werten einer bestimmten Spalte
    
    'Vorbereitung
        Dim i%, A()
        With Sheets(SheetName)
    'Action
        A = .Range(.Cells(z1, s), .Cells(z2, s)).Value
        ReDim Arr(1 To UBound(A, 1))
        For i = 1 To UBound(A, 1)
            Arr(i) = A(i, 1)
        Next
    'Finals
        End With
End Sub

Sub Load_Array1D_OfColumnPartTrimmedLines(Arr, SheetName$, z1%, z2%, s%)
    'Action     belegt (1-basiert) das bereits definierte 1D-Array Arr
    '           mit Werten einer bestimmten Spalte
    
    'Vorbereitung
        Dim i%, A()
        With Sheets(SheetName)
    'Action
        A = .Range(.Cells(z1, s), .Cells(z2, s)).Value
        ReDim Arr(1 To UBound(A, 1))
        For i = 1 To UBound(A, 1)
            Arr(i) = Trim(A(i, 1))
        Next
    'Finals
        End With
End Sub

Sub Load_Array1D_OfColumnPart_NoEmptyRows(Arr, SheetName$, z1%, z2%, s%)
    'Action     belegt das bereits definierte 1D-Array Arr
    '           mit nichtleeren CellContents einer bestimmten Spalte
    
    'Vorbereitung
        Dim T$, i%, A()
        With Sheets(SheetName)
    'Action
        A = .Range(.Cells(z1, s), .Cells(z2, s)).Value
        For i = 1 To UBound(A, 1)
            If A(i, 1) <> "" Then T = T + vbCrLf + A(i, 1)
        Next
        Arr = Split(T, vbCrLf)
    'Finals
        End With
End Sub

Sub T7_Load_ArrHelpers(ByRef H)
    'Called from    T7_Fill_Helpers
    
    Dim zLast&
    'Fill H from helpers
        zLast = Get_NrOfLastRowInColumnNr(47, "T7")
        With Sheets("T7")
            H = .Range(.Cells(5, 46), .Cells(zLast, 54)).Value
        End With
End Sub

Sub T7_Load_ArrLbArchiv(ByRef L)
    'Called from    T7_Fill_Helpers
    
    Dim F$, p$, i%, Arr1() As String
    p = ArrC(1) + "\prog\Label\LbArchiv\"
    'Fill L from LbArchiv
    F = Get_AllFileNames_Like_OfOneFolder(p, "*.jpg")
    Arr1 = Split(vbCrLf + F, vbCrLf)
    ReDim L(1 To UBound(Arr1))
    For i = 1 To UBound(Arr1)
        L(i) = Arr1(i)
    Next
End Sub

Sub Load_ArrT4Row2(ByRef Arr)
    'Called from    Get_BuchstabenKette_T4Zeile2, T4_Add_T4Row2_ToArrDg
    With Sheets("T4")
        Arr = .Range(.Cells(2, CInt(ArrC(39))), .Cells(2, CInt(ArrC(40)))).Value
    End With
End Sub

Sub TimeTest()
    'Dim t1&: t1 = GetTickCount(): ... : ShowTime t1
    Dim i%, t1&, ArrEventFolderPaths() As String
    t1 = GetTickCount()
    For i = 1 To 100
        Load_ArrPathsOfAllFolders_Events ArrEventFolderPaths
    Next
    ShowTime t1
    showArray ArrEventFolderPaths
End Sub

Sub Load_ArrPathsOfAllFolders_Events(Arr)
    'Called from    TimeTest
    'ArrEv..Paths   = 1D-Array, 1-basiert, alle EventFolderPaths (auch Liga)
    'Action         ermittelt Arr-Einträge
    'Duration       0,12 sec
    
    'Vorbereitung
        Dim s1$, s2$, C1%, C2%, C3%, i%, F, f1, fc, fso As Object
        Dim Arr1() As String, Arr2() As String, Arr3() As String
        ReDim Arr1(1 To 999): ReDim Arr2(1 To 99): ReDim Arr3(1 To 999): DoArrc
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set F = fso.GetFolder(ArrC(3))
        Set fc = F.subfolders
    'Sammeln aller EventOrdner OneLevel
        For Each f1 In fc
            If f1.path Like "*\#### Liga" Then
                C2 = C2 + 1: Arr2(C2) = f1.path    'arr2 sammelt Pfade von EventOrdnern wie "1982 Liga"
            Else
                C1 = C1 + 1: Arr1(C1) = f1.path    'arr1 sammelt Pfade aller anderen EventOrdner
            End If
        Next
    'Sammeln aller LigaOrdner OneLevel (...\Events\1982 Liga\1982 Buli1Süd Gernsbach-Bruchsal)
        For i = 1 To C2
            Set fso = CreateObject("Scripting.FileSystemObject")
            s1 = Arr2(i)                    '...\1973 Liga
            Set F = fso.GetFolder(s1)
            Set fc = F.subfolders
            For Each f1 In fc
                C3 = C3 + 1: Arr3(C3) = f1.path
            Next
        Next
    'Arr = arr1-Ordner + arr3-Ordner
        ReDim Arr(1 To C1 + C3)
        For i = 1 To C1
            Arr(i) = Arr1(i)
        Next
        For i = C1 + 1 To C1 + C3
            s2 = Arr3(i - C1)
            Arr(i) = s2
        Next
End Sub

Sub T4_Load_ArrDgTitles_FromT4ListOfDesignTitles(ByRef ArrDgTitles)
    Dim i%, z2&, Arr1(), ws As Worksheet, rng As Range
    Set ws = Sheets("T4"): DoArrc: z2 = Get_NrOfLastRowInColumnNr(CInt(ArrC(72)), "T4")
    Arr1 = ws.Range(ws.Cells(CInt(ArrC(71)) + 3, CInt(ArrC(72))), ws.Cells(z2, CInt(ArrC(72))))
    ReDim ArrDgTitles(1 To UBound(Arr1, 1))
    For i = 1 To UBound(Arr1, 1)
        ArrDgTitles(i) = Arr1(i, 1)
    Next
End Sub

Sub T4_Load_ArrDgTitles_FromDgs(ByRef ArrDgTitles)
    'ArrDgTitles    1D-Array, 1-basiert
    Dim i%, Arr1() As String
    Arr1 = Split(vbCrLf + Get_AllDgTitles_FromDgs, vbCrLf)
    'arr1 ist 0-basiert
    ReDim ArrDgTitles(1 To UBound(Arr1))
    For i = 1 To UBound(Arr1)
        ArrDgTitles(i) = Arr1(i)
    Next
    QuickSort ArrDgTitles
End Sub

Sub T4_Load_ArrDgTitles_FromT4SomeDgData(ByRef Arr1D() As String)
    'Called from    xxx
    'ArrDgTitles    1D-Array, 1-basiert; wird hier gefüllt
    
    'Vorbereitung
        Dim i%, zLast%, Arr2D()
        zLast = Get_NrOfLastRowInColumnNr(5, "T4")
    'Load 2D
        With Sheets("T4")
            Arr2D = .Range(.Cells(8, 5), .Cells(zLast, 5)).Value
        End With
        'showArray2D Arr2D: Stop
    'Load 1D
        Load_StringArray1D_from_1ColumnArray2D Arr1D, Arr2D
        'showArray Arr1D
End Sub

Sub Load_ArrNamesOfAllFolders_Leute(ByRef Arr() As String)
    Dim NamesOfLeuteFolders$: DoArrc
    NamesOfLeuteFolders = Get_Names_OfAllSubfolders_OneLevel(ArrC(4)) 'Folder "Leute"
    Arr = Split(NamesOfLeuteFolders, vbCrLf)
End Sub



Sub Load_ArrPathsOfAllFolders_Register(ByRef Arr() As String)
    Dim p$
    p = Get_Paths_ofAllSubfolders_OneLevel(ArrC(5)) 'Folder "Register"
    Arr = Split(p, vbCrLf)
End Sub



Sub Load_ArrPathsOfAllFolders_Leute(ByRef Arr() As String)
    Dim p$
    p = Get_Paths_ofAllSubfolders_OneLevel(ArrC(4)) 'Folder "Leute"
    Arr = Split(p, vbCrLf)
End Sub

Sub Load_ArrPathsOfAllFolders_Leute_ClubsNations(ByRef Arr() As String)
    Dim p1$, p2$
    p1 = Get_Paths_ofAllSubfolders_OneLevel(ArrC(4))    'Folder "Leute"
    p2 = Get_Paths_OfAllSubfolders_AllLevels(ArrC(11))  'Folder "ClubsNations"
    Arr = Split(p1 + vbCrLf + p2, vbCrLf)
End Sub

Sub Load_ArrPathsOfAllFolders_ClubsNations(ByRef Arr() As String)
    Dim F$
    F = Get_Paths_OfAllSubfolders_AllLevels(ArrC(11)) 'Folder "ClubsNations"
    Arr = Split(F, vbCrLf)
End Sub

Sub Load_ArrPathsOfAllFolders_Archiv(ByRef Arr() As String)
    Dim F$
    F = Get_Paths_OfAllSubfolders_AllLevels(ArrC(2)) 'Folder "Archiv Trampolin 1900-1999
    Arr = Split(F, vbCrLf)
End Sub

Sub Load_ArrDesign(ByRef ArrDesign())

    'FillArrDg?
    
    'Called from    Get_CharsOfOneActivatedDesign, T4_AutoFitCalc
    '               Add_Age_to_EachDesignName, Fill_ArrCCn, Fill_ListOfCompetitors
    'Status         Dg trägt nur Rohdaten, no Adds, no Checks
    
    Dim ws As Worksheet, rng As Range
    Set ws = Sheets("T4"): Set rng = ws.Range(ArrC(36)) 'Range of choosen design
    ArrDesign = rng.Value
End Sub

Sub T5_Load_Nn_to_Jhg(ByRef T5())
    'Called from    Update_Age_inDgs
    
    Dim zLast&
    zLast = Get_NrOfLastRowInColumnNr(3, "T5")
    With Sheets("T5")
    T5 = .Range(.Cells(8, 3), .Cells(zLast, 6)).Value
    End With
End Sub

Sub T4_Load_CompleteSheet(T4)
    'Called from    xxx
    
    'Vorbereitung
        Dim s&, z&, ws As Worksheet, r As Range
        Set ws = Sheets("T4") ': ws.Activate
        s = lastCol(ws): z = LastRow(ws)
    'Action
        T4 = ws.Range(ws.Cells(1, 1), ws.Cells(z, s)).Value
End Sub

Sub Load_ArrCompetitorsList(ByRef ArrComp())
    'Called from    Add_PersonData_ToCompetitorsList, Get_AllNames,
    '                   Reconstruct_PersonDataTxt_2_FromCompetitorsList
    'arrComp        2D-Array, T4-CompetitorsList (ohne Kopfzeile etc.)
    xxx
    'Vorbereitung
        Dim s1&, s2&, z1%, z2%, ws As Worksheet, r As Range
        z1 = ArrC(73): s1 = ArrC(74): s2 = ArrC(75): z2 = Get_NrOfLastRowInColumnNr(s1, "T4")
        Set ws = Sheets("T4"): ws.Activate
    'Fill arrComp
        Set r = ws.Range(ws.Cells(z1 + 3, s1), ws.Cells(z2, s2))
        ArrComp = r.Value 'T4 competitors list
End Sub

Sub T5_Load_OneLine(ByRef A() As String)
    'Called from    Check_T6TextBoxChange
    
    Dim N$, i%, z&, B() As String, c()
    With Sheets("T5")
    T6_Load_SixCellsLeftOfX B
    N = B(2) + " " + B(1) 'Name: Vn Nn
    z = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, N)
    c() = .Range(.Cells(z, 1), .Cells(z, 22)).Value
    ReDim A(1 To 22)
    For i = 1 To 22
        A(i) = CStr(c(1, i))
    Next
    End With
End Sub

Function Get_BuchstabenKette_T4Zeile2()
        Dim B$, i%, ArrZ2()
    'ArrZ2
        Load_ArrT4Row2 ArrZ2
        'showArray2D ArrZ2, "ArrZ2"
    'B
        For i = 1 To UBound(ArrZ2, 2)
            If ArrZ2(1, i) = "" Then B = B + "-" Else B = B + ArrZ2(1, i)
        Next
        B = Replace(B, " ", "") 'wegen "i " und " i"
        Get_BuchstabenKette_T4Zeile2 = B
        'show B
End Function


