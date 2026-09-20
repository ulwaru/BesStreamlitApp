Attribute VB_Name = "M_HelperBox"
Option Explicit

'In T4  Private Sub CellBox_KeyDown
Sub xxxMove_CellBox_toMyPosition()
    
    Dim cL!, cT!, cW!, cH!
    
    With Sheets("T4")
    'CellValues der angeklickten Zelle
        With .Cells(19, 17)
            cL = .Left: cT = .Top: cW = .Width: cH = .Height            'c CellValues
        End With
    'CellBox
        With .Shapes("CellBox")
            .Left = cL: .Top = cT: .Width = cW + 1: .Height = cH + 5    '
        End With
        CellBox.Value = Empty: CellBox.Visible = True: CellBox.Activate
    End With
End Sub

Sub Move_CellBox_toMyPosition()
    
    Dim cL As Single, cT As Single, cW As Single, cH As Single
    Dim shpCellBox As shape
    
    With Sheets("T4")
        ' 1. Position der Zelle auslesen (Q19)
        With .Cells(19, 17)
            cL = .Left
            cT = .Top
            cW = .Width
            cH = .Height
        End With
        
        ' 2. Shape-Referenz zuweisen
        Set shpCellBox = .Shapes("CellBox")
        
        ' 3. Shape positionieren, leeren, sichtbar machen und aktivieren
        With shpCellBox
            .Left = cL
            .Top = cT
            .Width = cW + 1
            .Height = cH + 5
            
            ' Textinhalt leeren (funktioniert für Textfelder & Zeichnungsformen)
            .TextFrame2.TextRange.text = ""
            
            ' Sichtbar machen
            .Visible = True
            
            ' Aktivieren/Markieren
            .Select
        End With
    End With
    
End Sub

Sub HelperBoxInit()
    'Called from    T4_SearchForBorders
    'HelperBox      = ListBox; trägt herausgefilterte Namen
    'Action         prüft, ob der UserKlick auf eine leere Zelle erfolgte,
    '               die zum NamenBereich (weiblich/männlich/mixed) gehört;
    '               falls ja, werden die Positionen von CellBox und HelperBox
    '               festgelegt, die CellBox angezeigt
    'ToDo           Kill FillArrC 65
        
    'Vorbereitung
        Dim MWX$, s%, s1%, sD%, sH%, z%, z1%, z2%, cL!, cT!, cW!, cH!
        Dim HelperBox As MSForms.ListBox, Arr1() As String, Target As Range
        Dim CellBox As MSForms.TextBox 'Shape
        With Sheets("T4")
        z = CInt(ArrC(104))          'T4 SelectionChange to ZeilenNr  z
        s = CInt(ArrC(105))          'T4 SelectionChange to SpaltenNr s
        Set Target = .Cells(z, s)   'T4 SelectionChange to Target
        z1 = CInt(ArrC(37))         'T4 SheetZeilenNr  of BorderTop         '171
        z2 = CInt(ArrC(38))         'T4 SheetZeilenNr  of BorderBottom      '239
        s1 = CInt(ArrC(39))         'T4 SheetSpaltenNr of BorderLeft        '247
    'Exit
        'If Target.Cells.count <> 1 Then Exit Sub
        If ActiveSheet.Cells(z, s).Interior.Color = vbWhite Then Exit Sub
        MWX = .Cells(2, s)
        If Not (MWX = "W" Or MWX = "M" Or MWX = "X") Then Exit Sub
        If z < z1 + 5 Then Exit Sub
        If z > z2 - 1 Then Exit Sub
        If Target.Value2 <> Empty Then Exit Sub
    'CellBox HelperBox
        Set CellBox = .CellBox
        Set HelperBox = .HelperBox:     Hide_HelperBox
    'Array für HelperBox bereitstellen
        If MWX = "M" Then FillArrC 65, "M": Fill_ArrM_Only
        If MWX = "W" Then FillArrC 65, "W": Fill_ArrW_Only
        If MWX = "X" Then FillArrC 65, "X": Fill_ArrX_Only
    'CellValues der angeklickten Zelle
        With .Cells(z, s)
            cL = .Left: cT = .Top: cW = .Width: cH = .Height            'c CellValues
        End With
    'CellBox
        With .Shapes("CellBox")
            .Left = cL: .Top = cT: .Width = cW + 1: .Height = cH + 5    '
        End With
        CellBox.Value = Empty: CellBox.Visible = True: CellBox.Activate
        'Repair
            Sheets("T4").Shapes("CellBox").Top = cT
    'HelperBox
        With .Shapes("HelperBox")
            .Left = cL + 0.5 * cW: .Top = cT + 18: .Width = 150: .Height = 150
        End With
    'Finals
        End With
End Sub

Sub Fill_ArrHelperBox(SomeKeys$, ByRef arrFilter)
    'Called from    CellBox_KeyDown
    'ArrFilter      = zu füllendes 1D-Array, 0-basiert; passend zu SomeKeys
    'SomeKeys       = bisher in die CellBox eingegebene Buchstaben
    'ArrX ArrW ArrM = 1D-Arrays: Aktive/Weiblich/Männlich stehen bereit
    'Action         Das Ergebnis dieses Macros ist eine NamenListe
    '               die in der HelperBox angezeigt wird (ListBox "HelperBox")
    
    Dim N$, L$, L2$, s$, s2$, Anz%, i%
    Anz = Len(SomeKeys): DoArrW
    
'showArray ArrW: Stop

    'L = Like-Inhalt
        If Anz > 0 Then
            L = UCase(SomeKeys)         'CHRISTIA  oder  SV
        End If
    's = gefilterte Liste
        If ArrC(65) = "X" Then
            For i = 1 To UBound(ArrX)
                If UCase(ArrX(i)) Like "*|" + L + "*" Then s = s + ArrX(i) + vbCrLf
                If UCase(ArrX(i)) Like "* " + L + "*" Then s = s + ArrX(i) + vbCrLf
                If Anz > 1 Then
                    L2 = Left(L, Len(L) - 1) + "* " + Right(L, 1) + "*"
                    If UCase(ArrX(i)) Like L2 Then s2 = s2 + ArrX(i) + vbCrLf
                End If
            Next
        End If
        If ArrC(65) = "W" Then
            For i = 1 To UBound(ArrW)
                If UCase(ArrW(i)) Like "*|" + L + "*" Then
                    s = s + ArrW(i) + vbCrLf
                ElseIf UCase(ArrW(i)) Like "* " + L + "*" Then s = s + ArrW(i) + vbCrLf
                End If
                If Anz > 1 Then
                    L2 = Left(L, Len(L) - 1) + "* " + Right(L, 1) + "*"
                    If UCase(ArrW(i)) Like L2 Then s2 = s2 + ArrW(i) + vbCrLf
                End If
            Next
        End If
        If ArrC(65) = "M" Then
            For i = 1 To UBound(ArrM)
                If UCase(ArrM(i)) Like "*|" + L + "*" Then
                    s = s + ArrM(i) + vbCrLf
                ElseIf UCase(ArrM(i)) Like "* " + L + "*" Then s = s + ArrM(i) + vbCrLf
                End If
                If Anz > 1 Then
                    L2 = Left(L, Len(L) - 1) + "* " + Right(L, 1) + "*"
                    If UCase(ArrM(i)) Like L2 Then s2 = s2 + ArrM(i) + vbCrLf
                End If
            Next
        End If
        s = s + vbCrLf + s2
        s = Delete_EmptyRowsInString(s)
        s = TAB_Simulation(s)   '|Alison   |Pester       |CAN|Kingston           |1963|w
        '                       '|Andrea   |Holmes       |GB |Dunstable          |1970|w

        
    If s = "" Then
        Sheets("T4").Shapes("HelperBox").Visible = False
        s = "No names"
        arrFilter = Split(s, vbCrLf)
    Else
        'show s
        'arrFilter als 1D-Array
            arrFilter = Split(s, vbCrLf)
            'sort
                For i = 0 To UBound(arrFilter)
                    If InStr(1, UCase(arrFilter(i)), L) = 2 Then _
                         arrFilter(i) = "1" + arrFilter(i) _
                    Else arrFilter(i) = "2" + arrFilter(i)
                Next
                QuickSort arrFilter
                For i = 0 To UBound(arrFilter)
                    arrFilter(i) = Mid(arrFilter(i), 3)
                Next
    End If
    'Repair
        'T4_Replace_Helperbox
    
    'showArray arrFilter
End Sub

Sub Fill_ArrX_Only()
    'Called from    HelperBoxInit
    'ArrX           globales Array
    
    'Vorbereitung
        Dim A$, i%, zLast&, B()
        zLast = Get_NrOfLastRowInColumnNr(3, "T5")
    'T5 in Array laden
        With Sheets("T5")
            B = Range(.Cells(8, 3), .Cells(zLast, 15)).Value
        End With
    'Fill
        For i = 1 To UBound(B, 1)
            A = A + "|" + CStr(B(i, 2)) + "|" + CStr(B(i, 1)) + "|" _
            + CStr(B(i, 7)) + "|" + CStr(B(i, 5)) + "|" + CStr(B(i, 4)) + "|" + CStr(B(i, 3)) _
            + "|" + CStr(B(i, 13)) + "|" + vbCrLf
        Next
        'show a
        ArrX = Split(A, vbCrLf)
End Sub

Sub Fill_ArrW_Only()
    'Called from    xxx
    'ArrW           globale Variable
    '               Zeilen wie "|Lea|Mai|D|Bonn|1962|w|May|"
    
    'Vorbereitung
        Dim w$, Nn$, Vn$, i&, zLast&, B()
        zLast = Get_NrOfLastRowInColumnNr(3, "T5"): With Sheets("T5")
    'T5 in Array laden
        B = Range(.Cells(8, 3), .Cells(zLast, 15)).Value
    'Fill
        For i = 1 To UBound(B, 1)
            Vn = CStr(B(i, 2)): Nn = CStr(B(i, 1))
            If B(i, 3) = "w" Then
                If Not w Like "*|" + Vn + "|" + Nn + "|*" Then w = w + "|" + Vn + "|" + Nn + "|" _
                + CStr(B(i, 7)) + "|" + CStr(B(i, 5)) + "|" + CStr(B(i, 4)) + "|" + CStr(B(i, 3)) _
                + "|" + CStr(B(i, 13)) + "|" + vbCrLf
            End If
        Next
        ArrW = Split(w, vbCrLf)
    'Finals
        End With
        'show w
        'QuickSort ArrW
        'showArray ArrW
End Sub

Sub Fill_ArrM_Only()
    'Called from    xxx
    
    'Vorbereitung
        Dim m$, Nn$, Vn$, i%, zLast&, B()
        zLast = Get_NrOfLastRowInColumnNr(3, "T5"): With Sheets("T5")
    'T5 in Array laden
        B = Range(.Cells(8, 3), .Cells(zLast, 15)).Value
    'Fill
        For i = 1 To UBound(B, 1)
            If B(i, 3) = "m" Then
                Vn = CStr(B(i, 2)): Nn = CStr(B(i, 1))
                If Not m Like "*|" + Vn + "|" + Nn + "|*" Then
                    m = m + "|" + Vn + "|" + Nn + "|" _
                    + CStr(B(i, 7)) + "|" + CStr(B(i, 5)) + "|" + CStr(B(i, 4)) + "|" + CStr(B(i, 3)) _
                    + "|" + CStr(B(i, 13)) + "|" + vbCrLf
                End If
            End If
        Next
        ArrM = Split(m, vbCrLf)
    'Finals
        End With
        'showArray ArrM
End Sub

Sub Fill_ArrX_ArrW_ArrM_FromT5()
    'Called from    DoArrW
    'ToDo           Einbau in Update_T4CompetitorsList_LLinksInEFolder
    'ArrX          =Public 1D-String-Array; alle        PersonenNamen aus T5 (PersonData)
    'ArrW          =Public 1D-String-Array; alle weibl. PersonenNamen aus T5 (PersonData)
    'ArrM          =Public 1D-String-Array; alle männl. PersonenNamen aus T5 (PersonData)
    
    'Vorbereitung
        Dim A$, m$, v$, w$, i%, zLast&, B()
        zLast = Get_NrOfLastRowInColumnNr(3, "T5"): v = vbCrLf: With Sheets("T5")
    'T5 in Array laden
        B = .Range(.Cells(8, 3), .Cells(zLast, 15)).Value
    'ArrX, ArrW, ArrM
        For i = 1 To UBound(B, 1)
            'ArrX       'Lea|Mai|D|Bonn|1962|w|May|
                A = A + "|" + CStr(B(i, 2)) + "|" + CStr(B(i, 1)) + "|" _
                + CStr(B(i, 7)) + "|" + CStr(B(i, 5)) + "|" + CStr(B(i, 4)) + "|" + CStr(B(i, 3)) _
                + "|" + CStr(B(i, 13)) + "|" + v
            'ArrW
                If B(i, 3) = "w" Then w = w + "|" + CStr(B(i, 2)) + "|" + CStr(B(i, 1)) + "|" _
                + CStr(B(i, 7)) + "|" + CStr(B(i, 5)) + "|" + CStr(B(i, 4)) + "|" + CStr(B(i, 3)) _
                + "|" + CStr(B(i, 13)) + "|" + v
            'ArrM
                If B(i, 3) = "m" Then m = m + "|" + CStr(B(i, 2)) + "|" + CStr(B(i, 1)) + "|" _
                + CStr(B(i, 7)) + "|" + CStr(B(i, 5)) + "|" + CStr(B(i, 4)) + "|" + CStr(B(i, 3)) _
                + "|" + CStr(B(i, 13)) + "|" + v
        Next
        ArrX = Split(A, v): QuickSort ArrX
        ArrW = Split(w, v): QuickSort ArrW
        ArrM = Split(m, v): QuickSort ArrM
    'Finals
        End With
        'show w
        'showArray ArrM
End Sub

Sub T4_Replace_Helperbox()
    Dim t1!, T2!, sSh%, zSh%
    zSh = CInt(ArrC(104))        'T4    ActualSelection SheetZeilenNr
    sSh = CInt(ArrC(105))        'T4    ActualSelection SheetSpaltenNr
    
    With Sheets("T4")
        .Shapes("CellBox").Left = .Cells(zSh, sSh).Left
        .Shapes("HelperBox").Left = .Cells(zSh, sSh).Left + .Shapes("CellBox").Width + 3
        t1 = .Shapes("CellBox").Top
        T2 = .Shapes("HelperBox").Top
        .Shapes("HelperBox").Top = t1
    End With
End Sub


