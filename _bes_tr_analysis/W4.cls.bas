Attribute VB_Name = "W4"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit 'MODUL: W4

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim s As Integer, z As Integer          's=Spalte, z=Zeile
    s = Target.Column: z = Target.Row       'momentan ausgewählte Zelle ermitteln
    If s = 2 Then doProc z                         'der Benutzer klickt ggf. auf ein Kästchen
    '    Select Case z
    '        Case 2: ListModules: Case 3: ListReferences: Case 4: ListProcedures: _
    '        Case 6: ListProcesses: Case 7: Color: Case 8: ColorIndex: _
    '        Case 9: ListSysInfo: Case 5: ListShapes ': Case 10: xxx
    '    End Select
    '    formatiereSpalteJ
    'End If
    If s = 10 And [B4] = "X" And s > 3 And Cells(z, s) <> "" Then
        Dim r As Long
        r = CodeFind(, Cells(z, s).Value)
    End If
End Sub

Sub doProc(z As Integer)
    Dim s As String, Arr1() As String, i As Integer
    'ggf. Reihenfolge ändern oder ergänzen
    s = s + "ListModules,ListReferences,ListProcedures,ListShapes,ListProcesses,Color,"
    s = s + "ColorIndex,ListSysInfo"
    Arr1 = Split(s, ",")
    If z < 2 Or z > 2 + UBound(Arr1) Then Exit Sub
    
    'Angebot neu schreiben
    ActiveSheet.Range("b:c").Clear
    For i = 0 To UBound(Arr1)
        Cells(2 + i, 3) = " " + Arr1(i) 'Start Kästchen in B2, ProcName in C2
        Cells(2 + i, 2).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium 'xlThick
    Next
    Cells(z, 2) = "X": Cells(z, 2).HorizontalAlignment = xlCenter: Cells(z, 2).Font.Bold = True
    ActiveSheet.Range("j:k").Clear
    ActiveSheet.Range("i:i").ClearContents
    Application.Run "W4." + Arr1(z - 2)
    formatiereSpalteJ
End Sub

Sub ListModules()
    Dim Arr1() As String, i As Integer
    Arr1 = Split(getListModules, vbCrLf)
    For i = 0 To UBound(Arr1)
        Cells(2 + i, 10) = Arr1(i)
        If i > 3 Then Cells(i, 9) = i - 3
    Next
    If i > 3 Then Cells(i, 9) = i - 3
    ActiveSheet.Range(Cells(4, 10), Cells(UBound(Arr1) + 1, 10)).Interior.Color = 10086143
End Sub

Sub ListReferences()
    Dim Arr1() As String, i As Integer
    Arr1 = Split(getReferences, vbCrLf)
    For i = 0 To UBound(Arr1)
        Cells(2 + i, 10) = Arr1(i)
        If i > 3 Then Cells(i, 9) = i - 3
    Next
    If i > 3 Then Cells(i, 9) = i - 3
    ActiveSheet.Range(Cells(4, 10), Cells(UBound(Arr1) + 1, 10)).Interior.Color = 10086143
End Sub

Sub ListProcedures()
    Dim arrMod() As String, Arr1() As String, Arr2() As String, i As Integer, c As Integer, C2 As Integer
    Dim s As String, s1 As String, s2 As String, s3 As String, LZ10 As Long
    arrMod = Split(getListModules, vbCrLf)
    Arr1 = Split(getFunctionAndSubNames, vbCrLf)

    For i = 0 To UBound(Arr1)
        s = Arr1(i) 'W3.myProc
        If s <> "" Then
            c = InStr(1, s, ".")
            s1 = s1 + Mid(s, c + 1) + " [" + Left(s, c - 1) + "] #° " 'myProc [W3] | ...
        End If
    Next
    For i = 2 To UBound(arrMod)
        s2 = arrMod(i): s3 = ""
        C2 = getAnzahlStringAinStringB("|", s2)
        If C2 > 1 Then
            s3 = Mid(s2, InStr(1, s2, ") |") + 4)       'W4 | "W4"
            s3 = Left(s3, InStr(1, s3, "|") - 2)        'W4
        Else
            s3 = Mid(s2, InStr(1, s2, "|") + 2)         'M01
        End If
        s1 = Replace(s1, "[" + s3 + "]", "[" + s2 + "]")
    Next
    Arr2 = Split(s1, " #° ")
    
    For i = 0 To UBound(Arr2)
        s2 = Arr2(i)
        If s2 = "" Then Exit For
        c = InStr(1, s2, " [")
        Cells(4 + i, 10) = Left(s2, c - 1)
        Cells(4 + i, 11) = Mid(s2, c)
    Next
    LZ10 = LetzteZeileInSpalteNr(10)
    [J2] = CStr(LZ10 - 3) + " Prozeduren in ActiveWorkbook:"
    nummeriereVonZeileBisZeileInSpalteNr 4, LZ10, 9
    ActiveSheet.Range(Cells(4, 10), Cells(LZ10, 10)).Interior.Color = 10086143
    ActiveSheet.Range(Cells(4, 10), Cells(LZ10, 11)).Sort Key1:=Range("j4")
End Sub

Sub ListProcesses()
    Dim Arr1() As String, i As Integer
    Arr1 = Split(getProcessList, vbCrLf)
    For i = 0 To UBound(Arr1)
        Cells(2 + i, 10) = Arr1(i)
        If i > 3 Then Cells(i, 9) = i - 3
    Next
    If i > 3 Then Cells(i, 9) = i - 3
    ActiveSheet.Range(Cells(4, 10), Cells(UBound(Arr1) + 1, 10)).Interior.Color = 10086143
    ActiveSheet.Range(Cells(4, 10), Cells(UBound(Arr1) + 1, 10)).Sort Key1:=Range("j4")
End Sub

Sub MyColorInfo()
    r.Interior.Color = RGB(200, 200, 200)
    If getColor(r) = 14348258 Then ClickOnGreen1 z, s
    'Green1 14348258 'Green2 11854022 'Green3 9359529 'Green4 3506772
    r.Interior.Color = 15921906 '(x-Farbe, hellgrau)
    r.Interior.Color = 14083324 '(w-Farbe, hellrot)
    r.Interior.Color = 15652797 '(m-Farbe, hellblau)
    r.Interior.ColorIndex = 0   'keine Füllung
    r.Interior.ColorIndex = 1   'Schwarz
    r.Interior.ColorIndex = 2   'Weiß
    r.Interior.ColorIndex = 3   'rot
    r.Font.Color = vbWhite
    r.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
    
    
    
End Sub

Sub Color()
    Dim i As Integer
    [J2] = "Interior.Color - Beispiele"
    Cells(4, 10).Interior.Color = RGB(200, 200, 200)
    Cells(4, 10) = "Cells(4, 10).Interior.Color = RGB(200, 200, 200)"
    Cells(5, 10).Interior.Color = RGB(128, 128, 128)
    Cells(5, 10) = "Cells(5, 10).Interior.Color = RGB(100, 100, 100)"
    Cells(6, 10).Interior.Color = 10086143
    Cells(6, 10) = "Cells(6, 10).Interior.Color = 10086143"
    Cells(7, 10).Interior.Color = RGB(255, 230, 153)
    Cells(7, 10) = "Cells(7, 10).Interior.Color = RGB(255, 230, 153)"
    Cells(8, 10).Interior.Color = RGB(169, 208, 142)
    Cells(8, 10) = "Cells(8, 10).Interior.Color = RGB(169, 208, 142)"
    For i = 10 To 15
        Cells(i, 10).Interior.Color = RGB(265 - i * (i - 9), 0, 0)
        Cells(i, 10) = "Cells(" + CStr(i) + ", 10)" + ".Interior.Color " + _
            "= RGB(" + CStr(265 - i * (i - 9)) + ", 0, 0)"
        'Cells(i, 9) = i - 3
    Next
End Sub

Sub ColorIndex()
    Dim i As Integer
    [J2] = "Interior.Color.Index"
    For i = 0 To 56
        Cells(4 + i, 10).Interior.ColorIndex = i
        Cells(4 + i, 10) = "Cells(" + CStr(4 + i) + ", 10)" + ".Interior.ColorIndex = " + CStr(i)
        Cells(4 + i, 9) = i
    Next
End Sub

Sub ListSysInfo()
    Dim s As String
    [J2] = "System- und Excel-Informationen:"
    [J4] = "Betriebssytem:"
    [J5] = "   " + Application.OperatingSystem
    [J6] = "Excel:"
    'Eine 64-bit-Version gibt es erst mit Excel 2010.
    'Alle anderen Versionen vor Excel 95 16-bit danach 32-bit-Versionen.
    [J7] = "   " + ExcelVersion()
End Sub

Sub Get_ListOfShapes_x_y_TEST()
    show Get_ListOfShapes_x_y("T4", "RoRe")
End Sub

Function Get_ListOfShapes_x_y(NameOfSheet$, NameOfShape$) As String
    Dim s$, v$, ws As Worksheet, shp As shape
    v = vbCrLf: Set ws = Sheets(NameOfSheet)
    For Each shp In ws.Shapes
        If shp.NAME = NameOfShape Then
            s = s + "Sheet '" + NameOfSheet + "', Shape '" + NameOfShape + "' at (" + Format(shp.Left, "00000.00") + "|" + Format(shp.Top, "00000.00") + ")" + v
        End If
    Next
    'show s
    Get_ListOfShapes_x_y = s
End Function

Sub Get_CountOfShapes_TEST()
    show CStr(Get_CountOfShapes("T4", "RoRe"))
End Sub

Function Get_CountOfShapes(NameOfSheet$, NameOfShape$) As Integer
    Dim c%, ws As Worksheet, shp As shape
    Set ws = Sheets(NameOfSheet)
    For Each shp In ws.Shapes
        If shp.NAME = NameOfShape Then c = c + 1
    Next
    'show s
    Get_CountOfShapes = c
End Function

Sub ListShapes()
    Dim s As String, ws As Worksheet, sH As shape, z As Long, LZ10 As Long
    [J2] = "Shapes dieses Workbook:": z = 3
    For Each ws In ThisWorkbook.Sheets
            For Each sH In ws.Shapes
                z = z + 1
                Sheets("W4").Cells(z, 10) = sH.NAME
                Sheets("W4").Cells(z, 11) = "in sheet(" + Chr(34) + ws.NAME + Chr(34) + "); " _
                      + "sh.Type = " + CStr(sH.Type) + " (" + TypeName(sH.Type) + "); ID = " + CStr(sH.Id)
            Next sH
    Next ws
    
    LZ10 = LetzteZeileInSpalteNr(10)
    [J2] = CStr(LZ10 - 3) + " Shapes in ActiveWorkbook:"
    nummeriereVonZeileBisZeileInSpalteNr 4, LZ10, 9
    ActiveSheet.Range(Cells(4, 10), Cells(LZ10, 10)).Interior.Color = 10086143
    ActiveSheet.Range(Cells(4, 10), Cells(LZ10, 11)).Sort Key1:=Range("j4")
End Sub

Sub formatiereSpalteJ()
    Columns("K:Z").ColumnWidth = 2
    Columns("i:k").EntireColumn.AutoFit
    If Columns("i:i").ColumnWidth < 2 Then Columns("i:i").ColumnWidth = 2
    ActiveSheet.Range("J2").HorizontalAlignment = xlCenter
    ActiveSheet.Range("J2").Interior.Color = 10086143
End Sub

Sub nummeriereVonZeileBisZeileInSpalteNr(vonZeile As Long, bisZeile As Long, SpaltenNr As Long)
    Dim i As Long
    For i = vonZeile To bisZeile
        Cells(i, SpaltenNr) = i - vonZeile + 1
    Next
End Sub

Function LetzteZeileInSpalteNr(SpNr As Long)
    LetzteZeileInSpalteNr = ActiveSheet.Cells(Rows.count, SpNr).End(xlUp).Row
End Function

Public Function TypeName(Nr) As String
    Dim T As String
    On Error Resume Next
    Select Case Nr
        Case 1: T = "msoAutoShape"
        Case 2: T = "msoCallout"
        Case 20: T = "msoCanvas"
        Case 3: T = "msoChart"
        Case 4: T = "msoComment"
        Case 21: T = "msoDiagram"
        Case 7: T = "msoEmbeddedOLEObject"
        Case 8: T = "msoFormControl"
        Case 5: T = "msoFreeform"
        Case 6: T = "msoGroup"
        Case 9: T = "msoLine"
        Case 10: T = "msoLinkedOLEObject"
        Case 11: T = "msoLinkedPicture"
        Case 16: T = "msoMedia"
        Case 12: T = "msoOLEControlObject"
        Case 13: T = "msoPicture"
        Case 14: T = "msoPlaceholder"
        Case 18: T = "msoScriptAnchor"
        Case -2: T = "msoShapeTypeMixed"
        Case 19: T = "msoTable"
        Case 17: T = "msoTextBox"
        Case 15: T = "msoTextEffect"
    End Select
    TypeName = T
End Function

