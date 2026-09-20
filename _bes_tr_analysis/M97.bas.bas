Attribute VB_Name = "M97"
Option Explicit 'M97

'Get_NamesOfAllOpenWindows
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowTextLengthA Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As LongPtr) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
'CursorPos
    Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    'GetCursorPos requires a variable declared as a custom data type
      ' that will hold two integers, one for x value and one for y value
      Type POINTAPI: xPos As Long: yPos As Long: End Type
'WIA_ConvertImage
    Public Enum wiaFormat: BMP = 0: gif = 1: JPeG = 2: PNG = 3: TIFF = 4: End Enum

Public Sub QuickSort(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1)
    On Error Resume Next
    'Sort a 1-Dimensional array
    ' SampleUsage: sort arrData
    '   QuickSortVector arrData
    ' Originally posted by Jim Rech 10/20/98 Excel.Programming
    ' Modifications, Nigel Heffernan:
    '       ' Escape failed comparison with an empty variant in the array
    '       ' Defensive coding: check inputs
    Dim i&, j&, varMid As Variant, varX As Variant
    If IsEmpty(SortArray) Then Exit Sub
    If InStr(TypeName(SortArray), "()") < 1 Then Exit Sub  'IsArray() is somewhat broken: Look for brackets in the type name
    If lngMin = -1 Then lngMin = LBound(SortArray)
    If lngMax = -1 Then lngMax = UBound(SortArray)
    If lngMin >= lngMax Then Exit Sub   ' no sorting required
    i = lngMin: j = lngMax
    varMid = Empty: varMid = SortArray((lngMin + lngMax) \ 2)
    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a default member or property
        i = lngMax: j = lngMin
    ElseIf IsEmpty(varMid) Then i = lngMax: j = lngMin
    ElseIf IsNull(varMid) Then i = lngMax: j = lngMin
    ElseIf varMid = "" Then i = lngMax: j = lngMin
    ElseIf VarType(varMid) = vbError Then i = lngMax: j = lngMin
    ElseIf VarType(varMid) > 17 Then i = lngMax: j = lngMin
    End If
    While i <= j
        While SortArray(i) < varMid And i < lngMax: i = i + 1: Wend
        While varMid < SortArray(j) And j > lngMin: j = j - 1: Wend
        If i <= j Then varX = SortArray(i): SortArray(i) = SortArray(j): SortArray(j) = varX: i = i + 1: j = j - 1 'Swap the item
    Wend
    If (lngMin < j) Then Call QuickSort(SortArray, lngMin, j)
    If (i < lngMax) Then Call QuickSort(SortArray, i, lngMax)
End Sub

Sub QuickSortDown(A)
    Dim s$, i%
    QuickSort A
    For i = UBound(A) To LBound(A) Step -1
        s = s + vbCrLf + A(i)
    Next
    s = Mid(s, 3): A = Split(s, vbCrLf)
End Sub

Public Sub QuickSort2D(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    'Sort a 2-Dimensional array
    'SampleUsage: sort arrData by the contents of column 3
    '   QuickSort2D arrData, , , 3
    'Posted by Jim Rech 10/20/98 Excel.Programming 'Modifications, Nigel Heffernan: 'Escape failed comparison with empty variant 'Defensive coding: check inputs
    On Error Resume Next
    Dim i&, j&, varMid As Variant, arrRowTemp As Variant, lngColTemp&
    If IsEmpty(SortArray) Then Exit Sub
    If InStr(TypeName(SortArray), "()") < 1 Then Exit Sub 'IsArray() is somewhat broken: Look for brackets in the type name
    If lngMin = -1 Then lngMin = LBound(SortArray, 1)
    If lngMax = -1 Then lngMax = UBound(SortArray, 1)
    If lngMin >= lngMax Then Exit Sub   'no sorting required
    i = lngMin: j = lngMax: varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)
    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a valid default member or property
        i = lngMax: j = lngMin
    ElseIf IsEmpty(varMid) Then i = lngMax: j = lngMin
    ElseIf IsNull(varMid) Then i = lngMax: j = lngMin
    ElseIf varMid = "" Then i = lngMax: j = lngMin
    ElseIf VarType(varMid) = vbError Then i = lngMax: j = lngMin
    ElseIf VarType(varMid) > 17 Then i = lngMax: j = lngMin
    End If
    While i <= j
        While SortArray(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j, lngColumn) And j > lngMin
            j = j - 1
        Wend
        If i <= j Then
            ' Swap the rows
            ReDim arrRowTemp(LBound(SortArray, 2) To UBound(SortArray, 2))
            For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
                arrRowTemp(lngColTemp) = SortArray(i, lngColTemp)
                SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
                SortArray(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp
            i = i + 1: j = j - 1
        End If
    Wend
    If (lngMin < j) Then Call QuickSort2D(SortArray, lngMin, j, lngColumn)
    If (i < lngMax) Then Call QuickSort2D(SortArray, i, lngMax, lngColumn)
End Sub

Sub T3_AddRechteck_T3RedRect()
    'Called from    [None] Einmalig verwendet zum Erzeugen eines Rechteck-Shapes
    Dim shp As shape
    Set shp = Sheets("T3").Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 200)
    With shp: .NAME = "T3RedRect": .Line.ForeColor.RGB = RGB(255, 0, 0): .Line.Weight = 4
             .Fill.Transparency = 1: .Visible = msoTrue: '.Visible = msoFalse
    End With
End Sub

Sub T4_AddRechteck_T4RedRect()
    'Called from    [None] Einmalig verwendet zum Erzeugen eines Rechteck-Shapes
    Dim shp As shape
    Set shp = Sheets("T4").Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 200)
    With shp: .NAME = "T4RedRect": .Line.ForeColor.RGB = RGB(255, 0, 0): .Line.Weight = 4
             .Fill.Transparency = 1: .Visible = msoTrue: '.Visible = msoFalse
    End With
End Sub

Function CellIsInVisibleRange(cell$)
    Dim c As Range
    Set c = ActiveSheet.Range(cell)
    CellIsInVisibleRange = Not Intersect(ActiveWindow.VisibleRange, c) Is Nothing
End Function

Function Get_PositionOfFirstLetter(s$) As Integer
    Dim i%
    For i = 1 To Len(s)
        If Mid(UCase(s), i, 1) Like "[A-Z]" Then Get_PositionOfFirstLetter = i: Exit Function
    Next
End Function

Function Get_PositionOfLastLetter(s$) As Integer
    Dim i%
    For i = Len(s) To 1 Step -1
        If Mid(UCase(s), i, 1) Like "[A-Z]" Then Get_PositionOfLastLetter = i: Exit Function
    Next
End Function

Function Get_PositionOfFirstLetterLeftOfPositionX(s$, x%) As Integer
    Dim i%
    For i = x - 1 To 1 Step -1
        If Mid(UCase(s), i, 1) Like "[A-Zäöüß]" Then Get_PositionOfFirstLetterLeftOfPositionX = i: Exit Function
    Next
End Function

Function Get_PositionOfFirstDigit(s$) As Integer
    Dim i%
    For i = 1 To Len(s)
        If IsNumeric(Mid(s, i, 1)) Then Get_PositionOfFirstDigit = i: Exit Function
    Next
End Function

Sub TEST_StringSizeArray()
    Dim s$, v$, i%, j%, L&, r As Range, ws As Worksheet, Arr1()
    Set ws = Sheets("T4"): Set r = ws.Range(ArrC(36)) 'Range of choosen design
    v = vbCrLf: Arr1 = r.Value
    For i = 1 To UBound(Arr1, 1)
    For j = 1 To UBound(Arr1, 2)
        L = L + Len(Arr1(i, j))
        s = s + CStr(Arr1(i, j))
    Next
    Next
    show CStr(L) + v + CStr(Len(s)) + v + v + s
    '[kw1] = s
End Sub

Sub MyShortcut_Create()
    Application.OnKey "+^j", "Proc_CalledBy_MyShortCut" '+ = Ctrl '^ = Shift '% = Alt
End Sub

Sub Proc_CalledBy_MyShortCut()
    show Selection.Value + " - Jahrgang ="
End Sub

Function Delete_LastChar(s$) As String
    Delete_LastChar = Left(s, Len(s) - 1)
End Function

Function Delete_FirstAndLastChar(s$) As String
    Delete_FirstAndLastChar = Mid(Left(s, Len(s) - 1), 2)
End Function

Sub TEST_WIA_ConvertImage()
    Dim p1$, p2$, A
    p1 = "F:\Archiv TR\Archiv Trampolin\Events\1965 WM02 London\19650000 WM02 London r04 v0012-04 ww.png"
    p2 = "F:\Archiv TR\Archiv Trampolin\Events\1965 WM02 London\19650000 WM02 London r04 v0012-04 ww.jpg"
    A = WIA_ConvertImage(p1, p2, JPeG, 100)
End Sub

Public Function WIA_ConvertImage(sInitialImage As String, sOutputImage As String, _
       lFormat As wiaFormat, Optional lQuality As Long = 85) As Boolean
    '---------------------------------------------------------------------------------------
    ' Procedure : WIA_ConvertImage
    ' Author    : Daniel Pineault, CARDA Consultants Inc.
    ' Website   : http://www.cardaconsultants.com
    ' Purpose   : Convert an image's format using WIA
    ' Copyright : The following is release as Attribution-ShareAlike 4.0 International
    '             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
    ' Req'd Refs: Uses Late Binding, so none required
    '
    ' Windows Image Acquisition (WIA)
    '             https://msdn.microsoft.com/en-us/library/windows/desktop/ms630368(v=vs.85).aspx
    '
    ' Input Variables:
    ' ~~~~~~~~~~~~~~~~
    ' sInitialImage : Fully qualified path and filename of the original image to resize
    ' sOutputImage  : Fully qualified path and filename of where to save the new image
    ' lFormat       : Format to convert the image into
    ' lQuality      : Quality level to be used for the conversion process (1-100)
    '
    ' Usage:
    ' ~~~~~~
    ' Call WIA_ConvertImage("C:\Users\Public\Pictures\Sample Pictures\Chrysanthemum.png", _
    '                       "C:\Users\MyUser\Desktop\Chrysanthemum_2.jpg", _
    '                       JPEG, 50)
    '
    ' Revision History:
    ' Rev       Date(yyyy/mm/dd)        Description
    ' **************************************************************************************
    ' 1         2017-01-18              Initial Release
    ' 2         2018-09-20              Updated Copyright
    '---------------------------------------------------------------------------------------
    On Error GoTo Error_Handler
    Dim oWIA                  As Object    'WIA.ImageFile
    Dim oIP                   As Object    'ImageProcess
    Dim sFormatID$, sExt$

    'Convert our Enum over to the proper value used by WIA
    Select Case lFormat
        Case 0
            sFormatID = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
            sExt = "BMP"
        Case 1
            sFormatID = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
            sExt = "GIF"
        Case 2
            sFormatID = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
            sExt = "JPEG"
        Case 3
            sFormatID = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
            sExt = "PNG"
        Case 4
            sFormatID = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
            sExt = "TIFF"
    End Select
    
    If lQuality > 100 Then lQuality = 100

    'Should check if the output file already exists and if so,
    'prompt the user to overwrite it or not

    Set oWIA = CreateObject("WIA.ImageFile")
    Set oIP = CreateObject("WIA.ImageProcess")

    oIP.Filters.Add oIP.FilterInfos("Convert").FilterID
    oIP.Filters(1).Properties("FormatID") = sFormatID
    oIP.Filters(1).Properties("Quality") = lQuality

    oWIA.LoadFile sInitialImage
    Set oWIA = oIP.Apply(oWIA)
    'Overide the specified ext with the appropriate one for the choosen format
    oWIA.SaveFile Left(sOutputImage, InStrRev(sOutputImage, ".")) & LCase(sExt)
    WIA_ConvertImage = True

Error_Handler_Exit:
    On Error Resume Next
    If Not oIP Is Nothing Then Set oIP = Nothing
    If Not oWIA Is Nothing Then Set oWIA = Nothing
    Exit Function

Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: WIA_ConvertImage" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Function Get_ColumnNrFromCellName(CellName$) As Integer
    
End Function

Function Get_ColumnLetter(intSpalte%) As String
    Get_ColumnLetter = Right(Columns(intSpalte).Address, _
    Len(Columns(intSpalte).Address) - InStrRev(Columns(intSpalte).Address, "$"))
End Function
      
Sub Get_Cursor_Pos()
    'Example routine to retrieve cursor position
    Dim MyCursor As POINTAPI
    GetCursorPos MyCursor
    MsgBox "X Position is : " & MyCursor.xPos & Chr(10) & "Y Position is : " & MyCursor.yPos
End Sub

Sub Set_Cursor_Pos()
    'Routine to set cursor position
    'Looping routine that positions the cursor
    Dim x%, y&
    For x = 1 To 480 Step 20
       SetCursorPos x, x
       For y = 1 To 40000: Next
    Next x
End Sub

Sub TEST_TAB_Simulation()
    Dim p$, p2$, s$
    p = "F:\Archiv TR\Archiv Trampolin\Deutsche VereinsMM (BuliE) 1973-2018.txt"
    p2 = "F:\Archiv TR\Archiv Trampolin\Deutsche VereinsMM (BuliE) 1973-2018x.txt"
    's = ReadFile(p)
    s = Read_UTF8TxtFile(p)
    s = TAB_Simulation(s)
    'show s
    writeStringToFile p2, s
    Beep
End Sub

Function TAB_Simulation(s2$, Optional Sender% = 0) As String
    'Called from    xxx
    's              mehrere Zeilen, je mit derselben Anzahl "|"-Trennern; kein "§"
    'Action         alle "|" stehen in jeder Zeile an der selben Position
    
    'Vorbereitung
        'On Error Resume Next
        If s2 = "" Then Exit Function
        Dim L$, s$, Sp$, A%, Counter%, C1%, cMax%, i%, j%
        Dim Arr1() As String, Arr2() As String, Arr3() As Integer
    's  = gesamter IncomeText
        s = s2: s = Delete_EmptyTrimmedRowsInString(s)
        s = translate_UTF8_to_ANSI(s)
    'Anzahl "|" in s
        A = anzAinB("|", Left(s, InStr(1, s + vbCrLf, vbCrLf))) 'Anzahl "|" in 1. Zeile
        ReDim Arr3(0 To A - 1)
    'Arr1 = Zeilen von s
        Arr1 = Split(s, vbCrLf): s = ""
        If Sender = 1 Then MyCountDown_MainStepsAllowed 50: MyCountDown_SubStepsMax UBound(Arr1)
        For i = 0 To UBound(Arr1)
            If Sender = 1 Then MyCountDown_OneMoreSubStep
            L = Arr1(i)                             'L = one Line = x...x|x...x|...|x...x
            'Ermitteln der max. Anzahl Zeichen pro Spalte
                Arr2 = Split(L, "|")
                For j = 0 To A - 1
                    'Schleife über alle Spalten einer Zeile
                    Sp = Arr2(j)    'Sp = "x...x"
                    If Len(Sp) > Arr3(j) Then Arr3(j) = Len(Sp)
                Next
        Next
    'xxx
        If Sender = 1 Then MyCountDown_MainStepsAllowed 300: MyCountDown_SubStepsMax UBound(Arr1)
        
        
        's aufsplitten; s= s1+s2+s3+s4
        
        For i = 0 To UBound(Arr1)
            If Sender = 1 Then MyCountDown_OneMoreSubStep
            L = Arr1(i)         'one Line 'Umlaut: ggf. 2 Zeichen
            'L = translate_UTF8_to_ANSI(L)
            Arr2 = Split(L, "|")
            For j = 0 To A - 1
                'Schleife über alle Spalten einer Zeile
                Sp = Arr2(j)    'Sp = "x...x"
                s = s + Sp + String(Arr3(j) - Len(Sp), " ") + "|"
            Next
            s = s + vbCrLf
        Next
        'show CStr(Len(s)): Stop: show CStr(Len(s2)): Stop
        'show s
        s = Delete_EmptyEndRowsInString(s)
    TAB_Simulation = s
End Function

Function Get_PathsOfOpenExplorerWindows_Events_ClubsNations() As String
    Dim s$, v$, i%, Arr1() As String
    s = Get_PathsOfOpenExplorerWindows: v = vbCrLf
    If s = "" Then Exit Function
    Arr1 = Split(s, v): s = ""
    For i = 0 To UBound(Arr1)
        If Arr1(i) Like "*\Events\*" Or Arr1(i) Like "*\ClubsNations\*" Then s = s + v + Arr1(i)
    Next
    'Finals
        If s <> "" Then s = Mid(s, 3)
        Get_PathsOfOpenExplorerWindows_Events_ClubsNations = s
        'show s
End Function

Function Get_PathsOfOpenExplorerWindows_ClubsNations() As String
    Dim s$, v$, i%, Arr1() As String
    s = Get_PathsOfOpenExplorerWindows: v = vbCrLf
    If s = "" Then Exit Function
    Arr1 = Split(s, v): s = ""
    For i = 0 To UBound(Arr1)
        If Arr1(i) Like "*\ClubsNations\*" Then s = s + v + Arr1(i)
    Next
    'Finals
        If s <> "" Then s = Mid(s, 3)
        Get_PathsOfOpenExplorerWindows_ClubsNations = s
        'show s
End Function

Function Get_PathsOfOpenExplorerWindows_Events() As String
    Dim s$, v$, i%, Arr1() As String
    s = Get_PathsOfOpenExplorerWindows: v = vbCrLf
    If s = "" Then Exit Function
    Arr1 = Split(s, v): s = ""
    For i = 0 To UBound(Arr1)
        If Arr1(i) Like "*\Events\*" Then s = s + v + Arr1(i)
    Next
    'Finals
        If s <> "" Then s = Mid(s, 3)
        Get_PathsOfOpenExplorerWindows_Events = s
        'show s
End Function

Function Get_PathsOfOpenExplorerWindows2() As String
    'enable "Microsoft Internet Controls" reference
    On Error GoTo jump1
    Dim explorer As SHDocVw.shellWindows: Set explorer = New SHDocVw.shellWindows
    Dim s$, i%
    For i = 0 To 99
        s = s + explorer.item(i).LocationURL + vbCrLf
    Next
jump1:
    s = Replace(s, "file:///", ""): s = Replace(s, "%20", " "): s = Replace(s, "/", "\")
    s = Replace(s, "%F6", "ö")
    s = Replace(s, "%E4", "ä")
    s = Replace(s, "%FC", "ü")
    s = Delete_EmptyEndRowsInString(s)
    Get_PathsOfOpenExplorerWindows2 = s
End Function

Function Get_PathsOfOpenExplorerWindows() As String
    'See    Get_PathsOfOpenExplorerWindows2
    Dim s$, v$, ShellApp As Object, window As Object, windowPath$
    v = vbCrLf
    'Initialize the Shell Application
        Set ShellApp = CreateObject("Shell.Application")
    'Loop through all open windows in the shell
    For Each window In ShellApp.Windows
        'Check if the window is a File Explorer window
        '(Document name is usually "File Explorer" or "folder")
        'We also check if the Name is "Windows Explorer" (internal class name)
        If InStr(1, window.NAME, "Explorer") > 0 Or InStr(1, window.FullName, "explorer.exe") > 0 Then
            'Extract the folder path
            On Error Resume Next
            windowPath = window.Document.folder.Self.path
            If Err.Number = 0 And windowPath <> "" Then s = s + v + windowPath
            On Error GoTo 0
        End If
    Next window
    'Finals
        If s <> "" Then s = Mid(s, 3)
        Get_PathsOfOpenExplorerWindows = s
        Set ShellApp = Nothing
End Function

Function Get_NamesOfAllOpenWindows() As String
    Const GWL_STYLE = -16           'Sets a new window style
    Const WS_VISIBLE = &H10000000   'The window is initially visible
    Const GW_HWNDNEXT = 2           'The retrieved handle identifies the window
    '                                below the specified window in the Z order
    Dim hwnd As LongPtr, s$, c%
    hwnd = FindWindow(vbNullString, vbNullString)
    While hwnd
        Dim sTitle As String
        sTitle = Space$(GetWindowTextLengthA(hwnd) + 1)
        sTitle = Left$(sTitle, GetWindowText(hwnd, sTitle, Len(sTitle)))
        If (GetWindowLong(hwnd, GWL_STYLE) And WS_VISIBLE) = WS_VISIBLE Then
            'only list visible windows
            If Len(Trim(sTitle)) > 0 Then
                'ignore blank window titles
                'Debug.Print sTitle
                c = c + 1
                s = s + Format(c, "00") + "   " + sTitle + vbCrLf
            End If
        End If
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Wend
    Get_NamesOfAllOpenWindows = s
End Function

Function Get_StdMinSek_Give_secs(Sekunden As String) As String
    Dim s As Long, m As Long, H As Long, s1 As String
    If Sekunden = "" Then Sekunden = "0"
    s = CLng(Sekunden)
    m = s \ 60 'ohne VnNn
    H = m \ 60
    If H > 0 Then
        s1 = CStr(H) + " Std., " + CStr(m - H * 60) + " Min." '+ ", " + CStr(s - h * 60 * 60 - m * 60) + " Sek."
        [E2] = CStr(s) + " - " + CStr(H * 60 * 60) + " - " + CStr(m * 60)
    Else    'h = 0
        If m > 0 Then
            s1 = CStr(m) + " Min., " + CStr(s - m * 60) + " Sek."
        Else
            s1 = CStr(s) + " Sek."
        End If
    End If
    Get_StdMinSek_Give_secs = s1
End Function

Sub TEST_ShellRun()
    'zeigt alle Ordner in Volume C (nicht "$...")
    MsgBox ShellRun("cmd.exe /c dir c:\")
End Sub
Public Function ShellRun(sCmd As String) As String
    'Run a shell command, returning the output as a string
    'Aufruf-Beispiel: MsgBox ShellRun("cmd.exe /c dir c:\"); nicht: ShellRun("dir c:\")
    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    'run command
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)
    ShellRun = oExec.StdOut.ReadAll
    Set oShell = Nothing
    Set oExec = Nothing
End Function

'-------------------------------------------- 'MODUL: W3
Function cmToPoints(cM#) As Double
    Dim valueCentimeters#, valuePoints#
    valueCentimeters = cM
    valuePoints = Application.CentimetersToPoints(valueCentimeters)
    cmToPoints = valuePoints
End Function


