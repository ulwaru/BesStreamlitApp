Attribute VB_Name = "M96"
Option Explicit 'M96

'TEST_Stoppuhr
    'The hash symbol '#' represents a preprocessor command (= command that is processed prior to compilation)
    #If Win64 Then
        Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    #Else
        Public Declare Function GetTickCount Lib "kernel32" () As Long
    #End If
'getShortName
    Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'getProcessList
    Private Const INVALID_HANDLE_VALUE = -1&
    Private Const TH32CS_SNAPPROCESS = &H2
    Private Type PROCESSENTRY32
        dwSize As Long: CntUsage As Long: th32ProcessID As Long: th32DefaultHeapID As Long
        th32ModuleID As Long: CntThreads As Long: th32ParentProcessID As Long: pcPriClassBase As Long: dwFlags As Long
        szExeFile As String * 1000 'Increase this limit if there are more than 1000 Process
    End Type
    Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
    Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
    Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, PE32 As PROCESSENTRY32) As Long
    Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, PE32 As PROCESSENTRY32) As Long
'MyBeep
    Private Declare Function Beep Lib "kernel32" (ByVal Fq As Long, ByVal Tm As Long) As Long
'[END] API Functions and Variables Declaration for Task_Manager_Process_List_To_Sheet

Sub Beep1():    Beep 1000, 30:  End Sub
Sub Beep2():    Beep 200, 100:  End Sub
Sub Beep3():    Beep 2000, 20:  End Sub

Sub Screen(c%)
    If c = 0 Then Application.ScreenUpdating = False Else Application.ScreenUpdating = True
End Sub

Sub Screen0()
    Application.ScreenUpdating = False
End Sub

Sub screen1()
    Application.ScreenUpdating = True
End Sub

Function Get_Height_OfOneImage(PathOfImage$) As Integer
    Dim s
    s = Get_WidthHeight_OfOneImage(PathOfImage) '"1000 x 1234"
    s = Mid(s, InStr(1, s, " x ") + 3)
    Get_Height_OfOneImage = CInt(s)
End Function

Function Get_WidthHeight_OfOneImage(PathOfImage$) As String
    'Vorbereitung
        Dim N$, NameOfFile$, PathOfSourceFolder$, s$, s2$, Arr1() As String
        Dim FldObj As Object, fso As Object
        Dim oShell As Object, oDir As Object, sFile As Object
        Arr1 = Split(PathOfImage, "\")
        NameOfFile = Arr1(UBound(Arr1))
        PathOfSourceFolder = Replace(PathOfImage, "\" + NameOfFile, "")
        Set fso = CreateObject("scripting.filesystemobject")
        Set FldObj = fso.GetFolder(PathOfSourceFolder)
        Set oDir = CreateObject("Scripting.FileSystemObject")
        Set oShell = CreateObject("Shell.Application")
        Set oDir = oShell.Namespace(FldObj.path)
        For Each sFile In oDir.items
            If NameOfFile = oDir.GetDetailsOf(sFile, 0) Then
                s = oDir.GetDetailsOf(sFile, 31): s = Mid(s, 2, Len(s) - 2)
                Get_WidthHeight_OfOneImage = s 'Abmessungen: Width x Height
                Exit For
            End If
        Next
    'Finals
        Set oDir = Nothing: Set oShell = Nothing
End Function

Function Get_WidthHeightOfAllImages_InsideOneFolder(PathOfSourceFolder$) As String
    'Vorbereitung
        Dim FldObj As Object, fso As Object
        Dim oShell As Object, oDir As Object, sFile As Object, s$, s2$
        Set fso = CreateObject("scripting.filesystemobject")
        Set FldObj = fso.GetFolder(PathOfSourceFolder)
        Set oDir = CreateObject("Scripting.FileSystemObject")
        Set oShell = CreateObject("Shell.Application")
        Set oDir = oShell.Namespace(FldObj.path)
        For Each sFile In oDir.items
            s2 = oDir.GetDetailsOf(sFile, 31) 'Abmessungen: Width x Height
            If s2 <> "" Then
                s2 = Mid(s2, 2, Len(s2) - 2)
                s = s + oDir.GetDetailsOf(sFile, 0) + "|" + s2 + vbCrLf
            End If
        Next
    'Finals
        Set oDir = Nothing: Set oShell = Nothing
        Get_WidthHeightOfAllImages_InsideOneFolder = s
End Function

Function GetImageWidthHeight_WIA(imagePath$) As Variant
    '--------------------------------------------------------------------------------------
    'Returns an array of integers that hold the image width and height in pixels.
    'The first element of the array corresponds to the width and the second to the height.
 
    'The function uses the Microsoft Windows Image Acquisition Library v2.0, which can be
    'found in the path: C:\Windows\System32\wiaaut.dll
    'However, the code is written in late binding, so no reference is required.
 
    'Written By:    Christos Samaras
    'Date:          18/02/2018
    'E-mail:        xristos.samaras@gmail.com
    'Site:          https://myengineeringworld.net
    '--------------------------------------------------------------------------------------
    'Declaring the necessary variables.
        Dim imgSize(1)  As Integer
        Dim wia         As Object
    'Check that the image file exists.
        If FileExists(imagePath) = False Then Exit Function
    'Check that the image file corresponds to an image format.
        If IsValidImageFormat(imagePath) = False Then Exit Function
    'Create the ImageFile object and check if it exists.
        On Error Resume Next
        Set wia = CreateObject("WIA.ImageFile")
        If wia Is Nothing Then Exit Function
        On Error GoTo 0
    'Load the ImageFile object with the specified File.
        wia.LoadFile imagePath
    'Get the necessary properties.
        imgSize(0) = wia.Width
        imgSize(1) = wia.Height
    'Release the ImageFile object.
        Set wia = Nothing
    'Return the array.
        GetImageWidthHeight_WIA = imgSize
End Function

Function Get_NrOfLastColumnInRowNr(NameOfSheet$, RowNr&) As Long
    Dim ws As Worksheet
    Set ws = Sheets(NameOfSheet)
    Get_NrOfLastColumnInRowNr = ws.Cells(RowNr, ws.Columns.count).End(xlToLeft).Column
End Function




Sub SelectMyFile_WinEplorer(PathOfFile$)
    shell "explorer.exe /select,""" & PathOfFile & """", vbNormalFocus
End Sub

Sub Copy_Folder1_To_Folder2(PathOfFolder1$, PathOfFolder2$)
    'PathOfFolder1  Ordner, dessen kompletter Inhalt (Files/UnterOrdner mit Files)
    '               in PathOfFolder2 kopiert werden soll
    'Vorbereitung
        Dim f1$, f2$, Hash1$, Hash2$, p1$, p2$, s1$, s2$, v$, i%, j%
        Dim ArrFolders() As String, arrFiles() As String
        v = vbCrLf: With Sheets("XX")
    's1 = Liste aller Folder
        s1 = PathOfFolder1 + v + Get_Paths_ofAllSubfoldersAllLevelsAsStringUseGlobalVar(PathOfFolder1)
        s1 = Delete_EmptyEndRowsInString(s1)
        'show s1
    
    ArrFolders = Split(s1, v)
    For i = 0 To UBound(ArrFolders)
        'Schleife über Folder1 und alle SubFolders
        f1 = ArrFolders(i)           'one FolderPath    'F1 = "G:\Archiv Photos\abc"
            .[e9] = "Folder " + CStr(i + 1) + " of " + CStr(UBound(ArrFolders) + 1)
            .[f9] = f1
        If Right(f1, 1) <> "\" Then f1 = f1 + "\"
        f2 = Replace(f1, PathOfFolder1, PathOfFolder2)  'F2 = "N:\Archiv Photos\abc"
        Create_FoldersUpToOneFolderPath f2
        s2 = Get_AllFileNamesOfOneFolder(f1)
        If s2 <> "" Then
            arrFiles = Split(s2, v)
            For j = 0 To UBound(arrFiles)
                p1 = f1 + arrFiles(j)                   'p1 = Pfad der OriginalDatei
                p2 = f2 + arrFiles(j)                   'p2 = Pfad der Kopie
                    .[e10] = "File " + CStr(j + 1) + " of " + CStr(UBound(arrFiles) + 1)
                    .[f10] = arrFiles(j)
                If Not FileExists(p2) Then
                    CopyFile p1, p2
                    Hash1 = PS_GetFileHash(p1, "SHA512")
                    Hash2 = PS_GetFileHash(p2, "SHA512")
                    If Hash1 <> Hash2 Then
                        Stop
                    End If
                End If
            Next
        End If
    Next
    Beep
    End With
End Sub

Sub TEST_PS_GetFileHash()
    Dim p$, s$, sHashAlgorithm$
    p = "G:\Archiv Photos\- noch einordnen\von Handy Inge\2020-08\CBRI0212.jpg"
    sHashAlgorithm = "SHA512" 'MACTripleDES MD5 RIPEMD160 SHA1 SHA256 SHA384 SHA512
    s = PS_GetFileHash(p, sHashAlgorithm)
    show s
End Sub

Function PS_GetFileHash(sFile As String, sHashAlgorithm As String) As String
    '---------------------------------------------------------------------------------------
    ' Procedure : PS_GetFileHash
    ' Author    : Daniel Pineault, CARDA Consultants Inc.
    ' Website   : http://www.cardaconsultants.com
    ' Purpose   : Returns the specified Hash for the supplied file.
    '             Utilizes PowerShell to perform the Hashing.
    '             Can handle files in excess of 2GB and is faster that standard VBA version.
    ' Copyright : The following is release as Attribution-ShareAlike 4.0 International
    '             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
    ' Req'd Refs: Late Binding  -> none required
    ' Dependencies: Requires PS_GetOutput()
    '
    ' Input Variables:
    ' ~~~~~~~~~~~~~~~~
    ' sFile             : Fully qualified path and filename to get the Hash of
    ' sHashAlgorithm    : Algorithm to use for the Hashing: MACTripleDES, MD5, RIPEMD160
    '                     SHA1, SHA256, SHA384 or SHA512
    '
    ' Usage:
    ' ~~~~~~
    ' PS_GetFileHash("C:\Temp\test.xlsx", "MD5")
    '   Returns -> 82BACEED94ACB83CA9B75F69244A2978
    ' PS_GetFileHash("C:\Temp\Office2013.iso", "SHA1")
    '   Returns -> F5D743B0909024F718F65C4C72DE4AA1BE6352C3
    '
    ' Revision History:
    ' Rev       Date(yyyy-mm-dd)        Description
    ' **************************************************************************************
    ' 1         2021-10-06              Initial Public Release
    '---------------------------------------------------------------------------------------
    Dim sPSCmd                As String
    On Error GoTo Error_Handler
    Select Case sHashAlgorithm
        Case "MACTripleDES", "MD5", "RIPEMD160", "SHA1", "SHA256", "SHA384", "SHA512"
        Case Else
            MsgBox "Unknown Algorithm", vbCritical Or vbOKOnly, "Operation Aborted"
            GoTo Error_Handler_Exit
    End Select
    sPSCmd = "(Get-FileHash '" & sFile & "' -Algorithm " & sHashAlgorithm & ").Hash"
    PS_GetFileHash = PS_GetOutput(sPSCmd)

Error_Handler_Exit:
    On Error Resume Next
    Exit Function

Error_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: PS_GetFileHash" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function

Public Function PS_GetOutput(ByVal sPSCmd As String) As String
    '---------------------------------------------------------------------------------------
    ' Procedure : PS_GetOutput
    ' Author    : Daniel Pineault, CARDA Consultants Inc.
    ' Website   : http://www.cardaconsultants.com
    ' Purpose   : Run a PowerShell command and return the response
    '               Improved version where the PS prompt is never displayed to the user
    ' Copyright : The following is release as Attribution-ShareAlike 4.0 International
    '             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
    ' Req'd Refs: Late Binding  -> none required
    '
    ' Input Variables:
    ' ~~~~~~~~~~~~~~~~
    ' sPSCmd : PowerShell command to run and return the value/response of
    '
    ' Usage:
    ' ~~~~~~
    ' ? PS_GetOutput("Get-ComputerInfo -Property 'OsName'")
    '   Returns
    '       OsName
    '       ------
    '       Microsoft Windows 10 Home
    '
    ' Revision History:
    ' Rev       Date(yyyy-mm-dd)        Description
    ' **************************************************************************************
    ' 1         2020-11-05              Initial Release
    '---------------------------------------------------------------------------------------
    'Setup the powershell command properly
    sPSCmd = "powershell -command " & sPSCmd & "|clip"
    'Execute the command which is being pushed to  the clipboard
    CreateObject("WScript.Shell").Run sPSCmd, 0, True
    'Get an instance of the clipboard to capture the save value
    With CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .GetFromClipboard
        PS_GetOutput = .GetText(1)
    End With
End Function

Sub TEST_ZeitStop()
    Dim s$, t1&, T2&, Arr(), Zeit$, i&, j&
    'Array erzeugen
        ReDim Arr(1 To 999, 1 To 10)
        For i = 1 To UBound(Arr, 1)
            For j = 1 To UBound(Arr, 2)
                Arr(i, j) = "Zeile (" + CStr(i) + "|" + CStr(j) + ")"
            Next
        Next
        'showArray2D Arr
    'StartTime
        t1 = GetTickCount()
        
    'Check the running time of this macro:
        'Functions
            's = CStr(Get_AllDgTitles_FromDgs)
            's = CStr(Get_AllDgTitles_FromDgs)
            's = CStr(TEST_ZeitStop_Macro(Arr))
        'Subs
            'TEST_T4_Load_ArrDgTitles_FromT4ListOfDesignTitles
            'TEST_T4_Load_ArrDgTitles_FromDgs
            Get_CharsOfOneActivatedDesign
    'Show
        't1 = GetTickCount
        ShowTime t1
End Sub

Sub Wait_MilliSec(SomeMilliSeconds&)
    Sleep SomeMilliSeconds
End Sub

Sub ShowTime(t1_TickCount&)
    Dim t1&, T2&
    t1 = t1_TickCount: T2 = GetTickCount
    show "Zeit = " + CStr((T2 - t1) / 1000) + " sec" + vbCrLf _
        + "StartTicks: [" + CStr(t1) + "]; EndTicks: " + "[" + CStr(T2) + "]"
End Sub

Function TEST_ZeitStop_Macro(Arr()) As String
    Dim s$, A&, i&, j&, c&
    For A = 1 To 999
        For i = 1 To UBound(Arr, 1)
            For j = 1 To UBound(Arr, 2)
                s = Arr(i, j) + vbCrLf
                If s Like "*Zeile (22|6)*" Then c = c + 1
            Next
        Next
    Next
    TEST_ZeitStop_Macro = CStr(c)
End Function

Sub Test_NameMyRange()
    NameMyRange "T8", "J10", "M16", "Test1"
End Sub

Sub NameMyRange(NameOfSheet$, FromCell1$, ToCell2$, NameOfMyRange$)
    Dim rng As Range
    Set rng = Sheets(NameOfSheet).Range(FromCell1 + ":" + ToCell2)
    ThisWorkbook.names.Add NAME:=NameOfMyRange$, RefersTo:=rng
End Sub

Sub ShowAllNamedRanges()
    Dim s$, c%, N As NAME
    For Each N In Application.ActiveWorkbook.names
        c = c + 1:
        s = s + "Name=[" + N.NAME + "]; refers to [" + N.RefersTo + "]" _
        + "; z1=" + CStr(Get_z1_FromRangeAddress(N.RefersTo)) _
        + "; z2=" + CStr(Get_z2_FromRangeAddress(N.RefersTo)) _
        + "; s1=" + CStr(Get_s1_FromRangeAddress(N.RefersTo)) _
        + "; s2=" + CStr(Get_s2_FromRangeAddress(N.RefersTo)) + vbCrLf
    Next
    show "Number of named Ranges: " + CStr(c) + vbCrLf + vbCrLf + s
End Sub

Sub TEST_Delete_NameOfNamedRange()
    Delete_NameOfNamedRange "Test1"
End Sub

Sub Delete_NameOfNamedRange(NameOfNamedRange$)
    Dim N As NAME
    For Each N In Application.ActiveWorkbook.names
        If N.NAME = NameOfNamedRange Then N.Delete: Exit Sub 'N.Delete 'löscht nur den Namen, keinen Inhalt
    Next
End Sub

Sub NamedRanges_Example()
    Dim s$, rng As Range, N As NAME
    Set rng = Range("J5:L11")
    ThisWorkbook.names.Add NAME:="Test1", RefersTo:=rng
    
    For Each N In Application.ActiveWorkbook.names
        s = s + "N = [" + N + "]; N.Name = [" + N.NAME + "]; N.RefersTo = [" _
            + N.RefersTo + "] " + vbCrLf
        'N.Delete 'löscht nur den Namen
        'Spalte einfügen erweitert den RefersTo-Range
    Next
        show s
End Sub

Function Get_Cell1_FromRangeAddress(RangeAddress$) As String
    Dim s$, c%
    s = RangeAddress                                            's = "='T8'!$J$5:$L$11" '"B5"
    c = InStr(1, s, ":"): If c > 0 Then s = Left(s, c - 1)      's = "='T8'!$J$5"       '"B5"
    c = InStr(1, s, "!"): If c > 0 Then s = Mid(s, c + 1)       's = "$J$5"             '"B5"
    Get_Cell1_FromRangeAddress = s                              '"$J$5"                 '"B5"
End Function

Function Get_Cell2_FromRangeAddress(RangeAddress$) As String
    If Not RangeAddress Like "*:*" Then Exit Function
    Get_Cell2_FromRangeAddress = Mid(RangeAddress, InStr(1, RangeAddress, ":") + 1)
End Function

Function Get_z1_FromRangeAddress(RangeAddress$) As Long
    Dim s$, c%
    s = RangeAddress                                            's = "='T8'!$J$5:$L$11" '"B5"
    c = InStr(1, s, ":"): If c > 0 Then s = Left(s, c - 1)      's = "='T8'!$J$5"       '"B5"
    c = InStr(1, s, "!"): If c > 0 Then s = Mid(s, c + 1)       's = "$J$5"             '"B5"
    Get_z1_FromRangeAddress = Range(s).Row                      '10                     '2
End Function

Function Get_z2_FromRangeAddress(RangeAddress$) As Long
    If Not RangeAddress Like "*:*" Then Exit Function
    Get_z2_FromRangeAddress = Range(Mid(RangeAddress, InStr(1, RangeAddress, ":") + 1)).Row
End Function

Function Get_s1_FromRangeAddress(RangeAddress$) As Long
    Dim s$, c%
    s = RangeAddress                                            's = "='T8'!$J$5:$L$11" '"B5"
    c = InStr(1, s, ":"): If c > 0 Then s = Left(s, c - 1)      's = "='T8'!$J$5"       '"B5"
    c = InStr(1, s, "!"): If c > 0 Then s = Mid(s, c + 1)       's = "$J$5"             '"B5"
    Get_s1_FromRangeAddress = Range(s).Column                   '10                     '2
End Function

Function Get_s2_FromRangeAddress(RangeAddress$) As Long
    If Not RangeAddress Like "*:*" Then Exit Function
    Get_s2_FromRangeAddress = Range(Mid(RangeAddress, InStr(1, RangeAddress, ":") + 1)).Column
End Function

Sub T8WatchRangeNamedTest1()
    'Called from    [None]/[T8]Worksheet_Change
    
    'Vorbereitung
        Dim Adr$, r2A$, T$, v$, s1&, s2&, z1$, z2&, r2 As Range
        v = vbCrLf: EE 0
    'r1 = Old Address of named Range "Test1"
        Static r1A As String 'makes r1A persistent for this procedure even after the procedure has completed
    'r2 = New Address of named Range "Test1"
        Set r2 = ThisWorkbook.names("Test1").RefersToRange
        r2A = r2.Address 'z. B.: r2A = "='T8'!$J$5:$L$11"
    'Keine Änderung von r1
        If r1A = "" Then r1A = r2A: Exit Sub
        If r2A = r1A Then Exit Sub
    'Änderung von r1
        'Wird das NamedRange (J5:L11) verschoben, ändern sich die 2 Zellwerte (J5 und L11);
        '   das [T8]Worksheet_Change-Event wird 2 x ausgelöst ('show t' 2 x ausgeführt)
        z1 = Get_z1_FromRangeAddress(r1A)
        z2 = Get_z2_FromRangeAddress(r1A)
        s1 = Get_s1_FromRangeAddress(r1A)
        s2 = Get_s2_FromRangeAddress(r1A)

        T = "OLD:  z1 = [" + CStr(z1) + "]; " + "z2 = [" + CStr(z2) + "]; " _
                + "s1 = [" + CStr(s1) + "]; " + "s2 = [" + CStr(s2) + "]" + v

        z1 = Get_z1_FromRangeAddress(r2A)
        z2 = Get_z2_FromRangeAddress(r2A)
        s1 = Get_s1_FromRangeAddress(r2A)
        s2 = Get_s2_FromRangeAddress(r2A)
        
        T = T + "NEW:  z1 = [" + CStr(z1) + "]; " + "z2 = [" + CStr(z2) + "]; " _
                + "s1 = [" + CStr(s1) + "]; " + "s2 = [" + CStr(s2) + "]" + v
        
        
        'r1A = r2.Address
        
        
        EE 1: show T
End Sub

Sub TEST_Stoppuhr()
    Dim lngTime As Long
    lngTime = GetTickCount()
    ' Dein Code
    MsgBox "Zeit = " & (GetTickCount - lngTime) / 1000
    MsgBox "Zeit = " & CStr((GetTickCount - lngTime) / 1000) + " sec"
End Sub

Sub MessungBeispiel()
    Dim startTime As Long
    Dim endTime As Long
    Dim duration As Long
    
    ' Startzeitpunkt erfassen
    startTime = GetTickCount()
    
    ' --- Dein Code hier ---
    ' Beispiel: Eine Schleife, die etwas Zeit benötigt
    Dim i As Long
    For i = 1 To 10000000
        ' Sinnlose Rechnung zur Zeitüberbrückung
    Next i
    ' ----------------------
    
    ' Endzeitpunkt erfassen
    endTime = GetTickCount()
    
    ' Dauer berechnen
    duration = endTime - startTime
    
    MsgBox "Die Prozedur dauerte " & duration & " Millisekunden."
End Sub


'-------------------------------------------- 'MODUL: M02
Sub errExcelNeustart()
    Dim ToDoTxt As String, arrToDo() As String, PathToDoTxt As String, s As String
    Dim pathActiveExcel As String, pathVBS As String
    'errNeustart-Vermerk in ToDo.txt schreiben
        ToDoTxt = ReadToDoTxt: arrToDo = Split(ToDoTxt, vbCrLf)
        ToDoTxt = Replace(ToDoTxt, vbCrLf + "--02--", " errNeustart" + vbCrLf + "--02--", 1, 1)
        PathToDoTxt = ActiveWorkbook.path + "\Reports\ToDo.txt"
        writeStringToFile PathToDoTxt, ToDoTxt
    'vbs-Datei erzeugen:
        pathActiveExcel = ActiveWorkbook.path + "\" + ActiveWorkbook.NAME
        pathVBS = ActiveWorkbook.path + "\HilfsDateien\ExcelNeustart.vbs"
        s = "WScript.Sleep 2000" + vbCrLf + "Dim ap1" + vbCrLf
        s = s + "Set ap1 = CreateObject(" + Chr(34) + "excel.application" + Chr(34) + ")" + vbCrLf
        s = s + "ap1.Workbooks.Open (" + Chr(34) + pathActiveExcel + Chr(34) + ")" + vbCrLf
        s = s + "ap1.Visible = True" + vbCrLf
        s = s + "Set ap1 = Nothing" + vbCrLf
        writeStringToFile pathVBS, s
            'Erzeugtes vbs-script:
                'WScript.Sleep 5000
                'Dim ap1
                'Set ap1 = CreateObject("excel.application")
                'ap1.Workbooks.Open ("G:\Archiv VBA\Bes FaceTags\Bes FaceTags v24.xlsm")
                'ap1.Visible = True
                'Set ap1 = Nothing
    'LogBuch
        LogBuch "errExcelNeustart; nextLineToDo: " + CStr(arrToDo(11))
    'NeuStart-Befehl absetzen:
        Application.Wait (Now + TimeValue("0:00:01"))
        vbScriptAufrufen pathVBS
    'Excel schließen'
        ThisWorkbook.Saved = True: Application.Quit
End Sub

'-------------------------------------------- 'MODUL: M02
Sub vbScriptAufrufen(pathVBS As String)
    Dim wshShell As Object
    Set wshShell = CreateObject("WScript.Shell")
    wshShell.Run Chr(34) + pathVBS + Chr(34)
    Set wshShell = Nothing
End Sub

Function ProcKindString(ProcKind As VBIDE.vbext_ProcKind) As String
    Select Case ProcKind
        Case vbext_pk_Get
            ProcKindString = "Property Get"
        Case vbext_pk_Let
            ProcKindString = "Property Let"
        Case vbext_pk_Set
            ProcKindString = "Property Set"
        Case vbext_pk_Proc
            ProcKindString = "Sub Or Function"
        Case Else
            ProcKindString = "Unknown Type: " & CStr(ProcKind)
    End Select
End Function

Sub setReferenceForVBIDE()
    If InStr(1, getReferences, "Microsoft Visual Basic for Applications Extensibility 5.3") = 0 Then
        ThisWorkbook.VBProject.References.AddFromGuid _
            GUID:="{0002E157-0000-0000-C000-000000000046}", Major:=5, Minor:=3
    End If
End Sub

Function ComponentTypeToString(ComponentType As VBIDE.vbext_ComponentType) As String
    'wird benutzt von getListModules
    Select Case ComponentType
        Case vbext_ct_ActiveXDesigner
            ComponentTypeToString = "ActiveX Designer"
        Case vbext_ct_ClassModule
            ComponentTypeToString = "Class Module"
        Case vbext_ct_Document
            ComponentTypeToString = "Document Module"
        Case vbext_ct_MSForm
            ComponentTypeToString = "UserForm"
        Case vbext_ct_StdModule
            ComponentTypeToString = "Code Module"
        Case Else
            ComponentTypeToString = "Unknown Type: " & CStr(ComponentType)
    End Select
End Function

Sub MsgBoxAllReferences()
    Dim xRef As Variant, s As String, c As Integer
    For Each xRef In ThisWorkbook.VBProject.References
        c = c + 1
        s = s + "(" + CStr(c) + ") " + xRef.NAME + " | " + xRef.Description + vbCrLf + "     " + xRef.fullPath + vbCrLf + vbCrLf
    Next xRef
    MsgBox "Es sind " + CStr(c) + " Referenzen gesetzt:" + vbCrLf + vbCrLf + s
End Sub

Function getAnzahlStringAinStringB(A As String, B As String) As Long
    Dim s As String
    s = Replace(B, A, "")
    getAnzahlStringAinStringB = (Len(B) - Len(s)) / Len(A)
End Function

Function ExcelVersion() As String
    Dim v, strVersion, strHauptVersion, strUnterVersion As String
    strVersion = CStr(ActiveWorkbook.CalculationVersion)
    strHauptVersion = Left(strVersion, Len(strVersion) - 4)
    strUnterVersion = Mid(strVersion, Len(strVersion) - 3)
    Select Case strHauptVersion
        Case "9": v = "Excel 2000: Unterversion " & strUnterVersion
        Case "10": v = "Excel 2002: Unterversion " & strUnterVersion
        Case "11": v = "Excel 2003: Unterversion " & strUnterVersion
        Case "12": v = "Excel 2007: Unterversion " & strUnterVersion
        Case Else: v = "Excel " & strHauptVersion & ": Unterversion " & strUnterVersion
    End Select
    'v = strHauptVersion + " - " + v: ExcelVersion = v
    ExcelVersion = v
End Function

Function xxxGet_LinkTargetPath(PathOfLink$) 'ByVal PathOfLink As String)
    show PathOfLink
    On Error Resume Next
    With CreateObject("Wscript.Shell").CreateShortcut(PathOfLink)
        Get_LinkTargetPath = .targetPath
        .Close
    End With
End Function

Function Get_LinkTargetPath(PathOfLink$)
    Dim shell, lnk
    
    ' Fehlerbehandlung, falls der Pfad ungültig ist
    On Error Resume Next
    
    Set shell = CreateObject("WScript.Shell")
    ' Erstellt das Link-Objekt anhand des Pfades p
    Set lnk = shell.CreateShortcut(PathOfLink)
    
    ' Rückgabewert der Funktion setzen
    Get_LinkTargetPath = lnk.targetPath
    
    ' Aufräumen
    Set lnk = Nothing
    Set shell = Nothing
    On Error GoTo 0
End Function







