Attribute VB_Name = "M99"
' Module -- Typ: Code Module -- Name: M99

'°[START] WinApi declarations|Bes FaceTags.xlsm|MODUL: M99
Option Explicit
'[START] API Functions and Variables Declaration for getProcessList
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
'[END] API Functions and Variables Declaration for getProcessList
'©1 = COPYRIGHT © DECISION MODELS LIMITED 2016. All rights reserved; Charles Williams 27 November 2016
'[START] ©1 WinApi declarations
#If VBA7 Then
        Private Declare PtrSafe Function GlobalMemoryStatusEx Lib "kernel32.dll" (ByRef lpBuffer As MEMORYSTATUSEX) As LongPtr
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal length As LongPtr)
        ' os version info
        Declare PtrSafe Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
    Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128                      '  Maintenance string for PSS usage
    End Type
    ' dwPlatforID Constants
    Private Const VER_PLATFORM_WIN32s = 0
    Private Const VER_PLATFORM_WIN32_WINDOWS = 1
    Private Const VER_PLATFORM_WIN32_NT = 2
#Else
    Private Declare Function GlobalMemoryStatusEx Lib "kernel32.dll" (ByRef lpBuffer As MEMORYSTATUSEX) As Long
    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal length As Long)
    ' os version info
    Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
    Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128                      '  Maintenance string for PSS usage
    End Type
    ' dwPlatforID Constants
    Private Const VER_PLATFORM_WIN32s = 0
    Private Const VER_PLATFORM_WIN32_WINDOWS = 1
    Private Const VER_PLATFORM_WIN32_NT = 2
#End If
'API Structures for status of memory
Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
Private Type MEMORYSTATUSEX
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As LARGE_INTEGER
    ullAvailPhys As LARGE_INTEGER
    ullTotalPageFile As LARGE_INTEGER
    ullAvailPageFile As LARGE_INTEGER
    ullTotalVirtual As LARGE_INTEGER
    ullAvailVirtual As LARGE_INTEGER
    ullAvailExtendedVirtual As LARGE_INTEGER
End Type
'[END] ©1 WinApi declarations
#If VBA7 Then 'Office 64 bit
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else 'Office 32 bit
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If
'für WriteStringToUTF16TxtFile
    Const adModeReadWrite = 3
    Const adTypeText = 2
    Const adSaveCreateOverWrite = 2

'°[END] WinApi declarations|Bes FaceTags.xlsm|MODUL: M99

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Sub ooo___M99() '°MODUL: M99
    '°not used in another procedure
    'showProcs "modul"
    'doArr
    'CodeFind "Uf2"
    Application.EnableEvents = True
    show getReferences
End Sub
' **********************************************************************

Sub openFile(PathOfFile$)
    ShellExecute 0, "Open", PathOfFile, vbNullString, vbNullString, 1
End Sub

Sub openFileOrFolder(PathOfFile$)
    ShellExecute 0, "Open", PathOfFile, vbNullString, vbNullString, 1
End Sub

Sub OpenFolder_SelectFile(PathOfFile$)
    shell "C:\Windows\explorer.exe /select," & Chr(34) & PathOfFile + Chr(34), vbMaximizedFocus
End Sub

Sub OpenExplorerFolder_BigIcons_SelectFile(PathOfFile As String)
    Dim ShellApp As Object
    Dim TargetFolder As String
    Dim TargetFile As String
    Dim WindowWidth As Long, WindowHeight As Long
    Dim WindowLeft As Long, WindowTop As Long
    Dim ScreenWidth As Long, ScreenHeight As Long
    Dim FolderWindow As Object
    Dim View As Object

    ' 1. Pfad und Dateiname trennen
    TargetFolder = Left(PathOfFile, InStrRev(PathOfFile, "\"))
    TargetFile = Mid(PathOfFile, InStrRev(PathOfFile, "\") + 1)

    ' 2. Bildschirmauflösung ermitteln
    ScreenWidth = GetSystemMetrics(0)  ' SM_CXSCREEN
    ScreenHeight = GetSystemMetrics(1) ' SM_CYSCREEN

    ' 3. Fenstergröße berechnen (~40% der Screen-Fläche -> ca. 63% Breite & Höhe)
    WindowWidth = ScreenWidth * 0.63
    WindowHeight = ScreenHeight * 0.63

    ' 4. Fensterposition berechnen (Links Unten)
    WindowLeft = 0
    WindowTop = ScreenHeight - WindowHeight - 40 ' -40px Buffer für die Windows-Taskleiste

    ' 5. Ordner im Explorer öffnen
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Explore TargetFolder

    ' Kurz warten, damit das Fenster geladen werden kann
    Application.Wait Now + TimeValue("00:00:01")

    ' 6. Geöffnetes Fenster suchen, positionieren und Anzeigemodus anpassen
    For Each FolderWindow In ShellApp.Windows
        If InStr(1, FolderWindow.LocationURL, Replace(TargetFolder, "\", "/"), vbTextCompare) > 0 Then
            With FolderWindow
                ' Position und Größe anpassen (Links Unten)
                .Left = WindowLeft
                .Top = WindowTop
                .Width = WindowWidth
                .Height = WindowHeight

                ' Datei im Ordner markieren/auswählen
                Set View = .Document
                On Error Resume Next
                View.SelectItem View.folder.ParseName(TargetFile), 1 Or 4 Or 8 ' SelectionFlags
                
                ' Ansicht auf Große Symbole stellen (1 = Große Symbole)
                View.CurrentViewMode = 1
                On Error GoTo 0
            End With
            Exit For
        End If
    Next FolderWindow
End Sub

Sub OpenFolder(PathOfFolder$)
    If PathOfFolder = "" Then Exit Sub
    ShellExecute 0, "Open", PathOfFolder, vbNullString, vbNullString, 1
End Sub

Sub OpenNewFolder_AndHighlightFile(PathOfFile$)
    ' Prüfen, ob die Datei überhaupt existiert, um Fehler zu vermeiden
    If Dir(PathOfFile) <> "" Then
        shell "explorer.exe /select," & Chr(34) & PathOfFile & Chr(34), vbNormalFocus
    Else
        MsgBox "Die Datei wurde nicht gefunden:" & vbCrLf & PathOfFile, vbExclamation, "Fehler"
    End If
End Sub

Sub OpenFolder_IfNecessary_HighlightFile(PathOfFile$)
    'Vorbereitung
        Dim PathOfFolder$, WindowWasFound As Boolean
        Dim ShellApp As Object, SingleWindow As Object, targetFolderItem As Object
        If Dir(PathOfFile) = "" Then MsgBox "Datei existiert nicht!", vbExclamation: Exit Sub
        Set ShellApp = CreateObject("Shell.Application")
        PathOfFolder = LCase(Left(PathOfFile, InStrRev(PathOfFile, "\") - 1))
        WindowWasFound = False

    For Each SingleWindow In ShellApp.Windows
        'Nur Explorer-SingleWindow prüfen (Typ muss "IShellFolderViewDual" ermöglichen)
        If InStr(1, SingleWindow.FullName, "explorer.exe", vbTextCompare) > 0 Then
            If DecodeURL(SingleWindow.LocationURL) = PathOfFolder Then
                
                ' 1. Ansicht auf Große Symbole stellen
                ' FVM_ICON = 1, FVM_SMALLICON = 2, FVM_LIST = 3, FVM_DETAILS = 4
                '   (manchmal variiert das je nach Win-Version)
                ' Für moderne Windows-Versionen (Große Symbole):
                On Error Resume Next
                SingleWindow.Document.CurrentViewMode = 1 ' 1 steht oft für Symbole/Icons
                ' Falls das nicht hilft, erzwinge Icon-Größe:
                SingleWindow.Document.IconSize = 96 ' 96 Pixel entspricht "Groß"
                On Error GoTo 0
                
                'Datei finden und markieren
                    'Wir holen uns das FolderItem-Objekt der Datei
                    'SelectItem mag oft keine reinen Text-Strings
                    Set targetFolderItem = SingleWindow.Document.folder.ParseName(NameOfPath(PathOfFile))
                If Not targetFolderItem Is Nothing Then
                    'SingleWindow in den Vordergrund
                    SingleWindow.Visible = True
                    'Datei selektieren (1 = Auswählen, 8 = Fokus geben, 16 = Scrollen zu Objekt)
                    SingleWindow.Document.SelectItem targetFolderItem, 1 + 8 + 16
                    WindowWasFound = True: Exit For
                End If
            End If
        End If
    Next
    'Falls kein SingleWindow offen war
        If Not WindowWasFound Then shell "explorer.exe /select," & Chr(34) & PathOfFile & Chr(34), vbNormalFocus
End Sub

Function DecodeURL(ByVal url As String) As String
    'Wandelt LocationURL (file:///C:/...) in normalen Pfad um
    url = Replace(url, "file:///", ""): url = Replace(url, "/", "\"): url = Replace(url, "%20", " ")
    DecodeURL = LCase(url) 'Kleinschreibung für besseren Vergleich
End Function

Sub CloseFolder(PathOfFolder$)
    Dim shellWindows As Object, explorerWindow As Object
    Set shellWindows = CreateObject("Shell.Application").Windows
    On Error Resume Next
    For Each explorerWindow In shellWindows
        If InStr(1, explorerWindow.Document.folder.Self.path, PathOfFolder, vbTextCompare) > 0 Then
            explorerWindow.Quit
            Exit For
        End If
    Next explorerWindow
End Sub

Function getProcessListDiff(s2 As String, s3 As String) As String
    Dim s1 As String, Arr1() As String, i As Integer
    Arr1 = Split(s2, vbCrLf): s1 = s3
    For i = 0 To UBound(Arr1)
        s1 = Replace(s1, Arr1(i) + vbCrLf, " ; ")
        s1 = Replace(s1, " ;  ; ", " ; ")
    Next
    getProcessListDiff = s1 'Replace(s1, vbCrLf, ";")
End Function

Function getProcessList()
    Dim PE32 As PROCESSENTRY32, Proc_Name As String, hSnapshot As Long
    Dim iRet1 As Integer, lRet As Long, s As String, c As Integer
    'listet Prozesse wie Task-Manager/Details
    'To get details about Threads of each process, refer MSDN for additonal Parameters
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapshot <> INVALID_HANDLE_VALUE Then
        PE32.dwSize = Len(PE32)
        lRet = Process32First(hSnapshot, PE32)
        'Fetch Each Process Details one by one
        Do While lRet
            iRet1 = InStr(1, PE32.szExeFile, VBA.stringS.Chr(0))
            If iRet1 > 0 Then
                c = c + 1
                Proc_Name = VBA.stringS.Left(PE32.szExeFile, iRet1 - 1)
                s = s + "[" + CStr(Proc_Name) + "] ProcessID = " + CStr(PE32.th32ProcessID) + vbCrLf
            End If
            lRet = Process32Next(hSnapshot, PE32)
        Loop
        CloseHandle hSnapshot
    End If
    getProcessList = "Im Moment laufen " + CStr(c) + " Prozesse:" + vbCrLf + vbCrLf + s
    
    'MsgBox getProcessList
    
End Function

Sub TEST_getListModules() '°MODUL: M99
    '°not used in another procedure
    Dim s$, s2$, s3$, c%
    s = getListModules
    c = InStr(1, s, "DieseArbeitsmappe|")
    s2 = Left(s, c - 1)
    s3 = s2 + TAB_Simulation(Mid(s, c))
    s3 = Replace(s3, "|00", "|  ")
    s3 = Replace(s3, "|0", "| ")
    show s3
    'showProcs "modul"
End Sub

Sub MsgBoxListModules2()
    MsgBox getListModules2
End Sub

Function getListModules2()
    Dim VBProj As VBIDE.VBProject, VBComp As VBIDE.VBComponent
    Dim s As String, pName As String, cType As String, Cname As String
    Dim doc As Boolean, thisWB As Boolean, cM As Boolean, Blatt As Boolean
    Dim t1 As String, T2 As String, T3 As String, T4 As String, c As Integer
    
    Set VBProj = ActiveWorkbook.VBProject
    'Set WS = ActiveWorkbook.Worksheets("W4")
    'Set Rng = WS.Range("A1")
    
    For Each VBComp In VBProj.VBComponents
        pName = VBComp.Properties("Name").Value
        cType = ComponentTypeToString(VBComp.Type)
        Cname = VBComp.NAME
        c = c + 1
        If cType = "Document Module" Then doc = True Else doc = False
        If pName = ThisWorkbook.NAME Then thisWB = True Else thisWB = False
        If cType = "Code Module" Then cM = True Else cM = False
        
        If thisWB Then
            t1 = cType + " | " + Cname + vbCrLf
        ElseIf doc And Not (thisWB) Then T2 = T2 + cType _
            + " | Sheets(" + CStr(Worksheets(pName).Index) + ")" _
            + " | " + Cname + " | " + Chr(34) + pName + Chr(34) + vbCrLf
        ElseIf cM Then T3 = T3 + cType + " | " + Cname + vbCrLf
        Else
            T4 = T4 + cType + " | " + Cname + vbCrLf _
                    + " | " + Chr(34) + pName + Chr(34) + vbCrLf
        End If
    Next VBComp
    s = CStr(c) + " Modules in ActiveWorkbook:" + vbCrLf + vbCrLf
    '[a7] = Worksheets(2).Name: [a8] = Worksheets(1).CodeName
    getListModules2 = s + t1 + T2 + T3 + T4
End Function

Function getListModules() '°MODUL: M99
    '°used in ListModules,ListProcedures,MsgBoxListModules,TEST_getListModules
    Dim VBProj As VBIDE.VBProject, VBComp As VBIDE.VBComponent
    Dim AnzMod%, c%, i%, Arr1() As String, Arr2() As String
    Dim cSize$, cType$, Cname$, pName$, s$, s2$, t1$, T2$, T3$, T4$, T5$, wbName$
    Dim MODd As Boolean, MODt As Boolean, MODc As Boolean, MODu As Boolean
    
    Set VBProj = ThisWorkbook.VBProject
    s2 = getModulSizes
    AnzMod = VBProj.VBComponents.count
    ReDim Arr1(1 To AnzMod)
    'VBComponents ggf. in array + sort:
    For Each VBComp In VBProj.VBComponents
        c = c + 1
        Cname = VBComp.NAME
        cType = ComponentTypeToString(VBComp.Type)
        'pName: Userform hat keine Methode VBComp.Properties("Name")
        If cType = "UserForm" Then pName = Cname Else pName = VBComp.Properties("Name").Value
        If pName = ThisWorkbook.NAME Then wbName = Cname 'ThisWorkbook 'DieseArbeitsmappe
        cSize = CStr(Len(getContentOfOneModule(Cname)) \ 1000 + 1)
        '.... Werte --> array
        Arr1(c) = Cname + "," + Format(cSize, "000") + " KB," + cType + "," + pName
    Next
    QuickSort Arr1
    'showArray arr1
    
    For i = 1 To UBound(Arr1)
        Arr2 = Split(Arr1(i), ",")
        If Arr2(0) = wbName Then
            t1 = Arr2(0) + "|" + Arr2(1) + "|" + Arr2(2) + "||" + vbCrLf
        ElseIf Arr2(2) = "Document Module" Then T2 = T2 + Arr2(0) + "|" + Arr2(1) + "|" _
            + Arr2(2) + "|Sheets(" + CStr(Worksheets(Arr2(0)).Index) + ")" _
            + "|Sheets(" + Chr(34) + Arr2(3) + Chr(34) + ")" + vbCrLf
        ElseIf Arr2(2) = "UserForm" Then T3 = T3 + Arr2(0) + "|" + Arr2(1) + "|" + Arr2(2) + "||" + vbCrLf
        ElseIf Arr2(2) = "Code Module" Then T4 = T4 + Arr2(0) + "|" + Arr2(1) + "|" + Arr2(2) + "||" + vbCrLf
        Else
            T5 = T5 + Arr2(0) + "|" + Arr2(1) + "|" + Arr2(2) + "||" + vbCrLf
        End If
    Next
    s = CStr(AnzMod) + " Modules in ThisWorkbook:" + vbCrLf + vbCrLf
    getListModules = s + t1 + T2 + T3 + T4 + T5
Exit Function
    
    
    c = 0
    For Each VBComp In VBProj.VBComponents
        Cname = VBComp.NAME
        cType = ComponentTypeToString(VBComp.Type)
        'Userform hat keine Methode VBComp.Properties("Name")
        If cType = "UserForm" Then pName = Cname Else pName = VBComp.Properties("Name").Value
        c = c + 1                                                           '1
        MODc = False: MODd = False: MODt = False: MODu = False
        If pName = ThisWorkbook.NAME Then MODt = True
        If cType = "Document Module" Then MODd = True
        If cType = "Code Module" Then MODc = True
        If cType = "UserForm" Then MODu = True
        If MODt Then
            t1 = t1 + Cname + "|" + tmpSize(s2, Cname) _
            + "|" + cType + "||" + vbCrLf
        ElseIf MODd And Not (MODt) Then _
            T2 = T2 + Cname + "|" + tmpSize(s2, Cname) _
            + "|" + cType _
            + "|Sheets(" + CStr(Worksheets(pName).Index) + ")" _
            + "|Sheets(" + Chr(34) + pName + Chr(34) + ")" + vbCrLf
        ElseIf MODc Then _
            T3 = T3 + Cname + "|" + tmpSize(s2, Cname) _
            + "|" + cType + "||" + vbCrLf
        ElseIf MODu Then _
            T4 = T4 + Cname + "|" + "??" _
            + "|" + cType + "||" + vbCrLf
        Else
            T5 = T5 + Cname + "|" + tmpSize(s2, Cname) _
            + "|" + cType + "||" + vbCrLf
        End If
    Next VBComp
    s = CStr(c) + " Modules in ThisWorkbook:" + vbCrLf + vbCrLf
    getListModules = s + t1 + T2 + T3 + T4 + T5
End Function

Function tmpSize(s2 As String, Cname As String) As String '°MODUL: M99
    '°used in getListModules
    'liefert "31 KB " aus zeile "DieseArbeitsmappe 31483"
    Dim s As String, C2 As Integer, C3 As Integer
    C2 = InStr(InStr(1, s2, Cname), s2, " ")
    C3 = InStr(C2, s2, "Bytes")
    s = Mid(s2, C2 + 1, C3 - C2 - 2)
    tmpSize = CStr(CLng(CLng(s) / 1024)) + " KB "
End Function

Function getModulNames() '°MODUL: M99
    '°not used in another procedure
    'nur ModulNamen
    Dim VBProj As VBIDE.VBProject, VBComp As VBIDE.VBComponent, T As String, Cname As String
    Set VBProj = ActiveWorkbook.VBProject
    For Each VBComp In VBProj.VBComponents
        Cname = VBComp.NAME
        T = T + "|" + Cname + "|"
    Next VBComp
    T = Replace(T, "||", "|")
    getModulNames = T
    Set VBProj = Nothing
End Function

Function ComponentTypeToString(ComponentType As VBIDE.vbext_ComponentType) As String '°MODUL: M99
    '°used in getFunctionAndSubNames,getListModules,getFunctionAndSubNames,getModuleNameOfProc
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

Function getReferences() As String
    Dim xRef As Variant, s As String, c As Integer
    For Each xRef In ThisWorkbook.VBProject.References
        c = c + 1
        s = s + xRef.NAME + " | " + xRef.Description + vbCrLf
    Next xRef
    getReferences = "In ActiveWorkbook sind " + CStr(c) + " Referenzen gesetzt:" + vbCrLf + vbCrLf + s
End Function

Function SearchStringInOneCodeModule(NameOfCodeModule$, FindWhat$, Optional wholeWordOnly As Boolean = False) '°MODUL: M99
    '°not used in another procedure
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim s$
    Dim sL As Long ' start line
    Dim eL As Long ' end line
    Dim sc As Long ' start column
    Dim EC As Long ' end column
    Dim Found As Boolean
    
    Set VBProj = ActiveWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(NameOfCodeModule)
    Set CodeMod = VBComp.CodeModule
   
    With CodeMod
        sL = 1
        eL = .CountOfLines
        sc = 1
        EC = 255
        Found = .Find(Target:=FindWhat, StartLine:=sL, StartColumn:=sc, _
            endline:=eL, EndColumn:=EC, _
            wholeWord:=True, MatchCase:=False, patternsearch:=False)
        Do Until Found = False
            s = s + "Line " & CStr(Format(sL, "0000")) _
                  + " Column " + CStr(Format(sc, "0000")) _
                  + " in CodeModul " + NameOfCodeModule + " [" + .lines(sL, 1) + "]" + vbCrLf
            eL = .CountOfLines
            sc = EC + 1
            EC = 255
            Found = .Find(Target:=FindWhat, StartLine:=sL, StartColumn:=sc, _
                endline:=eL, EndColumn:=EC, _
                wholeWord:=wholeWordOnly, MatchCase:=False, patternsearch:=False)
        Loop
    End With
    SearchStringInOneCodeModule = s
End Function
    

Public Function CodeFind(Optional FindMod As String = "", Optional FindProc As String = "", _ '°MODUL: M99
    '°used in Worksheet_SelectionChange
    Optional FindStr As String = "", Optional TypeOfSearch As Long = 0) As Long
    '---------------------------------------------------------------------------------------
    ' Procedure : CodeFind
    ' DateTime  : 7/5/2005 18:32
    ' Author    : Nelson Hochberg
    ' Purpose   : Find a module, a procedure and/or a string in code and highlight it
    ' Returns   : 0 if not found,  line number in module if found
    ' Syntax    : lngReturn = CodeFind ([FindMod],[FindProc],[FindStr],[TypeOfSearch])
    ' Arguments : Optional FindMod As String: Part of a name of a module
    '             Optional FindProc As String: Part of a name of a procedure
    '             Optional FindStr As String: Part of a string to search
    '             NOTE: One of the above three is required
    '             Optional TypeOfSearch As Long: -1 Find line number, 0 Find string,
    '                      >0 Continue search starting at line number: TypeOfSearch + 1
    ' Thanks    : To stevbe at Experts Exchange for the initial code.
    '---------------------------------------------------------------------------------------
    Dim vbc As VBIDE.VBComponent
    Dim cM As VBIDE.CodeModule
    Dim VBAEditor As VBIDE.VBE
    Dim VBProj As VBIDE.VBProject
    Dim StartLine As Long, startcol As Long, endline As Long, endcol As Long
    
    If FindMod <> "" Then
        CodeFind = FindModule(FindMod, vbc, cM)
            If CodeFind = False Then Exit Function
        If FindProc <> "" Then
            CodeFind = FindProcedure(FindProc, StartLine, startcol, endline, endcol, cM)
                If CodeFind = False Then Exit Function
            If FindStr <> "" Then
                CodeFind = FindString(FindStr, StartLine, startcol, endline, endcol, cM, TypeOfSearch)
                    If CodeFind = False Then Exit Function
            Else
                GoTo CodeLineFound
            End If
        Else
            StartLine = 1
            If FindStr <> "" Then
                CodeFind = FindString(FindStr, StartLine, startcol, endline, endcol, cM, TypeOfSearch)
                If CodeFind = False Then Exit Function
            Else
                GoTo CodeLineFound
            End If
        End If
    Else
        Set VBAEditor = Application.VBE
        Set VBProj = VBAEditor.ActiveVBProject
        For Each vbc In VBProj.VBComponents
    
    
            Set cM = vbc.CodeModule
            If FindProc <> "" Then
                CodeFind = FindProcedure(FindProc, StartLine, startcol, endline, endcol, cM)
                If CodeFind = False Then GoTo Nextvbc2 Else Exit For
            Else
                StartLine = 1
                If FindStr <> "" Then
                    CodeFind = FindString(FindStr, StartLine, startcol, endline, endcol, cM, TypeOfSearch)
                        If CodeFind = False Then GoTo Nextvbc2 Else Exit For
                Else
                    MsgBox "CodeFind: At least one of the following is required:" & vbCrLf & _
                        "    Module" & vbCrLf & "    Procedure" & vbCrLf & "    String"
                    CodeFind = False
                    Exit Function
                End If
            End If
Nextvbc2:
        Next vbc
        If CodeFind <> False Then
            If FindStr <> "" Then
                CodeFind = FindString(FindStr, StartLine, startcol, endline, endcol, cM, TypeOfSearch)
                If CodeFind = False Then Exit Function
            Else
                GoTo CodeLineFound
            End If
        End If
    End If
    
CodeLineFound:
    If CodeFind <> False Then
        If endline = -1 Then endline = 1
        If endcol = -1 Then endcol = 1
        cM.CodePane.show
        cM.CodePane.SetSelection StartLine, startcol, endline, endcol
    End If
End Function

Private Function FindModule(FindMod As String, vbc As VBComponent, cM As CodeModule) As Long '°MODUL: M99
    '°used in CodeFind
    FindModule = False
    For Each vbc In VBE.VBProjects(1).VBComponents
        'FindMod is VBComponent
        If InStr(vbc.NAME, FindMod) > 0 Then
            Set cM = vbc.CodeModule
            FindModule = 1
            Exit For
        End If
    Next vbc
End Function

Private Function FindProcedure(FindProc As String, StartLine As Long, startcol As Long, _ '°MODUL: M99
    '°used in CodeFind
                 endline As Long, endcol As Long, cM As CodeModule) As Long
    Dim lngFake As Long
    StartLine = 1
    startcol = 1
    If FindProc <> "" Then
        'search for procedure
        FindProcedure = False
        Do
            endline = -1
            endcol = -1
            If cM.Find(FindProc, StartLine, startcol, endline, endcol) = False Then Exit Do
            If InStr(cM.ProcOfLine(StartLine, lngFake), FindProc) Then
                FindProcedure = StartLine
                Exit Do
            End If
            startcol = endcol
        Loop
    End If
End Function

Private Function FindString(FindStr As String, StartLine As Long, startcol As Long, _ '°MODUL: M99
    '°used in CodeFind
                 endline As Long, endcol As Long, cM As CodeModule, TypeOfSearch As Long) As Long
    If FindStr <> "" Then
        If TypeOfSearch > 0 Then StartLine = TypeOfSearch
        startcol = 1: endline = -1: endcol = -1
        Do
            If cM.Find(FindStr, StartLine, startcol, endline, endcol) = False Then
                FindString = False
                Exit Function
            End If
            If TypeOfSearch >= 0 Then Exit Do
            If startcol = 1 Then Exit Do
            startcol = endcol
        Loop
    End If
    FindString = endline
End Function

Function getExcelMemory() As String '©1 '°MODUL: M99
    '°not used in another procedure
    'Find used and available Excel Virtual Mmory
    Dim MemStat As MEMORYSTATUSEX
    Dim dTotalVirt As Currency, dAvailVirt As Currency, dUsedVirt As Currency, lMB As Currency
    Dim strWindows As String, XL64 As String
    Dim jXLVersion As Long, nMajorVersion As Long, nBuildNumber As Long
    
    lMB = 1048576
    
    'Windows version, build and bitness
    strWindows = " 32 bit"
    If Len(Environ("PROGRAMFILES(x86)")) <> 0 Then strWindows = " (64 bit)"
    strWindows = strWinVersion2(nMajorVersion, nBuildNumber) & " Build " & nBuildNumber & strWindows
    'Excel version, build and bitness
    jXLVersion = Val(Application.Version)
    #If Win64 Then
        XL64 = strXLVersion(jXLVersion) & " Build " & CStr(Application.Build) & " (64 bit)"
    #Else
        XL64 = strXLVersion(jXLVersion) & " Build " & CStr(Application.Build) & " (32 bit)"
    #End If
    
    'virtual memory used and maximum available
    MemStat.dwLength = Len(MemStat)
    GlobalMemoryStatusEx MemStat
    '
    dTotalVirt = LargeIntToCurrency(MemStat.ullTotalVirtual) / lMB
    dAvailVirt = LargeIntToCurrency(MemStat.ullAvailVirtual) / lMB
    dUsedVirt = Round((dTotalVirt - dAvailVirt) / 1024, 2)
    dTotalVirt = Round(dTotalVirt / 1024, 2)
    '
    getExcelMemory = strWindows & "|" & XL64 & "|" & "z. Zt. " & CStr(dUsedVirt) & " GB" & "|" & CStr(dTotalVirt) & " GB"
    'MsgBox strWindows & vbCrLf & XL64 & vbCrLf & vbCrLf & "Currently using " & CStr(dUsedVirt) & " GB of Virtual Memory" & vbCrLf & "Maximum Available is " & CStr(dTotalVirt) & " GB Virtual Memory", vbOKOnly + vbInformation, "Excel Virtual Memory Usage"
End Function

Private Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency '©1 '°MODUL: M99
    '°used in getExcelMemory
    'copy 8 bytes from the large integer to an empty currency
    CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
    'adjust it
    LargeIntToCurrency = LargeIntToCurrency * 10000
End Function

Function strWinVersion2(nMajorVersion As Long, nBuildNumber As Long) As String '©1 '°MODUL: M99
    '°used in getExcelMemory
    'Function to return the OS Version
    Dim tOSVer As OSVERSIONINFO, strSP As String
    ' First set length of OSVERSIONINFO
    ' structure size
    tOSVer.dwOSVersionInfoSize = Len(tOSVer)
    ' Get version information
    GetVersionEx tOSVer
    ' Determine OS type
    With tOSVer
        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
            ' This is an NT version (NT/2000/XP)
            If .dwMajorVersion = 5 Then
                Select Case .dwMinorVersion
                    Case 0: strWinVersion2 = "Windows 2000 "
                    Case 1: strWinVersion2 = "Windows XP "
                    Case 2: strWinVersion2 = "Windows 2003 "
                    Case Else: strWinVersion2 = "Windows NT " & .dwMajorVersion & "." & .dwMinorVersion & " "
                End Select
            ElseIf .dwMajorVersion = 6 Then
                Select Case .dwMinorVersion
                    Case 0: strWinVersion2 = "Windows Vista "
                    Case 1: strWinVersion2 = "Windows 7 "
                    Case 2: strWinVersion2 = "Windows 8 "
                    Case Else: strWinVersion2 = "Windows 10 "
                End Select
            ElseIf .dwMajorVersion = 10 Then strWinVersion2 = "Windows 10 "
            Else
                strWinVersion2 = "Windows 10 "
            End If
        Else
            ' This is Windows 95/98/ME
            If .dwMajorVersion >= 5 Then
                strWinVersion2 = "Windows ME "
            ElseIf .dwMajorVersion = 4 And .dwMinorVersion > 0 Then
                strWinVersion2 = "Windows 98 "
            Else
                strWinVersion2 = "Windows 95 "
            End If
        End If
        nMajorVersion = .dwMajorVersion
        nBuildNumber = .dwBuildNumber
        'strSP = .szCSDVersion
        If Len(strSP) > 0 Then strWinVersion2 = strWinVersion2 & strSP
    End With
GoExit:
End Function

Function strXLVersion(jXLVersion As Long) As String '©1 '°MODUL: M99
    '°used in getExcelMemory
    'convert the Excel version number to a string
    Select Case jXLVersion
        Case 8: strXLVersion = "Excel 97"
        Case 9: strXLVersion = "Excel 2000"
        Case 10: strXLVersion = "Excel 2002"
        Case 11: strXLVersion = "Excel 2003"
        Case 12: strXLVersion = "Excel 2007"
        Case 14: strXLVersion = "Excel 2010"
        Case 15: strXLVersion = "Excel 2013"
        Case 16: strXLVersion = "Excel 2016"
        Case Else: strXLVersion = "Excel 20??"
    End Select
End Function

Function getModulSizes() '°MODUL: M99
    '°used in getListModules
    getModulSizes = getModulesContent(True)
End Function

Function getModulesContent(Optional SizesOnly As Boolean = False) As String '°MODUL: M99
    '°used in getModulSizes
  'gets ThisWorkbook's whole VBA project string
  'Set reference to Microsoft VBA Extensibility 5.5
  Dim VBProj As VBIDE.VBProject, VBComp As VBIDE.VBComponent, nLines As Long
  Dim VBMod As VBIDE.CodeModule, sMod As String, s As String, s2 As String
  s2 = "[NEW " + "MODULE:] " 'delimiter
  'get ref to ThisWorkbook project
  Set VBProj = ThisWorkbook.VBProject
  'loop through VBComponents collection
  For Each VBComp In VBProj.VBComponents
    Set VBMod = VBComp.CodeModule
    nLines = VBMod.CountOfLines
      If nLines <> 0 Then
        sMod = VBMod.lines(1, nLines) '1. bis letzte Zeile eines Moduls
        '1 Modul ist jetzt eingelesen
        If SizesOnly = False Then s = s + vbCrLf + vbCrLf + s2
        s = s + VBComp.NAME + " " + CStr(Len(sMod)) + " Bytes" & vbCrLf
        If SizesOnly = False Then s = s & vbCrLf & sMod
      Else 'just accum name of empty component
        s = s & vbCrLf & s2 + VBComp.NAME + " " + CStr(Len(sMod)) & vbCrLf & vbCrLf
      End If
  Next VBComp
  Do While InStr(1, s, vbCrLf + vbCrLf + vbCrLf) > 0
    s = Replace(s, vbCrLf + vbCrLf + vbCrLf, vbCrLf + vbCrLf)
  Loop
  getModulesContent = s
  Set VBProj = Nothing: Set VBComp = Nothing: Set VBMod = Nothing
End Function
'Sub TEST_()
'    show getContentOfOneModule("W1")
'End Sub

Function getContentOfOneModule(NameOfModule$) As String '°MODUL: M99
    '°used in getListModules
    'Set reference to Microsoft VBA Extensibility 5.5
    Dim VBProj As VBIDE.VBProject, VBComp As VBIDE.VBComponent, nLines As Long
    Dim VBMod As VBIDE.CodeModule, sMod As String, s As String
    'get ref to ThisWorkbook project
        Set VBProj = ThisWorkbook.VBProject
    'loop through VBComponents collection
        For Each VBComp In VBProj.VBComponents
            If VBComp.NAME = NameOfModule Then
                Set VBMod = VBComp.CodeModule
                nLines = VBMod.CountOfLines
                If nLines <> 0 Then
                    sMod = VBMod.lines(1, nLines) '1. bis letzte Zeile eines Moduls
                    '1 Modul ist jetzt eingelesen
                    s = "'Complete content of module " + VBComp.NAME + ":" + vbCrLf & sMod
                End If
                Exit For
            End If
        Next VBComp
    getContentOfOneModule = s
    Set VBProj = Nothing: Set VBComp = Nothing: Set VBMod = Nothing
End Function

Function getModulesContentPepUp() '°MODUL: M99
    '°used in show_CodeToTextFile,ShowCompleteCode
    Dim VBComp As Object, i As Long, startProc As Boolean, startModul As Boolean
    Dim s As String, s2 As String, s1 As String, s3 As String, s4 As String, ListUsedIn As String, ListVarNotUsed As String
    
    ListUsedIn = getListUsedIn
    ListVarNotUsed = getListVarNotUsed
    
    s = "'°Gesamter VBA-Code (Makros/Prozeduren aller Module) der Excel-Datei <" _
        & ThisWorkbook.NAME & ">" & vbCrLf
    
    s = s + "'°   Der Text in einer Zeile nach den beiden Zeichen '°" & vbCrLf
    s = s + "'°   ist ein dem Rohtext der Modulinhalte hinzugefügter Kommentar." & vbCrLf
    s = s & "'°****************************************************************************************" & vbCrLf & vbCrLf
    s = s + tmpListOfVarsNotUsed(ListVarNotUsed)
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        s = s + "'°##############################" & vbCrLf _
              + "'°MODUL: " & VBComp.NAME & vbCrLf _
              + "'°##############################" & vbCrLf & vbCrLf
        With VBComp.CodeModule
            i = 1: startModul = True
            s1 = ThisWorkbook.NAME & "|" & "MODUL: " & VBComp.NAME + vbCrLf
            s2 = vbCrLf + "    '°" + s1
            s = s + "'°[START] WinApi declarations|" + s1
            Do Until i > .CountOfLines
                If Left(.lines(i, 1), 4) = "Sub " Then startProc = True
                If Left(.lines(i, 1), 9) = "Function " Then startProc = True
                If Left(.lines(i, 1), 12) = "Private Sub " Then startProc = True
                If Left(.lines(i, 1), 17) = "Private Function " Then startProc = True
                If Left(.lines(i, 1), 11) = "Public Sub " Then startProc = True
                If Left(.lines(i, 1), 16) = "Public Function " Then startProc = True
                If startProc Then
                    If startModul Then s = s + "'°[END] WinApi declarations|" & ThisWorkbook.NAME & "|" & "MODUL: " & VBComp.NAME + vbCrLf
                    'Jetzt wird <'°MODUL: xx> hinter die Sub-Kopfzeile geschrieben:
                     s = s + vbCrLf + .lines(i, 1) + " '°MODUL: " & VBComp.NAME
                    'Jetzt wird <'°used in ...> unter die Sub-Kopfzeile geschrieben:
                     s4 = getUsedInLine(getProcNameOfProcLine(.lines(i, 1)), ListUsedIn) '.Lines(i, 1) = ProcHeader
                     s = s + vbCrLf + "    '°" + s4 + vbCrLf
                    'Jetzt wird <'°declared variables not used: ...> unter manche Sub-Kopfzeile geschrieben:
                     s3 = getVarNotUsedLine(getProcNameOfProcLine(.lines(i, 1)), ListVarNotUsed)
                     If s3 <> "" Then s = s + "    '°" + s3 + vbCrLf
                    startProc = False: startModul = False
                Else
                    'die VnNnlichen Zeilen bis zum Beginn der nächsten Prozedur:
                    If .lines(i, 1) <> "" Then s = s + .lines(i, 1) + vbCrLf
                End If
                i = i + 1
            Loop
        End With
        If startModul Then s = s + "'°[END] WinApi declarations|" & ThisWorkbook.NAME & "|" & "MODUL: " & VBComp.NAME + vbCrLf
        s = s + vbCrLf
    Next VBComp
    'Do While InStr(1, s, vbCrLf + vbCrLf + vbCrLf) > 0
        's = Replace(s, vbCrLf + vbCrLf + vbCrLf, vbCrLf + vbCrLf)
    'Loop
    getModulesContentPepUp = s
End Function

Function tmpListOfVarsNotUsed(ListVarNotUsed As String) As String '°MODUL: M99
    '°used in getModulesContentPepUp
    Dim s As String, s2 As String
    If ListVarNotUsed = "" Then Exit Function
    s2 = ListVarNotUsed
    s2 = Replace(s2, "declared", "- declared")
    s2 = Replace(s2, vbCrLf, vbCrLf + "'° ")
    s = "'° Procedures, that declare variables and do not use them:" + vbCrLf + "'° "
    s = s + s2 + vbCrLf
    s = s + "'°****************************************************************************************" & vbCrLf & vbCrLf
    tmpListOfVarsNotUsed = s
End Function

Function getVarNotUsedLine(ProcName As String, ListVarNotUsed As String) As String '°MODUL: M99
    '°used in getModulesContentPepUp
    Dim s As String, c As Long
    'liefert 1 VarNotUsed-Zeile für 1 proc
    s = ListVarNotUsed
    c = InStr(1, s, vbCrLf + ProcName + " ")
    If c = 0 Then Exit Function
    s = Mid(s, c + 2) 'Text vor ProcName wird abgeschnitten
    s = Replace(s, ProcName + " ", "") 'ProcName wird abgeschnitten
    c = InStr(1, s, vbCrLf)
    s = Left(s, c - 1) 's trägt jetzt die VarNotUsed-Zeile
    getVarNotUsedLine = s
End Function

Function getUsedInLine(ProcName As String, ListUsedIn As String) As String '°MODUL: M99
    '°used in getModulesContentPepUp
    Dim s As String, c As Long
    'liefert 1 usedIn-Zeile für 1 proc
    s = ListUsedIn
    c = InStr(1, s, vbCrLf + ProcName + " ") + 2
    s = Mid(s, c) 'Text vor ProcName wird abgeschnitten
    s = Replace(s, ProcName + " ", "") 'ProcName wird abgeschnitten
    c = InStr(1, s, vbCrLf)
    s = Left(s, c - 1) 's trägt jetzt die usedIn-Zeile
    getUsedInLine = s
End Function

Function getListVarNotUsed() '°MODUL: M99
    '°used in getModulesContentPepUp
    '°declared variables not used: Sub
    Dim s$, s1$, s2$, s3$, s4$, ProcName$, v$, TR$, procVars$
    Dim Arr1() As String, Arr2() As String, Arr3() As String
    Dim i%, j%, C1%
    
    s1 = vbCrLf + GetContentOfAllProcs 'enthält den Trenner tr
    TR = "'°" + "#'" 'Trenner/Delimiter
    Arr1 = Split(s1, vbCrLf + TR)
    
    For i = 0 To UBound(Arr1)                   'Schleife über alle WorkbookProcs
        s4 = Arr1(i)                            's4 = ein kompletter ProzedurTxt
        If Len(s4) > 4 Then
            ProcName = getProcNameOfProcLine(s4)    'ProcName = ein ProcName
            s4 = getProCnoComments(s4)          'Kommentare aus Proc entfernen
            procVars = getListVarsOfOneProc(s4)
            Arr2 = Split(procVars, vbCrLf)
            For j = 0 To UBound(Arr2)           'Schleife über alle Variablen der proc
                v = Arr2(j)                     'v = ein Variablenname
                'Reparaturen
                    v = Replace(v, "(", ""): v = Replace(v, ")", "")
                    v = Replace(v, "ByRef", ""): v = Replace(v, "ByVal", "")
                    v = Replace(v, "$", ""): v = Replace(v, "%", ""): v = Replace(v, "&", "")
                    v = Replace(v, "!", ""): v = Replace(v, Chr(34), ""):
                    v = Replace(v, "Sub ", ""): v = Replace(v, "Function ", "")
                    If Left(v, 1) = "." Then v = ""
                    If IsNumeric(Left(v, 1)) Then v = ""
                If v <> "" Then
                    C1 = InStr(1, s4, v)
                    s3 = Mid(s4, C1 + 1)        's3 = procTxt ohne erstes v (in Dim ...)
                    If Not WholeWordIsInStringS(v, s3, False) Then
                        'ganzes Wort V ist nicht in s3 enthalten
                    
                    
                        If InStr(1, s3, "(" + v) = 0 Then
                            If InStr(1, v, "(") = 0 Then s = s + v + ","
                        End If
                        
                        
                    End If
                End If
            Next
'            'Reparaturen
'                s = Replace(s, ",,", ","): s = Replace(s, ")", ""): s = Replace(s, "ByVal", "")
'                s = Replace(s, "$", ""): s = Replace(s, "!", ""): s = Replace(s, "%", "")
'                s = Replace(s, "&", ""): s = Replace(s, Chr(34), "")
'                s = Replace(s, "Sub", ""): s = Replace(s, "Function", "")
'                If s = "," Then s = ""
'                If Left(s, 1) = "," Then s = Mid(s, 2)
'                If IsNumeric(Left(s, 1)) Then s = ""
                If Right(s, 1) = "," Then s = Left(s, Len(s) - 1)
            If s <> "" Then
                s = ProcName + " declared variables not used: " + s
                's2 = s2 + Left(s, Len(s) - 1) + vbCrLf
                s2 = s2 + s + vbCrLf
                s = ""
            End If
        End If
    Next
    'sort s2
        Arr3 = Split(s2, vbCrLf): s2 = ""
        QuickSort Arr3
        For i = 0 To UBound(Arr3)
            s2 = s2 + Arr3(i) + vbCrLf
        Next
        show s2
    getListVarNotUsed = s2
End Function

Function getListVarsOfOneProc(completeProcTxt As String) As String '°MODUL: M99
    '°used in getListVarNotUsed
    '°declared variables not used: Sub,lefts1,InStr1
    Dim s As String, s1 As String, s2 As String
    Dim Arr1() As String, Arr2() As String, i As Integer, C1 As Integer
    s = Replace(completeProcTxt, ":", vbCrLf)
    s = Replace(s, "ReDim", "")
    Arr1 = Split(s, vbCrLf)
    For i = 0 To UBound(Arr1)           'Schleife über alle ProcZeilen
        s1 = Arr1(i)                    's1 = one ProcLine
        If InStr(1, s1, "Sub ") > 0 Or InStr(1, s1, "Function ") > 0 Then
            s1 = Replace(s1, "Optional ", "")
            s1 = Replace(s1, Left(s1, InStr(1, s1, "(")), "Dim ")
        End If
        If InStr(1, s1, "Dim ") > 0 Then
            s2 = s2 + s1 + ","          's2 = String aus proc-"Dim "-Zeilen
        End If
    Next
    s2 = Replace(s2, "Dim ", ",")
    s2 = Replace(s2, "()", "")
    Do While InStr(1, s2, " ,") > 0 Or InStr(1, s2, ", ") > 0
        s2 = Replace(s2, ", ", ","): s2 = Replace(s2, " ,", ",")
        's2 = ",FSO,oFolder,oSubfolder,queue As Collection,countFolders As Long,,s$,pFo As String,i As Long,"
    Loop
    
    Arr2 = Split(s2, ",")
    s1 = "": s2 = ""
    For i = 0 To UBound(Arr2)           'Schleife über alle Variablen einer proc
        s1 = Arr2(i)                    's1 = oneVariable '"FSO", "queue As Collection", ...
        If s1 <> "" Then
            C1 = InStr(2, s1, " ")
            If C1 > 0 Then s1 = Left(s1, C1 - 1)
            If Right(s1, 1) = "%" Or Right(s1, 1) = "&" Or Right(s1, 1) = "@" _
                Or Right(s1, 1) = "!" Or Right(s1, 1) = "#" Or Right(s1, 1) = "$" Then _
                s1 = Left(s1, Len(s1) - 1)
            s2 = s2 + s1 + vbCrLf
        End If
    Next
    getListVarsOfOneProc = s2
End Function

Function getProCnoComments(completeProcTxt As String) As String '°MODUL: M99
    '°used in getListVarNotUsed,GetSubFunCnames_insideOneProc
    Dim s As String, s1 As String, s2 As String
    Dim Arr1() As String, i As Integer, C1 As Integer
    s = Replace(completeProcTxt, " _" + vbCrLf, "")  'SonderZeilenumbruch entfernen
    Arr1 = Split(s, vbCrLf)
    For i = 0 To UBound(Arr1)           'Schleife über alle ProcZeilen
        s1 = Arr1(i)                    's1 = one ProcLine
        If s1 <> "" Then
            C1 = 0
            C1 = InStr(1, s1, "'")
            If C1 = 0 Then
                s2 = s2 + s1 + vbCrLf   's2 = Proczeilen ohne Comments
            Else
                'behandelt ' in " ' " fälschlich als CommentStart
                s2 = s2 + Left(s1, C1 - 1) + vbCrLf
            End If
        End If
    Next
    getProCnoComments = s2
End Function

Function Struc_GetStructureOfOneProc(NameOfParameter$) As String '°MODUL: M99
    '°not used in another procedure
    Application.EnableEvents = True
    Fb_Tick1 "Struc_GetStructureOfOneProc", "W1", "Y33", NameOfParameter
    Dim L$, NA$, NameOfProc$, NoSubs$, Nr$, sF$, T$
    Dim i%, j%, W1H3S1%, W1H3Z1%, Arr1() As String, Arr2() As String
    'NameOfParameter ist ggf. nur Teilstring
        W1H3S1 = CInt(ArrC(25)) 'SpaltenNr W1H3 LiOb (H3=HellgrüneFläche Nr 3)
        W1H3Z1 = CInt(ArrC(24)) 'ZeilenNr  W1H3 LiOb
        NameOfProc = GetOneProcName_ContainingMyString_LeftMost(NameOfParameter)
        ThisWorkbook.Sheets("W1").Cells(W1H3Z1 + 31, W1H3S1 + 11) = NameOfProc
    'Start Aufbau Struktur-Liste:
        T = "01 " + NameOfProc + "*"                    'T = Text; Struktur; Liste ProcNames
        
    'Untersuche alle NameOfProc mit "*", ob sie UP (UnterProcs) aufrufen;
    ' gefundene UP erhalten "*" (werden ihrerseits untersucht)
    Do While InStr(1, T, "*") > 0
        T = Delete_EmptyRowsInString(T)
        Arr1 = Split(T, vbCrLf)
        T = ""                                      'T    wird neu aufgebaut
        For i = 0 To UBound(Arr1)
            L = Arr1(i)                             'L  = "01 actualProcName*"; Line
            If Right(L, 1) = "*" Then
                Nr = Left(L, InStr(1, L, " ") - 1)  'Nr = "01"
                NA = Mid(L, InStr(1, L, " ") + 1)
                NA = Left(NA, Len(NA) - 1)          'Na  = "actualProcName", Name
                If InStr(1, T, NA + vbCrLf) = 0 And InStr(1, T, NA + "*") = 0 Then
                    T = T + Nr + " " + NA + vbCrLf      'L ohne *
                    sF = GetSubFunCnames_insideOneProc(NA)
                    If sF = "" Then NoSubs = NoSubs + "," + NA + ","
                    sF = Delete_EmptyRowsInString(sF)
                    Arr2 = Split(sF, vbCrLf)
                    For j = 0 To UBound(Arr2)
                        If InStr(1, T, Arr2(j) + vbCrLf) = 0 And InStr(1, T, Arr2(j) + "*") = 0 Then
                            T = T + Nr + Format(j + 1, "00") + " " + Arr2(j) + "*" + vbCrLf
                        Else
                            T = T + Nr + Format(j + 1, "00") + " " + Arr2(j) + vbCrLf
                        End If
                    Next
                    DoEvents
                Else
                    T = T + Nr + " " + NA + vbCrLf
                End If
            Else
                T = T + L + vbCrLf     'aktuelle Zeile L wird in Liste T übernommen
            End If
            
            DoEvents
            
        Next
        'DoEvents
    Loop
    Struc_GetStructureOfOneProc = Struc_PepUp(T, NoSubs)
    Fb_Tick2
End Function

Function Struc_PepUp(ListOfNumbersProcNames$, NoSubs$) As String '°MODUL: M99
    '°used in Struc_GetStructureOfOneProc
    Dim L$, MainProc$, NameOfProc$, NewL$, s$, s1$, SammN$, T$, u$
    Dim c%, C2%, i%, MaxL%, Arr1() As String
    s = ListOfNumbersProcNames
    s = Delete_EmptyRowsInString(s)
    Arr1 = Split(s, vbCrLf)
    'Zeilen wie L = "010302 W3H5_Fill_arrC_fromW3" sortieren:
        QuickSort Arr1
    'MaxL ermitteln:
        SammN = ","
        For i = 0 To UBound(Arr1)
            L = Arr1(i)
            c = InStr(1, L, " ")
            NameOfProc = Mid(L, c + 1)
            s1 = s1 + NameOfProc + vbCrLf 'Sammlung aller ProcNames
            If MainProc = "" Then MainProc = NameOfProc
            NewL = String(c - 3, " ") + NameOfProc
            If Len(NewL) > MaxL Then MaxL = Len(NewL)
            T = T + NewL + vbCrLf
        Next
        T = Replace(T, "  ", ". ")
        T = Delete_EmptyRowsInString(T)
        'T-Zeilen sind sortiert; Procs tieferen Levels sind eingerückt (". " pro Level)
    'Die "+" und "s. oben"-Markierung setzen (MaxL bekannt):
        Arr1 = Split(T, vbCrLf)
        T = "": C2 = 3
        For i = 0 To UBound(Arr1)
            L = Arr1(i)
            c = 2 * anzAinB(".", L) + 1
            If c > C2 Then u = "+" Else u = "-"         'u = "+": Unterprozedur von Proc(i-1)
            NameOfProc = Mid(L, c)
            If InStr(1, SammN, "," + NameOfProc + ",") = 0 _
                Or InStr(1, NoSubs, "," + NameOfProc + ",") > 0 Then
                SammN = SammN + NameOfProc + ","
                T = T + Format(i, "000") + u + "  " + L + vbCrLf
            Else
                T = T + Format(i, "000") + u + "  " + L + String(MaxL - Len(L) + 6, " ") + "s. oben" + vbCrLf
            End If
            C2 = c
        Next
    'Wiederholungen von UnterProcs löschen
        T = Struc_DeleteDoubles(T)
    'Vortext der Liste
        s = "Struktur der Prozedur '" + MainProc + "'" + vbCrLf + vbCrLf _
          + "     Aufrufe von " + CStr(anzAinB(",", SammN) - 2) + " verschiedenen Unter-/UnterUnterProzeduren;" + vbCrLf _
          + "     kam eine Prozedur (Sub oder Function) bereits weiter oben in der Liste vor," + vbCrLf _
          + "     werden ihre UnterProzeduren nicht erneut aufgeführt;" + vbCrLf _
          + "     stattdessen findet sich ein Hinweis auf das Erstvorkommen dieser Prozedur" + vbCrLf + vbCrLf
      
    Struc_PepUp = s + T
End Function

Function Struc_DeleteDoubles(Struc_Text$) As String '°MODUL: M99
    '°used in Struc_PepUp
    '°declared variables not used: s9
    'Vorbereitung:
        Dim G$, L1$, L2$, s$, s2$, s3$, s8$, s9$, T$
        Dim i%, j%, StartG1%, StartG2%, c&
        Dim Arr1() As String, Arr2() As String, Arr3() As String
        s = Struc_Text     'Zeilen wie "007+  . . . . W3H5_Fill_arrC_GiveNrVal"
    's reduzieren zu "W3H5_Fill_arrC_GiveNrVal" (arr1, arr2) und "+" (arr3)
        s8 = Replace(s, "s. oben", ""): s8 = Replace(s8, ". ", ""): s8 = Replace(s8, " ", "")
        's8 = Mid(s8, InStr(1, s8, "000")) 'Vortext weg
        'show s8
        Arr1 = Split(s8, vbCrLf): Arr2 = Split(s8, vbCrLf): Arr3 = Split(s8, vbCrLf)
        For i = 0 To UBound(Arr1)
            Arr3(i) = Mid(Arr1(i), 4, 1)        '-|+|...
            Arr1(i) = Mid(Arr1(i), 5)           'W3_Fill_arrC|Get_FirstID_ofSomeIDs|...
            Arr2(i) = Arr1(i)                   'W3_Fill_arrC|Get_FirstID_ofSomeIDs|...
        Next
    'WiederholungsStarts ermitteln
        For i = 0 To UBound(Arr1) - 1
            L1 = Arr1(i) + "|" + Arr1(i + 1)
            For j = 1 To UBound(Arr2) - 1
                L2 = Arr2(j) + "|" + Arr2(j + 1)
                If L1 = L2 And i < j And Arr3(i + 1) = "+" And Arr3(j + 1) = "+" Then
                    s2 = s2 _
                    + "i " + Format(i, "000") + ", " _
                    + "j " + Format(j, "000") + ", " + L1 + vbCrLf 'And arr3(i + 1) = "+"
                End If
            Next
        Next
        s2 = Delete_EmptyRowsInString(s2)
        '  = Liste; Zeilen wie "i 009, j 018, Fill_arrC_fromW3|LOG"
        '  = in s enthaltene Wiederholung einer Gruppe von Zeilen;
        '    Start bei 009 und 018; Zeilenanzahl der Gruppe noch unbekannt;
        '    die Zeilen             009+  . . . . . . W3H5_Fill_arrC_fromW3
        '                           010+  . . . . . . . LOG
        '                           011-  . . . . . . . GetCellNameOfGreenLiOb
        '                           012-  . . . . . . . Get_NrOfLastRowInColumnNr
        '    wiederholen sich in    018+  . . . W3H5_Fill_arrC_fromW3                                      s. oben
        '                           019+  . . . . LOG
        '                           020-  . . . . GetCellNameOfGreenLiOb
        '                           021-  . . . . Get_NrOfLastRowInColumnNr
        '    Hier wird Zeile 018 den Hinweis 'siehe 009' tragen,
        '    die Zeilen 019-021 werden gelöscht (sind Doubles)
    'Nummern zu löschender Wdh-Zeilen ermitteln:
        Arr2 = Split(s2, vbCrLf)
        For i = 0 To UBound(Arr2)
            StartG1 = CInt(Mid(Arr2(i), 3, 3))  '9
            StartG2 = CInt(Mid(Arr2(i), 10, 3)) '18
            For j = 1 To 99
                If Arr1(StartG1 + j) = Arr1(StartG2 + j) Then
                    If InStr(1, G, CStr(StartG2 + j)) = 0 Then G = G + Format(StartG2 + j, "000") + ","
                Else
                    j = 99
                End If
            Next
        Next
     'Wdh-Zeilen löschen:
        s = Delete_EmptyRowsInString(s)
        Arr1 = Split(s, vbCrLf)
        j = 0
        For i = 0 To UBound(Arr1)
            L1 = Arr1(i)
            If InStr(1, G, Left(L1, 3)) = 0 Then
                'Zeile L1 wird nicht gelöscht
                If InStr(1, L1, "s. oben") > 0 Then
                    'L2 (ProcName) ermitteln:
                        L2 = Trim(Replace(L1, "s. oben", ""))
                        L2 = Mid(L2, 7): L2 = Replace(L2, ". ", "") 'W3H5_Fill_arrC_GiveNrVal
                    's3 (Nr der Zeile mit L2) in T (StrukturText) suchen:
                        c = InStr(1, T, " " + L2 + vbCrLf)
                        s3 = Left(T, c): s3 = Replace(s3, ".", "")
                        s3 = Replace(s3, " ", ""): s3 = Right(s3, 3) '"007"
                    'Nr in L1 schreiben
                        L1 = Replace(L1, "s. oben", "s. " + s3)
                End If
                T = T + Format(j, "000") + Mid(L1, 5) + vbCrLf
                j = j + 1
            End If
        Next
    Struc_DeleteDoubles = T
End Function

Sub TEST_getListUsedIn() '°MODUL: M99
    '°not used in another procedure
    show getListUsedIn
End Sub

Function getListUsedIn() '°MODUL: M99
    '°used in getModulesContentPepUp,TEST_getListUsedIn
    Dim s As String, s1 As String, s2 As String, OneProcName As String, OneProcContent As String, s5 As String
    Dim Arr1() As String, Arr2() As String, TR As String
    Dim i As Integer, j As Integer
    
    s1 = getFunctionAndSubNames(False) 'False = ohne ModulName
    TR = "'°" + "#'" 'Trenner/Delimiter
    s2 = vbCrLf + GetContentOfAllProcs
    '  = alle Prozeduren; jede beginnt mit tr (als Delimiter für Split)
    Arr1 = Split(s1, vbCrLf)
    Arr2 = Split(s2, vbCrLf + TR)
    For i = 0 To UBound(Arr1)
        'Schleife über alle ProcNames
        OneProcName = Arr1(i)                   'OneProcName = ein ProzedurName
        If OneProcName <> "" Then
            s = s + vbCrLf + OneProcName + " used in "
            For j = 0 To UBound(Arr2)
                'Schleife über alle Procs
                OneProcContent = Arr2(j)        'OneProcContent = eine komplette Prozedur
                If OneProcContent <> "" Then
                    s5 = getProcNameOfProcLine(OneProcContent)
                    If WholeWordIsInStringS(OneProcName, OneProcContent, False) _
                       And OneProcName <> s5 Then
                        'OneProcName ist im Rumpf (nicht im Titel) von OneProcContent enthalten
                        'OneProcName wird also von proc s5 aufgerufen
                        s = s + s5 + ","
                    End If
                End If
            Next
        End If
    Next
    s = s + vbCrLf
    s = Replace(s, "," + vbCrLf, vbCrLf)
    s = Replace(s, " used in " + vbCrLf, " not used in another procedure" + vbCrLf)
    getListUsedIn = s
End Function

Function WordIsInCommentZone(wholeWord As String, stringS As String) As Boolean '°MODUL: M99
    '°used in WholeWordIsInStringS
    Dim w As String, s As String, z As String, Arr1() As String, i As Long
    w = wholeWord: s = stringS
    s = Replace(s, " _" + vbCrLf, "") 'Kommentar könnte per " _" mehrzeilig sein
    Arr1 = Split(s, vbCrLf)
    For i = 0 To UBound(Arr1)
        z = Arr1(i)
        If InStr(1, z, w) > 0 Then Exit For 'w ist in Zeile i von s
    Next
    'der Fall <z = "'" + w> ist nicht berücksichtigt (' vor w, aber nicht in CommentZone)
    
    If InStr(1, z, "'") = 0 Then Exit Function
    If InStr(1, z, w) > InStr(1, z, "'") Then WordIsInCommentZone = True
End Function

Function WholeWordIsInStringS(wholeWord As String, stringS As String, Optional includeCommentZone As Boolean = True) As Boolean '°MODUL: M99
    '°used in getListVarNotUsed,getListUsedIn,GetSubFunCnames_insideOneProc
    Dim w As String, s As String, c As Integer
    w = wholeWord: s = " " + stringS
    If w = "" Then Exit Function
    If InStr(1, s, w) = 0 Then Exit Function
    Do While InStr(1, s, w) > 0
        c = InStr(1, s, w) 'old in Folder --> c = 2
        If Mid(s, c - 1, 1) Like "[0-9a-zA-ZäöüßÄÖÜ_]" Or Mid(s, c + Len(w), 1) Like "[0-9a-zA-ZäöüßÄÖÜ_]" Then
            WholeWordIsInStringS = False 'das Zeichen vor oder nach w ist Buchstabe/Zahl
            s = Replace(s, w, "", 1, 1) 'erstes Vorkommen von w löschen
        Else 'es existiert w als Wort - ggf. in der CommentZone
            If includeCommentZone Then
                WholeWordIsInStringS = True
            Else
                If Not WordIsInCommentZone(w, s) Then WholeWordIsInStringS = True
            End If
            Exit Function
        End If
    Loop
End Function

Function getProcNameOfProcLine(proc As String) As String '°MODUL: M99
    '°used in getModulesContentPepUp,getListVarNotUsed,getListUsedIn
    '°declared variables not used: Sub
    Dim s As String, s1 As String
    s = proc
    If Len(s) < 5 Then Exit Function
    '"Sub ","Function ","Private Sub ","Private Function ","Public Sub ","Public Function "
    s1 = Left(s, InStr(1, s, "(") - 1) 'Sub ...ProcName
    s1 = Replace(s1, "Private ", ""): s1 = Replace(s1, "Public ", "")
    s1 = Replace(s1, "Sub ", ""): s1 = Replace(s1, "Function ", "")
    getProcNameOfProcLine = s1
End Function

Public Function GetListProcNames_Modul_LineNr() As String '°MODUL: M99
    '°used in GetContentOfOneProc,GetSubFunCnames_insideOneProc
    Dim Component As Object, NAME$, s$, Kind&, Index&
    For Each Component In Application.VBE.ActiveVBProject.VBComponents
        With Component.CodeModule
            'The Procedures
            Index = .CountOfDeclarationLines + 1
            Do While Index < .CountOfLines
                NAME = .ProcOfLine(Index, Kind)
                s = s + Component.NAME & "." & NAME + "." + CStr(Index)
                Index = .ProcStartLine(NAME, Kind) + .ProcCountLines(NAME, Kind) + 1
                s = s + "." + CStr(Index - 1) + vbCrLf
            Loop
        End With
    Next Component
    GetListProcNames_Modul_LineNr = s
End Function

Function Get_FirstLine_LikeMyString(SomeLines$, MyString$) As String
    Dim i&, Arr1() As String
    Arr1 = Split(SomeLines, vbCrLf)
    For i = 0 To UBound(Arr1)
        If Arr1(i) Like "*" + MyString + "*" Then Exit For
    Next
    If i < UBound(Arr1) + 1 Then Get_FirstLine_LikeMyString = Arr1(i)
End Function

Function Get_FirstLine_ContainingMyString(SomeLines$, MyString$) As String '°MODUL: M99
    '°used in GetContentOfOneProc,GetSubFunCnames_insideOneProc
    Dim s$, C1&, C2&, C3&
    s = vbCrLf + SomeLines + vbCrLf
    C3 = InStr(1, s, MyString)
    If C3 > 0 Then
        C1 = InStrRev(s, vbCrLf, C3) + 2    'Zeile first char
        C2 = InStr(C3, s, vbCrLf) - 1       'Zeile last  char
        Get_FirstLine_ContainingMyString = Mid(s, C1, C2 - C1 + 1)
    End If
End Function

Public Function GetContentOfOneProc(NameOfProc$, Optional NameOfModul$ = "-") As String '°MODUL: M99
    '°used in GetSubFunCnames_insideOneProc
    Dim L$, ListP$, i&, nr1&, nr2&, Arr1() As String
    ListP = GetListProcNames_Modul_LineNr
    If NameOfModul = "-" Then
        L = Get_FirstLine_ContainingMyString(ListP, "." + NameOfProc + ".")
    Else
        L = Get_FirstLine_ContainingMyString(ListP, NameOfModul + "." + NameOfProc + ".")
    End If
    If L <> "" Then
        Arr1 = Split(L, ".")
            NameOfModul = Arr1(0) 'NameOfProc = arr1(1)
            nr1 = CInt(Arr1(2)) '= Nr of ModulLine - ProcStart
            nr2 = CInt(Arr1(3)) '= Nr of ModulLine - ProcEnd
        With ThisWorkbook.VBProject.VBComponents(NameOfModul).CodeModule
            For i = nr1 To nr2
                GetContentOfOneProc = GetContentOfOneProc + .lines(i, 1) + vbCrLf
            Next
        End With
    End If
End Function

Function GetSubFunCnames_insideOneProc(NameOfProc$, Optional NameOfModul$ = "-") As String '°MODUL: M99
    '°used in Struc_GetStructureOfOneProc
        Dim ListP$, pos$, s$, G$, w$, W2$, c%, C1%, i%, Arr1() As String
    'Proc NameOfProc laden, säubern:
        s = GetContentOfOneProc(NameOfProc, NameOfModul)
        'Lösche 1. Zeile (Sub ..., Function ...)
            's = Mid(s, InStr(1, s, vbCrLf) + 2)
        s = getProCnoComments(s)
        s = DeleteQuotesContent(s)
        s = Delete_EmptyRowsInString(s)
    'Liste aller ProcNames
        ListP = GetListProcNames_Modul_LineNr
        ListP = Delete_EmptyRowsInString(ListP)
    Arr1 = Split(ListP, vbCrLf)
    For i = 0 To UBound(Arr1)
        'Schleife über alle existierenden ProcNames
        w = Arr1(i): w = Mid(w, InStr(1, w, ".") + 1)
        w = Left(w, InStr(1, w, ".") - 1)   'W = Wort, 1 ProcName der Liste
        If WholeWordIsInStringS(w, s) Then
            'Es existiert ein UnterProc - als Wort W, also mit W+" " | W+"(" | W+VbCrLf
            C1 = 1
            Do
                c = InStr(C1, s, w)
                'InStr findet ggf. Teilstring (W3H2_Fill_arrWB in W3H2_Fill_arrWB_6to7)
                W2 = Mid(s, c - 1, Len(w) + 2) 'Wort2 = W + 1 Zeichen li, + 1 Zeichen rechts
                If WholeWordIsInStringS(w, W2) Then
                    'W ist tasächlich ein ganzes Wort
                    pos = Format(c, "00000") + " "
                    G = G + pos + w + vbCrLf
                    Exit Do
                Else
                    'W ist kein ganzes Wort, nur Teilstring
                    'W befindet sich also weiter rechts in s
                    C1 = c + 1
                End If
            Loop
        End If
    Next
    G = Replace(G, Get_FirstLine_ContainingMyString(G, NameOfProc) + vbCrLf, "")
    G = Delete_EmptyRowsInString(G)
    'G enthält jetzt (unsortiert) alle Sub-/Function-Aufrufe innerhalb der proc; mit Pos
    If G <> "" Then
        Arr1 = Split(G, vbCrLf): G = ""
        QuickSort Arr1
        For i = 0 To UBound(Arr1)
            G = G + Mid(Arr1(i), 7) + vbCrLf
        Next
    End If
    GetSubFunCnames_insideOneProc = G
End Function

Function DeleteQuotesContent(SomeCodeLines$) As String '°MODUL: M99
    '°used in GetSubFunCnames_insideOneProc
    Dim L$, s$, s1$, s2$, qq$, c%, C1%, C2%, i%, Arr1() As String
    qq = Chr(34) 'quote, "-Zeichen
    s = SomeCodeLines
    Arr1 = Split(s, vbCrLf)
    For i = 0 To UBound(Arr1)
        L = Arr1(i)
        c = anzAinB(qq, L)
        Do While anzAinB(qq, L) > 1
            C1 = InStr(1, L, qq)
            C2 = InStr(C1 + 1, L, qq)
            s2 = Mid(L, C1, C2 - C1 + 1)
            L = Replace(L, s2, " ")
        Loop
        s1 = s1 + L + vbCrLf
    Next
    DeleteQuotesContent = s1
End Function

Function GetContentOfAllProcs() As String '°MODUL: M99
    '°used in getListVarNotUsed,getListUsedIn
    'liefert alle Prozeduren; jede beginnt mit "'°" + "#'" (als Delimiter für Split)
    Dim VBComp As Object, i As Long
    Dim s As String, startProc As Boolean, startModul As Boolean, whileProc As Boolean
    'doArrC
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        With VBComp.CodeModule
            'VBComp.CodeModule = z. B. "DieseArbeitsmappe"
            i = 1: startModul = True: whileProc = False
            Do Until i > .CountOfLines
                'Schleife über alle Zeilen eines Moduls
                If Left(.lines(i, 1), 4) = "Sub " Then startProc = True
                If Left(.lines(i, 1), 9) = "Function " Then startProc = True
                If Left(.lines(i, 1), 12) = "Private Sub " Then startProc = True
                If Left(.lines(i, 1), 17) = "Private Function " Then startProc = True
                If Left(.lines(i, 1), 11) = "Public Sub " Then startProc = True
                If Left(.lines(i, 1), 16) = "Public Function " Then startProc = True
                If startProc Then
                    'Zeile mit nächstem Prozedurstart ist gefunden
                    s = s + "'°" + "#'"
                    startProc = False: startModul = False: whileProc = True
                End If
                'die VnNnlichen Zeilen bis zum Beginn der nächsten Prozedur:
                If whileProc And Trim(.lines(i, 1)) <> "" Then s = s + .lines(i, 1) + vbCrLf
                i = i + 1
            Loop
        End With
        s = s + vbCrLf
    Next VBComp
    GetContentOfAllProcs = s
End Function

Function Convert_UmlauteToCharCodes(s$) As String
    '
    'Ä = Chr(196), Ö = Chr(214), Ü = Chr(220), ß = Chr(223)
    'ä = Chr(228), ö = Chr(246), ü = Chr(252)
    '195 Ã, 188 ¼
    '
    s = Replace(s, Chr(195) + Chr(188), Chr(252)) 'ü
    '... usw.
    
End Function

Function showStringAsCharCodes(s1$) As String
    Dim oneChar$, oneCode$, s$, i%
    s = s1 + vbCrLf + vbCrLf
    For i = 1 To Len(s1)
        oneChar = Mid(s1, i, 1)
        oneCode = Asc(Mid(s1, i, 1))
        s = s + Format(i, "000") + vbTab + oneChar + vbTab + CStr(oneCode) + vbCrLf
    Next
    show s
End Function

Sub Info_ASCII_ANSI_UNICODE_UTF8_Umlaute()
    'ASCII
    '
    '7 Bit, 128 Zeichen;
    'Zeichen 32-127: darstellbare Zeichen (vor allem die in der englischen Sprache
    '                verwendeten Buchstaben).
    '
    'ANSI
    '
    '8 Bit, 256 Zeichen;
    'Zeichen 0-127: ASCII-Zeichensatz; 128-255: Sonderzeichen (je nach ISO-Zeichensatz);
    'ISO-8859-1 schriftspezifischen Zeichen für westeuropäische und amerikanische Sprachen
    'ISO-8859-2 Zeichen für die meisten mitteleuropäischen und slawischen Sprachen
    '
    'Unicode
    '
    'soll Lösung für universellen Zeichensatz zur Darstellung möglichst aller Zeichen
    'aller Sprachen sein.
    '16 Bit, 65536 Zeichen; falls erstes Byte auf null gesetzt wird,
    'verhält Unicode sich wie ASCII.
    '
    'UTF -8
    '
    'In UTF-8 (8-bit Unicode Transformation Format) wird jedem Unicode-Zeichen
    'eine speziell kodierte Bytekette von variabler Länge zugeordnet.
    'UTF-8 unterstützt bis zu 4 Byte, mit denen sich 1.114.112 Zeichen abbilden lassen.
    'Unicode-Zeichen mit den Werten aus dem Bereich von 0 bis 127
    'werden in der UTF-8-Kodierung als ein Byte mit dem gleichen Wert wiedergegeben.
    'Insofern sind alle Daten, die ausschließlich echte ASCII-Zeichen verwenden,
    'in beiden Darstellungen identisch.
    'Unicode-Zeichen größer als 127 werden in der UTF-8-Kodierung zu Byteketten
    'der Länge zwei bis vier.
End Sub

Function translate_UTF8_to_ANSI(text As String) As String
    'UTF-8 was developed to create a more or less equivalent to ANSI
    'but without the many disadvantages ANSI had.
    'Both UTF-8 and ANSI expand from the basic set of characters put forth by ASCII;
    'so the two are basically equivalent when it comes to the first 127 characters.
    'ANSI is pretty limited.
    
    'The Open function from VBA works on ANSI encoded files only and binary;
    ' readFile uses "Open myFilePath"
    'If you wish to read/write an utf-8 file, you'll have to find another way;
    ' use Read_UTF8TxtFile, Write_UTF8TxtFile
    'A String in Excel and VBA is stored as utf-16
    '(VBA editor still use ANSI), so you only need to convert from utf-8 to utf-16.

    text = Replace(text, "Ã ", "à")     ' a  grave small
    text = Replace(text, "Ã„", "Ä")     ' a  umlaut big
    text = Replace(text, "Ã–", "Ö")     ' o  umlaut big
    text = Replace(text, "Ã¤", "ä")     ' a  umlaut small
    text = Replace(text, "Ã‡", "Ç")     ' c  cedilla big
    text = Replace(text, "Ã§", "ç")     ' c  cedilla small
    text = Replace(text, "Ã‰", "É")     ' e  acute big
    text = Replace(text, "Ã©", "é")     ' e  acute small
    text = Replace(text, "Ãª", "ê")     ' e  circumflex small
    text = Replace(text, "Ãˆ", "È")      ' e  grave big
    text = Replace(text, "Ã¨", "è")     ' e  grave small
    text = Replace(text, "Ã˜", "Ø")      ' o  slash big
    text = Replace(text, "Ã¸", "ø")     ' o  slash small
    text = Replace(text, "Ã¶", "ö")     ' o  umlaut small
    text = Replace(text, "ÃŸ", "ß")      ' sz ligature small
    text = Replace(text, "Ãœ", "Ü")      ' u  umlaut big
    text = Replace(text, "Ã¼", "ü")     ' u  umlaut small
    text = Replace(text, "Ã½", "ý")     ' y  acute small
    text = Replace(text, "â€“", "–")    ' Gedankenstrich
    text = Replace(text, "â€™", "’")    ' U+2019
    text = Replace(text, "â€š", "‚")      'U+201A
    text = Replace(text, "â€˜ ", "‘")     'U+2018
    text = Replace(text, "â€¦", "…")    ' U+2026
    translate_UTF8_to_ANSI = text
End Function

Function translate_ANSI_to_UTF8(text As String) As String
    'some chars only
    text = Replace(text, "à", "Ã ")    ' a  grave small
    text = Replace(text, "Ä", "Ã„")     ' a  umlaut big
    text = Replace(text, "Ö", "Ã–")     ' o  umlaut big
    text = Replace(text, "ä", "Ã¤")    ' a  umlaut small
    text = Replace(text, "Ç", "Ã‡")     ' c  cedilla big
    text = Replace(text, "ç", "Ã§")     ' c  cedilla small
    text = Replace(text, "É", "Ã‰")     ' e  acute big
    text = Replace(text, "é", "Ã©")     ' e  acute small
    text = Replace(text, "ê", "Ãª")     ' e  circumflex small
    text = Replace(text, "ö", "Ã¶")     ' o  umlaut small
    text = Replace(text, "ß", "ÃŸ")      ' sz ligature small
    text = Replace(text, "Ü", "Ãœ")      ' u  umlaut big
    text = Replace(text, "ü", "Ã¼")     ' u  umlaut small
    text = Replace(text, "–", "â€“")    ' Gedankenstrich
    text = Replace(text, "’", "â€™")    ' U+2019
    text = Replace(text, "‚", "â€š")      'U+201A
    text = Replace(text, "‘", "â€˜ ")     'U+2018
    text = Replace(text, "…", "â€¦")    ' U+2026
    translate_ANSI_to_UTF8 = text
End Function

Function translate_ANSI_to_UTF8hex(s As String) As String
    s = Replace(s, "Ä", "%C3%84", , , vbBinaryCompare)
    s = Replace(s, "Ö", "%C3%96", , , vbBinaryCompare)
    s = Replace(s, "Ü", "%C3%9c", , , vbBinaryCompare)
    s = Replace(s, "ä", "%C3%A4", , , vbBinaryCompare)
    s = Replace(s, "ö", "%C3%B6", , , vbBinaryCompare)
    s = Replace(s, "ü", "%C3%BC", , , vbBinaryCompare)
    s = Replace(s, "ß", "%DF", , , vbBinaryCompare)
    translate_ANSI_to_UTF8hex = s
End Function

Public Function Read_UTF8TxtFile(path As String, Optional CharSet As String = "utf-8")
  'A String in Excel and VBA is stored as utf-16
  Static obj As Object
  If obj Is Nothing Then Set obj = VBA.CreateObject("ADODB.Stream")
  obj.CharSet = CharSet
  obj.Open
  obj.LoadFromFile path
  Read_UTF8TxtFile = obj.ReadText()
  obj.Close
End Function

Public Sub Write_UTF8TxtFile(path As String, text As String, Optional CharSet As String = "utf-8")
  'A String in Excel and VBA is stored as utf-16
  Static obj As Object
  If obj Is Nothing Then Set obj = VBA.CreateObject("ADODB.Stream")
  obj.CharSet = CharSet
  obj.Open
  obj.WriteText text
  obj.SaveToFile path
  obj.Close
End Sub

Sub WriteStringToUTF16TxtFile(MyString, PathOfFile)
    'Action     schreibt TextFile als UTF-16 LE-BOM
    With CreateObject("ADODB.Stream")
        .Mode = adModeReadWrite                         'Const adModeReadWrite = 3
        .Type = adTypeText                              'Const adTypeText = 2
        .CharSet = "UTF-16": .Open: .WriteText MyString
        .SaveToFile PathOfFile, adSaveCreateOverWrite   'Const adSaveCreateOverWrite = 2
        .Close
    End With
End Sub

Sub WriteStringToAnsiTxtFile(pathForFileToWrite As String, StringToWrite As String)
    writeStringToFile pathForFileToWrite, StringToWrite
End Sub

Sub writeStringToFile(pathForFileToWrite$, StringToWrite$)
    'schreibt Ansi-TextFile
    'pathForFileToWrite = "C:\Users\ubes\Desktop\OUT\tmp.txt" 'kurzfristig
    Dim i As Long: i = FreeFile 'change Output to Append if you want to add to an existing file rather than creating a new file each time
    If pathForFileToWrite Like "*Roaming*" Then RoamingHinweis
    Open pathForFileToWrite For Output As i: Print #i, StringToWrite: Close i
End Sub

Sub WriteStringToFileAppend(pathForFileToWrite As String, StringToWrite As String)
    Dim i As Long: i = FreeFile 'change Output to Append if you want to add to an existing file rather than creating a new file each time
    If pathForFileToWrite Like "*Roaming*" Then RoamingHinweis
    Open pathForFileToWrite For Append As i: Print #i, StringToWrite: Close i
End Sub

Sub RoamingHinweis()
    MsgBox "Roaming-Hiweis" + vbCrLf _
    + "Diese Excel-Datei ist eine automatisch wiederhergestellte Datei." + vbCrLf _
    + "Ggf. möchten Sie die Original-Version dieser Datei starten."
    Stop
End Sub

Sub TEST_IsUTF16File()
    show IsUTF16File(ArrC(6)) 'Read me.txt
End Sub

Public Function IsUTF16File(filePath)
        Dim objFSO As Object, objStream As Object, intAsc1Chr%, intAsc2Chr%
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        If (objFSO.FileExists(filePath) = False) Then
            IsUTF16File = False
            Exit Function
        End If

        ' 1=Read-only, False==do not create if not exist, -1=Unicode 0=ASCII
        Set objStream = objFSO.OpenTextFile(filePath, 1, False, 0)
        intAsc1Chr = Asc(objStream.Read(1))
        intAsc2Chr = Asc(objStream.Read(1))
        objStream.Close

        If (intAsc1Chr = 255) And (intAsc2Chr = 254) Then
            IsUTF16File = True
        Else
            IsUTF16File = False
        End If

        Set objStream = Nothing
        Set objFSO = Nothing
End Function

Public Function ReadSmallFile(ByVal filePath As String) As String
     'Beliebige Datei auslesen und Inhalt als String zurückgeben
     Dim F As Integer, sInhalt As String
     If Dir$(filePath, vbNormal) <> "" Then 'Prüfen, ob Datei existiert
        F = FreeFile: Open filePath For Binary As #F ' Datei im Binärmodus öffnen
        sInhalt = Space$(LOF(F)) 'Größe ermitteln und Variable entsprechend mit Leerzeichen füllen
        Get #F, , sInhalt 'Gesamten Inhalt in einem "Rutsch" einlesen
        Close #F 'Datei schliessen
     End If
     ReadSmallFile = sInhalt
End Function

Function ReadFile3(Pfad$) As String '°MODUL: M99
    Dim s As String, i As Integer: i = FreeFile
    'readFile erkennt keine Pfade mit Umlauten etc., deshalb:
        Pfad = translate_UTF8_to_ANSI(Pfad)
    'Show Pfad
    Open Pfad For Input As #i: s = Input(LOF(i), i): Close #i
    ReadFile = s
    'Beep
End Function

Function ReadFile2(Pfad$) As String
    Dim fileNum%, DataLine$, s$
    fileNum = FreeFile()
    Open Pfad For Input As #fileNum
    While Not EOF(fileNum)
        Line Input #fileNum, DataLine ' read in data 1 line at a time
        s = s + DataLine + vbCrLf
    Wend
    ReadFile2 = s
End Function

Function ReadFile(Pfad$) As String
    Dim s As String
    Close #1
    Open Pfad For Input As #1
    Do Until EOF(1)
       Line Input #1, s 'Adding Line to read the whole line, not only first 128 positions
        ReadFile = ReadFile + s + vbCrLf
    Loop
    Close #1
End Function

Function GetOneProcName_ContainingMyString_LeftMost(MyString) As String '°MODUL: M99
    '°used in Struc_GetStructureOfOneProc
    Dim L$, m$, N$, T$, C1%, C2%, i%, Arr1() As String
    T = getFunctionAndSubNames(False)
    m = LCase(MyString)
    C2 = 99
    Arr1 = Split(T, vbCrLf)
    For i = 0 To UBound(Arr1)
        L = Arr1(i)     'one Line
        C1 = InStr(1, LCase(L), m)
        If C1 > 0 Then
            If C1 < C2 Then
                C2 = C1
                N = L
                If C2 = 1 Then Exit For
            End If
        End If
    Next
    GetOneProcName_ContainingMyString_LeftMost = N
End Function

Sub TEST_Show_AllProcNames_ContentContainingMyString_Sorted() '°MODUL: M99
    '°not used in another procedure
    Show_AllProcNames_ContentContainingMyString_Sorted "arrD[02]"
End Sub

Sub Show_AllProcNames_ContentContainingMyString_Sorted(MyString$) '°MODUL: M99
    '°used in TEST_Show_AllProcNames_ContentContainingMyString_Sorted
    Dim s$, s2$, i%, Arr1() As String
    s = GetAllProcNames_ContentContainingMyString_Sorted(MyString)
    Arr1 = Split(s, vbCrLf)
    For i = 0 To UBound(Arr1)
        If InStr(1, Arr1(i), "ContentContainingMyString_Sorted") = 0 Then
            If InStr(1, Arr1(i), "QuickSteps") = 0 Then
                If Len(Arr1(i)) > 0 Then s2 = s2 + Arr1(i) + vbCrLf
            End If
        End If
    Next
    s = "ProcNames of procs containing String '" + MyString + "':" + vbCrLf + vbCrLf + s2
    show s
End Sub

Function GetAllProcNames_ContentContainingMyString_Sorted(MyString$) '°MODUL: M99
    '°used in Show_AllProcNames_ContentContainingMyString_Sorted
    Dim A$, L$, LCont$, LL$, s2$, s3$, s1$, i%, Arr1() As String
    s1 = getFunctionAndSubNames(False)
    Arr1 = Split(s1, vbCrLf)
    For i = 0 To UBound(Arr1)
        L = Arr1(i)
        'LCont = LCase(GetContentOfOneProc(L))
        LCont = GetContentOfOneProc(L)
        LL = LCase(L)
        'If InStr(1, LCont, LCase(MyString)) > 0 Then
        If InStr(1, LCont, MyString) > 0 Then
            s2 = s2 + LL + "," + L + vbCrLf
        End If
    Next
    If s2 <> "" Then
        Arr1 = Split(s2, vbCrLf)
        QuickSort Arr1
        For i = 0 To UBound(Arr1)
            Arr1(i) = Mid(Arr1(i), InStr(1, Arr1(i), ",") + 1)
            s3 = s3 + Arr1(i) + vbCrLf
        Next
        
        GetAllProcNames_ContentContainingMyString_Sorted = s3
    End If
End Function

Function GetAllProcNames_ContainingMyString_Sorted(MyString$)
    'Vorbereitung
        Dim A$, L$, LL$, m$, M1$, M2$, s1$, s2$, s3$, c%, i%, Arr1() As String
    'All Subs
        s1 = getFunctionAndSubNames()
    '1 or 2 words: M or M1, M2
        m = Trim(MyString)                              'get color
            c = InStr(1, m, " ")
            If c > 0 Then
                M1 = Left(m, c - 1)                     'get
                M2 = Mid(m, c + 1)                      'color
            End If
    'Arr1 = Array of all Subs
        Arr1 = Split(s1, vbCrLf)
    's2 = FoundSubs
        For i = 0 To UBound(Arr1)
            L = Arr1(i)                                 'L  = Get_Color
            LL = LCase(L)                               'LL = get_color
            If c > 0 Then
                If InStr(1, LL, LCase(M1)) > 0 And InStr(1, LL, LCase(M2)) > 0 Then
                    s2 = s2 + LL + "," + L + vbCrLf
                End If
            Else
                If InStr(1, LL, LCase(m)) > 0 Then s2 = s2 + LL + "," + L + vbCrLf
            End If
        Next
        s2 = Delete_EmptyRowsInString(s2)
    If s2 <> "" Then
        Arr1 = Split(s2, vbCrLf)
        QuickSort Arr1
        For i = 0 To UBound(Arr1)
            Arr1(i) = Mid(Arr1(i), InStr(1, Arr1(i), ",") + 1)
            s3 = s3 + Arr1(i) + vbCrLf
        Next
        s3 = Delete_EmptyRowsInString(s3)
        GetAllProcNames_ContainingMyString_Sorted = s3
    End If
End Function

Public Function getFunctionAndSubNames(Optional withModulName As Boolean = False) As String '°MODUL: M99
    '°used in ListProcedures,GetAllProcNames_ContainingMyString_Sorted,getListUsedIn,GetOneProcName_ContainingMyString_LeftMost,GetAllProcNames_ContentContainingMyString_Sorted,GetAllProcNames_ContainingMyString_Sorted,showProcs,ShowProcNamesInsideProcNames
    'Verweis: Microsoft Visual Basic for Applications Extensibility 5.3 library
    Dim item As Variant, s As String
    For Each item In ThisWorkbook.VBProject.VBComponents
        If ComponentTypeToString(vbext_ct_StdModule) = "Code Module" Then
            s = s + getProceduresOfOneModule(item.NAME, withModulName) + vbCrLf
        End If
    Next item
    getFunctionAndSubNames = s
End Function

Public Function getModuleNameOfProc(myProcName As String) As String '°MODUL: M99
    '°not used in another procedure
    'Verweis: Microsoft Visual Basic for Applications Extensibility 5.3 library
    Dim item As Variant, s As String
    For Each item In ThisWorkbook.VBProject.VBComponents
        If ComponentTypeToString(vbext_ct_StdModule) = "Code Module" Then
            
            s = s + getProcNameAifInModuleB(item.NAME, myProcName) + vbCrLf 'item.Name = ein ModulName
            s = Trim(Replace(s, vbCrLf, ""))
            If s <> "" Then
                getModuleNameOfProc = s: Exit Function
            End If
        End If
    Next item
    getModuleNameOfProc = s
End Function

Function getProcNameAifInModuleB(ModulName As String, myProcName As String) '°MODUL: M99
    '°used in getModuleNameOfProc
    'Microsoft Visual Basic for Applications Extensibility 5.3 library
    Dim VBProj As VBIDE.VBProject, VBComp As VBIDE.VBComponent, CodeMod As VBIDE.CodeModule
    Dim LineNum As Long, ProcName As String, ProcKind As VBIDE.vbext_ProcKind, s As String
    Set VBProj = ActiveWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(ModulName)
    Set CodeMod = VBComp.CodeModule
    With CodeMod
        LineNum = .CountOfDeclarationLines + 1
        Do Until LineNum >= .CountOfLines
            ProcName = .ProcOfLine(LineNum, ProcKind)
            If myProcName = ProcName Then
                s = IIf(s = vbNullString, vbNullString, vbCrLf) & ModulName & "." & ProcName
                Exit Do
            End If
            LineNum = .ProcStartLine(ProcName, ProcKind) + .ProcCountLines(ProcName, ProcKind) + 1
        Loop
    End With
    getProcNameAifInModuleB = s
End Function

Sub TEST_getProceduresOfOneModule() '°MODUL: M99
    '°not used in another procedure
    show getProceduresOfOneModule("W1", 0)
End Sub

Function getProceduresOfOneModule(ModulName As String, Optional withModulName As Boolean = True) '°MODUL: M99
    '°used in getFunctionAndSubNames,TEST_getProceduresOfOneModule
    'Microsoft Visual Basic for Applications Extensibility 5.3 library
    Dim VBProj As VBIDE.VBProject, VBComp As VBIDE.VBComponent, CodeMod As VBIDE.CodeModule
    Dim LineNum As Long, ProcName As String, ProcKind As VBIDE.vbext_ProcKind, s As String
    Set VBProj = ActiveWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(ModulName)
    Set CodeMod = VBComp.CodeModule
    With CodeMod
        LineNum = .CountOfDeclarationLines + 1
        Do Until LineNum >= .CountOfLines
            ProcName = .ProcOfLine(LineNum, ProcKind)
            If withModulName Then
                s = s & IIf(s = vbNullString, vbNullString, vbCrLf) & ModulName & "." & ProcName
            Else
                s = s & IIf(s = vbNullString, vbNullString, vbCrLf) & ProcName
            End If
            LineNum = .ProcStartLine(ProcName, ProcKind) + .ProcCountLines(ProcName, ProcKind) + 1
        Loop
    End With
    getProceduresOfOneModule = s
End Function

Sub QuickSortDes(ByRef av As Variant, ByVal iBeg&, ByVal iEnd&)
    'from [URL]http://www.vba-programmer.com/Snippets/Code_VB/Quick_Sort_Single.html[/URL]
    Dim iLo&, iHi&, Temp As Variant, vSep As Variant
    iLo = iBeg: iHi = iEnd: vSep = av((iBeg + iEnd) / 2)
    Do
        Do While av(iLo) > vSep ' descending
            iLo = iLo + 1
        Loop
        Do While av(iHi) < vSep ' descending
            iHi = iHi - 1
        Loop
        If iLo <= iHi Then
            Temp = av(iLo): av(iLo) = av(iHi)
            av(iHi) = Temp: iLo = iLo + 1: iHi = iHi - 1
        End If
    Loop While iLo <= iHi
    If iBeg < iHi Then QuickSortDes av, iBeg, iHi
    If iLo < iEnd Then QuickSortDes av, iLo, iEnd
End Sub

Function getTextFrom2DArray(Arr, Optional NummernLinks As Boolean = True, Optional NummernOben As Boolean = True) As String '°MODUL: M99
    Dim i&, j&, MaxBreite%, E$, s$, si$, Title1$, zL%, sL%, zU%, sU%
    Dim Breite() As Integer, Arr2() As Integer
    'Einstellung
        MaxBreite = 40
    'arr(zL to zU, sL to sU)
        zL = LBound(Arr, 1): sL = LBound(Arr, 2): zU = UBound(Arr, 1): sU = UBound(Arr, 2)
    'maximale Länge eines Eintrages in Spalte i --> Breite(i)
        ReDim Breite(1 To sU)
        ReDim Arr2(1 To sU)
        For i = zL To zU
            For j = sL To sU
                If IsNull(Arr(i, j)) Then Arr(i, j) = vbNullString
                E = CStr(Arr(i, j)) '1 Eintrag
                If Breite(j) < 3 Then Breite(j) = 3
                If Len(E) > Breite(j) Then Breite(j) = Len(E)
            Next j
        Next i
    'Das 2D-Array listen:
        If NummernLinks Then s = "0001  |" Else s = "|"
        For i = zL To zU
            'si = Nummer links:
                If NummernLinks Then si = CStr(Format(CStr(i + 1), "0000"))
            For j = sL To sU
                E = CStr(Arr(i, j)) '1 Eintrag
                'ggf. dem Eintrag Leerzeichen anfügen (bis SOLL-Breite)
                    'If Breite(j) < 11 Then E = E + String(Breite(j) - Len(E), " ") Else E = Left(E + "           ", 11)
                    
                    
                    If Breite(j) < MaxBreite Then E = E + String(Breite(j) - Len(E), " ") Else E = Left(E + String(MaxBreite, " "), MaxBreite)
                    
                    
                s = s + E + "|"
                If NummernOben Then Arr2(j) = Len(E)
            Next j
            If NummernLinks Then
                If i < zU Then s = s + vbCrLf + si + "  |"
            Else
                If i < zU Then s = s + vbCrLf + "|"
            End If
        Next i
        If Right(s, 2) = vbCrLf + "|" Then s = Left(s, Len(s) - 2)
    'Title1-Zeile (SpaltenNummern):
        If NummernOben Then
            Title1 = "      |"
            For j = 1 To UBound(Arr2)
                Title1 = Title1 + CStr(j) + String(Arr2(j) - Len(CStr(j)), " ") + "|"
            Next
        End If
        If NummernOben Then s = Title1 + vbCrLf + "      " _
            + String(Len(Title1) - 5, "_") + vbCrLf + "      |" + vbCrLf + s
    getTextFrom2DArray = s
End Function

Sub showArray2D(ByRef Arr, Optional HeaderText$ = "") '°MODUL: M99
    Dim T$: If HeaderText <> "" Then T = HeaderText + vbCrLf + vbCrLf
    'G = getTextFrom2DArray(arr)
    show T + getTextFrom2DArray(Arr)
End Sub

Function LastRow(wks As Worksheet) As Long
    'Code funktioniert auch bei Autofilter
    Dim lngFirst As Long, lngLast As Long, lngTmp As Long
    With Application
        If .CountA(wks.Cells) = 0 Then Exit Function
        If .CountA(wks.Rows(wks.Rows.count)) Then
            LastRow = wks.Rows.count: Exit Function
        End If
        lngLast = wks.Rows.count
        Do While lngLast > lngFirst + 1
            lngTmp = (lngFirst + lngLast) \ 2
            If .CountA(wks.Rows(lngTmp).Resize(lngLast - lngTmp)) Then _
               lngFirst = lngTmp Else lngLast = lngTmp
        Loop
        If .CountA(wks.Rows(lngLast)) Then LastRow = lngLast Else LastRow = lngFirst
    End With
End Function

Function lastCol(wks As Worksheet) As Long
    'Code funktioniert auch bei Autofilter
    Dim lngFirst As Long, lngLast As Long, lngTmp As Long
    With Application
        If .CountA(wks.Cells) = 0 Then Exit Function
        If .CountA(wks.Columns(wks.Columns.count)) Then
            lastCol = wks.Columns.count: Exit Function
        End If
        lngLast = wks.Columns.count
        Do While lngLast > lngFirst + 1
            lngTmp = (lngFirst + lngLast) \ 2
            If .CountA(wks.Columns(lngTmp).Resize(, lngLast - lngTmp)) Then _
               lngFirst = lngTmp Else lngLast = lngTmp
        Loop
        If .CountA(wks.Columns(lngLast)) Then lastCol = lngLast Else lastCol = lngFirst
    End With
End Function

Sub Fenster_fixieren() '°MODUL: M99
    '°not used in another procedure
    'Ohne .Select das Fenster fixieren
    Dim rngZelle             As Range
    Set rngZelle = Range("A8")     'unterhalb und rechts der Zelle wird fixiert
    With ActiveWindow
        .SplitRow = rngZelle.Row
        .SplitColumn = rngZelle.Column
        .FreezePanes = True
    End With
    Set rngZelle = Nothing
End Sub

Sub ReplaceTextInCodeModuleA(NameOfModuleA As String, strFindWhat As String, strReplaceWith As String) '°MODUL: M99
    '°not used in another procedure
    'Search one code module for specific text; Replace every oldText 1x with newText
    Dim VBProj As VBProject, VBComp As VBComponent, CodeMod As CodeModule
    Dim sL As Long, eL As Long 'SL = start line     'EL = end line
    Dim sc As Long, EC As Long 'SC = start column   'EC = end column
    Dim strCodeLine As String, Found As Boolean
    Set VBProj = Application.VBE.ActiveVBProject
    Set VBComp = VBProj.VBComponents(NameOfModuleA)
    Set CodeMod = VBComp.CodeModule '    '.CodeModule
    With CodeMod
        sL = 1: eL = .CountOfLines: sc = 1: EC = 255 'Start/End Line/Column
        Found = .Find(Target:=strFindWhat, StartLine:=sL, StartColumn:=sc, _
            endline:=eL, EndColumn:=EC, _
            wholeWord:=False, MatchCase:=False, patternsearch:=False)
        If Found Then
            strCodeLine = CodeMod.lines(sL, 1)
            strCodeLine = Replace(strCodeLine, strFindWhat, strReplaceWith, Compare:=vbTextCompare) 'not case sensitive = vbTextCompare
            .ReplaceLine sL, strCodeLine
            'Debug.Print "Successfully Replaced: " & strFindWhat & " in VBA Module: " & NameOfModuleA & " with : " & strReplaceWith
        'Else
            'Debug.Print "Did not find: " & strFindWhat;
        End If
    End With
End Sub


