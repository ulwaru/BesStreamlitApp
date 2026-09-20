Attribute VB_Name = "M98"
Option Explicit 'M98

Sub ooo___M98()
    
    showProcs "first click"
    
    
    EE 1: Beep
End Sub

Sub Change_String1ToString2_Inside_SubFolderNames_OfOneSourceFolder_TEST()
    Dim p$
    DoArrc
    p = ArrC(3)  'Events
    'p = ArrC(4)  'Leute
    'p = ArrC(5)  'Register
    'Change_String1ToString2_Inside_SubFolderNames_OfOneSourceFolder "CZ", "CZE", p
    Change_String1ToString2_Inside_FileNames_OfSourceFolderAndSubFolders "CZ)", "CZE)", p
End Sub

Sub Change_String1ToString2_Inside_FileNames_OfSourceFolderAndSubFolders(String1$, String2$, PathOfSourceFolder$)
    Dim L$, N1$, N2$, p$, p1$, p2$, s$, i%, Arr1() As String
    p = PathOfSourceFolder
    L = p + "\*" + String1 + "*" 'LikeStringForPath
    s = Get_FilePaths_Like_insideSourceFolderAndSubFolders(p, L)
    s = Delete_EmptyRowsInString(s)
    If s = "" Then Beep: Exit Sub
    Arr1 = Split(s, vbCrLf)
    show "In " + CStr(UBound(Arr1) + 1) + " Dateien innerhalb des Ordners '" + getNameOfPath(p) _
        + "' wurde im Dateinamen '" + String1 + "' durch '" + String2 + "' ersetzt."
    For i = 0 To UBound(Arr1)
        p1 = Arr1(i) 'PathOfOneFile
        p2 = Replace(p1, String1, String2)
        RenameFolder p1, p2
    Next
    Beep
End Sub

Sub Change_String1ToString2_Inside_SubFolderNames_OfOneSourceFolder(String1$, String2$, PathOfSourceFolder$)
    Dim N1$, N2$, p$, p1$, p2$, s$, c%, i%, Arr1() As String
    p = PathOfSourceFolder
    s = Get_Paths_ofAllSubfolders_OneLevel(p)
    Arr1 = Split(s, vbCrLf)
    For i = 0 To UBound(Arr1)
        p1 = Arr1(i) 'PathOfOneSubFolderName
        N1 = getNameOfPath(p1)
        If N1 Like "*" + String1 + "*" Then
            c = c + 1
            N2 = Replace(N1, String1, String2)
            p2 = PathOfSourceFolder + "\" + N2
            RenameFolder p1, p2
        End If
    Next
    show "In " + CStr(c) + " UnterOrdner des Ordners '" + getNameOfPath(p) _
        + "' wurde im UnterOrdnerNamen '" + String1 + "' durch '" + String2 + "' ersetzt."
    Beep
End Sub

Sub DeleteContentOfColumnNr(NameOfSheet$, ColumnNr%)
    Worksheets(NameOfSheet).Columns(ColumnNr).ClearContents
End Sub

Function Get_FolderPath_Give_FilePath(PathOfFile$) As String
    If PathOfFile = "" Then Exit Function
    Get_FolderPath_Give_FilePath = Left(PathOfFile, InStrRev(PathOfFile, "\") - 1)
End Function

Sub CopyFile(source$, Destination$)
    'show source + vbCrLf + Destination
    FileCopy source, Destination
    'Stop
End Sub

Sub Paste_2DArrayToSheet(CellAsRange As Range, Arr As Variant)
    CellAsRange.Resize(UBound(Arr, 1), UBound(Arr, 2)).Value = Arr
End Sub

Sub Paste_2DArrayToCell_z_s(NameOfSheet$, z%, s%, Arr As Variant)
    Sheets(NameOfSheet).Cells(z, s).Resize(UBound(Arr, 1), UBound(Arr, 2)).Value = Arr
End Sub

Sub Paste_2DArray(NameOfSheet$, z%, s%, Arr As Variant)
    Sheets(NameOfSheet).Cells(z, s).Resize(UBound(Arr, 1), UBound(Arr, 2)).Value = Arr
End Sub

Sub Paste_1DArrayToCol(NameOfSheet$, z%, s%, Arr As Variant)
    'ganzes 1D-Array wird mit einem einzigen Zugriff auf ein Tabellenblatt geschrieben (mit 1 Schlag)
    'CellAsRange = einzelne Zelle (wird dann oberste Zelle der Spalte)
    Dim rg As Range
    Set rg = ThisWorkbook.Sheets(NameOfSheet).Cells(z, s)
    rg.Resize(UBound(Arr) - LBound(Arr) + 1, 1) = Application.Transpose(Arr)
End Sub

Sub Paste_1DArray_ToRow(NameOfSheet$, z%, s%, Arr As Variant)
    'ganzes 1D-Array wird mit einem einzigen Zugriff auf ein Tabellenblatt geschrieben (mit 1 Schlag)
    'CellAsRange = StartZelle in Zeile z
    Dim rg As Range
    Set rg = ThisWorkbook.Sheets(NameOfSheet).Cells(z, s)
    rg.Resize(1, UBound(Arr) - LBound(Arr) + 1) = Arr
End Sub

Sub TEST_CaptuVnNndOutFromShell_intoVariable()
    Dim sCmd$, p1$, p2$, qq$
    qq = Chr(34)
    p1 = "G:\Archiv VBA\Bes FaceTags\HilfsDateien\exiftool.exe"
    p2 = "G:\Archiv VBA\Bes FaceTags\test exiftool\rotation\4a rot90-6.jpg"
    'funktioniert, aber popup console
        sCmd = qq + p1 + qq + " -Orientation " + qq + p2 + qq
    'funktioniert
        'sCmd = "help"
    show ShellCapture(sCmd)
End Sub

Function ShellCapture(sCmd As String) As String
    Dim oShell   As New wshShell 'requires ref to Windows Script Host Object Model
    ShellCapture = oShell.Exec(sCmd).StdOut.ReadAll
End Function

Sub Wait_UntilFileSizeIsNotChanging_Step1sec(PathOfFile As String)
    Dim A As Long, B As Long, s As String ', i As Long
    s = "Size of " + Mid(PathOfFile, 1 + InStrRev(PathOfFile, "\"))
    Application.Wait (Now + TimeValue("0:00:01"))
    'OpenFolder PathOfFile
    DoEvents
    
    'Show PathOfFile
    
    
    A = FileSize(PathOfFile)
    'i = 10
    Do While A <> B
        [ar32] = CStr(A \ 1024 + 1) + " KB = " + s      '<-- adjust [ar32]
        Application.Wait (Now + TimeValue("0:00:01"))
        B = FileSize(PathOfFile)
        If A <> B Then A = B: B = 0
        DoEvents
    Loop
End Sub

Function FolderSize(PathOfFolder$) As String
    Dim fso As Object, fsoFolder As Object, c As Double
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsoFolder = fso.GetFolder(PathOfFolder)
    c = fsoFolder.size / 1024 / 1024 / 1024
    FolderSize = Format(c, "0.0") + " GB"
    Set fsoFolder = Nothing
    Set fso = Nothing
End Function

Function FileSize(ByVal filePath As String) As Long
  Dim size As Long
  On Local Error Resume Next
  size = FileLen(filePath)
  FileSize = IIf(Err = 0, size, -1)
  On Local Error GoTo 0
End Function

Function Get_TextWithTabLikeSpaces_LeftToEachStringX(txt$, StringX$) As String
        'Vorbereitung
            Dim A$, B$, z$, c%, i%, j%, max%, Arr1() As String
            A = "|°#": B = StringX
            txt = Delete_EmptyEndRowsInString(txt)
            txt = Replace(txt, B, A)
            txt = Replace99(txt, " " + A, A)
            Arr1 = Split(txt, vbCrLf)
            
        For j = 1 To 99
            For i = 0 To UBound(Arr1)
                z = Arr1(i)                         '1 Zeile
                c = InStr(1, z, A)                  'Position des ersten "|°#"
                If c > max Then max = c
            Next
            If max = 1 Then
                For i = 0 To UBound(Arr1)
                    z = Arr1(i)                     '1 Zeile
                    Arr1(i) = Replace(z, A, B, , 1) '1. Vorkommen ersetzen
                Next
            Else
                For i = 0 To UBound(Arr1)
                    z = Arr1(i)                     '1 Zeile
                    c = InStr(1, z, A)              'Position des ersten "|°#"
                        If c = 0 Then GoTo step2
                    z = Left(z, c - 1) + String(max - c + 1, " ") + Mid(z, c)
                    z = Replace(z, A, B, , 1)       '1. Vorkommen ersetzen
                    Arr1(i) = z
                Next
            End If
        Next
step2:
        Get_TextWithTabLikeSpaces_LeftToEachStringX = Join(Arr1, vbCrLf)
End Function

Sub Datentypen_Variables_short_signs()
    'shorts:    'String $, Integer(Int16) %, Long(Int32) &, Single !, Double #, Decimal @
    
    'Typ        Wertebereich                                                Speicherbedarf      Anfangswert         Beispiel
    
    'Boolean    Ja/Nein-Werte (True oder False)                             2 Byte              False               weiblich = True
    'Byte       Ganzzahlen 0 bis 255                                        1 Byte              0                   Alter = 42
    'Integer    Ganzzahlen -32.768 bis 32.767                               2 Byte              0                   Baujahr = 1950
    'Long       Ganzzahlen -2.147.483.648 bis 2.147.483.647                 4 Byte              0                   KundenNr = 23
    'LongLong   Ganzzahlen ca. ± 9 Trillionen                               8 Byte              0                   (64-Bit-Sys only)
    'Currency   skalierte Ganzzahlen ca. ± 9 Billionen, 4 Nachkommastellen  8 Byte              0                   Euro = -47.11
    'Single     Gleitkommazahl, einfache Genauigkeit                        4 Byte              0                   gerundet = 1.5
    'Double     Gleitkommazahl, doppelte Genauigkeit                        8 Byte              0                   Ergebnis = 1.75
    'Date       Datum (1. Jan 100 bis 31. Dez 9999) und Zeit                8 Byte              30.12.1899 00:00    Heute = #5/31/2023#
    'String     Texte                                                       Anz. Zeichen * 2    vbNullString        Info = "Hallo Welt"
    'Object     abhängig vom Objekt                                         4 Byte              Nothing             Set BMW = Auto
    'Variant    abhängig vom gerade aktuellen Inhalt                        unterschiedlich     Empty               diverses = -1
End Sub

Function FileExists(PathOfFile As String) As Boolean
    'nur Files! Liefert 'False' bei Ordnern
    If PathOfFile = "" Then Exit Function
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(PathOfFile)
'    Dim objFSO As Object
'    Set objFSO = CreateObject("Scripting.FileSystemObject")
'    FileExists = objFSO.FileExists(PathOfFile)
'    Set objFSO = Nothing
End Function

Function FileExists2(PathOfFile As String) As Boolean
    ' Prüft, ob die Datei existiert und kein Ordner ist
    If PathOfFile <> "" Then
        If Dir(PathOfFile) <> "" Then
            FileExists = True
        Else
            FileExists = False
        End If
    Else
        FileExists = False
    End If
End Function

Function FolderExists(ByVal PathOfFolder As String) As Boolean
    If PathOfFolder = "" Then Exit Function
    FolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(PathOfFolder)
    'FolderExists erkennt Pfade mit Umlauten
End Function

Function Replace99(MyString$, MyOld$, MyNew$) As String
    Dim i%
    For i = 1 To 99
        MyString = Replace(MyString, MyOld, MyNew)
        If InStr(1, MyString, MyOld) = 0 Then Exit For
    Next
    Replace99 = MyString
End Function

Sub ShowString(s As String) '°MODUL: M99
    'for use with userforms where 'Show' is reserved
    show s
End Sub

Function CountFilesInFolder(PathOfFolder As String, Optional strType As String)
    'EXAMPLE:  Call CountFilesInFolder("C:\Users\Ryan\", "*txt")
    Dim file As Variant, i%
    If Right(PathOfFolder, 1) <> "\" Then PathOfFolder = PathOfFolder & "\"
    file = Dir(PathOfFolder & strType)
    While (file <> "")
        i = i + 1: file = Dir
    Wend
    CountFilesInFolder = i
End Function

Sub showAt_TEST()
    showAt "Hallo", 100, 100, 300, 150
    showAt "Hallo", , 250
    showAt "Hallo", 393
    showAt "Hallo", , 100
End Sub

Sub showAt(s$, Optional L% = -1, Optional T% = -1, Optional w% = -1, Optional H% = -1)
    Dim p$, c%, TaskID#
    'Position of WinEditor
        If L > -1 Then CreateObject("WScript.Shell").RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Notepad\iWindowPosX", L, "REG_DWORD"
        If T > -1 Then CreateObject("WScript.Shell").RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Notepad\iWindowPosY", T, "REG_DWORD"
        If w > -1 Then CreateObject("WScript.Shell").RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Notepad\iWindowPosDX", w, "REG_DWORD"
        If H > -1 Then CreateObject("WScript.Shell").RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Notepad\iWindowPosDY", H, "REG_DWORD"
    'Write
        p = ThisWorkbook.path & "\old\tmpShow\"
        c = CountFilesInFolder(p)
        p = p + CStr(c + 1) + ".txt"
        writeStringToFile p, s
    'Open
        TaskID = shell("notepad.exe " & p, vbNormalFocus) 'öffnet sofort die Text-Datei
End Sub

Sub Read_SomeRegistryValues_Example()
    Dim L%, T%, w%, H%, v$ 'Left Top Width Height of WinEditor (Notepad.exe)
    v = vbCrLf
    'Read
        L = CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Notepad\iWindowPosX")
        T = CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Notepad\iWindowPosY")
        w = CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Notepad\iWindowPosDX")
        H = CreateObject("WScript.Shell").RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Notepad\iWindowPosDY")
        show L + v + T + v + w + v + H
End Sub

'sub ShowAt 200,200,500,500

Sub Write_SomeRegistryValues_Example()
    Dim L%, T%, w%, H% 'Left Top Width Height of WinEditor (Notepad.exe)
    'Write
        L = 200: T = 200: w = 500: H = 500
        CreateObject("WScript.Shell").RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Notepad\iWindowPosX", L, "REG_DWORD"
        CreateObject("WScript.Shell").RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Notepad\iWindowPosY", T, "REG_DWORD"
        CreateObject("WScript.Shell").RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Notepad\iWindowPosDX", w, "REG_DWORD"
        CreateObject("WScript.Shell").RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Notepad\iWindowPosDY", H, "REG_DWORD"
    'Show
        'show "Testfenster: L = 100: T = 100: W = 500: H = 500"
        show "Testfenster: L = 200: T = 200: W = 500: H = 500"
End Sub

Sub show(s As String)
    Dim F$, p$, ff$, v$, i&, max&, N&, TaskID#, A() As String
    v = vbCrLf
    p = ThisWorkbook.path & "\old - tmpShow-Files\"
    ff = Get_AllFileNamesOfOneFolder(p)
    A = Split(ff, v)
    For i = 0 To UBound(A)
        F = A(i): F = Replace(F, ".txt", "")
        If IsNumeric(F) Then
            N = CLng(F): If N > max Then max = N
        End If
    Next
    p = p + CStr(max + 1) + ".txt"
    writeStringToFile p, s
    'openFile p
        TaskID = shell("notepad.exe " & p, vbNormalFocus) 'öffnet sofort die Text-Datei
End Sub

Sub show16(s As String)
    Dim p$, p1$, p2$, TaskID#
    p = ThisWorkbook.path & "/helpers/tmpShow16.txt"
    WriteStringToUTF16TxtFile s, p
    'openFile p
    TaskID = shell("notepad.exe " & p, vbNormalFocus) 'öffnet sofort die Text-Datei
End Sub

Sub showArray(ByRef myArr() As String) '°MODUL: M98
    Dim s$, i&
    For i = LBound(myArr) To UBound(myArr)
                If CStr(myArr(i)) = "" Then s = s + " " + vbCrLf Else s = s + CStr(myArr(i)) + vbCrLf
        's = s + CStr(myArr(i)) + vbCrLf
    Next
    s = Delete_EmptyRowsInString(s)
    show s
End Sub

Sub showArray1DVariant(ByRef myArr())  '°MODUL: M98
    '°not used in another procedure
    Dim s$, i&
    For i = LBound(myArr) To UBound(myArr)
                If CStr(myArr(i)) = "" Then s = s + " " + vbCrLf Else s = s + CStr(myArr(i)) + vbCrLf
        's = s + CStr(myArr(i)) + vbCrLf
    Next
    s = Delete_EmptyRowsInString(s)
    show s
End Sub

Sub show2(N%, s$) '°MODUL: M98
    'Action:         Öffnet eine Datei in Notepad++
    'Merkwürdig     crash bei Versuch, diese proc aus 'Bes Ahnen' zu übernehmen
    '               verursacht durch eine 'TaskId = '-Zeile
    
    '°not used in another procedure
    Dim p1 As String, p2 As String, qq As String, TaskID#
    qq = Chr(34) 'quote, "-Zeichen
    p1 = "C:\Program Files\Notepad++\notepad++.exe"
    p2 = ThisWorkbook.path & "\HilfsDateien\npPlus" + CStr(N) + ".txt" 'HilfsDateien
    writeStringToFile p2, s
    
    
    
    'TaskId = x 'öffnet sofort die Text-Datei
    TaskID = shell("notepad.exe " & p2, vbNormalFocus) 'öffnet sofort die Text-Datei
    
    
    
End Sub

Sub showFile(Pfad$) 'ProjektModul+ '°MODUL: M98
    'showString textdatei
    shell "cmd /c " & Chr(34) & Pfad & Chr(34)
End Sub

Sub showFile2(p As String) '°MODUL: M99
    '°used in WriteT01
    Dim TaskID#
    TaskID = shell("notepad.exe " & p, vbNormalFocus) 'öffnet sofort die Text-Datei
End Sub

Sub showProcs(MyString$)
    'Vorbereitung
        Dim AllProcs$, m$, M1$, M2$, qq$, s$, s1$, s2$, FoundProcs$, v$
        Dim AnzAllProcs%, AnzFoundProcs%, Arr1() As String
        v = vbCrLf: m = Trim(MyString): qq = Chr(34) 'quote, "-Zeichen
    'Anzahl Prozeduren insgesamt
        AllProcs = getFunctionAndSubNames()
        AnzAllProcs = anzAinB(v, AllProcs)
    'All FoundProcs
        FoundProcs = GetAllProcNames_ContainingMyString_Sorted(MyString)
    'Found nothing
        If FoundProcs = "" Then
            s = "Der String  " + qq + m + qq + " ist in keiner der " _
            + CStr(AnzAllProcs) + " Prozeduren im Prozedurnamen enthalten."
            show s: Exit Sub
        End If
    'Found
        AnzFoundProcs = anzAinB(v, FoundProcs) + 1
    '1 Word in MyString
        If Not m Like "* *" Then
            s = "Der String  " + qq + m + qq + " ist in " + CStr(AnzFoundProcs) + " der " _
            + CStr(AnzAllProcs) + " Prozeduren im Prozedurnamen enthalten:" + v + v + FoundProcs
        End If
    '2 words in MyString
        If m Like "* *" Then
            M1 = Left(m, InStr(1, m, " ") - 1): M2 = Mid(m, InStr(1, m, " ") + 1)
            s = "Die Strings  " + qq + M1 + qq + " und " + qq + M2 + qq + "  sind in " _
            + CStr(AnzFoundProcs) + " der " + CStr(AnzAllProcs) + " Prozeduren " _
            + "im Prozedurnamen enthalten:" + v + v + FoundProcs
        End If
    'Finals
        show s
End Sub

Sub show_CodeToTextFile() '°MODUL: M98
    '°not used in another procedure
    Dim p$, s$, TaskID&
    p = ThisWorkbook.path & "\" & ThisWorkbook.NAME & "_VBA.txt"
    s = getModulesContentPepUp
    writeStringToFile p, s
    TaskID = shell("notepad.exe " & p, vbNormalFocus) 'öffnet sofort die Text-Datei
End Sub

Sub ShowCompleteCode() '°MODUL: M98
    '°not used in another procedure
    show getModulesContentPepUp
    'show tmpListOfVarsNotUsed(getListVarNotUsed)
    'show getListUsedIn
End Sub

Sub ShowProcNamesInsideProcNames() '°MODUL: M98
    '°not used in another procedure
    Dim AllProcs$, s$, i%, Arr1() As String
    AllProcs = getFunctionAndSubNames(False)
    Arr1 = Split(AllProcs, vbCrLf)
    For i = 0 To UBound(Arr1)
        
    Next
End Sub

Sub ListReferences() '°MODUL: M98
    '°used in Quick___W4,doProc
    Dim Arr1() As String, i As Integer
    Arr1 = Split(getReferences, vbCrLf)
    For i = 0 To UBound(Arr1)
        Cells(2 + i, 10) = Arr1(i)
        If i > 3 Then Cells(i, 9) = i - 3
    Next
    If i > 3 Then Cells(i, 9) = i - 3
    ActiveSheet.Range(Cells(4, 10), Cells(UBound(Arr1) + 1, 10)).Interior.Color = 10086143
End Sub
'--------------------------------------------------------------------------------------------
'START COPY MODUL AREA

Sub TEST_CopySomeCodeModules_FromOpenWorkbook1_to_OpenWorkbook2() '°MODUL: M98
    '°not used in another procedure
    Dim m$, p1$, p2$, i%, Arr1() As String
    Dim B As Boolean, Wb1 As Workbook, Wb2 As Workbook
    Dim Proj1 As VBIDE.VBProject, Proj2 As VBIDE.VBProject
    'alle p1-Module = "|M_p1|M_p2|M_Ahnen|ModulArrays|M_Data|M_Helpers|M_Draw|M_Stuff|M_Photos|M_Lines|M_move|Uf1|M_UserForm|M_Paste|M_p3|M_Einstellungen"
    'toDo = "|M_p1|M_p2"
    'all DocumentModules  = "|procs|W1ahn|W2dat|W3|W5bau|W7glo|W8csv"
    m = "|M_p2"
    p1 = "Bes Ahnen v087.xlsm" 'p1 = "G:\Archiv VBA\Bes Ahnen\Bes Ahnen v087.xlsm"
    p2 = "m2.xlsm" 'p2 = "G:\Archiv VBA\Bes Ahnen\m2.xlsm"
    Set Wb1 = Workbooks(p1)
    Set Wb2 = Workbooks(p2)
    Set Proj1 = Wb1.VBProject
    Set Proj2 = Wb2.VBProject
    Arr1 = Split(m, "|")
    For i = 1 To UBound(Arr1)
        B = CopyModule(Arr1(i), Proj1, Proj2, True)
    Next
End Sub

Sub TEST_CopySomeDocumentModules_FromOpenWorkbook1_to_OpenWorkbook2() '°MODUL: M98
    '°not used in another procedure
    'copySheets (incl. its Code)
    Dim m$, p1$, p2$, i%, Arr1() As String
    Dim Wb1 As Workbook, Wb2 As Workbook
    'all DocumentModules  = "|procs|W1ahn|W2dat|W3|W5bau|W7glo|W8csv"
    p1 = "Bes Ahnen v087.xlsm"  'p1 = "G:\Archiv VBA\Bes Ahnen\Bes Ahnen v087.xlsm"
    p2 = "m2.xlsm"              'p2 = "G:\Archiv VBA\Bes Ahnen\m2.xlsm"
    Set Wb1 = Workbooks(p1)
    Set Wb2 = Workbooks(p2)
    m = "|procs|W1ahn|W2dat|W3|W5bau|W7glo|W8csv"
    m = "|W2dat|W8csv"
    
    Arr1 = Split(m, "|")
    For i = 1 To UBound(Arr1)
        Wb1.Activate
        'WB1.Sheets("W3").Copy Before:=WB2.Sheets(1)
        Sheets(Arr1(i)).Copy After:=Wb2.Sheets(Wb2.Sheets.count)
    Next
End Sub

Function CopyModule(ModuleName As String, FromVBProject As VBIDE.VBProject, ToVBProject As VBIDE.VBProject, OverwriteExisting As Boolean) As Boolean
    '°used in TEST_CopySomeCodeModules_FromOpenWorkbook1_to_OpenWorkbook2
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' CopyModule
    ' This function copies a module from one VBProject to
    ' another. It returns True if successful or  False
    ' if an error occurs.
    '
    ' Parameters:
    ' --------------------------------
    ' FromVBProject         The VBProject that contains the module
    '                       to be copied.
    '
    ' ToVBProject           The VBProject into which the module is
    '                       to be copied.
    '
    ' ModuleName            The name of the module to copy.
    '
    ' OverwriteExisting     If True, the VBComponent named ModuleName
    '                       in ToVBProject will be removed before
    '                       importing the module. If False and
    '                       a VBComponent named ModuleName exists
    '                       in ToVBProject, the code will return
    '                       False.
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim VBComp As VBIDE.VBComponent
    Dim fName As String
    Dim CompName As String
    Dim s As String
    Dim SlashPos As Long
    Dim ExtPos As Long
    Dim TempVBComp As VBIDE.VBComponent
    
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Do some housekeeping validation.
    '''''''''''''''''''''''''''''''''''''''''''''
    If FromVBProject Is Nothing Then
        CopyModule = False
        Exit Function
    End If
    
    If Trim(ModuleName) = vbNullString Then
        CopyModule = False
        Exit Function
    End If
    
    If ToVBProject Is Nothing Then
        CopyModule = False
        Exit Function
    End If
    
    If FromVBProject.Protection = vbext_pp_locked Then
        CopyModule = False
        Exit Function
    End If
    
    If ToVBProject.Protection = vbext_pp_locked Then
        CopyModule = False
        Exit Function
    End If
    
    On Error Resume Next
    Set VBComp = FromVBProject.VBComponents(ModuleName)
    If Err.Number <> 0 Then
        CopyModule = False
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' FName is the name of the temporary file to be
    ' used in the Export/Import code.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    fName = Environ("Temp") & "\" & ModuleName & ".bas"
    If OverwriteExisting = True Then
        ''''''''''''''''''''''''''''''''''''''
        ' If OverwriteExisting is True, Kill
        ' the existing temp file and remove
        ' the existing VBComponent from the
        ' ToVBProject.
        ''''''''''''''''''''''''''''''''''''''
        If Dir(fName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
            Err.Clear
            Kill fName
            If Err.Number <> 0 Then
                CopyModule = False
                Exit Function
            End If
        End If
        With ToVBProject.VBComponents
            .Remove .item(ModuleName)
        End With
    Else
        '''''''''''''''''''''''''''''''''''''''''
        ' OverwriteExisting is False. If there is
        ' already a VBComponent named ModuleName,
        ' exit with a return code of False.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        Set VBComp = ToVBProject.VBComponents(ModuleName)
        If Err.Number <> 0 Then
            If Err.Number = 9 Then
                ' module doesn't exist. ignore error.
            Else
                ' other error. get out with return value of False
                CopyModule = False
                Exit Function
            End If
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Do the Export and Import operation using FName
    ' and then Kill FName.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    FromVBProject.VBComponents(ModuleName).Export fileName:=fName
    
    '''''''''''''''''''''''''''''''''''''
    ' Extract the module name from the
    ' export file name.
    '''''''''''''''''''''''''''''''''''''
    SlashPos = InStrRev(fName, "\")
    ExtPos = InStrRev(fName, ".")
    CompName = Mid(fName, SlashPos + 1, ExtPos - SlashPos - 1)
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Document modules (SheetX and ThisWorkbook)
    ' cannot be removed. So, if we are working with
    ' a document object, delete all code in that
    ' component and add the lines of FName
    ' back in to the module.
    ''''''''''''''''''''''''''''''''''''''''''''''
    Set VBComp = Nothing
    Set VBComp = ToVBProject.VBComponents(CompName)
    
    If VBComp Is Nothing Then
        ToVBProject.VBComponents.Import fileName:=fName
    Else
        If VBComp.Type = vbext_ct_Document Then
            ' VBComp is destination module
            Set TempVBComp = ToVBProject.VBComponents.Import(fName)
            ' TempVBComp is source module
            With VBComp.CodeModule
                .DeleteLines 1, .CountOfLines
                s = TempVBComp.CodeModule.lines(1, TempVBComp.CodeModule.CountOfLines)
                .InsertLines 1, s
            End With
            On Error GoTo 0
            ToVBProject.VBComponents.Remove TempVBComp
        End If
    End If
    Kill fName
    CopyModule = True
End Function

Public Sub CopySomeModulsToAnotherWorkbook() '°MODUL: M98
    '°not used in another procedure
    Dim Wb1 As Workbook, Wb2 As Workbook, m$, s$, i%, Arr1() As String
    Set Wb1 = Workbooks("Bes Ahnen v087.xlsm")
    Set Wb2 = Workbooks("m2.xlsm")
    'show getModulNames
    'copyModuls:
        m = "|M99|M98|M_Ahnen|M_Arrays|M_Data|M_Helpers|M_Draw|M_Stuff|M_Photos|M_Lines|M_move|Uf1|M_UserForm|M_Paste|M02|M_Einstellungen"
        m = "|M_Baum|M_CardsPos"
        Arr1 = Split(m, "|")
        For i = 1 To UBound(Arr1)
            Call CopyModule2(Wb1, Arr1(i), Wb2)
        Next
'     'copySheets (incl. its Code)
'        s = "|procs|W1ahn|W2dat|W3|W3|W7glo|W8csv"
'        ThisWorkbook.Activate
'        Sheets("W8").Copy Before:=WB2.Sheets(1)
'        'Sheets(arr1(i)).Copy After:=Workbooks("map.xlsx").Sheets(Workbooks("map.xlsx").Sheets.count)
End Sub

Public Sub CopyModule2(SourceWB As Workbook, strModuleName As String, TargetWB As Workbook) '°MODUL: M98
    '°used in CopySomeModulsToAnotherWorkbook
    ' Description:  copies a module from one workbook to another
    ' example: CopyModule2 Workbooks(ThisWorkbook), "Module2",
    '          Workbooks("Food Specials Rolling Depot Memo 46 - 01.xlsm")
    ' Notes:   If Module to be copied already exists, it is removed first,
    '          and afterwards copied
    Dim strFolder                       As String
    Dim strTempFile                     As String
    Dim fName                           As String
    If Trim(strModuleName) = vbNullString Then
        Exit Sub
    End If
    If TargetWB Is Nothing Then
        MsgBox "Error: Target Workbook " & TargetWB.NAME & " doesn't exist (or closed)", vbCritical
        Exit Sub
    End If
    strFolder = SourceWB.path
    If Len(strFolder) = 0 Then strFolder = CurDir
    ' create temp file and copy "Module2" into it
    strFolder = strFolder & "\"
    strTempFile = strFolder & "~tmpexport.bas"
    On Error Resume Next
    fName = Environ("Temp") & "\" & strModuleName & ".bas"
    If Dir(fName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
        Err.Clear
        Kill fName
        If Err.Number <> 0 Then
            MsgBox "Error copying module " & strModuleName & "  from Workbook " & SourceWB.NAME & " to Workbook " & TargetWB.NAME, vbInformation
            Exit Sub
        End If
    End If
    ' remove "Module2" if already exits in destination workbook
    With TargetWB.VBProject.VBComponents
        .Remove .item(strModuleName)
    End With
    ' copy "Module2" from temp file to destination workbook
    SourceWB.VBProject.VBComponents(strModuleName).Export strTempFile
    TargetWB.VBProject.VBComponents.Import strTempFile
    Kill strTempFile
    On Error GoTo 0
End Sub
'END COPY MODUL AREA
'--------------------------------------------------------------------------------------------

Function getRowNrWithMyPointsInside(MyPoints!) As Integer '°MODUL: M98
    '°not used in another procedure
    Dim i%, c%
    With ThisWorkbook.Sheets("W1")
        For i = 1 To 99
            'Zelle A100|A200|A300|...Top --> c = 200 (A200 < MyPoints < A300)
            If .Cells(i * 100, 1).Top > MyPoints Then c = (i - 1) * 100: Exit For
        Next
        For i = 1 To 99
            'Zelle A210|A220|A230|...Top --> c = 230 (A230 < MyPoints < A240)
            If .Cells(c + i * 10, 1).Top > MyPoints Then c = c + (i - 1) * 10: Exit For
        Next
        For i = 1 To 99
            'Zelle A231|A232|A233|...Top --> c = 237 (A237 < MyPoints < A238)
            If .Cells(c + i, 1).Top > MyPoints Then c = c + (i - 1): Exit For
        Next
        'MyPoints liegt in [A237]
    End With
    getRowNrWithMyPointsInside = c
End Function

Function getColNrWithMyPointsInside(MyPoints!) As Integer '°MODUL: M98
    '°not used in another procedure
    Dim i%, c%
    With ThisWorkbook.Sheets("W1")
        For i = 1 To 99
            If .Cells(1, i * 100).Left > MyPoints Then c = (i - 1) * 100: Exit For
        Next
        For i = 1 To 99
            If .Cells(1, c + i * 10).Left > MyPoints Then c = c + (i - 1) * 10: Exit For
        Next
        For i = 1 To 99
            If .Cells(1, c + i).Left > MyPoints Then c = c + (i - 1): Exit For
        Next
    End With
    getColNrWithMyPointsInside = c
End Function

Sub DruckbereichAufheben() '°MODUL: M98
    '°not used in another procedure
    'ThisWorkbook.Sheets("W1").PageSetup.PrintArea = Range("L12:z15").Address
    ThisWorkbook.Sheets("W1").PageSetup.PrintArea = ""
End Sub

Sub TEST_Linien() '°MODUL: M98
    '°not used in another procedure
    '   ActiveSheet.Shapes.AddLine(X1, Y1, X2, Y2).Name = <Name>
    'ActiveSheet.Shapes.AddLine(300, 100, 320, 100).Name = "L1"
    ActiveSheet.Shapes("L1").SetShapesDefaultProperties
    'ActiveSheet.Shapes("L1").Line.Weight = 1.5
    'ActiveSheet.Shapes("L1").Line.ForeColor.RGB = RGB(0, 0, 0)
End Sub

Sub openCsvFolder() '°MODUL: M98
    '°not used in another procedure
    openFileOrFolder ThisWorkbook.path + "\csv"
End Sub
'openFileOrFolder

Function FileExists3(ByVal FileToTest As String) As Boolean
    FileExists = (Dir(FileToTest) <> "")
End Function

Function FileExistsGA(ByVal filespec As String) As Boolean
  ' Karl Peterson MS VB MVP
  Dim Attr As Long
  ' Guard against bad FileSpec by ignoring errors
  ' retrieving its attributes.
  On Error Resume Next
  Attr = GetAttr(filespec)
  If Err.Number = 0 Then
    ' No error, so something was found.
    ' If Directory attribute set, then not a file.
    FileExistsGA = Not ((Attr And vbDirectory) = vbDirectory)
  End If
End Function

Sub DeleteFile(ByVal FileToDelete As String)
    'Called from    Delete_AllLinkFiles_insideOneFolder Create_OneLinkFile Update_LeuteLinks_inOneEventFolder
    On Error GoTo DeleteFile_Error
    If FileExists(FileToDelete) Then
        'First remove readonly attribute, if set
        SetAttr FileToDelete, vbNormal
        Kill FileToDelete
    End If
    
    On Error GoTo 0
    Exit Sub
DeleteFile_Error:
    show "Error " + CStr(Err.Number) + " (" + Err.Description _
        + ") in procedure DeleteFile, line " + CStr(Erl) + "." + vbCrLf + vbCrLf _
        + "Ggf. wurde wegen eines Zeitproblems 'Zugriff verweigert';" _
        + "deshalb: 'Application.Wait Now + #0:00:01#: Kill FileToDelete'"
        Application.Wait Now + #12:00:01 AM#: Kill FileToDelete
End Sub

Sub DeleteFolder(PathOfFolder$)
    On Error Resume Next
    'delete all files in folder
        Kill PathOfFolder + "\*.*"
    'delete empty folder
        RmDir PathOfFolder + "\"
    On Error GoTo 0
End Sub

Public Sub Timer2() '°MODUL: M98
    '°not used in another procedure
    Dim startTime As Long, diff As Long, p As String
    p = "G:\Archiv VBA"
    startTime = GetTickCount()
    'procs_DialogForFolder
    'TESTcount_Subfolders
    'procs_ListFolderPaths p
    'NonRecursiveMethod p
    'TEST2 p
    '[j22] = getFileCount(p)
    'test5 p
    diff = (GetTickCount() - startTime)
    show CStr(CDbl(diff / 1000)) & " seconds"
    Beep
End Sub

Function Get_1DArrayToString(ByRef myArr() As String) As String '°MODUL: M98
    '°not used in another procedure
    Dim s$, i&
    For i = LBound(myArr) To UBound(myArr)
        s = s + CStr(myArr(i)) + vbCrLf
    Next
    s = Delete_EmptyRowsInString(s)
    Get_1DArrayToString = s
End Function

Sub RenameFile(PathOfFile1 As String, PathOfFile2 As String) '°MODUL: M98
    '°not used in another procedure
    On Error Resume Next
    'show PathOfFile1 + vbCrLf + PathOfFile2
    Name PathOfFile1 As PathOfFile2
End Sub

Sub RenameFolder(PathOfFolder1 As String, PathOfFolder2 As String) '°MODUL: M98
    '°used in RenameFolders_Date10CharsTo8Chars
    'On Error Resume Next
    Name PathOfFolder1 As PathOfFolder2
    'Stop
End Sub

Function Get_Number_OfSubFolders_OneLevel(PathOfFolder$) As Long
    Dim oFSO As Object, folder As Object, subfolders As Object, c As Long
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set folder = oFSO.GetFolder(PathOfFolder)
    Set subfolders = folder.subfolders
    Get_Number_OfSubFolders_OneLevel = subfolders.count
    Set oFSO = Nothing: Set folder = Nothing: Set subfolders = Nothing
End Function

Sub Get_Number_OfFileNames_withMyString_inFolderAndSubFolders_TEST()
    Dim p$, MyString$
    p = "F:\Archiv TR\Archiv Trampolin\Events"
    MyString = "p123"
    show CStr(Get_Number_OfFileNames_withMyString_inFolderAndSubFolders(p, MyString))
End Sub

Function Get_Number_OfFileNames_withMyString_inFolderAndSubFolders(PathOfSourceFolder$, MyString$) As Long
    Dim c&
    'PathOfSourceFolder an Sub übergeben; dort rekursiv div. Subfolder statt PathOfSourceFolder
        Call Helper_Count01(PathOfSourceFolder, MyString, c) 'recursiv
    'Rückgabe
        Get_Number_OfFileNames_withMyString_inFolderAndSubFolders = c
End Function

Sub Helper_Count01(ByVal folderPath$, MyString$, c&)
    'Called from:   Get_Number_OfJPGs_inFolderAndSubfolders, (+recursiv)
    
    Dim fileName$, fullFilePath$, numFolders&, folders() As String, i As Long
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    fileName = Dir(folderPath & "*.*", vbDirectory)
    
    While Len(fileName) <> 0
        'If left(fileName, 1) <> "." Or fileName Like ".*[!.]*" Then
        If Left(fileName, 1) <> "." Then
            fullFilePath = folderPath & fileName
            If (GetAttr(fullFilePath) And vbDirectory) = vbDirectory Then
                ReDim Preserve folders(0 To numFolders) As String
                folders(numFolders) = fullFilePath
                numFolders = numFolders + 1
            Else
                'Insert the actions to be performed on each file
                If LCase(fileName) Like "*" + MyString + "*" Then c = c + 1
            End If
        End If
        fileName = Dir()
    Wend
    For i = 0 To numFolders - 1
        Helper_Count01 folders(i), MyString, c&
    Next i
End Sub

Sub Get_Number_OfJPGs_inFolderAndSubfolders_TEST()
    Dim p$
    p = "F:\Archiv TR\Archiv Trampolin\"
    show CStr(Get_Number_OfJPGs_inFolderAndSubfolders(p))
End Sub
Function Get_Number_OfJPGs_inFolderAndSubfolders(PathOfSourceFolder$) As Long
    Dim c&
    'PathOfSourceFolder an Sub übergeben; dort rekursiv div. Subfolder statt PathOfSourceFolder
        Call Count_Jpg_inFolderAndSubfolders(PathOfSourceFolder, c) 'recursiv
    'Rückgabe
        Get_Number_OfJPGs_inFolderAndSubfolders = c
End Function

Sub Helper_Count02(ByVal folderPath As String, c&)
    'Called from:   Get_Number_OfJPGs_inFolderAndSubfolders, (+recursiv)
    
    Dim fileName$, fullFilePath$, numFolders&, folders() As String, i As Long
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    fileName = Dir(folderPath & "*.*", vbDirectory)
    
    While Len(fileName) <> 0
        'If left(fileName, 1) <> "." Or fileName Like ".*[!.]*" Then
        If Left(fileName, 1) <> "." Then
            fullFilePath = folderPath & fileName
            If (GetAttr(fullFilePath) And vbDirectory) = vbDirectory Then
                ReDim Preserve folders(0 To numFolders) As String
                folders(numFolders) = fullFilePath
                numFolders = numFolders + 1
            Else
                'Insert the actions to be performed on each file
                If LCase(fileName) Like "*.jpg" Then c = c + 1
            End If
        End If
        fileName = Dir()
    Wend
    For i = 0 To numFolders - 1
        Helper_Count02 folders(i), c&
    Next i
    '
End Sub

Function Get_FileCount_String(ByVal folder As Variant, Optional ByVal FileFilter As String) As Variant
    'no Subfolders
    Dim Files As Object
    If FileFilter = "" Then FileFilter = "*.*"
    With CreateObject("Shell.Application")
        Set Files = .Namespace(folder).items
        Files.Filter 64, FileFilter
        Get_FileCount_String = Files.count
    End With
End Function

Function GetFileCount(localRoot, Optional fld, Optional count As Long) As Long '°MODUL: M98
    'recursiv; no global variables needed
    Dim fso, baseFolder, subFolder
    Set fso = CreateObject("Scripting.Filesystemobject")
    If IsMissing(fld) Then Set baseFolder = fso.GetFolder(localRoot) Else Set baseFolder = fld
    count = count + baseFolder.Files.count
    For Each subFolder In baseFolder.subfolders
        'GetFileCount localRoot, subFolder, count
        GetFileCount localRoot, subFolder, count
    Next
    GetFileCount = count
End Function

Function Delete_EmptyRowsInString(s As String) '°MODUL: M98
'    Do While InStr(1, s, vbCrLf + vbCrLf) > 0
'        s = Replace(s, vbCrLf + vbCrLf, vbCrLf)
'        s = Replace(s, vbCrLf + vbCrLf, vbCrLf)
'    Loop
    s = Replace99(s, vbCrLf + vbCrLf, vbCrLf)
    If Left(s, 2) = vbCrLf Then s = Mid(s, 3)
    If Right(s, 2) = vbCrLf Then s = Left(s, Len(s) - 2)
    Delete_EmptyRowsInString = s
End Function

Function Delete_EmptyEndRowsInString(s As String) '°MODUL: M98
    '°not used in another procedure
    Do While Right(s, 2) = vbCrLf
        s = Left(s, Len(s) - 2)
    Loop
    Delete_EmptyEndRowsInString = s
End Function

Function Delete_EmptyTrimmedRowsInString(s As String)
    s = Replace(s, "°", "$x%ß&")
    s = Replace99(s, " " + vbCr, vbCr)
    s = Replace99(s, " " + vbLf, vbLf)
    s = Replace(s, vbCr, "°")
    s = Replace(s, vbLf, "°")
    s = Replace99(s, "°°", "°")
    s = Replace(s, "°", vbCrLf)
    s = Replace(s, "$x%ß&", "°")
    If Left(s, 2) = vbCrLf Then s = Mid(s, 3)
    If Right(s, 2) = vbCrLf Then s = Left(s, Len(s) - 2)
    Delete_EmptyTrimmedRowsInString = s
End Function

Function Delete_EndReturnsInString(text As String) As String '°MODUL: M99
    '°not used in another procedure
    Do While Right(text, 1) = vbCr Or Right(text, 1) = vbLf
        text = Left(text, Len(text) - 1)
    Loop
    Delete_EndReturnsInString = text
End Function

Function Get_AllFileNamesOfOneFolder(PathOfFolder As String) As String '°MODUL: M98
    Dim fso As Object, objVerzeichnis As Object, objDateienliste As Object, objDatei As Object, s As String
    Set fso = CreateObject("scripting.FileSystemObject")
    Set objVerzeichnis = fso.GetFolder(PathOfFolder)
    Set objDateienliste = objVerzeichnis.Files
    For Each objDatei In objDateienliste
         If Not objDatei Is Nothing Then
              s = s + objDatei.NAME + vbCrLf
         End If
    Next objDatei
    s = Delete_EmptyEndRowsInString(s)
    Get_AllFileNamesOfOneFolder = s
End Function

Sub Load_AllFileNames_Like_InsideOneFolder(PathOfOneFolder$, MyString$, Array1D() As String)
    Dim s$: s = Get_AllFileNames_Like_OfOneFolder(PathOfOneFolder$, MyString$)
    Array1D = Split(vbCrLf + s, vbCrLf)
End Sub

Sub Load_AllLinkFileNames_InsideOneFolder(PathOfFolder$, Arr() As String)
    Dim s$
    s = Get_AllFileNames_Like_OfOneFolder(PathOfFolder, "*.lnk")
    Arr = Split(s, vbCrLf)
End Sub

Sub Count_Files_OfOneFolder_TEST()
    Dim p$, c%: DoArrc
    p = ArrC(3) + "\1996-04-20 BadWüM Gernsbach"
    c = Count_Files_OfOneFolder(p)
    show CStr(c)
End Sub

Function Count_Files_OfOneFolder(PathOfFolder$) As Integer
    'Called from    xxx

    'Vorbereitung
        Dim fso As Object, objFiles As Object, obj As Object, c%
    'FSO
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set objFiles = fso.GetFolder(PathOfFolder).Files
    'Count
        c = objFiles.count
    'Finals
        Set objFiles = Nothing: Set fso = Nothing: Set obj = Nothing
        Count_Files_OfOneFolder = c
End Function

Sub Load_AllFilePaths_LikeMyStringInFileName_OfOneFolder(PathOfOneFolder$, LikeString$, Arr() As String)
    Dim p$
    p = Get_AllFilePaths_LikeMyStringInFileName_OfOneFolder(PathOfOneFolder$, LikeString$)
    Arr = Split(p, vbCrLf)
End Sub

Function Get_AllFilePaths_LikeMyStringInFileName_OfOneFolder(PathOfOneFolder$, LikeString$) As String
    'Called from    xxx

    'Vorbereitung
        Dim fso As Object, objVerzeichnis As Object, objDateienliste As Object
        Dim objDatei As Object, s$, N$
    'FSO
        Set fso = CreateObject("scripting.FileSystemObject")
        Set objVerzeichnis = fso.GetFolder(PathOfOneFolder)
        Set objDateienliste = objVerzeichnis.Files
    'All FilePaths of one folder
        For Each objDatei In objDateienliste
             If Not objDatei Is Nothing Then
                N = CStr(objDatei.NAME)
                If N Like LikeString Then s = s + objDatei.path + vbCrLf
             End If
        Next objDatei
        s = Delete_EmptyRowsInString(s)
    'Finals
        Get_AllFilePaths_LikeMyStringInFileName_OfOneFolder = s
End Function

Function Get_AllFilePaths_WithMyStringInFileName_OfOneFolder(PathOfOneFolder$, MyString$) As String '°MODUL: M98
    '°not used in another procedure
    Dim fso As Object, objVerzeichnis As Object, objDateienliste As Object
    Dim objDatei As Object, s$, N$
    Set fso = CreateObject("scripting.FileSystemObject")
    Set objVerzeichnis = fso.GetFolder(PathOfOneFolder)
    Set objDateienliste = objVerzeichnis.Files
    For Each objDatei In objDateienliste
         If Not objDatei Is Nothing Then
            N = CStr(objDatei.NAME)
            If InStr(1, LCase(objDatei.NAME), LCase(MyString)) > 0 Then
              s = s + objDatei.path + vbCrLf
            End If
         End If
    Next objDatei
    s = Delete_EmptyRowsInString(s)
    Get_AllFilePaths_WithMyStringInFileName_OfOneFolder = s
End Function

Function Get_AllFilePaths_WithMyStringInFilePath_OfOneFolder(PathOfOneFolder$, MyString$) As String
    '°not used in another procedure
    Dim fso As Object, objVerzeichnis As Object, objDateienliste As Object
    Dim objDatei As Object, s As String
    Set fso = CreateObject("scripting.FileSystemObject")
    Set objVerzeichnis = fso.GetFolder(PathOfOneFolder)
    Set objDateienliste = objVerzeichnis.Files
    For Each objDatei In objDateienliste
         If Not objDatei Is Nothing Then
            If InStr(1, LCase(objDatei.NAME), LCase(MyString)) > 0 Then
              s = s + objDatei.path + vbCrLf
            End If
         End If
    Next objDatei
    s = Delete_EmptyRowsInString(s)
    Get_AllFilePaths_WithMyStringInFilePath_OfOneFolder = s
End Function

Sub ListFiles2a() '°MODUL: M98
    '°not used in another procedure
    Dim FileSystem As Object, HostFolder As String, s As String, CountFolders As Long, cFi As Long, cx As Long
    HostFolder = "G:\Archiv Photos\- noch einordnen\- PhotoSync Handy\iPh14 Ulli\"
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    ListFiles2b FileSystem.GetFolder(HostFolder), s, CountFolders, cFi, cx
    show s
End Sub

Sub ListFiles2b(folder, s As String, CountFolders As Long, cFi As Long, cx As Long) '°MODUL: M98
    '°used in ListFiles2a
    Dim subFolder, file
    For Each subFolder In folder.subfolders
        CountFolders = CountFolders + 1: [c15] = CountFolders: DoEvents
        ListFiles2b subFolder, s, CountFolders, cFi, cx
    Next
    For Each file In folder.Files
        'Operate on each file
        cFi = cFi + 1: [c16] = cFi: DoEvents
        If Right(file, 5) = ".jpg" Then
            s = s + file + vbCrLf
            cx = cx + 1: [C17] = cx: DoEvents
        End If
    Next
End Sub

Public Sub NonRecursiveMethod(p As String) '°MODUL: M98
    '°not used in another procedure
    Dim fso, oFolder, oSubfolder, oFile, queue As Collection
    Dim CountFolders As Long, cFi As Long, cx As Long, s As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set queue = New Collection
    queue.Add fso.GetFolder(p)
    Do While queue.count > 0
        Set oFolder = queue(1)
        queue.Remove 1 'dequeue
        '...insert any folder processing code here...
        CountFolders = CountFolders + 1: [j25] = CountFolders: DoEvents
        
        For Each oSubfolder In oFolder.subfolders
            queue.Add oSubfolder 'enqueue
        Next oSubfolder
        For Each oFile In oFolder.Files
            '...insert any file processing code here...
            cFi = cFi + 1: [J26] = cFi: DoEvents
            If Right(oFile, 5) = ".xlsm" Then
                s = s + oFile + vbCrLf
                cx = cx + 1: [J17] = cx: DoEvents
            End If
        Next oFile
    Loop
    show s
End Sub

Sub SammlungVBAprocsPepUp() '°MODUL: M98
    '°not used in another procedure
    Dim s As String, s1 As String
    s = getSammlungVBAprocs
    s = getNamesOfUsedExcelFiles(s) + s
            s1 = "Len = " + CStr(Len(s)) + " ": [A1] = 1
    s = deleteLinesPepUp(s)
            s1 = s1 + CStr(Len(s)) + " ": [A1] = 2
    s = deleteDoubledProcs(s)
            s1 = s1 + CStr(Len(s)) + " "
    show s1 + vbCrLf + s
    Beep
End Sub

Function deleteDoubledProcs(s3 As String) As String '°MODUL: M98
    '°used in SammlungVBAprocsPepUp
    Dim s As String, s1 As String, Arr1() As String
    Dim i As Long, c As Long
    '"Sub ","Function ","Private Sub ","Private Function ","Public Sub ","Public Function "
    s1 = s3
    'alle Procs in array:
    s1 = Replace(s1, vbCrLf + "'Gesamter VBA-", vbCrLf + "[%]'Gesamter VBA-")
    s1 = Replace(s1, vbCrLf + "'MODUL: ", vbCrLf + "[%]'MODUL:")
    s1 = Replace(s1, vbCrLf + "Sub ", vbCrLf + "[%]Sub ")
    s1 = Replace(s1, vbCrLf + "Function ", vbCrLf + "[%]Function ")
    s1 = Replace(s1, vbCrLf + "Private Sub ", vbCrLf + "[%]Private Sub ")
    s1 = Replace(s1, vbCrLf + "Private Function ", vbCrLf + "[%]Private Function ")
    s1 = Replace(s1, vbCrLf + "Public Sub ", vbCrLf + "[%]Public Sub ")
    s1 = Replace(s1, vbCrLf + "Public Function ", vbCrLf + "[%]Public Function ")
    Arr1 = Split(s1, "[%]")
    
    For i = 0 To UBound(Arr1)
        s = Arr1(i) 's = one complete proc
        If Left(s, 4) = "Sub " Or Left(s, 9) = "Function " _
                        Or Left(s, 8) = "Private " Or Left(s, 7) = "Public " Then
            c = InStr(1, s3, s)
            If c > 0 Then
                s3 = Replace(s3, s, "", Compare:=vbTextCompare) 'löscht s
                s3 = Insert(s3, s, c)
            End If
        End If
        [A1] = UBound(Arr1) - i
    Next
    deleteDoubledProcs = s3
End Function

Function Insert(original As String, added As String, pos As Long) As String '°MODUL: M98
    '°used in importiereFotoF1,importiereFotoF2,deleteDoubledProcs
    'If pos < 1 Then pos = 1
    If Len(original) < pos Then pos = Len(original) + 1
    On Error Resume Next
    If pos > 0 Then
        Insert = Mid(original, 1, pos - 1) & added & Mid(original, pos, Len(original) - pos + 1)
    Else
        Insert = original
    End If
    On Error GoTo 0
End Function

Function deleteLinesPepUp(s1 As String) As String '°MODUL: M98
    '°used in SammlungVBAprocsPepUp
    Dim s As String, s2 As String, C1 As Long, C2 As Long
    s = s1
    s = Replace(s, vbCrLf + "'°Gesamter VBA-Code (Makros", vbCrLf + "'Gesamter VBA-Code (Makros")
    s = Replace(s, vbCrLf + "'°MODUL: ", vbCrLf + "'MODUL: ")
    s = Replace(s, vbCrLf + "    '°", vbCrLf + "'°")
    s = Replace(s, ") '°", ")" + vbCrLf + "'°")
    s = Replace(s, "Option Explicit" + vbCrLf, "")
    Do While InStr(1, s, vbCrLf + "'°") > 0
        C1 = InStr(1, s, vbCrLf + "'°")
        C2 = InStr(C1 + 1, s, vbCrLf)
        s2 = Mid(s, C1, C2 - C1)
        s = Replace(s, s2, "", 1, 1)
    Loop
    Do While InStr(1, s, " " + vbCrLf) > 0
        s = Replace(s, " " + vbCrLf, vbCrLf)
    Loop
    Do While InStr(1, s, vbCrLf + vbCrLf + vbCrLf) > 0
        s = Replace(s, vbCrLf + vbCrLf + vbCrLf, vbCrLf + vbCrLf)
    Loop
    
    
    deleteLinesPepUp = s
End Function

Function getNamesOfUsedExcelFiles(s1 As String) As String '°MODUL: M98
    '°used in SammlungVBAprocsPepUp
    Dim s As String, anz1 As Integer, c As Integer
    Dim C1 As Long, C2 As Long, C3 As Long, C4 As Long
    s1 = vbCrLf + s1
    anz1 = procs_getAnzAinB(vbCrLf + "'°Gesamter VBA-Code (Makros", s1)
    s = "Prozeduren aus folgenden " + CStr(anz1) + " Excel-Dateien:" + vbCrLf
    c = 1
    Do While InStr(C4 + 1, s1, vbCrLf + "'°Gesamter VBA-Code (Makros") > 0
        C1 = InStr(C4 + 1, s1, vbCrLf + "'°Gesamter VBA-Code (Makros")
        C2 = InStr(C1, s1, "<")
        C3 = InStr(C1, s1, ">"): C4 = C3
        c = c + 1
        If c Mod 2 = 0 Then
            s = s + "   " + Mid(s1, C2 + 1, C3 - C2 - 1) + ", "
        Else
            s = s + Mid(s1, C2 + 1, C3 - C2 - 1) + ", " + vbCrLf
        End If
    Loop
    s = Left(s, Len(s) - 2) + vbCrLf + vbCrLf
    getNamesOfUsedExcelFiles = s
End Function

Function getSammlungVBAprocs() '°MODUL: M98
    '°used in SammlungVBAprocsPepUp
    Dim p As String
    p = "G:\Archiv VBA\SammlungVBAprocs.txt"
    getSammlungVBAprocs = ReadSmallFile(p)
End Function

Sub RenameModule(oldName As String, NewName As String) '°MODUL: M98
    '°used in Quick___M4
    ActiveWorkbook.VBProject.VBComponents(oldName).NAME = NewName
End Sub

Function procs_getAnzAinB(A As String, B As String) As Long '°MODUL: M98
    '°used in getNamesOfUsedExcelFiles
    Dim s As String
    s = Replace(B, A, "")
    procs_getAnzAinB = (Len(B) - Len(s)) / Len(A)
End Function

Function getRGB(rng As Range) As String '°MODUL: M98
    '°not used in another procedure
    Dim colorVal As Variant
    colorVal = Cells(rng.Row, rng.Column).Interior.Color
    getRGB = "RGB(" + CStr(colorVal Mod 256) & ", " & CStr((colorVal \ 256) Mod 256) & ", " & CStr(colorVal \ 65536) + ")"
End Function

Sub TEST_getColor() '°MODUL: M98
    '°not used in another procedure
    show CStr(getColor([au47], 0, "T6"))
    'Green1 14348258 'Green2 11854022 'Green3 9359529 'Green4 3506772
    'If Sheets(NameOfSheet).Cells(rng.Row, rng.Column).Interior.Color = 14348258 then ...
End Sub

Function getColor(rng As Range, Optional formatType As Integer = 0, Optional NameOfSheet As String = "") As Variant '°MODUL: M98
    '°used in TESTcolor,Get_ColorGreen1,Worksheet_SelectionChange,TEST_getColor
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Function            Color; Determine the Background Color Of a Cell
    '   rng                 Range to Determine Background Color of
    '   formatType          Default Value = 0 '0 Integer, 1 Hex, 2 RGB, 3 Excel Color Index
    '   Usage               Color(A1)      -->   9507341
    '                       Color(A1, 0)   -->   9507341
    '                       Color(A1, 1)   -->   91120D
    '                       Color(A1, 2)   -->   13, 18, 145
    '                       Color(A1, 3)   -->   6
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim colorVal As Variant
    If NameOfSheet = "" Then
        colorVal = Cells(rng.Row, rng.Column).Interior.Color
    Else
        colorVal = Sheets(NameOfSheet).Cells(rng.Row, rng.Column).Interior.Color
    End If
    Select Case formatType
        Case 1
            getColor = Hex(colorVal)
        Case 2
            getColor = (colorVal Mod 256) & ", " & ((colorVal \ 256) Mod 256) & ", " & (colorVal \ 65536)
        Case 3
            getColor = Cells(rng.Row, rng.Column).Interior.ColorIndex
        Case Else
            getColor = colorVal
    End Select
End Function

Function SheetExists(NameOfWorksheet$) As Boolean '°MODUL: M98
    '°not used in another procedure
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.NAME = NameOfWorksheet Then SheetExists = True
    Next
End Function














































































