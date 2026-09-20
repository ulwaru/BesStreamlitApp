Attribute VB_Name = "M04"
Option Explicit 'M04

Public ArrSubFolders() As String
Public CounterForSubFolders As Long

Sub zzz_M04()
    'doArr
    showProcs "last"
    Application.EnableEvents = True
    'show Get_AllPathsOfFoldersWithTonfotosIni_EverythingSearch
    'RenameModule "Modul1", "CharCheck"
    'Create_FaceImages
End Sub

Sub Get_CountOfPaths_OfAllSubfolders_AllLevels_TEST()
    Dim p$, v$, c%
    v = vbCrLf: DoArrc
    p = ArrC(3)                'T0 Path of Folder "Events"
    c = Get_CountOfPaths_OfAllSubfolders_AllLevels(p)
    show CStr(c)
End Sub

Function Get_CountOfPaths_OfAllSubfolders_AllLevels(SourceFolder$, Optional MyCountDown As Boolean = False) As Integer
    'Called from    T6_Get_PidPath_OfEachFileInOneFolder
    'Status         No global variables, no Collection
    
   'Vorbereitung
        Dim F$, s$, c%, FileSystem As Object
        F = SourceFolder
        If Right(F, 1) <> "/" Then F = F + "/"
    'FileSystemObject
        Set FileSystem = CreateObject("Scripting.FileSystemObject")
    'Alle Pfade in c zählen
        'If MyCountDown Then MyCountDown_MainStepsAllowed 20: MyCountDown_SubStepsMax AnzFold
        DoFolderRecursivGetCountOfPathsPartTwo c, FileSystem.GetFolder(F), s, MyCountDown
    'Finals
        Get_CountOfPaths_OfAllSubfolders_AllLevels = c
End Function

Sub DoFolderRecursivGetCountOfPathsPartTwo(c%, folder, s$, MyCountDown As Boolean)
    'Called from    Get_Paths_OfAllSubfolders_AllLevels, DoFolderRecursivGetPathsPartTwo
    Dim subFolder As Object
    For Each subFolder In folder.subfolders
        'If MyCountDown Then MyCountDown_OneMoreMainStep
        c = c + 1
        DoFolderRecursivGetCountOfPathsPartTwo c, subFolder, s, MyCountDown
    Next
End Sub


'#################
Sub Get_CountOfSubFoldersInsideOneSourceFolder_TEST()
    Dim s$
    DoArrc
    s = CStr(Get_CountOfSubFoldersInsideOneSourceFolder(ArrC(3)))
    show s
End Sub

Function Get_CountOfSubFoldersInsideOneSourceFolder(SourceFolder$) As Integer
    Dim oFSO As Object, folder As Object, subfolders As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set folder = oFSO.GetFolder(SourceFolder)
    Set subfolders = folder.subfolders
    Get_CountOfSubFoldersInsideOneSourceFolder = subfolders.count
    'release memory
        Set oFSO = Nothing: Set folder = Nothing: Set subfolders = Nothing
End Function

Sub Get_Paths_OfAllSubfolders_AllLevels_TEST()
    Dim p$, s$, v$
    v = vbCrLf: DoArrc
    p = ArrC(2)                'T0 Path of Folder "Events"
    s = Get_Paths_OfAllSubfolders_AllLevels(p)
    show CStr(anzAinB(v, s) + 1) + " SubFolders" + v + v + s
End Sub

Function Get_Paths_OfAllSubfolders_AllLevels(SourceFolder$, Optional MyCountDown As Boolean = False) As String
    'Called from    T6_Get_PidPath_OfEachFileInOneFolder
    'Status         No global variables, no Collection
    
   'Vorbereitung
        Dim F$, s$, c%, cF%, FileSystem As Object
        F = SourceFolder
        If Right(F, 1) <> "/" Then F = F + "/"
    'FileSystemObject
        Set FileSystem = CreateObject("Scripting.FileSystemObject")
    'Alle Pfade in s schreiben
        If MyCountDown Then
            MyCountDown_OneMoreMainStep
            'cF = Get_CountOfPaths_OfAllSubfolders_AllLevels(F)
            'MyCountDown_MainStepsAllowed 2: MyCountDown_SubStepsMax cF
        End If
        DoFolderRecursivGetPathsPartTwo FileSystem.GetFolder(F), s
    'Finals
        If Right(s, 2) = vbCrLf Then s = Left(s, Len(s) - 2)
        Get_Paths_OfAllSubfolders_AllLevels = s
End Function

Sub DoFolderRecursivGetPathsPartTwo(folder, s$)
    'Called from    Get_Paths_OfAllSubfolders_AllLevels, DoFolderRecursivGetPathsPartTwo
    Dim subFolder As Object
    For Each subFolder In folder.subfolders
        s = s + CStr(subFolder) + vbCrLf
        DoFolderRecursivGetPathsPartTwo subFolder, s
    Next
'    Dim File
'    For Each File In Folder.Files
'        ' Operate on each file
'    Next
End Sub

Sub ArrayÜbergeben_Part1_TEST()
    Dim Arr1(1 To 10, 1 To 10)
    Arr1(1, 1) = "Feld 1,1"
    Arr1(5, 5) = "Feld 5,5"
    Call ArrayÜbergeben_Part2_TEST(Arr1)
End Sub
Sub ArrayÜbergeben_Part2_TEST(Arr1)
    MsgBox Arr1(1, 1) & vbCrLf & Arr1(5, 5)
End Sub

Sub GetMy1DArrayFromMySub_Part1_TEST()
    Dim Arr1() As String
    Arr1 = GetMy1DArrayFromMySub_Part2_TEST()
    show Arr1(1)
End Sub
Function GetMy1DArrayFromMySub_Part2_TEST() As Variant
    Dim DeinArray() As String
    ReDim DeinArray(0 To 3)
    DeinArray(0) = "A"
    DeinArray(1) = "1"
    DeinArray(2) = "ccc"
    GetMy1DArrayFromMySub_Part2_TEST = DeinArray()
End Function

Function Get_Paths_ofAllSubfolders_OneLevel(PathOfSourceFolder$)
    Dim F, f1, fc, s, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set F = fso.GetFolder(PathOfSourceFolder)
    Set fc = F.subfolders
    For Each f1 In fc
        s = s + f1.path + vbCrLf
    Next
    If Right(s, 2) = vbCrLf Then s = Left(s, Len(s) - 2)
    Get_Paths_ofAllSubfolders_OneLevel = s
End Function

Function Get_Names_OfAllSubfolders_OneLevel(PathOfSourceFolder$)
    Dim F, f1, fc, s, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set F = fso.GetFolder(PathOfSourceFolder)
    Set fc = F.subfolders
    For Each f1 In fc
        s = s + f1.NAME + vbCrLf
    Next
    If Right(s, 2) = vbCrLf Then s = Left(s, Len(s) - 2)
    Get_Names_OfAllSubfolders_OneLevel = s
End Function

Sub CreateShortcut(p1$, N1$, p2$)
    'p1 = path of folder where the shortcut (.lnk-file) should be created
    'N1 = Name of lnk-file (that you will see inside Folder p1)
    'p2 = path of targetFolderOrFile 'DoubleClick on N1 jumps to p2-content
    If Right(p1, 1) <> "\" Then p1 = p1 + "\"
    With CreateObject("WScript.Shell")
        With .CreateShortcut(p1 + N1): .targetPath = p2: .Save: End With
    End With
End Sub

Sub MoveFile(SourceFilePath As String, DestinFilePath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.Filesystemobject")
    fso.MoveFile source:=SourceFilePath, Destination:=DestinFilePath
End Sub



'----------------------------------------------
'1 x
Sub RenameFolders_Date10CharsTo8Chars()
    'erledigt; ist 1 x durchgeführt
    Dim L10$, L8$, N$, p$, p1$, p2$, s$, T$, i%, Arr1() As String
    s = Get_Paths_ofAllSubfoldersAllLevelsAsStringUseGlobalVar("G:\Archiv Photos\- Photos ubes\")
    'Show s
    Arr1 = Split(s, vbCrLf)
    For i = 0 To UBound(Arr1)
        p1 = Arr1(i)         '1 FolderPath
        If InStr(1, p1, "\Photos Trampolin\") = 0 Then
            If p1 Like "*\####-##-## *" Then
                N = NameOfPath(p1)
                L10 = Left(N, 10)
                If L10 Like "####-##-##" Then
                    L8 = Replace(L10, "-", "")
                    p2 = Replace(p1, L10, L8)
                    RenameFolder p1, p2
                    't = t + p1 + vbCrLf + p2 + vbCrLf + vbCrLf
                End If
            End If
        End If
    Next
    'Show t
    Beep
End Sub
'----------------------------------------------

Function NameOfPath(PathOfFileOrFolder As String) As String
    Dim s As String: s = PathOfFileOrFolder
    If Right(s, 1) = "\" Then s = Left(s, Len(s) - 1)
    s = Mid(s, InStrRev(s, "\") + 1)
    NameOfPath = s
End Function

Sub Get_Paths_OfAllSubfolders_AllLevels_Collection_TEST()
    Dim s$: DoArrc
    s = Get_Paths_OfAllSubfolders_AllLevels_Collection(ArrC(3))
    show "Anzahl Ordner: " + CStr(anzAinB(":\", s)) + vbCrLf + vbCrLf + s
End Sub

Function Get_Paths_OfAllSubfolders_LikeMyString_AllLevels(PathOfSourceFolder$, LikeString$)
    Dim oFSO As Object, oFolder As Object, oFile As Object, sF
    Dim colFolders As New Collection, ws As Worksheet, p$, s$, i%
    p = PathOfSourceFolder
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(p)
    colFolders.Add oFolder             'start with this folder
    Do While colFolders.count > 0      'process all folders
        Set oFolder = colFolders(1)    'get a folder to process
        colFolders.Remove 1            'remove item at index 1
        If oFolder.path Like LikeString Then s = s + oFolder.path + vbCrLf
        'add any subfolders to the collection for processing
            For Each sF In oFolder.subfolders
                colFolders.Add sF
            Next sF
    Loop
    s = Replace(s, p + vbCrLf, "") 'Liste ohne PathOfSourceFolder
    If Right(s, 2) = vbCrLf Then s = Left(s, Len(s) - 2)
    Get_Paths_OfAllSubfolders_LikeMyString_AllLevels = s
End Function

Function Get_Paths_OfAllSubfolders_AllLevels_Collection(PathOfSourceFolder$, Optional MyCountDown As Boolean = False)
    Dim oFSO As Object, oFolder As Object, oFile As Object, sF
    Dim colFolders As New Collection, ws As Worksheet, p$, s$, i%
    p = PathOfSourceFolder
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(p)
    colFolders.Add oFolder             'start with this folder
    Do While colFolders.count > 0      'process all folders
        If MyCountDown Then MyCountDown_OneMoreSubStep
        Set oFolder = colFolders(1)    'get a folder to process
        colFolders.Remove 1            'remove item at index 1
        s = s + oFolder.path + vbCrLf
        'add any subfolders to the collection for processing
            For Each sF In oFolder.subfolders
                colFolders.Add sF
            Next sF
    Loop
    s = Replace(s, p + vbCrLf, "") 'Liste ohne PathOfSourceFolder
    If Right(s, 2) = vbCrLf Then s = Left(s, Len(s) - 2)
    Get_Paths_OfAllSubfolders_AllLevels_Collection = s
End Function

Function Get_Paths_ofAllSubfoldersAllLevelsAsStringUseGlobalVar(PathOfRootFolder$) As String
    Dim myArr, s$
    'Globale Variablen leeren
        CounterForSubFolders = 0
        Erase ArrSubFolders
    myArr = Get_Paths_OfAllSubfoldersAllLevelsAsArrayUseGlobalVar(PathOfRootFolder)
    
    s = Join(myArr, vbCrLf)
    s = Delete_EmptyEndRowsInString(s)
    Get_Paths_ofAllSubfoldersAllLevelsAsStringUseGlobalVar = s
End Function

Function Get_Paths_OfAllSubfoldersAllLevelsAsArrayUseGlobalVar(PathOfRootFolder$)
    'Public ArrSubFolders() As String
    'Public CounterForSubFolders As Long
    Dim fso As Object, fld As Object, sF As Object, myArr
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder(PathOfRootFolder)
    For Each sF In fld.subfolders
        ReDim Preserve ArrSubFolders(CounterForSubFolders)
        ArrSubFolders(CounterForSubFolders) = sF.path
        CounterForSubFolders = CounterForSubFolders + 1
        myArr = Get_Paths_OfAllSubfoldersAllLevelsAsArrayUseGlobalVar(sF.path)
    Next
    Get_Paths_OfAllSubfoldersAllLevelsAsArrayUseGlobalVar = ArrSubFolders
    Set sF = Nothing: Set fld = Nothing: Set fso = Nothing
End Function

Function anzAinB(A As String, B As String) As Long
    Dim s As String
    s = Replace(B, A, "")
    anzAinB = (Len(B) - Len(s)) / Len(A)
End Function

Function Get_NrOfLastRowInColumnNr(ColNr As Long, Optional NameOfSheet$ = "NoSheet") As Long
    'Called from:   xxx
    'Action:        xxx
    
    If NameOfSheet = "NoSheet" Then
        Get_NrOfLastRowInColumnNr = ActiveSheet.Cells(Rows.count, ColNr).End(xlUp).Row
    Else
        Get_NrOfLastRowInColumnNr = ThisWorkbook.Sheets(NameOfSheet).Cells(Rows.count, ColNr).End(xlUp).Row
    End If
End Function

Sub FillArrC(Nr%, Wert$)
    'Called from:   Fill_ArrC_Some, ...
    'Action:        Führt den Eintrag in ArrC und in die T8-Tabelle durch
    
    'Vorbereitung
        Dim z1%, s1%, r As Range
        DoArrc
    'Einstellungen
            'Cells(z1, s1) = "1" = erste dunkelgrüne Zelle; IndexNr 1
            s1 = 3    's1 = SpaltenNummer der IndexSpalte (ArrC-Nummern)
            z1 = 7    'z1 = ZeilenNummer  der 1. dunkelgrünen Zelle
    'Eintrag in Array:
        ArrC(Nr) = Wert
    'Eintrag in T8
        EE 0 'verhindert Event Worksheet_Change
        Set r = ThisWorkbook.Sheets("T8").Cells(z1 - 1 + Nr, s1 + 1): r.Value = Wert
        If Len(Wert) < 9 Then r.HorizontalAlignment = xlCenter Else r.HorizontalAlignment = xlLeft
        EE 1
End Sub

Function Get_T1Cell_HoldingMyText(MyText$) As String
    'Called from:   Worksheet_SelectionChange [in T1]
    'Status         nur für Sheets("T1")
    'Action:        liefert die T1-Zelle, welche MyText enthält
    '               als "123|45" [ZeilenNr|SpaltenNr]
    
    Dim r1 As Range
    With ThisWorkbook.Sheets("T1")
        Set r1 = .Cells.Find(What:=MyText, _
            After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End With
    If r1 Is Nothing Then
        'MyText nicht zu finden; ggf. ausgeblendet
        Get_T1Cell_HoldingMyText = ""
    Else
        Get_T1Cell_HoldingMyText = CStr(r1.Row) + "|" + CStr(r1.Column)
    End If
End Function

Function Get_T2Cell_HoldingMyText(MyText$) As String
    'Called from:   Worksheet_SelectionChange [in T1]
    'Status         nur für Sheets("T1")
    'Action:        liefert die T1-Zelle, welche MyText enthält
    '               als "123|45" [ZeilenNr|SpaltenNr]
    
    Dim r1 As Range
    With ThisWorkbook.Sheets("T2")
        Set r1 = .Cells.Find(What:=MyText, _
            After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End With
    If r1 Is Nothing Then
        'MyText nicht zu finden; ggf. ausgeblendet
        Get_T2Cell_HoldingMyText = ""
    Else
        Get_T2Cell_HoldingMyText = CStr(r1.Row) + "|" + CStr(r1.Column)
    End If
End Function

Function Get_RowNr_HoldingMyTextWhole(MySheet$, MyText$) As Integer
    'Called from:   xxx
    'Action:        xxx
    
    Dim r1 As Range
    With ThisWorkbook.Sheets(MySheet)
        Set r1 = .Cells.Find(What:=MyText, _
            After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End With
    If r1 Is Nothing Then
        'MyText nicht zu finden; ggf. ausgeblendet
        Get_RowNr_HoldingMyTextWhole = 0
    Else
        Get_RowNr_HoldingMyTextWhole = CInt(r1.Row)
    End If
End Function

Function Get_ColumnNr_HoldingMyTextWhole(MySheet$, MyText$) As Integer
    'Called from:   xxx
    'Action:        xxx
    
    Dim r1 As Range
    With ThisWorkbook.Sheets(MySheet)
        Set r1 = .Cells.Find(What:=MyText, _
            After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End With
    If r1 Is Nothing Then
        'MyText nicht zu finden; ggf. ausgeblendet
        Get_ColumnNr_HoldingMyTextWhole = 0
    Else
        Get_ColumnNr_HoldingMyTextWhole = CInt(r1.Column)
    End If
End Function

Function Get_CellName_HoldingMyText(MySheet$, MyText$) As String
    'Called from:   xxx
    'Action:        liefert die Zelle, welche MyText enthält; "D19"
    
    Dim s$, r1 As Range
    With ThisWorkbook.Sheets(MySheet)
        Set r1 = .Cells.Find(What:=MyText, _
            After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End With
    If r1 Is Nothing Then
        'MyText nicht zu finden; ggf. ausgeblendet
        s = ""
    Else
        s = Replace(CStr(r1.Address), "$", "")
    End If
    Get_CellName_HoldingMyText = s
End Function

Function Get_CellRowCol_HoldingMyText(MySheet$, MyText$) As String
    'Called from:   Worksheet_SelectionChange [in T1]
    'Action:        liefert die Zelle, welche MyText enthält
    '               als "123|45" [ZeilenNr|SpaltenNr]
    
    Dim r1 As Range
    With ThisWorkbook.Sheets(MySheet)
        Set r1 = .Cells.Find(What:=MyText, _
            After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End With
    If r1 Is Nothing Then
        'MyText nicht zu finden; ggf. ausgeblendet
        Get_CellRowCol_HoldingMyText = ""
    Else
        Get_CellRowCol_HoldingMyText = CStr(r1.Row) + "|" + CStr(r1.Column)
    End If
End Function

Function Get_T1Column_HoldingMyText(MyText$) As Integer
    'Called from:   xxx
    'Action:        xxx
    
    Dim r1 As Range
    With ThisWorkbook.Sheets("T1")
        Set r1 = .Cells.Find(What:=MyText, _
            After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End With
    If r1 Is Nothing Then
        'MyText nicht zu finden; ggf. ausgeblendet
        Get_T1Column_HoldingMyText = 0
    Else
        Get_T1Column_HoldingMyText = CInt(r1.Column)
    End If
End Function

Function Get_T1Row_HoldingMyText(MyText$) As Integer
    'Called from:   xxx
    'Action:        xxx
    
    Dim r1 As Range
    With ThisWorkbook.Sheets("T1")
        Set r1 = .Cells.Find(What:=MyText, _
            After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    End With
    If r1 Is Nothing Then
        'MyText nicht zu finden; ggf. ausgeblendet
        Get_T1Row_HoldingMyText = 0
    Else
        Get_T1Row_HoldingMyText = CInt(r1.Row)
    End If
End Function



