Attribute VB_Name = "M01"
' Module -- Typ: Code Module -- Name: M01

Option Explicit 'M01

Sub zzz_M01()
    
    showProcs "each"
    EE 1: Beep
End Sub

Sub TesteFTasten()
    Dim i As Long
    ' F1 bis F12 mit der Test-Prozedur verknüpfen
    For i = 1 To 12
        Application.OnKey "{F" & i & "}", "'FTasteGedrueckt " & i & "'"
    Next i
    
    MsgBox "F1 bis F12 wurden neu belegt!" & vbCrLf & _
           "Drücke jetzt nacheinander deine F-Tasten,", vbInformation, "Test aktiv"
End Sub

' Wird aufgerufen, wenn eine funktionierende F-Taste gedrückt wird
Sub FTasteGedrueckt(ByVal tasteNr As Integer)
    MsgBox "Taste F" & tasteNr & " ist FREI und funktioniert!", vbInformation, "Erfolg"
End Sub

Sub FTastenZuruecksetzen()
    Dim i As Long
    ' Standardbelegung aller F-Tasten wiederherstellen
    For i = 1 To 12
        Application.OnKey "{F" & i & "}"
    Next i
    MsgBox "Alle F-Tasten sind wieder im Originalzustand.", vbInformation, "Zurückgesetzt"
End Sub

Sub RunShowProcs()
    Dim s As String
    ' 1. Eingabedialog anzeigen
    s = InputBox("Suche Prozedurnamen mit:", "Texteingabe (F3)")
    If s = "" Then Exit Sub
    showProcs s
End Sub


Sub Load_AllVideos_Events_ClubsNations(Arr() As String)
    'enthält alle vIDs; können als Sprungziel für Links dienen
    Dim Like1$, Like2$, v1$, v2$
    DoArrc
    Like1 = "*v####-##*": Like2 = "*.jpg*"
    v1 = Get_FilePaths_Like_NotLike_insideSourceFolderAndSubFolders(ArrC(3), Like1, Like2)
    v2 = Get_FilePaths_Like_NotLike_insideSourceFolderAndSubFolders(ArrC(11), Like1, Like2)
    Arr = Split(v1 + v2, vbCrLf)
End Sub

Sub CountDownLabel_Create()
    Dim ws As Worksheet, shp As shape, visibleLeft As Double, visibleTop As Double
    Set ws = ActiveSheet
    'Aktuelle Scroll-Position (damit es da landet, wo du gerade hinschaust)
        visibleLeft = ActiveWindow.VisibleRange.Left: visibleTop = ActiveWindow.VisibleRange.Top
    'Altes Label löschen falls vorhanden
        On Error Resume Next
        ws.Shapes("CountDownLabel").Delete
        On Error GoTo 0
    'Shape an der AKTUELLEN Bildschirmposition erstellen
        Set shp = ws.Shapes.AddShape(msoShapeRectangle, visibleLeft, visibleTop, 150, 80)
    'Eigenschaften des Labels
        With shp: .NAME = "CountDownLabel": .Line.Weight = 2
        .Fill.ForeColor.RGB = RGB(240, 240, 240): .Line.ForeColor.RGB = RGB(100, 100, 100)
        With .TextFrame2: .TextRange.text = "9999": .TextRange.Font.size = 48
        .TextRange.Font.Bold = True: .TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter: .VerticalAnchor = msoAnchorMiddle
    End With: End With: DoEvents
End Sub

Sub CountDownLabel_Update(ByVal NeuerWert As String)
    On Error Resume Next
    ActiveSheet.Shapes("CountDownLabel").TextFrame2.TextRange.text = NeuerWert
    'Shape an der AKTUELLEN Bildschirmposition erstellen
        Dim ws As Worksheet, shp As shape, visibleLeft As Double, visibleTop As Double
        Set ws = ActiveSheet
        visibleLeft = ActiveWindow.VisibleRange.Left
        visibleTop = ActiveWindow.VisibleRange.Top
        Set shp = ws.Shapes.AddShape(msoShapeRectangle, visibleLeft, visibleTop, 150, 80)
    DoEvents: On Error GoTo 0
End Sub

Sub CountDownLabel_Delete()
    Dim ws As Worksheet: Set ws = ActiveSheet
    On Error Resume Next: ws.Shapes("CountDownLabel").Delete: On Error GoTo 0
End Sub

Sub CountFolders_TEST()
    Dim p$, s$: p = "F:\Archiv Trampolin 1900-1999"
    s = CountFolders(p)
    show s
End Sub

Function CountFolders(ByVal PathOfRootFolder As String) As Long
    'Hauptfunktion: Gibt die Gesamtanzahl aller Ordner und Unterordner zurück
    Dim fso As Object, startFolder As Object, totalFolders As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Prüfen, ob der Startordner überhaupt existiert
    If Not fso.FolderExists(PathOfRootFolder) Then
        CountFolders = -1 ' Rückgabewert -1 signalisiert einen ungültigen Pfad
        Exit Function
    End If
    
    Set startFolder = fso.GetFolder(PathOfRootFolder)
    totalFolders = 0
    ' Starte die rekursive Zählung
    Call CountFolders_CountSubFolders(startFolder, totalFolders)
    ' Ergebnis an die Funktion übergeben
    CountFolders = totalFolders
End Function

Sub CountFolders_CountSubFolders(ByRef currentFolder As Object, ByRef folderCount As Long)
    'Hilfs-Sub für die Rekursion (zählt im Hintergrund hoch)
    Dim subFolder As Object
    'Jeden Unterordner im aktuellen Verzeichnis durchlaufen
    For Each subFolder In currentFolder.subfolders
        folderCount = folderCount + 1 ' Zähler erhöhen
        'Rekursiver Aufruf: Sucht in diesem Unterordner weiter nach Ordnern
        Call CountFolders_CountSubFolders(subFolder, folderCount)
    Next subFolder
End Sub

Sub Repair_NonVbsPersonLinks_in_ClubsNations()
    Dim p$, s$, v$, i&, A() As String
    v = vbCrLf: DoArrc: p = ArrC(3)
    s = Get_NonVbsLinks_insideFolder(p)
    A = Split(s, v)
    For i = 0 To UBound(A)
        Create_OneVbsLink_Person A(i)        'one Old Link
    Next
    Beep
End Sub

Sub Show_NonVbsLinks_in_Leute()
    Dim p$, s$, v$, Anz&, CountDown&
    v = vbCrLf: DoArrc: CountDownLabel_Create: p = ArrC(4)
    s = Get_NonVbsLinks_insideFolder(p)
    CountDown = CountFolders(p): Anz = anzAinB(v, s): CountDownLabel_Delete
    show CStr(Anz) + " NonVbsLinks in " + CStr(CountDown) + " Subfolders of '" + p + "'" + v + v + s
End Sub

Sub Show_NonVbsLinks_in_Archiv()
    Dim p$, s$, v$, Anz&, CountDown&
    v = vbCrLf: DoArrc: CountDownLabel_Create: p = ArrC(2)
    s = Get_NonVbsLinks_insideFolder(p)
    CountDown = CountFolders(p): Anz = anzAinB(v, s): CountDownLabel_Delete
    show CStr(Anz) + " NonVbsLinks in " + CStr(CountDown) + " Subfolders of '" + p + "'" + v + v + s
End Sub

Sub Show_NonVbsLinks_in_ClubsNations()
    Dim p$, s$, v$, Anz&, CountDown&
    v = vbCrLf: DoArrc: CountDownLabel_Create: p = ArrC(11)
    s = Get_NonVbsLinks_insideFolder(p)
    CountDown = CountFolders(p): Anz = anzAinB(v, s): CountDownLabel_Delete
    show CStr(Anz) + " NonVbsLinks in " + CStr(CountDown) + " Subfolders of '" + p + "'" + v + v + s
End Sub

Sub Show_NonVbsLinks_in_Events()
    Dim p$, s$, v$, Anz&, CountDown&
    v = vbCrLf: DoArrc: CountDownLabel_Create: p = ArrC(3)
    s = Get_NonVbsLinks_insideFolder(p)
    CountDown = CountFolders(p): Anz = anzAinB(v, s): CountDownLabel_Delete
    show CStr(Anz) + " NonVbsLinks in " + CStr(CountDown) + " Subfolders of '" + p + "'" + v + v + s
End Sub

Function Get_NonVbsLinks_insideFolder(PathOfRootFolder$) As String
    Dim NonVbsLinks$, fso As Object, startFolder As Object, CountDown&
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Prüfen, ob der Startordner existiert
        If Not fso.FolderExists(PathOfRootFolder) Then
            MsgBox "Der angegebene Startordner wurde nicht gefunden!", vbCritical, "Fehler"
            Exit Function
        End If
        Set startFolder = fso.GetFolder(PathOfRootFolder)
        CountDown = CountFolders(PathOfRootFolder)
    'Starte die rekursive Suche und übergib die lokale Variable
        Call Get_NonVbsLinks_insideFolder_ScanFolder(startFolder, NonVbsLinks, CountDown)
    'Finals
        Get_NonVbsLinks_insideFolder = NonVbsLinks
End Function

Sub Get_NonVbsLinks_insideFolder_ScanFolder(ByRef currentFolder As Object, ByRef NonVbsLinks$, CountDown&)
    'Called from    Get_NonVbsLinks_insideFolder
    
    'Vorbereitung
    Dim fso As Object, subFolder As Object, file As Object, wshShell As Object, shortcut As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wshShell = CreateObject("WScript.Shell")
    
    ' 1. Alle Dateien im aktuellen Ordner prüfen
    For Each file In currentFolder.Files
        ' Prüfen auf .lnk-Datei
        If LCase(fso.GetExtensionName(file.path)) = "lnk" Then
            On Error Resume Next ' Falls eine Verknüpfung fehlerhaft/gesperrt ist
            Set shortcut = wshShell.CreateShortcut(file.path)
            
            ' Prüfen, ob die Beschreibung NICHT "VbsLink" ist
            If shortcut.Description <> "VbsLink" Then
                ' Pfad an den bestehenden String anhängen (mit Zeilenumbruch getrennt)
                If NonVbsLinks = "" Then
                    NonVbsLinks = file.path
                Else
                    NonVbsLinks = NonVbsLinks & vbCrLf & file.path
                End If
            End If
            On Error GoTo 0
        End If
    Next file
    
    ' 2. Rekursion: Unterordner durchlaufen
    For Each subFolder In currentFolder.subfolders
        'CountDown = CountDown - 1
        'CountDownLabel_Update CountDown
        Call Get_NonVbsLinks_insideFolder_ScanFolder(subFolder, NonVbsLinks, CountDown)
    Next subFolder
End Sub

Sub ShapePopupAndGo()
    ' Popup is a Shape with a text
    Worksheets("T4").Shapes("Popup").Visible = msoCTrue
    Application.OnTime Now + TimeSerial(0, 0, 5), "ShapePopupHide"
End Sub

Sub ShapePopupHide()
    Worksheets("Sheet1").Shapes("Popup").Visible = msoFalse
End Sub

Function Get_FolderPaths_OfLiga() As String
    Dim L1$, L2$, i%, Arr1() As String
    DoArrc
    L1 = Get_AllSubfolderPaths_LikeMyString_OneLevel(ArrC(3), "*Liga*") 'Level 1
    Arr1 = Split(L1, vbCrLf)
    For i = 0 To UBound(Arr1)
        L2 = L2 + Get_Paths_ofAllSubfolders_OneLevel(Arr1(i)) + vbCrLf
    Next
    'show L2
    Get_FolderPaths_OfLiga = L2
End Function

Function Get_AllSubfolderPaths_LikeMyString_OneLevel(PathOfSourceFolder$, MyString$)
    Dim F, f1, fc, s, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set F = fso.GetFolder(PathOfSourceFolder)
    Set fc = F.subfolders
    For Each f1 In fc
        If f1.path Like MyString Then s = s + f1.path + vbCrLf
    Next
    If Right(s, 2) = vbCrLf Then s = Left(s, Len(s) - 2)
    Get_AllSubfolderPaths_LikeMyString_OneLevel = s
End Function

Function Get_FirstSubfolderPath_LikeMyString_OneLevel(PathOfSourceFolder$, MyString$)
    Dim F, f1, fc, s, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set F = fso.GetFolder(PathOfSourceFolder)
    Set fc = F.subfolders
    For Each f1 In fc
        If f1.path Like MyString Then s = f1.path: Exit For
    Next
    Get_FirstSubfolderPath_LikeMyString_OneLevel = s
End Function

Function IsInArray(stringToBeFound As Integer, Arr As Variant) As Boolean
    'Called from    [None]
    '?              Ggf. nützlich?
    IsInArray = (UBound(Filter(Arr, stringToBeFound)) > -1)
    'Test also:
        'strSubNames = Filter(strNames, "smith")                    'CaseSensitive
        'strSubNames = Filter(strNames, "smith", , vbTextCompare)   'Not CaseSensitive
        
        
    'Dim arr As Variant, filterArr As Variant, aVal As Variant
    'arr = Array("Dragon", "Dog", "DRAGONfly", "Cat", "fly")
    'filterArr = Filter(arr, "Dragon")
    ''Print the contents of the filtered Array
    'For Each aVal In filterArr
    '    Debug.Print aVal
    'Next aVal
    ''Result: Dragon
    
    'arr = Array("Dragon", "Dog", "DRAGONfly", "Cat", "fly")
    'filterArr = Filter(arr, "Dragon", Compare:=vbTextCompare) 'Result: "Dragon", "DRAGONfly"
End Function

Function Get_NameOfFile(ByVal PathOfFolder$) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Get_NameOfGet_NameOfFileFolder = fso.GetFileName(PathOfFolder)
End Function

Function Get_NameOfFolder(ByVal PathOfFolder$) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Get_NameOfFolder = fso.GetFileName(PathOfFolder)
End Function

