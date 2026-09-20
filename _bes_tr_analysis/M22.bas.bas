Attribute VB_Name = "M22"
Option Explicit
'----------------------------------------------------------------
'API for Load_PathsOfAllSubfoldersAllLevels
    Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
    End Type
    Private Type WIN32_FIND_DATAW
        dwFileAttributes As Long
        ftCreationTime   As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime  As FILETIME
        nFileSizeHigh    As Long
        nFileSizeLow     As Long
        dwReserved0      As Long
        dwReserved1      As Long
        cFileName(0 To 519) As Byte        ' 260 WCHARs = 520 Bytes
        cAlternate(0 To 27) As Byte        ' 14 WCHARs = 28 Bytes
    End Type
    Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
    Private Const INVALID_HANDLE_VALUE     As Long = -1
    Private Declare PtrSafe Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByRef lpFindFileData As WIN32_FIND_DATAW) As LongPtr
    Private Declare PtrSafe Function FindNextFileW Lib "kernel32" (ByVal hFindFile As LongPtr, ByRef lpFindFileData As WIN32_FIND_DATAW) As LongPtr
    Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As Long
'----------------------------------------------------------------

Public Sub Load_PathsOfAllSubfoldersAllLevels(pSourceFolders$, Arr() As String, Optional NoFolderLike$ = "")
    Dim currentPath$, capacity&, count&, i&, basePaths() As String
    If Trim$(pSourceFolders) = "" Then Erase Arr: Exit Sub
    
    basePaths = Split(pSourceFolders, "|")
    capacity = 1000: count = 0: ReDim Arr(1 To capacity)
    
    For i = LBound(basePaths) To UBound(basePaths)
        currentPath = Trim$(basePaths(i))
        If Right$(currentPath, 1) = "\" Then currentPath = Left$(currentPath, Len(currentPath) - 1)
        If Len(currentPath) > 0 Then
            ' Prüfung für Basisordner
            If NoFolderLike = "" Or Not (LCase$(currentPath) Like LCase$(NoFolderLike)) Then
                count = count + 1
                If count > capacity Then
                    capacity = capacity + 1000
                    ReDim Preserve Arr(1 To capacity)
                End If
                Arr(count) = currentPath
                ScanSubfoldersAPI currentPath, Arr, count, capacity, NoFolderLike
            End If
        End If
    Next i
    
    If count > 0 Then
        ReDim Preserve Arr(1 To count)
    Else
        Erase Arr
    End If
End Sub

'----------------------------------------------------------------
' Rekursive Hilfsprozedur (API-Scan)
'----------------------------------------------------------------
Private Sub ScanSubfoldersAPI(ByVal sFolderPath$, ByRef Arr() As String, ByRef count&, ByRef capacity&, ByVal NoFolderLike$)
    #If VBA7 Then
        Dim hFind As LongPtr
    #Else
        Dim hFind As Long
    #End If
    
    Dim wfd As WIN32_FIND_DATAW
    Dim sName$, sFullPath$, searchPath$
    Dim nullPos As Long
    
    ' Vorbereitung des Pfads für lange Pfade (> 260 Zeichen)
    searchPath = "\\?\" & sFolderPath & "\*.*"
    
    hFind = FindFirstFileW(StrPtr(searchPath), wfd)
    
    If hFind <> INVALID_HANDLE_VALUE Then
        Do
            ' Byte-Array in String umwandeln
            sName = wfd.cFileName
            
            ' Null-Terminator im Unicode-String suchen
            nullPos = InStr(1, sName, vbNullChar, vbBinaryCompare)
            If nullPos > 1 Then
                sName = Left$(sName, nullPos - 1)
            Else
                sName = ""
            End If
            
            ' Systemverzeichnisse "." und ".." ignorieren
            If sName <> "." And sName <> ".." And Len(sName) > 0 Then
                ' Prüfen, ob es sich um einen Ordner handelt
                If (wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                    sFullPath = sFolderPath & "\" & sName
                    
                    ' Filter-Prüfung (NoFolderLike)
                    If NoFolderLike = "" Or Not (LCase$(sFullPath) Like LCase$(NoFolderLike)) Then
                        count = count + 1
                        
                        ' Array bei Bedarf vergrößern
                        If count > capacity Then
                            capacity = capacity + 1000
                            ReDim Preserve Arr(1 To capacity)
                        End If
                        
                        Arr(count) = sFullPath
                        
                        ' Rekursiver Aufruf für den Unterordner
                        ScanSubfoldersAPI sFullPath, Arr, count, capacity, NoFolderLike
                    End If
                End If
            End If
        Loop While FindNextFileW(hFind, wfd) <> 0
        
        FindClose hFind
    End If
End Sub

Sub T2_TextFeld2_Create(z%)
    'Called from    T2_BadLinks
    'z              = T2-ZeilenNr mit z.Top=Textfeld.Top
    
    'Vorbereitung
        Dim ws As Worksheet, shp As shape, tf As TextFrame2
        Dim NameOfTextBox$, T$, txt$, v$, L!, Tp!, w!, H!
        Set ws = Sheets("T2"): T = vbTab: v = vbCrLf
        NameOfTextBox = "T2_Textfeld2"
    'Einstellungen
        L = ws.Cells(1, 27).Left: Tp = ws.Cells(z, 1).Top
        w = ws.Cells(1, 35).Left - L: H = ws.Cells(z + 7, 1).Top - Tp
    'Falls das Textfeld bereits existiert, vorher löschen (vermeidet Duplikate)
        On Error Resume Next: ws.Shapes(NameOfTextBox).Delete: On Error GoTo 0
    'Textbox erstellen
        Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, L, Tp, w, H)
        shp.NAME = NameOfTextBox: Set tf = shp.TextFrame2
    'Innenabstände
        tf.MarginLeft = 2: tf.MarginRight = 2: tf.MarginTop = 2: tf.MarginBottom = 2
    'Tabstopps setzen
        With tf.TextRange.ParagraphFormat.TabStops
        'Vorhandene Standard-Tabs entfernen
            Do While .count > 0: .item(1).Clear: Loop
        'Tabstopp 1: Rechtsbündig bei 70 pt
            .Add Position:=85, Type:=msoTabStopRight
        'Tabstopp 2: Rechtsbündig
            .Add Position:=w - 2, Type:=msoTabStopRight
        End With
    'Text-Inhalt
        txt = "Änderungen" + v + v + "Bad Links" + T + "found:" + T + "0" + v _
              + T + "deleted:" + T + "0" + v + T + "repaired:" + T + "0" + v _
              + T + "still bad:" + T + "0" + v + v + "Good Links" + T + "total:" + T + "0"
    'Text zuweisen und formatieren (Consolas 8pt)
        With tf.TextRange: .text = txt: .Font.NAME = "Consolas": .Font.size = 8: End With
End Sub

Sub T2_ChangeTextBox2(Zeile&, NeuerWert As Variant)
    'Vorbereitung
        Dim shp As shape, tf As TextFrame2, targetPara As TextRange2
        Dim currentText$, labelText$, tabPos&
        On Error GoTo ErrorHandler
        Set shp = ActiveSheet.Shapes("T2_Textfeld2")
        Set tf = shp.TextFrame2
    'Prüfung, ob die angeforderte Zeile existiert
        If Zeile > tf.TextRange.Paragraphs.count Then Exit Sub
    'Referenz auf die spezifische Zeile (Paragraph)
        Set targetPara = tf.TextRange.Paragraphs(Zeile)
        currentText = targetPara.text
    'Suche den 2. Tabulator in dieser Zeile
        tabPos = InStr(InStr(1, currentText, vbTab) + 1, currentText, vbTab)
    
    If tabPos > 0 Then
        'Extrahiere alles bis einschließlich zum 2. Tab (z. B. " [Tab] Bad Links total [Tab]")
            labelText = Left(currentText, tabPos)
        'Ersetze den Text der Zeile (Label + neuer Wert + Zeilenumbruch falls vorhanden)
            If Right(currentText, 1) = vbCrLf Or Right(currentText, 1) = vbCr Then
                targetPara.text = labelText & NeuerWert & vbCr
            Else
                targetPara.text = labelText & NeuerWert
            End If
    End If
    Exit Sub
ErrorHandler:
    ' Baut Absätze auf, falls die Zeile noch nicht existierte
End Sub

Sub Fill_Nn_Vn_FromPersLink(NameOrPathOfLink$, Nn$, Vn$)
    'Called from    T2_BadLink_SameName, ...
    
    'Vorbereitung
        Dim N$, c%
    
    'N = NameOfLink                         'N = "May, Lea.lnk", "May, Lea (Bonn).lnk"
        If NameOrPathOfLink Like "*\*" Then N = NameOfPath(NameOrPathOfLink) Else N = NameOrPathOfLink
        N = Replace(N, ".lnk", "")          'N = "May, Lea", "May, Lea (Bonn)"
        c = InStr(1, N, " (")
        If c > 0 Then N = Left(N, c - 1)    'N = "May, Lea"
        c = InStr(1, N, ", ")
    'Fill Nn, Vn
        If c = 0 Then Nn = "": Vn = "" Else Nn = Left(N, c - 1): Vn = Mid(N, c + 2)
End Sub

Sub T2_BadLink_SameName(A() As String, PathOfLink$, a4&, a6&, Report$)
    'Called from    T2_BadLinks
    'SameName       Ggf. existiert noch ein älterer Link (= BadLink) zum gleichen Namen,
    '               z. B. "May, Lea.lnk" und "May, Lea (Bonn).lnk"
    'A()            holds all LinkFilesPaths of actual Folder
    'PathOfLink     = BadLink = eine der A()-Belegungen

    'Vorbereitung
        If PathOfLink Like "*\Leute\*" Then Exit Sub      'SprungZiel ist kein PersFolder
        Dim N$, NameOfLink$, Nn$, v$, Vn$, c&, i&: v = vbCrLf
    'Link = PersonLink?
        NameOfLink = NameOfPath(PathOfLink)                  'NameOfLink = "May, Lea (Bonn).lnk"
        Fill_Nn_Vn_FromPersLink NameOfLink, Nn, Vn
        If Nn = "" Then
            Report = Report + "   Link ist kein PersonLink" + v
            Exit Sub      'kein PersonName
        Else
            N = Nn + ", " + Vn
        End If
    'N = pure Name NnVn
    'Suche nach SameName
        For i = LBound(A) To UBound(A)
            If A(i) Like "*" + N + "*" Then
                If A(i) <> PathOfLink Then
                    Report = Report + "   Ein weiterer Link gleichen Namens existiert: " _
                    + NameOfPath(A(i)) + v _
                    + "   BadLink '" + NameOfLink + "' wird gelöscht" + v
                    DeleteFile PathOfLink
                    a4 = a4 + 1: T2_ChangeTextBox2 4, a4 'Zeile 4,     deleted
                    a6 = a6 - 1: T2_ChangeTextBox2 6, a6 'Zeile 6,   still bad
                    Exit Sub
                End If
            End If
        Next
        Report = Report + "   GoodLink gleichen Namens existiert nicht" + v
End Sub

Sub T2_Handle_CountOfEventsClubsNationsLeuteFolders(B(), F() As String, Report$)
    'Called from    T2_BadLinks
    Dim a11&, a12&, a13&, i&
    With Sheets("T2")
    'Anzahl ermitteln   'alle Events-, ClubsNations-, Leute-Ordner
        For i = LBound(F) To UBound(F)
            If F(i) Like "*\Events\*" Then a11 = a11 + 1        'a11 =  317 (+ "*\Events" --> 318)
            If F(i) Like "*\ClubsNations\*" Then a12 = a12 + 1  'a12 =  197       ClubsNations
            If F(i) Like "*\Leute\*" Then a13 = a13 + 1         'a13 = 2521       Leute
        Next
    'T2-Anzeige                                                             'Anzahl ...
        If a11 > 99 Then .Cells(B(4, 1) + 1, 10) = CStr(a11 + 1) + " Ordner"  'Event-Ordner
        If a12 > 99 Then .Cells(B(5, 1) + 1, 10) = CStr(a12 + 1) + " Ordner"  'ClubsNations
        If a13 > 99 Then .Cells(B(6, 1) + 1, 10) = CStr(a13 + 1) + " Ordner"  'Leute-Ordner
        Report = Report + " " + CStr(UBound(F)) + " Ordner sollen durchsucht werden" + vbCrLf + String(100, "-") + vbCrLf
    End With
End Sub


