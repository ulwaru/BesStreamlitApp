Attribute VB_Name = "M62"
Option Explicit 'M62
Private PathsOfSomeFolders As String

Sub zzz_M62()
    showProcs "brac"
    
    'RenameModule "Modul1", "M62"
    EE 1: Beep
End Sub

'------------ START [EventFileNames: Add ' - ' ...]

Sub T6_BMB_EventFileNames() 'T6_BMB_Events_BlankMinusBlankToEventFileNames
    'Called from    [T6-CellButton " EventFileNames: Add ' - ' ..."]
    'Privat         M62-Privat-Variable: PathsOfSomeFolders
    'Buttons        Ja - Nein - Abbrechen - Open EventFolder
    'Action         Vorbereitung einer AnzeigeListe mit Header und Buttons (ab T6-Spalte 37)
    
    'Vorbereitung
        Dim p$, s$, zT6%
        With Sheets("T6"): DoArrc: FillArrC 83, "1": Ticks1
    'MyCountDown Init
        'zT6 = T6-ZeilenNr, die den Cell-Button enthält (für T6_DoDoneRemarks-Einträge)
        zT6 = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 2, 4, " EventFileNames: Add ' - ' ...")
        MyCountDown_Init "|T6|" + CStr(zT6) + "|19|0|" '= |SheetName|z|s|StartNr| of CountDownShowCell
    'Get paths of all EventFolders (no SubFolder Dng etc.) and all ClubsNationsFolders
        PathsOfSomeFolders = Get_PathsOfAllDateEventFolders '+ vbCrLf + Get_Paths_OfAllSubfolders_AllLevels_Collection(ArrC(11))
        .Cells(zT6, 20) = anzAinB(":", PathsOfSomeFolders)
        p = Get_FirstLineOfStringX_AndDeleteFirstLine(PathsOfSomeFolders)
    'Format T6-DataArea [EventFileNames: Add ' - ' ...]
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 11 'Spalte, die ganz links zu sehen sein soll
        ActiveWindow.Zoom = 100
        T6_BMB_Format1DataArea
    'Check
        '.Cells(8, 42) = anzAinB(":", PathsOfSomeFolders) 'CountDown
        .Cells(5, 37) = " FileNames der EventOrdner werden durchsucht:      " + CStr(anzAinB(":", PathsOfSomeFolders))
        T6_Check_FileNamesInsideOneEventFolder p
    'Finals
        End With
End Sub

Sub T6_Check_FileNamesInsideOneEventFolder(PathOfEventFolder$)
    'Called from    T6_BMB_EventFileNames
    'Privat         M62-Privat-Variable: PathsOfSomeFolders
    'Action         xxx
    
    'Vorbereitung
        Dim A$, B$, f1$, f2$, NameOfEventFolder$, p$, sc$, T$, v$
        Dim c%, C1%, C2%, C3%, C4%, C5%, C6%, i%, zT6%, rg As Range
        Dim Arr1() As String, p1() As String, p2() As String, T6()
        v = vbCrLf: p = PathOfEventFolder: C3 = 0: With Sheets("T6")
        FillArrC 84, p
    'CountDown  (Anzahl Zeilen, die noch in 'PathsOfSomeFolders' stehen)
        c = anzAinB(":", PathsOfSomeFolders)
        If ArrC(83) = "1" Then   '"1" = EventFileNamesSearchIsRunning
            .Cells(5, 37) = " FileNames der EventOrdner werden durchsucht:      " + CStr(c)
        Else
            .Cells(8, 42) = c 'CountDownShowCell
        End If
    'Clear old content
        C1 = Get_NrOfLastRowInColumnNr(37, "T6")
        If C1 < 20 Then C1 = 20
        .Range(.Cells(11, 37), .Cells(C1, 43)).ClearContents
        .Range(.Cells(11, 37), .Cells(C1, 43)).Interior.ColorIndex = xlNone
    'List FileNames inside Eventfolder
        If p = "" Then
            'Alle Eventfolders sind verarbeitet
            .Cells(8, 41).Select 'CellButton 'Abbrechen'
            zT6 = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 2, 4, " EventFileNames: Add ' - ' ...")
            FillArrC 24, CStr(.Cells(zT6, 20)) 'CountOfChanges
            FillArrC 30, CStr(zT6) 'CountDownShowCell-ZeilenNr
            T6_DoDoneRemarks
            Exit Sub
        End If
    'Status: Es existiert ein PathOfEventFolder in p
    '   von diesem Ordner p werden alle files eingelesen, die mit "19" oder "20" beginnen
        f1 = Get_AllFileNames_Like_OfOneFolder(p, "[12][90]*")
        'If f1 = "" Then Stop: 'Keine Dateien im aktuellen EventFolder
    'Name of Eventfolder
        NameOfEventFolder = NameOfPath(p)
    'StartChars of FileNames
        sc = NameOfEventFolder          'sc = StartChars
        C1 = InStr(1, sc, " ")
        sc = T9_Get_Date8FromEventFolderName(sc) + Mid(sc, C1) + " -"
    'Status: Name of actual EventFolder:    "1959-07-09 Turnfest Basel_CH"
        '    sc =                           "19590709 Turnfest Basel_CH -"
        '    f1 = all FileNames of actual EventFolder (mit "19..." oder "20...")
    'FilNames mit " ." herausnehmen     '19870000 .DM Bonn #result ...
        Arr1 = Split(f1, v)
        For i = 0 To UBound(Arr1)
            'Schleife über alle FileNames des aktuellen EventFolders
            If InStr(1, Arr1(i), " .") = 0 Then f2 = f2 + Arr1(i) + v
        Next
        f2 = Delete_EmptyEndRowsInString(f2)
        '  = all FileNames of actual EventFolder (mit "19...", ohne " .")
        If f2 = "" Then
            'Keine DateiNamen zu bearbeiten --> Nächster EventFolder
            p = Get_FirstLineOfStringX_AndDeleteFirstLine(PathsOfSomeFolders)
            T6_Check_FileNamesInsideOneEventFolder p
            Exit Sub
        End If
    'Sind alle f2-FileNames bereits korrigiert? (alle mit " - ")
        Arr1 = Split(f2, v)
        For i = 0 To UBound(Arr1)
            'Schleife über alle f2-FileNames des aktuellen EventFolders
            If InStr(1, Arr1(i), " - ") = 0 Then C6 = C6 + 1
        Next
        If C6 = 0 Then
            'Keine DateiNamen zu bearbeiten --> Nächster EventFolder
            p = Get_FirstLineOfStringX_AndDeleteFirstLine(PathsOfSomeFolders)
            T6_Check_FileNamesInsideOneEventFolder p
            Exit Sub
        End If
    'Einige f2-FileNames haben noch kein " - "-Format
        ArrC(83) = "0"      '"1" = EventFileNamesSearchIsRunning
        T6_BMB_Format2DataArea
    'P1(), P2()     = Listen: FileNameOld, FileNameNew
        'Jetzt kann untersucht werden, ob die Files bereits " - " enthalten
        Arr1 = Split(v + f2, v): C4 = UBound(Arr1)
        ReDim p1(1 To C4): ReDim p2(1 To C4)
        For i = 1 To C4
            A = Arr1(i) 'FileNameOld (mit Date8-NameOfEventFolder + VnNnText)
            'c suchen (c = StartPosition VnNnText)
            C1 = Get_FirstPositionOf_Blank_LowLetter(A) 'Position von  " "+Kleinbuchstabe
            C2 = InStr(1, A, " (")
            If C1 > 0 And C2 = 0 Then c = C1
            If C1 = 0 And C2 > 0 Then c = C2
            If C1 > 0 And C2 > 0 Then c = Application.Min(C1, C2)
            If c > 0 Then
                'a=FileNameOld, b=FileNameNew
                B = sc + Mid(A, c)
                If A <> B Then C3 = 1   'Nicht alle DateiNamen haben das gewünschte Format
                p1(i) = A: p2(i) = B    'Pfad hierzu wäre: p + "\" + b
            Else
                OpenFolder PathOfEventFolder: Beep: Stop
            End If
        Next
        'showArray P1: showArray P2: Stop
        If C3 = 0 Then
            'Alle DateiNamen haben bereits das gewünschte Format
            'Nächster EventFolder
            p = Get_FirstLineOfStringX_AndDeleteFirstLine(PathsOfSomeFolders)
            T6_Check_FileNamesInsideOneEventFolder p
            Exit Sub
        End If
    'T6() füllen
        ReDim T6(1 To 3 * C4, 1 To 8)
        .Cells(5, 37) = "Sollen alle " + CStr(C4) + " EventFileNames geändert werden?"
        For i = 1 To C4
            T6(i + C5, 1) = "actual:":        T6(i + C5, 2) = p1(i):     T6(i + C5, 8) = " "
            T6(i + C5 + 1, 1) = "change to:": T6(i + C5 + 1, 2) = p2(i): T6(i + C5 + 1, 8) = " "
            Set rg = .Range(.Cells(i + C5 + 10, 37), .Cells(i + C5 + 11, 43)): rg.Interior.Color = Green1
            C5 = C5 + 2
        Next
    'FileNames in T6 schreiben
        Paste_2DArrayToCell_z_s "T6", 11, 37, T6 ': showArray2D T6
    'Finals
        End With
End Sub

Sub T6_BMB_Format1DataArea()
    'Called from    T6_BMB_EventFileNames
    'BMB            = BlankMinusBlank
    'Action         Vorbereitung einer AnzeigeListe mit Header und Buttons (ab T6-Spalte 37)
    'Buttons        Abbrechen
    
    'Vorbereitung
        Dim qq$, r As Range
        With Sheets("T6"): qq = Chr(34)
        'Beschriften
            .Cells(2, 37) = "EventFileNames: Add " + qq + " - " + qq
            .Cells(5, 37) = " FileNames der EventOrdner werden durchsucht: "
            .Cells(8, 41) = "Abbrechen"
        'Spaltenbreite
            .Columns(37).ColumnWidth = 10: .Columns(38).ColumnWidth = 2: .Columns(39).ColumnWidth = 10
            .Columns(40).ColumnWidth = 2: .Columns(41).ColumnWidth = 15: .Columns(42).ColumnWidth = 6
            .Columns(43).ColumnWidth = 25: .Columns(44).ColumnWidth = 2
        'Header
            Set r = .Range(.Cells(2, 37), .Cells(9, 43)): r.Font.size = 16: r.Font.Bold = True
            r.HorizontalAlignment = xlCenter: r.VerticalAlignment = xlCenter
            .Cells(5, 37).HorizontalAlignment = xlLeft '2. HeaderZeile
            Set r = .Range(.Cells(2, 37), .Cells(3, 43)): r.MergeCells = True: r.Interior.Color = Green1
            Set r = .Range(.Cells(5, 37), .Cells(6, 43)): r.MergeCells = True: r.Interior.Color = Green1
        'Buttons
            Set r = .Range(.Cells(8, 41), .Cells(9, 41)): r.MergeCells = True: r.Interior.Color = Green3
        'Listing Area
            .Range(.Cells(11, 37), .Cells(500, 37)).Font.size = 8
            .Range(.Cells(11, 38), .Cells(500, 38)).Font.size = 10
        'Select 1. Header-Zeile
            EE 0: .Cells(2, 37).Select: EE 1
        'Finals
            End With
End Sub

Sub T6_BMB_Format2DataArea()
    'Called from    T6_BMB_EventFileNames
    'BMB            = BlankMinusBlank
    'Action         Vorbereitung einer AnzeigeListe mit Header und Buttons (ab T6-Spalte 37)
    'Buttons        Ja - Nein - Abbrechen - Open EventFolder
    
    'Vorbereitung
        Dim qq$, r As Range
        With Sheets("T6"): qq = Chr(34)
        'Beschriften
            .Cells(5, 37) = "Sollen alle EventFileNames geändert werden?"
            .Cells(8, 37) = "Ja": .Cells(8, 39) = "Nein": .Cells(8, 43) = "Open EventFolder"
        'Header
            .Cells(5, 37).HorizontalAlignment = xlCenter '2. HeaderZeile
        'Buttons
            Set r = .Range(.Cells(8, 37), .Cells(9, 37)): r.MergeCells = True: r.Interior.Color = Green3
            Set r = .Range(.Cells(8, 39), .Cells(9, 39)): r.MergeCells = True: r.Interior.Color = Green3
            Set r = .Range(.Cells(8, 43), .Cells(9, 43)): r.MergeCells = True: r.Interior.Color = Green3
            Set r = .Range(.Cells(8, 42), .Cells(9, 42)): r.MergeCells = True: r.Font.size = 12
        'Finals
            End With
End Sub

Sub T6_BMB_CellButton_Ja(z%, s%)
    'Called from    [T6-CellButton 'Ja'] while doing [EventFileNames: Add ' - ' ...]
    'BMB            = BlankMinusBlank
    'Privat         M62-Privat-Variable: PathsOfSomeFolders
    
    'Vorbereitung
        If Sheets("T6").Cells(z, s) <> "Ja" Then Exit Sub
        Dim F$, C1%, C2%, C3%, i%, p(), p1() As String, p2() As String
        C1 = Get_NrOfLastRowInColumnNr(37, "T6"): With Sheets("T6")
    'P()
        p = .Range(.Cells(11, 38), .Cells(C1, 38)).Value
        C2 = (C1 - 9) \ 3 'Anzahl FileNames to change
    'P1(), P2()
        ReDim p1(1 To C2): ReDim p2(1 To C2): C3 = 1
        For i = 1 To UBound(p, 1) Step 3
            p1(C3) = ArrC(84) + "\" + p(i, 1)
            p2(C3) = ArrC(84) + "\" + p(i + 1, 1)
            C3 = C3 + 1
        Next 'showArray P1: showArray P2
    'Rename
        For i = 1 To UBound(p1)
            RenameFile p1(i), p2(i)
        Next
        .Cells(7, 37).Select
        End With
    'Nächster EventFolder
        F = Get_FirstLineOfStringX_AndDeleteFirstLine(PathsOfSomeFolders)
        T6_Check_FileNamesInsideOneEventFolder F
End Sub

Sub T6_BMB_CellButton_Nein(z%, s%)
    'Called from    [T6-CellButton 'Nein'] while doing [EventFileNames: Add ' - ' ...]
    'BMB            = BlankMinusBlank
    'Privat         M62-Privat-Variable: PathsOfSomeFolders
    
    If Sheets("T6").Cells(z, s) <> "Nein" Then Exit Sub
    Dim p$
    Sheets("T6").Cells(7, 39).Select
    'Nächster EventFolder
        p = Get_FirstLineOfStringX_AndDeleteFirstLine(PathsOfSomeFolders)
        T6_Check_FileNamesInsideOneEventFolder p
End Sub

Sub T6_BMB_CellButton_Abbrechen(z%, s%)
    'Called from    [T6-CellButton 'Abbrechen'] while doing [EventFileNames: Add ' - ' ...]
    'BMB            = BlankMinusBlank
    'Privat         M62-Privat-Variable: PathsOfSomeFolders
        
        If Sheets("T6").Cells(z, s) <> "Abbrechen" Then Exit Sub
    'Empty
        PathsOfSomeFolders = ""
    'T6 clear area
        Sheets("T6").Range("AK2:AR400").Clear
        Sheets("T6").Range("AK:AR").ColumnWidth = 2
End Sub

Sub T6_BMB_CellButton_OpenEventFolder(z%, s%)
    'Called from    [T6-CellButton 'Open EventFolder'] while doing [EventFileNames: Add ' - ' ...]
    'BMB            = BlankMinusBlank
    
    If Sheets("T6").Cells(z, s) <> "Open EventFolder" Then Exit Sub
        'ToDo:
            'if ArrC is not active then ...
    OpenFolder ArrC(84)
    Sheets("T6").Cells(7, 37).Select
End Sub

Function Get_FirstLineOfStringX_AndDeleteFirstLine(StringX$) As String
    'Privat         M62-Privat-Variable: PathsOfSomeFolders
    
    Dim p$, v$, c%
    v = vbCrLf
    c = InStr(1, StringX, v)
    If c = 0 Then
        'letzte Zeile der StringVariablen StringX
        p = StringX: StringX = "" ': If P = "" Then Stop
    Else
        p = Left(StringX, c - 1): StringX = Mid(StringX, c + 2)
    End If
    Get_FirstLineOfStringX_AndDeleteFirstLine = p
End Function

'------------ END [EventFileNames: Add ' - ' ...]

Sub T6_Create_DateSquares()
    'Called from    [T6-CellButton " Create DateSquares"]
    'DateSquare     jpg, 1000x1000, mit Datum, Hintergrund Schwarz, Schrift Weiß
    'Action         Erstellt 1 DateSquare für jedes (verschiedene) vorkommende Datum
    '               in den 'Leute'-Ordnern und allen UnterOrdnern von 'ClubsNations'
    '               (Event-Ordner erhalten keine DateSquares)
    
    'Vorbereitung
        Dim D$, dd$, FN$, p$, pFi$, v$, i%, j%, zT6%, Fi() As String, Fo() As String
        v = vbCrLf: DoArrc: Ticks1
        LogBuch "T6_Create_DateSquares was started"
    'MyCountDown Init
        'zT6 = T6-ZeilenNr, die den Cell-Button enthält (für T6_DoDoneRemarks-Einträge)
        zT6 = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 2, 4, " Create DateSquares")
        MyCountDown_Init "|T6|" + CStr(zT6) + "|19|99|" '= |SheetName|z|s|StartNr| of CountDownShowCell
    'Load
        Load_ArrPathsOfAllFolders_Leute_ClubsNations Fo 'Fo(1) = "Aaron, Syd (GB)"
        MyCountDown_OneMoreMainStep
        MyCountDown_MainStepsAllowed 98: MyCountDown_SubStepsMax UBound(Fo)
        
    For i = 0 To UBound(Fo)
        MyCountDown_OneMoreSubStep
        p = Fo(i)   'PathOfOneFolder
            'Sheets("T6").Cells(3, 40) = "[" + CStr(UBound(Fo) - i) + "] " + getNameOfPath(Fo(i))
        'fn = Get_AllFileNames_Like_OfOneFolder(p, "########*####-##*")
        FN = Get_AllFileNames_Like_OfOneFolder(p, "########*")
        If FN = "" Then GoTo Jump
        Fi = Split(FN, v): dd = "": pFi = ""
        For j = 0 To UBound(Fi)
            D = "|" + Left(Fi(j), 8) + "|"
            If Not dd Like "*" + D + "*" Then dd = dd + D: pFi = pFi + v + p + "\" + Fi(j)
        Next
        If dd = "" Then GoTo Jump
        Fi = Split(pFi, v)
        For j = 1 To UBound(Fi)
            Create_OneDateSquare_Give_Date8File Fi(j) 'pathOfFile with Date8 inside folder p
        Next
        'OpenFolder p: Stop: CloseFolder p
        DoEvents
Jump:
    Next
    'T6_DoDoneRemarks
        T6_DoDoneRemarks
    'Finals
        LogBuch "T6_Create_DateSquares has ended"
       Beep
    Beep
End Sub

Sub Create_OneDateSquare_Give_Date8File_TEST()
    Dim p$
    p = "F:\Archiv TR\Archiv Trampolin\Leute\Ramseger, Sabine (TV Gernsbach)\19780000 Buli Bruchsal-Gernsbach p0318-00 sf.jpg"
    Create_OneDateSquare_Give_Date8File p
    Beep
End Sub

Sub Create_OneDateSquare_Give_Date8File(PathOfFileWithDate8$)
    'Called from    T6_Create_DateSquares
    'DateSquare     jpg, 1000x1000, mit Datum, Hintergrund Schwarz, Schrift Weiß
    'Date8File      Datei, deren Name mit 8 Date-Ziffern beginnt
    'Path...Date8   Pfad zu irgendeinem Date8File im DateSquare-ZielOrdner
    'Action         Erstellt 1 DateSquare
    
    'Exit
        Dim p8$, p9$
        p8 = PathOfFileWithDate8
        p9 = Get_PathOfParentFolder(p8) + "\" + Left(getNameOfPath(p8), 8) + " ! .jpg"
        If FileExists(p9) Then Exit Sub
    'Vorbereitung
        Dim N8$, p$, p1$, p2$, p3$, p4$, p5$, pTool$, qq$, sCmd$, Text1$, Text2$, shResult%
        qq = Chr(34): DoArrc
        p = ArrC(1) + "\prog\Label\DateSquare\" 'Ordner für HelperJpgs
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        MyCountDown_Add1ToChanges
    'Text1, Text2
        N8 = getNameOfPath(p8)
        Text2 = Left(N8, 4) 'Jahreszahl
'        If Mid(N8, 5, 4) = "0000" Then
'            Text1 = "...."
'            ElseIf Mid(N8, 7, 2) = "00" Then Text1 = "00." + Mid(N8, 5, 2) + "."
'            Else: Text1 = Mid(N8, 7, 2) + "." + Mid(N8, 5, 2) + "."
'        End If
        If Mid(N8, 5, 4) Like "####" Then Text1 = Mid(N8, 7, 2) + "." + Mid(N8, 5, 2) + "."
        
        
    'OpenFolder for testing
        'OpenFolder Get_PathOfParentFolder(p8) ': show p9: Stop
        
        
    'Pfade
        p1 = p + "1.jpg": p2 = p + "2.jpg": p3 = p + "3.jpg": p4 = p + "4.jpg": p5 = p + "5.jpg"
    'Create Black1
        sCmd = pTool + "convert -size 1000x450 xc:black" + " " + qq + p1 + qq
        shResult = ShellAndWait(sCmd, 0, vbHide, PromptUser)
    'Create Text1
        sCmd = pTool + "convert -background none -fill white -font Arial " _
               + "-pointsize 200 -size 1000x250 -gravity center label:" + Text1 + " " + qq + p2 + qq
        shResult = ShellAndWait(sCmd, 0, vbHide, PromptUser)
    'Create Text2
        'sCmd = pTOOL + "convert -background none -fill white -font Arial " _
               + "-pointsize 300 -size 1000x250 -gravity center label:" + Text2 + " " + qq + p3 + qq
        sCmd = pTool + "convert -background none -fill white -font " + qq + "Arial-Bold" + qq + " " _
               + "-pointsize 300 -size 1000x250 -gravity center label:" + Text2 + " " + qq + p3 + qq
               
               '"Arial-Bold-Italic"
        shResult = ShellAndWait(sCmd, 0, vbHide, PromptUser)
    'Create Black2
        sCmd = pTool + "convert -size 1000x50 xc:black" + " " + qq + p4 + qq
        shResult = ShellAndWait(sCmd, 0, vbHide, PromptUser)
    'p1 p2 p3 p4 untereinander
        Hänge_jpg2_unter_jpg1 p1, p2, p5
        Hänge_jpg2_unter_jpg1 p5, p3, p5
        Hänge_jpg2_unter_jpg1 p5, p4, p5
        
        'show p5 + vbCrLf + p9: Stop
        CopyFile p5, p9
End Sub

Sub MyCountDown_Add1ToChanges()
    ArrC(24) = CStr(CInt(ArrC(24)) + 1)
    'If CInt(ArrC(24)) Mod 10 = 0 Then
    'N = ArrC(29)                'T0 MyCountDown ShowCell-NameOfSheet
    If ArrC(29) <> "0" Then
        Sheets(ArrC(29)).Cells(CInt(ArrC(30)), CInt(ArrC(31) + 1)) = ArrC(24)
    End If
    'DoEvents
End Sub
Sub MyCountDown_Add2ToChanges()
    ArrC(24) = CStr(CInt(ArrC(24)) + 2)
    'If CInt(ArrC(24)) Mod 10 = 0 Then
    Sheets(ArrC(29)).Cells(CInt(ArrC(30)), CInt(ArrC(31) + 1)) = ArrC(24)
    'DoEvents
End Sub
Sub MyCountDown_Add3ToChanges()
    ArrC(24) = CStr(CInt(ArrC(24)) + 3)
    'If CInt(ArrC(24)) Mod 10 = 0 Then
    Sheets(ArrC(29)).Cells(CInt(ArrC(30)), CInt(ArrC(31) + 1)) = ArrC(24)
    'DoEvents
End Sub
Sub MyCountDown_Add4ToChanges()
    ArrC(24) = CStr(CInt(ArrC(24)) + 4)
    'If CInt(ArrC(24)) Mod 10 = 0 Then
    Sheets(ArrC(29)).Cells(CInt(ArrC(30)), CInt(ArrC(31) + 1)) = ArrC(24)
    'DoEvents
End Sub

Function Get_FirstPositionOf_Blank_LowLetter(s$) As Integer
    If Not s Like "* [a-z]*" Then Exit Function
    Dim c%, C1%, i%
    C1 = 1
    For i = 1 To 99
        C1 = InStr(C1, s, " ")
        If C1 = 0 Then Exit Function
        If Mid(s, C1 + 1, 1) Like "[a-z]" Then c = C1: Exit For
        C1 = C1 + 1
    Next
    Get_FirstPositionOf_Blank_LowLetter = c
End Function

Sub T6_Add_Links_PersonLTVClubNation()
    'Called from    [T6-CellButton ' Add Links: Person-Club-LTV-Nation']
    'T5             Das Arbeitsblatt "T5" enthält Personendaten;
    '               pro Person existiert eine Zeile mit Daten zu dieser Person
    '               in den Spalten (Spalte 3: Nachname, Spalte 4: Vorname, ...)
    'Load           hier werden einzelne Spalten als 1-basiertes String-Array geladen:
    '               Nn, Vn, Club, LTV, Nation, Personfolder, VnNn
    'Action         Abarbeiten aller T5-Zeilen
    '               (1)  Nat=D   Link in PersonFolder to ClubFolder
    '               (2)  Nat=D   Link in PersonFolder to LTVFolder      CreateLTVFolder
    '               (3)  Nat=D   Link in PersonFolder to NationFolder
    '               (4)  Nat=D   Link in ClubFolder   to PersonFolder   CreateClubFolder
    '               (5)  Nat=D   Link in ClubFolder   to LTVFolder
    '               (6)  Nat=D   Link in ClubFolder   to NationFolder
    '               (7)  Nat=D   Link in LTVFolder    to PersonFolder
    '               (8)  Nat=D   Link in LTVFolder    to NationFolder
    '               (9)  Nat<>D  Link in NationFolder to PersonFolder   CreateNationFolder
    '               (10) Nat<>D  Link in PersonFolder to NationFolder
    '               (11) vh
    'Vorbereitung
        Dim Club() As String, PersonFolder() As String, LTV() As String, Nation() As String
        Dim Nn() As String, Vn() As String, VnNn() As String
        Dim ClubFolder$, LTVFolder$, NameOfPersonfolder$, NationFolder$, PersFolder$
        Dim pE1$, pE2$, PL$, v$, i%, zT6%
        v = vbCrLf: DoArrc: Ticks1: pE1 = ArrC(3): pE2 = ArrC(11): PL = ArrC(4)
        LogBuch "T6_Add_Links_PersonLTVClubNation was started"
    'MyCountDown Init
        'zT6 = T6-ZeilenNr, die den Cell-Button enthält (für T6_DoDoneRemarks-Einträge)
        zT6 = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 2, 4, " Add Links: Person-Club-LTV-Nation")
        MyCountDown_Init "|T6|" + CStr(zT6) + "|19|99|" '= |SheetName|z|s|StartNr| of CountDownShowCell
    'Load Nn Vn Club LTV Nation PersonFolder VnNn
        T5_Load_Club Club: T5_Load_LTV LTV: T5_Load_Nation Nation: T5_Load_Nn Nn
        T5_Load_Folder PersonFolder: T5_Load_Vn Vn: T5_Load_VnNn VnNn
    'Create Links
        MyCountDown_MainStepsAllowed 98: MyCountDown_SubStepsMax UBound(VnNn)
        For i = 1 To UBound(VnNn)
            'Schleife über alle T5-Zeilen
            MyCountDown_OneMoreSubStep
            NameOfPersonfolder = PersonFolder(i)
            PersFolder = PL + "\" + NameOfPersonfolder
            If Nation(i) = "D" Then
                NationFolder = pE2 + "\D"
                If LTV(i) <> "" Then
                        'LTV bekannt in T5
                        LTVFolder = pE2 + "\D\" + LTV(i)
                        If Not FolderExists(LTVFolder) Then CreateFolder LTVFolder
    '(2)                Nat=D, Link in PersonFolder to LTVFolder, CreateLTVFolder
    
    
    
                            OpenFolder PersFolder: Stop
                        Check_LtvLink_One PersFolder, LTVFolder, "-  " + LTV(i)
                        Create_OneLinkFile PersFolder, LTVFolder, "-  " + LTV(i)
                        
                        
                        
    '(3)                Nat=D, Link in PersonFolder to NationFolder
                        Create_OneLinkFile PersFolder, NationFolder, "-   D"
    '(8)                Nat=D, Link in LTVFolder to NationFolder
                            'Stop: CloseFolder PersFolder
                            'OpenFolder LTVFolder: Stop
                        Create_OneLinkFile LTVFolder, NationFolder, "-   D"
                            'Stop: CloseFolder LTVFolder
                        MyCountDown_Add3ToChanges
                    If Club(i) <> "" Then
                        'LTV bekannt, Club bekannt in T5
                        ClubFolder = pE2 + "\D\" + LTV(i) + "\" + Club(i)
                        If Not FolderExists(ClubFolder) Then CreateFolder ClubFolder
    '(1)                Nat=D   Link in PersonFolder to ClubFolder
                            'OpenFolder PersFolder: Stop
                        Create_OneLinkFile PersFolder, ClubFolder, "- " + Club(i)
                            'Stop: CloseFolder PersFolder
    '(4)                Nat=D, Link in ClubFolder to PersonFolder, CreateClubFolder
                            'OpenFolder ClubFolder: Stop
                        Create_OneLinkFile ClubFolder, PersFolder, Nn(i) + ", " + Vn(i)
    '(5)                Nat=D, Link in ClubFolder to LTVFolder
                        Create_OneLinkFile ClubFolder, LTVFolder, "-  " + LTV(i)
    '(6)                Nat=D, Link in ClubFolder to NationFolder
                        Create_OneLinkFile ClubFolder, NationFolder, "-   D"
                            'Stop: CloseFolder ClubFolder
                        MyCountDown_Add4ToChanges
                    Else
                        'LTV bekannt, Club unbekannt in T5
    '(7)                Nat=D   Link in LTVFolder    to PersonFolder
                            'OpenFolder LTVFolder: Stop
                        Create_OneLinkFile LTVFolder, PersFolder, Nn(i) + ", " + Vn(i)
                            'Stop: CloseFolder LTVFolder
                        MyCountDown_Add1ToChanges
                    End If
                End If
            Else
    '(9)        Nat<>D, Link in NationFolder to PersonFolder, CreateNationFolder
                NationFolder = pE2 + "\" + Nation(i)
                If Not FolderExists(NationFolder) Then CreateFolder NationFolder
                            'OpenFolder NationFolder: Stop
                Create_OneLinkFile NationFolder, PersFolder, Nn(i) + ", " + Vn(i)
                            'Stop: CloseFolder NationFolder
    '(10)       Nat<>D, Link in PersonFolder to NationFolder
                            'OpenFolder PersFolder: Stop
                Create_OneLinkFile PersFolder, NationFolder, "-   " + Nation(i)
                            'Stop: CloseFolder PersFolder
                MyCountDown_Add2ToChanges
            End If
            DoEvents
        Next
    'T6_DoDoneRemarks
        T6_DoDoneRemarks
    'Finals
        LogBuch "T6_Add_Links_PersonLTVClubNation has ended"
       Beep
End Sub

Sub T6_Update_AgeOfAllCompetitorsInsideAllDgs(z%)
    'Called from    [T6-CellButton ' Update Age of all Competitors in all Dgs']
    
    'Vorbereitung
        Dim IST$, L$, Nam$, Nam1$, Nam2$, Report$, SOLL$, T$, v$
        Dim C1%, C2%, C3%, C4%, cMax%, i%, j%, s1%, z1%
        Dim JH() As String, N() As String, m() As String, Vn() As String, so(), T4()
        v = vbCrLf: Ticks1
        LogBuch "T6_Update_AgeOfAllCompetitorsInsideAllDgs was started"
    'MyCountDown
        MyCountDown_Init "|T6|" + CStr(z) + "|19|18|"  '= |SheetName|z|s|StartNr| of CountDownShowCell
        'MyCountDown_MainStepsAllowed 20: MyCountDown_SubStepsMax AnzFold
        'MyCountDown_OneMoreSubStep 'MyCountDown_OneMoreMainStep
    'SO()   2D-Array; T4SomeDgData, Spalten 1-8: Title z1 s1 z2 s2 Folder - Names
            T4_Load_T4SomeDgData_1to8 so ': MyCountDown_OneMoreMainStep ': showArray2D SO
    'T4()   2D-Array; Sheet T4 komplett
            T4_Load_CompleteSheet T4:     MyCountDown_OneMoreMainStep
    'JH()   1D-Array; T5 PersonData, Spalte6 Jhg
            T5_Load_Jhg JH ':               MyCountDown_OneMoreMainStep ': showArray JH: Stop
    'VN()   1D-Array; T5 PersonData, Spalte22 VnNn
            T5_Load_VnNn Vn ':              MyCountDown_OneMoreMainStep ': showArray VN: Stop
    'MyCountDown
        'Anzahl CompetitorNameCells ermitteln (für MyCountDown_SubStepsMax)
            For i = 1 To UBound(so, 1): C1 = C1 + CInt(Left(CStr(so(i, 8)), 3)): Next
        MyCountDown_MainStepsAllowed 5: MyCountDown_SubStepsMax C1
    'Names and SheetCell
        For i = 1 To UBound(so, 1)
            'Schleife über alle T4SomeDgData-Zeilen
            z1 = CInt(so(i, 2)): s1 = CInt(so(i, 3))
            N = Split(CStr(so(i, 8)), "|") '004 Names: |#|Christel Vieweg|6|3|Einz|1|#|Roland ...
            'showArray N
            For j = 0 To 999 Step 6
                'Schleife über alle CompetitorNames des i. Dg
                C2 = C2 + 1: MyCountDown_OneMoreSubStep 'c2 = CompetitorNameCells mitzählen
                If 5 + j < UBound(N) Then
                    'String Nam wird groß/Add langsam, deshalb Teile Nam1, Nam2
                    If C2 < C1 \ 2 Then
                        Nam1 = Nam1 + "|0000|" + CStr(N(2 + j)) + "|" + CStr(z1 - 1 + CInt(N(3 + j))) + "|" + CStr(s1 - 1 + CInt(N(4 + j))) + "|" + CStr(Left(so(i, 1), 4)) + "|" + v
                        '    = je  + Zeile wie "|0000|Linda Ball|14|33|1964|"
                    Else
                        Nam2 = Nam2 + "|0000|" + CStr(N(2 + j)) + "|" + CStr(z1 - 1 + CInt(N(3 + j))) + "|" + CStr(s1 - 1 + CInt(N(4 + j))) + "|" + CStr(Left(so(i, 1), 4)) + "|" + v
                    End If
                Else
                    j = 999 '= Exit For j
                End If
                DoEvents
            Next
        Next
        Nam = Delete_EndReturnsInString(Nam1 + Nam2)
        '   = Zeilen wie "|0000|Linda Ball|14|33|1964|"
        '   show CStr(c1) + " Zellen mit CompetitorName" + v + v + Nam
        'sort
            N = Split(Nam, v): QuickSort N: Nam = Join(N, v)
    'Add    Jhg to Nam; change "|0000|" to "|1964|" if Jhg exists in JH()
        MyCountDown_MainStepsAllowed 10: MyCountDown_SubStepsMax UBound(Vn, 1)
        For i = 1 To UBound(Vn, 1)
            If JH(i) <> "" Then
                Nam = Replace(Nam, "|0000|" + Vn(i) + "|", "|" + JH(i) + "|" + Vn(i) + "|")
            End If
            MyCountDown_OneMoreSubStep
        Next
        'show CStr(c1) + " Zellen mit CompetitorName" + v + v + Nam
        'Nam = Zeilen wie    "|0000|Linda Ball|14|33|1964|"
        '      oder          "|1977|Ji Wallace|177|342|1999|"
    'Check  CellsWithCompetitorName; "Ji Wallace" or "Ji Wallace (22)"
        N = Split(Nam, v)
        MyCountDown_MainStepsAllowed 2: MyCountDown_SubStepsMax UBound(N)
        For i = 1 To UBound(N)
            MyCountDown_OneMoreSubStep
            L = N(i)                    'L = "|1977|Ji Wallace|177|342|1999|"
            m = Split(L, "|")
            'SOLL
                If m(1) = "0000" Then
                    SOLL = m(2)         'SOLL = SOLL-Content der NameCell
                Else
                    SOLL = m(2) + " (" + CStr(CInt(m(5)) - CInt(m(1))) + ")"
                End If
            'IST
                IST = T4(CInt(m(3)), CInt(m(4)))    'IST = IST-Content der NameCell
            'Compare
                If IST <> SOLL Then Report = Report + "T4-Cell(" + Format(m(3), "0000") _
                + ", " + Format(m(4), "0000") + ") = '" _
                + IST + "', sollte aber '" + SOLL + "' sein." + v: C3 = C3 + 1
        Next
    'T6_DoDoneRemarks
        FillArrC 24, CStr(C3) 'CountOfChanges
        T6_DoDoneRemarks
    'Report
        T = "Report to 'Update Age of all Competitors in all Dgs'"
        If Report = "" Then
            Report = T + v + v + "Alle CompetitorNameCells trugen bereits den korrekten Inhalt." _
            + "Keine Änderungen"
        Else
            'PepUp Report
            N = Split(Report, v)
            For i = 0 To UBound(N) - 1
                C4 = InStr(1, N(i), "',"): If C4 > cMax Then cMax = C4
            Next
            For i = 0 To UBound(N) - 1
                C4 = InStr(1, N(i), "',"): N(i) = Replace(N(i), "',", "'," + String(cMax - C4 + 1, " "))
            Next
            Report = Join(N, v)
            Report = T + v + "Count of Changes: " + CStr(C3) + v + v + Report
        End If
        show Report
    'Change DgNames to DgNames with Age
        If Not Report Like "*Keine Änderungen*" Then T4_Add_AgeToReportedDgNameCells Report
    'Zurück auf T6
        Sheets("T6").Activate
    'Finals
        LogBuch "T6_Update_AgeOfAllCompetitorsInsideAllDgs has ended"
       Beep
End Sub

Sub T4_Add_AgeToReportedDgNameCells(Report$)
    'Called from    T6_Update_AgeOfAllCompetitorsInsideAllDgs
        
    'Vorbereitung
        Dim L$, i%, s%, z%, A() As String
        With Sheets("T4"): .Activate
    'Action
        A = Split(Report, vbCrLf)
        For i = 2 To UBound(A)
            L = A(i)            'one Line
            If Left(L, 7) = "T4-Cell" Then
                z = CInt(Mid(L, 9, 4)): s = CInt(Mid(L, 15, 4)) 'DgNameCell (z, s)
                'Zeige DgNameCell ohne, dann mit Age
                    .Cells(z - 1, s).Select: RefreshScreen
                    .Cells(z, s).Select: RefreshScreen
                    .Cells(z + 1, s).Select: RefreshScreen
                    .Cells(z, s).Select: RefreshScreen
                'Dg ist aktiviert, DgNameCell aktualisiert und ausgewählt
                    ClickButton_CreateJpg_InActualDg
                    DoEvents 'OpenFolder ArrC(57)
            End If
        Next
    'Finals
        End With
End Sub

Sub T6_Update_EachResultJpg(zT6%)
    'Called from    [UserClick on T6-CellButton 'Add ID in open EventFolders'] T6_UserClickOnColumn2CellButton
        
    'Vorbereitung
        Dim c%, i%, s%, z%, L!, T!, z1() As String, s1() As String, r As Range
        Call DoArr: Ticks1
        LogBuch "T6_Update_EachResultJpg was started"
    'Anzahl Dgs
        c = Get_NrOfLastRowInColumnNr(2, "T4") - 7
    'MyCountDown
        MyCountDown_Init "|T6|" + CStr(zT6) + "|19|" + CStr(c) + "|" '= |SheetName|z|s|StartNr| of CountDownShowCell
        FillArrC 24, CStr(c) 'CountOfChanges
    'Updates?
        'T4SomeDgData, AgeOfAllCompetitors
    'z1(), s1()
        T4_Load_T4SomeDgData_OneCol 2, z1
        T4_Load_T4SomeDgData_OneCol 3, s1
    'Select each Dg
        With Worksheets("T4")
            'Selections to empty ArrDgOld
                .Activate ': .Cells(1, 2).Select: .Cells(1, 1).Select
        For i = 1 To UBound(z1) 'For i = 3 To 3
            'Select new Dg (1,1)
                z = CInt(z1(i)): s = CInt(s1(i))
                T4_LabelForCountDown_Set z, s
                ActiveWindow.ScrollRow = z - 2
                ActiveWindow.ScrollColumn = s - 2 'Spalte, die ganz links zu sehen sein soll
                RefreshScreen
                .Cells(z + 1, s + 1).Select 'jumps to Dg(2,2); Select
                RefreshScreen
                .Cells(z, s).Select 'Dg(1,1).Select
                RefreshScreen
            'Show CountDown
                'T4_LabelForCountDown_Set z, s
            'CreateResultsJpg
                ClickButton_CreateJpg_InActualDg
            'Selections to empty ArrDgOld
                '.Cells(1, 2).Select: .Cells(1, 1).Select
            MyCountDown_OneMoreMainStep
            DoEvents
        Next
        End With
    'Hide LabelForCountDown
        With Worksheets("T4").Shapes("LabelForCountDown"): .Left = 0: .Top = 9999: .Visible = msoFalse: End With
    'DoneRemarks to T6
        T6_DoDoneRemarks
    'Finals
        LogBuch "T6_Update_EachResultJpg has ended"
        Beep
End Sub

Sub T6_Show_AllPathsOfEventResultJpgs(z%)
    'Called from    T6_UserClickOnColumn2CellButton
    
    'Vorbereitung
        Dim p1$, p2$, pEvents$, v$
        Dim C1%, C2%, i%, Arr1() As String, r As Range
        With Worksheets("T6"): DoArr: Ticks1
        v = vbCrLf: pEvents = ArrC(3) 'Path of Folder "Events"
    'MyCountDown
        MyCountDown_Init "|T6|" + CStr(z) + "|19|26|"  '= |SheetName|z|s|StartNr| of CountDownShowCell
        C1 = Get_CountOfPaths_OfAllSubfolders_AllLevels(pEvents)
        MyCountDown_MainStepsAllowed 5: MyCountDown_SubStepsMax C1
        'If MyCountDown Then
        'MyCountDown_MainStepsAllowed 20: MyCountDown_SubStepsMax AnzFold
        'MyCountDown_OneMoreSubStep 'MyCountDown_OneMoreMainStep
    'pEvents SubFolders
        'p1 = Get_Paths_OfAllSubfolders_AllLevels(pEvents)
        p1 = Get_Paths_OfAllSubfolders_AllLevels_Collection(pEvents, True)
    'pEvents Files #
        Arr1 = Split(p1, v)
        MyCountDown_SubStepsMax UBound(Arr1): MyCountDown_MainStepsAllowed 20
        For i = 0 To UBound(Arr1)
            MyCountDown_OneMoreSubStep
            p2 = p2 + Get_AllFilePaths_LikeMyStringInFileName_OfOneFolder(Arr1(i), "*[#]*") + v
            .Cells(z, 20) = anzAinB("#", p2) 'CountOfChanges
        Next
        p2 = Delete_EmptyRowsInString(p2): C2 = anzAinB("#", p2): ArrC(24) = CStr(C2)
        p2 = CStr(C2) + " paths of #result-jpgs" + v + v + p2
        show p2 ': Stop
    'DoneRemarks to T6
        T6_DoDoneRemarks
    'Finals
        End With: Beep
End Sub

Sub T6_DoDoneRemarks()
    'Called from    T6_Update_AgeOfAllCompetitorsInsideAllDgs, T6_Show_AllPathsOfEventResultJpgs, ...
    
    'Vorbereitung
        Dim CountOfChanges%, z%, r As Range
        DoArrc
        z = CInt(ArrC(30))        'T0 MyCountDown CountDownShowCell-NrOfRow
        With Sheets("T6") ': .Activate: .Cells(1, 1).Select
    'LastDone
        Set r = .Cells(z, 18)
        r.Value = Format(Now(), "yyyymmdd_hhmmss")
        r.Font.size = 6: r.HorizontalAlignment = xlLeft
    'Days ago
        Set r = .Cells(z, 19)
        r.Value = 0:                        r.Font.size = 8: r.HorizontalAlignment = xlCenter
    'Count of changes (Anzahl ergänzter IDs)
        Set r = .Cells(z, 20): CountOfChanges = CInt(ArrC(24))
        r.Value = CountOfChanges:           r.Font.size = 8: r.HorizontalAlignment = xlCenter
    'Erledigt-Häkchen setzen
        .Cells(z, 16) = "ü"
    'Duration
        Set r = .Cells(z, 21)
        r.Value = T6_GetDuration():   r.Font.size = 8: r.HorizontalAlignment = xlCenter
    'CellButton (Green4) zurück auf Green3 setzen, Select weg
        Set r = .Range(.Cells(z, 2), .Cells(z, 15)): r.Interior.Color = Green3
        If ActiveSheet.NAME = "T6" Then EE 0: .Cells(z, 1).Select: EE 1
    'CountDownShowCell zurück auf Green1 setzen
        Set r = Sheets(ArrC(29)).Cells(CInt(ArrC(30)), CInt(ArrC(31)))
        r.Interior.Color = Green1
    'ArrC
        FillArrC 24, "0": FillArrC 25, "0": FillArrC 26, "0": FillArrC 27, "0"
        FillArrC 28, "0": FillArrC 29, "0": FillArrC 30, "0": FillArrC 31, "0"
    'Finals
        End With
End Sub

Sub T6_Check_PIDs_VIDs()
    'Called from    [T6-CellButton "Check PIDs VIDs"] T6_UserClickOnColumn2CellButton
    'PID            = IdentifikationsNr für ein Foto;  sollte eindeutig sein
    'VID            = IdentifikationsNr für ein Video; sollte eindeutig sein
    'Folder         E 'Events', I 'ClubsNations', L 'Leute',
    'Listen         Liste1 bis Liste5:
    '                    gleicher Aufbau (PidVid8|size10|Date8|Folder6|FilePath);
    '               no:  .lnk .url .pdf .srt .txt .dng .ini;
    '               nur: .jpg .avi .mp4 .mpg .wmv ...
    'Action         In E+I sollte keine Pid oder Vid mehrfach vorkommen;
    '                   falls doch, stehen diese in den Liste4-Zeilen
    '               In L befinden sich Kopien von Fotos/Videos aus E+I;
    '                   eine Datei und ihre Kopie sollten gleiche Pid/Vid haben;
    '                   auch sollten Original und Kopie die selbe Größe haben;
    '                   ist dies nicht der Fall, stehen diese in den Liste5-Zeilen
    '               (1) Liste aller SubFolderPaths, je von E/I/L
    '               (2) Liste1: PidVid8|size10|Date8|Folder6|FilePath aus E+I
    '               (3) Liste2: PidVid8|size10|Date8|Folder6|FilePath aus L
    '               (4) Liste3: PidVid8|size10|Date8|Folder6|FilePath aus E+I+L
    '               (5) Check FreePIDs, Check FreeVIDs
    '               (6) Liste4: E+I    NotUnique(Liste1); same Pid --> >=2 images; dito Vid
    '               (7) Liste5: E+I+L  PidGleich/SizeUngleich(Liste3); dito Vid
    '                           = Check, ob jpgs mit gleicher pId auch gleiche Größe haben;
    '                           falls ja gilt vermutlich: gleiche pId - gleiches jpg
 
    'Vorbereitung
        Dim C1$, C3$, C4$, C5$, c7$, Info$, NameOfButton$, Liste1$, Liste2$, Liste3$, Liste4$, Liste5$
        Dim p1$, p2$, p3$, s1$, s2$, s3$, GL$, FinalText$, txtFreePIDs$, txtFreeVIDs$, v$
        Dim CountOfChanges%, z%
        v = vbCrLf: DoArrc: NameOfButton = " Check PIDs VIDs": Ticks1
        p1 = ArrC(3): p2 = ArrC(11): p3 = ArrC(4) 'Path of Folder Events|ClubsNations|Leute
        z = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 2, 4, NameOfButton)
    'MyCountDown
         MyCountDown_Init "|T6|" + CStr(z) + "|35|82|" '= |SheetName|z|s|StartNr| of CountDownShowCell
         'T6_Create_MyCountDownBox
    '(1) Liste aller SubFolderPaths
         s1 = Get_Paths_OfAllSubfolders_AllLevels(p1, True)  'Events
         s2 = Get_Paths_OfAllSubfolders_AllLevels(p2, True)  'ClubsNations
         s3 = Get_Paths_OfAllSubfolders_AllLevels(p3, True)  'Leute
    '(2) Liste1: PidVid8|size10|Date8|Folder6|FilePath aus E+I
         Liste1 = Get_AllFilePathsOfFolderList_PidVid8size10Date8Folder6Path(s1 + v + s2)
         'show MyExt(Liste1) '".avi|.mp4|.mpg|.wmv|"
         'show "Liste1: Alle " + CStr(anzAinB(":", Liste1)) + " FilePaths (no .lnk .url .pdf .srt .txt .dng .ini) in Events/ClubsNations" + v + v + Liste1
    '(3) Liste2: PidVid8|size10|Date8|Folder6|FilePath aus L
         Liste2 = Get_AllFilePathsOfFolderList_PidVid8size10Date8Folder6Path(s3)
         'show "Liste2: Alle " + CStr(anzAinB(":", Liste2)) + " FilePaths (no .lnk .url .pdf .srt .txt .dng .ini) in Leute" + v + v + Liste2
    '(4) Liste3: PidVid8|size10|Date8|Folder6|FilePath aus E+I+L
         Liste3 = AddTwoListsAndSort(Liste1, Liste2)
         'show "Liste3: Alle " + CStr(anzAinB(":", Liste3)) + " FilePaths (no .lnk .url .pdf .srt .txt .dng .ini) in Events/ClubsNations/Leute" + v + v + Liste3
    '(5) Check FreePIDs         'Gute Gelegenheit, da alle PID-Files in Liste4 stehen
         txtFreePIDs = T6_Check_FreePIDs(Liste1)
         txtFreeVIDs = T6_Check_FreeVIDs(Liste1)
    '(6) Liste4: E+I    NotUnique(Liste1); same Pid --> >=2 images; dito Vid
         Liste4 = T6_Get_PIDsVIDsNotUnique(Liste1)      'CountOfPIDsNu = anzAinB(v + "p", v + Liste4)
         'show "Liste4: Alle " + CStr(CountOfChanges) + " Filepaths mit nicht eindeutiger PID in Events/ClubsNations" + v + v + Liste4
    '(7) Liste5: E+I+L  PidGleich/SizeUngleich(Liste3); dito Vid
         Liste5 = Get_IDsEqual_SizeNotEqual(Liste3)
    'Info
        GL = v + String(100, "-") + v 'gestrichelte Linie
        C1 = Add_FrontBlanks5(anzAinB(":", s1 + s2 + s3))   'Anz Ordner                 E+I+L
        C3 = Add_FrontBlanks5(anzAinB(":", Liste3))         'Anz Foto-/Videodateien     E+I+L
        C4 = Add_FrontBlanks5(anzAinB(":", Liste4))         'Anz NotUnique              E+I
        C5 = Add_FrontBlanks5(anzAinB(":", Liste5))         'Anz IDsEqual_SizeNotEqual  E+I+L
        c7 = Add_FrontBlanks5(CStr(CInt(C4) + CInt(C5)))    'CountOfChanges
        'FinalText     Text zu Liste4, Liste5
            If C4 > 0 Then FinalText = Trim(C4) + " FilePaths mit nicht eindeutiger ID:" + v + v + "   ID   |   Size   |  Date8 |Folder| FilePath" + v + v + Liste4 + GL
            If C5 > 0 Then FinalText = FinalText + Trim(C5) + " FilePaths mit IDsEqual_SizeNotEqual:" + v + v + "   ID   |   Size   |  Date8 |Folder| FilePath" + v + v + Liste5 + GL
        Info = "Info zu   T6-CellButton '" + Trim(NameOfButton) + "' (" + Format(Now(), "YYYYMMDD_hhmmss") + ")" + v _
            + "          PID = IdentifikationsNr für ein Foto; sollte eindeutig sein" + v _
            + "          VID = IdentifikationsNr für ein Video; sollte eindeutig sein" + GL _
            + txtFreePIDs + v + txtFreeVIDs + v _
            + C1 + " Ordner wurden durchsucht (in Events, ClubsNations, Leute)" + v _
            + C3 + " Foto-/VideoDateien wurden gefunden           (in E+I+L)" + v _
            + C4 + " FilePaths mit nicht eindeutiger ID gefunden  (in E+I)" + v _
            + C5 + " FilePaths mit IDsEqual_SizeNotEqual gefunden (in E+I+L)" + GL + FinalText
        Info = Replace(Info, "-----" + v + "-----", "-----")
        Info = Replace(Info, "-----" + v + v + "-----", "-----")
        show Info
    'DoneRemarks to T6
        CountOfChanges = CInt(c7)
        T6_Write_T6_DoDoneRemarks NameOfButton, CountOfChanges
    'Finals
        T6_Destroy_MyCountDownBox
        MyCountDown_End NameOfButton
End Sub

Function AddTwoListsAndSort(Liste1$, Liste2$) As String
    Dim v$, Arr1() As String: v = vbCrLf
    MyCountDown_OneMoreMainStep
    Arr1 = Split(Liste1 + v + Liste2, v)
    QuickSort Arr1: AddTwoListsAndSort = Join(Arr1, v)
End Function

Function MyExt(s$)
    Dim E$, L$, Last4$, v$, i&, Arr1() As String
    v = vbCrLf: Last4 = ".lnk.url.pdf.srt.txt.dng.ini.jpg"
    Arr1 = Split(s, v)
    For i = 0 To UBound(Arr1)
        L = Right(Arr1(i), 4):
        If Not Last4 Like "*" + L + "*" Then
            If Not E Like "*" + L + "*" Then E = E + L + "|"
        End If
    Next
    MyExt = E
End Function

Sub T6_Create_MyCountDownBox()
    'Called from    T6_Check_PIDs_VIDs
    
    'Vorbereitung
        Dim N$, s%, z%, r As Range
        N = ArrC(29)          'T0 MyCountDown CountDownShowCell-NameOfSheet
        z = CInt(ArrC(30))    'T0 MyCountDown CountDownShowCell-NrOfRow      3x3-Verbund Mitte
        s = CInt(ArrC(31))    'T0 MyCountDown CountDownShowCell-NrOfColumn   3x3-Verbund li
    'CountDownShowCell
        Set r = Sheets(N).Cells(z - 1, s): r.Font.Bold = True: r.Font.size = 20
        r.HorizontalAlignment = xlCenter: r.VerticalAlignment = xlCenter
    'Left of 3x3ShowCell
        Set r = Sheets(N).Cells(z, 34): r.Value = "----"
        r.HorizontalAlignment = xlCenter: r.VerticalAlignment = xlCenter
    'MergeCells
        Set r = Sheets(N).Range(Cells(z - 1, 35), Cells(z + 1, 37))
        r.MergeCells = True: r.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
End Sub

Sub T6_Destroy_MyCountDownBox()
    'Called from    T6_Check_PIDs_VIDs
    
    'Vorbereitung
        Dim N$, s%, z%, r As Range
        N = ArrC(29)          'T0 MyCountDown CountDownShowCell-NameOfSheet
        z = CInt(ArrC(30))    'T0 MyCountDown CountDownShowCell-NrOfRow
        s = CInt(ArrC(31))    'T0 MyCountDown CountDownShowCell-NrOfColumn
'    'CountDownShowCell
'        Set r = Sheets(N).Cells(z, s):  r.Clear
'    'Left of CountDownShowCell
'        Set r = Sheets(N).Cells(z, 34): r.Clear
    'MergeCells
        Set r = Sheets(N).Range(Cells(z - 1, 34), Cells(z + 1, 37)): r.Clear
        'r.MergeCells = False:
End Sub

Function Get_IDsEqual_SizeNotEqual(Liste3$) As String
    'Called from    T6_Check_PIDs_VIDs
    'Liste3         PidVid8|size10|Date8|Folder6|FilePath aus E+I+L
    
    'Vorbereitung
        Dim B$, L1$, Liste5$, ID1$, ID2$, v$, i&, Arr1() As String
        v = vbCrLf
        
    'Action
        Arr1 = Split(Liste3, v)      'Neubelegung von Arr1
        Liste5 = "": ID1 = "": B = ""
        MyCountDown_MainStepsAllowed 3
            On Error Resume Next
                MyCountDown_SubStepsMax UBound(Arr1)
            On Error GoTo 0
        For i = 0 To UBound(Arr1)
            'Schleife über alle Liste3-Zeilen PidVid8|size10|Date8|Folder6|FilePath
            MyCountDown_OneMoreSubStep
            L1 = Arr1(i)        'OneLine    'p0001-00|0000087544|19850000|Inside|F:\Archiv...jpg
            ID2 = Left(L1, 8)   'OnePid     'p0001-00 oder v0001-00
            If ID2 = ID1 Then
                'aktuelle pId (ID2) ist die der Vorzeile (ID1)
                B = B + L1 + v 'B-Container (Container für gleiche IDs)
            Else
                'aktuelle pId ist neu (nicht die der Vorzeile)
                'B-Container auswerten
                    If T6_PidsEqualGroupContainsDifferentSizes(B) Then Liste5 = Liste5 + B + v + "-----" + v
                'B-Container neu anlegen
                    B = L1 + v
            End If
            ID1 = ID2
        Next
        'letzten B-Container auswerten
            If T6_PidsEqualGroupContainsDifferentSizes(B) Then Liste5 = Liste5 + B + v + "-----"
    'Finals
        Get_IDsEqual_SizeNotEqual = Liste5
End Function

Function T4_Get_LinkNameToDisplay_ofEvFoldName(EvFoldName$) As String
    'Called from    T4_Add_EventLink_ToPersonFolders
    'EvFoldName     = "1986 DM"                  oder  "1984-09-15 LK D-GB Essen"
    'Action         Lieferung von "19860000 DM"  bzw.  "19840915 LK D-GB Essen"
    
    Dim E$, L$, s$
    E = EvFoldName: L = E     'L = LinkNameToDisplay
    If Len(E) < 10 Then       '"1986 DM" --> "19860000 DM"
        If Mid(L, 5, 1) = " " Then L = Left(L, 4) + "0000" + Mid(L, 5)
    Else
       'L =                                            "1984-09-15 LK D-GB"     "1984-09 LK D-GB"
        s = Left(L, 10)                               '"1984-09-15"             "1984-09 LK"
        s = Replace(s, "-", "")                       '"19840915"               "198409 LK"
        If Mid(s, 7, 1) = " " Then s = Left(s, 6) + "00" + Mid(s, 7)           '"19840900 LK"
        L = s + Mid(L, 11)                            '"19840915 LK D-GB"       "19840900 LK D-GB"
    End If
        L = Left(L, 9) + "! " + Mid(L, 10)            '"19840915 ..LK D-GB"     "19840900 ..LK D-GB"
    T4_Get_LinkNameToDisplay_ofEvFoldName = L
End Function





