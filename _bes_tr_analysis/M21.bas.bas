Attribute VB_Name = "M21"
Option Explicit

Sub zzz_M_21()
    showProcs "parent"
    'T2_ChangeTextBox2 5, 555
    EE 1: Beep
End Sub

Sub T2_BadLinks(B()) '<<<
    'Called from    T2_ButtonDoAll_wasClicked
    'Status         Box 2 ist angekreuzt: Repariere alle Links (broken .lnk files)
    'B()            = B(BoxNr, BoxItem) = B(1 to 6, 1 to 5) 'see T2_Load_BoxItems
    'BoxItems       z, s, BoxSelected, ZzAnzahl, ZzVisible
    
    'Vorbereitung
        Dim D$, OneFolderPath$, pSourceFolders$, PathOfLink$, Report$, s$, v$
        Dim a3&, a4&, a5&, a6&, a8&, i&, j&, A() As String, F() As String
        With Sheets("T2"): DoArrc: v = vbCrLf
        Report = "Report zu 'Repariere alle Links (broken .lnk files)'" + v _
        + "[Sub T2_BadLinks(B())]"
    'F()    'all folders to check all their subfolders
        If B(4, 3) = 1 Then pSourceFolders = pSourceFolders + ArrC(3) + "|"    'Events
        If B(5, 3) = 1 Then pSourceFolders = pSourceFolders + ArrC(11) + "|"   'ClubsNations
        If B(6, 3) = 1 Then pSourceFolders = pSourceFolders + ArrC(4)          'Leute
        Load_PathsOfAllSubfoldersAllLevels pSourceFolders, F, "*\zz*"          'Ausnahme \zz
        'show Join(F, v): Stop
    'Anzahl ermitteln/anzeigen   'alle Events-, ClubsNations-, Leute-Ordner
        T2_Handle_CountOfEventsClubsNationsLeuteFolders B, F, Report
    'TextFeld create    'mit korrekter Position   'alle Zahlen auf 0 gesetzt
        T2_TextFeld2_Create CInt(B(3, 1)) + 2
    'Alle Ordner sind gesammelt in F()
        For j = 1 To UBound(F)
            'Schleife über jeden Ordner der Ordnersammlung in F()
                'T2-Anzeige
                    .Cells(B(4, 1), 17) = j                     'Ordner j
                    .Cells(B(4, 1) + 1, 14) = NameOfPath(F(j))  'OrdnerName
                    .Cells(B(5, 1), 14) = UBound(F) - j         'CountDown
            OneFolderPath = F(j)
            'Pfade von .lnk-Files of one folder
                s = Get_AllFilePaths_LikeMyStringInFileName_OfOneFolder(OneFolderPath, "*.lnk")
                s = Delete_EmptyRowsInString(s)
                A = Split(s, v)     'A() holds all LinkFilePaths of actual Folder
            For i = 0 To UBound(A)
                'Schleife über jeden Link des Ordners j
                PathOfLink = A(i)
                D = Get_6DigitsAndInfo_OfLink(PathOfLink)
                ' = 111111|PathOfLink|exe|vbs|Ziel|ico|vbsLink = 6 Digits + 6 Info-items
                'BadLink is detected
                    If Left(D, 6) = "111111" Then
                        'Good Link      'count  'T2-Anzeige in T2Textfeld2
                            a8 = a8 + 1: T2_ChangeTextBox2 8, a8 'Zeile 8, Good Links total
                    Else
                        'Bad Link       'count  'T2-Anzeige
                            a3 = a3 + 1: T2_ChangeTextBox2 3, a3 'Zeile 3, Bad Links found
                            a6 = a6 + 1: T2_ChangeTextBox2 6, a6 'Zeile 6,       still bad
                            Report = Report + "BadLink " + CStr(a3) + ": " + D + v
                        'BadLink löschen, falls GoodLinkSameName existiert
                            T2_BadLink_SameName A, PathOfLink, a4, a6, Report
                        'BadLink reparieren
                            T2_BadLink_Repair D, a5, a6, Report
                    End If
                DoEvents
            Next i
            'Vorzeitig beenden
                'j = UBound(F)
        Next j
    'REPORT BadLinks
        End With
        show Report
        PepUp_BadLinksReport Report
    Call Beep ':     Stop
        
End Sub

Sub T2_BadLink_Repair(D$, a5&, a6&, Report$)
    'Called from    T2_BadLinks
    'D              = 111111|PathOfLink|exe|vbs|Ziel|ico|vbsLink = 6 Digits + 6 Info-items
    'a5             = aktueller Wert für Bad Links - repaired  (in T2TextFeld3, Zeile 5)
    'a6             = aktueller Wert für Bad Links - still bad (in T2TextFeld3, Zeile 6)
    'Action         Ermitteln des Link-Typs des OneBadLink; Typ-Reperatur
    
    'Vorbereitung
        Dim D2$, PF$, NL$, PL$, PathOfFolder$, Typ%, A() As String
        A = Split(D, "|")
        PL = A(1)                       'PL = PathOfLink = PathOfBadLink
        NL = NameOfPath(PL)             'NL = NameOfLink
        PF = Replace(PL, "\" + NL, "")  'PF = PathOfFolder
'    'Show while testing
'        PathOfFolder = Get_PathOfParentFolder(PL)
'        OpenFolder_BigIcons_SelectFile PL: Stop
    'Link-Typ ermitteln        '1 Person, 2 Event, 3 Club, 4 LTV, 5 Nation 6 Video 9 Sonst
        Typ = Get_LinkTypNr_1to9(PL)
    'Repair
        Select Case Typ
            Case 1: T2_BadLink_Repair_Typ1_Person PL
        
show Report: Stop
            
            Case 2: Repair_VbsLinks_OneFolder_LinkTyp2_Events PF, NL + "|", ""
            Case 3: Repair_VbsLinks_OneFolder_LinkTyp3_Club PF, NL + "|", ""
            Case 4: Repair_VbsLinks_OneFolder_LinkTyp4_LTV PF, NL + "|", ""
            Case 5: Repair_VbsLinks_OneFolder_LinkTyp5_Nation PF, NL + "|", ""
            Case 6: Repair_VbsLinks_OneFolder_LinkTyp6_Video PF, NL + "|", ""
            Case 9: Stop
        End Select
        
'        'Check, ob der Link jetzt OK ist
'                'NameOfLink ggf. geändert
'            D2 = Get_6DigitsAndInfo_OfLink(PL)
'            If Left(D, 6) = "111111" Then
'                'BadLink wurde erfolgreich repariert
                a5 = a5 + 1: T2_ChangeTextBox2 5, a5 'Zeile 6,   Bad Links repaired
                a6 = a6 - 1: T2_ChangeTextBox2 6, a6 'Zeile 6,   Bad Links still bad
'            Else
'                'BadLink ist nicht repariert
'                Stop
'            End If
'        Stop: CloseFolder PathOfFolder
    
'    If D Like "100***" Then T2_BadLink_Repair_100xxx LinkInfo, a5, a6
'    If D Like "11100*" Then T2_BadLink_Repair_11100x LinkInfo, a5, a6
'    If D Like "11101*" Then T2_BadLink_Repair_11101x LinkInfo, a5, a6
    
    'show Get_InfoOfOneLink_inOneLine(PL) + vbCrLf + vbCrLf + Get_InfoOfOneLink(PL)

End Sub

Sub T2_BadLink_Repair_Typ1_Person(PathOfLink$)
    'Called from    T2_BadLink_Repair
    'Typ1           Link, jumping to a PersonFolder
    'PathOfLink     Path of one BadLink
    'NameOfLink     Name of one BadLink; wird ggf. geändert
    'Action         ersetzt den BadLink durch einen GoodLink
    
    'Vorbereitung
        Dim D$, MyNameOfGoodLink$, MyPathOfGoodLink$, NI$, Nn$
        Dim PathOfLinkFolder$, PersFoldName$, v$, Vn$
        v = vbCrLf
        NI = Fill_NameItems_GiveNameOrPathOfLink(PathOfLink, Nn, Vn, PersFoldName)
    'Create
        Create_OneVbsLink_Person PathOfLink
    'Path of created Link should be MyPathOfGoodLink
        PathOfLinkFolder = Get_PathOfParentFolder(PathOfLink)
        MyPathOfGoodLink = PathOfLinkFolder + "\" + PersFoldName + ".lnk" 'SOLL-Name of GoodLink
        MyNameOfGoodLink = NameOfPath(MyPathOfGoodLink)
        If FileExists(MyPathOfGoodLink) Then
            D = Get_6DigitsAndInfo_OfLink(MyPathOfGoodLink)
            If Left(D, 6) = "111111" Then
                a5 = a5 + 1: T2_ChangeTextBox2 5, a5 'Zeile 6,   Bad Links repaired
                a6 = a6 - 1: T2_ChangeTextBox2 6, a6 'Zeile 6,   Bad Links still bad
                Report = Report + "   Link was changed to GoodLink: " + MyNameOfGoodLink + v
            Else
                Report = Report + "   Link was changed to BadLink: " + MyNameOfGoodLink + v
            End If
        Else
            Stop
        End If
    
End Sub

Sub Fill_NameItems_GiveNameOrPathOfLink(PathOfLink$, Nn$, Vn$, PersFoldName$)
    Fill_Nn_Vn_FromPersLink PathOfLink, Nn, Vn
    If Nn = "" Then Exit Sub
    PersFoldName = T5_Get_NameOfPersonfolder_Give_VnNn(Vn + " " + Nn)
End Sub

Function Get_6DigitsAndInfo_OfLink(PathOfLink$) As String
    'Called from    T2_BadLinks
    'Returns        111111|PathOfLink|exe|vbs|Ziel|ico|vbsLink = 6 Digits + 6 Info-items
    
    'Vorbereitung
        Dim D$, LinkInfo$, G() As String
    'Action
            LinkInfo = Get_InfoOfOneLink_inOneLine(PathOfLink) '6 items
        ' = "..\Events\..\Mai, Lea.lnk|..exe|..vbs|..\Leute\Mai, Lea (CH)|..ico|VbsLink"
        'Test one Link
            G = Split(LinkInfo, "|")
        'G(1) = LinkPfad, existiert (gerade eingelesen)
            If FileExists(G(1)) Then D = "1" Else D = "0"
        'G(2) = exe
            If G(2) = "C:\Windows\System32\wscript.exe" Then D = D + "1" Else D = D + "0"
        'G(3) = zz.vbs
            If G(3) = "F:\Archiv Trampolin 1900-1999\Leute\zzico\zz.vbs" Then D = D + "1" Else D = D + "0"
        'G(4) = SprungZiel, sollte existieren
            If Not FolderExists("F:\" + G(4)) Then
                If Not FileExists("F:\" + G(4)) Then D = D + "0" Else D = D + "1"
            Else
                D = D + "1"
            End If
        'G(5) = ico, sollte existieren
            If FileExists(G(5)) Then D = D + "1" Else D = D + "0"
        'G(6) = Description
            If G(6) = "VbsLink" Then D = D + "1" Else D = D + "0"
                    
    Get_6DigitsAndInfo_OfLink = D + LinkInfo
End Function

Sub T2_ClickInto_RemarksArea(z%, B())
    'Called from    Worksheet_SelectionChange
    'RemarksArea    = Spalten 27-34
    'z              = ZeilenNr der Zeile, in die der User klickte
    'Box            z. Zt. insgesamt 6 ankreuzbare Kästchen; Infos dazu in B()
    'B()            = B(i, j) = B(1 to 6, 1 to 5) 'siehe T2_Load_BoxItems
    '               j = (z,s,BoxAngekreuzt,ZzAnzahl,ZzVisible)
    'Action         blendet ZusatzZeilen ein/aus
    
    'Box 2
        If z = B(2, 1) Then         'Namen einer Person ...
            If B(2, 3) = 0 Then     'Box ist nicht angekreuzt
                If B(2, 5) = 1 Then 'ZusatzZeilen sind sichtbar
                    Rows(CStr(z + 1) + ":" + CStr(z + 8)).Hidden = True
                    B(2, 5) = 0
                Else
                    Rows(CStr(z + 1) + ":" + CStr(z + 8)).Hidden = False
                    B(2, 5) = 1: T2_ChangeNameTextfeld_Show
                End If
            End If
        End If
    'Box 3
        If z = B(3, 1) Then         'Repariere alle Links ...
            If B(3, 3) = 0 Then     'Box ist nicht angekreuzt
                If B(3, 5) = 1 Then 'ZusatzZeilen sind sichtbar
                    Rows(CStr(z + 1) + ":" + CStr(z + 8)).Hidden = True
                    B(3, 5) = 0
                Else
                    Rows(CStr(z + 1) + ":" + CStr(z + 8)).Hidden = False
                    B(3, 5) = 1: T2_ChangeNameTextfeld_Show
                End If
            End If
        End If
    EE 0: Sheets("T2").Cells(z, 26).Select: EE 1 'Select ändern
End Sub

Sub ChangeNameTextfeld_wasClicked()
    If Not T2_ButtonDoAll_isRed Then
        T2_ZusatzZeilenAusblenden_ChangeName
        Sheets("T2").Cells(T2_Get_z_ofBox_ChangeNameGlobal, 26).Select
        Beep1
    End If
End Sub


Sub T2_ZusatzZeilenEinblenden_ChangeName()
    Dim z&
    If Not T2_ChangeNameGlobal_ZusatzZeilen_areVisible Then
        z = T2_Get_z_ofBox_ChangeNameGlobal
        Sheets("T2").Rows(CStr(z + 2) + ":" + CStr(z + 9)).Hidden = False
    End If
End Sub

Sub T2_ZusatzZeilenAusblenden_ChangeName()
    Dim z&
    If T2_ChangeNameGlobal_ZusatzZeilen_areVisible Then
        z = T2_Get_z_ofBox_ChangeNameGlobal
        Sheets("T2").Rows(CStr(z + 2) + ":" + CStr(z + 9)).Hidden = True
    End If
End Sub

Function T2_ChangeNameGlobal_ZusatzZeilen_areVisible() As Boolean
    'Called from    T2_ZusatzZeilenEinblenden_ChangeName
    
    Dim zN$, vis As Boolean
    zN = T2_Get_z_ofBox_ChangeNameGlobal
    If RowIsVisible("T2", zN + 2) Then vis = True Else vis = False
    T2_ChangeNameGlobal_ZusatzZeilen_areVisible = vis
End Function

Function RowIsVisible(NameOfSheet As String, NrOfRow As Long) As Boolean
    'Called from    T2_ChangeNameGlobal_ZusatzZeilen_areVisible
    
    'Vorbereitung
        On Error GoTo ErrorHandler
        Dim targetSheet As Worksheet
        Set targetSheet = ThisWorkbook.Worksheets(NameOfSheet)
    'Returns True if the row is visible, False if hidden
        RowIsVisible = Not targetSheet.Rows(NrOfRow).Hidden
        Exit Function
ErrorHandler:
    'Returns False if sheet name or row number is invalid
        RowIsVisible = False
End Function

Function T2_Get_z_ofBox_ChangeNameGlobal() As String
    T2_Get_z_ofBox_ChangeNameGlobal = Get_RowNr_HoldingMyTextWhole("T2", "Namen einer Person ändern (global) ...")
End Function

Sub T2_ChangesInT4_NameOfPersonGlobal(z&, VnOld$, NnOld$, VnNew$, NnNew$, cc&, Coo$)
    'Called from    T2_ChangeNamePersGlobal
    'cc             wird hier gefüllt 'CountOfChanges
    'Coo            wird hier gefüllt 'Koordinaten der T4-VnNnOld-Zellen
    
    'Vorbereitung
        Dim NewContent$, OldContent$, zs$, ccT4&, i&, A() As String, B() As String
        T2_Anzeige_ChangeNameTextfeld_Item "5 T4:|0|NameOldCells": DoEvents '0 Namen gefunden
    'Zell-Koordinaten der Fundstellen, die NameOld enthalten
        Coo = T2_Get_AllCellCoo_ofT4CellsWithVnNn(VnOld, NnOld) '16,42|42,74|...|14,731
    'cc füllen
        A = Split(Coo, "|"): ccT4 = UBound(A) + 1: With Sheets("T4")
    'Schleife über alle NameOld-Fundstellen
        For i = 0 To UBound(A)
            zs = CStr(A(i)) 'Koordinaten der T4-VnNnOld-Zelle (z,s)   '16,42
            B = Split(zs, ",")
            'Zellinhalt (NameOld) ersetzen durch NameNew
                OldContent = .Cells(CLng(B(0)), CLng(B(1))).Value 'Lea Mai 'Lea Mai (17)
                NewContent = Replace(OldContent, VnOld + " " + NnOld, VnNew + " " + NnNew)
                .Cells(CLng(B(0)), CLng(B(1))) = NewContent
        Next
    'Finals
        End With 'T4
        T2_Anzeige_ChangeNameTextfeld_Item "5.2 " + CStr(ccT4) 'ccT4 Namen gefunden
        T2_Anzeige_ChangeNamePersonGlobal_RemarkAddToCC z, cc, ccT4
End Sub

Function T2_Get_AllCellCoo_ofT4CellsWithVnNn(Vn$, Nn$) As String
    'Called from    xxx
    'Action         Durchsucht alle T4-NamenZellen nach VnNn/VnNn + " (";
    '               findet "Lea Mai" und "Lea Mai (17)";
    '               liefert T4-Koordinaten der Fundstellen '16,42|42,74|...|14,731
    
    'Vorbereitung
        Dim ColNrs$, Coo$, N$, i&, s&, z&, A() As String, B()
    'Spalten-Nrn von T4-NamenSpalten ("W", "M" oder "X" in T4-Zeile2)
        ColNrs = Get_T4NameColumnNrs '22|31|35|...|744
        A = Split(ColNrs, "|")
    'Suche
        For i = 0 To UBound(A)
            s = CLng(A(i))
            'Array der Spalte s
                With Sheets("T4")
                B = .Range(.Cells(1, s), .Cells(Get_NrOfLastRowInColumnNr(s, "T4"), s)).Value
                End With
                For z = 1 To UBound(B, 1)
                    N = B(z, 1) 'CellContent (z,s) 'oft VnNn
                    If N = Vn + " " + Nn Then Coo = Coo + "|" + CStr(z) + "," + CStr(s)
                    If N Like Vn + " " + Nn + " (*" Then Coo = Coo + "|" + CStr(z) + "," + CStr(s)
                Next
        Next
    'Finals
        T2_Get_AllCellCoo_ofT4CellsWithVnNn = Mid(Coo, 2) '16,42|42,74|...|14,731
End Function

Function Get_T4NameColumnNrs() As String
    'Called from    T2_Get_AllCellCoo_ofT4CellsWithVnNn
    'Action         liefert alle Nrn von T4-NamenSpalten
    
    'Vorbereitung
        Dim s$, i&, sLast&, A()
    'T4-Zeile 2 laden
        sLast = Get_NrOfLastColumnInRowNr("T4", 2)
        With Sheets("T4")
            A = .Range(.Cells(2, 1), .Cells(2, sLast)).Value
        End With
    'Suche
        For i = 1 To UBound(A, 2)
            If A(1, i) = "W" Or A(1, i) = "M" Or A(1, i) = "X" Then s = s + "|" + CStr(i)
        Next
        Get_T4NameColumnNrs = Mid(s, 2)
End Function





Sub T2_ChangeNameTextfeld_Text_TEST()
    Dim MyText As String
    MyText = "##Änderungen##|hilfe"
    MyText = "Änderungen##" _
           + "T5:|5|Zellinhalte#Leute:|1|Ordner-Name#T4:|0|Namen gefunden#" _
           + "Icons:|22|umbenannt#|33|gelöscht#Links:|44|neu erzeugt|#|55|gelöscht"
    Call T2_ChangeNameTextfeld_Text(MyText)
End Sub

Sub T2_ChangeNameTextfeld_Text(ByVal MyText As String)
    'MyText         = Text, der in das Textfeld "ChangeNameTextfeld" geschrieben wird
    '               = "Text1:|5|Stichwort1#Text2:|22|Stichwort2#..."
    
    'Vorbereitung
        Dim formattedText$, shp As shape, tf As TextFrame2, Tab1#, Tab2#
    'Tabstopps definieren
        Tab1 = 1.4  '1. Tabstopp (erstes  '|'): Rechtsbündig bei 1,4 cm
        Tab2 = 1.6  '2. Tabstopp (zweites '|'): Linksbündig  bei 1,6 cm
    '1. Raute (#) durch Zeilenumbrüche und Striche (|) durch Tabulatoren ersetzen
        formattedText = Replace(MyText, "#", vbCrLf)
        formattedText = Replace(formattedText, "|", vbTab)
    '2. Textfeld auf Blatt T2 ansprechen
        Set shp = Worksheets("T2").Shapes("ChangeNameTextfeld")
        Set tf = shp.TextFrame2
    '3. Zuerst den Text zuweisen, damit das TextRange-Objekt vollständig initialisiert ist
        tf.TextRange.text = formattedText
    '4. Bestehende Tabstopps löschen (.TabStops statt .TabStops2 verwenden)
        With tf.TextRange.ParagraphFormat
            Do While .TabStops.count > 0: .TabStops(1).Clear: Loop
            .TabStops.Add Type:=msoTabStopRight, Position:=Application.CentimetersToPoints(Tab1)
            .TabStops.Add Type:=msoTabStopLeft, Position:=Application.CentimetersToPoints(Tab2)
        End With
End Sub

Sub T2_ChangeNameTextfeld_Hide()
    Worksheets("T2").Shapes("ChangeNameTextfeld").Visible = msoFalse
End Sub

Sub T2_ChangeNameTextfeld_Show()
    Dim zN%, L!, T!, w!, H!
    'ZeilenNr der ChangeNameBox
        zN = Get_RowNr_HoldingMyTextWhole("T2", "Namen einer Person ändern (global) ...")
    'Platzierung
        With Sheets("T2"): L = .Cells(zN + 2, 27).Left: T = .Cells(zN + 2, 27).Top
        w = .Cells(zN + 2, 35).Left - L: H = .Cells(zN + 9, 27).Top - T: End With
        T2_ChangeNameTextfeld_Set_LTWH L, T, w, H
    'Show
        Worksheets("T2").Shapes("ChangeNameTextfeld").Visible = msoTrue
End Sub

Sub T2_ChangeNameTextfeld_Set_LTWH(ByVal L!, ByVal T!, ByVal w!, ByVal H!)
    With Worksheets("T2").Shapes("ChangeNameTextfeld")
        .Left = L: .Top = T: .Width = w: .Height = H
    End With
End Sub

Sub T2_ChangeNameTextfeld_ExistenzAbsichern()
    'Vorbereitung
        Dim ws As Worksheet, shp As shape, tf As TextFrame2
        Dim L!, T!, w!, H!
        Const SHAPE_NAME As String = "ChangeNameTextfeld"
    '1. Referenz auf Tabellenblatt T2 setzen
        On Error Resume Next: Set ws = Worksheets("T2"): On Error GoTo 0
        If ws Is Nothing Then
            MsgBox "Das Tabellenblatt 'T2' wurde nicht gefunden!", vbCritical: Exit Sub
        End If
    '2. Prüfen, ob das Textfeld bereits existiert
        On Error Resume Next: Set shp = ws.Shapes(SHAPE_NAME): On Error GoTo 0
    '3. Erzeugen, falls nicht vorhanden
        If Not shp Is Nothing Then Exit Sub
        'AddTextbox(Ausrichtung, Links, Oben, Breite, Höhe) in Punkten
            L = Cells(10, 27).Left: T = Cells(10, 27).Top
            w = Cells(10, 35).Left - L: H = Cells(17, 27).Top - T
            Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, L, T, w, H)
            shp.NAME = SHAPE_NAME
        'Grundlegende Eigenschaften einrichten
            Set tf = shp.TextFrame2
            With tf: .WordWrap = msoFalse
                .MarginLeft = Application.CentimetersToPoints(0.1): .MarginRight = Application.CentimetersToPoints(0.1)
                .MarginTop = Application.CentimetersToPoints(0.1): .MarginBottom = Application.CentimetersToPoints(0.1)
            End With
        'Schriftart festlegen
            With tf.TextRange.Font: .NAME = "Calibri": .size = 8: End With
        'Beispieltext einfügen
            'tf.TextRange.text = "Textfeld erfolgreich in T2 erstellt."
End Sub

Sub ZellenVerbinden(NameOfSheet$, z1&, s1&, z2&, s2&)
    Application.DisplayAlerts = False 'Unterdrückt Warnmeldungen bezüglich Datenverlust beim Verbinden
    With Worksheets(NameOfSheet): With .Range(.Cells(z1, s1), .Cells(z2, s2))
        .Merge: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With: End With
    Application.DisplayAlerts = True
End Sub

Sub LeerzeilenEinfügen_unterhalb_LineX(NameOfSheet$, CountOfLines%, LineX%)
    'Vorbereitung
        If CountOfLines% < 1 Or LineX% < 1 Then Exit Sub
        Dim ws As Worksheet, calcState As XlCalculation, insertPos As Long
        Set ws = Worksheets(NameOfSheet$): insertPos = CLng(LineX%) + 1
    'Performance-Optimierung
        With Application: .ScreenUpdating = False: calcState = .Calculation: .Calculation = xlCalculationManual: End With
    '1. Neue Zeilen unterhalb von LineX einfügen
        ws.Rows(insertPos).Resize(CountOfLines%).Insert Shift:=xlDown
    '2. Nur die Hintergrundfarbe der neu eingefügten Zeilen entfernen
        ws.Rows(insertPos).Resize(CountOfLines%).Interior.ColorIndex = xlNone
    'Systemzustand wiederherstellen
        With Application: .Calculation = calcState: .ScreenUpdating = True: End With
End Sub

Sub Zeilenlöschen(NameOfSheet$, FromLine%, ToLine%)
    'Abfangen ungültiger Eingaben
        If FromLine% < 1 Or ToLine% < 1 Then Exit Sub
        If ToLine% < FromLine% Then Exit Sub
    'Vorbereitung
        Dim ws As Worksheet, calcState As XlCalculation, rowCount As Long
        Set ws = Worksheets(NameOfSheet$)
        rowCount = CLng(ToLine%) - CLng(FromLine%) + 1
    'Performance-Optimierung
        With Application: .ScreenUpdating = False: calcState = .Calculation: .Calculation = xlCalculationManual: End With
    'Zeilenbereich von FromLine im Umfang von rowCount Zeilen löschen
        ws.Rows(FromLine%).Resize(rowCount).Delete Shift:=xlUp
    'Systemzustand wiederherstellen
        With Application: .Calculation = calcState: .ScreenUpdating = True: End With
End Sub

Sub T2_DgNrnOben(z&)
    'Called from    T2_ButtonDoAll_wasClicked
    'Box            = angekreuztes Kästchen, das hier abgearbeitet wird
    'Box-Label      = "Fortlaufende Nummern oberhalb Dgs"
    'Action         Alle Dgs werden überprüft, ob sie oberhalb des Dg-Rahmens
    '               fortlaufende Nrn für die Dg-Spalten besitzen (1, 2, 3, ...)
    
    'Vorbereitung
        Dim cc&, i&, j&, k&, s1%, s2%, z1%, zLastT4&
        Dim r As Range, A(), B(), c() As Long
    'Anzeige in T2
        With Sheets("T2"): Ticks1   'Zeitnahme    'GetTickCount --> ArrC
        cc = 0: .Cells(z, 31) = cc  'Count of changes
        .Cells(z, 14) = ""          'Erledigt-Häkchen löschen
        .Cells(z, 4).Interior.Color = vbRed: End With 'Label rot
    'T4SomeDgData-Spalten 2,3,4 in A() laden
        zLastT4 = Get_NrOfLastRowInColumnNr(5, "T4")
        With Sheets("T4"): A = .Range(.Cells(8, 6), .Cells(zLastT4, 9)).Value
    'Alle Dgs prüfen
        For i = 1 To UBound(A, 1)
            'ZahlenReihe in B() laden
                z1 = A(i, 1): s1 = A(i, 2): s2 = A(i, 4)
                B = .Range(.Cells(z1 - 1, s1), .Cells(z1 - 1, s2)).Value
            'Reihe prüfen
                .Cells(19, 17) = "(" + CStr(z1) + ", " + CStr(s1) + ")"
                For j = 1 To UBound(B, 2)
                    If B(1, j) <> j Then
                        'Zahlenreihe des i-ten Dg unkorrekt
                        'korrekte Zahlenreihe erstellen
                            ReDim c(1 To UBound(B, 2))
                            For k = 1 To UBound(B, 2): c(k) = k: Next
                        'korrekte Zahlenreihe über i-tes Dg schreiben
                            Paste_1DArray_ToRow "T4", z1 - 1, s1, c
                        'Anzeige in T2: Count of changes
                            cc = cc + 1: Sheets("T2").Cells(z, 31) = cc
                        'For-Schleife beenden
                            j = UBound(B, 2)
                    End If
                Next
        Next
        End With
    'DoneRemarks to T2
        With Sheets("T2")
        'LastDone
            Set r = .Cells(z, 27)
            r.Value = Format(Now(), "yyyymmdd_hhmmss")
            r.Font.size = 6: r.HorizontalAlignment = xlLeft
        'Days ago
            Set r = .Cells(z, 29)
            r.Value = 0: r.Font.size = 8: r.HorizontalAlignment = xlCenter
        'Count of changes (Anzahl geänderter T5-Zeilen)
            Set r = .Cells(z, 31) 'r.Value = cc
            r.Font.size = 8: r.HorizontalAlignment = xlCenter
        'Erledigt-Häkchen setzen
            .Cells(z, 24) = "ü"
        'Duration
            .Cells(z, 33) = T6_GetDuration()
        'Label "Fortlaufende Nummern oberhalb Dgs" rot entfernen
            .Cells(z, 4).Interior.Color = Green1
        'Kästchen abkreuzen
            .Cells(z, 2).Select
        End With
End Sub

Sub T2_ChangeName_InputFields(zN%, z%, s%)
    'Called from    Worksheet_SelectionChange
    'zN             ZeilenNr der ChangeNameBox
    
    'Vorbereitung
        With Sheets("T2"): If .Cells(zN, 2) = "·" Then Exit Sub
    'ChangeName_InputFields sind sichtbar
        'mergedArea "ändern in"; bei Select: Weiterspringen
            If z = zN + 5 And s = 6 Then
                'mergedArea "ändern in" is selected (ggf. per ENTER in Zeile darüber)
                EE 0: .Cells(zN + 7, 6).Select: EE 1: Application.SendKeys "{F2}"
            End If
        'Ggf. Namen in Felder 3, 4 übertragen
            If .Cells(zN + 4, 6) <> "" And .Cells(zN + 7, 6) = "" Then .Cells(zN + 7, 6) = .Cells(zN + 4, 6).Value
            If .Cells(zN + 4, 14) <> "" And .Cells(zN + 7, 14) = "" Then .Cells(zN + 7, 14) = .Cells(zN + 4, 14).Value
        'Feld 1-4
            If (z = zN + 4 Or z = zN + 7) And (s = 6 Or s = 14) Then Application.SendKeys "{F2}"
    End With
End Sub

Sub Jump_FromT5_toT2_ChangeNameGlobal(zT5%)
    Dim Vn$, Nn$, zN%
    With Sheets("T5"): Nn = .Cells(zT5, 3): Vn = .Cells(zT5, 4): End With
    
    
    With Sheets("T2"): .Activate
    'ZeilenNr der ChangeNameBox ermitteln
        zN = Get_RowNr_HoldingMyTextWhole("T2", "Namen einer Person ändern (global) ...")
    'Vor- und Nachnamen in die 4 Input-Felder einsetzen
        .Cells(zN + 4, 6) = Vn: .Cells(zN + 7, 6) = Vn
        .Cells(zN + 4, 14) = Nn: .Cells(zN + 7, 14) = Nn
    'Kästchen ankreuzen (ZusatzZeilen einblenden)
        If .Cells(zN, 2) = "·" Then .Cells(zN, 2).Select
    End With
End Sub



Sub T2_Anzeige_ChangeNameTextfeld_Item(ByVal s As String)
    'Called from    T2_ChangeNamePersGlobal
    'Action         baut den Inhalt des T2-Textfeldes "ChangeNameTextfeld" auf
    '               T2_Anzeige_ChangeNameTextfeld_Item "3.1 T5:"            'Beginn Zeile 3
    '               T2_Anzeige_ChangeNameTextfeld_Item "3.2 7"              'Zeile 3, nach Tab1
    '               T2_Anzeige_ChangeNameTextfeld_Item "3.3 Zellinhalte"    'Zeile 3, nach Tab2
    '               oder direkt ganze Zeile:
    '               T2_Anzeige_ChangeNameTextfeld_Item "3 T5|5|Zellinhalte"
    '               erscheint dann im Textfeld in Zeile 3 als        T5:  7 Zellinhalte
    
    'Vorbereitung
        Dim parts() As String, lines() As String, lineParts() As String
        Dim cleanedLine$, currentText$, newValue$, i&, itemIdx&, lineIdx&
        parts = Split(Trim$(s), " ")
        If UBound(parts) < 0 Then Exit Sub
    'Neuer Wert setzen (falls kein Zusatztext übergeben wird -> leeres String "")
        If UBound(parts) >= 1 Then
            newValue = parts(1)
            For i = 2 To UBound(parts)
                newValue = newValue & " " & parts(i)
            Next i
        Else
            newValue = ""
        End If
    'Bisherigen Text laden
        currentText = ArrC(51)
    'Indizes & Zeilenverarbeitung
        If InStr(parts(0), ".") > 0 Then
            'Einzel-Item-Modus (z. B. "3.1 T5:")
            lineIdx = CLng(Split(parts(0), ".")(0))
            itemIdx = CLng(Split(parts(0), ".")(1)) - 1
            If itemIdx < 0 Then itemIdx = 0
            
            'Zeilen-Array laden oder erweitern
            If Len(currentText) = 0 Then
                ReDim lines(0 To lineIdx - 1)
            Else
                lines = Split(currentText, "#")
                If UBound(lines) < lineIdx - 1 Then ReDim Preserve lines(0 To lineIdx - 1)
            End If
            
            'Zielzeile aufteilen, Item setzen & leere Striche säubern
            lineParts = Split(lines(lineIdx - 1), "|")
            If UBound(lineParts) < itemIdx Then ReDim Preserve lineParts(0 To itemIdx)
            lineParts(itemIdx) = newValue
            cleanedLine = Join(lineParts, "|")
            Do While Right$(cleanedLine, 1) = "|"
                cleanedLine = Left$(cleanedLine, Len(cleanedLine) - 1)
            Loop
            lines(lineIdx - 1) = cleanedLine
        Else
            'Ganze-Zeile-Modus (z. B. "3 T5|5|Zellinhalte")
            lineIdx = CLng(parts(0))
            
            'Zeilen-Array laden oder erweitern
            If Len(currentText) = 0 Then
                ReDim lines(0 To lineIdx - 1)
            Else
                lines = Split(currentText, "#")
                If UBound(lines) < lineIdx - 1 Then ReDim Preserve lines(0 To lineIdx - 1)
            End If
            
            'Ganze Zeile direkt zuweisen
            lines(lineIdx - 1) = newValue
        End If
    'Text in ArrC zurückschreiben
        FillArrC 51, Join(lines, "#")
    'Textfeld aktualisieren
        Call T2_ChangeNameTextfeld_Text(ArrC(51))
End Sub

' <<<
Sub T2_ChangeNamePersGlobal(z&)
    'Called from    T2_ButtonDoAll_wasClicked
    'Box            = angekreuztes Kästchen, das hier abgearbeitet wird
    'Box-Label      = "Namen einer Person ändern (global) ..."
    'z              = ZeilenNr der Box
    'Changes        1) in T5:      Spalten 3,4,10,15,22
    '               2) in 'Leute': rename OldFolderName, NewFolderName
    '               3) in T4:      change each OldNameCell
    '               4) in 'Leute\zzico': xxx
    '               5) in 'Events': xxNewLinks
    
    'Vorbereitung
        Dim Coo$, NnNew$, NnOld$, VnNew$, VnOld$, NamePersFoldOld$, NamePersFoldNew$
        Dim i%, sT4%, zT4%, cc&
        T2_ChangeNamePersonGlobal_FirstHandlings z, cc: With Sheets("T2")
    'VnOld, NnOld, VnNew, NnNew aus T2-UserInputs übernehmen
        VnOld = .Cells(z + 4, 6): NnOld = .Cells(z + 4, 14)
        VnNew = .Cells(z + 7, 6): NnNew = .Cells(z + 7, 14): End With
    '1) Changes in T5 (PersonData)
        T2_ChangesInT5_NameOfPersonGlobal NnNew, NnOld, VnNew, VnOld, NamePersFoldOld, NamePersFoldNew
        If NamePersFoldOld = "" Then Stop
            T2_Anzeige_ChangeNameTextfeld_Item "3 T5:|5|Zellinhalte"
            T2_Anzeige_ChangeNamePersonGlobal_RemarkAddToCC z, cc, 5
    '2) Change name of folder in 'Leute'
        RenameFolder ArrC(4) + "\" + NamePersFoldOld, ArrC(4) + "\" + NamePersFoldNew
            T2_Anzeige_ChangeNameTextfeld_Item "4 Leute:|1|Ordner-Name"
            T2_Anzeige_ChangeNamePersonGlobal_RemarkAddToCC z, cc, 1
    '3) Changes in T4 'Alle Zellen mit NameOld ändern und Coo belegen
            T2_ChangesInT4_NameOfPersonGlobal z, VnOld, NnOld, VnNew, NnNew, cc, Coo
            'Coo wurde hier geliefert 'Koordinaten der OldNameZellen in div. Dgs
    '4) Change NameOfIcon
        T2_ChangeNameOfIcon_NameOfPersonGlobal z, cc&, NnOld, VnOld, NamePersFoldOld, NamePersFoldNew
    '5) Changes in 'Events': DeleteOldLinks CreateNewLinks
        T2_ChangesInEvents_NameOfPersonGlobal z, VnOld, NnOld, NamePersFoldNew, cc, Coo
    'Finals
        T2_FinalDoneRemarks_ChangeNamePersonGlobal z
End Sub

Sub T2_ChangeNameOfIcon_NameOfPersonGlobal(z&, cc&, NnOld$, VnOld$, NamePersFoldOld$, NamePersFoldNew$)
    'Called from    T2_ChangeNamePersGlobal
    
    'Vorbereitung
        T2_Anzeige_ChangeNameTextfeld_Item "6 Icon:|0|.ico files"
        If NamePersFoldOld = NamePersFoldNew Then Exit Sub
        Dim p$, pIconNew$, pIconOld$, ccIc&
    'Path of IconFolder
        p = ArrC(4) + "\zzico"
     'Path NameNew
        pIconOld = p + "\" + Replace(NamePersFoldOld, ",", "§") + ".ico"
        pIconNew = p + "\" + Replace(NamePersFoldNew, ",", "§") + ".ico"
    'Rename
        If FileExists(pIconNew) Then
            If FileExists(pIconOld) Then DeleteFile pIconOld: ccIc = ccIc + 1
        Else
            If FileExists(pIconOld) Then RenameFile pIconOld, pIconNew: ccIc = ccIc + 1
        End If
        T2_Anzeige_ChangeNameTextfeld_Item "6.2 " + CStr(ccIc)
        T2_Anzeige_ChangeNamePersonGlobal_RemarkAddToCC z, cc, ccIc
End Sub

Sub T2_ChangesInEvents_NameOfPersonGlobal(z&, VnOld$, NnOld$, NamePersFoldNew, cc&, Coo$)
    'Called from    T2_ChangeNamePersGlobal
    'Coo            Koordinaten der OldNameZellen in div. Dgs
    '               zum Auffinden der zugehörigen EventFolderNames
    'Action         NameOldLinks (in EventOrdnern) löschen, NameNewLinks erzeugen
    
    'Vorbereitung
        Dim EvFolNames$, L2_pJumpTo$, L3_pIcon$, N$, p$, PathOfNewLink$, ccEV&, i&, j&
        Dim A() As String, B() As String, NewLinkExists As Boolean
    'EventFolderNames
        EvFolNames = T2_Get_EventFolderNames_ofEachChangedDgNameCell(Coo)
        '          = 1969-08 Nissen-Cup11 Grenchen_CH '1966-05-08 DJM Nürnberg '...
        '   In jedem dieser EventOrdner liegen 1 od. 2 OldNameLinks: Ersetzen
        '   OldNameLinks enthalten NnOld, VnOld (...) 'Mai, Lea (...)
    'In allen gefundenen EventOrdnern suchen
        A = Split(EvFolNames, vbCrLf)
        For i = 0 To UBound(A)
            p = ArrC(3) + "\" + A(i)  'EventOrdnerPfad
            N = Get_AllFilePaths_LikeMyStringInFileName_OfOneFolder(p, "*" + NnOld + ", " + VnOld + "*")
            B = Split(N, vbCrLf)    'B() holds all file paths with NameOld
            T2_Anzeige_ChangeNameTextfeld_Item "7 Events:|0|OldLinks": ccEV = 0
            PathOfNewLink = ArrC(3) + "\" + A(i) + "\" + NamePersFoldNew + ".lnk"
            For j = 0 To UBound(B)
                'Schleife über alle OldName-Links im aktuellen EventOrdner
                If B(j) Like "*\" + NnOld + ", " + VnOld + ".lnk" _
                    Or B(j) Like "*\" + NnOld + ", " + VnOld + " (*" Then
                    'B(j) = "...\Mai, Lea.lnk" '= "...\Mai, Lea (TV Bern, CH).lnk
                    If B(j) = PathOfNewLink Then
                        NewLinkExists = True
                    Else
                        DeleteFile B(j): ccEV = ccEV + 1
                    End If
                End If
            Next
            If Not NewLinkExists Then
                'Create one new Link
                L2_pJumpTo = ArrC(4) + "\" + NamePersFoldNew
                L3_pIcon = ArrC(4) + "\zzico\" + Replace(NamePersFoldNew, ",", "§") + ".ico"
                Create_OneVbsLink PathOfNewLink, L2_pJumpTo, L3_pIcon
            End If
        Next
        T2_Anzeige_ChangeNameTextfeld_Item "7.2 " + CStr(ccEV)
        T2_Anzeige_ChangeNamePersonGlobal_RemarkAddToCC z, cc, ccEV
End Sub

Sub T2_Anzeige_ChangeNamePersonGlobal_RemarkAddToCC(z&, cc&, AddToCC&)
    'Eintrag in Zeile z, Spalte 31 (Count changes)
    cc = cc + AddToCC: Sheets("T2").Cells(z, 31) = cc: DoEvents
End Sub

Sub T2_ChangeNamePersonGlobal_FirstHandlings(z&, cc&)
    'Called from    T2_ChangeNamePersGlobal
    
    FillArrC 51, "" 'ChangeNameTextfeld-Inhalt      'Container leeren
    T2_Anzeige_ChangeNameTextfeld_Item "1.1 Änderungen"     'Textfeld neu belegen
    With Sheets("T2"): Ticks1                       'Zeitnahme 'GetTickCount --> ArrC
    cc = 0: .Cells(z, 31) = cc                      'Count of changes
    .Cells(z, 14) = ""                              'Erledigt-Häkchen löschen
    .Cells(z, 4).Interior.Color = vbRed             'Label rot
    T2_ChangeNameTextfeld_Show                      'Textfeld einblenden
    End With
End Sub

Sub T2_FinalDoneRemarks_ChangeNamePersonGlobal(z&)
    'Called from    T2_ChangeNamePersGlobal
    
    Dim r As Range: With Sheets("T2")
    'Label "Namen einer Person ändern (global) ..." wieder hellgrün
        .Cells(z, 4).Interior.Color = Green1
    'LastDone
        Set r = .Cells(z, 27)
        r.Value = Format(Now(), "yyyymmdd_hhmmss")
        r.Font.size = 6: r.HorizontalAlignment = xlLeft: r.VerticalAlignment = xlCenter
    'Days ago
        Set r = .Cells(z, 29)
        r.Value = 0: r.Font.size = 8: r.HorizontalAlignment = xlCenter
        ZellenVerbinden "T2", z, 29, z, 30
    'Count of changes (Anzahl geänderter T5-Zeilen)
        Set r = .Cells(z, 31) 'r.Value = cc 'cc wurde fortlaufend hochgezählt
        r.Font.size = 8: r.HorizontalAlignment = xlCenter
        ZellenVerbinden "T2", z, 31, z, 32
    'Erledigt-Häkchen setzen
        .Cells(z, 24) = "ü"
    'Duration
        Set r = .Cells(z, 33)
        r.Value = T6_GetDuration(): r.Font.size = 8
        ZellenVerbinden "T2", z, 33, z, 34
    'Label rot entfernen
        .Cells(z, 4).Interior.Color = Green1
    'Kästchen auf "·" stellen (nicht mit Select abkreuzen); Zusatzzeilen bleiben
        .Cells(z, 2) = "·": .Cells(z, 2).Font.NAME = "Calibri": .Cells(z, 2).Font.size = 11
    End With: Beep
End Sub

Sub T2_ChangesInT5_NameOfPersonGlobal(NnNew$, NnOld$, VnNew$, VnOld$, NamePersFoldOld$, NamePersFoldNew$)
    'Called from    T2_ChangeNamePersGlobal
        Dim sw$, zT5%
        zT5 = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, VnOld + " " + NnOld)
        If zT5 = 0 Then Exit Sub
        With Sheets("T5")
        'T5-Zeile 22:           Vn Nn
            .Cells(zT5, 22) = VnNew + " " + NnNew
        'T5-Zeile 10:           PersonenOrdner
            NamePersFoldOld = .Cells(zT5, 10)
            NamePersFoldNew = Replace(NamePersFoldOld, NnOld + ", " + VnOld, NnNew + ", " + VnNew)
            .Cells(zT5, 10) = NamePersFoldNew
        'T5-Zeile 3, Zeile 4:   Nachname, Vorname
            .Cells(zT5, 3) = NnNew: .Cells(zT5, 4) = VnNew
        'T5-Zeile 15:           Schreibweise
            sw = .Cells(zT5, 15)
            sw = Replace(sw, VnNew, ""): sw = Replace(sw, NnNew, "")
            If Not sw Like "*" + VnOld + "*" And VnOld <> VnNew Then sw = sw + " " + VnOld
            If Not sw Like "*" + NnOld + "*" And NnOld <> NnNew Then sw = sw + " " + NnOld
            sw = Replace99(sw, "  ", " "): .Cells(zT5, 15) = sw
'            'Bei Benutzung von NewName in einer Dg-Namenzelle:
'                'Ggf wurden Changes in T5 bereits teilweise durchgeführt,
'                'aber noch nicht im NewName auslösenden Dg, weshalb NameOld erneut in T5 landete
'                'und erneut in NameNew geändert wurde; NameNew existiert dadurch zweimal
'                    'zT5 = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, zT5, VnNew + " " + NnNew)
'                    'If zT5 > 0 Then .Range(.Cells(zT5, 3), .Cells(zT5, 22)) = ""
            End With
End Sub

Function T2_Get_EventFolderNames_ofEachChangedDgNameCell(Coo$)
    'Called from    T2_ChangeNamePersGlobal
    'Coo            = Koordinaten der T4-NameOld-Zellen '16,42|42,74|...|14,731
    
    'Vorbereitung
        Dim F$, zs$, v$, i&, s&, z&, A() As String, B() As String
        v = vbCrLf
    'Schleife über alle NameOld-Fundstellen
        A = Split(Coo, "|")
        For i = 0 To UBound(A)
            zs = CStr(A(i)) 'Koordinaten der T4-VnNnOld-Zelle (z,s)   '16,42
            B = Split(zs, ","): z = CLng(B(0)): s = CLng(B(1))
                'Dg(z,s) = Zelle, die NameOld enthielt und geändert wurde
            F = F + v + T4_Get_DgEventFolderName_by_zSh_sSh(z, s)
        Next
        'F = Folderliste '1969-08 Nissen-Cup11 Grenchen_CH '1966-05-08 DJM Nürnberg '...
    T2_Get_EventFolderNames_ofEachChangedDgNameCell = Mid(F, 3)
End Function

Function T2_ButtonDoAll_isRed() As Boolean
    'Called from    T2_KästchenAnAbkreuzen
    Dim r As Range: Set r = Range(Cells(2, 12), Cells(4, 17)):
    If r.Interior.Color = vbRed Then T2_ButtonDoAll_isRed = True
End Function


