Attribute VB_Name = "M_T7"
Option Explicit
    
Sub zzz_M_T7()
    Application.EnableEvents = True
    showProcs "with"
    'wh = Get_ImageWidthHeight(p) '001000 123456
End Sub

Sub T7_Check_TextBoxChange()
    'Called from    Worksheet_SelectionChange [T7; every UserClick]
    'helpers        Liste in T7, Spalten 46-54: |Nr|Nn|VnY|Age|Id|p0|TxtBox|-|
    'ArrC(85)       = aktuelle Id
    '               gefüllt von T7_Read_LbToDoFiles, T7_OneFieldLineWasClicked
    'ArrC(86)       = aktueller TextBoxContent Tbc
    '               gefüllt von T7_Read_LbToDoFiles, T7_OneFieldLineWasClicked
    'Action         schreibt geänderten TextBox-Inhalt in helpers
    
    'Vorbereitung
        Dim Age$, Id$, Nn$, P0$, Tbc1$, Tbc2$, Vn$, y$, Nr%, H9()
        With Sheets("T7"): DoArrc: Id = ArrC(85)
        Tbc1 = ArrC(86) '= "[Id:p0048-01]TV Bonn...§Event:...§Trainer..."
        Tbc2 = "[Id:" + Id + "]" + T7_Get_TextFromTextBoxOnSheetT7_AsOneLine
        If Tbc1 = Tbc2 Then Exit Sub
    'TextBoxContent has changed
        If Left(Tbc1, 13) <> Left(Tbc2, 13) Then Exit Sub 'Stop 'Ids verschieden
    'Actual helpersLine
        Nr = Get_RowNr_HoldingMyTextWholeInColumnX("T7", 51, 4, Id)
        If Nr = 0 Then Exit Sub
        Nn = .Cells(Nr, 47): Vn = .Cells(Nr, 48): y = .Cells(Nr, 49)
        Age = .Cells(Nr, 50): P0 = .Cells(Nr, 52)
    'Change TextBoxContent in actual helpers line
        .Cells(Nr, 53) = Tbc2
    'H9
        ReDim H9(1 To 1, 1 To 9)
        H9(1, 1) = Nr: H9(1, 2) = Nn: H9(1, 3) = Vn: H9(1, 4) = y
        H9(1, 5) = Age: H9(1, 6) = Id: H9(1, 7) = P0: H9(1, 8) = Tbc2: H9(1, 9) = "-"
    'LabelPhoto produzieren
        T7_Produce_PhotoWithLabel H9, 1
    'ArrC-Eintrag
        FillArrC 86, Tbc2
    'Neues LabelFoto aus LbProd in LbArchiv verschieben
        T7_Move_NewLabelPhoto_LbProd_LbArchiv H9, 1
    'Import
        T7_Import_Photo
        RefreshScreen
    'Finals
        End With
End Sub

Sub T7_Read_LbToDoFiles()
    'Called from    Worksheet_SelectionChange[T7]
    'Action         liest Dateien innerhalb des Ordners 'LbToDo'
    '               erstellt ein (neues) LabelPhoto zu jedem Link
    
    'Vorbereitung
        Dim FileNames$, F$, Fi$, Id$, P0$, p1$, Tb$, v$, i%, ArrF() As String, H9()
        v = vbCrLf: DoArrc
        p1 = ArrC(1) + "\prog\Label\LbToDo\"
    'Files of LbToDoFolder
        FileNames = Get_AllFileNamesOfOneFolder(p1)
        If FileNames = "" Then Exit Sub
    'Each file
        ArrF = Split(FileNames, vbCrLf)
        For i = 0 To UBound(ArrF)
            Fi = ArrF(i)     'one FileName (.lnk)
            If T7_ToDoFileNameIsCorrect(Fi) Then F = F + v + Fi
        Next
        If F = "" Then Exit Sub Else F = Mid(F, 3)
        'f = some good LnkFileNames
    'Nn Vn Y Age Id p0 Tb   '9HelpersData in each H9-Line
        T7_Fill_H9 F, H9    'f holds n LbToDoLnkFileNames 'H9(1 to n, 1 to 9)
    'Produce LabelPhotos and show each
        For i = 1 To UBound(H9, 1)
        
        
'            'p0 control
'                p0 = H9(i, 7) '
'                showArray2D H9, "i = " + CStr(i) + v + "p0 = " + p0 + v
                
                
            'LabelPhoto produzieren
                T7_Produce_PhotoWithLabel H9, i
            'Neues LabelPhoto zum ActivePhoto machen
                Id = H9(i, 6)
            'ArrC-Eintrag
                FillArrC 85, Id
                FillArrC 86, T7_Get_TextBoxContent_FromHelpers(Id)
            'Neues LabelFoto aus LbProd in LbArchiv verschieben
                T7_Move_NewLabelPhoto_LbProd_LbArchiv H9, i
            'Neue DataZeile zu helpers hinzufügen
                T7_Add_NewDataToHelpers H9, i
                T7_Fill_Helpers
            'Import
                T7_Import_Photo
                RefreshScreen
            'TextBox füllen
                Tb = CStr(H9(i, 8))
                T7_Fill_TextBoxOnSheetT7 Tb
            'Scroll to see Name in VisibleScrollLines
                T7_Scroll_ToActiveFieldLine
                T7_Highlight_ActiveFieldLine
            'Delete_OneLnkFile_inside LbToDo-Folder
                T7_Delete_LinkFileOfActivePhoto
                RefreshScreen
            'Wait
                WaitSecs 3
JumpToNext:
        Next
End Sub

Sub T7_Scroll_ToActiveFieldLine()
    Dim Id$, c%, cH%, Nr%, zLast%, L(): DoArrc
    zLast = Get_NrOfLastRowInColumnNr(47, "T7") 'helpers
    'count of VisibleScrollLines = 14
    cH = zLast - 4  '29 'count of HelpersLines
    Sheets("T7").ScrollBar1.max = cH
    Id = ArrC(85)
    Nr = Get_RowNr_HoldingMyTextWholeInColumnX("T7", 51, 4, Id) - 4
    c = Nr - 5      'Mid of VisibleScrollLines
    If c < 1 Then c = 1
    If c > cH - 13 Then c = cH - 13
    Sheets("T7").ScrollBar1.Value = c
End Sub

Sub T7_Fill_VisibleScrollLines()
    'Called from    ScrollBar1_Change
    
    'Vorbereitung
        Dim c%, cH%, cL%, zLast%, L()
        zLast = Get_NrOfLastRowInColumnNr(47, "T7") 'helpers
        With Sheets("T7")
    'ScrollBar
        cL = 14         '14 'count of VisibleScrollLines
        cH = zLast - 4  '29 'count of HelpersLines
        With .ScrollBar1: c = .Value: .max = cH - cL + 1: End With
    'L  = Array of VisibleScrollLines; 14 helpers-Zeilen ab aktuellem ScrollBar1.Value
        L = .Range(.Cells(4 + c, 46), .Cells(4 + c + cL - 1, 51)).Value
        Paste_2DArrayToCell_z_s "T7", 13, 3, L
        EE 0: .Cells(12, 2).Select: EE 1
        End With
    T7_Highlight_ActiveFieldLine
End Sub

Sub T7_Fill_Helpers()
    'Called from    T7_Read_LbToDoFiles
    'helpers        T7; Liste in Spalten 46-54 (graue Zellen);
    '               trägt die Daten aller LbArchiv-jpgs
    'Action         FileNames der LbArchiv-jpgs werden geladen (sortiert);
    '               pro jpg wird die 'helpers'-Zeile übernommen
    
    'Vorbereitung
        Dim FN$, Id$, i%, j%, Nr%, H(), HPaste(), L()
        Call DoArrc
    'Load helpers               'T7-Range     2D-Array
        T7_Load_ArrHelpers H    'ggf. doppelte Zeilen
    'Load LbArchiv              'FileNames    1D-Array, Base 1
        T7_Load_ArrLbArchiv L
    'PastArray erstellen        'für helpers new
        ReDim HPaste(1 To UBound(L), 1 To 9)
        For i = 1 To UBound(L)
            'Schleife über alle NameOfLnkFile inside LbArchiv
            FN = L(i) 'OneFileName in LbArchiv 'Mai, Lea (1985) p1234-01.jpg
            Id = Mid(FN, Len(FN) - 11, 8)
            Nr = Get_RowNr_HoldingMyTextWholeInColumnX("T7", 51, 4, Id) - 4
            'HPaste-Array füllen (new helpers)
                HPaste(i, 1) = i
                For j = 2 To 9
                    HPaste(i, j) = H(Nr, j)
                Next
        Next
    'Write new helpers
        T7_ClearHelpers
        Paste_2DArrayToCell_z_s "T7", 5, 46, HPaste
        T7_GreyHelpers
    'AnzHelpers
        Sheets("T7").[D12] = "Label-Photos in LbArchiv: " + CStr(UBound(L))
End Sub

Sub T7_Delete_LinkFileOfActivePhoto()
    'Called from    T7_Read_LbToDoFiles
    
    Dim Id$, p1$, s$, i%, Arr1() As String
    p1 = ArrC(1) + "\prog\Label\LbToDo\"
    Id = ArrC(85)
    s = Get_AllFileNames_Like_OfOneFolder(p1, "*" + Id + "*.lnk")
    Arr1 = Split(s, vbCrLf)
    DeleteFile p1 + Arr1(0)
End Sub

Sub T7_Fill_TextBoxOnSheetT7(Tb$)
    'Called from    T7_Read_LbToDoFiles
    'TextBoxContent [Id:p1234-01]TV Bonn ...§...§...
    
    Dim T$
    T = Mid(Tb, 14)
    T = Replace(T, "§", vbCrLf): T = Delete_EmptyRowsInString(T)
    T7_Put_TextIntoTextbox T
End Sub

Sub T7_Add_NewDataToHelpers(ByRef H9, Nr%)
    'Called from    T7_Read_LbToDoFiles
    'H9             = Array H9(1 to n, 1 to 9), 9 HelpersData per each H9-Line i
    '               nr Nn Vn Y Age Id p0 Tb -
    
    Dim z%: With Sheets("T7")
    'Nächste leere HelpersZeile
        z = Get_RowNr_HoldingMyTextWholeInColumnX("T7", 47, 4, "")
    'Zeile füllen
        .Cells(z, 46) = CInt(.Cells(z - 1, 46)) + 1
        .Cells(z, 47) = H9(Nr, 2)
        .Cells(z, 48) = H9(Nr, 3)
        .Cells(z, 49) = H9(Nr, 4)
        .Cells(z, 50) = H9(Nr, 5)
        .Cells(z, 51) = H9(Nr, 6)
        .Cells(z, 52) = H9(Nr, 7)
        .Cells(z, 53) = H9(Nr, 8)
        .Cells(z, 54) = "-"
    End With
End Sub

Sub T7_Move_NewLabelPhoto_LbProd_LbArchiv(ByRef H9, Nr%)
    'Called from    T7_Read_LbToDoFiles
    'H9             = Array H9(1 to n, 1 to 9), 9 HelpersData per each H9-Line i
    '               nr Nn Vn Y Age Id p0 Tb -
    
    Dim Id$, Nn$, p2$, p3$, Vn$, y$
    Id = H9(Nr, 6): Nn = H9(Nr, 2): Vn = H9(Nr, 3): y = H9(Nr, 4)
    p2 = ArrC(1) + "\prog\Label\LbProd\" + Id + "_i15.jpg"
    p3 = ArrC(1) + "\prog\Label\LbArchiv\" + Nn + ", " + Vn + " (" + y + ") " + Id + ".jpg"
    DeleteFile p3 'alte Version
    RenameFile p2, p3
End Sub

Sub T7_Produce_PhotoWithLabel(ByRef H9, Nr%)
    'Called from    T7_Read_LbToDoFiles
    'H9             = Array H9(1 to n, 1 to 9), 9 HelpersData per each H9-Line i
    '               nr Nn Vn Y Age Id p0 Tb -
    'Tb             = TextBoxContent
    '               = "[Id:p1164-01]['Pip', vh Maier]TV Brr (Pfalz)§Event: 1972...§Trainer§Line4"
    
    'Vorbereitung
        Dim Age$, Id$, L1$, L2$, L3$, L4$, N$, P0$, Tb$, y$
        Dim i%, Li%, Arr1() As String
        DoArrc
    'Alte LbProd-Files löschen
        Kill_AllFilesInOneFolder T7_Get_PathOfLbProd
    'L1 L2 L3 L4    (Label-ZusatzZeilen)
        Tb = Mid(H9(Nr, 8), 14)
        Arr1 = Split("§" + Tb + "§§§§", "§")
        L1 = Arr1(1): L2 = Arr1(2): L3 = Arr1(3): L4 = Arr1(4): Li = 4
        If L4 = "" Then Li = 3
        If L3 = "" Then Li = 2
    'N Y Age Id p0
        N = H9(Nr, 3) + " " + H9(Nr, 2): y = H9(Nr, 4): Age = H9(Nr, 5)
        Id = H9(Nr, 6): P0 = H9(Nr, 7)
    'Create
        Create_i14 Id, P0           'F1            Foto (1000xH)   .jpg+
        Create_i01 Id, N            'NamePerson    transparent     .png
        Create_i02 Id, y            'YearOfPhoto   transparent     .png
        Create_i03 Id, Age          'AgeOfPerson   transparent     .png
        Create_i04 Id, L1, "i04"    'TextLine1     transparent     .png
        Create_i05 Id, L2, "i05"    'TextLine2     transparent     .png
        Create_i06 Id, L3, "i06"    'TextLine3     transparent     .png    existiert ggf. nicht
        Create_i07 Id, L4, "i07"    'TextLine4     transparent     .png    existiert ggf. nicht
        Create_i08 Id               'i04+.+i07     TextEndversion  .png
        Create_i09 Id               'L LabelLeer   BackgroundWheat .jpg
        Create_i10 Id               'LN            L + Vn Nn
        Create_i11 Id               'i09+i01       LabelLeer +Name .jpg
        Create_i12 Id               'i11+i03       LNY       +Age  .jpg+
        Create_i13 Id, Li           'i12+i08       FertigesLabel   .jpg
        Create_i15 Id               'i14+i13       Foto+Label
        Beep
End Sub

Sub T7_Fill_H9(F$, ByRef H9)
    'Called from    T7_Read_LbToDoFiles
    'f              = n LbToDoLnkFileNames
                        '19720913 DTS a76 (Ute Czech) p1178-01 ub.jpg.lnk
                        '19851003 LK D-POL-UDSSR Aschendorf r3 (Roland Berger) v0037-01 nn.jpg.lnk
                        '20130525 DTF ehemalige r2 (Roland Berger) p0119-03 nn.jpg.lnk
                        '20240921_150338 OldiesTreff Kirchheim a2 (Kurt Schmidt) p0902-01 ub.jpg.lnk
    'H9             = Array H9(1 to n, 1 to 9); wird hier gefüllt
    '               9 HelpersData per each f-Link/H9-Line: nr Nn Vn Y Age Id p0 Tb -
    
    'Vorbereitung
        Dim Age$, E1$, E2$, Id$, L$, Nn$, P0$, p1$, p3$, Tb$, v$, Vn$, VnNn$, y$
        Dim i%, N%, ArrF() As String
        v = vbCrLf: ArrF = Split(v + F, v)
        N = UBound(ArrF)          'n = Anzahl LbToDoLnkFileNames
        ReDim H9(1 To N, 1 To 9)
        p1 = ArrC(1) + "\prog\Label\LbToDo\"
    'Get Nn Vn Y Age Id p0 Tb
        For i = 1 To N
            L = ArrF(i)     'one LnkFileName
            'p3 = Pfad zu Datei, von der ein Link in LbToDo gelegt wurde
               p3 = Get_LinkTargetPath(p1 + L)     'lnk points to p3; ggf. in 'Leute'
            Id = Get_IdFromPath(p3)                 'p1234-05
                VnNn = T7_Get_VnNnFromFileName(L)   'Lea May
            Nn = Get_Nn_FromVnNn(VnNn)              'May
            Vn = Get_Vn_FromVnNn(VnNn)              'Lea
            y = Get_YearOfPhotoFromFilename(L)      '1987
            Age = Get_AgeFromT5(Nn, Vn, y)          '- oder 23
            P0 = T7_Get_p0_FromHelpers(Id)          '- oder FullEventPathOfPhoto
                If P0 = "-" Then P0 = T7_Get_p0_FromE1E2(E1, E2, p3)
            Tb = T7_Get_TextBoxContent_FromHelpers(Id)          '- oder TextBoxContent
                If Tb = "-" Then Tb = T7_Get_TextBoxContent_FromT5_IfNoHelpersData(Nn, Vn, Id, P0)
        'Fill H9
            H9(i, 1) = i: H9(i, 2) = Nn: H9(i, 3) = Vn: H9(i, 4) = y
            H9(i, 5) = Age: H9(i, 6) = Id: H9(i, 7) = P0: H9(i, 8) = Tb: H9(i, 9) = "-"
JumpToNext:
        Next
        'showArray2D H9
End Sub

Function T7_Get_TextBoxContent_FromT5_IfNoHelpersData(Nn$, Vn$, Id$, P0$)
    'Called from    T7_Fill_H9
    'Tb             = TextBoxContent
    '               = "[Id:p1164-01]['Pip', vh Maier]TV Brr (Pfalz)§Event: 1972...§Trainer§Line4"
    'p0             = "..\Events\1968-11 WM Aaa\19681130 WM Aaa d8 (Lea May) v0133-08 ww.jpg"
    'T5             Sheets("T5") trägt jedwede PersonData: Nn Vn Jhg vh L1 L3 L4 ...
    
    'Vorbereitung
        Dim L1$, L2$, L3$, L4$, T$, z%
        z = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, Vn + " " + Nn)
        If z = 0 Then Stop
        With Sheets("T5")
        L1 = .Cells(z, 18): L3 = .Cells(z, 19): L4 = .Cells(z, 20)
        L2 = T7_Get_L2Event_OfPathP0(P0)
    'Fill Tb
        T = "[Id:" + Id + "]" + L1 + "§" + L2
        If L3 <> "-" Then T = T + "§" + L3
        If L4 <> "-" Then T = T + "§" + L4
    'Finals
        End With
        T7_Get_TextBoxContent_FromT5_IfNoHelpersData = T
End Function

Sub T7_WorksheetActivate()
    'Called from    Worksheet_Activate
    T7_ScrollBar1_SetWidthHeightEtc
End Sub
   
Sub T7_ScrollBar1_SetWidthHeightEtc()
    'Called from    T7_WorksheetActivate
    Dim sM%, zM%
    sM = Get_ColumnNr_HoldingMyTextWhole("T7", "sbMark")
    zM = Get_RowNr_HoldingMyTextWhole("T7", "sbMark")
    With Worksheets("T7")
        .Columns(sM + 7).ColumnWidth = 1.3 'Spaltenbreite links neben ScrollBar1
        .Columns(sM + 8).ColumnWidth = 2.7 'Spaltenbreite ScrollBar1
        With .ScrollBar1
            .Left = Cells(1, sM + 8).Left + 2:      .Width = Cells(zM + 1, sM + 8).Width - 4
            .Top = Cells(zM + 1, sM + 8).Top + 2.6: .Height = 205
        End With
    End With
End Sub

Sub T7_Open_ToDoFolder()
    'Called from    Worksheet_SelectionChange
    Dim p$: DoArrc: p = ArrC(1) + "\prog\Label\LbToDo": openFile (p): EE 0: [B5].Select: EE 1
End Sub

Sub T7_Open_LbArchivFolder()
    'Called from    Worksheet_SelectionChange
    Dim p$: DoArrc: p = ArrC(1) + "\prog\Label\LbArchiv": openFile (p): EE 0: [B5].Select: EE 1
End Sub

Sub T7_Highlight_ActiveFieldLine()
    'Called from    T7_Fill_VisibleScrollLines
    Dim Id$, z%
    With Sheets("T7"): DoArrc
    .Range(.Cells(13, 4), .Cells(26, 8)).Interior.ColorIndex = 0 'keine Füllung
    Id = ArrC(85)
    z = Get_RowNr_HoldingMyTextWholeInColumnX("T7", 8, 12, Id) 'Spalte 8 nach Zeile 12
    If z > 0 Then .Range(.Cells(z, 4), .Cells(z, 8)).Interior.Color = Green3
    End With
End Sub

Sub T7_OneFieldLineWasClicked(z%)
    'Called from    Worksheet_SelectionChange
    Dim Id$
    With Sheets("T7")
        .Range(.Cells(13, 4), .Cells(26, 8)).Interior.ColorIndex = 0 'keine Füllung
        .Range(.Cells(z, 4), .Cells(z, 8)).Interior.Color = Green3
        EE 0: .Cells(z, 2).Select: EE 1
        FillArrC 85, .Cells(z, 8)              'T7 Id of active Photo
        FillArrC 86, T7_Get_TextBoxContent_FromHelpers(ArrC(85))
    End With
    T7_ShowPhoto
End Sub

Function T7_Get_LePathOfActivePhoto() As String
    'Called from    Worksheet_SelectionChange  [UserClick Action:Le]
    'LePath         Pfad zum passenden jpg im passenden Leute-Folder
    'Action         Ermittlung von Nn, Vn um Leute-Ordner zu finden;
    '               im Leute-Ordner FileName per Id ermitteln
    
    'Vorbereitung
        Dim FN$, Nn$, pFo$, Vn$, z%
        DoArrc
    'Nn, Vn
        z = T7_Get_HelpersRowNrOfActivePhoto
        Nn = Sheets("T7").Cells(z, 47): Vn = Sheets("T7").Cells(z, 48)
    'Path of one LeuteFolder
        pFo = Get_FirstSubfolderPath_LikeMyString_OneLevel(ArrC(4), "*" + Nn + ", " + Vn + "*")
    'Path of file
        FN = Get_AllFileNames_Like_OfOneFolder(pFo, "*" + ArrC(85) + "*")
    'Finals
        T7_Get_LePathOfActivePhoto = pFo + "\" + FN
        EE 0: Sheets("T7").Cells(12, 2).Select: EE 1
End Function

Function T7_Get_HelpersRowNrOfActivePhoto() As Integer
    'Called from    T7_Get_EvPathOfActivePhoto, T7_Get_LePathOfActivePhoto,
    '               T7_Fill_TextBoxFromHelpers, T7_Import_Photo
    DoArrc
    T7_Get_HelpersRowNrOfActivePhoto = Get_RowNr_HoldingMyTextWholeInColumnX("T7", 51, 4, ArrC(85)) 'in helpers
End Function

Function T7_Get_EvPathOfActivePhoto() As String
    'Called from    Worksheet_SelectionChange [UserClick Action:Ev]
    
    Dim p$, z%
    z = T7_Get_HelpersRowNrOfActivePhoto
    p = Sheets("T7").Cells(z, 52)
    T7_Get_EvPathOfActivePhoto = p
    EE 0: Sheets("T7").Cells(12, 2).Select: EE 1
End Function

Sub T7_ShowPhoto()
    'Called from    T7_OneFieldLineWasClicked
    'Action         Foto zur aktiven Id wird angezeigt;
    '               update Label-TextBox
    Dim z%
    'Import
        T7_Import_Photo
    'CheckboxText aus Helpers3 holen
        T7_Fill_TextBoxFromHelpers
    'Zelle links neben ActiveFieldLine: Select
        z = Get_RowNr_HoldingMyTextWholeInColumnX("T7", 8, 12, ArrC(85))
        EE 0: Sheets("T7").Cells(z, 2).Select: EE 1
End Sub

Sub T7_Fill_TextBoxFromHelpers()
    'Called from    T7_ShowPhoto

    Dim T$, z%
    z = T7_Get_HelpersRowNrOfActivePhoto
    T = Sheets("T7").Cells(z, 53)       'TextBox-Inhalt aus helpers
    'Id     "[Id:p1234-05]"
        If Left(T, 4) = "[Id:" Then
            T = Mid(T, 14)
        Else
            Sheets("T7").Cells(z, 53) = "[Id:" + ArrC(85) + "]" + T
        End If
    T = Replace(T, "§", vbCrLf): T = Delete_EmptyRowsInString(T)
    T7_Put_TextIntoTextbox T
End Sub

Sub T7_Import_Photo()
    'Called from    T7_ShowPhoto
    'Action         importiert F1 mit oder ohne Label in Sheet T6
    
    'Vorbereitung
        Dim Age$, Id$, z%, Nn$, p$, p1$, p2$, Vn$, y$
        Dim ShapeF1 As Picture, ws As Worksheet
        Set ws = Sheets("T7")
    'Data from helpers
        z = T7_Get_HelpersRowNrOfActivePhoto
        
        
        If z = 0 Then
            Stop
            Exit Sub
        End If
        
        
        Id = ArrC(85): Nn = ws.Cells(z, 47): Vn = ws.Cells(z, 48): y = ws.Cells(z, 49)
        Age = ws.Cells(z, 50): p2 = ws.Cells(z, 52)
    'LabelPhoto könnte bereits vorliegen
        p1 = ArrC(1) + "\prog\Label\LbArchiv\" + Nn + ", " + Vn + " (" + y + ") " + Id + ".jpg"
    'p ermitteln (PathOfPhotoToImport)
        If FileExists(p1) Then p = p1 Else p = p2   'F0 im LbArchiv oder OriginalFoto in Events
        Application.ScreenUpdating = False
    'Position des importierten Fotos        'bei linken oberen Ecke
        ws.Activate: EE 0                   'der selektierten Zelle
        ws.Range("N3").Select: EE 1        'wird die linke obere Ecke des Fotos liegen
    'Foto in T7 einfügen
        'ws.Shapes("F1").Delete
        Delete_Shape_OnSheetX "F1", "T7"
        Set ShapeF1 = ws.Pictures.Insert(p) '"F1" --> T7
        ShapeF1.NAME = "F1"  'Name des eingefügten OriginalFotos: "Picture 10" --> "F1"
        Set ShapeF1 = Nothing
    'Foto Groesse und Position anpassen
        With ws.Shapes("F1")
            .Width = 280: .Left = .Left + 2: .Top = .Top + 1.5 'F1-Werte (Points)
        End With
    'F1B = Rechteck PhotoRahmenBlack FillTransparent
        With ws.Shapes("F1B"):                      .Top = ws.Shapes("F1").Top - 2
            .Left = ws.Shapes("F1").Left - 2:       .Width = ws.Shapes("F1").Width + 3.5
            .Height = ws.Shapes("F1").Height + 3.5: .ZOrder msoBringToFront
        End With
    'F1W = Rechteck PhotoRahmenBlack FillWhite (gleiche Werte wie F1B)
        With ws.Shapes("F1W"): .Top = ws.Shapes("F1B").Top: .Left = ws.Shapes("F1B").Left: .Width = ws.Shapes("F1B").Width: .Height = ws.Shapes("F1B").Height: .Visible = True: End With
End Sub

Function T7_ToDoFileNameIsCorrect(OneFileName$) As Boolean
    Dim F$, Nn$, v$, Vn$, VnNn$: v = vbCrLf: F = OneFileName
    '(1) f muss ein Link sein
        If Not F Like "*.lnk" Then
            show "Die Datei" + v + "  '" + F + "'" + v _
                + "im LbArchiv-Ordner ist kein Link (Verknüpfung)."
            Exit Function
        End If
    '(2) f muss '(Vorname Nachname)' enthalten
        If Not F Like "*([A-ZÄÖÜ][a-zäöüß.]* *[A-ZÄÖÜ][a-zäöüß]*)*" _
            Or F Like "*( *" Or F Like "* )*" Then
            show "Im Dateinamen" + v + "  '" + F + "'" + v _
                + "im LbArchiv-Ordner wurde kein gültiger Personenname" + v _
                + "  '...(Vorname Nachname)...'" + v + "erkannt."
            Exit Function
        End If
    '(3) 'Vorname Nachname' (aus f) muss in PersonData (T5) existieren
        VnNn = T7_Get_VnNnFromFileName(F): Nn = Get_Nn_FromVnNn(VnNn): Vn = Get_Vn_FromVnNn(VnNn)
        If Not Nn_Vn_ExistsIn_T5PersonData(Nn, Vn) Then
            show "Error: T7_ToDoFileNameIsCorrect" + v + v _
            + "Der Name '" + Nn + ", " + Vn + "' ist nicht in den T5-PersonData enthalten."
            Exit Function
        End If
    'Finals
        T7_ToDoFileNameIsCorrect = True
End Function

Function T7_Get_VnNnFromFileName(F$) As String
    Dim s$, C1%, C2%
    If Not F Like "*(*)*" Then Exit Function
    C1 = InStr(1, F, "("): C2 = InStr(C1, F, ")")
    s = CutVonBis(F, C1 + 1, C2 - 1)
    If Not s Like "[A-ZÄÖÜ][a-zäöüß.]* [A-ZÄÖÜ][a-zäöüß]*" Then Stop
    T7_Get_VnNnFromFileName = s
End Function

Function T7_Get_p0_FromE1E2(E1$, E2$, p3$) As String
    'FolderListen holen
        If E1 = "" Then 'E1, E2 nur 1 x laden
            E1 = Get_Paths_ofAllSubfoldersAllLevelsAsStringUseGlobalVar(ArrC(2) + "\Events\")
            E2 = Get_Paths_ofAllSubfoldersAllLevelsAsStringUseGlobalVar(ArrC(2) + "\ClubsNations\")
        End If
    'p0     (Pfad OriginalJpg)
        T7_Get_p0_FromE1E2 = T7_Get_OneFilePathInEvents(E1, E2, p3)
End Function

Function T7_Get_p0_FromHelpers(Id$) As String
    Dim z%
    'Photo-EventPfad p0 von Id könnte in helpers bereits vermerkt sein
        z = Get_RowNr_HoldingMyTextWholeInColumnX("T7", 51, 4, Id)
        If z > 0 Then T7_Get_p0_FromHelpers = Sheets("T7").Cells(z, 52) Else T7_Get_p0_FromHelpers = "-" 'p0
End Function

Function T7_Get_TextBoxContent_FromHelpers(Id$) As String
    'Called from    T7_Fill_H9Dim z%
    'TextBoxContent von Id könnte in helpers bereits vermerkt sein
    
    Dim z%
    z = Get_RowNr_HoldingMyTextWholeInColumnX("T7", 51, 4, Id)
    If z > 0 Then
        T7_Get_TextBoxContent_FromHelpers = Sheets("T7").Cells(z, 53)
    Else
        T7_Get_TextBoxContent_FromHelpers = "-" 'p0
    End If
End Function

Function T7_Get_L2Event_OfPathP0(P0$) As String
    'Called from    T6Read_ToDoFolder
    'p              = PathOfOriginalFile (in 'Events' oder 'ClubsNations')
    '               = "..\Events\2023-07-08 Treff..\20230708_170156 (iPh7)...jpg"
    '               = "..\ClubsNations\D - Baden - TSG Bruchsal\..jpg"
    
    'Vorbereitung
        Dim p$, E$, E2$, C1%, C2%, i%, Arr1() As String
        p = P0
        Arr1 = Split(p, "\")
        E = "Event: "
    'E1 OR E2? (\Events\ OR \ClubsNations\)
        For i = 0 To UBound(Arr1)
            If Arr1(i) = "Events" Then E = E + Arr1(i + 1): Exit For
            If Arr1(i) = "ClubsNations" Then
                E2 = Arr1(i + 1) '"D - Baden - TSG Bruchsal", "D - Bayern", "F"
                If E2 Like "- International" Then E = E + "International"
                If E2 Like "* - * - *" Then
                    'E2 = "D - Baden - TSG Bruchsal", "SU - Ukraine - Kiev"
                    C1 = InStr(1, E2, " - "): C2 = InStr(C1 + 3, E2, " - ")
                    E = E + "inside " + Mid(E2, C2 + 3) + " (" + CutVonBis(E2, C1 + 3, C2 - 1) + ")"
                    If Left(E2, 2) <> "D " Then E = Replace(E, ")", "/" + Left(E2, C1 - 1) + ")")
                ElseIf E2 Like "* - *" Then
                    'E2 = "D - Bayern", "SU - Ukraine"
                    C1 = InStr(1, E2, " - ")
                    E = E + "inside " + Mid(E2, C1 + 3)
                    If Left(E2, 2) <> "D " Then E = E + "/" + Left(E2, C1 - 1)
                ElseIf E2 Like "*zz *" Then
                    'E2 = "D - zz Hochschul-Szene"
                    C1 = InStr(1, E2, " - zz ")
                    E = E + Mid(E2, C1 + 6) + "/" + Left(E2, C1 - 1)
                ElseIf Not E2 Like "* - *" Then
                    'E2 = "D", "F", "USA"
                    E = E + "inside " + E2
                    If E2 = "CAN" Then
                        E2 = E2 + " (Kanada)"
                    ElseIf E2 = "CH" Then E2 = E2 + " (Schweiz)"
                    ElseIf E2 = "F" Then E2 = E2 + " (Frankreich)"
                    ElseIf E2 = "SA" Then E2 = E2 + " (Südafrika)"
                    ElseIf E2 = "SWE" Then E2 = E2 + " (Schweden)"
                    End If
                End If
            End If
        Next
        If E = "Event: " Then E = "Event: ?"
        T7_Get_L2Event_OfPathP0 = E
End Function

Function T7_Get_PathOfLbProd()
    T7_Get_PathOfLbProd = ArrC(1) + "\prog\Label\LbProd\"
End Function

Sub T7_GreyHelpers()
    'Called from    T7_Fill_Helpers
    'Action         T7-helpers-Range grau färben; Spalte 54 mittig
    
    Dim zLast%
    zLast = Get_NrOfLastRowInColumnNr(47, "T7"): With Sheets("T7")
    .Range(.Cells(4, 46), .Cells(zLast + 30, 54)).Interior.ColorIndex = 0 'keine Füllung
    .Range(.Cells(5, 46), .Cells(zLast, 54)).Interior.Color = 15921906 'Grau1
    .Range(.Cells(5, 54), .Cells(zLast, 54)).HorizontalAlignment = xlCenter
    '.Font.size = 6
    End With
End Sub

Sub T7_ClearHelpers()
    'Called from    T7_Fill_Helpers
    
    Dim zLast%
    With Sheets("T7")
        zLast = Get_NrOfLastRowInColumnNr(47, "T7")
        .Range(.Cells(5, 46), .Cells(zLast + 30, 54)).ClearContents
    End With
End Sub

Sub F1B_Click()
    T7_Open_ActiveLbPhoto
End Sub

Sub T7_Open_ActiveLbPhoto()
    'Called from    [UserClick onto Photo]
    'Action         LabelFoto anzeigen (mit Windows-Fotoanzeige)
    
    Dim p$
    p = T7_Get_PathOfActiveLbArchivPhoto
    If FileExists(p) Then openFile p
End Sub

Function T7_Get_PathOfActiveLbArchivPhoto() As String
    Dim Id$, Nn$, p$, Vn$, y$, z%
    With Sheets("T7"): DoArrc: Id = ArrC(85)
    z = T7_Get_HelpersRowNrOfActivePhoto
    Nn = .Cells(z, 47): Vn = .Cells(z, 48): y = .Cells(z, 49): End With
    p = ArrC(1) + "\prog\Label\LbArchiv\" + Nn + ", " + Vn + " (" + y + ") " + Id + ".jpg"
    T7_Get_PathOfActiveLbArchivPhoto = p
End Function

Function T7_Get_OneFilePathInEvents(E1$, E2$, p$) As String
    'Called from    xxx
    'p              = PathOfSourceFile (dessen Link im ToDo-Ordner liegt); P anfangs = ""
    '               = "..\Leute\Wiest, Arno (TSG Bruchsal)\19640000 (Arno Wiest)...jpg"
    '               = "..\Events\2023-07-08 Treff.. (bei L. Mai)\20230708_170156 (iPh7)...jpg"
    'E1             = Alle SubFolder in "Events"
    'E2             = Alle SubFolder in "ClubsNations"
    'Action         sucht das OriginalPhoto in den Ordnern 'Events', 'ClubsNations'

    'Vorbereitung
        Dim F$, Id$, p2$, PN$, v$, y$, i%, ArrFolders() As String, ArrJPGs() As String
        v = vbCrLf: p2 = ArrC(2) 'Path of Folder "Archiv Trampolin"
        Id = Get_IdFromPath(p)    'p0123-04
        PN = Get_NameFromPath(p)
    'Action
        If p Like "*\Events\*" Then
            F = p
        ElseIf p Like "*\ClubsNations\*" Then F = p
        ElseIf p Like "*\Leute\*" Then
                           F = T7_Get_PathOfOriginalJpg_Event1(E1, PN, Id)
            If F = "" Then F = T7_Get_PathOfOriginalJpg_Event2(E2, PN, Id)
        End If
        F = Delete_EmptyEndRowsInString(F)
        T7_Get_OneFilePathInEvents = F
End Function

Function T7_Get_PathOfOriginalJpg_Event1(E1$, PN$, Id$) As String
    'Called from    T7_Get_OneFilePathInEvents
    'PN             = NameOfSourceFile in 'Leute' = "2023 ... p0123-04 nn.jpg" , "1989-11-04 ..."
    '               Der Name eines Photos beginnt immer mit einer Jahreszahl (####)
    '               und besitzt eine Id (p####-##).
    'Id             = "p0123-04" = Teil von PN = Id des 'Leute'-jpg = Id des E1-jpg
    '               gleiche Id - gleiches Photo; E1-jpg = Original; 'Leute'-jpg = Kopie
    'E1             = Alle SubFolder in "Events"; auch SubSubFolder
    '               'Events' wird hier durchsucht; ggf. nur wenige oder 1 Subfolder
    '               SubFolderPath Level 1 startet immer mindestens mit Date ####,
    '                                 oft auch mit Date  ####-## oder ####-##-##
    '               SubFolderPath Level 2 kann ein Date haben (1988 Liga\1988-11-27 Buli..)
    'F              = PathOfFile des 'Events'-jpg, das die PN-Id besitzt ('Leute'-jpg)
    'Action         ggf. Lieferung von F
    
    'Vorbereitung
        Dim F$, Fol$, y$, v$, i%, Arr1() As String, ArrE1() As String
        v = vbCrLf: ArrE1 = Split(E1, v)
    'Sorte?
        If PN Like "####-##-##*" Then
            y = Left(PN, 10)
            'Suche alle E1-Zeilen mit Y
            If E1 Like "*\" + y + "*" Then
                For i = 0 To UBound(ArrE1)
                    If ArrE1(i) Like "*\" + y + "*" Then Fol = Fol + ArrE1(i) + v
                Next
            End If
        ElseIf PN Like "####-##*" Then
            y = Left(PN, 7)
            'Suche alle E1-Zeilen mit Y
            If E1 Like "*\" + y + "*" Then
                For i = 0 To UBound(ArrE1)
                    If ArrE1(i) Like "*\" + y + "*" Then Fol = Fol + ArrE1(i) + v
                Next
            End If
        ElseIf PN Like "####*" Then
            y = Left(PN, 4)
            'Suche alle E1-Zeilen mit Y
            If E1 Like "*\" + y + "*" Then
                For i = 0 To UBound(ArrE1)
                    If ArrE1(i) Like "*\" + y + "*" Then Fol = Fol + ArrE1(i) + v
                Next
            End If
        End If
        If Fol = "" Then Exit Function Else Fol = Delete_EmptyEndRowsInString(Fol)
        'Fol trägt alle Event1-FolderPaths, die mit PN-Date Y beginnen
    'Check ob ein Fol-jpg mit PN-Id existiert
        Arr1 = Split(Fol, v)
        For i = 0 To UBound(Arr1)
            F = Get_AllFilePaths_WithMyStringInFileName_OfOneFolder(Arr1(i), Id)
            If F <> "" Then Exit For
        Next
    T7_Get_PathOfOriginalJpg_Event1 = F
End Function

Function T7_Get_PathOfOriginalJpg_Event2(E2$, PN$, Id$) As String
    'Called from    T7_Get_OneFilePathInEvents
    'PN             = NameOfSourceFile in 'Leute' = "2023 ... p0123-04 nn.jpg" , "1989-11-04 ..."
    '               Der Name eines Photos beginnt immer mit einer Jahreszahl (####)
    '               und besitzt eine Id (p####-##)
    'Id             = "p0123-04" = Teil von PN = Id des 'Leute'-jpg = Id des E2-jpg
    '               gleiche Id - gleiches Photo; E2-jpg = Original; 'Leute'-jpg = Kopie
    'E2             = Alle SubFolder in "ClubsNations"; auch SubSubFolder
    '               alle E2-Folder werden hier durchsucht
    '               SubFolderPath Level 1 hat kein Date
    '               SubFolderPath Level 2 kann ein Date haben
    'F              = PathOfFile des E2-jpg, das die PN-Id besitzt ('Leute'-jpg)
    'Action         Lieferung von F
    
    'Vorbereitung
        Dim F$, i%, ArrE2() As String
    'Check ob ein E2-jpg mit PN-Id existiert
        ArrE2 = Split(E2, vbCrLf)
        For i = 0 To UBound(ArrE2)
            F = Get_AllFilePaths_WithMyStringInFileName_OfOneFolder(ArrE2(i), Id)
            If F <> "" Then Exit For
        Next
        F = Delete_EmptyEndRowsInString(F)
    T7_Get_PathOfOriginalJpg_Event2 = F
    If Right(F, 2) = vbCrLf Then Stop
End Function

Sub T7_Put_TextIntoTextbox(MyText$)
    Sheets("T7").Shapes("tx").TextFrame.Characters.text = MyText
End Sub

Function T7_Get_TextFromTextBoxOnSheetT7_AsOneLine() As String
    Dim s$
    With Sheets("T7")
    s = .Shapes("tx").TextFrame.Characters.Caption 'msoTextBox
    s = Replace(s, vbLf, "§")
    End With
    T7_Get_TextFromTextBoxOnSheetT7_AsOneLine = s
End Function


