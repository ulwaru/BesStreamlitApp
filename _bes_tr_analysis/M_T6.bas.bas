Attribute VB_Name = "M_T6"
Option Explicit 'M_T6

#If VBA7 Then 'Office 64 bit
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else 'Office 32 bit
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

Sub zzz_M_T6()

    showProcs "T5_Get"
    
    'show CStr(getColor([bb22], 0, "T6"))
    EE 0: Beep
End Sub

Sub Check_T6TextBoxChange()
    'Called from    Worksheet_SelectionChange [T6; every UserClick]
    'helpers3       Liste in T6, Spalte 65-70: |Id|p0|TxtBox|
    'Action         schreibt TextBox-Inhalt (falls neu) in helpers3
    '               (dies ist die einzige Sub, die das tut)
    
Exit Sub
    
    'Vorbereitung
        Dim Id$, Nn$, P0$, p1$, p2$, s$, tx1$, tx2$, v$, Vn$, y$, z3%
        Dim A() As String, PD() As String
        v = vbCrLf: With Sheets("T6")
    'Inhalt der 6 Zellen links des angekreuzten Kästchens
        T6_Load_SixCellsLeftOfX A 'showArray A
        If A(1) = "-" Then Exit Sub
        Nn = A(1): Vn = A(2): y = A(3): Id = A(5): P0 = A(6)
    'tx1    bisheriger TextBox-Inhalt aus Helpers3
        'ZeilenNr der zZt aktiven Helpers3-Zeile
            z3 = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 66, 4, Id)
        If z3 = 0 Then
            'Ausgewählte Person ist noch nicht in helpers3
            Stop
        Else
            tx1 = .Cells(z3, 69)
        End If
        If tx1 = "" Then Stop
    'tx2    aktueller TextBox-Inhalt (aus TextBox selbst)
        tx2 = Get_TextFrom_T6Textbox
        If tx2 = "" Then
            Stop
            Exit Sub
        Else
            tx2 = Replace(tx2, vbLf, "§")
        End If
    'TextBox was changed?
        If tx1 = tx2 Then
            'TextBox was not changed
            Exit Sub
        Else
            'TextBox was changed
            'neuen Text in Helpers3 schreiben
                .Cells(z3, 69) = tx2
            'NewLabel
                Produce_T6PhotoWithLabelAndShow
        End If
    End With
End Sub

Public Function DateiExistiert(PathOfFile As String) As Boolean
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    DateiExistiert = objFSO.FileExists(PathOfFile)
    Set objFSO = Nothing
End Function

Function Get_NameOfParentFolder(ByVal PathOfFolder$) As String
    Dim fso As Object, parentPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
        parentPath = fso.GetParentFolderName(PathOfFolder)
        If parentPath <> "" Then
            Get_NameOfParentFolder = fso.GetFileName(parentPath)
        End If
End Function

Function Get_PathOfParentFolder(PathOfFileOrFolder$) As String
    Dim c As Integer, s As String
    s = PathOfFileOrFolder
    If InStrRev(s, "\") < 3 Then Exit Function
    If Right(s, 1) = "\" Then s = Left(s, Len(s) - 1)
    's = "C:\MyFolder\MySubFolder" oder "C:\MyFolder\MyFile"
    c = InStrRev(s, "\"): s = Left(s, c - 1) 's = "C:\MyFolder"
    Get_PathOfParentFolder = s
End Function

Sub VerweisSetzen_VBA()
    'Bibliothek "Microsoft Scripting Runtime"
    Const strRef As String = "{420B2830-E718-11CF-893D-00A0C9054228}"
    Dim objRef As Object
    On Error GoTo jump1
    Set objRef = ThisWorkbook.VBProject.References
    objRef.AddFromGuid strRef, 1, 0
jump1:
    Set objRef = Nothing
    If Err.Number = 32813 Then Exit Sub
    If Err.Number <> 0 Then MsgBox "xError: " & Err.Number & " " & Err.Description
End Sub

Sub VerweisSetzen_VBE()
    'Bibliothek "Microsoft Visual Basic for Application Extensibility 5.3"
    Dim VBEobj As Object
    On Error Resume Next
    VBEobj = Application.VBE.ActiveVBProject.References.AddFromGuid("{0002E157-0000-0000-C000-000000000046}", 5, 3)
    On Error GoTo 0
End Sub

Function loescheLeerzeilen(s As String)
    Do While InStr(1, s, vbCrLf + vbCrLf) > 0
        s = Replace(s, vbCrLf + vbCrLf, vbCrLf)
    Loop
    loescheLeerzeilen = s
End Function

Function getNameOfPath(PathOfFileOrFolder As String) As String
    Dim s As String: s = PathOfFileOrFolder
    If Right(s, 1) = "\" Then s = Left(s, Len(s) - 1)
    s = Mid(s, InStrRev(s, "\") + 1)
    getNameOfPath = s
End Function

Function Add_EmptyRowAfterRowX(MyString$, RowNrX%) As String
    'Vorbereitung
        Dim s$, v$, A%, i%, c&
        v = vbCrLf: s = MyString: Add_EmptyRowAfterRowX = s
        A = anzAinB(v, s) 'Anzahl Zeilen = a + 1
    'Exit
        If RowNrX > A + 1 Then Exit Function
    'Get position where to add
        s = s + v
        For i = 1 To RowNrX
            c = InStr(c + 1, s, vbCrLf)
            If c = 0 Then Exit For
        Next
    'Add
        s = Left(s, c - 1) + v + Mid(s, c)
    'Finals
        Add_EmptyRowAfterRowX = s
End Function

Sub importiereFoto(pathFotoF1 As String)
    'F1 einfügen:
    Dim ws As Worksheet, fotF1 As Picture
    Dim F1L As Double, F1T As Double, F1W As Double, F1H As Double
    Set ws = ActiveSheet 'Tabellenblatt festlegen
    ActiveSheet.Range("Q18").Select
    'bei linken oberen Ecke der selektierten Zelle wird die linke obere Ecke des Fotos liegen
    Application.Wait Now + TimeSerial(0, 0, 1) 'wartet 1 Sekunde
    Delete_Shape_OnSheetX "F1", "T6"

    'pathFotoF1 = "G:\Archiv VBA\Bes FaceWith\2 faceOnly 1000x1500\Bentz, Leonhard 20191005 a2.jpg"
    Set fotF1 = ws.Pictures.Insert(pathFotoF1)             '"F1" hinzufügen (ggf. nicht Originalgröße)
    fotF1.NAME = "F1"  'Name des eingefügten OriginalFotos: Picture 10 --> "F1"
    Set fotF1 = Nothing
    '"F1" Groesse anpassen:
    ws.Shapes("F1").ScaleWidth 1, True '"F1" in Originalgröße darstellen (crop bezieht sich auf Originalgröße)
    ws.Shapes("F1").Width = 100
End Sub

Sub T6_AddRechteck()
    'Called from    [none]
    
    Dim shp As shape, L!, T!, w!, H!
    Delete_Shape_OnSheetX "F1B", "T6"
    With Sheets("T6")
    L = [M14].Left: T = [M14].Top: w = [M14].Width
    H = [M29].Top - [M14].Top
    Set shp = .Shapes.AddShape(msoShapeRectangle, L, T, w, H)
    With shp: .NAME = "F1B": .Line.ForeColor.RGB = RGB(0, 0, 0): .Line.Weight = 1
             .Fill.Transparency = 1: .Visible = msoTrue: '.Visible = msoFalse
    End With: End With
End Sub

Function Get_T6EvPathOfShowPhoto() As String
    Dim A() As String
    'Inhalt der 6 Zellen links des angekreuzten Kästchens
        T6_Load_SixCellsLeftOfX A 'showArray A
        Get_T6EvPathOfShowPhoto = A(6)
End Function

Function Get_T6LePathOfShowPhoto()
    'Called from    T6_Load_ArrFields, T6SomeActionsOnPhoto
    'LePath         Pfad zum passenden jpg im passenden Leute-Folder
    
    'Vorbereitung
        Dim pFo$, FN$, Id$, Nn$, PathOfPhoto$, Vn$, zX%, ArrFields()
        DoArrc
    'ArrFields
        T6_Load_ArrFields ArrFields, zX
        Nn = ArrFields(zX, 1)       'Mai
        Vn = ArrFields(zX, 2)       'Lea
        Id = ArrFields(zX, 4)       'p1234-05
    'Path of one LeuteFolder
        pFo = Get_FirstSubfolderPath_LikeMyString_OneLevel(ArrC(4), "*" + Nn + ", " + Vn + "*")
    'Path of file
        FN = Get_AllFileNames_Like_OfOneFolder(pFo, "*" + Id + "*")
    Get_T6LePathOfShowPhoto = pFo + "\" + FN
End Function

Function Get_IdFromPath(p$) As String
    'Called from    T7_Get_OneFilePathInEvents, T7_Fill_H9
    'P              = "...  p0123-04 nn.jpg", "...  p0123-04 nn.jpg.lnk"
    Dim Id$
    Id = CutVonBis(p, Len(p) - 14, Len(p) - 7)
    If Not Id Like "[p|v]####-##" Then show p: Stop
    Get_IdFromPath = Id
End Function

Sub Add_4Images_BelowEachOther(p1$, p2$, p3$, p4$, pOUT$)
    'Called from    Create_i08
    'Tool           gm, GraphicsMagick
    Dim sCmd$, qq$, c%, i%
    qq = Chr(34) 'quote, "-Zeichen
    'gm convert "p1" "p2" ... -append "pOUT"
    sCmd = "gm convert " + qq + p1 + qq + " " + qq + p2 + qq + " " + qq + p3 + qq + " " + qq + p4 + qq + " -append " + qq + pOUT + qq
    'Call Shell(sCmd)
    ShellAndWaitWithFlashingCmdWindow sCmd
End Sub

Sub WaitMax10SecUntilIcoJpgIsKilled()
    'Warte MaxSec Sekunden (1s = 1, 100ms = 0.1, 500ms = 0.5, 50ms = 0.05...)
    Dim PathOfFolder_KillOneIcoJpg$, dStopTime#
    PathOfFolder_KillOneIcoJpg = ArrC(1) + "\ico\Bitte das Nichtgewünschte jpg entfernen"
    dStopTime = Date + ((Timer + 10) / 86400)
    Do While dStopTime >= (Date + Timer / 86400)
        DoEvents
        If Not FileExists(PathOfFolder_KillOneIcoJpg + "\existing.jpg") Then Exit Do
        If Not FileExists(PathOfFolder_KillOneIcoJpg + "\new.jpg") Then Exit Do
    Loop
    On Error GoTo 0
    Beep
End Sub

Sub WaitMaxSecUntilFileExists(MaxSec As Double, PathOfFile$)
    'Warte MaxSec Sekunden (1s = 1, 100ms = 0.1, 500ms = 0.5, 50ms = 0.05...)
    Dim dStopTime As Double
    
    dStopTime = Date + ((Timer + MaxSec) / 86400)
    Do While dStopTime >= (Date + Timer / 86400)
        DoEvents
        If DateiExistiert(PathOfFile) And Not FileIsLocked(PathOfFile) Then Exit Do
    Loop
    
    dStopTime = Date + ((Timer + MaxSec) / 86400)
    Do While dStopTime >= (Date + Timer / 86400)
        DoEvents
        On Error Resume Next
        If IsNumeric(Get_Height_OfOneImage(PathOfFile)) Then Exit Do
    Loop
    
    On Error GoTo 0
    Beep
End Sub

Function FileIsLocked(strFileName As String) As Boolean
   On Error Resume Next
   ' If the file is already opened by another process,
   ' and the specified type of access is not allowed,
   ' the Open operation fails and an error occurs.
   Open strFileName For Binary Access Read Write Lock Read Write As #1
   Close #1
   ' If an error occurs, the document is currently open.
   If Err.Number <> 0 Then
      ' Display the error number and description.
      'MsgBox "Error #" & Str(Err.Number) & " - " & Err.Description
      FileIsLocked = True
      Err.Clear
   End If
End Function

Sub aufeinander(N1$, N2$)
    'Called from    Produce_T6PhotoWithLabel
    
    'Action         OriginalPhoto L1 [1000x1500] wird auf den
    '               leeren wheat-Hintergrund [1000x1750] gelegt
    'Vorbereitung
        Dim pFOLD$, p2$, p1$, LbSize$, pOUT$, pTool$, qq$, sCmd$
        qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        pFOLD = "F:\Archiv TR\prog\Label\LbProd\"
        p1 = " " + qq + pFOLD + N1 + qq
        p2 = " " + qq + pFOLD + N2 + qq
        pOUT = " " + qq + pFOLD + "L1.jpg" + qq
    'Create
        sCmd = pTool + "composite -geometry  +0+0 " + p1 + p2 + pOUT
        ShellAndWaitWithFlashingCmdWindow sCmd
End Sub

Sub T6_AddTextToTextbox(TextToAdd$)
    With Sheets("T6").Shapes("tx")
    .TextFrame.Characters.text = .TextFrame.Characters.Caption + TextToAdd
    End With
End Sub

Sub Fix_PositionOf_T6TextBox()
    With Sheets("T6").Shapes("tx")
        .Left = [W26].Left + 2: .Top = [W23].Top + 7: .Width = 275: .Height = 56.25
    End With
End Sub

Sub Open_T6LbProductionFolder()
    Dim p$: DoArrc: p = ArrC(1) + "\prog\Label\LbProd": openFile (p): EE 0: [B5].Select: EE 1
End Sub

Sub Open_T6LbArchivFolder()
    Dim p$: DoArrc: p = ArrC(1) + "\prog\Label\LbArchiv": openFile (p): EE 0: [B5].Select: EE 1
End Sub

Sub Kill_AllFilesInOneFolder(PathOfOneFolder$)
    Dim aFile As String
    If Right(PathOfOneFolder, 1) <> "\" Then PathOfOneFolder = PathOfOneFolder + "\"
    aFile = PathOfOneFolder + "*.*"
    If Len(Dir$(aFile)) > 0 Then Kill aFile
End Sub

Function Get_ArrT5_ZeilenNrOfNnVn(ArrT5, Nn$, Vn$) As Integer
    Dim i%
    For i = 1 To UBound(ArrT5, 1)
        If ArrT5(i, 1) = Nn Then
            If ArrT5(i, 2) = Vn Then Get_ArrT5_ZeilenNrOfNnVn = i: Exit Function
        End If
    Next
End Function

Function Nn_Vn_ExistsIn_T5PersonData(Nn$, Vn$) As Boolean
    Dim N$, z%
    N = Vn + " " + Nn
    z = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, N)
    If z > 0 Then Nn_Vn_ExistsIn_T5PersonData = True
End Function

Function NnVn_ExistsIn_T5(NnVn$) As Boolean
    Dim VnNn$
    VnNn = Get_VnNn_Give_NnVn(NnVn)
    If VnNn_ExistsIn_T5(VnNn) Then NnVn_ExistsIn_T5 = True
End Function

Function VnNn_ExistsIn_T5(VnNn$) As Boolean
    Dim z%
    z = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, VnNn)
    If z > 0 Then VnNn_ExistsIn_T5 = True
End Function


Function T5_Get_NameOfPersonfolder_Give_VnNn(VnNn$) As String
    Dim z%: z = Get_RowNr_HoldingMyTextPartInColumnX("T5", 22, 7, VnNn)
    If z = 0 Then Exit Function
    T5_Get_NameOfPersonfolder_Give_VnNn = Sheets("T5").Cells(z, 10)
End Function

Function T5_Get_PathOfPersonfolder_Give_VnNn(VnNn$) As String
    Dim z%: z = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, VnNn)
    If z = 0 Then Exit Function
    T5_Get_PathOfPersonfolder_Give_VnNn = ArrC(4) + "\" + Sheets("T5").Cells(z, 10)
End Function

Function T5_Get_mw_Give_NnVn(NnVn$) As String
    T5_Get_mw_Give_NnVn = T5_Get_mw_Give_VnNn(Get_VnNn_Give_NnVn(NnVn))
End Function

Function T5_Get_mw_Give_VnNn(VnNn$) As String
    Dim z%: z = Get_RowNr_HoldingMyTextPartInColumnX("T5", 22, 7, VnNn)
    If z = 0 Then Exit Function
    T5_Get_mw_Give_VnNn = Sheets("T5").Cells(z, 5)
End Function

Sub T5_Set_mw(mw$, VnNn$)
    Dim z%: z = Get_RowNr_HoldingMyTextPartInColumnX("T5", 22, 7, VnNn)
    If z > 0 Then Sheets("T5").Cells(z, 5) = mw
End Sub

Function T5_Get_NameOfPersonfolder_Give_NnVn(NnVn$) As String
    Dim p: p = T5_Get_PathOfPersonfolder_Give_NnVn(NnVn)
    T5_Get_NameOfPersonfolder_Give_NnVn = NameOfPath(p)
End Function

Function T5_Get_PathOfPersonfolder_Give_NnVn(NnKommaBlankVn$) As String
    'NnVn           'May, Lea 'May, Lea (Bonn) 'May, Lea.lnk
    Dim NnVn$, PathOfPersonfolder$, VnNn$, c%, z%
    'NnVn reduzieren auf "Nn, Vn", z. B. "May, Lea"
        NnVn = NnKommaBlankVn: c = InStr(1, NnVn, " (")
        If c > 0 Then NnVn = Left(NnVn, c - 1): NnVn = Replace(NnVn, ".lnk", "")
    'VnNn erstellen
        c = InStr(1, NnVn, ", ")
        If c = 0 Then Exit Function 'NnVn ist kein "Nn, Vn"
        VnNn = Mid(NnVn, c + 2) + " " + Left(NnVn, c - 1)
    'Mit VnNn in T5 suchen
        PathOfPersonfolder = T5_Get_PathOfPersonfolder_Give_VnNn(VnNn)
    'Finals
        T5_Get_PathOfPersonfolder_Give_NnVn = PathOfPersonfolder
End Function

Function T5_Get_T5RowNrOfOneVnNn(VnNn$) As Integer
    T5_Get_T5RowNrOfOneVnNn = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, VnNn)
End Function

Function T5_Get_T5RowNrOfOneNnVn(Nn$, Vn$) As Integer
    T5_Get_T5RowNrOfOneNnVn = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, Vn + " " + Nn)
End Function

Function Get_JhgFromT5(Nn$, Vn$) As String
    Dim N$, T$, z%
    N = Vn + " " + Nn
    z = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, N)
    If z = 0 Then Exit Function
    Get_JhgFromT5 = CStr(Sheets("T5").Cells(z, 6))
End Function

Function Get_AgeFromT5(Nn$, Vn$, y$) As String
    'Called from    T6_FillHelpers1, T6_FillHelpers2Helpers3
    
    Dim Jhg$: Jhg = Get_JhgFromT5(Nn$, Vn$)
    If Jhg = "" Then Get_AgeFromT5 = "-" Else Get_AgeFromT5 = CStr(CInt(y) - CInt(Jhg))
End Function

Function Get_T5_L1L2L3_OfNnVn(Nn$, Vn$) As String
    Dim N$, T$, z%
    N = Vn + " " + Nn
    z = Get_RowNr_HoldingMyTextPartInColumnX("T5", 22, 1, N)
    With Sheets("T5")
        T = .Cells(z, 18) + "§" + .Cells(z, 19) + "§" + .Cells(z, 20)
    End With
    Get_T5_L1ToL4_OfNnVn = T
End Function

Sub T6_FillHelpers2Helpers1()
    T6_FillHelpers2Helpers3
    T6_FillHelpers1
End Sub

Function Get_YearOfPhotoFromFilename(NameOfLnkFile$) As String
    'Called from    T6Read_ToDoFolder
    Dim y$
    y = Left(NameOfLnkFile, 4)
    If IsNumeric(y) Then
        If CInt(y) > 1800 Then Get_YearOfPhotoFromFilename = y
    End If
End Function

Sub Check_T6AllP0InHelpers1Helpers2()
    'Called from    T6_FillHelpers1

    'Vorbereitung
        Dim P0$, s$, v$, i%, H1(), H2()
        v = vbCrLf
        T6_Load_ArrHelpers1 H1
        T6_Load_ArrHelpers2 H2
    'p0 in Helpers1
        For i = 1 To UBound(H1)
            P0 = H1(i, 7)
            If P0 Like "*:\*" Then
                If Not FileExists(P0) Then
                    s = s + "File does not exist in helpers1:" + P0 + v
                End If
            End If
        Next
    'p0 in Helpers2
        For i = 1 To UBound(H2)
            P0 = H2(i, 7)
            If P0 Like "*:\*" Then
                If Not FileExists(P0) Then T6_Repair_p0 P0, s
            End If
        Next
        If s <> "" Then show "Check_T6AllP0InHelpers1Helpers2" + v + v + s
End Sub

Sub Check_T6AllAgeInHelpers1Helpers2()
    'Called from    T6_FillHelpers1

    'Vorbereitung
        Dim AgeT5$, AgeT6$, Nn$, s$, v$, Vn$, y$, i%, H1(), H2()
        v = vbCrLf
        T6_Load_ArrHelpers1 H1
        T6_Load_ArrHelpers2 H2
    'Age in Helpers1
        For i = 1 To UBound(H1)
            Nn = H1(i, 2): Vn = H1(i, 3): y = CStr(H1(i, 4))
            If Nn <> "-" Then
                AgeT5 = Get_AgeFromT5(Nn, Vn, y)
                AgeT6 = CStr(H1(i, 5))
                    If AgeT5 <> AgeT6 Then s = s + "Age has Changed: " _
                        + Nn + ", " + Vn + "; " + AgeT6 + " --> " + AgeT5 + v
            End If
        Next
    'Age in Helpers2
        For i = 1 To UBound(H2)
            Nn = H2(i, 2): Vn = H2(i, 3): y = CStr(H2(i, 4))
            AgeT5 = Get_AgeFromT5(Nn, Vn, y)
            AgeT6 = CStr(H2(i, 5))
                If AgeT5 <> AgeT6 Then
                    s = s + "Age has Changed: " _
                    + Nn + ", " + Vn + "; " + AgeT6 + " --> " + AgeT5 + v
                End If
        Next
        If s <> "" Then show s
End Sub

Sub T6_Repair_p0(P0$, s$)
    'Called from    Check_T6AllP0InHelpers1Helpers2
    'p0             PathOfFile mit "File does not exist"-Meldung
    '               ggf. wurde der jpg-Name nach der Labelproduktion geändert
    'Action         Id im gleichen p0-Ordner suchen;
    '               in helpers ändern, falls gefunden
    
    'Vorbereitung
        Dim Id$, p1$, p2$, v$, zF&, zH1&, zH2&, zH3&
        With Sheets("T6"): v = vbCrLf
    'Id
        Id = Mid(P0, Len(P0) - 14, 8) 'p0 = "... p0123-04 nn.jpg"
        If Not Id Like "[p|v]####-##" Then s = s + "Id not found in: " + P0 + v: Exit Sub
        'Korrekte Id in p0 gefunden
    'Find
        p1 = Get_PathOfParentFolder(P0)
        p2 = Get_AllFilePaths_WithMyStringInFileName_OfOneFolder(p1, Id)
        If Not FileExists(p2) Then s = s + "File does not exist: [" + p2 + "]" + v: Exit Sub
        'korrekten p0-FilePfad gefunden: p2
    'Change path in helpers1, helpers2
        zH1 = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 52, 4, Id)    'helpers1
        If zH1 > 0 Then .Cells(zH1, 53) = p2
        zH2 = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 61, 4, Id)    'helpers2
        If zH2 > 0 Then .Cells(zH2, 62) = p2
        zH3 = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 66, 4, Id)    'helpers3
        If zH3 > 0 Then .Cells(zH3, 68) = p2
        zF = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 8, 14, Id)     'Fields
        If zF > 0 Then .Cells(zF, 9) = p2
        If zH1 = 0 And zH2 = 0 Then s = s + "Id not found in helpers: " + Id + v
    'Finals
        End With
End Sub




