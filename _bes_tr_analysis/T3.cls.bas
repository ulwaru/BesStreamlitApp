Attribute VB_Name = "T3"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit 'T3
'Beep
    Private Declare Function Beep Lib "kernel32" (Optional ByVal dwFreq As Long = 900, Optional ByVal dwDuration As Long = 100) As Long

Sub zzz_T3()
    
    'showProcs "beep"
    
    
    EE 1: Beep
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range) 'T3
    'Called from    [UserClick to select another T3-Cell]
    
    'Vorab (bei jedem Klick auf eine andere T3-Zelle)
            Dim s%, z%
            z = Target.Row: s = Target.Column: DoArr
            FillArrC_T3PositionsOfListOfCompetitors
            Beep1 'Kontroll-Beep für SelectionChangeEvent-Aufruf
    'Select A1
        If z < 6 And s < 12 Then EE 0: [A1].Select: EE 1
    'CellButton 'Update this list"
        If z > 1 And z < 5 And s > 12 And s < 17 Then
            EE 0: Range("a1").Select: EE 1
            Call T3_Create_CompetitorsList: Exit Sub
        End If
    'CellButton 'Create textfile of this list"
        If z > 5 And z < 9 And s > 12 And s < 17 Then
            EE 0: Range("a1").Select: EE 1
            Call T3_Create_TextFile_ListOfCompetitors: Exit Sub
        End If
    'Search for
        If z = 11 And s = 13 Then Exit Sub 'Input "Suche" ermöglichen
        If z = 12 And s = 13 Then T3_SearchFor Range("M11").Value: Exit Sub
    'Jump
        'If z > 7 And s > 1 And s < 12 Then JumpFromCompetitorsListToMyDesign z, s: Exit Sub
    'Sort
        If z = 6 And s > 1 And s < 12 Then T3_Sort s: Exit Sub
    EE 1
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    'Beep 1111, 55: Beep 1111, 55
    Range_Change_TEST Target
End Sub

Sub Range_Change_TEST(ByVal Target As Range)
    Dim Area$, r As Range
    'Area to observe
        Area = "i8:i9"
    Set r = Intersect(Target, Me.Range(Area))
    EE 0
    'If r Is Nothing Then [M23] = "No" Else [M23] = r.Address + " was changed"
    If Not r Is Nothing Then Beep 1111, 55: Beep 1111, 55
    EE 1

End Sub

Private Sub Worksheet_Activate() 'Sobald dieses Blatt aktiviert/angezeigt wird
    ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
    Hide_Shape "T3", "T3RedRect"
End Sub

Sub T3_SearchFor(s$)
    Dim i%, z%, zLast%, Arr1()
    EE 0: [A1].Select
    zLast = Get_NrOfLastRowInColumnNr(2, "T3")
    Arr1 = Range(Cells(1, 2), Cells(zLast, 2))
    For i = 8 To zLast
        If LCase(Arr1(i, 1)) Like LCase(s) + "*" Then
            Range(Cells(i, 1), Cells(i, 12)).Select: EE 1: Exit Sub
        End If
    Next
    'String s wurde nicht gefunden
    If Len(s) = 1 Then EE 1: Exit Sub
    s = Left(s, Len(s) - 1)
    For i = 8 To zLast
        If LCase(Arr1(i, 1)) Like LCase(s) + "*" Then
            Range(Cells(i, 1), Cells(i, 12)).Select: EE 1: Exit For
        End If
    Next
    EE 1
End Sub

Sub JumpFromCompetitorsListToMyDesign(z%, s%)
    'Called from    [UserClick onto 1 line of T3CompetitorsList]
    
    
    Stop: Exit Sub 'überarbeiten
    
    
    Dim Adr$, Nn$, OvalValues$, Titel$, Vn$, zs$
    Titel = Trim(Cells(z, 9))
    Nn = Cells(z, 2) 'NachName Competitor
    Vn = Cells(z, 3) 'VorName  Competitor
    If Titel Like "####*" Then
        '(z,s)=(ZeilenNrTitel,SpaltenNrTitel)
            zs = Get_ZeSp_OfDesignTitleNotFolderName(Titel) '039064
            z = CInt(Left(zs, 3)): s = CInt(Right(zs, 3))   '39 '64
        'Springe kurz vor die linke obere Ecke des Designs
            Application.GoTo Reference:=Worksheets("T4").Cells(z - 2, s - 2), Scroll:=True
        'Aktiviere Design/arrC/...; zeige DesignButtons
            EE 1: Worksheets("T4").Cells(z - 1, s - 1).Select
        'Search for competitor
            FindCompetitorInsideDesign Vn + " " + Nn, Adr
            OvalValues = Get_OvalValues(Adr)
            'show OvalValues '|2973,5|410|71,75|35|2973,5|245|71,75|35|2754,75|135|3009,375|427,5|3009,375|427,5|3009,375|262,5|3009,375|262,5|387,8017|284,7633
            Draw_OvalShapeAroundName1 OvalValues: Draw_OvalShapeAroundName2 OvalValues
            Draw_Arrow1 OvalValues: Draw_Arrow2 OvalValues
            'Draw_Arrows Adr
        EE 1
    End If
End Sub

Sub FindCompetitorInsideDesign(Searchstring$, Adr$)
    'Called from    JumpFromCompetitorsListToMyDesign
    Dim r As Range, firstAddress$, s$
    DoArrc
    With Worksheets("T4").Range(ArrC(36)) 'Range of Design
        Set r = .Find(Searchstring, LookIn:=xlValues, LookAt:=xlPart)
        If Not r Is Nothing Then
            firstAddress = r.Address
            Do
                Set r = .FindNext(r)
                If s Like "*" + r.Address + "*" Then Exit Do
                s = s + r.Address + "|"
            Loop While Not r Is Nothing
        End If
    End With
    Adr = Replace(s, "$", "")
End Sub

Function Get_OvalValues(Adr$) As String
    'Called from    JumpFromCompetitorsListToMyDesign
    'Adr            = "CB29|CB18|" = CellAddresses
    'Action         Stellt alle benötigten Werte für Ovale und Pfeile bereit
    'OvalValues     =|L1|T1|W1|H1|L2|T2|W2|H2|x1|y1|xM1|yM1|xM2|yM2|LäPf1|LäPf2|DrawA
    '                 1  2  3  4  5  6  7  8  9  10 11  12  13  14   15    16    17

    'Vorbereitung
        If Adr = "" Then Exit Function
        Dim Cell1$, Cell2$, cV$, DrawA$, s$, c%
        Dim LäPf1!, LäPf2!, L1!, t1!, W1!, H1!, L2!, T2!, W2!, H2!
        Dim x1!, y1!, Ov1x1!, Ov1y1!, xM1!, yM1!, Ov2x1!, Ov2y1!, xM2!, yM2!
    'Cell1, Cell2       (Cell1 = "CB29", Cell2 = "CB18" oder "")
        c = InStr(1, Adr, "|"): Cell1 = Left(Adr, c - 1) '"CB29"
        If Len(Adr) = c Then Cell2 = "" Else Cell2 = Replace(Mid(Adr, c + 1), "|", "")
    'Oval1
        L1 = Get_LeftOfOvalInDesign(Cell1):  t1 = Get_TopOfOvalInDesign(Cell1)
        W1 = Get_WidthOfOvalInDesign(Cell1): H1 = Get_HeightOfOvalInDesign
        x1 = Cells(CInt(ArrC(37)), CInt(ArrC(39))).Left + 30   'Nähe DesignCellLiOb
        y1 = Cells(CInt(ArrC(37)), CInt(ArrC(39))).Top + 30     'Nähe DesignCellLiOb
        xM1 = L1 + 0.5 * W1: yM1 = t1 + 0.5 * H1
        LäPf1 = Sqr((xM1 - x1) ^ 2 + (yM1 - y1) ^ 2): cV = "yy"
        If CellIsInVisibleRange(Cell1) Then cV = "yy" Else cV = "ny" 'yes, no
    'Oval2
        If Cell2 <> "" Then
            L2 = Get_LeftOfOvalInDesign(Cell2): W2 = W1
            T2 = Get_TopOfOvalInDesign(Cell2):  H2 = H1
            xM2 = L2 + 0.5 * W2: yM2 = T2 + 0.5 * H2
            LäPf2 = Sqr((xM2 - x1) ^ 2 + (yM2 - y1) ^ 2)
            If CellIsInVisibleRange(Cell2) Then cV = Left(cV, 1) + "y" Else cV = Left(cV, 1) + "n"
        End If
    'DrawArrrow
        If cV = "yy" Then DrawA = "No" Else DrawA = "Yes"
    
    s = "|" + CStr(L1) + "|" + CStr(t1) + "|" + CStr(W1) + "|" + CStr(H1) _
      + "|" + CStr(L2) + "|" + CStr(T2) + "|" + CStr(W2) + "|" + CStr(H2) _
      + "|" + CStr(x1) + "|" + CStr(y1) + "|" + CStr(xM1) + "|" + CStr(yM1) _
      + "|" + CStr(xM2) + "|" + CStr(yM2) + "|" + CStr(LäPf1) + "|" + CStr(LäPf2) + "|" + DrawA
    Get_OvalValues = s
End Function

Sub Draw_OvalShapeAroundName1(OvalValues$)
    'Called from    JumpFromCompetitorsListToMyDesign
    'Action         Zeichnet das Oval1 (um Name1 herum)
    'OvalValues     =|L1|T1|W1|H1|L2|T2|W2|H2|x1|y1|xM1|yM1|xM2|yM2|LäPf1|LäPf2|DrawA
    '                 1  2  3  4  5  6  7  8  9  10 11  12  13  14   15    16    17

    'Vorbereitung
        If OvalValues = "" Then Exit Sub
        Dim Arr1() As String, sH As Object, L!, T!, w!, H!
    'Action
        Arr1 = Split(OvalValues, "|")
        L = Arr1(1): T = Arr1(2): w = Arr1(3): H = Arr1(4)
        Set sH = Sheets("T4").Shapes.AddShape(msoShapeOval, L, T, w, H)
        With sH: .NAME = "Oval1": .Fill.Transparency = 1
                 .Line.ForeColor.RGB = RGB(255, 0, 0): .Line.Weight = 4
        End With
End Sub

Sub Draw_OvalShapeAroundName2(OvalValues$)
    'Called from    JumpFromCompetitorsListToMyDesign
    'Action         Zeichnet das Oval2 (um Name2 herum)
    'OvalValues     =|L1|T1|W1|H1|L2|T2|W2|H2|x1|y1|xM1|yM1|xM2|yM2|LäPf1|LäPf2|DrawA
    '                 1  2  3  4  5  6  7  8  9  10 11  12  13  14   15    16    17

    'Vorbereitung
        If OvalValues = "" Then Exit Sub
        Dim Arr1() As String, sH As Object, L!, T!, w!, H!
    'Action
        Arr1 = Split(OvalValues, "|")
        
        L = Arr1(5): T = Arr1(6): w = Arr1(7): H = Arr1(8)
        If L = 0 Then Exit Sub
        Set sH = Sheets("T4").Shapes.AddShape(msoShapeOval, L, T, w, H)
        With sH: .NAME = "Oval2": .Fill.Transparency = 1
                 .Line.ForeColor.RGB = RGB(255, 0, 0): .Line.Weight = 4
        End With
End Sub

Sub Draw_Arrow1(OvalValues$)
    'Called from    JumpFromCompetitorsListToMyDesign
    'Action         Zeichnet Pfeil1 (zu Name1)
    'OvalValues     =|L1|T1|W1|H1|L2|T2|W2|H2|x1|y1|xM1|yM1|xM2|yM2|LäPf1|LäPf2|DrawA
    '                 1  2  3  4  5  6  7  8  9  10 11  12  13  14   15    16    17

    'Vorbereitung
        If OvalValues = "" Then Exit Sub
        Dim Arr1() As String, sH As Object, L!, T!, w!, H!, x1!, y1!, X2!, y2!, m!, c!, F!
        Arr1 = Split(OvalValues, "|")
        If Arr1(17) = "No" Then Exit Sub                    'DrawArrow? No.
    'Calc
        'm = (arr1(12) - arr1(10)) / (arr1(11) - arr1(9))    'm=(y2-y1)/(x2-x1)
        'c = arr1(10) - m * arr1(9)                          'y1=m*x1+c --> c=y1-m*x1
        F = 0.9 'Faktor für Verkürzung der Pfeillänge
        L = Arr1(15) * F        'L^2=(x2-x1)^2+(y2-y1)^2    'Wunschlänge Pfeil1
        If L > 470 Then F = 470 / Arr1(15)
    'Draw
        x1 = Arr1(9): y1 = Arr1(10)
        X2 = Arr1(9) + (Arr1(11) - Arr1(9)) * F
        y2 = Arr1(10) + (Arr1(12) - Arr1(10)) * F
        Set sH = Sheets("T4").Shapes.AddConnector(msoConnectorStraight, x1, y1, X2, y2)
        With sH: .NAME = "Arrow1": .Line.ForeColor.RGB = RGB(255, 0, 0): .Line.Weight = 4
                 .Line.EndArrowheadStyle = msoArrowheadOpen
        End With
End Sub

Sub Draw_Arrow2(OvalValues$)
    'Called from    JumpFromCompetitorsListToMyDesign
    'Action         Zeichnet Pfeil1 (zu Name1)
    'OvalValues     =|L1|T1|W1|H1|L2|T2|W2|H2|x1|y1|xM1|yM1|xM2|yM2|LäPf1|LäPf2|DrawA
    '                 1  2  3  4  5  6  7  8  9  10 11  12  13  14   15    16    17

    'Vorbereitung
        If OvalValues = "" Then Exit Sub
        Dim Arr1() As String, sH As Object, F!, L!, T!, w!, H!, x1!, y1!, X2!, y2!
        Arr1 = Split(OvalValues, "|")
        If Arr1(17) = "No" Then Exit Sub    'DrawArrow? No.
        If Arr1(16) = "0" Then Exit Sub     'no Oval2
    'Calc
        F = 0.9 'Faktor für Verkürzung der Pfeillänge
        L = Arr1(16) * F        'L^2=(x2-x1)^2+(y2-y1)^2    'Wunschlänge Pfeil1
        If L > 470 Then F = 470 / Arr1(16)
    'Draw
        x1 = Arr1(9): y1 = Arr1(10)
        X2 = Arr1(9) + (Arr1(13) - Arr1(9)) * F
        y2 = Arr1(10) + (Arr1(14) - Arr1(10)) * F
        Set sH = Sheets("T4").Shapes.AddConnector(msoConnectorStraight, x1, y1, X2, y2)
        With sH: .NAME = "Arrow2": .Line.ForeColor.RGB = RGB(255, 0, 0): .Line.Weight = 4
                 .Line.EndArrowheadStyle = msoArrowheadOpen
        End With
End Sub

Function Get_TopOfOvalInDesign(cell$) As Single
    Dim T!
    T = Sheets("T4").Range(cell).Top
    Get_TopOfOvalInDesign = T - 10
End Function

Function Get_LeftOfOvalInDesign(cell$) As Single
    Dim L!
    L = Sheets("T4").Range(cell).Left
    Get_LeftOfOvalInDesign = L - 10
End Function

Function Get_HeightOfOvalInDesign() As Single
    Get_HeightOfOvalInDesign = 35 'bei Zeilenhöhe 15
End Function

Function Get_WidthOfOvalInDesign(cell$) As Single
    Dim N$, s2%, w!
    'Cell       ="JQ103" = Name der T4Zelle, um die das ovalShape gelegt werden soll
    
    With Sheets("T4")
    'OvalWidth gem. Namen, nicht Zelle
        N = Sheets("T4").Range(cell)    'Lea Mai 'Lea Mai (19)
        'Namen in die erste Zelle der ersten Spalte rechts des Dg schreiben
            s2 = CInt(ArrC(40))             'T4 SpaltenNr of DesignBorderRight
            .Cells(1, s2 + 1) = N
        'Autofit
            .Columns(s2 + 1).EntireColumn.AutoFit 'TestSpalte
            'W = Columns(s2 + 1).ColumnWidth 'funktioniert nicht
        w = .Cells(1, s2 + 2).Left - Sheets("T4").Cells(1, s2 + 1).Left
        'ZellBreite wieder zurücksetzen
            .Columns(s2 + 1).ColumnWidth = 2
            DeleteContentOfColumnNr "T4", s2 + 1
        Get_WidthOfOvalInDesign = w + 20
    'OvalWidth gem. Zelle
        'Get_WidthOfOvalInDesign = .Range(Cell).Width + 20
    End With
End Function


