Attribute VB_Name = "T2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit 'T2

'Beep
    Private Declare Function Beep Lib "kernel32" (Optional ByVal dwFreq As Long = 900, Optional ByVal dwDuration As Long = 100) As Long

Sub zzz_T2()
    showProcs "RowNr"
    EE 1: Beep2
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range) 'T2
    'Called from    [UserClick to select another T2-Cell]
    
    'Vorab (bei jedem Klick auf eine andere T4-Zelle)
        'If Target.Cells.count <> 1 Then Exit Sub
        Beep1 'Kontroll-Beep für SelectionChangeEvent-Aufruf
        Dim s%, z%, zNa%, B()
        z = Target.Row: s = Target.Column: DoArrc
    'B(BoxNr 1-6, Items 1-5) 'Items (z, s, StatusSelect, AnzahlZz, StatusZz)
        T2_Load_BoxItems B
    'Kästchen an-/abkreuzen 'ZusatzZeilen ein-/ausblenden
        T2_KästchenAnAbkreuzen z, s, B
    'Klick in Remarks-Area
        If s > 26 And s < 35 Then T2_ClickInto_RemarksArea z, B
    'CellButton "Do all"
        If z = 2 And s = 12 Then T2_ButtonDoAll_wasClicked B: Exit Sub
    'Stop-Button dimmen
        If Cells(2, 12).Interior.Color <> vbRed Then T2_ButtonStop_Dimm
    EE 1
End Sub

Private Sub Worksheet_Activate() 'Sobald dieses Blatt aktiviert/angezeigt wird
    'ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
    'Zeilenhöhe
        SichtbareZeilenHoeheAnpassen
End Sub

Sub T2_Load_BoxItems(B())
    'Called from    Worksheet_SelectionChange
    'Box            z. Zt. insgesamt 6 ankreuzbare Kästchen; Infos dazu in B()
    'B()            = B(BoxNr, BoxItem) = B(1 to 6, 1 to 5)
    'BoxNr          1 = Box zu  Fortlaufende Nummern ...
    '               2 = Box zu  Namen einer Person .....
    '               3 = Box zu  Repariere alle Links ...
    '               4 = Box zu      Rep - Events
    '               5 = Box zu      Rep - ClubsNations
    '               6 = Box zu      Rep - Leute
    'BoxItem        1 = z = ZeilenNr  der Box
    '               2 = s = SpaltenNr der Box
    '               3 = SelectStatus  der Box (0 = nicht angekreuzt)
    '               4 = Anzahl ZusatzZeilen unterhalb z
    '               5 = Status Sichtbarkeit der ZusatzZeilen (0 = ausgeblendet)

    'Vorbereitung
        Dim c%, i%, zFo%, zNa%, zRe%, A()
    'Suche ZeilenNrn von Spalte-2-Boxen
        zFo = Get_RowNr_HoldingMyTextWholeInColumnX("T2", 4, 5, "Fortlaufende Nummern oberhalb Dgs")
        zNa = Get_RowNr_HoldingMyTextWholeInColumnX("T2", 4, 5, "Namen einer Person ändern (global) ...")
        zRe = Get_RowNr_HoldingMyTextWholeInColumnX("T2", 4, 5, "Repariere alle Links (broken .lnk files)")
    'Koordinaten der Box i: z in B(i, 1), s in B(i, 2)
        ReDim B(1 To 6, 1 To 5)
        B(1, 1) = zFo:     B(1, 2) = 2      'Fortlaufende Nummern ...
        B(2, 1) = zNa:     B(2, 2) = 2      'Namen einer Person .....
        B(3, 1) = zRe:     B(3, 2) = 2      'Repariere alle Links ...
        B(4, 1) = zRe + 3: B(4, 2) = 6      'Rep - Events
        B(5, 1) = zRe + 5: B(5, 2) = 6      'Rep - ClubsNations
        B(6, 1) = zRe + 7: B(6, 2) = 6      'Rep - Leute
    'B(i, 3)    'BoxSelectStatus            '1 = isSelected, 0 = isNotSelected
        A = Range(Cells(1, 1), Cells(30, 6)).Value
        For i = 1 To 6
            If A(B(i, 1), B(i, 2)) = "·" Then B(i, 3) = 0 Else B(i, 3) = 1
            If B(i, 2) = 2 And B(i, 3) = 1 Then c = c + 1   'Anzahl ActiveBoxen
            B(i, 4) = 0: B(i, 5) = 2 '0 ZusatzZeilen 'Sichtbarkeit ZusatzZeilen irrelevant
        Next ': showArray2D B
        [B4] = c 'Anzeige Anzahl aktiver Spalte-2-Boxen
    'B(i, 4)    'Anzahl ZusatzZeilen unterhalb der Zeile z
        B(2, 4) = 8: B(3, 4) = 8
    'B(i, 5)    'Status Sichtbarkeit der ZusatzZeilen
        If Rows(zNa + 2).Hidden Then B(2, 5) = 0 Else B(2, 5) = 1
        If Rows(zRe + 2).Hidden Then B(3, 5) = 0 Else B(3, 5) = 1 ': showArray2D B
End Sub

Sub T2_ButtonDoAll_wasClicked(B())
    'Called from    Worksheet_SelectionChange
    'B()            = B(BoxNr, BoxItem) = B(1 to 6, 1 to 5) 'see T2_Load_BoxItems
    'BoxItems       z, s, StatusSelect, AnzahlZz, StatusZz
    
    'Vorbereitung
        Dim T$, i&, z&, Boxes() As Integer
        Call T2_ButtonDoAll_UnDimm: DoEvents: T2_ButtonDoAll_Dimm: EE 0: [B2].Select: EE 1
        If [B4] = 0 Then Exit Sub
    'Button Do_All rot einfärben
        Call T2_ButtonDoAll_Red: T2_ButtonStop_UnDimm
    'Angekreuzte Kästchen ermitteln
        'T2_Get_BoxesSelected Boxes 'Boxes() = 1-Dim, 1-based 'Boxes(1) = 6
    'Boxes abarbeiten
        If B(1, 3) = 1 Then T2_DgNrnOben CLng(B(1, 1))            'Fortlaufende Nummern oberhalb Dgs
        If B(2, 3) = 1 Then T2_ChangeNamePersGlobal CLng(B(2, 1)) 'Namen einer Person ändern (global) ...
        If B(3, 3) = 1 Then T2_BadLinks B                         'Repariere alle Links (broken .lnk files)
    'Boxes sind abgearbeitet
        Call T2_ButtonDoAll_UnDimm: T2_ButtonStop_Dimm
End Sub

Sub T2_Get_BoxesSelected(Arr)
    'Called from    T2_ButtonDoAll_wasClicked

    'Vorbereitung
        Dim s$, i%, zLast&, A(), B() As String
        zLast = Get_NrOfLastRowInColumnNr(2, "T2")
    'Kästchen-Spalte in Array A() nehmen
        A = Range(Cells(6, 2), Cells(zLast, 2)).Value
    'Angekreuzte Kästchen suchen
        For i = 1 To UBound(A, 1)
            If A(i, 1) = "û" Then s = s + "|" + CStr(i + 5)
        Next
    'ZeilenNrn der Fundstellen in B() schreiben
        B = Split(s, "|")
    'Ziel-Array erstellen: Arr() = 1-Dim, 1-based
        ReDim Arr(1 To UBound(B)) As Integer
        For i = 1 To UBound(B): Arr(i) = B(i): Next
End Sub

Sub T2_ButtonDoAll_Dimm()
    Dim r As Range: Set r = Range(Cells(2, 12), Cells(4, 17)): r.Interior.Color = Green1: r.Font.Color = Green3
End Sub

Sub T2_ButtonDoAll_UnDimm()
    Dim r As Range: Set r = Range(Cells(2, 12), Cells(4, 17)): r.Interior.Color = Green3: r.Font.Color = vbBlack
End Sub

Sub T2_ButtonDoAll_Red()
    Dim r As Range: Set r = Range(Cells(2, 12), Cells(4, 17)): r.Interior.Color = vbRed: r.Font.Color = vbBlack
End Sub

Sub T2_ButtonStop_Dimm()
    Dim r As Range: Set r = Range(Cells(2, 19), Cells(4, 24)): r.Interior.Color = Green1: r.Font.Color = Green3
End Sub

Sub T2_ButtonStop_UnDimm()
    Dim r As Range: Set r = Range(Cells(2, 19), Cells(4, 24)): r.Interior.Color = Green3: r.Font.Color = vbBlack
End Sub

Sub T2_KästchenAnAbkreuzen(z%, s%, B())
    'Called from    Worksheet_SelectionChange
    'B()            = 2-Dim-Array mit (z,s,SelectStatus)  aller ankreuzbaren Kästchen
    '               Box 1   Fortlaufende Nummern ...    in Spalte 2
    '               Box 2   Namen einer Person .....    in Spalte 2
    '               Box 3   Repariere alle Links ...    in Spalte 2
    '               Box 4       Rep - Events            in Spalte 6
    '               Box 5       Rep - ClubsNations      in Spalte 6
    '               Box 6       Rep - Leute             in Spalte 6

    'Vorbereitung
        If Not (s = 2 Or s = 6) Then Exit Sub
        If T2_ButtonDoAll_isRed Then Exit Sub
        Dim Box%, i%, r As Range, BoxIsSelected As Boolean
        For i = 1 To UBound(B, 1)
            If B(i, 1) = z And B(i, 2) = s Then Box = i
        Next
        If Box = 0 Then Exit Sub
    'Box i wurde geklickt
        Set r = Sheets("T2").Cells(z, s)
        'Select ändern, Stop-Button dimmen
            EE 0: [B2].Select: EE 1: T2_ButtonStop_Dimm
        'Label hellgrün (bei Boxen mit s = 2))
            'If InStr(1, "|1|2|3|", "|" + CStr(Box) + "|") > 0 Then Cells(z, 4).Interior.Color = Green1: Range(Cells(z, 27), Cells(z, 34)).Interior.Color = Green1
    'Kästchen ankreuzen/abkreuzen
        If B(Box, 3) = 1 Then BoxIsSelected = True
        If BoxIsSelected Then 'Box abkreuzen
                r.Value = "·": r.Font.NAME = "Calibri":   r.Font.size = 11: B(Box, 3) = 0
        Else                    'Box ankreuzen
                r.Value = "û": r.Font.NAME = "Wingdings": r.Font.size = 24: B(Box, 3) = 1
        End If
        BoxIsSelected = Not BoxIsSelected
    'ZusatzZeilen einblenden/ausblenden
        If Box = 2 Then 'Namen einer Person ...
            If BoxIsSelected Then
                'T2_ZusatzZeilenEinblenden_ChangeName
                    Rows(CStr(z + 1) + ":" + CStr(z + 8)).Hidden = False
                T2_ChangeNameTextfeld_ExistenzAbsichern
                T2_ChangeNameTextfeld_Hide 'Textfeld noch nicht anzeigen
                T2_ChangeName_InputFields CInt(B(2, 1)), z, s
            Else
                'T2_ZusatzZeilenAusblenden_ChangeName
                    Rows(CStr(z + 1) + ":" + CStr(z + 8)).Hidden = True
            End If
        End If
        If Box = 3 Then 'Repariere alle Links ...
            If BoxIsSelected Then
                    Rows(CStr(z + 1) + ":" + CStr(z + 8)).Hidden = False
            Else
                    Rows(CStr(z + 1) + ":" + CStr(z + 8)).Hidden = True
            End If
        End If
End Sub

Sub SichtbareZeilenHoeheAnpassen()
    Dim ws As Worksheet, visRange As Range
    Set ws = Worksheets("T2")
    EE 0
    'Prüfen, ob sichtbare Zellen im UsedRange vorhanden sind
        On Error Resume Next
        Set visRange = ws.UsedRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
    ' Nur die sichtbaren Zeilen auf Höhe 15 setzen
        If Not visRange Is Nothing Then
            visRange.RowHeight = 15
        End If
    EE 1
End Sub

Sub T2_Count_BoxesSelect()
    'Vorbereitung
        Dim c%, i%, zLast&, A()
        zLast = Get_NrOfLastRowInColumnNr(2, "T2")
    'Spalte-2-Part in Array A() nehmen
        A = Range(Cells(6, 2), Cells(zLast, 2)).Value
    For i = 1 To UBound(A, 1)
        If A(i, 1) = "û" Then c = c + 1
    Next
    [B4] = c
    If c = 0 Then T2_ButtonDoAll_Dimm Else T2_ButtonDoAll_UnDimm
End Sub


