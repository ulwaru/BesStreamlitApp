Attribute VB_Name = "T4"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "HelperBox, 19, 12, MSForms, ListBox"
Attribute VB_Control = "CmdOpen, 21, 13, MSForms, CommandButton"
Attribute VB_Control = "CmdCreate, 22, 14, MSForms, CommandButton"
Attribute VB_Control = "CmdBack, 23, 15, MSForms, CommandButton"
Attribute VB_Control = "Cmdinfo, 24, 16, MSForms, CommandButton"
Attribute VB_Control = "CmdSingleRanking, 25, 17, MSForms, CommandButton"
Attribute VB_Control = "CellBox, 27, 18, MSForms, TextBox"
Option Explicit 'T4
'Beep
    Private Declare Function Beep Lib "kernel32" (Optional ByVal dwFreq As Long = 900, Optional ByVal dwDuration As Long = 100) As Long

Sub zzz_T4()
    
    showProcs "merge"
    
    'Scroll_ToSeeCell_z_s 32, 275, 3, 3
    'T1_ShowLogBuch
    'show CStr(getColor([nc410], 0, "T4"))
    '[FI106].Font.Color = 14083324 '(w-Farbe, hellrot)
    '[FQ112].Font.Color = 15652797 '(m-Farbe, hellblau)
    'showArray2D ArrDg
    EE 1: Beep2
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range) 'T4
    'Called from    [UserClick to select another T4-Cell]
    'ToDo           Create Buttons (UpdateListOfCompetitors, CreateTextFileOfListOfCompetitors)
    '               No fix values  (UpdateListOfCompetitors, CreateTextFileOfListOfCompetitors)
    
    'Vorab (bei jedem Klick auf eine andere T4-Zelle)
        'If Target.Cells.count <> 1 Then Exit Sub
        Beep1 'Kontroll-Beep für SelectionChangeEvent-Aufruf
        Ticks "Worksheet_SelectionChangeT4"
        Dim A$, s%, z%, z1Dg%, s1Dg%
        z = Target.Row: s = Target.Column: DoArr
        'z, s in ArrC eintragen
            FillArrC 106, ArrC(104): FillArrC 107, ArrC(105)
            FillArrC 104, CStr(z): FillArrC 105, CStr(s)
        'LastSelection
            T4_LastSelection z, s
        'Verschiebung entdecken: Zelle mit "List of DesignTitles"
            T4_Save_PositionOfHeaderCell_OfListOfDesignTitles_ToArrC
        'Vorbereitung
            z1Dg = CInt(ArrC(71))      'T4 ZeilenNr  Cell "List of DesignTitles"   '5
            s1Dg = CInt(ArrC(72))      'T4 SpaltenNr Cell "List of DesignTitles"   '2
        'Buttons, Ovale - Visible/Delete
            With Sheets("T4"): .Activate: .Shapes("T4RedRect").Visible = False
                'Vermeidbar?:
                .CmdCreate.Visible = False: .CmdOpen.Visible = False: .CmdBack.Visible = False
                .Cmdinfo.Visible = False: .CmdSingleRanking.Visible = False
            End With
            Call Hide_CellBox: Hide_HelperBox
            Delete_T4Shape "Oval1":  Delete_T4Shape "Oval2"
            Delete_T4Shape "Arrow1": Delete_T4Shape "Arrow2"
    'CellButtons
        'new    in T4-Zeilen 4/5
            If Cells(z, s) = "new" Then T4_CellBtn_new CLng(z), CLng(s): Exit Sub
        'LU ub
            If Cells(z, s) = "LU" Then T4_CellBtn_LU z, s: Exit Sub
            If Cells(z, s) = "ub" Then T4_CellBtn_ub z, s: Exit Sub
        'Spaltenbreite T4SomeDgData
            If z = 1 And s = 3 Then T4_Set_WidthOf_T4SomeDgData: Exit Sub
        'Line 2 Info (User clicks on a T4-Line2-Cell with value = "i")
            If z = 2 And Cells(z, s) = " i" Then T4Line2Info: Exit Sub
            If z = 2 And Cells(z, s) = "i " Then T4Line2Info: Exit Sub
        'Jump from ListOfDesignTitles to Design
            If z > z1Dg + 2 And s = s1Dg Then T4_JumpToMyDesign z, s, Target: Exit Sub
        'Update ListOfDesignTitles
            If (z = z1Dg Or z = z1Dg + 1) And s = s1Dg Then T4_Write_ListOfDesignTitles_FromT4SomeDgData: Exit Sub
        'Einzel -->, Synchron -->, Mannschaft -->, ...
            A = Cells(z, s)   'Chr(34) = " = Zeichen für PfeilRechts in Wingdings 3
            If A = "Einzel " + Chr(34) Then [E19].Select: Exit Sub
            If A Like "* " + Chr(34) + " " Or A Like "! *" Then JumpToDesignTyp A: Exit Sub
    'T4_Handle_UserClick
        If z > z1Dg + 2 And s > 19 Then T4_Handle_UserClick z, s
    EE 1
End Sub

Private Sub Worksheet_Activate() 'Sobald dieses Blatt aktiviert/angezeigt wird
    'ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
    'NeuErstellung aller T4ListOfDgTitles, T4SomeDgData-Zeilen
        DoArrc
        T4_Write_T4SomeDgData
        T4_Write_ListOfDesignTitles_FromT4SomeDgData
    'Was macht dies?
        StartVBS_HandleSameBaseNameFolders_InsideOneFolder ArrC(4) 'Leute
End Sub

Sub StartVBS_HandleSameBaseNameFolders_InsideOneFolder(PathOfOneFolder$)
    'vbs parameter
    Dim vbsPath$, shell As Object
    vbsPath = ArrC(6) + "\vbs\HandleSameBaseNameFolders_InsideOneFolder.vbs" 'helpers
    Set shell = CreateObject("WScript.Shell")
    'Startet das VBScript und übergibt den Pfad in Anführungszeichen
        shell.Run Chr(34) & vbsPath & Chr(34) & " " & Chr(34) & PathOfOneFolder & Chr(34)
End Sub

Sub T4_CellBtn_new(z&, s&)
    'Called from    Worksheet_SelectionChange[T4] (Klick auf "new" in T4-Zeilen 4/5)
    'Action         Ordnet alle Dgs dieser Gruppe nach Datum aufwärts;
    '               ermöglicht Erstellung eines neuen Dg unterhalb der Gruppe
    
    'Sort Dgs
        T4_Sort_Dgs_OfOneDgGroup CInt(z), CInt(s)
    'Neues Dg
        Create_NewDg z, s
End Sub

Sub T4_Sort_Dgs_OfOneDgGroup(z%, s%)
    'Called from    T4_CellBtn_new
    'Action         Ordnet alle Dgs dieser Gruppe nach Datum aufwärts
    'ToDo           1 Leerzeile zw. Dgs
    
    'Vorbereitung
        Dim Date8$, DgTitle$, GrpName$, L$, m$, R_old$, R_new$, RO$, T$, ub$, v$
        Dim i%, s1%, s2%, sLO%, sLU%, sRO%, sRU%
        Dim z1%, z2%, zLast%, zLO%, zLU%, zRO%, zRU%
        Dim ArrT() As String, ArrT2() As String, arrRO() As String, arrUB() As String
        With Sheets("T4"): v = vbCrLf
        zLast = Get_NrOfLastRowInColumnNr(s + 1, "T4")
        sRO = s + 1
    'z-Positionen von "RO" und "ub" in Spalte sRO
        For i = 8 To zLast
            T = CStr(.Cells(i, sRO))
            If T = "RO" Then
                'nur beim ersten Vorkommen von "RO":
                'sLO ermitteln (SpaltenNr DgAnfang)
                    If RO = "" Then sLO = T4_Get_sLO_From_zsRO(i, sRO)
                RO = RO + "|" + CStr(i)                 'zNr='| 8|78| 92|107|130|158|187|218
            End If
            If T = "ub" Then ub = ub + "|" + CStr(i)    'zNr='|75|89|104|127|155|184|215|261
        Next
    'Status: Alle DgPositionen (LO, RO, LU, ub) sind ermittelt
    'Alle Titel ermitteln (wegen Anordnung nach Datum)
        arrRO = Split(RO, "|"): T = ""
        arrUB = Split(ub, "|")
        For i = 1 To UBound(arrRO)
            zLO = CInt(arrRO(i)): zRO = zLO
            zLU = CInt(arrUB(i)): zRU = zLU: sRU = sRO: sLU = sLO
            'alle verfügbar für DgAktuell: zLO zRO zLU zRU sLO sRO sLU sRU
            'DgAußenNrn
                T4_Add_Dg_BorderNrs zLO, zRO, zLU, zRU, sLO, sRO, sLU, sRU
            
            DgTitle = CStr(.Cells(zLO + 1, sLO + 1))
            Date8 = T9_Get_Date8FromEventFolderName(DgTitle)
            'M = Meisterschaft (BadWüM, DJM, DSM, DM, JEM4, ...)
                m = Mid(DgTitle, InStr(1, DgTitle, " ") + 1)
                m = Left(m, InStr(1, m, " ") - 1)
            'a, b, c für SortierReihenfolge
                If m Like "*SM*" Then
                    m = "a"
                ElseIf m Like "*JM*" Then m = "b"
                Else: m = "c"
                End If
            T = T + "|" + Date8 + "|" + m + "|" + DgTitle + "|" + Format(i, "00") _
                  + "|" + Format(zLO - 1, "0000") + "|" + Format(sLO - 1, "0000") _
                  + "|" + Format(zRU, "0000") + "|" + Format(sRU, "0000") + v
        Next i
        'T bereinigen
            If Right(T, 2) = v Then T = Left(T, Len(T) - 2)
    'sort ArrT      9 Spalten: ArrT(0 to 8)
        ArrT = Split(T, v) 'ArrT(3)=Title
        'showArray ArrT
        QuickSort ArrT
        'showArray ArrT
        'ArrT(3)=Title herausnehmen
        '|19690800|c|1969-08 Nissen-Cup11 Grenchen/CH 6 (Einzel)|06|0116|0019|0125|0026
            For i = 0 To UBound(ArrT)
                L = ArrT(i) 'i. Zeile aus ArrT
                ArrT(i) = Left(L, 12) + Right(L, 22)
            Next
            'showArray ArrT
        For i = 0 To UBound(ArrT)
            R_old = R_old + Format(i + 1, "00")     '12345678  alte Reihenfolge
            R_new = R_new + Mid(ArrT(i), 13, 2)     '85674231  neue Reihenfolge
        Next
        'showArray ArrT
    'Cut and Paste nötig?
        If R_old = R_new Then Exit Sub 'Reihenfolge ist bereits sortiert
    'ArrT   9 Spalten: ArrT(0 to 8)
        'show T + v + v + Get_1DArrayToString(ArrT) + v + v + R_old + v + R_new
        '|19720506|a|08|0217|0178|0261|0211|1972-05-06 DSM Datteln (Einzel)
        '|19800426|a|05|0129|0178|0155|0211|1980-04-26 DSM Dillenburg (Einzel)
        '|19800426|b|06|0157|0178|0184|0211|1980-04-26 DJM Dillenburg (Einzel)
        '|19800510|c|07|0186|0178|0215|0211|1980-05-10 DM Dahn_Pfalz (Einzel)
        '|19870603|a|04|0106|0178|0127|0211|1987-06-03 DSM Berlin (Einzel)
        '|19870603|b|02|0077|0178|0089|0211|1987-06-03 DJM Berlin (Einzel)
        '|19870603|c|03|0091|0178|0104|0211|1987-06-03 DM Berlin (Einzel)
        '|19960420|c|01|0007|0178|0075|0211|1996-04-20 BadWüM Gernsbach (Einzel)
        'In Zeile 1 steht in der 4. bis 7. Spalte der Cut-Range1 (z1,s1,z2,s2)
        'Range1 wird unter die letzte T4-Zeile geschoben (wird das 1. Dg der Gruppe)
        'Cut-Range2 (in Zeile 2) wird unter das verschobene Range1 gesetzt, u.s.w.
    'Cut and Paste
        For i = 0 To UBound(ArrT)
            zLast = Get_NrOfLastRowInColumnNr(s + 1, "T4")
            L = ArrT(i) 'i. Zeile aus ArrT
            z1 = CInt(Mid(L, 16, 4)): s1 = CInt(Mid(L, 21, 4))
            z2 = CInt(Mid(L, 26, 4)): s2 = CInt(Mid(L, 31, 4)):  EE 0
            Cut_Range_PasteToCell z1, s1, z2, s2, zLast + 2, s1: EE 1
        Next
    'Neu geordnete Gruppe nach oben schieben
        z1 = Get_RowNr_HoldingMyTextWholeInColumnXSearchDown("T4", sRO, 7, "RO") - 1
        z2 = Get_NrOfLastRowInColumnNr(CLng(sRO), "T4")
        Cut_Range_PasteToCell z1, sLO - 1, z2, sRO, 7, sLO - 1
    'Update
        T4_Write_T4SomeDgData
        T4_Write_ListOfDesignTitles_FromDgs
        
        
        
        
        
        'GrpName = .Cells(z, s - 2)
    'Finals
        End With
    Beep
End Sub

Sub Cut_Range_PasteToCell(z1%, s1%, z2%, s2%, zPaste%, sPaste%)
    ' Deklaration der Quell- und Zielbereiche
    Dim sourceRange As Range
    Dim targetCell As Range
    
    ' Definition des Quellbereichs von (z1, s1) bis (z2, s2)
    Set sourceRange = Range(Cells(z1, s1), Cells(z2, s2))
    
    ' Definition der oberen linken Zielzelle
    Set targetCell = Cells(zPaste, sPaste)
    
    ' Bereich verschieben
    sourceRange.Cut Destination:=targetCell
End Sub

Function T4_Get_sLO_From_zsRO(zRO%, sRO%) As Integer
    'Vorbereitung
        Dim s%
        With Sheets("T4")
    'Zellen der Zeile zRO nach links nacheinander lesen
        For s = sRO To 1 Step -1
            If CStr(.Cells(zRO, s)) = "LO" Then
                T4_Get_sLO_From_zsRO = s: Exit Function
            End If
        Next
    'Finals
        End With
End Function

Sub T4_Set_WidthOf_T4SomeDgData()
    Dim w!
    w = Columns(4).ColumnWidth
    If w < 1 Then Columns("D:P").ColumnWidth = 2 Else Columns("D:P").ColumnWidth = 0.3
    [D1].Select
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    'Erledigt-Häkchen entfernen (bei CompetitorsList, da T4-Change)
        Sheets("T3").[r3] = "": Sheets("T3").[r7] = ""
End Sub

Sub T4_LastSelection(zSh%, sSh%)
    'Called from    Worksheet_SelectionChange[T4]
    'zSh, sSh       aktueller Click; T4-SheetZeilenNr, T4-SheetSpaltenNr
    'Action         schreibt die Werte des ActualClick (zSh, sSh)
    '               innerhalb einer neuen Klickfolge in ArrC
    
    'Vorbereitung
        Dim A$
    'Bisheriges a       'Last 3 Clicks
        A = ArrC(64)    '|0071|0380|0072|0380|1234|0567| = |zLast0|sLast0|zLast1|sLast1|zLast2|sLast2|
    'Neues a            neue Anordnung: ActualClick    --> Last0 of new a (Links)
        A = "|" + Format(zSh, "0000") + "|" + Format(sSh, "0000") + Left(A, 21)
    'Write
        FillArrC 64, A
End Sub

Sub T4_GotoLastSelection()
    Cells(CInt(Mid(ArrC(64), 12, 4)), CInt(Mid(ArrC(64), 17, 4))).Select
End Sub

Sub JumpToDesignTyp(A$)
    Dim z&, s&
    If A Like "*Einzel*" Then
        z = Get_RowNr_HoldingMyTextWhole("T4", "Einzel " + Chr(34))
        s = Get_ColumnNr_HoldingMyTextWhole("T4", "Einzel " + Chr(34))
        Cells(z + 17, s).Select
    ElseIf A Like "*Mannschaft*" Then
        z = Get_RowNr_HoldingMyTextWhole("T4", "Mannschaft " + Chr(34))
        s = Get_ColumnNr_HoldingMyTextWhole("T4", "Mannschaft " + Chr(34))
        Cells(z + 17, s).Select
    ElseIf A Like "*Synchron*" Then
        z = Get_RowNr_HoldingMyTextWhole("T4", "Synchron " + Chr(34))
        s = Get_ColumnNr_HoldingMyTextWhole("T4", "Synchron " + Chr(34))
        Cells(z + 17, s).Select
    End If
End Sub

Private Sub CellBox_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Called from    [User hits a key while cursor inside CellBox]
    'Status         User klickte auf eine T4-Design-Damen/Herren-Zelle;
    '               HelperBox wurde eingeblendet (HelperBoxInit durchgeführt)
    
    'Vorbereitung
        'MsgBox KeyCode
        Dim Temp$, s%, z%, cL!, cT!, cW!, arrFilter() As String, MyCell As Range
        DoArrc
        s = CInt(ArrC(105)): z = CInt(ArrC(104)): Set MyCell = Sheets("T4").Cells(z, s)
    'Einzelnen Tastendruck abfangen
        Select Case keyCode
            Case vbKeySubtract
                'Minuszeichentaste (-) auf der ZehnerTastatur
                CellBox.Value = CellBox.Value + "-"
            Case 189
                'Minuszeichentaste (-) auf der StandardTastatur
                CellBox.Value = CellBox.Value + "-"
            Case vbKeyBack
                'Rücktaste - letzten Buchstaben in CellBox löschen
                If Len(CellBox.Value) > 0 Then CellBox.Value = Left(CellBox.Value, Len(CellBox.Value) - 1)
            Case vbKeySpace, vbKeyA To vbKeyZ, vbKey0 To vbKey9, vbKeyDecimal
                'Buchstabe, Ziffer, Leerzeichen - hinzufügen
                CellBox.Value = WorksheetFunction.Proper(CellBox.Value & Chr(keyCode))
            Case vbKeyTab
                'HelperBox nicht benutzt; neuen Namen in T5 aufnehmen
                MyCell.Value2 = CellBox.Value
                Sheets("T4").Cells(z, s + 1).Select
            Case vbKeyReturn
                'Enter/Return - Einträge der CellBox in dahinterliegende Zelle übernehmen
                '             - CellBox verlassen, Boxen ausblenden
                If HelperBox.ListCount > 0 Then
                    'Angeklickten Namen aus der Helperbox übernehmen
                    MyCell.Value2 = HelperBox.List(0)
                Else
                    'HelperBox nicht benutzt; neuen Namen in T5 aufnehmen
                    MyCell.Value2 = CellBox.Value
                End If
                Call Hide_CellBox: Hide_HelperBox
                MyCell.Offset(1, 0).Activate
                Exit Sub
            Case vbKeyEscape
                Call Hide_CellBox
                Hide_HelperBox
                MyCell.Activate
            Case 190
                CellBox.Value = CellBox.Value & "."
            Case 222
                If Shift = 0 Then CellBox.Value = CellBox.Value & "ä"
                If Shift = 1 Then CellBox.Value = CellBox.Value & "Ä"
            Case 192
                If Shift = 0 Then CellBox.Value = CellBox.Value & "ö"
                If Shift = 1 Then CellBox.Value = CellBox.Value & "Ö"
            Case 186
                If Shift = 0 Then CellBox.Value = CellBox.Value & "ü"
                If Shift = 1 Then CellBox.Value = CellBox.Value & "Ü"
            Case 219: CellBox.Value = CellBox.Value & "ß"
            Case Else
                'CellBox.Value = WorksheetFunction.Proper(CellBox.Value & Chr(KeyCode))
        End Select
    'HelperBox wird neu belegt (nur passende Namen)
        HelperBox.Clear
        Fill_ArrHelperBox CellBox.Value, arrFilter()
        If arrFilter(0) = "No names" Then Hide_HelperBox: keyCode = 0: Exit Sub
        'On Error Resume Next 'falls arrFilter empty
        HelperBox.List = arrFilter
    'HelperBox neu formatieren
        With Sheets("T4").Cells(z, s)
            cL = .Left: cT = .Top: cW = .Width            'c CellValues
        End With
        With Sheets("T4").Shapes("HelperBox")
            .Left = cL + cW + 5
            '.Top = CT - 30
            .Height = (UBound(arrFilter) + 1) * 12 '200
                If .Height < 75 Then .Height = 75
                If .Height > 196 Then .Height = 196
            .Width = 15 + Len(arrFilter(0)) * 5.5
        End With
    'Repair
        T4_Replace_Helperbox 'Warum nötig?
    'HelperBox nur anzeigen, falls Inhalt vorhanden
        If HelperBox.ListCount = 0 Then
            Hide_HelperBox
        Else
            HelperBox.Visible = True
        End If
    keyCode = 0
End Sub

Private Sub HelperBox_Click()
    'Called from    [User clicks one line inside HelperBox]
    'CellBox        = TextBox; liegt direkt über der aktuell angeklickten Zelle
    'HelperBox      = ListBox; trägt herausgefilterte Zeilen (mit Namen, Jhg, ...)
    'L              = Line = 1 angeklickte Zeile der HelperBox-Zeilen
    '               = "Mai|Lea|D|TV Oos|1962|w||" oder "Meier|Max||||||"
    '               = 8 Items, gebildet aus dem aktuellen DatenStand in PersonData.txt
    'E              = Eintrag für die aktuell angeklickten Zelle; gebildet aus L
    'Action         Eintrag in die aktuell aktive NamenZelle, je nach HelperBox-Auswahl
    
    'Vorbereitung
        Dim Alter$, E$, Jhg$, L$, mw$, Nation$, Nn$, Verein$, Vn$, s%, sDg%, y%, z%
        Dim MyCell As Range, r1 As Range, Arr1() As String, MyColor
        DoArrDg
        z = CInt(ArrC(104))              'Sheet-ZeilenNr  der momentan aktiven Zelle
        s = CInt(ArrC(105))              'Sheet-SpaltenNr der momentan aktiven Zelle
        sDg = s - CInt(ArrC(39)) + 1    'Dg-SpaltenNr    der momentan aktiven Zelle
        Set MyCell = Sheets("T4").Cells(z, s)
    'L auswerten (angeklickte Zeile in der HelperBox)
        L = HelperBox.Value     'L = "Lea|Mai|D|TV Oos|1962|w"
        Arr1 = Split(L, "|")    '      0   1  2   3     4   5
        Vn = Trim(Arr1(0)):   Nation = Trim(Arr1(2)):   Jhg = Trim(CStr(Arr1(4)))
        Nn = Trim(Arr1(1)):   Verein = Trim(Arr1(3)):    mw = Trim(Arr1(5))
    'Verein     Falls VereinSpalte rechts neben NamenSpalte
        If Verein <> "" And ArrDg(5, sDg + 1) = "V" Then
            Cells(z, s + 1) = Verein        'Verein direkt ins Design schreiben
        End If
    'Nation     Falls NationSpalte rechts neben NamenSpalte
        If Nation <> "" And "STU" Like "*" + ArrDg(5, sDg + 1) + "*" Then
            Cells(z, s + 1) = Nation        'Nation direkt ins Design schreiben
            'ArrDg(zDg, sDg + 1) = Nation   'Nation in ArrDg schreiben
        End If
    'Jahrgang
        If Jhg = "" Then
            E = Vn + " " + Nn
        Else
            y = CInt(Left(ArrC(41), 4)) 'Y = Jahr des Events
            Alter = CStr(y - Jhg): E = Vn + " " + Nn + " (" + Alter + ")"
        End If
    'Liga           m/w/x-Farbe
        If ArrC(42) Like "* Liga\*" Then
            'Arr1(5) = Spalte für m/w innerhalb der angeklickten HelperZeile
            If ArrDg(2, 1) = "M0V1" Then
                Set r1 = Range(Cells(z, s), Cells(z, s + 1))
            Else
                Set r1 = Range(Cells(z, s), Cells(z, s + 8))
            End If
            If Arr1(5) = "" Then
                                      r1.Interior.Color = 15921906 '(x-Farbe, hellgrau)
            ElseIf Arr1(5) = "w" Then r1.Interior.Color = 14083324 '(w-Farbe, hellrot)
            ElseIf Arr1(5) = "m" Then r1.Interior.Color = 15652797 '(m-Farbe, hellblau)
            End If
        End If
jump1:
    MyCell.Value2 = E           '"Lea Mai" oder "Lea Mai (1962)" wird eingetragen
    Call Hide_CellBox:      CellBox.Value = ""
    Call Hide_HelperBox:    HelperBox.Clear
    T4_AutoFit
    MyCell.Offset(1, 0).Activate
End Sub

Sub CreateResultsTxt()
    'Called from    ClickButton_CreateJpg_InActualDg
    'ArrC           36   "CS8:DV38"    T4Range = ArrC(36)       'T4 Range of choosen design
    '               37   "8"           z1      = CInt(ArrC(37)) 'T4 RowNr    of BorderTop
    '               38   "38"          z2      = CInt(ArrC(38)) 'T4 RowNr    of BorderBottom
    '               39   "97"          s1      = CInt(ArrC(39)) 'T4 ColumnNr of BorderLeft
    '               40   "126"         s2      = CInt(ArrC(40)) 'T4 ColumnNr of BorderRight


    'Vorbereitung
        Dim A$, L$, NameOfEventFolder$, OneCell$, OneLine$, Rang$
        Dim SpInhalt$, SpLeer$, SpRang$, SpWert$
        Dim TitleOfResultJpg$, Trennlinie$, v$, ZeLeer$
        Dim Arr1(), Arr2() As String, r As Range, ws As Worksheet
        Dim Sp%, SpEinzel%, SpSynchron%
        Dim Ze%, ZeEinzel%, ZeSynchron%, KopfZeile%
        Set ws = Sheets("T4"): Set r = ws.Range(ArrC(36)) 'Range of choosen design
        Arr1 = r.Value: v = vbCrLf: ZeLeer = ",": SpWert = ","
        Trennlinie = String(155, "-") + v
    'TitleHandling
        TitleOfResultJpg = CStr(Arr1(2, 2))
        NameOfEventFolder = CStr(Arr1(1, 2))
        If NameOfEventFolder = "" Then NameOfEventFolder = Replace(TitleOfResultJpg, "/", "_")
        Arr1(1, 2) = "": Arr1(2, 2) = ""
    'Read Cells of Array      'letzte Zeile (MetaDaten oder leer) wird nicht ins Array aufgenommen
        For Ze = 1 To UBound(Arr1, 1) - 1
            OneLine = ""
            For Sp = 1 To UBound(Arr1, 2)
                'Schleife über alle Spalten einer Zeile
                OneCell = CStr(Arr1(Ze, Sp))
    'WertungsSpalten
                If InStr(1, OneCell, ",") > 0 And InStr(1, SpWert, "," + CStr(Sp) + ",") = 0 Then SpWert = SpWert + CStr(Sp) + ","
                
'                If InStr(1, OneCell, "Einzel") > 0 Then SpEinzel = Sp:     ZeEinzel = Ze
'                'SpEinzel auffüllen
'                    If Sp = SpEinzel And Ze > ZeEinzel And ZeSynchron = 0 Then
'                        If OneCell <> "Synchron" Then arr1(Ze, Sp) = "Einzel"
'                        OneCell = CStr(arr1(Ze, Sp))
'                    End If
'                If InStr(1, OneCell, "Synchron") > 0 Then SpSynchron = Sp: ZeSynchron = Ze
'                If InStr(1, OneCell, "Damen") > 0 Then SpDamen = Sp
'                If InStr(1, OneCell, "Herren") > 0 Then SpHerren = Sp
                
                OneLine = OneLine + OneCell + "|"
            Next
            'OneLine (1 Zeile der ResultTabelle) ist komplett gelesen
    'LeerZeilen
                L = Replace(OneLine, "|", ""): L = Replace(L, " ", "")
                If L = "" Then ZeLeer = ZeLeer + CStr(Ze) + ","
                'ZeLeer = ",1,2,3,5,20,"
    'Kopfzeile
                If KopfZeile = 0 Then
                    If InStr(1, OneLine, "Damen") > 0 Then
                        KopfZeile = Ze
                        ElseIf InStr(1, OneLine, "Jugendturner") > 0 Then KopfZeile = Ze
                        ElseIf InStr(1, OneLine, "Schüler") > 0 Then KopfZeile = Ze
                        ElseIf InStr(1, OneLine, "Synchron Mixed") > 0 Then KopfZeile = Ze
                        ElseIf InStr(1, OneLine, "LTV") > 0 Then KopfZeile = Ze
                        Else: KopfZeile = 0
                    End If
                End If
    'a wird gefüllt
                A = A + OneLine + v
        Next
        
    'LeerSpalten
        SpLeer = ","
        For Sp = 1 To UBound(Arr1, 2)
            SpInhalt = ""
            For Ze = 1 To UBound(Arr1, 1) - 1
                SpInhalt = SpInhalt + Trim(CStr(Arr1(Ze, Sp)))
            Next
            If SpInhalt = "" Then SpLeer = SpLeer + CStr(Sp) + ","
        Next
        'SpLeer = ",1,6,8,11,14,16,20,22,25,28,30,"
    'RangSpalten
        SpRang = ","
        For Sp = 1 To UBound(Arr1, 2)
            SpInhalt = ""
            If InStr(1, SpLeer, "," + CStr(Sp) + ",") = 0 Then
                For Ze = 1 To UBound(Arr1, 1)
                    OneCell = Trim(CStr(Arr1(Ze, Sp)))
                    If OneCell Like "#" Or OneCell Like "##" Or _
                        OneCell Like "###" Or OneCell = "" Then SpInhalt = SpInhalt + "1" _
                        Else SpInhalt = SpInhalt + "0"
                Next
                If Replace(SpInhalt, "1", "") = "" Then SpRang = SpRang + CStr(Sp) + ","
            End If
            
        Next
        'SpRang = ",3,17,"
        
'        arr2 = Split(SpRang, ",")
'        If UBound(arr2) = 2 Then SpRangDamen = CInt(arr2(1)): SpRangHerren = CInt(arr2(1))
'        If UBound(arr2) = 3 Then SpRangDamen = CInt(arr2(1)): SpRangHerren = CInt(arr2(2))
        
        
'    'Show
'        a = "results.txt                      (wurde für die maschinelle Lesbarkeit erstellt.)" + v + v _
'          + TrennLinie _
'          + "Title of ResultJpg:              " + TitleOfResultJpg + v _
'          + "Name of EventFolder:             " + NameOfEventFolder + v _
'          + "ZeilenNr   der Kopfzeile:        " + CStr(KopfZeile) + v _
'          + "ZeilenNrn  der Leerzeilen:       " + ZeLeer + v _
'          + "SpaltenNrn der Leerspalten:      " + SpLeer + v _
'          + "SpaltenNrn der Rang-Spalten:     " + SpRang + v _
'          + "SpaltenNrn der Werte-Spalten:    " + SpWert + v + v _
'          + TAB_Simulation(a) + v + TrennLinie
'        show a
        'N = "Name of EventFolder:             " + NameOfEventFolder
End Sub

Private Sub CmdInfo_Click()
    Dim p$
    p = ArrC(6) 'F:\Archiv Trampolin prog\helpers 'T0 Path of Folder "helpers"

    p = ArrC(6) + "\T4DgInfo.txt"
    openFile p
End Sub

Private Sub CmdBack_Click()
    Dim Title$, z%
    DoArrc
    Title = ArrC(41)            'T4Dg Titel to see inside design and jpg
    z = Get_RowNr_HoldingMyTextWholeInColumnX("T4", 5, 7, Title)
    If z = 0 Then z = 5
    Cells(z, 1).Select
End Sub

Private Sub CmdOpen_Click()
    DoArrc
    OpenFolder ArrC(57)
End Sub

Private Sub CmdSingleRanking_Click()
    'Called from    [UserClick on 'SingleRanking'-Button inside a Design]
    'SingleRanking  gibt es nur bei Mannschafts-Designs;
    '               ein ButtonClick blendet für jeden Competitor
    '               seinen Rang in der Einzelwertung ein/aus;
    '               Einblendung sinnvoll, falls Einzelwertungen vollständig vorhanden
    
    'Vorbereitung
        Dim z1%, s1%: DoArrc
        z1 = CInt(ArrC(37))         'T4 Dg ZeilenNr   of DesignBorder-Top
        s1 = CInt(ArrC(39))         'T4 Dg SpaltenNr  of DesignBorder-Left
        With Sheets("T4")
    'T4-SingleRanking-Vermerk-Zelle auf weiß setzen
        .Cells(z1, s1 + 24).Font.Color = vbWhite
    'Action
        If ArrDg(3, 1) = "SingleRanking is OFF" Then
            ArrDg(3, 1) = "SingleRanking is ON"
            .Cells(z1 + 2, s1) = "SingleRanking is ON"
            'SingleRanking einschalten
            T4_M2V8_SingleRanking
        Else
            ArrDg(3, 1) = "SingleRanking is OFF"
            .Cells(z1 + 2, s1) = "SingleRanking is OFF"
            'SingleRanking ausschalten
            'T4-Zellen leeren
            .Range(Cells(z1 + 6, s1 + 12), Cells(z1 - 2 + UBound(ArrDg), s1 + 12)) = "" 'Spalte 12
            .Range(Cells(z1 + 6, s1 + 25), Cells(z1 - 2 + UBound(ArrDg), s1 + 25)) = "" 'Spalte 25
        End If
    'Finals
        End With
        'Stop
End Sub

Private Sub CmdCreate_Click()
    'Called from    [UserClick on 'Create jpg'-Button inside a Design]
    ClickButton_CreateJpg_InActualDg
End Sub
    
Sub T4Line2Info()
    Dim p$
    p = ArrC(6) + "\T4Line2Info.txt"
    openFile p
End Sub

Sub T4_JumpToMyDesign(z%, s%, Target As Range)
    'Called from    Worksheet_SelectionChange[T4]
    'z, s           = T4-ZeilenNr/-SpaltenNr in 'List of DesignTitles'; s = 2
    'Action         Springt zum Dg des angeklickten DesignTitels;
    '               aktiviert das Dg
    'Status         "List of DesignTitles"  trägt                    jew.     DesignTitel+" "
    '               "T8SomeDgData"          trägt                    jew.     DesignTitel+"  "
    '               "List of competitors"   trägt in der EventSpalte jew. " "+DesignTitel+" "
    '               Innerhalb eines Designs trägt die Zelle (2,2)    jew.     DesignTitel
    Dim Titel$, s1%, z1%
    Titel = Trim(Cells(z, s)) 'Titel aus EventListe
    If Titel Like "####*" Then
        'Dg-Position von T4SomeDgData holen
            z1 = Cells(z, 6): s1 = Cells(z, 7) '[z1, s1] = "LO"
        EE 0
        Application.GoTo Reference:=Worksheets("T4").Cells(z1 - 2, s1 - 2), Scroll:=True
        EE 1
        Cells(z1, s1).Select
    End If
End Sub

Sub T4_Add_Dg_BorderNrs(zLO%, zRO%, zLU%, zRU%, sLO%, sRO%, sLU%, sRU%)
    Dim i%, r As Range
    With Sheets("T4")
    'Nrn schreiben oben
        Set r = .Range(Cells(zLO - 1, sLO), Cells(zRO - 1, sRO))
        r.Font.size = 6: r.HorizontalAlignment = xlCenter
        For i = 1 To sRO - sLO + 1
            .Cells(zLO - 1, sLO + i - 1) = i
        Next
    'Nrn schreiben links
        Set r = .Range(Cells(zLO, sLO - 1), Cells(zLU, sLU - 1))
        r.Font.size = 6: r.HorizontalAlignment = xlCenter
        For i = 1 To zLU - zLO + 1
            .Cells(zLO + i - 1, sLO - 1) = i
        Next
    End With
End Sub

Sub Fokus_VBE_zurücksetzen(vbeWarSichtbar As Boolean, aktuellesCodeFenster As Object)
    'Nur aufräumen, wenn der VBE vorab auch wirklich offen/sichtbar war
    If vbeWarSichtbar Then
        On Error Resume Next
        'Designer der Userform schließen
            ThisWorkbook.VBProject.VBComponents("UF5").DesignerWindow.Close
        'Nur zurückspringen, wenn wir ein Fenster gemerkt haben
            If Not aktuellesCodeFenster Is Nothing Then aktuellesCodeFenster.show
        On Error GoTo 0
    End If
End Sub

Function T4_Get_S_J_or_E_from_DgTitle(DgTitle$) As String
    'Called from    Create_NewDg
    
    'Vorbereitung
        Dim m$, c%
    'M = Meisterschaft (BadWüM, DJM, DSM, DM, JEM4, ...)
        m = Mid(DgTitle, InStr(1, DgTitle, " ") + 1)
        c = InStr(1, m, " ")
        If c > 0 Then m = Left(m, c - 1) 'DSM
    'Get S or J or E
        If m Like "*S*" Then
                                 T4_Get_S_J_or_E_from_DgTitle = "S"
        ElseIf m Like "*J*" Then T4_Get_S_J_or_E_from_DgTitle = "J"
        Else:                    T4_Get_S_J_or_E_from_DgTitle = "E"
        End If
End Function

Sub Copy_Range_PasteToCell(z1&, s1&, z2&, s2&, zPaste&, sPaste&)
    Dim srcRange As Range, destCell As Range
    With Sheets("T4"): EE 2
    Set srcRange = .Range(.Cells(z1, s1), .Cells(z2, s2))
    Set destCell = .Cells(zPaste, sPaste)
    'Kopieren und Einfügen in einer Zelle (schont die Zwischenablage)
        srcRange.Copy Destination:=destCell
    End With: EE 3
End Sub

Sub Create_NewDg(z&, s&)
    'Called from    T4_CellBtn_new
    
    'Vorbereitung
        Dim DgTitle$, GroupName$, m%, w%, s1&
    '2-NamenSpalten-Check
        s1 = Get_StartColumnOfMerge("T4", z, s - 1)
        GroupName = Sheets("T4").Cells(z, s1) 'E1VN1 E2N8
        If Mid(GroupName, 2, 1) <> "2" Then MsgBox "new: nur für Dgs mit 2 NamenSpalten": Exit Sub
    'VBE-CodeFenster merken
        Dim vbeWarSichtbar As Boolean, aktuellesCodeFenster As Object
        On Error Resume Next
        If Application.VBE.MainWindow.Visible = True Then vbeWarSichtbar = True: Set aktuellesCodeFenster = Application.VBE.ActiveCodePane
        On Error GoTo 0
    'UserEingabe (DgTitle; Anzahl m, w) per UserForm anbieten
        UF5.show
    'DgTitle, w, m      'Usereingaben übernehmen (DgTitle, Anzahl w, Anzahl m)
        DgTitle = UF5.TextBoxF1.text
        m = CInt(Val(UF5.TextBoxF2.text)): w = CInt(Val(UF5.TextBoxF3.text))
        Unload UF5 'UserForm zerstören
        Fokus_VBE_zurücksetzen vbeWarSichtbar, aktuellesCodeFenster
    'Abbruch?
        If DgTitle = "-Abbruch-" Then Unload UF5: Exit Sub '"1999-12-31 XY-Wettkampf Ort/CH"

    'Weiter geht's
        Dim DgGroup$, SJE$, sA&, sB&, sM&, sw&, K2() As String, r As Range
        Dim D&, i&, max&, s2&, z1&, z2&, z1E&, z2E&
        With Sheets("T4"): EE 0
    'Ränder Oberstes Dg: z1E, s1, z2E, s2;    Neues Dg: z1, s1, z2, s2
        s1 = Get_StartColumnOfMerge("T4", z, s - 1)
        s2 = s + 1:
        z1E = z + 3
        z2E = Get_NrOfLastRowInColumnNr(s2, "T4")
            If .Cells(z1E, s1) <> "LO" Then Stop
            If .Cells(z1E, s2) <> "RO" Then Stop
            If .Cells(z2E, s2) <> "ub" Then Stop
        z1 = z2E + 3
            max = IIf(w > m, w, m) 'Maximum(Anzahl weibl., Anzahl männl. Aktiver)
        z2 = z1 + 5 + max
    'DgGroup, benötigte Spalten
        DgGroup = .Cells(z1E - 3, s1) 'E2N4
        sw = FindNextColumnRight("T4", "W", 2, s1) 'SpaltenNr Name weibl. Aktive
        sM = FindNextColumnRight("T4", "M", 2, s1) 'SpaltenNr Name männl. Aktive
        sA = FindNextColumnRight("T4", "I", 2, s1) 'SpaltenNr Rang weibl. Aktive
        sB = FindNextColumnRight("T4", "J", 2, s1) 'SpaltenNr Rang männl. Aktive
    'Kopiere Kopf des ersten Dg der aktuellen DgGroup nach unten
        Copy_Range_PasteToCell 7, s1, 13, s2, z1 - 1, s1
    'Clean neuen DgKopf
        'Kopfzeile1     (z1)
            .Cells(z1, s1 + 1) = DgTitle
        'Kopfzeile2     (z1+1) DgGroup, Title
            ReDim K2(1 To s2 - s1 + 1) 'K2() = Kopfzeile2
            K2(1) = DgGroup 'Nur die ersten beiden Kopfzeile2-Zellen mit Text, Rest leer
            If Left(DgGroup, 1) = "E" Then
                K2(2) = DgTitle + " (Einzel)"
                ElseIf Left(DgGroup, 1) = "S" Then K2(2) = DgTitle + " (Synchron)"
                ElseIf Left(DgGroup, 1) = "M" Then K2(2) = DgTitle + " (Mannschaft)"
                Else: Stop
            End If '
            Paste_1DArray_ToRow "T4", CInt(z1) + 1, CInt(s1), K2
        'Kopfzeile4     (z1+3) Schü/Ju/Erwachsene?
            SJE = T4_Get_S_J_or_E_from_DgTitle(DgTitle)
            If SJE = "S" Then
                  .Cells(z1 + 3, sw) = "Schülerinnen":      .Cells(z1 + 3, sM) = "Schüler"
            ElseIf SJE = "J" Then
                  .Cells(z1 + 3, sw) = "Jugendturnerinnen": .Cells(z1 + 3, sM) = "Jugendturner"
            Else: .Cells(z1 + 3, sw) = "Turnerinnen":       .Cells(z1 + 3, sM) = "Turner"
            End If
        'Kopfzeile6     (z1+5) 1. Competitor-Zeile
            .Range(.Cells(z1 + 5, sA + 1), .Cells(z1 + 5, sB - 1)).ClearContents
            .Range(.Cells(z1 + 5, sB + 1), .Cells(z1 + 5, s2)).ClearContents
    'Rumpf hinzu        'Zeilen für Aktive
        For i = 1 To max - 1
            Copy_Range_PasteToCell z1 + 5, s1, z1 + 5, s2, z1 + 5 + i, s1
            .Cells(z1 + 5 + i, sA) = i + 1: .Cells(z1 + 5 + i, sB) = i + 1
        Next
        'Weisse Zellen
        If w > m Then
            D = w - m   'die letzten d hellroten Zeilen  --> weiss
            For i = 1 To D
                Copy_Range_PasteToCell z1 + 4, s1, z1 + 4, sB - 1, z2 - D - 1 + i, s1
            Next
        ElseIf m > w Then
            D = m - w   'die letzten d hellblauen Zeilen --> weiss
            For i = 1 To D
                Copy_Range_PasteToCell z1 + 4, sB, z1 + 4, s2, z2 - D - 1 + i, sB
            Next
        End If
    'SchlußZeile
        Copy_Range_PasteToCell z2E, s1, z2E, s2, z2, s1
    'AussenNrn links
        Set r = .Range(Cells(z1, s1 - 1), Cells(z2, s1 - 1))
        r.Font.size = 6: r.HorizontalAlignment = xlCenter: r.Font.NAME = "Calibri"
        For i = 1 To z2 - z1 + 1
            .Cells(z1 + i - 1, s1 - 1) = i
        Next
    'Scroll zum neuen Dg
        Scroll_ToSeeCell_z_s z1, s1, 5, 2
    'Finals
        EE 1: Beep: End With
End Sub

Function FindNextColumnLeft(NameOfSheet$, Searchstring$, z&, s&) As Long
    Dim ws As Worksheet, suchBereich As Range, gefundeneZelle As Range
    FindNextColumnLeft = 0
    If s <= 1 Then Exit Function
    Set ws = Worksheets(NameOfSheet)
    Set suchBereich = ws.Range(ws.Cells(z, 1), ws.Cells(z, s - 1))
    'Rückwärtsgerichtete, blitzschnelle Suche nach dem exakten Inhalt
    Set gefundeneZelle = suchBereich.Find(What:=Searchstring, After:=suchBereich.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=True)
    If Not gefundeneZelle Is Nothing Then
        FindNextColumnLeft = gefundeneZelle.Column
    End If
End Function

Function FindNextColumnRight(NameOfSheet$, Searchstring$, z&, s&) As Long
    Dim ws As Worksheet, suchBereich As Range, gefundeneZelle As Range
    Dim letzteSpalte As Long
    FindNextColumnRight = 0
    Set ws = Worksheets(NameOfSheet)
    letzteSpalte = ws.Cells(z, ws.Columns.count).End(xlToLeft).Column
    If s >= letzteSpalte Then Exit Function
    Set suchBereich = ws.Range(ws.Cells(z, s + 1), ws.Cells(z, letzteSpalte))
    'Suche starten
    Set gefundeneZelle = suchBereich.Find(What:=Searchstring, After:=suchBereich.Cells(1, suchBereich.Columns.count), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=True)
    If Not gefundeneZelle Is Nothing Then
        FindNextColumnRight = gefundeneZelle.Column
    End If
End Function

Function Get_StartColumnOfMerge(NameOfSheet$, z&, s&) As Long
    Dim zielZelle As Range
    ' Die entsprechende Zelle auf dem aktiven Tabellenblatt referenzieren
    Set zielZelle = ActiveSheet.Cells(z, s)
    ' Prüfen, ob die Zelle Teil eines Zellverbunds ist
    If zielZelle.MergeCells Then
        ' Wenn ja, die Nummer der ersten Spalte des Verbunds zurückgeben
        Get_StartColumnOfMerge = zielZelle.MergeArea.Column
    Else
        ' Wenn nein, 0 zurückgeben
        Get_StartColumnOfMerge = 0
    End If
End Function



