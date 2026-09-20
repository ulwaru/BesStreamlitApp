Attribute VB_Name = "T1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'Module -- Typ: Sheet Module -- Name: "T1"

Option Explicit

Sub zzz_T1()
    Application.EnableEvents = True
    showProcs "text"
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Vorbereitung
        Dim zs$, zs1$, zs2$, s%, z%, z9%, r As Range
        z = Target.Row: s = Target.Column
        zs1 = CStr(z) + "|" + CStr(s): zs2 = CStr(z) + "|" + CStr(s + 1)
        With Sheets("T1"): Set r = .Cells(z, s): End With
    'Vorab
        Call ZeilenHoehe15:        Beep1
        Mark_T1CellButtons_OpenShow
    'Click on W1-XBox "Menü ein"
        zs = Get_T1Cell_HoldingMyText(" Menü ein")
        If zs = zs1 Then XBox_MenüEin z, s - 1:      Exit Sub 'Klick auf Text
        If zs = zs2 Then XBox_MenüEin z, s:          Exit Sub 'Klick auf XBox
    'Click on W1-XBox "reduzierte Höhe"
        zs = Get_T1Cell_HoldingMyText(" reduzierte Höhe")
        If zs = zs1 Then XBox_MenüKlein z, s - 1:    Exit Sub 'Klick auf Text
        If zs = zs2 Then XBox_MenüKlein z, s:        Exit Sub 'Klick auf XBox
    'Click on W1-XBox "Gitterlinien"
        zs = Get_T1Cell_HoldingMyText(" Gitterlinien")
        If zs = zs1 Then XBox_Gitterlinien z, s - 1: Exit Sub 'Klick auf Text
        If zs = zs2 Then XBox_Gitterlinien z, s:     Exit Sub 'Klick auf XBox
    'Click on W1-XBox "Code rechts"
        zs = Get_T1Cell_HoldingMyText(" Code rechts")
        If zs = zs1 Then XBox_CodeRechts z, s - 1:   Exit Sub 'Klick auf Text
        If zs = zs2 Then XBox_CodeRechts z, s:       Exit Sub 'Klick auf XBox
    'Click on CellButton
        z9 = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "'Events'")
            If z = z9 And s = 19 Then T1_OpenFolder_Events z, s: Exit Sub
        z9 = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "Clubs")
            If z = z9 And s = 19 Then T1_OpenFolder_Clubs z, s: Exit Sub
        z9 = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "'Leute'")
            If z = z9 And s = 19 Then T1_OpenFolder_Leute z, s: Exit Sub
        z9 = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "'Register'")
            If z = z9 And s = 19 Then T1_OpenFolder_Register z, s: Exit Sub
        z9 = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "'ToDo'")
            If z = z9 And s = 19 Then T1_OpenFolder_ToDo z, s: Exit Sub
        z9 = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "Read me")
            If z = z9 And s = 19 Then T1_ShowReadMeTxt z, s: Exit Sub
        z9 = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "Logbuch")
            If z = z9 And s = 19 Then T1_ShowLogBuch: Exit Sub
        z9 = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "Terminkalender")
            If z = z9 And s = 19 Then T1_ShowTerminkalenderTxt z, s: Exit Sub
        z9 = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "Competitors")
            If z = z9 And s = 19 Then T1_ShowCompetitorsTxt: Exit Sub
        z9 = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "PersonData")
            If z = z9 And s = 19 Then T1_ShowPersonDataTxt: Exit Sub
    'Save a copy
        z9 = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "Save a copy")
            If z = z9 And s = 19 Then T1_Save_a_copy z, s: Exit Sub
    'Update Statistics
        z9 = Get_RowNr_HoldingMyTextPartInColumnX("T1", 35, 2, "Statistics")
            If z = z9 And s = 35 Then T1_UpdateStatistics: Exit Sub
    'Click on Green1
        If getColor(r) = 14348258 Then ClickOnGreen1 z, s
    'Write z, s to T1
        EE 0: [B1] = s: [A2] = z: EE 1
End Sub



