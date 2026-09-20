Attribute VB_Name = "T7"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "ScrollBar1, 2, 0, MSForms, ScrollBar"
Option Explicit 'T7

Sub zzz_T7()
    'showProcs "t7"
    Application.EnableEvents = True
    'RenameModule "ArrHelpers2", "T6"
    'T1_ShowLogBuch
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'Vorbereitung
        Dim s%, s1%, z%, z1%                'z = Zeile, s = Spalte
        z = Target.Row: s = Target.Column
    'Bei jedem Klick auf eine Zelle
        T7_Check_TextBoxChange
        T7_Read_LbToDoFiles
        Beep1
    'CellButtons
        If z = 9 And s = 3 Then T7_Open_ToDoFolder
        If z = 29 And s = 3 Then SelectMyFile_WinEplorer T7_Get_EvPathOfActivePhoto    'Ev
        If z = 31 And s = 3 Then SelectMyFile_WinEplorer T7_Get_LePathOfActivePhoto    'Le
        If z = 33 And s = 3 Then T7_Open_LbArchivFolder
        If z > 12 And z < 27 And s > 1 And s < 9 Then T7_OneFieldLineWasClicked z
End Sub

Public Sub Worksheet_Activate()
    T7_WorksheetActivate
End Sub

Private Sub ScrollBar1_Change()
    'Called from    [User changes ScrollBar-Value]
    T7_Fill_VisibleScrollLines '--> M_T7
End Sub


