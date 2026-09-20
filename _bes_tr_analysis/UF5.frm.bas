Attribute VB_Name = "UF5"
Attribute VB_Base = "0{6016D59C-F483-4D7D-94F7-0D434A22D26A}{D11555E3-1F72-437A-B5E4-E8B62D1D8FD3}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit 'UF5

' Die Konstante wohnt jetzt hier ganz oben in UF5
Private Const VORGABE_TEXT As String = "1999-12-31 XY-Wettkampf Ort/CH"

Private Sub UserForm_Initialize()
    'Einstellungen (Farbe/Größe)
        Call M_UF5.UF5_Einstellungen
    'Vorgabetexte beim Start setzen
        With Me.TextBoxF1: .text = VORGABE_TEXT: .ForeColor = RGB(160, 160, 160): End With
        With Me.TextBoxF2: .text = "10": .ForeColor = RGB(160, 160, 160): End With
        With Me.TextBoxF3: .text = "7": .ForeColor = RGB(160, 160, 160): End With
End Sub

Private Sub cmdErstellen_Click()
    Me.Hide
End Sub

Private Sub cmdAbbrechen_Click()
    TextBoxF1.text = "-Abbruch-": TextBoxF2.text = "-": TextBoxF3.text = "-"
    Me.Hide
End Sub

Private Sub TextBoxF1_Enter()
    'Vorgabetext grau
    If TextBoxF1.text = VORGABE_TEXT And TextBoxF1.ForeColor = RGB(160, 160, 160) Then
        TextBoxF1.text = "": TextBoxF1.ForeColor = vbBlack
    End If
End Sub

Private Sub TextBoxF2_Enter()
    'Vorgabetext grau
    If TextBoxF2.text = "10" And TextBoxF2.ForeColor = RGB(160, 160, 160) Then
        TextBoxF2.text = "": TextBoxF2.ForeColor = vbBlack
    End If
End Sub

Private Sub TextBoxF3_Enter()
    'Vorgabetext grau
    If TextBoxF3.text = "7" And TextBoxF3.ForeColor = RGB(160, 160, 160) Then
        TextBoxF3.text = "": TextBoxF3.ForeColor = vbBlack
    End If
End Sub

Private Sub TextBoxF1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(TextBoxF1.text) = "" Then
        TextBoxF1.text = VORGABE_TEXT: TextBoxF1.ForeColor = RGB(160, 160, 160)
    End If
End Sub

Private Sub TextBoxF2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Zahlen-Zwang
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub TextBoxF3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Zahlen-Zwang
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
