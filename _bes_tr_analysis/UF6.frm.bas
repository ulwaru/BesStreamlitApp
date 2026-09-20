Attribute VB_Name = "UF6"
Attribute VB_Base = "0{262238EF-83B5-48E4-8F73-FA19141E6BF5}{47518C9F-09F3-49D3-99D7-4E9A51926E91}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' Ein internes Feld, um die Auswahl zu speichern
Private m_Auswahl As String

' Diese Property erlaubt es MySub, das Ergebnis auszulesen
Public Property Get Auswahl() As String
    Auswahl = m_Auswahl
End Property

' Klick auf den "männlich" Button
Private Sub Btn_m_Click()
    m_Auswahl = "m"
    Me.Hide ' Formular nur verstecken, nicht entladen!
End Sub

' Klick auf den "weiblich" Button
Private Sub Btn_w_Click()
    m_Auswahl = "w"
    Me.Hide ' Formular nur verstecken, nicht entladen!
End Sub

Private Sub Label1_Click()

End Sub

' Falls der User die Box oben rechts über das "X" schließt
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True ' Schließen abbrechen
        m_Auswahl = "" ' Optional: Abbruch signalisieren
        Me.Hide ' Stattdessen nur verstecken
    End If
End Sub
