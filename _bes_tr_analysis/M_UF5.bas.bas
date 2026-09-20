Attribute VB_Name = "M_UF5"
Option Explicit

Sub UF5_Einstellungen()
    'Called from    xxx
    'Label          werden dynamisch erzeugt
    
    'Vorbereitung
        Dim lblF1 As Object, lblF2 As Object, lblF3 As Object
    'UserForm-Grunddesign
        With UF5: .Caption = "Erstelle eine neue Ergebnisliste"
            .Width = 353: .Height = 155: .BackColor = Green1: End With
    'Titel (Label)
        Set lblF1 = UF5.Controls.Add("Forms.Label.1", "LabelF1"): With lblF1
        .Caption = "Wettkampf-Titel:": .BackStyle = fmBackStyleTransparent
        .Width = 300: .Left = 20: .Top = 7: .ForeColor = vbBlack: End With
    'Titel (EingabeBox)
        With UF5.TextBoxF1: .text = "1999-12-31 XY-Wettkampf Ort/CH"
        .Width = 300: .Left = 20: .Top = 20: .ForeColor = RGB(160, 160, 160): End With
    'w (Label)
        Set lblF2 = UF5.Controls.Add("Forms.Label.1", "LabelF2"): With lblF2
        .Caption = "Anzahl Aktive weibl.:": .BackStyle = fmBackStyleTransparent
        .Width = 80: .Left = 20: .Top = 50: .ForeColor = vbBlack: End With
    'w (EingabeBox)
        With UF5.TextBoxF2: .Width = 40: .Left = 100: .Top = 47: End With
    'm (Label)
        Set lblF3 = UF5.Controls.Add("Forms.Label.1", "LabelF3"): With lblF3
        .Caption = "Anzahl Aktive männl.:": .BackStyle = fmBackStyleTransparent
        .Width = 80: .Left = 196: .Top = 50: .ForeColor = vbBlack: End With
    'm (EingabeBox)
        With UF5.TextBoxF3: .Width = 40: .Left = 280: .Top = 47: End With
    'Horizontale Linie
        Call CreateHorizontalLine(20, 77, 300)
    'Buttons
        With UF5.cmdErstellen: .Caption = "Erstellen": .Width = 70: .Left = 80:  .Top = 90: .Default = True: End With
        With UF5.cmdAbbrechen: .Caption = "Abbrechen": .Width = 70: .Left = 200: .Top = 90: .Cancel = True: End With
End Sub

Sub UF5_Verarbeiten()
    ' Werte aus den Textboxen in deine Variablen sichern
    If UF5.TextBoxF1.text = "1999-12-31 XY-Wettkampf Ort/CH" And UF5.TextBoxF1.ForeColor = RGB(160, 160, 160) Then
        strF1 = ""
    Else
        strF1 = UF5.TextBoxF1.text
    End If
    
    intF2 = CInt(Val(UF5.TextBoxF2.text))
    intF3 = CInt(Val(UF5.TextBoxF3.text))
    
    ' Ab hier stehen strF1, intF2 und intF3 perfekt befüllt im Modul bereit!
End Sub

Private Sub CreateHorizontalLine(Left!, topPosition!, lineLength!)
    Dim dynamicLine As MSForms.Label
    
    ' 1. Label dynamisch zur UserForm hinzufügen
    Set dynamicLine = UF5.Controls.Add("Forms.Label.1", "DynamicLine", True)
    
    ' 2. Eigenschaften für die waagerechte Linie setzen
    With dynamicLine
        .Height = 1                          ' Steuert die Dicke (1–2 Pixel ist ideal)
        .Width = lineLength                  ' Die Länge der Linie
        .Left = Left                           ' X-Position
        .Top = topPosition                   ' Y-Position (wird per Parameter übergeben)
        
        ' Design-Bereinigung, damit es wie eine echte Linie aussieht
        .BorderStyle = fmBorderStyleNone     ' Kein Rahmen
        .SpecialEffect = fmSpecialEffectFlat ' Keine 3D-Effekte
        .BackColor = RGB(128, 128, 128)      ' Linienfarbe (Grau)
        .Caption = ""                        ' Kein Text
    End With
End Sub

