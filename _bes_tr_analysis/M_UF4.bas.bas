Attribute VB_Name = "M_UF4"
Option Explicit

' Modulweite Variablen merken sich den Zustand der Schleifen und die Balkenbreite
Private CurrentI As Long
Private CurrentMaxI As Long
Private CurrentL1Text As String
Private MaxBarWidth As Single ' Speicher für die maximale Breite der Fortschrittsbalken

Sub x()
    UF4.show vbModeless
    UpdateCountdown 4, 99, "Hallo", "Change_LinkToPersonFolder_ToVbsLink"

End Sub


' =========================================================================
' 1. DEINE FLEXIBLEN EINSTELLUNGEN (HIER KANNST DU ALLES ANPASSEN)
' =========================================================================
Public Sub UF4_Einstellungen()
    'Called from Start_UF4_Countdown
    
    'Vorbereitung
        Dim UF4_Left!, UF4_Top!, UF4_Width!, UF4_Height!, Text_Height!, Text_Gap!
        Dim L1_Top!, L2_Top!, L1_LeftStart!
    ' --- MAßE & FARBEN FÜR DIE USERFORM (UF4) ---
    UF4_Left = 40: UF4_Top = 40: UF4_Width = 500: UF4_Height = 115
    
    With UF4: .Left = UF4_Left: .Top = UF4_Top: .Width = UF4_Width: .Height = UF4_Height: .BackColor = Green1: End With

    ' --- TEXTZEILEN L1 UND L2 (DIE INTERNEN FELDER L1T1 BIS L1T6) ---
    Text_Height = 14  ' Einheitliche Höhe (entspricht L1F1.Height)
    Text_Gap = 2      ' Abstand zwischen den Feldern (2pt)
    L1_Top = 6: L2_Top = 22: L1_LeftStart = 10
    
    'Die Breiten der 6 Einzelfelder je Zeile
    Dim w(1 To 6) As Single
    w(1) = 38   ' F1: "Folder" / "File"
    w(2) = 15   ' F2: "i=" / "j="
    w(3) = 20   ' F3: Ziffer (aktuell)
    w(4) = 10   ' F4: "of"
    w(5) = 20   ' F5: Ziffer (gesamt)
    w(6) = 310  ' F6: Info-Text (Restbreite)

    ' Dynamischer Aufbau der Textfelder für Zeile 1 und 2
    Dim i As Integer, j As Integer
    Dim currentLeft!
    Dim lbl As MSForms.Label
    
    For i = 1 To 2
        currentLeft = L1_LeftStart
        For j = 1 To 6
            Set lbl = GetOrCreateLabel("L" & i & "T" & j)
            With lbl
                .Top = IIf(i = 1, L1_Top, L2_Top)
                .Left = currentLeft
                If j = 6 Then .Left = currentLeft + 10
                .Width = w(j)
                .Height = Text_Height
                .BackStyle = fmBackStyleTransparent
                .ForeColor = &H0 ' Textfarbe Schwarz
                .Font.NAME = "Segoe UI"
                .Font.size = 9
                
                ' Standardbeschriftungen vergeben
                Select Case j
                    Case 1: .Caption = IIf(i = 1, "Folder", "File"): .TextAlign = fmTextAlignLeft
                    Case 2: .Caption = IIf(i = 1, "i =", "j ="): .TextAlign = fmTextAlignRight
                    Case 3: .Caption = "0": .TextAlign = fmTextAlignRight
                    Case 4: .Caption = "of": .TextAlign = fmTextAlignCenter
                    Case 5: .Caption = "0": .TextAlign = fmTextAlignRight
                    Case 6: .Caption = "": .TextAlign = fmTextAlignLeft
                End Select
            End With
            currentLeft = currentLeft + w(j) + Text_Gap
        Next j
        
        
    Next i

    ' --- FORTSCHRITTSBALKEN (ZEILE 3 UND ZEILE 4) ---
    Dim L3R_Top!, L4R_Top!
    Dim Box_Left!, Box_Width!, Box_Height!
    L3R_Top = 44      ' Top-Position Rahmen oben (Folders)
    L4R_Top = 62      ' Top-Position Rahmen unten (Files)
    Box_Left = 10     ' Linker Startpunkt der Boxen
    Box_Width = UF4.InsideWidth - 20 ' Boxen-Breite passt sich der Form an
    Box_Height = 11   ' Höhe des Rahmens
    
    ' Maximale Balkenbreite hier berechnen und modulweit merken (-2 Pixel für den Innenabstand)
    MaxBarWidth = Box_Width - 2
    
    ' Schriftzüge innerhalb des Rahmens ("Folders" / "Files")
    Dim Title_OffsetLeft!, Title_OffsetTop!
    Dim Title_Width!, Title_Height!
    Title_OffsetLeft = 4  ' Abstand vom linken Rahmenrand
    Title_OffsetTop = 1   ' Leicht nach unten versetzt (2 Pt tiefer)
    Title_Width = 100
    Title_Height = 9
    
    ' Die inneren, wandernden Balken
    Dim Bar_OffsetLeft!, Bar_OffsetTop!
    Dim Bar_Height!
    Bar_OffsetLeft = 1    ' Leicht eingerückt im Rahmen, damit er nicht überlappt
    Bar_OffsetTop = 1
    Bar_Height = 9
    
    Dim names(3 To 4) As String
    names(3) = "Folders"
    names(4) = "Files"
    
    Dim barTop!
    For i = 3 To 4
        barTop = IIf(i = 3, L3R_Top, L4R_Top)
        
        ' REIHENFOLGE GEÄNDERT: ZUERST die Balken im Hintergrund erzeugen...
        ' 3a. Fortschrittsbalken (L3B / L4B - nutzt dein globales Green2)
        Set lbl = GetOrCreateLabel("L" & i & "B")
        With lbl
            .Top = barTop + Bar_OffsetTop
            .Left = Box_Left + Bar_OffsetLeft
            .Width = 0 ' Startet bei Breite Null
            .Height = Bar_Height
            .BackColor = Green2
        End With
        
        ' 3b. Außenrahmen (L3R / L4R) darüberlegen
        Set lbl = GetOrCreateLabel("L" & i & "R")
        With lbl
            .Top = barTop
            .Left = Box_Left
            .Width = Box_Width
            .Height = Box_Height
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = &H0
            .BackStyle = fmBackStyleTransparent
        End With
        
        ' 3c. Innerer Text/Title (Nutzt dein globales Green4) - JETZT GANZ OBENAUF
        Set lbl = GetOrCreateLabel("L" & i & "Title")
        With lbl
            .Top = barTop + Title_OffsetTop
            .Left = Box_Left + Title_OffsetLeft
            .Width = Title_Width
            .Height = Title_Height
            .BackStyle = fmBackStyleTransparent ' Durchsichtig, damit der Balken durchschimmert
            .Caption = names(i)
            .ForeColor = Green4
            .Font.NAME = "Segoe UI"
            .Font.size = 7.5
            .Font.Bold = True
        End With
        
        ' Zur Sicherheit die Z-Ebene nochmals explizit nach vorne zwingen
        UF4.Controls("L" & i & "Title").ZOrder fmZOrderFront
    Next i
End Sub

' Hilfsfunktion für den dynamischen Aufbau (verhindert doppelte Controls)
Private Function GetOrCreateLabel(ByVal ControlName As String) As MSForms.Label
    On Error Resume Next
    Set GetOrCreateLabel = UF4.Controls(ControlName)
    On Error GoTo 0
    
    If GetOrCreateLabel Is Nothing Then
        Set GetOrCreateLabel = UF4.Controls.Add("Forms.Label.1", ControlName)
    End If
End Function

' =========================================================================
' 2. INITIALISIERUNG & ANZEIGE
' =========================================================================
Sub Start_UF4_Countdown() '(ByVal InitialTitle As String)
    Dim InitialTitle$, i%
    InitialTitle = "Hallo"
    'Alle alten Steuerelemente radikal vom Formular fegen
        For i = UF4.Controls.count - 1 To 0 Step -1
            UF4.Controls.Remove i
        Next i
    'Deine frisch konfigurierten Maße und Labels laden
        UF4_Einstellungen
    ' Titel setzen und Form nicht-modal anzeigen
        With UF4
            .Caption = InitialTitle
            .StartUpPosition = 0
            .show vbModeless
        End With
    DoEvents
End Sub

' =========================================================================
' 3. STEUERUNGS-SUB FÜR DEINE SCHLEIFEN (UPDATES WÄHREND DER LAUFZEIT)
' =========================================================================
Sub UpdateCountdown(ByVal CurrentStep As Long, _
                    ByVal MaxStep As Long, _
                    ByVal InfoText As String, _
                    Optional ByVal Title As String = "")
    
    Dim ProzentI As Double
    Dim ProzentJ As Double
    
    ' HOCHSICHERHEITS-CHECK: Falls Steuerelemente noch fehlen, erzeugen wir sie jetzt!
    On Error Resume Next
    Dim testCtrl As Object
    Set testCtrl = UF4.Controls("L1T3")
    On Error GoTo 0
    
    ' Wenn das Test-Label fehlt, initialisieren wir das Layout sofort nach
    If testCtrl Is Nothing Then
        UF4_Einstellungen
    End If
    
    ' Falls MaxBarWidth noch nicht belegt ist, sicheren Standardwert nutzen
    If MaxBarWidth <= 0 Then MaxBarWidth = UF4.InsideWidth - 22
    
    ' Unterscheidung: Kommt der Aufruf aus der i-Schleife (Ordner) oder j-Schleife (Dateien)?
    If Title <> "" Then
        ' --- ÄU?ERE SCHLEIFE (i) ---
        CurrentI = CurrentStep
        CurrentMaxI = MaxStep
        CurrentL1Text = InfoText
        
        With UF4
            .Caption = Title
            .Controls("L1T3").Caption = CurrentI
            .Controls("L1T5").Caption = CurrentMaxI
            .Controls("L1T6").Caption = CurrentL1Text
            
            ' Fortschritt i berechnen und Balken L3B dehnen
            If CurrentMaxI <= 0 Then CurrentMaxI = 1
            ProzentI = CurrentI / CurrentMaxI
            If ProzentI > 1 Then ProzentI = 1
            .Controls("L3B").Width = ProzentI * MaxBarWidth
            
            ' Dateianzeige (j) für den nächsten Ordnerschritt zurücksetzen
            .Controls("L2T3").Caption = "0"
            .Controls("L2T6").Caption = ""
            .Controls("L4B").Width = 0
        End With
    Else
        ' --- INNERE SCHLEIFE (j) ---
        If MaxStep <= 0 Then MaxStep = 1
        ProzentJ = CurrentStep / MaxStep
        If ProzentJ > 1 Then ProzentJ = 1
        
        With UF4
            ' Werte der i-Ebene stabil halten
            .Controls("L1T6").Caption = CurrentL1Text
            .Controls("L1T3").Caption = CurrentI
            
            ' Aktuelle j-Werte eintragen
            .Controls("L2T3").Caption = CurrentStep
            .Controls("L2T5").Caption = MaxStep
            .Controls("L2T6").Caption = InfoText
            
            ' Inneren Balken L4B dehnen
            .Controls("L4B").Width = ProzentJ * MaxBarWidth
        End With
    End If
    
    ' Form neu zeichnen und Windows ganz kurz Luft holen lassen
    UF4.Repaint
    DoEvents
End Sub
