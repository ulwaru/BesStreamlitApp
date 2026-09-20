Attribute VB_Name = "M_UF3"
Option Explicit 'M_UF3

Sub zzz_M_UF3()

    'showProcs "rgb"
    
    EE 1: Beep
    'show getRGB(Sheets("T5").Cells(2, 2))
End Sub

Sub UF3_CountDown(CountOfLines%, Title$)
    'Called from    T4_ManageBracs_EventsAndClubsNations_NoCopResults
    
    'Vorbereitung
        With UF3: UF3_Format
        UF3_ShowLines CountOfLines
        FillArrC 88, CStr(CountOfLines)
    'UF3 anzeigen
        .show vbModeless: .Caption = Title: .Left = 40: .Top = 40: .Width = 500: .BackColor = Green1
    'Finals
        End With
End Sub

Sub UF3_ShowLines(CountOfLines%)
    If CountOfLines < CInt(ArrC(88)) + 1 Then Exit Sub Else FillArrC 88, CStr(CountOfLines)
    With UF3
        .Height = 43 + 20 * CountOfLines
        If CountOfLines > 0 Then .A1.Visible = True: .B1.Visible = True: .C1.Visible = True
        If CountOfLines > 1 Then .A2.Visible = True: .B2.Visible = True: .C2.Visible = True
        If CountOfLines > 2 Then .a3.Visible = True: .B3.Visible = True: .C3.Visible = True
        If CountOfLines > 3 Then .a4.Visible = True: .B4.Visible = True: .C4.Visible = True
        If CountOfLines > 4 Then .a5.Visible = True: .B5.Visible = True: .C5.Visible = True
        If CountOfLines > 5 Then .a6.Visible = True: .B6.Visible = True: .C6.Visible = True
    End With
End Sub

Sub UF3_Format()
    'Called from    UF3_CountDown
    'Action         bereitet UF3_CountDown vor; 6 Zeilen, je Nr|Text|CountDown|erledigt;
    '               Positionen, Größen, Farben; zunächst nur erste Zeile sichtbar
    
    'Vorbereitung
        With UF3: .Frame1.Visible = False
    'A1 Laufende Nr.
        With .A1: .Left = 6: .Top = 10: .Width = 10: .Height = 12
        .BackColor = Green1: .Caption = "1": .Font.size = 8
        .ForeColor = vbBlack: .Visible = False: End With
    'B1 Text; Beschreibung
        With .B1: .Left = 20: .Top = 10: .Width = 430: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "B1Text": .Font.size = 8: End With
    'C1 CountDown
        With .C1: .Left = 460: .Top = 10: .Width = 20: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = 0: .Font.size = 10: .Font.Bold = True: End With
    'D1 Erledigt-Häkchen
        With .D1: .Left = 460: .Top = 10: .Width = 20: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "ü": .Font.size = 13: .Font.Bold = True: End With
        
    'A2 Laufende Nr.
        With .A2: .Left = 6: .Top = 30: .Width = 10: .Height = 12: .Visible = False
        .BackColor = Green1: .Caption = "2": .Font.size = 8: .ForeColor = vbBlack: End With
    'B2 Text; Beschreibung
        With .B2: .Left = 20: .Top = 30: .Width = 430: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "B2Text": .Font.size = 8: End With
    'C2 CountDown
        With .C2: .Left = 460: .Top = 30: .Width = 20: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = 0: .Font.size = 10: .Font.Bold = True: End With
    'D2 Erledigt-Häkchen
        With .D2: .Left = 460: .Top = 30: .Width = 20: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "ü": .Font.size = 13: .Font.Bold = True: End With
        
    'A3 Laufende Nr.
        With .a3: .Left = 6: .Top = 50: .Width = 10: .Height = 12: .Visible = False
        .BackColor = Green1: .Caption = "3": .Font.size = 8: .ForeColor = vbBlack: End With
    'B3 Text; Beschreibung
        With .B3: .Left = 20: .Top = 50: .Width = 430: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "B3Text": .Font.size = 8: End With
    'C3 CountDown
        With .C3: .Left = 460: .Top = 50: .Width = 20: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "": .Font.size = 10: .Font.Bold = True: End With
    'D3 Erledigt-Häkchen
        With .D3: .Left = 460: .Top = 50: .Width = 20: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "ü": .Font.size = 13: .Font.Bold = True: End With
        
    'A4 Laufende Nr.
        With .a4: .Left = 6: .Top = 70: .Width = 10: .Height = 12: .Visible = False
        .BackColor = Green1: .Caption = "4": .Font.size = 8: .ForeColor = vbBlack: End With
    'B4 Text; Beschreibung
        With .B4: .Left = 20: .Top = 70: .Width = 430: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "B4Text": .Font.size = 8: End With
    'C4 CountDown
        With .C4: .Left = 460: .Top = 70: .Width = 20: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "": .Font.size = 10: .Font.Bold = True: End With
    'D4 Erledigt-Häkchen
        With .D4: .Left = 460: .Top = 70: .Width = 20: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "ü": .Font.size = 13: .Font.Bold = True: End With
        
    'A5 Laufende Nr.
        With .a5: .Left = 6: .Top = 90: .Width = 10: .Height = 12: .Visible = False
        .BackColor = Green1: .Caption = "5": .Font.size = 8: .ForeColor = vbBlack: End With
    'B5 Text; Beschreibung
        With .B5: .Left = 20: .Top = 90: .Width = 430: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "B5Text": .Font.size = 8: End With
    'C5 CountDown
        With .C5: .Left = 460: .Top = 90: .Width = 20: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "": .Font.size = 10: .Font.Bold = True: End With
    'D5 Erledigt-Häkchen
        With .D5: .Left = 460: .Top = 90: .Width = 20: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "ü": .Font.size = 13: .Font.Bold = True: End With
        
    'A6 Laufende Nr.
        With .a6: .Left = 6: .Top = 110: .Width = 10: .Height = 12: .Visible = False
        .BackColor = Green1: .Caption = "6": .Font.size = 8: .ForeColor = vbBlack: End With
    'B6 Text; Beschreibung
        With .B6: .Left = 20: .Top = 110: .Width = 430: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "B6Text": .Font.size = 8: End With
    'C6 CountDown
        With .C6: .Left = 460: .Top = 110: .Width = 20: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "": .Font.size = 10: .Font.Bold = True: End With
    'D6 Erledigt-Häkchen
        With .D6: .Left = 460: .Top = 110: .Width = 20: .Height = 12: .Visible = False
        .BackColor = Green2: .Caption = "ü": .Font.size = 13: .Font.Bold = True: End With
    'Finals
        End With
End Sub

Sub UF3CountDown(LineNr%, max&, i%, Start%)
    Dim m%
    With UF3
    m = 10 * (Start + 1) * (max - i) \ max
    If m Mod 10 = 0 Then
        If LineNr = 1 Then
                               .C1 = m \ 10: DoEvents
        ElseIf LineNr = 2 Then .C2 = m \ 10: DoEvents
        ElseIf LineNr = 3 Then .C3 = m \ 10: DoEvents
        ElseIf LineNr = 4 Then .C4 = m \ 10: DoEvents
        ElseIf LineNr = 5 Then .C5 = m \ 10: DoEvents
        ElseIf LineNr = 6 Then .C6 = m \ 10: DoEvents
        End If
    End If
    End With
End Sub

Sub UF3_Show_Text_ForXSeconds_TEST()
    Dim qq$, T$, v$
    qq = Chr(34): v = vbCrLf
    T = "Ein Dg wurde hinzugefügt:" + v + qq _
        + "1988-10-27 JEM19 Salzgitter/D (Einzel)" + qq
    UF3_Show_Text_ForXSeconds T, 3
End Sub

Sub UF3_Show_Text_ForXSeconds(text$, XSeconds%)
    'Called from    xxx
    'L              = Anzahl Zeilen in Text
    
    'Vorbereitung
        Dim v$, L%
        v = vbCrLf: L = anzAinB(v, text) + 1
    'UF3 selbst
        With UF3: .show vbModeless: .Left = 40: .Top = 40: .Width = 500
        .BackColor = Green1: .Height = 70 + L * 25: End With
    'Label1
        With UF3.A1: .Left = 20: .Top = 10: .Width = 480: .Height = UF3.Height - 30
        .BackColor = Green1: .Caption = text: .Font.size = 20: .Visible = True
        .Font.NAME = "Arial": .Font.Bold = True: .ForeColor = vbBlack: End With
    'Ausblenden
        UF3.B1.Visible = False: UF3.C1.Visible = False: UF3.D1.Visible = False
    'Frame1
        With UF3.Frame1: .Left = 10: .Top = UF3.Height - 50
        .Height = 10: .Caption = "": .BackColor = RGB(90, 110, 250): .ZOrder 1 'Hintergrund
        .SpecialEffect = fmSpecialEffectFlat: .Visible = True: End With
    'Timer
        UF3_ShowForXsecs XSeconds
        UF3.Hide 'Absturz bei "Unload UF3"
End Sub

Sub UF3_ShowForXsecs(X_seconds%)
    'Called from    UF3_Show_Text_ForXSeconds
    
    Dim w!, startTime&, mSecDone&, mMax&, wMax&
    startTime = GetTickCount()
    wMax = 465: mMax = X_seconds * 1000 'Max of FrameWidth, MilliSeconds
    With UF3.Frame1
        Do While mSecDone < mMax
            mSecDone = (GetTickCount() - startTime) 'MilliSeconds gone since startTime
            DoEvents
            w = wMax * mSecDone / mMax
            .Width = w 'actual FrameWidth
        Loop
    End With
End Sub


