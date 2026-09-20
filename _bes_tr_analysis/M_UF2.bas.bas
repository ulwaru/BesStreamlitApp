Attribute VB_Name = "M_UF2"
Option Explicit 'M_UF2

'HideTitleBar
    Public Const GWL_STYLE = -16
    Public Const WS_CAPTION = &HC00000
    #If VBA7 Then
        Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
        Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
        Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
        Public Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    #Else
        Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
        Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
        Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
        Public Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    #End If

Sub zzz_M_UF2()
    
    Unload UF2
End Sub

Sub UF2_Test2()
    Unload UF2
    'UF2 selbst
        'HideTitleBar UF2
        UF2.Width = 500: UF2.Height = 55: UF2.Left = 55: UF2.Top = 40
    'Label1
        With UF2.Label1: .Left = 10: .Top = 8: .Width = 480: .Height = 70
            .BackColor = RGB(0, 0, 180)
            .Caption = "Irgendein Text": .Font.size = 20: .Font.Bold = True
            .Font.NAME = "Arial": .ForeColor = RGB(255, 0, 0)
        End With
    'Label2
        UF2.Label2.Visible = False
    'Frame1
        UF2.Frame1.Visible = False
    'Frame2
        UF2.Frame2.Visible = False
    'Show
        UF2.show 'vbModeless
End Sub

Sub UF2_Einstellungen()
    
    Exit Sub
    
    DoArrc
    With UF2:          .Width = 240: .Height = 55:
            '.Left = CSng(ArrC(69)): .Top = CSng(ArrC(70)) ':
        .Left = 100: .Top = 0
    With .Label1:      .Left = 14:   .Top = 8: .BackColor = RGB(240, 240, 240): End With
    With .Frame1:    .Top = 20:    .BackColor = RGB(240, 240, 240):           End With
    With .Label2: .Left = 221: .Top = 20:   .BackColor = RGB(240, 240, 240)
                       .Caption = "9": .Font.size = 16: .Font = "Arial": .Font.Bold = True
    End With: End With
End Sub

Sub ProgressBar_TEST()
    Dim CountDownNr$, text$, i&, iMax&, ProzentDone!
    iMax = 10 '000 'iMax
    UF2.Frame2.Width = 0: UF2.show 'vbModeless
    For i = 1 To iMax
        text = "Processing Row " & i & " of " & iMax
        ShowProgressBar iMax, CLng(i), text

            '--------------------------------------
            'the VnNn of your macro goes below here
            '
                Application.Wait (Now + TimeValue("0:00:01"))
            '
            '--------------------------------------
            
'            If i = iMax Then
'                Unload UF2
'            End If
    Next i
End Sub

Sub ShowProgressBar(FsMax&, i&, text$)
    Dim CountDownNr$, ProzentDone!
    
    If Not UserformIsLoaded("UF2") Then UF2.Frame2.Width = 0: UF2.show
    'Periodically update progress bar
        ProzentDone = i / FsMax
        With UF2
            .Label1.Caption = "Processing Row " & i & " of " & FsMax
            .Frame2.Width = ProzentDone * (.Frame1.Width)
            CountDownNr = CStr(10 - CInt(ProzentDone * 10)): If CountDownNr = "10" Then CountDownNr = "9"
            .Label2.Caption = CountDownNr
        End With
        'UF2.Repaint
        DoEvents
    If i = FsMax Then
        Unload UF2
    End If
End Sub

Public Function UserformIsLoaded(formName As String) As Boolean
    'Called from    ShowProgressBar
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.NAME = formName Then
            UserformIsLoaded = True
            Exit Function
        End If
    Next frm
    UserformIsLoaded = False
End Function

Sub HideTitleBar(frm As Object)
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = FindWindowA(vbNullString, frm.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
End Sub


