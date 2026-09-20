Attribute VB_Name = "UF2"
Attribute VB_Base = "0{3C3C5663-DF13-44A1-BA55-93636C9611B4}{9FE17CC4-9D71-444D-96C4-FB2681A19494}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'UserForm_MouseDown, UserForm_MouseMove
    Private xPosMouseDownOnUF2 As Single
    Private yPosMouseDownOnUF2 As Single

Sub ooo__UF2()
    
End Sub

Private Sub UserForm_Initialize()
    #If IsMac = False Then
        'hide the title bar if you're working on a windows machine.
        'Otherwise, just display it as you normally would
        Me.Height = Me.Height - 10
        M_UF2.HideTitleBar Me
    #End If
End Sub

Private Sub UserForm_Activate()
    'UF2_Einstellungen
End Sub

'UF2 kann während der Anzeige verschoben werden,
'   falls zu dem angeklickten Objekt (Uf2/ein Control darauf)
'   ein MouseDown- und ein MouseMove-Macro existiert
    'MouseDown UF2
    Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): xPosMouseDownOnUF2 = x: yPosMouseDownOnUF2 = y: End Sub
    Private Sub Frame2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): xPosMouseDownOnUF2 = x: yPosMouseDownOnUF2 = y: End Sub
    Private Sub Label1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): xPosMouseDownOnUF2 = x: yPosMouseDownOnUF2 = y: End Sub
    Private Sub Frame1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): xPosMouseDownOnUF2 = x: yPosMouseDownOnUF2 = y: End Sub
    Private Sub Label2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): xPosMouseDownOnUF2 = x: yPosMouseDownOnUF2 = y: End Sub
    'MouseMove UF2
    Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        If Button And 1 Then Me.Left = Me.Left + (x - xPosMouseDownOnUF2): Me.Top = Me.Top + (y - yPosMouseDownOnUF2)
    End Sub
    Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        If Button And 1 Then
            Me.Left = Me.Left + (x - xPosMouseDownOnUF2)
            Me.Top = Me.Top + (y - yPosMouseDownOnUF2)
            ArrC(69) = CStr(Me.Top)
        End If
    End Sub
    Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        If Button And 1 Then Me.Left = Me.Left + (x - xPosMouseDownOnUF2): Me.Top = Me.Top + (y - yPosMouseDownOnUF2)
    End Sub
    Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        If Button And 1 Then Me.Left = Me.Left + (x - xPosMouseDownOnUF2): Me.Top = Me.Top + (y - yPosMouseDownOnUF2)
    End Sub
    Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
        If Button And 1 Then Me.Left = Me.Left + (x - xPosMouseDownOnUF2): Me.Top = Me.Top + (y - yPosMouseDownOnUF2)
    End Sub


