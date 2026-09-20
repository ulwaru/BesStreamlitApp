Attribute VB_Name = "M23"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Private Const SW_RESTORE As Long = 9

Sub OpenFolder_BigIcons_SelectFile_TEST()
    OpenFolder_BigIcons_SelectFile "F:\Archiv Trampolin 1900-1999\ClubsNations\GB\Curtis, David (CEFN Forest, GB).lnk"
End Sub

Sub OpenFolder_BigIcons_SelectFile(PathOfFile As String)
    'Vorbereitung
        Dim w&, H&, L&, T&, Retries&, hwnd As LongPtr
    'Position, Größe
        L = 35: T = 295: w = 850: H = 700
    'Datei im Explorer öffnen & markieren
        shell "explorer.exe /select," & Chr(34) & PathOfFile & Chr(34), vbNormalFocus
    'Gezielt nach dem Explorer-Fenster ("CabinetWClass") suchen (alle 50ms, max 1 Sekunde)
        Do: Sleep 50: hwnd = FindWindow("CabinetWClass", vbNullString)
            Retries = Retries + 1: DoEvents
        Loop While hwnd = 0 And Retries < 20
    'Wenn gefunden: In den Vordergrund holen, Positionieren und Ansicht ändern
        If hwnd <> 0 Then
            ShowWindow hwnd, SW_RESTORE: SetForegroundWindow hwnd
            MoveWindow hwnd, L, T, w, H, 1
            'Kurz warten, bis das Fenster den Fokus verarbeitet hat, dann Tastenkürzel für "Große Symbole" senden
            Sleep 100: SendKeys "^+2", True
        End If
        MoveWindow hwnd, L, T, w, H, 1
End Sub


