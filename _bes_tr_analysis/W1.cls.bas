Attribute VB_Name = "W1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit 'W1

Sub zzz_W1()
    
    showProcs "file like"

    EE 1: Beep
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    ' 1. Ereignisse abschalten, um Endlosschleifen zu verhindern
    Application.EnableEvents = False
    
    ' --- IHR CODE START ---
    ' Beispiel:
    If Not Intersect(Target, Me.Range("A1:A10")) Is Nothing Then
        ' Ihre Logik hier
    End If
    ' --- IHR CODE ENDE ---

CleanExit:
    ' 2. Ereignisse IMMER wieder aktivieren!
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Fehler: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Sub FotosEinordnen()
    Dim p$, s$
    p = "G:\Archiv Photos\- noch einordnen\Fotos von hb - 202412-202511"
    s = Get_AllFileNamesOfOneFolder(p)
    show s
End Sub

Sub HeicToJpg()
    'Vorbereitung
        Dim pFo$, pIN$, pOUT$, pTool$, qq$, s$, sCmd$, i%, r%, Arr1() As String
        qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq
    'Folder holding some heic files
        pFo = "G:\Archiv Photos\- noch einordnen\Fotos von hb - 202412-202511"
    'Get heic files
        s = Get_AllFileNames_Like_OfOneFolder(pFo, "*.HEIC")
        'show s
        Arr1 = Split(s, vbCrLf)
        For i = 0 To UBound(Arr1)
            pIN = " " + qq + pFo + "\" + Arr1(i) + qq       'heic-file
            pOUT = Replace(pIN, ".HEIC", ".jpg")            'jpg-file
            sCmd = pTool + pIN + " -quality 95" + pOUT
            'show sCmd
            r = ShellAndWait(sCmd, 0, vbHide, PromptUser)
            DoEvents
        Next
End Sub

Sub DngToJpg()
    'Vorbereitung
        Dim pFo$, pIN$, pOUT$, pTool$, qq$, s$, sCmd$, i%, r%, Arr1() As String
        qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq
    'Folder holding some heic files
        pFo = "G:\Archiv Photos\- noch einordnen\- PhotoSync Handy\iPh14 Ulli\- AA iPh ub .dng"
    'Get heic files
        s = Get_AllFileNames_Like_OfOneFolder(pFo, "*.dng")
        'show s
        Arr1 = Split(s, vbCrLf)
        For i = 0 To UBound(Arr1)
            pIN = " " + qq + pFo + "\" + Arr1(i) + qq       'dng-file
            pOUT = " " + qq + pFo + "\" + Replace(Arr1(i), ".dng", ".jpg")            'jpg-file
            sCmd = pTool + pIN + " -quality 95" + pOUT
            'show sCmd
            r = ShellAndWait(sCmd, 0, vbHide, PromptUser)
            DoEvents
        Next
End Sub


