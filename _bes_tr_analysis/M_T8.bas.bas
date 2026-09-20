Attribute VB_Name = "M_T8"
Option Explicit 'M_T8

Sub zzz_M_T8()
    showProcs "zdg"
    Application.EnableEvents = True
    'RenameModule "Modul1", "PD"
    'FillArrC 59, ""
    'showArray2D ArrPd
    Beep
End Sub

Sub T4_ActivateEachPersNameOfEachDg()
    'Called from    xxx
    'sDgNameW       = CInt(ArrC(90))    'T4 Dg SpaltenNr Namen  weibl.
    'sDgNameM       = CInt(ArrC(91))    'T4 Dg SpaltenNr Namen  männl.
    'sDgNameX       = CInt(ArrC(92))    'T4 Dg SpaltenNr Namen  mix

    'Vorbereitung
        Dim i%, j%, k%, s1%, z1%, z2%, A() As Integer, ArrDgTitles() As String
        'Application.ScreenUpdating = False
        ReDim A(1 To 3): With Sheets("T4")
        'CountDownLabel_Create
    'Load
        T4_Load_ArrDgTitles_FromT4ListOfDesignTitles ArrDgTitles
        'showArray ArrDgTitles:Stop
        
    For i = 119 To UBound(ArrDgTitles)
        'Schleife über alle Dgs
        'CountDownLabel_Update UBound(ArrDgTitles) - i
        ActivateDesign_GiveTitle ArrDgTitles(i)
        A(1) = CInt(ArrC(90)): A(2) = CInt(ArrC(91)): A(3) = CInt(ArrC(92))   '0|0|3  '3|7|0
        z1 = CInt(ArrC(37)): z2 = CInt(ArrC(38)): s1 = CInt(ArrC(39))
        For j = 1 To 3
            'Schleife über alle Dg-NamenSpalten
            If A(j) > 0 Then
                For k = 6 To z2 - z1 + 1
                    'Schleife über alle Dg-NamenZeilen
                    If ArrDg(k, A(j)) <> "" Then .Cells(z1 - 1 + k, s1 - 1 + A(j)).Select
                    DoEvents
                    'CountDownLabel_Update UBound(ArrDgTitles) - i
                Next
            End If
        Next
    Next
    'Finals
        Application.ScreenUpdating = True
        End With: Beep
End Sub

Function GetArrPdZeilenNrOfOneName(Nn$, Vn$) As Integer
    Dim i%: DoArrPd
    For i = 1 To UBound(ArrPd, 1)
        If ArrPd(i, 1) = Nn Then
            If ArrPd(i, 2) = Vn Then GetArrPdZeilenNrOfOneName = i: Exit Function
        End If
    Next
End Function

Sub T4_ActivateEachDesign()
    Dim i%, ArrDgTitles() As String
    Application.ScreenUpdating = False
    T4_Load_ArrDgTitles_FromT4ListOfDesignTitles ArrDgTitles
        'showArray ArrDgTitles
    For i = 1 To UBound(ArrDgTitles)
        ActivateDesign_GiveTitle ArrDgTitles(i)
    Next
    Application.ScreenUpdating = True
    Beep
End Sub

Sub ActivateEachLigaDesign()
    Dim i%, ArrDgTitles() As String
    Application.ScreenUpdating = False
    T4_Load_ArrDgTitles_FromT4ListOfDesignTitles ArrDgTitles
        'showArray ArrDgTitles
    For i = 1 To UBound(ArrDgTitles)
        If ArrDgTitles(i) Like "*Buli*" Then
            ActivateDesign_GiveTitle ArrDgTitles(i)
        End If
    Next
    Application.ScreenUpdating = True
    Beep
End Sub



