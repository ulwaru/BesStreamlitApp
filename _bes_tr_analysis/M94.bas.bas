Attribute VB_Name = "M94"
Option Explicit 'M

Sub zzz_M()
    
    showProcs "load"
    show "aa"
    
    EE 1: Beep
End Sub

Function Get_TextOfAllOpenNotepadWindows() As String
    Dim objWMIService As Object
    Dim colProcesses As Object
    Dim objProcess As Object
    Dim fso As Object
    Dim txtStream As Object
    Dim commandLine As String
    Dim filePath As String
    Dim resultText As String
    Dim delimiter As String
    Dim startPos As Long
    
    delimiter = vbCrLf & String(40, "-") & vbCrLf
    resultText = ""
    
    On Error Resume Next
    ' WMI-Dienst und Dateisystem-Objekt initialisieren
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Alle aktuell geöffneten Notepad-Prozesse abfragen
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'notepad.exe'")
    
    If colProcesses.count = 0 Then
        Get_TextOfAllOpenNotepadWindows = "Fehler: Es sind aktuell keine Notepad-Fenster geöffnet."
        Exit Function
    End If
    
    ' Jeden einzelnen Notepad-Prozess durchlaufen
    For Each objProcess In colProcesses
        ' CommandLine enthält den vollen Aufruf, z.B.: notepad.exe "C:\Ordner\1234.txt"
        commandLine = objProcess.commandLine
        
        If Trim(commandLine) <> "" Then
            ' Wir suchen nach dem Dateipfad im Startbefehl
            startPos = InStr(commandLine, " ")
            If startPos > 0 Then
                ' Pfad extrahieren und eventuelle Anführungszeichen entfernen
                filePath = Mid(commandLine, startPos + 1)
                filePath = Replace(filePath, """", "")
                filePath = Trim(filePath)
                
                ' Prüfen, ob die extrahierte Datei wirklich existiert
                If fso.FileExists(filePath) Then
                    ' Datei im Hintergrund öffnen (1 = ForReading, -2 = System-Standard-Codierung)
                    Set txtStream = fso.OpenTextFile(filePath, 1, False, -2)
                    
                    If Not txtStream Is Nothing Then
                        ' Inhalt auslesen und mit dem Dateinamen als Überschrift anheften
                        resultText = resultText & "DATEI: " & fso.GetFileName(filePath) & vbCrLf & _
                                     txtStream.ReadAll & delimiter
                        txtStream.Close
                    End If
                    Set txtStream = Nothing
                End If
            End If
        End If
    Next objProcess
    On Error GoTo 0
    
    ' Ergebnisauswertung
    If resultText = "" Then
        Get_TextOfAllOpenNotepadWindows = "Hinweis: Notepad-Prozesse gefunden, aber die Dateipfade konnten nicht ausgelesen werden."
    Else
        Get_TextOfAllOpenNotepadWindows = resultText
    End If
    
    ' Objekte sauber abbauen
    Set colProcesses = Nothing
    Set objWMIService = Nothing
    Set fso = Nothing
End Function

Sub Join2d_TEST()
    Dim s$, T3()
    T3_Load_DataArea T3
    s = Join2d(T3, vbCrLf, "|")
    show s
End Sub

Public Function Join2d(ByRef InputArray As Variant, Optional RowDelimiter As String = vbCr, _
    Optional FieldDelimiter = vbTab, Optional SkipBlankRows As Boolean = False) As String

    ' Join up a 2-dimensional array into a string. Works like the standard
    '  VBA.Strings.Join, for a 2-dimensional array.
    ' Note that the default delimiters are those inserted into the string
    '  returned by ADODB.Recordset.GetString
    On Error Resume Next
    ' Coding note: we're not doing any string-handling in VBA.Strings -
    ' allocating, deallocating and (especially!) concatenating are SLOW.
    ' We're using the VBA Join & Split functions ONLY. The VBA Join,
    ' Split, & Replace functions are linked directly to fast (by VBA
    ' standards) functions in the native Windows code. Feel free to
    ' optimise further by declaring and using the Kernel string functions
    ' if you want to.
    ' ** THIS CODE IS IN THE PUBLIC DOMAIN **
    '   Nigel Heffernan   Excellerando.Blogspot.com
    
    Dim i&, j&, i_lBound&, i_uBound&, j_lBound&, j_uBound&
    Dim strBlankRow$, arrTemp1() As String, arrTemp2() As String
    
    i_lBound = LBound(InputArray, 1): i_uBound = UBound(InputArray, 1)
    j_lBound = LBound(InputArray, 2): j_uBound = UBound(InputArray, 2)
    ReDim arrTemp1(i_lBound To i_uBound): ReDim arrTemp2(j_lBound To j_uBound)
    
    For i = i_lBound To i_uBound
        For j = j_lBound To j_uBound
            arrTemp2(j) = InputArray(i, j)
        Next j
        arrTemp1(i) = Join(arrTemp2, FieldDelimiter)
    Next i
    
    If SkipBlankRows Then
        If Len(FieldDelimiter) = 1 Then
            strBlankRow = String(j_uBound - j_lBound, FieldDelimiter)
        Else
            For j = j_lBound To j_uBound
                strBlankRow = strBlankRow & FieldDelimiter
            Next j
        End If
        Join2d = Replace(Join(arrTemp1, RowDelimiter), strBlankRow, RowDelimiter, "")
        i = Len(strBlankRow & RowDelimiter)
        If Left(Join2d, i) = strBlankRow & RowDelimiter Then Mid$(Join2d, 1, i) = ""
    Else
        Join2d = Join(arrTemp1, RowDelimiter)
    End If
    Erase arrTemp1
End Function



