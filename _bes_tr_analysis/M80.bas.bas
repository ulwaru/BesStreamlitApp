Attribute VB_Name = "M80"
Option Explicit 'M80

Sub Repair_TEST()
    Repair "M43"
End Sub


' ==============================================================================
' 1. DETAIL-VERGLEICH (Zählt exakt Textzeilen im Header)
' ==============================================================================
Sub CompareCode()
    Dim EXCLUDE_MODULE As String
    EXCLUDE_MODULE = "M80" ' Das eigene Vergleichs-Modul vor Löschung schützen
    
    Dim Wb1 As Workbook, Wb2 As Workbook
    Dim comp1 As Object, comp2 As Object
    Dim name1 As String, name2 As String
    Dim sBody As String, sHeader As String
    
    Dim sortedNames() As String
    Dim modName As Variant
    Dim procResult As String
    Dim i As Long
    
    Dim moduleCount As Long
    Dim textLineCount As Long
    Dim lines() As String
    Dim lineItem As Variant
    
    Set Wb1 = ThisWorkbook
    name1 = Wb1.NAME
    name2 = "A2 corrupt.xlsm"
    
    On Error Resume Next
    Set Wb2 = Workbooks(name2)
    On Error GoTo 0
    
    If Wb2 Is Nothing Then
        MsgBox "Bitte stellen Sie sicher, dass '" & name2 & "' geöffnet ist!", vbCritical
        Exit Sub
    End If
    
    sBody = ""
    moduleCount = 0
    textLineCount = 0
    
    ' Alle Modulnamen aus wb1 sammeln und alphabetisch sortieren
    ReDim sortedNames(0 To Wb1.VBProject.VBComponents.count - 1)
    i = 0
    For Each comp1 In Wb1.VBProject.VBComponents
        sortedNames(i) = comp1.NAME
        i = i + 1
    Next comp1
    
    SortArrayAlphabetically sortedNames
    
    ' Sortierte Module durchgehen und Prozeduren vergleichen
    For Each modName In sortedNames
        If UCase(modName) <> UCase(EXCLUDE_MODULE) Then
            Set comp1 = Wb1.VBProject.VBComponents(modName)
            
            On Error Resume Next
            Set comp2 = Wb2.VBProject.VBComponents(modName)
            On Error GoTo 0
            
            If Not comp2 Is Nothing Then
                If NormalizeCode(GetModuleCode(comp1)) <> NormalizeCode(GetModuleCode(comp2)) Then
                    procResult = CompareProcedures(comp1, comp2, name1, name2)
                    If Trim(procResult) <> "" Then
                        moduleCount = moduleCount + 1
                        sBody = sBody & procResult
                        
                        ' Textzeilen exakt zählen
                        lines = Split(Trim(procResult), vbCrLf)
                        For Each lineItem In lines
                            If Trim(lineItem) <> "" Then textLineCount = textLineCount + 1
                        Next lineItem
                    Else
                        ' Falls echte Code-Abweichungen in Declarations vorliegen (ohne Attribute)
                        moduleCount = moduleCount + 1
                        textLineCount = textLineCount + 1
                        sBody = sBody & Left(comp1.NAME & Space(20), 20) & " - [Declarations]" & vbTab & "different Code" & vbCrLf
                    End If
                End If
            Else
                moduleCount = moduleCount + 1
                textLineCount = textLineCount + 1
                sBody = sBody & Left(comp1.NAME & Space(20), 20) & " - [Modul]" & vbTab & "exists in " & name1 & ", NOT in " & name2 & vbCrLf
            End If
            Set comp2 = Nothing
        End If
    Next modName
    
    ' Module prüfen, die NUR in wb2 existieren
    For Each comp2 In Wb2.VBProject.VBComponents
        If UCase(comp2.NAME) <> UCase(EXCLUDE_MODULE) Then
            On Error Resume Next
            Set comp1 = Wb1.VBProject.VBComponents(comp2.NAME)
            On Error GoTo 0
            
            If comp1 Is Nothing Then
                moduleCount = moduleCount + 1
                textLineCount = textLineCount + 1
                sBody = sBody & Left(comp2.NAME & Space(20), 20) & " - [Modul]" & vbTab & "exists in " & name2 & ", NOT in " & name1 & vbCrLf
            End If
            Set comp1 = Nothing
        End If
    Next comp2
    
    ' Header mit genauer Anzahl an Textzeilen generieren
    sHeader = "VergleichA1-A2.txt: Noch " & moduleCount & " Module (" & textLineCount & " Textzeilen) zu reparieren" & vbCrLf
    sHeader = sHeader & "==========================================================================" & vbCrLf
    sHeader = sHeader & "PROZEDUREN-VERGLEICH (DETAIL - NUR ABWEICHUNGEN): " & name1 & " <--> " & name2 & vbCrLf
    sHeader = sHeader & "==========================================================================" & vbCrLf & vbCrLf
    
    ' Ausgabe über die existierende Show-Sub
    show sHeader & sBody
End Sub

' ==============================================================================
' 2. REPARATUR-PROZEDUR
' ==============================================================================
Sub Repair(ByVal NameOfModul As String)
    Dim EXCLUDE_MODULE As String
    EXCLUDE_MODULE = "M80" ' Das eigene Vergleichs-Modul vor Löschung schützen
    
    Dim Wb1 As Workbook, Wb2 As Workbook
    Dim comp1 As Object, comp2 As Object
    Dim cm1 As Object, cm2 As Object
    Dim name1 As String, name2 As String
    
    If UCase(NameOfModul) = UCase(EXCLUDE_MODULE) Then
        MsgBox "Das Modul '" & EXCLUDE_MODULE & "' enthält das Reparatur-Skript selbst und kann nicht überschrieben werden!", vbExclamation
        Exit Sub
    End If
    
    Set Wb1 = ThisWorkbook
    name1 = Wb1.NAME
    name2 = "A2 corrupt.xlsm"
    
    On Error Resume Next
    Set Wb2 = Workbooks(name2)
    On Error GoTo 0
    
    If Wb2 Is Nothing Then
        MsgBox "Bitte stellen Sie sicher, dass '" & name2 & "' geöffnet ist!", vbCritical
        Exit Sub
    End If
    
    On Error Resume Next
    Set comp1 = Wb1.VBProject.VBComponents(NameOfModul)
    Set comp2 = Wb2.VBProject.VBComponents(NameOfModul)
    On Error GoTo 0
    
    If comp2 Is Nothing Then
        MsgBox "Das Modul '" & NameOfModul & "' existiert nicht in " & name2 & "!", vbExclamation
        Exit Sub
    End If
    
    If comp1 Is Nothing Then
        Set comp1 = Wb1.VBProject.VBComponents.Add(comp2.Type)
        comp1.NAME = NameOfModul
    End If
    
    Set cm1 = comp1.CodeModule
    Set cm2 = comp2.CodeModule
    
    ' Kompletter Austausch des Modulinhalts bei Abweichung
    If NormalizeCode(GetModuleCode(cm1)) <> NormalizeCode(GetModuleCode(cm2)) Then
        cm1.DeleteLines 1, cm1.CountOfLines
        If cm2.CountOfLines > 0 Then
            cm1.InsertLines 1, cm2.lines(1, cm2.CountOfLines)
        End If
    End If
    
    CompareCode
End Sub

' ==============================================================================
' 3. HILFSFUNKTIONEN
' ==============================================================================

Private Function CompareProcedures(comp1 As Object, comp2 As Object, name1 As String, name2 As String) As String
    Dim cm1 As Object, cm2 As Object
    Dim LineNum As Long, ProcName As String, ProcKind As Long
    Dim code1 As String, code2 As String
    Dim res As String
    Dim checkedProcs As Object
    Dim modNameFormatted As String
    Dim procKey As String
    
    Set cm1 = comp1.CodeModule
    Set cm2 = comp2.CodeModule
    Set checkedProcs = CreateObject("Scripting.Dictionary")
    
    modNameFormatted = Left(comp1.NAME & Space(20), 20)
    
    ' Prozeduren aus Modul 1 durchgehen
    LineNum = 1
    Do While LineNum <= cm1.CountOfLines
        ProcName = cm1.ProcOfLine(LineNum, ProcKind)
        If ProcName <> "" Then
            procKey = ProcName & "_" & ProcKind
            If Not checkedProcs.Exists(procKey) Then
                checkedProcs.Add procKey, True
                
                code1 = GetProcedureCode(cm1, ProcName, ProcKind)
                code2 = GetProcedureCode(cm2, ProcName, ProcKind)
                
                If code2 = "" Then
                    res = res & modNameFormatted & " - " & ProcName & vbTab & "exists in " & name1 & ", NOT in " & name2 & vbCrLf
                ElseIf NormalizeCode(code1) <> NormalizeCode(code2) Then
                    res = res & modNameFormatted & " - " & ProcName & vbTab & "different Code" & vbCrLf
                End If
            End If
            LineNum = LineNum + cm1.ProcCountLines(ProcName, ProcKind)
        Else
            LineNum = LineNum + 1
        End If
    Loop
    
    ' Prozeduren suchen, die NUR in Modul 2 existieren
    LineNum = 1
    Do While LineNum <= cm2.CountOfLines
        ProcName = cm2.ProcOfLine(LineNum, ProcKind)
        If ProcName <> "" Then
            procKey = ProcName & "_" & ProcKind
            If Not checkedProcs.Exists(procKey) Then
                checkedProcs.Add procKey, True
                res = res & modNameFormatted & " - " & ProcName & vbTab & "exists in " & name2 & ", NOT in " & name1 & vbCrLf
            End If
            LineNum = LineNum + cm2.ProcCountLines(ProcName, ProcKind)
        Else
            LineNum = LineNum + 1
        End If
    Loop
    
    CompareProcedures = res
End Function

Private Function GetProcedureCode(cM As Object, ByVal ProcName As String, ByVal ProcKind As Long) As String
    On Error Resume Next
    Dim StartLine As Long, countLines As Long
    StartLine = cM.ProcStartLine(ProcName, ProcKind)
    countLines = cM.ProcCountLines(ProcName, ProcKind)
    
    If Err.Number = 0 And StartLine > 0 And countLines > 0 Then
        GetProcedureCode = cM.lines(StartLine, countLines)
    Else
        GetProcedureCode = ""
    End If
    On Error GoTo 0
End Function

Private Function GetModuleCode(compOrCm As Object) As String
    On Error Resume Next
    Dim cM As Object
    Set cM = compOrCm.CodeModule
    If cM Is Nothing Then Set cM = compOrCm
    
    If cM.CountOfLines > 0 Then
        GetModuleCode = cM.lines(1, cM.CountOfLines)
    Else
        GetModuleCode = ""
    End If
    On Error GoTo 0
End Function

' Entfernt versteckte VB-Attribute, Leerzeichen & Leerzeilen für verlässliche Vergleiche
Private Function NormalizeCode(ByVal sCode As String) As String
    Dim lines() As String
    Dim i As Long
    Dim res As String
    Dim lineStr As String
    
    sCode = Replace(sCode, vbCr, "")
    lines = Split(sCode, vbLf)
    
    For i = LBound(lines) To UBound(lines)
        lineStr = Trim(lines(i))
        ' Ausblenden von unsichtbaren/automatischen VB-Attributen und Leerzeilen
        If lineStr <> "" And Not (lineStr Like "Attribute VB_*") Then
            res = res & lineStr & vbCrLf
        End If
    Next i
    
    NormalizeCode = Trim(res)
End Function

Private Sub SortArrayAlphabetically(ByRef Arr() As String)
    Dim i As Long, j As Long
    Dim Temp As String
    For i = LBound(Arr) To UBound(Arr) - 1
        For j = i + 1 To UBound(Arr)
            If UCase(Arr(i)) > UCase(Arr(j)) Then
                Temp = Arr(i)
                Arr(i) = Arr(j)
                Arr(j) = Temp
            End If
        Next j
    Next i
End Sub

Sub CopyOneSheetNoCode_TEST()
    CopyOneSheetNoCode "T9", ThisWorkbook.path & "\A2 corrupt.xlsm"
    EE 1: Beep
End Sub

Sub CopyOneSheetNoCode(ByVal NameOfSheet As String, PathOfExcelFileToCopyFrom$)
    Dim Wb1 As Workbook, Wb2 As Workbook
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim fileName As String
    Dim openedByScript As Boolean
    
    Set Wb1 = ThisWorkbook
    openedByScript = False
    
    ' Prüfen, ob der übergebene Pfad nicht leer ist
    If Trim(PathOfExcelFileToCopyFrom) = "" Then
        MsgBox "FEHLER: Es wurde kein Dateipfad angegeben!", vbCritical
        Exit Sub
    End If
    
    ' Dateinamen aus dem Vollpfad extrahieren (z. B. "A2 corrupt.xlsm")
    fileName = Mid(PathOfExcelFileToCopyFrom, InStrRev(PathOfExcelFileToCopyFrom, "\") + 1)
    
    ' 1. Prüfen, ob die Mappe bereits in Excel geöffnet ist
    On Error Resume Next
    Set Wb2 = Workbooks(fileName)
    On Error GoTo 0
    
    ' 2. Falls nicht geöffnet -> Mappe im Hintergrund (schreibgeschützt) öffnen
    If Wb2 Is Nothing Then
        If Dir(PathOfExcelFileToCopyFrom) = "" Then
            MsgBox "FEHLER: Die Datei wurde nicht gefunden:" & vbCrLf & PathOfExcelFileToCopyFrom, vbCritical
            Exit Sub
        End If
        
        On Error Resume Next
        Set Wb2 = Workbooks.Open(fileName:=PathOfExcelFileToCopyFrom, ReadOnly:=True)
        On Error GoTo 0
        
        If Wb2 Is Nothing Then
            MsgBox "FEHLER: Die Datei konnte nicht geöffnet werden!", vbCritical
            Exit Sub
        End If
        
        openedByScript = True
    End If
    
    ' Prüfen, ob das Tabellenblatt in beiden Arbeitsmappen existiert
    On Error Resume Next
    Set ws1 = Wb1.Worksheets(NameOfSheet)
    Set ws2 = Wb2.Worksheets(NameOfSheet)
    On Error GoTo 0
    
    If ws1 Is Nothing Or ws2 Is Nothing Then
        MsgBox "FEHLER: Das Blatt '" & NameOfSheet & "' fehlt in einer der Mappen!", vbExclamation
        If openedByScript Then Wb2.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Error Handling aktivieren
    On Error GoTo CleanUpError
    
    ' Performance- & Sicherheitseinstellungen
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    ' Shapes/Objekte auf dem Zielblatt löschen
    On Error Resume Next
    ws1.DrawingObjects.Delete
    On Error GoTo CleanUpError
    
    ' Zielblatt vollständig leeren & eventuellen Blattschutz aufheben
    ws1.Cells.Clear
    ws1.Unprotect
    
    ' Kopieren über die getrennte Copy & PasteSpecial-Methode
    ws2.UsedRange.Copy
    ws1.Range(ws2.UsedRange.Address).PasteSpecial xlPasteAll
    
    ' Kopierrahmen sicher aufheben
    Application.CutCopyMode = False

CleanUpExit:
    ' Einstellungen wieder zurücksetzen
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Falls die Quell-Mappe vom Skript geöffnet wurde -> wieder schließen
    If openedByScript And Not Wb2 Is Nothing Then
        Wb2.Close SaveChanges:=False
    End If
    
    Beep
    Exit Sub

CleanUpError:
    MsgBox "VBA-Fehler " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanUpExit
End Sub

