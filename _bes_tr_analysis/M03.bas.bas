Attribute VB_Name = "M03"
Option Explicit 'M03

Sub zzz_M03()
    showProcs "text"
    Application.EnableEvents = True
    'RenameModule "Tabelle1", "T4"
    'show ArrC(59)
End Sub

Sub Change_a0_to_a00()
    'wurde nur 1 x zur Bereinigung benutzt
    Dim Fi$, Fi2$, Fo$, s$, s2$, c%, i%, j%, Arr1() As String, Arr2() As String
    DoArr
    s = Get_Paths_ofAllSubfolders_OneLevel(ArrC(3)) + vbCrLf 'ArrC(3) = Path of Folder "Events"
    s = s + Get_Paths_ofAllSubfolders_OneLevel(ArrC(4))      'ArrC(4) = Path of Folder "Leute"
    Arr1 = Split(s, vbCrLf)
    For i = 0 To UBound(Arr1)
        'Schleife über alle Ordner in 'Events' und 'Leute'
        Fo = Arr1(i)            'path of one folder
        s = Get_AllFilePaths_WithMyStringInFileName_OfOneFolder(Fo, ".")
        Arr2 = Split(s, vbCrLf)
        For j = 0 To UBound(Arr2)
            'Schleife über alle Files eines Ordners
            Fi = Arr2(j)        'path of one file
            If Fi Like "* [a-z]# *" Then
                c = Get_PositionOf_LikeString_inside_MyString(" [a-z]# ", Fi)
                Fi2 = Left(Fi, c + 1) + "0" + Mid(Fi, c + 2)
                's2 = s2 + Fi + vbCrLf + Fi2 + vbCrLf + vbCrLf
                RenameFile Fi, Fi2
            ElseIf Fi Like "* [a-z]#.*" Then
                c = Get_PositionOf_LikeString_inside_MyString(" [a-z]#.", Fi)
                Fi2 = Left(Fi, c + 2) + "0" + Mid(Fi, c + 3)
                's2 = s2 + Fi + vbCrLf + Fi2 + vbCrLf + vbCrLf
                RenameFile Fi, Fi2
            End If
        Next
        'show s2
    Next
    Beep
End Sub

Function Get_PositionOf_LikeString_inside_MyString(LikeString$, MyString$) As Integer
    Dim i%, L%
    If MyString Like "*" + LikeString + "*" Then
        L = Len(LikeString) - 4 '[a-z] hat in MyString nur die Länge 1
        For i = 1 To Len(MyString) - L
          If Mid(MyString, i, L) Like LikeString Then Exit For
        Next
        Get_PositionOf_LikeString_inside_MyString = i
    Else
        Get_PositionOf_LikeString_inside_MyString = 0
    End If
End Function

Function Add_LTV_Nation_InListOfCompetitors(s$) As String
    'Called from    Update_T4CompetitorsList_LLinksInEFolder
    s = Delete_EmptyRowsInString(s)
    
    'Baden
        s = Replace(s, "|TSG Bruchsal|||", "|TSG Bruchsal|Baden|D|")
        s = Replace(s, "|TSG Bruchsal||", "|TSG Bruchsal|Baden|")
        s = Replace(s, "|TV Forst|||", "|TV Forst|Baden|D|")
        s = Replace(s, "|FT Freiburg|||", "|FT Freiburg|Baden|D|")
        s = Replace(s, "|TV Gernsbach|||", "|TV Gernsbach|Baden|D|")
        s = Replace(s, "|KuSG Leimen|||", "|KuSG Leimen|Baden|D|")
        s = Replace(s, "|TV Villingen|||", "|TV Villingen|Baden|D|")
        s = Replace(s, "|TV Villingen||", "|TV Villingen|Baden|")
        s = Replace(s, "|TSG Wiesloch|||", "|TSG Wiesloch|Baden|D|")
    'Bayern
        s = Replace(s, "|TV Erlangen|||", "|TV Erlangen|Bayern|D|")
        s = Replace(s, "|TSV Gauting|||", "|TSV Gauting|Bayern|D|")
        s = Replace(s, "|ASV Nürnberg|||", "|ASV Nürnberg|Bayern|D|")
        s = Replace(s, "|TV Obing|||", "|TV Obing|Bayern|D|")
    'Berlin
        s = Replace(s, "|PSV Berlin|||", "|PSV Berlin|Berlin|D|")
        s = Replace(s, "|PSV Berlin||", "|PSV Berlin|Berlin|")
        s = Replace(s, "|TSG Steglitz|||", "|TSG Steglitz|Berlin|D|")
        s = Replace(s, "|TSG Steglitz||", "|TSG Steglitz|Berlin|")
        s = Replace(s, "|TSV Spandau|||", "|TSV Spandau|Berlin|D|")
    'Bremen
        s = Replace(s, "|TVL Bremerhaven|||", "|TVL Bremerhaven|Bremen|D|")
    'Hamburg
        s = Replace(s, "|PolSV Hamburg|||", "|PolSV Hamburg|Hamburg|D|")
    'Mittelrhein
        s = Replace(s, "|TB Andernach|||", "|TB Andernach|Mittelrhein|D|")
        s = Replace(s, "|MTV Bad Kreuznach|||", "|MTV Bad Kreuznach|Mittelrhein|D|")
    'Niedersachsen
        s = Replace(s, "|MTV Elze|||", "|MTV Elze|Niedersachsen|D|")
        s = Replace(s, "|TB Osnabrück|||", "|TB Osnabrück|Niedersachsen|D|")
        s = Replace(s, "|TGJ Salzgitter|||", "|TGJ Salzgitter|Niedersachsen|D|")
        s = Replace(s, "|TGJ Salzgitter||", "|TGJ Salzgitter|Niedersachsen|")
        s = Replace(s, "|Wolfsburg|||", "|Wolfsburg|Niedersachsen|D|")
    'Pfalz
        s = Replace(s, "|TSG Mutterstadt|||", "|TSG Mutterstadt|Pfalz|D|")
    'Rheinland
        s = Replace(s, "|TuS Altenessen|||", "|TuS Altenessen|Rheinland|D|")
        s = Replace(s, "|TV Düren|||", "|TV Düren|Rheinland|D|")
        s = Replace(s, "|TV Düsseldorf|||", "|TV Düsseldorf|Rheinland|D|")
        s = Replace(s, "|PostSV Essen|||", "|PostSV Essen|Rheinland|D|")
        s = Replace(s, "|TB Solingen|||", "|TB Solingen|Rheinland|D|")
        s = Replace(s, "|TG Süchteln|||", "|TG Süchteln|Rheinland|D|")
    'Schleswig-Holstein
        s = Replace(s, "|VFL Pinneberg|||", "|VFL Pinneberg|Schleswig-Holstein|D|")
    'Schwaben
        s = Replace(s, "|TSV Leinfelden|||", "|TSV Leinfelden|Schwaben|D|")
        s = Replace(s, "|TV Sulzbach|||", "|TV Sulzbach|Schwaben|D|")
    'Westfalen
        s = Replace(s, "|TSG Datteln|||", "|TSG Datteln|Westfalen|D|")
        s = Replace(s, "|SU Witten-Annen|||", "|SU Witten-Annen|Westfalen|D|")
    '-
        s = Replace(s, "|GB|||", "|||GB|")
        
    'show TAB_Simulation(s)
    EE 0: Sheets("T4").[D2].Select: EE 1
    Add_LTV_Nation_InListOfCompetitors = s
End Function

Function Add_mw_Klasse_InListOfCompetitors(s$) As String
    'Called from    Update_T4CompetitorsList_LLinksInEFolder
    's              Zeilen wie "|Maier|Max| m |Tu|TV Bonn|||1999 DM Bonn |Synchron|1|"
    
    Dim L$, L2$, ListOfCompetitors$, NaVo$, sNew$, sOld$, Verein$
    Dim c&, C1&, C2&, C3&, i%, j%, Arr1() As String, Arr2() As String
    
    Arr1 = Split(s, vbCrLf)
    For i = 0 To UBound(Arr1)
        L = Arr1(i)         'one complete line
        
        If L Like "*|||D|*" Then
            'L enthält keinen Eintrag für Verein (und LTV)
            c = InStr(2, L, "|"): c = InStr(c + 1, L, "|")
            NaVo = Left(L, c)           '|Nachname|Vorname| '|Maier|Max|
            c = 1
            For j = 1 To 99
                c = InStr(c, s, NaVo)
                If c = 0 Then Exit For
                L2 = Mid(s, c, InStr(c + 1, s, vbCrLf) - c) 'gefundene Zeile mit NaVo
                Arr2 = Split(L2, "|")
                Verein = Arr2(5)
                If Verein <> "" Then
                    L2 = Replace(L2, "|||D|", "|" + Verein + "||D|")
                    s = Replace(s, L, L2)
                    Exit For
                Else
                    c = c + Len(L2) - 3
                End If
            Next
            
        End If

        If L Like "*|Synchron|*" Then
            'L könnte einen falschen w-Eintrag enthalten (bei einem Mix-Wettkampf)
            c = InStr(1, L, "| w |")
            NaVo = Left(L, c)           '|Nachname|Vorname| '|Maier|Max|
            'Suche nach NaVo mit "| m |"
            If InStr(1, s, NaVo + " m |") > 0 Then
                L2 = Replace(L, "| w |", "| m |")
                L2 = Replace(L2, "|Ti|", "|Tu|")
                L2 = Replace(L2, "|Juti|", "|Jutu|")
                L2 = Replace(L2, "|Schi|", "|Schü|")
                s = Replace(s, L, L2)
                Arr1(i) = L
            End If
        End If

        If L Like "|[A-Z]*|[A-Z]*|  ||*" Then
            'L enthält keinen Eintrag für m/w und Tu/Ti/...
            c = InStr(1, L, "|  ||")
            NaVo = Left(L, c)           '|Nachname|Vorname| '|Maier|K.|
            '    = Zeile von NaVo, bei der m/w fehlt
            sOld = NaVo + "  ||"
            
            'Suche nach NaVo mit "| m |" (vom Ende der Liste s nach oben)
            C1 = InStrRev(s, NaVo + " m |")
            If C1 > 0 Then
                's enthält |Na|Vo| m |
                C3 = InStr(C1 + Len(NaVo) + 5, s, "|") 'c3 = Position des "|" nach "...|Tu"
                sNew = Mid(s, C1, C3 - C1 + 1)
                s = Replace(s, sOld, sNew)
            Else
                C2 = InStrRev(s, NaVo + " w |")
                If C2 > 0 Then
                    C3 = InStr(C2 + Len(NaVo) + 5, s, "|")  'c3 = Position des "|" nach "...|Tu"
                    'show s
                    s = Replace(s, NaVo + " ||", Mid(s, C2, C3 - C2 + 1))
                End If
            End If
        End If
        
        If InStr(ListOfCompetitors, L) = 0 Then ListOfCompetitors = ListOfCompetitors + L + vbCrLf
    Next
    
    Add_mw_Klasse_InListOfCompetitors = ListOfCompetitors
End Function

Sub TEST_PowerShellOutputToString()
    ' CallPowerShell.vbs
    ' Call PowerShell from VBScript
    ' Author: ITomation (http://itomation.ca)
    ' Version 1.0 - 2015-11-27
    ' --------------------------------------------'
    Dim strPSCommand$, strDOSCommand$, objShell, objExec, s$, p$
    p = "G:\Archiv Photos\- Photos ubes\Photos BR\2023 BR\20230100 BR"
    
    'Construct PowerShell Command (PS syntax)
        strPSCommand = "get-acl '" + p + "' " _
            + "| foreach-object{ $_.Access } " _
            + "|select  -property IdentityReference, AccessControlType, FileSystemRights"
    'Consruct DOS command to pass PowerShell command (DOS syntax)
        strDOSCommand = "powershell -command " & strPSCommand & ""
    'Create shell object
        Set objShell = CreateObject("Wscript.Shell")
    'Execute the combined command
        Set objExec = objShell.Exec(strDOSCommand)
    'Read output into VBS variable
        s = objExec.StdOut.ReadAll
    show CStr(s)
End Sub

Public Sub TEST_ExecuteStringWithPowerShell()
    'Hier:  Umbenennungen; .JPG --> .jpg
    'Vorbereitung
        Dim p$, strGetFiles$, strRename$, varRetval As Variant
        'p= "G:\Archiv Photos\- Photos ubes\Photos BR\2023 BR"
        p = "G:\Archiv Photos\- Photos ubes"
        
    strGetFiles = "Get-Childitem -Path '" + p + "' -Filter '*.JPG' -Recurse -File "
    'Notes: If using -Recurse then the last '\*' in the path is not needed,
    '       otherwise needed if using -Filter.
    '       The use of -Filter and -File is just to reduce
    '       the number of files going through the pipeline
    '       and not needed if the -replace operator provides all the constraint required.
    
    strRename = " Rename-Item -NewName {$_.Name -replace '.JPG', '.jpg'} -PassThru "
    'Notes: The -replace comparison operator is not case sensitive on selection but writes as shown.
    '       If case sensitivity on selection is needed use -creplace instead.
    '       Rename-Item does not provide output and that is way the -PassThru param is added.
              
    varRetval = shell("Powershell.exe -noexit -Command " & strGetFiles & "|" & strRename, 1)
    'Show CStr(varRetval)
    'MsgBox "Enter 'exit' in PowerShell window to close the session."

End Sub

Sub TESTcolor()
    Dim s As String, r As Range
    Set r = Sheets("T1").Range("b2")
    s = CStr(getColor(r)) + vbCrLf + "RGB =" + getColor(r, 2) '255,230,153
    show s
End Sub

Sub SearchCodeModule()
    'liefert: Found at: Line: 45 Column: 12 (im Direktbereich)
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim FindWhat As String
    Dim sL As Long ' start line
    Dim eL As Long ' end line
    Dim sc As Long ' start column
    Dim EC As Long ' end column
    Dim Found As Boolean
    
    Set VBProj = ActiveWorkbook.VBProject
    Set VBComp = VBProj.VBComponents("W3")
    Set CodeMod = VBComp.CodeModule
    
    FindWhat = "Arbeitsauftraege("
    
    With CodeMod
        sL = 1
        eL = .CountOfLines
        sc = 1
        EC = 255
        Found = .Find(Target:=FindWhat, StartLine:=sL, StartColumn:=sc, _
            endline:=eL, EndColumn:=EC, _
            wholeWord:=True, MatchCase:=False, patternsearch:=False)
        Do Until Found = False
            Debug.Print "Found at: Line: " & CStr(sL) & " Column: " & CStr(sc)
            eL = .CountOfLines
            sc = EC + 1
            EC = 255
            Found = .Find(Target:=FindWhat, StartLine:=sL, StartColumn:=sc, _
                endline:=eL, EndColumn:=EC, _
                wholeWord:=True, MatchCase:=False, patternsearch:=False)
        Loop
    End With
End Sub

Public Function CodeFind(Optional FindMod As String = "", Optional FindProc As String = "", _
    Optional FindStr As String = "", Optional TypeOfSearch As Long = 0) As Long
    '---------------------------------------------------------------------------------------
    ' Procedure : CodeFind
    ' DateTime  : 7/5/2005 18:32
    ' Author    : Nelson Hochberg
    ' Purpose   : Find a module, a procedure and/or a string in code and highlight it
    ' Returns   : 0 if not found,  line number in module if found
    ' Syntax    : lngReturn = CodeFind ([FindMod],[FindProc],[FindStr],[TypeOfSearch])
    ' Arguments : Optional FindMod As String: Part of a name of a module
    '             Optional FindProc As String: Part of a name of a procedure
    '             Optional FindStr As String: Part of a string to search
    '             NOTE: One of the above three is required
    '             Optional TypeOfSearch As Long: -1 Find line number, 0 Find string,
    '                      >0 Continue search starting at line number: TypeOfSearch + 1
    ' Thanks    : To stevbe at Experts Exchange for the initial code.
    '---------------------------------------------------------------------------------------
    Dim vbc As VBIDE.VBComponent
    Dim cM As VBIDE.CodeModule
    Dim VBAEditor As VBIDE.VBE
    Dim VBProj As VBIDE.VBProject
    Dim StartLine As Long, startcol As Long, endline As Long, endcol As Long
    
    If FindMod <> "" Then
        CodeFind = FindModule(FindMod, vbc, cM)
            If CodeFind = False Then Exit Function
        If FindProc <> "" Then
            CodeFind = FindProcedure(FindProc, StartLine, startcol, endline, endcol, cM)
                If CodeFind = False Then Exit Function
            If FindStr <> "" Then
                CodeFind = FindString(FindStr, StartLine, startcol, endline, endcol, cM, TypeOfSearch)
                    If CodeFind = False Then Exit Function
            Else
                GoTo CodeLineFound
            End If
        Else
            StartLine = 1
            If FindStr <> "" Then
                CodeFind = FindString(FindStr, StartLine, startcol, endline, endcol, cM, TypeOfSearch)
                If CodeFind = False Then Exit Function
            Else
                GoTo CodeLineFound
            End If
        End If
    Else
        Set VBAEditor = Application.VBE
    '''''''''''''''''''''''''''''''''''''''''''
        Set VBProj = VBAEditor.ActiveVBProject
        For Each vbc In VBProj.VBComponents
    
    
            Set cM = vbc.CodeModule
            If FindProc <> "" Then
                CodeFind = FindProcedure(FindProc, StartLine, startcol, endline, endcol, cM)
                If CodeFind = False Then GoTo Nextvbc2 Else Exit For
            Else
                StartLine = 1
                If FindStr <> "" Then
                    CodeFind = FindString(FindStr, StartLine, startcol, endline, endcol, cM, TypeOfSearch)
                        If CodeFind = False Then GoTo Nextvbc2 Else Exit For
                Else
                    MsgBox "CodeFind: At least one of the following is required:" & vbCrLf & _
                        "    Module" & vbCrLf & "    Procedure" & vbCrLf & "    String"
                    CodeFind = False
                    Exit Function
                End If
            End If
Nextvbc2:
        Next vbc
        If CodeFind <> False Then
            If FindStr <> "" Then
                CodeFind = FindString(FindStr, StartLine, startcol, endline, endcol, cM, TypeOfSearch)
                If CodeFind = False Then Exit Function
            Else
                GoTo CodeLineFound
            End If
        End If
    End If
    
CodeLineFound:
    If CodeFind <> False Then
        If endline = -1 Then endline = 1
        If endcol = -1 Then endcol = 1
        cM.CodePane.show
        cM.CodePane.SetSelection StartLine, startcol, endline, endcol
    End If

 End Function

Private Function FindModule(FindMod As String, vbc As VBComponent, cM As CodeModule) As Long
    FindModule = False
    For Each vbc In VBE.VBProjects(1).VBComponents
        'FindMod is VBComponent
        If InStr(vbc.NAME, FindMod) > 0 Then
            Set cM = vbc.CodeModule
            FindModule = 1
            Exit For
        End If
    Next vbc
End Function

Private Function FindProcedure(FindProc As String, StartLine As Long, startcol As Long, _
                 endline As Long, endcol As Long, cM As CodeModule) As Long
    Dim lngFake As Long
    StartLine = 1
    startcol = 1
    If FindProc <> "" Then
        'search for procedure
        FindProcedure = False
        Do
            endline = -1
            endcol = -1
            If cM.Find(FindProc, StartLine, startcol, endline, endcol) = False Then Exit Do
            If InStr(cM.ProcOfLine(StartLine, lngFake), FindProc) Then
                FindProcedure = StartLine
                Exit Do
            End If
            startcol = endcol
        Loop
    End If
End Function

Private Function FindString(FindStr As String, StartLine As Long, startcol As Long, _
                 endline As Long, endcol As Long, cM As CodeModule, TypeOfSearch As Long) As Long
    If FindStr <> "" Then
        If TypeOfSearch > 0 Then StartLine = TypeOfSearch
        startcol = 1
        endline = -1
        endcol = -1
        Do
            If cM.Find(FindStr, StartLine, startcol, endline, endcol) = False Then
                FindString = False
                Exit Function
            End If
            If TypeOfSearch >= 0 Then Exit Do
            If startcol = 1 Then Exit Do
            startcol = endcol
        Loop
    End If
    FindString = endline
End Function

