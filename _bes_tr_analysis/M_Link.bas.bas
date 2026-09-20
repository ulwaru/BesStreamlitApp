Attribute VB_Name = "M_Link"
Option Explicit 'M_Link

Sub zzz_M_Link()
    'showProcs "SubFolders like"
        
    'show Get_TextOfAllOpenNotepadWindows
    EE 1: Beep
End Sub

Sub Create_OneLinkFile(FolderInWhitchTheLinkShouldBeCreated$, FileOrFolderToJumpTo$, _
        LinkNameToDisplay$, Optional Arg$ = "")
    'Called from    Create_LinkOfDgLeute
                    'Create_vhFolderAndLink
                    'T4_Add_EventLink_ToPersonFolders
                    'T4_Copy_°FilesOfOneEvent_ToPersonFolders
                    'T4_Create_EventLink_DateSquare_OfOneEventfolder_ToPersonFolders
                    'T4_Create_MissingBracLinks_OfOneEvent
                    'T6_Add_Links_PersonLTVClubNation
                    'Update_AllLinks_InsideStick_ToVolumeLetter_T
                    'Update_LeuteLinks_inOneEventFolder
                    'Update_Links_inside_FolderLeute
                    'Verify_OnePersonLinkOnly
    'Fil..ToJumpTo  =
    'Arg            = Argument; zusätzlich übergebener String; PathOfPersonFolder;
    '               if Arg="" then normalLink else vbsLink
    'Vorbereitung
        If FileOrFolderToJumpTo = "" Then Exit Sub
        Dim s$, v$: v = vbCrLf
        
    'Action
        s = FolderInWhitchTheLinkShouldBeCreated + "\" + LinkNameToDisplay + ".lnk"
        DeleteFile s
            'verweigert ggf. Zugriff; Zeitproblem?
        With CreateObject("WScript.Shell").CreateShortcut(s)
            .targetPath = FileOrFolderToJumpTo
            .arguments = Arg
            On Error GoTo Jump
            .Save
            Exit Sub
Jump:
        show "PROCEDURE:   Create_OneLinkFile" + v + "ERROR:       Pfad unkorrekt " _
           + "(ggf. wegen entsprechendem Eintrag in T5-Spalte 'Personenordner')" + v + v _
           + "FolderInWhitchTheLinkShouldBeCreated: " + FolderInWhitchTheLinkShouldBeCreated + v _
           + "FileOrFolderToJumpTo:                 " + FileOrFolderToJumpTo + v _
           + "LinkNameToDisplay:                    " + LinkNameToDisplay
        End With
End Sub


Sub Change_ExplorerWindow_ToBigIcons(PathOfFolder)
    Dim objShell As Object
    Dim objWindows As Object
    Dim objWin As Object
    Dim targetPath As String
    Dim windowFound As Boolean
    
    ' Pfad vereinheitlichen
    targetPath = LCase(Trim(PathOfFolder))
    If Right(targetPath, 1) = "\" And Len(targetPath) > 1 Then
        targetPath = Left(targetPath, Len(targetPath) - 1)
    End If
    
    Set objShell = CreateObject("Shell.Application")
    Set objWindows = objShell.Windows
    windowFound = False
    
    ' 1. Suchen, ob das Fenster bereits geöffnet ist
    For Each objWin In objWindows
        On Error Resume Next
        If InStr(1, objWin.FullName, "explorer.exe", vbTextCompare) > 0 Then
            Dim currentWinPath As String
            currentWinPath = LCase(Trim(objWin.Document.folder.Self.path))
            
            If Right(currentWinPath, 1) = "\" And Len(currentWinPath) > 1 Then
                currentWinPath = Left(currentWinPath, Len(currentWinPath) - 1)
            End If
            
            If currentWinPath = targetPath Then
                ' Ansicht umstellen
                objWin.Document.CurrentViewMode = 1
                objWin.Document.IconSize = 96
                windowFound = True
                Exit For
            End If
        End If
        On Error GoTo 0
    Next
    
    ' 2. FALLBACK: Wenn das Fenster NICHT offen war, neu öffnen und umstellen
    If Not windowFound Then
        ' Ordner im Explorer öffnen (1 = Normales Fenster focusiert)
        CreateObject("WScript.Shell").Run "explorer.exe """ & PathOfFolder & """", 1, False
        
        ' Kurz warten, bis das Fenster von Windows geladen wurde (0.5 Sekunden)
        Dim startTime As Double
        startTime = Timer
        Do While Timer < startTime + 0.5
            DoEvents
        Loop
        
        ' Das neue Fenster in der Shell-Sammlung suchen und umstellen
        For Each objWin In objShell.Windows
            On Error Resume Next
            If InStr(1, objWin.FullName, "explorer.exe", vbTextCompare) > 0 Then
                currentWinPath = LCase(Trim(objWin.Document.folder.Self.path))
                If Right(currentWinPath, 1) = "\" And Len(currentWinPath) > 1 Then
                    currentWinPath = Left(currentWinPath, Len(currentWinPath) - 1)
                End If
                
                If currentWinPath = targetPath Then
                    objWin.Document.CurrentViewMode = 1
                    objWin.Document.IconSize = 96
                    Exit For
                End If
            End If
            On Error GoTo 0
        Next
    End If
    
    Set objShell = Nothing
    Set objWindows = Nothing
End Sub

Sub Mark_MyFile(PathOfMyFile$)
    Dim objShell As Object
    Dim objWindows As Object
    Dim objWin As Object
    Dim fso As Object
    Dim parentFolderPath As String
    Dim fileName As String
    Dim folderOpened As Boolean
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(PathOfMyFile) Then
        MsgBox "Datei existiert nicht: " & PathOfMyFile, vbExclamation, "Fehler"
        Exit Sub
    End If
    
    parentFolderPath = LCase(Trim(fso.GetParentFolderName(PathOfMyFile)))
    fileName = fso.GetFileName(PathOfMyFile)
    
    Set objShell = CreateObject("Shell.Application")
    Set objWindows = objShell.Windows
    folderOpened = False
    
    ' 1. Wenn der Ordner bereits offen ist
    For Each objWin In objWindows
        On Error Resume Next
        If InStr(1, objWin.FullName, "explorer.exe", vbTextCompare) > 0 Then
            If LCase(Trim(objWin.Document.folder.Self.path)) = parentFolderPath Then
                
                ' Fenster in den Vordergrund holen
                Dim wshShell As Object
                Set wshShell = CreateObject("WScript.Shell")
                wshShell.AppActivate objWin.hwnd
                
                ' --- RADIKALE ABWAHL ALLER ELEMENTE ---
                Dim folderItem As Object
                For Each folderItem In objWin.Document.folder.items
                    ' Flag 0 = Hebt die Markierung für dieses Element explizit auf
                    objWin.Document.SelectItem folderItem, 0
                Next
                
                ' --- NEUE DATEI EXKLUSIV SELEKTIEREN ---
                ' Flag 1 = Auswählen, Flag 8 = In den Fokus scrollen
                Dim objItem As Object
                Set objItem = objWin.Document.folder.ParseName(fileName)
                
                If Not objItem Is Nothing Then
                    objWin.Document.SelectItem objItem, 1 + 8
                End If
                
                folderOpened = True
                Exit For
            End If
        End If
        On Error GoTo 0
    Next
    
    ' 2. Falls der Ordner noch geschlossen war
    If Not folderOpened Then
        CreateObject("WScript.Shell").Run "explorer.exe """ & parentFolderPath & """", 1, False
        
        ' Kurz warten, bis das Fenster bereit ist
        Dim startTime As Double
        startTime = Timer
        Do While Timer < startTime + 0.5
            DoEvents
        Loop
        
        ' Prozess für das neue Fenster wiederholen
        For Each objWin In objShell.Windows
            On Error Resume Next
            If LCase(Trim(objWin.Document.folder.Self.path)) = parentFolderPath Then
                
                ' Alle abwählen
                For Each folderItem In objWin.Document.folder.items
                    objWin.Document.SelectItem folderItem, 0
                Next
                
                ' Ziel-Datei auswählen
                objWin.Document.SelectItem objWin.Document.folder.ParseName(fileName), 1 + 8
                Exit For
            End If
            On Error GoTo 0
        Next
    End If
    
    Set fso = Nothing
    Set objShell = Nothing
    Set objWindows = Nothing
End Sub


Sub Show_InfoOfSomeLinks_TEST()
    Dim p$, v$: v = vbCrLf
'    p = "F:\Archiv Trampolin 1900-1999\Events\1959-07-09 Turnfest Basel_CH\Baechler, Kurt (TV Bern-Berna, CH).lnk"
'    p = p + v + "F:\Archiv Trampolin 1900-1999\Events\1959-07-09 Turnfest Basel_CH\Baechler, Kurt 2 (TV Bern-Berna, CH).lnk"
'    p = p + v + "F:\Archiv Trampolin 1900-1999\Events\1959-07-09 Turnfest Basel_CH\Baechler, Kurt 3.lnk"
'    p = p + v + "F:\Archiv TR\Archiv Trampolin\Events\1967-06-17 WM04 London_GB\Aaron, Syd (GB).lnk"
'    p = p + v + "F:\Archiv TR\Archiv Trampolin\Events\1972-09-23 WM07 Stuttgart\Czech, Ute (vh Latton Luxon Pitkamin, TGJ Salzgitter).lnk"
'    p = p + v + "F:\Archiv TR\Archiv Trampolin\Events\1972-09-23 WM07 Stuttgart\Czech, Ute (TGJ Salzgitter).lnk"
'    p = p + v + "F:\Archiv TR\Archiv Trampolin\Leute\Czech, Ute (vh Latton Luxon Pitkamin, TGJ Salzgitter)\19660508 ! DJM Nürnberg.lnk"
    
    p = "F:\Archiv Trampolin 1900-1999\Events\1959-07-14 Intern. Lehrgang Freiburg\Nissen, George (Cedar Rapids, USA).lnk"
    p = p + v + "F:\Archiv Trampolin 1900-1999\Events\1959-07-14 Intern. Lehrgang Freiburg\Schillinger, Roland (FT Freiburg).lnk"
    Show_InfoOfSomeLinks p
End Sub

Sub Show_InfoOfSomeLinks(Some_PathOfLink$)
    Dim p$, s$, v$, i%, L() As String
    v = vbCrLf: s = "ShowLinkInfo" + v + v
    L = Split(Some_PathOfLink, v)
    For i = 0 To UBound(L)
        p = L(i)
        s = s + Get_InfoOfOneLink(p)
    Next
    show s
End Sub

Sub Load_InfoOfOneLink(PathOfLink$, Arr() As String)
    'Called from    xxx
    'Action         liefert Arr(1 to 8)
    'Arr(1 to 8)    1 NameOfLink, 2 PathOfLinkFolder, 3 TargetPath, 4 CountOfArguments,
    '               5 Argument1, 6 Argument2, 7 PathOfIcon, 8 Description
    'Vorbereitung
        Dim A$, NameOfLink$, qq$, c%, Arg() As String, w As Object, L As Object
        ReDim Arr(1 To 8)
        Set w = CreateObject("WScript.Shell"): qq = Chr(34)
        Set L = w.CreateShortcut(PathOfLink)
    'Action
        '
        Arr(1) = NameOfPath(PathOfLink)                             'NameOfLink
        Arr(2) = Left(PathOfLink, InStrRev(PathOfLink, "\") - 1)    'PathOfLinkFolder
        Arr(3) = L.targetPath                                       'TargetPath
        Arr(7) = L.IconLocation                                     'PathOfIcon
        Arr(8) = L.Description                                      'Description
        If Right(Arr(1), 4) = ".lnk" Then Arr(1) = Left(Arr(1), Len(Arr(1)) - 4)
        If Right(Arr(7), 2) = ",0" Then Arr(7) = Left(Arr(7), Len(Arr(7)) - 2)
    'Get Arguments
        A = L.arguments
        ' = "F:\...\zzico\zz.vbs" "F:\..., John (GB)"
        If A = "" Then Arr(4) = "0": Arr(5) = "": Arr(6) = ""
        If Not A = "" Then
            A = Replace(A, qq, "|"):   A = Replace(A, "| ", "|")
            ' = C:\...\wscript.exe|F:\...\zzico\zz.vbs||F:\..., John (GB)|
            A = Replace(A, " |", "|"): A = Replace(A, "||", "|"):
            If Right(A, 1) = "|" Then A = Left(A, Len(A) - 1)
            If Left(A, 1) = "|" Then A = Mid(A, 2)
            ' = C:\...\wscript.exe|F:\...\zzico\zz.vbs|F:\..., John (GB)
            Arg = Split("|" + A, "|"): c = UBound(Arg)
            Arr(4) = CStr(c)                'CountOfArguments
            If c > 0 Then Arr(5) = Arg(1)   'Argument1
            If c > 1 Then Arr(6) = Arg(2)   'Argument2
        End If
End Sub

Sub Show_InfoOfOneLink_inOneLine(PathOfLink$)
    show Get_InfoOfOneLink_inOneLine(PathOfLink)
End Sub

Sub Kill_LinksWithBadTarget_OneFolder(PathOfOneFolder$)
    Dim s$, i&, A() As String
    s = Get_AllFilePaths_LikeMyStringInFileName_OfOneFolder(PathOfOneFolder, "*.lnk")
    A = Split(s, vbCrLf)
    For i = 0 To UBound(A)
        If Not TargetOfVbsLinkExists(A(i)) Then DeleteFile A(i)
    Next
End Sub

Function TargetOfVbsLinkExists(PathOfLink$) As Boolean
    'Vorbereitung
        Dim A$, p$, qq$, c%, w As Object, L As Object
        Dim D4$
        Set w = CreateObject("WScript.Shell"): qq = Chr(34)
        Set L = w.CreateShortcut(PathOfLink)
        
        A = L.arguments ' = "F:\...\zzico\zz.vbs" "F:\..., John (GB)"
        If A = "" Then Exit Function
        If A Like qq + "*" + qq + "* *" + qq Then
            c = InStr(1, A, qq + " " + qq)
            D4 = Mid(A, c + 3): D4 = Left(D4, Len(D4) - 1)
        End If
        p = "F:\" + D4
        If FolderExists(p) Then TargetOfVbsLinkExists = True
End Function

Function Get_InfoOfOneLink_inOneLine(PathOfLink$) As String
    'Vorbereitung
        Dim A$, qq$, c%, w As Object, L As Object
        Dim D$, D1$, D2$, D3$, D4$, D5$, D6$
        Set w = CreateObject("WScript.Shell"): qq = Chr(34)
        Set L = w.CreateShortcut(PathOfLink)
        
        D1 = PathOfLink         'PathOfLink
        D2 = L.targetPath       'TargetPath   'C:\Windows\System32\wscript.exe
        A = L.arguments ' = "F:\...\zzico\zz.vbs" "F:\..., John (GB)"
            If A = "" Then
                D3 = "": D4 = ""
            ElseIf A Like qq + "*" + qq + "* *" + qq Then
                c = InStr(1, A, qq + " " + qq): D3 = Mid(Left(A, c - 1), 2)
                D4 = Mid(A, c + 3): D4 = Left(D4, Len(D4) - 1)
            Else: D3 = A: D4 = ""
            End If
        D5 = L.IconLocation: If Right(D5, 2) = ",0" Then D5 = Left(D5, Len(D5) - 2)
        D6 = L.Description
        D = "|" + D1 + "|" + D2 + "|" + D3 + "|" + D4 + "|" + D5 + "|" + D6 + "|"
        Get_InfoOfOneLink_inOneLine = D
End Function

Sub Show_InfoOfOneLink(PathOfLink$)
    show Get_InfoOfOneLink(PathOfLink)
End Sub

Function Get_InfoOfOneLink(PathOfLink$) As String
    'Vorbereitung
        Dim A$, F$, p$, qq$, s$, v$, c%, i%, Arg() As String, w As Object, L As Object
        Set w = CreateObject("WScript.Shell"): qq = Chr(34)
        Set L = w.CreateShortcut(PathOfLink): v = vbCrLf
        F = Left(PathOfLink, InStrRev(PathOfLink, "\") - 1)
        p = L.IconLocation
        If Right(p, 2) = ",0" Then p = Left(p, Len(p) - 2)
    'Show
        s = s + "NameOfLink         = " + NameOfPath(PathOfLink) + v
        s = s + "PathOfLinkFolder   = " + F + v
        s = s + "TargetPath         = " + L.targetPath + v
    'Get Arguments
        A = L.arguments
        ' = "F:\...\zzico\zz.vbs" "F:\..., John (GB)"
        If A = "" Then
            c = 0
        Else
            A = Replace(A, qq, "|"):   A = Replace(A, "| ", "|")
            ' = C:\...\wscript.exe|F:\...\zzico\zz.vbs||F:\..., John (GB)|
            A = Replace(A, " |", "|"): A = Replace(A, "||", "|"):
            If Right(A, 1) = "|" Then A = Left(A, Len(A) - 1)
            If Left(A, 1) = "|" Then A = Mid(A, 2)
            ' = C:\...\wscript.exe|F:\...\zzico\zz.vbs|F:\..., John (GB)
            Arg = Split("|" + A, "|"): c = UBound(Arg)
        End If
    'Show Arguments/Items
        s = s + "CountOfArguments   = " + CStr(c) + v
        For i = 1 To c
            s = s + "Argument " + CStr(i) + "         = " + Arg(i) + v
        Next
    'Show
        s = s + "PathOfIcon         = " + p + v
        s = s + "Description        = " + L.Description + v + v + v
    'Finals
        Get_InfoOfOneLink = s
End Function

Function IsVbsLink(PathOfLink$) As Boolean
    'Called from    ChangeOneLink_ToVbsLinkToPerson
    
    'Vorbereitung
        Dim w As Object, L As Object
        Set w = CreateObject("WScript.Shell")
        'show PathOfLink
        Set L = w.CreateShortcut(PathOfLink)
    'Action
        If L.targetPath Like "*\zzico\zz.vbs" Then IsVbsLink = True
End Function

Sub Link_Handlings()
    'Check_AllFolders_Events         '+
        'Check_OneFolder            [acts on each SubFolder]
    'Check_AllFolders_ClubsNations   '+
        'Check_OneFolder            [acts on each SubFolder]
    'FullCheckFolder_Leute          '-
    'Check_AllFolders_Faces          '+
        'Check_OneFolder            [acts on this one Folder]
    'Create_AllPersonLinks_Register       '+
End Sub

Sub Check_AllLinks_PersFold_to_LTVFold()

End Sub

Sub Create_AllPersonLinks_Register()
    'Called from    xxx
    'Action         alle RegisterLinks neu erstellen, delete old
    
    'Vorbereitung
        Dim mw$, NameOfLink$, s$, v$, i%, max%
        Dim A() As String, B() As String, w As Object, L As Object
        v = vbCrLf:  DoArrc
        Set w = CreateObject("WScript.Shell")
    'Delete
        Delete_AllLinkFiles_insideOneFolder ArrC(5)
    'Load some T5-Columns + Adds
        '[1]Vn    [2]Nn    [3]mw    [4]Club  [5]LTV   [6]Nation   [7]Folder   [8]rn    [9]gn
        '[10]VnNn [11]NnVn [12]vh1  [13]vh2  [14]vh3  [15]LiVnNn  [16]LiNnVn  [17]LiRn
        T5_Load_1_to_17 A: max = UBound(A, 1)
    's = All LinkNames (= PersonFolderNames) for VbsLinks to create
        For i = 1 To UBound(A, 1)
            s = s + v + A(i, 15) + "|" + A(i, 3) + "|" + A(i, 7) + v _
                      + A(i, 16) + "|" + A(i, 3) + "|" + A(i, 7)
            If A(i, 17) <> "" Then s = s + v + A(i, 17) + "|" + A(i, 3) + "|" + A(i, 7)
        Next
        s = Mid(s, 3): B = Split(s, v): QuickSort B: max = UBound(B) + 1 ': showArray B
    'Create each VbsLink
        For i = 0 To UBound(B)
            NameOfLink = Left(B(i), InStr(1, B(i), "|") - 1) '+ ".lnk"
            Create_OneVbsLink_Person ArrC(5) + NameOfLink
            DoEvents
        Next
    'Finals
        Beep
End Sub

Sub Check_AllFolders_Faces()
    Dim p$, v$, Report$
    Call DoArrc: p = ArrC(2) + "\Faces"
    Call UF3_Format: UF3_CountDown 1, "FullCheck 1 folder"
    Check_OneFolder p, 1, 1, Report
    'Finals
        If Report = "" Then show "Check_AllFolders_Faces" + v + v + "- No Errors to report. -" _
                       Else show "Check_AllFolders_Faces" + v + v + Report
        UF3.Hide: Beep
End Sub

Sub Check_AllFolders_Events()
    Dim Report$, v$, i%, max%, F() As String
    Load_ArrPathsOfAllFolders_Events F
    max = UBound(F): v = vbCrLf
    Call UF3_Format: UF3_CountDown 1, "FullCheck " + CStr(max) + " folders"
    For i = 1 To UBound(F) '1 To UBound(F)
        Check_OneFolder F(i), i, max, Report
        DoEvents
    Next
    'Finals
        If Report = "" Then show "Check_AllFolders_Events" + v + v + "- No Errors to report. -" _
                       Else show "Check_AllFolders_Events" + v + v + Report
        UF3.Hide: Beep
End Sub

Sub Check_NameOfLink(PathOfLink$, FromFolder$, T5NameOfPersonFolder$, VnNn$, i%, Report$)
    'Called from    ChangeOneLink_ToVbsLinkToPerson
    'FromFolders    Events        NameOfLink = T5NameOfPersonFolder|vhClub|vhNation
    '               Register      NameOfLink = Nn23, Vn (geb Nn1, Club/Nation); Vn Nn123 ...
    '               Leute         NameOfLink = T5NameOfPersonFolder; NameOfFolder = Nn2, Vn (geb
    '               ClubsNations  NameOfLink = NnVn|vh [no Club, ..]
    '               Faces         NameOfLink = T5NameOfPersonFolder
    'Action         ändert ggf. den LinkName, je nach FromFolder, in welchem der Link liegt;
    '               z. B. "May, Lea" statt "May, Lea (TV Abc)" in einem ClubFolder;
    '               erkennt einen vh-LinkName
    
    'Vorbereitung
        Dim N1$, N2$, N3$, N4$, NnVn$, v$, c%, zT5%
        v = vbCrLf
        NnVn = Get_NnVn_Give_VnNn(VnNn)
        N1 = NameOfPath(PathOfLink) 'mit .lnk   'IST -Name
        N2 = T5NameOfPersonFolder   'Name des PersonenOrdners in 'Leute'
       'N3 = nur vhNnVn [no Club, no Nation]
       'N4 = vhNnVn(Club|Nation)
    'N1 könnte ein zulässiger vh-Name sein
        zT5 = T5_Get_T5RowNrOfOneVnNn(VnNn)
        If zT5 > 0 Then
            With Sheets("T5")
            If Len(.Cells(zT5, 14)) > 1 Then                'Eintrag vh existiert
                N3 = NnVn: N4 = N3
                If .Cells(zT5, 9) = "D" Then
                    If .Cells(zT5, 7) <> "" Then N4 = N3 + " (" + .Cells(zT5, 7) + ")"
                Else
                    N4 = N3 + " (" + .Cells(zT5, 9) + ")"
                End If
            End If
            End With
        Else
            Stop
        End If
    'Action
        Select Case FromFolder
            Case "Events"
                If N4 = "" Then N2 = N2 + ".lnk" Else N2 = N4 + ".lnk"
            Case "Faces"
                N2 = N2 + ".lnk"
            Case "ClubsNations"
                N2 = Get_LinkName_OfOneLinkInOneSubFolderOfClubsNations(PathOfLink, zT5, NnVn, N1, N2, N3, N4)
'            Case "Leute"
'                Stop
'            Case "Register"
'                Stop
        End Select
    'Rename
        If N1 <> N2 Then
            'Stop: OpenFolder Get_PathOfParentFolder(PathOfLink): Stop
            RenameFile PathOfLink, Replace(PathOfLink, N1, N2)
            Report = Report + "Folder Nr " + CStr(i) + ", [" + FromFolder + "] --> [Leute]" + v _
            + "  PathOfLink:   " + PathOfLink + v _
            + "  NameOfLink was changed: [" + N1 + "] --> [" + N2 + "]" + v
        End If
        'show REPORT
End Sub

Function Get_LinkName_OfOneLinkInOneSubFolderOfClubsNations(PathOfLink$, zT5%, NnVn$, N1$, N2$, N3$, N4$)
    'Called from    Check_NameOfLink
    'N1             = NameOfPath(PathOfLink) 'mit .lnk   'IST -Name
    'N2             = T5NameOfPersonFolder   'Name des PersonenOrdners in 'Leute'
    'N3             = vhNnVn ohne Club|Nation
    'N4             = vhNnVn mit  Club|Nation
    'SubFolder      = Club:   N2 = NnVn
    '               = LTV:    N2 = NnVnClub
    '               = Nation: N2 = NnVnClub
    
    'Vorbereitung
        Dim Club$, A() As String
        A = Split(PathOfLink, "\"): With Sheets("T5")
    'In welchem der Subfolder liegt der Link?
        'Club   F:\Archiv TR\Archiv Trampolin\ClubsNations\D\Baden\TV Abc\May, Lea.lnk
        'LTV    F:\Archiv TR\Archiv Trampolin\ClubsNations\D\Baden\May, Lea (TV Abc).lnk
        'Nation F:\Archiv TR\Archiv Trampolin\ClubsNations\D\May, Lea (TV Abc).lnk
        'Nation F:\Archiv TR\Archiv Trampolin\ClubsNations\AUS\May, Lea.lnk
        If PathOfLink Like "*\D\*" Then
            If UBound(A) = 7 Then 'Club
                If N3 = "" Then N2 = NnVn + ".lnk" Else N2 = N3 + ".lnk"
            ElseIf UBound(A) = 6 Then 'LTV
                If .Cells(zT5, 7) <> "" Then Club = " (" + .Cells(zT5, 7) + ")"
                If N4 = "" Then N2 = NnVn + Club + ".lnk" Else N2 = N4 + ".lnk"
            ElseIf UBound(A) = 5 Then 'Nation
                If .Cells(zT5, 7) <> "" Then Club = " (" + .Cells(zT5, 7) + ")"
                If N4 = "" Then N2 = NnVn + Club + ".lnk" Else N2 = N4 + ".lnk"
            End If
        Else
            If N3 = "" Then N2 = NnVn + ".lnk" Else N2 = N3 + ".lnk"
        End If
    'Finals
        Get_LinkName_OfOneLinkInOneSubFolderOfClubsNations = N2: End With
End Function

Function Base64EncodeString(ByVal text As String) As String
    'Wandelt Text in das von PowerShell benötigte UTF-16LE Base64 um
        Dim arrBytes() As Byte, xmlDoc As Object, xmlNode As Object
    'PowerShell erwartet UTF-16LE (Unicode in VBA)
        arrBytes = text
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")
    xmlNode.DataType = "bin.base64"
    xmlNode.nodeTypedValue = arrBytes
    'Das Ergebnis bereinigen (Zeilenumbrüche entfernen)
        Base64EncodeString = Replace(xmlNode.text, vbLf, "")
        Base64EncodeString = Replace(Base64EncodeString, vbCr, "")
End Function




