Attribute VB_Name = "M_Icon"
Option Explicit 'M_Icon

Sub zzz_M_Icon()
    
    showProcs "helper"
    
    EE 1: Beep
End Sub

Sub Copy_IconInUseJpgs_ToPersonFolder()
    Dim N$, p1$, p2$, pp$, s$, i%, Arr1() As String
    DoArrc
    s = Get_AllFileNames_Like_OfOneFolder(ArrC(1) + "\ico\Face_Jpgs", "*.jpg")
    'show s
    Arr1 = Split(s, vbCrLf)
    For i = 0 To UBound(Arr1)
        N = Arr1(i)             'Name of jpg-file
        pp = ArrC(4) + "\" + Replace(N, ".jpg", "") 'Path of PersonFolder
        p1 = ArrC(1) + "\ico\Face_Jpgs\" + N
        p2 = pp + "\zz_icon.jpg"
        If Not FileExists(p2) Then
            CopyFile p1, p2
        End If
    Next
    Beep
End Sub

Sub Search_ForPossible_zzjpg()
    'Called from    [once]
    'Vorbereitung
        Dim F$, p$, s$, i%, j%, Arr1() As String, Arr2() As String
        DoArrc
    'All PersonFolder
        s = Get_Paths_ofAllSubfolders_OneLevel(ArrC(4))
    Arr1 = Split(s, vbCrLf)
    For i = 0 To UBound(Arr1)
        p = Arr1(i)             'Path of one PersonFolder
        If p Like "*zzico*" Then GoTo GoNext
        If FileExists(p + "\zz_icon.jpg") Then GoTo GoNext
        If FileExists(p + "\zz.jpg") Then GoTo GoNext
        If Not (FileExists(p + "\zz_w.jpg") Or FileExists(p + "\zz_m.jpg")) Then Stop
        s = Get_FileNames_Like_NotLike_insideOneFolder(p, "*.jpg", "* .*")
        Arr2 = Split(s, vbCrLf): s = ""
        For j = 0 To UBound(Arr2)
            F = Arr2(j)      'one FileName
            If Not (F Like "*zz*" Or F Like "*°*" Or F Like "*Presse*") Then s = s + F + vbCrLf
        Next
        If s <> "" Then
            OpenFolder p
            Stop
            CloseFolder p
        End If
GoNext:
    Next
    Beep

End Sub

Sub Create_VbsLink_OfAll_zzjpg()
    Dim p$, p2$, s$, v$, c%, i%, Arr1() As String
    v = vbCrLf: DoArrc
    s = Get_Paths_ofAllSubfolders_OneLevel(ArrC(4))
    'show s
    Arr1 = Split(s, v)
    s = ""
    For i = 0 To UBound(Arr1)
        p = Arr1(i)         'Path of one PersonFolder
        p2 = p + "\zz.jpg"  'Path of zz.jpg in this PersonFolder
        If FileExists(p2) Then
            Create_JpgIconLink_OfOnePerson_zzJpg p
            c = c + 1
            s = s + getNameOfPath(p) + v
        End If
    Next
    s = "Proc 'Create_VbsLink_OfAll_zzjpg': " + CStr(c) + " new Icons/VbsLinks created" + v + v + s
    Call Beep:    show s
End Sub

Sub Create_JpgIconLink_OfOnePerson_zzJpg(NameOrPathOfPersonFolder$)
    'Called from    T6_Create_PersonLinkIcon_Of_zzJpg_in_OpenPersonFolder
    'Person.jpg     sollte im PersonFolder als "zz.jpg" bereits vorhanden sein
    '               bleibt letztlich als 'zz_icon.jpg' im PersonFolder
    'IconFolder     = Ordner "...\Leute\zzico"
    'FacesFolder    = Ordner "...\Archiv Trampolin 1900-1999\Faces"
    'icoJpgsFolder  = Ordner "...\Archiv Trampolin prog\ico\Face_Jpgs"
    'Action         legt (jpg|ico|lnk) der Person in (IconFolder|FacesFolder|icoJpgsFolder)
    
    'Vorbereitung
        Dim E$, JumpTo$, NameOfIcon$, NameOfPersonfolder$, PathOfIconFolder$, PathOfPersonfolder$
        Dim PathOfIcon$, PathOfJpg$, PathOfLink$, PathOfFolder_KillOneIcoJpg$
        DoArrc
    'Get PathOfPersonFolder, NameOfPersonFolder
        Get_NameAndPath_OfNameOrPath_OfFolder PathOfPersonfolder, NameOfPersonfolder, NameOrPathOfPersonFolder
    'Exit
        If Not FileExists(PathOfPersonfolder + "\zz.jpg") Then Exit Sub
    'Paths
        PathOfIconFolder = ArrC(4) + "\zzico\"
        NameOfIcon = Replace(NameOfPersonfolder, ",", "§") + ".ico" 'CreateShortCut mag kein Komma
        PathOfIcon = PathOfIconFolder & NameOfIcon
        PathOfLink = ArrC(2) + "\Faces\" + NameOfPersonfolder + ".lnk"
        PathOfJpg = ArrC(1) + "\ico\Face_Jpgs\" + NameOfPersonfolder + ".jpg"
    'Create icon
        CreateOnePersonIcon_via_zzJpg NameOfPersonfolder 'PersonFolder has zz.jpg
        Do While Not FileExists(PathOfPersonfolder + "\zz.ico"): DoEvents: Loop
    'Status
        'zz.ico wurde in PersonFolder erstellt; liegt neben zz.jpg
    'Icon in \Leute\zzico\ verschieben
        DeleteFile PathOfIcon
        RenameFile PathOfPersonfolder + "\zz.ico", PathOfIcon
    'Jpg in \ico\Face_Jpgs\ kopieren
        DeleteFile PathOfJpg
        CopyFile PathOfPersonfolder + "\zz.jpg", PathOfJpg
    'Jpg in \ico\Face_Jpgs\ umbenennen
        DeleteFile PathOfPersonfolder + "\zz_icon.jpg"
        DeleteFile PathOfPersonfolder + "\zz_w.jpg"
        DeleteFile PathOfPersonfolder + "\zz_m.jpg"
        RenameFile PathOfPersonfolder + "\zz.jpg", PathOfPersonfolder + "\zz_icon.jpg"
    'VbsLink in \Faces\ erzeugen
        JumpTo = PathOfPersonfolder
        Do While Not FileExists(PathOfIcon): DoEvents: Loop
        CreateShortCutWithIcon PathOfLink, PathOfIcon, JumpTo
    'Status
        'Ein Icon (z. B. "May, Lea (TV Bonn).ico") in Standard-Darstellung
        '(Kopie von z. B. "zz_w.ico") wurde gerade durch ein Bild der Person ersetzt,
        'jedoch wird (z. B. bei einem Personen-Link in einem Event-Ordner)
        'weiterhin das Standard-Icon dargestellt (kein Icon-Cache-Update);
        'nach dem nächsten PC-Neustart ist der Cache aktualisiert, das neue Bild wird angezeigt;
        'soll das neue Bild sofort dargestellt werden, run the following command lines as admin:
        '   taskkill /f /im explorer.exe
        '   del /f /s /q %localappdata%\Microsoft\Windows\Explorer\iconcache*
        '   Start explorer.exe
End Sub

Sub Update_IconCache()
    Dim cmd1$, cmd2$, cmd3$, r%
    cmd1 = "taskkill /f /im explorer.exe"
    cmd2 = "del /f /s /q %localappdata%\Microsoft\Windows\Explorer\iconcache*"
    cmd3 = "Start explorer.exe"
    
    r = ShellAndWait(cmd1, 0, vbHide, PromptUser)
    r = ShellAndWait(cmd2, 0, vbHide, PromptUser)
    r = ShellAndWait(cmd3, 0, vbHide, PromptUser)
End Sub

Sub CreateOnePersonIcon_via_zzJpg(NameOfPersonfolder$)
    'Called from    xxx
    'zz.jpg         = JpgOfPerson; liegt im PersonFolder
    'Action         erstellt zz.ico neben zz.jpg
    
    'Vorbereitung
        Dim p1$, p2$, p3$, PathOfPersonfolder$, qq$, shellCmd$, shResult%
        qq = Chr(34)
        PathOfPersonfolder = ArrC(4) + "\" + NameOfPersonfolder
        'If Not FileExists(PathOfPersonFolder + "\zz.jpg") Then Exit Sub
    'Pfade
        p1 = qq + "C:\Program Files\IrfanView\i_view64.exe" + qq + " "
        p2 = qq + PathOfPersonfolder + "\zz.jpg" + qq
        p3 = qq + PathOfPersonfolder + "\zz.ico" + qq
    'Icon erzeugen --> Ordner \test icon
        'cmd (IrfanView)
        shellCmd = p1 + p2 + " /advancedbatch /convert=" + p3
        'show shellCmd
        'Call Shell(shellCmd)
        shResult = ShellAndWait(shellCmd, 0, vbHide, PromptUser)
End Sub

Sub RefreshOpenExplorerWindow_TEST()
    'Called from    [none]
    Dim w As Object, p$
    Set w = CreateObject("Wscript.Shell")
    'p = "F:\Archiv TR\Archiv Trampolin\Events\1972-09-23 WM07 Stuttgart"
    p = "F:\Archiv TR\Archiv Trampolin\Leute\Schwebel, Robert (TV Gernsbach)"
    w.AppActivate p
    w.SendKeys "{F5}"
End Sub

Sub ChangeIconOfLink(PathOfLink$, PathOfIcon$)
    Dim wshShell As Object, L As Object
    Set wshShell = CreateObject("WScript.Shell")
    Set L = wshShell.CreateShortcut(PathOfLink)
    L.IconLocation = PathOfIcon
    L.Save
End Sub

Function GetIconLocationOfLink(PathOfLink$) As String
    Dim s$, wshShell As Object, L As Object
    'PathOfLink = "F:\Archiv TR\Archiv Trampolin\Events\1972-09-23 WM07 Stuttgart\Anderson, Robert (GB) - Copy.lnk"
    Set wshShell = CreateObject("WScript.Shell")
    Set L = wshShell.CreateShortcut(PathOfLink)
    GetIconLocationOfLink = L.IconLocation
End Function

Sub CreateShortCutWithIcon(PathOfLink$, PathOfIcon$, JumpTo$)
    'Called from    xxx
    'PathOfLink     = "...\Archiv Trampolin\Faces\May, Lea (TV Bonn).lnk"
    'Action         erstellt einen Link am Ort PathOfLink, der zu JumpTo springt;
    '               der Link wird mit dem Bild PathOfIcon angezeigt
    'Icon bereits im Ordner "...\Leute\zzico"?
        If Not FileExists(PathOfIcon) Then: Beep: show PathOfIcon: Stop
    DeleteFile PathOfLink
    Dim wshShell As Object, L As Object
    Set wshShell = CreateObject("WScript.Shell")
    Set L = wshShell.CreateShortcut(PathOfLink)
    L.targetPath = JumpTo
    L.WindowStyle = 1
    L.IconLocation = PathOfIcon  'for .ico file
    L.Description = "VbsLink"
    L.Save
End Sub

Sub Create_IconOfOneJpg_TEST()
    Dim p1$, p2$, p3$
    p1 = ArrC(1) + "\ico\WheatSquares\zzz_USA.jpg"
    p2 = Replace(p1, ".jpg", ".ico")
    p3 = ArrC(4) + "\zzico\zzy_LTV.ico"
    Create_IconOfOneJpg p1, p2
    CopyFile p2, p3
End Sub

Sub Create_IconOfOneJpg(PathOfJpg$, PathOfIcon$)
    'Vorbereitung
        Dim pIrf$, pIco$, pJpg$, qq$, sCmd$, shResult%
        qq = Chr(34)
    'Pfade
        pIrf = qq + "C:\Program Files\IrfanView\i_view64.exe" + qq + " "
        pJpg = qq + PathOfJpg + qq
        pIco = qq + PathOfIcon + qq
    'Icon erzeugen --> Ordner \test icon
        'cmd (IrfanView)
        sCmd = pIrf + pJpg + " /advancedbatch /convert=" + pIco
        'Show sCmd
        'Call Shell(sCmd)
        shResult = ShellAndWait(sCmd, 0, vbHide, PromptUser)
End Sub

Sub CreateIcons_OfAllJpgs_InOneFolder()
    'Vorbereitung
        Dim pIrf$, pIco$, pJpg$, qq$, shellCmd$
        qq = Chr(34)
    'Pfade
        pIrf = qq + "C:\Program Files\IrfanView\i_view64.exe" + qq + " "
        pJpg = qq + ArrC(1) + "\squares\*.jpg" + qq
        pIco = qq + ArrC(4) + "\squares\*.ico" + qq
    'Icon erzeugen --> Ordner \test icon
        'cmd (IrfanView)
        shellCmd = pIrf + pJpg + " /advancedbatch /convert=" + pIco
        'Show shellCmd
        Call shell(shellCmd)
End Sub

Sub Get_NameAndPath_OfNameOrPath_OfFolder(PathFo$, NameFo$, NameOrPath$)
        If NameOrPath Like "*:\*" Then
            PathFo = NameOrPath: NameFo = getNameOfPath(PathFo)
        Else
            NameFo = NameOrPath: PathFo = ArrC(4) + "\" + NameFo
        End If
End Sub

Sub IsClubLink_TEST()
    show IsClubLink("F:\Archiv TR\Archiv Trampolin\Leute\Abdel-Hadi, Samir (SuT Hemsbach)\- SuT Hemsbach.lnk")
End Sub

Function IsEventLink(PathOfOneLinkFile$) As Boolean
    If NameOfPath(PathOfOneLinkFile) Like "######## *" Then IsEventLink = True
End Function

Function IsClubLink(PathOfOneLinkFile$) As Boolean
    If NameOfPath(PathOfOneLinkFile) Like "- [A-Z]*" Then IsClubLink = True
End Function

Function IsLtvLink(PathOfOneLinkFile$) As Boolean
    If NameOfPath(PathOfOneLinkFile) Like "-  [A-Z]*" Then IsLtvLink = True
End Function

Function IsNationLink(PathOfOneLinkFile$) As Boolean
    If NameOfPath(PathOfOneLinkFile) Like "-   [A-Z]*" Then IsNationLink = True
End Function

Function IsPersLink(PathOfOneLinkFile$) As Boolean
    Dim p: p = NameOfPath(PathOfOneLinkFile)
    If p Like "*, *" Then IsPersLink = True: Exit Function
    If IsEventLink(p) Then IsPersLink = False: Exit Function
    If IsNationLink(p) Then IsPersLink = False: Exit Function
    If IsLtvLink(p) Then IsPersLink = False: Exit Function
    If IsClubLink(p) Then IsPersLink = False: Exit Function
    IsPersLink = True
End Function

Sub Separate_SomeLinkNames(clubs$, SomeLinkNames$, LnkPerson$, LnkEvent$, LnkClub$, LnkLTV$, LnkNation$, LnkVideo$)
    'Called from    Change_AllLinksOfOneFolder_ToVbsLinks
    'Clubs          = "|ASV Nürnberg|...|Zoetermeer|"
    'SomeLinkNames  = Zeilen wie "Froehlich, Ron (SA).lnk"
    'LnkPerson,...  LnkEvent, LnkClub, LnkLTV, LnkNation (werden hier gefüllt)
    'Action         xxx
    
    'Vorbereitung
        Dim M1$, M2$, N$, vID$, i&, A() As String
        vID = "|.avi|.divx|.mkv|.mp4|.mpg|.mov|.vob|.wmv|"

    'Split SomeLinkNames to LnkPerson, LnkEvent, LnkClub, LnkLTV, LnkNation
        A = Split(SomeLinkNames, vbCrLf)
        
        For i = 0 To UBound(A)
            N = A(i)                                'one LinkName
            M1 = Replace(N, ".lnk", "")             'LinkName ohne .lnk
            M2 = LCase("*|" + Right(M1, 4) + "|*")  'LikeString "*|.mp4|*"
            If N Like "-   [A-Z]*" Then
                LnkNation = LnkNation + N + "|"
            ElseIf N Like "-  [A-Z]*" Then LnkLTV = LnkLTV + N + "|"
            ElseIf vID Like M2 Then
                LnkVideo = LnkVideo + N + "|"
            ElseIf N Like "*, *" Then LnkPerson = LnkPerson + N + "|"
            
            ElseIf N Like "######## ! *" Then LnkEvent = LnkEvent + N + "|"
            
            ElseIf N Like "- [A-Z]*" Then LnkClub = LnkClub + N + "|"
            ElseIf clubs <> "" Then
                'In z. B. N = "...\Events\1980 Liga\1980 BadLiga\SSV Freiburg.lnk
                'startet der Name nicht mit "- " (wegen Reihenfolge)
                If clubs Like "*|" + Replace(N, ".lnk", "") + "|*" Then LnkClub = LnkClub + N + "|"
            Else: LnkPerson = LnkPerson + N + "|"
            End If
        Next
End Sub

Sub Separate_AllLinks_AllFolders()
    'Called from    [once]
    'Action         Check separation rules
    
    'Vorbereitung
        Dim Club$, Ev$, LTV$, Nat$, Pers$, VnNn$, s$, v$
        Dim cClub&, cEv&, cLTV&, cNat&, cPers&, cVnNn&, i&, A() As String
        With Sheets("T6"): DoArrc: v = vbCrLf
        's = Get_FileNames_Like_insideSourceFolderAndSubFolders(ArrC(2), "*.lnk")
        'writeStringToFile ArrC(2) + "\all_Links.txt", s
        s = ReadFile(ArrC(2) + "\all_Links.txt")
        s = Delete_EmptyEndRowsInString(s)
        'show s
    'split Links to LnkPerson, LnkVnNn, LnkEvent, LnkClub, LnkLTV, LnkNation
        A = Split(s, v): .[AO13] = UBound(A) + 1
        For i = 0 To UBound(A)
            If A(i) Like "-   [A-Z]*" Then
                Nat = Nat + A(i) + "|":                                 cNat = cNat + 1
            ElseIf A(i) Like "-  [A-Z]*" Then LTV = LTV + A(i) + "|":   cLTV = cLTV + 1
            ElseIf A(i) Like "*, *" Then Pers = Pers + A(i) + "|":      cPers = cPers + 1
            ElseIf A(i) Like "######## *" Then Ev = Ev + A(i) + "|":    cEv = cEv + 1
            ElseIf A(i) Like "- [A-Z]*" Then Club = Club + A(i) + "|":  cClub = cClub + 1
            Else: VnNn = VnNn + A(i) + "|":                             cVnNn = cVnNn + 1
            
            End If
            If i Mod 100 = 0 Then .[AO3] = UBound(A) - i + 1: .[AO5] = cPers: .[AO6] = cEv: _
                .[AO7] = cClub: .[AO8] = cLTV: .[AO9] = cNat: .[AO10] = cVnNn: _
                .[AO12] = cNat + cLTV + cPers + cEv + cClub + cVnNn
            DoEvents
        Next
    'genaue Werte
        .[AO3] = UBound(A) - i: .[AO5] = cPers: .[AO6] = cEv
        .[AO7] = cClub: .[AO8] = cLTV: .[AO9] = cNat: .[AO10] = cVnNn
        .[AO12] = cNat + cLTV + cPers + cEv + cClub + cVnNn
    show "Pers" + v + Pers
    show "Ev" + v + Ev
    show "Club" + v + Club
    show "LTV" + v + LTV
    show "Nat" + v + Nat
    show "VnNn" + v + VnNn
    Call Beep: Stop
    
    End With
End Sub

Function Get_AllT5Clubs() As String
    'Called from    Check_OneFolder
    'Action         liefert "|TV Bonn|Kiev|..."
    
    'Vorbereitung
        Dim s$, i%, Club() As String
        T5_Load_OneColumn Club, 7
        s = "|"
    'Action
        For i = 1 To UBound(Club)
            If Not s Like "*|" + Trim(Club(i)) + "|*" Then s = s + Trim(Club(i)) + "|"
        Next
        Get_AllT5Clubs = s
End Function

Sub RefreshAllIcons_RestartExplorerExe()
    Dim sCmd$, shResult%
    sCmd = "taskkill /F /IM explorer.exe & start explorer.exe"
    shResult = ShellAndWait(sCmd, 0, vbHide, PromptUser)
    'If explorer.exe did not start:
        Dim s: Ms = shell("cmd.exe", 1) 'opens CmdWindow to stay open
        'write into CmdWindow: start explorer.exe
End Sub

Sub RefreshShortcutIcon_TEST()
    Dim p$
    p = "F:\Archiv TR\Archiv Trampolin\Faces\Rydzewski, Andreas (TV Unterbach).lnk"
    RefreshShortcutIcon p
End Sub
Sub RefreshShortcutIcon(strFilePath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(strFilePath) Then
        ' Update the "Last Modified" date to right now
        fso.GetFile(strFilePath).DateLastModified = Now
    End If
End Sub

Sub Kill_LinkDoublesm_OfOneFolder(PathOfFolder$)
    Dim N1$, N2$, N3$, p$, pi$, c%, i%, A() As String
    pi = Get_AllFileNames_Like_OfOneFolder(PathOfFolder, "*.lnk")
    'show pi
    A = Split(pi, vbCrLf): N3 = "|"
    For i = 0 To UBound(A)
        N1 = A(i) 'one Name                 'N1 = Bahr, Gabi (TuRa Duisburg).lnk
        N2 = Replace(N1, ".lnk", "")
        c = InStr(1, N2, " (")
        If c > 0 Then N2 = Left(N2, c - 1)  'N2 = Bahr, Gabi
        If N3 Like "*|" + N2 + "|*" Then
            'NnVn existiert bereits in N3; N1 ist ein LinkDouble
            p = PathOfFolder + "\" + N1
            If FileExists(p) Then
                DeleteFile p
                'REPORT = REPORT + "File deleted: " + PathOfLink + v
                'Exit Sub
            Else
                Stop
            End If
        Else
            N3 = N3 + N2 + "|"
        End If
    Next
End Sub


