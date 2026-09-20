Attribute VB_Name = "M_LinkVbs"
Option Explicit

'----------------- API-Deklaration
'Create_OneVbsLink - Icon-Refresh
    #If VBA7 Then
        Private Declare PtrSafe Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, ByVal dwItem1 As LongPtr, ByVal dwItem2 As LongPtr)
    #Else
        Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)
    #End If
'-----------------

Sub zzz_M_LinkVBS()
    showProcs "nation"
    'show Get_TextOfAllOpenNotepadWindows
    EE 1: Beep
End Sub


Sub Create_OneVbsLink_Person(pLinkWithPersName$)
    Dim L1_pLink$, L2_pJumpTo$, L3_pIcon$, LinkWasCreated As Boolean
    'Verrify PersonName-Personfolder-(vh ..)-ifExist
        Fill_L1_L2_L3_IfLinkIsPersLink pLinkWithPersName, L1_pLink, L2_pJumpTo, L3_pIcon
    'Create Link
        If Not (L1_pLink = "" Or L2_pJumpTo = "" Or L3_pIcon = "") Then
            '##############
            Create_OneVbsLink L1_pLink, L2_pJumpTo, L3_pIcon
            '##############
            LinkWasCreated = True
        End If
        If Not LinkWasCreated Then Exit Sub
    'War das ein Link mit Verein?
        If L1_pLink Like "*).lnk" Then
            'Lösche .lnk gleichen Namens (ohne Verein);
            'zum Namen des gerade erzeugten VbsLink existiert bereits ein Link (ohne Verein)
                L1_pLink = Left(L1_pLink, InStrRev(L1_pLink, " (") - 1) + ".lnk"
                DeleteFile L1_pLink 'L1_pLink wurde neu belegt
        End If
End Sub

Sub StartVBS_Delete_OlderVersionsOfLink(PathOfFolder$)
    'vbs parameter
    Dim vbsPath$, myFold$, shell As Object
    vbsPath = ArrC(6) + "\vbs\Delete_OlderVersionsOfLink.vbs" 'helpers
    Set shell = CreateObject("WScript.Shell")
    'Startet das VBScript und übergibt den Pfad in Anführungszeichen
        shell.Run Chr(34) & vbsPath & Chr(34) & " " & Chr(34) & PathOfFolder & Chr(34)
End Sub

Sub Change_LinkToPersonFolder_ToVbsLink(PersonFolderOrName$)
    'Called from    xxx
    'Status         Hier werden nur absolute "F:\"-Pfade verwendet,
    '               relativ werden sie in Create_OneVbsLink
    'Status         Hier geht es nur um PersonLinks
    
    'Vorbereitung
        Dim JumpTo$, L1_pLink$, L2_pJumpTo$, L3_pIcon$
        Dim PathOfLink$, PathOfOneEventFolder$, v$, Register As Boolean
        Dim i%, j%, w%, E() As String, L() As String
        v = vbCrLf: DoArrc
    'All SubFolders of one SourceFolder
        Load_ArrPathsOfAllFolders_Events E ': showArray E
        'Load_ArrPathsOfAllFolders_ClubsNations E
        'Load Register --> E     'hat keine Ordner; ist selbst der Ordner mit Links
            'ReDim E(0 To 0): E(0) = ArrC(5): Register = True
    'Test: perform only one event
        Dim A&: A = 1: Stop
        For i = LBound(E) To UBound(E)
            If E(i) Like "*1969-06-07*" Then A = i: Stop
        Next
    'UF4
        UF4.show vbModeless
   '################################
    For i = A To A 'LBound(E) To UBound(E)
   '################################
        PathOfOneEventFolder = E(i)
        'UF4
            UpdateCountdown i, UBound(E), NameOfPath(E(i)), "Change_LinkToPersonFolder_ToVbsLink"
        'Open
            If Register Then
                OpenFolder E(i)
                Change_ExplorerWindow_ToBigIcons PathOfOneEventFolder
            End If
        Load_AllFilePaths_LikeMyStringInFileName_OfOneFolder PathOfOneEventFolder, "*.lnk", L
        w = 0
        For j = 0 To UBound(L)
            PathOfLink = L(j)       'Path of one .lnk in Folder j of Events
            DoEvents
            'UF4
                UpdateCountdown j, UBound(L), NameOfPath(PathOfLink)
            'OpenFolder(PathOfOneEventFolder) and MarkFile
                If Not Register Then Mark_MyFile PathOfLink
                Move_FolderWindow_XYWH PathOfOneEventFolder, 55, 225, 664, 800: w = w + 1
                'If w = 0 Then Move_FolderWindow_XYWH PathOfOneEventFolder, 50, 225, 664, 800: w = w + 1
            'Verrify PersonName-Personfolder-(vh ..)-ifExist
                Fill_L1_L2_L3_IfLinkIsPersLink PathOfLink, L1_pLink, L2_pJumpTo, L3_pIcon
            'Create Link
                If Not (L1_pLink = "" Or L2_pJumpTo = "" Or L3_pIcon = "") Then
                    '##############
                    Create_OneVbsLink L1_pLink, L2_pJumpTo, L3_pIcon
                    '##############
                End If
            Sheets("T4").Cells(42, 17) = "File:       j = " + CStr(j) _
                + " of " + CStr(UBound(L)) + "   " + NameOfPath(L1_pLink)
'            'Show_InfoOfOneLink L1_pLink
            If Not Register Then Change_ExplorerWindow_ToBigIcons PathOfOneEventFolder
            'Mark_MyFile L1_pLink
            'If w = 1 Then
                Call Window_OneNameOnly: w = w + 1
                Move_FolderWindow_XYWH PathOfOneEventFolder, 55, 225, 664, 800
            'End If
        Next j
        DoEvents
        CloseFolder E(i)
    Next i
    Unload UF4
    Beep
End Sub

Sub Create_OneVbsLink(PathOfLinkFile$, PathOfLinkTarget$, PathOfLinkIcon$)
    'Called from    Change_LinkToPersonFolder_ToVbsLink
    'Create          SelfReferencingLinks
    
    'Vorbereitung
        Dim ShellApp As Object, L As Object, fso As Object
        Dim pVbs$, pJump$, driveLetter$
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ShellApp = CreateObject("WScript.Shell")
    'Laufwerksbuchstaben ermitteln (z.B. "E:")
        driveLetter = fso.GetDriveName(PathOfLinkFile)
    'Pfade vorbereiten
        pVbs = Mid(ArrC(4), 3) + "\zzico\zz.vbs"
        '    = "\Archiv Trampolin 1900-1999\Leute\zzico\zz.vbs"
        pJump = Mid(PathOfLinkTarget, 4)
        '     = "\Archiv Trampolin 1900-1999\Leute\Nissen, Annie (USA)"
    'Alte Link-Datei löschen, falls sie existiert!
        'Das bricht die fehlerhafte "Dateiordner"-Kategorisierung auf.
        If fso.FileExists(PathOfLinkFile) Then fso.DeleteFile PathOfLinkFile, True
    'Jetzt den Link komplett frisch erstellen
        Set L = ShellApp.CreateShortcut(PathOfLinkFile)
    'Zieltyp sicher von "Dateiordner" auf "Anwendung" umstellen
        ' Wir übergeben wscript.exe UND das Skript mit absolutem Pfad in den TargetPath.
        ' Windows prüft beim Speichern: "Gibt es diese Datei?" -> Ja (weil E:\Archiv...\zz.vbs existiert).
        ' Ergebnis: Der Zieltyp springt GARANTIERT auf "Anwendung".
        L.targetPath = "C:\WINDOWS\system32\wscript.exe"
    'Für die Portabilität übergeben wir im Argument NUR den relativen Zielordner
        L.arguments = """" & driveLetter & pVbs & """ """ & pJump & """"
        L.IconLocation = PathOfLinkIcon & ",0"
        L.WorkingDirectory = "C:\WINDOWS\system32"
        L.Description = "VbsLink"
        L.Save
    'Bereinigen
        Set L = Nothing: Set fso = Nothing
    'System-Refresh erzwingen
        SHChangeNotify &H8000000, &H0, 0, 0
End Sub

Sub Fill_L1_L2_L3_IfLinkIsPersLink(PathOfLink$, L1_pLink$, L2_pJumpTo$, L3_pIcon$)
    'Called from    Change_LinkToPersonFolder_ToVbsLink, Create_OneVbsLink_Person
    'L1_pLink       trägt den bisherigen Pfad des Links; wird ggf. überschrieben
    'L2_pJumpTo     und L3_pIcon werden hier gefüllt
    
    'Vorbereitung
        Dim NnVn$, VnNn$, pLink$, pPersFold$, txt$, v$, c%
        v = vbCrLf: pLink = PathOfLink: L2_pJumpTo = "": L3_pIcon = ""
    'NnVn VnNn
        txt = Replace(NameOfPath(PathOfLink), ".lnk", ""): c = InStr(1, txt, " (")
        If c > 0 Then txt = Left(txt, c - 1)            'Mai, Lea
        'Txt trägt jetzt den Text des NameOfLink ohne Klammer
        Extract_VnNn_NnVn_FromTxt VnNn, NnVn, txt
        If VnNn = "" Then Exit Sub
        'gültige NnVn und VnNn existieren jetzt
    'pPersFold
        pPersFold = T5_Get_PathOfPersonfolder_Give_NnVn(NnVn) 'T5: aktuelle PersonenDaten
        If pPersFold = "" Then Stop
        If Not FolderExists(pPersFold) Then
            show "Sub Fill_L1_L2_L3_IfLinkIsPersLink(PathOfLink$, L1_pLink$, L2_pJumpTo$, L3_pIcon$)" _
                + v + v + "PathOfLink: " + PathOfLink _
                + v + "Folder existiert nicht: " + pPersFold  ': Stop
        End If
            'Falls Stop: Reparieren! (händisch oder prog)
            'pPersFold = T5_Get_PathOfPersonfolder_Give_NnVn(NnVn)
    'Status: gültiger pPersFold existiert
        L2_pJumpTo = pPersFold
        L1_pLink = Check_L1_pLink(pLink, pPersFold, NnVn, VnNn)
        L3_pIcon = Check_L3_pIcon(pPersFold, NnVn)
End Sub

Function Check_L1_pLink(pLink$, pPersFold$, NnVn$, VnNn$) As String
    'Called from    Check_Link
    'pLink          trägt den bisherigen Pfad des Links; wird ggf. überschrieben
    'DisplayName    nur:   "Mai, Lea",  "Mai, Lea (Bonn)"  oder "Mai, Lea (Bern, CH)"
    '               nicht: "Mai, Lea (vh Jun Jul ..)"
    'NnVn           Nn, Vn des pLink; sollten in pPersFold vorkommen
    'ClubNation     = "", " (Bonn)" oder " (Bern, CH)"
    '               = Teil2 des DisplayName wird durch pPersFold fixiert
    '               vh-Teile werden nicht in den DisplayName übernommen
    'pPersFol       gültiger Pfad des PersonFolder (Sprungziel des Link) existiert
    '               |*\Mai, Lea|*\Mai, Lea (vh Jun Jul, Bern, CH)|
    'Action         Ermittlung von L1_pLink über vh- und NnVn-Check;
    '               ggf. den pLink in Events ändern
    
    'Vorbereitung
        Dim ClubNation$, k$, L1_pLink$, N$, NamePart$, Nn$, Vn$, c%
    'ClubNation-Teil ermitteln     = "", " (Bonn)" oder " (Bern, CH)"
        N = NameOfPath(pPersFold)
        c = InStr(1, N, " (")
        If c = 0 Then
            ClubNation = ""
        Else
            'Klammerinhalt K
            k = Mid(N, c + 2): k = Left(k, Len(k) - 1)
            If k Like "vh *" Then
                c = InStr(1, k, ",")
                ClubNation = " (" + Mid(k, c + 2) + ")"
            Else
                ClubNation = " (" + k + ")"
            End If
        End If
    'L1_pLink
        'Die Anordnung VnNn/NnVn im LinkNameDisplay soll erhalten bleiben;
        'der Ordner "Register" enthält beide Anordnungen
        If pLink Like "*\" + NnVn + "*" Then
            L1_pLink = Get_PathOfParentFolder(pLink) + "\" + NnVn + ClubNation + ".lnk"
        ElseIf pLink Like "*\" + VnNn + "*" Then
            L1_pLink = Get_PathOfParentFolder(pLink) + "\" + VnNn + ClubNation + ".lnk"
        Else
            Stop
        End If
    'Change NameOfLink in EventFolder
        If pLink <> L1_pLink Then RenameFile pLink, L1_pLink
    'Finals
        Check_L1_pLink = L1_pLink
End Function

Function Check_L3_pIcon(pPersFold$, NnVn$) As String
    Dim mw$, N$, PathOfIcon$
    N = NameOfPath(pPersFold): DoArrc
    PathOfIcon = ArrC(4) + "\zzico\" + Replace(N, ",", "§") + ".ico"
    If Not FileExists(PathOfIcon) Then
        mw = T5_Get_mw_Give_NnVn(NnVn)
        If mw = "w" Then
            CopyFile ArrC(4) + "\zzico\zzw.ico", PathOfIcon
        ElseIf mw = "m" Then
            CopyFile ArrC(4) + "\zzico\zzm.ico", PathOfIcon
        Else: Stop
        End If
    End If
    Check_L3_pIcon = PathOfIcon
End Function

Sub Extract_VnNn_NnVn_FromTxt(VnNn$, NnVn$, txt$)
    'Called from    Check_Link
    
    'Vorbereitung
        Dim c%, D%
    'Ist txt ein NnVn oder VnNn?
        c = InStr(1, txt, ", ")
        If c = 0 Then
            'Txt ist kein NnVn
            D = InStr(1, txt, " ")
            If D = 0 Then
                Exit Sub 'NnVn und VnNn bleiben ""
            Else
                'Txt könnte ein VnNn sein
                If VnNn_ExistsIn_T5(txt) Then
                    VnNn = txt: NnVn = Get_NnVn_Give_VnNn(txt)
                Else
                    Exit Sub 'NnVn und VnNn bleiben ""
                End If
            End If
        Else
            'Txt könnte ein NnVn sein
            If NnVn_ExistsIn_T5(txt) Then
                NnVn = txt: VnNn = Get_VnNn_Give_NnVn(txt)
            Else
                Exit Sub 'NnVn und VnNn bleiben ""
            End If
        End If
End Sub

Sub Window_OneNameOnly()
    Dim objShellApp, objWindows, objWin, hwndExplorer, folderPath, dictSeenFolders
    Set dictSeenFolders = CreateObject("Scripting.Dictionary")
    Set objShellApp = CreateObject("Shell.Application")
    Set objWindows = objShellApp.Windows
    ' Wir laufen rückwärts durch die Fenster
    Dim i
    For i = objWindows.count - 1 To 0 Step -1
        Set objWin = objWindows.item(i)
        On Error Resume Next
        If InStr(1, objWin.FullName, "explorer.exe", 1) > 0 Then
            ' Wir nutzen DecodeURL, um das Pfad-Format absolut anzugleichen
            folderPath = DecodeURL(objWin.LocationURL)
            hwndExplorer = objWin.hwnd
            If folderPath <> "" And hwndExplorer <> 0 Then
                If dictSeenFolders.Exists(folderPath) Then
                    'TRICK: Wir nutzen deine funktionierende CloseFolderWindow-Logik!
                    'Statt des Pfads übergeben wir diesmal nichts, sondern rufen den API-Befehl
                    'direkt im Hintergrund auf. Das trennt die Verbindung zum anderen Fenster.
                    Dim cmd
                    cmd = "powershell -NoProfile -WindowStyle Hidden -Command """ & _
                          "Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class Win32 { [DllImport(\""user32.dll\"")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam); }'; [void][Win32]::PostMessage(" & hwndExplorer & ", 16, 0, 0)"""
                    ' WICHTIG: Asynchron starten (0, False), damit Windows Zeit hat zu atmen
                    CreateObject("WScript.Shell").Run cmd, 0, False
                    ' Sofort aufhören, damit nicht aus Versehen noch ein Fenster getroffen wird
                    Exit Sub
                Else
                    dictSeenFolders.Add folderPath, True
                End If
            End If
        End If
        On Error GoTo 0
    Next
End Sub

Sub Move_FolderWindow_XYWH(PathOfFolder, PosX, PosY, Breite, Hoehe)
    Dim objShellApp, objWindows, objWin, hwndExplorer, folderPath, cmd
    Dim flags, pX, pY, bR, hE
    Set objShellApp = CreateObject("Shell.Application")
    Set objWindows = objShellApp.Windows
    'Standard-Flags für die Windows-API (0x0040 = SWP_SHOWWINDOW)
        flags = &H40
    'Prüfen, ob Positionen übergeben wurden. Wenn nicht, aktuelle beibehalten (SWP_NOMOVE = 0x0002)
        If PosX = "" Or PosX = 0 Or PosY = "" Or PosY = 0 Then flags = flags Or &H2: pX = 0: pY = 0 Else pX = PosX: pY = PosY
    'Prüfen, ob Maße übergeben wurden. Wenn nicht, aktuelle beibehalten (SWP_NOSIZE = 0x0001)
        If Breite = "" Or Breite = 0 Or Hoehe = "" Or Hoehe = 0 Then flags = flags Or &H1: bR = 0: hE = 0 Else bR = Breite: hE = Hoehe
    'Fenster suchen und ansteuern
        For Each objWin In objWindows
            On Error Resume Next
            If InStr(1, objWin.FullName, "explorer.exe", 1) > 0 Then
                folderPath = DecodeURL(objWin.LocationURL): hwndExplorer = objWin.hwnd
                If folderPath = LCase(Trim(DecodeURL(PathOfFolder))) And hwndExplorer <> 0 Then
                    'Windows-API Aufruf zusammenbauen (Hex-Flags werden als Zahl übergeben)
                    cmd = "powershell -NoProfile -WindowStyle Hidden -Command """ & _
                          "Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class Win32 { [DllImport(\""user32.dll\"")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags); }'; " & _
                          "[void][Win32]::SetWindowPos(" & hwndExplorer & ", [IntPtr]::Zero, " & pX & ", " & pY & ", " & bR & ", " & hE & ", " & flags & ")"""
                    CreateObject("WScript.Shell").Run cmd, 0, False
                    Exit For
                End If
            End If
            On Error GoTo 0
        Next
End Sub

Sub Load_List_ClubLtvNation_FromT5(ByRef Arr() As String)
    'Called from    xxx
    'Status         Jeder Club kommt nur 1 x vor
    'Action         füllt Arr() '|ASV Nürnberg|Bayern|D|    '|Aalsmeer||NL|
    
    'Vorbereitung
        Dim B1$, B2$, B3$, B4$, List$, v$
        Dim i&, zLast&, r As Range, A()
        v = vbCrLf: DoArrc: With Sheets("T5")
    'Club-LTV-Nation aus T5 holen
        zLast = Get_T5_zLast: Set r = .Range(.Cells(8, 7), .Cells(zLast, 9))
        A = r.Value
        For i = 1 To UBound(A, 1)
            B1 = Trim(CStr(A(i, 1))): B2 = Trim(CStr(A(i, 2))): B3 = Trim(CStr(A(i, 3)))
            B4 = "|" + B1 + "|" + B2 + "|" + B3 + "|"
            If B1 <> "" And Not List Like "*" + B4 + "*" Then List = List + v + B4
        Next
        List = Mid(List, 3)
        Arr = Split(List, v):  QuickSort Arr
    'Finals
        End With
End Sub

Function HasVbsLinkDescription(PathOfLink$) As Boolean
    Dim L As Object, w As Object
    Set w = CreateObject("WScript.Shell")
    Set L = w.CreateShortcut(PathOfLink)
        If L.Description = "VbsLink" Then HasVbsLinkDescription = True
End Function

Sub Create_VbsLink_all_Club_to_D()
    'Called from    [none]
    'Action         xxx
    
    'Vorbereitung
        Dim Club$, ClubFolder$, LTV$, LTVFolder$, p$, i&, E() As String, L() As String
        Dim L1_pLink$, L2_pJumpTo$, L3_pIcon$
        Load_List_ClubLtvD_FromT5 L         '|ASV Nürnberg|Bayern|D|
    'D
        p = ArrC(11) + "\D"
        For i = 0 To UBound(L)
            E = Split(L(i), "|"): Club = E(1): LTV = E(2)
            ClubFolder = p + "\" + LTV + "\" + Club
            LTVFolder = p + "\" + LTV
            If Not FolderExists(LTVFolder) Then CreateFolder LTVFolder
            If Not FolderExists(ClubFolder) Then CreateFolder ClubFolder
            'Folders LTV, Club existieren
            'Anzeige einschalten?
                'Change_ExplorerWindow_ToBigIcons ClubFolder: Stop
            'Create Lnk in ClubFolder jumping to D-Folder
                L1_pLink = ClubFolder + "\" + "-   D.lnk"
                If Not HasVbsLinkDescription(L1_pLink) Then
                    L2_pJumpTo = p
                    L3_pIcon = ArrC(4) + "\zzico\" + "zzz_D.ico"
                    Create_OneVbsLink L1_pLink, L2_pJumpTo, L3_pIcon
                End If
            'Anzeige ausschalten?
                'Stop: CloseFolder ClubFolder
        Next
End Sub

Sub Create_VbsLink_all_Club_to_LTV()
    'Called from    [none]
    'Action         xxx
    
    'Vorbereitung
        Dim Club$, ClubFolder$, LTV$, LTVFolder$, p$, i&, E() As String, L() As String
        Dim L1_pLink$, L2_pJumpTo$, L3_pIcon$
        Load_List_ClubLtvD_FromT5 L         '|ASV Nürnberg|Bayern|D|
    'D
        p = ArrC(11) + "\D"
        For i = 0 To UBound(L)
            E = Split(L(i), "|"): Club = E(1): LTV = E(2)
            ClubFolder = p + "\" + LTV + "\" + Club
            LTVFolder = p + "\" + LTV
            If Not FolderExists(LTVFolder) Then CreateFolder LTVFolder
            If Not FolderExists(ClubFolder) Then CreateFolder ClubFolder
            'Folders LTV, Club existieren
            'Anzeige einschalten?
                'Change_ExplorerWindow_ToBigIcons ClubFolder: Stop
            'Create Lnk in ClubFolder jumping to D-Folder
                L1_pLink = ClubFolder + "\" + "-  " + LTV + ".lnk"
                If Not HasVbsLinkDescription(L1_pLink) Then
                    L2_pJumpTo = p + "\" + LTV
                    L3_pIcon = ArrC(4) + "\zzico\" + "zzy_LTV.ico"
                    Create_OneVbsLink L1_pLink, L2_pJumpTo, L3_pIcon
                End If
            'Anzeige ausschalten?
                'Stop: CloseFolder ClubFolder
        Next
End Sub

Sub Create_VbsLink_all_LTV_to_D()
    'Called from    [none]
    'Action         xxx
    
    'Vorbereitung
        Dim LTV$, LTVFolder$, DFolder$, i&, L() As String
        Dim L1_pLink$, L2_pJumpTo$, L3_pIcon$
    'Load
        Load_List_LTV_FromT5 L         'Baden|Bayern|...
    'LTV
        DFolder = ArrC(11) + "\D"
        For i = 0 To UBound(L)
            LTV = L(i): LTVFolder = DFolder + "\" + LTV
            If Not FolderExists(LTVFolder) Then CreateFolder LTVFolder
            'Folder LTV existiert
            'Anzeige einschalten?
                'Change_ExplorerWindow_ToBigIcons LTVFolder: Stop
            'Create Lnk in LTVFolder jumping to D-Folder
                L1_pLink = DFolder + "\" + LTV + "\-   D.lnk"
                If Not HasVbsLinkDescription(L1_pLink) Then
                    L2_pJumpTo = DFolder
                    L3_pIcon = ArrC(4) + "\zzico\" + "zzz_D.ico"
                    Create_OneVbsLink L1_pLink, L2_pJumpTo, L3_pIcon
                End If
            'Anzeige ausschalten?
                'Stop: CloseFolder LTVFolder
        Next
End Sub

Sub Load_List_LTV_FromT5(ByRef Arr() As String)
    'Called from    xxx
    'Status         Jeder LTV kommt nur 1 x vor
    'Action         füllt Arr() 'Baden|Bayern|...
    
    'Vorbereitung
        Dim LTV$, Nation$, B4$, List$, v$
        Dim i&, zLast&, r As Range, A()
        v = vbCrLf: DoArrc: With Sheets("T5")
    'LTV|Nation aus T5 holen
        zLast = Get_T5_zLast: Set r = .Range(.Cells(8, 8), .Cells(zLast, 9))
        A = r.Value
        For i = 1 To UBound(A, 1)
            LTV = Trim(CStr(A(i, 1))): Nation = Trim(CStr(A(i, 2)))
            B4 = "|" + LTV + "|" + Nation + "|"
            If LTV <> "" And Nation = "D" And Not List Like "*" + B4 + "*" Then List = List + v + B4
        Next
        List = Mid(List, 3): List = Replace(List, "|D|", ""): List = Replace(List, "|", "")
        Arr = Split(List, v): QuickSort Arr
    'Finals
        End With
End Sub

Sub Load_List_ClubLtvD_FromT5(ByRef Arr() As String)
    'Called from    xxx
    'Status         Jeder Club kommt nur 1 x vor
    'Action         füllt Arr() '|ASV Nürnberg|Bayern|D|    '|SV Moosbrunn|Baden|D|
    
    'Vorbereitung
        Dim B1$, B2$, B3$, B4$, List$, v$
        Dim i&, zLast&, r As Range, A()
        v = vbCrLf: DoArrc: With Sheets("T5")
    'Club-LTV-D aus T5 holen
        zLast = Get_T5_zLast: Set r = .Range(.Cells(8, 7), .Cells(zLast, 9))
        A = r.Value
        For i = 1 To UBound(A, 1)
            B1 = Trim(CStr(A(i, 1))): B2 = Trim(CStr(A(i, 2))): B3 = Trim(CStr(A(i, 3)))
            B4 = "|" + B1 + "|" + B2 + "|" + B3 + "|"
            If B1 <> "" And B3 = "D" And Not List Like "*" + B4 + "*" Then List = List + v + B4
        Next
        List = Mid(List, 3): Arr = Split(List, v): QuickSort Arr
    'Finals
        End With
End Sub


