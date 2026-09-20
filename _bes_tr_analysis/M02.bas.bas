Attribute VB_Name = "M02"
Option Explicit 'M02

'GetCursorPosDemo
    Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    'Create custom variable that holds two integers
        Type POINTAPI
            Xcoord As Long
            Ycoord As Long
        End Type

Sub zzz_M02()
    'showProcs "vnnn"
    
    'RenameModule "TR", "M02"
    'Update_Links_inside_FolderLeute
    'ActiveWorkbook.Sheets("T2").Cells(28, 3).Font.Color = RGB(150, 150, 150)
    'ActiveWorkbook.Sheets("T2").Column(3).FontSize = 16
    'T1_ShowLogBuch
    EE 1: Beep
End Sub

Sub GetCursorPosDemo()
    Dim llCoord As POINTAPI
    Dim rng As Range
    
    Do
    'Get the cursor positions
        GetCursorPos llCoord
        'Display the cursor position coordinates
        'MsgBox "X Position: " & llCoord.Xcoord & vbNewLine & "Y Position: " & llCoord.Ycoord
    'Range
        Set rng = GetRange(llCoord.Xcoord, llCoord.Ycoord)
        If Not rng Is Nothing Then
            'MsgBox "Cell under mouse is :" & rng.Address
            'rng.Select
            [M3] = "Cell(" + CStr(rng.Row) + ", " + CStr(rng.Column) + ")" _
                 + "  ScreenCoord(" + CStr(llCoord.Xcoord) + ", " + CStr(llCoord.Ycoord) + ")"
        Else
            'MsgBox "Not a valid location."
        End If
        DoEvents
    Loop
    
    
End Sub

Function GetRange(x As Long, y As Long) As Range
    On Error Resume Next
    Set GetRange = ActiveWindow.RangeFromPoint(x, y)
End Function

Sub shapeClicked()
    'Assign the macro to all the shapes you have (right click, "Assign Macro").
    Dim shape As shape, shapeName As String
    shapeName = Application.Caller
    Set shape = ActiveSheet.Shapes(shapeName)
    shape.TopLeftCell.Select
End Sub

Sub Convert_png_to_jpg()
    'Called from    [none]
    'Realized with  GraphicsMagic
    'ToDo           Sub Convert_png_to_jpg(PathOfFolder$, _
    '               Optional InclSubFolders01 As Integer = 1, _
    '               Optional DeleteOriginal01 As Integer = 1)
    'Action         Umwandlung von allen png-Files eines Ordners in jpg-Files;
    '               Original-png wird dabei nicht gelöscht
    '               (Video-snaps werden oft als png gespeichert)
    
    'Vorbereitung
        Dim qq$, s1$, shellCmd$, p1$, p2$, p3$, i%, Arr1() As String
        qq = Chr(34) 'quote, "-Zeichen
    'Einstellungen
        p1 = "F:\Archiv TR\Archiv Trampolin\Events\1965 WM02 London"
        'p1 = "G:\Archiv Photos\- Photos ubes\Photos Trampolin"
    'Alle File-Pfade holen
        s1 = Get_FilePaths_Like_insideSourceFolderAndSubFolders(p1)
        'show s1
    'Action
        Arr1 = Split(s1, vbCrLf)
        For i = 0 To UBound(Arr1)
            p2 = Arr1(i)            'Path of File
            If LCase(Right(p2, 4)) = ".png" Then
                '[F3] = "[" + CStr(UBound(arr1) - i) + "] " + Get_NameFromPath(p2)
                p3 = Replace(p2, ".png", ".jpg")
                shellCmd = "gm convert " + qq + p2 + qq + " " + qq + p3 + qq
                'Show ShellCmd
                Call shell(shellCmd)
            End If
        Next
    Beep
End Sub

Sub TEST_FileNameDoubles()
    'Called from    [none]

    Dim s1$, N$, p1$, c%, i%, Arr1() As String
    p1 = "G:\Archiv Photos\- Photos ubes\Photos Trampolin\Events"
    s1 = Get_FilePaths_Like_insideSourceFolderAndSubFolders(p1)
    
    show s1
    
    Arr1 = Split(s1, vbCrLf)
    For i = 0 To UBound(Arr1)
        N = Get_NameFromPath(Arr1(i))
        If LCase(Right(N, 4)) = ".jpg" Then
            c = anzAinB(N, s1)
            [f3] = UBound(Arr1) - i
            If c <> 1 Then Stop
        End If
    Next
    
    Beep
End Sub

Sub Add_00_To_pxx()
    'Called from    [none]
    'ParentFile     Date...p21...   (wird hier ggf. ergänzt zu 'Date...p21-00...')
    'ChildFiles     Date...p21-01..., Date...p21-02..., ... (Portrait-Ausschnitte)
    
        Dim D$, F$, N1$, N2$, N3$, PNr$, p1$, p2$, p3$, s1$, s2$, s3$, s4$, T$
        Dim c%, i%, j%, Arr1() As String, Arr2() As String
    'Einstellungen
        'p1 = "G:\Archiv Photos\- Photos ubes\Photos Trampolin\Events\1900- (Bruchsal)"
        p1 = "G:\Archiv Photos\- Photos ubes\Photos Trampolin\Events"
        'p1 = "G:\Archiv Photos\- Photos ubes\Photos Trampolin"
        
    'Alle "-01"-Files
        's1 = Get_FilePaths_Like_insideSourceFolderAndSubFolders(p1)
        s1 = Get_Paths_ofAllSubfoldersAllLevelsAsStringUseGlobalVar(p1)
        Arr1 = Split(s1, vbCrLf)
        For i = 0 To UBound(Arr1)
            p2 = Arr1(i)        'Path of one Folder
            s2 = Get_AllFilePaths_WithMyStringInFileName_OfOneFolder(p2, "-01")
            If s2 <> "" Then
                s3 = s3 + s2 + vbCrLf
            End If
        Next
        's3 = Liste aller Files mit "-01" in p1-Subfolders
        '     ggf. besteht Bedarf, das zugehörige ParentFile mit "-00" zu ergänzen
    '-00 anfügen?
        s3 = Delete_EmptyRowsInString(s3)
        Arr1 = Split(s3, vbCrLf)
        For i = 0 To UBound(Arr1)
            p3 = Arr1(i)                    'Path of one File with "-01 "
            N1 = Get_NameFromPath(p3)           'Name of File     with "-01 "
            F = Replace(p3, "\" + N1, "")    'Path of Folder
            D = GetDateOfFileName(N1)        '19850000 oder 20230708_160720
            PNr = GetPNrOfFileName(N1)       '"p5" oder "p123"
            s4 = Get_AllFileNamesOfOneFolder(F)
            
            'in s4 existiert ein FileName mit D und (P mit " " oder "-00 ")
            Arr2 = Split(s4, vbCrLf)
            For j = 0 To UBound(Arr2) - 1
                N2 = Arr1(j)                 'Name of File j in Folder F
                N3 = ""
                If InStr(1, N2, D) > 0 Then
                    'Datum von N2 = Datum von N1
                    If InStr(1, N2, " " + PNr + " ") > 0 Then
                        N3 = N2 'noch ohne "-00"
                        '  = Photo, dessen FileName im ganzen Archiv zu ändern ist
                    ElseIf InStr(1, N2, " " + PNr + ".") > 0 Then N3 = N2
                    End If
                End If
            Next
            
            T = T + "PNr=" + PNr + "; in: " + N1 + vbCrLf
                 
        Next
        
        
     
    show T
End Sub

Function GetDateOfFileName(N$) As String
    'Called from    Add_00_To_pxx

    If N Like "######## *" Then GetDateOfFileName = Left(N, 8)
    If N Like "########_###### *" Then GetDateOfFileName = Left(N, 15)
End Function

Function GetPNrOfFileName(NameOFile$) As String
    'Called from    Add_00_To_pxx
    'Action         liefert PNr = "p1234"
    'NameOFile      enthält "-01 "
    
        Dim N$, PNr$, C1%, C2%, i%
    N = NameOFile
    For i = 1 To Len(N) - 2
        If Mid(N, i, 2) = " p" And IsNumeric(Mid(N, i + 2, 1)) Then C1 = i + 1: Exit For
    Next
    If C1 = 0 Then Exit Function
    For i = C1 + 1 To Len(N)
        If Not IsNumeric(Mid(N, i, 1)) Then C2 = i - 1: Exit For
    Next
    If C2 = 0 Then Exit Function
    PNr = Mid(N, C1, C2 - C1 + 1)
    GetPNrOfFileName = PNr
End Function

Function GetTR_rn(N$) As String
    'Called from    Update_Links_inside_FolderLeute

    If InStr(1, N, "(rn ") = 0 And InStr(1, N, " rn ") = 0 Then Exit Function
    Dim m$, rn$, C1%, C2%       'Ma xrn Mi,...          'Ma xrn Mi)
                                'Ma (...xrn Mi,...     'Ma (...xrn Mi)
    If InStr(1, N, "(rn ") > 0 Then m = Replace(N, "(rn ", "xrn ")
    If InStr(1, N, ", rn ") > 0 Then m = Replace(N, ", rn ", ",xrn ")
    C1 = InStr(1, m, "xrn "): C2 = InStr(C1 + 1, m, ",")
    If C2 > 0 Then rn = Mid(m, C1 + 4, C2 - C1 - 4)        '...xrn Mi,...
    If C2 = 0 Then rn = Mid(m, C1 + 4, Len(m) - C1 - 4)    '...xrn Mi)
    GetTR_rn = rn
End Function

Function GetTR_Ort(N$) As String
    'Called from    Update_Links_inside_FolderLeute

    'N = "Maier, Max (vh Vh1, vh Vh2, vh Vh3, rn Ruf, Ort)"
    If InStr(1, N, "(") = 0 Then Exit Function
    Dim m$, Ort$, C1%, C2%
    C1 = InStrRev(N, "("): C2 = InStrRev(N, " ")
    If C1 < C2 Then C1 = C2                 '...xOrt)   'c1 = x-Position
    Ort = Mid(N, C1 + 1, Len(N) - C1 - 1)
    GetTR_Ort = Ort
End Function

Function Get_vh3_From_OneFolderName(N$) As String
    'Called from    Update_Links_inside_FolderLeute

    If InStr(1, N, ", vh") = 0 Then Exit Function
    If anzAinB("vh ", N) < 3 Then Exit Function
    Dim vh3$, C1%, C2%, C3%
    C1 = InStrRev(N, "vh "): C2 = InStr(C1 + 1, N, ",")
    If C2 > 0 Then vh3 = Mid(N, C1 + 3, C2 - C1 - 3)        '...vh Maa,...
    If C2 = 0 Then vh3 = Mid(N, C1 + 3, Len(N) - C1 - 3)    '...vh Maa)
    Get_vh3_From_OneFolderName = vh3
End Function

Function Get_vh2_From_OneFolderName(N$) As String
    'Called from    Update_Links_inside_FolderLeute

    If InStr(1, N, ", vh") = 0 Then Exit Function
    Dim vh2$, C1%, C2%
    C1 = InStr(1, N, ", vh"): C2 = InStr(C1 + 1, N, ",")    '           12   17 20
    If C2 > 0 Then vh2 = Mid(N, C1 + 5, C2 - C1 - 5)        'Mai (vh Maa, vh Moo,...
    If C2 = 0 Then vh2 = Mid(N, C1 + 5, Len(N) - C1 - 5)    'Mai (vh Maa, vh Moo)
    Get_vh2_From_OneFolderName = vh2
End Function

Function Get_vh1_From_OneFolderName(N$) As String
    'Called from    Update_Links_inside_FolderLeute

    If InStr(1, N, "(vh") = 0 Then Exit Function
    Dim vh1$, C1%, C2%
    C1 = InStr(1, N, "(vh"): C2 = InStr(C1, N, ",")         '    5   9  12
    If C2 > 0 Then vh1 = Mid(N, C1 + 4, C2 - C1 - 4)        'Mai (vh Maa,...
    If C2 = 0 Then vh1 = Mid(N, C1 + 4, Len(N) - C1 - 4)    'Mai (vh Maa)
    Get_vh1_From_OneFolderName = vh1
End Function

Function Get_Vn_FromVnNn(VnNn$) As String
    If VnNn <> "" Then Get_Vn_FromVnNn = Left(VnNn, InStr(1, VnNn, " ") - 1)
End Function

Function Get_Nn_FromVnNn(VnNn$) As String
    If VnNn <> "" Then Get_Nn_FromVnNn = Mid(VnNn, InStr(1, VnNn, " ") + 1)
End Function

Function Get_NnVn_Give_VnNn(VnNn$) As String
    If VnNn <> "" Then
        Get_NnVn_Give_VnNn = Get_Nn_FromVnNn(VnNn) + ", " + Get_Vn_FromVnNn(VnNn)
    End If
End Function

Function Get_VnNn_Give_NnVn(NnVn$) As String
    Dim c%
    If NnVn <> "" Then
        c = InStr(1, NnVn, ",")
        Get_VnNn_Give_NnVn = Mid(NnVn, c + 2) + " " + Left(NnVn, c - 1)
    End If
End Function

Function GetTR_Vorname(N$) As String
    'Called from    Update_Links_inside_FolderLeute

    Dim Vn$, C1%, C2%
    C1 = InStr(1, N, ","): C2 = InStr(1, N, "(")
    If C1 > 0 And C2 = 0 Then Vn = Mid(N, C1 + 2)               'Mau, Max
    If C1 > 0 And C2 > C1 Then Vn = Mid(N, C1 + 2, C2 - C1 - 3) 'Mau, Max (...
    GetTR_Vorname = Vn
End Function

Function GetTR_Nachname(N$) As String
    'Called from    Update_Links_inside_FolderLeute

    'N = 'OneName  'Mai, Max (vh Maa, vh Mia, vh Chu, Ort)
    '              'Mai  'Mai, Max  'Mai, Max (rn Maxi, Ort)
    Dim Nn$, C1%, C2%
    C1 = InStr(1, N, ","): C2 = InStr(1, N, "(")
    If C1 = 0 And C2 = 0 Then Nn = Trim(N)                      'Max
    If C1 = 0 And C2 > 0 Then Nn = Left(N, C2 - 2)              'Max (...
    If C1 > 0 And C2 > 0 And C2 < C1 Then Nn = Left(N, C2 - 2)  'Max (...,...
    If C1 > 0 And C2 = 0 Then Nn = Left(N, C1 - 1)              'Mau, Max
    If C1 > 0 And C2 > C1 Then Nn = Left(N, C1 - 1)             'Mau, Max (...
    GetTR_Nachname = Nn
End Function

Sub Delete_AllLinkFiles_insideOneFolder(p1$)
    'Called from    Update_Links_inside_FolderLeute, T7_Read_LbToDoFiles
    
    Dim s$, i%, Arr1() As String
    If Right(p1, 1) <> "\" Then p1 = p1 + "\"
    s = Get_AllFileNames_Like_OfOneFolder(p1, "*.lnk")
    Arr1 = Split(s, vbCrLf)
    For i = 0 To UBound(Arr1)
        DeleteFile p1 + Arr1(i)
    Next
End Sub

Sub RefreshScreen()
    Application.ScreenUpdating = True: DoEvents
End Sub

Sub Update_Links_inside_FolderLeute()
    'Called from    T1_UpdateStatistics

    'Vorbereitung
        Dim L$, N$, Nn$, Vn$, vh1$, vh2$, vh3$, rn$, Ort$, s1$, v$, p1$, p2$, T$
        Dim FileOrFolderToJumpTo$, NamesOfPersonFolders$
        Dim AnzLnk%, c%, i%, sReg%, zReg%, Arr1() As String
        v = vbCrLf: DoArr
        p1 = ArrC(5)  'Path of Folder "Links"
        p2 = ArrC(4)  'Path of Folder "Leute"
        sReg = Get_T1Column_HoldingMyText("UnterOrdner4:  'Register'") + 8
        zReg = Get_T1Row_HoldingMyText("UnterOrdner4:  'Register'") + 1
    'T1-Info
        EE 0
        With ActiveWorkbook.Sheets("T1")
        .Cells(zReg, sReg) = "'(00)"
    'Delete old Links
        Delete_AllLinkFiles_insideOneFolder p1
    NamesOfPersonFolders = Get_Names_OfAllSubfolders_OneLevel(p2)
    Arr1 = Split(NamesOfPersonFolders, vbCrLf)
            
    For i = 0 To UBound(Arr1)
        'Schleife über alle Namen der PersonenOrdner in 'Leute'
        N = Arr1(i) 'OneName  'Mai, Max (vh Maa, vh Mia, vh Chu, Ort)
        '                     'Mai  'Mai, Max  'Mai, Max (rn Maxi, Ort)
        Nn = GetTR_Nachname(N): Vn = GetTR_Vorname(N): rn = GetTR_rn(N)
        vh1 = Get_vh1_From_OneFolderName(N): vh2 = Get_vh2_From_OneFolderName(N): vh3 = Get_vh3_From_OneFolderName(N): Ort = GetTR_Ort(N)
        FileOrFolderToJumpTo = p2 + "\" + N
        If c Mod 10 = 0 Then .Cells(zReg, sReg) = "'(" + CStr((UBound(Arr1) - i) \ 10) + ")"
        
    'Link1      Maier, Max (vh Vh1, vh Vh2, vh Vh3, rn Ruf, Ort)
        L = N 'L = LinkNameToDisplay
        Create_OneLinkFile p1, FileOrFolderToJumpTo, L
        'T1-Info
            c = c + 1 ': .Cells(Z + 10, S + 11) = "'" + CStr(c)
    'Link2      Maier, Max (vh Vh1, vh Vh2, vh Vh3, rn Ruf, Ort)
        If Vn <> "" Then
            L = Vn + " " + Nn + Replace(N, Nn + ", " + Vn, "")
            Create_OneLinkFile p1, FileOrFolderToJumpTo, L
            'T1-Info
                c = c + 1 ': .Cells(Z + 10, S + 11) = "'" + CStr(c)
        End If
    'Link3      Ruf (Max Maier, vh Vh1, vh Vh2, vh Vh3, Ort)
        If rn <> "" Then
            L = rn + " (" + Vn + " " + Nn
            If vh1 <> "" Then L = L + ", vh " + vh1
            If vh2 <> "" Then L = L + ", vh " + vh2
            If vh3 <> "" Then L = L + ", vh " + vh3
            If Ort <> "" Then L = L + ", " + Ort
            L = L + ")"
            Create_OneLinkFile p1, FileOrFolderToJumpTo, L
            'T1-Info
                c = c + 1 ': .Cells(Z + 10, S + 11) = "'" + CStr(c)
        End If
    'Link4      Vh1, VN (geb NN, vh Vh2, vh Vh3, rn Ruf, Ort)
        If vh1 <> "" Then
            L = vh1 + ", " + Vn + " (geb " + Nn
            If vh2 <> "" Then L = L + ", vh " + vh2
            If vh3 <> "" Then L = L + ", vh " + vh3
            If rn <> "" Then L = L + ", rn " + rn
            If Ort <> "" Then L = L + ", " + Ort
            L = L + ")"
            Create_OneLinkFile p1, FileOrFolderToJumpTo, L
            'T1-Info
                c = c + 1 ': .Cells(Z + 10, S + 11) = "'" + CStr(c)
        End If
    'Link5      Vh2, VN (geb NN, vh Vh1, vh Vh3, rn Ruf, Ort)
        If vh2 <> "" Then
            L = vh2 + ", " + Vn + " (geb " + Nn + ", vh " + vh1
            If vh3 <> "" Then L = L + ", vh " + vh3
            If rn <> "" Then L = L + ", " + rn
            If Ort <> "" Then L = L + ", " + Ort
            L = L + ")"
            Create_OneLinkFile p1, FileOrFolderToJumpTo, L
            'T1-Info
                c = c + 1 ': .Cells(Z + 10, S + 11) = "'" + CStr(c)
        End If
    'Link6      Vh3, VN (geb NN, vh Vh1, vh Vh3, rn Ruf, Ort)
        If vh3 <> "" Then
            L = vh3 + ", " + Vn + " (geb " + Nn + ", vh " + vh1 + ", vh " + vh2
            If rn <> "" Then L = L + ", " + rn
            If Ort <> "" Then L = L + ", " + Ort
            L = L + ")"
            Create_OneLinkFile p1, FileOrFolderToJumpTo, L
            'T1-Info
                c = c + 1 ': .Cells(Z + 10, S + 11) = "'" + CStr(c)
        End If
    Next
    'finals
        .Cells(zReg, sReg) = ""
        .Cells(zReg, sReg + 3) = "'" + CStr(CountFilesInFolder(p1))
        EE 1: End With: Beep
End Sub

Function Get_pIDList()
    Dim s$, AnzPhot00Ev%, i%
    AnzPhot00Ev = CInt(ArrC(15))  'Anzahl 00-Photos in "Events"
    For i = 1 To AnzPhot00Ev + 300
    s = s + " p" + Format(i, "0000") '" p0001 p0002 ... "
    Next
    Get_pIDList = s
End Function

Function Get_vIDList()
    Dim s$, AnzVidGes%, i%
    AnzVidGes = CInt(ArrC(20))  'Anzahl Videos      gesamt
    For i = 1 To AnzVidGes + 100
    s = s + " v" + Format(i, "0000") '" v0001 v0002 ... "
    Next
    Get_vIDList = s
End Function

Function Get_FileExtension(PathOfFile$) As String
    Dim F() As String: F = Split(PathOfFile, ".")
    Get_FileExtension = LCase(F(UBound(F)))
End Function

Function isPhoto(NameOrPathOfFile$) As Boolean
    Dim E$, s$
    s = ";ani;bmp;dng;gif;heic;ico;jpe;jpeg;jpg;pcx;png;psd;tga;tif;tiff;webp;wmf;"
    E = Get_FileExtension(NameOrPathOfFile): If E = "" Then Exit Function
    If s Like "*;" + E + ";*" Then isPhoto = True
End Function

Function isVideo(NameOrPathOfFile$) As Boolean
    Dim E$, s$
    s = ";3g2;3gp;3gp2;3gpp;amr;amv;asf;avi;bdmv;bik;d2v;divx;drc;dsa;dsm;dss;dsv;evo;f4v;flc;fli;flic;flv;hdmov;ifo;ivf;m1v;m2p;m2t;m2ts;m2v;m4v;mkv;mp2v;mp4;mp4v;mpe;mpeg;mpg;mpls;mpv2;mpv4;mov;mts;ogm;ogv;pss;pva;qt;ram;ratdvd;rm;rmm;rmvb;roq;rpm;smil;smk;swf;tp;tpr;ts;vob;vp6;webm;wm;wmp;wmv;"
    E = Get_FileExtension(NameOrPathOfFile): If E = "" Then Exit Function
    If s Like "*;" + E + ";*" Then isVideo = True
End Function

Sub FillQuellenKürzel(p$, Shorty$, NoShorty$)
    'Called from    T1_UpdateStatistics
    'Shorty         Quellenkürzel; Variable wird fortlaufend ergänzt
    'p              Pfad einer einzelnen Datei
    'Action         p wird auf die Existenz eines gültigen QuellenKürzels untersucht;
    '               nur Fotos/Videos aus Ordnern 'Events' und 'ClubsNations'
    
    'Vorbereitung
        Dim E$, s2$, s3$, s4$, L%, c&
    'Exit
        If Not (p Like "*\Events\*" Or p Like "*\ClubsNations\*") Then Exit Sub
        If Not (isVideo(p) Or isPhoto(p)) Then Exit Sub
    'Falls FileName kein QuellenKürzel besitzt
        E = Get_FileExtension(p): L = Len(E)
        If Not p Like "* [a-z][a-z]." + E Then NoShorty = NoShorty + p + vbCrLf: Exit Sub
    'Action
        'p bezeichnet ein Foto oder Video und besitzt ein QuellenKürzel
        s2 = "|" + Right(p, L + 3) + " "            '= "|ub.jpg " oder "|nn.jpeg "
        'Shorty sammelt Quellen                     '|ub.jpg 1234|sf.mpg 0003|...
        c = InStr(1, Shorty, s2)
        If c = 0 Then
            's2 ist noch nicht in Shorty-Sammlung enthalten
            Shorty = Shorty + s2 + "0001" + "|"     'neues (1.) Kürzel "aa.jpg 0001"
        Else
            'Die bisherige Anzahl des Kürzels       wird um 1 erhöht
            s3 = Mid(Shorty, c, L + 10)             '= "|ub.jpg 1235|"
            s4 = s2 + Format(CInt(Mid(s3, L + 6, 4)) + 1, "0000") + "|"
            Shorty = Replace(Shorty, s3, s4)
        End If
        DoEvents
End Sub

Sub ShowNoShorty(NoShorty$)
    'Called from    T1_UpdateStatistics
    Dim T$, i%, Arr1() As String
    'Alle Ordner, die Files ohne Quellenkürzel enthalten, im Explorer öffnen
        Arr1 = Split(NoShorty, vbCrLf)
        For i = 0 To UBound(Arr1)
            OpenFolder Get_FolderPath_Give_FilePath(Arr1(i))
        Next
    'Hinweis anzeigen
        T = "ReperaturBedarf" + vbCrLf + vbCrLf _
           + "Während T1_UpdateStatistics() wurden Dateien entdeckt, die kein Quellenkürzel besitzen" + vbCrLf _
           + "(muss händisch repariert werden - bitte QuellenKürzel in FileName dazuschreiben)." _
           + vbCrLf + vbCrLf + String(105, "-") + vbCrLf
        If NoShorty <> "" Then show T + NoShorty + String(105, "-") + vbCrLf
    Beep
End Sub

Sub ShowPIDsDoubles(pIDsDoubles$)
    'Called from    T1_UpdateStatistics
    Dim T$
    T = "ReperaturBedarf" + vbCrLf + vbCrLf _
       + "T1_UpdateStatistics() --> Liste mehrfach benutzter IDs" + vbCrLf _
       + "(muss händisch repariert werden - per SearchEverythingTool in 'Archiv Trampolin')" _
       + vbCrLf + String(105, "-") + vbCrLf
    If pIDsDoubles <> "" Then show T + pIDsDoubles + String(105, "-") + vbCrLf
End Sub

Sub ShowVIDsDoubles(vIDsDoubles$)
    'Called from    T1_UpdateStatistics
    Dim T$
    T = "ReperaturBedarf" + vbCrLf + vbCrLf _
       + "T1_UpdateStatistics() --> Liste mehrfach benutzter IDs" + vbCrLf _
       + "(muss händisch repariert werden - per SearchEverythingTool in 'Archiv Trampolin')" _
       + vbCrLf + String(105, "-") + vbCrLf
    If vIDsDoubles <> "" Then show T + vIDsDoubles + String(105, "-") + vbCrLf
End Sub

Function Get_LinesOfString_Like(SourceString$, LikeString$) As String
    Dim s$, v$, i%, Arr1() As String
    v = vbCrLf
    Arr1 = Split(SourceString, v)
    For i = 0 To UBound(Arr1)
        If Arr1(i) Like LikeString Then s = s + v + Arr1(i)
    Next
    s = Mid(s, 3)
    Get_LinesOfString_Like = s
End Function





