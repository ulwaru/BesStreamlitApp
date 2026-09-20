Attribute VB_Name = "M11"
Option Explicit

Public AllErrors As String

Sub zzz_M11()
    showProcs " like OneFolder"
    
    'T1_ShowLogBuch
    'RenameModule "Modul1", "Load"
    'show Get_RangeOfFirstCellOfOneDesign("1969 LTV-Pokal Saarlouis")
    'LogBuch "Hallo"
    EE 1: Beep
End Sub

Sub Update_ReadMe_QUELLEN(Shorty$, ReadMeNEW$)
    'Called from    T1_UpdateStatistics
    'Shorty         Liste QuellenKürzel
    'ReadMeNEW      Update_ReadMe_ARCHIVTRAMPOLIN is done
    'Action         Aktualisiert den Teil-Bereich 'QUELLEN' aus der bisherigen
    '               Datei 'Read me.txt' mit Daten aus Shorty

    'Vorbereitung
        Dim k$, L$, r$, s$, t1$, T2$, T3$, T4$, v$
        Dim c&, C1&, C2&, i&, j&, A() As String, B() As String
        v = vbCrLf
    'Einstellung
        c = 55 'Soll-Anzahl Zeichen vor Beginn der Ziffern in einer t2-Zeile
    'ReadMeNEW
        r = ReadMeNEW               ': show ReadMeNEW: Stop
    'Shorty
        s = Format_Shorty(Shorty)   ': show s: Stop
        's = Shorty 'tmp
    'Kürzel+Namen aus r herauslösen
        C1 = InStr(1, r, "QUELLEN ("): C1 = InStr(C1, r, v + "a") + 2   'Start t2
        C2 = InStr(C1, r, v + v + "===") - 1                            'Ende  t2
        t1 = Left(r, C1 - 1)        'r-Teil 1
        T2 = CutVonBis(r, C1, C2)   'r-Teil 2, KürzelZeilen
        T3 = Mid(r, C2 + 3)         'r-Teil 3
        'show t2: Stop
    'B()    KürzelZeilen: nur Kürzel + Name/Beschreibung aus bisherigem ReadMe
        B = Split(T2, v)
        For i = 0 To UBound(B)
            B(i) = Left(B(i), c)
        Next
    'A()    t4 = t2-Ersatz
        A = Split(s, v)
        For i = 0 To UBound(A)
            k = Left(A(i), 2)                   'nn
            For j = 0 To UBound(B)
                L = B(j)                        'nn  (Quelle nicht bekannt)
                If Left(L, 2) = k Then A(i) = L + Mid(A(i), 4)
            Next
        Next
        T4 = Join(A, v)
        'show t4
    'Finals
        ReadMeNEW = t1 + T4 + T3
        'show ReadMeNEW
End Sub

Sub T1_OpenFolder_Events(z%, s%)
    Call DoArrc: OpenFolder ArrC(3) 'Path of Folder "Events"
    ActiveWorkbook.Sheets("T1").Cells(z, s + 12) = "ü": RefreshScreen
    Application.Wait (Now + TimeValue("0:00:01"))
    ActiveWorkbook.Sheets("T1").Cells(z - 1, s - 1).Select
End Sub

Sub T1_OpenFolder_Clubs(z%, s%)
    Call DoArrc: OpenFolder ArrC(11) 'Path of Folder "ClubsNations"
    ActiveWorkbook.Sheets("T1").Cells(z, s + 12) = "ü": RefreshScreen
    Application.Wait (Now + TimeValue("0:00:01"))
    ActiveWorkbook.Sheets("T1").Cells(z - 1, s - 1).Select
End Sub

Sub T1_OpenFolder_Leute(z%, s%)
    Call DoArrc: OpenFolder ArrC(4) 'Path of Folder "Leute"
    ActiveWorkbook.Sheets("T1").Cells(z, s + 12) = "ü": RefreshScreen
    Application.Wait (Now + TimeValue("0:00:01"))
    ActiveWorkbook.Sheets("T1").Cells(z - 1, s - 1).Select
End Sub

Sub T1_OpenFolder_Register(z%, s%)
    Call DoArrc: OpenFolder ArrC(5) 'Path of Folder "Register"
    ActiveWorkbook.Sheets("T1").Cells(z, s + 12) = "ü": RefreshScreen
    Application.Wait (Now + TimeValue("0:00:01"))
    ActiveWorkbook.Sheets("T1").Cells(z - 1, s - 1).Select
End Sub

Sub T1_OpenFolder_ToDo(z%, s%)
    'F:\Archiv TR\ToDo
    Call DoArrc: OpenFolder ArrC(1) + "\ToDo" 'Path of Folder "Register"
    ActiveWorkbook.Sheets("T1").Cells(z, s + 12) = "ü": RefreshScreen
    Application.Wait (Now + TimeValue("0:00:01"))
    ActiveWorkbook.Sheets("T1").Cells(z - 1, s - 1).Select
End Sub

Sub T1_Report_UpdateStatistics_OldNew_TEST()
    Dim DateTest$, s$, ArrTest() As String
    DateTest = "10.08.2025"
    s = ReadFile("F:\Archiv TR\prog\statistics\Statistics 20250804 values.txt")
    ArrTest = Split(vbCrLf + s, vbCrLf)
    T1_Report_UpdateStatistics_OldNew DateTest, ArrTest
End Sub

Sub T1_Report_UpdateStatistics_OldNew(DateOld$, ByRef StatOld() As String)
    'Called from    T1_UpdateStatistics
    'DateOld        UpdateStatistics-Datum vor Auslösen des Update
    'StatOld        UpdateStatistics-Werte vor Auslösen des Update (1D-Array)

    'Vorbereitung
        Dim A$, B$, DateNew$, DateNew8$, DateOld8$, G$, p$, sN$, so$, v$
        Dim i%, k#, c&, StatNew() As String
        v = vbCrLf: DoArrc: p = ArrC(1) + "\prog\statistics\"
    'DateOld
        DateOld8 = Mid(DateOld, 7, 4) + Mid(DateOld, 4, 2) + Mid(DateOld, 1, 2)
    'DateNew
        DateNew = Left(Sheets("T1").Cells(4, 47), 10)
        DateNew8 = Mid(DateNew, 7, 4) + Mid(DateNew, 4, 2) + Mid(DateNew, 1, 2)
    'StatOld
        For i = 1 To UBound(StatOld)
            If StatOld(i) <> "" Then so = so + v + StatOld(i)
        Next
        StatOld = Split(so, v)
    'StatNew        T1-UpdateStatistics-Werte laden
        Load_Array1D_OfColumnPart StatNew, "T1", 5, 35, 46
        For i = 1 To UBound(StatNew)
            If StatNew(i) <> "" Then sN = sN + v + StatNew(i)
        Next
        StatNew = Split(sN, v) 'showArray StatOld:showArray StatNew
    'Vordruck laden
        A = ReadFile(p + "Update Statistics old-new.txt")
        'openFile p + "Update Statistics old-new.txt": Stop
    'Vordruck füllen
        A = Replace(A, "x00", DateOld)
        For i = 1 To UBound(StatOld)
            A = Replace(A, "x" + Format(i, "00"), String(7 - Len(StatOld(i)), " ") + StatOld(i))
        Next
        A = Replace(A, "y00", DateNew)
        For i = 1 To UBound(StatNew)
            A = Replace(A, "y" + Format(i, "00"), String(7 - Len(StatNew(i)), " ") + StatNew(i))
        Next
    'Diff
        For i = 1 To UBound(StatNew)
            If StatNew(i) Like "*GB" Then
                'g = StatNew(i): g = Left(StatNew(i), Len(StatNew(i)) - 3)
                k = CDbl(Left(StatNew(i), Len(StatNew(i)) - 3)) - CDbl(Left(StatOld(i), Len(StatOld(i)) - 3))
                If k > 0 Then B = "+" Else B = "-"
                k = Format(Abs(k), "0.0")
                If k = 0 Then
                    A = Replace(A, "z" + Format(i, "00"), "")
                Else
                    A = Replace(A, "z" + Format(i, "00"), B + String(4 - Len(CStr(k)), " ") + CStr(k) + " GB")
                End If
            Else
                'showArray StatOld: showArray StatNew
                'Stop
                c = CLng(StatNew(i)) - CLng(StatOld(i))
                If c > 0 Then B = "+" Else B = "-"
                c = Abs(c)
                If c = 0 Then
                    A = Replace(A, "z" + Format(i, "00"), "")
                Else
                    A = Replace(A, "z" + Format(i, "00"), B + String(7 - Len(CStr(c)), " ") + CStr(c))
                End If
            End If
            
        Next
    'Finals
        show A
        writeStringToFile p + "Statistics " + DateNew8 + " values.txt", Join(StatOld, v)
        writeStringToFile p + "Statistics " + DateNew8 + " compare.txt", A
        'showArray stat
End Sub

Sub T1_UpdateStatistics()
    'Called from    Worksheet_SelectionChange[T1; UserClick onto CellButton 'Update Statistics'],
    '               Update_All
    'Scope          Archiv Trampolin
    
    'Vorbereitung
        Dim AnzFolEv$, AnzFolLe$, NoShorty$, p1$, PathsOfFoldersNoID$, Pid$, pIDsDoubles$, pIDsUsed$
        Dim ReadMeNEW$, s1$, s2$, s3$, s4$, s5$, Shorty$, SizeAr$, SizeEv$, SizeLe$, SizeLi$
        Dim T$, vID$, vIDsDoubles$, vIDsUsed$, DateOld$
        Dim AnzCropCn%, AnzCropEv%, AnzCropLe%, AnzFold%
        Dim AnzLnkCn%, AnzLnkEv%, AnzLnkLe%, AnzLnkLi%
        Dim AnzPhotCn%, AnzPhotEv%, AnzPhotLe%
        Dim AnzSnapEv%, AnzSnapCn%, AnzSnapLe%
        Dim AnzVidCn%, AnzVidCnT%, AnzVidEv%, AnzVidEvT%
        Dim c%, C2%, i%, s%, z%
        Dim AnzFilesGes&, AnzPhotGes&, AnzVidGes&, AnzLnkGes&
        Dim Arr1() As String, stat() As String, y() As String, StatOld() As String
        Dim oFSO As Object, oFolder As Object, oFile As Object, sF, colFolders As New Collection
        ReDim stat(1 To 32): DoArr
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        Set oFolder = oFSO.GetFolder(ArrC(2)) 'Path of Folder "Archiv Trampolin"
        z = Get_T1Row_HoldingMyText("Update Statistics")
        s = Get_T1Column_HoldingMyText("Update Statistics")
        LogBuch "'T1_UpdateStatistics' was started."
        With ActiveWorkbook.Sheets("T1")
    'ReportOnChanges
        DateOld = Left(Sheets("T1").Cells(4, 47), 10)
        Load_Array1D_OfColumnPart StatOld, "T1", 5, 35, 46
    'alte Einträge löschen/ändern
        T1_UpdateStatistics1 stat
    'SourceFolder in Collection schreiben
        colFolders.Add oFolder          'start with this folder (Folder "Archiv Trampolin")
        AnzFold = colFolders.count      '= zunächst 1
    Do While colFolders.count > 0       'process all folders
        Set oFolder = colFolders(1)     'get one folder to process
        colFolders.Remove 1             'remove item at index 1
        For Each oFile In oFolder.Files
        
        'z+5    [Photos  gesamt]
                AnzPhotGes = AnzPhotCn + AnzCropCn + AnzSnapCn + AnzPhotEv + AnzCropEv + AnzPhotLe + AnzCropLe + AnzSnapEv + AnzSnapLe
                If AnzPhotGes Mod 500 = 0 Then .Cells(z + 4, s + 11) = "'" + CStr(AnzPhotGes)
        'z+6    [Videos  gesamt]
                AnzVidGes = AnzVidEv + AnzVidEvT + AnzVidCn + AnzVidCnT
                If AnzVidGes Mod 5 = 0 Then .Cells(z + 5, s + 11) = "'" + CStr(AnzVidGes)
        'z+7    [Links   gesamt]
                AnzLnkGes = AnzLnkCn + AnzLnkEv + AnzLnkLe + AnzLnkLi
                If AnzLnkGes Mod 200 = 0 Then .Cells(z + 6, s + 11) = "'" + CStr(AnzLnkGes)
        'z+8    [Dateien gesamt]
                AnzFilesGes = AnzFilesGes + 1
                If AnzFilesGes Mod 500 = 0 Then .Cells(z + 7, s + 11) = "'" + CStr(AnzFilesGes) 'T1-Anzeige

            p1 = oFolder.path + "\" + oFile.NAME

    '---------------------------------------------------------
    
        'p1 = nacheinander jeder FilePath (inside "Archiv Trampolin" and SubFolders)
            
        'QuellenKürzel sammeln
            FillQuellenKürzel p1, Shorty, NoShorty '(ShortyVariablen werden fortlaufend ergänzt)
        'PathsOfFoldersNoID sammeln
            If isPhoto(p1) Then
                If Not p1 Like "* [p|v]####-## *" Then
                    'Der PhotoPfad p1 enthält keine pID; sein OrdnerPfad --> PathsOfFoldersNoID
                    If InStr(1, PathsOfFoldersNoID, oFolder.path) = 0 Then PathsOfFoldersNoID = PathsOfFoldersNoID + oFolder.path + vbCrLf
                End If
        'pIDsDoubles sammeln
                If p1 Like "*\Events\* p####-00 *" Then
                    Pid = Mid(p1, InStr(1, p1, "-00") - 5, 8)
                    If InStr(1, pIDsUsed, "|" + Pid + "|") = 0 Then pIDsUsed = pIDsUsed + "|" + Pid + "|" + vbCrLf Else pIDsDoubles = pIDsDoubles + "|" + Pid + "|" + vbCrLf
                End If
            ElseIf isVideo(p1) Then
                If Not p1 Like "* v####-## *" Then
                    'Der VideoPfad p1 enthält keine vID; sein OrdnerPfad --> PathsOfFoldersNoID
                    If InStr(1, PathsOfFoldersNoID, oFolder.path) = 0 Then PathsOfFoldersNoID = PathsOfFoldersNoID + oFolder.path + vbCrLf
                End If
        'vIDsDoubles sammeln
                If p1 Like "*\Events\* v####-00 *" Then
                    vID = Mid(p1, InStr(1, p1, "-00") - 5, 8)
                    If InStr(1, vIDsUsed, "|" + vID + "|") = 0 Then vIDsUsed = vIDsUsed + "|" + vID + "|" + vbCrLf Else vIDsDoubles = vIDsDoubles + "|" + vID + "|" + vbCrLf
                End If
            End If
            
        Select Case True
        
        'z+11   [Photos] [in 'Events']
                Case p1 Like "*\Events\*-00 *.jpg"
                    AnzPhotEv = AnzPhotEv + 1
                    If AnzPhotEv Mod 200 = 0 Then .Cells(z + 11, s + 11) = "'" + CStr(AnzPhotEv) 'T1-Anzeige
        'z+12   [Ausschnitte] [in 'Events']
                Case p1 Like "*\Events\*p####-## *.jpg" And Not p1 Like "*\Events\*p####-00 *.jpg"
                     AnzCropEv = AnzCropEv + 1
                     If AnzCropEv Mod 100 = 0 Then .Cells(z + 12, s + 11) = "'" + CStr(AnzCropEv) 'T1-Anzeige
        'z+15   [Snapshots] [in 'Events']
                Case p1 Like "*\Events\*v####*.jpg"
                     AnzSnapEv = AnzSnapEv + 1
                     If AnzSnapEv Mod 5 = 0 Then .Cells(z + 15, s + 11) = "'" + CStr(AnzSnapEv) 'T1-Anzeige
        'z+13   [Videos] [in 'Events']  .avi .mov .mp4 .mpg .wmv
                Case p1 Like "*\Events\*-00 *.avi", p1 Like "*\Events\*-00 *.mov", p1 Like "*\Events\*-00 *.mp4", p1 Like "*\Events\*-00 *.mpg", p1 Like "*\Events\*-00 *.wmv"
                    AnzVidEv = AnzVidEv + 1
                    If AnzVidEv Mod 5 = 0 Then .Cells(z + 13, s + 11) = "'" + CStr(AnzVidEv) 'T1-Anzeige
        'z+14   [Teil-Videos] [in 'Events']  .avi .mov .mp4 .mpg .wmv
                Case isVideo(p1) And p1 Like "*\Events\*-## *" And Not p1 Like "*\Events\*-00 *"
                    AnzVidEvT = AnzVidEvT + 1
                    If AnzVidEvT Mod 1 = 0 Then .Cells(z + 14, s + 11) = "'" + CStr(AnzVidEvT) 'T1-Anzeige
        
        'z+19   [Photos] [in '..Clubs..']
                Case p1 Like "*\ClubsNations\*-00 *.jpg"
                    AnzPhotCn = AnzPhotCn + 1
                    If AnzPhotCn Mod 20 = 0 Then .Cells(z + 19, s + 11) = "'" + CStr(AnzPhotCn) 'T1-Anzeige
        'z+20   [Ausschnitte] [in '..Clubs..']
                Case p1 Like "*\ClubsNations\*p####-## *.jpg" And Not p1 Like "*\Events\*p####-00 *.jpg"
                     AnzCropCn = AnzCropCn + 1
                     If AnzCropCn Mod 20 = 0 Then .Cells(z + 20, s + 11) = "'" + CStr(AnzCropCn) 'T1-Anzeige
        'z+23   [Snapshots] [in '..Clubs..']
                Case p1 Like "*\ClubsNations\*v####*.jpg"
                     AnzSnapCn = AnzSnapCn + 1
                     If AnzSnapCn Mod 1 = 0 Then .Cells(z + 23, s + 11) = "'" + CStr(AnzSnapCn) 'T1-Anzeige
        'z+21   [Videos] [in '..Clubs..']  .avi .mov .mp4 .mpg .wmv
                Case p1 Like "*-00 *.avi", p1 Like "*-00 *.mov", p1 Like "*-00 *.mp4", p1 Like "*-00 *.mpg", p1 Like "*-00 *.wmv"
                    AnzVidCn = AnzVidCn + 1
                    If AnzVidCn Mod 5 = 0 Then .Cells(z + 21, s + 11) = "'" + CStr(AnzVidCn) 'T1-Anzeige
        'z+22   [Teil-Videos] [in '..Clubs..']  .avi .mov .mp4 .mpg .wmv
                Case isVideo(p1) And p1 Like "*-## *" And Not p1 Like "*-00 *"
                    AnzVidCnT = AnzVidCnT + 1
                    If AnzVidCnT Mod 1 = 0 Then .Cells(z + 22, s + 11) = "'" + CStr(AnzVidCnT) 'T1-Anzeige
        
        'z+27   [Photos] [in 'Leute']
                Case p1 Like "*\Leute\*p####-00 *.jpg"
                    AnzPhotLe = AnzPhotLe + 1
                    If AnzPhotLe Mod 500 = 0 Then .Cells(z + 27, s + 11) = "'" + CStr(AnzPhotLe) 'T1-Anzeige
        'z+28   [Ausschnitte] [in 'Leute']
                Case p1 Like "*\Leute\*p####-## *.jpg" And Not p1 Like "*\Leute\*p####-00 *.jpg"
                    AnzCropLe = AnzCropLe + 1
                    If AnzCropLe Mod 500 = 0 Then .Cells(z + 28, s + 11) = "'" + CStr(AnzCropLe) 'T1-Anzeige
        'z+29   [Snapshots] [in 'Leute']
                Case p1 Like "*\Leute\*v####*.jpg"
                     AnzSnapLe = AnzSnapLe + 1
                     If AnzSnapLe Mod 10 = 0 Then .Cells(z + 29, s + 11) = "'" + CStr(AnzSnapLe) 'T1-Anzeige
        
        'z+16   [Links] [in 'Events']
                Case p1 Like "*\Events\*.lnk"
                    AnzLnkEv = AnzLnkEv + 1
                    If AnzLnkEv Mod 50 = 0 Then .Cells(z + 16, s + 11) = "'" + CStr(AnzLnkEv) 'T1-Anzeige
        'z+24   [Links] [in '..Clubs..']
                Case p1 Like "*\ClubsNations\*.lnk"
                    AnzLnkCn = AnzLnkCn + 1
                    If AnzLnkCn Mod 50 = 0 Then .Cells(z + 24, s + 11) = "'" + CStr(AnzLnkCn) 'T1-Anzeige
        'z+30   [Links] [in 'Leute']
                Case p1 Like "*\Leute\*.lnk"
                    AnzLnkLe = AnzLnkLe + 1
                     If AnzLnkLe Mod 10 = 0 Then .Cells(z + 30, s + 11) = "'" + CStr(AnzLnkLe) 'T1-Anzeige
        'z+32   [Links] [in 'Links']
                Case p1 Like "*\Register\*.lnk"
                    AnzLnkLi = AnzLnkLi + 1
                    If AnzLnkLi Mod 100 = 0 Then .Cells(z + 32, s + 11) = "'" + CStr(AnzLnkLi) 'T1-Anzeige
        
            End Select
            
        Next oFile
        
    '---------------------------------------------------------
     
        'add any subfolders to the collection for processing
            For Each sF In oFolder.subfolders
                colFolders.Add sF
            Next sF
    Loop
        End With
        
    'Einträge genauer Zahlen in ArrC
       '[...]-items has been filled in T1_UpdateStatistics1
       'STAT(1) = [Date/Time]   'STAT(2) = [SizeGes]
        stat(3) = "4":           stat(4) = AnzPhotGes:   stat(5) = AnzVidGes
        stat(6) = AnzLnkGes:     stat(7) = AnzFilesGes
        stat(8) = ""            'STAT(9) = [SizeEv]     'STAT(10) = [AnzFolEv]
        stat(11) = AnzPhotEv:    stat(12) = AnzCropEv:   stat(13) = AnzVidEv
        stat(14) = AnzVidEvT:    stat(15) = AnzSnapEv:   stat(16) = AnzLnkEv
       'STAT(17) = [SizeCn]     'STAT(18) = [AnzFolCn]
        stat(19) = AnzPhotCn:    stat(20) = AnzCropCn:   stat(21) = AnzVidCn
        stat(22) = AnzVidCnT:    stat(23) = AnzSnapCn:   stat(24) = AnzLnkCn
       'STAT(25) = [SizeLe]     'STAT(26) = [AnzFolLe]
        stat(27) = AnzPhotLe:    stat(28) = AnzCropLe:   stat(29) = AnzSnapLe
        stat(30) = AnzLnkLe     'STAT(31) = [SizeLi]
        stat(32) = AnzLnkLi
        
    'pIDs, vIDs zu einigen FileNames hinzufügen
        If PathsOfFoldersNoID <> "" Then
            PathsOfFoldersNoID = Delete_EmptyEndRowsInString(PathsOfFoldersNoID)
            Arr1 = Split(PathsOfFoldersNoID, vbCrLf)
            For i = 0 To UBound(Arr1)
                Add_IDs_toFilesInOneEventFolder Arr1(i), C2
                'c2 wird bei jeder hinzugefügten Id hochgezählt
                'c2 wird hier nicht weiter verwendet; nur in:
                'UserClick on T6-CellButton 'Add ID in open EventFolders'
            Next
        End If
    'Einträge genauer Zahlen in T1
        T1_UpdateStatistics2 stat
    'Eintrag STAT in ArrC
        FillArrC 32, Join(stat, "|")
    'Alte Links im Ordner 'Links' löschen und neu erzeugen
        Update_Links_inside_FolderLeute
    'Shorty
        Shorty = Replace(Shorty, "||", "|"): Shorty = Mid(Shorty, 2, Len(Shorty) - 2)
        y = Split(Shorty, "|"): QuickSort y: Shorty = Join(y, vbCrLf)
        'show Shorty: Stop
    'Update ReadMe
        ReadMeNEW = ReadFile(ArrC(7))       'Path of File   "Read me.txt"
        Update_ReadMe_ARCHIVTRAMPOLIN ReadMeNEW
        Update_ReadMe_QUELLEN Shorty, ReadMeNEW
        writeStringToFile ArrC(7), ReadMeNEW
    'Report on Changes
        T1_Report_UpdateStatistics_OldNew DateOld, StatOld
    'Finals
        If NoShorty <> "" Then ShowNoShorty NoShorty
        If pIDsDoubles <> "" Then ShowPIDsDoubles pIDsDoubles
        If vIDsDoubles <> "" Then ShowVIDsDoubles vIDsDoubles
        'ShowMissingPIDs 'ShowMissingVIDs
        'showArray STAT
        Beep
        LogBuch "'T1_UpdateStatistics' has ended."
End Sub 'T1_UpdateStatistics

Sub T1_UpdateStatistics1(stat() As String)
    'Called from    xxx
    'SourceFolder   Archiv Trampolin
    
    'Vorbereitung
        Dim T$, s%, z%
        Dim SizeAr$, SizeCn$, SizeEv$, SizeLe$, SizeLi$, AnzFolCn$, AnzFolEv$, AnzFolLe$
        z = Get_T1Row_HoldingMyText("Update Statistics")
        s = Get_T1Column_HoldingMyText("Update Statistics")
        With ActiveWorkbook.Sheets("T1")

    'Einige Vorab-Aktivitäten/-Einträge
        'Selection von CellButton wegändern
                EE 0: .Activate: .Cells(z - 1, s - 1).Select: EE 1
        'alte Einträge löschen (nur Zahlen und Häkchen)
                .Range(.Cells(z + 2, s + 11), .Cells(z + 32, s + 12)) = ""
        'z+1    [Date]
                T = Format(Now(), "dd.mm.yyyy hh:mm:ss") + " Uhr"
                stat(1) = T: .Cells(z + 1, s + 12) = T       'Neues Datum schreiben
        'z+2    [Ordner 'Archiv Trampolin'] [Größe]
                SizeAr = FolderSize(ArrC(2))  'Path of Folder "Archiv Trampolin"
                stat(2) = SizeAr: .Cells(z + 2, s + 11) = SizeAr
        'z+3    [UnterOrdner]
                stat(3) = "4": .Cells(z + 3, s + 11) = "'4" 'fixiert gesetzt
                .Cells(z + 3, s + 12) = "ü"     'Erledigt-Haken
        'z+4    [Photos  gesamt]    'wird in DoWhileSchleife gezählt
        'z+5    [Videos  gesamt]    'wird in DoWhileSchleife gezählt
        'z+6    [Links   gesamt]    'wird in DoWhileSchleife gezählt
        'z+7    [Dateien gesamt]    'wird in DoWhileSchleife gezählt
        'z+8    [Leerzeile]
        '------------------------
        'z+9    [UnterOrdner1: 'Events'] [Größe]
                SizeEv = FolderSize(ArrC(3))  'Path of Folder "Events"
                stat(9) = SizeEv: .Cells(z + 9, s + 11) = SizeEv
        'z+10   [Events] [Event-Ordner]
                    
                AnzFolEv = CStr(Get_Number_OfSubFolders_OneLevel(ArrC(3)))  'Path of Folder "Events"
                stat(10) = AnzFolEv: .Cells(z + 10, s + 11) = "'" + AnzFolEv
                .Cells(z + 10, s + 12) = "ü" 'Erledigt-Haken
        'z+11   [Events] [Photos]
        'z+12   [Events] [Ausschnitte]
        'z+13   [Events] [Videos]
        'z+14   [Events] [Teil-Videos]
        'z+15   [Events] [Snapshots]
        'z+16   [Events] [Links]
        '------------------------
        'z+17   [UnterOrdner2: 'ClubsNations']  [Größe]
                SizeCn = FolderSize(ArrC(11))  'Path of Folder "ClubsNations"
                stat(17) = SizeCn: .Cells(z + 17, s + 11) = SizeCn
        'z+18   [Club-/LTV-/Nation-Ordner] [Anzahl]
                AnzFolCn = CStr(Get_Number_OfSubFolders_OneLevel(ArrC(11)))
                stat(18) = AnzFolCn: .Cells(z + 18, s + 11) = "'" + AnzFolCn
                .Cells(z + 18, s + 12) = "ü" 'Erledigt-Haken
        'z+19   [ClubsNations] [Photos]
        'z+20   [ClubsNations] [Ausschnitte]
        'z+21   [ClubsNations] [Videos]
        'z+22   [ClubsNations] [Teil-Videos]
        'z+23   [ClubsNations] [Snapshots]
        'z+24   [ClubsNations] [Links]
        '------------------------
        'z+25   [UnterOrdner3: 'Leute']  [Größe]
                SizeLe = FolderSize(ArrC(4))  'Path of Folder "Leute"
                stat(25) = SizeLe: .Cells(z + 25, s + 11) = SizeLe
        'z+26   [Personen-Ordner]
                AnzFolLe = CStr(Get_Number_OfSubFolders_OneLevel(ArrC(4)))  'Path of Folder "Leute"
                stat(26) = AnzFolLe: .Cells(z + 26, s + 11) = "'" + AnzFolLe
                .Cells(z + 26, s + 12) = "ü" 'Erledigt-Haken
        'z+27   [Leute] [Photos]
        'z+28   [Leute] [Ausschnitte]
        'z+29   [Leute] [Snapshots]
        'z+30   [Leute] [Links]
        '------------------------
        'z+31   [UnterOrdner4: 'Register']  [Größe]
                SizeLi = FolderSize(ArrC(5))  'Path of Folder "Links"
                stat(31) = SizeLi: .Cells(z + 31, s + 11) = SizeLi
        'z+25   [Links]
        RefreshScreen
        End With
End Sub

Sub T1_UpdateStatistics2(stat() As String)
    'Action      Einträge genauer Zahlen in T1
    
    Dim i%, s%, z%, A(1 To 32, 1 To 2) As String
    z = Get_T1Row_HoldingMyText("Update Statistics")
    s = Get_T1Column_HoldingMyText("Update Statistics")
    
    'Array A für Paste erstellen
        For i = 2 To 32
            A(i, 1) = stat(i): A(i, 2) = "ü"
        Next
        A(1, 2) = stat(1): A(8, 1) = "": A(8, 2) = ""
    'Paste
        Paste_2DArrayToCell_z_s "T1", z + 1, s + 11, A
End Sub

Sub TEST_Update_ReadMe_ARCHIVTRAMPOLIN()
        Dim ReadMeNEW$
        ReadMeNEW = ReadFile(ArrC(7))       'Path of File   "Read me.txt"
        Update_ReadMe_ARCHIVTRAMPOLIN ReadMeNEW
        show ReadMeNEW

End Sub

Sub Update_ReadMe_ARCHIVTRAMPOLIN(ReadMeNEW$)
    'Called from    T1_UpdateStatistics
    'ReadMeNEW      trägt aktuell den Inhalt des bisherigen "Read me.txt"
    '               mit den Teilen ARCHIV TRAMPOLIN, PHOTOS/VIDEOS, QUELLEN, LEGENDE
    'Action         Aktualisiert den Bereich 'ARCHIV TRAMPOLIN' mit Daten aus ArrC
    '               (für die letztlich anstehende Überschreibung von 'Read me.txt')
    '               ReadMeNEW wird danach benutzt von 'Update_ReadMe_QUELLEN'
    
    'Vorbereitung
        Dim PartNew$, PartOld$, s$, stat$, c%, A() As String
    'STAT
        A = Split("|" + ArrC(32), "|") 'Daten von 'T1_UpdateStatistics'
    'PartOld
        c = InStr(1, ReadMeNEW, "Nachname1, Nachname2"): c = InStr(c, ReadMeNEW, "===")
        PartOld = Left(ReadMeNEW, c - 1)
        PartOld = Delete_EndReturnsInString(PartOld)
    'PartNew
        s = ReadFile(ArrC(6) + "\ReadMeVordruckArchivT.txt") 'ArrC(6) = Path of Folder "helpers"
        
        s = Replace(s, "{Date}", A(1))                          'Datum/Zeit Update
        s = Replace(s, "{SizeAr}", A(2))                        'Größe Folder "Archiv Trampolin"
        s = Replace(s, "{SizeEv}", A(9))                        'Größe Folder          "Events"
        s = Replace(s, "{SizeLe}", A(25))                        'Größe Folder          "Leute"
        s = Replace(s, "{SizeLi}", A(31))                        'Größe Folder          "Links"
        s = Replace(s, "{AnzPhotEv}", A(11))                     'Anzahl Photos      in "Events"
        s = Replace(s, "{AnzPhotLe}", A(27))                     'Anzahl Photos      in "Leute"
        s = Replace(s, "{AnzPhotGes}", A(4))                    'Anzahl Photos      gesamt
        s = Replace(s, "{AnzVidEv}", Right("   " + A(13), 4))    'Anzahl Videos      in "Events"
        s = Replace(s, "{AnzVidT}", Right("   " + A(14), 4))     'Anzahl Video-Teile in "Events"
        s = Replace(s, "{AnzVidGes}", A(5))                     'Anzahl Videos      gesamt
        s = Replace(s, "{AnzLnkEv}", Right("   " + A(16), 4))    'Anzahl Links       in "Events"
        s = Replace(s, "{AnzLnkLe}", Right("   " + A(30), 4))    'Anzahl Links       in "Leute"
        s = Replace(s, "{AnzLnkLi}", Right("   " + A(32), 4))    'Anzahl Links       in "Links"
        s = Replace(s, "{AnzLnkGes}", A(6))                     'Anzahl Links       gesamt
        s = Replace(s, "{AnzCropEv}", A(12))                     'Anzahl Ausschnitte in "Events"
        s = Replace(s, "{AnzCropLe}", A(28))                     'Anzahl Ausschnitte in "Leute"
        s = Replace(s, "{AnzSnapEv}", Right("   " + A(15), 4))   'Anzahl SnapShots   in "Events"
        s = Replace(s, "{AnzSnapLe}", Right("   " + A(29), 4))   'Anzahl SnapShots   in "Leute"
        s = Replace(s, "{AnzFoldEv}", A(10))                     'Anzahl Folder      in "Events"
        s = Replace(s, "{AnzFoldLe}", A(26))                     'Anzahl Folder      in "Leute"
        s = Replace(s, "{AnzFileGes}", A(7))                    'Anzahl Dateien     gesamt
        s = Replace(s, "{SizeCn}", A(17))
        s = Replace(s, "{AnzFoldCn}", A(18))
        s = Replace(s, "{AnzSnapCn}", Right("   " + A(23), 4))
        s = Replace(s, "{AnzLnkCn}", Right("   " + A(24), 4))
        s = Replace(s, "{AnzCropCn}", Right("   " + A(20), 4))
        s = Replace(s, "{AnzVidCn}", Right("   " + A(21), 4))
        s = Replace(s, "{AnzVidCnT}", Right("   " + A(22), 4))
        s = Replace(s, "{AnzPhotCn}", Right("   " + A(19), 4))
        s = Delete_EndReturnsInString(s)
        PartNew = s
    ReadMeNEW = Replace(ReadMeNEW, PartOld, s)
End Sub

Sub Show_ListOfFiles_DateLastModified()
    'Called from    xxx

    Dim oFSO As Object, oFolder As Object, oFile As Object, sF
    Dim colFolders As New Collection, F As Object, Fs As Object, i%
    Dim D$, D1$, s$, N$, PathOfFile$, PathOfSourceFolder$, Arr1() As String
    PathOfSourceFolder = ArrC(3)  'Path of Folder "Events"
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(PathOfSourceFolder)
    colFolders.Add oFolder          'start with this folder
    
    Do While colFolders.count > 0      'process all folders
        Set oFolder = colFolders(1)    'get a folder to process
        colFolders.Remove 1            'remove item at index 1
        For Each oFile In oFolder.Files
            N = oFile.NAME
            If N Like "*####-## [a-z][a-z].*" Then
                'File ist Photo oder Video
                PathOfFile = oFolder.path + "\" + N
                Set Fs = CreateObject("Scripting.FileSystemObject")
                Set F = Fs.GetFile(PathOfFile)
                D = F.DateLastModified
                D1 = Mid(D, 7, 4) + Mid(D, 4, 2) + Mid(D, 1, 2) _
                   + Mid(D, 12, 2) + Mid(D, 15, 2) + Mid(D, 18, 2)
                s = s + D1 + "   " + D + "   " + PathOfFile + vbCrLf
            End If
        Next oFile
        'add any subfolders to the collection for processing
        For Each sF In oFolder.subfolders
            colFolders.Add sF
        Next sF
    Loop
    s = Delete_EmptyEndRowsInString(s)
    Arr1 = Split(s, vbCrLf)
    QuickSortDes Arr1, 0, UBound(Arr1)
    For i = 0 To UBound(Arr1)
        Arr1(i) = Mid(Arr1(i), 18)
    Next
    s = "DateLastModified      PathOfFolder" + vbCrLf + vbCrLf
    show s + Join(Arr1, vbCrLf)
End Sub

Sub Show_ListOfFolders_DateLastModified()
    Dim D1$, D2$, p$, s$, i%, Arr1() As String, Arr2() As String
    s = Get_Paths_ofAllSubfoldersAllLevelsAsStringUseGlobalVar(ArrC(3))  'Path of Folder "Events"
    Arr1 = Split(s, vbCrLf)
    For i = 0 To UBound(Arr1)
        p = Arr1(i)     'One PathOfFolder
        D2 = CStr(FileDateTime(p))  '15.07.2023 15:03:27
        D1 = Mid(D2, 7, 4) + Mid(D2, 4, 2) + Mid(D2, 1, 2) _
           + Mid(D2, 12, 2) + Mid(D2, 15, 2) + Mid(D2, 18, 2)
        Arr1(i) = D1 + "   " + D2 + "   " + p
    Next
    QuickSortDes Arr1, 0, UBound(Arr1)
    For i = 0 To UBound(Arr1)
        Arr1(i) = Mid(Arr1(i), 18)
    Next
    s = "DateLastModified      PathOfFolder" + vbCrLf + vbCrLf
    show s + Join(Arr1, vbCrLf)
    Beep
End Sub

Function Get_ShortShorty(s1$) As String
    Dim s$, i%, Arr1() As String
    Arr1 = Split(s1, vbCrLf)
    For i = 0 To UBound(Arr1)
        If Left(Arr1(i), 3) Like "[a-z][a-z] " Then
            s = s + "|" + Left(Arr1(i), 2) + "|"
        End If
    Next
    s = Replace(s, "||", "|")
    Get_ShortShorty = s
End Function

Sub Format_Shorty_TEST()
    Dim s$, A() As String
    s = "ab.jpg 0252|cm.jpg 0026|cs.jpg 0023|cs.mp4 0001|ds.jpg 0088|fb.jpg 0015|fh.jpg 0024|fh.mp4 0006|hp.jpg 0001|ik.jpg 0054|kf.avi 0002|kf.jpg 0020|ms.jpg 0077|nn.avi 0002|nn.jpg 0688|nn.mp4 0008|nn.wmv 0001|se.jpg 0687|sf.avi 0037|sf.jpg 2081|sy.jpg 0181|th.jpg 0107|ub.avi 0004|ub.dng 0019|ub.jpg 1801|ul.jpg 0095|ul.mp4 0001|ww.jpg 0863|ww.mkv 0050|ww.mp4 0104|ww.mpg 0001|ww.png 0001|ww.wmv 0001|yw.jpg 0003"
    A = Split(s, "|"): s = Join(A, vbCrLf)
    s = Format_Shorty(s)
    show s
End Sub

Function Format_Shorty(Shorty$) As String
    'Called from    Update_ReadMe_QUELLEN
    'Shorty [anfangs]   ... [v] nn.avi 0002 [v] nn.jpeg 0688 [v] nn.mp4 0008 [v] ... (sortiert)
    'Shorty [am Ende]   Zeilen wie "ab  251 Photos", "ww   78 Photos,   14 Videos" [keine Leerzeilen]

    'Vorbereitung
        Dim ext$, pho$, PV$, L$, k$, T$, v$, vID$
        Dim c%, i%, j%, Arr1() As String, Arr2() As String, B()
        v = vbCrLf
    'Extensions
        pho = ";ani;bmp;dng;gif;heic;ico;jpe;jpeg;jpg;pcx;png;psd;tga;tif;tiff;webp;wmf;"
        vID = ";3g2;3gp;3gp2;3gpp;amr;amv;asf;avi;bdmv;bik;d2v;divx;drc;dsa;dsm;dss;dsv;evo;f4v;flc;fli;flic;flv;hdmov;ifo;ivf;m1v;m2p;m2t;m2ts;m2v;m4v;mkv;mp2v;mp4;mp4v;mpe;mpeg;mpg;mpls;mpv2;mpv4;mov;mts;ogm;ogv;pss;pva;qt;ram;ratdvd;rm;rmm;rmvb;roq;rpm;smil;smk;swf;tp;tpr;ts;vob;vp6;webm;wm;wmp;wmv;"
    'Get    |nn|video|0002|   |nn|photo|0688|   |nn|video|0008|
        Arr1 = Split(Shorty, v)
        For i = 0 To UBound(Arr1)
            PV = "xxxxx"                    'Extension unbekannt
            L = Arr1(i)                    'ab.jpg 0252    'ab.jpeg 0252
            ext = Mid(L, 4, Len(L) - 8)   'jpg            'jpeg
            If pho Like "*;" + ext + ";*" Then PV = "photo"
            If vID Like "*;" + ext + ";*" Then PV = "video"
            Arr1(i) = "|" + Left(L, 2) + "|" + PV + "|" + Right(L, 4) + "|"
        Next
    'Get    Arr1: |nn|00000|0002|   |nn|0688|00000|   |nn|00000|0008|
        For i = 0 To UBound(Arr1)
            L = Arr1(i)                    '|nn|video|0002|   '|nn|photo|0688|
            If Mid(L, 5, 5) = "photo" Then Arr1(i) = Left(L, 4) + Mid(L, 11, 5) + "0000|"
            If Mid(L, 5, 5) = "video" Then Arr1(i) = Left(L, 4) + "0000|" + Right(L, 5)
            If Mid(L, 5, 5) = "xxxxx" Then Arr1(i) = Left(L, 4) + "0000|0000|"
            'K = Kürzel = Sammlung Shorty
                If Not k Like "*" + Left(L, 4) + "*" Then k = k + Left(L, 4)
        Next
        'showArray Arr1: Stop
        k = Replace(k, "||", "|") '= "|ab|cm|cs|ds|fb|fh|hp|ik|kf|ms|nn|se|sf|sy|th|ub|ul|ww|yw|"
        'show K: Stop
    'B() anlegen; 2D-Array
        c = anzAinB("|", k) - 1: ReDim B(1 To c, 1 To 3)
    'B() mit K füllen
        Arr2 = Split(k, "|")
        For i = 1 To c
            B(i, 1) = Arr2(i): B(i, 2) = "0000": B(i, 3) = "0000"
        Next
        'showArray2D B: Stop '|nn|0000|0000|
    'B() mit SumPhotos|SumVideos füllen
        For i = 0 To UBound(Arr1)
            L = Arr1(i)                     '|nn|0000|0000|
            Arr2 = Split(L, "|")            'Neubelegung von Arr2
            For j = 1 To c
                If B(j, 1) = Arr2(1) Then
                    B(j, 2) = Format(CInt(B(j, 2)) + CInt(Arr2(2)), "0000")
                    B(j, 3) = Format(CInt(B(j, 3)) + CInt(Arr2(3)), "0000")
                End If
            Next
        Next
        'showArray2D B: Stop '|nn|0688|0010| = |Shorty|SumPhotos|SumVideos|
    'B() in Text umwandeln
        For i = 1 To c
            't = t + B(i, 1) + " xxx " + CStr(B(i, 2)) + " Fotos, " + CStr(B(i, 3)) + " Videos" + v
            T = T + B(i, 1) + " " + CStr(B(i, 2)) + " Fotos, " + CStr(B(i, 3)) + " Videos" + v
        Next
        T = Replace(T, ", 0000 Videos", "")
        T = Replace(T, "0000 Fotos,", "            ")
        T = Replace(T, "0001 Fotos,", "0001 Foto, ")
        T = Replace(T, "0001 Fotos", "0001 Foto")
        T = Replace(T, "0001 Videos", "0001 Video")
        T = Replace99(T, " 0", "  ")
        'show t ': Stop
    Format_Shorty = T
End Function

Sub Mark_T1CellButtons_OpenShow()
    'Called from    Worksheet_SelectionChange
    'Action         Erledigt-Häckchen bei CellButtons setzen/entfernen -
    '               je nachdem, ob ein Ordner/eine Datei gerade geöffnet ist oder nicht
    
    'Vorbereitung
        Dim w$, z%
        With ActiveWorkbook.Sheets("T1")
    'WindowNames
        w = Get_NamesOfAllOpenWindows
    'Häkchen
        z = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "'Events'")
            If InStr(1, w, "Events") > 0 Then .Cells(z, 31) = "ü" Else .Cells(z, 31) = ""
        z = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "Clubs")
            If InStr(1, w, "Clubs") > 0 Then .Cells(z, 31) = "ü" Else .Cells(z, 31) = ""
        z = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "'Leute'")
            If InStr(1, w, "Leute") > 0 Then .Cells(z, 31) = "ü" Else .Cells(z, 31) = ""
        z = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "'Register'")
            If InStr(1, w, "Register") > 0 Then .Cells(z, 31) = "ü" Else .Cells(z, 31) = ""
        z = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "'ToDo'")
            If InStr(1, w, "ToDo") > 0 Then .Cells(z, 31) = "ü" Else .Cells(z, 31) = ""
        z = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "Read me")
            If InStr(1, w, "Read me") > 0 Then .Cells(z, 31) = "ü" Else .Cells(z, 31) = ""
        z = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "Logbuch")
            If InStr(1, w, "Logbuch") > 0 Then .Cells(z, 31) = "ü" Else .Cells(z, 31) = ""
        z = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "Terminkalender")
            If InStr(1, w, "Terminkalender") > 0 Then .Cells(z, 31) = "ü" Else .Cells(z, 31) = ""
        z = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "Competitors")
            If InStr(1, w, "Competitors") > 0 Then .Cells(z, 31) = "ü" Else .Cells(z, 31) = ""
        z = Get_RowNr_HoldingMyTextPartInColumnX("T1", 19, 2, "PersonData")
            If InStr(1, w, "PersonData") > 0 Then .Cells(z, 31) = "ü" Else .Cells(z, 31) = ""
    'Finals
        End With
End Sub

Sub T1_ShowReadMeTxt(z%, s%)
    DoArr
    openFile ArrC(7)  'Path of File   "Read me.txt"
    ActiveWorkbook.Sheets("T1").Cells(z, s + 12) = "ü": RefreshScreen
    Application.Wait (Now + TimeValue("0:00:01"))
    ActiveWorkbook.Sheets("T1").Cells(z - 1, s - 1).Select
End Sub

Sub Wait1Sec()
    Application.Wait (Now + TimeValue("0:00:01"))
End Sub

Sub WaitSecs(N%)
    Application.Wait (Now + TimeSerial(0, 0, N))
End Sub

Sub T1_ShowLogBuch()
    'Called from:   Worksheet_SelectionChange[T1], T1_UpdateStatistics
    'Status         Mit 'Logbuch T' wird die TextZeile T der Datei Logbuch.txt angefügt
    'Action         Logbuch.txt enthält die neuesten Einträge; ggf. mehrere Monate;
    '               falls 'mehrere Monate', wird je ein MonatsLogbuch erstellt,
    '               wobei nur der jüngste Monat in 'Logbuch.txt' verbleibt

    'Vorbereitung
        Dim A$, D1$, D2$, D3$, D4$, L$, m$, M1$, N$, T$, TZ$, v$, p$
        Dim c%, i%, z%, s%, Arr1() As String
        v = vbCrLf: DoArr: p = ArrC(1) + "\prog\Logbuch\"
        L = ReadFile(p + "Logbuch.txt")
        z = Get_T1Row_HoldingMyText("Show                          'Logbuch.txt'")
        s = Get_T1Column_HoldingMyText("Show                          'Logbuch.txt'")
        ActiveWorkbook.Sheets("T1").Cells(z, s + 12) = "ü": RefreshScreen
    'MonatsLogbuch erstellen
        Arr1 = Split(L, v): D2 = "": D4 = ""
        For i = 0 To UBound(Arr1)
            'Schleife über alle Logbuch-Zeilen
            TZ = Arr1(i)                        'TZ = 1 TextZeile
            If TZ Like "########_*" Then
                D1 = Left(TZ, 6)                'd1 = YearMonth = "202310"
                If D2 = "" Then D2 = D1         'd2 = Monat der aktuellen Sammlung
                If Len(TZ) > c Then c = Len(TZ) 'c  = MaxZeilenLänge
                If D1 = D2 Then
                    'Gleicher Monat                    '
                    D3 = Left(TZ, 8)            'd3 = 8erDatum
                    If D4 = "" Then D4 = D3     'd4 = Tag der aktuellen Sammlung
                    If D3 = D4 Then
                        'Gleicher Tag
                        m = m + v + TZ
                    Else
                        m = m + v + v + TZ
                        D4 = D3
                    End If
                Else
                    A = String(c, "-")
                    m = A + v + Replace(m, v + v, v + A + v) + v + A + v
                    'show "Logbuch_" + d2 + v + v + M
                    writeStringToFile p + "Logbuch_" + D2 + ".txt", "Logbuch_" + D2 + v + v + m
                    m = TZ: c = 0: D2 = D1 'Neuer Monat
                End If
            End If
        Next
        A = String(c, "-"): m = A + Replace(m, v + v, v + A + v) + v + A
        m = "Logbuch des aktuellen Monats " + Mid(D1, 5, 2) + "." + Mid(D1, 1, 4) + v + v + m
        'show M
        writeStringToFile p + "Logbuch.txt", m
    'LogBuch öffnen
        Application.Wait (Now + TimeValue("0:00:01"))
        openFile p + "Logbuch.txt"
    'T1-Anzeige (Häkchen und ChangeSelect)
        N = ActiveSheet.NAME
        With ActiveWorkbook
        If N = "T1" Then
            .Sheets("T1").Cells(z, s + 12) = "ü": RefreshScreen
            .Sheets("T1").Cells(z - 1, s - 1).Select
        Else
            Application.ScreenUpdating = False
            .Sheets("T1").Activate
            .Sheets("T1").Cells(z, s + 12) = "ü": RefreshScreen
            .Sheets("T1").Cells(z - 1, s - 1).Select
            .Sheets(N).Activate
            Application.ScreenUpdating = True
        End If
        End With
End Sub

Sub T1_ShowTerminkalenderTxt(z%, s%)

    show "Textversion des Terminkalenders ist noch in Arbeit."
    Exit Sub
    
    'Vorbereitung
        DoArr
    'Terminkalender.txt öffnen
        openFile ArrC(2) + "\z Terminkalender.txt"
    'T1-Anzeige (Häkchen und ChangeSelect)
        If ActiveSheet.NAME = "T1" Then
            ActiveWorkbook.Sheets("T1").Cells(z, s + 12) = "ü": RefreshScreen
            ActiveWorkbook.Sheets("T1").Cells(z - 1, s - 1).Select
        ElseIf ActiveSheet.NAME = "T2" Then
            z = Get_RowNr_HoldingMyTextWhole("T2", "Show 'Terminkalender.txt'")
            s = Get_ColumnNr_HoldingMyTextWhole("T2", "Show 'Terminkalender.txt'")
            ActiveWorkbook.Sheets("T2").Cells(z, s - 2).Select
        End If
End Sub

Sub T1_ShowCompetitorsTxt()
    'Vorbereitung
        Dim s%, z%
        DoArr
    'Competitors.txt öffnen
        openFile ArrC(2) + "\Competitors.txt"
    'T1-Anzeige (Häkchen und ChangeSelect)
        If ActiveSheet.NAME = "T1" Then
            z = Get_T1Row_HoldingMyText("Show                  'Competitors.txt'")
            s = Get_T1Column_HoldingMyText("Show                  'Competitors.txt'")
            ActiveWorkbook.Sheets("T1").Cells(z, s + 12) = "ü": RefreshScreen
            ActiveWorkbook.Sheets("T1").Cells(z - 1, s - 1).Select
        End If
End Sub

Sub XBox_CodeRechts(z%, s%)
    'Called from:   Worksheet_SelectionChange [in T1]
    Dim w!, H!
    Application.EnableEvents = False
    With Sheets("T1")
        If .Cells(z, s) = "X" Then
            .Cells(z, s) = ""
            Application.WindowState = xlMaximized
        Else
            .Cells(z, s) = "X"
            With Application
                .ScreenUpdating = False
                .WindowState = xlMaximized
                w = .Width: H = .Height
                .WindowState = xlNormal
                .Top = 0: .Left = 0
                .Width = w / 2 + 10
                .Height = H - 10
                With .VBE.MainWindow
                    .Top = 0
                    .Left = 970     'Screen Pixel
                    .Width = 960    'Screen Pixel
                    .Height = 1032  'Screen Pixel
                    .Visible = True
                End With
                .ScreenUpdating = True
            End With
        End If
        .Activate
        .Cells(z - 5, s - 7).Select
    End With
    Application.EnableEvents = True
End Sub

Sub XBox_Gitterlinien(z%, s%)
    'Called from:   Worksheet_SelectionChange [in T1]

    Application.EnableEvents = False
    With Sheets("T1")
        If .Cells(z, s) = "X" Then
            .Cells(z, s) = "":  ActiveWindow.DisplayGridlines = 0
        Else
            .Cells(z, s) = "X": ActiveWindow.DisplayGridlines = 1
        End If
        .Cells(z - 5, s - 1).Select
    End With
    Application.EnableEvents = True
End Sub

Sub ToolBarRibbonShow()
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",true)"
End Sub

Sub ToolBarRibbonHide()
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",false)"
End Sub

Sub ToolBarRibbonMinimize()
    If CommandBars("Ribbon").Height > 150 Then CommandBars.ExecuteMso "MinimizeRibbon"
End Sub

Sub ToolBarRibbonMaximize()
    If CommandBars("Ribbon").Height < 150 Then CommandBars.ExecuteMso "MinimizeRibbon"
End Sub

Sub XBox_MenüEin(z%, s%)
    'Called from:   Worksheet_SelectionChange [in T1]
    
    Application.EnableEvents = False
    With Sheets("T1")
        If .Cells(z, s) = "X" Then
            .Cells(z, s) = "":  ToolBarRibbonHide
        Else
            .Cells(z, s) = "X": ToolBarRibbonShow
        End If
        .Cells(z - 3, s - 1).Select
    End With
    Application.EnableEvents = True
End Sub

Sub XBox_MenüKlein(z%, s%)
    'Called from:   Worksheet_SelectionChange [in T1]
    
    Application.EnableEvents = False
        With Sheets("T1")
        If .Cells(z, s) = "X" Then
            .Cells(z, s) = "":  ToolBarRibbonMaximize
        Else
            .Cells(z, s) = "X": ToolBarRibbonMinimize
        End If
        .Cells(z - 3, s - 7).Select
    End With
    Application.EnableEvents = True
End Sub

Sub T1_Save_a_copy(z%, s%)
    'Called from:   Worksheet_SelectionChange [in T1]
    'z              T1-ZeilenNr  des CellButtons "Save a copy"
    's              T1-SpaltenNr des CellButtons "Save a copy"
    
    Dim p$, qq$, T$
    Application.EnableEvents = False
    qq = Chr(34) 'quote, "-Zeichen
    T = Format(Now(), "yyyymmdd_hhmmss")
    p = ActiveWorkbook.path + "\old - Bes TR.xlsm\Bes TR " & T + ".xlsm"
    ActiveWorkbook.SaveCopyAs p
    With ActiveWorkbook.Sheets("T1")
        .Cells(z + 1, s) = "Bes TR " & T + ".xlsm"
        .Cells(z - 1, s - 1).Select
    End With
    LogBuch "xls-Copy was saved: " + p
    Application.EnableEvents = True
    Beep '800, 100
End Sub

Sub ZeilenHoehe15()
    Dim r As Range
    Set r = Sheets("T1").UsedRange: r.RowHeight = 15
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Get_FilePaths_Like_insideSourceFolderAndSubFolders_TEST()
    Dim F$, p$: DoArrc: p = ArrC(3) 'Events
    F = Get_FilePaths_Like_insideSourceFolderAndSubFolders(p, "*[#]result*")
    show F
End Sub

Function Get_FileNames_Like_insideSourceFolderAndSubFolders(PathOfSourceFolder$, LikeString$) As String
    Dim oFSO As Object, oFolder As Object, oFile As Object, sF
    Dim s1$, AnzFold%, i&, colFolders As New Collection
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(PathOfSourceFolder)
    colFolders.Add oFolder          'start with this folder
    AnzFold = colFolders.count
    Do While colFolders.count > 0      'process all folders
        Set oFolder = colFolders(1)    'get a folder to process
        colFolders.Remove 1            'remove item at index 1
        For Each oFile In oFolder.Files
            If oFile.NAME Like LikeString Then
                s1 = s1 + oFile.NAME + vbCrLf
            End If 'i = i + 1
        Next oFile
        'add any subfolders to the collection for processing
        For Each sF In oFolder.subfolders
            colFolders.Add sF
        Next sF
    Loop
    Get_FileNames_Like_insideSourceFolderAndSubFolders = s1
End Function

Public Function Get_AllFileNames_Like_OfOneFolder(PathOfOneFolder$, MyLikeString$) As String
    'Vorbereitung
        Dim fso As Object, folder As Object, file As Object, results() As String, count As Long
        Set fso = CreateObject("Scripting.FileSystemObject")
    'Falls Pfad nicht existiert, leeren String zurückgeben
        If Not fso.FolderExists(PathOfOneFolder) Then
            Get_AllFileNames_Like_OfOneFolder = ""
            show "Pfad existiert nicht: " + PathOfOneFolder
            Stop: Exit Function
        End If
    'FolderObject
        Set folder = fso.GetFolder(PathOfOneFolder)
        count = 0
    'Array dynamisch füllen
        For Each file In folder.Files
            If file.NAME Like MyLikeString Then
                ReDim Preserve results(count)
                results(count) = file.NAME
                count = count + 1
            End If
        Next file
    'Array zu einem String zusammenfügen
        If count > 0 Then
            Get_AllFileNames_Like_OfOneFolder = Join(results, vbCrLf)
        Else
            Get_AllFileNames_Like_OfOneFolder = ""
        End If
End Function

Function Get_FirstFileName_Like_OfOneFolder(PathOfOneFolder$, LikeString$)
    If PathOfOneFolder = "" Then Exit Function
    Dim fso As Object, objVerzeichnis As Object, objDateienliste As Object
    Dim objDatei As Object, s$
    Set fso = CreateObject("scripting.FileSystemObject")
    Set objVerzeichnis = fso.GetFolder(PathOfOneFolder)
    Set objDateienliste = objVerzeichnis.Files
    For Each objDatei In objDateienliste
         If Not objDatei Is Nothing Then
            If objDatei.NAME Like LikeString Then
              Get_FirstFileName_Like_OfOneFolder = objDatei.NAME: Exit Function
            End If
         End If
    Next objDatei
End Function

Function Get_FileNames_Like_NotLike_insideOneFolder(PathOfOneFolder$, LikeStringForFileName$, NotLikeStringForFileName$) As String
    If PathOfOneFolder = "" Then Exit Function
    Dim fso As Object, Fol As Object, Fils As Object
    Dim Fil As Object, s$
    Set fso = CreateObject("scripting.FileSystemObject")
    Set Fol = fso.GetFolder(PathOfOneFolder)
    Set Fils = Fol.Files
    For Each Fil In Fils
         If Not Fil Is Nothing Then
            If Fil.NAME Like LikeStringForFileName And Not Fil.NAME Like NotLikeStringForFileName Then
              s = s + Fil.NAME + vbCrLf
            End If
         End If
    Next Fil
    s = Delete_EmptyEndRowsInString(s)
    Get_FileNames_Like_NotLike_insideOneFolder = s
End Function

Function Get_FilePaths_Like_NotLike_insideSourceFolderAndSubFolders(PathOfSourceFolder$, LikeStringForFileName$, NotLikeStringForFileName$) As String
    Dim oFSO As Object, oFolder As Object, Fil As Object, sF
    Dim s1$, i&, colFolders As New Collection
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(PathOfSourceFolder)
    colFolders.Add oFolder          'start with this folder
    Do While colFolders.count > 0      'process all folders
        Set oFolder = colFolders(1)    'get a folder to process
        colFolders.Remove 1            'remove item at index 1
        For Each Fil In oFolder.Files
            If Fil.NAME Like LikeStringForFileName _
                And Not Fil.NAME Like NotLikeStringForFileName Then
                s1 = s1 + oFolder.path + "\" + Fil.NAME + vbCrLf '+ CStr(Fil.DateLastModified) + vbCrLf
            End If 'i = i + 1
        Next Fil
        'add any subfolders to the collection for processing
        For Each sF In oFolder.subfolders
            colFolders.Add sF
        Next sF
    Loop
    Get_FilePaths_Like_NotLike_insideSourceFolderAndSubFolders = s1
End Function

Function Get_FilePaths_Like_insideSourceFolderAndSubFolders(PathOfSourceFolder$, Optional LikeStringForFilePath$ = "*", Optional CountDown As Boolean = False) As String
    Dim oFSO As Object, oFolder As Object, oFile As Object, sF
    Dim s1$, AnzFold%, i&, colFolders As New Collection
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(PathOfSourceFolder)
    colFolders.Add oFolder          'start with this folder
    AnzFold = colFolders.count
    If CountDown Then MyCountDown_OneMoreMainStep
    If CountDown Then MyCountDown_MainStepsAllowed 20: MyCountDown_SubStepsMax AnzFold
    Do While colFolders.count > 0      'process all folders
        Set oFolder = colFolders(1)    'get a folder to process
        colFolders.Remove 1            'remove item at index 1
        If CountDown Then MyCountDown_OneMoreSubStep
            'to see countdown:
                'Sheets("T4").Cells(21, 22) = colFolders.count
        For Each oFile In oFolder.Files
            If oFolder.path + "\" + oFile.NAME Like LikeStringForFilePath Then
                s1 = s1 + oFolder.path + "\" + oFile.NAME + vbCrLf '+ CStr(oFile.DateLastModified) + vbCrLf
            End If 'i = i + 1
        Next oFile
        'add any subfolders to the collection for processing
        For Each sF In oFolder.subfolders
            colFolders.Add sF
        Next sF
    Loop
    Get_FilePaths_Like_insideSourceFolderAndSubFolders = s1
End Function

Sub Show_3DatesOfEveryFile_inFolderAndSubFolders()
    Dim p1$, p2$, s1$, s2$, s3$, i%, Arr1() As String
    'p1 = "G:\Archiv Photos\- Photos ubes\Photos BR\2023 BR\20230100 BR"
    p1 = "G:\Archiv Photos\- Photos ubes\Photos BR\1947 BR"
    s1 = Get_FilePaths_Like_insideSourceFolderAndSubFolders(p1)
    Arr1 = Split(s1, vbCrLf)
    For i = 0 To UBound(Arr1)
        p2 = Arr1(i) 'PathOfOneFile
        If Right(LCase(p2), 4) = ".jpg" Then
            s3 = Get_DateInFileHeader(p2)
                If s3 = "" Then s3 = "---------------"
                s3 = s3 + "  DateInFileHeader"
            s2 = s2 + p2 + vbCrLf _
                + Get_NameFromPath(p2) + vbCrLf + s3 + vbCrLf _
                + GetFileDateInfo(p2) + vbCrLf + vbCrLf
        End If
    Next
    show s2
End Sub

Function GetFileDateInfo(PathOfFile$)
    Dim Fs As Object, F, s$, s1$, s2$
    Set Fs = CreateObject("Scripting.FileSystemObject")
    Set F = Fs.GetFile(PathOfFile)
    s = F.DateCreated
    s = Replace(s, ":", "")
    s1 = Mid(s, 7, 4) + Mid(s, 4, 2) + Left(s, 2) + "_" + Mid(s, 12) + "  DateCreated" & vbCrLf
    s = F.DateLastModified
    s = Replace(s, ":", "")
    s2 = Mid(s, 7, 4) + Mid(s, 4, 2) + Left(s, 2) + "_" + Mid(s, 12) + "  Last Modified"
    GetFileDateInfo = s1 + s2
End Function

Function Get_DateInFileHeader(PathOfFile$) As String
    Dim p$, s$, s1$, sDate$, i%
    i = 0
    Open PathOfFile For Binary As #1
    Do While i < 1000
        Line Input #1, s
        i = i + 1
        If InStr(1, s, ":") > 0 Then
            On Error Resume Next
            sDate = Mid(s, InStr(1, s, ":") - 4, 19)
            If sDate Like "####:##:## ##:##:##" Then
                s1 = sDate: Exit Do
            End If
        End If
    Loop
    Close #1
    If s1 <> "" Then
        s1 = Replace(s1, ":", "")
        s1 = Replace(s1, " ", "_")
        Get_DateInFileHeader = s1
    End If
End Function

Sub Show_Infotipp_InformationenOfAllJpgs_InsideOneFolder()
    Dim s$, s1$, PathOfSourceFolder As Variant, objShell As Object, objFolder As Object, strFileName As Object
    PathOfSourceFolder = "G:\Archiv Photos\- Photos ubes\Photos BR\1976 BR"
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(PathOfSourceFolder)
    
    'Loop through all Items (files) in folder "C:Files"
    For Each strFileName In objFolder.items
        'Look only for image files
        If InStr(objFolder.GetDetailsOf(strFileName, 0), "jpg") Then
            '(0) Ruft den Namen des Elements ab.
            '(1) Ruft die Größe des Elements ab.
            '(2) Ruft den Typ des Elements ab.
            '(3) Ruft das Datum und die Uhrzeit der letzten Änderung des Elements ab.
            '(4) Ruft die Attribute des Elements ab.
            '(-1) Ruft die Infotippinformationen für das Element ab.
            'If Len(objFolder.GetDetailsOf(strFileName, -1)) > 0 Then
                s1 = objFolder.GetDetailsOf(strFileName, -1)
                's1 = objFolder.GetDetailsOf(strFileName, 4)
                s = s + PathOfSourceFolder + "\" + strFileName + vbCrLf _
                + s1 + vbCrLf + vbCrLf
            'End If
        End If
    Next
    show s
End Sub

Public Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If x > 2 And y > 2 And x < Label1.Width - 2 And y < Label1.Height - 2 Then
    Label1.BackColor = 255
    Else
    Label1.BackColor = &H80000005
    End If
End Sub

'-------------------------------------------- 'MODUL: M01
Public Function getShortName(ByVal longName As String) As String
    ' Ermittelt den kurzen Dateinamen
     Dim ShortName As String, sLen As Long
     ShortName = Space$(256)
     sLen = GetShortPathName(longName, ShortName, Len(ShortName))
     ShortName = Left$(ShortName, sLen)
     getShortName = ShortName
End Function

'-------------------------------------------- 'MODUL: M01
Sub LogBuch(s As String)
    Dim p As String
    Call DoArr: p = ArrC(1) + "\prog\Logbuch\Logbuch.txt"
    WriteStringToFileAppend p, Format(Now(), "YYYYMMDD_hhmmss") + " " + s
End Sub

'-------------------------------------------- 'MODUL: M01
Sub TextdateiStarten(Pfad)
    shell "cmd /c " & Chr(34) & Pfad & Chr(34)
End Sub

'-------------------------------------------- 'MODUL: M01
Function Get_ParentFolderPath_OfFolderPath(ordnerPfad As String) As String
    'wird in <Sub AuswahlOrdnerFotoQuelle_Click> verwendet
    Dim i1 As Integer, i2 As Integer, s As String
    s = ordnerPfad
    If Right(s, 1) = "\" Then s = Left(s, Len(s) - 1)       's = "C:\Archiv Photos\MyPhotos\Urlaub"
    i2 = InStrRev(ordnerPfad, "\"):                         'i2: Position des 1. "\" von rechts
    If i2 = 0 Then
        Exit Function                                       'd. h. getParentFolderPath = ""
    Else
        Get_ParentFolderPath_OfFolderPath = Left(s, i2 - 1)
    End If
End Function

Sub LoescheDatei(PathOfFile)
   On Error Resume Next
   Kill PathOfFile
End Sub

Public Function OrdnerExistiert(PathOfFolder As String) As Boolean
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    OrdnerExistiert = objFSO.FolderExists(PathOfFolder)
    Set objFSO = Nothing
End Function

Function CreateFolder(PathOfFolder$)
    'requires reference to Microsoft Scripting Runtime
    Dim fso As New FileSystemObject
    If Not fso.FolderExists(PathOfFolder) Then
        'show PathOfFolder
        fso.CreateFolder PathOfFolder
        'Stop
    End If
End Function

Sub Create_FoldersUpToOneFilePath(PathOfFile$)
    Dim lCtr As Long, arrpath() As String, PathOfFolder As String, p2 As String
    PathOfFolder = Left(PathOfFile, InStrRev(PathOfFile, "\") - 1)
    arrpath = Split(PathOfFolder, "\")
    p2 = arrpath(LBound(arrpath)) & "\"
    For lCtr = LBound(arrpath) + 1 To UBound(arrpath)
        p2 = p2 & arrpath(lCtr) & "\"
        If Dir(p2, vbDirectory) = "" Then MkDir p2
    Next
End Sub

Function Create_FoldersUpToOneFolderPath(path As String) As Boolean
    'Action     Erzeugt ggf. neuen Ordner (returns TRUE if folder was created)
    If path = "" Then Stop  'path sollte nicht "" sein
    Dim fso As New FileSystemObject
    'If the path exists as a file, the function fails.
        'If Right(path, 1) = "\" Then path = Left(path, Len(path) - 1)
        If fso.FileExists(path) Then
            Create_FoldersUpToOneFolderPath = False
            Exit Function
        End If
    'If the path already exists as a folder, don't do anything and return success.
        If fso.FolderExists(path) Then
            Create_FoldersUpToOneFolderPath = True
            Exit Function
        End If
    'recursively create the parent folder, then if successful create the top folder.
        If Create_FoldersUpToOneFolderPath(fso.GetParentFolderName(path)) Then
            If fso.CreateFolder(path) Is Nothing Then
                Create_FoldersUpToOneFolderPath = False
            Else
                Create_FoldersUpToOneFolderPath = True
            End If
        Else
            Create_FoldersUpToOneFolderPath = False
        End If
End Function

'-------------------------------------------- 'MODUL: M01
Function korrigiereFaceName(faceName As String) As String
    Dim s As String ', c1 As Integer
    s = faceName
    s = Replace(s, "*", ""): s = Replace(s, "|", "")            'kritische Zeichen aus faceName entfernen
    s = Replace(s, "  ", " ")                                   'Doppel-Leerzeichen löschen
    If Right(s, 1) = " " Then s = Left(s, Len(s) - 1)           'Leerzeichen am Ende von faceName löschen
    's = Replace(s, " (", "(")
    'c1 = InStr(1, s, "("): If c1 > 0 Then s = Left(s, c1 - 1)   'faceName-ZeichenLänge kürzen; "(..." weglassen
                                                                'manche faceNames unterscheiden sich nur MIT "(..."
    korrigiereFaceName = s
End Function

Function Get_FilePath_OfToolES() As String
    Get_FilePath_OfToolES = ArrC(8)  'Path of Tool   "es.exe"
End Function

Function Get_FilePath_OfExifTool() As String
    Get_FilePath_OfExifTool = ArrC(9)  'Path of Tool  "exiftool.exe"
End Function

'-------------------------------------------- 'MODUL: M01
Function Get_NameFromPath(PathOfFileOrFolder As String) As String
    Dim s$: s = PathOfFileOrFolder
    If Right(s, 1) = "\" Then s = Left(s, Len(s) - 1)
    s = Mid(s, InStrRev(s, "\") + 1)
    Get_NameFromPath = s
End Function

'-------------------------------------------- 'MODUL: M01
Sub tryRename(PathOfFile1 As String, PathOfFile2 As String)
    On Error Resume Next
    Name PathOfFile1 As PathOfFile2
End Sub

'-------------------------------------------- 'MODUL: M01
Sub writeAllVbaCodeToTxtFile()
    Dim VBComp As Object, objFile As Object
    Dim i As Long, TaskID#, T As String, s As String, s1 As String

    T = ThisWorkbook.path & "/" & ThisWorkbook.NAME & "_VBA.txt"
    Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(T, 2, True)
    s = "Gesamter VBA-Code (Makros/Prozeduren aller Module) der Excel-Datei <" & ThisWorkbook.NAME & ">" & vbCrLf
    s = s & "****************************************************************************************" & vbCrLf
    objFile.WriteLine s 'schreibt s in die 1. Zeile der Textdatei

    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        objFile.WriteLine "'##############################" & vbCrLf & "'MODUL: " & VBComp.NAME & vbCrLf & "'##############################"
        With VBComp.CodeModule
            i = 1: s1 = "'----------------------------------------- "
            s1 = s1 & "Quelle: " & ThisWorkbook.NAME & " --- " & "'MODUL: " & VBComp.NAME
            Do Until i > .CountOfLines
                If Left(.lines(i, 1), 4) = "Sub " Then objFile.WriteLine s1
                If Left(.lines(i, 1), 9) = "Function " Then objFile.WriteLine s1
                If Left(.lines(i, 1), 12) = "Private Sub " Then objFile.WriteLine s1
                If Left(.lines(i, 1), 17) = "Private Function " Then objFile.WriteLine s1
                If Left(.lines(i, 1), 11) = "Public Sub " Then objFile.WriteLine s1
                If Left(.lines(i, 1), 16) = "Public Function " Then objFile.WriteLine s1
                objFile.WriteLine .lines(i, 1)
                i = i + 1
            Loop
        End With
        objFile.WriteLine
    Next VBComp
    objFile.Close
    'TaskId = Shell("notepad.exe " & t, vbNormalFocus) 'öffnet sofort die Text-Datei
End Sub

'-------------------------------------------- 'MODUL: M01
Sub Get_NumberOfFiles_InsideOneFolder_TEST()
    Dim c As Long
    c = Get_NumberOfFiles_InsideOneFolder("G:\Archiv VBA\Bes FaceTags\Reports")
End Sub
'-------------------------------------------- 'MODUL: M01
Function Get_NumberOfFiles_InsideOneFolder(strDir As String, Optional strType As String) As Long
    'DEVELOPER: Ryan Wells (wellsr.com)
    'INPUT: Pass the procedure a string with your directory path and an optional
    '       file extension with the * wildcard
    'EXAMPLES: Call CountFilesInFolder("C:\Users\Ryan\")
    '          Call CountFilesInFolder("C:\Users\Ryan\", "*txt")
    Dim file As Variant, i As Long
    If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
    file = Dir(strDir & strType)
    While (file <> "")
        i = i + 1
        file = Dir
    Wend
    Get_NumberOfFiles_InsideOneFolder = i
End Function

'-------------------------------------------- 'MODUL: M01
Sub Get_Number_OfFiles_inFolderAndSubFolders_TEST()
    DoArr
    show Get_Number_OfFiles_inFolderAndSubFolders(ArrC(3))
    Beep
End Sub
'-------------------------------------------- 'MODUL: M01
Function Get_Number_OfFiles_inFolderAndSubFolders(localRoot, Optional fld, Optional count As Long) As Long
    Dim fso, F, baseFolder, subFolder, ftpFile, i
    Set fso = CreateObject("Scripting.Filesystemobject")
    If IsMissing(fld) Then Set baseFolder = fso.GetFolder(localRoot) Else Set baseFolder = fld
    count = count + baseFolder.Files.count
    For Each subFolder In baseFolder.subfolders
        Get_Number_OfFiles_inFolderAndSubFolders localRoot, subFolder, count
        displayNumberOfFacesInFolderFacePics count
        DoEvents
    Next
    Get_Number_OfFiles_inFolderAndSubFolders = count
End Function

'-------------------------------------------- 'MODUL: M01
Sub displayNumberOfFacesInFolderFacePics(count As Long)
    'wird von Get_Number_OfFiles_inFolderAndSubFolders für eine fortlaufende Anzeige benutzt
    [F13] = "Derzeit befinden sich " + CStr(count) + " Faces im Ordner <FacePics>"
End Sub

'-------------------------------------------- 'MODUL: M01
Sub Delete_T6Shape(NameOfShape$)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("T6")
    On Error Resume Next
    ws.Shapes(NameOfShape).Delete
    On Error GoTo 0
End Sub

Sub Delete_AllShapesWithMyNameOfOneSheet_TEST()
    Delete_AllShapesWithMyNameOfOneSheet "T4", "RoRe"
End Sub

Sub Delete_AllShapesWithMyNameOfOneSheet(NameOfSheet$, NameOfShape$)
    Dim i%, ws As Worksheet, shp As shape
    Set ws = ThisWorkbook.Sheets(NameOfSheet)
    For i = 1 To 9999
        Set shp = ws.Shapes(NameOfShape)
        If ShapeExists(shp) Then
            SearchAndShow_Shape NameOfSheet, NameOfShape
            'Stop
            shp.Delete
            'Stop
        Else
            Exit For
        End If
    Next
    Beep
End Sub

Sub SearchAndShow_Shape_TEST()
    SearchAndShow_Shape "T1", "Picture 2"
End Sub

Sub SearchAndShow_Shape(NameOfSheet$, NameOfShape$)
    Dim i%, s%, z%, L!, T!, ws As Worksheet, shp As shape
    Set ws = ThisWorkbook.Sheets(NameOfSheet)
    Set shp = ws.Shapes(NameOfShape): With shp
    'Exit
        If Not ShapeExists(shp) Then Exit Sub
    'SheetCoordinates Left, Top
        L = .Left: T = .Top
    's = SpaltenNr der ShapeZelle
        For i = 1 To 9999
            If ws.Cells(1, i).Left > L Then s = i - 1: Exit For 'SpaltenNr der ShapeZelle
        Next
    'z = ZeilenNr der ShapeZelle
        For i = 1 To 9999
            If ws.Cells(i, 1).Top > T Then z = i - 1:  Exit For 'ZeilenNr  der ShapeZelle
        Next
    'Select 2x3-Range, ShapeZelle in der Mitte
        ws.Activate
        'EE 0: ws.Range(Cells(z - 1, s - 1), Cells(z + 1, s + 1)).Select: EE 1
        EE 0: ws.Cells(z, s).Select: EE 1
    'shp.Visible
        shp.Visible = msoTrue
    'Finals
        End With
End Sub

Function ShapeExists(NameOfSheet As String, NameOfShape As String) As Boolean
    Dim shp As shape
    
    On Error Resume Next
    ' Versucht das Shape direkt zuzuweisen
    Set shp = ThisWorkbook.Sheets(NameOfSheet).Shapes(NameOfShape)
    On Error GoTo 0
    
    ' Wenn das Objekt zugewiesen werden konnte, ist ShapeExists True
    ShapeExists = Not shp Is Nothing
End Function

Sub Show_Shape(NameOfSheet$, NameOfShape$)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(NameOfSheet)
    On Error Resume Next
    ws.Shapes(NameOfShape).Visible = msoTrue
    On Error GoTo 0
End Sub

Sub Hide_Shape(NameOfSheet$, NameOfShape$)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(NameOfSheet)
    On Error Resume Next
    ws.Shapes(NameOfShape).Visible = msoFalse
    On Error GoTo 0
End Sub

Sub Delete_T4Shape(NameOfShape$)
    'Called from    Worksheet_SelectionChange[T4]
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("T4")
    On Error Resume Next
    ws.Shapes(NameOfShape).Delete
    On Error GoTo 0
End Sub

Sub Delete_Shape_OnSheetX(NameOfShape$, NameOfSheetX$)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(NameOfSheetX)
    On Error Resume Next
    ws.Shapes(NameOfShape).Delete
    On Error GoTo 0
End Sub

Public Function exeIsRunning(sExeName As String, Optional sComputer As String = ".") As Boolean
    On Error GoTo Error_Handler
    Dim objProcesses    As Object
    Set objProcesses = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _
        & sComputer & "\root\cimv2").ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" _
        & sExeName & "'")
    If objProcesses.count <> 0 Then exeIsRunning = True
Error_Handler_Exit:
    On Error Resume Next
    Set objProcesses = Nothing
    Exit Function
Error_Handler:
    MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
            "Error Number: exeIsRunning" & vbCrLf & _
            "Error Description: " & Err.Description, _
            vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function

'-------------------------------------------- 'MODUL: M01
Function VergroesserungPasst(ByVal hA1#, ByVal bA1#, ByVal hF1 As Long, ByVal bF1 As Long, ByVal xA1#, ByVal yA1#) As Boolean
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    'crop-Rechteck N --(Vergrößerung)--> crop-Quadrat N:
    If hA1 * hF1 > bA1 * bF1 Then bA1 = hA1 * hF1 / bF1 Else hA1 = bA1 * bF1 / hF1
    
    Left = CLng(bF1 * (xA1 - bA1 * 0.5))
    Top = CLng(hF1 * (yA1 - hA1 * 0.5))
    Right = CLng(bF1 * (1 - xA1 - bA1 * 0.5))
    Bottom = CLng(hF1 * (1 - yA1 - hA1 * 0.5))
    
    If Not (Left < 0 Or Top < 0 Or Right > bF1 Or Right < 0 Or Bottom > hF1 Or Bottom < 0) Then VergroesserungPasst = True
End Function

'-------------------------------------------- 'MODUL: M01
Sub shapeUmbenennen()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Shapes(1).NAME = "Q2"
End Sub

'-------------------------------------------- 'MODUL: M01
Function getLngAnzahlNichtLeererZeilenOfFile(PathOfFile As String) As Long
    Dim ZeileNr As Long, strZeile As String
    Open PathOfFile For Input As #1 'Datei zeilenweise einlesen:
    Do While Not EOF(1)             '(Funktioniert auch bei großen Dateien.)
        Line Input #1, strZeile
        If strZeile <> "" Then ZeileNr = ZeileNr + 1
    Loop
    Close #1
    getLngAnzahlNichtLeererZeilenOfFile = ZeileNr
End Function

'-------------------------------------------- 'MODUL: M01
Function getStrAnzNichtLeererZeilenInString(StringMitZeilen As String) As String
    Dim Arr1() As String, s As String, i As Long, c As Long
    s = StringMitZeilen
    Arr1 = Split(s, vbCrLf)
    For i = 0 To UBound(Arr1)
        If Arr1(i) <> "" Then c = c + 1
    Next
    getStrAnzNichtLeererZeilenInString = c
End Function

'-------------------------------------------- 'MODUL: M01
Sub AnzeigeFotoDatum(s As String)
    '19540000 abc.jpg  oder  20181224_181920 abc.jpg
    [i24] = Mid(s, 7, 2) + "." + Mid(s, 5, 2) + "." + Left(s, 4)
    If Mid(s, 9, 1) = "_" Then [n24] = Mid(s, 17) Else [n24] = Mid(s, 10)
End Sub

Function Green1() As Long
    Green1 = 14348258
End Function
Function Green2() As Long
    Green2 = 11854022
End Function
Function Green3() As Long
    Green3 = 9359529
End Function
Function Green4() As Long
    Green4 = 3506772
End Function

Function Get_FilePaths_Like_insideSourceFolderAndSubFolders_EverythingSearch(SourceFolder$) As String
    'Called from    xxx
    'Action         alle FilePfade sammeln (of all Sub- und SubSubFolders)
    '               (mit EverythingSearchCmdLineTool 'es.exe')
    'pOutTxt        EverythingSearchCmdLineTool schreibt alle FilePfade in pOutTxt
    '               (shellOutputIntoString klappt leider nicht)
    
    'Vorbereitung
        Dim pOutTxt$, qq$, s$, shellCmd1$, SuchFilter$, startTime As Long
        qq = Chr(34) 'quote, "-Zeichen
    
    SuchFilter = " path: " + qq + SourceFolder + qq
    pOutTxt = ThisWorkbook.path + "\helpers\AllFilePathsArchivTrampolin.txt"
    shellCmd1 = Chr(34) + Get_FilePath_OfToolES + Chr(34) + " -s " _
        + SuchFilter + " -export-txt " + Chr(34) + pOutTxt + Chr(34)
    Call shell(shellCmd1)
    startTime = GetTickCount()
    Do While exeIsRunning("es.exe")
        If GetTickCount() - startTime > 30000 Then Exit Do
        'do nothing; exit if exe is running more than 3 sec
    Loop
    s = ReadFile(pOutTxt)
    If Right(s, 2) = vbCrLf Then s = Left(s, Len(s) - 2)
    Get_AllPathsOfFoldersWithTonfotosIni_EverythingSearch = s
End Function

Sub ShowAllDgTitlesAndFolders()
    Dim A$, Title$, v$, zs$, i%, z%, s%, ArrDgTitles() As String
    v = vbCrLf
    T4_Load_ArrDgTitles_FromT4ListOfDesignTitles ArrDgTitles
    'showArray ArrDgTitles
    With Sheets("T4")
    For i = 1 To UBound(ArrDgTitles)
        Title = Trim(ArrDgTitles(i))
        zs = Get_ZeSp_OfDesignTitleNotFolderName(Title)         '039064
        z = CInt(Left(zs, 3)):     s = CInt(Right(zs, 3))       '39 '64
        A = A + "Titel  = [" + .Cells(z, s) + "]" + v + "Folder = [" + .Cells(z - 1, s) + "]" + v + v
    Next
    End With
    show A
End Sub

Function Get_RangeOfFirstCellOfOneDesign(Title$) As String
    Dim zs$, z%, s%
    zs = Get_ZeSp_OfDesignTitleNotFolderName(Trim(Title)) '039064
    z = CInt(Left(zs, 3)) - 1: s = CInt(Right(zs, 3)) - 1 '38 '63
    Get_RangeOfFirstCellOfOneDesign = Get_ColumnLetter(s) + CStr(z)
End Function




