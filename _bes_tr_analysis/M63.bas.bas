Attribute VB_Name = "M63"
Option Explicit 'M63

Sub zzz_M63()
    showProcs "open folder"
    
    'RenameModule "Modul1", "M62"
    EE 1: Beep
End Sub

Sub Update_AllLinks_InsideStick_ToVolumeLetter_T()
    'Called from    xxx
    'Scope          Nur Stick
    'ca. 24.000 Links
    'RelativePath       %comspec% /C start "" "JumpTo" [Eintrag in MyLink>Eigenschaften>Ziel]
    '   %comspec%       = absolute path to cmd.exe
    '   /C              = close the command prompt after execution
    '   start ""        = Text für die Titelleiste des CMD-Fensters ist ""
    '   "JumpTo"        = "..\"; 1 FolderLevel above Folder who holds .lnk
    '   z. B.           sei     PathOfMyLink = "..\Archiv Trampolin\Events\1965 DMM\MyLink.lnk"
    '                   dann öffnet   JumpTo = "..\"           den Folder 'Events'
    '                   dann öffnet   JumpTo = "..\..\"        den Folder 'Archiv Trampolin'
    '                   dann öffnet   JumpTo = "..\..\Leute"   den Folder 'Leute'
    
    'Vorbereitung
        Dim oFSO As Object, oFolder As Object, oFile As Object, sF
        Dim J1$, J2$, L2$, N2$, p1$, p2$, p3$, Report$
        Dim AnzFold%, C1&, C2&, colFolders As New Collection
        p1 = "F:\Archiv TR\Archiv Trampolin" '= PathOld
        p2 = "T:\Archiv Trampolin"           '= PathNew
        p3 = "F:\Archiv Trampo\Archiv Trampolin" '= PathOld
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        Set oFolder = oFSO.GetFolder(p2)
    colFolders.Add oFolder          'start with this folder
    AnzFold = colFolders.count
    Do While colFolders.count > 0      'process all folders
        Set oFolder = colFolders(1)    'get a folder to process
        colFolders.Remove 1            'remove item at index 1
        C1 = C1 + 1: [AN3] = C1: [an4] = oFolder.NAME
        For Each oFile In oFolder.Files
            
            If oFolder.path + "\" + oFile.NAME Like "*.lnk" Then
            
                L2 = oFolder.path + "\" + oFile.NAME    '"T:\Archiv Trampolin\..."
                '  = Link_PathNew
                J1 = Get_LinkTargetPath(L2)             '"F:\Archiv TR\Archiv Trampolin\..."
                '  = JumpTo_PathOld
                If Left(J1, 1) <> "T" Then
                    J2 = Replace(J1, p1, p2)
                    J2 = Replace(J2, p3, p2)
                    '  = JumpTo_PathNew
                    If J2 Like "T:\Archiv Trampolin\*" Then
                        N2 = Replace(oFile.NAME, ".lnk", "")
                        'Call beep: OpenFolder oFolder.path: Stop
                        Create_OneLinkFile oFolder.path, J2, N2
                        'Stop: CloseFolder oFolder.path
                        '[an4] = NameOfPath(N2)
                        DoEvents
                        C2 = C2 + 1: [AO3] = C2
                    Else
                        Report = Report + "L2 = " + L2 + vbCrLf + "J2 = " + J2 + vbCrLf + vbCrLf
                    End If
                End If
            End If
        Next oFile
        'add any subfolders to the collection for processing
        For Each sF In oFolder.subfolders
            colFolders.Add sF
        Next sF
        DoEvents
    Loop
    If Report <> "" Then show Report
End Sub

Sub T6_Add_IDs_toFilesInOpenEventFolders(z%)
    'Called from    [UserClick on T6-CellButton 'Add ID in open EventFolders']
    '               T6_UserClickOnColumn2CellButton
    'Scope          erreicht nur EventOrdner mit Namen like "####*" (1959 ..., 1962-09-00 ..., ...)
    'z              ZeilenNr des geklickten T6-CellButtons T6(z, 2)
    'EventFolders   "Events" and "ClubsNations"

    'Vorbereitung
        Dim p$, s1$, v$, C2%, i%, Arr1() As String, r As Range
        v = vbCrLf: DoArr: Ticks1 ': With Worksheets("T6")
    'MyCountDown Init
        FillArrC 30, CStr(z) 'CountDownShowCell-ZeilenNr
        MyCountDown_Init "|T6|" + CStr(z) + "|19|1|" '= |SheetName|z|s|StartNr| of CountDownShowCell
    'AllPathsOfOpenEventFolders = Pfade aller gerade im Explorer geöffneter EventFolder
        s1 = Get_PathsOfOpenExplorerWindows_Events_ClubsNations
        If s1 = "" Then Exit Sub
        Arr1 = Split(s1, v)
        'showArray arr1: Stop
        For i = 0 To UBound(Arr1)
            'Schleife über alle derzeit offenen EventFolder
            p = Arr1(i) 'Pfad zu einem EventFolder (im Explorer geöffnet)
            'c2 wird bei jeder hinzugefügten Id hochgezählt
            Add_IDs_toFilesInOneEventFolder p, C2
            FillArrC 24, CStr(C2) 'CountOfChanges
        Next i
    'Finals
        T6_DoDoneRemarks
        'End With:
        Beep
End Sub

Sub T6_Create_PersonLinkIcon_Of_zzJpg_in_OpenPersonFolder(z%)
    'Called from    [UserClick on T6-CellButton ' Create ico of zzJpg in open PersonFolder']
    '               T6_UserClickOnColumn2CellButton
    'Vorbereitung
        Dim p$, s$, i%, Arr1() As String
        Call DoArr: Ticks1
    'MyCountDown Init
        FillArrC 30, CStr(z) 'CountDownShowCell-ZeilenNr
        MyCountDown_Init "|T6|" + CStr(z) + "|19|1|" '= |SheetName|z|s|StartNr| of CountDownShowCell
    
    s = Get_PathsOfOpenExplorerWindows
    'show s
    Arr1 = Split(s, vbCrLf)
    For i = 0 To UBound(Arr1)
        p = Arr1(i)     'PathOfOneFolder (is open in ExplorerWindow)
        If p Like "*1999\Leute\*" Then
        
        
            If FileExists(p + "\zz.jpg") Then
                Create_JpgIconLink_OfOnePerson_zzJpg p
                FillArrC 24, CStr(CInt(ArrC(24)) + 1) 'CountOfChanges
            End If
        End If
    Next
    T6_DoDoneRemarks
    Beep
End Sub

Sub Delete_AllLinks_ClubLTVNation()
    Dim F$, p$, v$, i%, A() As String
    v = vbCrLf
    p = "F:\Archiv TR\Archiv Trampolin\Leute"
    F = Get_FilePaths_Like_insideSourceFolderAndSubFolders(p, "*- *.lnk")
    show F: Stop
    A = Split(F, v)
    For i = 0 To UBound(A)
        DeleteFile A(i)
    Next
    Beep
End Sub

Sub T6_Add_T5LTVNation_LikeExistingCombinations() 'Add in T5: Club|LTV|Nation-Combinations
    'Called from    Update_All
    'Status         In T5 (PersonData) ist in manchen Zeilen VereinX, LTV und Nation vermerkt,
    '               in anderen Zeilen nur der VereinX; dort kann LTV/Nation ergänzt werden
    'Action         Ergänzt in manchen Zeilen fehlende LTV/Nation-Angaben
    
    'Vorbereitung
        Dim LTV$, Nation$, Verein$, VLN$, s2$, c&, i%, zDR%, zLast%, zT6%
        Dim r As Range, Arr1() As String, A()
        With Sheets("T5"): zLast = Get_NrOfLastRowInColumnNr(3, "T5"): Ticks1
    'MyCountDown Init
        'zT6 = T6-ZeilenNr, die den Cell-Button enthält (für T6_DoDoneRemarks-Einträge)
        zT6 = Get_RowNr_HoldingMyTextWholeInColumnX("T6", 2, 4, " Add LTV/Nation in T5")
        MyCountDown_Init "|T6|" + CStr(zT6) + "|19|1|" '= |SheetName|z|s|StartNr| of CountDownShowCell
    'Array mit den T5-Spalten Verein/LTV/Nation
        A = .Range(.Cells(8, 7), .Cells(zLast, 9)).Value
    'Liste |Verein|LTV|Nation|  'falls T5-Zeile alle 3 Einträge kennt
        For i = 1 To UBound(A, 1)
            Verein = A(i, 1): LTV = A(i, 2): Nation = A(i, 3)
            If Verein <> "" And LTV <> "" And Nation <> "" Then
                'VLN = Sammlung von Verein-LTV-Nation-Kobinationen, wird laufend gefüllt
                '    = Zeilen wie "|TV Schwafheim|Rheinland|D|"
                If Not VLN Like "*|" + Verein + "|*" Then
                    VLN = VLN + "|" + Verein + "|" + LTV + "|" + Nation + "|" + vbCrLf
                End If
            End If
        Next
        'show TAB_Simulation(VLN) '(komplette Liste, bei der alle 3 Spalten (Verein, LTV, Nation) einen Eintrag enthalten)
    'Action
        For i = 1 To UBound(A, 1)
            'Schleife über alle T5-Zeilen (= A-Zeilen)
            Verein = A(i, 1): LTV = A(i, 2)
            If Verein <> "" And LTV = "" Then
            
                c = InStr(1, VLN, "|" + Verein + "|")
                If c > 0 Then
                    'Der Verein aus A ist in VLN enthalten;
                    'LTV und Nation können in A überschrieben werden
                    s2 = CutVonBis(VLN, c, InStr(c, VLN, vbCrLf) - 1)
                    '  = ganze VLN-Zeile mit jenem Verein aus A
                    '  = "|TV Forst|Baden|D|"
                    If s2 Like "|*|*|*|" Then
                        Arr1 = Split(s2, "|")
                        A(i, 2) = Arr1(2)  'LTV                    'Baden
                        A(i, 3) = Arr1(3)  'Nation                 'D
                        MyCountDown_Add1ToChanges
                    End If
                End If
            End If
        Next
        'showArray2D A
    'Paste
        Paste_2DArrayToCell_z_s "T5", 8, 7, A
    'Finals
        T6_DoDoneRemarks
        End With:  Beep
End Sub


