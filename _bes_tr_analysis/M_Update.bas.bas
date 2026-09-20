Attribute VB_Name = "M_Update"
Option Explicit 'M_Update

Sub zzz_M_Update()
    showProcs "name"
    Application.EnableEvents = True
    'DoArrC
    'T1_ShowLogBuch
    'show Get_Paths_OfAllSubfoldersAllLevelsAsArrayUseGlobalVar(ArrC(3))
End Sub

Sub Update_All()
    'Use UF1_ProgressBars
    LogBuch "Up01 Update_All was started"
    'Reihenfolge
        'DesignTitles
            
            
            T1_ShowLogBuch
            
            
        'ResultJpg_ToEventFolder_AndToDgLeuteFolders
        '   + LinksOfDgLeuteFolders_ToEventFolder + DgLeute_ToListOfCompetitors
            Update_Jpg_Links_Comp
                'search for all existing Designs;
                'for each Design: Create one results.jpg to EventFolder;
                '  get all CompetitorNames (+Verein/Nation necessary?)
                '  change Name to BirthName
                '  Copy jpg to each PersonFolder
                '  Update names of PersonFolders (should become unnecessary)
        
    LogBuch "Up01 Update_All has ended (Exit)"
    T1_ShowLogBuch
Beep
EE 1
Exit Sub
            
        
        'T4CompetitorsList
            Update_T4CompetitorsList_LLinksInEFolder
        'PersonData.txt
            Update_Pd_From_LeuteFolders_T4CompetitorsList
            T6_Add_T5LTVNation_LikeExistingCombinations
        'Name of PersonFolders
            Update_NameOfLeuteFolders_From_Pd
        'Statistics, Links, ReadMe
            T1_UpdateStatistics 'includes Update_Links_inside_FolderLeute, Update_ReadMe
        'Links of PersonFolders inside Events
            'Update_LinksOfPersonFolders_insideEvents
        'Delete
            'Delete_AllResultJpgs_InsideFolders_Events_Leute    '19660429 .WM03 Lafayette #result p2646-00 ub.jpg
            'Delete_AllEventFolderLinks_InsideFolders_Leute     '19660429 .WM03 Lafayette #event.lnk
            'Delete_AllPersonFolderLinks_InsideFolders_Leute    'Ball, Linda (GB).lnk
                'Don't delete SourceLink                        '19660429 WM03 Lafayette a00

        'LinksOfCompetitors to EventFolder
        
    'Macros (for use)
    '    Update_NamesOfPersonFolders
    '    Update_Terminkalender
    '    UpdateOnce_OrtToVerein_InsideNameOfPersonFolders
    '    T1_UpdateStatistics
    '    T1_UpdateStatistics1
    '    T1_UpdateStatistics2
End Sub

Sub Update_Jpg_Links_Comp()
    'Called from    Update_All
    'PathsOf        _LeuteFolder_All$   = Pfade aller PersonenOrdner (in 'Leute')
    '               _LeuteFolder_OneDg$ = Pfade aller PersonenOrdner der CompetitorsOneDg
    'ArrDgTitles    = 1D-Array, T4-'List of DesignTitles'
    
    Stop 'Control this sub 'better subName
    
    'Action         bearbeitet nacheinander alle existierenden Designs (Ergebnislisten)
    '(1) Loads      PathsOfLeuteFolderAll
    '(2) DgSchleife (Schleife über alle existierenden DesignTitles)
    '               ArrCCn      (2D-Array, As String, wird gefüllt, Competitor|Club|Nation)
    '               ArrDg       (2D-Array, wird von Fill_ArrCCn geladen)
    '               ArrPathsDg  (1D-Array, As String, hier erzeugt, wird in einer
    '                           Schleife (hier) von Fill_ArrPathsDg nach und nach gefüllt;
    '                           trägt dann alle Pfade zu den LeuteOrdnern,
    '                           in die das resultJPG hineinkopiert wird)
    '                           = Zeilen wie "...Archiv...\Leute\Lea, Mai (vh Juni, Bonn)"
    
    'Vorbereitung
        Dim AllNnVn_OneDg$, AllNameNation$, CCn$
        Dim DgTitles$, Report$, SpNamen$, Title$, v$, vhRenamings$, zs$
        Dim i%, i1%, i2%, s%, z%, ShowREPORT As Boolean
        Dim ArrCh(), ArrDgTitles() As String, ArrCCn() As String, ArrPathsDg() As String
        v = vbCrLf: DoArrc: Sheets("T4").Activate: Application.ScreenUpdating = False
        LogBuch "Up03 'Update_Jpg_Links_Comp' was started"
    'Einstellungen
        FillArrC 60, "1"   '"1" = write ResultJPGs, "0" = don't write ("0" for testing)
        ShowREPORT = 1      '1  = show REPORT,       0  = don't show
    'Logbuch
        If ArrC(60) = "0" Then LogBuch "Up03 'Write ResultJPGs' was set to 'No'" Else LogBuch "Up03 'Write ResultJPGs' was set to 'Yes'"
    'Report
        If ShowREPORT Then
            Report = "Das Makro 'Update_Jpg_Links_Comp' wurde aufgerufen." + v _
            + "Dort wurden folgende Einstellungen gewählt:   ResultJPGs erstellen: "
            If ArrC(60) = "0" Then Report = Report + "Nein" Else Report = Report + "Ja"
            Report = Report + v + String(46, " ") + "Report anzeigen:      Ja" + v + v
        End If
    'ArrDgTitles
        T4_Load_ArrDgTitles_FromT4ListOfDesignTitles ArrDgTitles
        
        
    'Anzahl zu bearbeitender Designs ggf. reduzieren:
        i1 = 1: i2 = UBound(ArrDgTitles) 'alle --> i1=1: i2=UBound(ArrDgTitles)
    For i = i1 To i2
        'DgSchleife (Schleife über alle existierenden DesignTitles)
            Title = ArrDgTitles(i)
        'In T4 (Spalte 2) wird ein rotes Rechteck über die einzelnen Titel wandern
            RedRectToNextDgListTitle Title
            Title = Trim(Title)
        'ProgressBar
            ShowProgressBar CLng(i2), CLng(i), Title
            Application.ScreenUpdating = False
            If ShowREPORT Then Report = Report + v + v + "--- Design  " + Format(i, "00") + "    [" + Title + "]" + v
        'DgActivate     (T4-Worksheet_SelectionChange wird ausgelöst, T4_SearchForBorders)
            'mit Select (wegen CreateJpg, CmdCreate_Click)
            zs = Get_ZeSp_OfDesignTitleNotFolderName(Title) '039064
            z = CInt(Left(zs, 3)): s = CInt(Right(zs, 3))   '39 '64
            EE 1: Cells(z, s).Select                  '(z, s) = Dg-Zelle, die den Titel trägt
            RedRectToNextDgListTitle Title + " "
        'EventFolder    (Check Existence of EventFolder (to write jpg and LeuteLinks into))
            If ArrC(57) <> ArrC(3) + "\" + Cells(z - 1, s) Then
                ShowInfo01NoFolder Title: GoTo jump1
            End If
            
            'OpenFolder ArrC(57)
            
        'ArrCCn          |Competitor|Club|Nation| (3 Spalten), lädt ArrDg
            Fill_ArrCCn ArrCCn 'neue Competitor|Club|Nation für ArrPathsDg; one Dg only
                'enthält Zeilen wie "|Alex Nic||USA|", "|Max Maier|TV Bonn||", dann Leerzeilen
                'showArray2D ArrCCn, "ArrCCn" + v + "[" + CStr(i) + "] " + Title
            If Not IsValidName(ArrCCn(1, 1)) Then
                'kommt vor, falls ein Design zwar eine Namenspalte, aber keine Namen trägt,
                'z. B. "1969 LTV-Pokal Saarlouis"
                'showArray2D ArrCCn: Stop
                GoTo jump1
            End If
        'JPG      (Dg-Knopf 'Create jpg' wird gedrückt; CmdCreate_Click)
            Sheets("T4").CmdCreate.Value = True
            If ShowREPORT Then AddToReport02 Report
        'ArrPathsDg  (Pj = FolderPathsForJpg)
            Fill_ArrPathsDg ArrCCn, ArrPathsDg  'ggf. neue Namen; Pd aktualisiert
            'showArray ArrPathsDg: Stop
        'Copy jpg to each Folder in LeuteFoldersForJpg
            CopyOneJpgToFolders ArrPathsDg, Report
        'Link of all LeuteFoldersForJpg to EventFolder (path in arrc(57))
            Create_LinkOfDgLeute ArrPathsDg, Report
        'FortschrittsBalken
            If UF2.Left <> 0 Then FillArrC 69, CStr(UF2.Left)
            If UF2.Top <> 0 Then FillArrC 70, CStr(UF2.Top)
jump1:
    Next
    'Finals
        Sheets("T4").Shapes("T4RedRect").Visible = False
        If ShowREPORT Then show Report
        EE 1: Beep
        Sheets("T4").Cells(1, 1).Select
        Application.GoTo Reference:=Worksheets("T4").Cells(1, 1), Scroll:=True
        LogBuch "Sub 'Update_Jpg_Links_Comp' has ended"
        Application.ScreenUpdating = True
End Sub

Sub Update_T4CompetitorsList_LLinksInEFolder()
    'Called from    [UserClick on T4-CellButton 'Update competitors list'], Update_All
    'Action         Sammelt Daten aus den T4-Designs; Zugriff auf PersonData.txt;
    '               Erzeugung von LeuteFolderLinks innerhalb des jew. EventFolders
    '(1)            FolderPaths_OfLiga
    '(2)            arrPD1, NnVn_LeuteFolderName ("Juni, Lea|Mai, Lea (vh Juni, TV Bonn)")
    '(3)            ListOfDesignTitles: Update
    '(4)            DgSchleife: Title, NameOfEventFolder, PathOfEventFolder
    
    'Vorbereitung
        Dim allEvents$, AllNames$, AllNnVn_OneDg$, Alter$, CompNnVn$
        Dim Jg$, L$, ListOfCompetitors$, NameOfEventFolder$, Nn$, NnVn$
        Dim NnVn_LeuteFolderName$, PathOfEventFolder$, pAllLiga$, pOneLiga$
        Dim sPaste$, Title$, v$, Vn$, Ye$, zs$
        Dim C1%, NumberOfLines%, i%, j%, Sp%, Ze%, zLast%
        Dim ws As Worksheet, r As Range
        Dim ArrDgTitles(), Arr1() As String, Arr2() As String, ArrComp(), ArrPD1()
        Set ws = Sheets("T4"): v = vbCrLf: DoArr: With Sheets("T4"): .Activate
        EE 0: .[E2:F3].Select: EE 1: RefreshScreen
    'FolderPaths_OfLiga
        pAllLiga = Get_FolderPaths_OfLiga
    'arrPD1 - PersonData.txt
        Load_ArrPD5 ArrPD1
        'showArray2D arrPD1
    'NnVn_LeuteFolderName    = "Juni, Lea|Mai, Lea (vh Juni, TV Bonn)"
        NnVn_LeuteFolderName = Get_NnVn_LeuteFolderName(ArrPD1)
    'ListOfDesignTitles
        T4_Write_ListOfDesignTitles_FromDgs
        Set r = ws.Range(Cells(8, 2), Cells(Get_NrOfLastRowInColumnNr(2, "T4"), 2))
        ArrDgTitles = r.Value 'Liste bereits existierender Designs
        
    'Get ListOfCompetitors (all Designs)
        For i = 1 To UBound(ArrDgTitles, 1)
        
'----------------------------------------------------------------------------------------

            'Schleife über alle existierenden Designs
            If (UBound(ArrDgTitles, 1) - i) Mod 5 = 0 Then Sheets("T4").[M2] = "0°" + Format((UBound(ArrDgTitles, 1) - i) / 5, "00")
            'Title
                Title = CStr(ArrDgTitles(i, 1))
                Title = Trim(Title)      '" " am Ende eines Titels löschen
                zs = Get_ZeSp_OfDesignTitleNotFolderName(Title) '039064
                    Ze = CInt(Left(zs, 3)): Sp = CInt(Right(zs, 3))   '39 '64
            'NameOfEventFolder
                NameOfEventFolder = Replace(Title, "/", "_")
                If .Cells(Ze - 1, Sp) <> "" Then NameOfEventFolder = .Cells(Ze - 1, Sp)
            'PathOfEventFolder
                PathOfEventFolder = Get_PathOfEventFolder_MayBeLiga(NameOfEventFolder, pAllLiga)
            'Design: Activate and analyze
                SearchForDesignBordersQuick Ze, Sp 'aktiviert Design 'z1,s1,z2,s2,... --> ArrC
            'ListOfCompetitors      wird hier fortlaufend gefüllt (über alle Designs)
            'AllNnVn_OneDg    wird hier nur für dieses Design gefüllt
                AllNnVn_OneDg = ""  'Alle "Nachname, Vorname", aber nur dieses einzelnen Designs
                Fill_ListOfCompetitors ListOfCompetitors, AllNnVn_OneDg
            'Erzeugung von LeuteFolderLinks innerhalb eines EventFolder
                Update_LeuteLinks_inOneEventFolder NameOfEventFolder, PathOfEventFolder, AllNnVn_OneDg, NnVn_LeuteFolderName
        Next
        
        ListOfCompetitors = Delete_EmptyRowsInString(ListOfCompetitors)
        'show ListOfCompetitors '|Vieweg|Christel| w | |Ti|FT Freiburg||| 1960-11 DM Frankfurt |Einzel|1|
    'Add m/w="", Klasse=""
        ListOfCompetitors = Add_mw_Klasse_InListOfCompetitors(ListOfCompetitors)
    'Add LTV, Nation
        ListOfCompetitors = Add_LTV_Nation_InListOfCompetitors(ListOfCompetitors)
    'arrComp = Array for pasting = arrCompFromCompetitorsInDesigns
        'ListOfCompetitors in arrComp übernehmen
        Arr1 = Split(ListOfCompetitors, vbCrLf)
        NumberOfLines = UBound(Arr1) + 1
        ReDim ArrComp(1 To NumberOfLines, 1 To 11)
        For i = 0 To UBound(Arr1)
            L = Arr1(i)         'one Line '|Maier|Max|m|19|Tu|TV Aaa|rhl|D|1980 DM Bbb|Einzel|1|
            Arr2 = Split(L, "|")
            If InStr(1, AllNames, "|" + Arr2(1) + "," + Arr2(2)) = 0 Then AllNames = AllNames + "|" + Arr2(1) + "," + Arr2(2)
            If InStr(1, allEvents, "|" + Arr2(9)) = 0 Then allEvents = allEvents + "|" + Arr2(9)
            For j = 1 To 11
                ArrComp(i + 1, j) = CStr(Arr2(j))
                If j = 4 Then ArrComp(i + 1, j) = "" 'Alter
            Next
        Next
        EE 0: [G3] = "  (" + CStr(NumberOfLines) + " lines, " _
                    + CStr(anzAinB("|", AllNames)) + " names, " _
                    + CStr(anzAinB("|", allEvents)) + " events)"
        EE 1
        'showArray2D arrComp
    'CompNnVn - NamensListe competitors
        For i = 1 To UBound(ArrComp, 1)
            NnVn = ArrComp(i, 1) + ArrComp(i, 2)
            If InStr(1, CompNnVn, NnVn) = 0 Then CompNnVn = CompNnVn + NnVn + " "
        Next
        'show CompNnVn 'ViewegChristel SchillingerRoland ...
    'Add Alter, Verein/Ort from PersonData to arrComp
            For i = 1 To UBound(ArrPD1, 1)
                'Schleife über alle Zeilen von arrPD1 (PersonData)
                Nn = Trim(ArrPD1(i, 1)): Vn = Trim(ArrPD1(i, 2)) 'Nach-, Vorname aus PersonData
                If InStr(1, CompNnVn, Nn + Vn) > 0 Then
                    'PD-Name ist auch in CompetitorList
                    'If Not (Trim(arrPD1(i, 4)) = "" And Trim(arrPD1(i, 6)) = "") Then
                        'arrPD1 soll immer arrComp überschreiben
                        For j = 1 To UBound(ArrComp, 1)
                           'Das ganze arrComp wird nach jedem Vorkommen von Nn, Vn durchsucht
                            If ArrComp(j, 1) = Nn Then
                                If ArrComp(j, 2) = Vn Then
                                    'arrComp, PersonData enthalten beide denselben Namen
                                    
                                    'ArrComp-Spalte 4: Alter
                                        Ye = "": Jg = "": ArrComp(j, 4) = ""
                                        Ye = Left(Trim(ArrComp(j, 9)), 4)     'YearEvent
                                        Jg = Trim(CStr(ArrPD1(i, 4)))         'Jahrgang
                                        If Jg <> "" Then Alter = CStr(CInt(Ye) - CInt(Jg)) Else Alter = ""
                                        ArrComp(j, 4) = Alter
                                    
                                    'Verein/Ort
                                        If ArrComp(j, 6) = "" Then ArrComp(j, 6) = Trim(ArrPD1(i, 6))
                                    'LTV
                                        If ArrComp(j, 7) = "" Then ArrComp(j, 7) = Trim(ArrPD1(i, 7))
                                    Ye = "": Jg = "": Alter = ""
                                End If
                            End If
                        Next
                Else
                
                
                End If
            Next
            'showArray2D arrComp
    'Clear old
        zLast = Get_NrOfLastRowInColumnNr(5, "T4")
        .Range(.Cells(8, 5), .Cells(zLast, 15)).ClearContents
        .Range(.Cells(8, 5), .Cells(zLast + 10, 15)).Borders.LineStyle = xlNone
    'Paste Array
        Paste_2DArrayToSheet Sheets("T4").Range("E8"), ArrComp
    'Hintergrund einfärben
        'show CStr(getColor([e8], 0, "T4")) '--> 14083324 (w-Farbe, rosa)
        'show CStr(getColor([e9], 0, "T4")) '--> 15652797 (m-Farbe, hellblau)
        zLast = Get_NrOfLastRowInColumnNr(5)
        For i = 8 To zLast
            If CStr(ArrComp(i - 7, 3)) = " w " Then .Range(.Cells(i, 5), .Cells(i, 15)).Interior.Color = 14083324
            If CStr(ArrComp(i - 7, 3)) = " m " Then .Range(.Cells(i, 5), .Cells(i, 15)).Interior.Color = 15652797
        Next
        For i = zLast To zLast + 100
            .Range(.Cells(i, 5), .Cells(i, 15)).Interior.ColorIndex = xlNone
        Next
    'Rahmen - each CellBorder: white
        With .Range(.Cells(8, 5), .Cells(zLast, 15)).Borders
            .LineStyle = xlContinuous: .Color = RGB(250, 250, 250): .Weight = xlThin
        End With
    'Finale
        .[M2] = "": EE 1: .[F5].Select: .[E5].Select: End With
End Sub

Function Get_AllNames(Optional NamesOfLeuteFolders$ = "") As String
    'Called from    Update_Pd_From_LeuteFolders_T4CompetitorsList
    'Zugriff auf    T4CompetitorsList, FolderLeute
    'Action         liefert komplette Namensliste incl. vH-Namen (Czech, Ute ...; Luxon, Ute ...)
    '               Zeilen wie "Marx, Karl", "Ley, A. (vh Abc, vh Def, rn Ghi, TV Jkl)"
    
    'Vorbereitung
        Dim L$, N$, Nation$, s$, s1$, s2$, v$, vh$, Verein$, zU$, i%, j%
        Dim Arr1() As String, Arr2() As String, ArrComp()
        v = vbCrLf: DoArr
    's1 NamesOfLeuteFolders
        s = NamesOfLeuteFolders
        If s = "" Then s = Get_Names_OfAllSubfolders_OneLevel(ArrC(4))
        's1 OrdnerNamen ohne 'Vorname, Nachname' entfernen
            Arr1 = Split(s, v)
            For i = 0 To UBound(Arr1)
                L = Arr1(i) 'one Line 'oneLeuteFolderName
                If L Like "*,*" Then s2 = s2 + L + v Else _
                    show "Bitte ändern: " + v + "Der Leute-Ordnername '" + L + "'" _
                    + v + "sollte die Form 'Aaa, Aaa (TV Aaa)' haben"
            Next
        s1 = Delete_EmptyRowsInString(s2) ': show s1
        '  = Zeilen wie "Aaron, Syd (GB)"
    'Load arrComp
        Load_ArrCompetitorsList ArrComp         ':showArray2D arrComp
    's1 erweitern (+ noch nicht in s1 enthaltene arrComp-Namen)
        For i = 1 To UBound(ArrComp, 1)
            Verein = "": Nation = "": zU = ""
            N = ArrComp(i, 1) + ", " + ArrComp(i, 2)
            'If InStr(1, vh, "|" + N + "|") = 0 Then 'alle vh-Namen weglassen
            Verein = ArrComp(i, 6)
            Nation = ArrComp(i, 8)
            If Nation = "D" Then
                If Verein <> "" Then zU = " (" + Verein + ")"
            Else
                If Verein = "" And Nation <> "" Then zU = " (" + Nation + ")"
                If Verein <> "" And Nation = "" Then zU = " (" + Verein + ")"
                If Verein <> "" And Nation <> "" Then zU = " (" + Verein + ", " + Nation + ")"
            End If
            If Not s1 Like "*" + N + "*" Then s1 = s1 + v + N + zU
            'End If
        Next
        'show s1  'Marx, Karl 'Ley, A. (vh Abc, vh Def, rn Ghi, TV Jkl) '...
    's1 Umlaute ersetzen (ä-->ae° etc.; wegen Sortierung)
        s1 = Replace(s1, v + "Ä", v + "Ae°"): s1 = Replace(s1, v + "Ö", v + "Oe°")
        s1 = Replace(s1, v + "Ü", v + "Ue°")
        s1 = Replace(s1, "ä", "ae°"): s1 = Replace(s1, "ö", "oe°"): s1 = Replace(s1, "ü", "ue°")
    'sort not case sensitive
        Arr1 = Split(s1, v): s1 = ""
        For i = 0 To UBound(Arr1)
            Arr1(i) = UCase(Left(Arr1(i), 5)) + Arr1(i)
        Next
        QuickSort Arr1
        For i = 0 To UBound(Arr1)
            'arr1(i) = Mid(arr1(i), 6)
            s1 = s1 + Mid(Arr1(i), 6) + v
        Next
    'Umlaute rückersetzen
        s1 = Replace(s1, v + "Ae°", v + "Ä"): s1 = Replace(s1, v + "Oe°", v + "Ö")
        s1 = Replace(s1, v + "Ue°", v + "Ü")
        s1 = Replace(s1, "ae°", "ä"): s1 = Replace(s1, "oe°", "ö"): s1 = Replace(s1, "ue°", "ü")
        s1 = Delete_EmptyRowsInString(s1)
    Get_AllNames = s1
End Function

Function Get_vh_From_AllLeuteFolders(Optional NamesOfLeuteFolders$ = "") As String
    'Called from    Get_AllNames
    'Action         liefert Liste der Geburtsnamen/weiteren Nachnamen von Personen;
    '               Zeilen wie "Latton, Ute = Czech, Ute"
    
    'Vorbereitung
        Dim N$, s$, v$, vh$, i%, Arr1() As String
        v = vbCrLf: DoArr
    's NamesOfLeuteFolders
        s = NamesOfLeuteFolders
        If s = "" Then s = Get_Names_OfAllSubfolders_OneLevel(ArrC(4))
    'Action
        Arr1 = Split(s, v)
        For i = 0 To UBound(Arr1)
            N = Arr1(i) 'oneLeuteFolderName '"Marx, Karl", "Ley, A. (vh Abc, vh Def, rn Ghi, TV Jkl)"
            vh = vh + Get_vh_GebNam_From_OneFolderName(N) + v
        Next
        s = Delete_EmptyRowsInString(vh) ': show s
    'Finals
        Get_vh_From_AllLeuteFolders = s
End Function

Function Get_vh_GebNam_From_OneFolderName(N$) As String
    'Called from    Get_AllNames
    'N              "Marx, K." oder "Müller, Lea (vh Nachname2, vh Nachname3, rn Ghi, TV Jkl)"
    'Action         liefert Liste aller weiteren Namen einer Person
    '               und die Zuordnung zum Geburtsnamen: "Nachname2, Lea = Müller, Lea"
    
        If Not N Like "*vh *" Then Exit Function
    'Vorbereitung
        Dim A$, Gn$, Nn$, NnVn$, s$, v$, Vn$, c%, i%, Arr1() As String
        v = vbCrLf
    'Action
        NnVn = Left(N, InStr(1, N, " (") - 1)       '"Müller, Lea"
            If Not NnVn Like "*,*" Then
                show "Bitte ändern:" + v + "Der Ordner '" + N + "' in 'Leute'" + v _
                    + "sollte mit 'Nachname, Vorname' beginnen"
                Exit Function
            End If
        Gn = Left(N, InStr(1, N, ",") - 1)          '"Müller"       Geburtsname
        Vn = Mid(NnVn, InStr(1, NnVn, ", ") + 2)    '"Ute"          Vorname
        Arr1 = Split(N, "vh ")
        For i = 1 To UBound(Arr1)
            A = Arr1(i)                             '"Nachname2, "  '"Nachname3, rn Ghi, TV Jkl)"
            c = InStr(1, A, ","): If c = 0 Then c = InStr(1, A, ")")
            Nn = Left(A, c - 1)                     '"Nachname2"    '"Nachname3"
            s = s + Nn + ", " + Vn + " = " + Gn + ", " + Vn + v
        Next
        Get_vh_GebNam_From_OneFolderName = s
End Function

Function Get_Nachname_From_OrdnerName(N$) As String
    'Called from    Update_Pd_From_LeuteFolders_T4CompetitorsList
    Get_Nachname_From_OrdnerName = Left(N, InStr(1, N, ",") - 1)
End Function

Function Get_Vorname_From_OrdnerName(N$) As String
    'Called from    Update_Pd_From_LeuteFolders_T4CompetitorsList
    'N              "Mai, Lea" oder "Mai, Lea (vh Juni, rn Jule, TV Bonn)"
    Dim C1%, C2%
    C1 = InStr(1, N, ","): C2 = InStr(1, N, "(")
    If C2 = 0 Then
        Get_Vorname_From_OrdnerName = Mid(N, C1 + 2)
    Else
        Get_Vorname_From_OrdnerName = Mid(N, C1 + 2, C2 - C1 - 3)
    End If
End Function

Function Get_Rufname_From_OrdnerName(N$) As String
    'Called from    Update_Pd_From_LeuteFolders_T4CompetitorsList
    'N              "Mai, Lea", "Mai, Lea (GB)", "Mai, Lea (vh Juni, rn Jule, TV Bonn)"
    Dim C1%, C2%
    If Not N Like "*(*rn *" Then Exit Function
    '"*rn *" kann ausserhalb der Klammer vorkommen: "Meier, Björn (TV Bonn)"
    C1 = InStr(1, N, "("): C1 = InStr(C1, N, "rn "): C2 = InStr(C1, N, ",")
    If C2 = 0 Then C2 = InStr(C1, N, ")")
    Get_Rufname_From_OrdnerName = CutVonBis(N, C1 + 3, C2 - 1)
End Function

Function Get_Verein_From_OrdnerName(N$) As String
    'Called from    Update_Pd_From_LeuteFolders_T4CompetitorsList
    'N              "Mai, Lea", "Mai, Lea (GB)", "Mai, Lea (vh Juni, rn Jule, TV Madrid, SPA)"
    Dim A$, Nation$, s$, Verein$, VereinNation$, C1%, C2%, i%, Arr1() As String
    If Not N Like "*(*" Then Exit Function
    C1 = InStr(1, N, "("): C2 = InStr(C1, N, ")")
    s = CutVonBis(N, C1 + 1, C2 - 1) 'Klammerinhalt
    
    Arr1 = Split(", " + s, ", ")
    For i = 1 To UBound(Arr1)
        A = Arr1(i)             '"vh Juni"  '"rn Jule"  '"TV Madrid"  '"SPA"
        If Left(A, 3) <> "vh " And Left(A, 3) <> "rn " Then VereinNation = VereinNation + A + "|"
    Next
    If Right(VereinNation, 1) = "|" Then VereinNation = Left(VereinNation, Len(VereinNation) - 1)
    'VereinNation = ""  "TV Madrid"  "SPA"  "TV Madrid|SPA"
    
    If VereinNation = "" Then Verein = "": Nation = ""
    C1 = InStr(1, VereinNation, "|")
    If C1 = 0 Then
        '"TV Madrid"  "SPA"
        If VereinNation Like "[A-Z]" Or VereinNation Like "[A-Z][A-Z]" _
            Or VereinNation Like "[A-Z][A-Z][A-Z]" Then
            Nation = VereinNation   '"SPA"
        Else
            Verein = VereinNation           '"TV Madrid"
        End If
    Else
        '"TV Madrid|SPA"
        Verein = Left(VereinNation, C1 - 1) '"TV Madrid"
        Nation = Mid(VereinNation, C1 + 1)  '"SPA"
    End If
    Get_Verein_From_OrdnerName = Verein
End Function

Function Get_Nation_From_OrdnerName(N$) As String
    'Called from    Update_Pd_From_LeuteFolders_T4CompetitorsList
    'N              "Mai, Lea", "Mai, Lea (GB)", "Mai, Lea (vh Juni, rn Jule, TV Madrid, SPA)"
    Dim A$, Nation$, s$, Verein$, VereinNation$, C1%, C2%, i%, Arr1() As String
    If Not N Like "*(*" Then Exit Function
    C1 = InStr(1, N, "("): C2 = InStr(C1, N, ")")
    s = CutVonBis(N, C1 + 1, C2 - 1) 'Klammerinhalt
    
    Arr1 = Split(", " + s, ", ")
    For i = 1 To UBound(Arr1)
        A = Arr1(i)             '"vh Juni"  '"rn Jule"  '"TV Madrid"  '"SPA"
        If Left(A, 3) <> "vh " And Left(A, 3) <> "rn " Then VereinNation = VereinNation + A + "|"
    Next
    If Right(VereinNation, 1) = "|" Then VereinNation = Left(VereinNation, Len(VereinNation) - 1)
    'VereinNation = ""  "TV Madrid"  "SPA"  "TV Madrid|SPA"
    
    If VereinNation = "" Then Verein = "": Nation = ""
    C1 = InStr(1, VereinNation, "|")
    If C1 = 0 Then
        '"TV Madrid"  "SPA"
        If VereinNation Like "[A-Z]" Or VereinNation Like "[A-Z][A-Z]" _
            Or VereinNation Like "[A-Z][A-Z][A-Z]" Then
            Nation = VereinNation   '"SPA"
        Else
            Verein = VereinNation           '"TV Madrid"
        End If
    Else
        '"TV Madrid|SPA"
        Verein = Left(VereinNation, C1 - 1) '"TV Madrid"
        Nation = Mid(VereinNation, C1 + 1)  '"SPA"
    End If
    Get_Nation_From_OrdnerName = Nation
End Function

Function Get_BirthFolderName(NamesOfLeuteFolders$, vh$, N$) As String
    'Called from    Update_Pd_From_LeuteFolders_T4CompetitorsList
    'vh             Zeilen wie "Latton, Ute = Czech, Ute"
    'N              Mai, Lea (vh Juni, Rufname Jule, TV Bonn)
    
    Dim Gn$, NAME$, s$, v$, vh2$, c%
    v = vbCrLf
    NAME = N: c = InStr(1, N, " ("): If c > 0 Then NAME = Left(N, c - 1) '"Latton, Ute"
    vh2 = v + vh + v
    c = InStr(1, vh2, v + NAME + " =")  'Position von v + "Latton, Ute =" in vh2
    If c = 0 Then
        s = N
    Else
        Gn = Mid(vh2, c + 2)                        'vh-Zeilen ab "Latton, Ute ="
        c = InStr(1, Gn, "="): Gn = Mid(Gn, c + 2)  '"Czech, Ute ..."
        c = InStr(1, Gn, v):  Gn = Left(Gn, c - 1)  '"Czech, Ute"   GeburtsName
        s = Get_FirstLine_LikeMyString(NamesOfLeuteFolders, Gn)
    End If
    Get_BirthFolderName = s
End Function

Function Get_PathOfEventFolder_MayBeLiga(NameOfEventFolder$, pAllLiga$) As String
    Dim p$, pOneLiga$: p = ArrC(3) + "\" + NameOfEventFolder
    If Not FolderExists(p) Then
        pOneLiga = Get_FirstLine_LikeMyString(pAllLiga, NameOfEventFolder)
        If pOneLiga <> "" Then
            p = pOneLiga
        Else
            Stop
            'NameOfEventfolder in 'Events' könnte geändert worden sein;
            'ggf. auch NameOfEventfolder im Design ändern
        End If
    End If
    Get_PathOfEventFolder_MayBeLiga = p
End Function

Sub ActivateDesign_GiveTitle(Title$)
    Dim Titel$, zs$, s%, z%
    Sheets("T4").Activate
    Titel = Trim(Title)
    If Titel Like "####*" Then
        zs = Get_ZeSp_OfDesignTitleNotFolderName(Titel) '039064
        z = CInt(Left(zs, 3)): s = CInt(Right(zs, 3))   '39 '64
        EE 1: Sheets("T4").Cells(z, s).Select
            '(z, s) = Dg-Zelle, die den Titel trägt
                'T4_SearchForBorders z, s
    End If
End Sub




