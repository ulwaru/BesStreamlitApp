Attribute VB_Name = "M_PD"
Option Explicit 'M_PD
'

Sub zzz_M_PD()
    'showProcs "add1"
    Application.EnableEvents = True
    'RenameModule "Modul1", "M10"
    'FillArrC 59, ""
    'showArray2D ArrPd
End Sub

Sub Add_PersonData_ToCompetitorsList()
    'Called from    [None]
    'Action         Überträgt Items      einer Person A aus PersonData.txt
    '               zu allen Vorkommen des Namens von A in CompetitorsList
    
    'Vorbereitung
        Dim Alter$, CompNnVn$, Jg$, L$, Nn$, NnVn$, PD1$, v$, Vn$, Ye$
        Dim C1%, i%, j%, ws As Worksheet, r As Range
        Dim Arr1() As String, Arr2() As String, ArrPD1(), ArrComp()
        v = vbCrLf: DoArr
    
    'arrPD1 - PersonData.txt
        Load_ArrPD5 ArrPD1
        'showArray2D arrPD1
    'arrComp - CompetitorsList (ohne Kopfzeile etc.)
        Load_ArrCompetitorsList ArrComp
        'showArray2D arrComp
    'CompNnVn - NamensListe competitors
        For i = 1 To UBound(ArrComp, 1)
            NnVn = ArrComp(i, 1) + ArrComp(i, 2)
            If InStr(1, CompNnVn, NnVn) = 0 Then CompNnVn = CompNnVn + NnVn + " "
        Next
        'show CompNnVn
    'arrComp ergänzen
        For i = 1 To UBound(ArrPD1, 1)
            Nn = Trim(ArrPD1(i, 1)): Vn = Trim(ArrPD1(i, 2)) 'Nach-, Vorname aus PersonData
            If InStr(1, CompNnVn, Nn + Vn) > 0 Then
                'PD-Name ist auch in CompetitorList
                If Not (Trim(ArrPD1(i, 4)) = "" And Trim(ArrPD1(i, 6)) = "") Then
                    'arrPD1 hat etwas zu bieten
                    For j = 1 To UBound(ArrComp, 1)
                       'Das ganze arrComp wird nach jedem Vorkommen von Nn, Vn durchsucht
                        If ArrComp(j, 1) = Nn Then
                            If ArrComp(j, 2) = Vn Then
                                Ye = Left(Trim(ArrComp(j, 9)), 4)     'YearEvent
                                Jg = Trim(CStr(ArrPD1(i, 4)))         'Jahrgang
                                If Jg <> "" Then Alter = CStr(CInt(Ye) - CInt(Jg))
                                ArrComp(j, 4) = Alter
                                ArrComp(j, 6) = Trim(ArrPD1(i, 6))    'Verein/Ort
                                ArrComp(j, 7) = Trim(ArrPD1(i, 7))    'LTV
                                Ye = "": Jg = "": Alter = ""
                            End If
                        End If
                    Next
                End If
            End If
        Next
        showArray2D ArrComp
End Sub

Sub Update_NameOfLeuteFolders_From_Pd()
    'Called from    Update_All
    'Action         Umbenennen einiger LeuteFolderName:
    '               bisheriger LeuteFolderName bis alle vh
    '               + Rufname von PersonDataTxt + Verein von PersonDataTxt
    '               + Nation von PersonDataTxt
    
    'Vorbereitung
        Dim LeuteFolderNames$, N1$, N2$, Nation$, Nn$, NnVn$, p$, PD$, Report$, Rufname$
        Dim v$, vh1$, vh2$, vh3$, vhNames$, vhrn$, Verein$, Vn$, xList$
        Dim i%, Nr%, C1&, C2&, arrLeute() As String
        v = vbCrLf
    'Names of LeuteFolders
        'Load_ArrNamesOfAllFolders_Leute arrLeute
        LeuteFolderNames = Get_Names_OfAllSubfolders_OneLevel(ArrC(4)) 'Folder "Leute"
        arrLeute = Split(LeuteFolderNames, vbCrLf)
    'Name and all vh-Names of each Person from LeuteFolders
        vhNames = Get_Name_vhNames_OfEachPerson_FromLeuteFolders(arrLeute) + v
        'show vhNames
    'Index NamePersonenOrdner
        For i = 0 To UBound(arrLeute)
            N1 = arrLeute(i)             'LeuteFolderName
            C1 = InStr(1, N1, "(")
            If C1 = 0 Then NnVn = N1 Else NnVn = Left(N1, C1 - 2)   'Meier, Max
            NnVn = Replace(NnVn, ", ", "")                      'MeierMax
            xList = xList + Format(i, "0000") + NnVn + "|"      '1234MeierMax|... IndexListe
        Next
    'Action
        For i = 1 To UBound(ArrPd, 1)
            'N= "": N2 = "": N2 = ""
            Nn = ArrPd(i, 1)       'Nachname (aus PersonData.txt)
            Vn = ArrPd(i, 2)       'Vorname  (aus PersonData.txt)
            Rufname = ArrPd(i, 8)  'Rufname  (aus PersonData.txt)
            Verein = ArrPd(i, 5)   'Verein   (aus PersonData.txt)
            Nation = ArrPd(i, 7)   'Nation   (aus PersonData.txt)
            C1 = InStr(1, xList, Nn + Vn)
            If C1 > 0 Then
                'PersonData-Name ist auch in LeuteFolderName
                Nr = Mid(xList, C1 - 4, 4)     'IndexNr für Name in arrLeute
                N1 = arrLeute(Nr) 'Aaron, Syd (GB)   'bisheriger LeuteFolderName
                If vhNames Like "*" + Nn + ", " + Vn + "*" Then
                    'vh existiert in LeuteFolderName
                    C1 = InStr(1, vhNames, Nn + ", " + Vn): C2 = InStr(C1, vhNames, v)
                    N2 = CutVonBis(vhNames, C1, C2 - 1)
                    If ArrPd(i, 8) <> "" Then N2 = N2 + ", rn " + ArrPd(i, 8) 'Rufname
                    If Verein <> "" Then N2 = N2 + ", " + Verein
                    If Nation <> "" And ArrPd(i, 7) <> "D" Then N2 = N2 + ", " + ArrPd(i, 7) 'Nation
                    N2 = N2 + ")"
                Else
                    'vh existiert nicht in LeuteFolderName
                    N2 = Nn + ", " + Vn
                    If Rufname <> "" Then N2 = N2 + " (rn " + Rufname
                    If N2 Like "*(*" Then
                        If Verein <> "" Then N2 = N2 + ", " + Verein
                    Else
                        If Verein <> "" Then N2 = N2 + " (" + Verein
                    End If
                    
                    If Nation <> "D" Then
                        If N2 Like "*(*" Then
                            If Nation <> "" Then N2 = N2 + ", " + Nation
                        Else
                            If Nation <> "" Then N2 = N2 + " (" + Nation
                        End If
                    End If
                    If N2 Like "*(*" Then N2 = N2 + ")"
                End If
                'LeuteFolderName ändern in ArrPd (Spalte 9)
                    ArrPd(i, 9) = N2
                'If N1 <> N2 Then
                    Report = Report + N1 + v + N2 + v + v
                'End If
            End If
        Next
        'show REPORT        'Kontrollmöglichkeit
    '-----------------------------------------------------------------------
    'Erstellung der neuen Datei 'PersonData.txt'
        Write_Pd_FromArrPd
    'Write new FolderNames (FolderName N1 in N2 ändern - in "Leute")
        p = ArrC(4) + "\"      'Path of Folder "Leute"
        Report = Report + v + v + "Änderungen von LeuteFolderNames:" + v + v
        For i = 1 To UBound(ArrPd, 1)
            'N1 = LeuteFolderName old finden
            NnVn = ArrPd(i, 1) + ", " + ArrPd(i, 2) 'NachnameVorname
            C1 = InStr(1, LeuteFolderNames + v, NnVn)
            If C1 > 0 Then 'also keine ZweitNamen
                C2 = InStr(C1, LeuteFolderNames + v, v)
                N1 = CutVonBis(LeuteFolderNames + v, C1, C2 - 1)
                N2 = ArrPd(i, 9)
                If N1 <> N2 Then
                    Report = Report + "N1: " + N1 + v + "N2: " + N2 + v + v
                    RenameFolder p + N1, p + N2
                End If
            End If
        Next
    'show REPORT
End Sub

Function Get_Name_vhNames_OfEachPerson_FromLeuteFolders(arrLeute) As String
    'Called from    Update_NameOfLeuteFolders_From_Pd
    'Status         Die aktuellste Info über Zweit-/Dritt-/...Namen = bei LeuteFolders
    'Action         liefert je linken Teil jedes LeuteFolderName bis zum letzten vh-Namen,
    '               z. B. "Mai, Lea (vh Juni" oder "Mai, Lea (vh Juni, vh Juli"
    
    Dim N$, vh$, i%
    For i = 0 To UBound(arrLeute)
        N = arrLeute(i)
        If N Like "* (vh *" Then
            vh = vh + Left(N, InStr(1, N, "(vh ") + 3) + Get_vh1_From_OneFolderName(N)
            If N Like "*vh*vh*" Then
                vh = vh + ", vh " + Get_vh2_From_OneFolderName(N)
                If N Like "*vh*vh*vh*" Then vh = vh + ", vh " + Get_vh3_From_OneFolderName(N)
            End If
            vh = vh + vbCrLf
        End If
    Next
    vh = Delete_EmptyEndRowsInString(vh)
    Get_Name_vhNames_OfEachPerson_FromLeuteFolders = vh
End Function

Function DaysAgo_GiveDate8(DateLike_20241231$) As Long
    Dim D$, D1 As Date, D2 As Date
    D2 = Date
    D = DateLike_20241231 ' "5/2/2025"
    D = Mid(D, 7, 2) + "/" + Mid(D, 5, 2) + "/" + Mid(D, 1, 4)
    D1 = CDate(D)
    DaysAgo_GiveDate8 = DateDiff("D", D1, D2)
End Function

Sub Update_Pd_From_LeuteFolders_T4CompetitorsList()
    'Called from    Update_All
    
    Stop 'Umstellung auf ArrPd
    
    
    'Nam            = Liste Namen, erzeugt aus LeuteOrdnerName + T4CompetitorsList
    '               = Zeilen wie "May, Karl" oder "Mai, Lea (vh Juni, rn Jule, TV Bonn)"
    'arrComp        = 2D-Array, T4 competitors list
    'ArrPD1      = 2D-Array, PersonData (von bisheriger TextDatei)
    'ArrPD2      = 2D-Array, PersonData (aktualisiert)
    'Action         Neue Daten werden in 'PersonData.txt' (= Referenzdatei) übernommen
    '               - im OrdnerNamen eines neu angelgten LeuteOrdners enthaltenen Infos
    '               - neue Daten aus der T4CompetitorsList
    '                 (entstehen bei Update_T4CompetitorsList_LLinksInEFolder
    '                 nach neu angelegter Ergebnisliste)

    'Vorbereitung
        Dim A$, Jhg$, KL$, L$, Nam$, NAME$, NameIndex$, NamesOfLeuteFolders$, Nation$
        Dim Nachname$, Ort$, Rufname$, T$, v$, Verein$, vh$, Vorname$
        Dim C1%, xNr%, i%, j%, k%, N%, ws As Worksheet, r As Range
        Dim Arr1() As String, ArrN() As String, ArrComp()
        Dim ArrPD1(), ArrPd2()
        v = vbCrLf: DoArr: EE 0
    'vh
        vh = Get_vh_From_AllLeuteFolders 'Zeilen wie "Latton, Ute = Czech, Ute"
    'Names of LeuteFolders
        NamesOfLeuteFolders = Get_Names_OfAllSubfolders_OneLevel(ArrC(4)) 'Folder "Leute"
    'Names of LeuteFolders + CompetitorsList
        Nam = Get_AllNames(NamesOfLeuteFolders) 'Mai, Lea (vh Juni, Rufname Jule, TV Bonn)
        '   = MaxListe Namen; aus LeuteFolders + CompetitorsList
        '   = alle verfügbaren, relevanten Infos zur Emittlung Namen/AnzahlNamen
        ArrN = Split(v + Nam, v): N = UBound(ArrN)      'N = Anzahl Namen
        'Das zu füllende PersonData-Array muss also genau (N,9) groß werden
        ReDim ArrPd2(1 To N, 1 To 9)
        '     ArrPD1 = 2D-Array, PersonData (von bisheriger TextDatei)
        '     ArrPD2 = 2D-Array, PersonData (aktualisiert)
    'Nam --> ArrPD2
        For i = 1 To N
            L = ArrN(i)     'one Line 'Mai, Lea (vh Juni, Rufname Jule, TV Bonn)
            Nachname = Get_Nachname_From_OrdnerName(L)
            Vorname = Get_Vorname_From_OrdnerName(L)
            Rufname = Get_Rufname_From_OrdnerName(L)
            Verein = Get_Verein_From_OrdnerName(L)
            Nation = Get_Nation_From_OrdnerName(L)
            ArrPd2(i, 1) = Nachname: ArrPd2(i, 2) = Vorname: ArrPd2(i, 5) = Verein
            ArrPd2(i, 7) = Nation: ArrPd2(i, 8) = Rufname
            ArrPd2(i, 9) = Get_BirthFolderName(NamesOfLeuteFolders, vh, L)
            NameIndex = NameIndex + Format(i, "0000") + Nachname + Vorname + " " '0123BardyFrank
            Verein = "": Nation = "": Rufname = ""
        Next
        'showArray2D ArrPD2, "ArrPD2 (Step 1 of filling process)"
        
        'Jetzt steht bereit:    ArrPD2 - aktuelle Liste aller Namen im PD-Raster
        '              inkl.    Namen aus neuen 'Leute'-Ordnern  ("Mai, Lea (vh Müller, Rufname Ma, Bonn)")
        '              inkl.    Namen aus neuen T4Ergebnislisten ("|Ball|Ron|GB|")
   '-------------------------------------------------------------------------------------------
        'Jetzt wird ArrPD2 befüllt
        '(1)           mit      Daten aus T4CompetitorsList
        '(2)           mit      Daten aus ArrC(59) = Zwischenspeicherung Jhg/Ort
        '(3)           mit      Daten aus der bisherigen Datei 'PersonData.txt'
        'danach        Erstellung des Strings PD2 (aus ArrPD2)
        '              Erstellung der neuen Datei 'PersonData.txt'
        
   '(1) ArrPD2-Befüllung mit Daten aus arrComp
        Set ws = Sheets("T4"): ws.Activate
        Set r = ws.Range(Cells(8, 5), Cells(Get_NrOfLastRowInColumnNr(5, "T4"), 14))
        ArrComp = r.Value                           'T4CompetitorsList (ohne Kopfzeile etc.)
        'arrComp --> ArrPD2 (PersonData)
        For i = 1 To UBound(ArrComp, 1)
            Nachname = ArrComp(i, 1): Vorname = ArrComp(i, 2)
            C1 = InStr(1, NameIndex, Nachname + Vorname)
            xNr = CInt(Mid(NameIndex, C1 - 4, 4))                        'NameIndex in ArrPD2
            If ArrComp(i, 3) <> "" Then ArrPd2(xNr, 3) = ArrComp(i, 3)   'm/w
            If ArrComp(i, 6) <> "" Then ArrPd2(xNr, 5) = ArrComp(i, 6)   'Verein
            If ArrComp(i, 7) <> "" Then ArrPd2(xNr, 6) = ArrComp(i, 7)   'Verband
            If ArrComp(i, 8) <> "" Then ArrPd2(xNr, 7) = ArrComp(i, 8)   'Nation
        Next
        'showArray2D ArrPD2, "ArrPD2 (Step 2 of filling process)"

        
        
'   '(2) ArrPD2-Befüllung mit Daten aus ArrC(59) = Zwischenspeicherung Jhg/Ort
'        'ArrC(59)       "|Ralf Widra j61|Carl Furrer j63 Harlow|...|Glenn Kelly j65 Melbourne"
'        If ArrC(59) <> "" Then
'            arr1 = Split(ArrC(59), "|")
'            For i = 1 To UBound(arr1)
'                L = arr1(i)             '1 Eintrag, "Carl Furrer j63 Harlow"
'                c1 = InStr(1, L, " j")
'                Name = left(L, c1 - 1): Jhg = "19" + Mid(L, c1 + 2, 2)
'                If Len(L) > c1 + 5 Then Ort = Mid(L, c1 + 5)
'                c1 = InStr(1, Name, " ")
'                Vorname = left(Name, c1 - 1): Nachname = Mid(Name, c1 + 1) 'Vorname, Nachname
'                'Eintrag in ArrPD2
'                c1 = InStr(1, NameIndex, Nachname + Vorname)
'                If c1 = 0 Then
'                    show "'" + Name + "' steht nicht im NameIndex:" + v + v + NameIndex
'                Else
'                    xNr = CInt(Mid(NameIndex, c1 - 4, 4))    'NameIndex in ArrPD2
'                    ArrPD2(xNr, 4) = Jhg
'                    If Ort <> "" Then ArrPD2(xNr, 6) = Ort
'                End If
'            Next
'        End If
'        'ArrC(59) wird am Prozedur-Ende geleert
   '(3) ArrPD2-Befüllung mit Daten aus PD1 - bisherige PersonData.txt
        'Load_ArrPD5 ArrPD1
        Load_ArrPD5 ArrPD1
        'showArray2D ArrPD1, "ArrPD1 (Step 3 of filling process)"
        'ArrPD1 --> ArrPD2 (PersonData); PD1-Daten überschreiben PD2-Daten
        Give_ErrInfoPersonData ArrPD1, NameIndex
        For i = 1 To UBound(ArrPD1, 1)
            'Schleife über die Zellen der bisherigen PersonData.txt
            Nachname = Trim(ArrPD1(i, 1))
            Vorname = Trim(ArrPD1(i, 2))
            'TxtOldZeile i steht in TxtNewZeile xNr
                C1 = InStr(1, NameIndex, Nachname + Vorname)
                If C1 > 0 Then
                        xNr = CInt(Mid(NameIndex, C1 - 4, 4))    'NameIndex in ArrPD2
                    'Write all NonEmptyTxtOldCell to TxtNew
                        For k = 3 To 9
                            If Trim(ArrPD1(i, k)) <> "" Then ArrPd2(xNr, k) = Trim(ArrPD1(i, k))
                        Next
                End If
        Next
        'showArray2D ArrPD2, "ArrPD2 (Step 4 of filling process)"
    'Erstellung der neuen Datei 'PersonData.txt'
        Write_Pd_FromArrPd
   'Finale
        EE 1 ':FillArrC 59, ""
End Sub

Sub Write_PersonDataTxt_Give_PdFull(StringOfCompletePersonDataTxt$)
    'Status     Einzige Prozedur, welche die Datei 'PersonData.txt' schreibt
    
    Dim AnzNamen$, p1$, p2$, Pd2$
    Pd2 = StringOfCompletePersonDataTxt
    AnzNamen = CStr(anzAinB(vbCrLf, Pd2) - 6)
    Pd2 = Replace(Pd2, "repräsentierten Personen", "repräsentierten Personen (z. Zt. " + AnzNamen + ")")
    If Pd2 Like "*" + vbCrLf + "|  *" Then
        show "'Sub Write_PersonDataTxt_Give_PdFull' versucht Leerzeilen in PersonData.txt zu schreiben."
        Stop
        show Pd2
        Stop
'        Exit Sub
    End If
    'Erstellung der neuen Datei 'PersonData.txt'
         p1 = ArrC(2) + "\z PersonData.txt"
         writeStringToFile p1, Pd2
    'Erstellung einer SicherungsDatei 'PersonData.txt'
         p2 = ArrC(1) + "\prog\old\PersonData\PersonData " + Format(Now(), "yyyymmdd_hhmmss") + ".txt"
         writeStringToFile p2, Pd2
End Sub

Function Get_PdPure_FromArrPd() As String
    'Called from    Update_NameOfLeuteFolders_From_Pd
    '               Update_Pd_From_LeuteFolders_T4CompetitorsList
    '               T6_Add_T5LTVNation_LikeExistingCombinations
    'Action         Erstellung des Strings Pd (aus ArrPd)

    'Vorbereitung
        Dim PD$, v$, i%, j%
        v = vbCrLf
    'Action
        For i = 1 To UBound(ArrPd, 1)
            For j = 1 To UBound(ArrPd, 2)
                PD = PD + "|" + Trim(ArrPd(i, j))
            Next
            PD = PD + "|" + v
        Next
        'Pd = Replace(Pd, "|m  |", "| m |"): Pd = Replace(Pd, "|w  |", "| w |")
            'show Pd
            'Stop
        PD = Replace(PD, "|||||||||||", "")
        PD = Sort_TextLines(PD)
        PD = Delete_EmptyRowsInString(PD)
        Get_PdPure_FromArrPd = PD
End Function

Sub Give_ErrInfoPersonData(arrTxtOLD, NameIndex$)
    'Vorbereitung
        Dim Nachname$, Vorname$, T$, t1$, v$, C1&, i%
        v = vbCrLf
        t1 = "Zeilen in PersonData.txt, die nicht in AllNames/arrTxtNEW landeten:" + v + v
        
        For i = 1 To UBound(arrTxtOLD, 1)
            'Schleife über die Zellen der bisherigen PersonData.txt
            Nachname = Trim(arrTxtOLD(i, 1))
            Vorname = Trim(arrTxtOLD(i, 2))
            'TxtOldZeile i steht in TxtNewZeile xNr
                C1 = InStr(1, NameIndex, Nachname + Vorname)
                If C1 = 0 Then
                    T = T + "|" + Trim(arrTxtOLD(i, 1)) + "|" + Trim(arrTxtOLD(i, 2)) _
                          + "|" + Trim(arrTxtOLD(i, 3)) + "|" + Trim(arrTxtOLD(i, 4)) _
                          + "|" + Trim(arrTxtOLD(i, 6)) + "|" + Trim(arrTxtOLD(i, 7)) _
                          + "|" + Trim(arrTxtOLD(i, 8)) + "|" + Trim(arrTxtOLD(i, 9)) + v
                End If
        Next
        If T <> "" Then show t1 + T: Stop
End Sub

Function Read_Pd() As String
    'Called from    Add_OneLineToPersonData, Load_ArrPD5,
    '               Update_LinksOfPersonFolders_insideEvents
    'Status         Einzige Prozedur, welche die Datei 'PersonData.txt' einliest

    Dim PD1$, C1%
    DoArrc
    'Read
        PD1 = ReadFile(ArrC(2) + "\z PersonData.txt")
        If PD1 Like "*" + vbCrLf + "|  *" Then Stop
        C1 = InStr(1, PD1, "|A")
        PD1 = Mid(PD1, C1): PD1 = Delete_EmptyEndRowsInString(PD1)
    'Leerzeichen raus (wurden durch TabSimulation produziert)
        PD1 = Replace99(PD1, "  ", " ")
        PD1 = Replace(PD1, " |", "|")
        PD1 = Replace(PD1, "| ", "|")
    Read_Pd = PD1
    'show PD1
End Function

Function Add_LeadingText_ToPdPure(DataLines$) As String
    'Called from    Add_OneLineToPersonData
    
    Dim Header$, s$, v$, C1%, C2%
    v = vbCrLf
    Header = "|Nachname|Vorname|m/w|Jhg.|Verein|Landesturnverband|Nat|Rufname|Ordnername in 'Leute'|Schreibweisen|"
    s = Header + v + DataLines
    s = TAB_Simulation(s)
    C1 = InStr(1, s, v) - 1
    s = "PersonData.txt  -  Daten zu den im Archiv repräsentierten Personen" + v + v _
        + "(Hier können Änderungen/Ergänzungen an den Personendaten vorgenommen werden.)" + v + v _
        + String(C1, "-") + v + s
    'Linie unterhalb des Headers einfügen
        C2 = InStr(1, s, "m/w"): C2 = InStr(C2, s, v)
        s = Left(s, C2 - 1) + v + String(C1, "-") + Mid(s, C2)
    s = Replace(s, "Schreibweisen|", "Schreibweisen|" + v + String(C1, "-"))
    Add_LeadingText_ToPdPure = s
End Function

Sub Add_OneLineToPersonData(CellBoxValue$)
    'Alte Version
    'Called from    CellBox_KeyDown
    'OneLine        = 9 Spalten
    '               = "|Nachname|Vorname|m/w|Jhg.|Verein|LTV|Nat|Rufname|Ordnername in 'Leute'|"

    'Vorbereitung
        Dim mw$, Nn$, OneLine$, p$, PD1$, Pd2$, Vn$, c%
        Call DoArrc:
    'CellBoxValue auswerten
        c = InStr(1, CellBoxValue, " ")
        If c > 0 Then
            Vn = Left(CellBoxValue, c - 1): Nn = Mid(CellBoxValue, c + 1)
            mw = "": If ArrC(65) = "Herren" Then mw = "m"
            
            OneLine = "|" + Nn + "|" + Vn + "|" + mw + "|||||||"
        End If
    'Write OneLine to PersonData.txt
        PD1 = Read_Pd                                  'nur DataZeilen, kein Vorspann
        Pd2 = PD1 + vbCrLf + OneLine                    '+ gewünschte Zeile
        Pd2 = xxSort_TextLines(Pd2)                       'Sortieren
        Pd2 = xxAdd_LeadingText_ToPdPure(Pd2)   '+ Vorspann, TabSimulation
        'show PD1                                       'Anzeigen
        xxWrite_PersonDataTxt Pd2
End Sub

Sub T1_ShowPersonDataTxt()
    'Vorbereitung
        Dim s%, z%
        DoArrc
    'PersonData.txt öffnen
        openFile ArrC(2) + "\z PersonData.txt"
    'T1-Anzeige (Häkchen und ChangeSelect)
        If ActiveSheet.NAME = "T1" Then
            z = Get_T1Row_HoldingMyText("Show                    'PersonData.txt'")
            s = Get_T1Column_HoldingMyText("Show                    'PersonData.txt'")
            ActiveWorkbook.Sheets("T1").Cells(z, s + 12) = "ü": RefreshScreen
            ActiveWorkbook.Sheets("T1").Cells(z - 1, s - 1).Select
        End If
End Sub


