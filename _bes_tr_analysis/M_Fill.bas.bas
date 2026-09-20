Attribute VB_Name = "M_Fill"
Option Explicit 'M_Fill

Public ArrC() As String         'Container
Public ArrX() As String         'Namensliste Aktive (weiblich + männlich)
Public ArrW() As String         'Namensliste weiblich
Public ArrM() As String         'Namensliste männlich
Public ArrDg()                  'Actual Design
Public SomeTicks As String
'Function TimeToMillisecond
    Private Type SYSTEMTIME
        wYear As Integer: wMonth As Integer: wDayOfWeek As Integer: wDay As Integer: wHour As Integer: wMinute As Integer: wSecond As Integer: wMilliseconds As Integer
    End Type
    Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Sub zzz_M_Fill()
    'showProcs "is name"
    Application.EnableEvents = True
    'RenameModule "Modul1", "M08"
    'showArray2D ArrDg
    Beep
End Sub
'---------------------------------------

Sub T4_Add_Nation_ToAllArrDg_TeamMembers()
    'Called from    FillArrDg
    
    'Vorbereitung
        Dim z%, s%
    'Exit
        If Not ArrDg(2, 1) Like "M*" Then Exit Sub  'Mannschaft
        If Not ArrDg(2, 1) Like "*N*" Then Exit Sub 'NationSpalte
    'Action
        'NationSpalte suchen
        For s = 2 To UBound(ArrDg, 2) - 1
            'Schleife über DgSpalten
            If "STU" Like "*" + ArrDg(5, s) + "*" Then
                'NationSpalte gefunden: s
                For z = 6 To UBound(ArrDg, 1) - 2
                    'Schleife über DgZeilen
                    If ArrDg(z, s) <> "" Then
                        'Zelle mit Eintrag in NationSpalte gefunden
                        If ArrDg(z + 1, s) = "" Then
                            'leere Zelle darunter gefunden
                            If ArrDg(z + 1, s + 1) <> "" Then
                                'zugehörige NamenSpalte (rechts daneben) besitzt Eintrag
                                ArrDg(z + 1, s) = ArrDg(z, s)
                            End If
                        End If
                    End If
                Next
            End If
        Next
    'showArray2D ArrDg: Stop
End Sub

Sub T4_Add_Nation_ToAllArrDg_SyPartner()
    'Called from    FillArrDg
    
    'Vorbereitung
        Dim i%, Sp%
    'Exit
        If Not ArrDg(2, 1) Like "S*" Then Exit Sub  'Synchron
        If Not ArrDg(2, 1) Like "*N*" Then Exit Sub 'NationSpalte
    'Action
        'Nation-Spalte suchen
        For Sp = 2 To UBound(ArrDg, 2) - 1
            If "STU" Like "*" + ArrDg(5, Sp) + "*" Then
                'Leere Nation-Zelle bei SyPartner suchen
                    For i = 6 To UBound(ArrDg, 1) - 2
                        If ArrDg(i, Sp) <> "" _
                        And (ArrDg(i + 1, Sp + 1) <> "" Or ArrDg(i + 1, Sp - 1) <> "") _
                        And ArrDg(i + 1, Sp) = "" Then ArrDg(i + 1, Sp) = ArrDg(i, Sp)
                    Next
            End If
        Next
    'showArray2D ArrDg
End Sub

Sub T4_Add_Rank_ToAllArrDg_SyPartner()
    'Called from    FillArrDg
    
    'Vorbereitung
        Dim i%, rg%, Sp%
    'Exit
        If Not ArrDg(2, 2) Like "*(Synchron)" Then Exit Sub
    'Action
        'Rang-Spalte suchen
        For Sp = 2 To UBound(ArrDg, 2) - 1
            If ArrDg(5, Sp) = "R" Then
                'Leere Rang-Zelle bei SyPartner suchen
                    For i = 6 To UBound(ArrDg, 1) - 2
                        'If (IsNumeric(ArrDg(i, Sp)) Or ArrDg(i, Sp)="-") And ArrDg(i + 1, Sp)="" And ArrDg(i + 1, Sp + 1) <> "" Then
                        If IsNumeric(ArrDg(i, Sp)) _
                          And ArrDg(i + 1, Sp + 1) <> "" _
                          And ArrDg(i + 1, Sp) = "" Then
                            'ArrDg(i + 1, SP)=ArrDg(i, SP)    'speichert 8,899..
                            rg = CInt(ArrDg(i, Sp))
                            ArrDg(i + 1, Sp) = rg
                        End If
                    Next
            End If
        Next
End Sub

Sub T4_Add_T4Row2_ToArrDg()
    Dim i%, ArrT4Row2()
    Load_ArrT4Row2 ArrT4Row2
    'Add values to DgRow 5
        For i = 1 To UBound(ArrT4Row2, 2)
            ArrDg(5, i) = ArrT4Row2(1, i)
            If ArrDg(5, i) = "" Then ArrDg(5, i) = "-"
        Next
End Sub

Sub DoArrDg()
    'Action:        ArrDg neu füllen, falls nötig
    Dim s$
    On Error GoTo err1
    s = CStr(ArrDg(1, 1))
    Exit Sub
err1:
    FillArrDg
End Sub

Sub Load_ArrPD5(ByRef ArrPD1())
    'Called from    Add_PersonData_ToCompetitorsList, Update_T4CompetitorsList_LLinksInEFolder
    'PD1           =TextFile PersonData.txt=nur PersonenZeilen; keine TabSimulatorLeerzeichen
    'ArrPD1        =zunächst leer (income); wird hier gefüllt
    
    'Vorbereitung
        Dim L$, PD1$, v$, C1%, i%, j%, Arr1() As String, Arr2() As String
        v = vbCrLf
        PD1 = Read_Pd 'PersonData.txt einlesen (ohne Vorspann, reine Datenzeilen)
    'ArrPD1 füllen
        Arr1 = Split(v + PD1, v): ReDim ArrPD1(1 To UBound(Arr1), 1 To 9)
        For i = 1 To UBound(Arr1)
            L = Arr1(i): Arr2 = Split(L, "|")
            For j = 1 To UBound(Arr2) - 1
                ArrPD1(i, j) = Arr2(j)
            Next
        Next
    'showArray2D arrPD1 '
End Sub

Sub DoArr()
    'Called from:   [mehrere procs]
    'Action:        ArrC neu füllen, falls nicht korrekt nutzbar
    DoArrc
    'DoArrW
End Sub

Sub DoArrc()
    'Called from:   DoArr
    'Action:        ArrC neu füllen, falls nicht korrekt nutzbar
    Dim s$
    On Error GoTo err1
    s = CStr(ArrC(1))
    If s Like "*Roaming*" Then GoTo err1
    Exit Sub
err1:
    Fill_ArrCFromT8
End Sub

Sub DoArrM()
    DoArrW
End Sub

Sub DoArrW()
    'Called from:   DoArrM
    'Action:        ArrW und ArrM neu füllen, falls nicht korrekt nutzbar
    Dim s$
    On Error GoTo Err
    s = CStr(ArrX(1)): s = CStr(ArrW(1)): s = CStr(ArrM(1)): Exit Sub
Err:
    Fill_ArrX_ArrW_ArrM_FromT5
End Sub

Sub FillArrC_DgColumnNrs_FromDgRow5()
    'Called from    FillArrDg
    'T4Zeile2       BuchstabenKette; wird durch FillArrDg in DgRow5 übernommen
    'DgRow5         "W"=sDgNameW   "S"=sDgNationW   "B"=sDgClubW   "I"=sDgRangW
    '               "M"=sDgNameM   "T"=sDgNationM   "C"=sDgClubM   "J"=sDgRangM
    '               "X"=sDgNameX   "U"=sDgNationX   "D"=sDgClubX   "K"=sDgRangX  "R"=sDgRangEi
    
    'Vorbereitung
        Dim i%, sDgNameM$, sDgNameW$, sDgNameX$, sDgNationM$, sDgNationW$, sDgNationX$
        Dim sDgClubM$, sDgClubW$, sDgClubX$, sDgRangM$, sDgRangW$, sDgRangX$, sDgRangEi$
        sDgNameM = "0": sDgNameW = "0": sDgNameX = "0": sDgNationM = "0"
        sDgNationW = "0": sDgNationX = "0": sDgClubM = "0": sDgClubW = "0"
        sDgClubX = "0": sDgRangM = "0": sDgRangW = "0": sDgRangX = "0": sDgRangEi = "0"
    'Get some items
        For i = 1 To UBound(ArrDg, 2)
            If ArrDg(5, i) = "W" Then
                                          sDgNameW = CStr(i)
            ElseIf ArrDg(5, i) = "M" Then sDgNameM = CStr(i)
            ElseIf ArrDg(5, i) = "I" Then sDgRangW = CStr(i)
            ElseIf ArrDg(5, i) = "J" Then sDgRangM = CStr(i)
            ElseIf ArrDg(5, i) = "S" Then sDgNationW = CStr(i)
            ElseIf ArrDg(5, i) = "T" Then sDgNationM = CStr(i)
            ElseIf ArrDg(5, i) = "B" Then sDgClubW = CStr(i)
            ElseIf ArrDg(5, i) = "C" Then sDgClubM = CStr(i)
            ElseIf ArrDg(5, i) = "X" Then sDgNameX = CStr(i)
            ElseIf ArrDg(5, i) = "U" Then sDgNationX = CStr(i)
            ElseIf ArrDg(5, i) = "D" Then sDgClubX = CStr(i)
            ElseIf ArrDg(5, i) = "K" Then sDgRangX = CStr(i)
            End If
        Next
     'FillArrC 90-101
        FillArrC 90, sDgNameW: FillArrC 91, sDgNameM: FillArrC 92, sDgNameX: FillArrC 93, sDgNationW
        FillArrC 94, sDgNationM: FillArrC 95, sDgNationX: FillArrC 96, sDgClubW
        FillArrC 97, sDgClubM: FillArrC 98, sDgClubX: FillArrC 99, sDgRangW: FillArrC 100, sDgRangM
        FillArrC 101, sDgRangX: FillArrC 102, sDgRangEi
End Sub
       
Sub Add_OneItemToArrPd(Nn$, Vn$, ItemNr%, folderName$)
    'Called from    Fill_ArrPathsDg
    
    Dim i%
    For i = 1 To UBound(ArrPd, 1)
        If ArrPd(i, 1) = Nn Then
            If ArrPd(i, 2) = Vn Then
                'Nn, Vn sind in ArrPd enthalten
                ArrPd(i, ItemNr) = folderName
                LogBuch "ArrPd erhält in Zeile [" + Nn + ", " + Vn + "], " _
                      + "Spalte [Ordnername in 'Leute'] den Eintrag [" + folderName + "]."
                Exit Sub
            End If
        End If
    Next
End Sub

Function Get_Name_Folder_IsInArrPd(Nn$, Vn$) As String
    'Called from    Fill_ArrPathsDg
    
    Dim i%
    For i = 1 To UBound(ArrPd, 1)
        If ArrPd(i, 1) = Nn Then
            If ArrPd(i, 2) = Vn Then
                'Nn, Vn sind in ArrPd enthalten
                If ArrPd(i, 9) = "" Then
                    Get_Name_Folder_IsInArrPd = "10"
                Else
                    Get_Name_Folder_IsInArrPd = ArrPd(i, 9)
                End If
                Exit Function
            End If
        End If
    Next
    Get_Name_Folder_IsInArrPd = "00"
End Function

Function CheckIfNameIsInArrPd(Nn$, Vn$) As Boolean
    'Called from    Fill_ArrPathsDg
    
    Dim i%
    For i = 1 To UBound(ArrPd, 1)
        If ArrPd(i, 1) = Nn Then
            If ArrPd(i, 2) = Vn Then
                'Nn, Vn sind in ArrPd enthalten (kein neuer Name)
                CheckIfNameIsInArrPd = True
                Exit Function
            End If
        End If
    Next
End Function

Function Get_FolderPath_FromArrPd(Nn$, Vn$) As String
    Dim j%
    
    'Search for first Nn|Vn in ArrPd
        For j = 1 To UBound(ArrPd, 1)
            If ArrPd(j, 1) = "" Then Exit Function
            If ArrPd(j, 1) = Nn Then
                If ArrPd(j, 2) = Vn Then
                    'Nn, Vn sind in ArrPd enthalten (kein neuer Name)
                    Get_FolderPath_FromArrPd = ArrC(4) + "\" + ArrPd(j, 9)
                    Exit Function
                End If
            End If
        Next
End Function

Sub Fill_ArrPathsDg(ArrCCn, ArrPathsDg)
    'Called from    Update_Jpg_Links_Comp
    'ArrCCn         2D-Array, As String, ist bereits gefüllt: (Competitor|Club|Nation) of oneDg
    'ArrPathsDg     1D-Array, As String, zunächst leer; wird hier gefüllt;
    '               trägt dann Pfade der FolderOfDgCompetitor to write resultJPG to

    'Vorbereitung
        Dim NAME$, NamFold$, Nation$, folderName$, Nn$, pFOLD$, Verein$, Vn$, i%
        Dim NameIsInArrPd As Boolean, FoldExists As Boolean
        ReDim ArrPathsDg(1 To UBound(ArrCCn, 1))
    'Fill
        For i = 1 To UBound(ArrCCn, 1)
            'Schleife über alle CompetitorNamen eines Dg
                NAME = ArrCCn(i, 1) 'nameCompetitor: Vn Nn, "Max Maier"
                Verein = ArrCCn(i, 2):    Nation = ArrCCn(i, 3)
                Vn = Get_Vn_FromVnNn(NAME): Nn = Get_Nn_FromVnNn(NAME)
                folderName = CalCnewFolderName(Nn, Vn, Verein, Nation)
            'Check for Name and his FolderName in ArrPd
                NamFold = Get_Name_Folder_IsInArrPd(Nn, Vn) '"10" "00" "Mai, Lea (Bonn)"
                If Right(NamFold, 1) <> "0" Then
                    'Name und auch sein FolderName sind in ArrPd
                    folderName = NamFold 'überschreibt CalcedFolderName
                End If
                pFOLD = ArrC(4) + "\" + folderName 'PathOfFolder
                'SOLL-FolderName/Folderpath sind ermittelt
            'FolderName in ArrPathsDg eintragen
                ArrPathsDg(i) = pFOLD
            'Ggf. FolderName in ArrPd eintragen
                If NamFold = "10" Then
                    'Name ist in ArrPd, sein FolderName nicht
                    'FolderName in ArrPd bei Name eintragen
                        Add_OneItemToArrPd Nn, Vn, 9, folderName
                    'Ggf. Folder in Events anlegen
                       FoldExists = Create_FoldersUpToOneFolderPath(ArrC(4) + "\" + folderName)
                       If Not FoldExists Then Stop
                End If
            'Ggf. neue Zeile in ArrPd eintragen
                If NamFold = "00" Then
                    'Name ist nicht in ArrPd
                    'neue Zeile in ArrPd mit Name, ..., FolderName
                        Add_OneLineToArrPd Nn, Vn, Verein, Nation, folderName
                    'Ggf. Folder in Events anlegen
                        Create_FoldersUpToOneFolderPath ArrC(4) + "\" + folderName
                End If
                'showArray2D ArrCCn, "ArrCCn"
                'show Join(ArrPathsDg, vbCrLf)
                'show NewDataForPd
        Next
        'Die Pfade von Ordnern aller Competitors sind jetzt in ArrPathsDg gesammelt
        'showArray2D ArrPd: Stop
    'Erstellung der neuen Datei 'PersonData.txt'
        Write_Pd_FromArrPd
End Sub

Function CalCnewFolderName(Nn$, Vn$, Verein$, Nation$) As String
    'Called from    Fill_ArrPathsDg
    
    Dim s$
    s = Nn + ", " + Vn
    If Verein = "" And Nation <> "" Then s = s + " (" + Nation + ")"
    If Verein <> "" Then
        If Nation = "" Then s = s + " (" + Verein + ")"
        If Nation <> "" And Nation <> "D" Then s = s + " (" + Verein + ", " + Nation + ")"
        If Nation = "D" Then s = s + " (" + Verein + ")"
    End If
    CalCnewFolderName = s
End Function

Sub MergeDoublesInArrPd(ArrPd)
    'Called from    Update_Jpg_Links_Comp
    'Action         Mehrfach vorkommende Namen zusammenführen
    
    'Vorbereitung
        Dim D$, N$, s$, v$, c%, i%, j%, Sp%, x1%
        v = vbCrLf
        'showArray2D ArrPd, "ArrPd vor MergeDoublesInArrPd"
    'Search for Doubles
        For i = 1 To UBound(ArrPd, 1)
            If ArrPd(i, 1) <> "" Then
                N = ArrPd(i, 1) + "|" + ArrPd(i, 2)
                If InStr(1, s, N) = 0 Then
                    s = s + N + "|" + CStr(i) + "|" + v
                Else
                    'Der Name N wurde in ArrPd bereits gefunden
                    D = D + N + v 'D=Doubles
                    If x1 = 0 Then x1 = i - 1
                    '  x1=ArrPd-Index des ersten der mehrfach vorkommenden Namen
                    'Merge
                        For j = 1 To 99
                            If ArrPd(x1 + j, 1) + "|" + ArrPd(x1 + j, 2) = N Then
                                For Sp = 1 To 10 'SpaltenNr
                                    If ArrPd(x1 + j, Sp) <> "" Then ArrPd(x1, Sp) = ArrPd(x1 + j, Sp)
                                    ArrPd(x1 + j, Sp) = ""
                                Next
                            Else
                                j = 99 'Exit For j
                            End If
                        Next j
                        x1 = 0
                End If
            End If
        Next
        'showArray2D ArrPd, "ArrPd nach MergeDoublesInArrPd"
        'show D
    'ArrPd neu
        For i = 1 To UBound(ArrPd, 1)
            If ArrPd(i, 1) <> "" Then
                c = c + 1
                For j = 1 To 10
                    ArrPd(c, j) = ArrPd(i, j)
                Next
            End If
        Next
        For i = c + 1 To UBound(ArrPd, 1)
            For j = 1 To 10
                ArrPd(i, j) = ""
            Next
        Next
        'showArray2D ArrPd, "ArrPd - keine Lücken"
End Sub

Sub Fill_ArrCCn(ArrCCn)
    'Called from    Update_Jpg_Links_Comp
    'ArrCCn         2D-Array, As String, ist angelegt, noch nicht Dim
    '               wird hier gefüllt (Competitor|Club|Nation), oneDg
    
    'Vorbereitung
        Dim AllNames$, AllNameNation$, AllNameVerein$, NAME$, Nation$, Verein$
        Dim SpNamen$, SpNation$, SpVerein$, v$, v1$, ArrCCn2() As String
        Dim c%, i%, N%, Sp%, Ze%, Arr1() As String, ArrDg()
        Nation = ""
    'In welcher DgSpalte stehen Namen, Verein, Nation?
        v = vbCrLf
        
        xxx
        SpNamen = ArrC(49): SpNation = ArrC(50): SpVerein = ArrC(51)
        If SpNamen = "," Then ReDim ArrCCn(1 To 1, 1 To 3): Exit Sub
        
    'ArrDg laden (2D-Array)
        Load_ArrDesign ArrDg
    'Liga
        If ArrC(42) Like "* Liga\*" Then
            'ArrCCn2
                ReDim ArrCCn2(1 To UBound(ArrDg, 1), 1 To 3) 'Daten zu Damen/Herren + Leerzeilen
            For Ze = 6 To UBound(ArrDg, 1) - 1
                'Schleife über alle DgZeilen Spalte 4 (=NamenSpalte)
                'Verein
                    v1 = Trim(ArrDg(Ze, 3))
                    If v1 <> "" Then Verein = v1
                'Name
                    NAME = Trim(ArrDg(Ze, 4)) '= Inhalt einer Liga-NamenZelle
                    c = InStr(1, NAME, " (")              'Lea Mai (20)
                    If c > 0 Then NAME = Left(NAME, c - 1) 'Lea Mai
                If IsValidName(NAME) Then
                    i = i + 1
                    ArrCCn2(i, 1) = NAME
                    ArrCCn2(i, 2) = Verein
                    ArrCCn2(i, 3) = "D"
                End If
            Next Ze
        
        
    'noLiga
        Else
            ReDim ArrCCn2(1 To 2 * UBound(ArrDg, 1), 1 To 3) 'Daten zu Damen/Herren + Leerzeilen
            Arr1 = Split(SpNamen, ",")
            For Sp = 1 To UBound(Arr1) - 1      ',4,   ',3,7,   ',3,7,11,15,
                'Schleife über alle Spalten des aktuellen Designs, die als NamenSpalte erkannt wurden
                If IsNumeric(Arr1(Sp)) Then
                    N = CInt(Arr1(Sp))             'N=SpNr einer NamenSpalte
                    'Namensliste beginnt immer in Zeile 6 (je mit Vorname+Nachname)
                    For Ze = 6 To UBound(ArrDg, 1) - 1
                        'Schleife über alle Zeilen einer NamenSpalte
                        'Bearbeitung einer einzelnen DesignZeile (mit Namen, Rang, ...)
                        NAME = Trim(ArrDg(Ze, N)) '= Inhalt einer NamenZelle
                        If IsValidName(NAME) Then
                            c = InStr(1, NAME, " (")              'Brett Austine (20)
                            If c > 0 Then NAME = Left(NAME, c - 1) 'Brett Austine
                            If InStr(1, AllNames, NAME) = 0 Then
                                'keine NameDupletten (Einzel+Synchron in one Design)
                                AllNames = AllNames + NAME + v
                                i = i + 1
                                ArrCCn2(i, 1) = NAME
                                If InStr(1, SpVerein, "," + CStr(N + 1) + ",") > 0 Then
                                    'falls VereinSpalte rechts neben der NamenSpalte
                                    Verein = Trim(ArrDg(Ze, N + 1))
                                    ArrCCn2(i, 2) = Verein
                                End If
                                If InStr(1, SpNation, "," + CStr(N + 1) + ",") > 0 Then
                                    'falls NationSpalte rechts neben der NamenSpalte
                                    Nation = Trim(ArrDg(Ze, N + 1))
                                    ArrCCn2(i, 3) = Nation
                                End If
                            End If
                        End If
                    Next Ze
                End If
            Next Sp
        End If
    'Leerzeilen vermeiden
        If i = 0 Then i = 1
        ReDim ArrCCn(1 To i, 1 To 3)
        For c = 1 To i
            ArrCCn(c, 1) = ArrCCn2(c, 1)
            ArrCCn(c, 2) = ArrCCn2(c, 2)
            ArrCCn(c, 3) = ArrCCn2(c, 3)
        Next
End Sub

Sub Format_T8_ArrC_Blue()
    Dim i%, r1 As Range, r2 As Range
    With Sheets("T8")
        Set r1 = .Cells(62, 4) 'Zellinhalt der Länge 1
        Set r2 = .Cells(7, 4) 'Zellinhalt der Länge > 10
        For i = 2 To UBound(ArrC)
            If Len(ArrC(i)) < 10 Then
                Set r1 = Union(r1, .Cells(6 + i, 4))
            Else
                Set r2 = Union(r2, .Cells(6 + i, 4))
            End If
        Next
    End With
    r1.HorizontalAlignment = xlCenter
    r2.HorizontalAlignment = xlLeft
End Sub

Function Sort_TextLines(SomeTextLines$, Optional delimiter$ = "") As String
    'Called from    Add_OneLineToPersonData
    'Action         Realisiert alphabetische Sortierung; AaäBbCc...
    
    Dim s$, ss$, c%, i%, Arr1() As String
        Arr1 = Split(SomeTextLines$, vbCrLf)
        For i = 0 To UBound(Arr1)
            'ss=20 Zeichen Sortstring; wird jeder Zeile vorangestellt
            ss = LCase(Left(Arr1(i), 20)) 'SortString
'            sS=Replace(sS, "ä", "ae"): sS=Replace(sS, "ö", "oezz")
'            sS=Replace(sS, "ü", "ue"): sS=Replace(sS, "ß", "ss")
            ss = Replace(ss, "ä", "aezz"): ss = Replace(ss, "ö", "oezz")
            ss = Replace(ss, "ü", "uezz"): ss = Replace(ss, "ß", "sszz")
            ss = Replace(ss, "é", "e")
            'Delimiter="|" --> "Bar|" > "Bardy|", aber "Bar |" < "Bardy |" (= SOLL-Wert)
            If delimiter <> "" Then ss = Replace(ss, delimiter, " " + delimiter)
            ss = Left(ss, 20)
            Arr1(i) = ss + Arr1(i)
        Next
        QuickSort Arr1
        
        
'showArray Arr1
        
        
        For i = 0 To UBound(Arr1)
            s = s + Mid(Arr1(i), 21) + vbCrLf
        Next
        s = Delete_EmptyEndRowsInString(s)
        Sort_TextLines = s
End Function

Function getPureTextFrom2DArray(Arr)
    Dim s$, i%, j%
    s = "|"
    For i = 1 To UBound(Arr, 1)
        For j = 1 To UBound(Arr, 2)
            s = s + Trim(Arr(i, j)) + "|"
        Next
        s = s + vbCrLf + "|"
    Next
    If Right(s, 3) = vbCrLf + "|" Then s = Left(s, Len(s) - 3)
    getPureTextFrom2DArray = s
End Function


Sub PasteArrC()
    Paste_1DArrayToCol "T8", 7, 4, ArrC
    Format_T8_ArrC_Blue
End Sub

Sub Fill_ArrCFromT8() '--Container--
    'Called from:   Workbook_Open, doArr
    'ArrC          =Container
    '               Erweiterung durch händisches Hinzufügen von Zeilen
    '                mit fortlaufenden Nummern in den Green3Zellen
    'Action:        Ermittelt erste und letzte blaue Zelle in W2
    '               (ZeilenNr LetzteBlaueZelle=ZeilenNr LetzteGreen3ZelleMitNr)
    '               Dimensioniert arrC (1 to AnzahlBlaueZellenInSheetT3)
    '               Schreibt diese Zellen in arrC

    Dim i%, BlueS%, BlueZ2%, BlueZ1%, Arr1() As Variant
    'Einstellung
        BlueZ1 = 7 'BlueZ1=ZeilenNr  der 1. blauen Zelle
        BlueS = 4 'BlueS =SpaltenNr der blauen Zellen
    'ZeilenNr  der letzten blauen Zelle, die eine Nummer trägt
        BlueZ2 = Get_NrOfLastRowInColumnNr(CLng(BlueS - 1), "T8")
    'hellblaue Zellen in T8-Spalte "D" in ein 2D-Array übernehmen:
        With ThisWorkbook.Sheets("T8")
        Arr1 = .Range(.Cells(BlueZ1, BlueS), .Cells(BlueZ2, BlueS)).Value
        End With
    'Variable "arrC" belegen:
        ReDim ArrC(1 To UBound(Arr1, 1))
        For i = 1 To UBound(Arr1, 1)
            ArrC(i) = Arr1(i, 1)
        Next
    'NoRoamingPath
        If ArrC(1) Like "*Roaming*" Then ArrC(1) = ArrC(10)
End Sub

Sub Fill_ArrCPaths()
    If ActiveWorkbook.path Like "*Roaming*" Then Stop
    FillArrC 1, Get_ParentFolderPath_OfFolderPath(ActiveWorkbook.path)   '"F:\Archiv Trampolin prog"
    FillArrC 2, Replace(ArrC(1), "prog", "1900-1999")
    FillArrC 3, ArrC(2) + "\Events"
    FillArrC 4, ArrC(2) + "\Leute"
    FillArrC 5, ArrC(2) + "\Register"
    FillArrC 6, ArrC(1) + "\prog\helpers"
    FillArrC 7, ArrC(2) + "\Read me.txt"
    FillArrC 8, ArrC(6) + "\es.exe"
    FillArrC 9, ArrC(6) + "\exiftool.exe"
    FillArrC 10, ArrC(1) + "\prog\Bes TR.xlsm"
    FillArrC 11, ArrC(2) + "\ClubsNations"
    'If Not FolderExists(ArrC(1)) Then Stop
End Sub

Sub Fill_ArrSex(arrSex)
    'Action     m/w aufgrund der Hintergrundfarbe feststellen
    
    'Vorbereitung
        Dim i%, j%, s1%, s2%, z1%, z2% ', arrSex() As String
        z1 = CInt(ArrC(37))       'T4 ZeilenNr   of BorderTop
        z2 = CInt(ArrC(38))       'T4 ZeilenNr   of BorderBottom
        s1 = CInt(ArrC(39))       'T4 SpaltenNr  of BorderLeft
        s2 = CInt(ArrC(40))       'T4 SpaltenNr  of BorderRight
        'ReDim arrSex(1 To z2 - z1 + 1, 1 To s2 - s1 + 1)
        'ReDim arrSex(1 To CInt(ArrC(38)) - CInt(ArrC(37)) + 1, 1 To CInt(ArrC(40)) - CInt(ArrC(39)) + 1)
    'Array erstellen
    For i = 1 To z2 - z1 + 1
        For j = 1 To s2 - s1 + 1
            If Cells(i + z1 - 1, j + s1 - 1).Interior.Color = 14083324 Then arrSex(i, j) = "w"
            If Cells(i + z1 - 1, j + s1 - 1).Interior.Color = 15652797 Then arrSex(i, j) = "m"
        Next j
    Next i
    'showArray2D arrSex
End Sub

Function IgnoreThisName(NAME$) As Boolean
    IgnoreThisName = True
    If NAME = "" Then Exit Function
    If NAME Like "*Ti" Then Exit Function
    If NAME Like "*Tu" Then Exit Function
    If NAME Like "*Juti*" Then Exit Function
    If NAME Like "*Jutu*" Then Exit Function
    If NAME Like "*Schi" Then Exit Function
    If NAME Like "*Schü" Then Exit Function
    If NAME Like "*Meisterschaft*" Then Exit Function
    If NAME Like "*klasse*" Then Exit Function
    If NAME Like "*Aufbau*" Then Exit Function
    If NAME Like "*Damen*" Then Exit Function
    If NAME Like "*Herren*" Then Exit Function
    IgnoreThisName = False
End Function

Sub Fill_ListOfCompetitors(ListOfCompetitors$, AllNnVn_OfOneDesign$)
    'Called from    Update_T4CompetitorsList_LLinksInEFolder
    'Status         Ein bestimmtes Design ist ausgewählt, AnalyzeDesign-Daten vorhanden
    'Action         Ergänzt ListOfCompetitors um weitere Zeilen (je nach AnzahlLeute/aktuellemDesign)
    
    'Vorbereitung
        Dim Alter$, k$, Klasse$, Nachname$, NAME$, OneLine$, Rang$, Sex$
        Dim SpNamen$, SpNation$, SpRang$, SpVerein$
        Dim LTV$, Nation$, Titel$, Typ$, Vorname$, Verein$, Verein1$, Verein2$
        Dim ArrDesign(), Arr1() As String, arrSex() As String
        Dim c%, C2%, s1%, s2%, z1%, z2%, Ze%, Sp%, N%, ZeTrenn%
        z1 = CInt(ArrC(37)): z2 = CInt(ArrC(38)): s1 = CInt(ArrC(39)): s2 = CInt(ArrC(40))
        Titel = ArrC(41)          'T4 Titel      to see inside design and jpg
        
        Stop 'ArrC(47) ... existiert so nicht mehr
        
        SpRang = ArrC(47)         'T4 SpaltenNrn der Rang  -Spalten ',3,8,
        SpNamen = ArrC(49)        'T4 SpaltenNrn der Namen -Spalten ',3,7,
        SpNation = ArrC(50)       'T4 SpaltenNrn der Nation-Spalten
        SpVerein = ArrC(51)       'T4 SpaltenNrn der Verein-Spalten ',3,8,
        ReDim arrSex(1 To z2 - z1 + 1, 1 To s2 - s1 + 1)
    
    'arrDesign laden (2D-Array)
        Load_ArrDesign ArrDesign
    'arrDesign ergänzen
        Call AddSomeValuesToArrDesign(ArrDesign)
        'showArray2D arrDesign ': Stop
    'arrSex laden (2D-Array)
        Fill_ArrSex arrSex()
        'showArray2D arrSex
    'ListOfCompetitors-TextZeilen erstellen
        'Designs können mehrere Namen-Spalten enthalten
        '  (Kopf: Schü.. Schi.. Jutu.. Juti.. Tu.. Ti.. Damen Herren Aktive)
        '  Dadurch ggf. auch mehrere Verein-/Rang-/Werte-Spalten
        'Designs, die Nachname, Vorname enthalten:
        Arr1 = Split(SpNamen, ",")
        For Sp = 1 To UBound(Arr1) - 1      ',3,7,   ',3,7,11,15,
            'Schleife über alle Spalten des aktuellen Designs, die als NamenSpalte erkannt wurden
            If IsNumeric(Arr1(Sp)) Then
                N = CInt(Arr1(Sp))             'N=Dg-SpNr einer NamenSpalte
                'Namensliste beginnt immer in Zeile 6 (je mit Vorname+Nachname)
                For Ze = 6 To UBound(ArrDesign, 1) - 1
                    'Schleife über alle Zeilen einer NamenSpalte innerhalb eines Designs
                    'Bearbeitung einer einzelnen DesignZeile (mit Namen, Rang, ...)
                    NAME = Trim(ArrDesign(Ze, N)) '= Inhalt einer NamenZelle
                    'Design könnte temporär (während des Eintragens von Namen)
                    '   noch Leerzeilen enthalten
                        If IgnoreThisName(NAME) Then GoTo jump1
                    c = InStr(1, NAME, " ")     'Brett Austine (20)
                    If c > 0 Then
                        'Nur Zeilen mit Vor- UND Nachnamen werden berücksichtigt
                '       Alter
                        C2 = InStr(1, NAME, " (")
                        If C2 > 0 Then Alter = Mid(NAME, C2 + 2, 2): NAME = Left(NAME, C2 - 1)
                '(1)    Vorname, Nachname
                        Vorname = Left(NAME, c - 1): Nachname = Mid(NAME, c + 1)
                '(2)    Sex
                        Sex = arrSex(Ze, N)
                '(3)    Typ
                        If Mid(ArrDesign(Ze, 1), 1, 1) = "E" Then Typ = "Einzel"
                        If Mid(ArrDesign(Ze, 1), 1, 1) = "Y" Then Typ = "Synchron"
                        If Mid(ArrDesign(Ze, 1), 1, 1) = "M" Then Typ = "Mannschaft"
                        If ArrC(41) Like "*(Einzel)" Then Typ = "Einzel"
                        If ArrC(41) Like "*(Synchron)" Then Typ = "Synchron"
                        If ArrC(41) Like "*(Mannschaft)" Then Typ = "Mannschaft"
                '(4)    Klasse
                            k = ArrDesign(4, N)
                        If Left(k, 8) = "Schüleri" Then Klasse = "Schi"
                        If k = "Schüler" Then Klasse = "Schü"
                        If Left(k, 9) = "Schüler (" Then Klasse = "Schü"
                        If Mid(k, 13, 1) = "Ze" Then Klasse = "Juti"
                        If k = "Jugendturner" Then Klasse = "Jutu"
                        If Mid(k, 12, 3) = "r (" Then Klasse = "Jutu"
                        If Left(k, 7) = "Turneri" Then Klasse = "Ti"
                        If k = "Turner" Then Klasse = "Tu"
                        If Left(k, 8) = "Turner (" Then Klasse = "Tu"
                        If Left(k, 5) = "Damen" Then Klasse = "Ti"
                        If Left(k, 5) = "Herre" Then Klasse = "Tu"
                            'Für MannschaftsWK
                            k = ArrDesign(4, N - 1)
                        If k = "Schülerinnen" Then Klasse = "Schi"
                        If k = "Schüler" Then Klasse = "Schü"
                        If k = "Jugendturnerinnen" Then Klasse = "Juti"
                        If k = "Jugendturner" Then Klasse = "Jutu"
                        If k = "Turnerinnen" Then Klasse = "Ti"
                        If k = "Turner" Then Klasse = "Tu"
                        If k = "Damen" Then Klasse = "Ti"
                        If k = "Herren" Then Klasse = "Tu"
                            'Für Sonderfall
                        If Titel = "1966-05-08 DJM Nürnberg" And N = 4 Then
                            If Ze > 43 And Ze < 52 Then Klasse = "Schi": Rang = "-"
                            If Ze > 51 Then Klasse = "Ju Mix" 'Synchron Mixed
                        End If
                '(5)    Verein, Nation
                        'Falls die VereinSpalte rechts neben der NamenSpalte steht:
                        If InStr(1, SpVerein, "," + CStr(N + 1) + ",") > 0 Then Verein = ArrDesign(Ze, N + 1)
                        'Falls die VereinSpalte links  neben der NamenSpalte steht:
                        If InStr(1, SpVerein, "," + CStr(N - 1) + ",") > 0 Then Verein = ArrDesign(Ze, N - 1)
                        'Falls die NationSpalte rechts neben der NamenSpalte steht:
                        If InStr(1, SpNation, "," + CStr(N + 1) + ",") > 0 Then Nation = ArrDesign(Ze, N + 1)
                        'Falls die NationSpalte links  neben der NamenSpalte steht:
                        If InStr(1, SpNation, "," + CStr(N - 1) + ",") > 0 Then Nation = ArrDesign(Ze, N - 1)
                '(6)    Rang
                        'Falls die RangSpalte direkt           links neben der NamenSpalte steht:
                        If InStr(1, SpRang, "," + CStr(N - 1) + ",") > 0 Then Rang = ArrDesign(Ze, N - 1)
                        'Falls die RangSpalte 2 Spalten weiter links neben der NamenSpalte steht:
                        If InStr(1, SpRang, "," + CStr(N - 2) + ",") > 0 Then Rang = ArrDesign(Ze, N - 2)
                        If Rang = "" Then Rang = "-"
                        If Len(SpRang) = 1 Then Rang = "-"
                '(7)    OneLine
                        If ZeTrenn = 0 Then 'keine Liga, keine Trennlinie
                            '|Nachname|Vorname|w/m|Klasse|Verein|LTV|Nat|Event|Typ|Rang|
                            OneLine = "|" + Nachname + "|" + Vorname + "| " + Sex + " |" + Alter _
                                   + " |" + Klasse + "|" + Verein + "|" + LTV + "|" + Nation _
                                   + "| " + Titel + " |" + Typ + "|" + Rang + "|"
                            AllNnVn_OfOneDesign = AllNnVn_OfOneDesign + Nachname + ", " + Vorname + vbCrLf
                        Else 'Liga
                            'Der jew. VereinsName "TV xx" ist im Design als "· TV xx ·" angegeben
                            '   (da er in der NamenSpalte steht, aber nicht als PersonenName gelten soll)
                                Verein1 = Trim(Replace(ArrDesign(ZeTrenn - 1, N), "·", ""))
                                Verein2 = Trim(Replace(ArrDesign(ZeTrenn + 1, N), "·", ""))
                                If Ze < ZeTrenn Then Verein = Verein1 Else Verein = Verein2
                            If NAME <> ArrDesign(ZeTrenn - 1, N) And NAME <> ArrDesign(ZeTrenn + 1, N) Then
                                'Einzel
                                OneLine = "|" + Nachname + "|" + Vorname + "| " + Sex + " |" + Alter + "|" + Klasse + "|" + Verein + "|" + LTV + "|" + Nation + "| " + Titel + " |" + Typ + "|" + Rang + "|"
                                ListOfCompetitors = ListOfCompetitors + vbCrLf + OneLine + vbCrLf
                                'Mannschaft
                                Typ = "Mannschaft"
                                If Ze < ZeTrenn Then
                                    Rang = CStr(ArrDesign(ZeTrenn - 1, N - 1))
                                Else
                                    Rang = CStr(ArrDesign(ZeTrenn + 1, N - 1))
                                End If
                                If Nachname Like "*·*" Then
                                    GoTo jump1
                                Else
                                    OneLine = "|" + Nachname + "|" + Vorname + "| " + Sex + " |" + Alter + "|" + Klasse + "|" + Verein + "|" + LTV + "|" + Nation + "| " + Titel + " |" + Typ + "|" + Rang + "|"
                                    AllNnVn_OfOneDesign = AllNnVn_OfOneDesign + Nachname + ", " + Vorname + vbCrLf
                                End If
                            End If
                        End If
                        'show OneLine
                '(8)    Zeile --> ListOfCompetitors
                        ListOfCompetitors = ListOfCompetitors + vbCrLf + OneLine + vbCrLf
                    End If
jump1:
                Next Ze
                    'show TAB_Simulation(ListOfCompetitors)
            End If
        Next Sp
'show s

End Sub

Sub CheckCells_ListOfDesignTitles_ListOfCompetitors()
    'Called from:   Worksheet_SelectionChange[T4]
    'Action         Verschiebung entdecken, ggf. neue Werte in ArrC schreiben
    '               bei Zelle mit "List of DesignTitles", mit "List of competitors"
    
    'Vorbereitung
        Dim MAA$, sCp%, s2Cp%, sDg%, zCp%, zDg%, DoFill As Boolean
        zDg = Get_RowNr_HoldingMyTextWhole("T4", "List of DesignTitles")         '5
        sDg = Get_ColumnNr_HoldingMyTextWhole("T4", "List of DesignTitles")      '2
        
        
        zCp = Get_RowNr_HoldingMyTextWhole("T4", "List of competitors")          '5
        sCp = Get_ColumnNr_HoldingMyTextWhole("T4", "List of competitors")       '5
            MAA = Sheets("T4").Cells(zCp, sCp).MergeArea.Address '$E$5:$O$5
        s2Cp = Range(Mid(MAA, InStr(1, MAA, ":") + 1)).Column             '15

    'Check
        If ArrC(71) = "" Or ArrC(75) = "" Then
            DoFill = True '(72), (73), (74) sind dann auch leer
        Else
            If zDg <> CInt(ArrC(71)) Or sDg <> CInt(ArrC(72)) Then DoFill = True
            If zCp <> CInt(ArrC(73)) Or sCp <> CInt(ArrC(74)) Then DoFill = True
            If s2Cp <> CInt(ArrC(75)) Then DoFill = True
        End If
    'Action
        If DoFill Then
            FillArrC 71, CStr(zDg): FillArrC 72, CStr(sDg)
            FillArrC 73, CStr(zCp): FillArrC 74, CStr(sCp): FillArrC 75, CStr(s2Cp)
        End If
End Sub

Sub FillArrC_T3PositionsOfListOfCompetitors()
    'Called from:   Worksheet_SelectionChange[T4]
    'Action         Verschiebung entdecken, ggf. neue Werte in ArrC schreiben
    '               bei Zelle mit "List of DesignTitles", mit "List of competitors"
    
    'Vorbereitung
        Dim MAA$, sCp%, s2Cp%, sDg%, zCp%, zDg%, DoFill As Boolean
        zCp = 2 'Get_RowNr_HoldingMyTextWhole("T4", "CompetitorsList")            '2
        sCp = 2 'Get_ColumnNr_HoldingMyTextWhole("T3", "CompetitorsList")         '2
            MAA = Sheets("T3").Cells(zCp, sCp).MergeArea.Address '$E$5:$O$5
        s2Cp = Range(Mid(MAA, InStr(1, MAA, ":") + 1)).Column             '15

    'Check
        If ArrC(75) = "" Then
            DoFill = True '(73), (74) sind dann auch leer
        Else
            If zCp <> CInt(ArrC(73)) Or sCp <> CInt(ArrC(74)) Then DoFill = True
            If s2Cp <> CInt(ArrC(75)) Then DoFill = True
        End If
    'Action
        If DoFill Then
            FillArrC 73, CStr(zCp): FillArrC 74, CStr(sCp): FillArrC 75, CStr(s2Cp)
        End If
End Sub


Sub T4_Save_PositionOfHeaderCell_OfListOfDesignTitles_ToArrC()
    'Called from:   Worksheet_SelectionChange[T4]
    'Action         Verschiebung der KopfZelle "List of DesignTitles" entdecken
    '               ggf. neue Werte in ArrC schreiben
    
    'Vorbereitung
        Dim sDg%, zDg%, DoFill As Boolean
        zDg = Get_RowNr_HoldingMyTextWhole("T4", "List of DesignTitles")         '5
        sDg = Get_ColumnNr_HoldingMyTextWhole("T4", "List of DesignTitles")      '2
    'Check
        'zLiDg=CInt(ArrC(71))      'T4 ZeilenNr  Cell "List of DesignTitles"
        'sLiDg=CInt(ArrC(72))      'T4 SpaltenNr Cell "List of DesignTitles"
        If ArrC(71) = "" Then
            DoFill = True '(72) ist dann auch leer
        Else
            If zDg <> CInt(ArrC(71)) Or sDg <> CInt(ArrC(72)) Then DoFill = True
        End If
    'Action
        If DoFill Then
            FillArrC 71, CStr(zDg): FillArrC 72, CStr(sDg)
        End If
End Sub

Sub TimeToMillisecond_TEST()
    show TimeToMillisecond
End Sub
Public Function TimeToMillisecond() As String
    Dim tSystem As SYSTEMTIME, sRet
    On Error Resume Next
    GetSystemTime tSystem
    sRet = Hour(Now) & ":" & Minute(Now) & ":" & Format(Second(Now), "00") & _
    "." & Format(tSystem.wMilliseconds, "000")
    TimeToMillisecond = sRet
End Function
    
Sub ShowSomeTicks()
    show SomeTicks
End Sub

Sub KillSomeTicksContent()
    SomeTicks = ""
End Sub

Sub Ticks(s$)
    SomeTicks = SomeTicks + CStr(GetTickCount()) + " | " + TimeToMillisecond + " | " + s + vbCrLf
End Sub

Sub TEST_MergeArea()
    show Sheets("T4").[JR13].MergeArea.Address '$JR$13:$JR$33
End Sub

Sub MergeSomeCells(NameOfSheet, ByVal z1&, ByVal s1&, ByVal z2&, ByVal s2&)
    With Sheets(NameOfSheet).Range(Cells(z1, s1), Cells(z2, s2))
        .Merge: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlTop
    End With
End Sub


