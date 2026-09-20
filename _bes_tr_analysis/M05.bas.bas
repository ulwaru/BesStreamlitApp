Attribute VB_Name = "M05"
Option Explicit 'M05

Public EE2_Holder As Boolean

Sub zzz_M05()
    showProcs "size"
    Application.EnableEvents = True
    'RenameModule "W2", "T3"
    'Show Get_ImageWidthHeight("G:\Archiv VBA\Bes FaceTags\1.JPG")
    'DrawW1Line "hallo", 600, 100, 700, 200
    'Get_TextWithTabLikeSpaces_LeftToEachStringX
    'showArray2D ArrT
End Sub

Sub Update_Terminkalender()
    'Called from    xxx
    'Tk             Terminkalender; wird aus Tk9 (Daten zu EvFol+EvTxt) erstellt
    'TkRaster       Einheitliches Format zur Erstellung des Terminkalenders Tk
    '               je RasterZeile: |Datum|Titel|Ort|Bem|Quelle|+|NameEventOrdner|Pfad|
    '               1 Termin kann zunächst aus mehreren RasterZeilen bestehen
    '               letztlich soll entstehen: 1 Termin - 1 RasterZeile
    'EvFol          Namen der EventOrdner;      werden ausgelesen und in TkRaster gebracht
    'EvTxt          bisherige Textversion des Tk; wird ausgelesen und in TkRaster gebracht
    'TkT2           Terminkalender auf dem Sheet Tk2; wird nicht ausgelesen; aus Tk9 erstellt
    'Tk1             = TkRaster Version1; EvFol+EvTxt im TkRaster; alle Termine
    'Tk2             = Teil2 von Tk1 = Alle RasterZeilen mit: Datum kommt in Tk1 mehrmals vor
    '                (Teil1 von Tk1 --> Tk9)
    'Tk3             = Teil1 von Tk2 = Liga; TkReady
    'Tk4             = Teil2 von Tk2 = MultiDate
    'Tk5             = Tk4, wobei: EvFol-/EvTxt-RasterZeilen zu 1 Event --> 1 RasterZeile
    'Tk9             = Endversion Termine im TkRaster; 1 RasterZeile pro Termin
    '               = (Teil1 von Tk1) + (...)
    
    'Vorbereitung
        Dim T$, Tk$, Tk1p$, Tk2p$, Tk1$, Tk2$, Tk3$, Tk4$, Tk5$, Tk9$, v$
        Dim AnzTk1%, AnzTk2%, AnzTk3%, AnzTk4%, AnzTk9%
        Dim c&, i%, Arr1() As String
        v = vbCrLf: DoArr
    'Tk1 - 2 Raster laden
        Tk1 = Get_Raster_EvFol + Get_Raster_EvTxt ': Show_Raster_EvFol ': Show_Raster_EvTxt
        Tk1 = Delete_EmptyRowsInString(Tk1)
        Arr1 = Split(Tk1, v): QuickSort Arr1: Tk1 = Join(Arr1, v)
        AnzTk1 = UBound(Arr1) + 1
        'ShowTk1 Tk1
    'Tk9        = Teil1 von Tk1 = Alle Termine mit: Datum kommt in Tk1 nur 1x vor
        Tk1p = Get_Tk1Parts(Tk1): c = InStr(1, Tk1p, "°°")
        Tk9 = Mid(Tk1p, 1, c - 1): Tk9 = Delete_EmptyRowsInString(Tk9)
        AnzTk9 = anzAinB(v, Tk9) + 1
        'ShowTk9 Tk9
    'Tk2        = Teil2 von Tk1 = Alle RasterZeilen mit: Datum kommt in Tk1 mehrmals vor, Liga
        Tk2 = Mid(Tk1p, c + 2): AnzTk2 = anzAinB(v, Tk2) + 1
        'ShowTk2 Tk2, AnzTk1, AnzTk2, AnzTk9
    'Tk3       = Tk2.1 = Liga = TkReady
        Tk2p = Get_Tk2Parts(Tk2): c = InStr(1, Tk2p, "°°")
        Tk3 = Mid(Tk2p, 1, c - 1): AnzTk3 = anzAinB(v, Tk3) + 1
    'Tk4       = Tk2.2 = MultiDate
        Tk4 = Mid(Tk2p, c + 2): AnzTk4 = anzAinB(v, Tk4) + 1
        'ShowTk3 Tk3, AnzTk1, AnzTk3, AnzTk4, AnzTk9
        'ShowTk4 Tk4, AnzTk1, AnzTk3, AnzTk4, AnzTk9
    'Tk5       = Tk4 ohne Doubles
        Tk5 = Get_Tk5(Tk4)
    'Tk9       = T9+T3+T5 = neue Version von Tk9
        Tk9 = Tk9 + v + Tk3 + v + Tk5
    'Tk
        Tk = Get_TkFromTk9(Tk9)
        't = "Anzahl Doubles     = " + ArrC(58) + v + "Anzahl Tk-Zeilen   = " + CStr(anzAinB(v, Tk) + 1) + v + "Anzahl StartZeilen = " + CStr(AnzTk1) + v + v
        'show t + Tk
       writeStringToFile ArrC(2) + "\z Terminkalender.txt", Tk
End Sub

Function Get_TkFromTk9(T9$) As String
    'Vorbereitung
        Dim Anz$, D$, K1$, K2$, L$, s$, s1$, T$, v$, y$, z$
        Dim C1%, C2%, C3%, i%, Arr1() As String, Arr2() As String
        v = vbCrLf: s1 = TAB_Simulation(T9): Arr1 = Split(s1, v)
        Anz = CStr(UBound(Arr1) + 1)
        QuickSort Arr1
    's1 anzeigen
        's1 = Join(arr1, v): show s1
    'RasterZeilen --> Termine
        s = "||Datum|Event|Ort|Bemerkung|Name des Ordners|" + v
        For i = 0 To UBound(Arr1)
            z = Arr1(i)         '1 Zeile
            Arr2 = Split(z, "|")
            s = s + Arr2(1) + "|" + Arr2(6) + "|" + Change_Date8ToDate10(Arr2(1)) + "|" _
                  + Arr2(2) + "|" + Arr2(3) + "|" + Arr2(4) + "|" + Arr2(7) + "|" + v
        Next
        s = TAB_Simulation(s): Arr1 = Split(s, v): s = ""
        L = String(Len(Arr1(0)) - 8, "-") 'TrennLinie; "---"-Zeile
        
        QuickSort Arr1
    'Date8-Sortierspalte entfernen, TrennLinien einfügen
        For i = 0 To UBound(Arr1)
            z = Mid(Arr1(i), 9)         'Z ohne Date8 am Anfang
            If i = 0 Then s = z + v
            If i > 0 Then
                If y = Mid(z, 10, 4) Then
                    s = s + z + v
                Else
                    s = s + L + v + z + v: y = Mid(z, 10, 4)
               End If
            End If
        Next
    'Kopfzeile: Items mittig
            K1 = Mid(Arr1(0), 9) 'Kopfzeile; "...|Datum     |..."
            K2 = "| |": Arr2 = Split(K1, "|")
            For i = 2 To 6
                z = Arr2(i)
                C3 = Len(z): C1 = Len(Trim(z)): C2 = C3 - C1 'c1|c2|c3=LenTxt|AnzLeerzeichen|LenGes
                K2 = K2 + String(C2 \ 2, " ") + Left(z, C1) + String(C2 - C2 \ 2, " ") + "|"
            Next
            s = Replace(s, K1, K2)
    'Finale
        s = "TERMINKALENDER 1958-1999" + String(Len(L) - 108, " ") _
          + "(" + Anz + " Termine; + = zu diesem Event existiert ein Ordner in 'Archiv Trampolin\Events')" + v + v + s + L
        Get_TkFromTk9 = s
End Function

Function Get_T5(T4) As String
    'Action     T4-Rasterzeilen werden getrennt in T5 (ready) und T6 (VnNn)
    'Z1, Z2     = 2 aufeinanderfolgenden RasterZeilen
    '           Falls Z1, Z2: mit |EvFol|, dann Z1=ready
    '           Falls Z1, Z2: Spalte1 bis Spalte3 gleich, dann Z1+Bem=ready, Z2=Duplette
        
    'Vorbereitung
        Dim a13$, a58$, b13$, b58$, Bem$, Doubles$, T5$, T6$, v$, z1$, z2$, i%
        Dim Arr1() As String, Arr2() As String, Arr3() As String
        v = vbCrLf
    'Action
        Arr1 = Split(T4 + v, v)
        For i = 0 To UBound(Arr1) - 1
            z1 = Arr1(i): z2 = Arr1(i + 1) 'Eine Zeile und ihre NachfolgeZeile
            Arr2 = Split(z1, "|"): Arr3 = Split(z2, "|")
            a13 = "|" + Arr2(1) + "|" + Arr2(2) + "|" + Arr2(3) + "|" 'Spalte1 bis Spalte3 von Z1
            b13 = "|" + Arr3(1) + "|" + Arr3(2) + "|" + Arr3(3) + "|" 'Spalte1 bis Spalte3 von Z2
            a58 = "|" + Arr2(5) + "|" + Arr2(6) + "|" + Arr2(7) + "|" + Arr2(8) + "|" 'Sp5 bis Sp8 von Z1
            b58 = "|" + Arr3(5) + "|" + Arr3(6) + "|" + Arr3(7) + "|" + Arr3(8) + "|" 'Sp5 bis Sp8 von Z2
            If z1 Like "*|EvFol|*" And z2 Like "*|EvFol|*" Then
                T5 = T5 + z1 + v
            ElseIf a13 = b13 Then
                'Spalte1 bis Spalte3 sind bei beiden RasterZeilen gleich
                If Len(Arr2(4)) > Len(Arr3(4)) Then Bem = Arr2(4) Else Bem = Arr3(4)
                T5 = T5 + a13 + Bem + a58 + v               'Spalte4 = Bem = Bemerkung
                Doubles = Doubles + z1 + v + z2 + v + v
                i = i + 1 'nächste Zeile überspringen
            Else
                T6 = T6 + z1 + v
            End If
        Next
        T5 = Delete_EmptyEndRowsInString(T5 + v + T6)
        ArrC(58) = CStr(CInt(ArrC(58)) + anzAinB("|", Doubles) / 18)
        Get_T5 = T5
        'show CStr(anzAinB("|", Doubles) / 18) + " T4-Doubles" + v + v + Doubles
        'show "T5 = T4.1 ready - " + CStr(anzAinB("|", T5) / 9) + v + v + T5
        'show "T6 = T4.2 MultiDate - " + CStr(anzAinB("|", T6) / 9) + v + v + T6
End Function

Function Get_Tk2Parts(Tk2$) As String
    'Action     Tk2-Rasterzeilen werden getrennt in Tk3 (Liga) und Tk4 (MultiDate)
    
    'Vorbereitung
        Dim Doubles$, Tk3$, Tk4$, v$, z1$, z2$, C3%, i%, Arr1() As String, Arr2() As String
        v = vbCrLf
    'Action
        Arr1 = Split(Tk2 + v, v)
        For i = 0 To UBound(Arr1) - 1
            z1 = Arr1(i): z2 = Arr1(i + 1)
            Arr2 = Split(z1, "|")
            If Arr2(4) Like "*[a-z]-[A-Z]*" And InStr(1, Arr2(4), " ") = 0 Then
                'Spalte4 (Bemerkung) enthält die Vereinsnamen einer Liga-Paarung
                C3 = InStr(1, z1, "|EvTxt|"): If C3 = 0 Then C3 = InStr(1, z1, "|EvFol|")
                If Left(z1, C3) = Left(z2, C3) Then
                    'Z2 ist eine Duplette; wird nicht übernommen (nicht in Tk3, nicht in Tk4)
                    If InStr(1, z1, "|EvFol|") > 0 Then Tk3 = Tk3 + z1 + v Else Tk3 = Tk3 + z2 + v
                    i = i + 1 'nächste Zeile überspringen
                    Doubles = Doubles + z1 + v + z2 + v + v
                Else
                    'Z1 ist eine Liga-Rasterzeile
                    Tk3 = Tk3 + z1 + v
                End If
            Else
                Tk4 = Tk4 + z1 + v
            End If
        Next
        ArrC(58) = CStr(anzAinB("|", Doubles) / 18)
        Tk3 = Delete_EmptyEndRowsInString(Tk3)
        Tk4 = Delete_EmptyEndRowsInString(Tk4)
        Get_Tk2Parts = Tk3 + "°°" + Tk4
        'show Doubles 'show Tk2 'show "Tk3 = Liga" + v + v + Tk3 'show "Tk4 = MultiDate von Tk2" + v + v + Tk4
End Function

Sub ShowTk1(Tk1$)
    Dim s$, v$, Arr1() As String
    v = vbCrLf: s = TAB_Simulation(Tk1): Arr1 = Split(s, v)
    QuickSort Arr1: s = Join(Arr1, v)
    show "Tk1" + v + v + "Alle Termine im TkRaster (" + CStr(UBound(Arr1) + 1) + " RasterZeilen);" + v _
        + "Jede RasterZeile enthält eine Version eines Termins;" + v _
        + "1 Termin kann zunächst aus mehreren RasterZeilen bestehen;" + v _
        + "letztlich soll entstehen: 1 Termin - 1 RasterZeile" + v + v + v + s
End Sub

Sub ShowTk2(Tk2$, AnzTk1%, AnzTk2%, AnzTk9%)
    Dim s$, v$, Arr1() As String
    v = vbCrLf: s = TAB_Simulation(Tk2): Arr1 = Split(s, v)
    QuickSort Arr1: s = Join(Arr1, v)
    show "Tk2" + v + v + "Teil der Tk1-Termine, wobei: Datum kommt in Tk1 mehrmals vor (" _
        + CStr(UBound(Arr1) + 1) + " RasterZeilen);" + v _
        + CStr(AnzTk2) + " [Tk2, MultiDate+Liga] + " + CStr(AnzTk9) + " [Tk9, Ready] = " + CStr(AnzTk1) + " [Tk1, All]" + v + v + v + Tk2
End Sub

Sub ShowTk3(Tk3$, AnzTk1%, AnzTk3%, AnzTk4%, AnzTk9%)
    Dim D$, s$, T$, v$, Arr1() As String
    v = vbCrLf: s = TAB_Simulation(Tk3): Arr1 = Split(s, v)
    QuickSort Arr1: Tk3 = Join(Arr1, v)
    D = CStr(AnzTk1 - AnzTk3 - AnzTk4 - AnzTk9)
    T = "Tk3" + v + v + "Teil1 der Tk2-Termine, Liga (" _
        + CStr(UBound(Arr1) + 1) + " RasterZeilen);" + v _
        + CStr(AnzTk4) + " [Tk4, MultiDate] + " _
        + CStr(AnzTk3) + " [Tk3, Liga] + " + D + " [Doubles] + " + CStr(AnzTk9) + " [Tk9, Ready] = " _
        + CStr(AnzTk1) + " [Tk1, All]"
    show T + v + v + v + Tk3
End Sub

Sub ShowTk4(Tk4$, AnzTk1%, AnzTk3%, AnzTk4%, AnzTk9%)
    Dim D$, s$, T$, v$, Arr1() As String
    v = vbCrLf: s = TAB_Simulation(Tk4): Arr1 = Split(s, v)
    QuickSort Arr1: Tk4 = Join(Arr1, v)
    D = CStr(AnzTk1 - AnzTk3 - AnzTk4 - AnzTk9)
    T = "Tk4" + v + v + "Teil2 der Tk2-Termine, MultiDate (" _
        + CStr(UBound(Arr1) + 1) + " RasterZeilen);" + v _
        + CStr(AnzTk4) + " [Tk4, MultiDate] + " _
        + CStr(AnzTk3) + " [Tk3, Liga] + " + D + " [Doubles] + " + CStr(AnzTk9) + " [Tk9, Ready] = " _
        + CStr(AnzTk1) + " [Tk1, All]"
    show T + v + v + v + Tk4
End Sub

Sub ShowTk9(Tk9$)
    Dim s$, v$, Arr1() As String
    v = vbCrLf: s = TAB_Simulation(Tk9): Arr1 = Split(s, v)
    QuickSort Arr1: s = Join(Arr1, v)
    show "Tk9" + v + v + "Teil der Tk1-Termine, bei denen ein Datum nur in 1 RasterZeile vorkommt (" _
        + CStr(UBound(Arr1) + 1) + " RasterZeilen);" + v _
        + "letztlich sollen hier auch die VnNnlichen Tk1-Termine landen" + v + v + v + Tk9
End Sub

Function Get_EventGroups(z$) As String
    'Z          trägt > 3 Events (1 Zeile, Trenner = "###", gleiches Date8)
    'Event      im Format |Date8|Name|Ort|Bem|Ev...|+|EvFolName|path|
    '           - Gleiches Event (gleiche Gruppe), falls Gleichheit besteht in
    '             (Name1InsideName2 OR Ort1InsideOrt2)
    
    'Vorbereitung
        Dim Group$, AllGroups$, MayBeKilled$, v$, z1$, gName$, gOrt$, NAME$, Ort$
        Dim bName As Boolean, bOrt As Boolean
        Dim i%, Arr1() As String, Arr2() As String
        v = vbCrLf

    'Gruppen abtrennen
        Arr1 = Split(z, "###")
        For i = 0 To UBound(Arr1)
            bName = False: bOrt = False
            z1 = Arr1(i)            '1 EventDataZeile von mehreren (getrennt durch "###")
            Arr2 = Split(z1, "|")
            NAME = Arr2(2): Ort = Arr2(3)
            If gName = "" Then gName = NAME
            If gOrt = "" Then gOrt = Ort

            If InStr(1, LCase(NAME), LCase(gName)) > 0 Then bName = True    'Name-OK für SameLine
            If InStr(1, LCase(gName), LCase(NAME)) > 0 Then bName = True    'Name-OK für SameLine
            If NAME = "DM" And gName = "DMM" Then bName = False             'Name-OK für Abtrennung
            If NAME = "DMM" And gName = "DM" Then bName = False             'Name-OK für Abtrennung
            If NAME = "BadM" And gName = "BadMM" Then bName = False         'Name-OK für Abtrennung
            If NAME = "BadMM" And gName = "BadM" Then bName = False         'Name-OK für Abtrennung
            
            If InStr(1, LCase(Left(Ort, 4)), LCase(Left(gOrt, 4))) > 0 Then bOrt = True
            If InStr(1, LCase(Left(gOrt, 4)), LCase(Left(Ort, 4))) > 0 Then bOrt = True
            If InStr(1, Replace(Ort, "/", "_"), Replace(gOrt, "/", "_")) > 0 Then bOrt = True
            If InStr(1, Replace(gOrt, "/", "_"), Replace(Ort, "/", "_")) > 0 Then bOrt = True
            
'If Z1 Like "*|BadM*" Then Stop
'Z mit ...|BadMM...###...|BadM... landet nicht in AllGroups
            
            If (bName And bOrt) Or (bName And Ort = "") Then
                'Z1 gehört zur aktuellen gName-gOrt-Gruppe                      'SameLine
                If Group = "" Then
                    Group = z1
                    
                    
'                    AllGroups = AllGroups + Z1 + v
'                    MayBeKilled = Z1
                    
                Else
'                    If InStr(1, AllGroups, MayBeKilled) > 0 And InStr(1, Group, MayBeKilled) > 0 Then
'                        AllGroups = Replace(AllGroups, MayBeKilled, "")
'                    End If
                    Group = Group + "###" + z1
                    
                    
                End If
            Else
                'Z1 gehört nicht zur aktuellen gName-gOrt-Gruppe    'Abtrennen!
                
                'AllGroups = AllGroups + Group + v
                AllGroups = AllGroups + z1 + v
'show AllGroups
                gName = NAME: gOrt = Ort
                Group = z1
            End If

        Next
        If InStr(1, AllGroups, z1) = 0 Then AllGroups = AllGroups + Group + v
        Get_EventGroups = AllGroups
End Function

Function Change_Date8ToDate10(Date8$) As String
    'Called from    xxx
    'Date8         = "19450312" --> Date10 = "12.03.1945"
    Change_Date8ToDate10 = Mid(Date8, 7, 2) + "." + Mid(Date8, 5, 2) + "." + Mid(Date8, 1, 4)
End Function

Function Change_Date2Point2Point4_ToDate8(Date10$) As String
    'Called from    xxx
    'Date10         = "12.03.1945" --> Date8 = "19450312"
    Change_Date2Point2Point4_ToDate8 = Mid(Date10, 7, 4) + Mid(Date10, 4, 2) + Mid(Date10, 1, 2)
End Function

Function Get_Raster_EvTxt()
    'Action     Terminkalender.txt --> einheitliche RasterZeilen
    '           alle "/" werden in "_" umgewandelt
    
    'Vorbereitung
        Dim s$, txt$, v$, i%, j%, Arr1() As String, Arr2() As String
        Call DoArr: v = vbCrLf
    'txtFile --> Txt
        txt = ReadFile(ArrC(2) + "\z Terminkalender.txt")
    'TrennLinien raus
        txt = Replace(txt, "--", "#"):   txt = Replace(txt, "#-", "#")
        txt = Replace99(txt, "#", ""): txt = Delete_EmptyRowsInString(txt)
        txt = Replace(txt, "/", "_")
    'Bestimmte OrtsNamen erkennbar machen
        If txt Like "*St. *" Then txt = Replace(txt, "St. ", "St.&") 'St. Ingbert, St. Petersburg, ...
    'Action
        Arr1 = Split(txt, v)
        For i = 2 To UBound(Arr1)
            Arr2 = Split(Arr1(i), "|")
            s = s + "|" + Change_Date2Point2Point4_ToDate8(Arr2(2)) + "|" _
                  + Trim(Arr2(3)) + "|" + Trim(Arr2(4)) + "|" + Trim(Arr2(5)) + "|EvTxt|" _
                  + Trim(Arr2(1)) + "|" + Trim(Arr2(6)) + "||" + v
        Next
        Get_Raster_EvTxt = s
        'Hier nicht 'Show_Raster_EvTxt' starten --> EndlosSchleife
End Function

Sub Show_Raster_EvTxt()
    Dim s$, s1$, s2$, s3$, v$
    v = vbCrLf
    s = Get_Raster_EvTxt
    s1 = "Ergebnis von 'Get_Raster_EvTxt'" + v + "(Textversion des Terminkalenders " _
       + "ausgelesen und in ein einheitliches Format - Raster - gebracht)" + v + v + v
    s2 = "| Datum| Titel| Ort| Bemerkung|Quell|+| Name des Event-Ordners| Pfad des Event-Ordners |" + v
    s3 = Replace(TAB_Simulation(s2 + s), "|19000000|", v + v + "|19000000|", 1, 1)
    show s1 + s3
End Sub

Function Get_Raster_EvFol() As String
    'Called from    Update_Terminkalender
    
    'Vorbereitung
        Dim Bem$, Date8$, EvFolName$, EvFolPaFP$, EvFolPath$, EvFolPaths$
        Dim NAME$, NOB$, Ort$, s$, v$
        Dim C1%, C2%, C3%, C4%, C5%, C6%, i%, ArrE() As String
        Call DoArr: v = vbCrLf
        
    EvFolPaths = Get_Paths_ofAllSubfoldersAllLevelsAsStringUseGlobalVar(ArrC(3)) 'ArrC(3) = Path of Folder "Events"
    ArrE = Split(EvFolPaths, v)
    For i = 0 To UBound(ArrE)
        'Schleife über alle EventFolderPfade
        EvFolPath = ArrE(i)         'oneLine    '...\1964-03-21 WM01 London_GB
        EvFolName = Get_NameFromPath(EvFolPath)   '    1964-03-21 WM01 London_GB
        EvFolPaFP = Get_ParentFolderPath_OfFolderPath(EvFolPath) 'ParentFolderPath
        
        If Not EvFolPath Like "*\1900- *\*" Then
        If Not EvFolName = "0000 new" Then
        If Not Mid(EvFolName, 6, 4) = "Liga" Then
           'Date8
            Date8 = Mid(EvFolName, 1, 4)           'SortierSpalte 1
            If Mid(EvFolName, 5, 1) = "-" Then
                If Len(EvFolName) = 5 Then Date8 = Date8 + "00" '"2000-"
                If Len(EvFolName) > 6 Then
                    If IsNumeric(Mid(EvFolName, 6, 2)) Then Date8 = Date8 + Mid(EvFolName, 6, 2) Else Date8 = Date8 + "00"
                End If
            Else
                Date8 = Date8 + "00"
            End If
            If Mid(EvFolName, 8, 1) = "-" Then Date8 = Date8 + Mid(EvFolName, 9, 2) Else Date8 = Date8 + "00"
            'Date8 besteht jetzt sicher aus 8 Zeichen; 19000000 19870100 19870123
            
            
           'NOB = Name, Ort, Bemerkung (ohne Date) (Data eines einzelne EventFolders)
            NOB = Mid(EvFolName, 5) 'Jahreszahl bleibt weg
            Bem = ""
            'Bestimmte OrtsNamen erkennbar machen
                If NOB Like "* St. *" Then NOB = Replace(NOB, " St. ", " St.&") 'St. Ingbert, St. Georgen, ...
                'If NOB Like "* Oos*" Then NOB = Replace(NOB, " Oos", "&Oos")
                'If NOB Like "* Gmünd*" Then NOB = Replace(NOB, " Gmünd", "&Gmünd")
                'If NOB Like "* Kreuznach*" Then NOB = Replace(NOB, " Kreuznach", "&Kreuznach")
            If Mid(NOB, 1, 2) = "- " Then
                NOB = Mid(NOB, 3)
            ElseIf NOB Like "-##-## *" Then NOB = Mid(NOB, 8)
            ElseIf NOB Like "-## *" Then NOB = Mid(NOB, 5)
            ElseIf NOB Like " *" Then NOB = Mid(NOB, 2)
            End If
            
            NAME = NOB
            If Mid(NOB, 1, 1) = "(" Then        '(3) Baden
                If Mid(NOB, 3) = ")" Then Ort = Mid(NOB, 4) Else Ort = Mid(NOB, 5)
            Else
                Ort = "": C4 = InStr(2, NOB, " (")
                C5 = InStr(1, NOB, "LK "): C6 = InStr(1, NOB, "-")
                C1 = InStr(1, NOB, " "): C2 = InStr(C1 + 1, NOB, " ")
                If C1 * C2 > 0 Then C3 = InStr(C2 + 1, NOB, " ") Else C3 = 0
                'c1 bis c6 tragen neue Werte (für dieses NOB)
                
                If C4 > 0 And C1 * C2 > 0 Then  'Treff Karlsdorf (bei Sven Flöß)    'c4 = Pos "("
                    Bem = Mid(NOB, C4 + 1): NOB = Left(NOB, C4 - 1)
                    NAME = Mid(NOB, 1, C1 - 1): Ort = Mid(NOB, C1 + 1)              'c1 = Pos 1." "
                ElseIf C1 * C2 * C3 > 0 Then    'DJM DM DTF Berlin
                    NAME = Mid(NOB, 1, C3 - 1): Ort = Mid(NOB, C3 + 1)
                ElseIf C1 * C2 > 0 Then         'DJM DJMSy Aachen
                    NAME = Mid(NOB, 1, C2 - 1): Ort = Mid(NOB, C2 + 1)
                ElseIf C5 > 0 Then              'JLK D-GB
                    If C2 = 0 Then NAME = NOB: Ort = ""
                ElseIf C1 > 0 And C6 > C1 Then            'Buli1Süd Adorf-Bdorf
                    NAME = Mid(NOB, 1, C1 - 1): Ort = Mid(NOB, C1 + 1, C6 - C1 - 1): Bem = Mid(NOB, C1 + 1)
                ElseIf C1 > 0 Then              'EM15 Dessau
                    NAME = Mid(NOB, 1, C1 - 1): Ort = Mid(NOB, C1 + 1)
                End If
            End If
            s = s + "|" + Date8 + "|" + NAME + "|" + Ort + "|" + Bem _
                  + "|EvFol|+|" + EvFolName + "|" + EvFolPaFP + "|" + v
        End If
        End If
        End If
    Next
    
    Get_Raster_EvFol = s
    'Hier nicht 'Show_Raster_EvFol' starten --> EndlosSchleife
End Function

Sub Show_Raster_EvFol()
    Dim s$, s1$, s2$, s3$, v$
    v = vbCrLf
    s = Get_Raster_EvFol
    s1 = "Ergebnis von 'Get_Raster_EvFol'" + v + "(Namen der EventFolder ausgelesen und " + _
         "in ein einheitliches Format - Raster - gebracht)" + v + v + v
    s2 = "| Datum| Titel| Ort| Bemerkung|Quell|+| Name des Event-Ordners| Pfad des Event-Ordners |" + v
    s3 = Replace(TAB_Simulation(s2 + s), "|19000000|", v + v + "|19000000|", 1, 1)
    show s1 + s3
End Sub

Function Get_Tk1Parts(t1$) As String
    'Vorbereitung
        Dim SoloDates$, MultiDates$, v$, z1$, z2$, z3$, i%, Arr1() As String
        v = vbCrLf
    'Action
        Arr1 = Split(v + t1 + v, v)
        For i = 1 To UBound(Arr1) - 1
            z1 = Left(Arr1(i - 1), 10)
            z2 = Left(Arr1(i), 10)      '|19631001|
            z3 = Left(Arr1(i + 1), 10)
            If z1 <> z2 And z2 <> z3 Then
                SoloDates = SoloDates + Arr1(i) + v
            Else
                MultiDates = MultiDates + Arr1(i) + v
            End If
        Next
        SoloDates = Delete_EmptyEndRowsInString(SoloDates)
        MultiDates = Delete_EmptyEndRowsInString(MultiDates)
        Get_Tk1Parts = SoloDates + "°°" + MultiDates
End Function

Function Get_ImageWidthHeight(imagePath$) As String
    Dim w!, H!, wia As Object
    If FileExists(imagePath) = False Then Exit Function
    'Create the ImageFile object and check if it exists.
        On Error Resume Next
        Set wia = CreateObject("WIA.ImageFile")
        If wia Is Nothing Then Exit Function
        On Error GoTo 0
    'Load the ImageFile object with the specified File.
        'show ImagePath
        wia.LoadFile imagePath
    'Get the necessary properties.
        w = wia.Width
        H = wia.Height
    'Release the ImageFile object.
        Set wia = Nothing
    Get_ImageWidthHeight = Format(w, "000000") + " " + Format(H, "000000")
End Function

Sub EE(c%)
    If c = 0 Then
        Application.EnableEvents = False
    ElseIf c = 1 Then Application.EnableEvents = True
    ElseIf c = 2 Then
        EE2_Holder = Application.EnableEvents: Application.EnableEvents = False
    ElseIf c = 3 Then Application.EnableEvents = EE2_Holder
    Else: Stop
    End If
    'EE2_Holder
End Sub

Sub OpenFolder_SourceImage()
    DoArr
    OpenFolder_SelectFile ArrC(2)
    SelectSourceImageGreen1TopLeft
End Sub

'======================================================================================


