Attribute VB_Name = "M44"
Option Explicit 'M_Calc

Sub zzz_M_Calc()

    showProcs "next"
    
    'show myPeekMessage
    'myPeekMessage = "##"
    'FillArrC 68, "D01D34D43"
    'If Sheets("T4").Cells(12, 31).Interior.Color = vbWhite Then Stop
    
    'DoArrW
    'showArray2D ArrDg
    'show CStr(ArrDg(1, 1))
    EE 1: Beep
End Sub

Function T4_Get_SumOfPfK1K2(z%, Arr_sPKK, mw$)
    'Called from    T4_Calc_M2V8_OneCompetitorSum
    
    'Vorbereitung
        Dim c%, s%, PF!, K1!, K2!, Sum#
        If mw = "Damen" Then c = 0 Else c = 3
    'Pf, K1, K2 für den Competitor in DgZeile z
        s = CInt(Arr_sPKK(c)):      If IsNull(ArrDg(z, s)) Then ArrDg(z, s) = 0
            PF = ArrDg(z, s)
        s = CInt(Arr_sPKK(c + 1)):  If IsNull(ArrDg(z, s)) Then ArrDg(z, s) = 0
            K1 = ArrDg(z, s)
        s = CInt(Arr_sPKK(c + 2)):  If IsNull(ArrDg(z, s)) Then ArrDg(z, s) = 0
            K2 = ArrDg(z, s)
        Sum = PF + K1 + K2: Sum = WorksheetFunction.Round(Sum, 2)
    'Finals
        T4_Get_SumOfPfK1K2 = Sum
End Function

Sub T4_Calc_M2V8_OneCompetitorSum(zMembers$, sPKK$)
    'Called from    T4_AutoCalcM2V8
    'zMembers       = "7|8|9|10|13|14|15|...|34|35|36"
    'sPKK           = "5,8,11,18,21,24" SpaltenNr Pfl K1 K2 (je w/m)
    
    'Vorbereitung
        Dim i%, j%, s%, s1%, z%, z1%, r As Range
        Dim PF!, K1!, K2!, Sum#, Arr_zMembers() As String, Arr_sPKK() As String
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39)): With Sheets("T4")
        zMembers = Replace99(zMembers, "||", "|")
        Arr_zMembers = Split(zMembers, "|")
        Arr_sPKK = Split(sPKK, ",")
    'Action
        For i = 0 To UBound(Arr_zMembers)           'Arr_zMembers:  7|8|9|10|13|14|15|...|34|35|36
            'Schleife über alle CompetitorZeilen z  'Arr_sPKK:  5,8,11,18,21,24
            z = CInt(Arr_zMembers(i))               'z = 7 (Competitor der DgZeile 7), dann 8, ...
            'Damen
                If ArrDg(z, 4) <> "" Then 'Namen
                Sum = T4_Get_SumOfPfK1K2(z, Arr_sPKK, "Damen")
                If Sum > 0 Then
                    'DgSpaltenNr für die CompetitorSumme
                        s = CInt(Arr_sPKK(2)) + 1   'rechts neben K2
                    'Write
                        T4_Write_OneSum_MayBeFixValue Sum, z, s
                End If
                End If
            'Herren
                If ArrDg(z, 17) <> "" Then 'Namen
                Sum = T4_Get_SumOfPfK1K2(z, Arr_sPKK, "Herren")
                If Sum > 0 Then
                    'DgSpaltenNr für die CompetitorSumme
                        s = CInt(Arr_sPKK(5)) + 1   'rechts neben K2
                    'Write
                        T4_Write_OneSum_MayBeFixValue Sum, z, s
                End If
                End If
        Next
    'Finals
        End With
End Sub

Sub T4_Write_OneSum_MayBeFixValue(Sum#, zDg%, sDg%)
    'Called from    T4_Calc_M2V8_Best3_OneCompetitorSum
    'FixValueCheck  In Dg(z|s) könnte ein Wert + "x" (z.B. 87,4x) eingegeben worden sein;
    '               d. h. der Wert soll fixiert seinen Wert behalten (FixValue, FontFarbe rot)
    '               und nicht durch den berechneten Wert überschrieben werden;
    '               in FixValueVermerk Dg(1|24) würde dann "(z|s)" zu finden sein;
    '               wird der Wert in Dg(z|s) gelöscht, verschwindet auch
    '               der Vermerk in Dg(1|4), Dg(z|s) erhält den berechneten Wert (FontFarbe schwarz)
    
    'Vorbereitung
        Dim s1%, z1%, r As Range
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39)): With Sheets("T4")
    'FixValue?
        If ArrDg(zDg, sDg) Like "*x*" Then
            'Dg(zDg,sDg) soll als FixValue gelten
                'Wert (ohne x) in ArrDg schreiben
                    ArrDg(zDg, sDg) = Replace(ArrDg(zDg, sDg), "x", "")
                'Wert (ohne x) in T4 schreiben
                    .Cells(z1 - 1 + zDg, s1 - 1 + sDg) = ArrDg(zDg, sDg)
                    .Cells(z1 - 1 + zDg, s1 - 1 + sDg).Font.Color = vbRed
            'Vermerk in Dg(1|4)
                If Not ArrDg(1, 4) Like "*(" + CStr(zDg) + "|" + CStr(sDg) + ")*" Then
                    '(z|s)-Vermerk steht noch nicht in ArrDg(1, 4)
                    '(z|s)-Vermerk in ArrDg(1, 4) schreiben
                        ArrDg(1, 4) = ArrDg(1, 4) + "(" + CStr(zDg) + "|" + CStr(sDg) + ")"
                    '(z|s)-Vermerk in T4 schreiben
                        Set r = .Cells(z1, s1 + 3)
                        r.Value = r.Value + "(" + CStr(zDg) + "|" + CStr(sDg) + ")"
                        r.Font.size = 6: r.Font.Color = vbWhite
                End If
        End If
    'FixValue löschen
        If ArrDg(zDg, sDg) = "" Then
            'In Dg(zDg,sDg) könnte ein FixValue gestanden haben: Vermerk löschen
            If ArrDg(1, 4) Like "*(" + CStr(zDg) + "|" + CStr(sDg) + ")*" Then
                'FixValue-Vermerk in ArrDg löschen
                    ArrDg(1, 4) = Replace(ArrDg(1, 4), "(" + CStr(zDg) + "|" + CStr(sDg) + ")", "")
                'FixValue-Vermerk in T4 löschen
                    .Cells(z1, s1 + 3) = ArrDg(1, 4)
                'T4-Wert von Rot auf Schwarz setzen
                    .Cells(z1 - 1 + zDg, s1 - 1 + sDg).Font.Color = vbBlack
            End If
        End If
    'Sum
        'Die OneCompetitorSum wird nur dann in Dg(z|s) eingetragen,
        'falls "(z|s)" nicht in Dg(1|4) vermerkt ist
        If Not ArrDg(1, 4) Like "*(" + CStr(zDg) + "|" + CStr(sDg) + ")*" Then
            ArrDg(zDg, sDg) = WorksheetFunction.Round(Sum, 2) 'Format(Sum, "0.0")
            .Cells(z1 - 1 + zDg, s1 - 1 + sDg) = WorksheetFunction.Round(Sum, 2)
        End If
    'Finals
        End With
End Sub

Sub T4_Calc_M2V8_TeamEndSum()
    'Called from    T4_AutoCalcM2V8
    'TeamEndSum     GesamtPunktzahl einer Mannschaft  =  P + K1 + K2
    'Action         Eintrag TeamEndSum in jeder Vereinszeile (= Zeile, die den VereinsNamen enthält)
    '               1. Vereinszeile: Ges(6,12) = P(6,5) + K1(6,8) + K2(6,11)
    
    'Vorbereitung
        Dim sPK1K2$, zClubs$, i%, j%, s%, s1%, z%, z1%
        Dim PF!, K1!, K2!, Sum#, Arr_zClubs() As String, Arr_sPK1K2() As String
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39)): With Sheets("T4")
    'Dg-ZeilenNrn der ClubZeilen '"|06|13|20|27|34|41|48|55|62|69"
        zClubs = ArrC(45)
        Arr_zClubs = Split(zClubs, "|")
    'Dg-SpaltenNrn für P K1 K2 P K1 K2
        sPK1K2 = "5,8,11,18,21,24"
        Arr_sPK1K2 = Split(sPK1K2, ",")
    'Action
        For i = 1 To UBound(Arr_zClubs)                     'Arr_zClubs:  06|13|20|27|34|41|48|55|62|69  zClubs
            'Schleife über alle ClubZeilen z     'Arr_sPK1K2:  5,8,11,18,21,24                sPK1K2
            z = CInt(Arr_zClubs(i)) '6 (Zeile 6 ist die erste ClubZeile, 13 die zweite)
            PF = ArrDg(z, CInt(Arr_sPK1K2(0)))
            K1 = ArrDg(z, CInt(Arr_sPK1K2(1)))
            K2 = ArrDg(z, CInt(Arr_sPK1K2(2)))
            Sum = PF + K1 + K2: Sum = WorksheetFunction.Round(Sum, 2)
            If Sum > 0 Then
                s = CInt(Arr_sPK1K2(2)) + 1     'Spalte für die TeamEndSum; rechts neben K2 (in der ClubZeile z)
                T4_Write_OneSum_MayBeFixValue Sum, z, s
'                ArrDg(z, s) = Sum
'                .Cells(z1 - 1 + z, s1 - 1 + s) = Sum
            End If
            
            'xxx fixValue
            'T4_Write_OneSum_MayBeFixValue Sum, z, s

            
            PF = ArrDg(z, CInt(Arr_sPK1K2(3)))
            K1 = ArrDg(z, CInt(Arr_sPK1K2(4)))
            K2 = ArrDg(z, CInt(Arr_sPK1K2(5)))
            Sum = PF + K1 + K2: Sum = WorksheetFunction.Round(Sum, 2)
            If Sum > 0 Then
                s = CInt(Arr_sPK1K2(5)) + 1
                ArrDg(z, s) = Sum
                .Cells(z1 - 1 + z, s1 - 1 + s) = Sum
            End If
        Next
        End With
End Sub

Sub Put_TeamSumNow_G(z%, s%, z1%, s1%, PF#, K1#, K2#)
    'Vorbereitung
        With Sheets("T4")
    'Action
        If ArrDg(z, s) Like "*x*" Then
            .Cells(z1 - 1 + z, s1 - 1 + s).Font.Color = vbRed
        Else
            ArrDg(z, s) = PF + K1 + K2
            If ArrDg(z, s) > 0 Then
                'select
                    'ee 0: .Cells(z1 - 1 + z, s1 - 1 + s).Select: ee 1: Stop
                .Cells(z1 - 1 + z, s1 - 1 + s) = ArrDg(z, s)
            End If
            .Cells(z1 - 1 + z, s1 - 1 + s).Font.Color = vbBlack
        End If
    'Finals
        End With
End Sub

Function Get_DgValueNoX(z%, SpNr$) As Double
    Dim A$, s%
    s = CInt(SpNr): A = Trim(Replace(ArrDg(z, s), "x", ""))
    If A = "" Then Get_DgValueNoX = 0 Else Get_DgValueNoX = CDbl(A)
End Function

Function T4_Get_zDgNrsOfMembersOfOneClub() As String
    'Called from    T4_AutoCalcM2V8
    
    'Vorbereitung
        Dim zClubs$, s$
        Dim i%, s1%, sDg%, sSh%, z1%, z2%, zClub1%, zClub2%, zDg%, zSh%
        Dim Arr_zClubs() As String
        z1 = CInt(ArrC(37)): z2 = CInt(ArrC(38)): s1 = CInt(ArrC(39))
    'Action
        zSh = CInt(ArrC(104))          'T4-SheetZeilenNr  der ActualSelectedCell (zSh, sSh)
        sSh = CInt(ArrC(105))          'T4-SheetSpaltenNr der ActualSelectedCell (zSh, sSh)
        zDg = zSh - z1 + 1
        sDg = sSh - s1 + 1
        If zDg > 6 Then
            zClubs = ArrC(45) + "|" + CStr(z2 - z1 + 2)     'T4-Dg-ZeilenNrn VereinsNamen '|6|13|20|...|181
            Arr_zClubs = Split(zClubs, "|")
            For i = 1 To UBound(Arr_zClubs)
                If Arr_zClubs(i) < zDg Then zClub1 = Arr_zClubs(i) Else zClub2 = Arr_zClubs(i): Exit For
            Next
            'show CStr(zDg) + " zwischen " + CStr(zClub1) + " und " + CStr(zClub2): Stop
            For i = zClub1 + 1 To zClub2 - 2: s = s + "|" + CStr(i): Next
            'show s: Stop '|182|183|184|185|186 'Dg-ZeilenNrn
        End If
        T4_Get_zDgNrsOfMembersOfOneClub = s
End Function

Sub T4_AutoCalcM2V8()
    'Called from    T4_AutoCalc
    'Best4          Bei MannschaftsWettkämpfen werden Mannschaftsergebnisse
    '               für 4 Durchgänge berechnet (Pflicht, Kür1, Kür2, Gesamt).
    '               Die Summe der 4 besten EinzelWertungen ergeben jeweils
    '               das TeamErgebnis; dies für P, K1, K2;
    '               daraus wird dann das TeamEndErgebnis P+K1+K2 berechnet
    'Best3          Ggf. sollen die 3 besten (nicht 4) EinzelWertungen summiert werden;
    '               dies gilt, falls "Best3" in DgZelle (1|3) vermerkt ist
    'm/w            ZeilenSpaltenStruktur(w) = ZeilenSpaltenStruktur(m)
    'CALC1          AAE, AddAddErgebnis (Haltung+Schwierigkeit=Kür1Ergebnis w, ...)
    'CALC2          SumP, SumK1, SumK2              (je Team, je m/w)
    'CALC3          TeamEndSum   (SumP+SumK1+SumK2) (je Team, je m/w)
    'CALC4          SingleEndSum (P+K1+K2)          (je Competitor)
    
    'Exit
        If ArrDg(2, 2) Like "1980-09-13 D*MM Stadtallendorf*" Then Exit Sub
    'Vorbereitung
        Dim sAllPKK$, zMembers$, zDgNrsOfMembersOfOneClub$, zMembersOneTeam$
        Dim i%, j%, s1%, sOnePKK%, z1%, z2%
        Dim Arr1() As String, Arr2() As String, Best3 As Boolean
        z1 = CInt(ArrC(37)): z2 = CInt(ArrC(38)): s1 = CInt(ArrC(39))
        
    'ArrDg 'wird bei jedem Klick in ein Dg neu geladen
        'ArrDg wurde gerade geladen (mit 'FillArrDg' in T4_ActionsOnDgClickNext)
    'Best3
        If ArrDg(1, 3) = "Best3" Then Best3 = True
        'showArray2D ArrDg
    'CALC1: AAE
        T4_Put_CalcedValue_TypAAE
    'ZeilenNrn der MannschaftsMitglieder
        zMembers = T4_Get_AllTeamMemberRowNrs
        '        = "7|8|9|10|11@15|...|29@32|33|34|35|36|37"
    'SpaltenNrn für Sum-Eintrag (P, K1, K2; immer je m/w bei M2V8)
        sAllPKK = "5,8,11,18,21,24" 'Spalten für Pfl K1 K2 Pfl K1 K2
    'CALC2: SumP_SumK1_SumK2_AllTeams
        Arr1 = Split(sAllPKK, ",")
        Arr2 = Split(zMembers, "@")       'Anzahl der Mannschaften = UBound(Arr2) + 1
        'showArray Arr2: Stop
        
        For i = 0 To UBound(Arr1)
            'Eintrag der MannschaftsSumme für Pflicht(i=0), Kür1(i=1), Kür2 (in T4Dg und ArrDg)
            For j = 0 To UBound(Arr2)
                'Schleife über alle Mannschaften
                'SpaltenNr für Durchgang i      (i=0 Pflicht w, ..., i=6 Kür2 m)
                    sOnePKK = Arr1(i)
                    'Für i = 0 wird Arr1(i) = 5 = Dg-SpaltenNr für Pflicht weiblich
                    'Für i = 1 wird Arr1(i) = 8 = Dg-SpaltenNr für Kür1    weiblich
                'ZeilenNrn des Team j
                    zMembersOneTeam = Arr2(j)
                    'Für j = 0 wird Arr2(j) = "7|8|9|10|11" = Dg-RowNr of each team1 member
                    'Für j = 1 wird Arr2(j) = "15|16|17|18" = Dg-RowNr of each team2 member
                If Best3 Then
    'Best3
                    T4_CalcAndWrite_SumP_SumK1_SumK2_Best3_OneTeam zMembersOneTeam, sOnePKK 'pro Team: GesamtWert für Pflicht, Kür1, Kür2
                Else
    'Best4
                    T4_CalcAndWrite_SumP_SumK1_SumK2_Best4_OneTeam zMembersOneTeam, sOnePKK 'für die Summe benötigte Dg-Zellen (z|s)
                End If
            Next
        Next
    'CALC3  TeamEndSum
        T4_Calc_M2V8_TeamEndSum
    'CALC4   SingleEndSum
        'Berechnung des Gesamtergebnisses für jeden einzelnen Wettkämpfer
        zMembers = Replace(zMembers, "@", "|")
        T4_Calc_M2V8_OneCompetitorSum zMembers, sAllPKK
    'Rang Einzel
        'T4_Put_DgLigaRangEinzel s1, z1, z2
End Sub

Sub T4_Load_StringTabelle_ToArray2D_SortDown(B, s$)
    'Called from    T4_CalcAndWrite_SumP_SumK1_SumK2_Best3_OneTeam                          'sorted s
    's              = StringTabelle                         'StringTabelle s        'B(), 2D-Array
    '                                                       '|24,20|7|5|278|369|    '|26,90|9|5|280|369|
    'B              = 2D-Array, zunächst leer;              '|23,50|8|5|279|369|    '|24,20|7|5|278|369|
    '                 wird hier gefüllt                     '|26,90|9|5|280|369|    '|23,50|8|5|279|369|
    '               = Daten von s,                          '||10|5|281|369|        '||||||
    '                 nach Spalte 1 absteigend sortiert     '||||||                 '||||||
    
    'Vorbereitung
        Dim AnzSp%, AnzZe%, C1%, C2%, i%, j%, A1() As String, A2() As String
        A1 = Split(s, vbCrLf)
    'B()
        AnzZe = UBound(A1) + 1
        AnzSp = anzAinB("|", A1(1)) - 1
        ReDim B(1 To AnzZe, 1 To AnzSp)
    'LeerZeilen
        For i = 0 To UBound(A1)
            If Left(A1(i), 2) = "||" Then A1(i) = "|00,00|" + String(AnzSp - 1, "|")
        Next
    'sort
        QuickSortDown A1
        'showArray A1: Stop
    'A1 1Dim --> A() 2Dim
        For i = 0 To UBound(A1)
            A2 = Split(A1(i), "|")  '|26,90|9|5|280|369| dann |24,20|7|5|278|369|
            For j = 1 To AnzSp
                If A2(1) = "00,00" Then A2(1) = ""
                B(i + 1, j) = A2(j)
            Next
        Next
        'showArray2D B: Stop
End Sub

Sub T4_CalcAndWrite_SumP_SumK1_SumK2_Best3_OneTeam(zTeam$, sSum%)
    'Called from    T4_AutoCalc
    'zTeam          = "7|8|9|10"    = ZeilenNrn, in denen die Werte stehen
    'sSum           = 5             = SpaltenNr, in der   die Werte stehen (5 Pflicht, 8 K1, ...)
    'Action         erreichte Summe für 1 Mannschaft/1 Durchgang (z.B. K1) wird eingetragen
    '               Erstellen eines 2D-Array mit zu addierenden Werten/Zellpositionen
    '               Best3 addieren, Streichwertungen grau
    
    'Exit
        Dim i%, Az() As String, Ex As Boolean
        Az = Split("|" + zTeam, "|"): Ex = True
        For i = 1 To UBound(Az)
            If ArrDg(CInt(Az(i)), sSum) <> "" Then Ex = False: Exit For
        Next
        If Ex Then Exit Sub
    'Vorbereitung
        Dim M1$, M2$, M3$, s$, v$, c%, j%, MaxLines%, MaxMembers%, s1%
        Dim sDgSum%, sShSum%, z1%, zDgSum%, zShSum%, Sum#, W1!, W2!, W3!
        Dim A(), A1() As String, A2() As String, B()
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39)): With Sheets("T4")
        v = vbCrLf
        MaxLines = UBound(Az) '= Anzahl Zeilen mit m/w-Farbe; ggf. Zeilen ohne Werte
    'Einstellung
        MaxMembers = 5
    'A()
        ReDim A(1 To MaxMembers, 1 To 5) '(1 To 5, Wert/zDg/sDg/zSh/sSh)
        For i = 1 To MaxLines
            A(i, 1) = CStr(Format(ArrDg(Az(i), sSum), "00.00"))    'Wert
            A(i, 2) = Az(i)                                 'zDg    Dg-ZeilenNr
            A(i, 3) = sSum                                  'sDg    Dg-SpaltenNr
            A(i, 4) = z1 - 1 + Az(i)                        'zSh    Sheet-ZeilenNr
            A(i, 5) = s1 - 1 + sSum                         'sSh    Sheet-SpaltenNr
        Next
        'Alle relevanten Werte befinden sich in A()
        'showArray2D a: Stop
    'Sum - DgPosition, CellPosition
        zDgSum = CInt(Az(1) - 1): sDgSum = sSum: zShSum = CInt(z1 - 2 + Az(1)): sShSum = s1 - 1 + sSum
    's                                                      'StringTabelle s        'B(), 2D-Array
        s = T4_Get_StringTabelle_FromArray2D(A)             '|24,20|7|5|278|369|    '|26,90|9|5|280|369|
        'show s: Stop                                       '|23,50|8|5|279|369|    '|24,20|7|5|278|369|
    'B()                                                    '|26,90|9|5|280|369|    '|23,50|8|5|279|369|
        If s = "" Then Exit Sub                             '||10|5|281|369|        '|||||
        T4_Load_StringTabelle_ToArray2D_SortDown B, s       '|||||                  '|||||
    'Sum Best3
        M1 = B(1, 1): If M1 = "" Then W1 = 0 Else W1 = CSng(M1)
        M2 = B(2, 1): If M2 = "" Then W2 = 0 Else W2 = CSng(M2)
        M3 = B(3, 1): If M3 = "" Then W3 = 0 Else W3 = CSng(M3)
        Sum = W1 + W2 + W3: Sum = WorksheetFunction.Round(Sum, 2)
    'Select
        'ee 0: .Cells(zShSum, sShSum).Select: ee 1: Stop
    'Write
        If Sum > 0 Then ArrDg(zDgSum, sDgSum) = Sum
        If Sum > 0 Then .Cells(zShSum, sShSum) = Sum
        .Cells(zShSum, sShSum).Font.Color = vbBlack
        If Sum > 0 Then ArrDg(zDgSum, sDgSum) = Sum
        If Sum > 0 Then .Cells(zShSum, sShSum) = Sum    'FontFarbe für Dg-Werte auf Sheet T4
        'showArray2D B: Stop
        If CStr(B(1, 1)) <> "" Then .Cells(CInt(B(1, 4)), CInt(B(1, 5))).Font.Color = vbBlack
        If CStr(B(2, 1)) <> "" Then .Cells(CInt(B(2, 4)), CInt(B(2, 5))).Font.Color = vbBlack
        If CStr(B(3, 1)) <> "" Then .Cells(CInt(B(3, 4)), CInt(B(3, 5))).Font.Color = vbBlack
        'Den 4.-höchsten Wert im Dg grau färben
            If CStr(B(4, 1)) <> "" Then
                'Select
                    'EE 0: .Cells(CInt(B(4, 4)), CInt(B(4, 5))).Select: EE 1: Stop
                    'showArray2D B: Stop
                    .Cells(CInt(B(4, 4)), CInt(B(4, 5))).Font.Color = RGB(99, 99, 99) 'Farbe grau
            End If
        'Den 5.-höchsten Wert im Dg grau färben
            If CStr(B(5, 1)) <> "" Then .Cells(CInt(B(5, 4)), CInt(B(5, 5))).Font.Color = RGB(99, 99, 99)  'Farbe grau
    'Finals
        End With
End Sub

Sub T4_CalcAndWrite_SumP_SumK1_SumK2_Best4_OneTeam(zTeam$, sSum%)
    'Called from    T4_AutoCalc
    'Status         Funktion wird bei jedem Dg-Klick aufgerufen
    'zTeam          = "7|8|9|10|11|12"  (DgRowNr of each member of 1 team)
    'sSum           = 5                 (DgColumnNr of values to add)
    'Action         writes Sum to T4    (Sum of best 4 values in [rows zTeam|column sSum])
    '               color black to best 4, color grey to VnNn
    'ArrDg          wird bei jedem Klick in ein Dg neu geladen
    '               wurde gerade geladen (mit 'FillArrDg' in T4_ActionsOnDgClickNext)
    
    'Vorbereitung
        Dim D$, W2$
        Dim c%, i%, j%, s%, s1%, z%, z1%, zDg%, zV%
        Dim G#, G1#, Sum#, w# '#=Double
        Dim Arr_zTeam() As String, ArrD() As String
        If zTeam = "" Then Exit Sub
        If Right(zTeam, 1) = "|" Then zTeam = Left(zTeam, Len(zTeam) - 1)
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39)): With Sheets("T4")
    'Arr_zTeam
        Arr_zTeam = Split(zTeam, "|")
    'DgZeilenNr der VereinsZeile
        zV = Arr_zTeam(0) - 1
    'D-Liste erstellen  'D = Durchgang (P, K1 oder K2) = Liste aller Werte + ihrer ZellPosition
        D = ""          'D = "23,95°07°05|...|28,00°12°05" (Zeilen 07-12, also 6 Aktive)
        '                     23,95°07°05 --> 23,95 P. des Aktiven Zeile 7, Spalte 5 (Pflicht)
        For i = 0 To UBound(Arr_zTeam)
            zDg = Arr_zTeam(i)    'Arr_zTeam(0) = "7" wegen zTeam = "7|8|9|10|11|12|"
            If ArrDg(zDg, sSum) = "" Then ArrDg(zDg, sSum) = 0
            D = D + Format(ArrDg(zDg, sSum), "00.00") _
                  + "°" + Format(zDg, "00") + "°" + Format(sSum, "00") + "|"
        Next
        D = Left(D, Len(D) - 1)
        'D trägt jetzt alle z. B. Kür1-Punkte einer Mannschaft
    'ArrD erstellen     (damit man sortieren kann)
        ArrD = Split(D, "|"): QuickSort ArrD
        'ArrD trägt jetzt eine sortierte D-Liste: 23,95°07°05|...|28,00°12°05
        c = 0: Sum = 0  'showArray ArrD
    'Sum        (max 4 höchste Werte addieren = Ergebnis; VnNn = Streichwertungen)
        For i = UBound(ArrD) To 0 Step -1
            'Schleife über die sortierte D-Liste (vom höchsten zum tiefsten Wert)
            c = c + 1
            w = CDbl(Left(ArrD(i), 5))      '28     [= CDbl("28,00")]
                'ArrD(i) mit NegativWert?   'ArrD(i) = "-31,90°10°11"
                'K1 wurde eingetragen; Einzelergebnis P muss neu berechnet werden
            z = CInt(Mid(ArrD(i), 7, 2))    '12     [von "28,00°12°05"]
            s = CInt(Mid(ArrD(i), 10, 2))   '5      [von "28,00°12°05"]
            'Der Wert W des Aktiven i steht in der Dg-Zelle (z, s)
            'FontFarbe für Dg-Werte auf Sheet T4
                If c < 5 Then               'W ist also einer der Best4-Werte
                    Sum = Sum + w
                    'Einfärben der 4 höchsten Wertungen: Schwarz
                    If w > 0 Then .Cells(z1 - 1 + z, s1 - 1 + s).Font.Color = vbBlack
                Else
                    'Einfärben der Streichnoten: Grau
                    If w > 0 Then .Cells(z1 - 1 + z, s1 - 1 + s).Font.Color = RGB(99, 99, 99) 'Farbe grau
                    'Einfärben der "0,0" in die m/w-Farbe (soll nicht sichtbar sein)
                    If w = 0 Then .Cells(z1 - 1 + z, s1 - 1 + s).Font.Color = .Cells(z1 - 1 + z, s1 - 1 + s).Interior.Color
                End If
        Next
        Sum = WorksheetFunction.Round(Sum, 2)
        'Sum ist jetzt berechnet (Mannschaftsergebnis, das in der Vereinszeile eingetragen wird)
        If Sum > 0 Then
                'Sum könnte in ArrDg bereits eingetragen sein und den Zusatz "x" tragen;
                'd. h. der Wert ist offiziell und soll nicht überschrieben werden,
                'ist aber nicht die Summe der zu addierenden Werte, wird deshalb markiert
            If ArrDg(zV, sSum) Like "*x*" Then
                .Cells(z1 + zV - 1, s1 - 1 + sSum).Font.Color = vbRed
            Else
                'Eintrag in das Dg auf Sheet T4
                    .Cells(z1 + zV - 1, s1 - 1 + sSum) = Sum
                    .Cells(z1 + zV - 1, s1 - 1 + sSum).Font.Color = vbBlack
                'Eintrag in ArrDg
                    ArrDg(zV, sSum) = Sum
            End If
        End If
    'Finals
        End With
        'showArray2D ArrDg
End Sub

Function T4_Get_StringTabelle_FromArray2D(A) As String
    'Called from    T4_CalcAndWrite_SumP_SumK1_SumK2_Best3_OneTeam
    'A              = 2D-Array, z. B.:      |24,20|7  |5  |278|369|
    '                                       |26,90|8  |5  |279|369|
    '                                       |23,50|9  |5  |280|369|
    '                                       |     |10 |5  |281|369|
    '                                       |     |   |   |   |   |
    's              = String, reduced
    '                           z. B.:      |24,20|7  |5  |278|369|
    '                                       |26,90|8  |5  |279|369|
    '                                       |23,50|9  |5  |280|369|
    
    'Vorbereitung
        Dim s$, i%, j%
    's
        For i = 1 To UBound(A, 1)
            For j = 1 To UBound(A, 2)
                s = s + "|" + CStr(A(i, j)) + "|"
            Next
            s = s + vbCrLf
        Next
        s = Replace(s, "||", "|"): s = Delete_EmptyRowsInString(s)
    'Finals
        T4_Get_StringTabelle_FromArray2D = s
End Function

Function Get_AllTeamSumColumnNrs() As String
    Dim s$, i%
    For i = 1 To UBound(ArrDg, 2)
        If ArrDg(5, i) = "E" Then s = s + "," + CStr(i)
    Next
    If Left(s, 1) = "," Then s = Mid(s, 2)
    Get_AllTeamSumColumnNrs = s '= "5,8,11,18,21,24"
End Function

Sub T4_AutoCalcLiga()
    'Called from    T4_AutoCalc
    'Wertung        Bei LigaWettkämpfen werden Mannschaftsergebnisse für
    '               4 Durchgänge berechnet (Pflicht, Kür1, Kür2, Gesamt).
    '               Die Summe der 4 besten EinzelWertungen ergeben jeweils
    '               die Mannschaftsergebnisse P, K1, K2, dann wird daraus
    '               das Gesamtergebnis P+K1+K2 berechnt;
    '               max. Anzahl Mannschaftsmitglieder bei Ligawettkämpfen = 7
        
    'Exit
        If ArrDg(2, 1) = "M0V1" Then Exit Sub
    'Vorbereitung
        Dim AllTeamNameRowNrs$, m$, i%, s1%, z1%, z2%, zDg%, Arr1() As String
        With Sheets("T4")
        z1 = CInt(ArrC(37)): z2 = CInt(ArrC(38)): s1 = CInt(ArrC(39))
    'ArrDg 'wird bei jedem Klick in ein Dg neu geladen
        'ArrDg wurde gerade geladen (mit 'FillArrDg' in T4_ActionsOnDgClickNext)
    'zClubs, ZeilenNrn der VereinsZeilen
        AllTeamNameRowNrs = T4_Get_AllTeamNameRowNrs_Liga     '"06|13|20|27|34|41|48|55|62|69"
        FillArrC 45, AllTeamNameRowNrs
    'ZeilenNrn der Mannschaften ermitteln
        m = T4_Get_AllTeamMemberRowNrs
        ' = "7|8|9|10|11|12|@15|...|29|@32|33|34|35|36|37|"
    '48 ArrayZellen pro Mannschaft: 8 x Zeile/Spalte/Wert x 3 Durchgänge (P,K1,K2)
        Arr1 = Split(m, "@")                'Anzahl der Mannschaften = UBound(Arr1)
        For i = 0 To UBound(Arr1)           'Arr1(i) = "7|8|9|10|11|12|" (Dg-ZeilenNummern)
            'Schleife über alle Mannschaften
            T4_Put_DgLigaSum Arr1(i)   'Dg-ZeilenNummern 1 Mannschaft
            '             Arr1(i) = "7|8|9|10|11|12|" (Dg-ZeilenNummern)
        Next
    'Rang Einzel
        T4_Put_DgLigaRangEinzel s1, z1, z2
    'Finals
        End With
End Sub

Function T4_Get_AllTeamNameRowNrs_Liga()
    'Called from    T4_AutoCalcLiga
    
    'Vorbereitung
        Dim VNames$, zV$, Sp%, Ze%, Arr1() As String
    'ArrDg 'wird bei jedem Klick in ein Dg neu geladen
        'ArrDg wurde gerade geladen (mit 'FillArrDg' in T4_ActionsOnDgClickNext)
        'showArray2D ArrDg: Stop
    'VereinZeilen zV ermitteln (ZeilenNrn in denen VereinsNamen stehen)
        For Sp = 2 To UBound(ArrDg, 2)
            'Alle Spalten durchsuchen (in Zeile 5) um die Vereinsspalte 'V' zu finden
            If ArrDg(5, Sp) = "D" Then
                'In SpaltenNr Sp steht ein Vereinsname
                '- und zwar in jeder Zeile >5 (aufgefüllt in FillArrDg)
                'In "K"-Spalten  steht eine Rangzahl
                '- und zwar in jeder Zeile >5 (aufgefüllt in FillArrDg)
                For Ze = 6 To UBound(ArrDg, 1) - 1
                    If InStr(1, VNames, ArrDg(Ze, Sp)) = 0 Then
                        VNames = VNames + ArrDg(Ze, Sp)
                        If ArrDg(Ze, Sp) <> "" Then
                            'In Zeile Ze steht ein Vereinsname
                            If InStr(1, zV, "|" + Format(Ze, "00") + "|") = 0 Then
                                zV = zV + "|" + Format(Ze, "00") + "|"
                            End If
                        End If
                    End If
                Next
            End If
        Next
        zV = Replace(zV, "||", "|"): zV = Mid(zV, 2, Len(zV) - 2)
        Arr1 = Split(zV, "|"): QuickSort Arr1: zV = Join(Arr1, "|")
        'zV ist jetzt ermittelt: zV = "06|13|20|27|34|41|48|55|62|69"
        T4_Get_AllTeamNameRowNrs_Liga = zV
End Function

Function T4_Get_AllTeamMemberRowNrs() As String
    'Called from    T4_AutoCalcM2V8, T4_AutoCalcLiga
    'Action         liefert alle TeamMemberZeilen einer Mannschaft (z|z|z|z|...);
    '               dito für alle Mannschaften, mit "@" als Trenner
    '               M = "7|8|9|10|11|12|@15|...|29|@32|33|34|35|36|37|"
    'Only if        es existiert eine LeerZeile zwischen den Mannschaften
    '               zClubs = ZeilenNrVereinsName; zClubs(w) = zClubs(m)
    '               TeamMemberZeilen reichen von zV+1 bis zNextV-2
    '               eine TeamMemberZeile/VereinsZeile kann leer sein

    'Vorbereitung
        Dim m$, VNames$, zClubs$, i%, j%, Sp%, Ze%, zDg%, zV1%, zV2%, Arr1() As String
    'ArrDg 'wird bei jedem Klick in ein Dg neu geladen
        'ArrDg wurde gerade geladen (mit 'FillArrDg' in T4_ActionsOnDgClickNext)
    'VereinZeilen zClubs (ZeilenNrn in denen VereinsNamen stehen)
        zClubs = ArrC(45): If zClubs = "-" Then Stop 'zClubs = "|06|13|20|27|34|41|48|55|62|69"
    'TeamMemberRowNrs
        'show zClubs: Stop
        'showArray2D ArrDg: Stop
        zClubs = zClubs + "|" + CStr(UBound(ArrDg, 1) + 1)    '|06|13|20|27|34|41|48|55|62|69|76
        If Left(zClubs, 1) = "|" Then zClubs = Mid(zClubs, 2) '06|13|20|27|34|41|48|55|62|69|76
        Arr1 = Split(zClubs, "|") '06|14|22|31|39
        
        For i = 0 To UBound(Arr1) - 1
            zV1 = CInt(Arr1(i))                     'zV1 =  6 = ZeilenNr eines Vereins
            zV2 = CInt(Arr1(i + 1))                 'zV2 = 13 = ZeilenNr NextVerein
            For j = 1 To zV2 - zV1 - 2
                m = m + CStr(zV1 + j) + "|"
            Next
            m = m + "@"
        Next
    'Finale
        m = Replace(m, "|@", "@"): m = Left(m, Len(m) - 1)
        T4_Get_AllTeamMemberRowNrs = m '7|8|9|10|11@14|15|16|17|18@...@70|71|72|73|74
End Function

Sub T4_Put_CalcedValue_TypAAE()
    'Called from    T4_AutoCalcM2V8
    'AAE            DgRow5-Buchstabenkette enthält ggf. AAE (nicht AAAE)
    '               WertInZelle
    'ArrDg          'wird bei jedem Klick in ein Dg neu geladen
    '               ArrDg wurde gerade geladen (mit 'FillArrDg' in T4_ActionsOnDgClickNext)
    'B              Buchstabenkette in DgRow5 (1 Buchstabe pro Spalte)
    '               = "iRMT-AAE-AAAE-AAAE-G-RWT-AAE-AAAE-AAAE-Gi"
    'AAE            steht im Beispiel in den Spalten 6,7,8 und 26,27,28
    '               A+A=E --> WertInSpalte6 + WertInSpalte7 = WertInSpalte8 (in einer Zeile >5)
    'Action         Falls einer der Werte A, A, E fehlt, wird dieser eingetragen
    
    'Vorbereitung
        Dim AAE$, i%, s%, s1%, z%, z1%, z2%, Arr1() As String
        z1 = CInt(ArrC(37)): z2 = CInt(ArrC(38)): s1 = CInt(ArrC(39)): With Sheets("T4")
        AAE = Get_StartColumsOfAAE() 'AAE = "" oder "6|26|" = je Position des ersten "s" eines AAE
        'show AAE '"6|9|19|22"
    'Action
        If AAE <> "" Then
            'Die Buchstabenfolge "AAE" existiert in B
            Arr1 = Split(AAE, "|")          '"6|9|19|22"
            For i = 0 To UBound(Arr1)
                s = CInt(Arr1(i))     '= 6 = DgSpaltenNr des ersten "A" eines AAE
                'Check pro Dg-Zeile, ob 2 der 3 Werte vorhanden sind
                For z = 6 To z2 - z1
                    'Schleife über alle relevanten Dg-Zeilen
                    If ArrDg(z, s) = "" And ArrDg(z, s + 1) <> "" And ArrDg(z, s + 2) <> "" Then
                        ArrDg(z, s) = ArrDg(z, s + 2) - ArrDg(z, s + 1)
                        .Cells(z1 - 1 + z, s1 - 1 + s) = ArrDg(z, s)
                    End If
                    If ArrDg(z, s + 1) = "" And ArrDg(z, s) <> "" And ArrDg(z, s + 2) <> "" Then
                       ArrDg(z, s + 1) = ArrDg(z, s + 2) - ArrDg(z, s)
                       .Cells(z1 - 1 + z, s1 - 1 + s + 1) = ArrDg(z, s + 1)
                    End If
                    If ArrDg(z, s + 2) = "" And ArrDg(z, s) <> "" And ArrDg(z, s + 1) <> "" Then
                       ArrDg(z, s + 2) = ArrDg(z, s) + ArrDg(z, s + 1)
                       .Cells(z1 - 1 + z, s1 - 1 + s + 2) = ArrDg(z, s + 2)
                    End If
                Next
            Next
        End If
    'Finals
        End With
End Sub

Function Get_StartColumsOfAAE() As String
    Dim B$, s$, c%, i%
    For i = 1 To UBound(ArrDg, 2)
        B = B + ArrDg(5, i)         'B = "iRVWEAAEAAEGR-RVMEAAEAAEGRi"
    Next
    B = Replace(B, "AAAE", "XXXX")
    Do While InStr(1, B, "AAE") > 0
        c = InStr(1, B, "AAE")
        s = s + "|" + CStr(c)
        B = Replace(B, "AAE", "XXX", 1, 1)
    Loop
    If Len(s) > 1 Then s = Mid(s, 2)
    Get_StartColumsOfAAE = s        's = "6|9|19|22"
End Function

Sub T4_Put_DgLigaSum(zDgs$)
    'Called from    T4_AutoCalc
    'Status         Sub ist nur bei Liga-Ergebnislisten aktiv
    '               Sub wird bei jedem Dg-Klick aufgerufen
    'zDgs           = "7|8|9|10|11|12|" (Dg-ZeilenNummern, 1 Mannschaft)
    'ArrDg          wird bei jedem Klick in ein Dg neu geladen
    '               wurde gerade geladen (mit 'FillArrDg' in T4_ActionsOnDgClickNext)
    
    'Exit
    
    
    Exit Sub
    
    
        If ArrDg(2, 1) = "M0V1" Then Exit Sub
        If ArrDg(2, 1) = "M0V3" Then Exit Sub
    'Vorbereitung
        Dim D$, v$, W2$
        Dim c%, i%, j%, s%, s1%, sDg%, sShSum%, z%, z1%, zDg%, zDgVer%, zShVer%
        Dim G#, G1#, Sum#, w# '#=Double
        Dim ArrzDgs() As String, ArrD() As String, Arr3(1 To 3)
        If zDgs = "" Then Exit Sub
        If Right(zDgs, 1) = "|" Then zDgs = Left(zDgs, Len(zDgs) - 1)
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39)): v = vbCrLf: With Sheets("T4")
    'ArrzDgs
        ArrzDgs = Split(zDgs, "|")
    'Arr3       3 DgSumSpalten (5,8,11) = E-Spalten, deren Werte aufsummiert werden sollen;
        '       in einem Liga-Dg sind es immer genau die Spalten 5, 8 und 11
        Arr3(1) = 5: Arr3(2) = 8: Arr3(3) = 11 'Arr3 trägt nur 3 Werte
        For j = 1 To 3
            'Schleife über die 3 DgSumSpalten (5,8,11)
                zDgVer = ArrzDgs(0) - 1             'VereinsZeile DgZeilenNr
                zShVer = z1 + zDgVer - 1            'VereinsZeile SheetZeilenNr
                sDg = Arr3(j)                       'SumSpalte    DgSpaltenNr
                sShSum = s1 - 1 + sDg               'SumSpalte    SheetSpaltenNr
            'D-Liste erstellen
                D = "" '= Durchgang (P, K1 oder K2) = Liste aller Werte + ihrer ZellPosition
                For i = 0 To UBound(ArrzDgs)
                    zDg = ArrzDgs(i)    'ArrzDgs(0) = "7" wegen zDgs = "7|8|9|10|11|12|"
                    If ArrDg(zDg, sDg) = "" Then ArrDg(zDg, sDg) = 0
                    D = D + Format(ArrDg(zDg, sDg), "00.0") _
                          + "°" + Format(zDg, "00") + "°" + Format(sDg, "00") + "|"
                          
                          Exit Sub
                Next
                D = Left(D, Len(D) - 1)
                ' = "23,9°07°05|...|28,1°12°05" (Zeilen 07-12, also 6 Aktive)
                '    23,9°07°05 --> 23,9 P. des Aktiven Zeile 7, Spalte 5 (Pflicht)
                'd trägt jetzt alle z. B. Kür1-Punkte einer Mannschaft (falls gerade j=2)
            'ArrD erstellen     (damit man sortieren kann)
                ArrD = Split(D, "|"): QuickSort ArrD
                'ArrD trägt jetzt eine sortierte D-Liste: 23,9°07°05|...|28,1°12°05
                c = 0: Sum = 0  'showArray ArrD
            'Sum        (max 4 höchste Werte addieren = Ergebnis; VnNn = Streichwertungen)
                For i = UBound(ArrD) To 0 Step -1
                    'Schleife über die sortierte D-Liste (vom höchsten zum tiefsten Wert)
                    c = c + 1: w = CDbl(Left(ArrD(i), 4))   '28,1
                        If Left(ArrD(i), 1) = "-" Then      'ArrD(i) = ""-31,9°10°11"
                            'K1 wurde eingetragen; Einzelergebnis muss neu berechnet werden
                            ArrD(i) = "00,0" + Mid(ArrD(i), 6)
                            z = CInt(Mid(ArrD(i), 6, 2)): s = CInt(Mid(ArrD(i), 9, 2))
                            G1 = CDbl(ArrDg(z, 5)) + CDbl(ArrDg(z, 8))
                            'Eintrag Einzelergebnis in das Dg auf Sheet T4
                                .Cells(z1 - 1 + z, s1 + 11) = Format(G1, "0.0")
                            'Eintrag Einzelergebnis in ArrDg
                                ArrDg(z, 11) = "0,00"
                                ArrDg(z, 12) = Format(G1, "0.0")
                        End If
                    z = CInt(Mid(ArrD(i), 6, 2)): s = CInt(Mid(ArrD(i), 9, 2))
                    'Der Wert W steht in der Dg-Zelle (z,s)
                    'FontFarbe für Dg-Werte auf dem Sheet T4
                        If c < 5 Then
                            Sum = Sum + w
                            'Einfärben der 4 höchsten Wertungen: Schwarz
                            If w > 0 Then .Cells(z1 - 1 + z, s1 - 1 + s).Font.Color = vbBlack
                        Else
                            'Einfärben der Streichnoten: Grau
                            If w > 0 Then .Cells(z1 - 1 + z, s1 - 1 + s).Font.Color = RGB(99, 99, 99) 'Farbe grau
                            'Einfärben der 0,0: m/w-Farbe (soll nicht sichtbar sein)
                            If w = 0 Then .Cells(z1 - 1 + z, s1 - 1 + s).Font.Color = .Cells(z1 - 1 + z, s1 - 1 + s).Interior.Color
                        End If
                    'Gesamt     Einzelergebnis      (Wert ändert sich, falls sich P/K1/K2 ändert)
                        'G1 = CDbl(ArrDg(z, 5)) + CDbl(ArrDg(z, 8)) + CDbl(ArrDg(z, 11))
                        If G1 > 0 Then
                            'Eintrag Einzelergebnis in das Dg auf Sheet T4
                                .Cells(z1 - 1 + z, s1 + 11) = Format(G1, "0.0")
                            'Eintrag Einzelergebnis in ArrDg
                                ArrDg(z, 12) = Format(G1, "0.0")
                        End If
                Next
                Sum = WorksheetFunction.Round(Sum, 2)
                If Sum > 0 Then
                    'Eintrag in das Dg auf Sheet T4
                        .Cells(zShVer, sShSum) = Format(Sum, "0.0")
                    'Eintrag in ArrDg
                        ArrDg(zDgVer, sDg) = Format(Sum, "0.0")
                End If
            'Gesamt     Mannschaftsergebnis (Wert ändert sich, falls sich Sum ändert)
                G = CDbl(ArrDg(zDgVer, 5)) + CDbl(ArrDg(zDgVer, 8)) + CDbl(ArrDg(zDgVer, 11))
                G = WorksheetFunction.Round(G, 2)
                If G > 0 Then
                    'Eintrag Mannschaftsergebnis in das Dg auf Sheet T4
                        .Cells(zShVer, s1 + 11) = G 'Format(G, "0.0")
                    'Eintrag Mannschaftsergebnis in ArrDg
                        ArrDg(zDgVer, 12) = G 'Format(G, "0.0")
                End If
        Next
    End With
    'showArray2D ArrDg
End Sub

Sub T4_Put_DgLigaRangEinzel(s1%, z1%, z2%)
    'Called from    T4_AutoCalc
    'Action         Schreibt bei LigaDesigns den Rang jeden Wettkämpfers in die vorletzte Dg-Spalte
    
    Exit Sub
    
    
    'Exit
        If ArrDg(2, 1) = "M0V1" Then Exit Sub
        If ArrDg(2, 1) = "M0V3" Then Exit Sub
    'Vorbereitung
        Dim A$, G$, Pkt$, s$, c%, i%, zDg%, Arr1() As String
        With Sheets("T4")
    'Sammeln der Gesamtpunkte aller Wettkämpfer
        For i = 7 To z2 - z1
            If ArrDg(i, 4) <> "" Then       'Spalte 4 enthält einen Namen
                G = CStr(ArrDg(i, 12))      'G = "85,3" = Gesamtpunkte des Wettkämpfers in Zeile i
                If CStr(G) = "" Then Stop
                If IsNumeric(G) Then        'IsNumeric("")=False
                    If CDbl(G) > 0 Then
                        s = s + Format(G, "000.0") + " " + Format(i, "00") + vbCrLf
                        ' = Liste, je  "085,3 09"-Zeilen (Gesamtpunkte, zDg) für Sortierung
                    Else
                        Exit Sub
                    End If
                Else
                    .Range(.Cells(z1 + 6, s1 + 12), .Cells(z2 - 1, s1 + 12)) = ""
                    'Falls ein Name (noch) keine GesamtPunkte hat, werden keine Ränge geschrieben
                    Exit Sub
                End If
            End If
        Next
        s = Delete_EmptyEndRowsInString(s)
    'Sort
        Arr1 = Split(s, vbCrLf)
        QuickSort Arr1
    'Rang eintragen
        For i = UBound(Arr1) To 0 Step -1
            c = c + 1                       'c   = 1          = Rang
            A = Arr1(i)                     'a   = "085,3 09" = Punkte, DgZeile
            zDg = CInt(Right(A, 2))         'zDg = 9          = DgZeile
            
            If c = 1 Then
                ArrDg(zDg, 13) = c: .Cells(z1 - 1 + zDg, s1 + 12) = c
            Else
                If Left(A, 5) = Pkt Then
                    'Punktgleichheit mit den Vorgänger-Punkten
                    ArrDg(zDg, 13) = c - 1: .Cells(z1 - 1 + zDg, s1 + 12) = c - 1
                Else
                    ArrDg(zDg, 13) = c: .Cells(z1 - 1 + zDg, s1 + 12) = c
                End If
            End If
            Pkt = Left(A, 5)                'Pkt = "085,3"    = Punkte
        Next
    'Finals
        End With
End Sub

Function Get_DgSpaltenNrn_GN(BuchstabenKette$) As String
    'Called from    T4_AutoCalc
    'Action         liefert s = "G16N03,G32N19"
    '                         = je DgSpaltenNr der Ges.-Spalte/zugehörige Namen-Spalte
    '               G16 = Der Wert für "Ges." steht in Spalte 16
    '               N03 = Der Name steht im Dg jeweils in Spalte 3
    '               2 G-Spalten z. B. bei '1988-05-10 WM15 Birmingham/USA (Mannschaft)'
    
    Dim B$, s$, i%
    B = BuchstabenKette '= "iRMT-AAE-AAAE-AAAE-G-RWT-AAE-AAAE-AAAE-Gi"
    For i = 1 To Len(B)
        If InStr(1, "MWXG", Mid(B, i, 1)) > 0 Then s = s + Format(i, "00")
    Next
    '
    If Len(s) = 4 Then          's = "0412" (z. B. bei Buli)
        s = "G" + Mid(s, 3, 2) + "N" + Mid(s, 1, 2)
        Get_DgSpaltenNrn_GN = s 's = "G12N04"
    End If
    If Len(s) = 8 Then          's = "03161932"
        s = "G" + Mid(s, 3, 2) + "N" + Mid(s, 1, 2) + ",G" + Mid(s, 7, 2) + "N" + Mid(s, 5, 2)
        Get_DgSpaltenNrn_GN = s 's = "G16N03,G32N19"
    End If
    'show s
End Function

Sub CalcRangEinzel(EinzW$, A1%, A2%, a3%, a4%, s1%, z1%)
    'Called from    CalcDg17
    'EinzW          = "|055,1|052,4|091,1|100,9|079,8|096,2|032,7|085,5|076,0|086,6|090,3|092,8|"
    '               = Endergebnisse Einzelwertung
    'Vorbereitung
        Dim s$, rg$, RM1$, RM2$, w$, Anz%, AnzM1%, AnzM2%, i%, j%
        Dim ArrE1() As String, ArrE2() As String, arrRM1() As String, arrRM2() As String
        AnzM1 = A2 - A1 + 1 'Anzahl Mitglieder Mannschaft1  '7
        AnzM2 = a4 - a3 + 1 'Anzahl Mitglieder Mannschaft2  '5
        Anz = AnzM1 + AnzM2                                 '12
        s = Delete_LastChar(EinzW) 'damit QuickSort funktioniert
        ArrE1 = Split(s, "|")
        'arrE1 --> |055,1|052,4|091,1|100,9|079,8|096,2|032,7|085,5|076,0|086,6|090,3|092,8
        ArrE2 = Split(s, "|"): QuickSort ArrE2 ': s = Join(arr1, "|")
        'arrE2 --> |032,7|052,4|055,1|076,0|079,8|085,5|086,6|090,3|091,1|092,8|096,2|100,9
    'RM1, RM2 (Rangliste Mannschaft1/Mannschaft2)
        For i = 1 To Anz '12
            w = ArrE1(i) '"055,1"
            'Rang für w feststellen:
            For j = 1 To Anz
                If ArrE2(j) = w Then rg = CStr(Anz - j + 1): Exit For
            Next
            If i < AnzM1 + 1 Then RM1 = RM1 + rg + "|" Else RM2 = RM2 + rg + "|"
        Next
        RM1 = Delete_LastChar(RM1)  'RM1 = "10|11|4|1|8|2|12"   Rangliste Mannschft1
        RM2 = Delete_LastChar(RM2)  'RM2 = "7|9|6|5|3"          Rangliste Mannschft2
        arrRM1 = Split(RM1, "|"): arrRM2 = Split(RM2, "|")
    'Paste
        Paste_1DArrayToCol "T4", z1 + A1 - 1, s1 + 1, arrRM1
        Paste_1DArrayToCol "T4", z1 + a3 - 1, s1 + 1, arrRM2
    
End Sub

Sub ColorBest4_CalcMGes8(A1%, a3%, s1%, z1%, W1$, W2$, W3$, W4$, W5$, W6$)
    'Called from    CalcDg17
    'A1 ... A4      = ZeilenNrDesign; A1/A2=1./letzterAktiverMannschaft1; A3/A4=ditoMannschaft2
    'W1 ... W6      = "25,7|0|28,2|27|...|"; Pfl K1 K2 von Mannschaft 1; dito Mannschaft 2
    'Range r2       = nach und nach entstehendes Range aller Zellen, die graue Schrift zeigen sollen
    'Streichwerte   = Werte, die nicht zu den Best4 gehören
    'Status         EE 0; Caller sets EE 1
    '               Schriftfarbe = Schwarz (im gesamten D01-Design)
    '               Alle Ges.-Werte aller Aktiven sind vorhanden
    'Action         In den Spalten Pflicht, Kür1End, Kür2End des aktuellen D01-Designs
    '                werden Streichwerte in grauer Schrift dargestellt
    '               MannschaftsErgebnisse von Pflicht, Kür1, Kür2 werden ggf. eingetragen
    '               MannschaftsEndErgebnisse werden ggf. eingetragen
    
    'Vorbereitung
        Dim Best$, M1Ges$, M2Ges$, i%, zM1%, zM2%, r1 As Range, r2 As Range, r3 As Range
        Dim arrSum() As String, ArrS() As Integer, arrZ() As Integer
        With Sheets("T4"): ReDim arrSum(1 To 6): ReDim ArrS(1 To 6): ReDim arrZ(1 To 6)
        zM1 = z1 + a3 - 5 'ZeilenNr Mannschaft1 ErgebnisZeile
        zM2 = z1 + a3 - 3 'ZeilenNr Mannschaft2 ErgebnisZeile
    'W1 (Pflicht-Werte Mannschaft 1)        'W1   = "025,7|024,4|026,7|027,9|028,2|028,2|000,0|"
        Best = Get_Best4(W1)                'Best = "|028,2|028,2|027,9|026,7|"
        Set r2 = Get_RangeUnion(W1, Best, z1 + A1 - 1, s1 + 4) '(,,zSheetTopCell,sSheetTopCell)
        arrSum(1) = Get_SumBest4(Best): arrZ(1) = zM1: ArrS(1) = s1 + 4
    'W2 (Kür1End-Werte Mannschaft 1)
        Best = Get_Best4(W2)
        Set r1 = Get_RangeUnion(W2, Best, z1 + A1 - 1, s1 + 8)
        Set r2 = Union(r2, r1)
        arrSum(2) = Get_SumBest4(Best): arrZ(2) = zM1: ArrS(2) = s1 + 8
    'W3 (Kür2End-Werte Mannschaft 1)
        Best = Get_Best4(W3)
        Set r1 = Get_RangeUnion(W3, Best, z1 + A1 - 1, s1 + 12)
        Set r2 = Union(r2, r1)
        arrSum(3) = Get_SumBest4(Best): arrZ(3) = zM1: ArrS(3) = s1 + 12
    'W4 (Pflicht-Werte Mannschaft 2)
        Best = Get_Best4(W4)
        Set r1 = Get_RangeUnion(W4, Best, z1 + a3 - 1, s1 + 4)
        Set r2 = Union(r2, r1)
        arrSum(4) = Get_SumBest4(Best): arrZ(4) = zM2: ArrS(4) = s1 + 4
    'W5 (Kür1End-Werte Mannschaft 2)
        Best = Get_Best4(W5)
        Set r1 = Get_RangeUnion(W5, Best, z1 + a3 - 1, s1 + 8)
        Set r2 = Union(r2, r1)
        arrSum(5) = Get_SumBest4(Best): arrZ(5) = zM2: ArrS(5) = s1 + 8
    'W6 (Kür2End-Werte Mannschaft 2)
        Best = Get_Best4(W6)
        Set r1 = Get_RangeUnion(W6, Best, z1 + a3 - 1, s1 + 12)
        Set r2 = Union(r2, r1)
        arrSum(6) = Get_SumBest4(Best): arrZ(6) = zM2: ArrS(6) = s1 + 12
    '6 MannschaftsErgebnisse (je Mannschaft: Pflicht, Kür1, Kür2)
        For i = 1 To 6
            Set r3 = .Cells(arrZ(i), ArrS(i)): If r3.Value = "" Then r3.Value = arrSum(i)
        Next
    '2 MannschaftsGesamtErgebnisse
        M1Ges = Format(CSng(arrSum(1)) + CSng(arrSum(2)) + CSng(arrSum(3)), "0.0")
        M2Ges = Format(CSng(arrSum(4)) + CSng(arrSum(5)) + CSng(arrSum(6)), "0.0")
        Set r3 = .Cells(zM1, s1 + 14): If r3.Value = "" Then r3.Value = M1Ges
        Set r3 = .Cells(zM2, s1 + 14): If r3.Value = "" Then r3.Value = M2Ges
    '2 MannschaftsRänge
        If M1Ges > M2Ges Then
            .Cells(zM1, s1 + 1) = "1": .Cells(zM2, s1 + 1) = "2"
        Else
            .Cells(zM2, s1 + 1) = "1": .Cells(zM1, s1 + 1) = "2"
        End If
    'Final
        'r2.Select
        r2.Font.Color = RGB(99, 99, 99) 'Grey
        End With
End Sub

Function Get_SumBest4(Best$) As String
    'Called from    SetColorBest4
    'Best           = "|028,2|028,2|027,9|026,7|"
    
    'Vorbereitung
        Dim i%, Sum#, Arr1() As String
        Arr1 = Split(Best, "|")
    'Sum
        For i = 1 To 4
            Sum = Sum + CSng(Arr1(i))
        Next
        Sum = WorksheetFunction.Round(Sum, 2)
    Get_SumBest4 = Sum 'Format(Sum, "0.0")
End Function

Function Get_RangeUnion(w, B, zs, ss) As Range
    'Called from    SetColorBest4
    'W              = "025,7|024,4|026,7|027,9|028,2|028,2|000,0|"
    'B              = "|028,2|028,2|027,9|026,7|" = Best4 von W1
    'sS             = SpaltenNrSheet (SpaltenNr, welche untereinander die Werte W1 trägt
    'zS             = ZeilenNrSheet  (ZeilenNr, welche den ersten W1-Wert trägt)
    'RangeUnion     = Range aller Zellen, die graue Schrift zeigen sollen
    
    'Vorbereitung
        Dim v$, c%, i%, r As Range, Arr1() As String
        With Sheets("T4")
    'Range
        Arr1 = Split("|" + w, "|")
        For i = 1 To UBound(Arr1) - 1
            v = Arr1(i)         'v = "25,7" = one Value
            If InStr(1, B, "|" + v + "|") = 0 Then
                'v ist kein Wert der Best4
                c = c + 1
                If c = 1 Then
                    Set r = .Cells(zs + i - 1, ss)
                Else
                    Set r = Union(r, .Cells(zs + i - 1, ss))
                End If
                B = Replace(B, "|" + v + "|", "|", , 1)
            End If
        Next
    'Final
        Set Get_RangeUnion = r
        End With
End Function

Function Get_Best4(WW$) As String
    'Called from    SetColorBest4
    'W              = "025,7|024,4|026,7|027,9|028,2|028,2|000,0|"
    
    'Vorbereitung
        Dim s$, w$, i%, Arr1() As String
        w = Left(WW, Len(WW) - 1)
        Arr1 = Split(w, "|"): QuickSort Arr1
        's = Join(arr1, "|") 's = "|000,0|024,4|025,7|026,7|027,9|028,2|028,2"
        s = "|"
    'Best4
        For i = UBound(Arr1) To UBound(Arr1) - 3 Step -1
            s = s + Arr1(i) + "|"
        Next 's = "|028,2|028,2|027,9|026,7|" = Sammlung Best4
    'Return
        Get_Best4 = s
End Function

Sub SetColWidth(s$)
    'Called from    CalcD34, CalcD43, ...
    's              = "02.00,1,17,33" = SpBreite 02.00 für die DesignSpaltenNrn 1, 17, 33
    
    Dim Nn$, N$, i%, s1%, w!, Arr1() As String
    s1 = CInt(ArrC(39))         's1 = T4-SpaltenNr of BorderLeft
    w = CSng(Left(s, 5))        'W  = Spaltenbreite
    Nn = Mid(s, 6) + ","        'nn = zu bearbeitende Design-SpaltenNummern ',1,17,33,
    Arr1 = Split(Nn, ",")
    With Sheets("T4")
    For i = 1 To UBound(Arr1)
        N = CStr(Arr1(i))
        If IsNumeric(N) Then
            .Columns(s1 - 1 + CInt(N)).ColumnWidth = w
        End If
    Next
    End With
End Sub

Sub T4_CalcM2V8_SingleRanking(EwOrEm$, Data$)
    'Called from    T4_M2V8_SingleRanking
    'SinglePlacing  = Rang Einzelwertung
    'Data           = Lines unsorted, each: EndSumOneCompetitor|zDg
    '               069,90|022  088,50|008  075,30|010  086,60|020  ...
    
    'Vorbereitung
        Dim L$, Rg1$, Rg2$, v$, c%, i%, sMW%, zDg%, zDg1%, zDg2%
        Dim Arr1() As String, Arr2() As String, Arr3() As String
        v = vbCrLf
        If Data = "" Then Exit Sub
        'show Data: Stop
    'Arr1: Data sorted
        Arr1 = Split(Data, v)
        QuickSortDown Arr1
        'showArray Arr1:  Stop
    'Arr2: Zeile Wert Rang
        ReDim Arr2(1 To UBound(ArrDg, 1))
        'Alle Arr2-Zeilen mit "Dg-Zeile 001|", "Dg-Zeile 002|", ... füllen
            For i = 1 To UBound(Arr2)
                Arr2(i) = "Dg-Zeile " + Format(i, "000") + "|"
            Next
        'Falls Dg-Name existiert: Arr2-Zeilen ergänzen:
            'mit "088,50|Rang 001|" (SingleCompetitorEndSum|Rang|)
            For i = 0 To UBound(Arr1)
                zDg = CInt(Right(Arr1(i), 3))
                Arr2(zDg) = Arr2(zDg) + Left(Arr1(i), 6) + "|" _
                            + "Rang " + Format(i + 1, "000") + "|"
            Next
        'Arr2:  'Dg-Zeile 001| ... 'Dg-Zeile 006| 'Dg-Zeile 007|088,50|Rang 002| ...
        'showArray Arr2:  Stop
    'Punktgleichheit
        For i = 0 To UBound(Arr1)
            If i > 0 Then
                If Left(Arr1(i), 6) = Left(Arr1(i - 1), 6) Then
                    'Arr1 ist nach PunktZahlen sortiert;
                    'gleiche PunktZahlen stehen dort also direkt aufeinanderfolgend
                    c = c + 1
                    zDg1 = CInt(Right(Arr1(i), 3)):     Rg1 = Mid(Arr2(zDg1), 26, 3)
                    zDg2 = CInt(Right(Arr1(i - 1), 3)): Rg2 = Mid(Arr2(zDg2), 26, 3)
                    If CInt(Rg2) < CInt(Rg1) Then Rg1 = Rg2
                    Arr2(zDg1) = Left(Arr2(zDg1), 25) + Rg1 + "|" + CStr(c)
                    Arr2(zDg2) = Left(Arr2(zDg2), 25) + Rg1 + "|" + CStr(c)
                End If
            End If
        Next
        'showArray Arr2: Stop
    'Arr3 für Paste ab Dg(7,13)
        ReDim Arr3(1 To UBound(Arr2))
        For i = 1 To UBound(Arr3)
            If i < UBound(Arr2) - 6 Then
                If Len(Arr2(i + 6)) > 13 Then Arr3(i) = CStr(CInt(Mid(Arr2(i + 6), 26, 3)))
            End If
        Next
        'showArray Arr3: Stop
    'Paste
        If EwOrEm = "Ew" Then sMW = 12 Else sMW = 25
        Paste_1DArrayToCol "T4", CInt(ArrC(37)) + 6, CInt(ArrC(39)) + sMW, Arr3
End Sub

Sub T4_M2V8_SingleRanking()
    'Called from    [UserClick Dg-Button 'SingleRanking'] CmdSingleRanking_Click
    'SinglePlacing  = Rang Einzelwertung
    '...RowNrs      = "7|8|9|10|11|14|15|16|17|...|154|155|156"
    
    'Vorbereitung
        Dim Em$, Ew$, v$, i%, z%, zCompetitors$, Arr1() As String
        v = vbCrLf
    'w-Competitors with SingleCompetitorEndSum
        For i = 7 To UBound(ArrDg, 1) - 1
            If ArrDg(i, 4) <> "" And ArrDg(i, 12) <> "" Then
                Ew = Ew + v + Format(ArrDg(i, 12), "000.00") + "|" + Format(i, "000")
            End If
        Next
        Ew = Mid(Ew, 3) 'Einzelergebnisse weiblich  088,50|007 ... 049,40|029 ...
    'm-Competitors with SingleCompetitorEndSum
        For i = 7 To UBound(ArrDg, 1) - 1
            If ArrDg(i, 17) <> "" And ArrDg(i, 25) <> "" Then
                Em = Em + v + Format(ArrDg(i, 25), "000.00") + "|" + Format(i, "000")
            End If
        Next
        Em = Mid(Em, 3) 'Einzelergebnisse männlich  090,20|007 ... 032,60|051 ...
        'show "Ew" + v + v + Ew: Stop:show "Em" + v + v + Em: Stop
    'Calc and write
        T4_CalcM2V8_SingleRanking "Ew", Ew
        T4_CalcM2V8_SingleRanking "Em", Em
End Sub




