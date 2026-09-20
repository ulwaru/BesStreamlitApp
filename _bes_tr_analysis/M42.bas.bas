Attribute VB_Name = "M42"
Option Explicit

Sub zzz_M42()
    EE 1
    
    showProcs "pfeil"
    
    'RenameModule "Tabelle1", "T6"
    'RenameModule "Modul1", "M41"
    Beep
End Sub

Sub zzz_M42_Struc()
    T4_AutoCalc
        T4_AutoCalcLiga     '+ Exit
            T4_Get_AllTeamNameRowNrs_Liga
            T4_Get_AllTeamMemberRowNrs
            T4_Put_DgLigaSum
            T4_Put_DgLigaRangEinzel
        T4_AutoCalcM2V8      '+ Exit
            T4_Put_CalcedValue_TypAAE
            T4_Get_AllTeamMemberRowNrs
            T4_CalcAndWrite_SumP_SumK1_SumK2_Best3_OneTeam
            T4_CalcAndWrite_SumP_SumK1_SumK2_Best4_OneTeam
            T4_Calc_M2V8_TeamEndSum
            T4_Calc_M2V8_OneCompetitorSum
        'Not M2V8, not Liga
        T4_CalcAndWrite_AAAE
            T4_Get_AAAE_ColumnNrs
        T4_CalcAndWrite_AAE
            T4_Get_AAE_ColumnNrs
            T4_Write_OneSum_MayBeFixValue
        T4_CalcAndWrite_EEEG
            T4_Get_EEEG_ColumnNrs
            T4_Write_OneSum_MayBeFixValue
        T4_CalcAndWrite_EEG
            T4_Get_EEG_ColumnNrs
            T4_Write_OneSum_MayBeFixValue
        T4_Set_ArrowsForFixValues
End Sub

Sub T4_AutoCalc()
    'Called from    T4_ActionsOnDgClickNext
    'Status         Sub wird bei jedem Dg-Klick aufgerufen
    '               Durchsuchung aller Zeilen: Berechnung, falls nur 1 der beteiligten Werte fehlt
    'B              = Buchstabenkette = Abfolge der Zeile2-Buchstaben
    '               iRMT-AAE-AAAE-G-RWT-AAE-AAAE-Gi             1976-07-03 WM09 Tulsa/Oklahoma/USA (Synchron)
    '               iRMT-E-AAE-AAE-G-RWT-E-AAE-AAE-Gi           1976-07-03 WM09 Tulsa/Oklahoma/USA (Einzel)
    '               iYRMT-E-AAE-AAE-G-RWT-E-AAE-AAE-Gi          1974-03-23 WM08 Johannesburg/SA
    '               iRMT-AAE-AAAE-AAAE-G-RWT-AAE-AAAE-AAAE-Gi   1982-05-13 WM12 Bozeman_USA (Synchron)
    '               iRVXEAAEAAEGRi                              1992-12-05 BuliE BadKreuznach
    
    'ToDo           Handling bei keinen/wenigen Competitors
    '               T4_Write_OneSum_MayBeFixValue Sum#, z%, s%
    
    'Andere Berechnungen für Mannschaftswettkampf M2V8 und Liga
        If ArrC(42) Like "* Liga\*" Then T4_AutoCalcLiga: Exit Sub
        If ArrDg(2, 1) = "M2V8" Then T4_AutoCalcM2V8: Exit Sub
    
    'Vorbereitung
        Dim AAAE$, AAE$, B$, D$, EEG$, EEEG$, FoldName$, Ges1$, GesAll$, m$, NAME$
        Dim s$, SpGN$, v$, Verein$, VM$, WM$
        Dim A%, C1%, C2%, C3%, C4%, i%, j%, k%, N%, RankNew%, RankOld%
        Dim s1%, s2%, sNam%, z1%, z2%, zDg%, A1#, A2#, a3#, a4#
        Dim Arr(), ArrZ2(), Arr1() As String, Arr2() As String
        'Ges1$,GesAll$,RankNew%,RankOld%
        z1 = CInt(ArrC(37)): z2 = CInt(ArrC(38)): s1 = CInt(ArrC(39)): s2 = CInt(ArrC(40))
        v = vbCrLf: With Sheets("T4")
    'Einstellung
    'ArrDg 'wird bei jedem Klick in ein Dg neu geladen
        'ArrDg wurde gerade geladen (mit 'FillArrDg' in T4_ActionsOnDgClickNext)
    'Buchstabenkette Zeile 2                'wird bei jedem Klick in ein Dg neu geladen
        B = Get_BuchstabenKette_T4Zeile2    'B = "iRMT-AAE-AAAE-AAAE-G-RWT-AAE-AAAE-AAAE-Gi"
        '                          bei Buli: B = "iRVXEAAEAAEGRi"
        'Beispiel Sy-Wettkampf:
        'B = "iRWT-AAE-AAAE-G-RMT-AAE-AAAE-Gi"      AAE: Pfl+SyAbzug=End 'Len(B)=31=AnzSpalten
        '              AAAE = "10|25|"              2 AAAE-Gruppen: 1. beginnt bei 10, 2. bei 25
    'DgSpaltenNr der NamenSpalten
        SpGN = Get_DgSpaltenNrn_GN(B)
        '    = "G16N03,G32N19" = je DgSpaltenNr der Ges.-Spalte/zugehörige Namen-Spalte
        '    = "G12N04" bei Buli; G12 = Der Wert für "Ges." steht in Spalte 12
    'FixValue
        T4_Handle_x_asFixValue B, z1, z2, s1, s2
    'AAAE
        T4_CalcAndWrite_AAAE B, z1, z2, s1  'B new = "iRWT-AAE-XXXE-G-RMT-AAE-XXXE-Gi" (keine "AAAE" mehr)
    'AAE
        T4_CalcAndWrite_AAE B, z1, z2, s1   'B new = "iRWT-XXE-XXXE-G-RMT-XXE-XXXE-Gi" (keine "AAE" mehr)
    'EEEG
        T4_CalcAndWrite_EEEG B, SpGN, z1, z2, s1 'B new = "iRWT-XXE-XXXE-G-RMT-XXE-XXXE-Gi" (keine "EEEG")
    'EEG
        T4_CalcAndWrite_EEG B, SpGN, z1, z2, s1  'B new = "iRWT-XXX-XXXX-X-RMT-XXX-XXXX-Xi" (keine "EEG" mehr)
    'Arrows for FixValues
        T4_Set_ArrowsForFixValues B, z1, z2, s1
    'Finals
        End With
End Sub

Sub T4_Handle_x_asFixValue(B$, z1%, z2%, s1%, s2%)
    'Called from    T4_AutoCalc
    'x-Wert         "34,55x"; Wert mit angehängtem x in einer E- oder G-Spalte
    'B              Buchstabenkette Zeile 2 oberhalb des aktuellen Dg --> Dg(5,i)
    'Action         Suche nach einem x-Wert;
    '               in der Spalte rechts des x-Wertes wird ein Pfeil geschrieben;
    '               in Dg(1,4) wird die Position des x-Wertes festgehalten
    
    
    Dim F$, Wert$, s&, z&, w#
    With Sheets("T4")
    'FixValue-Zelle vorbereiten
        F = ArrDg(1, 4)
        If F Like "*(*" Then
            F = "FixValues: " + Mid(F, InStr(1, F, "("))
        Else
            F = "FixValues: "
        End If
    'x suchen
        For s = 1 To s2 - s1 + 1
            If Mid(B, s, 1) = "E" Or Mid(B, s, 1) = "G" Then
                For z = 6 To z2 - z1
                    If ArrDg(z, s) Like "*#x" Then
                        F = F + "(" + CStr(z) + "|" + CStr(s) + ")"
                        Wert = CStr(ArrDg(z, s)): Wert = Replace(Wert, "x", "")
                        w = CDbl(Wert): ArrDg(z, s) = Wert
                        .Cells(z1 - 1 + z, s1 - 1 + s).Value = w
                        .Cells(z1 - 1 + z, s1 - 1 + s).Font.Color = vbRed
                    End If
                Next
            End If
        Next
    'FixValue-Zelle füllen
        ArrDg(1, 4) = F: .Cells(z1, s1 + 3) = F
    'Pfeil setzen
    'T4_Set_ArrowsForFixValues B, z1, z2, s1
    End With
    
End Sub

Sub T4_CalcAndWrite_AAAE(B$, z1%, z2%, s1%)
    'Called from    T4_AutoCalc
    'B              = "iRWT-AAE-AAAE-G-RMT-AAE-AAAE-Gi" (Len(B)=31=AnzSpalten; Sy-Wettkampf)
    'AAAE           A+A+A=E:        Haltung+SyAbzug+Schwierigkeit=End
    '               = "10|25|"      2 AAAE-Gruppen: 1. beginnt bei 10 (Ti), 2. bei 25 (Tu)
    
    'Vorbereitung
        Dim AAAE$, D$, A%, i%, zDg%, A1#, A2#, a3#, a4#, Arr1() As String
        With Sheets("T4")
    'AAAE
        AAAE = T4_Get_AAAE_ColumnNrs(B) 'AAAE = "" oder "10|15|30|35|" = je Position des ersten "A" eines AAAE
                              'B wurde reduziert  'B old = "iRWT-AAE-AAAE-G-RMT-AAE-AAAE-Gi"
                                                  'B new = "iRWT-AAE-XXXE-G-RMT-AAE-XXXE-Gi"
        If AAAE = "" Then Exit Sub
        ' z.B. AAAE = "6|":  Die DgBuchstabenZeile enthält nur 1x "AAAE" oder "AAA-E";
        '                    das 1. A beginnt bei 6
    'Action
        Arr1 = Split(AAAE, "|")
        For i = 0 To UBound(Arr1) - 1
            A = CInt(Arr1(i))     '= 10 = Dg-SpaltenPosition des ersten "A" des i. AAAE
            'Check pro Dg-Zeile, ob 3 der 4 Werte vorhanden sind
            For zDg = 6 To z2 - z1
                'Schleife über alle relevanten Dg-Zeilen
                'A1 bis A4 sind die 4 Werte, die zu A, A, A und E gehören
                If Not IsNumeric(CStr(ArrDg(zDg, A + 0))) Then D = "0" Else D = "1":         A1 = ArrDg(zDg, A)
                If Not IsNumeric(CStr(ArrDg(zDg, A + 1))) Then D = D + "0" Else D = D + "1": A2 = ArrDg(zDg, A + 1)
                If Not IsNumeric(CStr(ArrDg(zDg, A + 2))) Then D = D + "0" Else D = D + "1": a3 = ArrDg(zDg, A + 2)
    
                If Mid(B, A + 3, 1) = "-" Then 'AAA-E
                    If Not IsNumeric(CStr(ArrDg(zDg, A + 4))) Then D = D + "0" Else D = D + "1": a4 = ArrDg(zDg, A + 4)
                Else
                    If Not IsNumeric(CStr(ArrDg(zDg, A + 3))) Then D = D + "0" Else D = D + "1": a4 = ArrDg(zDg, A + 3)
                End If
    
                'Calc
                If D = "0111" Then
                                       A1 = a4 - A2 - a3: ArrDg(zDg, A + 0) = A1: .Cells(z1 - 1 + zDg, s1 + A - 1) = A1
                ElseIf D = "1011" Then A2 = a4 - A1 - a3: ArrDg(zDg, A + 1) = A2: .Cells(z1 - 1 + zDg, s1 + A + 0) = A2
                ElseIf D = "1101" Then
                    'Die Werte A1, A2, A4 sind vorhanden
                    'A3 ist nicht numerisch
                    a3 = a4 - A1 - A2: ArrDg(zDg, A + 2) = a3: .Cells(z1 - 1 + zDg, s1 + A + 1) = a3
                            ''old
                            'If CStr(ArrDg(zDg, a + 2)) = "" Then
                            '    'Der 3. Wert in AAAE (2. Kür) soll "" bleiben, nicht "0,00"
                            '    .Cells(z1 - 1 + zDg, s1 + a + 1) = ""
                            'Else
                            '    A3 = A4 - A1 - A2: ArrDg(zDg, a + 2) = A3
                            '    .Cells(z1 - 1 + zDg, s1 + a + 1) = A3
                            'End If
                ElseIf D = "1110" Then
                    a4 = A1 + A2 + a3
                    If Mid(B, A + 3, 1) = "-" Then 'AAA-E
                        ArrDg(zDg, A + 4) = a4: .Cells(z1 - 1 + zDg, s1 + A + 3) = a4
                    Else
                        ArrDg(zDg, A + 3) = a4: .Cells(z1 - 1 + zDg, s1 + A + 2) = a4
                    End If
                End If
            Next
        Next
    'Finals
        End With
End Sub

Sub T4_CalcAndWrite_AAE(B$, z1%, z2%, s1%)
    'Called from    T4_AutoCalc
    'B              = "iRWT-AAE-XXXE-G-RMT-AAE-XXXE-Gi" (Len(B)=31=AnzSpalten; Sy-Wettkampf)
    'AAE            A + A = E:      Haltung + SyAbzugBeiPflicht = End
    '               = "6|21|"       2 AAE-Gruppen: 1. beginnt bei Dg-Spalte 6 (Ti), 2. bei 21 (Tu)
    '                               AAE in den Spalten 6,7,8 und 21,22,23
    
    'Vorbereitung
        Dim AAE$, D$, i%, sDg%, zDg%, A1#, A2#, a3#, a4#, Arr1() As String
    'AAE
        AAE = T4_Get_AAE_ColumnNrs(B) 'AAE = "" oder "6|21|" = je SpaltenNr des ersten "A" eines AAE
                            'B wurde reduziert  'B old = "iRWT-AAE-AAAE-G-RMT-AAE-AAAE-Gi"
                                                'B new = "iRWT-AAE-XXXE-G-RMT-AAE-XXXE-Gi"
        If AAE = "" Then Exit Sub
        With Sheets("T4")
    'Action
        Arr1 = Split(AAE, "|")
        For i = 0 To UBound(Arr1) - 1
            sDg = CInt(Arr1(i))     '= 6 = SpaltenPosition des ersten "A" eines AAE
            'Check pro Dg-Zeile, ob 2 der 3 Werte vorhanden sind
            For zDg = 6 To z2 - z1
                'Schleife über alle relevanten Dg-Zeilen
                'Zu welchem A oder E (in A+A=E) ist ein Wert vorhanden?
                If Not IsNumeric(CStr(ArrDg(zDg, sDg))) Then D = "0" Else D = "1": A1 = ArrDg(zDg, sDg) '1. Wert des AAE
                If Not IsNumeric(CStr(ArrDg(zDg, sDg + 1))) Then D = D + "0" Else D = D + "1": A2 = ArrDg(zDg, sDg + 1)
                If Not IsNumeric(CStr(ArrDg(zDg, sDg + 2))) Then D = D + "0" Else D = D + "1": a3 = ArrDg(zDg, sDg + 2)
                'Calc
                If D = "011" Then
                    A1 = a3 - A2: ArrDg(zDg, sDg + 0) = A1: .Cells(z1 - 1 + zDg, s1 - 1 + sDg) = A1
                ElseIf D = "101" Then
                    A2 = a3 - A1: ArrDg(zDg, sDg + 1) = A2: .Cells(z1 - 1 + zDg, s1 + 0 + sDg) = A2
                ElseIf D = "110" Then
                    'Der 3. D-Wert von AAE, der E-Wert in ArrDg(zDg, sDg3), ist hier "0",
                    'also NotNumeric, d.h.  ""  oder  "29,4x" (FixValue);
                    'falls G-Wert = ""       ---> Eintrag des berechneten Wertes A3
                    'falls G-Wert = "29,4x"  ---> Eintrag 29,4; rot
                    a3 = A1 + A2  '2 Vorwerte (Haltung, Schwierigkeit) werden addiert
                    T4_Write_OneSum_MayBeFixValue a3, zDg, sDg + 2 'regelt auch Vermerk
                End If
            Next
        Next
    'Finals
        End With
End Sub

Sub T4_CalcAndWrite_EEEG(B$, SpGN$, z1%, z2%, s1%)
    'Called from    T4_AutoCalc
    'B              = "iRWT-XXE-XXXE-XXXE-G-RMT-XXE-XXXE-XXXE-Gi" (Len(B)=41=AnzSpalten; Sy)
    'EEEG           E+E+E=G:      Pflicht + K1 + K2 = Gesamt
    '               = "8|13|18|20|@28|33|38|40|"    2 EEEG-Gruppen: 1. beginnt bei Dg-Spalte 8 (Ti), 2. bei 28 (Tu)
    'Action         Berechnet alle "Ges."-Werte (alle Competitors; m/w); trägt diese ein
    
    'Vorbereitung
        Dim A$, D$, E1$, E2$, E3$, E4$, EEEG$, Ges1$, GesAll$, WM$
        Dim i%, j%, RankNew%, RankOld%, sDg1%, sDg2%, sDg3%, sDg4%
        Dim sDgNameM%, sDgNameW%, sNam%, zDg%
        Dim A1#, A2#, a3#, a4#, Arr1() As String, Arr2() As String, r As Range
        With Sheets("T4")
    'EEEG
        EEEG = T4_Get_EEEG_ColumnNrs(B)           'EEEG = "" oder "8|13|18|20|@28|33|38|40|"
                    'B wurde reduziert  'B old = "iRWT-XXE-XXXE-XXXE-G-RMT-XXE-XXXE-XXXE-Gi"
                                        'B new = "iRWT-XXX-XXXX-XXXX-X-RMT-XXX-XXXX-XXXX-Xi"
        If EEEG = "" Then Exit Sub
    'sDgNameW, sDgNameM
        sDgNameW = CInt(ArrC(90))       'T4 Dg SpaltenNr Namen  weibl.
        sDgNameM = CInt(ArrC(91))       'T4 Dg SpaltenNr Namen  männl.
    'Action
        Arr2 = Split(EEEG, "@")         'Dg-SpaltenNummern (m@w)
        For j = 0 To UBound(Arr2)       'J=0 Damen, J=1 Herren
            'Nur Herren in M1N8
                If ArrDg(2, 1) = "M1N8" Then GoTo End_j
            EEEG = Arr2(j)              '"8|13|18|20|" bei J=0 Damen
            Arr1 = Split(EEEG, "|")     'Dg-SpaltenNummern (m|w)
            'Check pro Dg-Zeile, ob 3 der 4 Werte (EEEG) vorhanden sind
            For zDg = 6 To z2 - z1
                'Keine Überschrift als Name
                    If j = 0 And IsNumeric(sDgNameW) Then
                        If Not IsValidName(CStr(ArrDg(zDg, sDgNameW))) Then GoTo End_zDg
                    End If
                    If j = 1 And IsNumeric(sDgNameM) Then
                        If Not IsValidName(CStr(ArrDg(zDg, sDgNameM))) Then GoTo End_zDg
                    End If
                'Schleife über alle relevanten Dg-Zeilen
                A1 = 0: A2 = 0: a3 = 0: a4 = 0
                sDg1 = CInt(Arr1(0))      '=  8 = DgSpaltenPosition des ersten  "E" eines EEEG
                sDg2 = CInt(Arr1(1))      '= 13 = DgSpaltenPosition des zweiten "E" eines EEEG
                sDg3 = CInt(Arr1(2))      '= 18 = DgSpaltenPosition des dritten "E" eines EEEG
                sDg4 = CInt(Arr1(3))      '= 20 = DgSpaltenPosition des         "G" eines EEEG
                
                'Ist ein Wert in Spalte sDg1, sDg2, sDg3, sDg4 vorhanden oder nicht?
                'A1, A2, A3 dienen nur der A4-Berechnung, ggf. nicht der A1-, A2-, A3-Anzeige
                'Produce 0111 1011 1101 1110
                    E1 = CStr(ArrDg(zDg, sDg1)) 'E1 = "", "-", "0" oder "17,4"
                    E2 = CStr(ArrDg(zDg, sDg2)): E3 = CStr(ArrDg(zDg, sDg3)): E4 = CStr(ArrDg(zDg, sDg4))
                    If E1 = "" Then D = "0" Else D = "1"
                    If E2 = "" Then D = D + "0" Else D = D + "1"
                    If E3 = "" Then D = D + "0" Else D = D + "1"
                    If E4 = "" Or E4 Like "*x*" Then D = D + "0" Else D = D + "1"
                'A1 A2 A3 numeric
                    If E1 = "-" Then
                        A1 = 0: Set r = .Cells(z1 - 1 + zDg, s1 - 1 + sDg1)
                        If j = 0 Then r.Font.Color = 14083324 Else r.Font.Color = 15652797 '(m-Farbe, hellblau)
                        ElseIf E1 <> "" Then A1 = ArrDg(zDg, sDg1)
                    End If
                    If E2 = "-" Then
                        A2 = 0: Set r = .Cells(z1 - 1 + zDg, s1 - 1 + sDg2)
                        If j = 0 Then r.Font.Color = 14083324 Else r.Font.Color = 15652797 '(m-Farbe, hellblau)
                        ElseIf E2 <> "" Then A2 = ArrDg(zDg, sDg2)
                    End If
                    If E3 = "-" Then
                        a3 = 0: Set r = .Cells(z1 - 1 + zDg, s1 - 1 + sDg3)
                        If j = 0 Then r.Font.Color = 14083324 Else r.Font.Color = 15652797 '(m-Farbe, hellblau)
                        ElseIf E3 <> "" Then a3 = ArrDg(zDg, sDg3)
                    End If
                'Calc       Eine Berechnung kann erfolgen, falls genau einer der 4 Werte nicht vorhanden ist
                    If D = "0111" Then
                                            A1 = a4 - a3 - A2
                                            ArrDg(zDg, sDg1) = A1: A = A1
                                            If E1 = "-" Then A = "-"
                                            .Cells(z1 - 1 + zDg, s1 + sDg1 - 1) = A
                    ElseIf D = "1011" Then
                                            A2 = a4 - a3 - A1
                                            ArrDg(zDg, sDg2) = A2: A = A2
                                            If E2 = "-" Then A = "-"
                                            .Cells(z1 - 1 + zDg, s1 + sDg2 - 1) = A
                    ElseIf D = "1101" Then
                                            a3 = a4 - A1 - A2
                                            ArrDg(zDg, sDg3) = a3: A = a3
                                            If a3 < 0.001 Then A = "" '-3,55...E-15 --> 0
                                            If E3 = "-" Then A = "-"
                                            .Cells(z1 - 1 + zDg, s1 + sDg3 - 1) = A
                    ElseIf D = "1110" Or D = "1100" Then 'ggf. keine Kür2
                        'Der 4. Wert, der G-Wert in ArrDg(zDg, sDg4), ist hier "0",
                        '   also not numeric, d.h.  ""  oder  "29,4x" (FixValue);
                        '   falls G-Wert = ""       ---> Eintrag des berechneten Wertes A4
                        '   falls G-Wert = "29,4x"  ---> Eintrag 29,4; rot
                            a4 = A1 + A2 + a3 '3 Vorwerte (Pflicht, K1, K2) werden addiert
                            T4_Write_OneSum_MayBeFixValue a4, zDg, sDg4 'regelt auch Vermerk
                    End If
                'Check_EqualRank
                    If j = 0 Then WM = "W"
                    If j = 1 And WM = "W" Then WM = "M"
                    Ges1 = "": GesAll = "": RankOld = 0: RankNew = 0
                    'G-Werte der Damen, dann Herren
                    If CStr(ArrDg(zDg, sDg4)) <> "" Then
                        'If CStr(ArrDg(zDg, sDg4)) = "61,4" Then Stop
                        Ges1 = CStr(ArrDg(zDg, sDg4)) '"72,10"
                        sNam = CInt(Mid(SpGN, InStr(1, SpGN, CStr(sDg4)) + 3, 2))
                        RankNew = CInt(ArrDg(zDg, sNam - 1))
                        If InStr(1, GesAll, "|" + Ges1 + "|") > 0 Then
                            'Gesamtwert = Gesamtwert der Zeile zuvor
                            If RankNew <> RankOld Then
                                ArrDg(zDg, sNam - 1) = RankOld
                                .Cells(z1 - 1 + zDg, s1 + sNam - 2) = RankOld
                                'showArray2D ArrDg
                            End If
                        End If
                        GesAll = GesAll + "|" + Ges1 + "|"
                        RankOld = RankNew
                    End If
End_zDg:
            Next zDg
End_j:
        Next j
        'B enthält keine "EEEG" mehr         B = "iRNTWXXXWXXXXWXXXXWXWRNTWXXXWXXXXWXXXXWXi"
    'Finals
        End With
End Sub

Sub T4_Set_ArrowsForFixValues(B$, z1%, z2%, s1%)
    'Called from    T4_AutoCalc
    
    'Vorbereitung
        Dim F$, i%, sDg%, sGes1%, sGes2%, Sp%, zDg%, Ze%, Arr1() As String, r As Range
        With Sheets("T4")
    'FixValues to Array
        F = ArrDg(1, 4)                     'F = "FixValues: (17|32)(23|32)(27|32)(28|32)(29|32)"
        If F = "FixValues: " Then Exit Sub  'nothing to do
        F = Replace(F, "FixValues: ", "")   '(17|32)(23|32)(27|32)(28|32)(29|32)
        F = Mid(Replace(F, ")(", "|"), 2)   '17|32|23|32|27|32|28|32|29|32)
        F = Left(F, Len(F) - 1)             '17|32|23|32|27|32|28|32|29|32
        Arr1 = Split(F, "|")
    'Pfeile löschen     'rechts einer Ges.-Spalte
        'Pfeile in ArrDg suchen
        For Sp = 1 To UBound(ArrDg, 2)
            If Mid(B, Sp, 1) = "-" Or Mid(B, Sp, 1) = "i" Then
                For Ze = 6 To UBound(ArrDg, 1) - 1
                    If ArrDg(Ze, Sp) = "!" Then
                        ArrDg(Ze, Sp) = "": Set r = .Cells(z1 + Ze - 1, s1 - 1 + Sp)
                        'Einen Pfeil löschen
                            r.Value = "": r.Font.NAME = "Calibri"
                        'Wert vorn rot auf schwarz setzen
                            Set r = .Cells(z1 + Ze - 1, s1 - 2 + Sp)
                            r.Font.Color = vbBlack
                    End If
                Next
            End If
        Next
    'Pfeil anzeigen
        'FixValues können in E- und G-Spalten auftreten
        'Pfeile werden rechts eines FixValue vermerkt
        For i = 0 To UBound(Arr1) - 1 Step 2    '17|32|23|32|27|32|28|32|29|32
            zDg = Arr1(i): sDg = Arr1(i + 1)    'Dg-Zelle eines FixValue (rot)
            'Eintrag "!" (Pfeil) in ArrDg
                If ArrDg(zDg, sDg + 1) = "" Then ArrDg(zDg, sDg + 1) = "!"
            'Eintrag "!" (Pfeil) in T4
                Set r = .Cells(z1 - 1 + zDg, s1 + sDg)
                If r.Value = "" Then
                    r.Font.NAME = "Wingdings 3": r.Value = "!"
                    r.HorizontalAlignment = xlRight: r.VerticalAlignment = xlCenter
                End If
            'Einfärben des Wertes in T4
                Set r = .Cells(z1 - 1 + zDg, s1 + sDg - 1): r.Font.Color = vbRed
        Next
    'Finals
        End With
End Sub

Sub T4_CalcAndWrite_EEG(B$, SpGN$, z1%, z2%, s1%)
    'Called from    T4_AutoCalc
    'B              = "iRWT-XXE-XXXE-G-RMT-XXE-XXXE-Gi"     '2 EEG-Gruppen, keine EEEG
    'EEG            E + E = G:      Pflicht + K1 = Gesamt
    '               = "8|13|15|@23|28|30|"    1. EEG-Gruppe beginnt bei Dg-Spalte 8 (Ti), 2. bei 23 (Tu)
    'Action         Ermittlung von EEG, Write G, Check_EqualRank, Check_CorrectSum
    
    'Vorbereitung
        Dim D$, EEG$, Ges1$, GesAll$, sNam$, WM$
        Dim i%, j%, RankNew$, RankOld%, sDg1%, sDg2%, sDg3%, sDg4%, zDg%, sRank%
        Dim A1#, A2#, a3#, Arr1() As String, Arr2() As String
    'EEG
        EEG = T4_Get_EEG_ColumnNrs(B)             'EEG = "" oder "8|13|15|@23|28|30|"
                    'B wurde reduziert  'B old = "iRWT-XXE-XXXE-G-RMT-XXE-XXXE-Gi"
                                        'B new = "iRWT-XXX-XXXX-X-RMT-XXX-XXXX-Xi"
        If EEG = "" Then Exit Sub
        With Sheets("T4")
    'Action
        Arr2 = Split(EEG, "@")         'Dg-SpaltenNummern (m@w)
        For j = 0 To UBound(Arr2)       'j=0 Damen, j=1 Herren
            EEG = Arr2(j)               '"8|13|15|" bei j=0 Damen; (zDg,8)+(zDg,13)=(zDg,15)
            Arr1 = Split(EEG, "|")      'Dg-SpaltenNummern (m|w)
            'Check pro Dg-Zeile, ob 3 der 4 Werte (EEG) vorhanden sind
            For zDg = 6 To z2 - z1
                'Schleife über alle relevanten Dg-Zeilen
                    A1 = 0: A2 = 0: a3 = 0:
                'DgSpaltenPositionen
                    sDg1 = CInt(Arr1(0))      '=  8 = DgSpaltenPosition des ersten  "E" eines EEG
                    sDg2 = CInt(Arr1(1))      '= 13 = DgSpaltenPosition des zweiten "E" eines EEG
                    sDg3 = CInt(Arr1(2))      '= 15 = DgSpaltenPosition des         "G" eines EEG
                'Ist ein Wert in Spalte sDg1, sDg2, sDg3 vorhanden ("1") oder nicht ("0")?
                    'D = "011", "101", "110" oder "111"
                    If Not IsNumeric(CStr(ArrDg(zDg, sDg1))) Then D = "0" Else D = "1":         A1 = ArrDg(zDg, sDg1)
                    If Not IsNumeric(CStr(ArrDg(zDg, sDg2))) Then D = D + "0" Else D = D + "1": A2 = ArrDg(zDg, sDg2)
                    If Not IsNumeric(CStr(ArrDg(zDg, sDg3))) Then D = D + "0" Else D = D + "1": a3 = ArrDg(zDg, sDg3)
                'Calc       Eine Berechnung kann erfolgen,
                '           falls genau einer der 3 Werte (A1, A2, A3) nicht vorhanden ist
                    If D = "011" Then
                                          A1 = a3 - A2: ArrDg(zDg, sDg1) = A1: .Cells(z1 - 1 + zDg, s1 + sDg1 - 1) = A1
                    ElseIf D = "101" Then A2 = a3 - A1: ArrDg(zDg, sDg2) = A2: .Cells(z1 - 1 + zDg, s1 + sDg2 - 1) = A2
                    ElseIf D = "110" Then
                        'Der 3. Wert, der G-Wert in ArrDg(zDg, sDg3), ist hier "0",
                        '   also not numeric, d.h.  ""  oder  "29,4x" (FixValue);
                        '   falls G-Wert = ""       ---> Eintrag des berechneten Wertes A3
                        '   falls G-Wert = "29,4x"  ---> Eintrag 29,4; rot; Vermerk
                            a3 = A1 + A2   '2 Vorwerte (Pflicht, K1) werden addiert
                            T4_Write_OneSum_MayBeFixValue a3, zDg, sDg3 'regelt auch Vermerk
                    End If
                    
'                'Check_EqualRank
'                    If j = 0 Then WM = "W"
'                    If j = 1 And WM = "W" Then WM = "M"
'                    Ges1 = "": GesAll = "": RankOld = 0: RankNew = 0
'                    'G-Werte der Damen, dann Herren
'                    If CStr(ArrDg(zDg, sDg3)) <> "" Then
'                        Ges1 = CStr(ArrDg(zDg, sDg3))
'                        '    = "72,10"; Gesamtwert des Aktiven in DgZeile zDg
'                        sNam = CInt(Mid(SpGN, InStr(1, SpGN, CStr(sDg3)) + 3, 2))
'                        '    = SpaltenNr der NamenSpalte
'                            'SpGN = "G16N03,G32N19"
'                            '     = je DgSpaltenNr der Ges.-Spalte/zugehörige Namen-Spalte
'                            '     = "G12N04" bei Buli; G12 = Der Wert für "Ges." steht in Spalte 12
'                        If ArrDg(5, sNam - 1) = "R" Then sRank = sNam - 1
'                        If ArrDg(5, sNam - 2) = "R" Then sRank = sNam - 2
'                        RankNew = CInt(ArrDg(zDg, sRank))
'                        If InStr(1, GesAll, "|" + Ges1 + "|") > 0 Then
'                            'Gesamtwert = Gesamtwert der Zeile zuvor
'                            If RankNew <> RankOld Then
'                                'Rang ins Array schreiben
'                                    ArrDg(zDg, sRank) = RankOld
'                                'Rang ins Dg schreiben
'                                    .Cells(z1 - 1 + zDg, s1 - 1 + sRank) = RankOld
'                            End If
'                        End If
'                        GesAll = GesAll + "|" + Ges1 + "|"
'                        RankOld = RankNew
'                    End If
                    
                'Check_CorrectSum
                    If a3 > 0 And CStr(Round(A1 + A2, 2)) <> CStr(Round(a3, 2)) Then
                        'a3 (= G-Wert) stimmt nicht mit der Summe der Vorwerte überein
                        If CStr(A1 + A2) <> "0" Then
                             'Pfeil anzeigen
                            With .Cells(z1 - 1 + zDg, s1 + sDg3)
                                .Font.NAME = "Wingdings 3": .Value = "!"
                                .HorizontalAlignment = xlRight
                                .VerticalAlignment = xlCenter
                            End With
                            ArrDg(zDg, sDg3 + 1) = "!"
                        End If
                    Else
                        'a3 (= G-Wert) stimmt mit der Summe der Vorwerte überein
                        .Cells(z1 - 1 + zDg, s1 + sDg3).Value = "" 'keinen Pfeil anzeigen
                        ArrDg(zDg, sDg3 + 1) = ""
                    End If
            Next
        Next
    'Finals
        End With
End Sub

Function T4_Get_EEEG_ColumnNrs(B$) As String
    'Called from    T4_AutoCalc
    'B              = "iRWT-XXE-XXXE-XXXE-G-RMT-XXE-XXXE-XXXE-Gi"   Sy PK1K2
    
    'Exit
        If Not B Like "*E*E*E*G*" Then Exit Function
    'Vorbereitung
        Dim B1$, s$, C1%, C2%, i%
    'Exit
        C1 = InStr(1, B, "G")
        B1 = Left(B, C1)        '= "iRWT-XXE-XXXE-XXXE-G"
        If Not B1 Like "*E*E*E*G" Then Exit Function 'kein Sy-Wettkampf mit 2 Küren
    'Action
        For i = 1 To 99
            C1 = InStr(1, B, "E")
            C2 = InStr(1, B, "G")
            If C1 > 0 And C1 < C2 Then
                'Es existiert ein E und weiter rechts ein G
                s = s + CStr(C1) + "|"
                B = Replace(B, "E", "X", 1, 1)
            Else
                s = s + CStr(C2) + "|"
                B = Replace(B, "G", "X", 1, 1)
                If InStr(1, B, "E") = 0 Then Exit For Else s = s + "@"
            End If
        Next
    'Finals
        T4_Get_EEEG_ColumnNrs = s
        'show s: Stop   's = "8|13|18|20|@28|33|38|40|"
        'show B: Stop   'B = "iRWT-XXX-XXXX-XXXX-X-RMT-XXX-XXXX-XXXX-Xi"
End Function

Function T4_Get_EEG_ColumnNrs(B$) As String
    'Called from    T4_CalcAndWrite_EEG
    'B              = "iRWT-XXE-XXXE-G-RMT-XXE-XXXE-Gi"     '2 EEG-Gruppen, keine EEEG
    'EEG            E + E = G:      Pflicht + K1 = Gesamt
    '               = "8|13|15|@23|28|30|"    1. EEG-Gruppe beginnt bei Dg-Spalte 8 (Ti), 2. bei 23 (Tu)
    
    If Not B Like "*E*E*G*" Then Exit Function
    Dim s$, C1%, C2%, i%
    For i = 1 To 99
        C1 = InStr(1, B, "E")
        C2 = InStr(1, B, "G")
        If C1 > 0 And C1 < C2 Then
            'Es existiert ein E und weiter rechts ein G
            s = s + CStr(C1) + "|"
            B = Replace(B, "E", "X", 1, 1)
        Else
            s = s + CStr(C2) + "|"
            B = Replace(B, "G", "X", 1, 1)
            If InStr(1, B, "E") = 0 Then Exit For Else s = s + "@"
        End If
    Next
    'MN4: Best-5-Teams-Final: nur EGEG, kein EEG
        If anzAinB("|", s) = 2 Or anzAinB("|", s) = 4 Then s = ""
    T4_Get_EEG_ColumnNrs = s
End Function


Function T4_Get_AAAE_ColumnNrs(B$) As String
    'Called from    T4_CalcAndWrite_AAAE
    'B              = "iRWT-AAE-AAAE-G-RMT-AAE-AAAE-Gi" Sy PK1
    
    If Not B Like "*AAA*E*" Then Exit Function
    Dim s$, c%, i%
    For i = 1 To 99
        c = InStr(1, B, "AAAE")
        If c > 0 Then
            s = s + CStr(c) + "|"
            B = Replace(B, "AAAE", "XXXE", 1, 1)
        Else
            Exit For
        End If
    Next
    
    For i = 1 To 99
        c = InStr(1, B, "AAA-E")
        If c > 0 Then
            s = s + CStr(c) + "|"
            B = Replace(B, "AAA-E", "XXX-E", 1, 1)
        Else
            Exit For
        End If
    Next
    T4_Get_AAAE_ColumnNrs = s
End Function

Function T4_Get_AAE_ColumnNrs(B$) As String
    'Called from    T4_CalcAndWrite_AAE
    'B              = "iRWT-AAE-XXXE-G-RMT-AAE-XXXE-Gi" Sy PK1
    
    If Not B Like "*AAE*" Then Exit Function
    Dim s$, c%, i%
    For i = 1 To 99
        c = InStr(1, B, "AAE")
        If c > 0 Then
            s = s + CStr(c) + "|"
            B = Replace(B, "AAE", "XXE", 1, 1)
        Else
            Exit For
        End If
    Next
    T4_Get_AAE_ColumnNrs = s
End Function

Function T4_Get_DgFolderName_GiveDgTitle(DgTitle$) As String
    'Called from    T4_ActionsOnDgClickFirst
    '               T8_Synchronize_T8SomeDgData_With_T4ListOfDesignTitles
    
    'Vorbereitung
        Dim F$, T$, z%, s%
    'Get (zSh, sSh) of T4-DgTitle
        T = Trim(DgTitle)
        z = Get_RowNr_HoldingMyTextWhole("T4", T)      'LookAt:=xlWhole
        s = Get_ColumnNr_HoldingMyTextWhole("T4", T)
    'Exit       DgTitle nicht gefunden
        If z = 0 Then Stop 'show txt1 + "Ein Design mit dem Titel '" + T + "' existiert in T4 nicht.": Exit Function
    'DgTitle gefunden
        'DgFolderName steht 1 Zelle über DgTitle
        F = Sheets("T4").Cells(z - 1, s)
        If F <> "" Then T4_Get_DgFolderName_GiveDgTitle = F Else Stop
End Function


