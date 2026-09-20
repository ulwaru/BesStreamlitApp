Attribute VB_Name = "M45"
Option Explicit 'M45

'Beep
    Private Declare Function Beep Lib "kernel32" (Optional ByVal dwFreq As Long = 900, Optional ByVal dwDuration As Long = 100) As Long

Sub zzz_M45()

    showProcs "bad link"
        
    'RenameModule "Modul1", "M41"
    'Beep 400, 666
    EE 1: Beep
End Sub

Sub T4_Write_ClubToT5()
    'Called from    T4_ActionsOnDgClickNext
    'Action         Fills T5ClubCell of actual competitor if T5ClubCell is empty

    'Vorbereitung
        Dim BCD$, Club$, VnNn$
        Dim c%, s1%, sDg%, sDgNameM%, sDgNameW%, sDgNameX%, z1%, zDg%, zT5%
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39))
    'Check Dg-Wechsel
            If ArrC(46) <> ArrC(47) Then Exit Sub 'Last1Click was not inside same Dg
    'Get LastCell (DgZelle, die gerade verlassen wurde)
        zDg = ArrC(106) - z1 + 1: sDg = ArrC(107) - s1 + 1 'zShLast2S 'sShLast2S
            If zDg < 1 Then Exit Sub
            If sDg < 1 Then Exit Sub
            If ArrDg(zDg, sDg) = "" Then Exit Sub
    'sDg = ClubSpalte?
        BCD = CStr(ArrDg(5, sDg))
        'In ArrDg, Zeile 5, stehen in den Spalten EinzelBuchstaben, z. B.:
        '"B" für Verein (Aktive weiblich), "C" für Verein (Aktive männlich), "D" für Verein (Aktive mixed)
        
        If "BCD" Like "*" + BCD + "*" Then
            If zDg > 5 Then
                'User verließ gerade die DgClubZelle (zDg, sDg)
                Sheets("T4").Cells(CInt(ArrC(106)), CInt(ArrC(107)) - 1).Select
                Sheets("T4").Cells(CInt(ArrC(106)), CInt(ArrC(107)) - 1).Select
            End If
        End If
        
'        If BCD = "B" Then
'            '"B" für Verein (Aktive weiblich)
'            Club = CStr(ArrDg(zDg, sDg))
'            sDgNameW = CInt(ArrC(90))    'T4 Dg SpaltenNr Namen weibl.
'            VnNn = CStr(ArrDg(zDg, sDgNameW))
'            c = InStr(1, VnNn, " ("): If c > 0 Then VnNn = Left(VnNn, c - 1)
'            zT5 = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, VnNn)
'            If Sheets("T5").Cells(zT5, 7) = "" And Club <> "" Then Sheets("T5").Cells(zT5, 7) = Club
'         ElseIf BCD = "C" Then
'            '"C" für Verein (Aktive männlich)
'            Club = CStr(ArrDg(zDg, sDg))
'            sDgNameM = CInt(ArrC(91))    'T4 Dg SpaltenNr Namen männlich
'            VnNn = CStr(ArrDg(zDg, sDgNameM))
'            c = InStr(1, VnNn, " ("): If c > 0 Then VnNn = Left(VnNn, c - 1)
'            zT5 = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, VnNn)
'            If Sheets("T5").Cells(zT5, 7) = "" And Club <> "" Then Sheets("T5").Cells(zT5, 7) = Club
'         ElseIf BCD = "D" Then
'            '"C" für Verein (Aktive mixed)
'            Club = CStr(ArrDg(zDg, sDg))
'            sDgNameX = CInt(ArrC(91))    'T4 Dg SpaltenNr Namen mixed
'            VnNn = CStr(ArrDg(zDg, sDgNameX))
'            c = InStr(1, VnNn, " ("): If c > 0 Then VnNn = Left(VnNn, c - 1)
'            zT5 = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, VnNn)
'            If Sheets("T5").Cells(zT5, 7) = "" And Club <> "" Then Sheets("T5").Cells(zT5, 7) = Club
'        End If
End Sub

Sub T4_Write_NationToT5()
    'Called from    T4_ActionsOnDgClickNext
    'Action         Fills T5NationCell of actual competitor if T5NationCell is empty
    
    'Vorbereitung
        Dim Nat$, c%, s1%, sDg%, z1%, zDg%
        Dim sDgNameM%, sDgNameW%, sDgNameX%, sDgNationM%, sDgNationW%, sDgNationX%
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39))
        sDgNameW = CInt(ArrC(90))    'T4 Dg SpaltenNr Namen  weibl.
        sDgNameM = CInt(ArrC(91))    'T4 Dg SpaltenNr Namen  männl.
        sDgNameX = CInt(ArrC(92))    'T4 Dg SpaltenNr Namen  mix
        sDgNationW = CInt(ArrC(93))  'T4 Dg SpaltenNr Nation weibl.
        sDgNationM = CInt(ArrC(94))  'T4 Dg SpaltenNr Nation männl.
        sDgNationX = CInt(ArrC(95))  'T4 Dg SpaltenNr Nation mix
    'Check Dg-Wechsel
            If ArrC(46) <> ArrC(47) Then Exit Sub 'Last1Click was not inside same Dg
    'Check if NationCell exists
        If sDgNationW + sDgNationM + sDgNationX = 0 Then Exit Sub
    'Get LastCell (DgZelle, die gerade verlassen wurde)
        zDg = ArrC(106) - z1 + 1: sDg = ArrC(107) - s1 + 1 'zShLast2S 'sShLast2S
            If zDg < 0 Then Exit Sub
            If sDg < 0 Then Exit Sub
            If ArrDg(zDg, sDg) = "" Then Exit Sub
    'Check if LastCell was a NationCell
        'showArray2D ArrDg: Stop
        If sDg = sDgNationW Then
            If ArrDg(zDg, sDgNameW) = "" Then Exit Sub Else T4_DgClick_CheckOneName_LastSelected zDg, sDgNameW
        ElseIf sDg = sDgNationM Then
            If ArrDg(zDg, sDgNameM) = "" Then Exit Sub Else T4_DgClick_CheckOneName_LastSelected zDg, sDgNameM
        ElseIf sDg = sDgNationX Then
            If ArrDg(zDg, sDgNameX) = "" Then Exit Sub Else T4_DgClick_CheckOneName_LastSelected zDg, sDgNameX
        End If
End Sub


Sub T4_DgClick_CheckOneName_LastSelected(Optional zDg%, Optional sDg%)
    'Called from    T4_ActionsOnDgClickNext
    'zDg, sDg       ArrDg(zDg, sDg) = CompleteDgNameCellContent
    'Status         Ein Dg ist aktiv, ArrDg existiert;
    '               OnEveryDgClick: Sub wird nur abgearbeitet, falls gerade
    '               eine CompetitorNameCell verlassen wurde (LastSelectedCell);
    'Action         CompetitorName komplett versorgen;
    '               N = CompetitorNameCell-Inhalt = "Lea Mai" oder "Lea Mai (15)"
    '               (1) Zusatzangaben in NameCell herausfiltern und abarbeiten
    '                   N = "Lea Mai", ggf. mit (15) j=1978 w o=Bonn s=May n=Lea May
    '                       (15)        Age; wird ignoriert; gem T5-Jhg berechnet
    '                       j=1978      Jahrgang wird in T5-Jhg geschrieben
    '                       o=Bonn      Ort/Club wird in T5-Club geschrieben
    '                       s=May       Schreibweise wird in T5 hinzugefügt
    '                       n=Lea May   globale Ersetzung des Namens; "Lea May" statt "Lea Mai"
    '                                   in allen Dgs, T4SomeDgData, T5, Foldername
    '               (2) N new --> PersonalFolder in 'Leute' erstellen;
    '                         --> neue Zeile in T5 anlegen (PersonData);
    '               (3) N old --> ggf. T5 ergänzen (Jhg, m/w, Verein, Nation, Schreibweise);
    '                         --> ggf. NameOfPersonalFolder ändern (in T5 und 'Leute')
    '               (4) Dg    --> ggf. Age, Club, Nation eintragen
    '               (5) legt pro Competitor einen Link seines PersonenOrdners in den EventOrdner;
    '               (6) update Names in T4SomeDgData
    
    'Vorbereitung
        'Ticks "T4_DgClick_CheckNames Start"
        Dim Age$, Club$, Jhg$, mw$, N$, NameOfPersonalFolder$, Nat$, NewName$
        Dim Nn$, PathOfEventFolder$, PathOfPersonalFolder$, sw$, Title$, Vn$
        Dim s1%, s2%, z1%, zT5%
        Title = ArrDg(2, 2): With Sheets("T4")
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39)) '(z1,s1) = DgEcke li ob
        s2 = CInt(ArrC(40))
    'ArrDg 'wird bei jedem Klick in ein Dg neu geladen
        'ArrDg wurde gerade geladen (mit 'FillArrDg' in T4_ActionsOnDgClickNext)
    'Exit
        If Not T4_zDg_sDg_IsValidNameCell(zDg, sDg, z1, s1) Then Exit Sub
    'Status:    Last2SelectedCell = valid PersonNameCell (zDg, sDg)
                'Data benötigt:
                    'für T5             |Nn|Vn|m/w|Jhg|Club|LTV|Nation|NamePersonenOrdner|
                    'für T4SomeDgData   #Nn|Vn|zDg|sDg|ESyM|Rang
                'Data vorhanden:        ValidNameCell(zDg, sDg)
        N = ArrDg(zDg, sDg) '= gesamte UserEingabe, CompleteDgNameCellContent
    'Get Items from actual DgNameCellContent only   'mw, Club, D from Dg
        T4_Get_Items_FromValidDgNameCellOnly N, zDg, sDg, Nn, Vn, Age, Jhg, mw, Club, Nat, sw, NewName
    'mw
        If mw <> "" And Nn <> "" And Vn <> "" Then T5_Set_mw mw, Vn + " " + Nn
        If mw = "" And Nn <> "" And Vn <> "" Then mw = T5_Get_mw_Give_VnNn(Vn + " " + Nn)
        If mw = "" Then mw = T4_Get_mw_viaUF6(Vn, Nn): T5_Set_mw mw, Vn + " " + Nn
        'Zeile einfärben falls nur 1 Personenspalte (Mix)
            If Mid(ArrDg(2, 1), 2, 1) = "1" Then
                Dim r As Range: EE 2: Set r = .Range(.Cells(z1 - 1 + zDg, s1 + 2), .Cells(z1 - 1 + zDg, s2 - 1))
                If mw = "w" Then r.Interior.Color = 14083324 Else r.Interior.Color = 15652797 'hellblau/Blau1 (m-Farbe)
                With r.Borders: .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = 2: End With: EE 3
            End If
    'T5 Create new line for actual competitor if necessary
        T4_Create_NewLineInT5_ifNecessary zT5, zDg, sDg, Nn, Vn, mw, NewName 'fills zT5
    'T5 Add content to zT5-Line
        T4_Add_SomeContent_ToOneLineInT5 Nn, Vn, zT5, Jhg, Club, Nat, sw
        T5_Check_LTV_Nat_ViaClubCombination Club
    'Write to Dg: Vn Nn Age
        T4_Write_NewDgNameCellContent Nn, Vn, Jhg, Age, zT5, zDg, sDg 'Fills Jhg, Age
    'Write to Dg: Nation
        T4_Write_Nation_toDg Nat, zT5, zDg, sDg
    'Write to Dg: Club
        T4_Write_Club_toDg Club, zT5, zDg, sDg
Jump:  'jump from 3 lines down
    'Change CompetitorMainFolder
        T4_CreateOrChange_PersonalFolder zT5, Nn, Vn, NewName, Nat, Club, NameOfPersonalFolder, PathOfPersonalFolder
        If NameOfPersonalFolder = "" Then Stop: GoTo Jump: Stop 'time to repair
    'Create PersonLink inside EventFolder
        Create_OneVbsLink_Person ArrC(57) + "\" + NameOfPersonalFolder
    'T4SomeDgData
        T4_Change_T4SomeDgData_OneNameData Nn, Vn, zDg, sDg
    'NewName-Action (global change)
        T4_NewName_Action Vn, Nn, NewName
    'Finals
        End With
End Sub

Function T4_Get_mw_viaUF6(Vn$, Nn$) As String
    'm/w abfragen
        Dim mw$, frm As UF6: Set frm = New UF6
        frm.Label1.Caption = "Bitte geben sie für " + Vn + " " + Nn + " an:"
        frm.show: mw = frm.Auswahl
        T4_Get_mw_viaUF6 = mw '"", "w", "m"
        Unload frm: Set frm = Nothing
    'm/w eintragen
        T5_Set_mw mw, Vn + " " + Nn
End Function

Sub T4_Get_Items_FromValidDgNameCellOnly(N$, zDg%, sDg%, Nn$, Vn$, Age$, Jhg$, mw$, Club$, Nat$, sw$, NewName$)
    'Called from    T4_DgClick_CheckOneName_LastSelected
    
    'HelperBox-Zeile?
        If anzAinB("|", N) = 7 Then
            'N = HelperBox-Zeile --> nothing to cut
            Dim Arr1() As String
            Arr1 = Split(N, "|"): Vn = Arr1(0): Nn = Arr1(1): Nat = Arr1(2)
            Club = Arr1(3): Jhg = Arr1(4): mw = Arr1(5): N = ""
        Else
            T4_Cut_ItemsN N, zDg, sDg, Nn, Vn, Age, Jhg, mw, Club, Nat, sw, NewName
        End If
End Sub

Sub T4_Cut_ItemsN(N$, zDg%, sDg%, Nn$, Vn$, Age$, Jhg$, mw$, Club$, Nat$, sw$, NewName$)
    'Called from    T4_Get_Items_FromValidDgNameCellOnly
    'N
    'Dg(zDg, sDg)   Last2SelectedCell = valid CompetitorNameCell
    'Cut            Nn Vn Age j= x= c= o= v= s= n=
    'GetFromDg      Club Nat mw

    'Vorbereitung
        Dim NnVn$, C1%
    'Age                    " (17)" aus N rausnehmen bei "Lea Mai (17)"
        Age = T4_Cut_OutOfDgNameCellContent_Age(N)    '"" oder "17"
    'Jhg                    "75", "j=" aus N rausnehmen bei z. B. "Maier75" oder " j=75"
        Jhg = T4_Cut_OutOfDgNameCellContent_j_Jhg(N)    '"" oder "1975"
    'mw                     Einzelnes "w" bzw. "m" aus N rausnehmen
        mw = T4_Cut_OutOfDgNameCellContent_x_mw(N)
        If mw = "" Then mw = T4_Get_mw_FromDg(sDg)
    'Nn, Vn                 (N = "Lea Mai", "Lea Mai (15)", "Lea Mai j=78 o=Bonn s=May")
        NnVn = T4_Cut_OutOfDgNameCellContent_VnNn(N)
        C1 = InStr(1, NnVn, "|"): Nn = Left(NnVn, C1 - 1): Vn = Mid(NnVn, C1 + 1)
    'Club
        Club = T4_Cut_OutOfDgNameCellContent_c_Club(N)          'c= o= v=
        If Club = "" Then Club = T4_Get_Club_FromDgClubColumn(zDg, sDg)
    'Nat
        Nat = T4_Get_Nation_FromDgNationColumn(zDg, sDg)
        If Nat = "" And isGermanEvent(CStr(ArrDg(2, 2))) Then Nat = "D"
    'Sw
        sw = T4_Cut_OutOfDgNameCellContent_s_Spelling(N)     's= (Schreibweise)
    'NewName
        NewName = T4_Cut_OutOfDgNameCellContent_n_NewNameGlobal(N) 'n= 'NewName-Action: s. unten
    'N sollte jetzt leer sein
End Sub

Function T4_zDg_sDg_IsValidNameCell(zDg%, sDg%, z1%, s1%) As Boolean
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'zDg, sDg       Cell inside ActualDg
        
    'Exit
        If zDg = 0 Then
            'Dg-Wechsel?
                If ArrC(46) <> ArrC(47) Then Exit Function 'Last1Click was not inside same Dg
            'Last2Click inside ActualDg
                zDg = ArrC(106) - z1 + 1 '(zShLast2S ..
                sDg = ArrC(107) - s1 + 1 '.., sShLast2S) = DgZelle, die gerade verlassen wurde
        End If
        If zDg < 6 Then Exit Function
        If sDg < 1 Then Exit Function
        'DoArrDg
        If Not "WMX" Like "*" + CStr(ArrDg(5, sDg)) + "*" Then Exit Function
        If ArrDg(zDg, sDg) = "" Then Exit Function
        If Not IsValidName(CStr(ArrDg(zDg, sDg))) Then Exit Function
    'Finals
        T4_zDg_sDg_IsValidNameCell = True
End Function

Sub T4_Change_T4SomeDgData_OneNameData(Nn$, Vn$, zDg%, sDg%)
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'T4SomeDgData   012 Names: *|#|Vn Nn|zDg|sDg|ESyM|Rang|*
    'Dg(zDg, sDg)   = LastSelectedCell; actual competitor
    'Action         Calc   ESyM, Rang
    '               Write  T4DgNameCellContent
    '               Change T4SomeDgData
    'Vorbereitung
        Dim ESyM$, names$
        Dim Rang$, ActualNameData$, Title$
        Dim C1%, i%, zT4%
        Title = ArrDg(2, 2)
    'ESyM
        ESyM = T4_Get_ESyM
    'Rang
        If "IJK" Like "*" + ArrDg(5, sDg - 1) + "*" Then Rang = CStr(ArrDg(zDg, sDg - 1))
        If "IJK" Like "*" + ArrDg(5, sDg - 2) + "*" Then Rang = CStr(ArrDg(zDg, sDg - 2))
    'T4SomeDgData   'Names-Zelle enthält je *#|Lea Mai|6|3|Einz|1|*
        'T4 ActualNameData
            ActualNameData = "#|" + Vn + " " + Nn + "|" + CStr(zDg) + "|" _
               + CStr(sDg) + "|" + ESyM + "|" + Rang + "|"
        'zT4 = ZeilenNr EventTitle in T4Spalte5 (= 1. Spalte von T4SomeDgData)
            zT4 = Get_RowNr_HoldingMyTextWholeInColumnX("T4", 5, 7, Title) 'Sheet,Col,After,Txt
            
            
            
            If zT4 = 0 Then Stop 'sollte nicht vorkommen
            'T4SomeDgData neu?
            
            
        'Names      (Content of NamesCell inside actual line of T4SomeDgData)
            With Sheets("T4"): names = .Cells(zT4, 12)
            C1 = InStr(1, names, "#|" + Vn + " " + Nn + "|")   'c1 = Start of ActualName
            If C1 = 0 Then
                'Name war bisher nicht in der Namensliste
                'Anfangsnummer anpassen, NameData hinten dazuschreiben
                    C1 = CInt(Left(names, 3))     '000    '123
                    If C1 = 0 Then
                        names = "001 Names: |" + ActualNameData
                    Else
                        names = Format(CStr(C1 + 1), "000") + Mid(names, 4) + ActualNameData
                    End If
                'Names in T4 schreiben
                    .Cells(zT4, 12) = names
            Else
                'Name war bereits in der Namensliste; nothing new in ActualNameData
                'Nothing to do
            End If
            End With
End Sub

Sub T4_Add_SomeContent_ToOneLineInT5(Nn$, Vn$, zT5%, Jhg$, Club$, Nat$, sw$)
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'zT5            = ZeilenNr des aktuellen competitors in T5
    
    'Vorbereitung
        Dim Sw2$, T$, z1%, s1%
        With Sheets("T5")
    'Dg-Ecke li ob; (z1,s1)
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39))
    'Action     'bestehende T5-Einträge werden nicht geändert
                'leere T5-Items werden ergänzt
            'Jhg
                If .Cells(zT5, 6) = "" And Jhg <> "" Then
                    .Cells(zT5, 6) = Jhg
'                    t = "Ein Jahrgang wurde neu in T5 übernommen:" + vbCrLf _
'                        + Vn + " " + Nn + ", Jhg. " + Jhg
'                    UF3_Show_Text_ForXSeconds t, 3
                End If
            'Club
                If .Cells(zT5, 7) = "" And Club <> "" Then .Cells(zT5, 7) = Club
            'Nat
                If .Cells(zT5, 9) = "" And Nat <> "" Then .Cells(zT5, 9) = Nat
            'Sw
                Sw2 = .Cells(zT5, 15)
                If Sw2 = "-" Then
                    .Cells(zT5, 15) = sw
                Else
                    If Not Sw2 Like "*" + sw + "*" Then .Cells(zT5, 15) = Trim(Sw2 + " " + sw)
                End If
    'Jhg zentriert in T5
        .Cells(zT5, 6).HorizontalAlignment = xlCenter 'Jhg
    'Finals
        End With
End Sub

Function T4_LastClickWasInsideActualDg() As Boolean
    'Called from    T4_DgClick_CheckOneName_LastSelected
    
    'Vorbereitung
        Dim s1%, s2%, sSh%, z1%, z2%, zSh%
    'DgCorners
        z1 = CInt(ArrC(37)): z2 = CInt(ArrC(38)): s1 = CInt(ArrC(39)): s2 = CInt(ArrC(40))
    'Action
        zSh = CInt(ArrC(104)): sSh = CInt(ArrC(105)) 'zShLast1S, sShLast1S
        If zSh < z1 Or zSh > z2 Then Exit Function
        If sSh < s1 Or sSh > s2 Then Exit Function
        T4_LastClickWasInsideActualDg = True
End Function




Function T4_Get_zDgLast1Name() As Integer
    'Called from    xxx
    Dim z1%, zSh%: z1 = CInt(ArrC(37)): zSh = T4_Get_zShLast1Name
    T4_Get_zDgLast1Name = zSh - z1 + 1
End Function
Function T4_Get_sDgLast1Name() As Integer
    'Called from    xxx
    Dim s1%, sSh%: s1 = CInt(ArrC(39)): sSh = T4_Get_sShLast1Name
    T4_Get_sDgLast1Name = sSh - s1 + 1
End Function
Function T4_Get_zDgLast2Name()
    'Called from    xxx
    'Status         T4_LastClickWasInsideActualDg = True
    Dim z1%, zSh%: z1 = CInt(ArrC(37)): zSh = T4_Get_zShLast2Name
    T4_Get_zDgLast2Name = zSh - z1 + 1
End Function
Function T4_Get_sDgLast2Name()
    'Called from    xxx
    'Status         T4_LastClickWasInsideActualDg = True
    Dim s1%, sSh%: s1 = CInt(ArrC(39)): sSh = T4_Get_sShLast2Name
    T4_Get_sDgLast2Name = sSh - s1 + 1
End Function





Function T4_Get_zDgLast1()
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'Status         T4_LastClickWasInsideActualDg = True
    Dim z1%, zSh%: z1 = CInt(ArrC(37)): zSh = T4_Get_zShLast1
    T4_Get_zDgLast1 = zSh - z1 + 1
End Function
Function T4_Get_sDgLast1()
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'Status         T4_LastClickWasInsideActualDg = True
    Dim s1%, sSh%: s1 = CInt(ArrC(39)): sSh = T4_Get_sShLast1
    T4_Get_sDgLast1 = sSh - s1 + 1
End Function
Function T4_Get_zDgLast2()
    'Called from    xxx
    'Status         T4_LastClickWasInsideActualDg = True
    Dim z1%, zSh%: z1 = CInt(ArrC(37)): zSh = T4_Get_zShLast2
    T4_Get_zDgLast2 = zSh - z1 + 1
End Function
Function T4_Get_sDgLast2()
    'Called from    xxx
    'Status         T4_LastClickWasInsideActualDg = True
    Dim s1%, sSh%: s1 = CInt(ArrC(39)): sSh = T4_Get_sShLast2
    T4_Get_sDgLast2 = sSh - s1 + 1
End Function





Function T4_Get_mw_FromDg(sDg%) As String
    'Called from    T4_DgClick_CheckOneName_LastSelected
    
    If ArrDg(5, sDg) = "W" Then T4_Get_mw_FromDg = "w"
    If ArrDg(5, sDg) = "M" Then T4_Get_mw_FromDg = "m"
End Function

Function T4_Get_mw_FromT5(zT5%) As String
    'Called from    T4_DgClick_CheckOneName_LastSelected
    
    Dim mw$
    If zT5 = 0 Then Exit Function
    T4_Get_mw_FromT5 = Sheets("T5").Cells(zT5, 5)
End Function
            
Function T4_Get_Club_FromDgClubColumn(zDg%, sDg%) As String
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'Dg(zDg, sDg)   = LastSelectedCell; actual competitor
    
    'Vorbereitung
        Dim Club$, SpClub$, i%, s%, sClb%, z%, sClub() As String
        SpClub = ","         'T4 Dg-SpaltenNrn der Verein-Spalten = ",4,20,"
        If CInt(ArrC(96)) > 0 Then SpClub = SpClub + CStr(ArrC(96)) + ","
        If CInt(ArrC(97)) > 0 Then SpClub = SpClub + CStr(ArrC(97)) + ","
        If CInt(ArrC(98)) > 0 Then SpClub = SpClub + CStr(ArrC(98)) + ","
        If SpClub = "," Then Exit Function
    'Club    Welche der beiden sCLUB-Spalten gehört zu sDg?
    '       SpClub eines Competitors ist max. +/- 2 Spalten vom Namen entfernt (je w/m)
            sClub = Split(SpClub, ",")
        For i = 1 To UBound(sClub) - 1
            sClb = CInt(sClub(i))
            If Abs(sDg - sClb) < 3 Then Club = CStr(ArrDg(zDg, sClb)): Exit For
        Next
    'Finals
        T4_Get_Club_FromDgClubColumn = Club
End Function
        
Function T4_Get_Nation_FromDgNationColumn(zDg%, sDg%) As String
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'Dg(zDg, sDg)   = LastSelectedCell; actual competitor; valid NameCell
    
    'Vorbereitung
        Dim Nat$, sDgNameM%, sDgNameW%, sDgNameX%, sDgNationM%, sDgNationW%, sDgNationX
        sDgNameW = CInt(ArrC(90)):   sDgNameM = CInt(ArrC(91)):   sDgNameX = CInt(ArrC(92))
        sDgNationW = CInt(ArrC(93)): sDgNationM = CInt(ArrC(94)): sDgNationX = CInt(ArrC(95))
    'Existiert eine Nation-Spalte?
        If sDgNationW + sDgNationM + sDgNationX = 0 Then Exit Function
    'Existiert nur eine NationW-Spalte?
        If sDgNationW > 0 And sDgNationM = 0 Then sDgNationM = sDgNationW
    'Nat    Welche der beiden sNATION-Spalten gehört zu sDg?
        If sDg = sDgNameW Then
                                   Nat = CStr(ArrDg(zDg, sDgNationW))
        ElseIf sDg = sDgNameM Then Nat = CStr(ArrDg(zDg, sDgNationM))
        ElseIf sDg = sDgNameX Then Nat = CStr(ArrDg(zDg, sDgNationX))
        End If
    'Finals
        T4_Get_Nation_FromDgNationColumn = Nat
End Function
        
Function isGermanEvent(DgTitle$) As Boolean
    Dim s$, i%, Arr1() As String
    s = "DM,DSM,DJM,BadM,BadSM,BadJM"
    Arr1 = Split(s, ",")
    For i = 0 To UBound(Arr1)
        If DgTitle Like "*" + Arr1(i) + "*" Then isGermanEvent = True: Exit Function
    Next
End Function

Function T4_Cut_OutOfDgNameCellContent_c_Club(N$) As String
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'N              DgNameCellContent, reduced: No Jhg, Age, Nn, Vn, mw
    'Action         Herausnahme aus N: Zusatzangabe Club/Ort/Verein
    '               Rückgabe: "", Club
    
    'Exit
        If Not N Like "*=*" Then Exit Function
        If Not (N Like "*c=*" Or N Like "*o=*" Or N Like "*v=*") Then Exit Function
    'Vorbereitung
        Dim Club$, c%
    'Get Club
        c = InStr(1, N, "c=") + InStr(1, N, "o=") + InStr(1, N, "v=")
        If Not Mid(N, c + 1, 1) = "=" Then Stop
        Club = Mid(N, c + 2) '"Bonn", "Bad X",  "Bad X s=... n=...",
        c = InStr(1, Club, "="): If c > 0 Then Club = Left(Club, c - 2)
    'Cut Club
        N = Trim(Replace(N, Club, "")): N = Trim(Replace(N, "c=", ""))
        N = Trim(Replace(N, "o=", "")): N = Trim(Replace(N, "v=", ""))
    'Return
        T4_Cut_OutOfDgNameCellContent_c_Club = Club
End Function

Function T4_Cut_OutOfDgNameCellContent_s_Spelling(N$) As String
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'N              DgNameCellContent, reduced: No Jhg, Age, Nn, Vn, mw, Club
    'Action         Herausnahme aus N: Zusatzangabe s= (Schreibweise)
    '               Rückgabe: "", Sw
    
    'Exit
        If Not N Like "*s=*" Then Exit Function
    'Vorbereitung
        Dim sw$, c%
    'Action
        c = InStr(1, N, "s=")
        sw = Mid(N, c + 2) '"May Lui", "May Lui n=...",
        c = InStr(1, sw, "="): If c > 0 Then sw = Trim(Left(sw, c - 2))
    'Cut
        N = Trim(Replace(N, sw, "")): N = Trim(Replace(N, "s=", ""))
    'Return
        T4_Cut_OutOfDgNameCellContent_s_Spelling = sw
End Function

Function T4_Cut_OutOfDgNameCellContent_x_mw(N$) As String
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'N              DgNameCellContent with valid name, reduziert: no Age Jhg m/w
    'Action         Herausnahme aus N: m, w
    '               Rückgabe: "", "w" oder "m"
    
    '"w" oder "m" als einzeln angefügt
        If N Like "* w *" Then N = Trim(Replace(N, " w ", " ")): T4_Cut_OutOfDgNameCellContent_x_mw = "w": Exit Function
        If N Like "* w" Then N = Left(N, Len(N) - 2): T4_Cut_OutOfDgNameCellContent_x_mw = "w": Exit Function
        If N Like "* m *" Then N = Trim(Replace(N, " m ", " ")): T4_Cut_OutOfDgNameCellContent_x_mw = "m": Exit Function
        If N Like "* m" Then N = Left(N, Len(N) - 2): T4_Cut_OutOfDgNameCellContent_x_mw = "m": Exit Function
    'Exit
        If Not N Like "*x=*" Then Exit Function
    'Vorbereitung
        Dim mw$
    'Action
        If N Like "*x=w*" Then mw = "w": N = Replace(N, "x=w", "")
        If N Like "*x=m*" Then mw = "m": N = Replace(N, "x=m", "")
        N = Trim(Replace99(N, "  ", " "))
    'Return
        T4_Cut_OutOfDgNameCellContent_x_mw = mw
End Function

Function T4_Cut_OutOfDgNameCellContent_n_NewNameGlobal(N$) As String
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'N              DgNameCellContent, reduced: No Jhg, Age, Nn, Vn, mw, Club
    'Action         Herausnahme aus N: Zusatzangabe n= (change ActualName to NewName)
    '               Rückgabe: "", NewName
    
    'Exit
        If Not N Like "*n=*" Then Exit Function
    'Vorbereitung
        Dim NewName$, c%
    'Get NewName
        c = InStr(1, N, "n=")
        NewName = Mid(N, c + 2) '"May Lui", "May Lui o=...",
        c = InStr(1, NewName, "="): If c > 0 Then NewName = Trim(Left(NewName, c - 2))
        
    'Cut
        N = Trim(Replace(N, NewName, "")): N = Trim(Replace(N, "n=", ""))
    'Return
        T4_Cut_OutOfDgNameCellContent_n_NewNameGlobal = NewName
End Function

Function T4_Cut_OutOfDgNameCellContent_Age(N$)
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'N              CompleteDgNameCellContent; wird hier ggf. reduziert
    
    'Exit
        If Not N Like "*(*" Then Exit Function
    'Vorbereitung
        Dim Age$, C1%, C2%
        N = Trim(N)
    'Get Age
        C1 = InStr(1, N, "("): C2 = InStr(1, N, ")"): If C2 = 0 Then Stop
        Age = CutVonBis(N, C1 + 1, C2 - 1) '17
        If Not IsNumeric(Age) Then Stop
    'Cut
        N = Trim(Replace(N, "(" + Age + ")", ""))
    'Return
        T4_Cut_OutOfDgNameCellContent_Age = Age
End Function
        
Function T4_Cut_OutOfDgNameCellContent_j_Jhg(N$) As String
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'N              DgNameCellContent, reduziert um Age (kein " (19)")
    '               wird hier ggf. weiter reduziert
    'Action         Herausnahme aus N: Zusatzangabe Jhg
    '               Rückgabe: "", "1975"
    'Vorbereitung
        Dim Jhg$, c%, i%, pos%
        'N = CompleteDgNameCellContent '"Lea May", "Lea2 May78", "Lea Mayj=78", "Lea May (12)"
    'Exit
        If Not N Like "*##*" Then Exit Function     'Kein Jhg enthalten
        If N Like "*(##)*" Then Exit Function       'Jhg bereits erkannt
    'Nur genau 1x zwei aufeinander folgende Ziffern
        For i = 1 To Len(N) - 1
            If Mid(N, i, 2) Like "##" Then c = c + 1: pos = i
        Next
        If c <> 1 Then Stop 'Mehr als 2 aufeinanderfolgende Ziffern (### oder ## ... ##)
    'Jhg
        Jhg = Mid(N, pos, 2)
    'Cut (reduce N)
        N = Trim(Replace(N, Jhg, "")): N = Trim(Replace(N, "j=", "")): N = Replace99(N, "  ", " ")
    'Return
        T4_Cut_OutOfDgNameCellContent_j_Jhg = "19" + Jhg
End Function

Function T4_Cut_OutOfDgNameCellContent_VnNn(N$) As String
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'N              DgNameCellContent with valid name, reduziert: no Age Jhg m/w
    '               = "Lea Mai", "Lea Mai o=Bonn s=May"
    'Action         Herausnahme aus N: Nn, Vn
    '               Rückgabe: Nn|Vn
    
    'Vorbereitung
        Dim Nn$, NnVn$, Vn$, c%
    'Action
        c = InStr(1, N, "="): If c > 0 Then NnVn = Trim(Left(N, c - 2)) Else NnVn = N
        c = InStr(1, NnVn, " "): If c = 0 Then Stop
        Vn = Left(NnVn, c - 1): Nn = Mid(NnVn, c + 1)
    'Cut
        N = Trim(Replace(N, Vn + " " + Nn, ""))
    'Return
        T4_Cut_OutOfDgNameCellContent_VnNn = Nn + "|" + Vn
End Function

Sub T4_Create_NewLineInT5_ifNecessary(zT5%, zDg%, sDg%, Nn$, Vn$, mw$, NewName$)
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'zT5            wird hier gefüllt; besitzt dann Wert > 0
    'Action         stellt Existenz einer T5-Zeile sicher (zum aktuellen competitor)
    
    'zT5
        If NewName <> "" Then
            zT5 = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, NewName)
        End If
        If zT5 = 0 Then
            zT5 = Get_RowNr_HoldingMyTextWholeInColumnX("T5", 22, 7, Vn + " " + Nn)
        End If
    'Exit
        If zT5 > 0 Then Exit Sub
    'Create
        '|Nachname|Vorname| ist nicht in T5 enthalten
        'zT5
            zT5 = Get_NrOfLastRowInColumnNr(3, "T5") + 1
        'Neue Zeile in T5 anlegen und formatieren
            T4_DgHandleNameCell_UnknownName zT5, zDg, sDg, Vn, Nn, mw
End Sub

Sub T4_DgHandleNameCell_UnknownName(zT5%, zDg%, sDg%, Vn$, Nn$, mw$) ', JhgDg$, VerDg$, NatDg$, Sw$)
    'Called from    T4_Add_SomeContent_ToOneLineInT5, T4_DgClick_CheckNames
    'Action         Dg-NamenZelle in T4 neu füllen
    '               Neue Zeile in T5 anlegen
    'Status T4      User könnte Zusatzangaben gemacht haben, die im Dg nicht mehr sichtbar sein sollen;
    '               alle Zusatzangaben wurden bereits ausgewertet;
    '               die Dg-Zelle soll letztlich nur Nn, Vn, ggf. Alter tragen, sonst nichts
    'Status T5      CompetitorName in Dg existiert noch nicht in T5
    
    'Vorbereitung
        Dim Alter$, i%, s1%, y%, z1%, r As Range
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39))
'        Y = CInt(Left(ArrDg(2, 2), 4)) 'Year of event
'    'T4
'        With Sheets("T4")
'            If JhgDg = "" Then
'                'Zum Aktiven in T5-Zeile zT5 ist kein Jahrgang vermerkt
'                .Cells(z1 - 1 + zDg, s1 - 1 + sDg) = Vn + " " + Nn 'NamenZelle ohne Jhg
'            Else
'                Alter = CStr(Y - CInt(JhgDg))
'                .Cells(z1 - 1 + zDg, s1 - 1 + sDg) = Vn + " " + Nn + " (" + Alter + ")"
'            End If
'        End With
    'T5 - Anlegen einer neuen Zeile in T5-PersonData (Name, Verein, ...)
        With Sheets("T5")
            'Hintergrundfarbe der T5-Zeile
                Set r = .Range(.Cells(zT5, 3), .Cells(zT5, 10))
                If mw = "w" Then r.Interior.Color = 14083324 Else r.Interior.Color = 15652797 'hellblau/Blau1 (m-Farbe)
                Set r = .Range(.Cells(zT5, 12), .Cells(zT5, 22))
                If mw = "w" Then r.Interior.Color = 14083324 Else r.Interior.Color = 15652797 'hellblau/Blau1 (m-Farbe)
            'BorderAround
                For i = 3 To 22
                    Set r = .Cells(zT5, i)
                    r.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, ColorIndex:=2
                Next
            'ZeilenNr
                Set r = .Cells(zT5, 2): r.Value = zT5 - 7
                r.HorizontalAlignment = xlCenter: r.Font.size = 6
            'Nachname, Vorname
                .Cells(zT5, 3) = Nn:      .Cells(zT5, 4) = Vn
            'm/w
                Set r = .Cells(zT5, 5): r.Value = mw: r.HorizontalAlignment = xlCenter
            'Jahrgang, Verein
                Set r = .Cells(zT5, 6) ': r.Value = JhgDg:
                r.HorizontalAlignment = xlCenter
'            'Club/Verein
'                .Cells(zT5, 7) = VerDg
            'Nation
                Set r = .Cells(zT5, 9) ': r.Value = NatDg
                r.HorizontalAlignment = xlCenter
            'ZusatzSpalte 11    'Hilfsspalte, hellgraue Schrift
                Set r = .Cells(zT5, 1): r.Font.Color = RGB(155, 155, 155): r.HorizontalAlignment = xlLeft: r.Font.size = 6
            'ZusatzSpalten 12-21
                Set r = .Range(.Cells(zT5, 12), .Cells(zT5, 21)): r.Value = "-": r.HorizontalAlignment = xlCenter
            'ZusatzSpalte 22: Vn Nn
                Set r = .Cells(zT5, 22): r.Value = Vn + " " + Nn: r.Font.size = 8
            'Zusatzspalte 23
                .Cells(zT5, 23) = "|"
'            'Schreibweise
'                If Sw <> "" Then .Cells(zT5, 15) = Sw
        End With
End Sub

Sub T4_CreateOrChange_PersonalFolder(zT5%, Nn$, Vn$, NewName$, Nat$, Club$, NameOfPersonalFolder$, PathOfPersonalFolder$)
    'Called from            T4_DgClick_CheckOneName_LastSelected
    'NameOfPersonalFolder   wird hier belegt
    'PathOfPersonalFolder   wird hier belegt
    
    'Vorbereitung
        Dim Fi$, K1$, K2$, N1$, N2$, N3$, NnNew$, p1$, p2$, pi$, PathOfEventFolder$, v$, VnNew$
        Dim c%, i%, j%, ArrFi() As String, ArrPi() As String
        With Sheets("T5"): v = vbCrLf
    'NewName
        If NewName <> "" Then
            c = InStr(1, NewName, " ")
            VnNew = Left(NewName, c - 1): NnNew = Left(NewName, c + 1)
        End If
    'Existenz des PersonenOrdners in 'Leute' sicherstellen
    'Name des PersonFolder ermitteln (--> FolderName in 'Leute'):
        'NameOfPersonalFolder, PathOfPersonalFolder
        N1 = .Cells(zT5, 10)                                        'NamePersFoldOld in T5
        N2 = T4_Calc_NameOfPersonalFolder(Nn, Vn, Nat, Club)        'NamePersFoldNew
        If NewName <> "" Then N3 = T4_Calc_NameOfPersonalFolder(NnNew, VnNew, Nat, Club)
        
        If Not IsValidName(N2) Then Exit Sub
        
        p1 = ArrC(4) + "\" + N1 '     PfadAlt PersonenOrdner
        p2 = ArrC(4) + "\" + N2 'ggf. PfadNeu PersonenOrdner
        PathOfPersonalFolder = "": NameOfPersonalFolder = ""
        If NewName <> "" Then
            If FolderExists(ArrC(4) + "\" + N3) Then
                NameOfPersonalFolder = N3
                PathOfPersonalFolder = ArrC(4) + "\" + N3
            End If
        End If
        
        
    'N1 = ""
        If N1 = "" Then
            If NameOfPersonalFolder = "" Then
                NameOfPersonalFolder = N2: PathOfPersonalFolder = p2
            End If
            If Not FolderExists(PathOfPersonalFolder) Then CreateFolder PathOfPersonalFolder
            'Eintrag in T5 (PersonData)
                .Cells(zT5, 10) = NameOfPersonalFolder
    'Exit
            Exit Sub
        Else
            NameOfPersonalFolder = N1: PathOfPersonalFolder = p1
        End If
    'N1 <> "", N2 <> ""
        'N1 oder N2 könnte besser sein:
        'L. May 'L. May () 'L. May (GB) 'L. May (Club, GB) 'L. May (vh June, Club, GB)
        If N1 Like "*(*" And N2 Like "*(*" Then
            K1 = CutVonBis(N1, InStr(1, N1, "(") + 1, InStr(1, N1, ")") - 1) 'N1-Klammerinhalt
            K2 = CutVonBis(N2, InStr(1, N2, "(") + 1, InStr(1, N2, ")") - 1) 'N2-Klammerinhalt
            If N1 Like "*" + K2 + "*" Then NameOfPersonalFolder = N1
            If N2 Like "*" + K1 + "*" Then NameOfPersonalFolder = N2
        End If
        If N1 Like "*()" Then
            NameOfPersonalFolder = N2
        ElseIf N1 Like "*(*[A-Z]*)*" And Not N2 Like "*(*[A-Z]*)*" Then NameOfPersonalFolder = N1
        ElseIf N2 Like "*(*[A-Z]*)*" And Not N1 Like "*(*[A-Z]*)*" Then NameOfPersonalFolder = N2
        ElseIf N1 Like "*(*" And Not N2 Like "*(*" Then NameOfPersonalFolder = N1
        ElseIf N2 Like "*(*" And Not N1 Like "*(*" Then NameOfPersonalFolder = N2
        End If
        'NameOfPersonalFolder ist jetzt festgelegt: N1 oder N2
    'N1
        If NameOfPersonalFolder = N1 Then
            PathOfPersonalFolder = p1: If Not FolderExists(p1) Then CreateFolder p1
    'N2
        ElseIf NameOfPersonalFolder = N2 Then
            PathOfPersonalFolder = p2
            'Eintrag in T5 (PersonData)
                .Cells(zT5, 10) = N2
            'N1 in N2 ändern
                If Not FolderExists(p2) Then RenameFolder p1, p2
        End If
    'Mehrere PersonenOrdner?
        'pi = Get_AllSubfolderPaths_LikeMyString_OneLevel(ArrC(4), "*\" + Nn + "*, " + Vn + "*")
        '[Nn + ",] statt [Nn + "*,] sonst wird "Roth, Christiane" in "Rother, Christiane" gefunden
        pi = Get_AllSubfolderPaths_LikeMyString_OneLevel(ArrC(4), "*\" + Nn + ", " + Vn + "*")
        'show pi: Stop
        ArrPi = Split(pi, v)
        If UBound(ArrPi) = 0 Then Exit Sub 'nur 1 PersonalFolder
        For i = 0 To UBound(ArrPi)
            If ArrPi(i) <> PathOfPersonalFolder Then
                'Dateien verschieben
                    Fi = Get_AllFileNamesOfOneFolder(ArrPi(i))
                    ArrFi = Split(Fi, v)
                    For j = 0 To UBound(ArrFi)
                        RenameFile ArrPi(i) + "\" + ArrFi(j), PathOfPersonalFolder + "\" + ArrFi(j)
                    Next
                'Delete
                    DeleteFolder ArrPi(i)
            End If
        Next
    'Finals
        End With
End Sub

Function T4_Calc_NameOfPersonalFolder(Nn$, Vn$, Nat$, Club$) As String
    'Called from    T4_DgClick_CheckNames
    'Action         SOLL-Namen eines PersonenOrdners für 'Leute' ermitteln
    
    'Vorbereitung
        Dim G$, H$, p1$, rn$, vh$, z%
        z = T5_Get_T5RowNrOfOneNnVn(Nn, Vn)
    'Geburts-/Vh-Name
        With Sheets("T5")
            If z > 0 Then
                G = .Cells(z, 13) 'G = GeburtsName = "Maier",  "-"
                H = .Cells(z, 14) 'H = Vh-Name     = "Müller", "-"
                If Club = "" Then Club = .Cells(z, 7)
                If Club = "" Then Club = .Cells(z, 8)
                If Nat = "" Then Nat = .Cells(z, 9)
                rn = .Cells(z, 12) 'Rufname
                If rn = "-" Then rn = ""
            End If
        End With
        If G = "-" Then G = "": H = ""
    'Calc
        p1 = Nn + ", " + Vn
        'Ggf. Club/Nation anfügen
            If Club <> "" And Nat = "D" Then
                                                 p1 = p1 + " (" + Club + ")"
            ElseIf Club = "" And Nat <> "D" Then p1 = p1 + " (" + Nat + ")"
            ElseIf Club <> "" And Nat = "" Then p1 = p1 + " (" + Club + ")"
            ElseIf Club <> "" And Nat <> "D" Then p1 = p1 + " (" + Club + ", " + Nat + ")"
            End If
            
        If rn <> "" Then
            'Die Person besitzt einen Rufnamen
            If p1 Like "*(*" Then
                p1 = Replace(p1, "(", "(" + "rn " + rn + ", ")
            Else
                p1 = p1 + " (rn " + rn + ")"
            End If
        End If
        
        If G <> "" Then
            'Die Person besitzt neben dem Geburtsnamen noch weitere Nachnamen
            vh = Replace(H, "/", " ")
            p1 = Replace(p1, Nn, G)
            If p1 Like "*(*" Then
                p1 = Replace(p1, "(", "(" + "vh " + vh + ", ")
            Else
                p1 = p1 + " (vh " + vh + ")"
            End If
        End If
        p1 = Replace(p1, " ()", "")
        T4_Calc_NameOfPersonalFolder = p1
End Function

Function IsValidName(CellContent$) As Boolean
    'Called from            T4_zDg_sDg_IsValidNameCell, T4_Cut_OutOfDgNameCellContent_j_Jhg
    '                       T4_CreateOrChange_PersonalFolder
    
    'Vorbereitung
        Dim s$, c%, Arr1() As String
        s = Trim(CellContent)
        c = InStr(1, s, "="): If c > 0 Then s = Trim(Left(s, c - 2))
        c = InStr(1, s, " ("): If c > 0 Then s = Trim(Left(s, c - 1))
    'Action
        If s = "" Then Exit Function
        'M. May; Rien de Ruiter; X. van Helsing
        If Not (s Like "*[A-ZÄÖÜÅ]*[A-ZÄÖÜÅ]*") Then Exit Function
        If s Like "*Juti" Then Exit Function
        If s Like "*Jutu" Then Exit Function
        If s Like "*Ti" Then Exit Function
        If s Like "*Tu" Then Exit Function
        If s Like "*Juti *" Then Exit Function
        If s Like "*Jutu *" Then Exit Function
        If s Like "*Ti *" Then Exit Function
        If s Like "*Tu *" Then Exit Function
        If s Like "*klasse Tu*" Then Exit Function
        If s Like "*Jhg*" Then Exit Function
        If s Like "*Jugend*" Then Exit Function
        If s Like "Turner*" Then Exit Function
        If s Like "Schülerinnen*" Then Exit Function
        If s Like "*Damen*" Then Exit Function
        If s Like "*Herren*" Then Exit Function
        If anzAinB("|", s) = 7 Then IsValidName = True: Exit Function 'HelperBox-Zeile
        If Not (s Like "[A-ZÄÖÜÅ]* *[A-ZÄÖÜÅ]*" Or s Like "*[A-ZÄÖÜÅ]*, [A-ZÄÖÜÅ]*") Then
           Beep 400, 666: Stop
            show "Bitte geben sie Vorname und Nachname ein.": Exit Function
        End If
    'Finals
        IsValidName = True
End Function

Sub T5_Check_LTV_Nat_ViaClubCombination(Club$)
    'Called from    T4_Change_T4SomeDgData_OneNameData
        
    'Vorbereitung
        Dim LTV$, Nat$, zLast%, i%, z%, zT5%, CLN()
        zLast = Get_NrOfLastRowInColumnNr(3, "T5")
        With Sheets("T5")
    'Exit
        If Club = "" Then Exit Sub
    'T5-Bereich Spalte 7 bis 9 (Club, LTV, Nation) in Array nehmen
        CLN = .Range(.Cells(8, 7), .Cells(zLast, 9)).Value
    'Search in Array CLN()  'Club LTV Nation
        For i = 1 To UBound(CLN, 1)
            'Schleife über alle T5-Zeilen
            If CStr(CLN(i, 1)) = Club Then
                If CStr(CLN(i, 2)) <> "" Then LTV = CStr(CLN(i, 2))
                If CStr(CLN(i, 3)) <> "" Then Nat = CStr(CLN(i, 3))
                If LTV <> "" And Nat <> "" Then Exit For
            End If
        Next
        'If LTV = "" Then Exit Sub
    'Write LTV to all CLN(i, 2) where CLN(i, 1) = Club
        For i = 1 To UBound(CLN, 1)
            If CStr(CLN(i, 1)) = Club Then CLN(i, 2) = LTV: CLN(i, 3) = Nat
        Next
    'Paste
        Paste_2DArrayToCell_z_s "T5", 8, 7, CLN
    'Finals
        End With
End Sub

Sub T5_Write_Nat(Nat$, Nn$, Vn$)
    'Called from    T4_Change_T4SomeDgData_OneNameData
    
    'Exit
        If Nat = "" Then Exit Sub
    'Vorbereitung
        Dim zT5%
        With Sheets("T5")
    'Write
        zT5 = T5_Get_T5RowNrOfOneNnVn(Nn, Vn)
        If zT5 = 0 Then Stop
        If .Cells(zT5, 9) = "" Then .Cells(zT5, 9) = Nat
    'Finals
        End With
End Sub

Sub T4_NewName_Action(Vn$, Nn$, NewName$)
    'Called from    T4_DgClick_CheckOneName_LastSelected
    
    'Exit
        If NewName = "" Then Exit Sub
    'Vorbereitung
        Dim VnNew$, NnNew$
        VnNew = Left(NewName, InStr(1, NewName, " ") - 1)
        NnNew = Mid(NewName, InStr(1, NewName, " ") + 1)
    'Changes
        T5_ChangeNameGlobal Vn, Nn, VnNew, NnNew
End Sub

Sub T4_Write_Club_toDg(Club1$, zT5%, zDg%, sDg%)
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'Club1          = ClubOfCompetitor; wird hier gefüllt
    'zT5            = CompetitorRowNr in T5
    'zDg, sDg       Dg(zDg, sDg) CompetitorNameCell

    'Exit
        If ArrDg(2, 1) = "M0V1" Then Exit Sub
        If ArrDg(2, 1) = "M2V8" Then Exit Sub
    'Vorbereitung
        Dim Club2$, s1%, sClub%, z1%
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39))
    'sClub         'Dg-SpaltenNr der zum aktuellen Namen gehörigen Dg-Club-Spalte
        If sDg = CInt(ArrC(90)) Then sClub = CInt(ArrC(96)) 'sDgNameW --> sDgClubW
        If sDg = CInt(ArrC(91)) Then sClub = CInt(ArrC(97)) 'sDgNameM --> sDgClubM
        If sDg = CInt(ArrC(92)) Then sClub = CInt(ArrC(98)) 'sDgNameX --> sDgClubX
        If sClub = 0 Then Exit Sub 'sClub = DgSpaltenNr Verein
    'Write Club of actual competitor to Dg
        If Club1 = "" Then                       'Club1 = aus Dg ermittelt
            Club2 = Sheets("T5").Cells(zT5, 7)   'Club2 = aus T5 ermittelt
            If Club2 <> "" Then
                If ArrDg(zDg, sClub) = "" Then
                    EE 0
                    Sheets("T4").Cells(z1 - 1 + zDg, s1 - 1 + sClub) = Club2
                    EE 1 'EE: sonst Worksheet_Change
                End If
            End If
        Else
            EE 0: Sheets("T4").Cells(z1 - 1 + zDg, s1 - 1 + sClub) = Club1: EE 1 'EE: sonst Worksheet_Change
        End If
End Sub

Sub T4_Write_Nation_toDg(Nat1$, zT5%, zDg%, sDg%)
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'Nat1           = NationOfCompetitor; wird hier gefüllt
    'zT5            = CompetitorRowNr in T5
    'zDg, sDg       Dg(zDg, sDg) CompetitorNameCell

    'Vorbereitung
        Dim Nat2$, s1%, St%, z1%
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39))
    'sT         'Dg-SpaltenNr der zum aktuellen Namen gehörigen Dg-Nation-Spalte
        If sDg = CInt(ArrC(90)) Then St = CInt(ArrC(93)) 'sDgNameW, sDgNationW
        If sDg = CInt(ArrC(91)) Then St = CInt(ArrC(94)) 'sDgNameM, sDgNationM
        If sDg = CInt(ArrC(92)) Then St = CInt(ArrC(95)) 'sDgNameX, sDgNationX
        If St = 0 Then Exit Sub
    'Write Nat of actual competitor to Dg
        If Nat1 = "" Then                       'Nat1 = Nation, aus Dg ermittelt
            Nat2 = Sheets("T5").Cells(zT5, 9)   'Nat2 = Nation, aus T5 ermittelt
            If Nat2 <> "" Then
                If ArrDg(zDg, St) = "" Then
                    EE 0: Sheets("T4").Cells(z1 - 1 + zDg, s1 - 1 + St) = Nat2: EE 1
                End If
            End If
        Else
            EE 0: Sheets("T4").Cells(z1 - 1 + zDg, s1 - 1 + St) = Nat1: EE 1
        End If
End Sub

Sub T4_Write_NewDgNameCellContent(Nn$, Vn$, Jhg$, Age$, zT5%, zDg%, sDg%)
    'Called from    T4_DgClick_CheckOneName_LastSelected
    'Dg(zDg, sDg)   = LastSelectedCell; actual competitor
    'Jhg            = "" or "1970" if given with Name in LastSelectedCell
    'Age            = "" or "17"   if given with Name in LastSelectedCell
    
    'Vorbereitung
        Dim s1%, y%, z1%, r As Range
        z1 = CInt(ArrC(37)): s1 = CInt(ArrC(39))
        Set r = Sheets("T4").Cells(z1 - 1 + zDg, s1 - 1 + sDg)
    'Alter
        If Age <> "" Then r.Value = Vn + " " + Nn + " (" + Age + ")": Exit Sub
        If Trim(Jhg) = "" Then Jhg = Trim(CStr(Sheets("T5").Cells(zT5, 6)))
        If Jhg = "" Then
            'Zum Aktiven in T5-Zeile zT5 ist kein Jahrgang vermerkt
            'Darstellung der Dg-NamenZelle: ohne Jhg
            r.Value = Vn + " " + Nn
        Else
            y = CInt(Left(ArrDg(2, 2), 4))      'Year of event
            Age = CStr(y - CInt(Jhg))
            'Darstellung der Dg-NamenZelle: mit Jhg
            r.Value = Vn + " " + Nn + " (" + Age + ")"
        End If
End Sub




