Attribute VB_Name = "M_UF1"
Option Explicit 'M_UF1

Sub zzz_M_UF1()
    'showProcs "tab"
    Application.EnableEvents = True
    'RenameModule "Modul1", "M_Dg"
    'Dim tic&: tic = GetTickCount()   ...   ShowTime tic
    'T1_ShowLogBuch
End Sub

Sub UF1_Test()
    ShowUF1 "Harald Müller"
End Sub

Sub ShowUF1(VnNn$)
    'Called from    xxx
    
    'Vorbereitung
        Dim Nat$, N() As String
        With UF1: FillArrC 49, "-": FillArrC 50, "-"
    'L1, ComboBox füllen
        Nat = "ARG - Argentinien|AUS - Australien|AUT - Österreich|BEL - Belgien|BLR - Belarus|BRA - Brasilien|BUL - Bulgarien|CAN - Kanada|CH  - Schweiz|CSR - Tschechoslowakei|CZE - Tschechien|D   - Deutschland|DK  - Dänemark|ENG - England|F   - Frankreich|GB  - Großbritannien|GEO - Georgien   |GRE - Griechenland|HKG - Hongkong|ITA - Italien|JAP - Japan|LAT - Lettland|NL  - Niederlande|NOR - Norwegen|NZL - Neuseeland|POL - Polen|POR - Portugal|RUS - Russland|SA  - Südafrika|SCO - Schottland|SPA - Spanien|SU  - Sowjetunion|SVK - Slowakei|SWE - Schweden|UKR - Ukraine|USA - Vereinigte Staaten von Amerika|UZB - Usbekistan|WAL - Wales"
        N = Split(Nat, "|"): .ComboBox1.List = N: .ComboBox1.ListIndex = 11
        .L1 = VnNn: .m = False: .w = False
    'Finals
        .show (1): End With
End Sub

Sub UF1_OK_Button_wasClicked()
    Dim mw$, Nation$
    mw = "-":                                       If UF1.m = True Then mw = "m"
    Nation = Trim(Left(UF1.ComboBox1.Value, 3)):    If UF1.w = True Then mw = "w"
    FillArrC 49, mw: FillArrC 50, Nation: UF1.Hide
End Sub


