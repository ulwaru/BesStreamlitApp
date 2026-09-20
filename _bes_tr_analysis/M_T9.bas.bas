Attribute VB_Name = "M_T9"
Option Explicit 'M19

Function Get_PathsOfAllDateEventFolders() As String
    'Called from    T6_Check_FolderPaths, T6_BMB_EventFileNames, T6_ChangeEFN_Start,
    '               T9_Check_neue_Events_Termine_im_Ordner_Events, T9_Load_DetailsOfSomeDateEventFolders

    'Vorbereitung
        Dim L$, LN$, m$, p$, PN$, s$, v$, i%, A() As String
        v = vbCrLf: DoArrc
    'Liste M: Pfade von Ordnern in Events
        m = Get_Paths_OfAllSubfolders_LikeMyString_AllLevels(ArrC(3), "*\####*")  'Folder "Events"
            'AllPaths inside "Events"; all Subfolders AND all SubSubFolders
            'Subfolder XOR SubSubFolder is the DateFolder
    'Liste s: Sammlung DateFolder
        A = Split(m + v, v)
        For i = 0 To UBound(A) - 1
            L = A(i)        'one Path inside "Events"; L1 XOR L2
                'L1 = Level1-Subfolder          L2 = Level2-Subfolder
                '"*\Events\1987 DM Bonn"        "*\Events\1987 DM Bonn\VnNnImages"
                '"*\Events\1987 Liga"           "*\Events\1987 Liga\1987-10-12 Buli2 ADorf-BDorf"
            LN = NameOfPath(L)
            p = Get_ParentFolderPath_OfFolderPath(L) 'ParentFolderPath of L
            PN = NameOfPath(p)
            If PN = "Events" Then
                'L ist ein L1
                    'In welchen Fällen gilt L = DateFolder?
                        'L = DateFolder, falls L einen SubFolder ohne "\####" besitzt
                        If m Like "*" + L + "\*" Then
                            'L besitzt einen SubFolder
                            If Not m Like "*" + L + "\####*" Then s = s + L + v
                        'L = DateFolder, falls L keinen SubFolder besitzt
                        Else
                            'L besitzt keinen SubFolder
                            s = s + L + v       'Sammlung DateFolder
                        End If
            Else
                'L ist ein L2
                    'In welchen Fällen gilt L = DateFolder?
                    'L = DateFolder, falls LN ein Datum besitzt
                        If LN Like "####*" Then s = s + L + v
            End If
        Next
        s = Delete_EmptyRowsInString(s)
        Get_PathsOfAllDateEventFolders = s
        'show s
End Function

Sub T9_Load_DetailsOfSomeDateEventFolders_TEST()
    Dim T()
    T9_Load_DetailsOfSomeDateEventFolders T
    showArray2D T
End Sub

Sub T9_Load_DetailsOfSomeDateEventFolders(T, Optional SomeDateEventFolders$ = "")
    'Called from    T9_UpdateTerminalender, T9_Add_EventFoldersToTerminalender
    'T              2D-Array im T9-Raster, anfangs leer
    
    'Vorbereitung
        Dim L$, LN$, p$, s$, v$, i%, A() As String
        v = vbCrLf
    'Liste s: Sammlung DateFolder
        s = SomeDateEventFolders
        If s = "" Then s = Get_PathsOfAllDateEventFolders
        s = Delete_EmptyRowsInString(s)
    'T() füllen     'FolderTermine im T9-Raster
        A = Split(v + s, v): ReDim T(1 To UBound(A), 1 To 10)
        For i = 1 To UBound(A)
            L = A(i)        'one Line
            LN = NameOfPath(L)
            p = Get_ParentFolderPath_OfFolderPath(L)
            T(i, 1) = T9_Get_Date8FromEventFolderName(LN)        'Date8
            T(i, 2) = Get_Date10FromDate8(CStr(T(i, 1)))         'Date10
            T(i, 8) = 1                                          'OrdnerSymbol
            T(i, 9) = LN                                         'OrdnerName
            T(i, 10) = L                                         'OrdnerPfad
            LN = Mid(LN, InStr(1, LN, " ") + 1)                  'Neubelegung LN: ohne Datum
            T(i, 6) = T9_Get_NationFromEventFolderNameVnNn(LN)   'Nation 'LN neu: ohne Nation
            T(i, 3) = T9_Get_EventFromEventFolderNameVnNn(LN)    'Event  'LN neu: ohne Event
            T(i, 4) = LN                                         'Ort
        Next
End Sub

Function T9_Get_EventFromEventFolderNameVnNn(N$)
    'Called from    T9_Load_DetailsOfSomeDateEventFolders
    'N              = Teil/VnNn eines EventFolderName
    '               (ohne Datum, ohne Nation; nur noch Event Ort)
    '               EM15 Eindhoven 'Buli1Süd Bruchsal-Gernsbach
    '               NordbadenLiga BadenOos-Hemsbach 'Buli1Süd(3) Freiburg-Wiesloch
    'Event          wird hier aus N ermittelt, dann aus N entfernt
    'Ort            wird letztlich in N stehen
    
    'Vorbereitung
        Dim s$, AnzLZ%, C1%, C2%, Liga As Boolean
    'Liga?
        If N Like "*Buli*" Or N Like "*Liga*" Then Liga = True
        If N Like "*BuliE*" Then Liga = False
    'Eines der Leerzeichen in N ist der Trenner zw. Event und Ort
    'Bestimmte Leerzeichen sollen nicht als Trenner wirken
    '   werden ersetzt durch @ (werden später rückersetzt)
        If N Like "*LK *" Then N = Replace(N, "LK ", "LK@")
        If N Like "*. *" Then N = Replace(N, ". ", ".@")
        If N Like "*sch Gmünd*" Then N = Replace(N, "sch Gmünd", "sch@Gmünd")
    'Leerzeichen in N
        AnzLZ = anzAinB(" ", N)
        If AnzLZ > 0 Then C1 = InStr(1, N, " ")
        If AnzLZ > 1 Then C2 = InStr(C1 + 1, N, " ")
        Stop
        If C1 = 0 Then
            'N enthält kein Leerzeichen; N ist gesuchtes Event s
            s = N: N = "" 'wird vom Caller übernommen
        Else
            C2 = InStr(C1 + 1, N, " ")
            If C2 = 0 Then
                'N besitzt genau 1 Leerzeichen (keine 2 Leerzeichen)
                'C1 = Position des Leerzeichens
                If Liga Then
                    s = N: N = ""
                Else
                    s = Left(N, C1 - 1): N = Mid(N, C1 + 1)
                End If
            Else
                'N besitzt 2 Leerzeichen
                'C1 = Position des 1. Leerzeichens
                'C2 = Position des 2. Leerzeichens
                If Liga Then
                    s = Left(N, C2 - 1): N = Mid(N, C2 + 1)
                Else
                    'DJM DM Ort
                    If Mid(N, C2 - 1, 1) Like "[A-Z]" Then
                        s = Left(N, C2 - 1)     'Event
                        N = Mid(N, C2 + 1)      'Ort
                    Else
                        Stop
                    End If
                End If
            End If
        End If
    'Finals
        s = Replace(s, "@", " "): N = Replace(N, "@", " ")
        T9_Get_EventFromEventFolderNameVnNn = s
End Function

Function T9_Get_OrtFromEventFolderNameVnNn(N$)
    Dim s$, c%
    c = InStr(1, N, " ")
    If c = 0 Then s = "" Else s = Mid(N, c + 1)
    T9_Get_OrtFromEventFolderNameVnNn = s
End Function

Function T9_Get_NationFromEventFolderNameVnNn(N$)
    Dim s$
    If N Like "*_[A-Z]" Then
        s = Right(N, 1): N = Left(N, Len(N) - 2)
    ElseIf N Like "*_[A-Z][A-Z]" Then s = Right(N, 2): N = Left(N, Len(N) - 3)
    ElseIf N Like "*_[A-Z][A-Z][A-Z]" Then s = Right(N, 3): N = Left(N, Len(N) - 4)
    Else: s = "D"
    End If
    T9_Get_NationFromEventFolderNameVnNn = Trim(s)
End Function

Function Get_Date10FromDate8(d10)
    Get_Date10FromDate8 = Mid(d10, 7, 2) + "." + Mid(d10, 5, 2) + "." + Mid(d10, 1, 4)
End Function

Function T9_Get_Date8FromEventFolderName(N$) As String
    'FolderName     1989 Treff, 1989-07 Treff, 1989-07-14 Treff
    '               - International, ARG, ...
    If Not N Like "[12][90]*" Then T9_Get_Date8FromEventFolderName = "": Exit Function
    Dim D1$, D2$, c%
    c = InStr(1, N, " ")
    D1 = Left(N, c - 1)
    If Len(D1) = 4 Then
        D2 = D1 + "0000"
    ElseIf Len(D1) = 7 Then D2 = Left(D1, 4) + Mid(D1, 6, 2) + "00"
    ElseIf Len(D1) = 10 Then D2 = Replace(D1, "-", "")
    Else: Stop
    End If
    T9_Get_Date8FromEventFolderName = D2
End Function


