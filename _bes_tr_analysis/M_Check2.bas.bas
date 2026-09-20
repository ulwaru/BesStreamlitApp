Attribute VB_Name = "M_Check2"
Option Explicit 'M_Check2

Sub zzz_M_Check2()
    
    showProcs "ico"

    EE 1: Beep
End Sub

Sub Check_ClubLtvNationLink_All_EachOneInsidePersFold()
    Dim Club$, JumpTo$, LinkNameToDisplay$, LTV$, Nation$, p$, PathOfIcon$, VnNn$
    Dim i%, zT5%, L() As String
    With Sheets("T5")
    'Clubs = Get_AllT5Clubs
    Load_ArrNamesOfAllFolders_Leute L
    For i = LBound(L) To UBound(L)
        p = ArrC(4) + "\" + L(i) 'Path of one PersFold
        'VnNn = Get_VnNn_From_NameOfPersFold(L(i))
        'zT5 = Get_RowNr_HoldingMyTextPartInColumnX("T5", 22, 7, VnNn)
        zT5 = Get_RowNr_HoldingMyTextPartInColumnX("T5", 10, 7, L(i))
            If zT5 = 0 Then Stop
        Club = .Cells(zT5, 7)
        LTV = .Cells(zT5, 8)
        Nation = .Cells(zT5, 9)
        
        
        If Nation = "D" Then
            OpenFolder p: Stop
'            LinkNameToDisplay = "- " + Club
'            JumpTo = ArrC(11) + "\D\" + LTV + "\" + Club
'            PathOfIcon = PathOfIcon = ArrC(4) + "\zzico\zzy_Club.ico"
'            Check_ClubLink_One p, LinkNameToDisplay, JumpTo, PathOfIcon
            LinkNameToDisplay = "-  " + LTV
            JumpTo = ArrC(11) + "\D\" + LTV
            PathOfIcon = ArrC(4) + "\zzico\zzy_LTV.ico"
            Check_LtvLink_One p, LinkNameToDisplay, JumpTo, PathOfIcon
            Stop: CloseFolder p
        End If
'            LinkNameToDisplay = "-  " + Nation
'            JumpTo = ArrC(11) + "\" + Nation
'            PathOfIcon = ArrC(4) + "\zzico\zzz_" + Nation + ".ico"
'                If Not FileExists(PathOfIcon) Then Stop
'            Check_NationLink_One p, LinkNameToDisplay, JumpTo, PathOfIcon
        
    Next
    'Finals
        End With
End Sub

Sub Check_LtvLink_One_TEST()
    Dim FolderToHoldLink$, LinkNameToDisplay$, JumpTo$, PathOfIcon$
    DoArrc
        FolderToHoldLink = ArrC(4) + "\Treiter, Kurt (TSG Mutterstadt)"
        LinkNameToDisplay = "-  Pfalz"
        JumpTo = ArrC(11) + "\D\Pfalz"
        PathOfIcon = ArrC(4) + "\zzico\zzy_LTV.ico"
    Check_LtvLink_One FolderToHoldLink, LinkNameToDisplay, JumpTo, PathOfIcon
End Sub

Sub Check_LtvLink_One(FolderToHoldLink$, LinkNameToDisplay$, JumpTo$, PathOfIcon$)
    Dim PL$
    'Path of link
        PL = FolderToHoldLink + "\" + LinkNameToDisplay + ".lnk" 'PathOfLink
    'Exist?
        If FileExists(PL) Then
        'Vbs?
            If Not IsVbsLink(PL) Then
                DeleteFile PL
                Create_OneVbsLink FolderToHoldLink, LinkNameToDisplay, JumpTo, PathOfIcon
            End If
        Else
            Create_OneVbsLink FolderToHoldLink, LinkNameToDisplay, JumpTo, PathOfIcon
        End If
End Sub

Sub Check_UsedIconJpg_All_EachOneInsidePersFold()
    Dim p$, i%, L() As String
    Call DoArrc
    Load_ArrNamesOfAllFolders_Leute L
    For i = 0 To UBound(L)
        p = ArrC(4) + "\" + L(i) 'Path of one LeuteFolder
        'OpenFolderGrosseSymbole p: OpenFolderGrosseSymbole p: OpenFolderGrosseSymbole p
        'WaitSecs 2 ': Stop
        Check_UsedIconJpg_One p
        'WaitSecs 2: OpenFolderGrosseSymbole p:
        'CloseFolder p: CloseFolder p: CloseFolder p
        DoEvents
    Next
    Beep
End Sub

Sub Check_UsedIconJpg_One(PathOfFolder$)
    'Called from    Check_UsedIconJpg_All_EachOneInsidePersFold
    
    'Vorbereitung
        Dim mw$, N1$, N2$, pM1$, pM2$, pW1$, pW2$, pU1$, pU2$
        DoArrc
        If Not PathOfFolder Like "*\Leute\*" Then Exit Sub
        
        N1 = getNameOfPath(PathOfFolder)            'NameOfFolder
        N2 = Replace(N1, ",", "§")                  'NameOfIco
        pM1 = ArrC(1) + "\ico\Face_Jpgs\zzm.jpg"    'Path of Jpg m to copy from
        pM2 = PathOfFolder + "\zz_m.jpg"            'Path of Jpg m to copy to
        pW1 = ArrC(1) + "\ico\Face_Jpgs\zzw.jpg"    'Path of Jpg w to copy from
        pW2 = PathOfFolder + "\zz_w.jpg"            'Path of Jpg w to copy to
        pU1 = ArrC(4) + "\zzico\zzx_u.jpg"          'Path of SquareUsedIcon to copy from
        pU2 = PathOfFolder + "\zz used.jpg"         'Path of SquareUsedIcon to copy to
    'Existiert das SquareUsedIcon im PersonFolder?
        If Not FileExists(pU2) Then CopyFile pU1, pU2
    'Existiert das FaceJpg im PersonFolder?
        If Not FileExists(PathOfFolder + "\zz_icon.jpg") Then
            'Es existiert noch kein icon zu VnNn    'm/w-Icon
            If Not (FileExists(pM2) Or FileExists(pW2)) Then
                mw = T5_Get_mw_Give_VnNn(Get_VnNn_From_NameOfPersFold(N1))
                If mw = "w" Then
                    CopyFile pW1, pW2
                Else
                    CopyFile pM1, pM2
                End If
            End If
        End If
End Sub


