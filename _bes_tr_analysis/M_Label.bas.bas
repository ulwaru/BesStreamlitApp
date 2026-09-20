Attribute VB_Name = "M_Label"
Option Explicit 'M_Label

Sub zzz_M_Label()

    showProcs "jpg ico"
    
    'RenameModule "Modul1", "M08"
    'T1_ShowLogBuch
    'Sortier-Reihenfolge Explorer:
        '!#$%&(),.'-;[]°^_`{}~´+=§°µ0123aAbBcC
        'zaz|zA a|zAa
    EE 1: Beep
End Sub

Sub Create_WhiteInJpg_ToTransparentInPng_TEST()
    'Called from    xxx

    'magick input.jpg -transparent white output.png
    
    'Vorbereitung
        Dim p1$, p2$, pFront$, pIN$, pOUT$, pTool$, qq$, sCmd$, shResult%
        qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        p1 = ArrC(1) + "\ico\WheatSquares\1.jpg"
        p2 = ArrC(1) + "\ico\WheatSquares\2.png"
        
        pIN = " " + qq + p1 + qq
        pOUT = " " + qq + p2 + qq
    'Create
        sCmd = pTool + pIN + " -transparent white " + pOUT
        shResult = ShellAndWait(sCmd, 0, vbHide, PromptUser)
End Sub

Sub ResizeImage_TEST()
        'Called from    xxx
        'magick input.jpg -transparent white output.png
    
    'Vorbereitung
        Dim pFront$, pIN$, pOUT$, pTool$, qq$, sCmd$, shResult%
        qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        
        pIN = " " + qq + ArrC(1) + "\ico\WheatSquares\2.png" + qq
        pOUT = " " + qq + ArrC(1) + "\ico\WheatSquares\3.png" + qq
    'Create
        sCmd = pTool + pIN + " -resize 500x " + pOUT
        shResult = ShellAndWait(sCmd, 0, vbHide, PromptUser)
End Sub

Sub Create_JpgWith_FrontPng_BackJpg_TEST()
    'Called from    xxx
    
    'Vorbereitung
        Dim pFront$, pBack$, pOUT$, pTool$, qq$, sCmd$, shResult%
        qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        pBack = " " + qq + ArrC(1) + "\ico\WheatSquares\Icon.jpg" + qq
        pFront = " " + qq + ArrC(1) + "\ico\WheatSquares\used.png" + qq
        pOUT = " " + qq + ArrC(1) + "\ico\WheatSquares\4.jpg" + qq
    'Create
        sCmd = pTool + "composite -gravity center -geometry +5+80 " + pFront + pBack + pOUT
        shResult = ShellAndWait(sCmd, 0, vbHide, PromptUser)
End Sub

Sub Create_SomeWheatSquares()
    'Create_WheatSquareWithText "Event", 300, 300
    'Create_WheatSquareWithText "Video", 300, 300
    'Create_WheatSquareWithText "Quelle", 300, 300
    'Create_JpgArrow_L_R_U_D "Icon", 300, 300, "R", 600, 0, -200
    'Create_JpgArrow_L_R_U_D "Club", 300, 300, "U", 600, 0, -100
    'Create_JpgArrow_L_R_U_D "LTV", 300, 300, "U", 600, 0, -100
    Create_JpgArrow_L_R_U_D "Event", 300, 300, "U", 600, 0, -100
    'Create_JpgArrow_L_R_U_D "GB", 300, 300, "U", 600, 0, -100
    'Create_JpgArrow_L_R_U_D "F", 300, 300, "U", 600, 0, -100
    'Create_JpgArrow_L_R_U_D "USA", 300, 300, "U", 600, 0, -100
    'Create_JpgArrow_L_R_U_D "Leute", 300, 300, "R", 600, 0, -100
End Sub

Sub Create_WheatSquareWithText(text$, FontSize%, CenterPlusY%)
    'Called from    xxx
    
    Create_Square_wheat_1000x1000
    Create_TextBlack_BackTransparent text, FontSize
    Create_Jpg_FrontPng_BackJpg text, CenterPlusY
    DeleteFile ArrC(1) + "\ico\WheatSquares\" + text + ".png"
    DeleteFile ArrC(1) + "\ico\WheatSquares\wheat.jpg"
End Sub

Sub Create_Jpg_FrontPng_BackJpg(text$, CenterPlusY%)
    'Called from    xxx
    
    'Vorbereitung
        Dim pFront$, pBack$, pOUT$, pTool$, qq$, sCmd$, shResult%
        qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        pBack = " " + qq + ArrC(1) + "\ico\WheatSquares\wheat.jpg" + qq
        pFront = " " + qq + ArrC(1) + "\ico\WheatSquares\" + text + ".png" + qq
        pOUT = " " + qq + ArrC(1) + "\ico\WheatSquares\" + text + ".jpg" + qq
    'Create
        sCmd = pTool + "composite -gravity center -geometry +0+" + CStr(CenterPlusY) + pFront + pBack + pOUT
        shResult = ShellAndWait(sCmd, 0, vbHide, PromptUser)
End Sub

Sub CopyToFolder_WheatSquareJpg_Leute(PathOfFolder$)
    'Called from    Repair_VbsLinks_OneFolder_TEST
    'Action         kopiert, falls nötig, das Leute-WheatSquareJpg in den PathOfFolder
    
    'Vorbereitung
        Dim p1$, p2$
        If PathOfFolder Like "*\Leute\*" Then Exit Sub
        If PathOfFolder Like "*\Faces" Then Exit Sub
        If PathOfFolder Like "*\Register\*" Then Exit Sub
    'Action
        p1 = ArrC(4) + "\zzico\zzx_Leute.jpg"
        p2 = PathOfFolder + "\a Leute.jpg"
        If Not FileExists(p2) Then CopyFile p1, p2
End Sub

Sub Create_WheatSquareJpgAndIcon_Nation_TEST()
    Create_WheatSquareJpgAndIcon_Nation "Event"
End Sub

Sub Create_WheatSquareJpgAndIcon_Nation(NationShorty$)
    'Called from    xxx
    
    'Vorbereitung
        Dim pFront$, pBack$, pOUT$, pTool$, qq$, sCmd$, shResult%
    
        qq = Chr(34): DoArrc
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        pBack = " " + qq + ArrC(1) + "\ico\WheatSquares\" + NationShorty + ".jpg" + qq
        pFront = " " + qq + ArrC(1) + "\ico\WheatSquares\Arrow.png" + qq
        pOUT = " " + qq + ArrC(1) + "\ico\WheatSquares\zzz_" + NationShorty + ".jpg" + qq
    'Pfeil-png erstellen
        Create_PngArrow_L_R_U_D "U", 600
    'TextSquare erstellen
        Create_WheatSquareWithText NationShorty, 300, 300
    'Create ArrowRightJpg
        sCmd = pTool + "composite -gravity center -geometry +" _
            + CStr(0) + "+" + CStr(-100) + pFront + pBack + pOUT
        shResult = ShellAndWait(sCmd, 0, vbHide, PromptUser)
    'Delete
        DeleteFile ArrC(1) + "\ico\WheatSquares\Arrow.png"
        DeleteFile ArrC(1) + "\ico\WheatSquares\" + NationShorty + ".jpg"
    'Icon
        Dim p1$, p2$, p3$
        p1 = ArrC(1) + "\ico\WheatSquares\zzz_" + NationShorty + ".jpg"
        p2 = Replace(p1, ".jpg", ".ico")
        p3 = ArrC(4) + "\zzico\zzz_" + NationShorty + ".ico"
        Create_IconOfOneJpg p1, p2
        CopyFile p2, p3
    End Sub

Sub Create_JpgArrow_L_R_U_D(text$, TextSize%, TextPlusY%, LRUD$, ArrowSize%, ArrowPlusX%, ArrowPlusY%)
    'Called from    Create_SomeWheatSquares
    
    'Vorbereitung
        Dim pFront$, pBack$, pOUT$, pTool$, qq$, sCmd$, shResult%
        qq = Chr(34): DoArrc
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        pBack = " " + qq + ArrC(1) + "\ico\WheatSquares\" + text + ".jpg" + qq
        pFront = " " + qq + ArrC(1) + "\ico\WheatSquares\Arrow.png" + qq
        pOUT = " " + qq + ArrC(1) + "\ico\WheatSquares\" + text + ".jpg" + qq
    'Pfeil-png erstellen
        Create_PngArrow_L_R_U_D LRUD$, ArrowSize
    'TextSquare erstellen
        Create_WheatSquareWithText text, TextSize, TextPlusY
    'Create ArrowRightJpg
        sCmd = pTool + "composite -gravity center -geometry +" _
            + CStr(ArrowPlusX) + "+" + CStr(ArrowPlusY) + pFront + pBack + pOUT
        shResult = ShellAndWait(sCmd, 0, vbHide, PromptUser)
    'Delete
        DeleteFile ArrC(1) + "\ico\WheatSquares\Arrow.png"
        'DeleteFile ArrC(1) + "\ico\WheatSquares\" + Text + ".jpg"
End Sub

Sub Create_PngArrow_L_R_U_D(LRUD_LeftRightUpDown$, ArrowSize%)
    'Called from    xxx
    'Arrow          (Left|Right|Up|Down) = Wingdings (ï|ð|ñ|ò)
    'Png            = .png-File, transparent
    'Action         Erstellt xx
    
    'Vorbereitung
        Dim A$, B$, pOUT$, pTool$, qq$, sz$, sCmd$, shResult%
        qq = Chr(34): DoArrc
        A = LRUD_LeftRightUpDown: B = Mid("ïðñò", InStr(1, "LRUD", A), 1)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        pOUT = " " + qq + ArrC(1) + "\ico\WheatSquares\Arrow.png" + qq
        sz = " " + CStr(ArrowSize) + " "
    'Create Text
        sCmd = pTool + "convert -background none -fill black -font Wingdings " _
               + "-pointsize" + sz + "label:" + qq + B + qq + pOUT
        shResult = ShellAndWait(sCmd, 0, vbHide, PromptUser)
End Sub

Sub Create_TextBlack_BackTransparent_TEST()
    Create_TextBlack_BackTransparent "used", 200
End Sub

Sub Create_TextBlack_BackTransparent(text$, FontSize%)
    'Called from    Create_WheatSquareWithText Create_TextBlack_BackTransparent_TEST
    '               z. B.: Create_TextBlack_BackTransparent "Hallo", 60
    'Action         Erstellt xx
    
    'Vorbereitung
        Dim pOUT$, pTool$, qq$, sz$, sCmd$, shResult%
        qq = Chr(34): DoArrc
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        pOUT = " " + qq + ArrC(1) + "\ico\WheatSquares\" + text + ".png" + qq
        sz = " " + CStr(FontSize) + " "
    'Create Text
        sCmd = pTool + "convert -background none -fill black -font Arial-Bold " _
               + "-pointsize" + sz + "label:" + qq + text + qq + pOUT
        shResult = ShellAndWait(sCmd, 0, vbHide, PromptUser)
End Sub

Sub Create_Square_wheat_1000x1000()
    'Called from    xxx
    'Action         Erstellt BackgroundImage im Ordner 'squares', Farbe = wheat, 500x500
    
    'Vorbereitung
        Dim PathOfFolderSquares$, pOUT$, pTool$
        Dim qq$, sCmd$, shResult%
        qq = Chr(34): DoArrc
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        PathOfFolderSquares = ArrC(1) + "\ico\WheatSquares"
        pOUT = " " + qq + PathOfFolderSquares + "\wheat.jpg" + qq
    'Create BackgroundImage
        sCmd = pTool + "convert -size 1000x1000 xc:wheat" + pOUT
        shResult = ShellAndWait(sCmd, 0, vbHide, PromptUser)
End Sub

Sub Hänge_jpg2_unter_jpg1(p1$, p2$, p3$)
    'Called from    xxx

    'gm: GraphicsMagick
    Dim shellCmd$, qq$, shResult%
    qq = Chr(34) 'quote, "-Zeichen
    'gm convert "p1" "p2" -append "p3"
    shellCmd = "gm convert " + qq + p1 + qq + " " + qq + p2 + qq + " -append " + qq + p3 + qq
    'Shell (shellCmd)
    shResult = ShellAndWait(shellCmd, 0, vbHide, PromptUser)
End Sub

Sub T4_LabelForCountDown_Set(z%, s%)
    'Called from    T6_Update_EachResultJpg
    'Action         zeigt transparentes Label mit großer grüner CountDown-Nr
    '               rechts unterhalb der T4-Zelle (z, s)
    
    'Vorbereitung
        Dim N$, NrToDisplay%, L!, T!, lbl As MSForms.Label
        N = "LabelForCountDown" 'NameOfShape
        NrToDisplay = CInt(ArrC(25))     'T0 MyCountDown RunDownNr
    'Left/Top of LabelForCountDown
        With Sheets("T4"): L = .Cells(z, s).Left + 70: T = .Cells(z, s).Top + 15: End With
    'Create if Label does not exist
        If Not T4_LabelForCountDown_Exists Then T4_LabelForCountDown_Create
    'Lbl Shape
        With Sheets("T4").Shapes(N): .Left = L: .Top = T: .Visible = msoTrue: End With
    'Lbl Object
        Set lbl = ActiveSheet.OLEObjects(N).Object
        lbl.Caption = CStr(NrToDisplay)
End Sub

Sub T4_LabelForCountDown_Hide()
    'Sheets("T4").Shapes("LabelForCountDown").Visible = msoFalse
End Sub

Sub T4_LabelForCountDown_Show()
    'Sheets("T4").Shapes("LabelForCountDown").Visible = msoTrue
End Sub

Sub T4_LabelForCountDown_Create()
    'Called from    xxx

    Dim L!, T!, lbl As OLEObject
    Set lbl = Sheets("T4").OLEObjects.Add(ClassType:="Forms.Label.1", _
        Link:=False, DisplayAsIcon:=False, Left:=0, Top:=9999, Width:=120, Height:=100)
    With lbl.Object: .BackStyle = 0: .Caption = "333"
        .Font.size = 74: .Font.Bold = True: .ForeColor = Green4: End With
    With lbl: .NAME = "LabelForCountDown"
        .ShapeRange.Fill.Transparency = 1#: .Visible = True: End With 'Lbl.Delete
End Sub

Function T4_LabelForCountDown_Exists() As Boolean
    'Called from    xxx

    On Error GoTo J1
    With Sheets("T4").Shapes("LabelForCountDown"): .Left = .Left + 10: End With
    T4_LabelForCountDown_Exists = True: Exit Function
J1: On Error GoTo 0
End Function


Sub LabelTEST()
    'Called from    xxx

    Dim Sht  As Worksheet
    Dim MyLbl As OLEObject
    
    'set the worksheet object
        Set Sht = ThisWorkbook.Worksheets("T4")
    'set the Active-X label object
        Set MyLbl = Sht.OLEObjects("Label1")
    'change the Caption to the value in Range "F4")
        MyLbl.Object.Caption = Sht.Range("rv11").Value
        
        MyLbl.Left = [sa8].Left
        MyLbl.Top = [sa8].Top
        MyLbl.Height = 10
End Sub

Sub Create_i15(Id$)
    'Called from    xxx

    'Vorbereitung
        Dim p$, USEp1$, USEp2$, USEp3$
        p = ArrC(1) + "\prog\Label\LbProd\" + Id + "_"
        USEp1 = p + "i14.jpg"           'F1 = Foto (1000xH)             .jpg
        USEp2 = p + "i13.jpg"           'FertigesLabel                  .jpg
        USEp3 = p + "i15.jpg"           'FertigesLabel unterhalb Foto   .jpg
    'Create
        Add_4Images_BelowEachOther USEp1, USEp2, "", "", USEp3
End Sub

Sub Create_i14(Id$, pF0$)
    'Called from    Produce_T6PhotoWithLabel
    'F1             aktuell ausgewähltes Foto mit Width=1000
    'Action         Erstellt F1 im Ordner 'LbProd'
    
    'Vorbereitung
        Dim p$, pIN$, pOUT$, pTool$, qq$, sCmd$, USEp1$, USEp2$
        qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        p = ArrC(1) + "\prog\Label\LbProd\" + Id + "_"
        USEp1 = pF0 'Get_T6EvPathOfShowPhoto
        USEp2 = p + "i14.jpg"               'F1 = Foto (1000xHeight)
        pIN = " " + qq + USEp1 + qq: pOUT = " " + qq + USEp2 + qq
    'Create
        sCmd = pTool + "convert" + pIN + " -resize 1000x" + pOUT
        ShellAndWaitWithFlashingCmdWindow sCmd
End Sub

Sub Create_i13(Id$, Zi%)
    'Called from    Produce_T6PhotoWithLabel
    'Zi             = Anzahl Zeilen
    
    'Vorbereitung
        Dim geo$, p$, pFront$, pBack$, LbSize$, pOUT$, pTool$, qq$, sCmd$, USEp1$, USEp2$, USEp3$
        qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        p = ArrC(1) + "\prog\Label\LbProd\" + Id + "_"
        USEp1 = p + "i12.jpg"               'BACK   LNY+Age
        USEp2 = p + "i08.png"               'FRONT  TextEndversion (i04+...+i07)
        USEp3 = p + "i13.jpg"               'OUT    LNYA+TextEndversion
        pBack = " " + qq + USEp1 + qq: pFront = " " + qq + USEp2 + qq: pOUT = " " + qq + USEp3 + qq
    'Falls i08 nicht existiert, dann i12
        If FileExists(USEp2) Then
            'Create
                If Zi = 1 Then geo = "+13+67"
                If Zi = 2 Then geo = "+13+59"
                If Zi = 3 Then geo = "+13+49"
                If Zi = 4 Then geo = "+13+24"
                sCmd = pTool + "composite -gravity south -geometry " + geo + pFront + pBack + pOUT
                ShellAndWaitWithFlashingCmdWindow sCmd 'Call Shell(sCmd)
        Else
            CopyFile USEp1, USEp3
        End If
End Sub

Sub Create_i12(Id$)
    'Called from    xxx

    'Vorbereitung
        Dim p$, pFront$, pBack$, LbSize$, pOUT$, pTool$, qq$, sCmd$, USEp1$, USEp2$, USEp3$
        qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        p = ArrC(1) + "\prog\Label\LbProd\" + Id + "_"
        USEp1 = p + "i11.jpg"               'BACK   LN+Year
        USEp2 = p + "i03.png"               'FRONT  Age
        USEp3 = p + "i12.jpg"               'OUT    LNY+Age
        pBack = " " + qq + USEp1 + qq: pFront = " " + qq + USEp2 + qq: pOUT = " " + qq + USEp3 + qq
    'Create
        sCmd = pTool + "composite -gravity east -geometry +50+40" + pFront + pBack + pOUT
        ShellAndWaitWithFlashingCmdWindow sCmd
End Sub

Sub Create_i11(Id$)
    'Called from    xxx

    'Vorbereitung
        Dim p$, pFront$, pBack$, LbSize$, pOUT$, pTool$, qq$, sCmd$, USEp1$, USEp2$, USEp3$
        qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        p = ArrC(1) + "\prog\Label\LbProd\" + Id + "_"
        USEp1 = p + "i10.jpg"               'BACK   L+Name
        USEp2 = p + "i02.png"               'FRONT  Year
        USEp3 = p + "i11.jpg"               'OUT    LN+Year
        pBack = " " + qq + p + "i10.jpg" + qq
        pFront = " " + qq + p + "i02.png" + qq      'Jahr
        pOUT = " " + qq + USEp3 + qq
    'Create
        sCmd = pTool + "composite -gravity west -geometry +50+40" + pFront + pBack + pOUT
        ShellAndWaitWithFlashingCmdWindow sCmd
End Sub

Sub Create_i10(Id$)
    'Called from    xxx

    'Vorbereitung
        Dim p$, pFront$, pBack$, LbSize$, pOUT$, pTool$, qq$, sCmd$
        qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        p = ArrC(1) + "\prog\Label\LbProd\" + Id + "_"
        pBack = " " + qq + p + "i09.jpg" + qq       'LabelLeerBackgroundWheat
        pFront = " " + qq + p + "i01.png" + qq      'VornameNachname
        pOUT = " " + qq + p + "i10.jpg" + qq
    'Create
        sCmd = pTool + "composite -gravity North -geometry +0+25" + pFront + pBack + pOUT
        ShellAndWaitWithFlashingCmdWindow sCmd
End Sub

Sub Create_i09(Id$)
    'Called from    Produce_T6PhotoWithLabel
    'Action         Erstellt BackgroundImage, Farbe = wheat, 1000x250
    
    'Vorbereitung
        Dim p1$, pOUT$, pTool$, qq$, sCmd$: qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        p1 = ArrC(1) + "\prog\Label\LbProd\" + Id + "_i09.jpg"
        pOUT = " " + qq + p1 + qq
    'Create
        sCmd = pTool + "convert -size 1000x250 xc:wheat" + pOUT
        ShellAndWaitWithFlashingCmdWindow sCmd
End Sub

Sub Create_i08(Id$)
    'Called from    Produce_T6PhotoWithLabel
    'Action         Erstellt TextEndversion (1-4 Zeilen), background transparent
    
    Dim p$, p1$, p2$, p3$, p4$, pOUT$
    p = ArrC(1) + "\prog\Label\LbProd\": pOUT = p + Id + "_i08.png"
    p1 = p + Id + "_i04.png": p2 = p + Id + "_i05.png"
    p3 = p + Id + "_i06.png": p4 = p + Id + "_i07.png"
    If Not FileExists(p1) Then p1 = ""
    If Not FileExists(p2) Then p2 = ""
    If Not FileExists(p3) Then p3 = ""
    If Not FileExists(p4) Then p4 = ""
    Add_4Images_BelowEachOther p1, p2, p3, p4, pOUT
End Sub

Sub Create_i07(Id$, OneLine$, iOUT$): Create_i04 Id, OneLine, iOUT: End Sub
Sub Create_i06(Id$, OneLine$, iOUT$): Create_i04 Id, OneLine, iOUT: End Sub
Sub Create_i05(Id$, OneLine$, iOUT$): Create_i04 Id, OneLine, iOUT: End Sub

Sub Create_i04(Id$, OneLine$, iOUT$)
    'Called from    Produce_T6PhotoWithLabel
    
    'Vorbereitung
        If OneLine = "" Then Exit Sub
        Dim LbText$, LbSize$, p1$, pOUT$, pTool$, qq$, sCmd$
        qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        p1 = ArrC(1) + "\prog\Label\LbProd\" + Id + "_" + iOUT + ".png"
        pOUT = " " + qq + p1 + qq
        LbText = qq + OneLine + qq
    'Create
        sCmd = pTool + "convert -background none -fill blue -font Arial " _
               + "-pointsize 24 -size 630x30 -gravity center label:" + LbText + pOUT
        ShellAndWaitWithFlashingCmdWindow sCmd
End Sub

Sub Create_i03(Id$, Age$)
    'Called from    Produce_T6PhotoWithLabel
    
    'Vorbereitung
        Dim LbText$, LbSize$, p1$, pOUT$, pTool$, qq$, sCmd$
        qq = Chr(34): If Age = "" Then Age = "?"
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        p1 = ArrC(1) + "\prog\Label\LbProd\" + Id + "_i03.png"
        pOUT = " " + qq + p1 + qq
        LbText = qq + "(" + Age + ")" + qq
        'LbSize = "1000x100" 'Breite x Höhe
    'Create
        sCmd = pTool + "convert -background none -fill black -font Arial-Bold " _
               + "-pointsize 60 label:" + LbText + pOUT
        ShellAndWaitWithFlashingCmdWindow sCmd
End Sub

Sub Create_i02(Id$, Year$)
    'Called from    Produce_T6PhotoWithLabel
    
    'Vorbereitung
        Dim LbText$, LbSize$, p1$, pOUT$, pTool$, qq$, sCmd$
        qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        p1 = ArrC(1) + "\prog\Label\LbProd\" + Id + "_i02.png"
        pOUT = " " + qq + p1 + qq
        LbText = qq + Year + qq
    'Create
        sCmd = pTool + "convert -background none -fill black -font Arial-Bold " _
               + "-pointsize 60 label:" + LbText + pOUT
        ShellAndWaitWithFlashingCmdWindow sCmd
End Sub

Sub Create_i01(Id$, NAME$)
    'Called from    Produce_T6PhotoWithLabel
    
    'Vorbereitung
        Dim p1$, pOUT$, pTool$, qq$, sCmd$: qq = Chr(34)
        pTool = qq + "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe" + qq + " "
        p1 = ArrC(1) + "\prog\Label\LbProd\" + Id + "_i01.png"
        pOUT = " " + qq + p1 + qq
    'Create
        sCmd = pTool + "convert -background none -fill black -font Arial-Bold " _
               + "-size 1000x78 -pointsize 66.7 -gravity north label:" + qq + NAME + qq + pOUT
        ShellAndWaitWithFlashingCmdWindow sCmd
End Sub



