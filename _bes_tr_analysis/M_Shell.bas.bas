Attribute VB_Name = "M_Shell"
Option Explicit 'M_Shell

Sub zzz_M_Shell()
    Application.EnableEvents = True
    'showProcs "files"
    Beep
End Sub

Sub renameJpgs()
    'Vorbereitung
        Dim FileNamesLike$, pFolder$, p1$, p2$, s$
        Dim i%, subfolders%, Arr1() As String
        pFolder = "F:\Archiv TR\Lieferungen von\20240922 von ik Ivonne Kraft\" _
                    + "Ordner 01 Bilderbuch1"
        subfolders = 0
        FileNamesLike = "* c##.jpg"
    'Get list PathOfMyFiles
        s = Get_PathOfMyFiles(pFolder, subfolders, FileNamesLike)
        show s
    'Rename
        Arr1 = Split(s, vbCrLf)
        For i = 0 To UBound(Arr1)
            p1 = Arr1(i)
            'p2 = Replace(p1, "a.jpg", "a PhotoFromA4Page.jpg")
            'p2 = Replace(p1, "b.jpg", "b Scan A4 600 dpi.jpg")
            'p2 = Replace(p1, "c0 300dpi1000px.jpg", "c SmallFromA4Scan.jpg")
             p2 = Replace(p1, " c", " d ")
            RenameFile p1, p2
            DoEvents
        Next
        Beep
End Sub

Sub Create_Jpg300Dpi1000Px_FromBigA4Jpg()
    'Vorbereitung
        Dim FileNamesLike$, pFolder$, pIN$, pOUT$, pTool$, s$, shellCmd$
        Dim i%, subfolders%, Arr1() As String
        pTool = "C:\Program Files\ImageMagick-7.1.0-Q16-HDRI\magick.exe"
    
    'In welchem folder sollen files gesucht werden
        pFolder = "F:\Archiv TR\Lieferungen von\20240922 von ik Ivonne Kraft\Ordner 01 Bilderbuch1"
    'Auch in dessen SubFolders? 0 = Nein, 1 = Ja
        subfolders = 0
    'Für welche files sollen die Pfade gelistet werden?
        FileNamesLike = "* b *.jpg"
    'Get list PathOfMyFiles
        s = Get_PathOfMyFiles(pFolder, subfolders, FileNamesLike)
        'show s
        If s = "" Then Beep: Exit Sub
    'Create jpgs 300 dpi, 1000 px
        Arr1 = Split(s, vbCrLf)
        For i = 0 To UBound(Arr1)
            pIN = Arr1(i)
            pOUT = Replace(pIN, " b A4Scan600dpi.jpg", " c SmallFromA4Scan.jpg")
            If Not FileExists(pOUT) Then
                shellCmd = Chr(34) + pTool + Chr(34) + " convert -density 300 -resize 1000 " _
                            + Chr(34) + pIN + Chr(34) + " " + Chr(34) + pOUT + Chr(34)
                '        = pTool convert -density 300 -resize 1000 pIN pOUT
                Call shell(shellCmd)
                'wait
                    Do
                        If Not FileExists(pOUT) Then
                            DoEvents
                        Else
                            Call Beep: Exit Do
                        End If
                    Loop
            End If
        Next
        Beep
End Sub

