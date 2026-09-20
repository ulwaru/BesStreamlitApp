Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Option Explicit 'DieseArbeitsmappe

Sub zzz_DieseArbeitsmappe()
    Application.EnableEvents = True
    
    showProcs "holding"
    
    Beep
End Sub

Sub MyToDoList()
    'Get_OneNameItems_DgOnly
    'T4_Load_ArrDgTitles_FromT4ListOfDesignTitles
    'T4_Sort_Paste_Format_ListOfDesignTitles
    'T4_Write_ListOfDesignTitles_FromDgs
    'T8_Synchronize_T8SomeDgData_With_T4ListOfDesignTitles
    'No CharCheck please
    'PersonDataTxt leer
    'Kontrolle:  Get_Nation_From_OrdnerName  Get_Rufname_From_OrdnerName
    '            Get_Verein_From_OrdnerName  GetTR_rn

    '"Cells(8, 5)" ändern
    'ArrC(59) weg
    'ShowMissingVIDs_toFileName
    'Update_Links_inside_FolderLeute_insideOneEventFolder
    'Update Links in Folder 'Links' --> Update Statistics
    'vbScriptAufrufen
    'ua kürzel
    'Kill? Add_OneLineToPersonData Read_Pd Load_ArrPD5

End Sub

Private Sub Workbook_Open()
    Dim s$
    'F4 mit dem Makro verknüpfen
        Application.OnKey "{F3}", "RunShowProcs"
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    ActivateT1
    'ThisWorkbook.path
        s = ThisWorkbook.path
        If s Like "*.xlsb*" Then
            MsgBox "Dies ist nicht die Originaldatei 'Bes TR.xlsm'" + vbCrLf _
                 + "('" + ThisWorkbook.NAME + "')"
        End If
    'ArrC aktivieren
        DoArrc
    'Pfade aktualisieren
        Fill_ArrCPaths
    'Roaming
        If s Like "*Roaming*" Then
            LogBuch "ThisWorkbook.path was changed from '" + s + "' to '" + ArrC(1) + "'"
        End If
    LogBuch vbCrLf + "Workbook 'Bes TR.xlsm' was opened"
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'Beim Schließen der Datei die Standard-Funktion von F4 wiederherstellen
        Application.OnKey "{F4}"
    LogBuch "Workbook 'Bes TR.xlsm' was closed"
End Sub

Sub ActivateT1()
    Dim s%, z%
    With Worksheets("T1")
        .Activate
        Worksheets("T1").Select
        z = Get_T1Row_HoldingMyText(" Code rechts")
        s = Get_T1Column_HoldingMyText(" Code rechts")
        If .Cells(z, s - 1) = "X" Then .Cells(z, s - 1) = "" Else .Cells(z, s - 1) = "X"
        XBox_CodeRechts z, s - 1
    End With
End Sub

Sub StartVBS_toCreateTextFile_HoldingAllFilesInsideFolderEvents()
    DeleteFile ArrC(6) + "\AllFilesInsideFolderEvents.txt"
    vbScriptAufrufen ArrC(6) + "\vbs\FilesInEvents.vbs"
End Sub

Sub StartVBS_toCreateTextFile_HoldingAllFilesInsideFolderLeute()
    DeleteFile ArrC(6) + "\AllFilesInsideFolderLeute.txt"
    vbScriptAufrufen ArrC(6) + "\vbs\FilesInLeute.vbs"
End Sub

