Attribute VB_Name = "Housekeeping"

Private Sub GoToA1()
'Purpose: Go to cell A1 of each sheet on open

    Dim ws As Worksheet
    ' Looping through all the sheets and setting the cell A1
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ws.Cells(1, 1) = "Cirkul, Inc."
        Application.Goto ws.Cells(1, 1), True
    Next

    ' Going to the summary tab
    Application.Goto ThisWorkbook.Sheets("Summary").Range("A1"), True

End Sub


Private Sub HideSupportingSheets()
    
    ThisWorkbook.Worksheets("Reference").Visible = False
    ThisWorkbook.Worksheets("ReadMe").Visible = False
    ThisWorkbook.Worksheets("ChangeLogs").Visible = False
    ThisWorkbook.Worksheets("Macros").Visible = False

    ' Going to the summary tab
    Application.Goto ThisWorkbook.Sheets("Summary").Range("A1"), True

End Sub

Private Sub ExportAsDist()
    
    'Setting all the sheets to A1
    GoToA1
    ThisWorkbook.Worksheets("Macros").Activate
   
    Dim masterWB As Workbook
    Set masterWB = Application.ThisWorkbook
    
    'Checking if user is on a PC or Mac
    Dim fileFormatCode
    fileFormatCode = Application.InputBox("For PC - enter 51" & vbCrLf & "For Mac - enter 52", "Are you on a PC or Mac?")
    Dim tgtPath As String
    
    'PC User
    If fileFormatCode = 51 Then
        tgtPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") + "\"
    End If
    
    'Mac User
    If fileFormatCode = 52 Then
        Dim macUserName As String
        macUserName = Environ("USER")
        tgtPath = "/Users/" + macUserName + "/Desktop/"
    End If
        
    
    'Message Box asking for export
    Dim Msg, Style, Title, Response
    Msg = "A distribution version of this file will be exported to your desktop. Do you want to proceed?"
    Style = vbOKCancel
    Title = "Export"
    Response = MsgBox(Msg, Style, Title)
    
    If Response = vbOK Then
        
        Application.DisplayAlerts = False
                    
        'Saving a copy of the current file (.xlsm) on desktop
        Dim nameText, timeStamp, newFileName, newWB As String
        nameText = "Cirkul Operating Model (Dist.) "
        timeStamp = Format(CStr(Now), "mm-dd-yyyy hh-mm AM/PM")
        newFileName = nameText + timeStamp
        newWB = tgtPath + newFileName + ".xlsm"
        ActiveWorkbook.SaveCopyAs (newWB)
        
        'Opening the new copy (.xlsm) and making changes to it
        Workbooks.Open (newWB)
        Dim distWB_xlsm As Workbook
        
        If fileFormatCode = 51 Then
            Set distWB_xlsm = Workbooks(newFileName)
        End If
        
        If fileFormatCode = 52 Then
            Set distWB_xlsm = Workbooks.Open(newWB)
        End If
        
        ' Deleting supporting sheets
        distWB_xlsm.Worksheets("Macros").Delete
        distWB_xlsm.Worksheets("ReadMe").Delete
        distWB_xlsm.Worksheets("ChangeLogs").Delete
        distWB_xlsm.Worksheets(1).Activate
        
        
        'Saving the new copy as normal excel file
        outputWB = tgtPath + newFileName + ".xlsx"
        distWB_xlsm.SaveAs fileName:=outputWB, FileFormat:=51  'This worked on both Windows & mac
        ActiveWorkbook.Close
                
        'Deleting the copied .xlsm file & returing to the current file
        output_xlsm_Path = tgtPath + newFileName + ".xlsm"
        Kill (output_xlsm_Path)
        masterWB.Worksheets("Macros").Activate
        
        'Final message
        Application.DisplayAlerts = True
        finalMsgBox = MsgBox("A distribution version of this file has been exported to your Desktop", vbOKOnly, "Success!")
            
    End If

    
End Sub
