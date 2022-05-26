Attribute VB_Name = "Scenarios_DTC_Retention"
Sub Scenarios_DTC_Retention()
    
    'Purpose: Update borders in the retention charts
    Dim tgtWB As Workbook
    Dim fileName As String
    Dim strtRow_LiveCase, strtRow_BaseCase, strtRow_UpCase, strtRow_DownCase, strtRow_DTC_Sales, n_months, i As Integer
    Dim strtCol_DTC_Sales, strtCol_Scenario_DTC_Retention As Integer
    
    
    fileName = "Cirkul Operating Model (Live).xlsx"
    Set tgtWB = Workbooks(fileName)
    n_months = 51  'Feb-18 to April-2022 are 51 months
    
    strtRow_LiveCase = 14
    strtRow_BaseCase = 112
    strtRow_UpCase = 210
    strtRow_DownCase = 308
    strtCol_Scenario_DTC_Retention = 3
    
    'Base Case - Filling down
    tgtWB.Worksheets("Scenarios_DTC-Retention").Activate
    
    For i = 1 To n_months
        Range(Cells(strtRow_BaseCase, strtCol_Scenario_DTC_Retention + i - 1), Cells(strtRow_BaseCase + n_months - i, strtCol_Scenario_DTC_Retention + i - 1)).FillDown
    Next i
    
    'Base Case Filling last data point
    Range(Cells(strtRow_BaseCase, strtCol_Scenario_DTC_Retention + n_months - 2), Cells(strtRow_BaseCase, strtCol_Scenario_DTC_Retention + n_months - 1)).FillRight
    
    'BORDERS - Scenarios DTC Retention
    Call ind_Table_Borders("Scenarios_DTC-Retention", strtRow_LiveCase, strtCol_Scenario_DTC_Retention, n_months)
    Call ind_Table_Borders("Scenarios_DTC-Retention", strtRow_BaseCase, strtCol_Scenario_DTC_Retention, n_months)
    Call ind_Table_Borders("Scenarios_DTC-Retention", strtRow_UpCase, strtCol_Scenario_DTC_Retention, n_months)
    Call ind_Table_Borders("Scenarios_DTC-Retention", strtRow_DownCase, strtCol_Scenario_DTC_Retention, n_months)
         
    MsgBox ("Done!")
End Sub

Sub ind_Table_Borders(sheetName, strtRow, strtCol, n_months)
    
    'Purpose: Applies border to a individual tables
    
    Worksheets(sheetName).Activate
    
    'Removing existing borders
    Range(Cells(strtRow, 1), Cells(strtRow + n_months, n_months + 3)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    'Applying borders under current month
    Range(Cells(strtRow + n_months - 1, 1), Cells(strtRow + n_months - 1, strtCol - 1)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    'Applying borders to the waterfall
    For i = 0 To (n_months - 1)
        Cells(strtRow + n_months - 1 - i, strtCol + i).Select
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
    Next i

    
End Sub


