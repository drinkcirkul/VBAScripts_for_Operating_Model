Attribute VB_Name = "DTCSales_Revenue_RetentionPer"
Sub DTCSales_Rev_Retetnion_Per_Update()
Attribute DTCSales_Rev_Retetnion_Per_Update.VB_ProcData.VB_Invoke_Func = " \n14"
    'Purpose: To update the section B - % Revenue Retention of the DTC Sales Sheet

    Dim strtRow, strtCol, n_months, n_totalCols, i As Integer
    'Dim n_months As Integer
    
    strtRow = 109  'Header row of section B where it says Month, Cohort etc.
    strtCol = 7    ' % retention numbers start from column G
    n_months = 50  '2/2018 to 3/2022 are 50 months
    n_totalCols = 101
    
    Worksheets("DTC Sales").Activate
    
    'Removing existing borders
    Range(Cells(strtRow, 1), Cells(strtRow + n_months, n_months + 6)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    'Filling down
    For i = 1 To n_months
        Range(Cells(strtRow + n_months - i, strtCol + i - 1), Cells(strtRow + n_months - i + 1, strtCol + i - 1)).FillDown
    Next i
    
    'Filling last data point
    Range(Cells(strtRow + 1, strtCol + n_months - 2), Cells(strtRow + 1, strtCol + n_months - 1)).FillRight
    
    'Applying borders to the waterfall
    For i = 1 To n_months
        Cells(strtRow + n_months - i + 1, strtCol + i - 1).Select
            With Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
    
            With Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
    Next i
    
    'Adding a bottom border
    Range(Cells(strtRow + n_months, 1), Cells(strtRow + n_months, 6)).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
     
    Application.CutCopyMode = False
    Cells(strtRow + 1, 2).Select
    
    MsgBox ("Done!")
     
End Sub