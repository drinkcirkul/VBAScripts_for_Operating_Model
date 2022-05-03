Attribute VB_Name = "Adjustments_3SM_Monthly"
Sub Adjustments()
Attribute Adjustments.VB_ProcData.VB_Invoke_Func = " \n14"
    
    'Purpose: To update adjustment values of Revenue, COGS and Retained Earnings in the 3SM-Monthly Sheet
    'Created by: Vikas Singh on 28-June-2021
       
    'Key variables
    Dim i, n_months, nrow_Revenue, nrow_COGS, nrow_BSCheck, nrow_REAdj  As Integer
    Dim QBO_Revenue_Vals, QBO_COGS_Vals As Variant
    
    'Inputs - change here
    '------------------------------------------------------------------------------------------------------------------
    n_months = 44          '# of months for which you're runing the macro (e.g. 44 months - from 02/2018 to 09/2021)
    nrow_Revenue = 110     'Row number of 'Revenue, Total' in 3SM-Monthly sheet
    nrow_COGS = 127        'Row number of 'Cost of Goods Sold, Total' in 3SM-Monthly sheet
    nrow_BSCheck = 219     'Row number of 'Balance Sheet Check' in 3SM-Monthly sheet
    nrow_REAdj = 272       'Row 'Adjustments' in the Retainted Earnigns section int he Supporting Schedules
    '------------------------------------------------------------------------------------------------------------------
      
    QBO_Revenue_Vals = Worksheets("3SM-Monthly").Range(Cells(nrow_Revenue + 2, 3), Cells(nrow_Revenue + 2, 3 + n_months)).Value
    QBO_COGS_Vals = Worksheets("3SM-Monthly").Range(Cells(nrow_COGS + 2, 3), Cells(nrow_COGS + 2, 3 + n_months)).Value
    
    ' Goal seeking Revenue Numbers
    For i = 1 To n_months
        Worksheets("3SM-Monthly").Cells(nrow_Revenue, i + 2).GoalSeek Goal:=QBO_Revenue_Vals(1, i), ChangingCell:=Worksheets("3SM-Monthly").Cells(nrow_Revenue - 1, i + 2)
    Next i
    
    ' Goal seeking COGS Numbers
    For i = 1 To n_months
        Worksheets("3SM-Monthly").Cells(nrow_COGS, i + 2).GoalSeek Goal:=QBO_COGS_Vals(1, i), ChangingCell:=Worksheets("3SM-Monthly").Cells(nrow_COGS - 1, i + 2)
    Next i
    
    
    ' Goal seeking Retained Earnings Numbers until current month
    'For i = 1 To n_months
    '    Worksheets("3SM-Monthly").Cells(nrow_BSCheck, i + 2).GoalSeek Goal:=0, ChangingCell:=Worksheets("3SM-Monthly").Cells(nrow_REAdj, i + 2)
    'Next i
    
    'Goal seeking Retained Earnings Numbers for following months
    'For i = (n_months + 1) To 95
    '    Worksheets("3SM-Monthly").Cells(nrow_BSCheck, i + 2).GoalSeek Goal:=0, ChangingCell:=Worksheets("3SM-Monthly").Cells(nrow_REAdj, i + 2)
    'Next i
    
    
    'Finish message
    MsgBox ("Done!")
    
    
End Sub
