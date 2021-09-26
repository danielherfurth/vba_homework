Option Explicit

Sub get_uniques(source_rng As Range, target_rng As Range)

'Using advance filter for extacting unique items in the source range
source_rng.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=target_rng, Unique:=True
        
End Sub


Sub call_uniques()
Dim sh As Worksheet
Dim wb As Workbook

For Each sh In ThisWorkbook.Worksheets

'Calling get_uniques macro
Call get_uniques(sh.Range("A:A"), sh.Range("J1"))

' setting header values
sh.Range("K1") = "<annual_diff>"
sh.Range("L1") = "<pct_chng>"
sh.Range("M1") = "<annual_vol>"

Next sh

End Sub

Sub do_the_calcs()

On Error Resume Next
' i is loop iterator
' iter is a tracker of how many rows in range _
    of each ticker
' end_ : where range ends
' beg_ : where range begins
' vol is total stock volume per ticker
Dim i As Long, last_cell As Long, _
end_ As Long, beg_ As Long, _
vol As Long, iter As Long

Dim sh As Worksheet

' percent change and annual difference
Dim pct_change, diff

Call start


For Each sh In ThisWorkbook.Worksheets
    i = 2
    beg_ = 2
    end_ = WorksheetFunction.CountIf(sh.Range("A:A"), sh.Range("J" & i)) + 1
    last_cell = sh.Cells(Rows.Count, 10).End(xlUp).Row

Do Until i = last_cell + 1
    ' calculations
    diff = sh.Cells(end_, 6) - sh.Cells(beg_, 3)
    pct_change = FormatPercent(diff / sh.Cells(beg_, 3))
    vol = WorksheetFunction.Sum(sh.Range("G" & beg_ & ":G" & end_))
    
    sh.Range("K" & i) = diff
    sh.Range("L" & i) = pct_change
    sh.Range("M" & i) = vol
    
    ' using vba for the color because _
    conditional formatting is a huge performance drag _
    on workbooks this large
    If diff < 0 Then
        sh.Range("K" & i).Interior.Color = vbRed
    Else:
        sh.Range("K" & i).Interior.Color = vbGreen
    End If

    
    ' beginning of next range is 1 row after current end
    beg_ = end_ + 1
    
    'count how many rows for each ticker to _
    define ranges
    iter = WorksheetFunction.CountIf(sh.Range("A:A"), sh.Range("J" & i + 1))
    end_ = end_ + iter
    
    ' code used for debugging range selection
        'Dim w_rng
        'w_rng = sh.Range("A" & beg_ & ":A" & end_).Select
        'Selection.Activate
    
    i = i + 1
    'Debug.Print i
Loop
    
Next sh

Call finish
End Sub

Sub clear()

' used to clear cols K:M while testing function
Dim sh As Worksheet
For Each sh In ThisWorkbook.Worksheets

sh.Range("K:M").Cells.clear

Next sh

End Sub

Sub start()
' runs before starting other code to speed up runtime
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.StatusBar = "working on it . . ."
Application.EnableEvents = False
End Sub

Sub finish()
' setting everything back after code runs
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.EnableEvents = True
Application.StatusBar = Null
Application.EnableEvents = True

End Sub