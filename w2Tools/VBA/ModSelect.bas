Attribute VB_Name = "ModSelect"
Option Explicit

Dim prcnt
Dim rnd
Dim aMyArray
Dim aMyArrayInvert
Dim irow As Double
Dim icol As Double
Dim vbMsg As VbMsgBoxResult

Dim startRow As Long
Dim startCol As Long
Dim numCols As Long
Dim numRows As Long
Dim counter As Long

Private Sub multiply_selection()
'SCRIPT USED TO MULTIPLY A SELECTION BY THE USER'S INPUT AND
'ROUND THE VALUE BASED ON USER INPUT

On Error GoTo errTrap

prcnt = CDbl(InputBox("Multiply selection by...?", "Multiply Selection", "1.0"))

rnd = CDbl(InputBox("Enter the desired precission, decimal places.", "Decimal Places to Round", "4"))

If ActiveWindow.RangeSelection.Rows.count = 1 And ActiveWindow.RangeSelection.Columns.count = 1 Then
    
    aMyArray = Round(ActiveWindow.Selection * prcnt, rnd)

Else

counter = 1
frmProgress.LabelProgressForm.Caption = "Multiplying and rounding.  Please wait..."
frmProgress.Show vbModeless

'SETTING AN ARRAY WITH SELECTION VALUES AND APPLYING MULTIPLY AND ROUND VALUE
    aMyArray = ActiveWindow.Selection
    
    For irow = LBound(aMyArray, 1) To UBound(aMyArray, 1)
    
        For icol = LBound(aMyArray, 2) To UBound(aMyArray, 2)
            
            If aMyArray(irow, icol) <> Empty Then
            
                aMyArray(irow, icol) = Round(aMyArray(irow, icol) * prcnt, rnd)
            
            End If
            
        Next icol
        counter = counter + 1
        modProgress.UpdateProgressBar (counter / (UBound(aMyArray, 1) * UBound(aMyArray, 2)))
    Next irow
    
End If

'SETTING SELECTION AREA EQUAL TO ARRAY
ActiveWindow.Selection = aMyArray

endForm:
Unload frmProgress

Exit Sub

errTrap:

    If prcnt = Empty Or rnd = Empty Then
    
        Exit Sub
        
    Else
    
        MsgBox "Error Description :" & Err.Description
        Err.Clear
        GoTo endForm
        
    End If
    
End Sub

Private Sub round_selection()
'ROUND THE VALUE BASED ON USER INPUT

On Error GoTo errTrap

rnd = CDbl(InputBox("Enter the desired precission, decimal places.", "Decimal Places to Round", "4"))

If ActiveWindow.RangeSelection.Rows.count = 1 And ActiveWindow.RangeSelection.Columns.count = 1 Then
    
    aMyArray = Round(ActiveWindow.Selection, rnd)

Else

counter = 1
frmProgress.LabelProgressForm.Caption = "Rounding.  Please wait..."
frmProgress.Show vbModeless

'SETTING AN ARRAY WITH SELECTION VALUES AND APPLYING MULTIPLY AND ROUND VALUE
    aMyArray = ActiveWindow.Selection
    
    For irow = LBound(aMyArray, 1) To UBound(aMyArray, 1)
    
        For icol = LBound(aMyArray, 2) To UBound(aMyArray, 2)
            
            If aMyArray(irow, icol) <> Empty Then
            
                aMyArray(irow, icol) = Round(aMyArray(irow, icol), rnd)
            
            End If
            
        Next icol
        counter = counter + 1
        modProgress.UpdateProgressBar (counter / (UBound(aMyArray, 1) * UBound(aMyArray, 2)))
    Next irow
    
End If

'SETTING SELECTION AREA EQUAL TO ARRAY
ActiveWindow.Selection = aMyArray

endForm:
Unload frmProgress

Exit Sub

errTrap:

    If prcnt = Empty Or rnd = Empty Then
    
        Exit Sub
        
    Else
    
        MsgBox "Error Description :" & Err.Description
        Err.Clear
        GoTo endForm
        
    End If

End Sub

Private Sub switch_selection()
'SWITCH SELECTED VALUES
'ONLY VALUES ADJACENT TO ONE ANOTHER

On Error GoTo errTrap
    
If ActiveWindow.RangeSelection.Rows.count = 1 And ActiveWindow.RangeSelection.Columns.count = 1 Then Exit Sub

aMyArray = ActiveWindow.Selection

aMyArrayInvert = aMyArray

vbMsg = MsgBox("Do you want to invert the rows?" & vbCr & "Selecting 'No' inverts the columns.", vbYesNoCancel, "Invert Selection")

counter = 1
frmProgress.LabelProgressForm.Caption = "Pivot table.  Please wait..."
frmProgress.Show vbModeless

Select Case vbMsg

    Case vbCancel
    
        Exit Sub
    
    Case vbYes
    
        For icol = LBound(aMyArray, 2) To UBound(aMyArray, 2)
        
            For irow = UBound(aMyArray, 1) To LBound(aMyArray, 1) Step -1
            
                aMyArrayInvert(UBound(aMyArray, 1) - irow + 1, icol) = aMyArray(irow, icol)
                
            counter = counter + 1
            modProgress.UpdateProgressBar (counter / (UBound(aMyArray, 1) * UBound(aMyArray, 2)))
            Next irow
        Next icol

    Case vbNo

        For irow = LBound(aMyArray, 1) To UBound(aMyArray, 1)
        
            For icol = UBound(aMyArray, 2) To LBound(aMyArray, 2) Step -1
            
                aMyArrayInvert(irow, UBound(aMyArray, 2) - icol + 1) = aMyArray(irow, icol)
                
            counter = counter + 1
            modProgress.UpdateProgressBar (counter / (UBound(aMyArray, 1) * UBound(aMyArray, 2)))
            Next icol
        Next irow
        
End Select

'SETTING SELECTION AREA EQUAL TO ARRAY
ActiveWindow.Selection = aMyArrayInvert

endForm:
Unload frmProgress

Exit Sub

errTrap:

    MsgBox "Error Description :" & Err.Description
    Err.Clear
    GoTo endForm

End Sub
Private Sub InterpYX()
'USE InterpYX TO INTERPOLATE VALUES FROM TWO COLUMN HIGHLIGHT RANGE
'YX ASSUMES THE VALUES TOO COMPUTE ARE IN COLUMN 1
Dim icnt As Long ' Integer
Dim iX1 As Double
Dim iX2 As Double
Dim iY1 As Double
Dim iY2 As Double
Dim dSlope As Double
Dim dIntercept As Double
Dim mySht As Object
Set mySht = ActiveWorkbook.ActiveSheet
'CHECK NUMBER OF COLUMNS SELECTED
If ActiveWindow.RangeSelection.Columns.count <> 2 Then
    MsgBox "Select two continuous columns.", vbInformation + vbCritical, "Linear Interpolation"
    Exit Sub
End If
'ASSIGN VARIABLES WITH KNOWN VALUES
iX1 = mySht.Cells(ActiveWindow.RangeSelection.Row, ActiveWindow.RangeSelection.Column + 1)
iY1 = mySht.Cells(ActiveWindow.RangeSelection.Row, ActiveWindow.RangeSelection.Column)
iX2 = mySht.Cells(ActiveWindow.RangeSelection.Row + ActiveWindow.RangeSelection.Rows.count - 1, ActiveWindow.RangeSelection.Column + 1)
iY2 = mySht.Cells(ActiveWindow.RangeSelection.Row + ActiveWindow.RangeSelection.Rows.count - 1, ActiveWindow.RangeSelection.Column)
dSlope = (iY2 - iY1) / (iX2 - iX1)
dIntercept = iY1 - dSlope * iX1
'INTERPOLATE

counter = 0
frmProgress.LabelProgressForm.Caption = "Interpolating.  Please wait..."
frmProgress.Show vbModeless

For icnt = ActiveWindow.RangeSelection.Row To ActiveWindow.RangeSelection.Row + ActiveWindow.RangeSelection.Rows.count - 1
    mySht.Cells(icnt, ActiveWindow.RangeSelection.Column).Value = dSlope * Cells(icnt, ActiveWindow.RangeSelection.Column + 1) + dIntercept
    counter = counter + 1
    modProgress.UpdateProgressBar (counter / ActiveWindow.RangeSelection.Rows.count)

Next icnt

Unload frmProgress

Set mySht = Nothing
End Sub

Private Sub InterpXY()
'USE InterpXY TO INTERPOLATE VALUES FROM TWO COLUMN HIGHLIGHT RANGE
'XY ASSUMES THE VALUES TOO COMPUTE ARE IN COLUMN 2
Dim icnt As Integer
Dim iX1 As Double
Dim iX2 As Double
Dim iY1 As Double
Dim iY2 As Double
Dim dSlope As Double
Dim dIntercept As Double
Dim mySht As Object
Set mySht = ActiveWorkbook.ActiveSheet
'CHECK NUMBER OF COLUMNS SELECTED
If ActiveWindow.RangeSelection.Columns.count <> 2 Then
    MsgBox "Select two continuous columns.", vbInformation + vbCritical, "Linear Interpolation"
    Exit Sub
End If
'ASSIGN VARIABLES WITH KNOWN VALUES
iX1 = mySht.Cells(ActiveWindow.RangeSelection.Row, ActiveWindow.RangeSelection.Column)
iY1 = mySht.Cells(ActiveWindow.RangeSelection.Row, ActiveWindow.RangeSelection.Column + 1)
iX2 = mySht.Cells(ActiveWindow.RangeSelection.Row + ActiveWindow.RangeSelection.Rows.count - 1, ActiveWindow.RangeSelection.Column)
iY2 = mySht.Cells(ActiveWindow.RangeSelection.Row + ActiveWindow.RangeSelection.Rows.count - 1, ActiveWindow.RangeSelection.Column + 1)
dSlope = (iY2 - iY1) / (iX2 - iX1)
dIntercept = iY1 - dSlope * iX1
'INTERPOLATE

counter = 0
frmProgress.LabelProgressForm.Caption = "Interpolating.  Please wait..."
frmProgress.Show vbModeless

For icnt = ActiveWindow.RangeSelection.Row To ActiveWindow.RangeSelection.Row + ActiveWindow.RangeSelection.Rows.count - 1
    mySht.Cells(icnt, ActiveWindow.RangeSelection.Column + 1).Value = dSlope * Cells(icnt, ActiveWindow.RangeSelection.Column) + dIntercept
    counter = counter + 1
    modProgress.UpdateProgressBar (counter / ActiveWindow.RangeSelection.Rows.count)
Next icnt

Unload frmProgress

Set mySht = Nothing
End Sub


