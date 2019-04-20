VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSpr 
   Caption         =   "Add Spr.opt Data"
   ClientHeight    =   8565
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4728
   OleObjectBlob   =   "frmSpr.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdAdd_Click()
'
'This command uses selected data in the list boxes to decide the data range to add as a new series in the selected plot
'
    Dim irow As Integer
    Dim icol As Integer
    Dim mycol As Integer
    Dim myrowstart As Integer
    Dim myrowend As Integer
    Dim mysheet As Object
    Dim mySeries As Object
    Dim myXval As String
    Dim myVal As String
    
On Error GoTo errTrap
    
    Set mysheet = Sheets(frmSpr.lstSpr.Text)
    

'Determine column number for selected segment
    icol = 1

    Do While mysheet.Cells(1, icol) <> ""
    
        If mysheet.Cells(1, icol).Value = frmSpr.lstSeg Then
        
            mycol = icol
            
        End If
    
        icol = icol + 1
        
    Loop
'Determine column number for selected segment
'
'Determine row numbers for selected jday
    irow = 2

    Do While mysheet.Cells(irow, 1) <> ""
    
        If mysheet.Cells(irow, 1) = frmSpr.lstConst.Text And mysheet.Cells(irow, 2) = frmSpr.lstJday.Text Then
    
            myrowend = irow
            
        End If
    
        irow = irow + 1
    
    Loop
'Reverse to catch rowend
    Do Until irow = 1
    
        If mysheet.Cells(irow, 1) = frmSpr.lstConst.Text And mysheet.Cells(irow, 2) = frmSpr.lstJday.Text Then
    
            myrowstart = irow
            
        End If
    
        irow = irow - 1
    
    Loop

'Determine row numbers for selected jday
'
'Set address and add new series to selected plot
    myXval = mysheet.Cells(myrowstart, mycol).Address() & ":" & mysheet.Cells(myrowend, mycol).Address()
    myVal = mysheet.Cells(myrowstart, mycol - 1).Address() & ":" & mysheet.Cells(myrowend, mycol - 1).Address()
'    MsgBox myXval & vbCr & myVal
    
    Set mySeries = ActiveChart.SeriesCollection.NewSeries

    mySeries.Name = frmSpr.lstSeg.Text & " at " & frmSpr.lstJday
    mySeries.XValues = "='" & frmSpr.lstSpr.Text & "'!" & myXval
    mySeries.Values = "='" & frmSpr.lstSpr.Text & "'!" & myVal

'    ActiveChart.SeriesCollection.NewSeries
'    ActiveChart.SeriesCollection(1).Name = "=""NEW SERIES"""
'    ActiveChart.SeriesCollection(1).XValues = "='" & frmSpr.lstSpr.Text & "'!" & myXval
'    ActiveChart.SeriesCollection(1).Values = "='" & frmSpr.lstSpr.Text & "'!" & myVal
'Set address and add new series to selected plot

ExitSub:

    Set mysheet = Nothing
    Set mySeries = Nothing

    Exit Sub
    
errTrap:
    MsgBox Err.Description & vbCr & "Check to see if chart is selected.", vbCritical, "Error Message"
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub cmdClose_Click()
'
'
'
    Unload frmSpr

End Sub

Private Sub lstConst_Click()
'
'Selecting an entry will populate the segments available within the worksheet connected to the spr.opt ASCII file
'
    Dim icol As Integer
    Dim mysheet As Object
    
    Set mysheet = Sheets(frmSpr.lstSpr.Text)
    
    icol = 1
    
    frmSpr.lstSeg.Clear
    frmSpr.lstJday.Clear
    
    Do While mysheet.Cells(1, icol) <> ""
    
        If Left(Trim(mysheet.Cells(1, icol).Value), 4) = "Seg_" Then
        
            frmSpr.lstSeg.AddItem Trim(mysheet.Cells(1, icol).Value)
            
        End If
        
        icol = icol + 1
        
    Loop
    
    Set mysheet = Nothing
End Sub

Private Sub lstSeg_Click()
'
'Selecting an entry populates the julian days available corresponding to the selected consitituent
'
    Dim irow As Integer
    Dim mysheet As Object

    Set mysheet = Sheets(frmSpr.lstSpr.Text)

    irow = 1
    
    frmSpr.lstJday.Clear
    
    Do While mysheet.Cells(irow, 2) <> ""
    
        'If mysheet.Cells(irow, 1) = frmSpr.lstConst.Text And mysheet.Cells(irow, 2) <> mysheet.Cells(irow + 1, 2) Then
        If mysheet.Cells(irow, 2) <> mysheet.Cells(irow + 1, 2) Then
        
            frmSpr.lstJday.AddItem mysheet.Cells(irow + 1, 2).Value
            
        End If
        
        irow = irow + 1
        
    Loop
    
    frmSpr.lstJday.RemoveItem (frmSpr.lstJday.ListCount - 1) 'Remove empty
    
    Set mysheet = Nothing
    
End Sub

Private Sub lstSpr_Click()
'
'Selecting an entry will check the selected worksheet name, row 1 column 1 (range A1), for the text "Constituent".
'Having the correct worksheet and data in the correct location will populate the listbox lstConst with available constituent names.
'
'The spr.opt data output has to be connected to a worksheet starting in range A1.  The first row is the header info and the first column
'is the constituent name.
'
    Dim irow As Integer
    Dim icnt As Integer
    Dim iupper As Integer
    Dim mysheet As Object
    Dim frm As Object
    Dim aConst
    Dim myConst

    Set mysheet = Sheets(frmSpr.lstSpr.Text)

'Check to see if selected worksheet is from spr.opt
    If mysheet.Range("A1").Value <> "Constituent" Then
    
        frmSpr.lstConst.Clear
        frmSpr.lstJday.Clear
        frmSpr.lstSeg.Clear
        
        Exit Sub
    
    End If
'Check to see if selected worksheet is from spr.opt

    frmSpr.lstConst.Clear

'Determine constituents
    irow = 2
    
    aConst = Array(mysheet.Cells(irow, 1).Value)
    
    Do While mysheet.Cells(irow, 1) <> ""
    
        If mysheet.Cells(irow, 1).Value <> mysheet.Cells(irow + 1, 1).Value Then
        
            ReDim Preserve aConst(UBound(aConst) + 1)

            aConst(UBound(aConst)) = mysheet.Cells(irow + 1, 1).Value
            
        End If
        
        irow = irow + 1
        
    Loop
'
'Determining the number of constituents
    For Each myConst In aConst
    
        If aConst(LBound(aConst)) = myConst Then
        
            iupper = iupper + 1
            
            Exit For
            
        End If
        
    Next
'Determining the number of constituents
'Add item to listbox based on upper integer variable iupper
    For icnt = LBound(aConst) To iupper
    
        If aConst(icnt) <> Empty Then frmSpr.lstConst.AddItem aConst(icnt)
        
    Next icnt

'Determine constituents
    
'    frmSpr.lstConst.RemoveItem (frmSpr.lstConst.ListCount - 1) 'Remove empty
    
    Set mysheet = Nothing
    Set frm = Nothing
    
End Sub

Private Sub UserForm_Initialize()
'
'Adding all the worksheet names during form initializing
'
    Dim mysheet
    
    For Each mysheet In Application.ActiveWorkbook.Worksheets
    
        lstSpr.AddItem mysheet.Name
        
    Next
'
End Sub
