Attribute VB_Name = "exportData"
Option Explicit

Dim fso As Object
Dim npt As Object

Dim lngRow As Long
Dim lngCol As Long
Dim lngCnt As Long
Dim counter As Long

Dim arySelection As Variant
Dim strWkBkPath As String
Dim strWkShtName As String
Dim strMsg As String
Dim strMsgTitle As String
Dim msgType As VbMsgBoxStyle

Private Sub exportTimeSeries()
'EXPORT TIME SERIES DATA FOR W2 FORMAT, 10F8.0
'FOLLOW FORMAT STRUCTURE FOR W2
'FIRST THREE SELECTED ROWS ARE IGNORED BY W2 AND THEREFORE LOADED AS TEXT IN THIS ROUTINE
'THIS SCRIPT ASSUMES THE FIRST COLUMN IS JULIAN DAY AND EACH ADDITIONAL COLUMN IS TIME SERIES DATA

Dim strLine As String
Const intRowStart As Integer = 3

'CHECKING REQUIREMENTS TO RUN
If minSelected(5, 2) Then GoTo errTrap
'CHECKING TO SEE IS THE DATA IS A NUMBER
If Not valIsNumeric(intRowStart) Then GoTo errTrap

'WORKBOOK PATH AND SHEET NAME
strWkBkPath = Application.ActiveWorkbook.Path
strWkShtName = Application.ActiveSheet.Name

'LOADING ARRAY FROM SELECTION
arySelection = ActiveWindow.Selection

'FORMATTING VALUES IN THE SELECTION ARRAY ASSUMING THE FIRST ROW IS THE TITLE
Call setPrecision

'SET FILE SYSTEM OBJECT AND CREATE A TEXT FILE
Set fso = CreateObject("Scripting.FileSystemObject")
Set npt = fso.CreateTextFile(strWkBkPath & "\" & strWkShtName, True)

'WRITE TEXT1
npt.writeline arySelection(1, 1)

counter = 0
frmProgress.LabelProgressForm.Caption = "Exporting time series.  Please wait..." 'SHOW THE PROGRESS BAR
frmProgress.Show vbModeless

'For...Next TO EXPORT FROM ARRAY TO OUTPUT FILE
For lngRow = LBound(arySelection, 1) + 1 To UBound(arySelection, 1)
  strLine = Empty
  For lngCol = LBound(arySelection, 2) To UBound(arySelection, 2)
    strLine = RTrim(strLine & arySelection(lngRow, lngCol))
    'UPDATE THE PROGRESS BAR
    counter = counter + 1
    modProgress.UpdateProgressBar (counter / ((UBound(arySelection, 1) - 1) * UBound(arySelection, 2)))
  Next lngCol
  npt.writeline strLine
Next lngRow

Call exportMessage

npt.Close

ExitSub:
ActiveCell.Select
Application.Cursor = xlDefault

Unload frmProgress 'UNLOAD THE PROGRESS BAR

Exit Sub

'ERROR TRAP FOR ANY ISSUES THAT COME UP
errTrap:
    MsgBox strMsg, msgType, strMsgTitle
    Err.Clear
    GoTo ExitSub
End Sub

Private Sub exportTimeVarying()
'EXPORT TIME VARYING DATA FOR W2 FORMAT, 10F8.0
'FOLLOW FORMAT STRUCTURE FOR W2
'W2 IGNORES THE FIRST THREE LINES IN THE INPUT FILE; THEREFORE,
'THE FIRST TWO SELECTED ROWS ARE WRITTEN OUT AS LINES 1 AND 3 WITH A BLANK LINE FOR 2
'ROW THREE IN THE SELECTED RANGE IS THE JULIAN DAY FOR THAT COLUMN AND IS PREPENDED TO THE FIRST DATA LINE

Dim strLine As String
Const intRowStart As Integer = 3


'CHECKING REQUIREMENTS TO RUN
If minSelected(5, 2) Then GoTo errTrap
'CHECKING TO SEE IS THE DATA IS A NUMBER
If Not valIsNumeric(2) Then GoTo errTrap

'WORKBOOK PATH AND SHEET NAME
strWkBkPath = Application.ActiveWorkbook.Path
strWkShtName = Application.ActiveSheet.Name

'LOADING ARRAY FROM SELECTION
arySelection = ActiveWindow.Selection

'FORMATTING VALUES IN THE SELECTION ARRAY
Call setPrecision

'SET FILE SYSTEM OBJECT AND CREATE A TEXT FILE
Set fso = CreateObject("Scripting.FileSystemObject")
Set npt = fso.CreateTextFile(strWkBkPath & "\" & strWkShtName, True)

'WRITE TEXT1 AND HEADER INFORMATION
npt.writeline arySelection(1, 1)
npt.writeline
strLine = "    JDAY"
For lngCnt = 1 To 9
  strLine = strLine & arySelection(2, 1)
Next lngCnt
npt.writeline strLine

counter = 0
frmProgress.LabelProgressForm.Caption = "Exporting time varying.  Please wait..." 'SHOW THE PROGRESS BAR
frmProgress.Show vbModeless

'For...Next TO EXPORT FROM ARRAY TO OUTPUT FILE
For lngCol = LBound(arySelection, 2) To UBound(arySelection, 2)
  strLine = arySelection(intRowStart, lngCol) 'THIS IS THE JULIAN DAY IN ROW 3
  counter = counter + 3
  For lngRow = LBound(arySelection, 1) + intRowStart To UBound(arySelection, 1)
    strLine = strLine & arySelection(lngRow, lngCol)
    If (lngRow - intRowStart) Mod 9 = 0 And lngRow < UBound(arySelection, 1) Then  'THIS DETERMINES WHEN TO ADD A LINE
      strLine = strLine & vbCr & Space(8)
    End If
    'UPDATE THE PROGRESS BAR
    counter = counter + 1
    modProgress.UpdateProgressBar (counter / (UBound(arySelection, 1) * UBound(arySelection, 2)))
  Next lngRow
  npt.writeline strLine
Next lngCol


Call exportMessage

npt.Close

ExitSub:
ActiveCell.Select
Application.Cursor = xlDefault

Unload frmProgress 'UNLOAD THE PROGRESS BAR

Exit Sub

'ERROR TRAP FOR ANY ISSUES THAT COME UP
errTrap:
    MsgBox strMsg, msgType, strMsgTitle
    Err.Clear
    GoTo ExitSub
End Sub

Private Sub exportVertProfile()
'EXPORT TIME VARYING DATA FOR W2 FORMAT, 10F8.0
'FOLLOW FORMAT STRUCTURE FOR W2
'W2 IGNORES THE FIRST THREE LINES IN THE INPUT FILE; THEREFORE,
'THE FIRST TWO SELECTED ROWS ARE WRITTEN OUT AS LINES 1 AND 3 WITH A BLANK LINE FOR 2
'ROW THREE IN THE SELECTED RANGE IS THE JULIAN DAY FOR THAT COLUMN AND IS PREPENDED TO THE FIRST DATA LINE

Dim strLine As String
Dim strConLbl As String
Const intRowStart As Integer = 2

'CHECKING REQUIREMENTS TO RUN
If minSelected(5, 1) Then GoTo errTrap
'CHECKING TO SEE IS THE DATA IS A NUMBER
If Not valIsNumeric(intRowStart) Then GoTo errTrap

'WORKBOOK PATH AND SHEET NAME
strWkBkPath = Application.ActiveWorkbook.Path
strWkShtName = Application.ActiveSheet.Name

'LOADING ARRAY FROM SELECTION
arySelection = ActiveWindow.Selection

'FORMATTING VALUES IN THE SELECTION ARRAY
Call setPrecision

'SET FILE SYSTEM OBJECT AND CREATE A TEXT FILE
Set fso = CreateObject("Scripting.FileSystemObject")
Set npt = fso.CreateTextFile(strWkBkPath & "\" & strWkShtName, True)

'WRITE TEXT1
npt.writeline arySelection(1, 1)

'strConLbl = strConLbl & Space(1) & "C1"
'For lngCnt = 2 To 9
'  strConLbl = strConLbl & Space(6) & "C" & lngCnt
'Next lngCnt

counter = 0
frmProgress.LabelProgressForm.Caption = "Exporting Vertical Profile.  Please wait..." 'SHOW THE PROGRESS BAR
frmProgress.Show vbModeless

'For...Next TO EXPORT FROM ARRAY TO OUTPUT FILE
For lngCol = LBound(arySelection, 2) To UBound(arySelection, 2)
  npt.writeline
  'WRITE EMPTY LINE AND LABELS FOR EACH COLUMN
  strLine = arySelection(intRowStart, lngCol) 'THIS IS THE CONSTITUENT LABEL
  
  For lngCnt = 1 To 9
    strLine = strLine & "      C" & lngCnt
  Next lngCnt
  npt.writeline strLine
  strLine = Space(8)
  
  counter = counter + 2
  
'For...Next TO EXPORT FROM ARRAY TO OUTPUT FILE
  For lngRow = LBound(arySelection, 1) + intRowStart To UBound(arySelection, 1)
    strLine = strLine & arySelection(lngRow, lngCol)
    If (lngRow - intRowStart) Mod 9 = 0 And lngRow < UBound(arySelection, 1) Then  'THIS DETERMINES WHEN TO ADD A LINE
      strLine = strLine & vbCr & Space(8)
    End If
    'UPDATE THE PROGRESS BAR
    counter = counter + 1
    modProgress.UpdateProgressBar (counter / (UBound(arySelection, 1) * UBound(arySelection, 2)))
  Next lngRow
  npt.writeline strLine
Next lngCol

Call exportMessage

npt.Close

ExitSub:
ActiveCell.Select
Application.Cursor = xlDefault

Unload frmProgress 'UNLOAD THE PROGRESS BAR

Exit Sub

'ERROR TRAP FOR ANY ISSUES THAT COME UP
errTrap:
    MsgBox strMsg, msgType, strMsgTitle
    Err.Clear
    GoTo ExitSub
End Sub

Private Sub exportLongProfile()
'EXPORT TIME VARYING DATA FOR W2 FORMAT, 10F8.0
'FOLLOW FORMAT STRUCTURE FOR W2
'W2 IGNORES THE FIRST THREE LINES IN THE INPUT FILE; THEREFORE,
'THE FIRST TWO SELECTED ROWS ARE WRITTEN OUT AS LINES 1 AND 3 WITH A BLANK LINE FOR 2
'ROW THREE IN THE SELECTED RANGE IS THE JULIAN DAY FOR THAT COLUMN AND IS PREPENDED TO THE FIRST DATA LINE

Dim strLine As String
Dim strConLbl As String
Dim lngUbound As Long
Const intRowStart As Integer = 4

'CHECKING REQUIREMENTS TO RUN
If minSelected(5, 2) Then GoTo errTrap
'CHECKING TO SEE IF THE DATA IS A NUMBER BY CHECKING VALUES STARTING AT LBOUND + 2 ROWS
If Not valIsNumeric(2) Then GoTo errTrap

'WORKBOOK PATH AND SHEET NAME
strWkBkPath = Application.ActiveWorkbook.Path
strWkShtName = InputBox("Enter file name.", "Longitudinal File Name", Application.ActiveSheet.Name)
If strWkShtName = "" Then strWkShtName = Application.ActiveSheet.Name

'LOADING ARRAY FROM SELECTION
arySelection = ActiveWindow.Selection

'GETTING CONSTITUENT NAME BEFORE SETTING PRECISION
strConLbl = arySelection(2, 1) & Space(16) 'THIS IS THE CONSTITUENT LABEL LIMITED TO 13 CHARACTERS
strConLbl = Left(strConLbl, 13)

'FORMATTING VALUES IN THE SELECTION ARRAY
Call setPrecision

'SET FILE SYSTEM OBJECT AND CREATE A TEXT FILE
Set fso = CreateObject("Scripting.FileSystemObject")
Set npt = fso.CreateTextFile(strWkBkPath & "\" & strWkShtName, True)

'WRITE TEXT1
npt.writeline arySelection(1, 1)

counter = 0
frmProgress.LabelProgressForm.Caption = "Exporting Longitudinal Profile.  Please wait..." 'SHOW THE PROGRESS BAR
frmProgress.Show vbModeless

strConLbl = strConLbl & Space(1) & "C1"
For lngCnt = 2 To 9
  strConLbl = strConLbl & Space(6) & "C" & lngCnt
  'counter = counter + 1
Next lngCnt

On Error GoTo errTrap

'For...Next TO EXPORT FROM ARRAY TO OUTPUT FILE
For lngCol = LBound(arySelection, 2) To UBound(arySelection, 2)

  If arySelection(4, lngCol) <= 0 Then GoTo nextCol
  'WRITE LABELS FOR EACH COLUMN
  npt.writeline "Segment " & Trim(CLng(arySelection(3, lngCol))) 'SEGMENT NUMBER
  npt.writeline strConLbl
  strLine = Space(8)
  
  lngUbound = (CLng(arySelection(4, lngCol)) + intRowStart)
  
  For lngRow = LBound(arySelection, 1) + intRowStart To lngUbound
    strLine = strLine & arySelection(lngRow, lngCol)
     If (lngRow - intRowStart) Mod 9 = 0 And lngRow < lngUbound Then  'THIS DETERMINES WHEN TO ADD A LINE
      strLine = strLine & vbCr & Space(8)
    End If
  Next lngRow
  npt.writeline strLine
  
nextCol:

    'UPDATE THE PROGRESS BAR
    counter = counter + 1
    'modProgress.UpdateProgressBar (counter / (UBound(arySelection, 2) * lngUbound))
    modProgress.UpdateProgressBar (counter / UBound(arySelection, 2))


Next lngCol

Call exportMessage

npt.Close

ExitSub:
ActiveCell.Select
Application.Cursor = xlDefault

Unload frmProgress 'UNLOAD THE PROGRESS BAR

Exit Sub

'ERROR TRAP FOR ANY ISSUES THAT COME UP
errTrap:
    strMsg = Err.Description
    msgType = vbCritical
    strMsgTitle = "Error"
    MsgBox strMsg, msgType, strMsgTitle
    Err.Clear
    GoTo ExitSub
End Sub

Private Sub exportBathymetry()
'EXPORT BATHYMETRY DATA FOR W2 FORMAT, 10F8.0
'FOLLOW FORMAT STRUCTURE FOR W2
'W2 IGNORES THE FIRST THREE LINES IN THE INPUT FILE
'
'

Dim strLine As String

'CHECKING REQUIREMENTS TO RUN
If minSelected(5, 2) Then GoTo errTrap
'CHECKING TO SEE IS THE DATA IS A NUMBER
'If Not valIsNumeric(1) Then GoTo errTrap

'WORKBOOK PATH AND SHEET NAME
strWkBkPath = Application.ActiveWorkbook.Path
strWkShtName = Application.ActiveSheet.Name

'LOADING ARRAY FROM SELECTION
arySelection = ActiveWindow.Selection

'FORMATTING VALUES IN THE SELECTION ARRAY
Call setPrecision

'SET FILE SYSTEM OBJECT AND CREATE A TEXT FILE
Set fso = CreateObject("Scripting.FileSystemObject")
Set npt = fso.CreateTextFile(strWkBkPath & "\" & strWkShtName, True)

counter = 0
frmProgress.LabelProgressForm.Caption = "Exporting bathymetry.  Please wait..." 'SHOW THE PROGRESS BAR
frmProgress.Show vbModeless

'WRITE TEXT1
npt.writeline arySelection(1, 1)
npt.writeline

'START GEOMETRY SECTION 1 OUTPUT
For lngRow = 3 To 6
  npt.writeline LTrim(arySelection(lngRow, 1)) 'WRITE GEOMETRY SECTION HEADER
  strLine = Empty
  For lngCol = 2 To UBound(arySelection, 2)
    strLine = strLine & arySelection(lngRow, lngCol)
    If (lngCol - 1) Mod 10 = 0 And lngCol < UBound(arySelection, 2) Then 'THIS DETERMINES WHEN TO ADD A LINE
      strLine = strLine & vbCr
    End If
    'counter = counter + 1 'PROGRESS BAR COUNTER
  Next lngCol
  npt.writeline strLine & vbCr
Next lngRow
'START GEOMETRY SECTION 2 OUTPUT
npt.writeline LTrim(arySelection(7, 1)) 'WRITE GEOMETRY SECTION HEADER
lngCol = 1
strLine = Empty
For lngRow = 8 To UBound(arySelection, 1)
  strLine = strLine & arySelection(lngRow, lngCol)
  If (lngRow - 7) Mod 10 = 0 And lngRow < UBound(arySelection, 1) Then 'THIS DETERMINES WHEN TO ADD A LINE
    strLine = strLine & vbCr
  End If
  'counter = counter + 1 'PROGRESS BAR COUNTER
Next lngRow
npt.writeline strLine & vbCr
'START GEOMETRY SECTION 3 OUTPUT
For lngCol = 2 To UBound(arySelection, 2)
  npt.writeline "Segment " & CLng(arySelection(2, lngCol)) & " widths [B]    Branch " & CLng(arySelection(7, lngCol))
  strLine = Empty
  For lngRow = 8 To UBound(arySelection, 1)
    strLine = strLine & arySelection(lngRow, lngCol)
    If (lngRow - 7) Mod 10 = 0 And lngRow < UBound(arySelection, 1) Then 'THIS DETERMINES WHEN TO ADD A LINE
      strLine = strLine & vbCr
    End If
    'UPDATE PROGRESS BAR
    counter = counter + 1
    modProgress.UpdateProgressBar (counter / ((UBound(arySelection, 1) - 7) * (UBound(arySelection, 2) - 1)))
  Next lngRow
  npt.writeline strLine & vbCr
Next lngCol

Call exportMessage

npt.Close

ExitSub:
ActiveCell.Select
Application.Cursor = xlDefault

Unload frmProgress 'UNLOAD PROGRESS BAR

Exit Sub

'ERROR TRAP FOR ANY ISSUES THAT COME UP
errTrap:
    strMsg = Err.Description
    msgType = vbCritical
    strMsgTitle = "Error"
    MsgBox strMsg, msgType, strMsgTitle
    Err.Clear
    GoTo ExitSub
End Sub

Private Sub exportCsv()
'
'
'
Dim strLine As String

'CHECKING REQUIREMENTS TO RUN
If minSelected(5, 1) Then GoTo errTrap
'CHECKING TO SEE IS THE DATA IS A NUMBER
'If Not valIsNumeric(1) Then GoTo errTrap

'WORKBOOK PATH AND SHEET NAME
strWkBkPath = Application.ActiveWorkbook.Path
strWkShtName = Application.ActiveSheet.Name

'LOADING ARRAY FROM SELECTION
arySelection = ActiveWindow.Selection

'FORMATTING VALUES IN THE SELECTION ARRAY
'Call setPrecision

'SET FILE SYSTEM OBJECT AND CREATE A TEXT FILE
Set fso = CreateObject("Scripting.FileSystemObject")
Set npt = fso.CreateTextFile(strWkBkPath & "\" & strWkShtName, True)

counter = 1
frmProgress.LabelProgressForm.Caption = "Exporting to CSV.  Please wait..." 'SHOW THE PROGRESS BAR
frmProgress.Show vbModeless

'START GEOMETRY OUTPUT
For lngRow = LBound(arySelection, 1) To UBound(arySelection, 1)
  strLine = Empty
  For lngCol = LBound(arySelection, 2) To UBound(arySelection, 2)
    If lngRow = 1 And lngCol = 1 Then strLine = "$"
    strLine = strLine & arySelection(lngRow, lngCol) & ","
  'UPDATE PROGRESS BAR
  counter = counter + 1
  modProgress.UpdateProgressBar (counter / (UBound(arySelection, 1) * UBound(arySelection, 2)))
  Next lngCol
  npt.writeline strLine
Next lngRow

Call exportMessage

npt.Close

ExitSub:
ActiveCell.Select
Application.Cursor = xlDefault

Unload frmProgress 'UNLOAD PROGRESS BAR

Exit Sub

'ERROR TRAP FOR ANY ISSUES THAT COME UP
errTrap:
    strMsg = Err.Description
    msgType = vbCritical
    strMsgTitle = "Error"
    MsgBox strMsg, msgType, strMsgTitle
    Err.Clear
    GoTo ExitSub
End Sub

Private Function setPrecision()
'SETTING THE VALUE TO NUMBER OF DECIMAL PLACES TO FIT IN 10F8 FORMAT
Dim aryVal As String * 8
For lngCol = LBound(arySelection, 2) To UBound(arySelection, 2)
  For lngRow = LBound(arySelection, 1) + 1 To UBound(arySelection, 1)
    Select Case arySelection(lngRow, lngCol)
      Case -1E+99 To -10000000000#
        RSet aryVal = Format(arySelection(lngRow, lngCol), "0.#E+##")
      Case -10000000000# To -1000000
        RSet aryVal = Format(arySelection(lngRow, lngCol), "0.##E+##")
      Case -1000000 To -10000
        RSet aryVal = Format(arySelection(lngRow, lngCol), "0.#")
      Case -10000 To -1000
        RSet aryVal = Format(arySelection(lngRow, lngCol), "0.##")
      Case -1000 To -100
        RSet aryVal = Format(arySelection(lngRow, lngCol), "0.###")
      Case -100 To -10
        RSet aryVal = Format(arySelection(lngRow, lngCol), "0.####")
      Case -10 To -0.001
        RSet aryVal = Format(arySelection(lngRow, lngCol), "0.#####")
      Case -0.001 To -0.000000001
        RSet aryVal = Format(arySelection(lngRow, lngCol), "0.##E+##")
      Case -0.000000001 To -1E-99
        RSet aryVal = Format(arySelection(lngRow, lngCol), "0.#E+##")
        
      Case 0
        RSet aryVal = Format(arySelection(lngRow, lngCol), "0.#")
        
      Case 1E-99 To 0.000000001
        RSet aryVal = Format(arySelection(lngRow, lngCol), "0.##E+##")
      Case 0.000000001 To 0.001
        RSet aryVal = Format(arySelection(lngRow, lngCol), "0.###E+##")
      Case 0.001 To 10000000
        RSet aryVal = Format(arySelection(lngRow, lngCol), "0.######")
      Case 10000000 To 10000000000#
        RSet aryVal = Format(arySelection(lngRow, lngCol), "0.###E+##")
      Case 10000000000# To 1E+99
        RSet aryVal = Format(arySelection(lngRow, lngCol), "0.##E+##")
        
      Case Else
        RSet aryVal = arySelection(lngRow, lngCol)
        
    End Select
    arySelection(lngRow, lngCol) = aryVal
  Next lngRow
Next lngCol
End Function

Private Function valIsNumeric(rowShift As Integer) As Boolean
'CHECKING THE DATA TYPE WITHIN THE ARRAY ASSIGNGED BY THE CALLING SUBROUTINE
Dim startRow As Long
Dim startCol As Long
Dim numCols As Long
Dim numRows As Long
'
On Error Resume Next

counter = 1
frmProgress.LabelProgressForm.Caption = "Checking values.  Please wait..."
frmProgress.Show vbModeless
'
startRow = ActiveCell.Row
startCol = ActiveCell.Column
numRows = ActiveWindow.Selection.Rows.count
numCols = ActiveWindow.Selection.Columns.count
valIsNumeric = True
'
For lngCol = startCol To (numCols + ActiveCell.Column - 1)
  For lngRow = (startRow + rowShift) To (startRow + numRows - 1)
    If Not IsNumeric(ActiveSheet.Cells(lngRow, lngCol)) Or ActiveSheet.Cells(lngRow, lngCol) = "" Then
      With ActiveSheet.Cells(lngRow, lngCol)
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 65535
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
      End With
      valIsNumeric = False
      strMsg = "Numerical check equals FALSE for value(s) in selection.  Those cells have been highlighted yellow."
      strMsgTitle = "Value Type Check"
      msgType = vbCritical
    End If
    
    counter = counter + 1
    modProgress.UpdateProgressBar (counter / (numRows * numCols))
    
  Next lngRow
Next lngCol

Unload frmProgress

End Function

Private Sub numericCheck()
'CHECKING VALUES TO SEE IF THEY ARE NUMERICAL
If valIsNumeric(0) Then
  strMsg = "Numerical check equals TRUE for value(s) in selection."
  msgType = vbInformation
End If

MsgBox strMsg, msgType, strMsgTitle

End Sub

Private Function minSelected(numRow As Long, numCol As Long) As Boolean
'CHECKING THE SELECTED RANGE TO SEE IF IT MEETS THE MINIMUM NUMBER OF COLOMNS AND ROWS
If ActiveWindow.Selection.Rows.count < numRow Or ActiveWindow.Selection.Columns.count < numCol Then
  strMsg = "Selection contains less than the minimum required for export."
  strMsgTitle = "Minimum Required Selection"
  msgType = vbCritical
  minSelected = True
End If
End Function

Private Function exportMessage()
'MESSAGE TO USER WHERE FILE WAS CREATED
'
MsgBox "Selected data exported to " & strWkBkPath & "\" & strWkShtName, vbInformation + vbDefaultButton1, "Export Data File"
End Function
