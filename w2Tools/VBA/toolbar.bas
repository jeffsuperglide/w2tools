Attribute VB_Name = "toolbar"
Option Explicit

Sub Auto_Open()
'CREATING W2 TOOLBAR ON FILE OPEN

Dim i As Integer
Dim oW2 As Object
Dim cb As Object
Dim cbExists As Boolean
Dim aCaption() As Variant
Dim aStyle() As Variant
Dim aOnAction() As Variant
Dim aBeginGroup() As Variant
Dim aToolTip() As Variant
Dim aTag() As Variant
Dim aEnabled() As Variant
Dim aType As Variant

On Error GoTo errTrap

Application.Visible = True
'Setting object for toolbar creation

For Each cb In CommandBars
    If cb.Name = "W2" Then
        cbExists = True
        Set cb = Nothing
    End If
Next

If Not cbExists Then
    Set oW2 = CommandBars.Add(Name:="W2", Position:=msoBarTop)
Else
    Exit Sub
End If

'TOP LEVEL ON TOOLBAR
aType = Array(msoControlPopup, msoControlPopup, msoControlButton, msoControlButton)
aStyle = Array(msoControlPopup, msoControlPopup, msoControlButton, msoControlButton)
aCaption = Array("Modify", "Export", "Check Value(s)", "RUN EXE")
aToolTip = Array("Modify Selection", "Export Selection", "Vaules Numeric?", "RUN EXE")
aOnAction = Array("", "", "exportData.numericCheck", "toolbar.w2_pre_exe")
aBeginGroup() = Array(True, True, True, True)
With oW2

    For i = LBound(aCaption) To UBound(aCaption)
    
        '.Controls.Add Type:=msoControlPopup
        .Controls.Add Type:=aType(i)
        
        If .Controls(i + 1).Type = msoControlButton Then
        
            .Controls(i + 1).Style = msoButtonCaption
            
        End If
        
        .Controls(i + 1).Caption = aCaption(i)
        .Controls(i + 1).TooltipText = aToolTip(i)
        .Controls(i + 1).OnAction = aOnAction(i)
        .Controls(i + 1).BeginGroup = aBeginGroup(i)
    
    Next i

End With

'SECOND LEVEL ON TOOLBAR WITH ACTION
aCaption = Array("Multiply/Round", "Rounding", "Invert Rows/Columns", "Interpolate (X?)", "Interpolate (?X)")
aToolTip = Array("Multiply/Round", "Rounding", "Invert Rows/Columns", "Interpolate (X?)", "Interpolate (?X)")
aTag = Array("11", "12", "13", "14", "15")
aEnabled = Array(True, True, True, True, True)
aOnAction() = Array("ModSelect.multiply_selection", "ModSelect.round_selection", "ModSelect.switch_selection", "ModSelect.InterpXY", "ModSelect.InterpYX")
aBeginGroup() = Array(False, False, False, False, False)
With oW2.Controls(1)
 
    For i = LBound(aCaption) To UBound(aCaption)
    
        .CommandBar.Controls.Add Type:=msoControlButton 'MULTIPLY/ROUND SELECTION
        '.CommandBar.Controls(i+1).Style = msoButtonCaption
        .CommandBar.Controls(i + 1).Caption = aCaption(i)
        .CommandBar.Controls(i + 1).TooltipText = aToolTip(i)
        .CommandBar.Controls(i + 1).OnAction = aOnAction(i)
        .CommandBar.Controls(i + 1).Tag = aTag(i)
        .CommandBar.Controls(i + 1).BeginGroup = aBeginGroup(i)
        .CommandBar.Controls(i + 1).Enabled = aEnabled(i)

    Next i
    
End With

aCaption = Array("Time Series (xF8.0)", "Time Varying (10F8.0)", "Vert Profile IC (10F8.0)", "Long Profile IC (10F8.0)", "Bathymetry (10F8.0)", "CSV format")
aToolTip = Array("tin, qin, met", "wsc, euh, tuh, cuh, qwd", "vpr", "lpr", "bth", "Any selection to CSV")
aTag = Array("21", "22", "23", "24", "25", "26")
aEnabled = Array(True, True, True, True, True, True)
aOnAction() = Array("exportData.exportTimeSeries", "exportData.exportTimeVarying", "exportData.exportVertProfile", "exportData.exportLongProfile", "exportData.exportBathymetry", "exportData.exportCsv")
aBeginGroup() = Array(False, False, False, False, False, False)

With oW2.Controls(2)
 
    For i = LBound(aCaption) To UBound(aCaption)
    
        .CommandBar.Controls.Add Type:=msoControlButton 'EXPORT OPTIONS
        '.CommandBar.Controls(i+1).Style = msoButtonCaption
        .CommandBar.Controls(i + 1).Caption = aCaption(i)
        .CommandBar.Controls(i + 1).TooltipText = aToolTip(i)
        .CommandBar.Controls(i + 1).OnAction = aOnAction(i)
        .CommandBar.Controls(i + 1).Tag = aTag(i)
        .CommandBar.Controls(i + 1).BeginGroup = aBeginGroup(i)
        .CommandBar.Controls(i + 1).Enabled = aEnabled(i)

    Next i
    
End With

oW2.Visible = True

ExitSub:

Exit Sub

'ERROR TRAP FOR ANY ISSUES THAT COME UP
errTrap:
    
    MsgBox "Error Description: " & Err.Description & vbCr & "Check input data.", vbCritical
    Err.Clear
    GoTo ExitSub


End Sub

Sub Auto_Close()
'DELETING TOOLBAR WHEN ADDIN IS CLOSED/REMOVED
    Dim cb
    On Error GoTo errTrap
    Application.Caption = "Microsoft Excel"
    For Each cb In CommandBars
        If cb.Name = "W2" Then
            Application.CommandBars("W2").Delete
            Set cb = Nothing
        End If
    Next
    Exit Sub
    'ERROR TRAP FOR ANY ISSUES THAT COME UP
errTrap:
        MsgBox "Error Description: " & Err.Description & vbCr & "Check input data.", vbCritical
        Err.Clear
End Sub

Private Sub w2_pre_exe()
'SETTING PARAMETERS FOR A FILE DIALOG
'AND OPENING FOR THE USER TO RUN EXECUTABLES

Dim fd As FileDialog
Dim fdselection As Variant
Dim RetVal

'OPEN FILE DIALOG BOX
Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
With fd

    .AllowMultiSelect = False
    .Filters.Clear
    .Filters.Add "All Executables", "*.exe", 1
    .InitialFileName = Application.ActiveWorkbook.Path
    .InitialView = msoFileDialogViewDetails
    .Title = "Run Executable"
    
    If .Show = -1 Then
    
        For Each fdselection In .SelectedItems
'RUNNING EXECUTABLE IN A SHELL
            RetVal = Shell(fdselection, vbNormalFocus)
            
        Next
        
    End If
    
End With

Set fd = Nothing

End Sub

