Attribute VB_Name = "modProgress"
Option Explicit
Sub UpdateProgressBar(PctDone As Single)
    With frmProgress

        ' Update the Caption property of the Frame control.
        '.FrameProgress.Caption = Format(PctDone, "0%")
        .LabelProgressPercent.Caption = Format(PctDone, "0%")
        
        ' Widen the Label control.
        '.LabelProgress.Width = PctDone * (.FrameProgress.Width - 10)
        .LabelProgress.Width = PctDone * (.FrameProgress.Width)
    End With

    ' The DoEvents allows the UserForm to update.
    DoEvents
End Sub

