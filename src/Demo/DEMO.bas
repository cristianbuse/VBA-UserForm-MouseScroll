Attribute VB_Name = "DEMO"
Option Explicit

Sub ShowDemoForm()
    With New DemoForm
        .Show vbModal
    End With
End Sub
