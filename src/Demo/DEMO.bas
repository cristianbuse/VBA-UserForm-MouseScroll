Attribute VB_Name = "DEMO"
Option Explicit

Sub ShowDemoForm()
    With New DemoForm
        .Caption = "MODELESS form"
        .Show vbModeless
    End With
End Sub
