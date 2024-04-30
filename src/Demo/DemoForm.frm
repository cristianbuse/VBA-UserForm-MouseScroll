VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DemoForm 
   Caption         =   "UserForm1"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13215
   OleObjectBlob   =   "DemoForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DemoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnClose_Click()
    Me.Hide
End Sub

Private Sub CheckBox1_Click()
    EnableMouseScroll Me, CheckBox1.Value, CheckBox2.Value, CheckBox3.Value
End Sub

Private Sub CheckBox2_Click()
    EnableMouseScroll Me, CheckBox1.Value, CheckBox2.Value, CheckBox3.Value
End Sub

Private Sub CheckBox3_Click()
    EnableMouseScroll Me, CheckBox1.Value, CheckBox2.Value, CheckBox3.Value
End Sub

Private Sub CommandButton1_Click()
    MsgBox "Demo"
End Sub

'Note that the error will close the Form in Microsoft Word because
'   Application.EnableCancelKey is set to wdCancelDisabled and the
'   Run-time error Dialog (End/Debug) is not shown
'In Excel the error will display the Run-time error Dialog (End/Debug)
Private Sub CommandButton2_Click()
    Debug.Print 1 / 0
End Sub

Private Sub CommandButton3_Click()
    With New DemoForm
        .Top = Me.Top + 30
        .Caption = "MODAL form"
        .Show vbModal
    End With
End Sub

Private Sub CommandButton5_Click()
    With New DemoForm
        .Top = Me.Top + 30
        On Error Resume Next
        .Caption = "MODELESS form"
        .Show vbModeless
        If Err.Number <> 0 Then
            MsgBox Err.Description, vbInformation, "Cannot display"
        End If
        On Error GoTo 0
    End With
End Sub

Private Sub CommandButton4_Click()
    Debug.Print "Input: " & InputBox("Demo", "Demo", "Demo")
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Top = Application.Top + Application.Height / 2 - Me.Height / 2
    AddDemoData
    EnableMouseScroll Me
End Sub

Private Sub AddDemoData()
    Dim i As Long
    Dim tValue As String
    
    For i = 1 To 100
        ListBox1.AddItem i
        ComboBox1.AddItem i
        TextBox1.Value = TextBox1.Value & vbNewLine & i
    Next i
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
