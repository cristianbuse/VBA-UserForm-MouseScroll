VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DemoForm 
   Caption         =   "TEST CAPTION"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13905
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
    Unload Me
End Sub

Private Sub CheckBox1_Click()
    HookMouseToForm Me, CheckBox1.Value
End Sub

Private Sub CommandButton1_Click()
    MsgBox "Demo"
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Top = Application.Top + Application.Height / 2 - Me.Height / 2
    AddDemoData
    HookMouseToForm Me
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
