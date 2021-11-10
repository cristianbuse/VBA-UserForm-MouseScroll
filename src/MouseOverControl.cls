VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MouseOverControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''=============================================================================
''' VBA UserForm MouseScroll
''' --------------------------------------------------------
''' https://github.com/cristianbuse/VBA-UserForm-MouseScroll
''' --------------------------------------------------------
''' MIT License
'''
''' Copyright (c) 2019 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================
''
''==============================================================================
'' Description:
''    This object holds a single reference to a MSForms.UserForm or a
''       MSForms.Control that is capable of having a MouseMove event.
''    The reference is set using the InitFromControl method and can be any of
''       the 13 types listed below.
''    When the MouseMove event is triggered, the hovered control is passed to
''       the SetHoveredControl method in the MouseScroll module.
'' Notes:
''    This class implements Scroll events for the container types (frame, form
''       and multiPage) in order to supress any scrolls that are related to the
''       controls getting the focus.
'' Requires:
''    - MouseScroll: module enabling Mouse Message processing
''==============================================================================

'@PredeclaredID
Option Explicit

'Control types (MSForms.Control or MSForms.UserForm)
Private WithEvents m_CheckBox As MSForms.CheckBox
Attribute m_CheckBox.VB_VarHelpID = -1
Private WithEvents m_ComboBox As MSForms.ComboBox
Attribute m_ComboBox.VB_VarHelpID = -1
Private WithEvents m_CommandButton As MSForms.CommandButton
Attribute m_CommandButton.VB_VarHelpID = -1
Private WithEvents m_Frame As MSForms.Frame
Attribute m_Frame.VB_VarHelpID = -1
Private WithEvents m_Image As MSForms.Image
Attribute m_Image.VB_VarHelpID = -1
Private WithEvents m_Label As MSForms.Label
Attribute m_Label.VB_VarHelpID = -1
Private WithEvents m_ListBox As MSForms.ListBox
Attribute m_ListBox.VB_VarHelpID = -1
Private WithEvents m_MultiPage As MSForms.MultiPage
Attribute m_MultiPage.VB_VarHelpID = -1
Private WithEvents m_OptionButton As MSForms.OptionButton
Attribute m_OptionButton.VB_VarHelpID = -1
Private WithEvents m_TabStrip As MSForms.TabStrip
Attribute m_TabStrip.VB_VarHelpID = -1
Private WithEvents m_TextBox As MSForms.TextBox
Attribute m_TextBox.VB_VarHelpID = -1
Private WithEvents m_ToggleButton As MSForms.ToggleButton
Attribute m_ToggleButton.VB_VarHelpID = -1
Private WithEvents m_UserForm As MSForms.UserForm
Attribute m_UserForm.VB_VarHelpID = -1

#Const DETECT_COMMON_CONTROLS = 0 'Turn to 1 to detect ListView controls
#Const DETECT_INTERNET_CONTROLS = 0 'Turn to 1 to detect WebBrowser controls

#If DETECT_COMMON_CONTROLS Then
Private WithEvents m_ListView As MSComctlLib.ListView
Attribute m_ListView.VB_VarHelpID = -1
#End If

#If DETECT_INTERNET_CONTROLS Then
Private WithEvents m_WebBrowser As SHDocVw.WebBrowser
Attribute m_WebBrowser.VB_VarHelpID = -1
Private WithEvents m_HTMLDocument As MSHTML.HTMLDocument
Attribute m_HTMLDocument.VB_VarHelpID = -1
#End If

'Late-binded control/form
Private m_control As Object

#If VBA7 Then
    Private m_hWndForm As LongPtr
#Else
    Private m_hWndForm As Long
#End If

#If VBA7 Then
Public Function CreateFromControl(ByVal formControl As MSForms.Control, ByVal hWndForm As LongPtr) As MouseOverControl
#Else
Public Function CreateFromControl(ByVal formControl As MSForms.Control, ByVal hWndForm As Long) As MouseOverControl
#End If
    Dim moControl As MouseOverControl
    Set moControl = New MouseOverControl
    moControl.InitFromControl formControl, hWndForm
    Set CreateFromControl = moControl
End Function

#If VBA7 Then
Public Function CreateFromForm(ByVal uForm As MSForms.UserForm, ByVal hWndForm As LongPtr) As MouseOverControl
#Else
Public Function CreateFromForm(ByVal uForm As MSForms.UserForm, ByVal hWndForm As Long) As MouseOverControl
#End If
    Dim moControl As MouseOverControl
    Set moControl = New MouseOverControl
    moControl.InitFromForm uForm, hWndForm
    Set CreateFromForm = moControl
End Function

#If VBA7 Then
Public Sub InitFromControl(ByVal formControl As MSForms.Control, ByVal hWndForm As LongPtr)
#Else
Public Sub InitFromControl(ByVal formControl As MSForms.Control, ByVal hWndForm As Long)
#End If
    Set m_control = formControl
    m_hWndForm = hWndForm
    Select Case TypeName(formControl)
        Case "CheckBox"
            Set m_CheckBox = formControl
        Case "ComboBox"
            Set m_ComboBox = formControl
        Case "CommandButton"
            Set m_CommandButton = formControl
        Case "Frame"
            Set m_Frame = formControl
        Case "Image"
            Set m_Image = formControl
        Case "Label"
            Set m_Label = formControl
        Case "ListBox"
            Set m_ListBox = formControl
        Case "MultiPage"
            Set m_MultiPage = formControl
        Case "OptionButton"
            Set m_OptionButton = formControl
        Case "TabStrip"
            Set m_TabStrip = formControl
        Case "TextBox"
            Set m_TextBox = formControl
        Case "ToggleButton"
            Set m_ToggleButton = formControl
        Case Else
            #If DETECT_COMMON_CONTROLS Then
            If TypeName(formControl) Like "ListView*" Then Set m_ListView = formControl
            #End If
            #If DETECT_INTERNET_CONTROLS Then
            If TypeName(formControl) Like "WebBrowser*" Then Set m_WebBrowser = formControl
            #End If
    End Select
End Sub

#If VBA7 Then
Public Sub InitFromForm(ByVal uForm As MSForms.UserForm, ByVal hWndForm As LongPtr)
#Else
Public Sub InitFromForm(ByVal uForm As MSForms.UserForm, ByVal hWndForm As Long)
#End If
    Set m_control = uForm
    m_hWndForm = hWndForm
    Set m_UserForm = uForm
End Sub

'MouseMove Events
Private Sub m_CheckBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl Me
End Sub
Private Sub m_ComboBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl Me
End Sub
Private Sub m_CommandButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl Me
End Sub
Private Sub m_Frame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl Me
End Sub
Private Sub m_WebBrowser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    #If DETECT_INTERNET_CONTROLS Then
    Set m_HTMLDocument = m_WebBrowser.Document
    #End If
End Sub
Private Sub m_HTMLDocument_onmousemove()
    SetHoveredControl Me
End Sub
Private Sub m_Image_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl Me
End Sub
Private Sub m_Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl Me
End Sub
Private Sub m_ListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl Me
End Sub
Private Sub m_ListView_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
    SetHoveredControl Me
End Sub
Private Sub m_MultiPage_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl Me
End Sub
Private Sub m_OptionButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl Me
End Sub
Private Sub M_TabStrip_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl Me
End Sub
Private Sub m_TextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl Me
End Sub
Private Sub m_ToggleButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl Me
End Sub
Private Sub m_UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl Me
End Sub

'Scroll Event for Containers. This cancels out any scroll that is triggered by a
'   control getting the focus - especially TextBoxes on .SetFocus
Private Sub m_UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
    If ActionY = fmScrollActionFocusRequest Or ActionY = fmScrollActionControlRequest Then ActualDy = 0
    If ActionX = fmScrollActionFocusRequest Or ActionX = fmScrollActionControlRequest Then ActualDx = 0
End Sub
Private Sub m_Frame_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
    If ActionY = fmScrollActionFocusRequest Or ActionY = fmScrollActionControlRequest Then ActualDy = 0
    If ActionX = fmScrollActionFocusRequest Or ActionX = fmScrollActionControlRequest Then ActualDx = 0
End Sub
Private Sub m_MultiPage_Scroll(ByVal Index As Long, ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
    If ActionY = fmScrollActionFocusRequest Or ActionY = fmScrollActionControlRequest Then ActualDy = 0
    If ActionX = fmScrollActionFocusRequest Or ActionX = fmScrollActionControlRequest Then ActualDx = 0
End Sub

Public Function GetControl() As Object
    Set GetControl = m_control
End Function

#If VBA7 Then
Public Property Get FormHandle() As LongPtr
#Else
Public Property Get FormHandle() As Long
#End If
    FormHandle = m_hWndForm
End Property

Private Sub Class_Terminate()
    Set m_CheckBox = Nothing
    Set m_ComboBox = Nothing
    Set m_CommandButton = Nothing
    Set m_Frame = Nothing
    Set m_Image = Nothing
    Set m_Label = Nothing
    Set m_ListBox = Nothing
    Set m_MultiPage = Nothing
    Set m_OptionButton = Nothing
    Set m_TabStrip = Nothing
    Set m_TextBox = Nothing
    Set m_ToggleButton = Nothing
    Set m_UserForm = Nothing
End Sub
