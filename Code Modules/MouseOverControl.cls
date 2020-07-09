VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MouseOverControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''=============================================================================
''' VBA UserForm MouseScroll
'''---------------------------------------------------------
''' https://github.com/cristianbuse/VBA-UserForm-MouseScroll
'''---------------------------------------------------------
'''
''' Copyright (c) 2019 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to deal
''' in the Software without restriction, including without limitation the rights
''' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
''' copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in all
''' copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
''' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
''' SOFTWARE.
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

Public Sub InitFromControl(ByVal objControl As Object)
    Select Case TypeName(objControl)
        Case "CheckBox"
            Set m_CheckBox = objControl
        Case "ComboBox"
            Set m_ComboBox = objControl
        Case "CommandButton"
            Set m_CommandButton = objControl
        Case "Frame"
            Set m_Frame = objControl
        Case "Image"
            Set m_Image = objControl
        Case "Label"
            Set m_Label = objControl
        Case "ListBox"
            Set m_ListBox = objControl
        Case "MultiPage"
            Set m_MultiPage = objControl
        Case "OptionButton"
            Set m_OptionButton = objControl
        Case "TabStrip"
            Set m_TabStrip = objControl
        Case "TextBox"
            Set m_TextBox = objControl
        Case "ToggleButton"
            Set m_ToggleButton = objControl
        Case Else
            If TypeOf objControl Is MSForms.UserForm Then
                Set m_UserForm = objControl
            Else
                'Ignore
            End If
    End Select
End Sub

Public Sub TerminateReferences()
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

'MouseMove Events
Private Sub m_CheckBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl m_CheckBox
End Sub
Private Sub m_ComboBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl m_ComboBox
End Sub
Private Sub m_CommandButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl m_CommandButton
End Sub
Private Sub m_Frame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl m_Frame
End Sub
Private Sub m_Image_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl m_Image
End Sub
Private Sub m_Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl m_Label
End Sub
Private Sub m_ListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl m_ListBox
End Sub
Private Sub m_MultiPage_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl m_MultiPage
End Sub
Private Sub m_OptionButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl m_OptionButton
End Sub
Private Sub M_TabStrip_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl m_TabStrip
End Sub
Private Sub m_TextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl m_TextBox
End Sub
Private Sub m_ToggleButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl m_ToggleButton
End Sub
Private Sub m_UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    SetHoveredControl m_UserForm
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
