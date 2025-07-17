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
''    - This object holds a single reference to a MSForms.UserForm or a
''      MSForms.Control that is capable of having a MouseMove event
''    - The reference is set using the InitFromControl method and can be any of
''      the 13 types listed below
''    - When the MouseMove event is triggered, the hovered control is passed to
''      the SetHoveredControl method in the MouseScroll module
''    - This class is also used as an asyncronous callback via the Terminate
'       event. In this case events are not used. This choice removed the need
''      to use a new class just for the Terminate event
'' Notes:
''    This class implements Scroll events for the container types (frame, form
''       and multiPage) in order to supress any scrolls that are related to the
''       controls getting the focus.
'' Requires:
''    - MouseScroll: module enabling Mouse Message processing
''==============================================================================

Option Explicit

#If Mac = 0 Then
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    #End If
#End If
#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If
#If Win64 Then
    Private Const PTR_SIZE As Long = 8
#Else
    Private Const PTR_SIZE As Long = 4
#End If

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

Implements DataObject
Implements IFontEventsDisp

Private m_data As MSForms.DataObject
Private m_callbackPtr As LongPtr
Private m_control As Object 'Late-binded control/form
Private m_hWndForm As LongPtr
Private m_isCallback As Boolean

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
        Case "CheckBox":      Set m_CheckBox = formControl
        Case "ComboBox":      Set m_ComboBox = formControl
        Case "CommandButton": Set m_CommandButton = formControl
        Case "Frame":         Set m_Frame = formControl
        Case "Image":         Set m_Image = formControl
        Case "Label":         Set m_Label = formControl
        Case "ListBox":       Set m_ListBox = formControl
        Case "MultiPage":     Set m_MultiPage = formControl
        Case "OptionButton":  Set m_OptionButton = formControl
        Case "TabStrip":      Set m_TabStrip = formControl
        Case "TextBox":       Set m_TextBox = formControl
        Case "ToggleButton":  Set m_ToggleButton = formControl
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
Private Sub m_CheckBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single):                                 SetHoveredControl Me: End Sub
Private Sub m_ComboBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single):                                 SetHoveredControl Me: End Sub
Private Sub m_CommandButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single):                            SetHoveredControl Me: End Sub
Private Sub m_Frame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single):                                    SetHoveredControl Me: End Sub
Private Sub m_Image_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single):                                    SetHoveredControl Me: End Sub
Private Sub m_Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single):                                    SetHoveredControl Me: End Sub
Private Sub m_ListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single):                                  SetHoveredControl Me: End Sub
Private Sub m_ListView_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS): SetHoveredControl Me: End Sub
Private Sub m_MultiPage_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single):           SetHoveredControl Me: End Sub
Private Sub m_OptionButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single):                             SetHoveredControl Me: End Sub
Private Sub M_TabStrip_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single):            SetHoveredControl Me: End Sub
Private Sub m_TextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single):                                  SetHoveredControl Me: End Sub
Private Sub m_ToggleButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single):                             SetHoveredControl Me: End Sub
Private Sub m_UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single):                                 SetHoveredControl Me: End Sub

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

Public Property Let IsAsyncCallback(ByVal newValue As Boolean)
    m_isCallback = newValue
End Property

Private Sub Class_Terminate()
    If m_isCallback Then ProcessMouseData
End Sub

Private Property Get MemLongPtr(ByVal memAddress As LongPtr) As LongPtr
    CopyMemory MemLongPtr, ByVal memAddress, PTR_SIZE
End Property
Private Property Let MemLongPtr(ByVal memAddress As LongPtr, ByVal newValue As LongPtr)
    CopyMemory ByVal memAddress, newValue, PTR_SIZE
End Property

'https://github.com/cristianbuse/VBA-StateLossCallback
#If VBA7 Then
Public Sub InitStateLossCallback(ByVal memAddress As LongPtr)
#Else
Public Sub InitStateLossCallback(ByVal memAddress As Long)
#End If
    m_callbackPtr = memAddress
    Dim dataInterfacePtr As LongPtr: dataInterfacePtr = VarPtr(DataObject) + PTR_SIZE
    Dim dataVirtTablePtr As LongPtr: dataVirtTablePtr = MemLongPtr(dataInterfacePtr)
    Dim dataAddRefAddr As LongPtr:   dataAddRefAddr = dataVirtTablePtr + PTR_SIZE * 1
    Dim dataReleaseAddr As LongPtr:  dataReleaseAddr = dataVirtTablePtr + PTR_SIZE * 2
    Dim dataClearAddr As LongPtr:    dataClearAddr = dataVirtTablePtr + PTR_SIZE * 7
    Dim dataGetClipAddr As LongPtr:  dataGetClipAddr = dataVirtTablePtr + PTR_SIZE * 12
    '
    MemLongPtr(VarPtr(m_data)) = dataInterfacePtr            'Unmanaged reference
    MemLongPtr(dataReleaseAddr) = MemLongPtr(dataClearAddr)  'Main reroute
    MemLongPtr(dataAddRefAddr) = MemLongPtr(dataGetClipAddr) 'See GetFromClipboard
End Sub


'*******************************************************************************
'Called when the m_data instance is being cleared by VB
'*******************************************************************************
Private Sub DataObject_Clear()
    If Not m_data Is Nothing Then Exit Sub
    '
    Dim eFont As stdole.IFontEventsDisp: Set eFont = Me
    MemLongPtr(MemLongPtr(ObjPtr(eFont)) + PTR_SIZE * 7) = m_callbackPtr
    eFont.FontChanged vbNullString
End Sub

'*******************************************************************************
'Called instead of AddRef when an external cast to DataObject is executed
'Avoids incrementing the reference count thus making sure there are no leaks
'*******************************************************************************
Private Sub DataObject_GetFromClipboard(): End Sub

'The following methods are required by the 2 implemented interfaces but are not used
Private Function DataObject_GetFormat(ByVal Format As Variant) As Boolean: End Function
Private Function DataObject_GetText(Optional ByVal Format As Variant) As String: End Function
Private Sub DataObject_PutInClipboard(): End Sub
Private Sub DataObject_SetText(ByVal Text As String, Optional ByVal Format As Variant): End Sub
Private Function DataObject_StartDrag(Optional ByVal OKEffect As Variant) As MSForms.fmDropEffect: End Function
Private Sub IFontEventsDisp_FontChanged(ByVal PropertyName As String): End Sub
