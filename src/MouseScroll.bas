Attribute VB_Name = "MouseScroll"
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
''    Allows forms and form controls to be scrolled using the mouse wheel.
''    Simultaneoulsy tracks all forms that called the EnableMouseScroll method!
''
''    Installs a Mouse Hook by calling SetWindowsHookEx API with ID WH_MOUSE = 7
''       and the address of the MouseProc callback function.
''    The Mouse Hook is active as long as there is at least one form that
''       previously enabled scrolling (i.e. called EnableMouseScroll method)
''    Another option would be to use ID WH_MOUSE_LL = 14 which would require a
''       LowLevelMouseProc callback but unlike the WH_MOUSE hook which is local
''       (hooked on the current thread only) the WH_MOUSE_LL hook is actually
''       global and very slow.
''    The system calls the MouseProc function whenever the Excel Application
''       calls the GetMessage or PeekMessage functions and there is a mouse
''       message to be processed.
''    This module also holds a collection of MouseOverControls that call back
''       the SetHoveredControl method in this module whenever a MouseMove event
''       is triggered.
'' Notes:
''    MouseProc hook works properly with MODAL UserForms only!
''    Modeless Forms will cause unhooking! This is done on purpose to prevent
''       crashes!
''    No need to call DisableMouseScroll method. This is done automatically!
''       However, in Ms Word is recommanded to call it from the form's terminate
''       event (i.e. UserForm_Terminate)
''    Hold down SHIFT key when scrolling the mouse wheel, for Horizontal Scroll!
''    Hold down CTRL key when scrolling the mouse wheel, for Zoom!
'' Warning:
''    Do not debug code while the hook is active to avoid crashes!
'' Requires:
''    - MouseOverControl: Container that tracks MouseMove events
''==============================================================================

Option Explicit
Option Private Module

'API declarations
'*******************************************************************************
#If Mac Then
    'No Mac functionality implemented
#Else
    'Windows API functionality
    #If VBA7 Then
        Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
        Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
        Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
        Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
        Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
        Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
        Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
        Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As LongPtr) As Long
        Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
        Private Declare PtrSafe Function IsWindowEnabled Lib "user32" (ByVal hwnd As LongPtr) As Long
        Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As LongPtr) As Long
        Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
        Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
        Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
        Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
    #Else
        Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
        Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
        Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
        Private Declare Function GetActiveWindow Lib "user32" () As Long
        Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
        Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
        Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
        Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
        Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
        Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As Long) As Long
        Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
        Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
        Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    #End If
#End If
'*******************************************************************************

'Id of the hook procedure to be installed with SetWindowsHookExA for MouseProc
Private Const WH_MOUSE As Long = 7

'Necessary API structs and constants for MouseProc Callback
'https://msdn.microsoft.com/en-us/library/windows/desktop/ms644988(v=vs.85).aspx
'*******************************************************************************
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    #If VBA7 Then
        hwnd As LongPtr
    #Else
        hwnd As Long
    #End If
    wHitTestCode As Long
    #If VBA7 Then
        dwExtraInfo As LongPtr
    #Else
        dwExtraInfo As Long
    #End If
End Type

'nCode
Private Const HC_ACTION As Long = 0
Private Const HC_NOREMOVE As Long = 3

'wParam
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_XBUTTONDOWN As Long = &H20B
Private Const WM_XBUTTONUP As Long = &H20C
Private Const WM_XBUTTONDBLCLK As Long = &H20D
Private Const WM_MOUSEHWHEEL As Long = &H20E

'lParam
Private Type MOUSEHOOKSTRUCTEX
    tagMOUSEHOOKSTRUCT As MOUSEHOOKSTRUCT
    mouseData As Long 'DWORD
End Type
'*******************************************************************************

'Necessary struct and constants to calculate the number of lines/pages to scroll
'https://msdn.microsoft.com/en-us/library/ms997498.aspx
'*******************************************************************************
Private Const WHEEL_DELTA As Long = 120
Private Const SPI_GETWHEELSCROLLLINES As Long = &H68

Private Type SCROLL_AMOUNT
    lines As Single
    pages As Single
End Type
'*******************************************************************************

'm_hHookMouse - Hook handle obtained from a previous call to SetWindowsHookEx
'   - Used when calling UnhookWindowsHookEx in order to remove the hook
'm_hWndMainOwner - Main UserForm's Owner Handle (to track Modal state)
'*******************************************************************************
#If VBA7 Then
    Private m_hHookMouse As LongPtr
    Private m_hWndMainOwner As LongPtr
#Else
    Private m_hHookMouse As Long
    Private m_hWndMainOwner As Long
#End If
'*******************************************************************************

'Window handles for all forms with scrolling enabled. Always instantiated
Private m_hWndAllForms As New Collection

'Collection of sub-collections of MouseOverControls (one for each form)
Private m_controls As New Collection

'Keeps track of the passScrollAtMargins option for each form
Private m_passScrollColl As New Collection

'The last control that was hovered (could be the UserForm itself)
Private m_lastHoveredControl As MouseOverControl

'The last ComboBox that was used
Private m_lastCombo As MSForms.ComboBox
Private m_isLastComboOn As Boolean

'The previous state of Application.EnableCancelKey (if available)
Private m_enableCancelKey As Long

'The Scrollable Control Types/Categories
Private Enum CONTROL_TYPE
    ctNone = 0
    ctCombo = 1
    ctList = 2
    ctFrame = 3
    ctPage = 4
    ctMulti = 5
    ctForm = 6
    ctText = 7
    ctOther = 8
End Enum

'The type of Action possible when Mouse Wheel is turned (see MouseProc func)
Private Enum SCROLL_ACTION
    saScrollY = 1
    saScrollX = 2
    saZoom = 3
End Enum

'If the current hovered control cannot scroll anymore, then pass (or not) the
'   scroll to the Parent Control/Form. Variable set in SetHoveredControl()
Private m_passScrollToParentAtMargins As Boolean

'*******************************************************************************
'Enables mouse wheel scroll for the specified UserForm
'*******************************************************************************
Public Function EnableMouseScroll(ByVal uForm As MSForms.UserForm _
    , Optional ByVal passScrollToParentAtMargins As Boolean = True _
) As Boolean
    If uForm Is Nothing Then Exit Function
    If Not HookMouse Then Exit Function
    '
    AddForm uForm, passScrollToParentAtMargins
    ResetLast
    EnableMouseScroll = True
End Function

'*******************************************************************************
'Disables mouse wheel scroll for a specific UserForm. Can be called, optionally,
'   from a form's teminate event but is not needed
'*******************************************************************************
Public Sub DisableMouseScroll(ByVal uForm As MSForms.UserForm)
    RemoveForm GetFormHandle(uForm)
    ResetLast
End Sub

'*******************************************************************************
'Resets cached controls
'*******************************************************************************
Private Sub ResetLast()
    Set m_lastHoveredControl = Nothing
    Set m_lastCombo = Nothing
End Sub

'*******************************************************************************
'Hooks Mouse messages to the MouseProc procedure
'The MouseProc callback will manipulate controls/forms by calling methods like
'   ScrollY and ScrollX
'*******************************************************************************
Private Function HookMouse() As Boolean
    If m_hHookMouse <> 0 Then
        HookMouse = True
        Exit Function
    End If
    '
    Dim isHookSuccessful As Boolean
    '
    #If Mac Then
    #Else
        m_hHookMouse = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, 0, GetCurrentThreadId())
    #End If
    isHookSuccessful = (m_hHookMouse <> 0)
    If isHookSuccessful Then
        On Error Resume Next
        m_enableCancelKey = CallByName(Application, "EnableCancelKey", VbGet)
        CallByName Application, "EnableCancelKey", VbLet, 0
        On Error GoTo 0
        Debug.Print "Mouse hooked " & Now
    End If
    '
    HookMouse = isHookSuccessful
End Function

'*******************************************************************************
'UnHooks Mouse
'*******************************************************************************
Private Sub UnHookMouse()
    If m_hHookMouse <> 0 Then
        #If Mac Then
        #Else
            UnhookWindowsHookEx m_hHookMouse
        #End If
        On Error Resume Next
        CallByName Application, "EnableCancelKey", VbLet, m_enableCancelKey
        On Error GoTo 0
        m_hHookMouse = 0
        Set m_hWndAllForms = Nothing
        Set m_controls = Nothing
        Set m_passScrollColl = Nothing
        Set m_lastHoveredControl = Nothing
        Set m_lastCombo = Nothing
        Debug.Print "Mouse unhooked " & Now
    End If
End Sub

'*******************************************************************************
'Adds the form handle to m_hWndAllForms collection
'Adds the passScrollAtMargins option to m_passScrollColl collection
'Adds a sub-collection of MouseMove controls to m_controls collection
'*******************************************************************************
Private Sub AddForm(ByVal uForm As MSForms.UserForm, ByVal passScrollAtMargins As Boolean)
    #If VBA7 Then
        Dim hWndForm As LongPtr
    #Else
        Dim hWndForm As Long
    #End If
    Dim key_ As String
    '
    hWndForm = GetFormHandle(uForm)
    key_ = CStr(hWndForm)
    '
    If CollectionHasKey(m_hWndAllForms, key_) Then
        m_controls.Remove key_
        m_passScrollColl.Remove key_
    Else
        m_hWndAllForms.Add hWndForm, key_
    End If
    m_passScrollColl.Add passScrollAtMargins, key_
    '
    If m_controls.Count = 0 Then
        'Keep track of the owner of the first form only
        m_hWndMainOwner = GetOwnerHandle(hWndForm)
    End If
    '
    Dim subControls As Collection
    Set subControls = New Collection
    m_controls.Add subControls, key_
    '
    Dim frmCtrl As MSForms.Control
    '
    For Each frmCtrl In uForm.Controls
        subControls.Add MouseOverControl.CreateFromControl(frmCtrl, hWndForm)
    Next frmCtrl
    subControls.Add MouseOverControl.CreateFromForm(uForm, hWndForm)
End Sub

'*******************************************************************************
'Removes a form (by window handle) from the internal collections
'*******************************************************************************
#If VBA7 Then
Private Sub RemoveForm(ByVal hWndForm As LongPtr)
#Else
Private Sub RemoveForm(ByVal hWndForm As Long)
#End If
    If CollectionHasKey(m_hWndAllForms, hWndForm) Then
        Dim key_ As String: key_ = CStr(hWndForm)
        m_hWndAllForms.Remove key_
        m_controls.Remove key_
        m_passScrollColl.Remove key_
    End If
    If m_hWndAllForms.Count = 0 Then UnHookMouse
End Sub

'*******************************************************************************
'Removes any form that has been destroyed
'*******************************************************************************
Private Sub RemoveDestroyedForms()
    Dim v As Variant
    '
    For Each v In m_hWndAllForms
        If Not CBool(IsWindow(v)) Then
            RemoveForm v
        End If
    Next v
End Sub

'*******************************************************************************
'Returns a boolean indicating if a Collection has a specific key
'Parameters:
'   - coll: a collection to check for key
'   - key_: the key being searched for
'Does not raise errors
'*******************************************************************************
Private Function CollectionHasKey(ByVal coll As Collection, ByVal key_ As String) As Boolean
    On Error Resume Next
    coll.Item key_
    CollectionHasKey = (Err.Number = 0)
    On Error GoTo 0
End Function

'*******************************************************************************
'Called by MouseMove capable controls (MouseOverControl) stored in m_controls
'*******************************************************************************
Public Sub SetHoveredControl(ByVal moCtrl As MouseOverControl)
    Set m_lastHoveredControl = moCtrl
    On Error Resume Next
    m_passScrollToParentAtMargins = m_passScrollColl(CStr(moCtrl.FormHandle))
    On Error GoTo 0
    UpdateLastCombo
End Sub

'*******************************************************************************
'Keeps track of last combo box to avoid scrolling other controls while the combo
'   is expanded
'*******************************************************************************
Private Sub UpdateLastCombo()
    On Error Resume Next
    Set m_lastCombo = m_lastHoveredControl.GetControl
    On Error GoTo 0
    If Not m_lastCombo Is Nothing Then
        m_isLastComboOn = (m_lastCombo.TopIndex >= 0)
    End If
End Sub

'*******************************************************************************
'Callback hook function - monitors mouse messages
'*******************************************************************************
#If Mac Then
#Else
#If VBA7 Then
Private Function MouseProc(ByVal ncode As Long _
                         , ByVal wParam As Long _
                         , ByRef lParam As MOUSEHOOKSTRUCTEX) As LongPtr
#Else
Private Function MouseProc(ByVal ncode As Long _
                         , ByVal wParam As Long _
                         , ByRef lParam As MOUSEHOOKSTRUCTEX) As Long
#End If
    'Unhook if a VBE window is active
    If IsVBEActive Then GoTo Unhook
    '
    RemoveDestroyedForms
    '
    'Unhook if no form handles
    If m_hWndAllForms.Count = 0 Then GoTo Unhook
    '
    'Unhook if the top Owner is active (Modeless Form)
    If CBool(IsWindowEnabled(m_hWndMainOwner)) Then GoTo Unhook
    '
    If m_lastHoveredControl Is Nothing Then GoTo NextHook
    'Ignore input if Window matching last hovered control is not Active
    If Not CBool(IsWindowEnabled(m_lastHoveredControl.FormHandle)) Then GoTo NextHook
    '
    'The nCode could either be negative, HC_ACTION or HC_NOREMOVE
    'HC_NOREMOVE is passed when the Application calls the PeekMessage function
    '   with a PM_NOREMOVE flag which means that the mouse message has not been
    '   removed from the message queue
    'In case of negative or HC_NOREMOVE nCode the function will pass the message
    '   to the CallNextHookEx function and return it's value
    If ncode = HC_ACTION Then
        If wParam = WM_MOUSEWHEEL Or wParam = WM_MOUSEHWHEEL Then
            If TypeName(m_lastHoveredControl.GetControl) Like "ListView*" Then GoTo NextHook
            '
            Dim scrollAmount As SCROLL_AMOUNT
            Dim scrollAction As SCROLL_ACTION
            '
            scrollAmount = GetScrollAmount(GetWheelDelta(lParam.mouseData))
            scrollAction = GetScrollAction(yWheel:=(wParam = WM_MOUSEWHEEL))
            '
            If m_isLastComboOn Then
                m_passScrollToParentAtMargins = False
                Call ScrollY(m_lastCombo, scrollAmount)
            Else
                Select Case scrollAction
                Case saScrollY
                    Call ScrollY(m_lastHoveredControl.GetControl, scrollAmount)
                Case saScrollX
                    If m_isLastComboOn Then GoTo NextHook
                    Call ScrollX(m_lastHoveredControl.GetControl, scrollAmount)
                Case saZoom
                    If m_isLastComboOn Then GoTo NextHook
                    Call Zoom(m_lastHoveredControl.GetControl, scrollAmount)
                End Select
            End If
            '
            MouseProc = -1
            Exit Function
        Else
            'Here you could implement logic for:
            'WM_MOUSEMOVE
            'WM_LBUTTONDOWN
            'WM_LBUTTONUP
            'WM_LBUTTONDBLCLK
            'WM_RBUTTONDOWN
            'WM_RBUTTONUP
            'WM_RBUTTONDBLCLK
            'WM_MBUTTONDOWN
            'WM_MBUTTONUP
            'WM_MBUTTONDBLCLK
            '
            'Mouse side buttons example:
            If wParam = WM_XBUTTONDOWN Then
                Const HIGH_VALUE  As Single = 10000000
                '
                If lParam.mouseData = &H20000 Then
                    scrollAmount.lines = HIGH_VALUE
                    ScrollY m_lastHoveredControl.GetControl, scrollAmount
                ElseIf lParam.mouseData = &H10000 Then
                    scrollAmount.lines = -HIGH_VALUE
                    ScrollY m_lastHoveredControl.GetControl, scrollAmount
                End If
                MouseProc = -1
                Exit Function
            End If
            '
            'For now, just passing the message to (CallNextHookEx)
        End If
    End If
    '
NextHook:
    MouseProc = CallNextHookEx(0, ncode, wParam, ByVal lParam)
Exit Function
Unhook:
    UnHookMouse
    GoTo NextHook
End Function
#End If

'*******************************************************************************
'Get the type of scroll action by reading Shift and Control key states
'*******************************************************************************
Private Function GetScrollAction(ByVal yWheel As Boolean) As SCROLL_ACTION
    If yWheel Then
        If IsShiftKeyDown() Then
            GetScrollAction = saScrollX
        ElseIf IsControlKeyDown() Then
            GetScrollAction = saZoom
        Else
            GetScrollAction = saScrollY
        End If
    Else
        If IsShiftKeyDown() Then
            GetScrollAction = saScrollY
        ElseIf IsControlKeyDown() Then
            GetScrollAction = saZoom
        Else
            GetScrollAction = saScrollX
        End If
    End If
End Function

'*******************************************************************************
'Get the wheel delta from mouseData Double Word's HiWord
'The LoWord is undefined and reserved
'*******************************************************************************
Private Function GetWheelDelta(ByVal dwMouseData As Long) As Integer
    GetWheelDelta = HiWord(dwMouseData)
End Function

'*******************************************************************************
'Function to retrieve the High Word (16-bit) from a Double Word (32-bit)
'*******************************************************************************
Private Function HiWord(ByVal dWord As Long) As Integer
    HiWord = VBA.Int(dWord / &H10000)
End Function

'*******************************************************************************
'Get the scroll amount (lines or pages) for a mouse wheel scroll value
'*******************************************************************************
Private Function GetScrollAmount(ByVal scrollValue As Integer) As SCROLL_AMOUNT
    Dim systemScrollLines As Long: systemScrollLines = GetUserScrollLines()
    Dim scrollAmount As SCROLL_AMOUNT
    '
    If systemScrollLines = -1 Then
        scrollAmount.pages = scrollValue / WHEEL_DELTA
    Else
        scrollAmount.lines = scrollValue / WHEEL_DELTA * systemScrollLines
    End If
    '
    GetScrollAmount = scrollAmount
End Function

'*******************************************************************************
'Get the number of scroll lines (or page = -1) that are set in the system
'*******************************************************************************
Private Function GetUserScrollLines() As Long
    Dim result As Long: result = 3 'default
    '
    #If Mac Then
    #Else
        SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, result, 0
    #End If
    GetUserScrollLines = result
End Function

'*******************************************************************************
'Vertically scroll a control or the hooked Form itself
'*******************************************************************************
Private Sub ScrollY(ByVal ctrl As Object, ByRef scrollAmount As SCROLL_AMOUNT)
    Const scrollPointsPerLine As Single = 6
    Dim ctrlType As CONTROL_TYPE: ctrlType = GetControlType(ctrl)
    '
    Select Case ctrlType
        Case ctNone
            Exit Sub
        Case ctCombo, ctList
            Call ListScrollY(ctrl, scrollAmount, ctrlType)
        Case ctFrame, ctPage, ctMulti, ctForm
            If ctrlType = ctMulti Then
                Set ctrl = ctrl.SelectedItem
                ctrlType = ctPage
            End If
            '
            Dim lastScrollTop As Single
            Dim newScrollTop As Single
            Dim maxScroll As Single
            '
            'Store the Top position of the scroll. Can throw - must guard
            On Error Resume Next
            lastScrollTop = ctrl.ScrollTop
            If Err.Number <> 0 Then
                Err.Clear
                Exit Sub
            End If
            On Error GoTo 0
            '
            'Compute the new Top position
            newScrollTop = lastScrollTop _
                - scrollAmount.lines * scrollPointsPerLine _
                - scrollAmount.pages * ctrl.InsideHeight
            '
            'Clamp the new scroll value
            maxScroll = ctrl.ScrollHeight - ctrl.InsideHeight
            If newScrollTop > maxScroll Then newScrollTop = maxScroll
            If newScrollTop < 0 Then newScrollTop = 0
            '
            'Apply new scroll if needed
            If ctrl.ScrollTop <> newScrollTop Then
                ctrl.ScrollTop = newScrollTop
                If ctrlType = ctForm Then ctrl.Repaint
            End If
            '
            If m_passScrollToParentAtMargins Then
                'If scroll hasn't changed pass scroll to parent control
                If ctrl.ScrollTop = lastScrollTop And ctrlType <> ctForm Then
                    If ctrlType = ctPage Then Set ctrl = ctrl.Parent 'Multi
                    Call ScrollY(ctrl.Parent, scrollAmount)
                End If
            End If
        Case ctText
            Call TBoxScrollY(ctrl, scrollAmount)
        Case Else
            'Control is not scrollable. Pass scroll to parent
            Dim parentCtrlType As CONTROL_TYPE
            '
            On Error Resume Next 'Necessary during Form Init
            parentCtrlType = GetControlType(ctrl.Parent)
            On Error GoTo 0
            If parentCtrlType <> ctNone Then ScrollY ctrl.Parent, scrollAmount
    End Select
End Sub

'*******************************************************************************
'Vertically scroll a ComboBox or a ListBox control
'*******************************************************************************
Private Sub ListScrollY(ByVal ctrl As Object _
                      , ByRef scrollAmount As SCROLL_AMOUNT _
                      , ByVal ctrlType As CONTROL_TYPE _
)
    Dim lastTopIndex As Long: lastTopIndex = ctrl.TopIndex
    Dim newTopIndex As Long
    '
    If scrollAmount.lines <> 0 Then
        newTopIndex = lastTopIndex - scrollAmount.lines
    Else
        Dim linesPerPage As Long
        '
        If ctrlType = ctCombo Then
            linesPerPage = ctrl.ListRows
        Else
            ctrl.TopIndex = ctrl.ListCount - 1
            linesPerPage = VBA.Int(ctrl.ListCount - ctrl.TopIndex)
            ctrl.TopIndex = lastTopIndex
        End If
        newTopIndex = lastTopIndex - scrollAmount.pages * linesPerPage
    End If
    '
    'Clamp the new scroll top
    If newTopIndex < 0 Then
        newTopIndex = 0
    ElseIf newTopIndex >= ctrl.ListCount Then
        newTopIndex = ctrl.ListCount - 1
    End If
    '
    On Error Resume Next 'could fail for undropped ComboBox
    If lastTopIndex <> newTopIndex Then ctrl.TopIndex = newTopIndex
    If Err.Number <> 0 Then
        Err.Clear
        Call ScrollY(ctrl.Parent, scrollAmount)
        Exit Sub
    End If
    On Error GoTo 0
    '
    If m_passScrollToParentAtMargins Then
        If ctrl.TopIndex = lastTopIndex Then
            Call ScrollY(ctrl.Parent, scrollAmount)
        End If
    End If
End Sub

'*******************************************************************************
'Vertically scroll a TextBox control
'*******************************************************************************
Private Sub TBoxScrollY(ByVal tbox As MSForms.TextBox, ByRef scrollAmount As SCROLL_AMOUNT)
    If Not tbox.MultiLine Then
        Call ScrollY(tbox.Parent, scrollAmount)
        Exit Sub
    End If
    '
    'Store current Selection (to be reverted later)
    tbox.SetFocus
    Dim lastY As Long: lastY = tbox.CurY
    Dim selectionStart As Long: selectionStart = tbox.SelStart
    Dim selectionLength As Long: selectionLength = tbox.SelLength
    '
    'Compute line metrics
    Dim deltaLines As Long
    Dim lineHeight As Single: lineHeight = GetTextBoxLineHeight(tbox)
    Dim linesPerPage As Long: linesPerPage = VBA.Int(tbox.Height / lineHeight)
    '
    'Lines to scroll up/down
    If scrollAmount.lines <> 0 Then
        deltaLines = -scrollAmount.lines
    Else
        deltaLines = -scrollAmount.pages * linesPerPage
    End If
    '
    'Jump to top/bottom line of the "page"
    Const topOffsetPt As Single = 3 'the extra 3 points at the top of a tbox
    If deltaLines > 0 Then
        tbox.CurY = PointsToHiMeter(topOffsetPt + linesPerPage * lineHeight)
    ElseIf deltaLines < 0 Then
        tbox.CurY = PointsToHiMeter(topOffsetPt)
    End If
    '
    Dim lastLine As Long: lastLine = tbox.CurLine
    Dim newline As Long: newline = lastLine + deltaLines
    '
    'Clamp the new scroll line
    If newline < 0 Then
        newline = 0
    ElseIf newline >= tbox.LineCount Then
        newline = tbox.LineCount - 1
    End If
    If lastLine <> newline Then tbox.CurLine = newline
    '
    'Restore Selection. Must hide (or disable) textBox first, to lock scroll
    tbox.Visible = False
    tbox.SelStart = selectionStart
    tbox.SelLength = selectionLength
    tbox.Visible = True
    tbox.SetFocus
    '
    If m_passScrollToParentAtMargins Then
        Dim currentY As Long: currentY = tbox.CurY
        'Adjustment in case the top of the textbox is outside the parent scroll
        Const topAdjust As Long = 1734040
        '
        If currentY > topAdjust Then currentY = currentY - topAdjust
        If lastY > topAdjust Then lastY = lastY - topAdjust
        '
        If currentY = lastY Then Call ScrollY(tbox.Parent, scrollAmount)
    End If
End Sub

'*******************************************************************************
'Get the row height for a TextBox by using the AutoSize feature
'*******************************************************************************
Private Function GetTextBoxLineHeight(ByVal tbox As MSForms.TextBox) As Single
    tbox.SetFocus
    'Store Size and appearance
    Dim oldHeight As Single: oldHeight = tbox.Height
    Dim oldWidth As Single: oldWidth = tbox.Width
    Dim isVisible As Boolean: isVisible = tbox.Visible
    Dim isAutoSized As Boolean: isAutoSized = tbox.AutoSize
    Dim borderSt As fmBorderStyle: borderSt = tbox.BorderStyle
    Dim spEffect As fmSpecialEffect: spEffect = tbox.SpecialEffect
    Dim scBars As fmScrollBars: scBars = tbox.ScrollBars
    Dim linesCount As Long: linesCount = tbox.LineCount
    '
    Dim lineHeight As Single
    Const topOffsetPt As Single = 3 'the extra 3 points at the top of a tbox
    '
    'Make sure AutoSize is triggered
    If isVisible Then tbox.Visible = False
    If isAutoSized Then tbox.AutoSize = False
    If tbox.WordWrap Then
        Select Case scBars
            Case fmScrollBars.fmScrollBarsHorizontal
                tbox.ScrollBars = fmScrollBarsNone
            Case fmScrollBars.fmScrollBarsBoth
                tbox.ScrollBars = fmScrollBarsVertical
        End Select
    End If
    tbox.BorderStyle = fmBorderStyleNone
    tbox.SpecialEffect = fmSpecialEffectFlat
    tbox.AutoSize = True
    '
    'If the last line is empty then the AutoSize is ignoring it and an
    '   adjustment is needed for the total line count
    If VBA.Right$(tbox.text, 2) = vbNewLine Then linesCount = linesCount - 1
    lineHeight = (tbox.Height - topOffsetPt) / linesCount
    '
    'Restore TextBox state
    tbox.AutoSize = isAutoSized
    If tbox.BorderStyle <> borderSt Then tbox.BorderStyle = borderSt
    If tbox.SpecialEffect <> spEffect Then tbox.SpecialEffect = spEffect
    If tbox.ScrollBars <> scBars Then tbox.ScrollBars = scBars
    tbox.Height = oldHeight
    tbox.Width = oldWidth
    tbox.Visible = isVisible
    tbox.SetFocus
    '
    'Return result
    GetTextBoxLineHeight = lineHeight
End Function

'*******************************************************************************
'Horizontally scroll a control or the hooked Form itself
'Code is very similar to the ScrollY method with main difference being that
'   all values are relative to the Left instead of the Top side
'*******************************************************************************
Private Sub ScrollX(ByVal ctrl As Object, ByRef scrollAmount As SCROLL_AMOUNT)
    Const scrollPointsPerColumn As Single = 15
    Dim ctrlType As CONTROL_TYPE: ctrlType = GetControlType(ctrl)
    '
    Select Case ctrlType
        Case ctNone
            Exit Sub
        Case ctList
            Call ListScrollX(ctrl, scrollAmount)
        Case ctFrame, ctPage, ctMulti, ctForm
            If ctrlType = ctMulti Then
                Set ctrl = ctrl.SelectedItem
                ctrlType = ctPage
            End If
            '
            Dim lastScrollLeft As Single
            Dim newScrollLeft As Single
            Dim maxScroll As Single
            '
            'Store the Left position of the scroll. Can throw - must guard
            On Error Resume Next
            lastScrollLeft = ctrl.ScrollLeft
            If Err.Number <> 0 Then
                Err.Clear
                Exit Sub
            End If
            On Error GoTo 0
            '
            'Compute the new Left position
            newScrollLeft = lastScrollLeft _
                - scrollAmount.lines * scrollPointsPerColumn _
                - scrollAmount.pages * ctrl.InsideWidth
            '
            'Clamp the new scroll value
            maxScroll = ctrl.ScrollWidth - ctrl.InsideWidth
            If newScrollLeft > maxScroll Then newScrollLeft = maxScroll
            If newScrollLeft < 0 Then newScrollLeft = 0
            '
            'Apply new scroll if needed
            If ctrl.ScrollLeft <> newScrollLeft Then
                ctrl.ScrollLeft = newScrollLeft
                If ctrlType = ctForm Then ctrl.Repaint
            End If
            '
            'If scroll hasn't changed pass scroll to parent control
            If m_passScrollToParentAtMargins Then
                If ctrl.ScrollLeft = lastScrollLeft And ctrlType <> ctForm Then
                    If ctrlType = ctPage Then Set ctrl = ctrl.Parent 'Multi
                    ScrollX ctrl.Parent, scrollAmount
                End If
            End If
        Case Else
            'Control is not scrollable. Pass scroll to parent
            Dim parentCtrlType As CONTROL_TYPE
            '
            On Error Resume Next 'Necessary during Form Init
            parentCtrlType = GetControlType(ctrl.Parent)
            On Error GoTo 0
            If parentCtrlType <> ctNone Then ScrollX ctrl.Parent, scrollAmount
    End Select
End Sub

'*******************************************************************************
'Horizontally scroll a ListBox control
'*******************************************************************************
Private Sub ListScrollX(ByVal lbox As MSForms.Control, ByRef scrollAmount As SCROLL_AMOUNT)
    Const WM_KEYDOWN As Long = &H100
    Const VK_LEFT = &H25
    Const VK_RIGHT = &H27
    Const colsPerPage As Long = 15
    '
    Dim msgCount As Long
    '
    msgCount = scrollAmount.lines + scrollAmount.pages * colsPerPage
    lbox.SetFocus
    If msgCount > 0 Then
        'A single left key will considerably move the scroll bar
        PostMessage lbox.[_GethWnd], WM_KEYDOWN, VK_LEFT, 0
    Else
        Dim i As Long
        '
        For i = 1 To Math.Abs(msgCount)
            PostMessage lbox.[_GethWnd], WM_KEYDOWN, VK_RIGHT, 0
        Next i
    End If
End Sub

'*******************************************************************************
'Zooms controls using mouse scroll
'*******************************************************************************
Private Sub Zoom(ByVal ctrl As Object, ByRef scrollAmount As SCROLL_AMOUNT)
    Const minZoom As Integer = 10
    Const maxZoom As Integer = 400
    Dim ctrlType As CONTROL_TYPE: ctrlType = GetControlType(ctrl)
    '
    Select Case ctrlType
        Case ctNone
            Exit Sub
        Case ctFrame, ctPage, ctMulti, ctForm
            If ctrlType = ctMulti Then
                Set ctrl = ctrl.SelectedItem
                ctrlType = ctPage
            End If
            '
            Dim lastZoom As Single
            Dim newZoom As Single
            '
            lastZoom = ctrl.Zoom
            '
            'Compute the new zoom
            newZoom = lastZoom + scrollAmount.lines * 5 + scrollAmount.pages * 25
            '
            'Clamp the new zoom value
            If newZoom > maxZoom Then newZoom = maxZoom
            If newZoom < minZoom Then newZoom = minZoom
            '
            'Apply new zoom if needed
            If lastZoom <> newZoom Then
                ctrl.Zoom = newZoom
                If ctrlType = ctForm Then ctrl.Repaint
            End If
            '
            'If zoom hasn't changed pass zoom to parent control
            If m_passScrollToParentAtMargins Then
                If ctrl.Zoom = lastZoom And ctrlType <> ctForm Then
                    If ctrlType = ctPage Then Set ctrl = ctrl.Parent 'Multi
                    Zoom ctrl.Parent, scrollAmount
                End If
            End If
        Case Else
            'Control cannot be zoomed. Pass zoom to parent
            Dim parentCtrlType As CONTROL_TYPE
            '
            On Error Resume Next 'Necessary during Form Init
            parentCtrlType = GetControlType(ctrl.Parent)
            On Error GoTo 0
            If parentCtrlType <> ctNone Then Zoom ctrl.Parent, scrollAmount
    End Select
End Sub

'*******************************************************************************
'Get enum of Control Type
'*******************************************************************************
Private Function GetControlType(ByVal objControl As Object) As CONTROL_TYPE
    If objControl Is Nothing Then
        GetControlType = ctNone
        Exit Function
    End If
    Select Case TypeName(objControl)
        Case "ComboBox"
            GetControlType = ctCombo
        Case "Frame"
            GetControlType = ctFrame
        Case "ListBox"
            GetControlType = ctList
        Case "MultiPage"
            GetControlType = ctMulti
        Case "Page"
            GetControlType = ctPage
        Case "TextBox"
            GetControlType = ctText
        Case Else
            If TypeOf objControl Is MSForms.UserForm Then
                GetControlType = ctForm
            Else
                GetControlType = ctOther
            End If
    End Select
End Function

'*******************************************************************************
'Returns the Window Handle for a UserForm
'https://docs.microsoft.com/en-us/windows/desktop/api/shlwapi/nf-shlwapi-iunknown_getwindow
'*******************************************************************************
#If VBA7 Then
Private Function GetFormHandle(ByVal objForm As MSForms.UserForm) As LongPtr
#Else
Private Function GetFormHandle(ByVal objForm As MSForms.UserForm) As Long
#End If
    #If Mac Then
    #Else
        IUnknown_GetWindow objForm, VBA.VarPtr(GetFormHandle)
    #End If
End Function

'*******************************************************************************
'Returns a Window Owner's Handle
'*******************************************************************************
#If VBA7 Then
Private Function GetOwnerHandle(ByVal hwnd As LongPtr) As LongPtr
#Else
Private Function GetOwnerHandle(ByVal hwnd As Long) As Long
#End If
    Const GW_OWNER As Long = 4
    #If Mac Then
    #Else
        GetOwnerHandle = GetWindow(hwnd, GW_OWNER)
    #End If
End Function

'*******************************************************************************
'Get Shift/Control Key State
'https://docs.microsoft.com/en-us/windows/desktop/api/winuser/nf-winuser-getkeystate
'https://docs.microsoft.com/en-us/windows/desktop/inputdev/virtual-key-codes
'*******************************************************************************
Private Function IsShiftKeyDown() As Boolean
    Const VK_SHIFT As Long = &H10
    '
    IsShiftKeyDown = CBool(GetKeyState(VK_SHIFT) And &H8000) 'hi-order bit only
End Function
Private Function IsControlKeyDown() As Boolean
    Const VK_CONTROL As Long = &H11
    '
    IsControlKeyDown = CBool(GetKeyState(VK_CONTROL) And &H8000)
End Function

'*******************************************************************************
'Convert between HiMetric and Points
'1) 1 hiMetric = 0.00001 meters (1E-5)
'2) 1 inch = 0.0254 meters
'3) 1 inch = 72 points (in computing)
'1)+2)+3) => 1 hiMetric = 1 / 100000 / 0.0254 * 72 = 0.0283464... points
'*******************************************************************************
Private Function HiMetricToPoints(ByVal hiMetric As Long) As Single
    HiMetricToPoints = hiMetric * 0.0283464
End Function
Private Function PointsToHiMeter(ByVal pts As Single) As Long
    PointsToHiMeter = CLng(pts / 0.0283464)
End Function

'*******************************************************************************
'Returns the String Caption of a Window identified by a handle
'*******************************************************************************
#If VBA7 Then
    Private Function GetWindowCaption(ByVal hwnd As LongPtr) As String
#Else
    Private Function GetWindowCaption(ByVal hwnd As Long) As String
#End If
    Dim bufferLength As Long: bufferLength = GetWindowTextLength(hwnd)
    GetWindowCaption = VBA.Space$(bufferLength)
    GetWindowText hwnd, GetWindowCaption, bufferLength + 1
End Function

'*******************************************************************************
'Checks if the ActiveWindow is a VBE Window
'*******************************************************************************
Private Function IsVBEActive() As Boolean
    #If Mac Then
    #Else
    IsVBEActive = VBA.InStr(1, GetWindowCaption(GetActiveWindow()) _
        , "Microsoft Visual Basic", vbTextCompare) <> 0
    #End If
End Function
