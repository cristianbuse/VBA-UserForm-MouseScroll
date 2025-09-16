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
''    Works with both MODAL and MODELESS UserForms!
''    Simultaneoulsy tracks all forms that called the EnableMouseScroll method!
''    Hold down SHIFT key when scrolling the mouse wheel, for Horizontal Scroll!
''    Hold down CTRL key when scrolling the mouse wheel, for Zoom!
'' Notes:
''    - Installs a Mouse Hook by calling SetWindowsHookEx API with ID
''      WH_MOUSE = 7 and the address of the MouseProc callback function
''    - The Mouse Hook is active as long as there is at least one form that
''      previously enabled scrolling (i.e. called EnableMouseScroll method)
''      Relevant forms are tracked automatically by checking if the form's main
''      window is still enabled and if there are any references left pointing
'       to the form's object. When all the forms that called EnableMouseScroll
''      are destroyed then the mouse hook is removed automatically. No need to
''      call DisableMouseScroll method although you could do it in the form's
''      Terminate event if you wish to
''    - Another option would be to use ID WH_MOUSE_LL = 14 which would require a
''      LowLevelMouseProc callback but unlike the WH_MOUSE hook which is local
''      (hooked on the current thread only) the WH_MOUSE_LL hook is actually
''      global and very slow
''    - The system calls the MouseProc function whenever the Excel Application
''      calls the GetMessage or PeekMessage functions and there is a mouse
''      message to be processed
''    - This module also holds a collection of MouseOverControls that call back
''      the SetHoveredControl method in this module whenever a MouseMove event
''      is triggered
''    - You can debug code safely while hook is on
'' Requires:
''    - MouseOverControl: Container that tracks MouseMove events
''==============================================================================

Option Explicit
Option Private Module

#Const x64 = Win64
#Const x32 = (x64 = 0)

#If Mac Then 'Placeholders
    Public Function EnableMouseScroll(ByVal uForm As MSForms.UserForm _
                                    , Optional ByVal passScrollToParentAtMargins As Boolean = True _
                                    , Optional ByVal useShiftForPerpendicularScroll As Boolean = True _
                                    , Optional ByVal useCtrlToZoom As Boolean = True) As Boolean
    End Function
    Public Sub DisableMouseScroll(ByVal uForm As MSForms.UserForm): End Sub
    Public Sub SetHoveredControl(ByVal moCtrl As MouseOverControl): End Sub
    Public Sub ProcessMouseData(): End Sub
#Else

Private Type POINTAPI
    x As Long
    y As Long
End Type

'API declarations
#If VBA7 Then
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
    Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function IsChild Lib "user32" (ByVal hWndParent As LongPtr, ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function IsWindowEnabled Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function ShowWindowAsync Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    #If Win64 Then
        Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal Point As LongLong) As LongPtr
    #Else
        Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr
    #End If
#Else
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
    Private Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hwnd As Long) As Long
    Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function ShowWindowAsync Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
#End If

#If VBA7 = 0 Then
    Private Enum LongPtr: [_]: End Enum
    Private Enum LONG_PTR: [_]: End Enum
#End If

#If x64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize As Long = 8
#Else
    Private Const NullPtr As Long = 0&
    Private Const PtrSize As Long = 4
#End If

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

'Id of the hook procedure to be installed with SetWindowsHookExA for MouseProc
Private Const WH_MOUSE As Long = 7

'https://msdn.microsoft.com/en-us/library/windows/desktop/ms644988(v=vs.85).aspx
Private Type MOUSEHOOKSTRUCT
    pt As POINTAPI
    hwnd As LongPtr
    wHitTestCode As Long
    dwExtraInfo As LongPtr
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
Public Type MOUSEHOOKSTRUCTEX
    tagMOUSEHOOKSTRUCT As MOUSEHOOKSTRUCT
    mouseData As Long 'DWORD
End Type

'Necessary struct and constants to calculate the number of lines/pages to scroll
'https://msdn.microsoft.com/en-us/library/ms997498.aspx
Private Const WHEEL_DELTA As Long = 120
Private Const SPI_GETWHEELSCROLLLINES As Long = &H68

Private Type SCROLL_AMOUNT
    lines As Single
    pages As Single
End Type

'Window handles for all forms with scrolling enabled. Always instantiated
Private m_hWndAllForms As New Collection

'Collection of sub-collections of MouseOverControls (one for each form)
Private m_controls As New Collection

'Keeps track of options for each form
Private m_options As New Collection

'The last control that was hovered (could be the UserForm itself)
Private m_lastHoveredControl As MouseOverControl

'If VBE took focus then the form needs to retake focus via an async call
Private m_needsActivation As Boolean
Private m_needsHooking As Boolean

'The last ComboBox that was used
Private m_lastCombo As MSForms.ComboBox
Private m_isLastComboOn As Boolean

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

Private Enum SCROLL_OPTIONS
    soNone = 0
    soPassScrollToParentAtMargins = 1 'If the current hovered control cannot
                                      ' scroll anymore, then pass (or not) the
                                      ' scroll to the Parent Control/Form
    soUseShiftForPerpendicularScroll = 2
    soUseCtrlToZoom = 4
End Enum
Private m_lastSO As SCROLL_OPTIONS

Private m_hookOnFlag As Byte 'See HookMouseIfNeeded

'*******************************************************************************
'Enables mouse wheel scroll for the specified UserForm
'*******************************************************************************
Public Function EnableMouseScroll(ByVal uForm As MSForms.UserForm _
                                , Optional ByVal passScrollToParentAtMargins As Boolean = True _
                                , Optional ByVal useShiftForPerpendicularScroll As Boolean = True _
                                , Optional ByVal useCtrlToZoom As Boolean = True) As Boolean
    If uForm Is Nothing Then Exit Function
    HookMouseIfNeeded
    '
    AddForm uForm, passScrollToParentAtMargins _
                 , useShiftForPerpendicularScroll _
                 , useCtrlToZoom
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
Private Sub HookMouseIfNeeded()
#If x64 Then
    Const ctrlASM As Long = &H894C
    Const lateBindOffset As LongLong = 55
#Else
    Const ctrlASM As Long = &HD0FF
    Const lateBindOffset As Long = 22
#End If
    Static hHookAddr As LongPtr
    Static aPtr As LongPtr
    Static mProcObj As New MouseOverControl
    Static tID As Long
    Dim needsASM As Boolean
    Dim uHookAddr As LongPtr
    Dim flagAddr As LongPtr
    Dim oPtrAddr As LongPtr
    Dim mPtr As LongPtr: mPtr = VBA.Int(AddressOf MouseProcEntryASM)
    '
    needsASM = (aPtr = NullPtr)
    If Not needsASM Then needsASM = (MemLongPtr(aPtr + 10) And &HFFFF) <> ctrlASM
    '
    If needsASM Then
#If x64 Then
        aPtr = MemLongPtr(MemLongPtr(ObjPtr(mProcObj)) + PtrSize * 14) 'mProcObj.MouseProcASM
        MemLongPtr(aPtr) = &H8244C8948^       '48 89 4C 24 08       ;MOV QWORD PTR [RSP+08],RCX
        MemLongPtr(aPtr + 5) = &H1024548948^  '48 89 54 24 10       ;MOV QWORD PTR [RSP+10],RDX
        MemLongPtr(aPtr + 10) = &H182444894C^ '4C 89 44 24 18       ;MOV QWORD PTR [RSP+18],R8
        MemLongPtr(aPtr + 15) = &HB948        '48 B9                ;MOV RCX,0...
        hHookAddr = aPtr + 17
        MemLongPtr(aPtr + 25) = &HB848        '48 B8                ;MOV RAX,0...
        uHookAddr = aPtr + 27
        MemLongPtr(aPtr + 35) = &H55          '55                   ;PUSH RBP
        MemLongPtr(aPtr + 36) = &HEC8B48      '48 8B EC             ;MOV RBP,RSP
        MemLongPtr(aPtr + 39) = &H20EC8348    '48 83 EC 20          ;SUB RSP,0x20
        MemLongPtr(aPtr + 43) = &HD0FF        'FF D0                ;CALL RAX ;UnHookMouse
        MemLongPtr(aPtr + 45) = &H20C48348    '48 83 C4 20          ;ADD RSP,0x20
        MemLongPtr(aPtr + 49) = &H5D          '5D                   ;POP RBP
        MemLongPtr(aPtr + 50) = &HB848        '48 B8                ;MOV RAX,0...
        flagAddr = aPtr + 52
        MemLongPtr(aPtr + 60) = &H3880        '80 38 00             ;CMP BYTE PTR [RAX],0x0
        MemLongPtr(aPtr + 63) = &H2F74        '74 2F                ;JE ... ;RET
        MemLongPtr(aPtr + 65) = &HC6          'C6 00 00             ;MOV BYTE PTR [RAX],00
        MemLongPtr(aPtr + 68) = &HB948        '48 B9                ;MOV RCX,0...
        oPtrAddr = aPtr + 70
        MemLongPtr(aPtr + 78) = &H824548B48^  '48 8B 54 24 08       ;MOV RDX,QWORD PTR [RSP+08]
        MemLongPtr(aPtr + 83) = &H1024448B4C^ '4C 8B 44 24 10       ;MOV R8,QWORD PTR [RSP+10]
        MemLongPtr(aPtr + 88) = &H18244C8B4C^ '4C 8B 4C 24 18       ;MOV R9,QWORD PTR [RSP+18
        MemLongPtr(aPtr + 93) = &H18B48       '48 8B 01             ;MOV RAX,QWORD PTR [RCX] ;vtbl
        MemLongPtr(aPtr + 96) = &H55          '55                   ;PUSH RBP
        MemLongPtr(aPtr + 97) = &HEC8B48      '48 8B EC             ;MOV RBP,RSP
        MemLongPtr(aPtr + 100) = &H20EC8348   '48 83 EC 20          ;SUB RSP,0x20
        MemLongPtr(aPtr + 104) = &H6850FF     'FF 50 68             ;CALL QWORD PTR [RAX+68] ;MouseProc
        MemLongPtr(aPtr + 107) = &H20C48348   '48 83 C4 20          ;ADD RSP,0x20
        MemLongPtr(aPtr + 111) = &H5D         '5D                   ;POP RBP
        MemLongPtr(aPtr + 112) = &HC3         'C3                   ;RET
#Else
        aPtr = VBA.Int(AddressOf MouseProcASM1) 'Only enough for 33 bytes
        Dim aPtr2 As Long: aPtr2 = VBA.Int(AddressOf MouseProcASM2)
        '
        MemLongPtr(aPtr) = &H68               '68                   ;PUSH 0...
        hHookAddr = aPtr + 1
        MemLongPtr(aPtr + 5) = &HB8           'B8                   ;MOV EAX,0...
        uHookAddr = aPtr + 6
        MemLongPtr(aPtr + 10) = &HD0FF        'FF D0                ;CALL EAX ;UnHookMouse
        MemLongPtr(aPtr + 12) = &HB9          'B9                   ;MOV ECX,0...
        flagAddr = aPtr + 13
        MemLongPtr(aPtr + 17) = &H13980       '80 39 01             ;CMP BYTE PTR [ECX],01
        MemLongPtr(aPtr + 20) = &H374         '74 03                ;JE ... ;Skip RET
        MemLongPtr(aPtr + 22) = &HCC2         'C2 0C 00             ;RET 000C
        MemLongPtr(aPtr + 25) = &HB8          'B8                   ;MOV EAX,0...
        MemLongPtr(aPtr + 26) = aPtr2
        MemLongPtr(aPtr + 30) = &HE0FF        'FF E0                ;JMP EAX
        '
        MemLongPtr(aPtr2) = &H1C6             'C6 01 00             ;MOV BYTE PTR [ECX],00
        MemLongPtr(aPtr2 + 3) = &H2434FF      'FF 34 24             ;PUSH DWORD PTR [ESP] ;push return address
        MemLongPtr(aPtr2 + 6) = &HB9          'B9                   ;MOV ECX,0...
        oPtrAddr = aPtr2 + 7
        MemLongPtr(aPtr2 + 11) = &H4244C89    '89 4C 24 04          ;MOV DWORD PTR [ESP+04],ECX ;mProcObj
        MemLongPtr(aPtr2 + 15) = &H18B        '8B 01                ;MOV EAX,DWORD PTR [ECX]
        MemLongPtr(aPtr2 + 17) = &H34408B     '8B 40 34             ;MOV EAX,DWORD PTR [EAX+34]
        MemLongPtr(aPtr2 + 20) = &HE0FF       'FF E0                ;JMP EAX
#End If
        MemLongPtr(flagAddr) = VarPtr(m_hookOnFlag)
        MemLongPtr(oPtrAddr) = ObjPtr(mProcObj)
        MemLongPtr(uHookAddr) = GetProcAddress(GetModuleHandle("user32"), "UnhookWindowsHookEx")
        tID = GetCurrentThreadId()
    End If
    MemLongPtr(mPtr + lateBindOffset) = aPtr 'In case of recompilation
    If m_hookOnFlag = 1 Then Exit Sub
    m_hookOnFlag = 1
    MemLongPtr(hHookAddr) = SetWindowsHookEx(WH_MOUSE, mPtr, 0, tID)
End Sub
#If x64 Then
Private Sub MouseProcEntryASM(): End Sub 'Takes care of Break Mode by default
#Else
Private Sub MouseProcEntryASM(a1, a2, a3): End Sub 'Args for correct RET 0x0C
Private Sub MouseProcASM1(): End Sub
Private Sub MouseProcASM2(): End Sub
#End If

'*******************************************************************************
'UnHooks Mouse
'*******************************************************************************
Private Sub UnHookMouse()
    m_hookOnFlag = 0
    Set m_hWndAllForms = Nothing
    Set m_controls = Nothing
    Set m_options = Nothing
    Set m_lastHoveredControl = Nothing
    Set m_lastCombo = Nothing
    Debug.Print "Unhooked " & Now()
End Sub

'*******************************************************************************
'Adds the form handle to m_hWndAllForms collection
'Adds the passScrollAtMargins option to m_passScrollColl collection
'Adds a sub-collection of MouseMove controls to m_controls collection
'*******************************************************************************
Private Sub AddForm(ByVal uForm As MSForms.UserForm _
                  , ByVal passScrollAtMargins As Boolean _
                  , ByVal useShiftForPerpendicularScroll As Boolean _
                  , ByVal useCtrlToZoom As Boolean)
    Dim hWndForm As LongPtr
    Dim keyValue As String
    Dim so As SCROLL_OPTIONS
    '
    hWndForm = GetFormHandle(uForm)
    keyValue = CStr(hWndForm)
    '
    If CollectionHasKey(m_hWndAllForms, keyValue) Then
        m_controls.Remove keyValue
        m_options.Remove keyValue
    Else
        m_hWndAllForms.Add hWndForm, keyValue
    End If
    '
    If passScrollAtMargins Then so = so Or soPassScrollToParentAtMargins
    If useShiftForPerpendicularScroll Then so = so Or soUseShiftForPerpendicularScroll
    If useCtrlToZoom Then so = so Or soUseCtrlToZoom
    m_options.Add so, keyValue
    '
    Dim subControls As Collection
    Set subControls = New Collection
    m_controls.Add subControls, keyValue
    '
    Dim frmCtrl As MSForms.Control
    '
    For Each frmCtrl In uForm.Controls
        subControls.Add MouseOverControl.CreateFromControl(frmCtrl, hWndForm)
    Next frmCtrl
    subControls.Add MouseOverControl.CreateFromForm(uForm, hWndForm), keyValue
End Sub
Private Function MouseOverControl() As MouseOverControl
    Static moc As MouseOverControl
    If moc Is Nothing Then Set moc = New MouseOverControl
    Set MouseOverControl = moc
End Function

'*******************************************************************************
'Removes a form (by window handle) from the internal collections
'*******************************************************************************
Private Sub RemoveForm(ByVal hWndForm As LongPtr)
    If CollectionHasKey(m_hWndAllForms, hWndForm) Then
        Dim keyValue As String: keyValue = CStr(hWndForm)
        m_hWndAllForms.Remove keyValue
        m_controls.Remove keyValue
        m_options.Remove keyValue
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
        If CBool(IsWindow(v)) Then
            Dim s As String:      s = CStr(v)
            Dim iUnk As IUnknown: Set iUnk = m_controls(s)(s).GetControl
            Dim ptr As LongPtr:   ptr = ObjPtr(iUnk)
            Dim refCount As Long
            '
            Set iUnk = Nothing
            '
            refCount = CLng(MemLongPtr(ptr + PtrSize) And &H7FFFFFFF)
            If refCount = 2 Then RemoveForm v
        Else
            RemoveForm v
        End If
    Next v
End Sub

'*******************************************************************************
'Returns a boolean indicating if a Collection has a specific key
'Parameters:
'   - coll: a collection to check for key
'   - keyValue: the key being searched for
'Does not raise errors
'*******************************************************************************
Private Function CollectionHasKey(ByVal coll As Collection _
                                , ByVal keyValue As String) As Boolean
    On Error Resume Next
    coll.Item keyValue
    CollectionHasKey = (Err.Number = 0)
    On Error GoTo 0
End Function

'*******************************************************************************
'Called by MouseMove capable controls (MouseOverControl) stored in m_controls
'*******************************************************************************
Public Sub SetHoveredControl(ByVal moCtrl As MouseOverControl)
    Set m_lastHoveredControl = moCtrl
    On Error Resume Next
    m_lastSO = m_options(CStr(moCtrl.FormHandle))
    On Error GoTo 0
    If m_needsActivation Then
        Const SW_SHOW As Long = 5
        ShowWindowAsync moCtrl.FormHandle, SW_SHOW
        m_needsActivation = False
        m_needsHooking = True
    End If
    If m_needsHooking Then
        HookMouseIfNeeded
        m_needsHooking = False
    End If
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
Public Sub ProcessMouseData(ByVal ncode As Long _
                          , ByVal wParam As Long _
                          , ByRef lParam As MOUSEHOOKSTRUCTEX)
    RemoveDestroyedForms
    If m_hWndAllForms.Count = 0 Then
        UnHookMouse
        Exit Sub
    End If
    '
    If m_lastHoveredControl Is Nothing Then GoTo DelayHookAsync
    Dim fHWnd As LongPtr: fHWnd = m_lastHoveredControl.FormHandle
    '
    If Not CBool(IsWindowEnabled(fHWnd)) Then GoTo DelayHookAsync
    UpdateLastCombo
    If Not m_isLastComboOn Then
        Dim pHWnd As LongPtr: pHWnd = GetWindowUnderCursor()
        Dim className As String: className = Space$(&HFF)
        '
        If IsChild(fHWnd, pHWnd) = 0 Then GoTo DelayHookAsync
        className = Left$(className, GetClassName(pHWnd, className, Len(className)))
        If Not (className Like "F3 Server*") Then GoTo DelayHookAsync
    End If
    '
    If wParam = WM_MOUSEWHEEL Or wParam = WM_MOUSEHWHEEL Then
        Dim scrollAmount As SCROLL_AMOUNT
        Dim scrollAction As SCROLL_ACTION
        '
        scrollAmount = GetScrollAmount(GetWheelDelta(lParam.mouseData))
        scrollAction = GetScrollAction(yWheel:=(wParam = WM_MOUSEWHEEL))
        '
        If m_isLastComboOn Then
            m_lastSO = m_lastSO And Not soPassScrollToParentAtMargins
            Call ScrollY(m_lastCombo, scrollAmount)
        Else
            Select Case scrollAction
            Case saScrollY
                Call ScrollY(m_lastHoveredControl.GetControl, scrollAmount)
            Case saScrollX
                If m_isLastComboOn Then GoTo DelayHookAsync
                Call ScrollX(m_lastHoveredControl.GetControl, scrollAmount)
            Case saZoom
                If m_isLastComboOn Then GoTo DelayHookAsync
                Call Zoom(m_lastHoveredControl.GetControl, scrollAmount)
            End Select
        End If
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
        'Mouse drag by scroll wheel example:
        Static lastX As Single
        Static lastY As Single
        Const sLines As Long = 3 'Constant number of lines to scroll - change as needed
        Const VK_MBUTTON As Long = &H4
        '
        If wParam = WM_MBUTTONDOWN Then
            lastX = lParam.tagMOUSEHOOKSTRUCT.pt.x
            lastY = lParam.tagMOUSEHOOKSTRUCT.pt.y
        End If
        '
        If GetKeyState(VK_MBUTTON) And &H8000 Then
            If IsShiftKeyDown() Then
                scrollAmount.lines = sLines * Sgn(lastX - lParam.tagMOUSEHOOKSTRUCT.pt.x)
                If m_isLastComboOn Then GoTo DelayHookAsync
                Call ScrollX(m_lastHoveredControl.GetControl, scrollAmount)
            Else
                scrollAmount.lines = sLines * Sgn(lastY - lParam.tagMOUSEHOOKSTRUCT.pt.y)
                Call ScrollY(m_lastHoveredControl.GetControl, scrollAmount)
            End If
            lastX = lParam.tagMOUSEHOOKSTRUCT.pt.x
            lastY = lParam.tagMOUSEHOOKSTRUCT.pt.y
        End If
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
        End If
    End If
    '
    DoEvents
    'Make sure VBE is not activated as this would make the forms lose focus
    Const VBELabel As String = "Microsoft Visual Basic for Applications*"
    Dim foreHWnd As LongPtr: foreHWnd = GetForegroundWindow()
    If foreHWnd <> fHWnd Then
        m_needsActivation = GetWindowCaption(foreHWnd) Like VBELabel
    End If
    If Not m_needsActivation Then HookMouseIfNeeded
Exit Sub
DelayHookAsync:
    m_needsHooking = True
End Sub

'*******************************************************************************
'Returns the String Caption of a Window identified by a handle
'*******************************************************************************
Private Function GetWindowCaption(ByVal hwnd As LongPtr) As String
    Dim bufferLength As Long: bufferLength = GetWindowTextLength(hwnd)
    GetWindowCaption = Space$(bufferLength)
    GetWindowText hwnd, GetWindowCaption, bufferLength + 1
End Function

'*******************************************************************************
'Get the type of scroll action by reading Shift and Control key states
'*******************************************************************************
Private Function GetScrollAction(ByVal yWheel As Boolean) As SCROLL_ACTION
    If yWheel Then
        If IsShiftKeyDown() And CBool(m_lastSO And soUseShiftForPerpendicularScroll) Then
            GetScrollAction = saScrollX
        ElseIf IsControlKeyDown() And CBool(m_lastSO And soUseCtrlToZoom) Then
            GetScrollAction = saZoom
        Else
            GetScrollAction = saScrollY
        End If
    Else
        If IsShiftKeyDown() And CBool(m_lastSO And soUseShiftForPerpendicularScroll) Then
            GetScrollAction = saScrollY
        ElseIf IsControlKeyDown() And CBool(m_lastSO And soUseCtrlToZoom) Then
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
    SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, result, 0
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
            If m_lastSO And soPassScrollToParentAtMargins Then
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
                      , ByVal ctrlType As CONTROL_TYPE)
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
    If m_lastSO And soPassScrollToParentAtMargins Then
        If ctrl.TopIndex = lastTopIndex Then
            Call ScrollY(ctrl.Parent, scrollAmount)
        End If
    End If
End Sub

'*******************************************************************************
'Vertically scroll a TextBox control
'*******************************************************************************
Private Sub TBoxScrollY(ByVal tbox As MSForms.TextBox _
                      , ByRef scrollAmount As SCROLL_AMOUNT)
    If Not tbox.MultiLine Then
        ScrollY tbox.Parent, scrollAmount
        Exit Sub
    End If
    tbox.SetFocus
    '
    'Store current state
    Dim selectionStart As Long:  selectionStart = tbox.SelStart
    Dim selectionLength As Long: selectionLength = tbox.SelLength
    Dim startY As Long:          startY = tbox.CurY
    Dim startLine As Long:       startLine = tbox.CurLine
    '
    'Determine line characteristics
    With tbox
        .CurLine = 0
        .CurY = 0
        Dim minY As Long:  minY = .CurY
        Dim currY As Long: currY = minY
        Dim lastY As Long
        Dim i As Long
        '
        For i = 1 To .LineCount - 1
            lastY = currY
            .CurLine = i
            currY = .CurY
            If currY = lastY Then Exit For
        Next i
        Dim linesPerPage As Long: linesPerPage = i - 1
        '
        If (linesPerPage = 0) Or (linesPerPage = .LineCount - 1) Then
            tbox.SelStart = selectionStart
            tbox.SelLength = selectionLength
            ScrollY tbox.Parent, scrollAmount
            Exit Sub
        End If
        '
        .CurLine = .LineCount - 1
        Dim lastSelStart As Long: lastSelStart = .SelStart
        .CurLine = 0
        .Visible = False
        .SelStart = lastSelStart
        .SelLength = 0
        .Visible = True
        .SetFocus
        '
        Dim bottomY As Long: bottomY = .CurY
        Dim hmPerLine As Single
        Dim topAdjust As Long
        '
        .CurLine = .LineCount - 1
        .Visible = False
        .SelStart = 0
        .SelLength = 0
        .Visible = True
        .SetFocus
        '
        If bottomY > minY Then
            hmPerLine = (bottomY - minY) / (.LineCount - 1)
        Else
            hmPerLine = (minY - .CurY) / (.LineCount - linesPerPage - 1)
            minY = VBA.Int(bottomY - hmPerLine * (.LineCount - 1))
        End If
        If hmPerLine = 0 Then Exit Sub
        '
        topAdjust = .CurY - minY + (.LineCount - linesPerPage - 1) * hmPerLine
        If Abs(topAdjust) = 1 Then topAdjust = 0 'Rounding error
    End With
    If startY > tbox.LineCount * hmPerLine Then startY = startY - topAdjust
    '
    'Lines to scroll up/down
    Dim deltaLines As Long
    If scrollAmount.lines <> 0 Then
        deltaLines = -scrollAmount.lines
    Else
        deltaLines = -scrollAmount.pages * VBA.Int(linesPerPage)
    End If
    '
    'Adjust for 1 line scroll here
    'deltaLines = Sgn(deltaLines)
    '
    Dim topLine As Long: topLine = startLine - (startY - minY) / hmPerLine
    Dim newline As Long: newline = topLine + deltaLines
    '
    'Clamp the new scroll line
    If newline < 0 Then
        newline = 0
    ElseIf newline >= tbox.LineCount Then
        newline = tbox.LineCount - 1
    End If
    tbox.CurLine = newline
    '
    'Restore Selection. Must hide (or disable) textBox first, to lock scroll
    tbox.Visible = False
    tbox.SelStart = selectionStart
    tbox.SelLength = selectionLength
    tbox.Visible = True
    If Abs(startLine - newline - linesPerPage) < 2 Then GetParent(tbox).Repaint
    tbox.SetFocus
    '
    If m_lastSO And soPassScrollToParentAtMargins Then
        currY = tbox.CurY
        If currY > tbox.LineCount * hmPerLine Then currY = currY - topAdjust
        If Abs(currY - startY) < 2 Then ScrollY tbox.Parent, scrollAmount
    End If
End Sub
Private Function GetParent(ByVal tbox As MSForms.TextBox) As Object
    Dim p As Object: Set p = tbox.Parent
    Dim o As Object
    '
    On Error Resume Next
    Do
        Set o = Nothing
        Set o = p.Parent
        If Not o Is Nothing Then Set p = o
    Loop Until o Is Nothing
    On Error GoTo 0
    Set GetParent = p
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
            If m_lastSO And soPassScrollToParentAtMargins Then
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
Private Sub ListScrollX(ByVal lbox As MSForms.Control _
                      , ByRef scrollAmount As SCROLL_AMOUNT)
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
            If m_lastSO And soPassScrollToParentAtMargins Then
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
Private Function GetFormHandle(ByVal objForm As MSForms.UserForm) As LongPtr
    IUnknown_GetWindow objForm, VarPtr(GetFormHandle)
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
'Returns the handle for the window currently under cursor
'*******************************************************************************
Private Function GetWindowUnderCursor() As LongPtr
    Dim pt As POINTAPI: GetCursorPos pt
    #If Win64 Then
        GetWindowUnderCursor = WindowFromPoint(MemLongPtr(VarPtr(pt)))
    #Else
        GetWindowUnderCursor = WindowFromPoint(pt.x, pt.y)
    #End If
End Function

Private Property Get MemLongPtr(ByVal addr As LongPtr) As LongPtr
    Static pa(0 To 0) As PointerAccessor
    With pa(0)
        If .sa.cDims = 0 Then
            InitSafeArray .sa
            WritePtrNatively pa, VarPtr(.sa)
        End If
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        MemLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property
Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Static pa(0 To 0) As PointerAccessor
    With pa(0)
        If .sa.cDims = 0 Then
            InitSafeArray .sa
            WritePtrNatively pa, VarPtr(.sa)
        End If
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property
Private Sub InitSafeArray(ByRef sa As SAFEARRAY_1D)
    Const FADF_AUTO = &H1
    Const FADF_FIXEDSIZE = &H10
    sa.cDims = 1
    sa.cLocks = 1
    sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
End Sub
'https://github.com/WNKLER/RefTypes/discussions/3#discussion-8595790
Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal ptr As LongPtr)
    ptrs(0) = ptr
End Sub

#End If 'End of #If Mac
