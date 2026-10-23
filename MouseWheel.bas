Attribute VB_Name = "MouseWheel"
Option Explicit

' Mouse-wheel scrolling for MSForms ListBox and ComboBox controls, which ignore the
' wheel on their own. Windows only; on Mac every call is a no-op.
'
'   WheelScroll ctl, frm    Call from each control's MouseMove. Makes ctl the scroll
'                           target and installs a low-level mouse hook if one isn't
'                           live. Pass Nothing (from the form's own MouseMove) to
'                           stop scrolling while the pointer is over blank space.
'   StopWheelScroll         Removes the hook. Call from the form's QueryClose.
'
' The hook only lives while the pointer is over the form (or an open combo dropdown,
' which is a separate popup window owned by the form). The first mouse event outside
' it removes the hook, so the sheet's own wheel scrolling is never swallowed and no
' hook is left pointing at VBA code while you work in the VBE.

#If Win64 Then
    Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal point As LongLong) As LongPtr
#Else
    Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long) As LongPtr
#End If
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
    ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hhk As LongPtr) As Long
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hhk As LongPtr, _
    ByVal nCode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetAncestor Lib "user32" (ByVal hwnd As LongPtr, ByVal gaFlags As Long) As LongPtr
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr

Private Const WH_MOUSE_LL As Long = 14
Private Const HC_ACTION As Long = 0
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const GA_ROOT As Long = 2
Private Const GW_OWNER As Long = 4
Private Const WHEEL_DELTA As Long = 120
Private Const LINES_PER_NOTCH As Long = 3

Private Type POINTAPI
    x As Long
    y As Long
End Type

#If Win64 Then
' WindowFromPoint takes POINT by value, which x64 passes as one 8-byte argument.
Private Type PackedPoint
    value As LongLong
End Type
#End If

Private Type MSLLHOOKSTRUCT
    pt As POINTAPI
    mouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As LongPtr
End Type

Private hookHandle As LongPtr
Private formHandle As LongPtr
Private target As Object

'=======================================================================================
' Entry
'=======================================================================================

Public Sub WheelScroll(ctl As Object, frm As Object)
#If Mac Then
    Exit Sub
#Else
    Set target = ctl
    If hookHandle <> 0 Or ctl Is Nothing Then Exit Sub

    formHandle = FindWindow("ThunderDFrame", frm.Caption)
    If formHandle = 0 Then Exit Sub
    hookHandle = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, Application.HinstancePtr, 0)
#End If
End Sub

Public Sub StopWheelScroll()
#If Mac Then
    Exit Sub
#Else
    If hookHandle <> 0 Then UnhookWindowsHookEx hookHandle
    hookHandle = 0
    formHandle = 0
    Set target = Nothing
#End If
End Sub

'=======================================================================================
' Hook
'=======================================================================================

' Runs for every mouse event system-wide while the hook is live, so it stays cheap and
' must never raise: an error escaping a hook procedure takes Excel down.
Private Function MouseProc(ByVal nCode As Long, ByVal wParam As LongPtr, _
                           ByRef info As MSLLHOOKSTRUCT) As LongPtr
    On Error Resume Next
    If nCode = HC_ACTION Then
        If Not PointerOnForm(info.pt) Then
            StopWheelScroll
        ElseIf wParam = WM_MOUSEWHEEL And Not target Is Nothing Then
            ScrollTarget info.mouseData \ &H10000
            MouseProc = 1   ' handled; don't pass the wheel on to Excel
            Exit Function
        End If
    End If
    MouseProc = CallNextHookEx(0, nCode, wParam, info)
End Function

' The wheel delta is the signed high word of mouseData: +120 per notch away from the
' user, -120 towards. Touchpads send smaller deltas, so any movement is at least a line.
Private Sub ScrollTarget(ByVal delta As Long)
    Dim lines As Long, newTop As Long

    lines = -(delta \ WHEEL_DELTA) * LINES_PER_NOTCH
    If lines = 0 Then lines = IIf(delta > 0, -1, 1)
    If target.ListCount = 0 Then Exit Sub

    newTop = target.TopIndex + lines
    If newTop < 0 Then newTop = 0
    If newTop > target.ListCount - 1 Then newTop = target.ListCount - 1
    target.TopIndex = newTop
End Sub

Private Function PointerOnForm(pt As POINTAPI) As Boolean
    Dim root As LongPtr

    root = GetAncestor(WindowAt(pt), GA_ROOT)
    PointerOnForm = (root = formHandle) Or (GetWindow(root, GW_OWNER) = formHandle)
End Function

Private Function WindowAt(pt As POINTAPI) As LongPtr
#If Win64 Then
    Dim packed As PackedPoint
    LSet packed = pt
    WindowAt = WindowFromPoint(packed.value)
#Else
    WindowAt = WindowFromPoint(pt.x, pt.y)
#End If
End Function
