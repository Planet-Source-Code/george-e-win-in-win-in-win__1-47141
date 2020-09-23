Attribute VB_Name = "modstuff"
Option Explicit
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ScrollWindowByNum& Lib "user32" Alias "ScrollWindow" (ByVal hwnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, ByVal lpRect As Long, ByVal lpClipRect As Long)
Private Declare Function GetWindowRect& Lib "user32" (ByVal hwnd As Long, lpRect As RECT)
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex As Long)
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const SWP_NOSIZE As Integer = &H1
Private Const SWP_NOZORDER  As Integer = &H4
Private Const SWP_NOMOVE  As Integer = &H2
Private Const SWP_DRAWFRAME  As Integer = &H20
Private Const GWL_STYLE  As Integer = (-16)
Private Const WS_THICKFRAME = &H40000
Private Const WS_DLGFRAME = &H400000
Private Const WS_POPUP = &H80000000
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_MAXIMIZE = &H1000000
Private Const GWL_WNDPROC  As Integer = (-4)
Private Const WS_VSCROLL = &H200000
Private Const WS_HSCROLL = &H100000
Private Const SWP_FRAMECHANGED  As Integer = &H20
Private Const SB_HORZ  As Integer = 0
Private Const SB_VERT  As Integer = 1
Private Const SB_BOTH  As Integer = 3
Private Const SB_LINEDOWN  As Integer = 1
Private Const SB_LINEUP  As Integer = 0
Private Const SB_PAGEDOWN  As Integer = 3
Private Const SB_PAGEUP  As Integer = 2
Private Const SB_THUMBTRACK  As Integer = 5
Private Const SB_ENDSCROLL  As Integer = 8
Private Const WM_HSCROLL  As Integer = &H114
Private Const WM_VSCROLL  As Integer = &H115
Private Const WM_DESTROY  As Integer = &H2
Private Const SIF_ALL  As Integer = &H17
Private Const SIF_DISABLENOSCROLL  As Integer = &H8
Private Const SM_CXVSCROLL  As Integer = 2
Private Const SM_CYHSCROLL  As Integer = 3
Public Const GWL_EXSTYLE  As Integer = (-20)
Public Const WS_BORDER = &H800000
Public Const WS_CHILD = &H40000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_GROUP = &H20000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_TABSTOP = &H10000
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_VISIBLE = &H10000000
Public Const WM_CLOSE  As Integer = &H10
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_STATICEDGE = &H20000
Public Const SWP_NOACTIVATE  As Integer = &H10
Public Const SW_HIDE  As Integer = 0
Public Const SW_MAXIMIZE  As Integer = 3
Public Const SW_MINIMIZE  As Integer = 6
Public Const SW_NORMAL  As Integer = 1
Public Const SW_RESTORE  As Integer = 9
Public Const SW_SHOW  As Integer = 5
Public Const HWND_TOPMOST  As Integer = -1
Public Const HWND_NOTOPMOST  As Integer = -2
Public Const GW_HWNDNEXT  As Integer = 2

Private s As SCROLLINFO
Private OriginHeight As Long
Private OriginWidth As Long
Public old_parent As Long
Public child_hwnd As Long

Public Sub ControlSize(ControlName As Control, SetTrue As Boolean)

  '* Sizable BorderStyle property for a Control *

  Dim dwStyle As Long

    dwStyle = GetWindowLong(ControlName.hwnd, GWL_STYLE)
    If SetTrue Then
        dwStyle = dwStyle Or WS_THICKFRAME
      Else 'SETTRUE = FALSE/0
        dwStyle = dwStyle - WS_THICKFRAME
    End If
    dwStyle = SetWindowLong(ControlName.hwnd, GWL_STYLE, dwStyle)
    SetWindowPos ControlName.hwnd, ControlName.Parent.hwnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

End Sub

Public Sub ControlDialog(ControlName As Control, SetTrue As Boolean)

  '* Fixed Dialog BorderStyle property for a Control *

  Dim dwStyle As Long

    dwStyle = GetWindowLong(ControlName.hwnd, GWL_STYLE)
    If SetTrue Then
        dwStyle = dwStyle Or WS_DLGFRAME
      Else 'SETTRUE = FALSE/0
        dwStyle = dwStyle - WS_DLGFRAME
    End If
    dwStyle = SetWindowLong(ControlName.hwnd, GWL_STYLE, dwStyle)
    SetWindowPos ControlName.hwnd, ControlName.Parent.hwnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

End Sub

Public Sub ControlModal(ControlName As Control, SetTrue As Boolean)

  '* Make a Control in Modal mode like "form.show 1" *
  '* so you can't switch to another control or unload the
  '* form until all control's modal are set to false

  Dim dwStyle As Long

    dwStyle = GetWindowLong(ControlName.hwnd, GWL_STYLE)
    If SetTrue Then
        dwStyle = dwStyle Or WS_POPUP
      Else 'SETTRUE = FALSE/0
        dwStyle = dwStyle - WS_POPUP
    End If
    dwStyle = SetWindowLong(ControlName.hwnd, GWL_STYLE, dwStyle)
    SetWindowPos ControlName.hwnd, ControlName.Parent.hwnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

End Sub

Public Sub ControlCaption(ControlName As Control, SetTrue As Boolean)

  '* Add a Title-bar like form to a Control *
  '* (from FreeVBCode.com (Great to combine with controlsize)

  Dim dwStyle As Long

    dwStyle = GetWindowLong(ControlName.hwnd, GWL_STYLE)
    If SetTrue Then
        dwStyle = dwStyle Or WS_CAPTION
      Else 'SETTRUE = FALSE/0
        dwStyle = dwStyle - WS_CAPTION
    End If
    dwStyle = SetWindowLong(ControlName.hwnd, GWL_STYLE, dwStyle)
    SetWindowPos ControlName.hwnd, ControlName.Parent.hwnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

End Sub

Public Sub ControlSysMenu(ControlName As Control, SetTrue As Boolean)

  '* Enable the System-Menu for Control with Title-bar *
  '* Call ControlCaption First

  Dim dwStyle As Long

    dwStyle = GetWindowLong(ControlName.hwnd, GWL_STYLE)
    If SetTrue Then
        dwStyle = dwStyle Or WS_SYSMENU
      Else 'SETTRUE = FALSE/0
        dwStyle = dwStyle - WS_SYSMENU
    End If
    dwStyle = SetWindowLong(ControlName.hwnd, GWL_STYLE, dwStyle)
    SetWindowPos ControlName.hwnd, ControlName.Parent.hwnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

End Sub

Public Sub ControlMinBox(ControlName As Control, SetTrue As Boolean)

  '* Enable the Minimize-Box for Control with Title-bar *

  Dim dwStyle As Long

    dwStyle = GetWindowLong(ControlName.hwnd, GWL_STYLE)
    If SetTrue Then
        dwStyle = dwStyle Or WS_MINIMIZEBOX
      Else 'SETTRUE = FALSE/0
        dwStyle = dwStyle - WS_MINIMIZEBOX
    End If
    dwStyle = SetWindowLong(ControlName.hwnd, GWL_STYLE, dwStyle)
    SetWindowPos ControlName.hwnd, ControlName.Parent.hwnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

End Sub

Public Sub ControlMaxBox(ControlName As Control, SetTrue As Boolean)

  '* Enable the Maximize-Box for Control with Title-bar *

  Dim dwStyle As Long

    dwStyle = GetWindowLong(ControlName.hwnd, GWL_STYLE)
    If SetTrue Then
        dwStyle = dwStyle Or WS_MAXIMIZEBOX
      Else 'SETTRUE = FALSE/0
        dwStyle = dwStyle - WS_MAXIMIZEBOX
    End If
    dwStyle = SetWindowLong(ControlName.hwnd, GWL_STYLE, dwStyle)
    SetWindowPos ControlName.hwnd, ControlName.Parent.hwnd, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME

End Sub

Public Function InstanceToWnd(ByVal target_pid As Long) As Long

  Dim test_hwnd As Long
  Dim test_pid As Long
  Dim test_thread_id As Long

  ' Get the first window handle.

    test_hwnd = FindWindow(ByVal 0&, ByVal 0&)

    ' Loop until we find the target or we run out
    ' of windows.
    Do While test_hwnd <> 0
        ' See if this window has a parent. If not,
        ' it is a top-level window.
        If GetParent(test_hwnd) = 0 Then
            ' This is a top-level window. See if
            ' it has the target instance handle.
            test_thread_id = GetWindowThreadProcessId(test_hwnd, test_pid)

            If test_pid = target_pid Then
                ' This is the target.
                InstanceToWnd = test_hwnd
                Exit Do '>---> Loop
            End If
        End If

        ' Examine the next window.
        test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    Loop

End Function

Public Sub SetScrollBar(hObj As Long, sbPos As ScrollBarConstants, Optional bShowAlways As Boolean = False)

  Dim lStyle As Long, rc As RECT, OldProc As Long

    lStyle = sbPos * &H100000
    SetWindowLong hObj, GWL_STYLE, GetWindowLong(hObj, GWL_STYLE) Or lStyle
    SetWindowPos hObj, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
    Call GetWindowRect(hObj, rc)
    OriginHeight = rc.bottom - rc.top + GetSystemMetrics(SM_CYHSCROLL) * (sbPos And vbHorizontal)
    OriginWidth = rc.right - rc.left + GetSystemMetrics(SM_CXVSCROLL) * (sbPos And vbVertical) / 2
    s.cbSize = Len(s)
    s.fMask = SIF_ALL
    If bShowAlways Then
        s.fMask = s.fMask Or SIF_DISABLENOSCROLL
    End If
    s.nMin = 0
    s.nPos = 0
    OldProc = SetWindowLong(hObj, GWL_WNDPROC, AddressOf WndProc)
    SetProp hObj, "OLDPROC", OldProc
    SetProp hObj, "SB_POS", sbPos
    SetProp hObj, "ORIGIN_WIDTH", OriginWidth
    SetProp hObj, "ORIGIN_HEIGHT", OriginHeight

End Sub

Public Sub AdjustScrollInfo(hObj As Long)

  Dim sb As Long, rc As RECT

    sb = GetProp(hObj, "SB_POS")
    Call GetWindowRect(hObj, rc)
    If (sb And vbVertical) = vbVertical Then
        Call GetScrollInfo(hObj, SB_VERT, s)
        s.nMax = GetProp(hObj, "ORIGIN_HEIGHT")
        s.nPage = rc.bottom - rc.top + 1
        If s.nPage > s.nMax - s.nPos + 1 Then
            s.nPage = s.nMax - s.nPos + 1
        End If
        Call SetScrollInfo(hObj, SB_VERT, s, True)
    End If
    If (sb And vbHorizontal) = vbHorizontal Then
        Call GetScrollInfo(hObj, SB_HORZ, s)
        s.nMax = GetProp(hObj, "ORIGIN_WIDTH")
        s.nPage = rc.right - rc.left + 1
        If s.nPage > s.nMax - s.nPos + 1 Then
            s.nPage = s.nMax - s.nPos + 1
        End If
        Call SetScrollInfo(hObj, SB_HORZ, s, True)
    End If

End Sub

Public Function WndProc(ByVal hOwner As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Dim nOldPos As Long, n As Long

    Select Case wMsg
      Case WM_VSCROLL, WM_HSCROLL
        GetScrollInfo hOwner, wMsg - WM_HSCROLL, s
        nOldPos = s.nPos
        Select Case GetLoWord(wParam)
          Case SB_LINEDOWN
            s.nPos = s.nPos + s.nPage \ 10
          Case SB_LINEUP
            s.nPos = s.nPos - s.nPage \ 10
          Case SB_PAGEDOWN
            s.nPos = s.nPos + s.nPage
          Case SB_PAGEUP
            s.nPos = s.nPos - s.nPage
          Case SB_THUMBTRACK
            s.nPos = GetHiWord(wParam)
          Case SB_ENDSCROLL
            If s.nPos = 0 Then
                AdjustScrollInfo hOwner
                Exit Function '>---> Bottom
            End If
        End Select
        SetScrollInfo hOwner, wMsg - WM_HSCROLL, s, True
        GetScrollInfo hOwner, wMsg - WM_HSCROLL, s
        If wMsg = WM_VSCROLL Then
            ScrollWindowByNum hOwner, 0, nOldPos - s.nPos, 0, 0
          Else 'NOT WMSG...
            ScrollWindowByNum hOwner, nOldPos - s.nPos, 0, 0, 0
        End If
      Case WM_DESTROY
        RemoveProp hOwner, "SB_POS"
        RemoveProp hOwner, "ORIGIN_WIDTH"
        RemoveProp hOwner, "ORIGIN_HEIGHT"
        Call SetWindowLong(hOwner, GWL_WNDPROC, GetProp(hOwner, "OLDPROC"))
      Case Else
    End Select
    WndProc = CallWindowProc(GetProp(hOwner, "OLDPROC"), hOwner, wMsg, wParam, lParam)

End Function

Private Function GetHiWord(dw As Long) As Long

    If dw And &H80000000 Then
        GetHiWord = (dw \ 65535) - 1
      Else 'NOT DW...
        GetHiWord = dw \ 65535
    End If

End Function

Private Function GetLoWord(dw As Long) As Long

    If dw And &H8000& Then
        GetLoWord = &H8000 Or (dw And &H7FFF&)
      Else 'NOT DW...
        GetLoWord = dw And &HFFFF&
    End If

End Function

':) Ulli's VB Code Formatter V2.16.6 (2005-Jul-24 04:27 AM) 108 + 307 = 415 Lines
