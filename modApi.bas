Attribute VB_Name = "modApi"
Option Explicit
' Settings
Public number_of_screens As Long
Public show_on_screen As Long
Public one_screen_maximise As Long
' ----

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal lWinIni As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "Shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal hwnd As Long, ByVal lptpm As Any) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetCBTSHLHook& Lib "dscbtshl" (ByVal Hook&, ByVal AdrCBT&, ByVal AdrSHL&)
Public Declare Function GetFileName& Lib "dscbtshl" (ByVal hwnd&, ByVal FileName$)
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function AdjustWindowRectEx Lib "user32" (lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Const GWL_STYLE = (-16)     ' Get Window's Style
Const GWL_EXSTYLE = (-20)   ' Get Window's Extended Style
Public Const HCBT_MOVESIZE = 0
Public Const HCBT_MINMAX = 1
Public Const HCBT_CREATEWND = 3
Public Const HCBT_DESTROYWND = 4
Public Const HCBT_ACTIVATE = 5
Public Const HCBT_SYSCOMMAND = 8
Public Const HCBT_SETFOCUS = 9
Public Const SM_CXICON = 11
Public Const SM_CYICON = 12
Public Const SM_CXSMICON = 49
Public Const SM_CYSMICON = 50
Public Const LR_DEFAULTCOLOR = &H0
Public Const LR_MONOCHROME = &H1
Public Const LR_COLOR = &H2
Public Const LR_COPYRETURNORG = &H4
Public Const LR_COPYDELETEORG = &H8
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_DEFAULTSIZE = &H40
Public Const LR_VGACOLOR = &H80
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_CREATEDIBSECTION = &H2000
Public Const LR_COPYFROMRESOURCE = &H4000
Public Const LR_SHARED = &H8000&
Public Const IMAGE_ICON = 1
Public Const WM_SETICON = &H80
Public Const ICON_SMALL = 0
Public Const ICON_BIG = 1
Public Const GW_OWNER = 4
Public Const MF_CHECKED = &H8&
Public Const MF_APPEND = &H100&
Public Const TPM_LEFTALIGN = &H0&
Public Const TPM_RETURNCMD = &H100&
Public Const TPM_RIGHTBUTTON = &H2&
Public Const MF_DISABLED = &H2&
Public Const MF_GRAYED = &H1&
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const SM_CXSCREEN = 0        ' Width of screen
Public Const SM_CYSCREEN = 1        ' Height of screen
Public Const SM_CXFULLSCREEN = 16   ' Width of window client area
Public Const SM_CYFULLSCREEN = 17   ' Height of window client area
Public Const SM_CXVIRTUALSCREEN = 78
Public Const SM_CYVIRTUALSCREEN = 79
Public Const SM_CMONITORS = 80
Public Const SM_SAMEDISPLAYFORMAT = 81
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Const SWP_NOSIZE = 1
Public Const SWP_NOMOVE = 2
Public Const SWP_NOZORDER = 4
Public Const SWP_NOACTIVATE = 10
Public Const SW_NORMAL = 1
Public Const SW_MAXIMIZE = 3
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_THICKFRAME = &H40000
Public Const WS_CAPTION = &HC00000                  ' WS_BORDER Or WS_DLGFRAME
Public Const WS_EX_CLIENTEDGE = &H200
'Styles
Const WS_CHILD = &H40000000
Const WS_CHILDWINDOW = WS_CHILD
Const WS_CLIPCHILDREN = &H2000000
Const WS_CLIPSIBLINGS = &H4000000
Const WS_DISABLED = &H8000000
Const WS_GROUP = &H20000
Const WS_MINIMIZE = &H20000000
Const WS_ICONIC = WS_MINIMIZE
Const WS_MAXIMIZE = &H1000000
Const WS_MAXIMIZEBOX = &H10000
Const WS_MINIMIZEBOX = &H20000
Const WS_POPUP = &H80000000
Const WS_SYSMENU = &H80000
Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Const WS_SIZEBOX = WS_THICKFRAME
Const WS_TABSTOP = &H10000
Const WS_VISIBLE = &H10000000

'Extended Styles
Const WS_EX_ACCEPTFILES = &H10&
Const WS_EX_APPWINDOW = &H40000
Const WS_EX_COMPOSITED = &H2000000
Const WS_EX_CONTEXTHELP = &H400
Const WS_EX_CONTROLPARENT = &H10000
Const WS_EX_DLGMODALFRAME = &H1
Const WS_EX_LAYERED = &H80000
Const WS_EX_LAYOUTRTL = &H400000
Const WS_EX_LEFT = &H0
Const WS_EX_MDICHILD = &H40
Const WS_EX_NOACTIVATE = &H8000000
Const WS_EX_NOINHERITLAYOUT = &H100000
Const WS_EX_NOPARENTNOTIFY = &H4
Const WS_EX_TOPMOST = &H8
Const WS_EX_TOOLWINDOW = &H80
Const WS_EX_WINDOWEDGE = &H100
Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
Const WS_EX_RIGHT = &H1000
Const WS_EX_RIGHTSCROLLBAR = &H0
Const WS_EX_RTLREADING = &H2000
Const WS_EX_STATICEDGE = &H20000
Const WS_EX_TRANSPARENT = &H20
Const WS_EX_LEFTSCROLLBAR = &H4000
Const WS_EX_LTRREADING = &H0

Const SPI_GETWORKAREA = 48
Public Function cbtCallback(ByVal hwnd As Long, ByVal ncode As Long) As Long

    Dim height As Long
    Dim width As Long
    Dim rtn As Long
    Dim WinEst As WINDOWPLACEMENT
    Dim R As RECT
    Dim lpClassName As String
    Dim lngCurStyle As Long
    WinEst.Length = Len(WinEst)
    
    If ncode = HCBT_MINMAX And one_screen_maximise = 1 Then
        rtn = GetWindowPlacement(hwnd, WinEst)
        If WinEst.showCmd = SW_MAXIMIZE Then
            lngCurStyle = GetWindowLong(hwnd, GWL_STYLE)
            If Not (lngCurStyle Or WS_CHILD) = lngCurStyle Then
                Call SystemParametersInfo(SPI_GETWORKAREA, vbNull, R, 0)
                R.Right = R.Right / 2
                If WinEst.rcNormalPosition.Left > R.Right Then
                    R.Left = R.Right
                    R.Right = R.Right * 2
                End If
                MoveWindow hwnd, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, True
            End If
        End If
'    ElseIf ncode = HCBT_CREATEWND Then
'        If Not (lngCurStyle Or WS_DLGFRAME) = lngCurStyle Then
'            rtn = GetWindowPlacement(hwnd, WinEst)
'            Debug.Print "WS_DLGFRAME"
'            Debug.Print WinEst.rcNormalPosition.Bottom
'            Debug.Print WinEst.rcNormalPosition.Top
'            Debug.Print WinEst.rcNormalPosition.Left
'            Debug.Print WinEst.rcNormalPosition.Right
'        End If
'
'        lpClassName = Space(256)
'        rtn = GetClassName(hwnd, lpClassName, 256)
'        If Trim(lpClassName) = "TNASTYNAGSCREEN" & vbNullChar Then
'            rtn = GetWindowPlacement(hwnd, WinEst)
'
'        End If
'        Debug.Print lpClassName
'    Else
'        showStyle hwnd
'        Debug.Print "cbt=" & ncode, hwnd
    End If

End Function
Public Function shlCallback(ByVal hwnd As Long, ByVal ncode As Long) As Long

    Debug.Print "shl=" & ncode, hwnd

End Function

Public Sub showStyle(hwnd As Long)

    Dim lngCurStyle As Long

    lngCurStyle = GetWindowLong(hwnd, GWL_STYLE)
    If (lngCurStyle Or WS_BORDER) = lngCurStyle Then Debug.Print "WS_BORDER"
    If (lngCurStyle Or WS_CHILD) = lngCurStyle Then Debug.Print "WS_CHILD"
    If (lngCurStyle Or WS_DLGFRAME) = lngCurStyle Then Debug.Print "WS_DLGFRAME"
    If (lngCurStyle Or WS_POPUP) = lngCurStyle Then Debug.Print "WS_POPUP"


'Styles
'Const WS_BORDER = &H800000
'Const WS_CAPTION = &HC00000
'Const WS_CHILD = &H40000000
'Const WS_CHILDWINDOW = WS_CHILD
'Const WS_CLIPCHILDREN = &H2000000
'Const WS_CLIPSIBLINGS = &H4000000
'Const WS_DISABLED = &H8000000
'Const WS_DLGFRAME = &H400000
'Const WS_GROUP = &H20000
'Const WS_MINIMIZE = &H20000000
'Const WS_ICONIC = WS_MINIMIZE
'Const WS_MAXIMIZE = &H1000000
'Const WS_MAXIMIZEBOX = &H10000
''Const WS_MINIMIZEBOX = &H20000
'Const WS_POPUP = &H80000000
'Const WS_SYSMENU = &H80000
'Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)'
'Const WS_THICKFRAME = &H40000
'Const WS_SIZEBOX = WS_THICKFRAME
'Const WS_TABSTOP = &H10000
'Const WS_VISIBLE = &H10000000

'Extended Styles
'Const WS_EX_ACCEPTFILES = &H10&
'Const WS_EX_APPWINDOW = &H40000
'Const WS_EX_CLIENTEDGE = &H200
'Const WS_EX_COMPOSITED = &H2000000
'Const WS_EX_CONTEXTHELP = &H400
'Const WS_EX_CONTROLPARENT = &H10000
'Const WS_EX_DLGMODALFRAME = &H1
'Const WS_EX_LAYERED = &H80000
'Const WS_EX_LAYOUTRTL = &H400000
'Const WS_EX_LEFT = &H0
'Const WS_EX_MDICHILD = &H40
'Const WS_EX_NOACTIVATE = &H8000000
'Const WS_EX_NOINHERITLAYOUT = &H100000
'Const WS_EX_NOPARENTNOTIFY = &H4
'Const WS_EX_TOPMOST = &H8
'Const WS_EX_TOOLWINDOW = &H80
'Const WS_EX_WINDOWEDGE = &H100
'Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
'Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
'Const WS_EX_RIGHT = &H1000
'Const WS_EX_RIGHTSCROLLBAR = &H0
'Const WS_EX_RTLREADING = &H2000
'Const WS_EX_STATICEDGE = &H20000
'Const WS_EX_TRANSPARENT = &H20
'Const WS_EX_LEFTSCROLLBAR = &H4000
'Const WS_EX_LTRREADING = &H0

End Sub
