Attribute VB_Name = "modApi"
Option Explicit
' Settings
Public number_of_screens As Long
Public show_on_screen As Long
Public one_screen_maximise As Long
' ----

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
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
Public Declare Function AdjustWindowRectEx Lib "user32" (lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetWinEventHook Lib "user32.dll" (ByVal eventMin As Long, ByVal eventMax As Long, ByVal hmodWinEventProc As Long, ByVal pfnWinEventProc As Long, ByVal idProcess As Long, ByVal idThread As Long, ByVal dwFlags As Long) As Long
Public Declare Function UnhookWinEvent Lib "user32.dll" (ByVal lHandle As Long) As Long

Public Type WINEVENTPROC
    hWinEventHook As Long
    event As Long
    hwnd As Long
    idObject As Long
    idChild As Long
    idEventThread As Long
    dwmsEventTime As Long
End Type

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
Public Const WS_CHILD = &H40000000
Public Const WM_USER = &H400

Public Const EVENT_MIN = 1&
Public Const ROLE_COMBOBOX = &H2E&
Public Const OB_ACCELERATORCHANGE = &H8012&
Public Const SYS_ALERT = 2&
Public Const WINEVENT_SKIPOWNPROCESS = 2&
Public Const OB_REORDER = &H8004&
Public Const EVENT_SYSTEM_SWITCHSTART = 20
Public Const GWL_WNDPROC = (-4)

Const SPI_GETWORKAREA = 48
Public OldWindowProc As Long
Public hMenu As Long
Public settings As clsINI
Public nid As NOTIFYICONDATA
Public LHook As Long
Sub Main()

    Dim uRegMsg As Long
    Dim i As Long
    Dim lhWndTop As Long
    Dim lhWnd As Long
    Dim cx As Long
    Dim cy As Long
    Dim hIconLarge As Long
    Dim hIconSmall As Long
      
    Load frmMain
    frmMain.Caption = " Multi Monitor Tools Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Set settings = New clsINI
    
    hMenu = CreatePopupMenu()
    AppendMenu hMenu, MF_STRING, 1, "&Setup"
    AppendMenu hMenu, MF_SEPARATOR, 2, ByVal 0&
    AppendMenu hMenu, MF_STRING, 3, "E&xit"
    
    ' Find VB's hidden parent window:
    lhWnd = frmMain.hwnd
    lhWndTop = lhWnd
    Do While Not (lhWnd = 0)
       lhWnd = GetWindow(lhWnd, GW_OWNER)
       If Not (lhWnd = 0) Then
          lhWndTop = lhWnd
       End If
    Loop
    
    cx = GetSystemMetrics(SM_CXICON)
    cy = GetSystemMetrics(SM_CYICON)
    hIconLarge = LoadImageAsString(App.hInstance, "AAA", IMAGE_ICON, cx, cy, LR_LOADMAP3DCOLORS)
    SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
    SendMessageLong frmMain.hwnd, WM_SETICON, ICON_BIG, hIconLarge
    
    cx = GetSystemMetrics(SM_CXSMICON)
    cy = GetSystemMetrics(SM_CYSMICON)
    hIconSmall = LoadImageAsString(App.hInstance, "AAA", IMAGE_ICON, cx, cy, LR_LOADMAP3DCOLORS)
    SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
    SendMessageLong frmMain.hwnd, WM_SETICON, ICON_SMALL, hIconSmall
    
    With nid
        .cbSize = Len(nid)
        .hwnd = frmMain.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_USER
        .hIcon = hIconSmall
        .szTip = frmMain.Caption & vbNullChar
    End With

    Shell_NotifyIcon NIM_ADD, nid
    
    OldWindowProc = SetWindowLong(frmMain.hwnd, GWL_WNDPROC, AddressOf WindowProc)
'    Call SetCBTSHLHook(-1, AddressOf cbtCallback, AddressOf shlCallback)
    Call SetCBTSHLHook(-1, AddressOf cbtCallback, 0)
    LHook = SetWinEventHook(EVENT_MIN, OB_ACCELERATORCHANGE, 0, AddressOf WinEventFunc, 0, 0, WINEVENT_SKIPOWNPROCESS)
'    LHook = SetWinEventHook(SYS_ALERT, OB_ACCELERATORCHANGE, 0&, AddressOf WinEventFunc, 0, 0, WINEVENT_SKIPOWNPROCESS)

End Sub
Sub shutdown()
    
    UnhookWinEvent LHook
    Call SetCBTSHLHook(0, 0, 0)
    DestroyMenu hMenu
    Shell_NotifyIcon NIM_DELETE, nid
    SetWindowLong frmMain.hwnd, GWL_WNDPROC, OldWindowProc
    Unload frmMain

End Sub
Public Function cbtCallback(ByVal hwnd As Long, ByVal ncode As Long) As Long

    Dim height As Long
    Dim width As Long
    Dim rtn As Long
    Dim WinEst As WINDOWPLACEMENT
    Dim R As RECT
    Dim lpClassName As String
    Dim lngCurStyle As Long
    Dim sBuf As String
    Dim iPos As Long
    Dim ClassName As String
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
    ElseIf ncode = HCBT_CREATEWND Then
        sBuf = String$(255, 0)
        GetClassName hwnd, sBuf, 255
        iPos = InStr(sBuf, Chr$(0))
        If iPos > 1 Then
            ClassName = Left$(sBuf, iPos - 1)
        End If
        If (InStr(ClassName, "#32770")) Then
            Debug.Print "Dialog created"
        End If
        If (InStr(ClassName, "#32771")) Then
            Debug.Print "Task switch created"
        End If
    Else
        Debug.Print "cbt=" & ncode, hwnd
    End If

End Function
Public Function shlCallback(ByVal hwnd As Long, ByVal ncode As Long) As Long

    Debug.Print "shl=" & ncode, hwnd

End Function
Public Function WinEventFunc(ByVal HookHandle As Long, ByVal LEvent As Long, ByVal hwnd As Long, ByVal idObject As Long, ByVal idChild As Long, ByVal idEventThread As Long, ByVal dwmsEventTime As Long) As Long
    
    Dim rRect As RECT
    Dim pPoint As POINTAPI
    Dim onescreen As Long
    Dim onscreen As Long

    If LEvent = EVENT_SYSTEM_SWITCHSTART Then
        onescreen = GetSystemMetrics(SM_CXFULLSCREEN) / number_of_screens
        If show_on_screen = 0 Then
            GetCursorPos pPoint
            onscreen = RoundToValue(pPoint.x / onescreen, 1, True)
        Else
            onscreen = show_on_screen
        End If
        GetWindowRect hwnd, rRect
        SetWindowPos hwnd, vbNull, ((onscreen - 1) * onescreen) + (onescreen / 2) - (rRect.Right - rRect.Left) / 2, (GetSystemMetrics(SM_CYFULLSCREEN) / 2) - (rRect.Bottom - rRect.Top) / 2, 0, 0, SWP_NOSIZE
    End If
    WinEventFunc = 0
    
End Function
Public Function WindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim Result As Long
    Dim Pt As POINTAPI
    
    If frmMain.hwnd = hwnd And Msg = WM_USER Then
        Select Case lParam
            Case WM_LBUTTONUP
                frmMain.showSetup
            Case WM_LBUTTONDBLCLK
                frmMain.showSetup
            Case WM_RBUTTONUP
                Result = SetForegroundWindow(frmMain.hwnd)
                GetCursorPos Pt
                Result = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, Pt.x, Pt.y, frmMain.hwnd, ByVal 0&)
                Select Case Result
                Case 1
                    frmMain.showSetup
                Case 3
                    shutdown
                Case Else
                    WindowProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)
                End Select
        End Select
    Else
        WindowProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)
    End If

End Function
Public Function RoundToValue(ByVal nValue, nCeiling As Double, Optional RoundUp As Boolean = True) As Double
    
    Dim tmp As Integer
    Dim tmpVal
    If Not IsNumeric(nValue) Then Exit Function
    nValue = CDbl(nValue)
    
    tmpVal = ((nValue / nCeiling) + (-0.5 + (RoundUp And 1)))
    tmp = Fix(tmpVal)
    tmpVal = CInt((tmpVal - tmp) * 10 ^ 0)
    nValue = tmp + tmpVal / 10 ^ 0

    RoundToValue = nValue * nCeiling
       
End Function
