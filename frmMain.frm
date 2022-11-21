VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " mmTools"
   ClientHeight    =   3780
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4245
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   252
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   283
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame fraMain 
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4215
      Begin VB.CheckBox chkSetup 
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   1080
         Width           =   255
      End
      Begin VB.ComboBox cboSetup 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   3000
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox cboSetup 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblSetup 
         Caption         =   "One screen maximize:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lblSetup 
         Caption         =   "Number of Screens:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblSetup 
         Caption         =   "Show on Screen:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   2880
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Begin VB.Menu mnuSetup 
         Caption         =   "&Setup"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nid As NOTIFYICONDATA
Dim hMenu As Long
Dim settings As clsINI
Private Sub cboSetup_Click(Index As Integer)

    Dim i As Long
    
    Select Case Index
    Case 0
        settings.SetKey "mmtools", "number_of_screens", cboSetup(Index).ListIndex + 1
        number_of_screens = cboSetup(Index).ListIndex + 1
        cboSetup(1).Clear
        cboSetup(1).AddItem "Follow"
        For i = 1 To number_of_screens
            cboSetup(1).AddItem i
        Next i
        cboSetup(1).ListIndex = show_on_screen
    Case 1
        settings.SetKey "mmtools", "show_on_screen", cboSetup(Index).ListIndex
        show_on_screen = cboSetup(Index).ListIndex
    End Select
    
End Sub
Private Sub chkSetup_Click()

    settings.SetKey "mmtools", "one_screen_maximise", chkSetup.Value
    one_screen_maximise = chkSetup.Value

End Sub
Private Sub cmdOk_Click()
    
    Me.Hide

End Sub
Private Sub Timer1_Timer()

    Dim hwnd As Long
    Dim rRect As RECT
    Dim pPoint As POINTAPI
    Dim onescreen As Long
    Dim onscreen As Long
    
    onescreen = GetSystemMetrics(SM_CXFULLSCREEN) / number_of_screens
    hwnd = FindWindow("#32771", "")
    If hwnd <> 0 Then
        If show_on_screen = 0 Then
            GetCursorPos pPoint
            onscreen = RoundToValue(pPoint.x / onescreen, 1, True)
        Else
            onscreen = show_on_screen
        End If
        GetWindowRect hwnd, rRect
        SetWindowPos hwnd, vbNull, ((onscreen - 1) * onescreen) + (onescreen / 2) - (rRect.Right - rRect.Left) / 2, (GetSystemMetrics(SM_CYFULLSCREEN) / 2) - (rRect.Bottom - rRect.Top) / 2, 0, 0, SWP_NOSIZE
    End If

End Sub
Private Sub Form_Load()
       
    Dim uRegMsg As Long
    Dim i As Long
    Dim lhWndTop As Long
    Dim lhWnd As Long
    Dim cx As Long
    Dim cy As Long
    Dim hIconLarge As Long
    Dim hIconSmall As Long
      
    Me.Caption = " Multi Monitor Tools Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    Set settings = New clsINI
    number_of_screens = CInt(settings.GetKey("mmtools", "number_of_screens", "2"))
    show_on_screen = CInt(settings.GetKey("mmtools", "show_on_screen", "0"))
    one_screen_maximise = CInt(settings.GetKey("mmtools", "one_screen_maximise", "1"))
    
    
    For i = 1 To 9
        cboSetup(0).AddItem i
    Next i
    cboSetup(0).ListIndex = number_of_screens - 1
    
    cboSetup(1).Clear
    cboSetup(1).AddItem "Follow"
    For i = 1 To number_of_screens
        cboSetup(1).AddItem i
    Next i
    cboSetup(1).ListIndex = show_on_screen
    
    chkSetup.Value = one_screen_maximise
    
    hMenu = CreatePopupMenu()
    AppendMenu hMenu, MF_STRING, 1, "&Setup"
    AppendMenu hMenu, MF_SEPARATOR, 2, ByVal 0&
    AppendMenu hMenu, MF_STRING, 3, "E&xit"
    
    ' Find VB's hidden parent window:
    lhWnd = hwnd
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
    SendMessageLong hwnd, WM_SETICON, ICON_BIG, hIconLarge
    
    cx = GetSystemMetrics(SM_CXSMICON)
    cy = GetSystemMetrics(SM_CYSMICON)
    hIconSmall = LoadImageAsString(App.hInstance, "AAA", IMAGE_ICON, cx, cy, LR_LOADMAP3DCOLORS)
    SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
    SendMessageLong hwnd, WM_SETICON, ICON_SMALL, hIconSmall
    
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = hIconSmall
        .szTip = Me.Caption & vbNullChar
    End With

    Shell_NotifyIcon NIM_ADD, nid
    
'    Call SetCBTSHLHook(-1, AddressOf cbtCallback, AddressOf shlCallback)
    Call SetCBTSHLHook(-1, AddressOf cbtCallback, 0)
    
    Me.Visible = False

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim Result As Long
    Dim msg As Long
    Dim Pt As POINTAPI
    
    If Me.ScaleMode = vbPixels Then
        msg = x
    Else
        msg = x / Screen.TwipsPerPixelX
    End If
    Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
            showSetup
        Case WM_LBUTTONDBLCLK    '515 restore form window
            showSetup
        Case WM_RBUTTONUP       '517 display popup menu
            Result = SetForegroundWindow(Me.hwnd)
            GetCursorPos Pt
            Result = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, Pt.x, Pt.y, Me.hwnd, ByVal 0&)
            Select Case Result
            Case 1
                showSetup
            Case 3
                Unload Me
            End Select
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Call SetCBTSHLHook(0, 0, 0)
    DestroyMenu hMenu
    Shell_NotifyIcon NIM_DELETE, nid

End Sub
Private Function RoundToValue(ByVal nValue, nCeiling As Double, Optional RoundUp As Boolean = True) As Double
    
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
Private Sub showSetup()

    Dim pPoint As POINTAPI
    Dim onescreen As Long
    Dim onscreen As Long
    
    onescreen = GetSystemMetrics(SM_CXFULLSCREEN) / number_of_screens
    
    GetCursorPos pPoint
    onscreen = RoundToValue(pPoint.x / onescreen, 1, True) - 1
    SetForegroundWindow Me.hwnd
    frmMain.Left = (onscreen * onescreen) * Screen.TwipsPerPixelX + (onescreen * Screen.TwipsPerPixelX - frmMain.width) / 2
    frmMain.Top = (Screen.height - frmMain.height) / 2
    Me.Visible = True

End Sub
