VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " mmTools"
   ClientHeight    =   5460
   ClientLeft      =   150
   ClientTop       =   135
   ClientWidth     =   4245
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   364
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   283
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
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
      Height          =   4935
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
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboSetup_Click(Index As Integer)

    Dim i As Long
    Dim store As Long
    
    If Index = 0 Then
        cboSetup(1).Clear
        cboSetup(1).AddItem "Follow"
        For i = 1 To cboSetup(Index).ListIndex + 1
            cboSetup(1).AddItem i
        Next i
        If show_on_screen < cboSetup(1).ListCount Then
            cboSetup(1).ListIndex = show_on_screen
        Else
            cboSetup(1).ListIndex = -1
        End If
    End If
    
End Sub
Private Sub cmdOk_Click()
    
    number_of_screens = cboSetup(0).ListIndex + 1
    show_on_screen = cboSetup(1).ListIndex
    one_screen_maximise = chkSetup.Value
    
    settings.SetKey "mmtools", "one_screen_maximise", chkSetup.Value
    settings.SetKey "mmtools", "number_of_screens", cboSetup(0).ListIndex + 1
    settings.SetKey "mmtools", "show_on_screen", cboSetup(1).ListIndex
    Me.Visible = False

End Sub
Public Sub showSetup()

    Dim pPoint As POINTAPI
    Dim onescreen As Long
    Dim onscreen As Long
    Dim i As Long
    
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
    
    onescreen = GetSystemMetrics(SM_CXFULLSCREEN) / number_of_screens
    
    GetCursorPos pPoint
    onscreen = RoundToValue(pPoint.x / onescreen, 1, True) - 1
    SetForegroundWindow Me.hwnd
    frmMain.Left = (onscreen * onescreen) * Screen.TwipsPerPixelX + (onescreen * Screen.TwipsPerPixelX - frmMain.width) / 2
    frmMain.Top = (Screen.height - frmMain.height) / 2
    Me.Visible = True

End Sub
