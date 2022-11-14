VERSION 5.00
Begin VB.Form compact 
   BorderStyle     =   0  'None
   Caption         =   "Stardisk's Player Compact Mode"
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   17
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   301
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2040
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   9
      Top             =   0
      Width           =   1215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   45
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00B0B0B0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4260
      TabIndex        =   8
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   960
      Top             =   120
   End
   Begin VB.CommandButton Command7 
      Caption         =   "+"
      Height          =   255
      Left            =   1740
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "-"
      Height          =   255
      Left            =   1260
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblVolume 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   1530
      TabIndex        =   11
      ToolTipText     =   "Громкость"
      Top             =   45
      Width           =   195
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0:00:00 / 0:00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   3300
      TabIndex        =   6
      Top             =   45
      Width           =   915
   End
End
Attribute VB_Name = "compact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCaption = 2

Dim symbol As Integer
Dim step As Long

Private Sub ButtonColor()
    If main.wmp.playState = 3 Or main.wmp.playState = 9 Then Command1.BackColor = "&H00C0C0C0": Command2.BackColor = "&H8000000F": Command3.BackColor = "&H8000000F"
    If main.wmp.playState = 2 Then Command1.BackColor = "&H8000000F": Command2.BackColor = "&H00C0C0C0": Command3.BackColor = "&H8000000F"
    If main.wmp.playState < 2 Or main.wmp.playState = 8 Then Command1.BackColor = "&H8000000F": Command2.BackColor = "&H8000000F": Command3.BackColor = "&H00C0C0C0"
End Sub

Private Sub SavePosition()
    WriteINIKey "CompactPosition", "X", Me.Left, main.presets
    WriteINIKey "CompactPosition", "Y", Me.Top, main.presets
End Sub

Private Sub Command1_Click()
    main.cmdplay_Click
    ButtonColor
End Sub

Private Sub Command2_Click()
    main.cmdpause_Click
    ButtonColor
End Sub

Private Sub Command3_Click()
    main.cmdstop_Click
    ButtonColor
End Sub

Private Sub Command4_Click()
    main.cmdprev_Click
    symbol = 0
    ButtonColor
End Sub

Private Sub Command5_Click()
    main.cmdnext_Click
    symbol = 0
    ButtonColor
End Sub

Private Sub Command6_Click()
On Error GoTo err
main.hsvolume.Value = main.hsvolume.Value - step
lblVolume = main.wmp.settings.volume
Exit Sub
err:
main.hsvolume.Value = 0
lblVolume = main.wmp.settings.volume
End Sub

Private Sub Command7_Click()
On Error GoTo err
main.hsvolume.Value = main.hsvolume.Value + step
lblVolume = main.wmp.settings.volume
Exit Sub
err:
main.hsvolume.Value = 100
lblVolume = main.wmp.settings.volume
End Sub

Private Sub Command8_Click()
    SavePosition
    Unload Me
    main.WindowState = 0
    main.Show
End Sub

Private Sub Form_Activate()
  If ReadINIKey("CompactPosition", "X", main.presets) = "" Then
    compact.Left = 0
  Else
    compact.Left = ReadINIKey("CompactPosition", "X", main.presets)
  End If
  
  If ReadINIKey("CompactPosition", "Y", main.presets) = "" Then
    compact.Top = 0
  Else
    compact.Top = ReadINIKey("CompactPosition", "Y", main.presets)
  End If
  ButtonColor
End Sub

Private Sub Form_Load()
  Call SetWindowPos(compact.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
  Timer1_Timer
  lblVolume = main.wmp.settings.volume
  step = ReadINIKey("General", "CompactStepVolume", main.presets)
  ButtonColor
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCaption, 0&)
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCaption, 0&)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCaption, 0&)
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCaption, 0&)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    Label2 = main.lblTiming
    Label1 = main.lblTitle
    Label1.Left = Label1.Left - 5
    If Label1.Left < 0 - Label1.Width Then Label1.Left = Picture1.Width
    Shape1.Width = 80 / main.hsprogress.Max * main.hsprogress.Value
End Sub
