VERSION 5.00
Begin VB.Form toRemove 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   240
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Text            =   "-3000"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "toRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type IMediaType
    Control As IMediaControl
    Position As IMediaPosition
    Audio As IBasicAudio
    
End Type

Dim IMedia As IMediaType

Private Sub LoadMP3(Filename$, Optional ByVal Volume&)
    Set IMedia.Control = New FilgraphManager
    IMedia.Control.RenderFile Filename
    Set IMedia.Audio = IMedia.Control
    IMedia.Audio.Volume = Volume
    Set IMedia.Position = IMedia.Control
    IMedia.Position.Rate = 1
    IMedia.Position.CurrentPosition = 0
End Sub

Private Sub DestoryMP3()
    Set IMedia.Audio = Nothing
    Set IMedia.Position = Nothing
    Set IMedia.Control = Nothing
End Sub

Private Sub Command1_Click()
    IMedia.Control.Run
    
    
    
End Sub

Private Sub Command2_Click()
    IMedia.Control.Stop
    
End Sub

Private Sub Command3_Click()
    IMedia.Audio.Volume = Text1.Text
End Sub

Private Sub Command4_Click()
    LoadMP3 (Text2.Text)
    Timer1.Enabled = True
    
End Sub

Private Sub Command5_Click()
    IMedia.Control.Pause
    
    
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = IMedia.Position.CurrentPosition & "/" & IMedia.Position.Duration
End Sub
