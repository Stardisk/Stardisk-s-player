VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stardisk's Player"
   ClientHeight    =   5700
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRepeat 
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4680
      Width           =   375
   End
   Begin VB.HScrollBar hsprogress 
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4320
      Width           =   4455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "..."
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      Top             =   0
      Width           =   375
   End
   Begin VB.HScrollBar hsvolume 
      Height          =   255
      LargeChange     =   10
      Left            =   3000
      Max             =   100
      TabIndex        =   8
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   240
   End
   Begin VB.Frame playPartWrapper 
      Caption         =   "������ �����"
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CheckBox cmdApply 
         Caption         =   "�������� ������"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1200
         Width           =   2055
      End
      Begin MSMask.MaskEdBox txtEndTime 
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtStartTime 
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "������ ��������"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1620
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�� �������:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "� �������:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   375
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   120
      Pattern         =   "*.mp3;*.wma;*.ogg;*.wav;*.m3u"
      TabIndex        =   2
      Top             =   1800
      Width           =   4455
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   120
      TabIndex        =   18
      Top             =   5160
      Width           =   2895
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "0.00 ��"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   4080
      Width           =   585
   End
   Begin VB.Label lblItems 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "000 / 000"
      Height          =   195
      Left            =   3870
      TabIndex        =   23
      Top             =   4080
      Width           =   690
   End
   Begin VB.Label lblBitrate 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 ����/�"
      Height          =   195
      Left            =   2040
      TabIndex        =   22
      Top             =   4080
      Width           =   645
   End
   Begin VB.Label lblVolume 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0%"
      Height          =   195
      Left            =   4320
      TabIndex        =   16
      Top             =   4920
      Width           =   270
   End
   Begin VB.Label lblTiming 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "00:00 / 00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   15
      Top             =   5160
      Width           =   1500
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   15
      Left            =   4080
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   15
      URL             =   ""
      rate            =   -10
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "invisible"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   26
      _cy             =   26
   End
   Begin VB.Menu mnuMain 
      Caption         =   "������� ����� ��� ������ ����"
      WindowList      =   -1  'True
      Begin VB.Menu mnuRename 
         Caption         =   "������������� ���������� ����"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuMove 
         Caption         =   "����������� �..."
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "������� ���������� ����"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuOpenLocation 
         Caption         =   "������� ������������ �����"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "����� ���� � ������� �����"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowPlayPart 
         Caption         =   "�������� ""������ �����"""
      End
      Begin VB.Menu mnuShowTrackInCaption 
         Caption         =   "�������� ����� � ���������"
      End
      Begin VB.Menu mnuMinimizeToCompact 
         Caption         =   "����������� � ���������� ����"
      End
      Begin VB.Menu mnuCloseToCompact 
         Caption         =   "��������� � ���������� �����"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGotoCompact 
         Caption         =   "������� � ���������� �����"
      End
      Begin VB.Menu mnuCompactVolume 
         Caption         =   "��� ��������� � ���������� ������"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "� ���������"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "�����"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)     '��� sleep��
Dim repeating, playing, showTrackInCaption As Boolean                                   '���������� ��� ������ ��������������� � �������
Dim starttime, endtime As Integer                                   '��������� � �������� ����� ���������������
Public presets As String                                            '���� � ����������
Dim repeatlist As String                                            '���� � ������ ����� � ������ ����������� ������
Dim tmplistindex As Integer                                         '���������� ��� ����������� ������ ��������� ������ ��� �����������/��������/�������������� �����
Dim tmpPosition As Variant                                          '���������� ��� ����������� �� ����� ����� ������ ��������������� ��� �������������� �����

Private Sub ButtonColor()                                           '������ ���� ������
    If wmp.playState = 3 Or wmp.playState = 9 Then cmdPlay.BackColor = "&H00C0C0C0": cmdPause.BackColor = "&H8000000F": cmdStop.BackColor = "&H8000000F" ' ����������� ���������������
    If wmp.playState = 2 Then cmdPlay.BackColor = "&H8000000F": cmdPause.BackColor = "&H00C0C0C0": cmdStop.BackColor = "&H8000000F"                      ' ���������� �����
    If wmp.playState < 2 Or wmp.playState = 8 Then cmdPlay.BackColor = "&H8000000F": cmdPause.BackColor = "&H8000000F": cmdStop.BackColor = "&H00C0C0C0"                      ' ���������� ����
End Sub

Private Sub TimeToSeconds()                                         '�������������� ������� ��:�� � �������
    starttime = CInt(Mid(txtStartTime, 1, 2)) * 60 + CInt(Mid(txtStartTime, 4, 2))
    endtime = CInt(Mid(txtEndTime, 1, 2)) * 60 + CInt(Mid(txtEndTime, 4, 2))
End Sub

Private Sub SecondsToTime()                                         '�������������� ������ � ��:��
    txtStartTime = Right(TimeSerial(0, 0, starttime), 5)
    txtEndTime = Right(TimeSerial(0, 0, endtime), 5)
End Sub

Private Sub ActionsOnPlay()                                                 '�������� ��� ���������������
On Error Resume Next
    lblTitle = "": lblBitrate = ""                                          '������� �������� � �������, ������� ����� �������� �� ����� �� ����������� �����
    TimeToSeconds                                                           '��������� ��:�� � �������
    cmdApply.Value = 0                                                      '��������� ������
    wmp.URL = Dir1.Path & "\" & File1.filename                              '�������� ������ ���� � ���������� �����
    lblItems = File1.ListIndex + 1 & " / " & File1.ListCount                '���������� ����� �������� �������� � ����� ����� ���������
    lblSize = Format(FileLen(Dir1.Path & "\" & File1.filename) / 1048576, "0.00") & " ��"   '���������� ��� �����
    playing = True                                                          '�������� ���� ���������������
    hsprogress.Value = 0

    If wmp.currentMedia.getItemInfo("Author") <> "" Then                                                    '���� � ����� ��� ����� �� ������
        lblTitle = wmp.currentMedia.getItemInfo("Title") & " - " & wmp.currentMedia.getItemInfo("Author")       '����� �������� ����� � ������
    Else                                                                                                    '�����
        lblTitle = wmp.currentMedia.getItemInfo("Title")                                                        '����� ������ �������� �����, ��� ������ ���� ���� ������������� ��� �����
    End If
    
    If showTrackInCaption = True Then main.Caption = lblTitle.Caption
    
    lblBitrate = Format(wmp.currentMedia.getItemInfo("bitrate") / 1000, "#") & " ����/�"
    
    If ReadINIKey(StrConv(wmp.currentMedia.sourceURL, vbLowerCase), "Start", repeatlist) <> "" Then     '���� ���� � ����� ���� � ������ ��� ��������
        starttime = ReadINIKey(StrConv(wmp.currentMedia.sourceURL, vbLowerCase), "Start", repeatlist)   '��������� ��������� �������� �������
        endtime = ReadINIKey(StrConv(wmp.currentMedia.sourceURL, vbLowerCase), "End", repeatlist)       '��������� ������������� �������� �������
        SecondsToTime
        If repeating = True Then wmp.Controls.currentPosition = starttime                               '���� ���� ������� �������, ������������� ���������� � ������� �������
    Else
        starttime = 0                                        '���� ���� � ����� ��� � ������ ��������, �� ������ ��� �� 0
        endtime = 0
        txtStartTime = "00:00"
        txtEndTime = "00:00"
    End If
    Timer1.Enabled = True                                           '�������� ������ �������� �������
    WriteINIKey "General", "LastIndex", File1.ListIndex, presets    '���������� ��������� ���������������� ����
    ButtonColor
End Sub

Private Sub Form_Resize()                                                   '������������ � ���������� �����
    If Me.WindowState = 1 Then                                                  '���� ������� ���� �������������
        If mnuMinimizeToCompact.Checked = True Then
            Me.Hide                                                                     '�������� ���
            compact.Show                                                                '���������� ���������� �����
        End If
    End If
End Sub

Private Sub mnuAbout_Click()                                        '���� "� ���������"
    MsgBox "Stardisk's Player v." & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
    "�����: ����������� ���� �������� aka Stardisk", vbInformation
End Sub

Private Sub mnuCloseToCompact_Click()
    If mnuCloseToCompact.Checked = False Then
        mnuCloseToCompact.Checked = True
        WriteINIKey "View", "CloseToCompact", "1", presets
    Else
        mnuCloseToCompact.Checked = False
        WriteINIKey "View", "CloseToCompact", "0", presets
    End If
End Sub

Private Sub mnuCompactVolume_Click()                                '���� ��������� ���� ��������� ��� ������ � ���������� ������
    Dim tmp As Variant
    tmp = ReadINIKey("General", "CompactStepVolume", presets)
    tmp = InputBox("�������, �� ������� ��������� ����� ���������� ������� ��������� ��� ������� ������ ���������� ���������� � ���������� ������:", , tmp)
    If tmp = "" Then Exit Sub
    If IsNumeric(tmp) = False Then MsgBox "������� �����": Exit Sub
    If tmp < 1 Or tmp > 100 Then MsgBox "����� ������ ���� �� 1 �� 100": Exit Sub
    
    WriteINIKey "General", "CompactStepVolume", CInt(tmp), presets
End Sub

Private Sub mnuGotoCompact_Click()                                  '���� ������������ � ���������� �����
    Unload Me
End Sub

Private Sub mnuMinimizeToCompact_Click()
    If mnuMinimizeToCompact.Checked = False Then
        mnuMinimizeToCompact.Checked = True
        WriteINIKey "View", "MinimizeToCompact", "1", presets
    Else
        mnuMinimizeToCompact.Checked = False
        WriteINIKey "View", "MinimizeToCompact", "0", presets
    End If
End Sub

Private Sub mnuMove_Click()
    If File1.ListIndex = -1 Then Exit Sub
    Dim tmp As String
    Dim tmpfilesize As String
    tmp = InputBox("���� ����������?", "", ReadINIKey("General", "LastMovePath", presets))  '����������� ���� ��� �����������
    If StrConv(tmp, vbLowerCase) = StrConv(Dir1.Path, vbLowerCase) Then MsgBox "���� � ��� ��� ��� �����.": Exit Sub '���� ��� ������ ��� �� ����, ��� � ��������, ����� ������ � ������� �� ������
    If tmp <> "" Then                                                   '���� �� ���� ������ ������
        cmdstop_Click                                                   '������������� ��������������� ��� ������������ �����
        wmp.URL = ""                                                    '������� ����� �� ���
        Timer1.Enabled = False                                          '������������� ������
        tmplistindex = File1.ListIndex                                  '���������� ����� ����������������� �����
        tmpPosition = wmp.Controls.currentPosition                      '���������� ����� ���������������
        tmpfilesize = FileLen(Dir1.Path & "\" & File1.filename)         '���������� ������ �����
        On Error GoTo err
            Name Dir1.Path & "\" & File1.filename As tmp & "\" & File1.filename '���������� ����
        On Error GoTo 0
        While FileLen(tmp & "\" & File1.filename) <> tmpfilesize        '���������, ������������ �� ���� ��� ���
            Sleep (1000)                                                '���� ��� �� �������������, ���� �������
        Wend
        WriteINIKey "General", "LastMovePath", tmp, presets             '���������� ��������� �������������� ���� ��� �����������
        File1.Refresh                                                   '��������� ������ ������
        If tmplistindex > File1.ListCount - 1 Then File1.ListIndex = tmplistindex - 1 Else File1.ListIndex = tmplistindex ' ���� ���� ��� ��������� � ������, ��������� �� ����������, ���� �� ��� ���������, �������� ���, ������� ���� ������ ���� �������
        Timer1.Enabled = True
    End If
    Exit Sub
err:
    MsgBox "�� ������� ����������� ����."
End Sub

Private Sub mnuOpenLocation_Click()
    Shell "explorer.exe /select, " & wmp.URL, vbNormalFocus
End Sub

Private Sub mnuRename_Click()                                       '���� �������������� �����
On Error Resume Next
    If File1.ListIndex = -1 Then Exit Sub                           '���� ������� ���� �� ������, �������
    Dim newName As String                                           '���������� � ������������ ����� ��� �����
    newName = InputBox("����� ��� ��� ����� " & File1.List(File1.ListIndex), , File1.List(File1.ListIndex))
    If newName = "" Then Exit Sub                                   '�������, ���� ������ ������

    tmplistindex = File1.ListIndex                                  '���������� ����� ����������������� �����
    tmpPosition = wmp.Controls.currentPosition                      '���������� ����� ���������������
    cmdstop_Click                                                   '������������� ���������������, ����� ���� �� �������������
    Name Dir1.Path & "\" & File1.List(File1.ListIndex) As Dir1.Path & "\" & newName '���������������
    File1.Refresh                                                   '��������� ������ ������
    File1.ListIndex = tmplistindex                                  '��������� �� ����������� �����
    wmp.URL = Dir1.Path & "\" & newName                             '�������� ������ ����� ��� ���������������� �����
    wmp.Controls.currentPosition = tmpPosition                      '��������� ��������������� � ���� �����, �� ������� �� ��� ������������
    If playing = True Then wmp.Controls.play                        '�������������
End Sub

Public Sub cmdstop_Click()                                          '������ ��������� ���������������
    wmp.Controls.stop                                               '������������� ���������������
    playing = False                                                 '������� ���� ���������������
    ButtonColor                                                     '������ ������� ����
    If showTrackInCaption = True Then main.Caption = "Stardisk's Player"
End Sub

Private Sub cmdApply_Click()                                        '������ �������
On Error GoTo err
If File1.ListIndex < 0 Then cmdApply.Value = 0: Exit Sub            '���� ���� �� ������, �������
If cmdApply.Value = 1 Then                                          '���� ������ ������
    TimeToSeconds
    repeating = True
    If starttime <> 0 And endtime <> 0 Then wmp.Controls.currentPosition = starttime
    Label5.Caption = "������ �������"
    Label5.FontBold = True
    txtStartTime.Enabled = False
    txtEndTime.Enabled = False
    
    If starttime > 0 Or endtime > 0 Then
        WriteINIKey StrConv(wmp.currentMedia.sourceURL, vbLowerCase), "Start", CStr(starttime), repeatlist
        WriteINIKey StrConv(wmp.currentMedia.sourceURL, vbLowerCase), "End", CStr(endtime), repeatlist
    End If
    Exit Sub
Else
    repeating = False
    Label5.Caption = "������ ��������"
    Label5.FontBold = False
    txtStartTime.Enabled = True
    txtEndTime.Enabled = True
    Exit Sub
End If
err:
    MsgBox "�������� ������."
    cmdApply.Value = 0
End Sub

Public Sub cmdpause_Click()                                     '������ �����
    wmp.Controls.pause                                          '���������������� ���������������
    ButtonColor                                                 '�������� ���� ������
End Sub

Public Sub cmdplay_Click()                                      '������ ��������������
    If wmp.playState = 1 And repeating = True Then wmp.Controls.currentPosition = starttime
    wmp.Controls.play
    playing = True
    ButtonColor
    If showTrackInCaption = True Then main.Caption = lblTitle.Caption
End Sub

Public Sub cmdprev_Click()
    If File1.ListIndex <= 0 Then File1.ListIndex = File1.ListCount - 1 Else File1.ListIndex = File1.ListIndex - 1
End Sub

Public Sub cmdnext_Click()
On Error GoTo err
    File1.ListIndex = File1.ListIndex + 1: Exit Sub
err:
    If File1.ListCount > 0 Then File1.ListIndex = 0
End Sub

Private Sub mnuClose_Click()
    End
End Sub

Private Sub mnuDelete_Click()                                           '���� �������� �����
'On Error Resume Next
    If File1.ListIndex > -1 Then                                        '���� ���� ������
        If MsgBox("������� " & File1.filename & "?", vbYesNo + vbQuestion) = vbYes Then '������ ������ ��� ��������
            cmdstop_Click                                               '������������� �������������
            Timer1.Enabled = False                                      '������������� ������
            wmp.URL = ""
            lblTitle = "�������� �����..."                              '����� ��� ���� ���������
            tmplistindex = File1.ListIndex                              '���������� ����� ����� �����
            Me.Enabled = False                                          '��������� �����, ���� ������������ ���� �� ������
            mnuMain.Enabled = False                                     '��������� ����, ��� �� �������
            Kill Dir1.Path & "\" & File1.filename                       '������� ����
            While Len(Dir(Dir1.Path & "\" & File1.filename)) > 0        '���������, �������� �� ���� ��� ���
                Sleep (200)                                                 '���� �� ��������, ���� 1000 ��
                DoEvents
            Wend
            Me.Enabled = True                                           '�������� �����
            mnuMain.Enabled = True                                      '�������� ����
            File1.Refresh                                               '��������� ������
            If tmplistindex > File1.ListCount - 1 Then File1.ListIndex = tmplistindex - 1 Else File1.ListIndex = tmplistindex ' ���� ���� ��� ��������� � ������, ��������� �� ����������, ���� �� ��� ���������, �������� ���, ������� ���� ������ ���� �������
            Timer1.Enabled = True
        End If
    End If
End Sub

Private Sub Command8_Click()                                            '������ ������� ����� ����
    On Error Resume Next
    Dir1.Path = InputBox("������� ���� � �����:", "", Dir1.Path)
End Sub

Private Sub Dir1_Change()                                               '��� �������� �� ������ � ���� ���������
    File1.Path = Dir1.Path                                              '���������� ����� � ������� �����
    WriteINIKey "General", "LastPath", Dir1.Path, presets               '���������� ���� � ��������� ��������� ����� � ���������
End Sub

Private Sub Drive1_Change()                                             '����� ����� � ���� ���������
On Error GoTo err
    Dir1.Path = Drive1.Drive                                            '���������� ����� �� �����
    Exit Sub
err:
    MsgBox "���������� ��������� ����������. ��������, ���� ����������� ��� ���������."
    Drive1.Drive = Left(App.Path, 2)
End Sub

Private Sub File1_Click()                                               '���� �� ����� � ���� ���������
    If File1.ListIndex > -1 And File1.ListIndex < File1.ListCount Then ActionsOnPlay '���� ����� �������� ������ ����������, �� ��������� �������� ��� ��������������� - �������� ����� � �.�.
End Sub

Private Sub Form_Load()                                                 '�������� ��� �������
On Error Resume Next
    If App.PrevInstance = True Then End                                 '���� ��� ������� ���� �����, ��������� ������
    presets = App.Path & "\presets.ini"                                 '���������� ���� � ���������� � ����������
    repeatlist = App.Path & "\repeatlist.ini"                           '���������� ���� � ������ ���������� � ����������
    txtStartTime = "00:00"                                              '����� ������ � ������
    txtEndTime = "00:00"
    Dir1.Path = ReadINIKey("General", "LastPath", presets)              '��������� �����, � ������� ���� ��������� ���
    File1.ListIndex = ReadINIKey("General", "LastIndex", presets)       '�������� ����, �� ������� ������������ � ��������� ���
    hsvolume.Value = ReadINIKey("General", "Volume", presets)           '������������� ������� ��������� �������� �� ��������
    wmp.settings.volume = hsvolume.Value                                '������������� ��������������� ���������
    If ReadINIKey("General", "CompactStepVolume", presets) = "" Then WriteINIKey "General", "CompactStepVolume", "5", presets '���� � ���������� �� ������ ��� ��������� ������� ����������� ������, ������������� �� ��������� 5
    If ReadINIKey("View", "ShowPlayPart", presets) = "1" Then mnuShowPlayPart_Click
    If ReadINIKey("View", "ShowTrackInCaption", presets) = "1" Then mnuShowTrackInCaption_Click
    If ReadINIKey("View", "MinimizeToCompact", presets) = "1" Then mnuMinimizeToCompact_Click
    If ReadINIKey("View", "CloseToCompact", presets) = "1" Then mnuCloseToCompact_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)                              '���� ������� ���� ��������� ���������
    If mnuCloseToCompact.Checked = True Then
        Cancel = 1                                                          '������ ���������� ���������
        Me.Hide                                                             '������ �����
        compact.Show                                                        '���������� ���������� �����
    End If
End Sub

Private Sub hsprogress_Scroll()                                         '��������� ������� ��������������� ��������� �������
    wmp.Controls.currentPosition = hsprogress.Value
End Sub

Private Sub hsvolume_Change()                                           '��������� ��������� ��������
    wmp.settings.volume = hsvolume.Value                                    '������ ��������� ������
    lblVolume = "���������: " & hsvolume.Value & "%"                        '����� ������� �������� ���������
    WriteINIKey "General", "Volume", hsvolume.Value, presets                '���������� ��������� � ����
End Sub

Private Sub mnuSearch_Click()                                           '����� ����� � ������
    Dim tmp As String
    Dim i As Integer                                                    '���������� ��� �����
    tmp = InputBox("�������� ����� ��� ��� �����: ", , "")              '������ ����� �����
    If tmp = "" Then Exit Sub                                           '���� �������� ������ ��������, �������
    tmp = StrConv(tmp, vbLowerCase)                                     '������ ������ ������� �������� ������
    For i = 0 To File1.ListCount - 1                                    '���� ����
        If InStr(StrConv(File1.List(i), vbLowerCase), tmp) > 0 Then         '���� ������� ����������� �������� � ��������� �������� ������ ������
            File1.ListIndex = i                                             '������ ��������� �� ����
            Exit Sub                                                        '�������
        End If
    Next
    MsgBox "�� �������"                                                 '����� �� �������, ���� ������ �� �������
End Sub

Private Sub mnuShowPlayPart_Click()
    If mnuShowPlayPart.Checked = True Then
        mnuShowPlayPart.Checked = False
        WriteINIKey "View", "ShowPlayPart", "0", presets
        playPartWrapper.Visible = False
        lblTitle.Left = 8
        lblTitle.Top = 344
        lblTitle.Width = 193
        lblTitle.Height = 30
        lblTitle.Alignment = 0
        main.Height = 6480
        If repeating = True Then
            cmdApply.Value = 0
            cmdApply_Click
        End If
    Else
        mnuShowPlayPart.Checked = True
        WriteINIKey "View", "ShowPlayPart", "1", presets
        playPartWrapper.Visible = True
        lblTitle.Left = 168
        lblTitle.Top = 368
        lblTitle.Width = 137
        lblTitle.Height = 105
        lblTitle.Alignment = 1
        main.Height = 8000
    End If
End Sub

Private Sub mnuShowTrackInCaption_Click()
    If mnuShowTrackInCaption.Checked = False Then
        mnuShowTrackInCaption.Checked = True
        WriteINIKey "View", "ShowTrackInCaption", "1", presets
        showTrackInCaption = True
        If playing = True Then main.Caption = lblTitle.Caption
    Else
        mnuShowTrackInCaption.Checked = False
        WriteINIKey "View", "ShowTrackInCaption", "0", presets
        showTrackInCaption = False
        If playing = True Then main.Caption = "Stardisk's Player"
    End If
End Sub

Private Sub Timer1_Timer()                                              '������ ����������� ������� ��������������� � �������� ������� ���������
'On Error GoTo err
    lblTiming.Caption = wmp.Controls.currentPositionString & " / " & wmp.currentMedia.durationString ' ���������� ������� ��������� ��������������� � ����� �����
    hsprogress.Max = wmp.currentMedia.duration
    If hsprogress.Max <> 0 Then hsprogress.Value = CInt(wmp.Controls.currentPosition)             '������� �������
    
    If repeating = False And playing = True And wmp.playState = 1 Then '���� ���� ������� ��������, � ��������������� �������
        If chkRepeat.Value = 1 Then
            wmp.Controls.currentPosition = 0                    '���� ����� ����� "��������� ����������", �� ��������� ��������������� �� ������
            wmp.Controls.play
        Else
            cmdnext_Click                                                '� ����� ��� ��������� ��������������� ��-�� ��������� ����� �������� ������ "����.����"
        End If
    End If
    
    If repeating = True And playing = True Then                         '���� ����� ������� � ��������������� ��������
        If endtime <> 0 Then                                      '���� ����� ��������� �� ����� ����
            If wmp.Controls.currentPosition >= endtime Then wmp.Controls.currentPosition = starttime     '���� ��������������� ��������� ����� ���������, ������������� �� ������ ���������������
        Else                                                            '���� ����� ��������� ����� ����
            If wmp.Controls.currentPosition + 1 >= wmp.currentMedia.duration Then wmp.Controls.currentPosition = starttime '���� ������� ������� + 1 ������� ������ ����� �����, ������������� �� ������
        End If
    End If
    Exit Sub
err:
    lblTiming.Caption = "--.-- / --.--"                                 '����� ��������, ���� �� ������� �������� ����� ���������������
End Sub
