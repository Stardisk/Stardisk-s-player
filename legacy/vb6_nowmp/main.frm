VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stardisk's Player"
   ClientHeight    =   6015
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   5295
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
   ScaleHeight     =   401
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkShuffle 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Width           =   375
   End
   Begin VB.HScrollBar hsRate 
      Height          =   255
      LargeChange     =   10
      Left            =   3000
      Max             =   200
      Min             =   50
      TabIndex        =   12
      Top             =   5640
      Value           =   100
      Width           =   2175
   End
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
      TabIndex        =   8
      Top             =   4680
      Width           =   375
   End
   Begin VB.HScrollBar hsprogress 
      Height          =   255
      LargeChange     =   20
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   5055
   End
   Begin VB.HScrollBar hsvolume 
      Height          =   255
      LargeChange     =   10
      Left            =   3000
      Max             =   100
      TabIndex        =   11
      Top             =   5160
      Width           =   2175
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
      Left            =   120
      Top             =   3600
   End
   Begin VB.Frame playPartWrapper 
      Caption         =   "Играть кусок"
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CheckBox cmdApply 
         Caption         =   "Включить повтор"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1200
         Width           =   2055
      End
      Begin MSMask.MaskEdBox txtEndTime 
         Height          =   375
         Left            =   1200
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "До момента:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "С момента:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
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
      Pattern         =   "*.mp3;*.wma;*.ogg;*.wav"
      TabIndex        =   2
      Top             =   1800
      Width           =   5055
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label lblVolume 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   2565
      TabIndex        =   20
      Top             =   5205
      Width           =   240
   End
   Begin VB.Label lblRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "x1"
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
      Left            =   2640
      TabIndex        =   25
      Top             =   5700
      Width           =   150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   27.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2550
      TabIndex        =   24
      ToolTipText     =   "Громкость"
      Top             =   4995
      Width           =   300
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "0.00 МБ"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   4080
      Width           =   585
   End
   Begin VB.Label lblItems 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "000 / 000"
      Height          =   195
      Left            =   4440
      TabIndex        =   22
      Top             =   4080
      Width           =   690
   End
   Begin VB.Label lblBitrate 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 кбит/с"
      Height          =   195
      Left            =   2280
      TabIndex        =   21
      Top             =   4080
      Width           =   645
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
      Left            =   3675
      TabIndex        =   19
      Top             =   4680
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "`"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   27.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2430
      TabIndex        =   26
      ToolTipText     =   "Громкость"
      Top             =   5490
      Width           =   555
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Файл"
      Begin VB.Menu mnuRename 
         Caption         =   "Переименовать выделенный файл"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuMove 
         Caption         =   "Переместить в..."
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Удалить выделенный файл"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuOpenLocation 
         Caption         =   "Открыть расположение файла"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Найти файл в текущей папке"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuGoTo 
         Caption         =   "Перейти к адресу..."
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "О программе"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Выйти"
      End
   End
   Begin VB.Menu mnuSep0 
      Caption         =   "Вид"
      Begin VB.Menu mnuShowPlayPart 
         Caption         =   "Показать ""Играть кусок"""
      End
      Begin VB.Menu mnuShowTrackInCaption 
         Caption         =   "Название трека в заголовке"
      End
      Begin VB.Menu mnuMinimizeToCompact 
         Caption         =   "Сворачивать в компактный режи"
      End
      Begin VB.Menu mnuCloseToCompact 
         Caption         =   "Закрывать в компактный режим"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGotoCompact 
         Caption         =   "Перейти в компактный режим"
      End
      Begin VB.Menu mnuCompactVolume 
         Caption         =   "Шаг громкости в компактном режиме"
      End
      Begin VB.Menu mnuRestoreCompact 
         Caption         =   "Восстановить расположение компактного режима"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)     'для sleepов
Dim repeating, endOfTrack, showTrackInCaption As Boolean                                   'переменные для флагов воспроизведения и повтора
Dim starttime, endtime As Integer                                   'начальное и конечное время воспроизведения
Dim volume As Integer
Public presets As String                                            'путь к настройкам
Dim repeatlist As String                                            'путь к списку начал и концов повторяемых кусков
Dim tmplistindex As Integer                                         'переменная для запоминания номера выбранной строки при перемещении/удалении/переименовании файла
Dim tmpPosition As Variant                                          'переменная для запоминания на каком месте играло воспроизведение при переименовании файла
Dim currentPlayingFile As String

Public playingStatus As Byte ' 0 - stop; 1 - play; 2 - paused
Dim realCurrentPosition As Single
Dim rate As Single

Private Type IMediaType
    Control As IMediaControl
    Position As IMediaPosition
    Audio As IBasicAudio
    Event As IMediaEvent
End Type

' странные переменные для IMedia.Event.GetEvent
Dim evcode As Long
Dim param1 As LONG_PTR, param2 As LONG_PTR

Dim IMedia As IMediaType

Private Sub LoadMP3(filename$)
    On Error GoTo err
    resetFormWhileSwitchingTracks
    currentPlayingFile = filename
    Set IMedia.Control = New FilgraphManager
    Call IMedia.Control.RenderFile(filename)
    'Call IMedia.Control.RenderFile("http://orangepi.stardisk.xyz/radio?play")
    Set IMedia.Audio = IMedia.Control
    IMedia.Audio.volume = volume
    Set IMedia.Position = IMedia.Control
    Set IMedia.Event = IMedia.Control
    changeRate
    IMedia.Position.CurrentPosition = 0
    ActionsOnPlay
    Exit Sub
err:
    MsgBox "Ошибка [" & err.number & "]: " & err.Description
End Sub

Private Sub DestroyMP3()
    Set IMedia.Audio = Nothing
    Set IMedia.Position = Nothing
    Set IMedia.Control = Nothing
    Set IMedia.Event = Nothing
End Sub

Private Sub ButtonColor()                                           'задать цвет кнопок
    Select Case playingStatus
        Case 1: cmdPlay.BackColor = "&H00C0C0C0": cmdPause.BackColor = "&H8000000F": cmdStop.BackColor = "&H8000000F" ' подстветить воспроизведение
        Case 2: cmdPlay.BackColor = "&H8000000F": cmdPause.BackColor = "&H00C0C0C0": cmdStop.BackColor = "&H8000000F" ' подсветить паузу
        Case 0: cmdPlay.BackColor = "&H8000000F": cmdPause.BackColor = "&H8000000F": cmdStop.BackColor = "&H00C0C0C0"                      ' подсветить стоп
    End Select
End Sub

Private Function TimeToSeconds(mmss As String) As Integer                                         'преобразование времени мм:сс в секунды
    TimeToSeconds = CInt(Mid(mmss, 1, 2)) * 60 + CInt(Mid(mmss, 4, 2))
End Function

Private Function SecondsToTime(timeInSeconds As Variant) As String                                         'преобразование секунд в мм:сс
    SecondsToTime = addZeros(timeInSeconds \ 60, 2) & ":" & addZeros(timeInSeconds Mod 60, 2)
End Function

Private Function addZeros(number As Integer, requiredLength As Integer)
    Dim str As String
    str = CStr(number)
    Do While (Len(str) < requiredLength)
        str = "0" & str
    Loop
    addZeros = str
End Function

Private Sub changeCurrentPositionManually(toValue As Integer)
    hsprogress.Value = toValue
    IMedia.Position.CurrentPosition = toValue
    realCurrentPosition = toValue
    lblTiming.Caption = SecondsToTime(realCurrentPosition) & " / " & SecondsToTime(IMedia.Position.Duration)  ' показываем текущее положение воспроизведения и общую длину
End Sub

Private Sub resetFormWhileSwitchingTracks()
    Timer1.Enabled = False
    lblBitrate = ""                                          'убираем название и битрейт, которые могли остаться на форме от предыдущего трека
    cmdApply.Value = 0                                                      'выключить повтор
    DestroyMP3
End Sub

Private Sub ActionsOnPlay()                                                 'действия при воспроизведении
'On Error Resume Next
    endOfTrack = False
    IMedia.Control.Run
    lblItems = File1.ListIndex + 1 & " / " & File1.ListCount                'показываем номер текущего элемента и общее число элементов
    lblSize = Format(FileLen(currentPlayingFile) / 1048576, "0.00") & " МБ"   'показываем вес файла
    playingStatus = 1
    changeCurrentPositionManually (0)
    hsprogress.Max = IMedia.Position.Duration
    
    If showTrackInCaption = True Then main.Caption = File1.filename
    
    lblBitrate = "~" & Fix(FileLen(currentPlayingFile) / IMedia.Position.Duration / 1000 * 8) & " кбит/с"
    
    Dim loweredFilePath As String
    loweredFilePath = StrConv(currentPlayingFile, vbLowerCase)
    
    If ReadINIKey(loweredFilePath, "Start", repeatlist) <> "" Then     'если путь к песне есть в списке для повторов
        starttime = ReadINIKey(loweredFilePath, "Start", repeatlist)   'считываем начальное значение повтора
        endtime = ReadINIKey(loweredFilePath, "End", repeatlist)       'считываем окончательное значение повтора
        txtStartTime = SecondsToTime(starttime)
        txtEndTime = SecondsToTime(endtime)
        If repeating = True Then changeCurrentPositionManually (starttime)                              'если флаг повтора включен, воспроизводим повторение с момента повтора
    Else
        starttime = 0                                        'если пути к песне нет в списке повторов, то ставим все на 0
        endtime = 0
        txtStartTime = "00:00"
        txtEndTime = "00:00"
    End If
    Timer1.Enabled = True                                           'включаем таймер счетчика времени
    WriteINIKey "General", "LastIndex", File1.ListIndex, presets    'записываем последний воспроизведенный трек
    ButtonColor
End Sub

Private Sub Form_Resize()                                                   'переключение в компактный режим
    If Me.WindowState = 1 Then                                                  'если главное окно сворачивается
        If mnuMinimizeToCompact.Checked = True Then
            Me.Hide                                                                     'скрываем его
            compact.Show                                                                'показываем компактный режим
        End If
    End If
End Sub

Private Sub changeRate()
    rate = hsRate.Value / 100
    IMedia.Position.rate = rate
    lblRate = "x" & rate
    changeCurrentPositionManually (realCurrentPosition)
End Sub

Private Sub hsRate_Change()
    changeRate
End Sub

Private Sub mnuAbout_Click()                                        'меню "о программе"
    MsgBox App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
    "Автор: Александров Олег Игоревич aka Stardisk", vbInformation
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

Private Sub mnuCompactVolume_Click()                                'меню установки шага громкости для кнопок в компактном режиме
    Dim tmp As Variant
    tmp = ReadINIKey("General", "CompactStepVolume", presets)
    tmp = InputBox("Укажите, на сколько процентов будет изменяться уровень громкости при нажатии кнопок управления громкостью в компактном режиме:", , tmp)
    If tmp = "" Then Exit Sub
    If IsNumeric(tmp) = False Then MsgBox "Введите число": Exit Sub
    If tmp < 1 Or tmp > 100 Then MsgBox "Число должно быть от 1 до 100": Exit Sub
    
    WriteINIKey "General", "CompactStepVolume", CInt(tmp), presets
End Sub

Private Sub mnuGoTo_Click()
    On Error GoTo err
    Dim newPath As String
    
    newPath = InputBox("Введите путь к папке:", "", Dir1.Path)
    If newPath <> "" Then
        If GetAttr(newPath) And vbDirectory Then
            Dir1.Path = newPath
            Drive1.Drive = Left(newPath, 2)
        Else
            MsgBox "Not directory"
        End If
    End If
    Exit Sub
err:
    MsgBox "Неверный адрес"
End Sub

Private Sub mnuGotoCompact_Click()                                  'меню переключения в компактный режим
    Me.Hide
    compact.Show
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
    tmp = InputBox("Куда перемещаем?", "", ReadINIKey("General", "LastMovePath", presets))  'запрашиваем путь для перемещения
    If StrConv(tmp, vbLowerCase) = StrConv(Dir1.Path, vbLowerCase) Then MsgBox "Файл и так уже тут лежит.": Exit Sub 'если был введен тот же путь, что и исходный, пишем ошибку и выходим из модуля
    If tmp <> "" Then                                                   'если не была нажата отмена
        cmdstop_Click                                                   'останавливаем воспроизведение для освобождения файла
        resetFormWhileSwitchingTracks
        tmplistindex = File1.ListIndex                                  'запоминаем номер воспроизводящейся песни
        tmpfilesize = FileLen(Dir1.Path & "\" & File1.filename)         'записываем размер файла
        On Error GoTo err
            Name Dir1.Path & "\" & File1.filename As tmp & "\" & File1.filename 'перемещаем файл
        On Error GoTo 0
        While FileLen(tmp & "\" & File1.filename) <> tmpfilesize        'проверяем, переместился ли файл или нет
            Sleep (1000)                                                'если еще не докопировался, ждем секунду
        Wend
        WriteINIKey "General", "LastMovePath", tmp, presets             'записываем последний использованный путь для перемещения
        File1.Refresh                                                   'обновляем список файлов
        If tmplistindex > File1.ListCount - 1 Then File1.ListIndex = tmplistindex - 1 Else File1.ListIndex = tmplistindex ' если файл был последним в списке, переходим на предыдущий, если не был последним, выбираем тот, который стал теперь этим номером
    End If
    Exit Sub
err:
    MsgBox "Не удалось переместить файл."
End Sub

Private Sub mnuOpenLocation_Click()
    Shell "explorer.exe /select, " & currentPlayingFile, vbNormalFocus
End Sub

Private Sub mnuRename_Click()                                       'меню переименования файла
On Error Resume Next
    If File1.ListIndex = -1 Then Exit Sub                           'если никакой файл не выбран, выходим
    Dim newName As String                                           'спрашиваем у пользователя новое имя файла
    newName = InputBox("Новое имя для файла " & File1.List(File1.ListIndex), , File1.List(File1.ListIndex))
    If newName = "" Then Exit Sub                                   'выходим, если нажата отмена

    tmplistindex = File1.ListIndex                                  'запоминаем номер воспроизводящейся песни
    cmdstop_Click                                                   'останавливаем воспроизведение, чтобы файл не использовался
    Name Dir1.Path & "\" & File1.List(File1.ListIndex) As Dir1.Path & "\" & newName 'переименовываем
    File1.Refresh                                                   'обновляем список файлов
    File1.ListIndex = tmplistindex                                  'переходим на запомненный номер
End Sub

Public Sub cmdstop_Click()                                          'кнопка остановки воспроизведения
    IMedia.Control.Stop                                               'останавливаем воспроизведение
    changeCurrentPositionManually (0)
    playingStatus = 0
    ButtonColor                                                     'меняем кнопкам цвет
    If showTrackInCaption = True Then main.Caption = App.Title
End Sub

Private Sub cmdApply_Click()                                        'кнопка повтора
On Error GoTo err
If File1.ListIndex < 0 Then cmdApply.Value = 0: Exit Sub            'если файл не выбран, выходим
If cmdApply.Value = 1 Then                                          'если кнопка нажата
    starttime = TimeToSeconds(txtStartTime)
    endtime = TimeToSeconds(txtEndTime)
    If starttime >= endtime Then
        MsgBox "Время начала должно быть меньше времени окончания"
        cmdApply.Value = 0
    Else
        repeating = True
        If starttime <> 0 And endtime <> 0 Then changeCurrentPositionManually (starttime)
        txtStartTime.Enabled = False
        txtEndTime.Enabled = False
        
        If starttime > 0 Or endtime > 0 Then
            WriteINIKey StrConv(currentPlayingFile, vbLowerCase), "Start", CStr(starttime), repeatlist
            WriteINIKey StrConv(currentPlayingFile, vbLowerCase), "End", CStr(endtime), repeatlist
        End If
    End If
    Exit Sub
Else
    repeating = False
    txtStartTime.Enabled = True
    txtEndTime.Enabled = True
    Exit Sub
End If
err:
    MsgBox "Неверные данные."
    cmdApply.Value = 0
End Sub

Public Sub cmdpause_Click()                                     'кнопка паузы
    IMedia.Control.Pause                                          'приостанавливаем воспроизведение
    playingStatus = 2
    ButtonColor                                                 'изменяем цвет кнопок
End Sub

Public Sub cmdplay_Click()                                      'кнопка вопроизведения
    If repeating = True Then changeCurrentPositionManually (starttime)
    IMedia.Control.Run
    playingStatus = 1
    ButtonColor
    If showTrackInCaption = True Then main.Caption = File1.filename
End Sub

Public Sub cmdPrev_Click()
    If File1.ListIndex <= 0 Then File1.ListIndex = File1.ListCount - 1 Else File1.ListIndex = File1.ListIndex - 1
End Sub

Public Sub cmdnext_Click()
    On Error GoTo err
    If chkShuffle.Value = 1 Then File1.ListIndex = Int((Rnd * (File1.ListCount - 1))) Else:  File1.ListIndex = File1.ListIndex + 1
    Exit Sub
err:
    If File1.ListCount > 0 Then File1.ListIndex = 0
End Sub

Private Sub mnuClose_Click()
    End
End Sub

Private Sub mnuDelete_Click()                                           'меню удаления файла
'On Error Resume Next
    If File1.ListIndex > -1 Then                                        'если файл выбран
        If MsgBox("Удалить " & File1.filename & "?", vbYesNo + vbQuestion) = vbYes Then 'задаем вопрос про удаление
            cmdstop_Click                                               'останавливаем воспроизвение
            resetFormWhileSwitchingTracks
            DestroyMP3
            tmplistindex = File1.ListIndex                              'записываем номер этого трека
            Me.Enabled = False                                          'отключаем форму, дабы пользователь ничо не кликал
            Kill currentPlayingFile                       'удаляем файл
            Sleep (500)                                                 'если не удалился, ждем 1000 мс
            Me.Enabled = True                                           'включаем форму
            File1.Refresh                                               'обновляем список
            If tmplistindex > File1.ListCount - 1 Then File1.ListIndex = tmplistindex - 1 Else File1.ListIndex = tmplistindex ' если файл был последним в списке, переходим на предыдущий, если не был последним, выбираем тот, который стал теперь этим номером
        End If
    End If
End Sub

Private Sub Dir1_Change()                                               'при переходе по папкам в окне программы
    File1.Path = Dir1.Path                                              'показываем файлы в текущей папке
    WriteINIKey "General", "LastPath", Dir1.Path, presets               'записываем путь к последней выбранной папке в настройки
End Sub

Private Sub Drive1_Change()                                             'смена диска в окне программв
On Error GoTo err
    Dir1.Path = Drive1.Drive                                            'показываем папки на диске
    Exit Sub
err:
    MsgBox "Невозможно прочитать информацию. Возможно, диск отсутствует или поврежден."
    Drive1.Drive = Left(App.Path, 2)
End Sub

Private Sub File1_Click()                                               'клик по файлу в окне программы
    If File1.ListIndex > -1 And File1.ListIndex < File1.ListCount Then LoadMP3 (Dir1.Path & "\" & File1.filename)
End Sub

Private Sub loadDefaults()
    WriteINIKey "General", "LastPath", App.Path, presets
    WriteINIKey "General", "CompactStepVolume", "5", presets
    
    hsvolume.Value = 100
    mnuShowPlayPart_Click
    
    MsgBox 1
End Sub

Private Sub Form_Load()                                                 'действия при запуске
On Error Resume Next
    Dim prevPath As String
    If App.PrevInstance = True Then End                                 'если уже запущен один плеер, закрываем второй
    presets = App.Path & "\presets.ini"                                 'записываем путь к настройкам в переменную
    repeatlist = App.Path & "\repeatlist.ini"                           'записываем путь к списку повторялок в переменную
    If Dir(presets) = "" Then loadDefaults
    prevPath = ReadINIKey("General", "LastPath", presets)
    Drive1.Drive = Left(prevPath, 2)
    Dir1.Path = prevPath              'открываем папку, в которой были последний раз
    File1.ListIndex = ReadINIKey("General", "LastIndex", presets)       'выбираем файл, на котором остановились в последний раз
    hsvolume.Value = ReadINIKey("General", "Volume", presets)           'устанавливаем бегунку громкости значение из настроек
    changeRate
    
    If ReadINIKey("View", "ShowPlayPart", presets) = "1" Then mnuShowPlayPart_Click
    If ReadINIKey("View", "ShowTrackInCaption", presets) = "1" Then mnuShowTrackInCaption_Click
    If ReadINIKey("View", "MinimizeToCompact", presets) = "1" Then mnuMinimizeToCompact_Click
    If ReadINIKey("View", "CloseToCompact", presets) = "1" Then mnuCloseToCompact_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)                              'если главное окно программы закрывают
    If mnuCloseToCompact.Checked = True Then
        Cancel = 1                                                          'отмена завершения программы
        Me.Hide                                                             'прячем форму
        compact.Show                                                        'показываем компактный режим
    End If
End Sub

Private Sub hsprogress_scroll()                                         'изменение позиции воспроизведения движением бегунка
    changeCurrentPositionManually (hsprogress.Value)
End Sub

Private Sub changeVolume()
    Dim preVolume As Integer
    preVolume = -10000
    If hsvolume.Value > 0 Then preVolume = -2000 + ((Log(hsvolume.Value) / Log(10) - 1) * 2000)                           'задаем громкость плееру
    If IMedia.Audio Is Nothing <> True Then IMedia.Audio.volume = preVolume
    
    volume = preVolume
    lblVolume = hsvolume.Value                                            'пишем текущее значение громкости
    WriteINIKey "General", "Volume", hsvolume.Value, presets                'записываем громкость в файл
End Sub

Private Sub hsvolume_Change()                                           'изменение громкости бегунком
    changeVolume
End Sub
Private Sub hsvolume_Scroll()                                           'изменение громкости бегунком
    changeVolume
End Sub

Private Sub mnuRestoreCompact_Click()
    WriteINIKey "CompactPosition", "X", 0, main.presets
    WriteINIKey "CompactPosition", "Y", 0, main.presets
    mnuGotoCompact_Click
    MsgBox "Компактный режим находится в верхнем левом углу.", vbInformation
End Sub

Private Sub mnuSearch_Click()                                           'поиск файла в списке
    Dim tmp As String
    Dim i As Integer                                                    'переменная для цикла
    tmp = InputBox("Название файла или его часть: ", , "")              'запрос имени файла
    If tmp = "" Then Exit Sub                                           'если получено пустое значение, выходим
    tmp = StrConv(tmp, vbLowerCase)                                     'делаем нижний регистр значению поиска
    For i = 0 To File1.ListCount - 1                                    'ищем файл
        If InStr(StrConv(File1.List(i), vbLowerCase), tmp) > 0 Then         'если находим запрошенное значение в очередном элементе списка файлов
            File1.ListIndex = i                                             'ставим выделение на него
            Exit Sub                                                        'выходим
        End If
    Next
    MsgBox "Не найдено"                                                 'пишем не найдено, если ничего не нашлось
End Sub

Private Sub mnuShowPlayPart_Click()
    If mnuShowPlayPart.Checked = True Then
        mnuShowPlayPart.Checked = False
        WriteINIKey "View", "ShowPlayPart", "0", presets
        playPartWrapper.Visible = False
        main.Height = 6795
        If repeating = True Then
            cmdApply.Value = 0
            cmdApply_Click
        End If
    Else
        mnuShowPlayPart.Checked = True
        WriteINIKey "View", "ShowPlayPart", "1", presets
        playPartWrapper.Visible = True
        main.Height = 7755
    End If
End Sub

Private Sub mnuShowTrackInCaption_Click()
    If mnuShowTrackInCaption.Checked = False Then
        mnuShowTrackInCaption.Checked = True
        WriteINIKey "View", "ShowTrackInCaption", "1", presets
        showTrackInCaption = True
        If playingStatus <> 0 Then main.Caption = File1.filename
    Else
        mnuShowTrackInCaption.Checked = False
        WriteINIKey "View", "ShowTrackInCaption", "0", presets
        showTrackInCaption = False
        main.Caption = App.Title
    End If
End Sub

Private Sub Timer1_Timer()                                              'таймер отображения времени воспроизведения и движения бегунка прогресса
'On Error GoTo err
    ' этот код нужен лишь для определения конца воспроизведения... если evcode после цикла равен 1, значит достигнут конец файла
    evcode = 999
    Do While evcode <> 0
        On Error Resume Next
        IMedia.Event.GetEvent evcode&, param1, param2, 0
        If evcode = 1 Then endOfTrack = True
    Loop
    IMedia.Event.FreeEventParams evcode&, param1, param2
    
    If repeating = False And playingStatus = 1 And endOfTrack = True Then 'если флаг повтора отключен, а воспроизведения включен
        If chkRepeat.Value = 1 Then
            changeCurrentPositionManually (0)
            IMedia.Control.Run
        Else
            cmdnext_Click                      'а иначе при остановке воспроизведения из-за окончания трека нажимаем кнопку "след.трек"
        End If
    Else
        If playingStatus = 1 Then realCurrentPosition = realCurrentPosition + rate
        lblTiming.Caption = SecondsToTime(realCurrentPosition) & " / " & SecondsToTime(IMedia.Position.Duration)  ' показываем текущее положение воспроизведения и общую длину
        If hsprogress.Max <> 0 Then hsprogress.Value = CInt(realCurrentPosition)             'двигаем бегунок
    End If
    
    If repeating = True And playingStatus = 1 Then                         'если флаги повтора и воспроизведения включены
        If realCurrentPosition < starttime Then
            changeCurrentPositionManually (starttime)
        Else
            If endtime <> 0 Then                                      'если время окончания не равно нулю
                If realCurrentPosition >= endtime Then changeCurrentPositionManually (starttime)      'если воспроизведение превысило время окончания, перескакиваем на начало воспроизведения
            Else                                                            'если время окончания равно нулю
                If endOfTrack = True Then changeCurrentPositionManually (starttime): endOfTrack = False
            End If
        End If
    End If
    
    Exit Sub
err:
    lblTiming.Caption = "--.-- / --.--"                                 'пишем черточки, если не удается получить время воспроизведения
End Sub
