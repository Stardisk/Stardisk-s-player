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
      Caption         =   "Играть кусок"
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CheckBox cmdApply 
         Caption         =   "Включить повтор"
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
         Caption         =   "Повтор выключен"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1620
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "До момента:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "С момента:"
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
      Caption         =   "0.00 МБ"
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
      Caption         =   "0 кбит/с"
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
      Caption         =   "Нажмите здесь для вывода меню"
      WindowList      =   -1  'True
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
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
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
      Begin VB.Menu mnuAbout 
         Caption         =   "О программе"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Выйти"
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
Dim repeating, playing, showTrackInCaption As Boolean                                   'переменные для флагов воспроизведения и повтора
Dim starttime, endtime As Integer                                   'начальное и конечное время воспроизведения
Public presets As String                                            'путь к настройкам
Dim repeatlist As String                                            'путь к списку начал и концов повторяемых кусков
Dim tmplistindex As Integer                                         'переменная для запоминания номера выбранной строки при перемещении/удалении/переименовании файла
Dim tmpPosition As Variant                                          'переменная для запоминания на каком месте играло воспроизведение при переименовании файла

Private Sub ButtonColor()                                           'задать цвет кнопок
    If wmp.playState = 3 Or wmp.playState = 9 Then cmdPlay.BackColor = "&H00C0C0C0": cmdPause.BackColor = "&H8000000F": cmdStop.BackColor = "&H8000000F" ' подстветить воспроизведение
    If wmp.playState = 2 Then cmdPlay.BackColor = "&H8000000F": cmdPause.BackColor = "&H00C0C0C0": cmdStop.BackColor = "&H8000000F"                      ' подсветить паузу
    If wmp.playState < 2 Or wmp.playState = 8 Then cmdPlay.BackColor = "&H8000000F": cmdPause.BackColor = "&H8000000F": cmdStop.BackColor = "&H00C0C0C0"                      ' подсветить стоп
End Sub

Private Sub TimeToSeconds()                                         'преобразование времени мм:сс в секунды
    starttime = CInt(Mid(txtStartTime, 1, 2)) * 60 + CInt(Mid(txtStartTime, 4, 2))
    endtime = CInt(Mid(txtEndTime, 1, 2)) * 60 + CInt(Mid(txtEndTime, 4, 2))
End Sub

Private Sub SecondsToTime()                                         'преобразование секунд в мм:сс
    txtStartTime = Right(TimeSerial(0, 0, starttime), 5)
    txtEndTime = Right(TimeSerial(0, 0, endtime), 5)
End Sub

Private Sub ActionsOnPlay()                                                 'действия при воспроизведении
On Error Resume Next
    lblTitle = "": lblBitrate = ""                                          'убираем название и битрейт, которые могли остаться на форме от предыдущего трека
    TimeToSeconds                                                           'переводим мм:сс в секунды
    cmdApply.Value = 0                                                      'выключить повтор
    wmp.URL = Dir1.Path & "\" & File1.filename                              'сообщаем плееру путь к выбранному файлу
    lblItems = File1.ListIndex + 1 & " / " & File1.ListCount                'показываем номер текущего элемента и общее число элементов
    lblSize = Format(FileLen(Dir1.Path & "\" & File1.filename) / 1048576, "0.00") & " МБ"   'показываем вес файла
    playing = True                                                          'включаем флаг воспроизведения
    hsprogress.Value = 0

    If wmp.currentMedia.getItemInfo("Author") <> "" Then                                                    'если у песни тег автор не пустой
        lblTitle = wmp.currentMedia.getItemInfo("Title") & " - " & wmp.currentMedia.getItemInfo("Author")       'пишем название песни и автора
    Else                                                                                                    'иначе
        lblTitle = wmp.currentMedia.getItemInfo("Title")                                                        'пишем только название песни, при пустом теге туда подставляется имя файла
    End If
    
    If showTrackInCaption = True Then main.Caption = lblTitle.Caption
    
    lblBitrate = Format(wmp.currentMedia.getItemInfo("bitrate") / 1000, "#") & " кбит/с"
    
    If ReadINIKey(StrConv(wmp.currentMedia.sourceURL, vbLowerCase), "Start", repeatlist) <> "" Then     'если путь к песне есть в списке для повторов
        starttime = ReadINIKey(StrConv(wmp.currentMedia.sourceURL, vbLowerCase), "Start", repeatlist)   'считываем начальное значение повтора
        endtime = ReadINIKey(StrConv(wmp.currentMedia.sourceURL, vbLowerCase), "End", repeatlist)       'считываем окончательное значение повтора
        SecondsToTime
        If repeating = True Then wmp.Controls.currentPosition = starttime                               'если флаг повтора включен, воспроизводим повторение с момента повтора
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

Private Sub mnuAbout_Click()                                        'меню "о программе"
    MsgBox "Stardisk's Player v." & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
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

Private Sub mnuGotoCompact_Click()                                  'меню переключения в компактный режим
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
    tmp = InputBox("Куда перемещаем?", "", ReadINIKey("General", "LastMovePath", presets))  'запрашиваем путь для перемещения
    If StrConv(tmp, vbLowerCase) = StrConv(Dir1.Path, vbLowerCase) Then MsgBox "Файл и так уже тут лежит.": Exit Sub 'если был введен тот же путь, что и исходный, пишем ошибку и выходим из модуля
    If tmp <> "" Then                                                   'если не была нажата отмена
        cmdstop_Click                                                   'останавливаем воспроизведение для освобождения файла
        wmp.URL = ""                                                    'убираем адрес из вмп
        Timer1.Enabled = False                                          'останавливаем таймер
        tmplistindex = File1.ListIndex                                  'запоминаем номер воспроизводящейся песни
        tmpPosition = wmp.Controls.currentPosition                      'запоминаем место воспроизведения
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
        Timer1.Enabled = True
    End If
    Exit Sub
err:
    MsgBox "Не удалось переместить файл."
End Sub

Private Sub mnuOpenLocation_Click()
    Shell "explorer.exe /select, " & wmp.URL, vbNormalFocus
End Sub

Private Sub mnuRename_Click()                                       'меню переименования файла
On Error Resume Next
    If File1.ListIndex = -1 Then Exit Sub                           'если никакой файл не выбран, выходим
    Dim newName As String                                           'спрашиваем у пользователя новое имя файла
    newName = InputBox("Новое имя для файла " & File1.List(File1.ListIndex), , File1.List(File1.ListIndex))
    If newName = "" Then Exit Sub                                   'выходим, если нажата отмена

    tmplistindex = File1.ListIndex                                  'запоминаем номер воспроизводящейся песни
    tmpPosition = wmp.Controls.currentPosition                      'запоминаем место воспроизведения
    cmdstop_Click                                                   'останавливаем воспроизведение, чтобы файл не использовался
    Name Dir1.Path & "\" & File1.List(File1.ListIndex) As Dir1.Path & "\" & newName 'переименовываем
    File1.Refresh                                                   'обновляем список файлов
    File1.ListIndex = tmplistindex                                  'переходим на запомненный номер
    wmp.URL = Dir1.Path & "\" & newName                             'сообщаем плееру новое имя переименованного файла
    wmp.Controls.currentPosition = tmpPosition                      'запускаем воспроизведение с того места, на котором он был переименован
    If playing = True Then wmp.Controls.play                        'воспроизводим
End Sub

Public Sub cmdstop_Click()                                          'кнопка остановки воспроизведения
    wmp.Controls.stop                                               'останавливаем воспроизведение
    playing = False                                                 'убираем флаг воспроизведения
    ButtonColor                                                     'меняем кнопкам цвет
    If showTrackInCaption = True Then main.Caption = "Stardisk's Player"
End Sub

Private Sub cmdApply_Click()                                        'кнопка повтора
On Error GoTo err
If File1.ListIndex < 0 Then cmdApply.Value = 0: Exit Sub            'если файл не выбран, выходим
If cmdApply.Value = 1 Then                                          'если кнопка нажата
    TimeToSeconds
    repeating = True
    If starttime <> 0 And endtime <> 0 Then wmp.Controls.currentPosition = starttime
    Label5.Caption = "Повтор включен"
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
    Label5.Caption = "Повтор выключен"
    Label5.FontBold = False
    txtStartTime.Enabled = True
    txtEndTime.Enabled = True
    Exit Sub
End If
err:
    MsgBox "Неверные данные."
    cmdApply.Value = 0
End Sub

Public Sub cmdpause_Click()                                     'кнопка паузы
    wmp.Controls.pause                                          'приостанавливаем воспроизведение
    ButtonColor                                                 'изменяем цвет кнопок
End Sub

Public Sub cmdplay_Click()                                      'кнопка вопроизведения
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

Private Sub mnuDelete_Click()                                           'меню удаления файла
'On Error Resume Next
    If File1.ListIndex > -1 Then                                        'если файл выбран
        If MsgBox("Удалить " & File1.filename & "?", vbYesNo + vbQuestion) = vbYes Then 'задаем вопрос про удаление
            cmdstop_Click                                               'останавливаем воспроизвение
            Timer1.Enabled = False                                      'останавливаем таймер
            wmp.URL = ""
            lblTitle = "Удаление файла..."                              'пишем что файл удаляется
            tmplistindex = File1.ListIndex                              'записываем номер этого трека
            Me.Enabled = False                                          'отключаем форму, дабы пользователь ничо не кликал
            mnuMain.Enabled = False                                     'отключаем меню, шоб не кликали
            Kill Dir1.Path & "\" & File1.filename                       'удаляем файл
            While Len(Dir(Dir1.Path & "\" & File1.filename)) > 0        'проверяем, удалился ли файл или нет
                Sleep (200)                                                 'если не удалился, ждем 1000 мс
                DoEvents
            Wend
            Me.Enabled = True                                           'включаем форму
            mnuMain.Enabled = True                                      'включаем меню
            File1.Refresh                                               'обновляем список
            If tmplistindex > File1.ListCount - 1 Then File1.ListIndex = tmplistindex - 1 Else File1.ListIndex = tmplistindex ' если файл был последним в списке, переходим на предыдущий, если не был последним, выбираем тот, который стал теперь этим номером
            Timer1.Enabled = True
        End If
    End If
End Sub

Private Sub Command8_Click()                                            'кнопка ручного ввода пути
    On Error Resume Next
    Dir1.Path = InputBox("Введите путь к папке:", "", Dir1.Path)
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
    If File1.ListIndex > -1 And File1.ListIndex < File1.ListCount Then ActionsOnPlay 'если номер элемента списка корректный, то выполняем действия при воспроизведении - загрузка файла и т.п.
End Sub

Private Sub Form_Load()                                                 'действия при запуске
On Error Resume Next
    If App.PrevInstance = True Then End                                 'если уже запущен один плеер, закрываем второй
    presets = App.Path & "\presets.ini"                                 'записываем путь к настройкам в переменную
    repeatlist = App.Path & "\repeatlist.ini"                           'записываем путь к списку повторялок в переменную
    txtStartTime = "00:00"                                              'пишем нолики в повтор
    txtEndTime = "00:00"
    Dir1.Path = ReadINIKey("General", "LastPath", presets)              'открываем папку, в которой были последний раз
    File1.ListIndex = ReadINIKey("General", "LastIndex", presets)       'выбираем файл, на котором остановились в последний раз
    hsvolume.Value = ReadINIKey("General", "Volume", presets)           'устанавливаем бегунку громкости значение из настроек
    wmp.settings.volume = hsvolume.Value                                'устанавливаем соответствующую громкость
    If ReadINIKey("General", "CompactStepVolume", presets) = "" Then WriteINIKey "General", "CompactStepVolume", "5", presets 'если в настройках не указан шаг громкости кнопкам компактного режима, устанавливаем по умолчанию 5
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

Private Sub hsprogress_Scroll()                                         'изменение позиции воспроизведения движением бегунка
    wmp.Controls.currentPosition = hsprogress.Value
End Sub

Private Sub hsvolume_Change()                                           'изменение громкости бегунком
    wmp.settings.volume = hsvolume.Value                                    'задаем громкость плееру
    lblVolume = "Громкость: " & hsvolume.Value & "%"                        'пишем текущее значение громкости
    WriteINIKey "General", "Volume", hsvolume.Value, presets                'записываем громкость в файл
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

Private Sub Timer1_Timer()                                              'таймер отображения времени воспроизведения и движения бегунка прогресса
'On Error GoTo err
    lblTiming.Caption = wmp.Controls.currentPositionString & " / " & wmp.currentMedia.durationString ' показываем текущее положение воспроизведения и общую длину
    hsprogress.Max = wmp.currentMedia.duration
    If hsprogress.Max <> 0 Then hsprogress.Value = CInt(wmp.Controls.currentPosition)             'двигаем бегунок
    
    If repeating = False And playing = True And wmp.playState = 1 Then 'если флаг повтора отключен, а воспроизведения включен
        If chkRepeat.Value = 1 Then
            wmp.Controls.currentPosition = 0                    'если стоит галка "повторять композицию", то переводим воспроизведение на начало
            wmp.Controls.play
        Else
            cmdnext_Click                                                'а иначе при остановке воспроизведения из-за окончания трека нажимаем кнопку "след.трек"
        End If
    End If
    
    If repeating = True And playing = True Then                         'если флаги повтора и воспроизведения включены
        If endtime <> 0 Then                                      'если время окончания не равно нулю
            If wmp.Controls.currentPosition >= endtime Then wmp.Controls.currentPosition = starttime     'если воспроизведение превысило время окончания, перескакиваем на начало воспроизведения
        Else                                                            'если время окончания равно нулю
            If wmp.Controls.currentPosition + 1 >= wmp.currentMedia.duration Then wmp.Controls.currentPosition = starttime 'если текущая позиция + 1 секунда больше длины файла, перескакиваем на начало
        End If
    End If
    Exit Sub
err:
    lblTiming.Caption = "--.-- / --.--"                                 'пишем черточки, если не удается получить время воспроизведения
End Sub
