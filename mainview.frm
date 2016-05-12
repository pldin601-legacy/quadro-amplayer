VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "SysTray.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8475
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "mainview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command20 
      Caption         =   "S"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   38
      ToolTipText     =   "Возобновить таймер"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command14 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   37
      ToolTipText     =   "Выключить таймер (мало ресурсов)"
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox BackGround 
      BackColor       =   &H00808080&
      Height          =   5895
      Left            =   0
      ScaleHeight     =   5835
      ScaleWidth      =   8415
      TabIndex        =   4
      Top             =   240
      Width           =   8475
      Begin VB.Frame Frame3 
         BackColor       =   &H00808080&
         Caption         =   "Karaoke Extenstions"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   3300
         TabIndex        =   51
         Top             =   3240
         Width           =   4935
         Begin VB.CommandButton Command24 
            Caption         =   "Рекламма"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3780
            TabIndex        =   55
            Top             =   240
            Width           =   1035
         End
         Begin VB.CommandButton Command23 
            Caption         =   "Фу!!!"
            Height          =   255
            Left            =   2640
            TabIndex        =   54
            Top             =   240
            Width           =   1155
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Овации"
            Height          =   255
            Left            =   1740
            TabIndex        =   53
            Top             =   240
            Width           =   915
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Аплодисменты"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   1635
         End
      End
      Begin VB.Timer Timer2 
         Interval        =   5000
         Left            =   3480
         Top             =   3960
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00000000&
         Height          =   675
         Left            =   180
         ScaleHeight     =   615
         ScaleWidth      =   2895
         TabIndex        =   44
         Top             =   2580
         Width           =   2955
         Begin Project1.TrackAll TrackAll2 
            Height          =   135
            Left            =   1920
            TabIndex        =   50
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   238
         End
         Begin Project1.TrackAll TrackAll1 
            Height          =   135
            Left            =   1920
            TabIndex        =   49
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   238
         End
         Begin Project1.MTimer MTimer6 
            Height          =   330
            Left            =   180
            TabIndex        =   46
            Top             =   240
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stop:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1440
            TabIndex        =   48
            Top             =   300
            Width           =   375
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Start:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1440
            TabIndex        =   47
            Top             =   60
            Width           =   375
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Таймер"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Left            =   180
            TabIndex        =   45
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         Height          =   675
         Left            =   180
         ScaleHeight     =   615
         ScaleWidth      =   2895
         TabIndex        =   39
         Top             =   3240
         Width           =   2955
         Begin Project1.MTimer MTimer5 
            Height          =   330
            Left            =   1500
            TabIndex        =   40
            Top             =   240
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
         End
         Begin Project1.MTimer MTimer4 
            Height          =   330
            Left            =   180
            TabIndex        =   41
            Top             =   240
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Стартовое время"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Left            =   180
            TabIndex        =   43
            Top             =   0
            Width           =   1125
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Окончание"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Left            =   1500
            TabIndex        =   42
            Top             =   0
            Width           =   720
         End
      End
      Begin SysTrayCtl.cSysTray cSysTray1 
         Left            =   2640
         Top             =   1020
         _ExtentX        =   900
         _ExtentY        =   900
         InTray          =   0   'False
         TrayIcon        =   "mainview.frx":0442
         TrayTip         =   "MegaBox"
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   20000
         Left            =   60
         Top             =   4020
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   3960
         Width           =   3495
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00808080&
         Caption         =   "Воспроизведение"
         Height          =   1215
         Left            =   120
         TabIndex        =   28
         Top             =   4260
         Width           =   3495
         Begin VB.CommandButton Command4 
            Caption         =   "Старт"
            CausesValidation=   0   'False
            Height          =   435
            Left            =   180
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Воспоизведение"
            Top             =   660
            UseMaskColor    =   -1  'True
            Width           =   1515
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Пауза"
            Height          =   375
            Left            =   1740
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   660
            Width           =   795
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Стоп"
            Height          =   375
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   660
            Width           =   795
         End
         Begin VB.CommandButton Command7 
            Caption         =   "<< Предыдущая"
            Height          =   375
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   300
            Width           =   1575
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Следующая >>"
            Height          =   375
            Left            =   1740
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00808080&
         Caption         =   "Кнопки"
         Height          =   1635
         Left            =   3900
         TabIndex        =   17
         Top             =   3840
         Width           =   4335
         Begin VB.CommandButton Command9 
            Caption         =   "Настройка"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Громкость"
            Height          =   375
            Left            =   1200
            TabIndex        =   26
            Top             =   240
            Width           =   1155
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Поиск в Explorrer"
            Height          =   375
            Left            =   2340
            TabIndex        =   25
            Top             =   240
            Width           =   1875
         End
         Begin VB.CommandButton Command12 
            Caption         =   "ReRegister"
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton Command13 
            Caption         =   "О программе"
            Height          =   375
            Left            =   1200
            TabIndex        =   23
            Top             =   600
            Width           =   1155
         End
         Begin VB.CommandButton Command15 
            Caption         =   "Добавить файлы"
            Height          =   375
            Left            =   2340
            TabIndex        =   22
            Top             =   600
            Width           =   1875
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Сохранить список"
            Height          =   555
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   1095
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Загрузить список"
            Height          =   555
            Left            =   1200
            TabIndex        =   20
            Top             =   960
            Width           =   1155
         End
         Begin VB.CommandButton Command18 
            Caption         =   "Удалить"
            Height          =   555
            Left            =   2340
            TabIndex        =   19
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton Command19 
            Caption         =   "Очистить"
            Height          =   555
            Left            =   3300
            TabIndex        =   18
            Top             =   960
            Width           =   915
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   1575
         Left            =   180
         ScaleHeight     =   1515
         ScaleWidth      =   2895
         TabIndex        =   8
         Top             =   1020
         Width           =   2955
         Begin Project1.MTrack MTrack1 
            Height          =   315
            Left            =   180
            TabIndex        =   9
            Top             =   1020
            Width           =   915
            _ExtentX        =   1826
            _ExtentY        =   556
         End
         Begin Project1.MTimer MTimer3 
            Height          =   330
            Left            =   1500
            TabIndex        =   10
            Top             =   1020
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
         End
         Begin Project1.MTimer MTimer2 
            Height          =   330
            Left            =   1500
            TabIndex        =   11
            Top             =   360
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
         End
         Begin Project1.MTimer MTimer1 
            Height          =   330
            Left            =   180
            TabIndex        =   12
            Top             =   360
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   582
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Позиция"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Left            =   180
            TabIndex        =   16
            Top             =   120
            Width           =   585
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Длительность"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Left            =   1500
            TabIndex        =   15
            Top             =   120
            Width           =   915
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Осталось времени"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Left            =   1500
            TabIndex        =   14
            Top             =   780
            Width           =   1200
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Песня #"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Left            =   180
            TabIndex        =   13
            Top             =   780
            Width           =   525
         End
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   2400
         Left            =   3300
         OLEDropMode     =   1  'Manual
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   120
         Width           =   4935
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   555
         Left            =   3300
         TabIndex        =   5
         Top             =   2640
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   979
         _Version        =   327682
         BorderStyle     =   1
         TickFrequency   =   30
      End
      Begin MCI.MMControl MMControl1 
         Height          =   315
         Left            =   3540
         TabIndex        =   7
         Top             =   4440
         Visible         =   0   'False
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   556
         _Version        =   393216
         PrevVisible     =   0   'False
         NextVisible     =   0   'False
         PlayVisible     =   0   'False
         PauseVisible    =   0   'False
         BackVisible     =   0   'False
         StepVisible     =   0   'False
         StopVisible     =   0   'False
         RecordVisible   =   0   'False
         EjectVisible    =   0   'False
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin MSComDlg.CommonDialog COD 
         Left            =   3540
         Top             =   4980
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   "Списки MegaBox 1.01 (*.mbx)|*.mbx"
      End
      Begin VB.Image Image1 
         Height          =   915
         Left            =   180
         Picture         =   "mainview.frx":0894
         Stretch         =   -1  'True
         Top             =   60
         Width           =   2925
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   34
         Top             =   5580
         Width           =   8475
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7980
      TabIndex        =   3
      ToolTipText     =   "Показать в SysTray"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7740
      TabIndex        =   2
      ToolTipText     =   "Свернуть"
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8220
      TabIndex        =   1
      ToolTipText     =   "Выйти из программы"
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ФИГ ВАМ!!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   540
      TabIndex        =   36
      Top             =   960
      Visible         =   0   'False
      Width           =   7275
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "  Quadrosoft MegaBox Multimedia Player. Version 1.01.002 (новогодняя версия)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
   Begin VB.Menu lstMenu 
      Caption         =   "001"
      Visible         =   0   'False
      Begin VB.Menu ыПрибавить 
         Caption         =   "Добавить файлы"
         Shortcut        =   ^A
      End
      Begin VB.Menu llFind 
         Caption         =   "Поиск"
         Shortcut        =   ^F
      End
      Begin VB.Menu удалить 
         Caption         =   "Удалить из списка"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu ва1 
         Caption         =   "-"
      End
      Begin VB.Menu слеар 
         Caption         =   "Очистить список"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu пробелка 
         Caption         =   "-"
      End
      Begin VB.Menu стартонуть 
         Caption         =   "Воспоизвести"
         Shortcut        =   {F5}
      End
      Begin VB.Menu xxsep 
         Caption         =   "-"
      End
      Begin VB.Menu mmKill 
         Caption         =   "Стереть файл с диска"
         Shortcut        =   ^K
      End
      Begin VB.Menu llCopy 
         Caption         =   "Копировать в другое место"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldX As Long
Dim OldY As Long

Sub DeleteItem()
On Local Error Resume Next
List1.RemoveItem List1.ListIndex
List1.ListIndex = SE - 1
End Sub

Sub PlayFile()
Dim FileName As String
Dim EndTime As String
Dim Secnds As Currency

Timer1.Enabled = False
On Local Error Resume Next
MMControl1.Command = "Close"
FileName = List1.List(List1.ListIndex)
MTrack1.UnWait

MMControl1.FileName = FileName
MMControl1.Command = "Open"
MMControl1.Command = "Play"
   If Dialog1.Hallo.Value = 1 Then
     MMControl1.To = MMControl1.Length / 2
     MMControl1.Command = "Seek"
     MMControl1.Command = "Play"
     Timer1.Enabled = True
   End If
Label4.Caption = "Идет воспроизведение."
MTrack1.Track = List1.ListIndex + 1
Text1.Text = MMControl1.FileName
MTimer4.TimeSet = Mid$(Time$, 1, 5)



Secnds = Fix(CCur(MMControl1.Length) / CCur(1000))


MTimer5.TimeSet = Format$(TimeSerial(Hour(Now), Minute(Now), Second(Now) + Secnds), "hh:mm")

End Sub

Private Sub bIc_Click()
ListView1.View = lvwIcon
End Sub

Sub RandomPlay(Playing As Integer)
Dim Ml, RPlay As Integer
List1.Selected(Playing) = False

For Ml = 0 To List1.ListCount - 1
If List1.Selected(Ml) = True Then Ok = 1
Next

If Ok = 0 Then Exit Sub

Do
RPlay = (List1.ListCount - 1) * Rnd(-Timer)
DoEvents
Loop While Not List1.Selected(RPlay) = True
List1.ListIndex = RPlay
PlayFile
End Sub

Private Sub Command1_Click()
End
End Sub



Private Sub Command10_Click()
Mois = Shell("SndVol32", vbNormalFocus)
End Sub

Private Sub Command11_Click()
D = Shell("Explorer.exe", vbNormalFocus)
End Sub


Private Sub Command12_Click()
Kill "C:\Windows\System\register.dat"
Dialog2.Show
Command12.Enabled = False
Unload Form1
End Sub

Private Sub Command13_Click()
frmAbout.Show
End Sub

Private Sub Command14_Click()
MMControl1.UpdateInterval = 0
MTimer1.Reset
MTimer2.Reset
MTimer3.Reset
Slider1.Enabled = False
Command14.Enabled = False
Command20.Enabled = True
End Sub

Private Sub Command15_Click()
Dialog.Show
End Sub

Private Sub Command16_Click()
On Error Resume Next
COD.FileName = ""
COD.ShowSave
 If COD.FileName <> "" Then
 If Err Then Exit Sub
 Open COD.FileName For Output As #1
 Label4.Caption = "Идет сохранение..."
 For LBLS = 0 To List1.ListCount - 1
  Print #1, List1.List(LBLS)
 Next
 Close
 Label4.Caption = "Сохранено."
 End If
Close
End Sub


Private Sub Command17_Click()
On Local Error Resume Next
COD.ShowOpen
If COD.FileName <> "" Then
 If Err Then Exit Sub
 Open COD.FileName For Input As #1
 Label4.Caption = "Идет открытие..."
 List1.Clear
 Do While Not EOF(1)
 Line Input #1, Itms$
 List1.AddItem Itms$
 Loop

Close #1
Label4.Caption = "Список открыт."

End If
End Sub

Private Sub Command18_Click()
DeleteItem
End Sub

Private Sub Command19_Click()
List1.Clear
End Sub

Private Sub Command2_Click()
Form1.WindowState = 1
End Sub

Private Sub Command20_Click()
MTimer1.InitTimer
MTimer2.InitTimer
MTimer3.InitTimer
MMControl1.UpdateInterval = 1000
Slider1.Enabled = True
Command14.Enabled = True
Command20.Enabled = False

End Sub

Private Sub Command21_Click()
MMControl1.Command = "Close"
MMControl1.FileName = App.Path + "\applause.wav"
MMControl1.Command = "Open"
MMControl1.Command = "Play"
End Sub

Private Sub Command22_Click()
MMControl1.Command = "Close"
MMControl1.FileName = App.Path + "\ovations.wav"
MMControl1.Command = "Open"
MMControl1.Command = "Play"

End Sub


Private Sub Command23_Click()
MMControl1.Command = "Close"
MMControl1.FileName = App.Path + "\foo.wav"
MMControl1.Command = "Open"
MMControl1.Command = "Play"

End Sub


Private Sub Command3_Click()
cSysTray1.InTray = True
Form1.Hide
MMControl1.UpdateInterval = 0
End Sub

Private Sub Command4_Click()
Call PlayFile
End Sub

Private Sub Command5_Click()
MMControl1.Command = "Pause"
Dialog1.Hallo.Value = 0
Timer1.Enabled = False
End Sub

Private Sub Command6_Click()
MMControl1.Command = "Stop"
MMControl1.Command = "Close"
MMControl1_StatusUpdate
Call MTrack1.Waiting
End Sub


Private Sub Command7_Click()
On Error Resume Next
List1.ListIndex = List1.ListIndex - 1
PlayFile

End Sub

Private Sub Command8_Click()
On Error Resume Next
List1.ListIndex = List1.ListIndex + 1
PlayFile
End Sub

Private Sub Command9_Click()
Dialog1.Show
End Sub

Private Sub cSysTray1_MouseDown(Button As Integer, Id As Long)

Form1.Show
cSysTray1.InTray = False
MMControl1.UpdateInterval = 1000

End Sub

Private Sub Form_Load()
MMControl1_StatusUpdate
Call MTrack1.Waiting
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
OldX = X
OldY = y

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 1 Then
MoveX = X - OldX
MoveY = y - OldY

Form1.Move Form1.Left + MoveX, Form1.Top + MoveY
End If
End Sub



Private Sub Image1_Click()
Form4.Show
End Sub

Private Sub List1_DblClick()
Command4_Click
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then PopupMenu lstMenu, , List1.Left + X, List1.Top + y

End Sub


Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)

For Listov = 1 To Data.Files.Count
Dat = Data.Files(Listov)
List1.AddItem Dat
List1.Selected(List1.ListCount - 1) = True
Next

End Sub


Private Sub lstIc_Click()
ListView1.View = lvwList
End Sub

Private Sub llCopy_Click()
MMControl1.Command = "Close"
Dest = InputBox("Введите путь и имя назначения (C:\KARAOKE\melody1.mp3)", "Copy to Disk", App.Path + "\*.*")
If Dest = "" Then Exit Sub
On Error Resume Next

FileCopy List1.List(List1.ListIndex), Dest

If Err <> 0 Then Beep


End Sub

Private Sub llFind_Click()
If List1.ListCount < 12 Then MsgBox ("Сам ищи!!!"): Exit Sub
End Sub


Private Sub MMControl1_Done(NotifyCode As Integer)
If Dialog1.RAND.Value = False Then If List1.ListIndex = List1.ListCount - 1 Then Exit Sub
If NotifyCode = 1 Then
 MMControl1.Command = "Close"
 Label4.Caption = "Идет смена мелодии..."
  If Dialog1.UpDown.Value = True Then Call Command8_Click
  If Dialog1.DownUp.Value = True Then Call Command7_Click
  If Dialog1.RAND.Value = True Then RandomPlay (List1.ListIndex)
 Label4.Caption = "Идет воспроизведение."
Else
  Call MTrack1.Waiting
End If

End Sub

Private Sub MMControl1_StatusUpdate()
On Local Error Resume Next

MMControl1.TimeFormat = vbMCIFormatMilliseconds

If (MMControl1.Length / 1000) >= 60 Then MMControl1.UpdateInterval = 1000
If (MMControl1.Length / 1000) < 60 Then MMControl1.UpdateInterval = 100
If (MMControl1.Length / 1000) < 30 Then MMControl1.UpdateInterval = 10
If (MMControl1.Length / 1000) < 3 Then MMControl1.UpdateInterval = 1

Slider1.Max = MMControl1.Length / 1000
Slider1.Value = MMControl1.Position / 1000

aSeconds = Fix(MMControl1.Position / 1000) Mod 60
aMinutes = Fix((MMControl1.Position / 1000) / 60)

bSeconds = Fix(MMControl1.Length / 1000) Mod 60
bMinutes = Fix((MMControl1.Length / 1000) / 60)


cSeconds = Fix((MMControl1.Length - MMControl1.Position) / 1000) Mod 60
cMinutes = Fix(((MMControl1.Length - MMControl1.Position) / 1000) / 60)

If Dialog1.Check3.Value = 1 Then MTimer1.TimeSet = Format$(aMinutes, "00") + ":" + Format$(aSeconds, "00") Else MTimer1.Off
If Dialog1.Check2.Value = 1 Then MTimer2.TimeSet = Format$(bMinutes, "00") + ":" + Format$(bSeconds, "00") Else MTimer2.Off
If Dialog1.Check1.Value = 1 Then MTimer3.TimeSet = Format$(cMinutes, "00") + ":" + Format$(cSeconds, "00") Else MTimer3.Off


End Sub


Private Sub mmKill_Click()
Dim title As String
MMControl1.Command = "Close"
title = List1.List(List1.ListIndex)
On Error Resume Next
Kill List1.List(List1.ListIndex)
List1.RemoveItem List1.ListIndex

If Err = 0 Then
    Label4.Caption = "Файл был стерен с диска."
Else
    Call MsgBox("Удаление файла невозможно!", vbExclamation, title)
End If

End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
 
 Case 37
  Form1.Left = Form1.Left - Screen.TwipsPerPixelX * 4
  
 Case 38
  Form1.Top = Form1.Top - Screen.TwipsPerPixelY * 4

 Case 39
  Form1.Left = Form1.Left + Screen.TwipsPerPixelX * 4

 Case 40
  Form1.Top = Form1.Top + Screen.TwipsPerPixelY * 4
End Select

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case KeyCode
 
 Case 37
  Form1.Left = Form1.Left - Screen.TwipsPerPixelX * 4
  
 Case 38
  Form1.Top = Form1.Top - Screen.TwipsPerPixelY * 4

 Case 39
  Form1.Left = Form1.Left + Screen.TwipsPerPixelX * 4

 Case 40
  Form1.Top = Form1.Top + Screen.TwipsPerPixelY * 4
End Select

End Sub





Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
OldX = X
OldY = y

End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button Then
MoveX = X - OldX
MoveY = y - OldY
  
  If Button = 2 Then Form1.Move 225 * Fix((Form1.Left + MoveX) / 225), 225 * Fix((Form1.Top + MoveY) / 225)
  If Button = 1 Then Form1.Move Form1.Left + MoveX, Form1.Top + MoveY
End If

End Sub


Private Sub sIc_Click()
ListView1.View = lvwSmallIcon
End Sub







Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
MMControl1.UpdateInterval = 0
MMControl1.Command = "Stop"
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
MMControl1.To = CCur(Slider1.Value) * CCur(1000)
MMControl1.Command = "Seek"
MMControl1.Command = "Play"
MMControl1.UpdateInterval = 1000

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Command8_Click
End Sub

Private Sub Timer2_Timer()
MTimer6.TimeSet = Format$(Now, "hh:mm")
End Sub

Private Sub слеар_Click()
List1.Clear
End Sub

Private Sub удалить_Click()
DeleteItem
End Sub


Private Sub ыПрибавить_Click()
Command15_Click
End Sub


