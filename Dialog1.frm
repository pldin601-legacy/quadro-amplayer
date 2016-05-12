VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form Dialog1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3795
   ClientLeft      =   2760
   ClientTop       =   3465
   ClientWidth     =   7695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      Caption         =   "Справка"
      Height          =   375
      Left            =   5040
      TabIndex        =   31
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Programming"
      Height          =   375
      Left            =   5040
      TabIndex        =   30
      ToolTipText     =   "Программирование таймера"
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Считать"
      Height          =   375
      Left            =   6540
      TabIndex        =   29
      Top             =   3300
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   5340
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   27
      Top             =   3240
      Width           =   1095
      Begin Project1.MTimer MTimer1 
         Height          =   330
         Left            =   60
         TabIndex        =   28
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Индикатор"
      Height          =   1275
      Left            =   4680
      TabIndex        =   22
      Top             =   1800
      Width           =   2895
      Begin VB.CheckBox Check3 
         Caption         =   "Время: позиция"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   900
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Время: длительность"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Время: осталось"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   300
         Value           =   1  'Checked
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "r"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      Caption         =   "5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7080
      TabIndex        =   16
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7380
      TabIndex        =   15
      Top             =   0
      Width           =   315
   End
   Begin VB.CheckBox Hallo 
      Caption         =   "Режим ознакомления"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   2115
   End
   Begin VB.Frame Frame4 
      Caption         =   "Окна"
      Height          =   1275
      Left            =   2520
      TabIndex        =   11
      Top             =   480
      Width           =   2115
      Begin VB.OptionButton Option7 
         Caption         =   "MegaBox Forms"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   780
         Value           =   -1  'True
         Width           =   1515
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Windows 95 standart"
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   1875
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Настройка шкалы"
      Height          =   1275
      Left            =   2520
      TabIndex        =   10
      Top             =   1800
      Width           =   2115
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   660
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   327681
         Value           =   30
         AutoBuddy       =   -1  'True
         BuddyControl    =   "Label3"
         BuddyDispid     =   196625
         OrigLeft        =   600
         OrigTop         =   660
         OrigRight       =   840
         OrigBottom      =   915
         Max             =   120
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds = 1 Tick"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "30"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   660
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Частота шкалы"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Свойства DragDrop"
      Height          =   1275
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   2355
      Begin VB.OptionButton Option5 
         Caption         =   "Полное перемещение (MegaBox 1.01)"
         Height          =   435
         Left            =   180
         TabIndex        =   9
         Top             =   720
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Перенос контура (Windows 95)"
         Enabled         =   0   'False
         Height          =   435
         Left            =   180
         TabIndex        =   8
         Top             =   240
         Width           =   1995
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Воспроизведение"
      Height          =   1275
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2355
      Begin VB.OptionButton RAND 
         Caption         =   "Случайным образом"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   900
         Width           =   1935
      End
      Begin VB.OptionButton DownUp 
         Caption         =   "Снизу-вверх"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton UpDown 
         Caption         =   "Сверху-вниз"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Длительность заданий (час: мин) "
      Height          =   195
      Left            =   2580
      TabIndex        =   26
      Top             =   3360
      Width           =   2595
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   60
      X2              =   7620
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   60
      X2              =   7620
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Настройка программы"
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
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldX As Long, OldY As Long



Private Sub CancelButton_Click()
UpDown.Value = True
Dialog1.Hide
End Sub

Private Sub Command4_Click()
Dim Seco As Long
What = MsgBox("Счет длительности всех заданий может затянуться на долго. Продолжить?", vbInformation + vbYesNo, "Счет заданий")
If What = 7 Then Exit Sub
Form2.Show 0
Form1.MMControl1.Command = "Close"

For Opens = 0 To Form1.List1.ListCount - 1
Form2.ProgressBar1.Max = Form1.List1.ListCount - 1
Form2.ProgressBar1.Value = Opens
Form2.Label1.Caption = "Счет..." + Format$(100 / (Form1.List1.ListCount - 1) * Opens, "0") + "%"
Form2.MTimer1.TimeSet = Format$(Seco / 3600, "00") + ":" + Format$((Seco / 60) Mod 60, "00")
Form1.MMControl1.FileName = Form1.List1.List(Opens)
If Form2.Check1.Value = 1 Then MsgBox "Счет отменен", vbInformation: Form2.Hide: Exit Sub
DoEvents
Form1.MMControl1.Command = "Open"
Form1.MMControl1.TimeFormat = vbMCIFormatMilliseconds
Seco = Seco + (Form1.MMControl1.Length / 1000)
Form1.MMControl1.Command = "Close"
Next
MTimer1.TimeSet = Format$(Seco / 3600, "00") + ":" + Format$((Seco / 60) Mod 60, "00")
Form2.Hide
End Sub

Private Sub Command5_Click()
Form3.Show
End Sub

Private Sub Hallo_Click()
If Hallo.Value = 1 Then Frame1.Enabled = False
If Hallo.Value = 0 Then Frame1.Enabled = True
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
OldX = X
OldY = y
End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
 If Button = 1 Then
     MoveX = X - OldX
     MoveY = y - OldY
     Dialog1.Move Dialog1.Left + MoveX, Dialog1.Top + MoveY
 End If
End Sub


Private Sub OKButton_Click()
Dialog1.Hide
End Sub


Private Sub UpDown1_Change()
' Label3.Caption = Format$(UpDown.Value, "0")
Form1.Slider1.TickFrequency = Val(Label3.Caption)
End Sub
