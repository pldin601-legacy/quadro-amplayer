VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Программирование таймера"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2220
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2220
      Top             =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Время выключения"
      Height          =   1335
      Left            =   2640
      TabIndex        =   4
      Top             =   180
      Width           =   1875
      Begin ComCtl2.UpDown UpDown2 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   327681
         BuddyControl    =   "HourOff"
         BuddyDispid     =   196622
         OrigLeft        =   240
         OrigTop         =   480
         OrigRight       =   480
         OrigBottom      =   735
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   13
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   327681
         BuddyControl    =   "MinOff"
         BuddyDispid     =   196623
         OrigLeft        =   1140
         OrigTop         =   480
         OrigRight       =   1380
         OrigBottom      =   735
         Max             =   59
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Мин"
         Height          =   195
         Index           =   1
         Left            =   900
         TabIndex        =   17
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Час"
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   16
         Top             =   240
         Width           =   300
      End
      Begin VB.Label HourOff 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   480
         Width           =   255
      End
      Begin VB.Label MinOff 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Left            =   900
         TabIndex        =   14
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Время включения"
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   180
      Width           =   1875
      Begin ComCtl2.UpDown UpDown2 
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   327681
         BuddyControl    =   "HourOn"
         BuddyDispid     =   196614
         OrigLeft        =   180
         OrigTop         =   480
         OrigRight       =   420
         OrigBottom      =   735
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   450
         _Version        =   327681
         BuddyControl    =   "MinOn"
         BuddyDispid     =   196615
         OrigLeft        =   1080
         OrigTop         =   480
         OrigRight       =   1320
         OrigBottom      =   735
         Max             =   59
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65537
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Мин"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Час"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   480
         Width           =   45
      End
      Begin VB.Label MinOn 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   480
         Width           =   255
      End
      Begin VB.Label HourOn 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Left            =   420
         TabIndex        =   5
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Подробнее >>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   1140
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   660
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ОК"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
Timer2.Enabled = True
Form3.Hide
Form1.TrackAll1.TimeCDSet = Format$(Val(HourOn.Caption), "00") + ":" + Format$(Val(MinOn.Caption), "00")
Form1.TrackAll2.TimeCDSet = Format$(Val(HourOff.Caption), "00") + ":" + Format$(Val(MinOff.Caption), "00")
End Sub

Private Sub Command2_Click()
Form1.TrackAll1.TimeCDSet = "00:00"
Form1.TrackAll2.TimeCDSet = "00:00"
Unload Form3
End Sub

Private Sub Timer1_Timer()
If Format$(Now, "h:m") = HourOff.Caption + ":" + MinOff.Caption Then Form1.MMControl1.Command = "Stop": Timer1.Enabled = False

End Sub


Private Sub Timer2_Timer()
If Format$(Now, "h:m") = HourOn.Caption + ":" + MinOn.Caption Then Form1.PlayFile: Timer2.Enabled = False
End Sub


