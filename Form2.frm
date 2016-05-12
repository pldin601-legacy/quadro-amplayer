VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2340
      ScaleHeight     =   315
      ScaleWidth      =   915
      TabIndex        =   4
      Top             =   60
      Width           =   975
      Begin Project1.MTimer MTimer1 
         Height          =   330
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   300
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   60
      Width           =   1035
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ָהוע סקוע..."
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   885
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Check1.Value = 1
End Sub
