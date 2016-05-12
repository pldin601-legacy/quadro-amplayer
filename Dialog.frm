VERSION 5.00
Begin VB.Form Dialog 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3600
   ClientLeft      =   2760
   ClientTop       =   3465
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3600
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Выделить все"
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   1500
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Text            =   "All Files (*.multimedia)"
      Top             =   780
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   780
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2520
      TabIndex        =   4
      Top             =   3060
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   120
      MultiSelect     =   2  'Extended
      Pattern         =   "*.mid;*.wav;*.avi;*.mpg;*.dat;*.mp3;*.mp2"
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Добавить файлы"
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
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   6975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Шаблон :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   765
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldX As Long, OldY As Long


Option Explicit

Private Sub CancelButton_Click()
Dialog.Hide
End Sub

Private Sub Combo1_Click()
On Local Error Resume Next
Dir1.Path = Combo1.List(Combo1.ListIndex)
If Err Then Combo1.RemoveItem Combo1.ListIndex: Beep
End Sub


Private Sub Combo2_Click()
File1.Pattern = Combo2.Text
End Sub


Private Sub Command1_Click()
Dim dm

For dm = 0 To File1.ListCount - 1
File1.Selected(dm) = True
Next
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Local Error GoTo Driver
Dir1.Path = Drive1.Drive
Exit Sub

Driver:
Drive1.Drive = Dir1.Path
Resume 0
End Sub

Private Sub File1_DblClick()
OKButton_Click
End Sub


Private Sub File1_KeyPress(KeyAscii As Integer)
OKButton_Click
End Sub



Private Sub Form_Load()
Combo2.AddItem "*.MID; Midi Files"
Combo2.AddItem "*.WAV; Wave Files"
Combo2.AddItem "*.MP3; Music Files"
Combo2.AddItem "*.MPG; Video Files"
Combo2.AddItem "*.AVI; Video for Windows Files"
Combo2.AddItem "*.DAT; Video CD Files"
Combo2.AddItem "*.RMI; Riff Midi Files"
End Sub

Private Sub Form_Paint()
Dim lPath$

Combo1.Clear

On Local Error Resume Next
Open App.Path + "\history.mb1" For Input As #5
If Err Then Close: Exit Sub
' On Local Error GoTo 0
Do While Not EOF(5)
Line Input #5, lPath$
Combo1.AddItem lPath$
Loop
Close

End Sub


Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
OldX = X
OldY = y
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

Dim MoveX As Long
Dim MoveY As Long

 If Button = 1 Then
     MoveX = X - OldX
     MoveY = y - OldY
     Dialog.Move Dialog.Left + MoveX, Dialog.Top + MoveY
 End If
End Sub

Private Sub OKButton_Click()

Dim FilesAdd As Currency
Dim Separator$, Modo As Integer, Yest As Integer

If Right$(Dir1.Path, 1) <> "\" Then Separator$ = "\"

    For FilesAdd = 0 To File1.ListCount - 1
    If File1.Selected(FilesAdd) = True Then Form1.List1.AddItem File1.Path + Separator$ + File1.List(FilesAdd):    Form1.List1.Selected(Form1.List1.ListCount - 1) = True


    Next
Combo1.AddItem Dir1.Path
Dialog.Hide


   Open App.Path + "\history.mb1" For Output As #1
   For Modo = 0 To Combo1.ListCount - 1
   Print #1, Combo1.List(Modo)
   Next
   Close
End Sub

Private Sub Text1_Change()
File1.Pattern = Combo2.Text
End Sub
