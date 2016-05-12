VERSION 5.00
Begin VB.Form Dialog2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Регистрация приветствует вас!!!"
   ClientHeight    =   1245
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   315
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "А свой номер ты знаешь"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "О, имя твое, юзерок"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   1560
   End
End
Attribute VB_Name = "Dialog2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
End
End Sub

Private Sub Form_Load()
Dim MyName$, MyCode$
On Error Resume Next



Open "C:\Windows\System\register.dat" For Input As #1
If Err > 0 Then Exit Sub
If Err = 0 Then
 Line Input #1, MyName$
 Line Input #1, MyCode$
 If MyCode$ = "8-064-345-45-45" Then Form1.Show: Unload Dialog2
 
End If
Close


End Sub

Private Sub OKButton_Click()

Dim Batut As Integer
If Text2.Text = "8-064-345-45-45" Then
 Open "C:\Windows\system\register.dat" For Output As #1
 Print #1, Text1.Text
 Print #1, Text2.Text
 Close
 Dialog2.Hide
 Form1.Show
 Unload Dialog2

Else
 MsgBox ("Кодик-то не правельный введен!!! HA-HA-HA.")
 Form1.Show
 Dialog2.Hide
 Form1.Label9.Visible = True
 Form1.BackGround.Enabled = False
 For Batut = Form1.BackGround.Top To Form1.Height
  DoEvents
  Form1.BackGround.Top = Batut
 Next
 Unload Dialog1
End If
End Sub
