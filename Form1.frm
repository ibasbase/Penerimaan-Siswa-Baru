VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4050
   ScaleWidth      =   3810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1560
      Top             =   3240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      Picture         =   "Form1.frx":1C59
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "Form1.frx":5B09
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Text            =   "Pilih..."
      Top             =   1560
      Width           =   2295
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   3240
      Width           =   615
      URL             =   "3 RPL 2 - Teman.mp3"
      rate            =   1
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
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1085
      _cy             =   873
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aa_Click()

End Sub

Private Sub Command1_Click()
If Combo1 = "Admin" And Text1 = "myadmin" Then
Form2.Show
Me.Hide
ElseIf Combo1 = "User" And Text1 = "myuser" Then
Form9.Show
Me.Hide
Else
MsgBox "Password Salah!", vbCritical, "Informasi"
Text1 = ""
End If
End Sub

Private Sub Form_Load()
With Combo1
.AddItem "Admin"
.AddItem "User"
End With
End Sub

Private Sub Text1_keypress(keyascii As Integer)
If keyascii = 13 Then
Command1.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Format(Time, "hh:mm:ss")
Label4.Caption = Format(Date, "dd-mm-yyyy")
End Sub

