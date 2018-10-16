VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00FFFF80&
   Caption         =   "Form2"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7425
   LinkTopic       =   "Form2"
   ScaleHeight     =   8520
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      Caption         =   "Siswa"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.Frame Frame4 
         Caption         =   "SEARCH"
         Height          =   1335
         Left            =   120
         TabIndex        =   24
         Top             =   6480
         Width           =   3735
         Begin VB.CommandButton Command7 
            Caption         =   "Refresh"
            Height          =   615
            Left            =   1800
            Picture         =   "Form2.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Search"
            Height          =   615
            Left            =   120
            Picture         =   "Form2.frx":3EB0
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Browse"
         Height          =   375
         Left            =   1560
         TabIndex        =   23
         Top             =   2400
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3360
         Top             =   4320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Print"
         Height          =   615
         Left            =   240
         Picture         =   "Form2.frx":7A56
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Top             =   960
         Width           =   2295
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0000FFFF&
         Caption         =   "Date Time"
         Height          =   1455
         Left            =   2040
         TabIndex        =   16
         Top             =   2880
         Width           =   1815
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   195
            Left            =   1680
            TabIndex        =   21
            Top             =   960
            Width           =   75
         End
         Begin VB.Timer Timer1 
            Interval        =   1
            Left            =   1320
            Top             =   480
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Date Time"
            BeginProperty Font 
               Name            =   "Algerian"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   975
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form2.frx":B4E0
         Height          =   1695
         Left            =   120
         TabIndex        =   15
         Top             =   4800
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2990
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   2040
         Top             =   4440
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=UJIKOM.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=UJIKOM.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select*from siswa"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H000000FF&
         Caption         =   "Button"
         Height          =   1935
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   1935
         Begin VB.CommandButton Command2 
            Caption         =   "Next"
            Height          =   615
            Left            =   960
            Picture         =   "Form2.frx":B4F5
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Save"
            Height          =   615
            Left            =   120
            Picture         =   "Form2.frx":EAC3
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   600
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "CONFIRM"
            BeginProperty Font 
               Name            =   "Algerian"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   3615
         Left            =   3960
         Stretch         =   -1  'True
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Siswa Foto"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Siswa Tgl Lahir"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Siswa Alamat"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Siswa Nama"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Kopetensi Kode"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Siswa Nisn"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim save As Integer

Private Sub Combo1_Click()
If Combo1 = "Ibas" Then
Picture1.Picture = LoadPicture("base\aa.jpg")
ElseIf Combo1 = "Base" Then
Picture1.Picture = LoadPicture("base\bb.jpg")
ElseIf Combo1 = "Sabi" Then
Picture1.Picture = LoadPicture("base\dd.jpg")
ElseIf Combo1 = "Owo" Then
Picture1.Picture = LoadPicture("base\ee.jpg")
ElseIf Combo1 = "Nibras" Then
Picture1.Picture = LoadPicture("base\asasasas.jpg")
ElseIf Combo1 = "Abilowo" Then
Picture1.Picture = LoadPicture("base\abab.jpg")
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Image1 = "" Then
MsgBox "Masih ada data yang belum terisi!", , "Informasi"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset("siswa_nisn") = Text1
Adodc1.Recordset("kopetensi_kode") = Text2
Adodc1.Recordset("siswa_nama") = Text3
Adodc1.Recordset("siswa_alamat") = Text4
Adodc1.Recordset("siswa_tgl_lahir") = Text5
Adodc1.Recordset("siswa_foto") = CommonDialog1.FileName
Adodc1.Recordset.Update
MsgBox "Data telah tersimpan", , "Informasi"
Command1.Visible = False
Command1.Enabled = False
End If
End Sub

Private Sub Command2_Click()
Form3.Show
Me.Hide
End Sub

Private Sub Command3_Click()
With Adodc1.Recordset
.Delete
.MoveFirst
End With
End Sub

Private Sub Command4_Click()
DataReport5.Show
End Sub

Private Sub Command5_Click()
On Error GoTo err
CommonDialog1.ShowOpen
Image1.Picture = LoadPicture(CommonDialog1.FileName)
err:
End Sub

Private Sub Command6_Click()
Adodc1.RecordSource = "select*from siswa where siswa_nisn like '" & Text6.Text & "'"
Adodc2.Refresh
End Sub

Private Sub Command7_Click()
Adodc1.RecordSource = "select*from siswa"
Adodc1.Refresh
End Sub

Private Sub Label7_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Image1 = "" Then
MsgBox "Masih ada data yang belum terisi!", , "Informasi"
Else
Command1.Visible = True
MsgBox "Silakan simpan"
End If
End Sub

Private Sub Text1_keypress(keyascii As Integer)
If keyascii = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Text2_keypress(keyascii As Integer)
If keyascii = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub Text3_keypress(keyascii As Integer)
If keyascii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text4_keypress(keyascii As Integer)
If keyascii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Text5_keypress(keyascii As Integer)
If keyascii = 13 Then
Text6.SetFocus
End If
End Sub


Private Sub Timer1_Timer()
Label8.Caption = Format(Time, "hh:mm:ss")
Label9.Caption = Format(Date, "dd-mm-yyyy")
End Sub
