VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form5"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4200
   LinkTopic       =   "Form5"
   ScaleHeight     =   6840
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      Caption         =   "bidang studi"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.Frame Frame4 
         Caption         =   "SEARCH"
         Height          =   1335
         Left            =   120
         TabIndex        =   16
         Top             =   5160
         Width           =   3735
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   3255
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Refresh"
            Height          =   615
            Left            =   1800
            Picture         =   "Form5.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Search"
            Height          =   615
            Left            =   240
            Picture         =   "Form5.frx":3EB0
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   13
         Text            =   "Pilih..."
         Top             =   480
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form5.frx":7A56
         Height          =   1575
         Left            =   120
         TabIndex        =   12
         Top             =   3600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2778
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
         Left            =   120
         Top             =   3240
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
         RecordSource    =   "select*from bidang_studi"
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
      Begin VB.Frame Frame3 
         BackColor       =   &H0000FFFF&
         Caption         =   "Date Time"
         Height          =   1335
         Left            =   2040
         TabIndex        =   8
         Top             =   1680
         Width           =   1815
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   195
            Left            =   1680
            TabIndex        =   14
            Top             =   1080
            Width           =   75
         End
         Begin VB.Timer Timer1 
            Interval        =   1
            Left            =   1320
            Top             =   600
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label4 
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
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H000000FF&
         Caption         =   "Button"
         Height          =   1935
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1935
         Begin VB.CommandButton Command4 
            Caption         =   "Print"
            Height          =   615
            Left            =   120
            Picture         =   "Form5.frx":7A6B
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Next"
            Height          =   615
            Left            =   960
            Picture         =   "Form5.frx":B4F5
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Save"
            Height          =   615
            Left            =   120
            Picture         =   "Form5.frx":EAC3
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   600
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm"
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
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bidang Nama"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bidang Kode"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1 = "AK" Then
Text1 = "Akutansi"
ElseIf Combo1 = "AP" Then
Text1 = "Administrasi Perkantoran"
ElseIf Combo1 = "RPL" Then
Text1 = "Rekayasa Perangkat Lunak"
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Combo1 = "" Or Text1 = "" Then
MsgBox "Masih ada data yang belum terisi!", , "Informasi"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset("bidang_kode") = Combo1
Adodc1.Recordset("bidang_nama") = Text1
Adodc1.Recordset.Update
MsgBox "Data Telah Tersimpan", , "Informasi"
Command1.Visible = False
Command1.Enabled = False
End If
End Sub

Private Sub Command2_Click()
Form8.Show
Me.Hide
End Sub

Private Sub Command3_Click()
With Adodc1.Recordset
.Delete
.MoveFirst
End With
End Sub

Private Sub Command4_Click()
DataReport1.Show
End Sub

Private Sub Command5_Click()
Adodc1.RecordSource = "select*from bidang_studi where bidang_kode like '" & Text2.Text & "'"
Adodc1.Refresh
End Sub

Private Sub Command6_Click()
Adodc1.RecordSource = "select*from bidang_studi"
Adodc1.Refresh
End Sub

Private Sub Form_Load()
With Combo1
.AddItem "AK"
.AddItem "AP"
.AddItem "RPL"
End With
End Sub

Private Sub Label3_Click()
If Combo1 = "" Or Text1 = "" Then
MsgBox "Masih ada data yang belum terisi!", , "Informasi"
Else
Command1.Visible = True
MsgBox "Silahkan Simpan!", , "Informasi"
End If
End Sub

Private Sub Timer1_Timer()
Label5.Caption = Format(Time, "hh:mm:ss")
Label6.Caption = Format(Date, "dd-mm-yyyy")
End Sub
