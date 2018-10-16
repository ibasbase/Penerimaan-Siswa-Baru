VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00FFFF00&
   Caption         =   "P"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form6"
   ScaleHeight     =   7050
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      Caption         =   "kopetensi keahlian"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.Frame Frame4 
         Caption         =   "SEARCH"
         Height          =   1455
         Left            =   120
         TabIndex        =   18
         Top             =   5280
         Width           =   3855
         Begin VB.CommandButton Command6 
            Caption         =   "Refresh"
            Height          =   615
            Left            =   1680
            Picture         =   "Form6.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Search"
            Height          =   615
            Left            =   120
            Picture         =   "Form6.frx":3EB0
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Text            =   "Pilih..."
         Top             =   840
         Width           =   2415
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   2760
         Top             =   2880
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
         CommandType     =   8
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
         Caption         =   "Adodc3"
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   2040
         Top             =   3120
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
         CommandType     =   8
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
         RecordSource    =   "select*from guru"
         Caption         =   "Adodc2"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form6.frx":7A56
         Height          =   1575
         Left            =   120
         TabIndex        =   14
         Top             =   3600
         Width           =   3855
         _ExtentX        =   6800
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
         Left            =   2040
         Top             =   2880
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
         RecordSource    =   "select*from kopetensi_keahlian"
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
         TabIndex        =   10
         Top             =   1560
         Width           =   1935
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   195
            Left            =   1800
            TabIndex        =   16
            Top             =   1080
            Width           =   75
         End
         Begin VB.Timer Timer1 
            Interval        =   1
            Left            =   1440
            Top             =   600
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label5 
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
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H000000FF&
         Caption         =   "Button"
         Height          =   1935
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1935
         Begin VB.CommandButton Command4 
            Caption         =   "Print"
            Height          =   615
            Left            =   120
            Picture         =   "Form6.frx":7A6B
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Next"
            Height          =   615
            Left            =   960
            Picture         =   "Form6.frx":B4F5
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Save"
            Height          =   615
            Left            =   120
            Picture         =   "Form6.frx":EAC3
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   600
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label4 
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
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   1200
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Form6.frx":12207
         DataField       =   "kopetensi_kode"
         DataSource      =   "Adodc3"
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "kopetensi_kode"
         Text            =   "Pilih..."
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Kopetensi Nama"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bidang Kode"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kopetensi Kode"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1 = "AK" Then
Text1 = "Pembukuan"
ElseIf Combo1 = "AP" Then
Text1 = "Sekertaris"
ElseIf Combo1 = "RPL" Then
Text1 = "Komputer"
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Text1 = "" Or DataCombo1 = "" Or Combo1 = "" Then
MsgBox "Masih ada data yang belum terisi!", , "Informasi"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset("kopetensi_kode") = DataCombo1
Adodc1.Recordset("bidang_kode") = Combo1
Adodc1.Recordset("kopetensi_nama") = Text1
Adodc1.Recordset.Update
MsgBox "Data Telah Tersimpan!", , "Informasi"
Command1.Visible = False
Command1.Enabled = False
End If
End Sub

Private Sub Command2_Click()
Form7.Show
Me.Hide
End Sub

Private Sub Command3_Click()
With Adodc1.Recordset
.Delete
.MoveFirst
End With
End Sub

Private Sub Command4_Click()
DataReport3.Show
End Sub

Private Sub Command5_Click()
Adodc1.RecordSource = "select*from nilai where siswa_nisn like '" & Text1.Text & "'"
Adodc1.Refresh
End Sub

Private Sub DataCombo1_Click(Area As Integer)
muncul
End Sub
Public Sub muncul()
Adodc2.RecordSource = "select*from siswa where kopetensi_kode ='" & DataCombo1.Text & "'"
Adodc2.Refresh
End Sub

Private Sub Form_Load()
With Combo1
.AddItem "AK"
.AddItem "AP"
.AddItem "RPL"
End With
End Sub

Private Sub Label4_Click()
If DataCombo1 = "" Or Text1 = "" Or Combo1 = "" Then
MsgBox "Masih ada data yang belum terisi!", , "Informasi"
Else
Command1.Visible = True
MsgBox "Silakan simpan", , "Informasi"
End If
End Sub

Private Sub Timer1_Timer()
Label6.Caption = Format(Time, "hh:mm:ss")
Label7.Caption = Format(Date, "dd-mm-yyyy")
End Sub
