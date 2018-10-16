VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form3"
   ClientHeight    =   8730
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5040
   LinkTopic       =   "Form3"
   ScaleHeight     =   8730
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      Caption         =   "Wali Murid"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.Frame Frame4 
         Caption         =   "SEARCH"
         Height          =   1335
         Left            =   120
         TabIndex        =   28
         Top             =   6960
         Width           =   3735
         Begin VB.CommandButton Command6 
            Caption         =   "Refresh"
            Height          =   615
            Left            =   1800
            Picture         =   "Form3.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Search"
            Height          =   615
            Left            =   120
            Picture         =   "Form3.frx":3EB0
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   3495
         End
      End
      Begin MSAdodcLib.Adodc Adodc3 
         Height          =   330
         Left            =   2280
         Top             =   5160
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
         Left            =   3480
         Top             =   4800
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
         RecordSource    =   ""
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
      Begin VB.Frame Frame3 
         BackColor       =   &H0000FFFF&
         Caption         =   "Date Time"
         Height          =   1335
         Left            =   2280
         TabIndex        =   22
         Top             =   3480
         Width           =   2415
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   195
            Left            =   2280
            TabIndex        =   26
            Top             =   1080
            Width           =   75
         End
         Begin VB.Timer Timer1 
            Interval        =   1
            Left            =   1560
            Top             =   600
         End
         Begin VB.Label Label12 
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
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H000000FF&
         Caption         =   "Button"
         Height          =   1935
         Left            =   120
         TabIndex        =   18
         Top             =   3480
         Width           =   2175
         Begin VB.CommandButton Command4 
            Caption         =   "Print"
            Height          =   615
            Left            =   120
            Picture         =   "Form3.frx":7A56
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Next"
            Height          =   615
            Left            =   1080
            Picture         =   "Form3.frx":B4E0
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Save"
            Height          =   615
            Left            =   120
            Picture         =   "Form3.frx":EAAE
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   600
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label9 
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
            TabIndex        =   19
            Top             =   240
            Width           =   1575
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form3.frx":121F2
         Height          =   1455
         Left            =   120
         TabIndex        =   17
         Top             =   5520
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   2566
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
         Left            =   2280
         Top             =   4800
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
         RecordSource    =   "select*from wali_murid"
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
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   16
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   14
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Top             =   1200
         Width           =   2655
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Form3.frx":12207
         DataField       =   "siswa_nisn"
         DataSource      =   "Adodc3"
         Height          =   315
         Left            =   2040
         TabIndex        =   10
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "siswa_nisn"
         Text            =   "Pilih..."
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Wali Telepon"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Wali Alamat"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Wali Pekerjaan Ibu"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Wali Nama Ibu"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Wali Pekerjaan Ayah"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Wali Nama Ayah"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Siswa Nisn"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Wali Id"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Text1 = "" Or DataCombo1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
MsgBox "Masih ada data yang belum terisi!", , "Informasi"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset("wali_id") = Text1
Adodc1.Recordset("siswa_nisn") = DataCombo1
Adodc1.Recordset("wali_nama_ayah") = Text2
Adodc1.Recordset("wali_pekerjaan_ayah") = Text3
Adodc1.Recordset("wali_nama_ibu") = Text4
Adodc1.Recordset("wali_pekerjaan_ibu") = Text5
Adodc1.Recordset("wali_alamat") = Text6
Adodc1.Recordset("wali_telepon") = Text7
Adodc1.Recordset.Update
MsgBox "Data Telah Tersimpan", , "Informasi"
Command1.Visible = False
Command1.Enabled = False
End If
End Sub

Private Sub Command2_Click()
Form5.Show
Me.Hide
End Sub

Private Sub Command3_Click()
With Adodc1.Recordset
.Delete
.MoveFirst
End With
End Sub

Private Sub Command4_Click()
DataReport7.Show
End Sub

Private Sub Command5_Click()
Adodc1.RecordSource = "select*from wali_murid where wali_id like '" & Text8.Text & "'"
Adodc1.Refresh
End Sub

Private Sub Command6_Click()
Adodc1.RecordSource = "select*from wali_murid"
Adodc1.Refresh
End Sub

Private Sub DataCombo1_Click(Area As Integer)
Adodc2.RecordSource = "select*from siswa where siswa_nisn='" & DataCombo1.Text & "'"
Adodc2.Refresh
End Sub

Private Sub Label9_Click()
If Text1 = "" Or DataCombo1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Or Text7 = "" Then
MsgBox "Masih ada data yang belum terisi!", , "Informasi"
Else
Command1.Visible = True
MsgBox "Silahkan Simpan!", , "Informasi"
End If
End Sub

Private Sub Timer1_Timer()
Label10.Caption = Format(Time, "hh:mm:ss")
Label11.Caption = Format(Time, "dd-mm-yyyy")
End Sub
