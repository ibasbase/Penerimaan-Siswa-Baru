VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form9 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form9"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14490
   LinkTopic       =   "Form9"
   ScaleHeight     =   7980
   ScaleWidth      =   14490
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14175
      Begin TabDlg.SSTab SSTab1 
         Height          =   5055
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   8916
         _Version        =   393216
         Tabs            =   7
         Tab             =   1
         TabsPerRow      =   7
         TabHeight       =   520
         TabCaption(0)   =   "Bidang Studi"
         TabPicture(0)   =   "Form9.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "DataGrid1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Adodc1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame8"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Guru"
         TabPicture(1)   =   "Form9.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Adodc2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "DataGrid2"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame7"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Kopetensi Keahlian"
         TabPicture(2)   =   "Form9.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Adodc3"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "DataGrid3"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Frame6"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "Nilai"
         TabPicture(3)   =   "Form9.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Adodc4"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "DataGrid4"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "Frame5"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).ControlCount=   3
         TabCaption(4)   =   "Siswa"
         TabPicture(4)   =   "Form9.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Adodc5"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "DataGrid5"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "Frame4"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).ControlCount=   3
         TabCaption(5)   =   "Standar Kopetensi"
         TabPicture(5)   =   "Form9.frx":008C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Adodc6"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).Control(1)=   "DataGrid6"
         Tab(5).Control(1).Enabled=   0   'False
         Tab(5).Control(2)=   "Frame3"
         Tab(5).Control(2).Enabled=   0   'False
         Tab(5).ControlCount=   3
         TabCaption(6)   =   "Wali Murid"
         TabPicture(6)   =   "Form9.frx":00A8
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Adodc7"
         Tab(6).Control(0).Enabled=   0   'False
         Tab(6).Control(1)=   "DataGrid7"
         Tab(6).Control(1).Enabled=   0   'False
         Tab(6).Control(2)=   "Frame2"
         Tab(6).Control(2).Enabled=   0   'False
         Tab(6).ControlCount=   3
         Begin VB.Frame Frame8 
            Caption         =   "SEARCH"
            Height          =   1455
            Left            =   -73560
            TabIndex        =   33
            Top             =   3240
            Width           =   3495
            Begin VB.CommandButton Command14 
               Caption         =   "Refresh"
               Height          =   735
               Left            =   1680
               Picture         =   "Form9.frx":00C4
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   600
               Width           =   1695
            End
            Begin VB.CommandButton Command13 
               Caption         =   "Search"
               Height          =   735
               Left            =   120
               Picture         =   "Form9.frx":3F74
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox Text7 
               Height          =   285
               Left            =   120
               TabIndex        =   34
               Top             =   240
               Width           =   3255
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "SEARCH"
            Height          =   1455
            Left            =   1440
            TabIndex        =   29
            Top             =   3240
            Width           =   3855
            Begin VB.CommandButton Command12 
               Caption         =   "Refresh"
               Height          =   735
               Left            =   1920
               Picture         =   "Form9.frx":7B1A
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   600
               Width           =   1815
            End
            Begin VB.CommandButton Command11 
               Caption         =   "Search"
               Height          =   735
               Left            =   120
               Picture         =   "Form9.frx":B9CA
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox Text6 
               Height          =   285
               Left            =   120
               TabIndex        =   30
               Top             =   240
               Width           =   3615
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "SEARCH"
            Height          =   1575
            Left            =   -73680
            TabIndex        =   25
            Top             =   3240
            Width           =   3975
            Begin VB.CommandButton Command10 
               Caption         =   "Refresh"
               Height          =   615
               Left            =   2040
               Picture         =   "Form9.frx":F570
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   720
               Width           =   1815
            End
            Begin VB.CommandButton Command9 
               Caption         =   "Search"
               Height          =   615
               Left            =   120
               Picture         =   "Form9.frx":13420
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   720
               Width           =   1935
            End
            Begin VB.TextBox Text5 
               Height          =   285
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   3735
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "SEARCH"
            Height          =   1455
            Left            =   -73680
            TabIndex        =   21
            Top             =   3240
            Width           =   4095
            Begin VB.CommandButton Command8 
               Caption         =   "Refresh"
               Height          =   615
               Left            =   1920
               Picture         =   "Form9.frx":16FC6
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   600
               Width           =   1935
            End
            Begin VB.CommandButton Command7 
               Caption         =   "Search"
               Height          =   615
               Left            =   120
               Picture         =   "Form9.frx":1AE76
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox Text4 
               Height          =   285
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   3735
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "SEARCH"
            Height          =   1455
            Left            =   -73560
            TabIndex        =   17
            Top             =   3240
            Width           =   3975
            Begin VB.CommandButton Command6 
               Caption         =   "Refresh"
               Height          =   615
               Left            =   1920
               Picture         =   "Form9.frx":1EA1C
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   600
               Width           =   1935
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Search"
               Height          =   615
               Left            =   120
               Picture         =   "Form9.frx":228CC
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   3735
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "SEARCH"
            ClipControls    =   0   'False
            Height          =   1455
            Left            =   -73560
            TabIndex        =   13
            Top             =   3240
            Width           =   3735
            Begin VB.CommandButton Command4 
               Caption         =   "Refresh"
               Height          =   615
               Left            =   1920
               Picture         =   "Form9.frx":26472
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   600
               Width           =   1695
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Search"
               Height          =   615
               Left            =   120
               Picture         =   "Form9.frx":2A322
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   3495
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "SEARCH"
            Height          =   1455
            Left            =   -73560
            TabIndex        =   9
            Top             =   3240
            Width           =   3735
            Begin VB.CommandButton Command2 
               Caption         =   "Refresh"
               Height          =   615
               Left            =   1800
               Picture         =   "Form9.frx":2DEC8
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   600
               Width           =   1815
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Search"
               Height          =   615
               Left            =   120
               Picture         =   "Form9.frx":31D78
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   600
               Width           =   1695
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   120
               TabIndex        =   10
               Top             =   240
               Width           =   3495
            End
         End
         Begin MSDataGridLib.DataGrid DataGrid7 
            Bindings        =   "Form9.frx":3591E
            Height          =   2655
            Left            =   -74880
            TabIndex        =   8
            Top             =   480
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   4683
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
         Begin MSAdodcLib.Adodc Adodc7 
            Height          =   330
            Left            =   -74880
            Top             =   3360
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
            Caption         =   "Adodc7"
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
         Begin MSDataGridLib.DataGrid DataGrid6 
            Bindings        =   "Form9.frx":35933
            Height          =   2655
            Left            =   -74880
            TabIndex        =   7
            Top             =   480
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   4683
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
         Begin MSAdodcLib.Adodc Adodc6 
            Height          =   330
            Left            =   -74880
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
            RecordSource    =   "select*from standar_kopetensi"
            Caption         =   "Adodc6"
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
         Begin MSDataGridLib.DataGrid DataGrid5 
            Bindings        =   "Form9.frx":35948
            Height          =   2655
            Left            =   -74880
            TabIndex        =   6
            Top             =   480
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   4683
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
         Begin MSAdodcLib.Adodc Adodc5 
            Height          =   375
            Left            =   -74880
            Top             =   3240
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   661
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
            Caption         =   "Adodc5"
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
         Begin MSDataGridLib.DataGrid DataGrid4 
            Bindings        =   "Form9.frx":3595D
            Height          =   2655
            Left            =   -74880
            TabIndex        =   5
            Top             =   480
            Width           =   13335
            _ExtentX        =   23521
            _ExtentY        =   4683
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
         Begin MSAdodcLib.Adodc Adodc4 
            Height          =   330
            Left            =   -74880
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
            RecordSource    =   "select*from nilai"
            Caption         =   "Adodc4"
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
         Begin MSDataGridLib.DataGrid DataGrid3 
            Bindings        =   "Form9.frx":35972
            Height          =   2655
            Left            =   -74880
            TabIndex        =   4
            Top             =   480
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   4683
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
         Begin MSAdodcLib.Adodc Adodc3 
            Height          =   330
            Left            =   -74880
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
            RecordSource    =   "select*from kopetensi_keahlian"
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
         Begin MSDataGridLib.DataGrid DataGrid2 
            Bindings        =   "Form9.frx":35987
            Height          =   2655
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   4683
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
         Begin MSAdodcLib.Adodc Adodc2 
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
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   -74880
            Top             =   3240
            Width           =   1215
            _ExtentX        =   2143
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "Form9.frx":3599C
            Height          =   2655
            Left            =   -74880
            TabIndex        =   2
            Top             =   480
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   4683
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
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc7.RecordSource = "select*from wali_murid where wali_id like '" & Text1.Text & "'"
Adodc7.Refresh
End Sub

Private Sub Command10_Click()
Adodc3.RecordSource = "select*from kopetensi_keahlian"
Adodc3.Refresh
End Sub

Private Sub Command11_Click()
Adodc2.RecordSource = "select*from guru where guru_kode like '" & Text6.Text & "'"
Adodc2.Refresh
End Sub

Private Sub Command12_Click()
Adodc2.RecordSource = "select*from guru"
Adodc2.Refresh
End Sub

Private Sub Command13_Click()
Adodc1.RecordSource = "select*from bidang_studi where bidang_kode like '" & Text7.Text & "'"
Adodc1.Refresh
End Sub

Private Sub Command14_Click()
Adodc1.RecordSource = "select*from bidang_studi"
Adodc1.Refresh
End Sub

Private Sub Command2_Click()
Adodc7.RecordSource = "select*from wali_murid"
Adodc7.Refresh
End Sub

Private Sub Command3_Click()
Adodc6.RecordSource = "select*from standar_kopetensi where sk_kode like '" & Text2.Text & "'"
Adodc6.Refresh
End Sub

Private Sub Command4_Click()
Adodc6.RecordSource = "select*from standar_kopetensi"
Adodc6.Refresh
End Sub

Private Sub Command5_Click()
Adodc5.RecordSource = "select*from siswa where siswa_nisn like '" & Text3.Text & "'"
Adodc5.Refresh
End Sub

Private Sub Command6_Click()
Adodc5.RecordSource = "select*from siswa"
Adodc5.Refresh
End Sub

Private Sub Command7_Click()
Adodc4.RecordSource = "select*from nilai where siswa_nisn like '" & Text4.Text & "'"
Adodc4.Refresh
End Sub

Private Sub Command8_Click()
Adodc4.RecordSource = "select*from nilai"
Adodc4.Refresh
End Sub

Private Sub Command9_Click()
Adodc3.RecordSource = "select*from kopetensi_keahlian where kopetensi_kode like '" & Text5.Text & "'"
Adodc3.Refresh
End Sub
