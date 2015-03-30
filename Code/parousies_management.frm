VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form parousies_management 
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   11550
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14430
   FillStyle       =   0  'Solid
   LinkTopic       =   "entity_management"
   MDIChild        =   -1  'True
   ScaleHeight     =   11550
   ScaleWidth      =   14430
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000014&
      DisabledPicture =   "parousies_management.frx":0000
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   12600
      MaskColor       =   &H80000014&
      Picture         =   "parousies_management.frx":268B
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   """Εξαγωγή στοιχείων σε Excel"""
      Top             =   9840
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc ado_athlites_tmimatos 
      Height          =   375
      Left            =   10560
      Top             =   5880
      Visible         =   0   'False
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
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=poseidon.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=poseidon.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Αθλητές_Τμημάτων"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dt_athlites_tmimatos 
      Bindings        =   "parousies_management.frx":4D16
      Height          =   1575
      Left            =   6480
      TabIndex        =   63
      Top             =   6600
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2778
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   161
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
            LCID            =   1032
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
            LCID            =   1032
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Παρουσιολόγιο του Μήνα <<"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7695
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Visible         =   0   'False
      Width           =   13575
      Begin MSAdodcLib.Adodc ado_parousiologia 
         Height          =   375
         Left            =   8280
         Top             =   4560
         Visible         =   0   'False
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=poseidon.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=poseidon.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Παρουσιολόγια"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox txt_minas_per 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00 ""€"""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   12000
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   6480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton sin_apous 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ανανέωση"
         DisabledPicture =   "parousies_management.frx":4D3A
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   12495
         MaskColor       =   &H80000014&
         Picture         =   "parousies_management.frx":9B68
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   360
         Width           =   975
      End
      Begin MSAdodcLib.Adodc ado_analytiko_par 
         Height          =   375
         Left            =   120
         Top             =   7200
         Width           =   12260
         _ExtentX        =   21616
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=poseidon.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=poseidon.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "GiaThnProvoliStoixeiwnParousiologiou"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid dt_parousiologia 
         Bindings        =   "parousies_management.frx":E996
         Height          =   1455
         Left            =   10080
         TabIndex        =   64
         Top             =   4920
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   18
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   35
         BeginProperty Column00 
            DataField       =   "id_par"
            Caption         =   "id_par"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   4
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   8
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "id_tmimatos"
            Caption         =   "id_tmimatos"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "id_athliti"
            Caption         =   "id_athliti"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "id_mina"
            Caption         =   "id_mina"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "d1"
            Caption         =   "d1"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "d2"
            Caption         =   "d2"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "d3"
            Caption         =   "d3"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "d4"
            Caption         =   "d4"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "d5"
            Caption         =   "d5"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "d6"
            Caption         =   "d6"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "d7"
            Caption         =   "d7"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "d8"
            Caption         =   "d8"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "d9"
            Caption         =   "d9"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "d10"
            Caption         =   "d10"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column14 
            DataField       =   "d11"
            Caption         =   "d11"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column15 
            DataField       =   "d12"
            Caption         =   "d12"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column16 
            DataField       =   "d13"
            Caption         =   "d13"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column17 
            DataField       =   "d14"
            Caption         =   "d14"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column18 
            DataField       =   "d15"
            Caption         =   "d15"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column19 
            DataField       =   "d16"
            Caption         =   "d16"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column20 
            DataField       =   "d17"
            Caption         =   "d17"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column21 
            DataField       =   "d18"
            Caption         =   "d18"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column22 
            DataField       =   "d19"
            Caption         =   "d19"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column23 
            DataField       =   "d20"
            Caption         =   "d20"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column24 
            DataField       =   "d21"
            Caption         =   "d21"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column25 
            DataField       =   "d22"
            Caption         =   "d22"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column26 
            DataField       =   "d23"
            Caption         =   "d23"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column27 
            DataField       =   "d24"
            Caption         =   "d24"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column28 
            DataField       =   "d25"
            Caption         =   "d25"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column29 
            DataField       =   "d26"
            Caption         =   "d26"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column30 
            DataField       =   "d27"
            Caption         =   "d27"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column31 
            DataField       =   "d28"
            Caption         =   "d28"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column32 
            DataField       =   "d29"
            Caption         =   "d29"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column33 
            DataField       =   "d30"
            Caption         =   "d30"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column34 
            DataField       =   "d31"
            Caption         =   "d31"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   734,74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   734,74
            EndProperty
            BeginProperty Column04 
               Button          =   -1  'True
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column05 
               Button          =   -1  'True
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column20 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column21 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column22 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column23 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column24 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column25 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column26 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column27 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column28 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column29 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column30 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column31 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column32 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column33 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column34 
               ColumnWidth     =   615,118
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dt_analytiko_par 
         Bindings        =   "parousies_management.frx":E9B6
         Height          =   6855
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   12240
         _ExtentX        =   21590
         _ExtentY        =   12091
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BackColor       =   -2147483628
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   18
         TabAction       =   2
         RowDividerStyle =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   39
         BeginProperty Column00 
            DataField       =   "id_tmimatos"
            Caption         =   "id_tmimatos"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "id_athliti"
            Caption         =   "id_athliti"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "OE"
            Caption         =   "Ονοματεπώνυμο"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "EG"
            Caption         =   "Γέννηση"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "id_mina"
            Caption         =   "id_mina"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "d1"
            Caption         =   "1"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "d2"
            Caption         =   "2"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "d3"
            Caption         =   "3"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "d4"
            Caption         =   "4"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "d5"
            Caption         =   "5"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "d6"
            Caption         =   "6"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "d7"
            Caption         =   "7"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "d8"
            Caption         =   "8"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "d9"
            Caption         =   "9"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column14 
            DataField       =   "d10"
            Caption         =   "10"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column15 
            DataField       =   "d11"
            Caption         =   "11"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column16 
            DataField       =   "d12"
            Caption         =   "12"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column17 
            DataField       =   "d13"
            Caption         =   "13"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column18 
            DataField       =   "d14"
            Caption         =   "14"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column19 
            DataField       =   "d15"
            Caption         =   "15"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column20 
            DataField       =   "d16"
            Caption         =   "16"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column21 
            DataField       =   "d17"
            Caption         =   "17"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column22 
            DataField       =   "d18"
            Caption         =   "18"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column23 
            DataField       =   "d19"
            Caption         =   "19"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column24 
            DataField       =   "d20"
            Caption         =   "20"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column25 
            DataField       =   "d21"
            Caption         =   "21"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   "0,000E+00"
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column26 
            DataField       =   "d22"
            Caption         =   "22"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column27 
            DataField       =   "d23"
            Caption         =   "23"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column28 
            DataField       =   "d24"
            Caption         =   "24"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column29 
            DataField       =   "d25"
            Caption         =   "25"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column30 
            DataField       =   "d26"
            Caption         =   "26"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column31 
            DataField       =   "d27"
            Caption         =   "27"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   "0,000E+00"
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column32 
            DataField       =   "d28"
            Caption         =   "28"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column33 
            DataField       =   "d29"
            Caption         =   "29"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column34 
            DataField       =   "d30"
            Caption         =   "30"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column35 
            DataField       =   "d31"
            Caption         =   "31"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "1"
               FalseValue      =   "0"
               NullValue       =   "0"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column36 
            DataField       =   "id_par"
            Caption         =   "id_par"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column37 
            DataField       =   "sm"
            Caption         =   " Σύνολα"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column38 
            DataField       =   "P"
            Caption         =   "Ο.Τ."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Ναι"
               FalseValue      =   "Όχι"
               NullValue       =   "Όχι"
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   7
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            SizeMode        =   1
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1995,024
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   734,74
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column07 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column08 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column09 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column10 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column11 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column12 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column13 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column14 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column15 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column16 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column17 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column18 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column19 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column20 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column21 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column22 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column23 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column24 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column25 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column26 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column27 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column28 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column29 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column30 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column31 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column32 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column33 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column34 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column35 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   255,118
            EndProperty
            BeginProperty Column36 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   734,74
            EndProperty
            BeginProperty Column37 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column38 
               Alignment       =   2
               DividerStyle    =   3
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   404,787
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton bt_print 
         BackColor       =   &H80000014&
         Caption         =   "Εκτύπω&ση"
         DisabledPicture =   "parousies_management.frx":E9D6
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   12495
         MaskColor       =   &H80000014&
         Picture         =   "parousies_management.frx":133DB
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   6840
         Width           =   960
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Επέλεξε Μήνα Παρουσιολογίου για το τρέχων τμήμα και το τρέχων Αθλητικό Έτος"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   38
      Top             =   3000
      Width           =   13575
      Begin MSDataGridLib.DataGrid dt_athl_sm 
         Bindings        =   "parousies_management.frx":17DE0
         Height          =   615
         Left            =   10200
         TabIndex        =   68
         Top             =   120
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1085
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
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
               LCID            =   1032
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
               LCID            =   1032
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
      Begin VB.OptionButton Option1 
         Caption         =   "Σεπτέμβριος"
         Height          =   400
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Οκτώβριος"
         Height          =   400
         Index           =   1
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Νοέμβριος"
         Height          =   400
         Index           =   2
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Δεκέμβριος"
         Height          =   400
         Index           =   3
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ιανουάριος"
         Height          =   400
         Index           =   4
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Φεβρουάριος"
         Height          =   400
         Index           =   5
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Μάρτιος"
         Height          =   400
         Index           =   6
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Απρίλιος"
         Height          =   400
         Index           =   7
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Μάϊος"
         Height          =   400
         Index           =   8
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ιούνιος"
         Height          =   400
         Index           =   9
         Left            =   10920
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ιούλιος"
         Height          =   400
         Index           =   10
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Σεπτέμβριος"
         Height          =   495
         Index           =   21
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Οκτώβριος"
         Height          =   495
         Index           =   20
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Νοέμβριος"
         Height          =   495
         Index           =   19
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Δεκέμβριος"
         Height          =   495
         Index           =   18
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ιανουάριος"
         Height          =   495
         Index           =   17
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Φεβρουάριος"
         Height          =   495
         Index           =   16
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Μάρτιος"
         Height          =   495
         Index           =   15
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Απρίλιος"
         Height          =   495
         Index           =   14
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   3960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Μάϊος"
         Height          =   495
         Index           =   13
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ιούνιος"
         Height          =   495
         Index           =   12
         Left            =   10920
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3960
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ιούλιος"
         Height          =   495
         Index           =   11
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3960
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc ado_athl_sm 
         Height          =   375
         Left            =   9600
         Top             =   120
         Visible         =   0   'False
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=poseidon.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=poseidon.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "ΑθλητέςΣυνδρομέςΜήνα"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.CommandButton kl_bt 
      BackColor       =   &H80000014&
      Caption         =   "Κ&λείσιμο"
      DisabledPicture =   "parousies_management.frx":17DFA
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   12495
      MaskColor       =   &H80000014&
      Picture         =   "parousies_management.frx":1D872
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   360
      Width           =   1080
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Στοιχεία Τμήματος"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   13575
      Begin VB.TextBox txt_id_mn 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00 ""€"""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame7 
         Caption         =   "Εβδομαδιαίο Πρόγραμμα Τμήματος"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   4920
         TabIndex        =   19
         Top             =   360
         Width           =   7340
         Begin MSMask.MaskEdBox im_l 
            Height          =   300
            Index           =   1
            Left            =   1800
            TabIndex        =   5
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_l 
            Height          =   300
            Index           =   0
            Left            =   840
            TabIndex        =   3
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_l 
            Height          =   300
            Index           =   2
            Left            =   2760
            TabIndex        =   7
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_l 
            Height          =   300
            Index           =   3
            Left            =   3720
            TabIndex        =   9
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_l 
            Height          =   300
            Index           =   4
            Left            =   4680
            TabIndex        =   11
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_l 
            Height          =   300
            Index           =   5
            Left            =   5640
            TabIndex        =   13
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_en 
            Height          =   300
            Index           =   0
            Left            =   840
            TabIndex        =   2
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_en 
            Height          =   300
            Index           =   1
            Left            =   1800
            TabIndex        =   4
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_en 
            Height          =   300
            Index           =   2
            Left            =   2760
            TabIndex        =   6
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_en 
            Height          =   300
            Index           =   3
            Left            =   3720
            TabIndex        =   8
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_en 
            Height          =   300
            Index           =   4
            Left            =   4680
            TabIndex        =   10
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_en 
            Height          =   300
            Index           =   5
            Left            =   5640
            TabIndex        =   12
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Λήξη "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Έναρξη "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label5 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Δευτέρα"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   26
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Σάββατο"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   5640
            TabIndex        =   31
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Παρασκευή"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   4680
            TabIndex        =   30
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Πέμπτη"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   3720
            TabIndex        =   29
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Τετάρτη"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   28
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Τρίτη"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   27
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Γενικά Στοιχεία"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   4815
         Begin VB.TextBox txt_id_tm 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txt_ae 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox txt_kt 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txt_pa 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox txt_pb 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox tmp_poso_mina 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   2040
            Width           =   2415
         End
         Begin VB.TextBox tmp_poso_eggrafis 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Προπονητής Β"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   840
            TabIndex        =   25
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ποσό Μηνιαίας Συνδρομής"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   24
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ποσό Εγγραφής"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   720
            TabIndex        =   23
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Προπονητής Α"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   840
            TabIndex        =   22
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Κατηγορία Τμήματος"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   21
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Αθλητικό Έτος"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   20
            Top             =   240
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "parousies_management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public a_new_tm_is_ready_to_add As Integer
Public id_τμ, new_addition, col_affected, defined_col As Integer
Public f_n As String
Public rs_ado_ae As ADODB.Recordset
Public rs_ado_athl As ADODB.Recordset
Public rs_ado_athl_tm As ADODB.Recordset
Public rs_ado_athl_tm_ids As ADODB.Recordset
Public rs_ado_kt As ADODB.Recordset
Public rs_ado_prop As ADODB.Recordset
Public rs_ado_tm_ids As ADODB.Recordset
Public rs_ado_tmimata As ADODB.Recordset
Public s_sort As String
Public for_search As Integer
Public at, tt As String

Const strChecked = "ώ"
Const strUnChecked = "q"

Dim mn As Integer

Private Sub bt_print_Click()
    
    Call sin_apous_Click
    Rep_ΠαρουσίεςΑνάΜήνα_1_Τμήματος.Show

End Sub

Private Sub dt_athlites_HeadClick(ByVal ColIndex As Integer)
    
    defined_col = ColIndex
    
End Sub

Private Sub dt_athlites_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If rs_ado_tmimata.AbsolutePosition >= 1 And rs_ado_tmimata.AbsolutePosition <= rs_ado_tmimata.RecordCount Then
    
    If Trim(rs_ado_tmimata.Fields(1).Value) <> "" Then
        tmp_am.Text = rs_ado_tmimata.Fields(1).Value
    Else
        tmp_am.Text = ""
    End If
    If Trim(rs_ado_tmimata.Fields(3).Value) <> "" Then
        tmp_onoma.Text = rs_ado_tmimata.Fields(3).Value
    Else
        tmp_onoma.Text = ""
    End If
    If Trim(rs_ado_tmimata.Fields(2).Value) <> "" Then
        tmp_eponimo.Text = rs_ado_tmimata.Fields(2).Value
    Else
        tmp_eponimo.Text = ""
    End If
    If Trim(rs_ado_tmimata.Fields(4).Value) <> "" Then
        tmp_odos.Text = rs_ado_tmimata.Fields(4).Value
    Else
        tmp_odos.Text = ""
    End If
    If Trim(rs_ado_tmimata.Fields(5).Value) <> "" Then
        tmp_arithmos.Text = rs_ado_tmimata.Fields(5).Value
    Else
        tmp_arithmos.Text = ""
    End If
    'If Trim(rs_ado_tmimata.Fields(6).Value) <> "" Then
    '    tmp_perioxi.Text = rs_ado_tmimata.Fields(6).Value
    'Else
    '    tmp_perioxi.Text = ""
    'End If
    'If Trim(rs_ado_tmimata.Fields(8).Value) <> "" Then
    '    Me.co_pe.Text = rs_ado_tmimata.Fields(8).Value
    'Else
    '    Me.co_pe.Text = ""
    'End If
    'If Trim(rs_ado_tmimata.Fields(7).Value) <> "" Then
    '    Me.co_dimoi.Text = rs_ado_tmimata.Fields(7).Value
    'Else
    '    Me.co_dimoi.Text = ""
    'End If
    If Trim(rs_ado_tmimata.Fields(9).Value) <> "" Then
        tmp_tk.Text = rs_ado_tmimata.Fields(9).Value
    Else
        tmp_tk.Text = ""
    End If
    If Trim(rs_ado_tmimata.Fields(10).Value) <> "" Then
        tmp_til_oikias.Text = rs_ado_tmimata.Fields(10).Value
    Else
        tmp_til_oikias.Text = ""
    End If
    If Trim(rs_ado_tmimata.Fields(11).Value) <> "" Then
        tmp_kinito.Text = rs_ado_tmimata.Fields(11).Value
    Else
        tmp_kinito.Text = ""
    End If
    If Trim(rs_ado_tmimata.Fields(12).Value) <> "" Then
        tmp_fax.Text = rs_ado_tmimata.Fields(12).Value
    Else
        tmp_fax.Text = ""
    End If
    If Trim(rs_ado_tmimata.Fields(13).Value) <> "" Then
        tmp_email.Text = rs_ado_tmimata.Fields(13).Value
    Else
        tmp_email.Text = ""
    End If
    'If Trim(rs_ado_tmimata.Fields(14).Value) <> "" Then
    '    Me.MaskEdBox1.Text = rs_ado_tmimata.Fields(14).Value
    'Else
    '    Me.MaskEdBox1.Text = "00/00/0000"
    'End If
    'If Trim(rs_ado_tmimata.Fields(15).Value) <> "" Then
    '    Me.ado_pateres.Recordset.Filter = "[id] LIKE '" & Str(rs_ado_tmimata.Fields(15).Value) & "'"
    'Else
    '    Me.ado_pateres.Recordset.Filter = "[id] LIKE '" & Str(-1) & "'"
    'End If
    'If Trim(rs_ado_tmimata.Fields(16).Value) <> "" Then
    '    Me.ado_miteres.Recordset.Filter = "[id] LIKE '" & Str(rs_ado_tmimata.Fields(16).Value) & "'"
    'Else
    '    Me.ado_miteres.Recordset.Filter = "[id] LIKE '" & Str(-1) & "'"
    'End If
    'If Trim(rs_ado_tmimata.Fields(17).Value) <> "" Then
    '    Me.co_sxolia.Text = rs_ado_tmimata.Fields(17).Value
    'Else
    '    Me.co_sxolia.Text = ""
    'End If
    
    End If

End Sub

Private Sub Command2_Click()
    
    
    answ = InputBox("Εισάγετε το πλήθος των προπονήσεων για το μήνα: " & Trim(parousies_management.txt_minas_per), "Εφαρμογή Διαχείρισης ΠΟΣΕΙΔΩΝΑ")
    
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    Dim oChart As Object     ' Excel Chart
    Dim r As Integer
   
    Clipboard.Clear
    Dim sData As Variant
    sData = ""
    If Me.ado_analytiko_par.Recordset.RecordCount >= 1 Then
        r = Me.ado_analytiko_par.Recordset.RecordCount
        Me.ado_analytiko_par.Recordset.MoveFirst
        sData = "ΣΥΝΟΛΟ ΠΑΡΟΥΣΙΩΝ ΑΝΑ ΑΘΛΗΤΗ ΣΤΟ ΤΜΗΜΑ ΤΟΥ ΓΙΑ ΤΟΝ ΕΠΙΛΕΓΜΕΝΟ ΜΗΝΑ: " & Trim(parousies_management.txt_minas_per) & vbCr
        sData = sData & "Ονοματεπώνυμο Αθλητή" & vbTab & "1" & vbTab & "2" & vbTab & "3" & vbTab & "4" & vbTab & "5" & vbTab & "6" & vbTab & "7" & vbTab & "8" & vbTab & "9" & vbTab & "10" & vbTab & "11" & vbTab & "12" & vbTab & "13" & vbTab & "14" & vbTab & "15" & vbTab & "16" & vbTab & "17" & vbTab & "18" & vbTab & "19" & vbTab & "19" & vbTab & "20" & vbTab & "21" & vbTab & "22" & vbTab & "23" & vbTab & "24" & vbTab & "25" & vbTab & "26" & vbTab & "27" & vbTab & "28" & vbTab & "29" & vbTab & "30" & vbTab & "31" & vbTab & "" & vbTab & "% Συμμετοχής Αθλητή στις προπονήσεις μηνιαίως" & vbCr
        For i = 0 To Me.ado_analytiko_par.Recordset.RecordCount - 1
            sData = sData & Me.ado_analytiko_par.Recordset.Fields("OE").Value & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d1").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d2").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d3").Value))
            sData = sData & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d4").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d5").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d6").Value))
            sData = sData & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d7").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d8").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d9").Value))
            sData = sData & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d10").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d11").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d12").Value))
            sData = sData & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d13").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d14").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d15").Value))
            sData = sData & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d16").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d17").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d18").Value))
            sData = sData & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d19").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d20").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d21").Value))
            sData = sData & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d22").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d23").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d24").Value))
            sData = sData & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d25").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d26").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d27").Value))
            sData = sData & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d28").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d29").Value)) & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d30").Value))
            sData = sData & vbTab & Abs(CInt(Me.ado_analytiko_par.Recordset.Fields("d31").Value))
            sData = sData & vbCr
            Me.ado_analytiko_par.Recordset.MoveNext
        Next i
   End If
   Clipboard.SetText sData
   oBook.Worksheets(1).range("A1").Select
   oBook.Worksheets(1).Paste
   oSheet.range("A1:AG2").Font.Bold = True
   oSheet.range("A2:AG2").Font.ColorIndex = 3
   oSheet.range("A1").ColumnWidth = 25
   oSheet.range("B3:AG3").ColumnWidth = 5
   oBook.Worksheets(1).Columns.HorizontalAlignment = -4131
   
   For i = 3 To r + 2
    oSheet.range("AH" & i).Formula = "=Sum(B" & i & ":AG" & i & ")"
    myrange = oSheet.range("AH" & i)
    If answ <> "" Then
        oSheet.range("AI" & i).Formula = "=round(100*" & myrange & "/" & Val(answ) & ", 2)"
        If oSheet.range("AI" & i).Value = 0 Then
            'oSheet.Rows(i).Select
            oSheet.Rows(i).Interior.Color = vbYellow
        End If
    End If
   Next i
   
   For i = 2 To 34
    myrange = oBook.Worksheets(1).range(oBook.Worksheets(1).Cells(3, i), oBook.Worksheets(1).Cells(r + 3, i))
    results = oExcel.WorksheetFunction.Sum(myrange)
    oBook.Worksheets(1).Cells(r + 3, i).Formula = results
   Next i
   results = oExcel.WorksheetFunction.CountIf(oBook.Worksheets(1).range("AH3:AH" & r + 2), "=0")
   oSheet.range("AH" & r + 4).Formula = results
   oSheet.range("AI" & r + 4).Formula = "Πλήθος ΜΗ ΣΥΜΜΕΤΕΧΟΝΤΩΝ Αθλητών τον τρέχοντα μήνα"
   oSheet.range("AH" & r + 4 & ":AI" & r + 4).Font.Bold = True
   oSheet.range("AH" & r + 4 & ":AI" & r + 4).Font.ColorIndex = 3
   
    With oExcel.ActiveWindow
        .SplitColumn = 1
        .SplitRow = 2
    End With
    oExcel.ActiveWindow.FreezePanes = True
    oExcel.ScreenUpdating = True
    
    oExcel.Visible = True

End Sub

Private Sub Form_Deactivate()

    Unload Me

End Sub

Private Sub Form_Load()

    

    Me.Top = 100
    Me.Left = 100
    Me.Height = 12100
    Me.Width = 13800

    'ΕΝΗΜΕΡΩΣΗ ΓΕΝΙΚΩΝ ΣΤΟΙΧΕΙΩΝ ΤΜΗΜΑΤΟΣ
    Me.txt_id_tm = tmima_management.ado_tmimata.Recordset.Fields("TID").Value
    Me.txt_ae = tmima_management.co_ae.Text
    Me.txt_kt = tmima_management.co_kt.Text
    Me.txt_pa = tmima_management.co_propa.Text
    Me.txt_pb = tmima_management.co_propb.Text
    Me.tmp_poso_eggrafis = tmima_management.tmp_poso_eggrafis.Text
    Me.tmp_poso_mina = tmima_management.tmp_poso_mina.Text
    'ΕΝΗΜΕΡΩΣΗ ΕΒΔΟΜΑΔΙΑΙΟΥ ΠΡΟΓΡΑΜΜΑΤΟΣ ΤΜΗΜΑΤΟΣ
    For i = 0 To 5
        Me.im_en(i).Text = tmima_management.im_en(i).Text
        Me.im_l(i).Text = tmima_management.im_l(i).Text
    Next i
    'ΕΝΗΜΕΡΩΣΗ ΤΩΝ ΑΘΛΗΤΩΝ ΤΟΥ ΤΜΗΜΑΤΟΣ
    Me.ado_athlites_tmimatos.Recordset.Filter = "[id_Τμήματος] = " & Trim(Me.txt_id_tm.Text)
    
End Sub



Private Sub im_en_GotFocus(Index As Integer)
    
    Me.im_en(Index).SelStart = 0
    Me.im_en(Index).SelLength = Len(Me.im_en(Index).Text)
    
End Sub

Private Sub im_en_LostFocus(Index As Integer)

    If IsDate(Me.im_en(Index).Text) = False Then
        Me.im_en(Index).Text = "00:00"
        Me.im_en(Index).SelStart = 0
        Me.im_en(Index).SelLength = 5
        Me.im_en(Index).SelText = "00:00"
        Me.im_en(Index).SetFocus
    
    Else
        If Me.im_l(Index).Text <> "00:00" And Me.im_l(Index).Text <= Me.im_en(Index).Text Then
            MsgBox "Η ΩΡΑ ΕΝΑΡΞΗΣ πρέπει να είναι ΜΙΚΡΟΤΕΡΗ από την ΩΡΑ ΛΗΞΗΣ!", vbCritical, "Μήνυμα Λάθους"
            Me.im_en(Index).Text = "00:00"
            Me.im_en(Index).SelStart = 0
            Me.im_en(Index).SelLength = 5
            Me.im_en(Index).SelText = "00:00"
            Me.im_en(Index).SetFocus
        End If
    End If
    
End Sub

Private Sub im_l_GotFocus(Index As Integer)

    Me.im_l(Index).SelStart = 0
    Me.im_l(Index).SelLength = Len(Me.im_l(Index).Text)

End Sub

Private Sub im_l_LostFocus(Index As Integer)
    
    If IsDate(Me.im_l(Index).Text) = False Then
            Me.im_l(Index).Text = "00:00"
            Me.im_l(Index).SelStart = 0
            Me.im_l(Index).SelLength = 5
            Me.im_l(Index).SelText = "00:00"
            Me.im_l(Index).SetFocus
    
    Else
    
    If Me.im_l(Index).Text <> "00:00" And Me.im_l(Index).Text <= Me.im_en(Index).Text Then
        MsgBox "Η ΩΡΑ ΛΗΞΗΣ πρέπει να είναι ΜΕΓΑΛΥΤΕΡΗ από την ΩΡΑ ΕΝΑΡΞΗΣ!", vbCritical, "Μήνυμα Λάθους"
        Me.im_l(Index).Text = "00:00"
        Me.im_l(Index).SelStart = 0
        Me.im_l(Index).SelLength = 5
        Me.im_l(Index).SelText = "00:00"
        Me.im_l(Index).SetFocus
    End If
    
    End If
    
End Sub

Private Sub kl_bt_Click()

    Unload Me

End Sub

Private Sub sear_moth_Click()
    
    athlet_management.flag_mitera = 1
    meli_management.met_st.Enabled = True
    meli_management.Show
    
End Sub

Private Sub sear_pat_Click()

    athlet_management.flag_pateras = 1
    meli_management.met_st.Enabled = True
    meli_management.Show
    
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Or KeyAscii = 32 Then 'Enter/Space
        With MSFlexGrid1
            Call TriggerCheckbox(.Row, .Col)
        End With
    End If
    
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        With MSFlexGrid1
            If .MouseRow <> 0 And .MouseCol <> 0 Then
                Call TriggerCheckbox(.MouseRow, .MouseCol)
            End If
        End With
    End If
    
End Sub

Private Sub Option1_Click(Index As Integer)

    Dim tm, minas, id_athl As Integer
    
    Me.Frame1.Visible = True
    Me.Frame1.Caption = "Παρουσιολόγιο Αθλητών του Μήνα <<" & Option1(Index).Caption & ">>"
    Me.txt_minas_per.Text = Option1(Index).Caption
    
    mn = Index

    tm = Trim(Me.txt_id_tm.Text)
    minas = ((Index + 8) Mod 12) + 1
    Me.ado_parousiologia.Recordset.Requery
    Me.ado_parousiologia.Recordset.Filter = "[id_tmimatos] = " & tm & " AND [id_mina] = " & minas
    Me.txt_id_mn.Text = minas
    If Me.ado_parousiologia.Recordset.RecordCount >= 1 Then
        '
        'ΕΛΕΓΧΟΣ ΓΙΑ ΤΟ ΑΝ ΤΟ ΠΛΗΘΟΣ ΤΩΝ ΑΘΛΗΤΩΝ ΤΟΥ ΤΜΗΜΑΤΟΣ ΣΥΜΠΙΠΤΕΙ ΜΕ ΑΥΤΟ ΤΟΥ ΠΑΡΟΥΣΙΟΛΟΓΙΟΥ
        'MsgBox "parousiologia=" & Me.ado_parousiologia.Recordset.RecordCount & " athlites=" & Me.ado_athlites_tmimatos.Recordset.RecordCount
        If Me.ado_parousiologia.Recordset.RecordCount <> Me.ado_athlites_tmimatos.Recordset.RecordCount Then
            Me.ado_athlites_tmimatos.Recordset.MoveFirst
            For i = 1 To Me.ado_athlites_tmimatos.Recordset.RecordCount
                id_athl = Me.ado_athlites_tmimatos.Recordset.Fields("id_Αθλητή").Value
                Me.ado_parousiologia.Recordset.MoveFirst
                Me.ado_parousiologia.Recordset.Find "[id_athliti] like '" & str(id_athl) & "'", , adSearchForward, 0
                If Me.ado_parousiologia.Recordset.EOF Then 'ΔΕΝ ΤΟ ΕΧΕΙΣ ΒΡΕΙ
                    Me.ado_parousiologia.Recordset.AddNew
                    Me.ado_parousiologia.Recordset.Fields("id_tmimatos").Value = tm
                    Me.ado_parousiologia.Recordset.Fields("id_athliti").Value = id_athl
                    Me.ado_parousiologia.Recordset.Fields("id_mina").Value = minas
                    Me.ado_parousiologia.Recordset.UpdateBatch adAffectCurrent
                End If
                Me.ado_athlites_tmimatos.Recordset.MoveNext
            Next i
            Me.ado_parousiologia.Recordset.Requery
            Me.ado_parousiologia.Refresh
            Me.ado_analytiko_par.Recordset.Requery
            Me.ado_analytiko_par.Refresh
            Me.ado_analytiko_par.Recordset.Filter = "[id_tmimatos] = " & tm & " AND [id_mina] = " & minas
        Else
            Me.ado_analytiko_par.Recordset.Filter = "[id_tmimatos] = " & tm & " AND [id_mina] = " & minas
        End If
        '
    Else 'ΠΑΡΟΥΣΙΟΛΟΓΙΟ ΓΙΑ ΤΟ ΤΡΕΧΟΝΤΑ ΜΗΝΑ ΔΕΝ ΒΡΕΘΗΚΕ, ΟΠΟΤΕ ΔΗΜΙΟΥΡΓΕΙΤΑΙ
        'MsgBox "Δεν έχουν βρεθεί παρουσιολόγια για τον Μήνα = " & minas
        For i = 1 To Me.ado_athlites_tmimatos.Recordset.RecordCount
            If i = 1 Then
                Me.ado_athlites_tmimatos.Recordset.MoveFirst
            Else
                Me.ado_athlites_tmimatos.Recordset.MoveNext
            End If
            id_athl = Me.ado_athlites_tmimatos.Recordset.Fields("id_Αθλητή").Value
            Me.ado_parousiologia.Recordset.AddNew
            Me.ado_parousiologia.Recordset.Fields("id_tmimatos").Value = tm
            Me.ado_parousiologia.Recordset.Fields("id_athliti").Value = id_athl
            Me.ado_parousiologia.Recordset.Fields("id_mina").Value = minas
            Me.ado_parousiologia.Recordset.UpdateBatch adAffectCurrent
        Next i
        Me.ado_parousiologia.Recordset.Requery
        Me.ado_parousiologia.Refresh
        Me.ado_analytiko_par.Recordset.Requery
        Me.ado_analytiko_par.Refresh
        Me.ado_analytiko_par.Recordset.Filter = "[id_tmimatos] = " & tm & " AND [id_mina] = " & minas
        'CallRefreshParousiologio
    End If
    
    'ΕΝΗΜΕΡΩΣΗ ΣΧΕΤΙΚΑ ΜΕ ΤΗΝ ΟΙΚΟΝΟΜΙΚΗ ΤΑΚΤΟΠΟΙΗΣΗ ΤΩΝ ΑΘΛΗΤΩΝ ΓΙΑ ΤΟΝ ΤΡΕΧΟΝΤΑ ΜΗΝΑ
    If Me.ado_analytiko_par.Recordset.RecordCount >= 1 Then
        Dim ms_str As String
        ms_str = "Μήνας" & minas
        'Me.ado_athl_sm.Recordset.Filter = "[id_Τμήματος] = " & tm & " AND [Μήνας] = " & minas
        Me.ado_athl_sm.Recordset.Filter = "[id_Τμήματος] = " & tm & " AND [" & ms_str & "] = True"
        Me.ado_analytiko_par.Recordset.MoveFirst
        For i = 1 To Me.ado_analytiko_par.Recordset.RecordCount
            If Me.ado_analytiko_par.Recordset.Fields("P").Value = False Then
                id_athl = Me.ado_analytiko_par.Recordset.Fields("id_athliti").Value
                Me.ado_athl_sm.Recordset.Find "[id_Αθλητή] like '" & str(id_athl) & "'", , adSearchForward
                If Not Me.ado_athl_sm.Recordset.EOF Then 'ΤΟ ΕΧΕΙΣ ΒΡΕΙ
                    Me.ado_analytiko_par.Recordset.Fields("P").Value = True
                    Me.ado_analytiko_par.Recordset.UpdateBatch adAffectCurrent
                    Me.ado_athl_sm.Recordset.MoveFirst
                Else
                    If Me.ado_athl_sm.Recordset.RecordCount >= 1 Then
                        Me.ado_athl_sm.Recordset.MoveFirst
                    End If
                End If
            Else
                Me.dt_analytiko_par.SelBookmarks.Add Me.ado_analytiko_par.Recordset.Bookmark
            End If
            Me.ado_analytiko_par.Recordset.MoveNext
        Next i
    End If
    
    Me.Command2.Visible = True
    

End Sub

Private Sub Text1_Change()
    
    'Me.bt_ins_athl.Enabled = False
    'Me.bt_del_athl.Enabled = False
    'Me.bt_up_athl.Enabled = True
    'Me.bt_can_athl.Enabled = True
    
End Sub

Private Sub taksin_Click()

End Sub

Private Sub sin_apous_Click()

    Dim rw As Integer
    
    Me.dt_analytiko_par.Visible = False
    rw = Me.ado_analytiko_par.Recordset.AbsolutePosition
    Me.ado_analytiko_par.Recordset.Requery
    Me.ado_analytiko_par.Refresh
    Me.dt_analytiko_par.Refresh
    Me.ado_athl_sm.Recordset.Requery
    Me.ado_athl_sm.Refresh
    Me.dt_athl_sm.Refresh
    Call Option1_Click(mn)
    If Me.ado_analytiko_par.Recordset.RecordCount >= 1 Then
        Me.ado_analytiko_par.Recordset.MoveFirst
        Me.ado_analytiko_par.Recordset.Move rw - 1
    End If
    Me.dt_analytiko_par.Visible = True
    
End Sub

Private Sub tmp_poso_eggrafis_GotFocus()
    
    Me.tmp_poso_eggrafis.SelStart = 0
    Me.tmp_poso_eggrafis.SelLength = Len(Me.tmp_poso_eggrafis.Text)

End Sub

Private Sub tmp_poso_mina_GotFocus()
        
    Me.tmp_poso_mina.SelStart = 0
    Me.tmp_poso_mina.SelLength = Len(Me.tmp_poso_mina.Text)
    
End Sub

Private Sub txt_im_eis_Change()
    
    'Me.bt_ins_athl.Enabled = False
    'Me.bt_del_athl.Enabled = False
    'Me.bt_up_athl.Enabled = True
    'Me.bt_can_athl.Enabled = True

End Sub

Private Sub TriggerCheckbox(iRow As Integer, iCol As Integer)
        With MSFlexGrid1
            If .TextMatrix(iRow, iCol) = strUnChecked Then
                .TextMatrix(iRow, iCol) = strChecked
            Else
                .TextMatrix(iRow, iCol) = strUnChecked
            End If
        End With
End Sub
