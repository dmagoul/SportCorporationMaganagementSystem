VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form tmima_management 
   BackColor       =   &H00C0C0FF&
   ClientHeight    =   10995
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14760
   FillStyle       =   0  'Solid
   LinkTopic       =   "entity_management"
   MDIChild        =   -1  'True
   ScaleHeight     =   10995
   ScaleWidth      =   14760
   Begin MSDataGridLib.DataGrid dt_parousiologia 
      Bindings        =   "tmima_management.frx":0000
      Height          =   1815
      Left            =   9720
      TabIndex        =   71
      Top             =   6720
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3201
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
   Begin VB.CommandButton insert_bt 
      BackColor       =   &H80000014&
      Caption         =   "Προσ&θήκη"
      DisabledPicture =   "tmima_management.frx":0020
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      MaskColor       =   &H80000014&
      Picture         =   "tmima_management.frx":4DAF
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton save_command 
      BackColor       =   &H80000014&
      Caption         =   "&Αποθήκευση"
      DisabledPicture =   "tmima_management.frx":9B3E
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      MaskColor       =   &H80000014&
      Picture         =   "tmima_management.frx":E73D
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton bt_print 
      BackColor       =   &H80000014&
      Caption         =   "Εκτυπώ&σεις, Υπολογισμοί"
      DisabledPicture =   "tmima_management.frx":1333C
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      MaskColor       =   &H80000014&
      Picture         =   "tmima_management.frx":17D41
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   10200
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid dt_athl_tm_ids 
      Bindings        =   "tmima_management.frx":1C746
      Height          =   375
      Left            =   12240
      TabIndex        =   53
      Top             =   4080
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin MSAdodcLib.Adodc ado_athl_tm_ids 
      Height          =   330
      Left            =   11760
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton sear_bt 
      BackColor       =   &H80000014&
      Caption         =   "Α&ναζήτηση"
      DisabledPicture =   "tmima_management.frx":1C764
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      MaskColor       =   &H80000014&
      Picture         =   "tmima_management.frx":21B18
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton del_bt 
      BackColor       =   &H80000014&
      Caption         =   "&Διαγραφή"
      DisabledPicture =   "tmima_management.frx":26ECC
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      MaskColor       =   &H80000014&
      Picture         =   "tmima_management.frx":2BB7C
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000014&
      Caption         =   "&Καθαρισμός"
      DisabledPicture =   "tmima_management.frx":3082C
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      MaskColor       =   &H80000014&
      Picture         =   "tmima_management.frx":35370
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton up_bt 
      BackColor       =   &H80000014&
      Caption         =   "&Ενημέρωση"
      DisabledPicture =   "tmima_management.frx":39EB4
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      MaskColor       =   &H80000014&
      Picture         =   "tmima_management.frx":419AE
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton kl_bt 
      BackColor       =   &H80000014&
      Caption         =   "Κ&λείσιμο"
      DisabledPicture =   "tmima_management.frx":494A8
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13680
      MaskColor       =   &H80000014&
      Picture         =   "tmima_management.frx":4EF20
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton canc_bt 
      BackColor       =   &H80000014&
      Caption         =   "Ακύ&ρωση"
      DisabledPicture =   "tmima_management.frx":54998
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      MaskColor       =   &H80000014&
      Picture         =   "tmima_management.frx":598B5
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Τμήματα"
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
      Height          =   4935
      Left            =   120
      TabIndex        =   25
      Top             =   6000
      Width           =   11655
      Begin VB.CommandButton taksin 
         BackColor       =   &H80000014&
         Caption         =   "&Ταξινόμηση"
         DisabledPicture =   "tmima_management.frx":5E7D2
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MaskColor       =   &H80000014&
         Picture         =   "tmima_management.frx":6424A
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   4200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dt_tmimata 
         Bindings        =   "tmima_management.frx":68F51
         Height          =   3375
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   5953
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
      Begin MSAdodcLib.Adodc ado_tmimata 
         Height          =   375
         Left            =   120
         Top             =   3720
         Width           =   11460
         _ExtentX        =   20214
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=poseidon.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=poseidon.mdb;Mode=ReadWrite;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   $"tmima_management.frx":68F6B
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
      Begin MSAdodcLib.Adodc ado_tm_ids 
         Height          =   330
         Left            =   7080
         Top             =   4800
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
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
         RecordSource    =   "Τμήματα"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "tmima_management.frx":692D1
         Height          =   255
         Left            =   9600
         TabIndex        =   44
         Top             =   4800
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   450
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Height          =   6015
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   14655
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000014&
         Caption         =   "Εκτύπωση Παρουσιολογίου"
         DisabledPicture =   "tmima_management.frx":692EA
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9820
         MaskColor       =   &H80000014&
         Picture         =   "tmima_management.frx":6DCEF
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton cancel_cur_rec 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "Ακύρωση Τρέχουσας Ε&γγραφής"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7440
         MaskColor       =   &H00000000&
         Picture         =   "tmima_management.frx":726F4
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Ακύρωση Τρέχουσας Εγγραφής"
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         Caption         =   "Αθλητές  Τμήματος"
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
         Height          =   3375
         Left            =   4920
         TabIndex        =   43
         Top             =   1680
         Width           =   9615
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "tmima_management.frx":728B3
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9160
            MaskColor       =   &H00008000&
            Picture         =   "tmima_management.frx":74F3E
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Εκτύπωση Προϋπολογισμού Εσόδων"
            Top             =   2400
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "tmima_management.frx":775C9
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9160
            MaskColor       =   &H00008000&
            Picture         =   "tmima_management.frx":7BFCE
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Εκτύπωση Προϋπολογισμού Εσόδων"
            Top             =   2040
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton bt_can_athl 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            Caption         =   "<-"
            DisabledPicture =   "tmima_management.frx":809D3
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9160
            MaskColor       =   &H000000FF&
            Picture         =   "tmima_management.frx":80C91
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Αναίρεση τελευταίας ενέργειας"
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton bt_del_athl 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Caption         =   "x"
            DisabledPicture =   "tmima_management.frx":80F4F
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9160
            MaskColor       =   &H000000FF&
            Picture         =   "tmima_management.frx":8106D
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Διαγραφή Επιλεγμένου Αθλητή από το Τρέχων Τμήμα"
            Top             =   1320
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton bt_up_athl 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "V"
            DisabledPicture =   "tmima_management.frx":81624
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9160
            MaskColor       =   &H000000FF&
            Picture         =   "tmima_management.frx":82D49
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Ενημέρωση στη ΒΔ των Αλλαγών"
            Top             =   960
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton bt_ins_athl 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "+"
            DisabledPicture =   "tmima_management.frx":8446E
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9160
            MaskColor       =   &H000000FF&
            Picture         =   "tmima_management.frx":8462D
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Προσθήκη Νέου Αθλητή στο Τρέχων Τμήμα"
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CheckBox ch_ib 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   8160
            TabIndex        =   23
            Top             =   300
            Width           =   255
         End
         Begin VB.TextBox Text1 
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
            Height          =   345
            Left            =   7400
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   240
            Width           =   1455
         End
         Begin MSAdodcLib.Adodc ado_athl 
            Height          =   375
            Left            =   8400
            Top             =   0
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
            RecordSource    =   "ΓιαΤηνΕισαγωγήΑθλητώνΣεΤμήμα"
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.TextBox txt_pm 
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
            Height          =   345
            Left            =   6130
            TabIndex        =   22
            Top             =   260
            Width           =   1250
         End
         Begin VB.TextBox txt_pe 
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
            Height          =   345
            Left            =   4880
            TabIndex        =   21
            Top             =   260
            Width           =   1245
         End
         Begin MSDataListLib.DataCombo co_athl 
            Bindings        =   "tmima_management.frx":84CB8
            Height          =   345
            Left            =   435
            TabIndex        =   19
            Top             =   240
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   609
            _Version        =   393216
            IntegralHeight  =   0   'False
            Locked          =   -1  'True
            ListField       =   "ΑΜOE"
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataGridLib.DataGrid dt_athl_tm 
            Bindings        =   "tmima_management.frx":84CCF
            Height          =   2295
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   8985
            _ExtentX        =   15849
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
            HeadLines       =   1
            RowHeight       =   18
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
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
         Begin MSAdodcLib.Adodc ado_athl_tm 
            Height          =   375
            Left            =   110
            Top             =   2880
            Width           =   9030
            _ExtentX        =   15928
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
            RecordSource    =   "ΑθλητέςΤμήματα"
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSMask.MaskEdBox txt_im_eis 
            Height          =   345
            Left            =   3440
            TabIndex        =   20
            Top             =   260
            Width           =   1450
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label11 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   345
            Left            =   240
            TabIndex        =   62
            Top             =   390
            Width           =   135
         End
         Begin VB.Label Label9 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   240
            TabIndex        =   61
            Top             =   180
            Width           =   255
         End
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
         TabIndex        =   35
         Top             =   360
         Width           =   7095
         Begin MSMask.MaskEdBox im_l 
            Height          =   300
            Index           =   1
            Left            =   1800
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
            Index           =   0
            Left            =   840
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
            Index           =   2
            Left            =   2760
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
            Index           =   3
            Left            =   3720
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
         Begin MSMask.MaskEdBox im_l 
            Height          =   300
            Index           =   4
            Left            =   4680
            TabIndex        =   15
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
            TabIndex        =   17
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
            Index           =   1
            Left            =   1800
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
            Index           =   2
            Left            =   2760
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
            Index           =   3
            Left            =   3720
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
         Begin MSMask.MaskEdBox im_en 
            Height          =   300
            Index           =   4
            Left            =   4680
            TabIndex        =   14
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
            TabIndex        =   16
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   45
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
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
         Height          =   4695
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   4935
         Begin VB.TextBox tmp_poso_mina_plain 
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
            Height          =   375
            Left            =   3000
            TabIndex        =   56
            Top             =   2760
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox tmp_poso_eggrafis_plain 
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
            Height          =   375
            Left            =   3000
            TabIndex        =   55
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
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
            Height          =   375
            Left            =   1560
            TabIndex        =   5
            Top             =   2760
            Width           =   1335
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
            Height          =   375
            Left            =   1560
            TabIndex        =   4
            Top             =   2280
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo co_ae 
            Bindings        =   "tmima_management.frx":84CE9
            Height          =   345
            Left            =   1560
            TabIndex        =   0
            Top             =   360
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   609
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "περιγραφή"
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSAdodcLib.Adodc ado_ae 
            Height          =   735
            Left            =   2280
            Top             =   0
            Visible         =   0   'False
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   1296
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
            RecordSource    =   "Αθλητικά_Έτη"
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataListLib.DataCombo co_kt 
            Bindings        =   "tmima_management.frx":84CFE
            Height          =   345
            Left            =   1560
            TabIndex        =   1
            Top             =   840
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   609
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "περιγραφή"
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSAdodcLib.Adodc ado_kt 
            Height          =   375
            Left            =   3480
            Top             =   840
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=poseidon.mdb;Mode=ReadWrite;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=poseidon.mdb;Mode=ReadWrite;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   $"tmima_management.frx":84D13
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataListLib.DataCombo co_propa 
            Bindings        =   "tmima_management.frx":84DAA
            Height          =   345
            Left            =   1560
            TabIndex        =   2
            Top             =   1320
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   609
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "OE"
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSAdodcLib.Adodc ado_prop 
            Height          =   375
            Left            =   3480
            Top             =   1320
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
            RecordSource    =   "Ονοματεπώνυμα_Προπονητών"
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataListLib.DataCombo co_propb 
            Bindings        =   "tmima_management.frx":84DC1
            Height          =   345
            Left            =   1560
            TabIndex        =   3
            Top             =   1800
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   609
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "OE"
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Left            =   120
            TabIndex        =   42
            Top             =   1800
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
            Left            =   240
            TabIndex        =   41
            Top             =   2760
            Width           =   1215
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
            Left            =   0
            TabIndex        =   40
            Top             =   2280
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
            Left            =   120
            TabIndex        =   39
            Top             =   1320
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
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   1335
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
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton parousies 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Καταχώρηση Παρουσιών"
         DisabledPicture =   "tmima_management.frx":84DD8
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11400
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tmima_management.frx":8A850
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000009&
         Height          =   345
         Left            =   10680
         TabIndex        =   57
         Top             =   2040
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc ado_parousiologia 
      Height          =   375
      Left            =   11880
      Top             =   6240
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
End
Attribute VB_Name = "tmima_management"
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

Private Sub ado_athl_tm_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    'If Not adReason = 10 And pRecordset.AbsolutePosition >= 1 Then
    If pRecordset.AbsolutePosition >= 1 Then
        If Not pRecordset.EOF Then
          If Trim(Me.ado_athl_tm.Recordset.Fields(1).Value) <> "" Then
                'If Not rs_ado_athl.EOF Then
                If Not Me.ado_athl.Recordset.EOF Then
                    Me.ado_athl.Recordset.MoveFirst
                    Me.ado_athl.Recordset.Find "[id] = '" & Me.ado_athl_tm.Recordset.Fields(1).Value & "'"
                    If Not Me.ado_athl.Recordset.EOF Then
                        co_athl.Text = Me.ado_athl.Recordset.Fields(1).Value
                    End If
                End If
            Else
                co_athl.Text = ""
            End If
            If Trim(Me.ado_athl_tm.Recordset.Fields(4).Value) <> "" Then
                Me.txt_im_eis.Text = Me.ado_athl_tm.Recordset.Fields(4).Value
            Else
                Me.txt_im_eis.Text = "00/00/0000"
            End If
            If Trim(Me.ado_athl_tm.Recordset.Fields(5).Value) <> "" Then
                Me.txt_pe.Text = Me.ado_athl_tm.Recordset.Fields(5).Value
            Else
                Me.txt_pe.Text = ""
            End If
            If Trim(Me.ado_athl_tm.Recordset.Fields(6).Value) <> "" Then
                Me.txt_pm.Text = Me.ado_athl_tm.Recordset.Fields(6).Value
            Else
                Me.txt_pm.Text = ""
            End If
            If Trim(Me.ado_athl_tm.Recordset.Fields(8).Value) <> "" Then
                If Me.ado_athl_tm.Recordset.Fields(8).Value Like "NAI" Then
                    Me.ch_ib.Value = 1
                Else
                    Me.ch_ib.Value = 0
                End If
            Else
                Me.ch_ib.Value = 0
            End If
        End If
    End If
    If pRecordset.RecordCount >= 1 Then
        Me.ado_athl_tm.Caption = "Αθλητής " & pRecordset.AbsolutePosition & " από " & pRecordset.RecordCount
    Else
        Me.ado_athl_tm.Caption = "Αθλητής " & 0 & " από " & 0
    End If
    bt_ins_athl.Enabled = True
    bt_up_athl.Enabled = False
    bt_del_athl.Enabled = True
    
End Sub

Private Sub ado_tmimata_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    If Not adReason = 10 And pRecordset.AbsolutePosition >= 1 Then
        If Not pRecordset.EOF Then
            If Trim(pRecordset.Fields(0).Value) <> "" Then
                If Not Me.ado_ae.Recordset.EOF Then
                    Me.ado_ae.Recordset.MoveFirst
                    Me.ado_ae.Recordset.Find "[id_αθλητικού_έτους] = '" & pRecordset.Fields(0).Value & "'"
                    If Not Me.ado_ae.Recordset.EOF Then
                        co_ae.Text = Me.ado_ae.Recordset.Fields(1).Value
                    End If
                End If
            Else
                co_ae.Text = ""
            End If
            If Trim(pRecordset.Fields(2).Value) <> "" Then
                If Not rs_ado_kt.EOF Then
                    rs_ado_kt.MoveFirst
                   rs_ado_kt.Find "[id_κατηγορίας_τμήματος] = '" & pRecordset.Fields(2).Value & "'"
                    If Not rs_ado_kt.EOF Then
                        co_kt.Text = rs_ado_kt.Fields(1).Value
                    End If
                End If
            Else
                co_kt.Text = ""
            End If
            If Trim(pRecordset.Fields(4).Value) <> "" Then
                If Not rs_ado_prop.EOF Then
                    rs_ado_prop.MoveFirst
                    rs_ado_prop.Find "[id] = '" & pRecordset.Fields(4).Value & "'"
                    If Not rs_ado_prop.EOF Then
                        co_propa.Text = rs_ado_prop.Fields(1).Value
                    End If
                End If
            Else
                tmp_propA.Text = ""
            End If
            If Trim(pRecordset.Fields(6).Value) <> "" Then
                If Not rs_ado_prop.EOF Then
                    rs_ado_prop.MoveFirst
                    rs_ado_prop.Find "[id] = '" & pRecordset.Fields(6).Value & "'"
                    If Not rs_ado_prop.EOF Then
                        co_propb.Text = rs_ado_prop.Fields(1).Value
                    End If
                End If
            Else
                co_propb.Text = ""
            End If
            If Trim(Me.ado_tmimata.Recordset.Fields(8).Value) <> "" Then
                Me.tmp_poso_eggrafis_plain.Text = Trim(str(pRecordset.Fields(8).Value))
                Me.tmp_poso_eggrafis.Text = Trim(str(pRecordset.Fields(8).Value)) + ",00 €"
            Else
                tmp_poso_eggrafis.Text = ""
            End If
            If Trim(Me.ado_tmimata.Recordset.Fields(9).Value) <> "" Then
                'Me.tmp_poso_mina.Text = Str(rs_ado_tmimata.Fields(9).Value)
                Me.tmp_poso_mina_plain.Text = Trim(str(pRecordset.Fields(9).Value))
                Me.tmp_poso_mina.Text = Trim(str(pRecordset.Fields(9).Value)) + ",00 €"
            Else
                Me.tmp_poso_mina.Text = ""
            End If
            'ΕΝΗΜΕΡΩΣΗ ΤΟΥ ΕΒΔΟΜΑΔΙΑΙΟΥ ΠΡΟΓΡΑΜΜΑΤΟΣ
            i = 0
            For j = 0 To 5
                If Trim(pRecordset.Fields(11 + i).Value) <> "" Then
                    Me.im_en(j).Text = pRecordset.Fields(11 + i).Value
                Else
                    Me.im_en(j).Text = "00:00"
                End If
                If Trim(pRecordset.Fields(12 + i).Value) <> "" Then
                    Me.im_l(j).Text = pRecordset.Fields(12 + i).Value
                Else
                    Me.im_l(j).Text = "00:00"
                End If
                i = i + 2
            Next j
        End If
        'REFRESH του datagrid ΑθλητέςΤμήματα -- BEGIN
        Me.Label9.Caption = ""
        Me.Label11.Caption = ""
        Me.ado_athl_tm.Recordset.Filter = "[id_Τμήματος] = '" & pRecordset.Fields(10).Value & "'"
        If Not Me.ado_athl_tm.Recordset.EOF Then
        If Me.ado_athl_tm.Recordset.RecordCount > 0 Then
            Me.co_athl.Enabled = True
            Me.txt_im_eis.Enabled = True
            Me.txt_pe.Enabled = True
            Me.txt_pm.Enabled = True
            Me.ch_ib.Enabled = True
            Me.dt_athl_tm.Row = 0
            Me.dt_athl_tm.Col = 1
            Me.ado_athl_tm.Caption = "Αθλητής " & Me.dt_athl_tm.Row + 1 & " από " & Me.ado_athl_tm.Recordset.RecordCount
        Else
            If Me.ado_athl_tm.Recordset.RecordCount = 0 Then
                Me.ado_athl_tm.Caption = "Αθλητής " & 0 & " από " & 0
                Me.co_athl.Enabled = False
                Me.bt_ins_athl.Enabled = True
                Me.txt_im_eis.Enabled = False
                Me.txt_pe.Enabled = False
                Me.txt_pm.Enabled = False
                Me.ch_ib.Enabled = False
                Me.bt_up_athl.Enabled = False
            End If
        End If
        If Trim(Me.ado_athl_tm.Recordset.Fields(1).Value) <> "" Then
                    If Not rs_ado_athl.EOF Then
                        rs_ado_athl.MoveFirst
                        rs_ado_athl.Find "[id] = '" & Me.ado_athl_tm.Recordset.Fields(1).Value & "'"
                        If Not rs_ado_athl.EOF Then
                            co_athl.Text = rs_ado_athl.Fields(1).Value
                        End If
                    End If
                Else
                    co_athl.Text = ""
                End If
                If Trim(Me.ado_athl_tm.Recordset.Fields(4).Value) <> "" Then
                    Me.txt_im_eis.Text = Trim(Me.ado_athl_tm.Recordset.Fields(4).Value)
                Else
                    Me.txt_im_eis.Text = ""
                End If
                If Trim(Me.ado_athl_tm.Recordset.Fields(5).Value) <> "" Then
                    Me.txt_pe.Text = Trim(Me.ado_athl_tm.Recordset.Fields(5).Value)
                Else
                    Me.txt_pe.Text = ""
                End If
                If Trim(Me.ado_athl_tm.Recordset.Fields(6).Value) <> "" Then
                    Me.txt_pm.Text = Trim(Me.ado_athl_tm.Recordset.Fields(6).Value)
                Else
                    Me.txt_pm.Text = ""
                End If
                
        Me.dt_athl_tm.Columns(0).Visible = False
        Me.dt_athl_tm.Columns(1).Visible = False
        Me.dt_athl_tm.Columns(2).Caption = "Α.Μ. Αθλητή"
        Me.dt_athl_tm.Columns(2).Width = 1000
        Me.dt_athl_tm.Columns(3).Caption = "Ονοματεπώνυμο Αθλητή"
        Me.dt_athl_tm.Columns(3).Width = 2000
        Me.dt_athl_tm.Columns(4).Caption = "Ημ/νία Εισαγωγής"
        Me.dt_athl_tm.Columns(4).Width = 1450
        Me.dt_athl_tm.Columns(5).Caption = "Ποσό Εγγραφής"
        Me.dt_athl_tm.Columns(5).Width = 1250
        Me.dt_athl_tm.Columns(6).Caption = "Ποσό Μήνα"
        Me.dt_athl_tm.Columns(6).Width = 1250
        Me.dt_athl_tm.Columns(7).Visible = False
        Me.dt_athl_tm.Columns(8).Caption = "Ιατρική Βεβαίωση"
        Me.dt_athl_tm.Columns(8).Width = 1450
        For i = 9 To 23
            Me.dt_athl_tm.Columns(i).Visible = False
        Next i
        'REFRESH του datagrid ΑθλητέςΤμήματα -- END
        End If
    End If
    '*******************
    Me.ado_tmimata.Caption = "Τμήμα " & pRecordset.AbsolutePosition & " από " & pRecordset.RecordCount

End Sub

Private Sub bt_can_athl_Click()

    If Not Me.ado_athl_tm.Recordset.EOF Then
        pr_c = Me.ado_athl_tm.Recordset.AbsolutePosition
        Me.ado_athl_tm.Recordset.MoveFirst
        If pr_c >= 1 Then
            Me.ado_athl_tm.Recordset.Move pr_c - 1
        End If
        Me.co_athl.Locked = True
        Me.bt_up_athl.Enabled = False
        Me.bt_ins_athl.Enabled = True
        Me.bt_del_athl.Enabled = True
    Else
        Me.co_athl.Text = ""
        Me.co_athl.Enabled = False
        Me.txt_im_eis.Text = "00/00/0000"
        Me.txt_im_eis.Enabled = False
        Me.txt_pe.Text = ""
        Me.txt_pe.Enabled = False
        Me.txt_pm.Text = ""
        Me.txt_pm.Enabled = False
        Me.ch_ib.Value = 0
        Me.ch_ib.Enabled = False
        Me.co_athl.Locked = True
        Me.bt_up_athl.Enabled = False
        Me.bt_ins_athl.Enabled = True
        Me.bt_del_athl.Enabled = False
    End If
    new_addition = 0
    Me.bt_can_athl.Enabled = False

End Sub

Private Sub bt_del_athl_Click()

    If Not Me.ado_athl_tm.Recordset.EOF Then
        ms = MsgBox("Είσαι σίγουρος; Μαζί με τον Αθλητή θα ΔΙΑΓΡΑΦΟΥΝ και όλες οι σχετικές μ' αυτόν πληροφορίες (π.χ. παρουσίες στο τμήμα). Επέλεξε ΝΑΙ ή ΟΧΙ ...", vbYesNo, "Παράθυρο διαγραφής")
        If ms = 6 Then
            Dim tm As String
            Dim athl As String
            Dim cur_id As String
            tm = Trim(str(Me.ado_tmimata.Recordset.Fields(10).Value))
            athl = Trim(str(Me.ado_athl_tm.Recordset.Fields(1).Value))
            cur_id = Trim(str(Me.ado_athl_tm.Recordset.Fields(0).Value))
            Me.ado_athl_tm_ids.Recordset.MoveFirst
            Me.ado_athl_tm_ids.Recordset.Find "[id] = '" & Val(cur_id) & "'"
            col_affected = Me.ado_athl_tm.Recordset.AbsolutePosition
            If Me.ado_athl_tm.Recordset.AbsolutePosition = Me.ado_athl_tm.Recordset.RecordCount Then
                col_affected = col_affected - 1
            End If
            'rs_ado_athl_tm_ids.Delete adAffectCurrent
            Me.ado_athl_tm_ids.Recordset.Delete adAffectCurrent
            Me.ado_athl_tm_ids.Recordset.MoveNext
            Me.ado_athl_tm_ids.Recordset.Requery
            Me.ado_athl_tm_ids.Refresh
            Me.dt_athl_tm_ids.Refresh
            Me.ado_athl_tm.Recordset.Requery
            Me.ado_athl_tm.Refresh
            Me.dt_athl_tm.Refresh
            '
            Set Me.dt_athl_tm.DataSource = Me.ado_athl_tm
            '
            Me.ado_athl_tm.Recordset.Filter = "[id_Τμήματος] = '" & Me.ado_tmimata.Recordset.Fields(10).Value & "'"
            If Not Me.ado_athl_tm.Recordset.EOF Then
                Me.ado_athl_tm.Recordset.MoveFirst
                If col_affected - 1 >= 1 Then
                    Me.ado_athl_tm.Recordset.Move col_affected - 1
                End If
                If col_affected - 1 >= 0 Then
                    'Me.dt_athl_tm.Row = col_affected - 1
                    Me.dt_athl_tm.SelBookmarks.Add Me.ado_athl_tm.Recordset.Bookmark
                Else
                    'Me.dt_athl_tm.Row = 0
                    Me.dt_athl_tm.SelBookmarks.Add Me.ado_athl_tm.Recordset.Bookmark
                End If
                Me.dt_athl_tm.Col = 1
            Else
                Me.co_athl.Text = ""
                Me.txt_im_eis.Text = "00/00/0000"
                Me.txt_pe.Text = ""
                Me.txt_pm.Text = ""
                Me.ch_ib.Value = 0
            End If
            
            Me.dt_athl_tm.Columns(0).Visible = False
            Me.dt_athl_tm.Columns(1).Visible = False
            Me.dt_athl_tm.Columns(2).Caption = "Α.Μ. Αθλητή"
            Me.dt_athl_tm.Columns(2).Width = 1000
            Me.dt_athl_tm.Columns(3).Caption = "Ονοματεπώνυμο Αθλητή"
            Me.dt_athl_tm.Columns(3).Width = 2000
            Me.dt_athl_tm.Columns(4).Caption = "Ημ/νία Εισαγωγής"
            Me.dt_athl_tm.Columns(4).Width = 1500
            Me.dt_athl_tm.Columns(5).Caption = "Ποσό Εγγραφής"
            Me.dt_athl_tm.Columns(5).Width = 1500
            Me.dt_athl_tm.Columns(6).Caption = "Ποσό Μήνα"
            Me.dt_athl_tm.Columns(6).Width = 1500
            Me.dt_athl_tm.Columns(7).Visible = False
            
            ' ΔΙΑΓΡΑΦΗ ΠΑΡΟΥΣΙΩΝ ΣΤΟ ΤΜΗΜΑ
            Me.ado_parousiologia.Recordset.Requery
            Me.ado_parousiologia.Refresh
            Me.ado_parousiologia.Recordset.Filter = "[id_tmimatos] like '" & str(tm) & "'" & " AND [id_athliti] like '" & str(athl) & "'"
            cn = Me.ado_parousiologia.Recordset.RecordCount
            If cn >= 1 Then
                Me.ado_parousiologia.Recordset.MoveLast
                For i = 1 To cn
                    Me.ado_parousiologia.Recordset.Delete adAffectCurrent
                    If i <> cn Then
                        Me.ado_parousiologia.Recordset.MovePrevious
                    End If
                Next i
            Me.ado_parousiologia.Recordset.Requery
            End If
            '
        Else
            MsgBox "Ακύρωση Διαγραφής!"
            'Me.ado_tm_ids.Recordset.Delete adAffectCurrent
        End If
    Else
        MsgBox "Δεν υπάρχει εγγραφή προς ΔΙΑΓΡΑΦΗ!", vbCritical, "Μήνυμα Λάθους"
    End If
    
End Sub

Private Sub bt_ins_athl_Click()

If a_new_tm_is_ready_to_add = 0 Then
    new_addition = 1
    Me.co_athl.Text = ""
    Me.co_athl.Enabled = True
    Me.txt_im_eis.Text = "00/00/0000"
    Me.txt_im_eis.Enabled = True
    'Me.txt_pe.Text = ""
    Me.txt_pe.Text = Me.tmp_poso_eggrafis
    Me.txt_pe.Enabled = True
    'Me.txt_pm.Text = ""
    Me.txt_pm.Text = Me.tmp_poso_mina
    Me.txt_pm.Enabled = True
    Me.ch_ib.Value = 0
    Me.ch_ib.Enabled = True
    Me.Label9.Caption = ""
    
    Me.bt_del_athl.Enabled = False
    Me.co_athl.Enabled = True
    Me.bt_up_athl.Enabled = False
    Me.bt_can_athl.Enabled = True
    Me.co_athl.Locked = False
Else
    MsgBox "Αδύνατη η προσθήκη ΑΘΛΗΤΩΝ σε ΤΜΗΜΑ που τώρα δημιουργείται! Αποθηκεύστε πρώτα το τμήμα και ξαναπροσπαθήστε!"
End If
    
End Sub

Private Sub bt_print_Click()
    
    f_n = Me.Name
    frm_pr_tm.Show

End Sub

Private Sub bt_up_athl_Click()

    'ΠΡΟΣΘΗΚΗ ΝΕΟΥ ΑΘΛΗΤΗ ΣΕ ΤΜΗΜΑ
    If new_addition = 1 Then
        new_addition = 0
        Me.co_athl.Locked = True
        Me.bt_up_athl.Enabled = False
        Me.bt_ins_athl.Enabled = True
        Me.bt_del_athl.Enabled = True
        Me.bt_can_athl.Enabled = False
        'ΝΑ ΕΛΕΓΞΩ ΑΝ Ο ΑΘΛΗΤΗΣ ΕΙΝΑΙ ΗΔΗ ΣΤΟ ΤΜΗΜΑ
        Dim nm As String
        Dim kl, ib As Integer
        Dim im_eis, pe, pm As String
        im_eis = Me.txt_im_eis.Text
        pe = Me.txt_pe.Text
        pm = Me.txt_pm.Text
        ib = Me.ch_ib.Value
        nm = Trim(Me.co_athl.Text)
        If ado_athl.Recordset.RecordCount >= 1 Then
            ado_athl.Recordset.MoveFirst
            ado_athl.Recordset.Find "[ΑΜOE] Like '" & Trim(nm) & "'"
            kl = ado_athl.Recordset.Fields(0).Value
        End If
        If ado_athl_tm.Recordset.RecordCount >= 1 Then
            Me.ado_athl_tm.Recordset.MoveFirst
            Me.ado_athl_tm.Recordset.Find "[id_Αθλητή] = '" & kl & "'"
            If Not Me.ado_athl_tm.Recordset.EOF Then
                MsgBox "Αδύνατη η εισαγωγή του αθλητή στο τμήμα ... ο αθλητής υπάρχει ήδη ..."
                Exit Sub
            End If
        End If
        '
        'Να βρω το υποψήφιο id για την αποθήκευση στον πίνακα ΑΘΛΗΤΕΣ_ΤΜΗΜΑΤΩΝ
        Me.ado_athl_tm_ids.Recordset.Sort = "[" & Trim(Me.ado_athl_tm_ids.Recordset.Fields(0).Name) & "]"
        If Not Me.ado_athl_tm_ids.Recordset.EOF Then
            Me.ado_athl_tm_ids.Recordset.MoveLast
            id_αθλ_τμ = Me.ado_athl_tm_ids.Recordset.Fields(0).Value
            id_αθλ_τμ = id_αθλ_τμ + 1
        End If
        'Αποθήκευση στα ΑθλητέςΤμημάτων
        Me.ado_athl_tm_ids.Recordset.AddNew
        'Αποθήκευση id
        Me.ado_athl_tm_ids.Recordset.Fields(0).Value = id_αθλ_τμ
        'Αποθήκευση id_Τμήματος
        If a_new_tm_is_ready_to_add = 1 Then
            'rs_ado_athl_tm_ids.Fields(1).Value = id_τμ
            MsgBox "Αδύνατη η προσθήκη ΑΘΛΗΤΩΝ σε ΤΜΗΜΑ που τώρα δημιουργείται! Αποθηκεύστε πρώτα το τμήμα και ξαναπροσπαθήστε!"
            Exit Sub
        Else
            Me.ado_athl_tm_ids.Recordset.Fields(1).Value = Me.ado_tmimata.Recordset.Fields(10).Value
        End If
        'Αποθήκευση id_Αθλητή
        If Trim(Me.co_athl.Text) <> "" Then
            If Me.co_athl.SelectedItem >= 1 Then
                'rs_ado_athl.MoveFirst
                'rs_ado_athl.Move Me.co_athl.SelectedItem - 1
                'rs_ado_athl_tm_ids.Fields(2).Value = rs_ado_athl.Fields(0).Value
                Me.ado_athl_tm_ids.Recordset.Fields(2).Value = kl
            End If
        Else
            Me.ado_athl_tm_ids.Recordset.Fields(2).Value = ""
        End If
        If Trim(im_eis) <> "" Then
            Me.ado_athl_tm_ids.Recordset.Fields(3).Value = Trim(im_eis)
        End If
        If Trim(pe) <> "" Then
            Me.ado_athl_tm_ids.Recordset.Fields(4).Value = Trim(pe)
        End If
        If Trim(pm) <> "" Then
            Me.ado_athl_tm_ids.Recordset.Fields(5).Value = Trim(pm)
        End If
        Me.ado_athl_tm_ids.Recordset.Fields(6).Value = ib
        'Me.ado_athl_tm_ids.Recordset.UpdateBatch adAffectCurrent
        Me.ado_athl_tm_ids.Recordset.UpdateBatch adAffectAllChapters
        Me.ado_athl_tm_ids.Recordset.Requery
        Me.ado_athl_tm_ids.Refresh
        Me.dt_athl_tm_ids.Refresh
        Me.ado_athl_tm.Recordset.Requery
        Me.ado_athl_tm.Refresh
        Me.dt_athl_tm.Refresh
        '
        Set rs_ado_athl_tm_ids = Me.ado_athl_tm_ids.Recordset
        Set Me.dt_athl_tm_ids.DataSource = Me.ado_athl_tm_ids
        '
        Set rs_ado_athl_tm = Me.ado_athl_tm.Recordset
        Set Me.dt_athl_tm.DataSource = Me.ado_athl_tm
        '
        Me.ado_athl_tm.Recordset.Filter = "[id_Τμήματος] = '" & Me.ado_tmimata.Recordset.Fields(10).Value & "'"
        If Not Me.ado_athl_tm.Recordset.EOF Then
            Me.ado_athl_tm.Recordset.MoveFirst
            Me.ado_athl_tm.Recordset.Find "[id_Αθλητή] = '" & kl & "'"
            Me.dt_athl_tm.SelBookmarks.Add Me.ado_athl_tm.Recordset.Bookmark
            Me.dt_athl_tm.Col = 1
            If Me.ado_athl_tm.Recordset.RecordCount > 0 Then
                Me.ado_athl_tm.Caption = "Αθλητής " & Me.ado_athl_tm.Recordset.AbsolutePosition & " από " & Me.ado_athl_tm.Recordset.RecordCount
            End If
        End If
        '
        Me.dt_athl_tm.Columns(0).Visible = False
        Me.dt_athl_tm.Columns(1).Visible = False
        Me.dt_athl_tm.Columns(2).Caption = "Α.Μ. Αθλητή"
        Me.dt_athl_tm.Columns(2).Width = 1000
        Me.dt_athl_tm.Columns(3).Caption = "Ονοματεπώνυμο Αθλητή"
        Me.dt_athl_tm.Columns(3).Width = 2000
        Me.dt_athl_tm.Columns(4).Caption = "Ημ/νία Εισαγωγής"
        Me.dt_athl_tm.Columns(4).Width = 1450
        Me.dt_athl_tm.Columns(5).Caption = "Ποσό Εγγραφής"
        Me.dt_athl_tm.Columns(5).Width = 1250
        Me.dt_athl_tm.Columns(6).Caption = "Ποσό Μήνα"
        Me.dt_athl_tm.Columns(6).Width = 1250
        Me.dt_athl_tm.Columns(7).Visible = False
        Me.dt_athl_tm.Columns(8).Caption = "Ιατρική Βεβαίωση"
        Me.dt_athl_tm.Columns(8).Width = 1450
        For i = 9 To 23
            Me.dt_athl_tm.Columns(i).Visible = False
        Next i
    'ΤΕΛΟΣ --- ΠΡΟΣΘΗΚΗ ΝΕΟΥ ΑΘΛΗΤΗ ΣΕ ΤΜΗΜΑ
    Else
    'ΑΡΧΗ --- ΕΝΗΜΕΡΩΣΗ ΠΑΛΙΟΥ ΑΘΛΗΤΗ ΤΜΗΜΑΤΟΣ
        Me.bt_del_athl.Enabled = True
        If Not Me.ado_athl_tm.Recordset.EOF And Not Me.ado_athl_tm_ids.Recordset.EOF Then
            Me.ado_athl_tm_ids.Recordset.Filter = "[id_Τμήματος] = '" & Me.ado_athl_tm.Recordset.Fields(7).Value & "'"
            Me.ado_athl_tm_ids.Recordset.Find "[id_Αθλητή] = '" & Me.ado_athl_tm.Recordset.Fields(1).Value & "'"
            If Trim(Me.txt_im_eis.Text) <> "" Then
                Me.ado_athl_tm_ids.Recordset.Fields(3).Value = Trim(Me.txt_im_eis.Text)
            End If
            If Trim(Me.txt_pe.Text) <> "" Then
                Me.ado_athl_tm_ids.Recordset.Fields(4).Value = Trim(txt_pe.Text)
            End If
            If Trim(Me.txt_pm.Text) <> "" Then
                Me.ado_athl_tm_ids.Recordset.Fields(5).Value = Trim(txt_pm.Text)
            End If
            Me.ado_athl_tm_ids.Recordset.Fields(6).Value = Me.ch_ib.Value
            col_affected = Me.ado_athl_tm.Recordset.AbsolutePosition
            Me.ado_athl_tm_ids.Recordset.UpdateBatch adAffectCurrent
            Me.ado_athl_tm_ids.Recordset.Requery
            Me.ado_athl_tm_ids.Refresh
            Me.dt_athl_tm_ids.Refresh
    
            Me.ado_athl_tm.Recordset.Requery
            Me.ado_athl_tm.Refresh
            Me.dt_athl_tm.Refresh
            
            Me.ado_athl_tm.Recordset.Filter = "[id_Τμήματος] = '" & Me.ado_tmimata.Recordset.Fields(10).Value & "'"
            Me.ado_athl_tm.Recordset.MoveFirst
            If col_affected > 1 Then
                Me.ado_athl_tm.Recordset.Move col_affected - 1
            End If
            Me.dt_athl_tm.SelBookmarks.Add Me.ado_athl_tm.Recordset.Bookmark
            Me.dt_athl_tm.Col = 1
            If Me.ado_athl_tm.Recordset.RecordCount > 0 Then
                Me.ado_athl_tm.Caption = "Αθλητής " & Me.ado_athl_tm.Recordset.AbsolutePosition & " από " & Me.ado_athl_tm.Recordset.RecordCount
            End If
            
            Me.dt_athl_tm.Columns(0).Visible = False
            Me.dt_athl_tm.Columns(1).Visible = False
            Me.dt_athl_tm.Columns(2).Caption = "Α.Μ. Αθλητή"
            Me.dt_athl_tm.Columns(2).Width = 1000
            Me.dt_athl_tm.Columns(3).Caption = "Ονοματεπώνυμο Αθλητή"
            Me.dt_athl_tm.Columns(3).Width = 2000
            Me.dt_athl_tm.Columns(4).Caption = "Ημ/νία Εισαγωγής"
            Me.dt_athl_tm.Columns(4).Width = 1450
            Me.dt_athl_tm.Columns(5).Caption = "Ποσό Εγγραφής"
            Me.dt_athl_tm.Columns(5).Width = 1250
            Me.dt_athl_tm.Columns(6).Caption = "Ποσό Μήνα"
            Me.dt_athl_tm.Columns(6).Width = 1250
            Me.dt_athl_tm.Columns(7).Visible = False
            Me.dt_athl_tm.Columns(8).Caption = "Ιατρική Βεβαίωση"
            Me.dt_athl_tm.Columns(8).Width = 1450
            For i = 9 To 23
                Me.dt_athl_tm.Columns(i).Visible = False
            Next i
        End If
    'ΤΕΛΟΣ --- ΕΝΗΜΕΡΩΣΗ ΠΑΛΙΟΥ ΑΘΛΗΤΗ ΤΜΗΜΑΤΟΣ
    End If
    
End Sub

Private Sub canc_bt_Click()

    a_new_tm_is_ready_to_add = 0
    Me.insert_bt.Enabled = True
    Me.up_bt.Enabled = True
    Me.sear_bt.Enabled = False
    
    Me.ado_tmimata.Refresh
    
    Me.ado_tmimata.Recordset.Sort = s_sort
    
    If Not Me.ado_tmimata.Recordset.EOF Then
        Me.ado_tmimata.Recordset.MoveFirst
    Else
        co_ae.Text = ""
        co_kt.Text = ""
        co_propa.Text = ""
        co_propb.Text = ""
        tmp_poso_eggrafis.Text = ""
        tmp_poso_mina.Text = ""
    End If
    
    Me.dt_tmimata.Columns(0).Visible = False
    Me.dt_tmimata.Columns(1).Caption = "Αθλητικό Έτος"
    Me.dt_tmimata.Columns(1).Width = 1500
    Me.dt_tmimata.Columns(2).Visible = False
    Me.dt_tmimata.Columns(3).Caption = "Κατηγορία Τμήματος"
    Me.dt_tmimata.Columns(3).Width = 2000
    Me.dt_tmimata.Columns(4).Visible = False
    Me.dt_tmimata.Columns(5).Caption = "Α Προπονητής"
    Me.dt_tmimata.Columns(5).Width = 2500
    Me.dt_tmimata.Columns(6).Visible = False
    Me.dt_tmimata.Columns(7).Caption = "Β Προπονητής"
    Me.dt_tmimata.Columns(7).Width = 2500
    For i = 8 To Me.ado_tmimata.Recordset.Fields.Count - 1
        Me.dt_tmimata.Columns(i).Visible = False
    Next i

End Sub

Private Sub cancel_cur_rec_Click()

    
    a_new_tm_is_ready_to_add = 0

    Dim id_s As Integer
    
    If Me.ado_tmimata.Recordset.RecordCount >= 1 Then
        id_s = Me.ado_tmimata.Recordset.Fields(10).Value
        Me.ado_tmimata.Recordset.MoveFirst
        Me.ado_tmimata.Recordset.Find "[TID] like '" & str(id_s) & "'", , adSearchForward
        If Not Me.ado_tmimata.Recordset.EOF Then 'ΤΟ ΕΧΕΙΣ ΒΡΕΙ
            '
            If Trim(Me.ado_tmimata.Recordset.Fields(4).Value) <> "" Then
                If Me.ado_prop.Recordset.RecordCount >= 1 Then
                    Me.ado_prop.Recordset.MoveFirst
                    Me.ado_prop.Recordset.Find "[id] = '" & Me.ado_tmimata.Recordset.Fields(4).Value & "'"
                    If Not Me.ado_prop.Recordset.EOF Then
                        co_propa.Text = Me.ado_prop.Recordset.Fields(1).Value
                    Else
                        co_propa.Text = ""
                    End If
                Else
                    co_propa.Text = ""
                End If
            Else
                co_propa.Text = ""
            End If
            If Trim(Me.ado_tmimata.Recordset.Fields(6).Value) <> "" Then
                If Me.ado_prop.Recordset.RecordCount >= 1 Then
                    Me.ado_prop.Recordset.MoveFirst
                    Me.ado_prop.Recordset.Find "[id] = '" & Me.ado_tmimata.Recordset.Fields(6).Value & "'"
                    If Not Me.ado_prop.Recordset.EOF Then
                        co_propb.Text = Me.ado_prop.Recordset.Fields(1).Value
                    Else
                        co_propb.Text = ""
                    End If
                Else
                    co_propb.Text = ""
                End If
            Else
                co_propb.Text = ""
            End If
            '
            Me.canc_bt.Enabled = True
        Else 'ΔΕΝ ΤΟ ΒΡΕΙΣ
            Me.ado_tmimata.Recordset.Requery
            Me.ado_tmimata.Refresh
            Me.dt_tmimata.Refresh
            Me.ado_tmimata.Recordset.MoveFirst
            Me.dt_tmimata.Columns(0).Visible = False
            Me.dt_tmimata.Columns(1).Caption = "Αθλητικό Έτος"
            Me.dt_tmimata.Columns(1).Width = 1500
            Me.dt_tmimata.Columns(2).Visible = False
            Me.dt_tmimata.Columns(3).Caption = "Κατηγορία Τμήματος"
            Me.dt_tmimata.Columns(3).Width = 2000
            Me.dt_tmimata.Columns(4).Visible = False
            Me.dt_tmimata.Columns(5).Caption = "Α Προπονητής"
            Me.dt_tmimata.Columns(5).Width = 2500
            Me.dt_tmimata.Columns(6).Visible = False
            Me.dt_tmimata.Columns(7).Caption = "Β Προπονητής"
            Me.dt_tmimata.Columns(7).Width = 2500
            For i = 8 To Me.ado_tmimata.Recordset.Fields.Count - 1
                Me.dt_tmimata.Columns(i).Visible = False
            Next i
        End If
    Else
        '
    End If
    
    Me.insert_bt.Enabled = True
    Me.save_command.Enabled = False
    Me.up_bt.Enabled = True

End Sub

Private Sub ch_ib_Click()
    
    'bt_up_athl.Enabled = True
    
    Me.bt_ins_athl.Enabled = False
    Me.bt_del_athl.Enabled = False
    Me.bt_up_athl.Enabled = True
    Me.bt_can_athl.Enabled = True

End Sub

Private Sub co_athl_Change()
    
    'ΝΑ ΕΛΕΓΞΩ ΑΝ Ο ΚΗΔΕΜΟΝΑΣ ΤΟΥ ΑΘΛΗΤΗ ΕΙΝΑΙ ΜΕΛΟΣ
    Dim nm As String
    Dim mp, mm As Boolean
    Dim ad As Integer
    nm = Trim(Me.co_athl.Text)
    'If ado_athl.Recordset.RecordCount >= 1 Then
    If nm <> "" Then
        ado_athl.Recordset.MoveFirst
        ado_athl.Recordset.Find "[ΑΜOE] Like '" & Trim(nm) & "'"
        If Not ado_athl.Recordset.EOF Then
            mp = ado_athl.Recordset.Fields(4).Value
            mm = ado_athl.Recordset.Fields(5).Value
            ad = ado_athl.Recordset.Fields(6).Value
            If mp = True Or mm = True Then
                Me.Label9.Caption = "M"
                Me.txt_pe.Text = "0,00 €"
            Else
                Me.Label9.Caption = " "
                Me.txt_pe.Text = Me.tmp_poso_eggrafis.Text
            End If
            If ad <> 0 Then
                Me.Label11.Caption = "A"
                Me.txt_pm.Text = str(Val(Me.tmp_poso_mina.Text) / 2) + ",00 €"
            Else
                Me.Label11.Caption = ""
                Me.txt_pm.Text = str(Val(Me.tmp_poso_mina.Text)) + ",00 €"
            End If
        End If
    Else
        Me.Label11.Caption = ""
        Me.Label9.Caption = ""
    End If
    '
    Me.bt_ins_athl.Enabled = True
    Me.bt_del_athl.Enabled = True
    If Me.co_athl.Enabled = True Then
        Me.bt_up_athl.Enabled = True
        Me.bt_ins_athl.Enabled = False
        Me.bt_del_athl.Enabled = False
        Me.bt_can_athl.Enabled = True
    End If
    
End Sub

Private Sub co_athl_GotFocus()
    
    Me.bt_ins_athl.Enabled = False
    Me.bt_del_athl.Enabled = False

End Sub

Private Sub Command1_Click()

        for_search = 1
        
        Me.co_ae.Text = ""
        Me.co_kt.Text = ""
        Me.co_propa.Text = ""
        Me.co_propb.Text = ""
        Me.tmp_poso_eggrafis.Text = ""
        Me.tmp_poso_mina.Text = ""
        Me.tmp_poso_eggrafis_plain.Text = ""
        Me.tmp_poso_mina_plain.Text = ""
        
        Me.up_bt.Enabled = False
        Me.save_command.Enabled = False
        Me.sear_bt.Enabled = True
        Me.insert_bt.Enabled = False
        Me.del_bt.Enabled = False
        Me.canc_bt.Enabled = True
        
End Sub

Private Sub Command2_Click()


    Me.dt_tmimata.Refresh
    If Not Me.ado_tmimata.Recordset.EOF Then
        Rep_Καρτέλα_1_Τμήματος.Show
    Else
        MsgBox "Δεν υπάρχει Τμήμα προς Εκτύπωση", , "Μήνυμα Προειδοποίησης"
    End If
            
End Sub

Private Sub Command3_Click()

    Rep_ΈντυποΑπουσιώνΑνάΜήνα_1_Τμήματος.Show

End Sub

Private Sub Command4_Click()

   Dim oExcel As Object
   Dim oBook As Object
   Dim oSheet As Object
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add
   Set oSheet = oBook.Worksheets(1)
   
   
   tmp = Me.ado_athl_tm.Recordset.AbsolutePosition
   
   Clipboard.Clear
   Dim sData As Variant
   sData = ""
   If Me.ado_athl_tm.Recordset.RecordCount >= 1 Then
        Me.ado_athl_tm.Recordset.MoveFirst
        sData = "ΑΜ Αθλητή" & vbTab & "Επώνυμο" & vbTab & "Όνομα" & vbTab & "Γέννηση" & vbTab & "Οδός" & vbTab & "Αριθμός" & vbTab & "Περιοχή" & vbTab & "Δήμος" & vbTab & "Περιφερειακή Ενότητα" & vbTab & "Ταχυδρομικός Κώδικας" & vbTab & "Τηλέφωνο Οικίας" & vbTab & "Κινητό Τηλέφωνο" & vbTab & "Αριθμός Fax" & vbTab & "Email" & vbTab & "Παρατηρήσεις" & vbCr
        For i = 0 To Me.ado_athl_tm.Recordset.RecordCount - 1
            sData = sData & Me.ado_athl_tm.Recordset.Fields(2) _
            & vbTab & Me.ado_athl_tm.Recordset.Fields(10) _
            & vbTab & Me.ado_athl_tm.Recordset.Fields(11) _
            & vbTab & Me.ado_athl_tm.Recordset.Fields(9) _
            & vbTab & Me.ado_athl_tm.Recordset.Fields(12) _
            & vbTab & Me.ado_athl_tm.Recordset.Fields(13) _
            & vbTab & Me.ado_athl_tm.Recordset.Fields(14) _
            & vbTab & Me.ado_athl_tm.Recordset.Fields(15) _
            & vbTab & Me.ado_athl_tm.Recordset.Fields(16) _
            & vbTab & Me.ado_athl_tm.Recordset.Fields(17) _
            & vbTab & Me.ado_athl_tm.Recordset.Fields(18) _
            & vbTab & Me.ado_athl_tm.Recordset.Fields(19) _
            & vbTab & Me.ado_athl_tm.Recordset.Fields(20) _
            & vbTab & Me.ado_athl_tm.Recordset.Fields(21) _
            & vbTab & Me.ado_athl_tm.Recordset.Fields(22) _
            & vbCr
            Me.ado_athl_tm.Recordset.MoveNext
        Next i
        Me.ado_athl_tm.Recordset.MoveFirst
        Me.ado_athl_tm.Recordset.Move tmp - 1
   End If
   Clipboard.SetText sData
   oBook.Worksheets(1).Range("A1").Select
   oBook.Worksheets(1).Paste
   oSheet.Range("A1:Z1").Font.Bold = True
   oSheet.Range("A1:Z1").EntireColumn.AutoFit
   oBook.Worksheets(1).Columns.HorizontalAlignment = -4131
   oExcel.Visible = True
   
End Sub

Private Sub del_bt_Click()
    
    Dim ms As String

    If Not Me.ado_tmimata.Recordset.EOF Then
        ms = MsgBox("Είσαι σίγουρος; (ΝΑΙ ή ΟΧΙ)", vbYesNo, "Παράθυρο διαγραφής")
        tm_id = Me.ado_tmimata.Recordset.Fields(10).Value
        If ms = 6 And Me.ado_athl_tm.Recordset.RecordCount = 0 Then
            col_affected = Me.ado_tmimata.Recordset.AbsolutePosition
            If Me.ado_tmimata.Recordset.AbsolutePosition = Me.ado_tmimata.Recordset.RecordCount Then
                col_affected = col_affected - 1
            End If
            'For i = 0 To Me.ado_tm_ids.Recordset.RecordCount - 1
            '    If Me.ado_tm_ids.Recordset.Fields(0).Value <> tm_id Then
            '        Me.ado_tm_ids.Recordset.MoveNext
            '    Else
            '        i = Me.ado_tm_ids.Recordset.RecordCount
            '    End If
            'Next i
            Me.ado_tm_ids.Recordset.Find "[id] LIKE " & tm_id, , adSearchForward, 1
            Me.ado_tm_ids.Recordset.Delete adAffectCurrent
            Me.ado_tm_ids.Recordset.Requery
            Me.ado_tm_ids.Refresh
            Me.DataGrid2.Refresh
            Me.ado_tmimata.Recordset.Requery
            Me.ado_tmimata.Refresh
            Me.dt_tmimata.Refresh
            Me.ado_tmimata.Recordset.Sort = s_sort
            Me.ado_tmimata.Recordset.Move col_affected - 1, 0
            Me.dt_tmimata.SelBookmarks.Add Me.ado_tmimata.Recordset.Bookmark
            Me.dt_tmimata.Columns(0).Visible = False
            Me.dt_tmimata.Columns(1).Caption = "Αθλητικό Έτος"
            Me.dt_tmimata.Columns(1).Width = 1500
            Me.dt_tmimata.Columns(2).Visible = False
            Me.dt_tmimata.Columns(3).Caption = "Κατηγορία Τμήματος"
            Me.dt_tmimata.Columns(3).Width = 2000
            Me.dt_tmimata.Columns(4).Visible = False
            Me.dt_tmimata.Columns(5).Caption = "Α Προπονητής"
            Me.dt_tmimata.Columns(5).Width = 2500
            Me.dt_tmimata.Columns(6).Visible = False
            Me.dt_tmimata.Columns(7).Caption = "Β Προπονητής"
            Me.dt_tmimata.Columns(7).Width = 2500
            For i = 8 To Me.ado_tmimata.Recordset.Fields.Count - 1
                Me.dt_tmimata.Columns(i).Visible = False
            Next i
        Else
            If ms = 6 And Me.ado_athl_tm.Recordset.RecordCount <> 0 Then
                MsgBox "ΑΠΑΓΟΡΕΥΕΤΑΙ η ΔΙΑΓΡΑΦΗ ΤΜΗΜΑΤΟΣ που περιέχει ΑΘΛΗΤΕΣ!", vbCritical, "Μήνυμα Λάθους"
            Else
                MsgBox "Ακύρωση ΔΙΑΓΡΑΦΗΣ!", vbOKOnly, "Παράθυρο Ακύρωσης"
            End If
        End If
    Else
        MsgBox "Δεν υπάρχει εγγραφή προς ΔΙΑΓΡΑΦΗ!", vbCritical, "Μήνυμα Λάθους"
    End If
    
    
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

Private Sub dt_athl_tm_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim c, r As Integer
    Dim t As String
    If KeyCode = vbKeyDown And Me.ado_athl_tm.Recordset.AbsolutePosition < Me.ado_athl_tm.Recordset.RecordCount Then
        c = Me.dt_athl_tm.Col
        Me.ado_athl_tm.Recordset.MoveNext
        Me.dt_athl_tm.Col = c
        Me.co_athl.SetFocus
        Me.dt_athl_tm.SetFocus
    End If
    If KeyCode = vbKeyUp And Me.ado_athl_tm.Recordset.AbsolutePosition > 1 Then
        c = Me.dt_athl_tm.Col
        Me.ado_athl_tm.Recordset.MovePrevious
        Me.dt_athl_tm.Col = c
        Me.co_athl.SetFocus
        Me.dt_athl_tm.SetFocus
    End If
    c = Me.dt_athl_tm.Col
    If KeyCode = vbKeyRight And (c = 2 Or c = 3 Or c = 4 Or c = 5 Or c = 6 Or c = 8) Then
        If c = 6 Then
            t = Me.dt_athl_tm.Columns(c + 2).Text
            Me.dt_athl_tm.Col = c + 2
        Else
            If c = 3 Or c = 4 Or c = 5 Then
                t = Me.dt_athl_tm.Columns(c + 1).Text
                Me.dt_athl_tm.Col = c + 1
                Me.dt_athl_tm.SelText = t
            Else
                If c = 2 Then
                    t = Me.dt_athl_tm.Columns(c + 1).Text
                    Me.dt_athl_tm.Col = c + 1
                End If
            End If
        End If
        Me.dt_athl_tm.SetFocus
    End If
    If KeyCode = vbKeyLeft And (c = 2 Or c = 3 Or c = 4 Or c = 5 Or c = 6 Or c = 8) Then
        Select Case c
            Case 2
    
            Case 3, 5, 6
                t = Me.dt_athl_tm.Columns(c - 1).Text
                Me.dt_athl_tm.Col = c - 1
                Me.dt_athl_tm.SelText = t
            Case 4
                t = Me.dt_athl_tm.Columns(c - 1).Text
                Me.dt_athl_tm.Col = c - 1
            Case Else
                t = Me.dt_athl_tm.Columns(c - 2).Text
                Me.dt_athl_tm.Col = c - 2
                Me.dt_athl_tm.SelText = t
        End Select
        Me.dt_athl_tm.SetFocus
    End If
    
End Sub

Private Sub dt_athl_tm_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Me.ado_athl_tm.Recordset.AbsolutePosition >= 1 And Me.ado_athl_tm.Recordset.AbsolutePosition <= Me.ado_athl_tm.Recordset.RecordCount Then
        If Not Me.ado_athl_tm.Recordset.EOF Then
          If Trim(Me.ado_athl_tm.Recordset.Fields(1).Value) <> "" Then
                If Not rs_ado_athl.EOF Then
                    rs_ado_athl.MoveFirst
                    rs_ado_athl.Find "[id] = '" & Me.ado_athl_tm.Recordset.Fields(1).Value & "'"
                    If Not rs_ado_athl.EOF Then
                        co_athl.Text = rs_ado_athl.Fields(1).Value
                    End If
                End If
            Else
                co_athl.Text = ""
            End If
            If Trim(Me.ado_athl_tm.Recordset.Fields(4).Value) <> "" Then
                Me.txt_im_eis.Text = Me.ado_athl_tm.Recordset.Fields(4).Value
            Else
                Me.txt_im_eis.Text = "00/00/0000"
            End If
            If Trim(Me.ado_athl_tm.Recordset.Fields(5).Value) <> "" Then
                Me.txt_pe.Text = Me.ado_athl_tm.Recordset.Fields(5).Value
            Else
                Me.txt_pe.Text = ""
            End If
            If Trim(Me.ado_athl_tm.Recordset.Fields(6).Value) <> "" Then
                Me.txt_pm.Text = Me.ado_athl_tm.Recordset.Fields(6).Value
            Else
                Me.txt_pm.Text = ""
            End If
            If Trim(Me.ado_athl_tm.Recordset.Fields(8).Value) <> "" Then
                If Me.ado_athl_tm.Recordset.Fields(8).Value Like "NAI" Then
                    Me.ch_ib.Value = 1
                Else
                    Me.ch_ib.Value = 0
                End If
            Else
                Me.ch_ib.Value = 0
            End If
        End If
        bt_ins_athl.Enabled = True
        bt_up_athl.Enabled = False
        bt_del_athl.Enabled = True
        Me.bt_can_athl.Enabled = False
    End If
    
End Sub

Private Sub dt_del_athl_Click()

    
End Sub

Private Sub dt_tmimata_HeadClick(ByVal ColIndex As Integer)

    defined_col = ColIndex

End Sub

Private Sub dt_tmimata_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Me.ado_tmimata.Recordset.AbsolutePosition >= 1 And Me.ado_tmimata.Recordset.AbsolutePosition <= Me.ado_tmimata.Recordset.RecordCount Then
            If Trim(Me.ado_tmimata.Recordset.Fields(0).Value) <> "" Then
                If Not Me.ado_ae.Recordset.EOF Then
                    Me.ado_ae.Recordset.MoveFirst
                    Me.ado_ae.Recordset.Find "[id_αθλητικού_έτους] = '" & Me.ado_tmimata.Recordset.Fields(0).Value & "'"
                    If Not Me.ado_ae.Recordset.EOF Then
                        co_ae.Text = Me.ado_ae.Recordset.Fields(1).Value
                    End If
                End If
            Else
                co_ae.Text = ""
            End If
            If Trim(Me.ado_tmimata.Recordset.Fields(2).Value) <> "" Then
                If Not rs_ado_kt.EOF Then
                    rs_ado_kt.MoveFirst
                   rs_ado_kt.Find "[id_κατηγορίας_τμήματος] = '" & Me.ado_tmimata.Recordset.Fields(2).Value & "'"
                    If Not rs_ado_kt.EOF Then
                        co_kt.Text = rs_ado_kt.Fields(1).Value
                    End If
                End If
            Else
                co_kt.Text = ""
            End If
            If Trim(Me.ado_tmimata.Recordset.Fields(4).Value) <> "" Then
                'If Not rs_ado_prop.EOF Then
                If Me.ado_prop.Recordset.RecordCount >= 1 Then
                    Me.ado_prop.Recordset.MoveFirst
                    Me.ado_prop.Recordset.Find "[id] = '" & Me.ado_tmimata.Recordset.Fields(4).Value & "'"
                    If Not Me.ado_prop.Recordset.EOF Then
                        co_propa.Text = Me.ado_prop.Recordset.Fields(1).Value
                    Else
                        co_propa.Text = ""
                    End If
                Else
                    co_propa.Text = ""
                End If
            Else
                co_propa.Text = ""
            End If
            If Trim(Me.ado_tmimata.Recordset.Fields(6).Value) <> "" Then
                'If Not rs_ado_prop.EOF Then
                If Me.ado_prop.Recordset.RecordCount >= 1 Then
                    Me.ado_prop.Recordset.MoveFirst
                    Me.ado_prop.Recordset.Find "[id] = '" & Me.ado_tmimata.Recordset.Fields(6).Value & "'"
                    If Not Me.ado_prop.Recordset.EOF Then
                        co_propb.Text = Me.ado_prop.Recordset.Fields(1).Value
                    Else
                        co_propb.Text = ""
                    End If
                Else
                    co_propb.Text = ""
                End If
            Else
                co_propb.Text = ""
            End If
            If Trim(Me.ado_tmimata.Recordset.Fields(8).Value) <> "" Then
                Me.tmp_poso_eggrafis_plain.Text = Trim(str(Me.ado_tmimata.Recordset.Fields(8).Value))
                Me.tmp_poso_eggrafis.Text = Trim(str(Me.ado_tmimata.Recordset.Fields(8).Value)) + ",00 €"
            Else
                tmp_poso_eggrafis.Text = ""
            End If
            If Trim(Me.ado_tmimata.Recordset.Fields(9).Value) <> "" Then
                Me.tmp_poso_mina_plain.Text = Trim(str(Me.ado_tmimata.Recordset.Fields(9).Value))
                Me.tmp_poso_mina.Text = Trim(str(Me.ado_tmimata.Recordset.Fields(9).Value)) + ",00 €"
            Else
                Me.tmp_poso_mina.Text = ""
            End If
            'ΕΝΗΜΕΡΩΣΗ ΤΟΥ ΕΒΔΟΜΑΔΙΑΙΟΥ ΠΡΟΓΡΑΜΜΑΤΟΣ
            i = 0
            For j = 0 To 5
                If Trim(Me.ado_tmimata.Recordset.Fields(11 + i).Value) <> "" Then
                    Me.im_en(j).Text = Me.ado_tmimata.Recordset.Fields(11 + i).Value
                Else
                    Me.im_en(j).Text = "00:00"
                End If
                If Trim(Me.ado_tmimata.Recordset.Fields(12 + i).Value) <> "" Then
                    Me.im_l(j).Text = Me.ado_tmimata.Recordset.Fields(12 + i).Value
                Else
                    Me.im_l(j).Text = "00:00"
                End If
                i = i + 2
            Next j
            'REFRESH του datagrid ΑθλητέςΤμήματα
            Me.Label9.Caption = ""
            Me.Label11.Caption = ""
            Me.ado_athl_tm.Recordset.Filter = "[id_Τμήματος] = '" & Me.ado_tmimata.Recordset.Fields(10).Value & "'"
            If Me.ado_athl_tm.Recordset.RecordCount > 0 Then
                Me.bt_can_athl.Enabled = False
                Me.co_athl.Enabled = True
                Me.txt_im_eis.Enabled = True
                Me.txt_pe.Enabled = True
                Me.txt_pm.Enabled = True
                Me.ch_ib.Enabled = True
                Me.dt_athl_tm.Row = 0
                Me.dt_athl_tm.Col = 1
                Me.ado_athl_tm.Caption = "Αθλητής " & Me.dt_athl_tm.Row + 1 & " από " & Me.ado_athl_tm.Recordset.RecordCount
                ''
                If Trim(Me.ado_athl_tm.Recordset.Fields(1).Value) <> "" Then
                    If Not rs_ado_athl.EOF Then
                        rs_ado_athl.MoveFirst
                        rs_ado_athl.Find "[id] = '" & Me.ado_athl_tm.Recordset.Fields(1).Value & "'"
                        If Not rs_ado_athl.EOF Then
                            co_athl.Text = rs_ado_athl.Fields(1).Value
                        End If
                    End If
                Else
                    co_athl.Text = ""
                End If
                If Trim(Me.ado_athl_tm.Recordset.Fields(4).Value) <> "" Then
                    Me.txt_im_eis.Text = Trim(Me.ado_athl_tm.Recordset.Fields(4).Value)
                Else
                    Me.txt_im_eis.Text = "00/00/0000"
                End If
                If Trim(Me.ado_athl_tm.Recordset.Fields(5).Value) <> "" Then
                    Me.txt_pe.Text = Trim(Me.ado_athl_tm.Recordset.Fields(5).Value)
                Else
                    Me.txt_pe.Text = ""
                End If
                If Trim(Me.ado_athl_tm.Recordset.Fields(6).Value) <> "" Then
                    Me.txt_pm.Text = Trim(Me.ado_athl_tm.Recordset.Fields(6).Value)
                Else
                    Me.txt_pm.Text = ""
                End If
                If Trim(Me.ado_athl_tm.Recordset.Fields(8).Value) <> "" Then
                    If Trim(Me.ado_athl_tm.Recordset.Fields(8).Value) Like "ΝΑΙ" Then
                        Me.ch_ib.Value = 1
                    Else
                        Me.ch_ib.Value = 0
                    End If
                Else
                    Me.ch_ib.Value = 0
                End If
                ''
            Else
                If Me.ado_athl_tm.Recordset.RecordCount = 0 Then
                    Me.ado_athl_tm.Caption = "Αθλητής " & 0 & " από " & 0
                    co_athl.Text = ""
                    Me.co_athl.Enabled = False
                    Me.txt_im_eis.Text = "00/00/0000"
                    Me.txt_im_eis.Enabled = False
                    Me.txt_pe.Text = ""
                    Me.txt_pe.Enabled = False
                    Me.txt_pm.Text = ""
                    Me.txt_pm.Enabled = False
                    Me.ch_ib.Value = 0
                    Me.ch_ib.Enabled = False
                    Me.bt_ins_athl.Enabled = True
                    Me.bt_up_athl.Enabled = False
                    Me.bt_can_athl.Enabled = False
                    Me.bt_del_athl.Enabled = False
                End If
            End If
            'ΤΕΛΟΣ - refresh του datagrid αθλητές-τμήματα
            Me.dt_athl_tm.Columns(0).Visible = False
            Me.dt_athl_tm.Columns(1).Visible = False
            Me.dt_athl_tm.Columns(2).Caption = "Α.Μ. Αθλητή"
            Me.dt_athl_tm.Columns(2).Width = 1000
            Me.dt_athl_tm.Columns(3).Caption = "Ονοματεπώνυμο Αθλητή"
            Me.dt_athl_tm.Columns(3).Width = 2000
            Me.dt_athl_tm.Columns(4).Caption = "Ημ/νία Εισαγωγής"
            Me.dt_athl_tm.Columns(4).Width = 1450
            Me.dt_athl_tm.Columns(5).Caption = "Ποσό Εγγραφής"
            Me.dt_athl_tm.Columns(5).Width = 1250
            Me.dt_athl_tm.Columns(6).Caption = "Ποσό Μήνα"
            Me.dt_athl_tm.Columns(6).Width = 1250
            Me.dt_athl_tm.Columns(7).Visible = False
            Me.dt_athl_tm.Columns(8).Caption = "Ιατρική Βεβαίωση"
            Me.dt_athl_tm.Columns(8).Width = 1450
            For i = 9 To 23
                Me.dt_athl_tm.Columns(i).Visible = False
            Next i
    End If

End Sub

Private Sub Form_Load()
    
    a_new_tm_is_ready_to_add = 0
    s_sort = ""
    for_search = 0
    
    parousies.UseMaskColor = True
    parousies.MaskColor = vbRed
    
    Me.ado_ae.Refresh
    Set Me.co_ae.RowSource = Me.ado_ae
    Me.ado_athl.Refresh
    Set rs_ado_athl = ado_athl.Recordset
    Me.ado_athl_tm.Refresh
    Set Me.dt_athl_tm.DataSource = Me.ado_athl_tm
    Me.ado_athl_tm_ids.Refresh
    Set rs_ado_athl_tm_ids = ado_athl_tm_ids.Recordset
    Set Me.dt_athl_tm_ids.DataSource = Me.ado_athl_tm_ids
    Me.ado_kt.Refresh
    Set rs_ado_kt = ado_kt.Recordset
    Set Me.co_kt.RowSource = Me.ado_kt
    Me.ado_prop.Refresh
    Set rs_ado_prop = ado_prop.Recordset
    Set Me.co_propa.RowSource = Me.ado_prop
    Set Me.co_propb.RowSource = Me.ado_prop
    Me.ado_tmimata.Refresh
    Set Me.dt_tmimata.DataSource = Me.ado_tmimata
    Set Me.co_athl.RowSource = Me.ado_athl
  
  
    Me.Height = 11500
    Me.Width = 15000
    
    new_addition = 0
    Me.bt_up_athl.Enabled = False
    
    s_sort = "[" & Trim(Me.ado_tmimata.Recordset.Fields(1).Name) & "] DESC, [" & Trim(Me.ado_tmimata.Recordset.Fields(3).Name) & "] ASC"
    Me.ado_tmimata.Recordset.Sort = s_sort
    '******************************************************
    If Not Me.ado_tmimata.Recordset.EOF Then
        Me.ado_tmimata.Recordset.MoveFirst
        If Trim(Me.ado_tmimata.Recordset.Fields(0).Value) <> "" Then
            If Not Me.ado_ae.Recordset.EOF Then
                Me.ado_ae.Recordset.MoveFirst
                Me.ado_ae.Recordset.Find "[id_αθλητικού_έτους] = '" & Me.ado_tmimata.Recordset.Fields(0).Value & "'"
                If Not Me.ado_ae.Recordset.EOF Then
                    co_ae.Text = Me.ado_ae.Recordset.Fields(1).Value
                End If
            End If
        Else
            co_ae.Text = ""
        End If
        If Trim(Me.ado_tmimata.Recordset.Fields(2).Value) <> "" Then
            If Not rs_ado_kt.EOF Then
                rs_ado_kt.MoveFirst
                rs_ado_kt.Find "[id_κατηγορίας_τμήματος] = '" & Me.ado_tmimata.Recordset.Fields(2).Value & "'"
                If Not rs_ado_kt.EOF Then
                    co_kt.Text = rs_ado_kt.Fields(1).Value
                End If
            End If
        Else
            co_kt.Text = ""
        End If
        If Trim(Me.ado_tmimata.Recordset.Fields(4).Value) <> "" Then
            If Not rs_ado_prop.EOF Then
                rs_ado_prop.MoveFirst
                rs_ado_prop.Find "[id] = '" & Me.ado_tmimata.Recordset.Fields(4).Value & "'"
                If Not rs_ado_prop.EOF Then
                    co_propa.Text = rs_ado_prop.Fields(1).Value
                End If
            End If
        Else
            tmp_propA.Text = ""
        End If
        If Trim(Me.ado_tmimata.Recordset.Fields(6).Value) <> "" Then
            If Not rs_ado_prop.EOF Then
                rs_ado_prop.MoveFirst
                rs_ado_prop.Find "[id] = '" & Me.ado_tmimata.Recordset.Fields(6).Value & "'"
                If Not rs_ado_prop.EOF Then
                    co_propb.Text = rs_ado_prop.Fields(1).Value
                End If
            End If
        Else
            co_propb.Text = ""
        End If
        If Trim(Me.ado_tmimata.Recordset.Fields(8).Value) <> "" Then
            tmp_poso_eggrafis.Text = Val(Me.ado_tmimata.Recordset.Fields(8).Value) & ",00 €"
        Else
            tmp_poso_eggrafis.Text = ""
        End If
        If Trim(Me.ado_tmimata.Recordset.Fields(9).Value) <> "" Then
            Me.tmp_poso_mina.Text = Val(Me.ado_tmimata.Recordset.Fields(9).Value) & ",00 €"
        Else
            Me.tmp_poso_mina.Text = ""
        End If
        'ΕΝΗΜΕΡΩΣΗ ΤΟΥ ΕΒΔΟΜΑΔΙΑΙΟΥ ΠΡΟΓΡΑΜΜΑΤΟΣ
        i = 0
        For j = 0 To 5
            If Trim(Me.ado_tmimata.Recordset.Fields(11 + i).Value) <> "" Then
                Me.im_en(j).Text = Me.ado_tmimata.Recordset.Fields(11 + i).Value
            Else
                Me.im_en(j).Text = "00:00"
            End If
            If Trim(Me.ado_tmimata.Recordset.Fields(12 + i).Value) <> "" Then
                Me.im_l(j).Text = Me.ado_tmimata.Recordset.Fields(12 + i).Value
            Else
                Me.im_l(j).Text = "00:00"
            End If
            i = i + 2
        Next j
    End If
    '***************************************************************
       
    If Me.ado_tmimata.Recordset.RecordCount > 0 Then
        Me.dt_tmimata.Row = 0
        Me.dt_tmimata.Col = 1
    End If
    If Me.ado_tmimata.Recordset.RecordCount > 0 Then
        Me.ado_tmimata.Caption = "Τμήμα " & Me.dt_tmimata.Row + 1 & " από " & Me.ado_tmimata.Recordset.RecordCount
    End If
    
    Me.dt_tmimata.Columns(0).Visible = False
    Me.dt_tmimata.Columns(1).Caption = "Αθλητικό Έτος"
    Me.dt_tmimata.Columns(1).Width = 1500
    Me.dt_tmimata.Columns(2).Visible = False
    Me.dt_tmimata.Columns(3).Caption = "Κατηγορία Τμήματος"
    Me.dt_tmimata.Columns(3).Width = 2000
    Me.dt_tmimata.Columns(4).Visible = False
    Me.dt_tmimata.Columns(5).Caption = "Α Προπονητής"
    Me.dt_tmimata.Columns(5).Width = 2500
    Me.dt_tmimata.Columns(6).Visible = False
    Me.dt_tmimata.Columns(7).Caption = "Β Προπονητής"
    Me.dt_tmimata.Columns(7).Width = 2500
    For i = 8 To Me.ado_tmimata.Recordset.Fields.Count - 1
        Me.dt_tmimata.Columns(i).Visible = False
    Next i
    
    'REFRESH του datagrid ΑθλητέςΤμήματα
    Me.ado_athl_tm.Recordset.Filter = "[id_Τμήματος] = '" & Me.ado_tmimata.Recordset.Fields(10).Value & "'"
    If Me.ado_athl_tm.Recordset.RecordCount > 0 Then
        Me.dt_athl_tm.Row = 0
        Me.dt_athl_tm.Col = 1
        Me.ado_athl_tm.Caption = "Αθλητής " & Me.dt_athl_tm.Row + 1 & " από " & Me.ado_athl_tm.Recordset.RecordCount
        Me.bt_can_athl.Enabled = False
    Else
        If Me.ado_athl_tm.Recordset.RecordCount = 0 Then
            Me.ado_athl_tm.Caption = "Αθλητής " & 0 & " από " & 0
            Me.co_athl.Text = ""
            Me.co_athl.Enabled = False
            Me.txt_im_eis.Text = "00/00/0000"
            Me.txt_im_eis.Enabled = False
            Me.txt_pe.Text = ""
            Me.txt_pe.Enabled = False
            Me.txt_pm.Text = ""
            Me.txt_pm.Enabled = False
            Me.ch_ib.Value = 0
            Me.ch_ib.Enabled = False
            Me.bt_up_athl.Enabled = False
            Me.bt_can_athl.Enabled = False
            Me.bt_del_athl.Enabled = False
        End If
    End If
        
    Me.dt_athl_tm.Columns(0).Visible = False
    Me.dt_athl_tm.Columns(1).Visible = False
    Me.dt_athl_tm.Columns(2).Caption = "Α.Μ. Αθλητή"
    Me.dt_athl_tm.Columns(2).Width = 1000
    Me.dt_athl_tm.Columns(3).Caption = "Ονοματεπώνυμο Αθλητή"
    Me.dt_athl_tm.Columns(3).Width = 2000
    Me.dt_athl_tm.Columns(4).Caption = "Ημ/νία Εισαγωγής"
    Me.dt_athl_tm.Columns(4).Width = 1450
    Me.dt_athl_tm.Columns(5).Caption = "Ποσό Εγγραφής"
    Me.dt_athl_tm.Columns(5).Width = 1250
    Me.dt_athl_tm.Columns(6).Caption = "Ποσό Μήνα"
    Me.dt_athl_tm.Columns(6).Width = 1250
    Me.dt_athl_tm.Columns(7).Visible = False
    Me.dt_athl_tm.Columns(8).Caption = "Ιατρική Βεβαίωση"
    Me.dt_athl_tm.Columns(8).Width = 1450
    For i = 9 To 23
        Me.dt_athl_tm.Columns(i).Visible = False
    Next i
    
    bt_ins_athl.Enabled = True
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    
    Set rs_ado_ae = Nothing
    Set rs_ado_athl = Nothing
    Set rs_ado_athl_tm = Nothing
    Set Me.dt_athl_tm.DataSource = Nothing
    Set rs_ado_athl_tm_ids = Nothing
    Set Me.dt_athl_tm_ids.DataSource = Nothing
    Set rs_ado_kt = Nothing
    Set rs_ado_prop = Nothing
    'Set rs_ado_tm_ids = Nothing
    'Set Me.DataGrid2.DataSource = Nothing
    Set rs_ado_tmimata = Nothing
    Set Me.dt_tmimata.DataSource = Nothing
    Set Me.co_ae.RowSource = Nothing
    Set Me.co_kt.RowSource = Nothing
    Set Me.co_propa.RowSource = Nothing
    Set Me.co_propb.RowSource = Nothing
    Set Me.co_athl.RowSource = Nothing

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

Private Sub insert_bt_Click()

    a_new_tm_is_ready_to_add = 1
    'Να βρω το υποψήφιο id_τμήματος
    Me.ado_tm_ids.Recordset.Sort = "[" & Trim(Me.ado_tm_ids.Recordset.Fields(0).Name) & "]"
    If Not Me.ado_tm_ids.Recordset.EOF Then
        Me.ado_tm_ids.Recordset.MoveLast
        id_τμ = Me.ado_tm_ids.Recordset.Fields(0).Value
        id_τμ = id_τμ + 1
    End If
    'Καθαρισμός πεδίων
    co_ae.Text = ""
    co_kt.Text = ""
    co_propa.Text = ""
    co_propb.Text = ""
    tmp_poso_eggrafis.Text = ""
    tmp_poso_mina.Text = ""
    'ΚΑΘΑΡΙΣΜΟΣ ΤΟΥ ΕΒΔΟΜΑΔΙΑΙΟΥ ΠΡΟΓΡΑΜΜΑΤΟΣ
    For j = 0 To 5
        Me.im_en(j).Text = "00:00"
        Me.im_l(j).Text = "00:00"
    Next j
    'ΚΑΘΑΡΙΣΜΟΣ ΤΩΝ ΑΘΛΗΤΩΝ ΤΟΥ ΤΜΗΜΑΤΟΣ
    'Me.ado_athl_tm.Recordset.Filter = "[id_Τμήματος] = '" & Me.ado_tmimata.Recordset.Fields(10).Value & "'"
    Me.ado_athl_tm.Recordset.Filter = "[id_Τμήματος] = '" & -1 & "'"
    If Me.ado_athl_tm.Recordset.RecordCount > 0 Then
        Me.dt_athl_tm.Row = 0
        Me.dt_athl_tm.Col = 1
        Me.ado_athl_tm.Caption = "Αθλητής " & Me.dt_athl_tm.Row + 1 & " από " & Me.ado_athl_tm.Recordset.RecordCount
        ''
        If Trim(Me.ado_athl_tm.Recordset.Fields(1).Value) <> "" Then
            If Not rs_ado_athl.EOF Then
                rs_ado_athl.MoveFirst
                rs_ado_athl.Find "[id] = '" & Me.ado_athl_tm.Recordset.Fields(1).Value & "'"
                If Not rs_ado_athl.EOF Then
                    co_athl.Text = rs_ado_athl.Fields(1).Value
                End If
            End If
        Else
            co_athl.Text = ""
        End If
        If Trim(Me.ado_athl_tm.Recordset.Fields(4).Value) <> "" Then
            Me.txt_im_eis.Text = Trim(Me.ado_athl_tm.Recordset.Fields(4).Value)
        Else
            Me.txt_im_eis.Text = "00/00/0000"
        End If
        If Trim(Me.ado_athl_tm.Recordset.Fields(5).Value) <> "" Then
            Me.txt_pe.Text = Trim(Me.ado_athl_tm.Recordset.Fields(5).Value)
        Else
            Me.txt_pe.Text = ""
        End If
        If Trim(Me.ado_athl_tm.Recordset.Fields(6).Value) <> "" Then
            Me.txt_pm.Text = Trim(Me.ado_athl_tm.Recordset.Fields(6).Value)
        Else
            Me.txt_pm.Text = ""
        End If
        If Trim(Me.ado_athl_tm.Recordset.Fields(8).Value) <> "" Then
            If Trim(Me.ado_athl_tm.Recordset.Fields(8).Value) Like "ΝΑΙ" Then
                Me.ch_ib.Value = 1
            Else
                Me.ch_ib.Value = 0
            End If
        Else
            Me.ch_ib.Value = 0
        End If
        ''
    Else
        If Me.ado_athl_tm.Recordset.RecordCount = 0 Then
            Me.ado_athl_tm.Caption = "Αθλητής " & 0 & " από " & 0
            co_athl.Text = ""
            Me.co_athl.Enabled = False
            Me.txt_im_eis.Text = "00/00/0000"
            Me.txt_im_eis.Enabled = False
            Me.txt_pe.Text = ""
            Me.txt_pe.Enabled = False
            Me.txt_pm.Text = ""
            Me.txt_pm.Enabled = False
            Me.ch_ib.Value = 0
            Me.ch_ib.Enabled = False
            Me.bt_ins_athl.Enabled = True
            Me.bt_up_athl.Enabled = False
        End If
    End If
    'ΤΕΛΟΣ - refresh του datagrid αθλητές-τμήματα
    
    Me.co_ae.SetFocus
    
    Me.insert_bt.Enabled = False
    Me.save_command.Enabled = True
    Me.canc_bt.Enabled = True
    Me.up_bt.Enabled = False
    
End Sub

Private Sub kl_bt_Click()

    Unload Me

End Sub

Private Sub parousies_Click()

    parousies_management.Show

End Sub

Private Sub save_command_Click()

    a_new_tm_is_ready_to_add = 0
    'Αποθήκευση στα τμήματα
    Me.ado_tm_ids.Recordset.AddNew
    'Αποθήκευση id
    Me.ado_tm_ids.Recordset.Fields(0).Value = id_τμ
    'Αποθήκευση ΛΟΙΠΩΝ ΣΤΟΙΧΕΙΑ
    If Trim(Me.co_ae.Text) <> "" Then
        If Me.co_ae.SelectedItem >= 1 Then
            Me.ado_ae.Recordset.MoveFirst
            Me.ado_ae.Recordset.Move Me.co_ae.SelectedItem - 1
            Me.ado_tm_ids.Recordset.Fields(1).Value = Me.ado_ae.Recordset.Fields(0).Value
        End If
    Else
        Me.ado_tm_ids.Recordset.Fields(1).Value = 1
    End If
    If Trim(Me.co_kt.Text) <> "" Then
        If Me.co_kt.SelectedItem >= 1 Then
            rs_ado_kt.MoveFirst
            rs_ado_kt.Move Me.co_kt.SelectedItem - 1
            Me.ado_tm_ids.Recordset.Fields(2).Value = rs_ado_kt.Fields(0).Value
        End If
    Else
        Me.ado_tm_ids.Recordset.Fields(2).Value = 3
    End If
    If Trim(Me.co_propa.Text) <> "" Then
        If Me.co_propa.SelectedItem >= 1 Then
            rs_ado_prop.MoveFirst
            rs_ado_prop.Move Me.co_propa.SelectedItem - 1
            Me.ado_tm_ids.Recordset.Fields(3).Value = rs_ado_prop.Fields(0).Value
        End If
    Else
        Me.ado_tm_ids.Recordset.Fields(3).Value = 0
    End If
    If Trim(Me.co_propb.Text) <> "" Then
        If Me.co_propb.SelectedItem >= 1 Then
            rs_ado_prop.MoveFirst
            rs_ado_prop.Move Me.co_propb.SelectedItem - 1
            Me.ado_tm_ids.Recordset.Fields(4).Value = rs_ado_prop.Fields(0).Value
        End If
        Else
            Me.ado_tm_ids.Recordset.Fields(4).Value = 0
    End If
    If Trim(Me.tmp_poso_eggrafis.Text) <> "" Then
        Me.ado_tm_ids.Recordset.Fields(5).Value = Me.tmp_poso_eggrafis_plain.Text
    End If
    If Trim(Me.tmp_poso_mina.Text) <> "" Then
        Me.ado_tm_ids.Recordset.Fields(6).Value = Me.tmp_poso_mina_plain.Text
    End If
    'ΕΝΗΜΕΡΩΣΗ ΤΟΥ ΕΒΔΟΜΑΔΙΑΙΟΥ ΠΡΟΓΡΑΜΜΑΤΟΣ
    i = 0
    For j = 0 To 5
        If Trim(Me.im_en(j).Text) <> "" Then
            Me.ado_tm_ids.Recordset.Fields(7 + i).Value = Me.im_en(j).Text
        Else
            Me.ado_tm_ids.Recordset.Fields(7 + i).Value = ""
        End If
        If Trim(Me.im_l(j).Text) <> "" Then
            Me.ado_tm_ids.Recordset.Fields(8 + i).Value = Me.im_l(j).Text
        Else
            Me.ado_tm_ids.Recordset.Fields(8 + i).Value = ""
        End If
        i = i + 2
    Next j
        
    Me.ado_tm_ids.Recordset.UpdateBatch adAffectCurrent
    Me.ado_tm_ids.Recordset.Requery
    Me.ado_tm_ids.Refresh
    Me.DataGrid2.Refresh
    Me.ado_tmimata.Recordset.Requery
    Me.ado_tmimata.Refresh
    Me.dt_tmimata.Refresh
    Me.ado_tmimata.Recordset.Sort = "[" & Trim(Me.ado_tmimata.Recordset.Fields(1).Name) & "] DESC, [" & Trim(Me.ado_tmimata.Recordset.Fields(3).Name) & "] ASC"
    'Me.ado_tmimata.Recordset.MoveLast
    'Me.dt_tmimata.Col = 1
    '
    'For i = 0 To Me.ado_tmimata.Recordset.RecordCount - 1
    '        If Me.ado_tmimata.Recordset.Fields(10).Value <> tm_id Then
    '            Me.ado_tmimata.Recordset.MoveNext
    '        Else
    '            i = Me.ado_tmimata.Recordset.RecordCount
    '            Me.dt_tmimata.Col = 1
    '        End If
    '    Next i
    '
    Me.ado_tmimata.Recordset.Find "[TID] LIKE " & id_τμ, , adSearchForward, 1
    Me.dt_tmimata.SelBookmarks.Add Me.ado_tmimata.Recordset.Bookmark
    If Me.ado_tmimata.Recordset.RecordCount > 0 Then
        Me.ado_tmimata.Caption = "Τμήμα " & Me.ado_tmimata.Recordset.AbsolutePosition & " από " & Me.ado_tmimata.Recordset.RecordCount
    End If
    Me.dt_tmimata.Columns(0).Visible = False
    Me.dt_tmimata.Columns(1).Caption = "Αθλητικό Έτος"
    Me.dt_tmimata.Columns(1).Width = 1500
    Me.dt_tmimata.Columns(2).Visible = False
    Me.dt_tmimata.Columns(3).Caption = "Κατηγορία Τμήματος"
    Me.dt_tmimata.Columns(3).Width = 2000
    Me.dt_tmimata.Columns(4).Visible = False
    Me.dt_tmimata.Columns(5).Caption = "Α Προπονητής"
    Me.dt_tmimata.Columns(5).Width = 2500
    Me.dt_tmimata.Columns(6).Visible = False
    Me.dt_tmimata.Columns(7).Caption = "Β Προπονητής"
    Me.dt_tmimata.Columns(7).Width = 2500
    For i = 8 To Me.ado_tmimata.Recordset.Fields.Count - 1
        Me.dt_tmimata.Columns(i).Visible = False
    Next i
        
    Me.save_command.Enabled = False
    Me.canc_bt.Enabled = True
    Me.insert_bt.Enabled = True
    Me.up_bt.Enabled = True
    
End Sub

Private Sub sear_bt_Click()

    'ΑΠΑΡΑΙΤΗΤΕΣ ΑΡΧΙΚΟΠΟΙΗΣΕΙΣ
    s_string = ""
    
    'ΚΡΙΤΗΡΙΟ ΑΘΛΗΤΙΚΟΥ ΕΤΟΥΣ
    If Trim(co_ae.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Αθλητικά_Έτη.Περιγραφή] LIKE '*" & Trim(co_ae.Text) & "*'"
        Else
            s_string = "[Αθλητικά_Έτη.Περιγραφή] LIKE '*" & Trim(co_ae.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΚΑΤΗΓΟΡΙΑ ΤΜΗΜΑΤΟΣ
    If Trim(co_kt.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Κατηγορίες_Τμημάτων.περιγραφή] LIKE '" & Trim(co_kt.Text) & "'"
        Else
            s_string = "[Κατηγορίες_Τμημάτων.περιγραφή] LIKE '" & Trim(co_kt.Text) & "'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΠΡΟΠΟΝΗΤΗ Α
    If Trim(co_propa.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΟΕΑ] LIKE '*" & Trim(co_propa.Text) & "*'"
        Else
            s_string = "[ΟΕΑ] LIKE '*" & Trim(co_propa.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΠΡΟΠΟΝΗΤΗ Β
    If Trim(co_propb.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΟΕΒ] LIKE '*" & Trim(co_propb.Text) & "*'"
        Else
            s_string = "[ΟΕΒ] LIKE '*" & Trim(co_propb.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΠΟΣΟΥ ΕΓΓΡΑΦΗΣ
    If Trim(tmp_poso_eggrafis_plain.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [PE] LIKE '" & Trim(tmp_poso_eggrafis_plain.Text) & "'"
        Else
            s_string = "[PE] LIKE '" & Trim(tmp_poso_eggrafis_plain.Text) & "'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΠΟΣΟΥ ΜΗΝΙΑΙΑΣ ΣΥΝΔΡΟΜΗΣ
    If Trim(tmp_poso_mina_plain.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [PMS] LIKE '" & Trim(tmp_poso_mina_plain.Text) & "'"
        Else
            s_string = "[PMS] LIKE '" & Trim(tmp_poso_mina_plain.Text) & "'"
        End If
    End If
    '
    If s_string <> "" Then
        Me.ado_tmimata.Recordset.Filter = Trim(s_string)
    End If
    If s_sort <> "" Then
        Me.ado_tmimata.Recordset.Sort = Trim(s_sort)
    End If
    
    Me.canc_bt.Enabled = True
    Me.up_bt.Enabled = True
    Me.del_bt.Enabled = True
        
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

Private Sub taksin_Click()
  
    If Me.dt_tmimata.Col >= 0 Then
        Me.ado_tmimata.Recordset.Sort = "[" & Trim(Me.ado_tmimata.Recordset.Fields(dt_tmimata.Col).Name) & "]"
        s_sort = "[" & Trim(Me.ado_tmimata.Recordset.Fields(dt_tmimata.Col).Name) & "]"
    Else
        Me.ado_tmimata.Recordset.Sort = "[" & Trim(Me.ado_tmimata.Recordset.Fields(defined_col).Name) & "]"
        s_sort = "[" & Trim(Me.ado_tmimata.Recordset.Fields(defined_col).Name) & "]"
    End If
    
End Sub

Private Sub Text1_Change()
    
    'Me.bt_ins_athl.Enabled = False
    'Me.bt_del_athl.Enabled = False
    'Me.bt_up_athl.Enabled = True
    'Me.bt_can_athl.Enabled = True
    
End Sub

Private Sub tmp_poso_eggrafis_GotFocus()
    
    Me.tmp_poso_eggrafis.SelStart = 0
    Me.tmp_poso_eggrafis.SelLength = Len(Me.tmp_poso_eggrafis.Text)

End Sub

Private Sub tmp_poso_eggrafis_LostFocus()
        
    If tmp_poso_eggrafis.Text <> "" Then
        tmp_poso_eggrafis_plain.Text = Val(tmp_poso_eggrafis.Text)
        'If Trim(Me.tmp_poso_eggrafis_plain.Text) <> "0" And for_search = 0 Then
        If Trim(Me.tmp_poso_eggrafis_plain.Text) <> "0" Then
            If FormatCurrency(Val(Me.tmp_poso_eggrafis_plain), 2, vbFalse) = False Then
                tmp_poso_eggrafis.SelStart = 0
                tmp_poso_eggrafis.SelLength = Len(tmp_poso_eggrafis.Text)
                'tmp_poso_eggrafis.SelText = "00/00/0000"
                tmp_poso_eggrafis.SelText = ""
                tmp_poso_eggrafis_plain.Text = ""
                tmp_poso_eggrafis.SetFocus
            Else
                'tmp_poso_eggrafis_plain.Text = Trim(tmp_poso_eggrafis.Text)
                Me.tmp_poso_eggrafis.Text = Me.tmp_poso_eggrafis_plain.Text + ",00 €"
            End If
        Else
            If Trim(Me.tmp_poso_eggrafis_plain.Text) = "0" Then
                tmp_poso_eggrafis_plain.Text = 0
                Me.tmp_poso_eggrafis.Text = "0,00 €"
            End If
        End If
    Else
        tmp_poso_eggrafis_plain.Text = ""
    End If
    
End Sub

Private Sub tmp_poso_mina_GotFocus()
        
    Me.tmp_poso_mina.SelStart = 0
    Me.tmp_poso_mina.SelLength = Len(Me.tmp_poso_mina.Text)
    
End Sub

Private Sub tmp_poso_mina_LostFocus()
        
    If Me.tmp_poso_mina.Text <> "" Then
        tmp_poso_mina_plain.Text = Val(Me.tmp_poso_mina.Text)
        If Me.tmp_poso_mina_plain.Text <> "0" And for_search = 0 Then
            If FormatCurrency(Val(Me.tmp_poso_mina_plain.Text)) = False Then
                tmp_poso_mina.Text = ""
                tmp_poso_mina_plain.Text = ""
                tmp_poso_mina.SetFocus
            Else
                'tmp_poso_mina_plain.Text = Trim(tmp_poso_mina.Text)
                Me.tmp_poso_mina.Text = Me.tmp_poso_mina_plain.Text + ",00 €"
            End If
        Else
            If Trim(Me.tmp_poso_mina_plain.Text) = "0" Then
                tmp_poso_mina_plain.Text = 0
                Me.tmp_poso_mina.Text = "0,00 €"
            End If
        End If
    Else
        tmp_poso_mina_plain.Text = ""
    End If
    
End Sub

Private Sub txt_im_eis_Change()
    
    'Me.bt_ins_athl.Enabled = False
    'Me.bt_del_athl.Enabled = False
    'Me.bt_up_athl.Enabled = True
    'Me.bt_can_athl.Enabled = True

End Sub

Private Sub txt_im_eis_GotFocus()
    
    'Me.bt_ins_athl.Enabled = False
    'Me.bt_del_athl.Enabled = False
    'Me.bt_up_athl.Enabled = True
    'Me.bt_can_athl.Enabled = True
    
    at = Me.txt_im_eis.Text
    txt_im_eis.SelStart = 0
    txt_im_eis.SelLength = Len(txt_im_eis.Text)


End Sub

Private Sub txt_im_eis_LostFocus()
    
    tt = Me.txt_im_eis.Text
    If IsDate(Me.txt_im_eis.Text) = False And Me.txt_im_eis.Text <> "00/00/0000" Then
        Me.txt_im_eis.SelStart = 0
        Me.txt_im_eis.SelLength = 10
        Me.txt_im_eis.SelText = "00/00/0000"
        Me.txt_im_eis.SetFocus
    Else
        If Trim(at) <> Trim(tt) Then
            Me.bt_ins_athl.Enabled = True
            Me.bt_del_athl.Enabled = True
        End If
    End If

End Sub

Private Sub txt_pe_Change()
    
    Me.bt_ins_athl.Enabled = False
    Me.bt_del_athl.Enabled = False
    Me.bt_up_athl.Enabled = True
    Me.bt_can_athl.Enabled = True

End Sub

Private Sub txt_pe_GotFocus()
    
    'Me.bt_ins_athl.Enabled = False
    'Me.bt_del_athl.Enabled = False
    'Me.bt_up_athl.Enabled = True
    'Me.bt_can_athl.Enabled = True
    
    at = Me.txt_pe.Text
    
    txt_pe.SelStart = 0
    txt_pe.SelLength = Len(txt_pe.Text)

End Sub

Private Sub txt_pe_LostFocus()
    
    tt = Me.txt_pe.Text
    If Val(Me.txt_pe.Text) = 0 And Trim(Me.txt_pe.Text) <> 0 And Trim(Me.txt_pe.Text) <> "" Then
        Me.txt_pe.Text = ""
        Me.txt_pe.SetFocus
    Else
        If Trim(at) <> Trim(tt) Then
            'Me.bt_ins_athl.Enabled = True
            'Me.bt_del_athl.Enabled = True
            txt_pe.Text = Trim(str(Val(txt_pe.Text)) & ",00 €")
        End If
    End If
    
End Sub

Private Sub txt_pm_Change()
    
    Me.bt_ins_athl.Enabled = False
    Me.bt_del_athl.Enabled = False
    Me.bt_up_athl.Enabled = True
    Me.bt_can_athl.Enabled = True

End Sub

Private Sub txt_pm_GotFocus()
    
    'Me.bt_ins_athl.Enabled = False
    'Me.bt_del_athl.Enabled = False
    'Me.bt_up_athl.Enabled = True
    'Me.bt_can_athl.Enabled = True
    
    at = Me.txt_pm.Text
    
    txt_pm.SelStart = 0
    txt_pm.SelLength = Len(txt_pm.Text)

End Sub

Private Sub txt_pm_LostFocus()
    
    tt = Me.txt_pm.Text
    If Val(Me.txt_pm.Text) = 0 And Trim(Me.txt_pm.Text) <> 0 And Trim(Me.txt_pm.Text) <> "" Then
        Me.txt_pm.Text = ""
        Me.txt_pm.SetFocus
    Else
        If Trim(at) <> Trim(tt) Then
            txt_pm.Text = Trim(str(Val(txt_pm.Text)) & ",00 €")
        End If
    End If
    
End Sub

Private Sub up_bt_Click()

    'On Error GoTo up_bt_Click_l


    Dim id As Integer
    
    If Not Me.ado_tmimata.Recordset.EOF And Not Me.ado_tm_ids.Recordset.EOF Then
        tm_id = Me.ado_tmimata.Recordset.Fields(10).Value
        Me.ado_tm_ids.Recordset.MoveFirst
        For i = 0 To Me.ado_tmimata.Recordset.RecordCount - 1
            If Me.ado_tm_ids.Recordset.Fields(0).Value <> tm_id Then
                Me.ado_tm_ids.Recordset.MoveNext
            Else
                i = Me.ado_tmimata.Recordset.RecordCount
            End If
        Next i
        If Me.co_ae.SelectedItem >= 1 Then
            Me.ado_ae.Recordset.MoveFirst
            Me.ado_ae.Recordset.Move Me.co_ae.SelectedItem - 1
            Me.ado_tm_ids.Recordset.Fields(1).Value = Me.ado_ae.Recordset.Fields(0).Value
        End If
        If Me.co_kt.SelectedItem >= 1 Then
            rs_ado_kt.MoveFirst
            rs_ado_kt.Move Me.co_kt.SelectedItem - 1
            Me.ado_tm_ids.Recordset.Fields(2).Value = rs_ado_kt.Fields(0).Value
        End If
        If Me.co_propa.SelectedItem >= 1 Then
            rs_ado_prop.MoveFirst
            rs_ado_prop.Move Me.co_propa.SelectedItem - 1
            Me.ado_tm_ids.Recordset.Fields(3).Value = rs_ado_prop.Fields(0).Value
        End If
        If Me.co_propb.SelectedItem >= 1 Then
            rs_ado_prop.MoveFirst
            rs_ado_prop.Move Me.co_propb.SelectedItem - 1
            Me.ado_tm_ids.Recordset.Fields(4).Value = rs_ado_prop.Fields(0).Value
        Else
            Me.ado_tm_ids.Recordset.Fields(4).Value = 0
        End If
        If Trim(Me.tmp_poso_eggrafis.Text) <> "" Then
            Me.ado_tm_ids.Recordset.Fields(5).Value = Me.tmp_poso_eggrafis_plain.Text
        End If
        If Trim(Me.tmp_poso_mina.Text) <> "" Then
            Me.ado_tm_ids.Recordset.Fields(6).Value = Me.tmp_poso_mina_plain.Text
        End If
        'ΕΝΗΜΕΡΩΣΗ ΤΟΥ ΕΒΔΟΜΑΔΙΑΙΟΥ ΠΡΟΓΡΑΜΜΑΤΟΣ
        i = 0
        For j = 0 To 5
            If Trim(Me.im_en(j).Text) <> "" Then
                Me.ado_tm_ids.Recordset.Fields(7 + i).Value = Me.im_en(j).Text
            Else
                Me.ado_tm_ids.Recordset.Fields(7 + i).Value = ""
            End If
            If Trim(Me.im_l(j).Text) <> "" Then
                Me.ado_tm_ids.Recordset.Fields(8 + i).Value = Me.im_l(j).Text
            Else
                Me.ado_tm_ids.Recordset.Fields(8 + i).Value = ""
            End If
            i = i + 2
        Next j
        '
        If Me.ado_tmimata.Recordset.AbsolutePosition >= 1 Then
            col_aff = Me.ado_tmimata.Recordset.AbsolutePosition
        Else
            col_aff = 1
        End If
        Me.ado_tm_ids.Recordset.UpdateBatch adAffectCurrent
        Me.ado_tm_ids.Recordset.Requery
        Me.ado_tm_ids.Refresh
        Me.DataGrid2.Refresh
        Set Me.DataGrid2.DataSource = Me.ado_tm_ids
        
        Me.ado_tmimata.Recordset.Requery
        Me.ado_tmimata.Refresh
        Me.dt_tmimata.Refresh
        If Not Me.ado_tmimata.Recordset.EOF Then
            Me.ado_tmimata.Recordset.Sort = "[" & Trim(Me.ado_tmimata.Recordset.Fields(1).Name) & "] DESC, [" & Trim(Me.ado_tmimata.Recordset.Fields(3).Name) & "] ASC"
            Me.ado_tmimata.Recordset.Find "[TID] LIKE " & tm_id, , adSearchForward, 1
            Me.dt_tmimata.SelBookmarks.Add Me.ado_tmimata.Recordset.Bookmark
        End If
        If Me.ado_tmimata.Recordset.RecordCount > 0 Then
            Me.ado_tmimata.Caption = "Τμήμα " & Me.ado_tmimata.Recordset.AbsolutePosition & " από " & Me.ado_tmimata.Recordset.RecordCount
        End If
        Me.dt_tmimata.Columns(0).Visible = False
        Me.dt_tmimata.Columns(1).Caption = "Αθλητικό Έτος"
        Me.dt_tmimata.Columns(1).Width = 1500
        Me.dt_tmimata.Columns(2).Visible = False
        Me.dt_tmimata.Columns(3).Caption = "Κατηγορία Τμήματος"
        Me.dt_tmimata.Columns(3).Width = 2000
        Me.dt_tmimata.Columns(4).Visible = False
        Me.dt_tmimata.Columns(5).Caption = "Α Προπονητής"
        Me.dt_tmimata.Columns(5).Width = 2500
        Me.dt_tmimata.Columns(6).Visible = False
        Me.dt_tmimata.Columns(7).Caption = "Β Προπονητής"
        Me.dt_tmimata.Columns(7).Width = 2500
        For i = 8 To Me.ado_tmimata.Recordset.Fields.Count - 1
            Me.dt_tmimata.Columns(i).Visible = False
       Next i
    Else
        MsgBox "Ακύρωση Ενημέρωσης! Απαιτείται πρώτα προσθήκη εγγραφής!", vbCritical, "Μήνυμα Λάθους!"
        Me.co_ae.SetFocus
    End If
    '
'up_bt_Click_l:
    
End Sub
