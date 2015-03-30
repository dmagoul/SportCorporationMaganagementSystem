VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form meli_management 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Διαχείριση Μελών"
   ClientHeight    =   10830
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11775
   ForeColor       =   &H000000FF&
   LinkTopic       =   "entity_management"
   MDIChild        =   -1  'True
   ScaleHeight     =   10830
   ScaleWidth      =   11775
   Begin MSAdodcLib.Adodc ado_dimoi 
      Height          =   375
      Left            =   9600
      Top             =   360
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
      RecordSource    =   "Δήμοι_ταξινομημένοι"
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
   Begin MSAdodcLib.Adodc ado_pe 
      Height          =   375
      Left            =   9600
      Top             =   720
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
      RecordSource    =   "ΠεριφερειακέςΕνότητες_ταξινομημένες"
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
   Begin MSAdodcLib.Adodc ado_jobs 
      Height          =   375
      Left            =   9600
      Top             =   1080
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
      RecordSource    =   "Επαγγέλματα_ταξινομημένα"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Όλα τα Μέλη"
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
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   11535
      Begin MSDataGridLib.DataGrid tmp_dt_meli 
         Bindings        =   "meli_management.frx":0000
         Height          =   735
         Left            =   8520
         TabIndex        =   75
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1296
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin MSAdodcLib.Adodc tmp_ado_meli 
         Height          =   375
         Left            =   8520
         Top             =   3480
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
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
         RecordSource    =   "Μέλη"
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
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000014&
         Caption         =   "Ε&ξαγωγή Επιλεγμένων Μελών στο Excel"
         DisabledPicture =   "meli_management.frx":001B
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         MaskColor       =   &H80000014&
         Picture         =   "meli_management.frx":26A6
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   4080
         Width           =   2895
      End
      Begin VB.CommandButton bt_print 
         BackColor       =   &H80000014&
         Caption         =   "Εκτύπω&ση Επιλεγμένων Μελών"
         DisabledPicture =   "meli_management.frx":4D31
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         MaskColor       =   &H80000014&
         Picture         =   "meli_management.frx":56A4
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   4080
         Width           =   2895
      End
      Begin VB.CommandButton taksin 
         BackColor       =   &H80000014&
         Caption         =   "&Ταξινόμηση"
         DisabledPicture =   "meli_management.frx":A0A9
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MaskColor       =   &H80000014&
         Picture         =   "meli_management.frx":AA1C
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton met_st 
         BackColor       =   &H80000014&
         Caption         =   "&Μεταφορά Στοιχείων Μέλους"
         DisabledPicture =   "meli_management.frx":F723
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8760
         MaskColor       =   &H80000014&
         Picture         =   "meli_management.frx":143AC
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4080
         Width           =   2655
      End
      Begin MSDataGridLib.DataGrid dt_meli 
         Bindings        =   "meli_management.frx":19035
         Height          =   3255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin MSAdodcLib.Adodc ado_meli 
         Height          =   375
         Left            =   120
         Top             =   3600
         Width           =   11320
         _ExtentX        =   19976
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
         RecordSource    =   "Μέλη"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Στοιχεία Μέλους"
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
      Height          =   6135
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   11535
      Begin VB.CommandButton cancel_cur_rec 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Cancel          =   -1  'True
         Caption         =   "Ακύρωση Τρέχουσας Ε&γγραφής"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8640
         MaskColor       =   &H80000014&
         Picture         =   "meli_management.frx":1904C
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Ακύρωση Τρέχουσας Εγγραφής"
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton insert_bt 
         BackColor       =   &H80000014&
         Caption         =   "Προσ&θήκη"
         DisabledPicture =   "meli_management.frx":1920B
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MaskColor       =   &H80000014&
         Picture         =   "meli_management.frx":1DF9A
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton save_command 
         BackColor       =   &H80000014&
         Caption         =   "&Αποθήκευση"
         DisabledPicture =   "meli_management.frx":22D29
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         MaskColor       =   &H80000014&
         Picture         =   "meli_management.frx":27928
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton canc_bt 
         BackColor       =   &H80000014&
         Caption         =   "Ακύ&ρωση"
         DisabledPicture =   "meli_management.frx":2C527
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7080
         MaskColor       =   &H80000014&
         Picture         =   "meli_management.frx":31444
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton kl_bt 
         BackColor       =   &H80000014&
         Caption         =   "Κ&λείσιμο"
         DisabledPicture =   "meli_management.frx":36361
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10440
         MaskColor       =   &H80000014&
         Picture         =   "meli_management.frx":36CD4
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton up_bt 
         BackColor       =   &H80000014&
         Caption         =   "&Ενημέρωση"
         Default         =   -1  'True
         DisabledPicture =   "meli_management.frx":3C74C
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         MaskColor       =   &H80000014&
         Picture         =   "meli_management.frx":44246
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000014&
         Caption         =   "&Καθαρισμός"
         DisabledPicture =   "meli_management.frx":4BD40
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4680
         MaskColor       =   &H80000014&
         Picture         =   "meli_management.frx":4BE88
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton del_bt 
         BackColor       =   &H80000014&
         Caption         =   "&Διαγραφή"
         DisabledPicture =   "meli_management.frx":509CC
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         MaskColor       =   &H80000014&
         Picture         =   "meli_management.frx":557D0
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton sear_bt 
         BackColor       =   &H80000014&
         Caption         =   "Α&ναζήτηση"
         DisabledPicture =   "meli_management.frx":5A5D4
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5640
         MaskColor       =   &H80000014&
         Picture         =   "meli_management.frx":5F988
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   5160
         Width           =   975
      End
      Begin VB.Frame Frame5 
         Caption         =   "Στοιχεία Διεύθυνσης"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2100
         Left            =   4800
         TabIndex        =   37
         Top             =   3000
         Width           =   4695
         Begin VB.TextBox tmp_tk 
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
            Left            =   3120
            TabIndex        =   20
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox tmp_perioxi 
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
            Left            =   1200
            TabIndex        =   21
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox tmp_arithmos 
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
            Left            =   1200
            TabIndex        =   19
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox tmp_odos 
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
            Left            =   1200
            TabIndex        =   18
            Top             =   240
            Width           =   3015
         End
         Begin MSDataListLib.DataCombo co_pe 
            Bindings        =   "meli_management.frx":64D3C
            Height          =   345
            Left            =   1200
            TabIndex        =   23
            Top             =   1635
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   609
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "περιγραφή"
            BoundColumn     =   "id_πε"
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
         Begin MSDataListLib.DataCombo co_dimoi 
            Bindings        =   "meli_management.frx":64D51
            Height          =   345
            Left            =   1200
            TabIndex        =   22
            Top             =   1320
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   609
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "περιγραφή"
            BoundColumn     =   "id_δήμου"
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
            Caption         =   "Οδός"
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
            TabIndex        =   43
            Top             =   240
            Width           =   945
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Αριθμός"
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
            Top             =   600
            Width           =   945
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Περιοχή"
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
            TabIndex        =   41
            Top             =   960
            Width           =   945
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Δήμος"
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
            Left            =   240
            TabIndex        =   40
            Top             =   1320
            Width           =   825
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Π. Ε."
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
            Left            =   360
            TabIndex        =   39
            Top             =   1680
            Width           =   705
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Τ.Κ."
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
            Left            =   2040
            TabIndex        =   38
            Top             =   600
            Width           =   1065
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Ιδιότητες Μελών"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   4800
         TabIndex        =   52
         Top             =   1680
         Width           =   4695
         Begin VB.CheckBox ch_tm_en 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3600
            TabIndex        =   76
            Top             =   800
            Width           =   210
         End
         Begin VB.CheckBox ch_dr 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   73
            Top             =   800
            Width           =   210
         End
         Begin VB.CheckBox ch_g 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3600
            TabIndex        =   68
            Top             =   550
            Width           =   210
         End
         Begin VB.CheckBox ch_m_ds 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   15
            Top             =   240
            Width           =   210
         End
         Begin VB.CheckBox ch_e 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3600
            TabIndex        =   17
            Top             =   240
            Width           =   210
         End
         Begin VB.CheckBox ch_p 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   16
            Top             =   550
            Width           =   210
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Τμ. Ενηλίκων"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2040
            TabIndex        =   77
            Top             =   800
            Width           =   1485
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Δρομέας"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   800
            Width           =   1485
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Γονέας"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2040
            TabIndex        =   69
            Top             =   550
            Width           =   1485
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Μέλος ΔΣ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   285
            Width           =   1485
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Εθελοντής"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2040
            TabIndex        =   54
            Top             =   285
            Width           =   1485
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Προπονητής"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   550
            Width           =   1485
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Μέλος Σωματείου"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   4800
         TabIndex        =   45
         Top             =   360
         Width           =   4695
         Begin VB.CheckBox ch_m_gs 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   11
            Top             =   360
            Width           =   210
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3480
            TabIndex        =   14
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox tmp_am 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3480
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin MSMask.MaskEdBox im_eg 
            Height          =   375
            Left            =   1320
            TabIndex        =   13
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   16777215
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
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Μέλος ΓΣ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Αριθμός Απόδειξης"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   2400
            TabIndex        =   51
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Αριθμός Μητρώου"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   2400
            TabIndex        =   47
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ημερομηνία Εγγραφής"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   0
            TabIndex        =   46
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Στοιχεία Επικοινωνίας"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   120
         TabIndex        =   32
         Top             =   3000
         Width           =   4575
         Begin VB.TextBox tmp_email 
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
            Left            =   1440
            TabIndex        =   10
            Top             =   1320
            Width           =   3015
         End
         Begin VB.TextBox tmp_fax 
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
            Left            =   1440
            TabIndex        =   9
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox tmp_kinito 
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
            Left            =   1440
            TabIndex        =   8
            Top             =   600
            Width           =   3015
         End
         Begin VB.TextBox tmp_til_oikias 
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
            Left            =   1440
            TabIndex        =   7
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Τηλ.οικίας"
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
            Left            =   -240
            TabIndex        =   36
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητό"
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
            Left            =   -120
            TabIndex        =   35
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
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
            Left            =   -120
            TabIndex        =   34
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
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
            Left            =   -120
            TabIndex        =   33
            Top             =   1320
            Width           =   1455
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
         Height          =   2730
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   4575
         Begin VB.TextBox tmp_eponimo 
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
            Left            =   1440
            TabIndex        =   2
            Top             =   1080
            Width           =   2655
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            Caption         =   "..."
            Enabled         =   0   'False
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
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   1080
            Width           =   375
         End
         Begin VB.CheckBox ch_en 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4200
            TabIndex        =   55
            Top             =   360
            Width           =   210
         End
         Begin VB.TextBox txt_ea_adt 
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
            Left            =   2520
            TabIndex        =   4
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox txt_adt 
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
            Left            =   1440
            TabIndex        =   3
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox txt_kod 
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox tmp_onoma 
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
            Left            =   1440
            TabIndex        =   1
            Top             =   720
            Width           =   3015
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   375
            Left            =   1440
            TabIndex        =   5
            Top             =   1800
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
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
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo co_jobs 
            Bindings        =   "meli_management.frx":64D69
            Height          =   345
            Left            =   1440
            TabIndex        =   6
            Top             =   2160
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   609
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "περιγραφή"
            BoundColumn     =   "id_επαγγέλματος"
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
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Caption         =   "Ενεργό"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3360
            TabIndex        =   56
            Top             =   405
            Width           =   765
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ΑΔΤ / Εκδ. Αρχή"
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
            TabIndex        =   50
            Top             =   1440
            Width           =   1365
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Κωδικός"
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
            Left            =   360
            TabIndex        =   49
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Επάγγελμα"
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
            Left            =   480
            TabIndex        =   44
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ημ/νία Γέννησης"
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
            Left            =   -120
            TabIndex        =   31
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Όνομα"
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
            Left            =   360
            TabIndex        =   30
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Επώνυμο"
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
            Left            =   360
            TabIndex        =   29
            Top             =   1080
            Width           =   1005
         End
      End
   End
End
Attribute VB_Name = "meli_management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public it_is_a_new_record As Integer
Public initial_string As String
Public final_string As String
Public id_α, defined_col As Integer
Public rs_ado_dimoi As ADODB.Recordset
Public rs_ado_jobs As ADODB.Recordset
Public rs_ado_pe As ADODB.Recordset
Public for_search As Integer
Public rs_meli As ADODB.Recordset

Private Sub ado_meli_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    If Not adReason = 10 And pRecordset.AbsolutePosition >= 1 Then
        If Trim(Me.ado_meli.Recordset.Fields(0).Value) <> "" And Me.it_is_a_new_record = 0 Then
            txt_kod.Text = Me.ado_meli.Recordset.Fields(0).Value
        Else
            If Me.it_is_a_new_record = 1 Then
                '
            Else
                txt_kod.Text = ""
            End If
        End If
        If Trim(pRecordset.Fields(1).Value) <> "" Then
            tmp_am.Text = pRecordset.Fields(1).Value
        Else
            tmp_am.Text = ""
        End If
        If Trim(pRecordset.Fields(3).Value) <> "" And Me.it_is_a_new_record = 0 Then
            tmp_onoma.Text = pRecordset.Fields(3).Value
        Else
            If Me.it_is_a_new_record = 1 Then
                '
            Else
                tmp_onoma.Text = ""
            End If
        End If
        If Trim(pRecordset.Fields(2).Value) <> "" Then
            tmp_eponimo.Text = pRecordset.Fields(2).Value
        Else
            tmp_eponimo.Text = ""
        End If
        If Trim(pRecordset.Fields(4).Value) <> "" Then
            tmp_odos.Text = pRecordset.Fields(4).Value
        Else
            tmp_odos.Text = ""
        End If
        If Trim(pRecordset.Fields(5).Value) <> "" Then
            tmp_arithmos.Text = pRecordset.Fields(5).Value
        Else
            tmp_arithmos.Text = ""
        End If
        If Trim(pRecordset.Fields(6).Value) <> "" Then
            tmp_perioxi.Text = pRecordset.Fields(6).Value
        Else
            tmp_perioxi.Text = ""
        End If
        If Trim(pRecordset.Fields(8).Value) <> "" Then
            Me.co_pe.Text = pRecordset.Fields(8).Value
        Else
            Me.co_pe.Text = ""
        End If
        If Trim(pRecordset.Fields(7).Value) <> "" Then
            Me.co_dimoi.Text = pRecordset.Fields(7).Value
        Else
            Me.co_dimoi.Text = ""
        End If
        If Trim(pRecordset.Fields(9).Value) <> "" Then
            tmp_tk.Text = pRecordset.Fields(9).Value
        Else
            tmp_tk.Text = ""
        End If
        If Trim(pRecordset.Fields(10).Value) <> "" Then
            tmp_til_oikias.Text = pRecordset.Fields(10).Value
        Else
            tmp_til_oikias.Text = ""
        End If
        If Trim(pRecordset.Fields(11).Value) <> "" Then
            tmp_kinito.Text = pRecordset.Fields(11).Value
        Else
            tmp_kinito.Text = ""
        End If
        If Trim(pRecordset.Fields(12).Value) <> "" Then
            tmp_fax.Text = pRecordset.Fields(12).Value
        Else
            tmp_fax.Text = ""
        End If
        If Trim(pRecordset.Fields(13).Value) <> "" Then
            tmp_email.Text = pRecordset.Fields(13).Value
        Else
            tmp_email.Text = ""
        End If
        If Trim(pRecordset.Fields(14).Value) <> "" Then
            Me.MaskEdBox1.Text = pRecordset.Fields(14).Value
        Else
            Me.MaskEdBox1.Text = "00/00/0000"
        End If
        If pRecordset.Fields(15).Value = True Then
            Me.ch_en.Value = 1
        Else
            Me.ch_en.Value = 0
        End If
        If Trim(pRecordset.Fields(16).Value) <> "" Then
            Me.im_eg.Text = pRecordset.Fields(16).Value
        Else
            Me.im_eg.Text = "00/00/0000"
        End If
        If Trim(pRecordset.Fields(17).Value) <> "" Then
            Me.co_jobs.Text = pRecordset.Fields(17).Value
        Else
            Me.co_jobs.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(18).Value) <> "" Then
            Me.txt_adt.Text = Me.ado_meli.Recordset.Fields(18).Value
        Else
            Me.txt_adt.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(19).Value) <> "" Then
            Me.txt_ea_adt.Text = Me.ado_meli.Recordset.Fields(19).Value
        Else
            Me.txt_ea_adt.Text = ""
        End If
        If Me.ado_meli.Recordset.Fields(20).Value = True Then
            Me.ch_m_gs.Value = 1
        Else
            Me.ch_m_gs.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields(21).Value = True Then
            Me.ch_m_ds.Value = 1
        Else
            Me.ch_m_ds.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields(22).Value = True Then
            Me.ch_p.Value = 1
        Else
            Me.ch_p.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields(23).Value = True Then
            Me.ch_e.Value = 1
        Else
            Me.ch_e.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields(25).Value = True Then
            Me.ch_g.Value = 1
        Else
            Me.ch_g.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields("Δρομέας").Value = True Then
            Me.ch_dr.Value = 1
        Else
            Me.ch_dr.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields("Τμήμα_Ενηλίκων").Value = True Then
            Me.ch_tm_en.Value = 1
        Else
            Me.ch_tm_en.Value = 0
        End If
    End If
    If pRecordset.RecordCount > 0 Then
        Me.ado_meli.Caption = "Μέλος " & pRecordset.AbsolutePosition & " από " & pRecordset.RecordCount
    Else
        Me.ado_meli.Caption = "Μέλος " & 0 & " από " & 0
    End If
    Set rs_meli = pRecordset
    
End Sub

Private Sub bt_print_Click()
    
    If MDIForm1.s_string <> "" Then
        Poseidon_DB.rsσυνοπτική_εκτύπωση_επιλεγμένων_μελών.Filter = MDIForm1.s_string
    Else
        Poseidon_DB.rsσυνοπτική_εκτύπωση_επιλεγμένων_μελών.Filter = ""
    End If
    If MDIForm1.s_sort <> "" Then
        Poseidon_DB.rsσυνοπτική_εκτύπωση_επιλεγμένων_μελών.Sort = MDIForm1.s_sort
    Else
        Poseidon_DB.rsσυνοπτική_εκτύπωση_επιλεγμένων_μελών.Sort = "[id]"
    End If
    Rep_συνοπτική_εκτύπωση_επιλεγμένων_μελών.Sections("ReportHeader").Controls("Label12").Caption = MDIForm1.rep_lbl
    
    Rep_συνοπτική_εκτύπωση_επιλεγμένων_μελών.Orientation = rptOrientPortrait
    Rep_συνοπτική_εκτύπωση_επιλεγμένων_μελών.Show

End Sub

Private Sub canc_bt_Click()

    Me.it_is_a_new_record = 0
    Me.Command2.Enabled = False
    
    Me.txt_kod.Locked = True
    
    Me.ado_meli.Refresh
    If Me.ado_meli.Recordset.RecordCount >= 1 Then
        Me.insert_bt.Enabled = True
        Me.up_bt.Enabled = True
        Me.del_bt.Enabled = True
        Me.sear_bt.Enabled = False
        s_sort = ""
        s_string = ""
        rep_lbl = ""
        MDIForm1.s_sort = ""
        MDIForm1.s_string = ""
        MDIForm1.rep_lbl = ""
        for_search = 0
    
        Me.dt_meli.Refresh
        Me.ado_meli.Refresh
    
        Me.ado_meli.Recordset.Sort = "[id]"
        Me.ado_meli.Recordset.MoveFirst
        '***************************************************************

        Me.dt_meli.Columns(0).Caption = "Κωδικός"
        Me.dt_meli.Columns(0).Width = 1500
        Me.dt_meli.Columns(1).Visible = False
        Me.dt_meli.Columns(2).Caption = "Επώνυμο"
        Me.dt_meli.Columns(2).Width = 3000
        Me.dt_meli.Columns(3).Caption = "Όνομα"
        Me.dt_meli.Columns(3).Width = 2000
        For i = 4 To 16
            Me.dt_meli.Columns(i).Visible = False
        Next i
        Me.dt_meli.Columns(17).Caption = "Επάγγελμα"
        Me.dt_meli.Columns(17).Width = 2000
        For i = 18 To Me.ado_meli.Recordset.Fields.Count - 1
            Me.dt_meli.Columns(i).Visible = False
        Next i
    
        Me.tmp_onoma.SetFocus
        Me.Command1.Enabled = True
        If athlet_management.flag_mitera = 1 Or athlet_management.flag_pateras = 1 Then
            Me.met_st.Enabled = True
        End If
        Me.canc_bt.Enabled = False
        Me.save_command.Enabled = False
        Me.cancel_cur_rec.Enabled = True
    End If

End Sub

Private Sub cancel_cur_rec_Click()

    Dim id_mel As Integer
    Dim f_s As String
    
    Me.it_is_a_new_record = 0
    If Me.ado_meli.Recordset.RecordCount >= 1 Then
        id_mel = Me.txt_kod
        Me.ado_meli.Recordset.MoveFirst
        Me.ado_meli.Recordset.Find "[id] like '" & str(id_mel) & "'", , adSearchForward
        If Not Me.ado_meli.Recordset.EOF Then 'ΤΟ ΕΧΕΙΣ ΒΡΕΙ
            Me.tmp_onoma.SetFocus
            Me.canc_bt.Enabled = True
        Else 'ΔΕΝ ΤΟ ΒΡΕΙΣ
            Me.ado_meli.Recordset.Requery
            Me.ado_meli.Refresh
            Me.ado_meli.Refresh
            Me.ado_meli.Recordset.MoveFirst
            Me.dt_meli.Columns(0).Caption = "Κωδικός"
            Me.dt_meli.Columns(0).Width = 1500
            Me.dt_meli.Columns(1).Visible = False
            Me.dt_meli.Columns(2).Caption = "Επώνυμο"
            Me.dt_meli.Columns(2).Width = 3000
            Me.dt_meli.Columns(3).Caption = "Όνομα"
            Me.dt_meli.Columns(3).Width = 2000
            For i = 4 To 16
                Me.dt_meli.Columns(i).Visible = False
            Next i
            Me.dt_meli.Columns(17).Caption = "Επάγγελμα"
            Me.dt_meli.Columns(17).Width = 2000
            For i = 18 To Me.ado_meli.Recordset.Fields.Count - 1
                Me.dt_meli.Columns(i).Visible = False
            Next i
        End If
    Else
        '
    End If

End Sub

Private Sub co_dimoi_Click(Area As Integer)
    
    initial_string = Left(Me.co_dimoi.Text, Me.co_dimoi.SelStart)
    final_string = Right(Me.co_dimoi.Text, Len(Me.co_dimoi.Text) - Me.co_dimoi.SelStart)
    
End Sub

Private Sub co_dimoi_GotFocus()
    
    initial_string = ""
    final_string = ""
    If Me.co_dimoi.Text <> "" Then
        initial_string = Left(Me.co_dimoi.Text, Me.co_dimoi.SelStart)
        final_string = Right(Me.co_dimoi.Text, Len(Me.co_dimoi.Text) - Me.co_dimoi.SelStart)
    End If

End Sub

Private Sub co_dimoi_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyTab Then
        'Me.co_sxolia.BoundText = initial_string & final_string
        Me.co_dimoi.Text = initial_string & final_string
    End If

End Sub

Private Sub co_dimoi_KeyPress(KeyAscii As Integer)

    'SendKeys "{Esc}"
    'SendKeys "%{Up}"
    
    initial_string = Left(Me.co_dimoi.Text, Me.co_dimoi.SelStart)
    final_string = Right(Me.co_dimoi.Text, Len(Me.co_dimoi.Text) - Me.co_dimoi.SelStart)
    initial_string = initial_string & Chr(KeyAscii)
    
End Sub

Private Sub co_dimoi_KeyUp(KeyCode As Integer, Shift As Integer)

    initial_string = Left(Me.co_dimoi.Text, Me.co_dimoi.SelStart)
    final_string = Right(Me.co_dimoi.Text, Len(Me.co_dimoi.Text) - Me.co_dimoi.SelStart)
    
End Sub

Private Sub co_dimoi_LostFocus()

    'Me.co_dimoi.BoundText = initial_string & final_string
    Me.co_dimoi.Text = initial_string & final_string
    
End Sub

Private Sub co_jobs_Click(Area As Integer)
    
    initial_string = Left(Me.co_jobs.Text, Me.co_jobs.SelStart)
    final_string = Right(Me.co_jobs.Text, Len(Me.co_jobs.Text) - Me.co_jobs.SelStart)

End Sub

Private Sub co_jobs_GotFocus()
    
    initial_string = ""
    final_string = ""
    If Me.co_jobs.Text <> "" Then
        initial_string = Left(Me.co_jobs.Text, Me.co_jobs.SelStart)
        final_string = Right(Me.co_jobs.Text, Len(Me.co_jobs.Text) - Me.co_jobs.SelStart)
    End If

End Sub

Private Sub co_jobs_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyTab Then
        'Me.co_sxolia.BoundText = initial_string & final_string
        Me.co_jobs.Text = initial_string & final_string
    End If
    
End Sub

Private Sub co_jobs_KeyPress(KeyAscii As Integer)
    
    'SendKeys "{Esc}"
    'SendKeys "%{Up}"
    
    initial_string = Left(Me.co_jobs.Text, Me.co_jobs.SelStart)
    final_string = Right(Me.co_jobs.Text, Len(Me.co_jobs.Text) - Me.co_jobs.SelStart)
    initial_string = initial_string & Chr(KeyAscii)
    
End Sub

Private Sub co_jobs_KeyUp(KeyCode As Integer, Shift As Integer)

    initial_string = Left(Me.co_jobs.Text, Me.co_jobs.SelStart)
    final_string = Right(Me.co_jobs.Text, Len(Me.co_jobs.Text) - Me.co_jobs.SelStart)
    
End Sub

Private Sub co_jobs_LostFocus()
    
    'Me.co_jobs.BoundText = initial_string & final_string
    Me.co_jobs.Text = initial_string & final_string
    
End Sub

Private Sub co_pe_Click(Area As Integer)
    
    initial_string = Left(Me.co_pe.Text, Me.co_pe.SelStart)
    final_string = Right(Me.co_pe.Text, Len(Me.co_pe.Text) - Me.co_pe.SelStart)
    
End Sub

Private Sub co_pe_GotFocus()
    
    initial_string = ""
    final_string = ""
    If Me.co_pe.Text <> "" Then
        initial_string = Left(Me.co_pe.Text, Me.co_pe.SelStart)
        final_string = Right(Me.co_pe.Text, Len(Me.co_pe.Text) - Me.co_pe.SelStart)
    End If

End Sub

Private Sub co_pe_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyTab Then
        'Me.co_sxolia.BoundText = initial_string & final_string
        Me.co_pe.Text = initial_string & final_string
    End If

End Sub

Private Sub co_pe_KeyPress(KeyAscii As Integer)

    'SendKeys "{Esc}"
    'SendKeys "%{Up}"
    
    initial_string = Left(Me.co_pe.Text, Me.co_pe.SelStart)
    final_string = Right(Me.co_pe.Text, Len(Me.co_pe.Text) - Me.co_pe.SelStart)
    initial_string = initial_string & Chr(KeyAscii)
    
End Sub

Private Sub co_pe_KeyUp(KeyCode As Integer, Shift As Integer)

    initial_string = Left(Me.co_pe.Text, Me.co_pe.SelStart)
    final_string = Right(Me.co_pe.Text, Len(Me.co_pe.Text) - Me.co_pe.SelStart)
    
End Sub

Private Sub co_pe_LostFocus()

    'Me.co_pe.BoundText = initial_string & final_string
    Me.co_pe.Text = initial_string & final_string
    
End Sub

Private Sub Command1_Click()


    Me.up_bt.Enabled = False
    Me.del_bt.Enabled = False
    Me.cancel_cur_rec.Enabled = False
    
        for_search = 1
        txt_kod.Locked = False
        txt_kod.Text = ""
        Me.ch_en = 0
        tmp_onoma.Text = ""
        tmp_eponimo.Text = ""
        txt_adt = ""
        txt_ea_adt = ""
        Me.MaskEdBox1.Text = "  /  /    "
        Me.co_jobs.Text = ""
        tmp_til_oikias.Text = ""
        tmp_kinito.Text = ""
        tmp_fax.Text = ""
        tmp_email.Text = ""
        Me.ch_m_gs = 0
        Me.tmp_am.Text = ""
        Me.im_eg.Text = "  /  /    "
        Me.Text4.Text = ""
        Me.ch_m_ds = 0
        Me.ch_p = 0
        Me.ch_e = 0
        Me.ch_g = 0
        Me.ch_dr = 0
        Me.ch_tm_en = 0
        Me.tmp_odos.Text = ""
        Me.tmp_arithmos.Text = ""
        Me.tmp_tk.Text = ""
        tmp_perioxi.Text = ""
        Me.co_dimoi.Text = ""
        Me.co_pe.Text = ""
    
        Me.insert_bt.Enabled = False
        Me.up_bt.Enabled = False
        Me.save_command.Enabled = False
        Me.del_bt.Enabled = False
        Me.sear_bt.Enabled = True
        Me.canc_bt.Enabled = True
        
End Sub

Private Sub Command2_Click()
    
    Dim f_s As String
    
    If Me.ado_meli.Recordset.RecordCount >= 1 And Me.tmp_eponimo.Text <> "" And it_is_a_new_record = 1 Then
        f_s = "[Επώνυμο] LIKE '" & Trim(Me.tmp_eponimo.Text) & "*'"
        Me.ado_meli.Recordset.Filter = f_s
        If Not Me.ado_meli.Recordset.EOF Then 'ΤΟ ΕΧΕΙΣ ΒΡΕΙ
            Me.canc_bt.Enabled = True
            'Me.del_bt.Enabled = True
        Else 'ΔΕΝ ΤΟ ΕΧΕΙΣ ΒΡΕΙ
            'Me.ado_meli.Refresh
            'Me.dt_meli.Columns(0).Caption = "Κωδικός"
            'Me.dt_meli.Columns(0).Width = 1500
            'Me.dt_meli.Columns(1).Visible = False
            'Me.dt_meli.Columns(2).Caption = "Επώνυμο"
            'Me.dt_meli.Columns(2).Width = 4000
            'Me.dt_meli.Columns(3).Caption = "Όνομα"
            'Me.dt_meli.Columns(3).Width = 4000
            'For i = 4 To Me.ado_meli.Recordset.Fields.Count - 1
            '    Me.dt_meli.Columns(i).Visible = False
            'Next i
            'If Me.ado_meli.Recordset.RecordCount - 1 >= 0 Then
            '    Me.dt_meli.Row = Me.ado_meli.Recordset.RecordCount - 1
            'End If
        End If
    End If

End Sub

Private Sub Command3_Click()

   Dim oExcel As Object
   Dim oBook As Object
   Dim oSheet As Object
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add
   Set oSheet = oBook.Worksheets(1)
   
   Clipboard.Clear
   Dim sData As Variant
   sData = ""
   If Me.ado_meli.Recordset.RecordCount >= 1 Then
        Me.ado_meli.Recordset.MoveFirst
        sData = "ΑΜ Μέλους" & vbTab & "Επώνυμο" & vbTab & "Όνομα" & vbTab & "Τηλέφωνο" & vbTab & "Κινητό" & vbTab & "Επάγγελμα" & vbCr
        For i = 0 To Me.ado_meli.Recordset.RecordCount - 1
            sData = sData & Me.ado_meli.Recordset.Fields(1) & vbTab & Me.ado_meli.Recordset.Fields(2) & vbTab & Me.ado_meli.Recordset.Fields(3) & vbTab & Me.ado_meli.Recordset.Fields(10) & vbTab & Me.ado_meli.Recordset.Fields(11) & vbTab & Me.ado_meli.Recordset.Fields(17) & vbCr
            Me.ado_meli.Recordset.MoveNext
        Next i
   End If
   Clipboard.SetText sData
   oBook.Worksheets(1).Range("A1").Select
   oBook.Worksheets(1).Paste
   oSheet.Range("A1:Z1").Font.Bold = True
   oSheet.Range("A1:Z1").EntireColumn.AutoFit
   oBook.Worksheets(1).Columns.HorizontalAlignment = -4131
   oExcel.Visible = True
   

End Sub

Private Sub dt_meli_HeadClick(ByVal ColIndex As Integer)

    defined_col = ColIndex
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    
    Set rs_ado_dimoi = Nothing
    Set rs_ado_jobs = Nothing
    Set rs_ado_meli = Nothing
    Set Me.dt_meli.DataSource = Nothing
    Set rs_ado_pe = Nothing
    Set Me.co_jobs.RowSource = Nothing
    Set Me.co_dimoi.RowSource = Nothing
    Set Me.co_pe.RowSource = Nothing
    
    
End Sub

Private Sub im_eg_GotFocus()

    With im_eg
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub im_eg_LostFocus()

    With Me.im_eg
        'If IsDate(.Text) = False And for_search = 1 Then
        If IsDate(.Text) = False Then
            Dim imera, year, minas As Variant
            .SelStart = 0
            .SelLength = 2
            imera = .SelText
            If Not (imera >= 1 And imera <= 31) Then
                imera = 0
                .SelText = "  "
            End If
            .SelStart = 3
            .SelLength = 2
            minas = .SelText
            If Not (minas >= 1 And minas <= 12) Then
                minas = 0
                .SelText = "  "
            End If
            .SelStart = 6
            .SelLength = 4
            year = .SelText
            If Not (year >= 0) Then
                year = 0
                .SelText = "  "
            End If
            .SelStart = 0
            .SelLength = 10
            If Val(imera) <> 0 Then
                st = imera & "/"
            Else
                st = "00/"
            End If
            If Val(minas) <> 0 Then
                st = st & minas & "/"
            Else
                st = st & "00/"
            End If
            If Val(year) <> 0 Then
                st = st & year
            Else
                st = st & "0000"
            End If
        End If
        If st = .Text Then
            If IsDate(.Text) = False And (.Text <> "__/__/____") And (.Text <> "  /  /    ") And st <> "00/00/0000" Then
                MsgBox "Λάθος τιμή ημερομηνίας!", vbCritical, "Μήνυμα λάθους"
                .SelStart = 0
                .SelLength = 10
                .SetFocus
            End If
        End If
    End With
    
End Sub

Private Sub MaskEdBox1_GotFocus()

    With MaskEdBox1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub met_st_Click()
If Not Me.ado_meli.Recordset.EOF Then
    melos_id = Me.ado_meli.Recordset.Fields(0).Value
    If melos_id >= 0 And athlet_management.flag_pateras = 1 Then
        athlet_management.ado_pateres.Refresh
        athlet_management.dt_father.Columns(0).Visible = False
        athlet_management.dt_father.Columns(1).Caption = "Α.Μ."
        athlet_management.dt_father.Columns(1).Width = 1000
        athlet_management.dt_father.Columns(2).Caption = "Επώνυμο"
        athlet_management.dt_father.Columns(2).Width = 1500
        athlet_management.dt_father.Columns(3).Caption = "Όνομα"
        athlet_management.dt_father.Columns(3).Width = 1200
        Set athlet_management.rs_ado_pateres = athlet_management.ado_pateres.Recordset
        For i = 4 To athlet_management.rs_ado_pateres.Fields.Count - 1
            athlet_management.dt_father.Columns(i).Visible = False
        Next i
        'athlet_management.rs_ado_pateres.Filter = "[id] LIKE '" & str(melos_id) & "'"
        athlet_management.ado_pateres.Recordset.Filter = "[id] LIKE '" & str(melos_id) & "'"
        athlet_management.flag_pateras = 0
    End If
    If melos_id >= 0 And athlet_management.flag_mitera = 1 Then
        athlet_management.ado_miteres.Refresh
        athlet_management.dt_mother.Columns(0).Visible = False
        athlet_management.dt_mother.Columns(1).Caption = "A.Μ."
        athlet_management.dt_mother.Columns(1).Width = 1000
        athlet_management.dt_mother.Columns(2).Caption = "Επώνυμο"
        athlet_management.dt_mother.Columns(2).Width = 1500
        athlet_management.dt_mother.Columns(3).Caption = "Όνομα"
        athlet_management.dt_mother.Columns(3).Width = 1200
        Set athlet_management.rs_ado_miteres = athlet_management.ado_miteres.Recordset
        For i = 4 To athlet_management.rs_ado_miteres.Fields.Count - 1
            athlet_management.dt_mother.Columns(i).Visible = False
        Next i
        'athlet_management.rs_ado_miteres.Filter = "[id] LIKE '" & str(melos_id) & "'"
        athlet_management.ado_miteres.Recordset.Filter = "[id] LIKE '" & str(melos_id) & "'"
        athlet_management.flag_mitera = 0
    End If
    
    ' ΜΕΤΑΦΟΡΑ ΥΠΟΨΗΦΙΩΝ ΙΔΙΩΝ ΣΤΟΙΧΕΙΩΝ ΑΠΟ ΓΟΝΕΑ ΣΕ ΠΑΙΔΙ, εφόσον δε γίνεται ΑΝΑΖΗΤΗΣΗ
If athlet_management.for_search = 0 Then
    If Trim(Me.ado_meli.Recordset.Fields(4).Value) <> "" And Trim(athlet_management.tmp_odos.Text) = "" Then
        athlet_management.tmp_odos.Text = Me.ado_meli.Recordset.Fields(4).Value
    End If
    If Trim(Me.ado_meli.Recordset.Fields(5).Value) <> "" And Trim(athlet_management.tmp_arithmos.Text) = "" Then
        athlet_management.tmp_arithmos.Text = Me.ado_meli.Recordset.Fields(5).Value
    End If
    If Trim(Me.ado_meli.Recordset.Fields(6).Value) <> "" And Trim(athlet_management.tmp_perioxi.Text) = "" Then
        athlet_management.tmp_perioxi.Text = Me.ado_meli.Recordset.Fields(6).Value
    End If
    If Trim(Me.ado_meli.Recordset.Fields(8).Value) <> "" And Trim(athlet_management.co_pe.Text) = "" Then
        athlet_management.co_pe.Text = Me.ado_meli.Recordset.Fields(8).Value
    End If
    If Trim(Me.ado_meli.Recordset.Fields(7).Value) <> "" And Trim(athlet_management.co_dimoi.Text) = "" Then
        athlet_management.co_dimoi.Text = Me.ado_meli.Recordset.Fields(7).Value
    End If
    If Trim(Me.ado_meli.Recordset.Fields(9).Value) <> "" And Trim(athlet_management.tmp_tk.Text) = "" Then
        athlet_management.tmp_tk.Text = Me.ado_meli.Recordset.Fields(9).Value
    End If
    If Trim(Me.ado_meli.Recordset.Fields(10).Value) <> "" And Trim(athlet_management.tmp_til_oikias.Text) = "" Then
        athlet_management.tmp_til_oikias.Text = Me.ado_meli.Recordset.Fields(10).Value
    End If
    If Trim(Me.ado_meli.Recordset.Fields(11).Value) <> "" And Trim(athlet_management.tmp_kinito.Text) = "" Then
        athlet_management.tmp_kinito.Text = Me.ado_meli.Recordset.Fields(11).Value
    End If
    If Trim(Me.ado_meli.Recordset.Fields(12).Value) <> "" And Trim(athlet_management.tmp_fax.Text) = "" Then
        athlet_management.tmp_fax.Text = Me.ado_meli.Recordset.Fields(12).Value
    End If
    If Trim(Me.ado_meli.Recordset.Fields(13).Value) <> "" And Trim(athlet_management.tmp_email.Text) = "" Then
        athlet_management.tmp_email.Text = Me.ado_meli.Recordset.Fields(13).Value
    End If
End If
    '
    'meli_management.Hide
    Unload meli_management
End If
End Sub

Private Sub del_bt_Click()
    
    Dim ms As String
    
    If Not Me.ado_meli.Recordset.EOF And Me.ado_meli.Recordset.AbsolutePosition >= 1 Then
    ms = MsgBox("Είσαι σίγουρος; (ΝΑΙ ή ΟΧΙ)", vbYesNo, "Παράθυρο διαγραφής")
    If ms = 6 Then
        Me.ado_meli.Recordset.Delete
        '''''''''
        'Έχουν μείνει μέλη μετά τη ΔΙΑΓΡΑΦΗ
        If Me.ado_meli.Recordset.RecordCount >= 1 Then
        'ΔΕΝ έχουν μείνει αθλητές μετά τη ΔΙΑΓΡΑΦΗ, άρα REQUERY
        Else
            Me.ado_meli.Refresh
            Me.ado_meli.Refresh
            If Me.ado_meli.Recordset.RecordCount >= 1 Then
                Me.ado_meli.Recordset.MoveFirst
                Me.dt_meli.Columns(0).Caption = "Κωδικός"
                Me.dt_meli.Columns(0).Width = 1500
                Me.dt_meli.Columns(1).Visible = False
                Me.dt_meli.Columns(2).Caption = "Επώνυμο"
                Me.dt_meli.Columns(2).Width = 3000
                Me.dt_meli.Columns(3).Caption = "Όνομα"
                Me.dt_meli.Columns(3).Width = 2000
                For i = 4 To 16
                    Me.dt_meli.Columns(i).Visible = False
                Next i
                Me.dt_meli.Columns(17).Caption = "Επάγγελμα"
                Me.dt_meli.Columns(17).Width = 2000
                For i = 18 To Me.ado_meli.Recordset.Fields.Count - 1
                    Me.dt_meli.Columns(i).Visible = False
                Next i
                Me.canc_bt.Enabled = False
                Me.sear_bt.Enabled = False
                Me.insert_bt.Enabled = True
                'ΔΕΝ έχουν μείνει αθλητές μετά τη ΔΙΑΓΡΑΦΗ, άρα καθαρισμός όλων των πεδίων της φόρμας
                Else
                    '
                    Me.save_command.Enabled = False
                    Me.sear_bt.Enabled = False
                    Me.insert_bt.Enabled = True
                    Me.canc_bt.Enabled = False
                    Me.up_bt.Enabled = False
                    Me.Command1.Enabled = False
                    Me.del_bt.Enabled = False
                End If
            End If
        '''''''''
        End If
    Else
        MsgBox "Δεν υπάρχει εγγραφή προς ΔΙΑΓΡΑΦΗ!", vbCritical, "Μήνυμα Λάθους"
    End If
    
End Sub

Private Sub dt_meli_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If Me.ado_meli.Recordset.AbsolutePosition >= 1 And Me.ado_meli.Recordset.AbsolutePosition <= Me.ado_meli.Recordset.RecordCount Then
        If Trim(Me.ado_meli.Recordset.Fields(0).Value) <> "" And Me.it_is_a_new_record = 0 Then
            txt_kod.Text = Me.ado_meli.Recordset.Fields(0).Value
        Else
            If Me.it_is_a_new_record = 1 Then
                '
            Else
                txt_kod.Text = ""
            End If
        End If
        If Trim(Me.ado_meli.Recordset.Fields(1).Value) <> "" Then
            tmp_am.Text = Me.ado_meli.Recordset.Fields(1).Value
        Else
            tmp_am.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(3).Value) <> "" And Me.it_is_a_new_record = 0 Then
            tmp_onoma.Text = Me.ado_meli.Recordset.Fields(3).Value
        Else
            If Me.it_is_a_new_record = 1 Then
                '
            Else
                tmp_onoma.Text = ""
            End If
        End If
        If Trim(Me.ado_meli.Recordset.Fields(2).Value) <> "" Then
            tmp_eponimo.Text = Me.ado_meli.Recordset.Fields(2).Value
        Else
            tmp_eponimo.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(4).Value) <> "" Then
            tmp_odos.Text = Me.ado_meli.Recordset.Fields(4).Value
        Else
            tmp_odos.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(5).Value) <> "" Then
            tmp_arithmos.Text = Me.ado_meli.Recordset.Fields(5).Value
        Else
            tmp_arithmos.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(6).Value) <> "" Then
            tmp_perioxi.Text = Me.ado_meli.Recordset.Fields(6).Value
        Else
            tmp_perioxi.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(8).Value) <> "" Then
            Me.co_pe.Text = Me.ado_meli.Recordset.Fields(8).Value
        Else
            Me.co_pe.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(7).Value) <> "" Then
            Me.co_dimoi.Text = Me.ado_meli.Recordset.Fields(7).Value
        Else
            Me.co_dimoi.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(9).Value) <> "" Then
            tmp_tk.Text = Me.ado_meli.Recordset.Fields(9).Value
        Else
            tmp_tk.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(10).Value) <> "" Then
            tmp_til_oikias.Text = Me.ado_meli.Recordset.Fields(10).Value
        Else
            tmp_til_oikias.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(11).Value) <> "" Then
            tmp_kinito.Text = Me.ado_meli.Recordset.Fields(11).Value
        Else
            tmp_kinito.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(12).Value) <> "" Then
            tmp_fax.Text = Me.ado_meli.Recordset.Fields(12).Value
        Else
            tmp_fax.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(13).Value) <> "" Then
            tmp_email.Text = Me.ado_meli.Recordset.Fields(13).Value
        Else
            tmp_email.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(14).Value) <> "" Then
            Me.MaskEdBox1.Text = Me.ado_meli.Recordset.Fields(14).Value
        Else
            Me.MaskEdBox1.Text = "00/00/0000"
        End If
        If Me.ado_meli.Recordset.Fields(15).Value = True Then
            Me.ch_en.Value = 1
        Else
            Me.ch_en.Value = 0
        End If
        If Trim(Me.ado_meli.Recordset.Fields(16).Value) <> "" Then
            Me.im_eg.Text = Me.ado_meli.Recordset.Fields(16).Value
        Else
            Me.im_eg.Text = "00/00/0000"
        End If
        If Trim(Me.ado_meli.Recordset.Fields(17).Value) <> "" Then
            Me.co_jobs.Text = Me.ado_meli.Recordset.Fields(17).Value
        Else
            Me.co_jobs.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(18).Value) <> "" Then
            Me.txt_adt.Text = Me.ado_meli.Recordset.Fields(18).Value
        Else
            Me.txt_adt.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(19).Value) <> "" Then
            Me.txt_ea_adt.Text = Me.ado_meli.Recordset.Fields(19).Value
        Else
            Me.txt_ea_adt.Text = ""
        End If
        If Me.ado_meli.Recordset.Fields(20).Value = True Then
            Me.ch_m_gs.Value = 1
        Else
            Me.ch_m_gs.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields(21).Value = True Then
            Me.ch_m_ds.Value = 1
        Else
            Me.ch_m_ds.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields(22).Value = True Then
            Me.ch_p.Value = 1
        Else
            Me.ch_p.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields(23).Value = True Then
            Me.ch_e.Value = 1
        Else
            Me.ch_e.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields(25).Value = True Then
            Me.ch_g.Value = 1
        Else
            Me.ch_g.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields("Δρομέας").Value = True Then
            Me.ch_dr.Value = 1
        Else
            Me.ch_dr.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields("Τμήμα_Ενηλίκων").Value = True Then
            Me.ch_tm_en.Value = 1
        Else
            Me.ch_tm_en.Value = 0
        End If
    End If

End Sub

Private Sub Form_Load()
 
 
    for_search = 0
    s_sort = ""
    MDIForm1.s_string = ""
    MDIForm1.s_sort = ""
    MDIForm1.rep_lbl = ""
    Me.it_is_a_new_record = 0
    Me.ado_dimoi.Refresh
    Set rs_ado_dimoi = ado_dimoi.Recordset
    Set Me.co_dimoi.RowSource = Me.ado_dimoi
    Me.ado_jobs.Refresh
    Set rs_ado_jobs = ado_jobs.Recordset
    Set Me.co_jobs.RowSource = Me.ado_jobs
    Me.ado_pe.Refresh
    Set rs_ado_pe = ado_pe.Recordset
    Set Me.co_pe.RowSource = Me.ado_pe
         
    Me.Height = 11500
    Me.Width = 11930
 
    Me.ado_meli.Recordset.Sort = "[" & Trim(Me.ado_meli.Recordset.Fields(0).Name) & "]"
    If Not Me.ado_meli.Recordset.EOF Then
        Me.ado_meli.Recordset.MoveFirst
        txt_kod.Text = Me.ado_meli.Recordset.Fields(0).Value
        If Trim(Me.ado_meli.Recordset.Fields(1).Value) <> "" Then
            tmp_am.Text = Me.ado_meli.Recordset.Fields(1).Value
        Else
            tmp_am.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(3).Value) <> "" Then
            tmp_onoma.Text = Me.ado_meli.Recordset.Fields(3).Value
        Else
            tmp_onoma.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(2).Value) <> "" Then
            tmp_eponimo.Text = Me.ado_meli.Recordset.Fields(2).Value
        Else
            tmp_eponimo.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(4).Value) <> "" Then
            tmp_odos.Text = Me.ado_meli.Recordset.Fields(4).Value
        Else
            tmp_odos.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(5).Value) <> "" Then
            tmp_arithmos.Text = Me.ado_meli.Recordset.Fields(5).Value
        Else
            tmp_arithmos.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(6).Value) <> "" Then
            tmp_perioxi.Text = Me.ado_meli.Recordset.Fields(6).Value
        Else
            tmp_perioxi.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(8).Value) <> "" Then
            Me.co_pe.Text = Me.ado_meli.Recordset.Fields(8).Value
        Else
            Me.co_pe.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(7).Value) <> "" Then
            Me.co_dimoi.Text = Me.ado_meli.Recordset.Fields(7).Value
        Else
            Me.co_dimoi.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(9).Value) <> "" Then
            tmp_tk.Text = Me.ado_meli.Recordset.Fields(9).Value
        Else
            tmp_tk.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(10).Value) <> "" Then
            tmp_til_oikias.Text = Me.ado_meli.Recordset.Fields(10).Value
        Else
            tmp_til_oikias.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(11).Value) <> "" Then
            tmp_kinito.Text = Me.ado_meli.Recordset.Fields(11).Value
        Else
            tmp_kinito.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(12).Value) <> "" Then
            tmp_fax.Text = Me.ado_meli.Recordset.Fields(12).Value
        Else
            tmp_fax.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(13).Value) <> "" Then
            tmp_email.Text = Me.ado_meli.Recordset.Fields(13).Value
        Else
            tmp_email.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(14).Value) <> "" Then
            Me.MaskEdBox1.Text = Me.ado_meli.Recordset.Fields(14).Value
        Else
            'Me.MaskEdBox1.Text = "00/00/0000"
            Me.MaskEdBox1.Text = "  /  /    "
        End If
        If Me.ado_meli.Recordset.Fields(15).Value = True Then
            Me.ch_en.Value = 1
        Else
            Me.ch_en.Value = 0
        End If
        If Trim(Me.ado_meli.Recordset.Fields(16).Value) <> "" Then
            Me.im_eg.Text = Me.ado_meli.Recordset.Fields(16).Value
        Else
            'Me.im_eg.Text = "00/00/0000"
            Me.im_eg.Text = "  /  /    "
        End If
        If Trim(Me.ado_meli.Recordset.Fields(17).Value) <> "" Then
            Me.co_jobs.Text = Me.ado_meli.Recordset.Fields(17).Value
        Else
            Me.co_jobs.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(18).Value) <> "" Then
            Me.txt_adt.Text = Me.ado_meli.Recordset.Fields(18).Value
        Else
            Me.txt_adt.Text = ""
        End If
        If Trim(Me.ado_meli.Recordset.Fields(19).Value) <> "" Then
            Me.txt_ea_adt.Text = Me.ado_meli.Recordset.Fields(19).Value
        Else
            Me.txt_ea_adt.Text = ""
        End If
        If Me.ado_meli.Recordset.Fields(20).Value = True Then
            Me.ch_m_gs.Value = 1
        Else
            Me.ch_m_gs.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields(21).Value = True Then
            Me.ch_m_ds.Value = 1
        Else
            Me.ch_m_ds.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields(22).Value = True Then
            Me.ch_p.Value = 1
        Else
            Me.ch_p.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields(23).Value = True Then
            Me.ch_e.Value = 1
        Else
            Me.ch_e.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields(25).Value = True Then
            Me.ch_g.Value = 1
        Else
            Me.ch_g.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields("Δρομέας").Value = True Then
            Me.ch_dr.Value = 1
        Else
            Me.ch_dr.Value = 0
        End If
        If Me.ado_meli.Recordset.Fields("Τμήμα_Ενηλίκων").Value = True Then
            Me.ch_tm_en.Value = 1
        Else
            Me.ch_tm_en.Value = 0
        End If
    End If
    If Me.ado_meli.Recordset.RecordCount > 0 Then
        Me.dt_meli.Row = 0
        Me.dt_meli.Col = 1
    End If
    If Me.ado_meli.Recordset.RecordCount > 0 Then
        Me.ado_meli.Caption = "Μέλος " & Me.dt_meli.Row + 1 & " από " & Me.ado_meli.Recordset.RecordCount
    Else
        Me.ado_meli.Caption = "Μέλος " & 0 & " από " & 0
    End If
    Me.dt_meli.Columns(0).Caption = "Κωδικός"
    Me.dt_meli.Columns(0).Width = 1500
    Me.dt_meli.Columns(1).Visible = False
    Me.dt_meli.Columns(2).Caption = "Επώνυμο"
    Me.dt_meli.Columns(2).Width = 3000
    Me.dt_meli.Columns(3).Caption = "Όνομα"
    Me.dt_meli.Columns(3).Width = 2000
    For i = 4 To 16
        Me.dt_meli.Columns(i).Visible = False
    Next i
    Me.dt_meli.Columns(17).Caption = "Επάγγελμα"
    Me.dt_meli.Columns(17).Width = 2000
    For i = 18 To Me.ado_meli.Recordset.Fields.Count - 1
        Me.dt_meli.Columns(i).Visible = False
    Next i
    
End Sub

Private Sub insert_bt_Click()


    'Me.txt_kod.Locked = True

    Me.it_is_a_new_record = 1
    Me.Command2.Enabled = True
    'Να βρω το υποψήφιο id_μέλους
    If Me.tmp_ado_meli.Recordset.RecordCount >= 1 Then
        Me.tmp_ado_meli.Recordset.Sort = "[" & Trim(Me.tmp_ado_meli.Recordset.Fields(0).Name) & "]"
        Me.tmp_ado_meli.Recordset.MoveLast
        id_α = Me.tmp_ado_meli.Recordset.Fields(0).Value
        id_α = id_α + 1
    Else
        id_α = 1
    End If
    'Καθαρισμός πεδίων
    txt_kod.Text = id_α
    tmp_am.Text = ""
    tmp_onoma.Text = ""
    tmp_eponimo.Text = ""
    Me.txt_adt.Text = ""
    Me.txt_ea_adt.Text = ""
    tmp_odos.Text = ""
    tmp_arithmos.Text = ""
    tmp_perioxi.Text = ""
    Me.co_pe.Text = "Ιωαννίνων"
    Me.co_dimoi.Text = "Ιωαννιτών"
    tmp_tk.Text = ""
    tmp_til_oikias.Text = ""
    tmp_kinito.Text = ""
    tmp_fax.Text = ""
    tmp_email.Text = ""
    Me.MaskEdBox1.Text = "  /  /    "
    Me.ch_en.Value = 1  'Με την προσθήκη μέλους, αυτό θεωρείται ΕΝΕΡΓΟ
    Me.im_eg.Text = "  /  /    "
    Me.co_jobs.Text = ""
    Me.ch_m_ds.Value = 0
    Me.ch_m_gs.Value = 0
    Me.ch_e.Value = 0
    Me.ch_g.Value = 1
    Me.ch_p.Value = 0
    Me.ch_dr.Value = 0
    
    ' ΜΕΤΑΦΟΡΑ ΥΠΟΨΗΦΙΩΝ ΙΔΙΩΝ ΣΤΟΙΧΕΙΩΝ ΑΠΟ ΠΑΙΔΙ ΣΕ ΓΟΝΕΑ
    If Me.met_st.Enabled = True And (athlet_management.flag_mitera = 1 Or athlet_management.flag_pateras = 1) Then
        If Trim(athlet_management.tmp_odos.Text) <> "" Then
            Me.tmp_odos.Text = athlet_management.tmp_odos.Text
        End If
        If Trim(athlet_management.tmp_arithmos.Text) <> "" Then
            Me.tmp_arithmos.Text = athlet_management.tmp_arithmos.Text
        End If
        If Trim(athlet_management.tmp_perioxi.Text) <> "" Then
            Me.tmp_perioxi.Text = athlet_management.tmp_perioxi.Text
        End If
        If Trim(athlet_management.co_pe.Text) <> "" Then
            Me.co_pe.Text = athlet_management.co_pe.Text
        End If
        If Trim(athlet_management.co_dimoi.Text) <> "" Then
            Me.co_dimoi.Text = athlet_management.co_dimoi.Text
        End If
        If Trim(athlet_management.tmp_tk.Text) <> "" Then
            Me.tmp_tk.Text = athlet_management.tmp_tk.Text
        End If
        If Trim(athlet_management.tmp_til_oikias.Text) <> "" Then
            Me.tmp_til_oikias.Text = athlet_management.tmp_til_oikias.Text
        End If
        If Trim(athlet_management.tmp_kinito.Text) <> "" Then
            Me.tmp_kinito.Text = athlet_management.tmp_kinito.Text
        End If
        If Trim(athlet_management.tmp_fax.Text) <> "" Then
            Me.tmp_fax.Text = athlet_management.tmp_fax.Text
        End If
        If Trim(athlet_management.tmp_email.Text) <> "" Then
            Me.tmp_email.Text = athlet_management.tmp_email.Text
        End If
    End If
    '

    Me.tmp_onoma.SetFocus
    
    Me.insert_bt.Enabled = False
    Me.save_command.Enabled = True
    Me.canc_bt.Enabled = True
    Me.up_bt.Enabled = False
    Me.met_st.Enabled = False
    Me.del_bt.Enabled = False
    Me.Command1.Enabled = False
    Me.cancel_cur_rec.Enabled = False
    
End Sub

Private Sub kl_bt_Click()

    Me.it_is_a_new_record = 0
    Unload Me

End Sub

Private Sub MaskEdBox1_LostFocus()
        
        
    With Me.MaskEdBox1
        'If IsDate(.Text) = False And for_search = 1 Then
        If IsDate(.Text) = False Then
            Dim imera, year, minas As Variant
            .SelStart = 0
            .SelLength = 2
            imera = .SelText
            If Not (imera >= 1 And imera <= 31) Then
                imera = 0
                .SelText = "  "
            End If
            .SelStart = 3
            .SelLength = 2
            minas = .SelText
            If Not (minas >= 1 And minas <= 12) Then
                minas = 0
                .SelText = "  "
            End If
            .SelStart = 6
            .SelLength = 4
            year = .SelText
            If Not (year >= 0) Then
                year = 0
                .SelText = "  "
            End If
            .SelStart = 0
            .SelLength = 10
            If Val(imera) <> 0 Then
                st = imera & "/"
            Else
                st = "00/"
            End If
            If Val(minas) <> 0 Then
                st = st & minas & "/"
            Else
                st = st & "00/"
            End If
            If Val(year) <> 0 Then
                st = st & year
            Else
                st = st & "0000"
            End If
        End If
        If st = .Text Then
            If IsDate(.Text) = False And (.Text <> "__/__/____") And (.Text <> "  /  /    ") And st <> "00/00/0000" Then
                MsgBox "Λάθος τιμή ημερομηνίας!", vbCritical, "Μήνυμα λάθους"
                .SelStart = 0
                .SelLength = 10
                .SelText = "  /  /    "
                .SetFocus
            End If
        End If
    End With
    
End Sub

Private Sub save_command_Click()

    Me.it_is_a_new_record = 0
    Me.Command2.Enabled = False
    'Αποθήκευση στα ΜΕΛΗ
    Me.tmp_ado_meli.Recordset.AddNew
    Me.tmp_ado_meli.Recordset.Fields(0).Value = id_α
    'Αποθήκευση ΟΝΟΜΑΤΟΣ
    If Trim(Me.tmp_onoma.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields("Όνομα").Value = Me.tmp_onoma.Text
    End If
    'Αποθήκευση ΕΠΩΝΥΜΟΥ
    If Trim(Me.tmp_eponimo.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields("Επώνυμο").Value = Me.tmp_eponimo.Text
    End If
    'Αποθήκευση ΑΡΙΘΜΟΥ ΔΕΛΤΙΟΥ ΤΑΥΤΟΤΗΤΑΣ
    If Trim(Me.txt_adt.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(18).Value = Me.txt_adt.Text
    End If
    'Αποθήκευση ΕΚΔΟΥΣΑΣ ΑΡΧΗΣ ΑΡΙΘΜΟΥ ΔΕΛΤΙΟΥ ΤΑΥΤΟΤΗΤΑΣ
    If Trim(Me.txt_ea_adt.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(19).Value = Me.txt_ea_adt.Text
    End If
    'Αποθήκευση ΟΔΟΥ
    If Trim(Me.tmp_odos.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(4).Value = Me.tmp_odos.Text
    End If
    'Αποθήκευση ΑΡΙΘΜΟΥ
    If Trim(Me.tmp_arithmos.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(5).Value = Me.tmp_arithmos.Text
    End If
    'Αποθήκευση ΠΕΡΙΟΧΗΣ
    If Trim(Me.tmp_perioxi.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(6).Value = Me.tmp_perioxi.Text
    End If
    'Αποθήκευση ΔΗΜΟΥ
    If Trim(Me.co_dimoi.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(7).Value = Me.co_dimoi.Text
        Me.ado_dimoi.Recordset.Requery
        Me.ado_dimoi.Refresh
        Me.co_dimoi.ReFill
        Me.co_dimoi.Text = Trim(Me.tmp_ado_meli.Recordset.Fields(7).Value)
        If Me.co_dimoi.MatchedWithList = False Then
        'ΕΙΣΑΓΩΓΗ ΝΕΟΥ ΔΗΜΟΥ
            If Me.ado_dimoi.Recordset.RecordCount >= 1 Then
                Me.ado_dimoi.Recordset.Sort = "[" & Trim(Me.ado_dimoi.Recordset.Fields(0).Name) & "]"
                Me.ado_dimoi.Recordset.MoveLast
                id = Me.ado_dimoi.Recordset![id_δήμου]
            Else
                id = 0
            End If
            Me.ado_dimoi.Recordset.AddNew
            Me.ado_dimoi.Recordset.Fields(0) = id + 1
            Me.ado_dimoi.Recordset.Fields(1) = Trim(Me.co_dimoi.Text)
            Me.ado_dimoi.Recordset.UpdateBatch adAffectCurrent
            Me.ado_dimoi.Recordset.Requery
        End If
    End If
    If Trim(Me.co_pe.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(8).Value = Me.co_pe.Text
        Me.ado_pe.Recordset.Requery
        Me.ado_pe.Refresh
        Me.co_pe.ReFill
        Me.co_pe.Text = Trim(Me.tmp_ado_meli.Recordset.Fields(8).Value)
        If Me.co_pe.MatchedWithList = False Then
        'ΕΙΣΑΓΩΓΗ ΝΕΑΣ ΠΕΡΙΦΕΡΕΙΑΚΗΣ ΕΝΟΤΗΤΑΣ
            If Me.ado_pe.Recordset.RecordCount >= 1 Then
                Me.ado_pe.Recordset.Sort = "[" & Trim(Me.ado_pe.Recordset.Fields(0).Name) & "]"
                Me.ado_pe.Recordset.MoveLast
                id = Me.ado_pe.Recordset![id_πε]
            Else
                id = 0
            End If
            Me.ado_pe.Recordset.AddNew
            Me.ado_pe.Recordset.Fields(0) = id + 1
            Me.ado_pe.Recordset.Fields(1) = Trim(Me.co_pe.Text)
            Me.ado_pe.Recordset.UpdateBatch adAffectCurrent
            Me.ado_pe.Recordset.Requery
        End If
    End If
    If Trim(Me.tmp_tk.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(9).Value = Me.tmp_tk.Text
    End If
    If Trim(Me.tmp_til_oikias.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(10).Value = Me.tmp_til_oikias.Text
    End If
    If Trim(Me.tmp_kinito.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(11).Value = Me.tmp_kinito.Text
    End If
    If Trim(Me.tmp_fax.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(12).Value = Me.tmp_fax.Text
    End If
    If Trim(Me.tmp_email.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(13).Value = Me.tmp_email.Text
    End If
    If Trim(Me.MaskEdBox1.Text) <> "00/00/0000" Then
        Me.tmp_ado_meli.Recordset.Fields(14).Value = Me.MaskEdBox1.Text
    End If
    If Trim(Me.tmp_am.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(1).Value = Me.tmp_am.Text
    End If
    Me.tmp_ado_meli.Recordset.Fields(15).Value = Me.ch_en.Value
    If Trim(Me.im_eg.Text) <> "00/00/0000" Then
        Me.tmp_ado_meli.Recordset.Fields(16).Value = Me.im_eg.Text
    End If
    If Trim(Me.co_jobs.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(17).Value = Me.co_jobs.Text
        If Me.co_jobs.MatchedWithList = False Then
        'ΕΙΣΑΓΩΓΗ ΝΕΟΥ ΕΠΑΓΓΕΛΜΑΤΟΣ
            If Me.ado_jobs.Recordset.RecordCount >= 1 Then
                Me.ado_jobs.Recordset.Sort = "[" & Trim(Me.ado_jobs.Recordset.Fields(0).Name) & "]"
                Me.ado_jobs.Recordset.MoveLast
                id = Me.ado_jobs.Recordset![id_επαγγέλματος]
            Else
                id = 0
            End If
            Me.ado_jobs.Recordset.AddNew
            Me.ado_jobs.Recordset.Fields(0) = id + 1
            Me.ado_jobs.Recordset.Fields(1) = Trim(Me.co_jobs.Text)
            Me.ado_jobs.Recordset.UpdateBatch adAffectCurrent
            Me.ado_jobs.Recordset.Requery
        End If
    End If
    Me.tmp_ado_meli.Recordset.Fields(20).Value = Me.ch_m_gs.Value
    Me.tmp_ado_meli.Recordset.Fields(21).Value = Me.ch_m_ds.Value
    Me.tmp_ado_meli.Recordset.Fields(22).Value = Me.ch_p.Value
    Me.tmp_ado_meli.Recordset.Fields(23).Value = Me.ch_e.Value
    Me.tmp_ado_meli.Recordset.Fields(25).Value = Me.ch_g.Value
    Me.tmp_ado_meli.Recordset.Fields("Δρομέας").Value = Me.ch_dr.Value
    Me.tmp_ado_meli.Recordset.Fields("Τμήμα_Ενηλίκων").Value = Me.ch_tm_en.Value
    '
    Me.tmp_ado_meli.Recordset.UpdateBatch adAffectCurrent
    Me.tmp_ado_meli.Recordset.Requery
    Me.tmp_ado_meli.Refresh
    Me.ado_meli.Recordset.Requery
    Me.ado_meli.Refresh
    'Me.dt_meli.Refresh
    'Me.ado_meli.Recordset.Sort = "[" & Trim(Me.ado_meli.Recordset.Fields(0).Name) & "]"
    'Me.ado_meli.Recordset.MoveLast
    
    Me.ado_meli.Recordset.Filter = MDIForm1.s_string
    If MDIForm1.s_sort <> "" Then
        Me.ado_meli.Recordset.Sort = MDIForm1.s_sort
    Else
        Me.ado_meli.Recordset.Sort = "[id]"
    End If
    If Me.ado_meli.Recordset.RecordCount >= 1 Then
        Me.ado_meli.Recordset.Find "[id] = " & id_α
        If Not Me.ado_meli.Recordset.EOF Then
            mv = Me.ado_meli.Recordset.AbsolutePosition
            Me.ado_meli.Recordset.MoveFirst
            Me.ado_meli.Recordset.Move mv - 1
        Else
            Me.ado_meli.Recordset.MoveFirst
        End If
    Else
        txt_kod.Text = ""
        Me.ch_en = 0
        tmp_onoma.Text = ""
        tmp_eponimo.Text = ""
        txt_adt = ""
        txt_ea_adt = ""
        Me.MaskEdBox1.Text = "  /  /    "
        Me.co_jobs.Text = ""
        tmp_til_oikias.Text = ""
        tmp_kinito.Text = ""
        tmp_fax.Text = ""
        tmp_email.Text = ""
        Me.ch_m_gs = 0
        Me.tmp_am.Text = ""
        Me.im_eg.Text = "  /  /    "
        Me.Text4.Text = ""
        Me.ch_m_ds = 0
        Me.ch_p = 0
        Me.ch_e = 0
        Me.ch_g = 0
        Me.ch_dr = 0
        Me.tmp_odos.Text = ""
        Me.tmp_arithmos.Text = ""
        Me.tmp_tk.Text = ""
        tmp_perioxi.Text = ""
        Me.co_dimoi.Text = ""
        Me.co_pe.Text = ""
    End If
    
    Me.dt_meli.Columns(0).Caption = "Κωδικός"
    Me.dt_meli.Columns(0).Width = 1500
    Me.dt_meli.Columns(1).Visible = False
    Me.dt_meli.Columns(2).Caption = "Επώνυμο"
    Me.dt_meli.Columns(2).Width = 3000
    Me.dt_meli.Columns(3).Caption = "Όνομα"
    Me.dt_meli.Columns(3).Width = 2000
    For i = 4 To 16
        Me.dt_meli.Columns(i).Visible = False
    Next i
    Me.dt_meli.Columns(17).Caption = "Επάγγελμα"
    Me.dt_meli.Columns(17).Width = 2000
    For i = 18 To Me.ado_meli.Recordset.Fields.Count - 1
        Me.dt_meli.Columns(i).Visible = False
    Next i
    
    Me.save_command.Enabled = False
    Me.canc_bt.Enabled = True
    Me.insert_bt.Enabled = True
    Me.up_bt.Enabled = True
    Me.cancel_cur_rec.Enabled = True
    Me.del_bt.Enabled = True
    Me.Command1.Enabled = True
    If athlet_management.flag_mitera = 1 Or athlet_management.flag_pateras = 1 Then
        Me.met_st.Enabled = True
    End If
    
    'Err.Clear
    
End Sub

Private Sub sear_bt_Click()


    Me.txt_kod.Locked = True

    s_string = ""
    rep_lbl = ""
    'ΚΡΙΤΗΡΙΟ ΚΩΔΙΚΟΥ
    If Trim(Me.txt_kod.Text) <> "" Then
        s_string = "[id] LIKE " & Trim(Me.txt_kod.Text)
    End If
    'ΚΡΙΤΗΡΙΟ ΕΝΕΡΓΟΥ ΜΕΛΟΥΣ
    If s_string <> "" Then
            If ch_en.Value = 0 Then
                's_string = s_string & " AND [Ενεργό] LIKE FALSE"
            Else
                s_string = s_string & " AND [Ενεργό] LIKE TRUE"
                rep_lbl = rep_lbl & " (ΕΝΕΡΓΑ)"
            End If
    Else
            If ch_en.Value = 0 Then
                's_string = s_string & "[Ενεργό] LIKE FALSE"
            Else
                s_string = s_string & "[Ενεργό] LIKE TRUE"
                rep_lbl = rep_lbl & " (ΕΝΕΡΓΑ)"
            End If
    End If
    'ΚΡΙΤΗΡΙΟ ΟΝΟΜΑΤΟΣ
    If Trim(tmp_onoma.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Όνομα] LIKE '*" & Trim(tmp_onoma.Text) & "*'"
        Else
            s_string = "[Όνομα] LIKE '*" & Trim(tmp_onoma.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΕΠΩΝΥΜΟΥ
    If Trim(tmp_eponimo.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Επώνυμο] LIKE '*" & Trim(tmp_eponimo.Text) & "*'"
        Else
            s_string = "[Επώνυμο] LIKE '*" & Trim(tmp_eponimo.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΑΔΤ
    If Trim(txt_adt.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΑΔΤ] LIKE '*" & Trim(txt_adt.Text) & "*'"
        Else
            s_string = "[ΑΔΤ] LIKE '*" & Trim(txt_adt.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΕΚΔΟΥΣΑ ΑΡΧΗ ΑΔΤ
    If Trim(txt_ea_adt.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΕκδΑρχήΑΔΤ] LIKE '*" & Trim(txt_ea_adt.Text) & "*'"
        Else
            s_string = "[ΕκδΑρχήΑΔΤ] LIKE '*" & Trim(txt_ea_adt.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΗΜΕΡΟΜΗΝΙΑ ΓΕΝΝΗΣΗΣ
    If Trim(MaskEdBox1.Text) <> "00/00/0000" Then
        Dim imera, year, minas As String
        Me.MaskEdBox1.SelStart = 0
        Me.MaskEdBox1.SelLength = 2
        imera = Me.MaskEdBox1.SelText
        Me.MaskEdBox1.SelStart = 3
        Me.MaskEdBox1.SelLength = 2
        minas = Me.MaskEdBox1.SelText
        Me.MaskEdBox1.SelStart = 6
        Me.MaskEdBox1.SelLength = 4
        year = Me.MaskEdBox1.SelText
        Me.MaskEdBox1.SelStart = 0
        Me.MaskEdBox1.SelLength = 10
        st1 = ""
        If Val(imera) <> 0 Then
            st1 = imera & "/"
            st5 = imera & "/"
        Else
            st1 = ""
            st5 = ""
        End If
        If Val(minas) <> 0 Then
            If st1 = "" Then
                st2 = "*/" & minas & "/"
                st6 = "%/" & minas & "/"
            Else
                st2 = minas & "/"
                st6 = minas & "/"
            End If
        Else
            If st1 = "" Then
                st2 = ""
                st6 = ""
            Else
                st2 = "*/"
                st6 = "%/"
            End If
        End If
        If Val(year) <> 0 Then
            If st1 = "" And st2 = "" Then
                st3 = "*/" & year & "*"
                st7 = "%/" & year & "%"
            Else
                st3 = "" & year & "*"
                st7 = "" & year & "%"
            End If
        Else
            st3 = "*"
            st7 = "%"
        End If
        If Val(imera) <> 0 Then
            st1 = imera & "/*"
        Else
            st1 = ""
        End If
        If Val(minas) <> 0 Then
            st2 = "*/" & minas & "/*"
        Else
            st2 = ""
        End If
        If Val(year) <> 0 Then
            st3 = "*/" & year & "*"
        Else
            st3 = ""
        End If
        st = Trim(st1) & Trim(st2) & Trim(st3)
        st4 = st5 + st6 + st7
        ''
        If s_string <> "" Then
            's_string = s_string & " AND [Γέννηση] LIKE '" & Trim(st) & "'"
            's_string = s_string & " AND [Γέννηση] LIKE '" & st1 & "' and [Γέννηση] like '" & st2 & "' and [Γέννηση] like '" & st3 & "'"
            If st1 <> "" Then
                s_string = s_string & "AND [Γέννηση] LIKE '" & st1 & "'"
            End If
            If st2 <> "" Then
                s_string = s_string & " AND [Γέννηση] LIKE '" & st2 & "'"
            End If
            If st3 <> "" Then
                    s_string = s_string & " AND [Γέννηση] LIKE '" & st3 & "'"
            End If
            s_string2 = s_string2 & " AND [Γέννηση] LIKE '" & Trim(st4) & "'"
        Else
            If st1 <> "" Then
                s_string = "[Γέννηση] LIKE '" & st1 & "'"
            End If
            If st2 <> "" Then
                If s_string <> "" Then
                    s_string = s_string & " AND [Γέννηση] LIKE '" & st2 & "'"
                Else
                    s_string = "[Γέννηση] LIKE '" & st2 & "'"
                End If
                
            End If
            If st3 <> "" Then
                If s_string <> "" Then
                    s_string = s_string & " AND [Γέννηση] LIKE '" & st3 & "'"
                Else
                    s_string = "[Γέννηση] LIKE '" & st3 & "'"
                End If
            End If
            s_string2 = "[Γέννηση] LIKE '" & Trim(st4) & "'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΕΠΑΓΓΕΛΜΑΤΟΣ
    If Trim(co_jobs.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Επάγγελμα] LIKE '%" & Trim(co_jobs.Text) & "%'"
        Else
            s_string = "[Επάγγελμα] LIKE '%" & Trim(co_jobs.Text) & "%'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΤΗΛΕΦΩΝΟΥ ΟΙΚΙΑΣ
    If Trim(tmp_til_oikias.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΤηλέφωνοΟικίας] LIKE '*" & Trim(tmp_til_oikias.Text) & "*'"
        Else
            s_string = "[ΤηλέφωνοΟικίας] LIKE '*" & Trim(tmp_til_oikias.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΚΙΝΗΤΟΥ ΤΗΛΕΦΩΝΟΥ
    If Trim(tmp_kinito.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΚινητόΤηλέφωνο] LIKE '*" & Trim(tmp_kinito.Text) & "*'"
        Else
            s_string = "[ΚινητόΤηλέφωνο] LIKE '*" & Trim(tmp_kinito.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΑΡΙΘΜΟΥ FAX
    If Trim(tmp_fax.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΑριθμόςΦαξ] LIKE '*" & Trim(tmp_fax.Text) & "*'"
        Else
            s_string = "[ΑριθμόςΦαξ] LIKE '*" & Trim(tmp_fax.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ EMAIL
    If Trim(tmp_email.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΌνομαEmail] LIKE '*" & Trim(tmp_email.Text) & "*'"
        Else
            s_string = "[ΌνομαEmail] LIKE '*" & Trim(tmp_email.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΜΕΛΟΥΣ ΓΕΝΙΚΗΣ ΣΥΝΕΛΕΥΣΗΣ
    If s_string <> "" Then
            If ch_m_gs.Value = 0 Then
                's_string = s_string & " AND [ΜέλοςΓΣ] LIKE FALSE"
            Else
                s_string = s_string & " AND [ΜέλοςΓΣ] LIKE TRUE"
                rep_lbl = rep_lbl & " (ΜΕΛΗ ΓΕΝΙΚΗΣ ΣΥΝΕΛΕΥΣΗΣ)"
            End If
    Else
            If ch_m_gs.Value = 0 Then
                's_string = s_string & "[ΜέλοςΓΣ] LIKE FALSE"
            Else
                s_string = s_string & "[ΜέλοςΓΣ] LIKE TRUE"
                rep_lbl = rep_lbl & " (ΜΕΛΗ ΓΕΝΙΚΗΣ ΣΥΝΕΛΕΥΣΗΣ)"
            End If
    End If
    'ΚΡΙΤΗΡΙΟ ΑΜ Μέλους
    If Trim(tmp_am.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΑΜ_Μέλους] LIKE '*" & Trim(tmp_am.Text) & "*'"
        Else
            s_string = "[ΑΜ_Μέλους] LIKE '*" & Trim(tmp_am.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΗΜΕΡΟΜΗΝΙΑ ΕΓΓΡΑΦΗΣ
    If Trim(im_eg.Text) <> "00/00/0000" Then
        Me.im_eg.SelStart = 0
        Me.im_eg.SelLength = 2
        imera = Me.im_eg.SelText
        Me.im_eg.SelStart = 3
        Me.im_eg.SelLength = 2
        minas = Me.im_eg.SelText
        Me.im_eg.SelStart = 6
        Me.im_eg.SelLength = 4
        year = Me.im_eg.SelText
        Me.im_eg.SelStart = 0
        Me.im_eg.SelLength = 10
        st1 = ""
        If Val(imera) <> 0 Then
            st1 = imera & "/"
            st5 = imera & "/"
        Else
            st1 = ""
            st5 = ""
        End If
        If Val(minas) <> 0 Then
            If st1 = "" Then
                st2 = "*/" & minas & "/"
                st6 = "%/" & minas & "/"
            Else
                st2 = minas & "/"
                st6 = minas & "/"
            End If
        Else
            If st1 = "" Then
                st2 = ""
                st6 = ""
            Else
                st2 = "*/"
                st6 = "%/"
            End If
        End If
        If Val(year) <> 0 Then
            If st1 = "" And st2 = "" Then
                st3 = "*/" & year & "*"
                st7 = "%/" & year & "%"
            Else
                st3 = "" & year & "*"
                st7 = "" & year & "%"
            End If
        Else
            st3 = "*"
            st7 = "%"
        End If
        If Val(imera) <> 0 Then
            st1 = imera & "/*"
        Else
            st1 = ""
        End If
        If Val(minas) <> 0 Then
            st2 = "*/" & minas & "/*"
        Else
            st2 = ""
        End If
        If Val(year) <> 0 Then
            st3 = "*/" & year & "*"
        Else
            st3 = ""
        End If
        st = Trim(st1) & Trim(st2) & Trim(st3)
        st4 = st5 + st6 + st7
        ''
        If s_string <> "" Then
            If st1 <> "" Then
                s_string = s_string & "AND [ΗμΕγγραφής] LIKE '" & st1 & "'"
            End If
            If st2 <> "" Then
                s_string = s_string & " AND [ΗμΕγγραφής] LIKE '" & st2 & "'"
            End If
            If st3 <> "" Then
                    s_string = s_string & " AND [ΗμΕγγραφής] LIKE '" & st3 & "'"
            End If
            s_string2 = s_string2 & " AND [ΗμΕγγραφής] LIKE '" & Trim(st4) & "'"
        Else
            If st1 <> "" Then
                s_string = "[ΗμΕγγραφής] LIKE '" & st1 & "'"
            End If
            If st2 <> "" Then
                If s_string <> "" Then
                    s_string = s_string & " AND [ΗμΕγγραφής] LIKE '" & st2 & "'"
                Else
                    s_string = "[ΗμΕγγραφής] LIKE '" & st2 & "'"
                End If
            End If
            If st3 <> "" Then
                If s_string <> "" Then
                    s_string = s_string & " AND [ΗμΕγγραφής] LIKE '" & st3 & "'"
                Else
                    s_string = "[ΗμΕγγραφής] LIKE '" & st3 & "'"
                End If
            End If
            s_string2 = "[ΗμΕγγραφής] LIKE '" & Trim(st4) & "'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ Αριθμού Απόδειξης
    If Trim(Text4.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΑρΑπόδειξης] LIKE '*" & Trim(Text4.Text) & "*'"
        Else
            s_string = "[ΑρΑπόδειξης] LIKE '*" & Trim(Text4.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΜΕΛΟΥΣ ΔΣ
    If s_string <> "" Then
            If ch_m_ds.Value = 0 Then
                's_string = s_string & " AND [ΜέλοςΔΣ] LIKE FALSE"
            Else
                s_string = s_string & " AND [ΜέλοςΔΣ] LIKE TRUE"
                rep_lbl = rep_lbl & " (ΜΕΛΗ ΔΙΟΙΚΗΤΙΚΟΥ ΣΥΜΒΟΥΛΙΟΥ)"
            End If
    Else
            If ch_m_ds.Value = 0 Then
                's_string = s_string & "[ΜέλοςΔΣ] LIKE FALSE"
            Else
                s_string = s_string & "[ΜέλοςΔΣ] LIKE TRUE"
                rep_lbl = rep_lbl & " (ΜΕΛΗ ΔΙΟΙΚΗΤΙΚΟΥ ΣΥΜΒΟΥΛΙΟΥ)"
            End If
    End If
    'ΚΡΙΤΗΡΙΟ ΠΡΟΠΟΝΗΤΗ
    If s_string <> "" Then
            If ch_p.Value = 0 Then
                's_string = s_string & " AND [Προπονητής] LIKE FALSE"
            Else
                s_string = s_string & " AND [Προπονητής] LIKE TRUE"
                rep_lbl = rep_lbl & " (ΠΡΟΠΟΝΗΤΕΣ)"
            End If
    Else
            If ch_p.Value = 0 Then
                's_string = s_string & "[Προπονητής] LIKE FALSE"
            Else
                s_string = s_string & "[Προπονητής] LIKE TRUE"
                rep_lbl = rep_lbl & " (ΠΡΟΠΟΝΗΤΕΣ)"
            End If
    End If
    'ΚΡΙΤΗΡΙΟ ΕΘΕΛΟΝΤΗ
    If s_string <> "" Then
            If ch_e.Value = 0 Then
                's_string = s_string & " AND [Εθελοντής] LIKE FALSE"
            Else
                s_string = s_string & " AND [Εθελοντής] LIKE TRUE"
                rep_lbl = rep_lbl & " (ΕΘΕΛΟΝΤΕΣ)"
            End If
    Else
            If ch_e.Value = 0 Then
                's_string = s_string & "[Εθελοντής] LIKE FALSE"
            Else
                s_string = s_string & "[Εθελοντής] LIKE TRUE"
                rep_lbl = rep_lbl & " (ΕΘΕΛΟΝΤΕΣ)"
            End If
    End If
    'ΚΡΙΤΗΡΙΟ ΓΟΝΕΑ
    If s_string <> "" Then
            If ch_g.Value = 0 Then
                's_string = s_string & " AND [Γονέας] LIKE FALSE"
            Else
                s_string = s_string & " AND [Γονέας] LIKE TRUE"
                rep_lbl = rep_lbl & " (Γονείς)"
            End If
    Else
            If ch_g.Value = 0 Then
                's_string = s_string & "[Γονέας] LIKE FALSE"
            Else
                s_string = s_string & "[Γονέας] LIKE TRUE"
                rep_lbl = rep_lbl & " (Γονείς)"
            End If
    End If
    'ΚΡΙΤΗΡΙΟ ΔΡΟΜΕΑ
    If s_string <> "" Then
            If ch_dr.Value = 0 Then
                's_string = s_string & " AND [Δρομέας] LIKE FALSE"
            Else
                s_string = s_string & " AND [Δρομέας] LIKE TRUE"
                rep_lbl = rep_lbl & " (Δρομείς)"
            End If
    Else
            If ch_dr.Value = 0 Then
                's_string = s_string & "[Δρομέας] LIKE FALSE"
            Else
                s_string = s_string & "[Δρομέας] LIKE TRUE"
                rep_lbl = rep_lbl & " (Δρομείς)"
            End If
    End If
    'ΚΡΙΤΗΡΙΟ ΤΜΗΜΑ_ΕΝΗΛΙΚΩΝ
    If s_string <> "" Then
            If ch_tm_en.Value = 0 Then
                's_string = s_string & " AND [Τμήμα_Ενηλίκων] LIKE FALSE"
            Else
                s_string = s_string & " AND [Τμήμα_Ενηλίκων] LIKE TRUE"
                rep_lbl = rep_lbl & " (Ενήλικες Αθλητές Κολύμβησης)"
            End If
    Else
            If ch_tm_en.Value = 0 Then
                's_string = s_string & "[Τμήμα_Ενηλίκων] LIKE FALSE"
            Else
                s_string = s_string & "[Τμήμα_Ενηλίκων] LIKE TRUE"
                rep_lbl = rep_lbl & " (Ενήλικες Αθλητές Κολύμβησης)"
            End If
    End If
    'ΚΡΙΤΗΡΙΟ ΟΔΟΥ
    If Trim(tmp_odos.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Οδός] LIKE '*" & Trim(tmp_odos.Text) & "*'"
        Else
            s_string = "[Οδός] LIKE '*" & Trim(tmp_odos.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΑΡΙΘΜΟΥ ΟΔΟΥ
    If Trim(tmp_arithmos.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Αριθμός] LIKE '*" & Trim(tmp_arithmos.Text) & "*'"
        Else
            s_string = "[Αριθμός] LIKE '*" & Trim(tmp_arithmos.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΤΑΧΥΔΡΟΜΙΚΟΥ ΚΩΔΙΚΑ
    If Trim(tmp_tk.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Ταχυδρομικός Κώδικας] LIKE '*" & Trim(tmp_tk.Text) & "*'"
        Else
            s_string = "[Ταχυδρομικός Κώδικας] LIKE '*" & Trim(tmp_tk.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΠΕΡΙΟΧΗΣ
    If Trim(tmp_perioxi.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Περιοχή] LIKE '*" & Trim(tmp_perioxi.Text) & "*'"
        Else
            s_string = "[Περιοχή] LIKE '*" & Trim(tmp_perioxi.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΔΗΜΟΥ
    If Trim(co_dimoi.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Δήμος] LIKE '*" & Trim(co_dimoi.Text) & "*'"
        Else
            s_string = "[Δήμος] LIKE '*" & Trim(co_dimoi.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΠΕΡΙΦΕΡΕΙΑΚΗΣ ΕΝΟΤΗΤΑΣ
    If Trim(co_pe.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Περιφερειακή Ενότητα] LIKE '*" & Trim(co_pe.Text) & "*'"
        Else
            s_string = "[Περιφερειακή Ενότητα] LIKE '*" & Trim(co_pe.Text) & "*'"
        End If
    End If
    '
    'rs_ado_meli.Filter = s_string
    Me.ado_meli.Recordset.Filter = s_string
    MDIForm1.s_string = s_string
    MDIForm1.rep_lbl = rep_lbl
    If s_sort <> "" Then
        Me.ado_meli.Recordset.Sort = Trim(s_sort)
    End If
    '
    Me.canc_bt.Enabled = True
    Me.up_bt.Enabled = True
    Me.insert_bt.Enabled = True
    Me.del_bt.Enabled = True
    Me.cancel_cur_rec.Enabled = True
    '
            
End Sub

Private Sub taksin_Click()

    If Me.ado_meli.Recordset.RecordCount >= 1 Then
        If Me.dt_meli.Col >= 0 Then
            Me.ado_meli.Recordset.Sort = "[" & Trim(Me.ado_meli.Recordset.Fields(Me.dt_meli.Col).Name) & "]"
            s_sort = "[" & Trim(Me.ado_meli.Recordset.Fields(Me.dt_meli.Col).Name) & "]"
        Else
            Me.ado_meli.Recordset.Sort = "[" & Trim(Me.ado_meli.Recordset.Fields(defined_col).Name) & "]"
            s_sort = "[" & Trim(Me.ado_meli.Recordset.Fields(defined_col).Name) & "]"
        End If
        MDIForm1.s_sort = s_sort
        Me.ado_meli.Caption = "Μέλος 1 από " & Me.ado_meli.Recordset.RecordCount
        Me.canc_bt.Enabled = True
    End If
  
End Sub

Private Sub Text4_GotFocus()

    With Text4
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub tmp_am_GotFocus()

    With tmp_am
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub tmp_arithmos_GotFocus()

    With tmp_arithmos
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub tmp_email_GotFocus()

    With tmp_email
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub tmp_eponimo_GotFocus()

    With tmp_eponimo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub tmp_fax_GotFocus()

    With tmp_fax
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub tmp_kinito_GotFocus()

    With tmp_kinito
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub tmp_odos_GotFocus()

    With tmp_odos
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub tmp_onoma_GotFocus()

    With tmp_onoma
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub tmp_perioxi_GotFocus()

    With tmp_perioxi
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub tmp_til_oikias_GotFocus()

    With tmp_til_oikias
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub tmp_tk_GotFocus()

    With tmp_tk
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txt_adt_GotFocus()

    With txt_adt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txt_ea_adt_GotFocus()

    With txt_ea_adt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub up_bt_Click()


If Me.ado_meli.Recordset.RecordCount >= 1 Then


    Me.tmp_ado_meli.Recordset.MoveFirst
    id_α = Me.ado_meli.Recordset.Fields(0).Value
    Me.tmp_ado_meli.Recordset.Find "[id] = " & id_α
If Not Me.tmp_ado_meli.Recordset.EOF Then


    If Trim(Me.tmp_onoma.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(3).Value = Me.tmp_onoma.Text
    Else
        Me.tmp_ado_meli.Recordset.Fields(3).Value = ""
    End If
    If Trim(Me.tmp_eponimo.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(2).Value = Me.tmp_eponimo.Text
    Else
        Me.tmp_ado_meli.Recordset.Fields(2).Value = ""
    End If
    If Trim(Me.tmp_odos.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(4).Value = Me.tmp_odos.Text
    Else
        Me.tmp_ado_meli.Recordset.Fields(4).Value = ""
    End If
    If Trim(Me.tmp_arithmos.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(5).Value = Me.tmp_arithmos.Text
    Else
        Me.tmp_ado_meli.Recordset.Fields(5).Value = ""
    End If
    If Trim(Me.tmp_perioxi.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(6).Value = Me.tmp_perioxi.Text
    Else
        Me.tmp_ado_meli.Recordset.Fields(6).Value = ""
    End If
    If Me.co_dimoi.Text <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(7).Value = Me.co_dimoi.Text
            Me.ado_dimoi.Recordset.Requery
            Me.ado_dimoi.Refresh
            Me.co_dimoi.ReFill
            Me.co_dimoi.Text = Trim(Me.tmp_ado_meli.Recordset.Fields(7).Value)
            If Me.co_dimoi.MatchedWithList = False Then
            'ΕΙΣΑΓΩΓΗ ΝΕΟΥ ΔΗΜΟΥ
                Dim id As Integer
                If Me.ado_dimoi.Recordset.RecordCount >= 1 Then
                    Me.ado_dimoi.Recordset.Sort = "[id_δήμου]"
                    Me.ado_dimoi.Recordset.MoveLast
                    id = Me.ado_dimoi.Recordset![id_δήμου]
                Else
                    id = 0
                End If
            Me.ado_dimoi.Recordset.AddNew
            Me.ado_dimoi.Recordset.Fields(0) = id + 1
            Me.ado_dimoi.Recordset.Fields(1) = Trim(Me.co_dimoi.Text)
            Me.ado_dimoi.Recordset.UpdateBatch adAffectCurrent
            Me.ado_dimoi.Recordset.Requery
        End If
    Else
        Me.tmp_ado_meli.Recordset.Fields(7).Value = ""
    End If
    If Trim(Me.co_pe.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(8).Value = Me.co_pe.Text
        Me.ado_pe.Recordset.Requery
        Me.ado_pe.Refresh
        Me.co_pe.ReFill
        Me.co_pe.Text = Trim(Me.tmp_ado_meli.Recordset.Fields(8).Value)
        If Me.co_pe.MatchedWithList = False Then
        'ΕΙΣΑΓΩΓΗ ΝΕΑΣ ΠΕΡΙΦΕΡΕΙΑΚΗΣ ΕΝΟΤΗΤΑΣ
            If Me.ado_pe.Recordset.RecordCount >= 1 Then
                Me.ado_pe.Recordset.Sort = "[" & Trim(Me.ado_pe.Recordset.Fields(0).Name) & "]"
                Me.ado_pe.Recordset.MoveLast
                id = Me.ado_pe.Recordset![id_πε]
            Else
                id = 0
            End If
            Me.ado_pe.Recordset.AddNew
            Me.ado_pe.Recordset.Fields(0) = id + 1
            Me.ado_pe.Recordset.Fields(1) = Trim(Me.co_pe.Text)
            Me.ado_pe.Recordset.UpdateBatch adAffectCurrent
            Me.ado_pe.Recordset.Requery
        End If
        Else
            Me.tmp_ado_meli.Recordset.Fields(8).Value = ""
    End If
    If Trim(Me.tmp_tk.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(9).Value = Me.tmp_tk.Text
    Else
        Me.tmp_ado_meli.Recordset.Fields(9).Value = ""
    End If
    If Trim(Me.tmp_til_oikias.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(10).Value = Me.tmp_til_oikias.Text
    Else
        Me.tmp_ado_meli.Recordset.Fields(10).Value = ""
    End If
    If Trim(Me.tmp_kinito.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(11).Value = Me.tmp_kinito.Text
    Else
        Me.tmp_ado_meli.Recordset.Fields(11).Value = ""
    End If
    If Trim(Me.tmp_fax.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(12).Value = Me.tmp_fax.Text
    Else
        Me.tmp_ado_meli.Recordset.Fields(12).Value = ""
    End If
    If Trim(Me.tmp_email.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(13).Value = Me.tmp_email.Text
    Else
        Me.tmp_ado_meli.Recordset.Fields(13).Value = ""
    End If
    If Trim(Me.MaskEdBox1.Text) <> "00/00/0000" Then
        Me.tmp_ado_meli.Recordset.Fields(14).Value = Me.MaskEdBox1.Text
    Else
        Me.tmp_ado_meli.Recordset.Fields(14).Value = "00/00/0000"
    End If
    If Trim(Me.tmp_am.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(1).Value = Me.tmp_am.Text
    Else
        Me.tmp_ado_meli.Recordset.Fields(1) = ""
    End If
    Me.tmp_ado_meli.Recordset.Fields(15).Value = Me.ch_en.Value
    If Trim(Me.im_eg.Text) <> "00/00/0000" Then
        Me.tmp_ado_meli.Recordset.Fields(16).Value = Me.im_eg.Text
    Else
        Me.tmp_ado_meli.Recordset.Fields(16).Value = "00/00/0000"
    End If
    If Me.co_jobs.Text <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(17).Value = Me.co_jobs.Text
            If Me.co_jobs.MatchedWithList = False Then
            'ΕΙΣΑΓΩΓΗ ΝΕΟΥ ΕΠΑΓΓΕΛΜΑΤΟΣ
                If Me.ado_jobs.Recordset.RecordCount >= 1 Then
                    Me.ado_jobs.Recordset.Sort = "[id_επαγγέλματος]"
                    Me.ado_jobs.Recordset.MoveLast
                    id = Me.ado_jobs.Recordset![id_επαγγέλματος]
                Else
                    id = 0
                End If
            Me.ado_jobs.Recordset.AddNew
            Me.ado_jobs.Recordset.Fields(0) = id + 1
            Me.ado_jobs.Recordset.Fields(1) = Trim(Me.co_jobs.Text)
            Me.ado_jobs.Recordset.UpdateBatch adAffectCurrent
            Me.ado_jobs.Recordset.Requery
        End If
    Else
        Me.tmp_ado_meli.Recordset.Fields(17).Value = ""
    End If
    If Trim(Me.txt_adt.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(18).Value = Me.txt_adt.Text
    Else
        Me.tmp_ado_meli.Recordset.Fields(18).Value = ""
    End If
    If Trim(Me.txt_ea_adt.Text) <> "" Then
        Me.tmp_ado_meli.Recordset.Fields(19).Value = Me.txt_ea_adt.Text
    Else
        Me.tmp_ado_meli.Recordset.Fields(19).Value = ""
    End If
    Me.tmp_ado_meli.Recordset.Fields(20).Value = Me.ch_m_gs.Value
    Me.tmp_ado_meli.Recordset.Fields(21).Value = Me.ch_m_ds.Value
    Me.tmp_ado_meli.Recordset.Fields(22).Value = Me.ch_p.Value
    Me.tmp_ado_meli.Recordset.Fields(23).Value = Me.ch_e.Value
    Me.tmp_ado_meli.Recordset.Fields(25).Value = Me.ch_g.Value
    Me.tmp_ado_meli.Recordset.Fields("Δρομέας").Value = Me.ch_dr.Value
    Me.tmp_ado_meli.Recordset.Fields("Τμήμα_Ενηλίκων").Value = Me.ch_tm_en.Value
    '
    Me.tmp_ado_meli.Recordset.UpdateBatch adAffectCurrent
    
    Me.tmp_ado_meli.Recordset.Requery
    Me.tmp_ado_meli.Refresh
    Me.ado_meli.Recordset.Requery
    Me.ado_meli.Refresh
    
    Me.ado_meli.Recordset.Filter = MDIForm1.s_string
    If MDIForm1.s_sort <> "" Then
        Me.ado_meli.Recordset.Sort = MDIForm1.s_sort
    Else
        Me.ado_meli.Recordset.Sort = "[id]"
    End If
    If Me.ado_meli.Recordset.RecordCount >= 1 Then
        Me.ado_meli.Recordset.Find "[id] = " & id_α
        If Not Me.ado_meli.Recordset.EOF Then
            mv = Me.ado_meli.Recordset.AbsolutePosition
            Me.ado_meli.Recordset.MoveFirst
            Me.ado_meli.Recordset.Move mv - 1
        Else
            Me.ado_meli.Recordset.MoveFirst
        End If
    End If
    
    Me.dt_meli.Columns(0).Caption = "Κωδικός"
    Me.dt_meli.Columns(0).Width = 1500
    Me.dt_meli.Columns(1).Visible = False
    Me.dt_meli.Columns(2).Caption = "Επώνυμο"
    Me.dt_meli.Columns(2).Width = 3000
    Me.dt_meli.Columns(3).Caption = "Όνομα"
    Me.dt_meli.Columns(3).Width = 2000
    For i = 4 To 16
        Me.dt_meli.Columns(i).Visible = False
    Next i
    Me.dt_meli.Columns(17).Caption = "Επάγγελμα"
    Me.dt_meli.Columns(17).Width = 2000
    For i = 18 To Me.ado_meli.Recordset.Fields.Count - 1
        Me.dt_meli.Columns(i).Visible = False
    Next i
    
    
End If
End If
    
End Sub
