VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_pr_tm 
   BackColor       =   &H00C0C0FF&
   Caption         =   "ΕΚΤΥΠΩΣΕΙΣ, ΥΠΟΛΟΓΙΣΜΟΙ Κατηγορίας ΤΜΗΜΑΤΑ,  ΠΑΡΟΥΣΙΕΣ"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21075
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   21075
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000014&
      DisabledPicture =   "frm_pr_tm.frx":0000
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   14280
      MaskColor       =   &H80000014&
      Picture         =   "frm_pr_tm.frx":268B
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   """Εξαγωγή στοιχείων σε Excel"""
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000014&
      DisabledPicture =   "frm_pr_tm.frx":4D16
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6720
      MaskColor       =   &H80000014&
      Picture         =   "frm_pr_tm.frx":73A1
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   """Εξαγωγή στοιχείων σε Excel"""
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000014&
      DisabledPicture =   "frm_pr_tm.frx":9A2C
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6720
      MaskColor       =   &H80000014&
      Picture         =   "frm_pr_tm.frx":C0B7
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   """Εξαγωγή στοιχείων σε Excel"""
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000014&
      DisabledPicture =   "frm_pr_tm.frx":E742
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6720
      MaskColor       =   &H80000014&
      Picture         =   "frm_pr_tm.frx":10DCD
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   """Εξαγωγή στοιχείων σε Excel"""
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton kl_bt 
      BackColor       =   &H80000014&
      Caption         =   "Κ&λείσιμο"
      DisabledPicture =   "frm_pr_tm.frx":13458
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
      Left            =   5520
      MaskColor       =   &H80000014&
      Picture         =   "frm_pr_tm.frx":18ED0
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   10260
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Σεμπτέμβριος"
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   14
      Top             =   2880
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ιούλιος"
      Height          =   375
      Index           =   10
      Left            =   4920
      TabIndex        =   13
      Top             =   2880
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ιούνιος"
      Height          =   375
      Index           =   9
      Left            =   3120
      TabIndex        =   12
      Top             =   4320
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Μάιος"
      Height          =   375
      Index           =   8
      Left            =   3120
      TabIndex        =   11
      Top             =   3960
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Απρίλιος"
      Height          =   375
      Index           =   7
      Left            =   3120
      TabIndex        =   10
      Top             =   3600
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Μάρτιος"
      Height          =   375
      Index           =   6
      Left            =   3120
      TabIndex        =   9
      Top             =   3240
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Φεβρουάριος"
      Height          =   375
      Index           =   5
      Left            =   3120
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ιανουάριος"
      Height          =   375
      Index           =   4
      Left            =   1440
      TabIndex        =   7
      Top             =   4320
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Δεκέμβριος"
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Νοέμβριος"
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   5
      Top             =   3600
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Οκτώβριος"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "... συνέχισε ..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   5000
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Υπολογισμός Συνολικών Παρουσιών ΟΛΩΝ ΤΩΝ ΤΜΗΜΑΤΩΝ ανά Μήνα"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   6375
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   170
         TabIndex        =   16
         Top             =   300
         Width           =   4695
      End
   End
   Begin VB.CommandButton bt_epib 
      Caption         =   "... συνέχισε ..."
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   840
      ItemData        =   "frm_pr_tm.frx":1E948
      Left            =   1320
      List            =   "frm_pr_tm.frx":1E94A
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
   Begin MSAdodcLib.Adodc ado_sin_parousiwn_ana_ae_mina 
      Height          =   375
      Left            =   12120
      Top             =   2040
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
      RecordSource    =   "ΣύνολαΠαρουσιών2"
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
   Begin MSDataGridLib.DataGrid dt_sin_parousiwn_ana_ae_mina 
      Bindings        =   "frm_pr_tm.frx":1E94C
      Height          =   540
      Left            =   7200
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   953
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Enabled         =   0   'False
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7440
      Top             =   3360
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
      RecordSource    =   "ΣύνολαΠαρουσιών"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_pr_tm.frx":1E978
      Height          =   780
      Left            =   7200
      TabIndex        =   17
      Top             =   2760
      Visible         =   0   'False
      Width           =   12900
      _ExtentX        =   22754
      _ExtentY        =   1376
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Enabled         =   0   'False
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
   Begin MSAdodcLib.Adodc ado_sin_parousiwn_ana_ae_mina2 
      Height          =   375
      Left            =   12360
      Top             =   6480
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
      RecordSource    =   "ΣύνολοΠαρουσιώνΑνάΤμήμα2"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   15240
      Top             =   6480
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
      RecordSource    =   "ΣύνολοΠαρουσιώνΑνάΤμήμα"
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "frm_pr_tm.frx":1E98D
      Height          =   4155
      Left            =   7200
      TabIndex        =   22
      Top             =   6720
      Visible         =   0   'False
      Width           =   13740
      _ExtentX        =   24236
      _ExtentY        =   7329
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Enabled         =   -1  'True
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
   Begin MSDataGridLib.DataGrid dt_sin_parousiwn_ana_ae_mina3 
      Bindings        =   "frm_pr_tm.frx":1E9A2
      Height          =   2580
      Left            =   7200
      TabIndex        =   24
      Top             =   4080
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4551
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Enabled         =   -1  'True
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
   Begin MSAdodcLib.Adodc ado_plithos_paidiwn_poy_apoysiazoyn 
      Height          =   375
      Left            =   17400
      Top             =   3840
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
      RecordSource    =   "ΣύνολοΠαρουσιώνΑνάΤμήμα_ver2_5"
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
   Begin MSDataGridLib.DataGrid dt_plithos_paidiwn_poy_apoysiazoyn 
      Bindings        =   "frm_pr_tm.frx":1E9CF
      Height          =   2580
      Left            =   14760
      TabIndex        =   25
      Top             =   4080
      Visible         =   0   'False
      Width           =   6160
      _ExtentX        =   10874
      _ExtentY        =   4551
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Enabled         =   -1  'True
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
   Begin MSAdodcLib.Adodc ado_paidia_poy_irthan_ton_mina 
      Height          =   375
      Left            =   16920
      Top             =   1920
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
      RecordSource    =   "ΣύνολοΠαρουσιώνΑνάΤμήνα_ver2_7"
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
   Begin MSDataGridLib.DataGrid dt_paidia_poy_irthan_ton_mina 
      Bindings        =   "frm_pr_tm.frx":1EA01
      Height          =   540
      Left            =   14280
      TabIndex        =   27
      Top             =   2160
      Visible         =   0   'False
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   953
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Enabled         =   0   'False
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
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Επιλογή Εκτύπωσης:"
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
      Left            =   1320
      TabIndex        =   20
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frm_pr_tm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id_tmimatos_report As Integer
Dim minas As Integer

Private Sub bt_epib_Click()

    If List1.ListIndex = 0 Then
        Rep_Καρτέλα_1_Τμήματος.Show
        'frm_pr_tm.Hide
    Else
        If List1.ListIndex = 1 Then
            Rep_Καρτέλα_Τμήματος.Show
            'frm_pr_tm.Hide
        Else
            If List1.ListIndex = 2 Then
                Rep_ΈντυποΑπουσιώνΑνάΜήνα_1_Τμήματος.Show
                'frm_pr_tm.Hide
            Else
                MsgBox "Κάντε την επιλογή σας στο παραπάνω πλαίσιο για να συνεχίσετε...", , "Εφαρμογή Διαχείρισης ΠΟΣΕΙΔΩΝΑ"
            End If
        End If
    End If
    
End Sub

Private Sub Command1_Click()

    idae = tmima_management.ado_ae.Recordset.Fields(0).Value
    
    If Option1(0).Value = True Then
        minas = 9
    ElseIf Option1(1).Value = True Then
        minas = 10
    ElseIf Option1(2).Value = True Then
        minas = 11
    ElseIf Option1(3).Value = True Then
        minas = 12
    ElseIf Option1(4).Value = True Then
        minas = 1
    ElseIf Option1(5).Value = True Then
        minas = 2
    ElseIf Option1(6).Value = True Then
        minas = 3
    ElseIf Option1(7).Value = True Then
        minas = 4
    ElseIf Option1(8).Value = True Then
        minas = 5
    ElseIf Option1(9).Value = True Then
        minas = 6
    ElseIf Option1(10).Value = True Then
        minas = 7
    End If
    If minas <> 0 Then
    
        Me.ado_sin_parousiwn_ana_ae_mina.Recordset.Filter = "[id_ΑθλητικούΈτους] = " & idae & " and [id_mina] = " & minas
        Me.dt_sin_parousiwn_ana_ae_mina.Columns(0).Visible = False
        Me.dt_sin_parousiwn_ana_ae_mina.Columns(1).Visible = False
        Me.dt_sin_parousiwn_ana_ae_mina.Columns(2).Width = 6975
        Me.dt_sin_parousiwn_ana_ae_mina.Columns(2).Caption = "Σύνολο Παρουσιών ΟΛΩΝ ΤΩΝ ΤΜΗΜΑΤΩΝ ΜΑΖΙ Επιλεγμένου Μήνα: "
        Me.dt_sin_parousiwn_ana_ae_mina.Visible = True
        
        
        Me.ado_paidia_poy_irthan_ton_mina.Recordset.Filter = "[id_ΑθλητικούΈτους] = " & idae & " and [id_mina] = " & minas
        Me.dt_paidia_poy_irthan_ton_mina.Columns(0).Visible = False
        Me.dt_paidia_poy_irthan_ton_mina.Columns(1).Visible = False
        Me.dt_paidia_poy_irthan_ton_mina.Columns(2).Width = 3000
        Me.dt_paidia_poy_irthan_ton_mina.Columns(2).Caption = "με ΠΑΡΟΝΤΕΣ Αθλητές: "
        Me.dt_paidia_poy_irthan_ton_mina.Visible = True
        
        
        Me.Adodc1.Recordset.Filter = "[id_ΑθλητικούΈτους] = " & idae & " and [id_mina] = " & minas
        Me.DataGrid1.Caption = "Σύνολο Παρουσιών ΟΛΩΝ ΤΩΝ ΤΜΗΜΑΤΩΝ ΜΑΖΙ ανά ΗΜΕΡΑ του ΜΗΝΑ:                                                                                                                                              "
        Me.DataGrid1.Columns(0).Visible = False
        Me.DataGrid1.Columns(1).Visible = False
        Me.DataGrid1.Columns(2).Width = 400
        Me.DataGrid1.Columns(2).Caption = "1"
        Me.DataGrid1.Columns(3).Width = 400
        Me.DataGrid1.Columns(3).Caption = "2"
        Me.DataGrid1.Columns(4).Width = 400
        Me.DataGrid1.Columns(4).Caption = "3"
        Me.DataGrid1.Columns(5).Width = 400
        Me.DataGrid1.Columns(5).Caption = "4"
        Me.DataGrid1.Columns(6).Width = 400
        Me.DataGrid1.Columns(6).Caption = "5"
        Me.DataGrid1.Columns(7).Width = 400
        Me.DataGrid1.Columns(7).Caption = "6"
        Me.DataGrid1.Columns(8).Width = 400
        Me.DataGrid1.Columns(8).Caption = "7"
        Me.DataGrid1.Columns(9).Width = 400
        Me.DataGrid1.Columns(9).Caption = "8"
        Me.DataGrid1.Columns(10).Width = 400
        Me.DataGrid1.Columns(10).Caption = "9"
        Me.DataGrid1.Columns(11).Width = 400
        Me.DataGrid1.Columns(11).Caption = "10"
        Me.DataGrid1.Columns(12).Width = 400
        Me.DataGrid1.Columns(12).Caption = "11"
        Me.DataGrid1.Columns(13).Width = 400
        Me.DataGrid1.Columns(13).Caption = "12"
        Me.DataGrid1.Columns(14).Width = 400
        Me.DataGrid1.Columns(14).Caption = "13"
        Me.DataGrid1.Columns(15).Width = 400
        Me.DataGrid1.Columns(15).Caption = "14"
        Me.DataGrid1.Columns(16).Width = 400
        Me.DataGrid1.Columns(16).Caption = "15"
        Me.DataGrid1.Columns(17).Width = 400
        Me.DataGrid1.Columns(17).Caption = "16"
        Me.DataGrid1.Columns(18).Width = 400
        Me.DataGrid1.Columns(18).Caption = "17"
        Me.DataGrid1.Columns(19).Width = 400
        Me.DataGrid1.Columns(19).Caption = "18"
        Me.DataGrid1.Columns(20).Width = 400
        Me.DataGrid1.Columns(20).Caption = "19"
        Me.DataGrid1.Columns(21).Width = 400
        Me.DataGrid1.Columns(21).Caption = "20"
        Me.DataGrid1.Columns(22).Width = 400
        Me.DataGrid1.Columns(22).Caption = "21"
        Me.DataGrid1.Columns(23).Width = 400
        Me.DataGrid1.Columns(23).Caption = "22"
        Me.DataGrid1.Columns(24).Width = 400
        Me.DataGrid1.Columns(24).Caption = "23"
        Me.DataGrid1.Columns(25).Width = 400
        Me.DataGrid1.Columns(25).Caption = "24"
        Me.DataGrid1.Columns(26).Width = 400
        Me.DataGrid1.Columns(26).Caption = "25"
        Me.DataGrid1.Columns(27).Width = 400
        Me.DataGrid1.Columns(27).Caption = "26"
        Me.DataGrid1.Columns(28).Width = 400
        Me.DataGrid1.Columns(28).Caption = "27"
        Me.DataGrid1.Columns(29).Width = 400
        Me.DataGrid1.Columns(29).Caption = "28"
        Me.DataGrid1.Columns(30).Width = 400
        Me.DataGrid1.Columns(30).Caption = "29"
        Me.DataGrid1.Columns(31).Width = 400
        Me.DataGrid1.Columns(31).Caption = "30"
        Me.DataGrid1.Columns(32).Width = 400
        Me.DataGrid1.Columns(32).Caption = "31"
        Me.DataGrid1.Visible = True
        Me.Command6.Visible = True
        Me.Command1.Enabled = False
        
        
        Me.ado_sin_parousiwn_ana_ae_mina2.Recordset.Filter = "[id_ΑθλητικούΈτους] = " & idae & " and [id_mina] = " & minas
        Me.dt_sin_parousiwn_ana_ae_mina3.Columns(0).Width = 2000
        Me.dt_sin_parousiwn_ana_ae_mina3.Columns(0).Caption = "Τμήμα"
        Me.dt_sin_parousiwn_ana_ae_mina3.Columns(1).Visible = False
        Me.dt_sin_parousiwn_ana_ae_mina3.Columns(2).Width = 5300
        Me.dt_sin_parousiwn_ana_ae_mina3.Columns(2).Caption = "Σύνολο Παρουσιών ΑΝΑ ΤΜΗΜΑ Επιλεγμένου MHNA: "
        Me.dt_sin_parousiwn_ana_ae_mina3.Columns(3).Visible = False
        Me.dt_sin_parousiwn_ana_ae_mina3.Columns(4).Visible = False
        Me.dt_sin_parousiwn_ana_ae_mina3.Visible = True
        Me.Command2.Visible = True
        
        
        Me.ado_plithos_paidiwn_poy_apoysiazoyn.Recordset.Filter = "[id_ΑθλητικούΈτους] = " & idae & " and [id_mina] = " & minas
        Me.dt_plithos_paidiwn_poy_apoysiazoyn.Columns(0).Visible = False
        Me.dt_plithos_paidiwn_poy_apoysiazoyn.Columns(1).Visible = False
        Me.dt_plithos_paidiwn_poy_apoysiazoyn.Columns(2).Width = 2000
        Me.dt_plithos_paidiwn_poy_apoysiazoyn.Columns(2).Caption = "Τμήμα: "
        Me.dt_plithos_paidiwn_poy_apoysiazoyn.Columns(3).Visible = False
        Me.dt_plithos_paidiwn_poy_apoysiazoyn.Columns(4).Width = 2200
        Me.dt_plithos_paidiwn_poy_apoysiazoyn.Columns(4).Caption = "#ΑΘΛΗΤΩΝ ΑΠΟΥΣΙΑΣ: "
        Me.dt_plithos_paidiwn_poy_apoysiazoyn.Columns(5).Width = 1500
        Me.dt_plithos_paidiwn_poy_apoysiazoyn.Columns(5).Caption = "στο ΣΥΝΟΛΟ: "
        Me.dt_plithos_paidiwn_poy_apoysiazoyn.Visible = True
        Me.Command3.Visible = True


        Me.Adodc2.Recordset.Filter = "[id_ΑθλητικούΈτους] = " & idae & " and [id_mina] = " & minas
        Me.DataGrid3.Caption = "Σύνολο Παρουσιών ΑΝΑ ΤΜΗΜΑ και ΑΝΑ ΗΜΕΡΑ του ΜΗΝΑ:                                                                                                                                                                                       "
        Me.DataGrid3.Columns(0).Width = 1000
        Me.DataGrid3.Columns(0).Caption = "Τμήμα"
        Me.DataGrid3.Columns(1).Visible = False
        Me.DataGrid3.Columns(2).Width = 400
        Me.DataGrid3.Columns(2).Caption = "1"
        Me.DataGrid3.Columns(3).Width = 400
        Me.DataGrid3.Columns(3).Caption = "2"
        Me.DataGrid3.Columns(4).Width = 400
        Me.DataGrid3.Columns(4).Caption = "3"
        Me.DataGrid3.Columns(5).Width = 400
        Me.DataGrid3.Columns(5).Caption = "4"
        Me.DataGrid3.Columns(6).Width = 400
        Me.DataGrid3.Columns(6).Caption = "5"
        Me.DataGrid3.Columns(7).Width = 400
        Me.DataGrid3.Columns(7).Caption = "6"
        Me.DataGrid3.Columns(8).Width = 400
        Me.DataGrid3.Columns(8).Caption = "7"
        Me.DataGrid3.Columns(9).Width = 400
        Me.DataGrid3.Columns(9).Caption = "8"
        Me.DataGrid3.Columns(10).Width = 400
        Me.DataGrid3.Columns(10).Caption = "9"
        Me.DataGrid3.Columns(11).Width = 400
        Me.DataGrid3.Columns(11).Caption = "10"
        Me.DataGrid3.Columns(12).Width = 400
        Me.DataGrid3.Columns(12).Caption = "11"
        Me.DataGrid3.Columns(13).Width = 400
        Me.DataGrid3.Columns(13).Caption = "12"
        Me.DataGrid3.Columns(14).Width = 400
        Me.DataGrid3.Columns(14).Caption = "13"
        Me.DataGrid3.Columns(15).Width = 400
        Me.DataGrid3.Columns(15).Caption = "14"
        Me.DataGrid3.Columns(16).Width = 400
        Me.DataGrid3.Columns(16).Caption = "15"
        Me.DataGrid3.Columns(17).Width = 400
        Me.DataGrid3.Columns(17).Caption = "16"
        Me.DataGrid3.Columns(18).Width = 400
        Me.DataGrid3.Columns(18).Caption = "17"
        Me.DataGrid3.Columns(19).Width = 400
        Me.DataGrid3.Columns(19).Caption = "18"
        Me.DataGrid3.Columns(20).Width = 400
        Me.DataGrid3.Columns(20).Caption = "19"
        Me.DataGrid3.Columns(21).Width = 400
        Me.DataGrid3.Columns(21).Caption = "20"
        Me.DataGrid3.Columns(22).Width = 400
        Me.DataGrid3.Columns(22).Caption = "21"
        Me.DataGrid3.Columns(23).Width = 400
        Me.DataGrid3.Columns(23).Caption = "22"
        Me.DataGrid3.Columns(24).Width = 400
        Me.DataGrid3.Columns(24).Caption = "23"
        Me.DataGrid3.Columns(25).Width = 400
        Me.DataGrid3.Columns(25).Caption = "24"
        Me.DataGrid3.Columns(26).Width = 400
        Me.DataGrid3.Columns(26).Caption = "25"
        Me.DataGrid3.Columns(27).Width = 400
        Me.DataGrid3.Columns(27).Caption = "26"
        Me.DataGrid3.Columns(28).Width = 400
        Me.DataGrid3.Columns(28).Caption = "27"
        Me.DataGrid3.Columns(29).Width = 400
        Me.DataGrid3.Columns(29).Caption = "28"
        Me.DataGrid3.Columns(30).Width = 400
        Me.DataGrid3.Columns(30).Caption = "29"
        Me.DataGrid3.Columns(31).Width = 400
        Me.DataGrid3.Columns(31).Caption = "30"
        Me.DataGrid3.Columns(32).Width = 400
        Me.DataGrid3.Columns(32).Caption = "31"
        Me.DataGrid3.Columns(33).Visible = False
        Me.DataGrid3.Columns(34).Visible = False
        Me.DataGrid3.Visible = True
        Me.Command4.Visible = True
        
    Else
        MsgBox "Επιλέξτε πρώτα μήνα για να συνεχίσετε...", , "Εφαρμογή Διαχείρισης ΠΟΣΕΙΔΩΝΑ"
    End If
    
End Sub

Private Sub Command2_Click()
   
   Dim oExcel As Object
   Dim oBook As Object
   Dim oSheet As Object
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add
   Set oSheet = oBook.Worksheets(1)
   
   Clipboard.Clear
   Dim sData As Variant
   sData = ""
   If Me.ado_sin_parousiwn_ana_ae_mina2.Recordset.RecordCount >= 1 Then
        Me.ado_sin_parousiwn_ana_ae_mina2.Recordset.MoveFirst
        sData = "ΣΥΝΟΛΟ ΠΑΡΟΥΣΙΩΝ ΑΝΑ ΤΜΗΜΑ ΓΙΑ ΤΟΝ ΕΠΙΛΕΓΜΕΝΟ ΜΗΝΑ:" & vbCr
        sData = sData & "Αθλητικό Έτος" & vbTab & "Τμήμα" & vbTab & "Μήνας" & vbTab & "Σύνολο Παρουσιών" & vbCr
        For i = 0 To Me.Adodc2.Recordset.RecordCount - 1
            sData = sData & tmima_management.co_ae.Text & vbTab & Me.ado_sin_parousiwn_ana_ae_mina2.Recordset.Fields(0) & vbTab & minas & vbTab & Me.ado_sin_parousiwn_ana_ae_mina2.Recordset.Fields(2) & vbCr
            Me.ado_sin_parousiwn_ana_ae_mina2.Recordset.MoveNext
        Next i
   End If
   Clipboard.SetText sData
   oBook.Worksheets(1).Range("A1").Select
   oBook.Worksheets(1).Paste
   oSheet.Range("A1:D2").Font.Bold = True
   oSheet.Range("A2:D2").Font.ColorIndex = 3
   oSheet.Range("A2:d3").ColumnWidth = 25
   oBook.Worksheets(1).Columns.HorizontalAlignment = -4131
   oExcel.Visible = True
   

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
   If Me.ado_plithos_paidiwn_poy_apoysiazoyn.Recordset.RecordCount >= 1 Then
        Me.ado_plithos_paidiwn_poy_apoysiazoyn.Recordset.MoveFirst
        sData = "ΠΛΗΘΟΣ ΑΘΛΗΤΩΝ που ΑΠΕΙΧΑΝ ΤΟΝ ΕΠΙΛΕΓΜΕΝΟ ΜΗΝΑ:" & vbCr
        sData = sData & "Αθλητικό Έτος" & vbTab & "Τμήμα" & vbTab & "Μήνας" & vbTab & "Πλήθος Αθλητών που ΔΕΝ ΗΡΘΕ τον επιλεγμένο μήνα" & vbTab & "(από το ΣΥΝΟΛΟ των ΕΓΓΕΓΡΑΜΜΕΝΩΝ Αθλητών)" & vbCr
        For i = 0 To Me.ado_plithos_paidiwn_poy_apoysiazoyn.Recordset.RecordCount - 1
            sData = sData & tmima_management.co_ae.Text & vbTab & Me.ado_plithos_paidiwn_poy_apoysiazoyn.Recordset.Fields(2) & vbTab & minas & vbTab & Me.ado_plithos_paidiwn_poy_apoysiazoyn.Recordset.Fields(4) & vbTab & Me.ado_plithos_paidiwn_poy_apoysiazoyn.Recordset.Fields(5) & vbCr
            Me.ado_plithos_paidiwn_poy_apoysiazoyn.Recordset.MoveNext
        Next i
   End If
   Clipboard.SetText sData
   oBook.Worksheets(1).Range("A1").Select
   oBook.Worksheets(1).Paste
   oSheet.Range("A1:E2").Font.Bold = True
   oSheet.Range("A2:E2").Font.ColorIndex = 3
   oSheet.Range("A2:E3").ColumnWidth = 35
   oSheet.Range("D2").WrapText = True
   oSheet.Range("E2").WrapText = True
   oBook.Worksheets(1).Columns.HorizontalAlignment = -4131
   oExcel.Visible = True
   


End Sub

Private Sub Command4_Click()
   
   Dim oExcel As Object
   Dim oBook As Object
   Dim oSheet As Object
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add
   Set oSheet = oBook.Worksheets(1)
   
   Clipboard.Clear
   Dim sData As Variant
   sData = ""
   If Me.Adodc2.Recordset.RecordCount >= 1 Then
        Me.Adodc2.Recordset.MoveFirst
        sData = "" & vbTab & "" & vbTab & "" & vbTab & "ΣΥΝΟΛΟ ΠΑΡΟΥΣΙΩΝ ANA ΤΜΗΜΑ και ΑΝΑ ΗΜΕΡΑ για τον Επιλεγμένο ΜΗΝΑ: " & vbCr
        sData = sData & "" & vbTab & "" & vbTab & vbTab & "ΗΜΕΡΕΣ:" & vbCr
        sData = sData & "Αθλητικό Έτος" & vbTab & "Τμήμα" & vbTab & "Μήνας" & vbTab & "1" & vbTab & "2" & vbTab & "3" & vbTab & "4" & vbTab & "5" & vbTab & "6" & vbTab & "7" & vbTab & "8" & vbTab & "9" & vbTab & "10" & vbTab & "11" & vbTab & "12" & vbTab & "13" & vbTab & "14" & vbTab & "15" & vbTab & "16" & vbTab & "17" & vbTab & "18" & vbTab & "19" & vbTab & "20" & vbTab & "21" & vbTab & "22" & vbTab & "23" & vbTab & "24" & vbTab & "25" & vbTab & "26" & vbTab & "27" & vbTab & "28" & vbTab & "29" & vbTab & "30" & vbTab & "31" & vbCr
        For i = 0 To Me.Adodc2.Recordset.RecordCount - 1
            sData = sData & tmima_management.co_ae.Text & vbTab & Me.Adodc2.Recordset.Fields(0) & vbTab & minas & vbTab & Me.Adodc2.Recordset.Fields(2) & vbTab & Me.Adodc2.Recordset.Fields(3) & vbTab & Me.Adodc2.Recordset.Fields(4) & vbTab & Me.Adodc2.Recordset.Fields(5) & vbTab & Me.Adodc2.Recordset.Fields(6) & vbTab & Me.Adodc2.Recordset.Fields(7) & vbTab & Me.Adodc2.Recordset.Fields(8) & vbTab & Me.Adodc2.Recordset.Fields(9) & vbTab & Me.Adodc2.Recordset.Fields(10) & vbTab & Me.Adodc2.Recordset.Fields(11) & vbTab & Me.Adodc2.Recordset.Fields(12) & vbTab & Me.Adodc2.Recordset.Fields(13) & vbTab & Me.Adodc2.Recordset.Fields(14) & vbTab & Me.Adodc2.Recordset.Fields(15) & vbTab & Me.Adodc2.Recordset.Fields(16) & vbTab & Me.Adodc2.Recordset.Fields(17) & vbTab & Me.Adodc2.Recordset.Fields(18) & vbTab & Me.Adodc2.Recordset.Fields(19) & vbTab & Me.Adodc2.Recordset.Fields(20) & vbTab & Me.Adodc2.Recordset.Fields(21) & vbTab & Me.Adodc2.Recordset.Fields(22) & vbTab & Me.Adodc2.Recordset.Fields(23)
            sData = sData & vbTab & Me.Adodc2.Recordset.Fields(24) & vbTab & Me.Adodc2.Recordset.Fields(25) & vbTab & Me.Adodc2.Recordset.Fields(26) & vbTab & Me.Adodc2.Recordset.Fields(27) & vbTab & Me.Adodc2.Recordset.Fields(28) & vbTab & Me.Adodc2.Recordset.Fields(29) & vbTab & Me.Adodc2.Recordset.Fields(30) & vbTab & Me.Adodc2.Recordset.Fields(31) & vbCr
            Me.Adodc2.Recordset.MoveNext
        Next i
   End If
   Clipboard.SetText sData
   oBook.Worksheets(1).Range("A1").Select
   oBook.Worksheets(1).Paste
   oSheet.Range("A1:AH3").Font.Bold = True
   oSheet.Range("A3:C3").Font.ColorIndex = 3
   oSheet.Range("D1:AH1").merge
   oSheet.Range("D2:AH2").merge
   oSheet.Range("A3:C4").ColumnWidth = 20
   oSheet.Range("D3:AH3").ColumnWidth = 5
   oBook.Worksheets(1).Columns.HorizontalAlignment = -4131
   oExcel.Visible = True
   
End Sub

Private Sub Command6_Click()
   
   Dim oExcel As Object
   Dim oBook As Object
   Dim oSheet As Object
   Set oExcel = CreateObject("Excel.Application")
   Set oBook = oExcel.Workbooks.Add
   Set oSheet = oBook.Worksheets(1)
   
   Clipboard.Clear
   Dim sData As Variant
   sData = ""
   If Me.Adodc1.Recordset.RecordCount >= 1 Then
        Me.Adodc1.Recordset.MoveFirst
        sData = "" & vbTab & "" & vbTab & "ΣΥΝΟΛΟ ΠΑΡΟΥΣΙΩΝ ΑΝΑ ΗΜΕΡΑ ΓΙΑ ΟΛΑ ΤΑ ΤΜΗΜΑΤΑ ΜΑΖΙ: " & vbCr
        sData = sData & "" & vbTab & "" & vbTab & "ΗΜΕΡΕΣ:" & vbCr
        sData = sData & "Αθλητικό Έτος" & vbTab & "Μήνας" & vbTab & "1" & vbTab & "2" & vbTab & "3" & vbTab & "4" & vbTab & "5" & vbTab & "6" & vbTab & "7" & vbTab & "8" & vbTab & "9" & vbTab & "10" & vbTab & "11" & vbTab & "12" & vbTab & "13" & vbTab & "14" & vbTab & "15" & vbTab & "16" & vbTab & "17" & vbTab & "18" & vbTab & "19" & vbTab & "20" & vbTab & "21" & vbTab & "22" & vbTab & "23" & vbTab & "24" & vbTab & "25" & vbTab & "26" & vbTab & "27" & vbTab & "28" & vbTab & "29" & vbTab & "30" & vbTab & "31" & vbCr
        For i = 0 To Me.Adodc1.Recordset.RecordCount - 1
            sData = sData & tmima_management.co_ae.Text & vbTab & minas & vbTab & Me.Adodc1.Recordset.Fields(2) & vbTab & Me.Adodc1.Recordset.Fields(3) & vbTab & Me.Adodc1.Recordset.Fields(4) & vbTab & Me.Adodc1.Recordset.Fields(5) & vbTab & Me.Adodc1.Recordset.Fields(6) & vbTab & Me.Adodc1.Recordset.Fields(7) & vbTab & Me.Adodc1.Recordset.Fields(8) & vbTab & Me.Adodc1.Recordset.Fields(9) & vbTab & Me.Adodc1.Recordset.Fields(10) & vbTab & Me.Adodc1.Recordset.Fields(11) & vbTab & Me.Adodc1.Recordset.Fields(12) & vbTab & Me.Adodc1.Recordset.Fields(13) & vbTab & Me.Adodc1.Recordset.Fields(14) & vbTab & Me.Adodc1.Recordset.Fields(15) & vbTab & Me.Adodc1.Recordset.Fields(16) & vbTab & Me.Adodc1.Recordset.Fields(17) & vbTab & Me.Adodc1.Recordset.Fields(18) & vbTab & Me.Adodc1.Recordset.Fields(19) & vbTab & Me.Adodc1.Recordset.Fields(20) & vbTab & Me.Adodc1.Recordset.Fields(21) & vbTab & Me.Adodc1.Recordset.Fields(22) & vbTab & Me.Adodc1.Recordset.Fields(23)
            sData = sData & vbTab & Me.Adodc1.Recordset.Fields(24) & vbTab & Me.Adodc1.Recordset.Fields(25) & vbTab & Me.Adodc1.Recordset.Fields(26) & vbTab & Me.Adodc1.Recordset.Fields(27) & vbTab & Me.Adodc1.Recordset.Fields(28) & vbTab & Me.Adodc1.Recordset.Fields(29) & vbTab & Me.Adodc1.Recordset.Fields(30) & vbTab & Me.Adodc1.Recordset.Fields(31) & vbCr
            Me.Adodc1.Recordset.MoveNext
        Next i
   End If
   Clipboard.SetText sData
   oBook.Worksheets(1).Range("A1").Select
   oBook.Worksheets(1).Paste
   oSheet.Range("A1:AG3").Font.Bold = True
   oSheet.Range("A4:B4").Font.ColorIndex = 3
   oSheet.Range("C1:AG1").merge
   oSheet.Range("C2:AG2").merge
   oSheet.Range("A3:B3").ColumnWidth = 15
   oSheet.Range("C3:AG3").ColumnWidth = 5
   oBook.Worksheets(1).Columns.HorizontalAlignment = -4131
   oExcel.Visible = True
   

End Sub

Private Sub Form_Load()
    
    frm_pr_tm.List1.AddItem "Εκτύπωση Καρτέλας Τρέχοντος Τμήματος (Γενικά Στοιχεία)"
    frm_pr_tm.List1.AddItem "Εκτύπωση Καρτέλας Όλων των Τμημάτων (Γενικά Στοιχεία)"
    frm_pr_tm.List1.AddItem "Εκτύπωση Παρουσιολογίου Τμήματος ανά Μήνα"
    
    Me.Label1 = "για το Αθλητικό Έτος: " & tmima_management.co_ae.Text & "."

End Sub

Private Sub kl_bt_Click()

    Unload Me

End Sub

Private Sub List1_Click()

    Me.dt_sin_parousiwn_ana_ae_mina.Visible = False
    Me.dt_paidia_poy_irthan_ton_mina.Visible = False
    Me.DataGrid1.Visible = False
    Me.Command6.Visible = False
    Me.Command1.Enabled = False
    Me.DataGrid3.Visible = False
    Me.Command4.Visible = False
    Me.dt_plithos_paidiwn_poy_apoysiazoyn.Visible = False
    Me.Command3.Visible = False
    
    Me.dt_sin_parousiwn_ana_ae_mina3.Visible = False
    Me.DataGrid3.Visible = False
    Me.Command2.Visible = False

    For i = 0 To 10
        Me.Option1(i).FontBold = False
        Me.Option1(i).Value = False
    Next i

End Sub

Private Sub Option1_Click(Index As Integer)

    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = False
    Next i
    
    Me.DataGrid3.Visible = False
    Me.Command2.Visible = False

    For i = 0 To 10
        Me.Option1(i).FontBold = False
    Next i
    Me.Option1(Index).FontBold = True

    Me.dt_sin_parousiwn_ana_ae_mina.Visible = False
    Me.ado_sin_parousiwn_ana_ae_mina.Recordset.Requery
    Me.dt_paidia_poy_irthan_ton_mina.Visible = False
    Me.DataGrid1.Visible = False
    Me.Adodc1.Recordset.Requery
    Me.dt_sin_parousiwn_ana_ae_mina3.Visible = False
    Me.Command6.Visible = False
    Me.Command2.Visible = False
    Me.Command4.Visible = False
    Me.Command1.Enabled = True
    Me.Command3.Visible = False
    Me.dt_plithos_paidiwn_poy_apoysiazoyn.Visible = False

End Sub
