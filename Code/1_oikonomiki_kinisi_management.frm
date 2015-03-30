VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form MIA_oikonomiki_kinisi_management 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Στοιχεία Οικονομικής Κίνησης"
   ClientHeight    =   8265
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12585
   ForeColor       =   &H000000FF&
   LinkTopic       =   "entity_management"
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   12585
   Begin VB.CommandButton bt_cl_dt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "x"
      DisabledPicture =   "1_oikonomiki_kinisi_management.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Picture         =   "1_oikonomiki_kinisi_management.frx":02E0
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Ενημέρωση Αιτιολογίας"
      Top             =   1200
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   330
   End
   Begin MSDataGridLib.DataGrid dt_paroysies_athliti 
      Bindings        =   "1_oikonomiki_kinisi_management.frx":4D0E
      Height          =   2220
      Left            =   7440
      TabIndex        =   68
      Top             =   1200
      Visible         =   0   'False
      Width           =   4810
      _ExtentX        =   8493
      _ExtentY        =   3916
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
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
   Begin VB.CommandButton bt_refr_esoda 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ανανέωση"
      DisabledPicture =   "1_oikonomiki_kinisi_management.frx":4D32
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10680
      MaskColor       =   &H80000014&
      Picture         =   "1_oikonomiki_kinisi_management.frx":9B60
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "ΠΡΟΣ ΤΟ ΠΑΡΟΝ ΤΟ ΚΟΥΜΠΙ ΑΝΑΝΕΩΝΕΙ ΑΘΛΗΤΕΣ ΚΑΙ ΜΕΛΗ ΜΟΝΟ"
      Top             =   100
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dt_anal_py_eks 
      Bindings        =   "1_oikonomiki_kinisi_management.frx":E98E
      Height          =   375
      Left            =   4440
      TabIndex        =   59
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
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
   Begin MSDataGridLib.DataGrid dt_anal_py_es 
      Bindings        =   "1_oikonomiki_kinisi_management.frx":E9AC
      Height          =   375
      Left            =   1200
      TabIndex        =   58
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
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
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000014&
      Caption         =   "Εκτύπωση "
      DisabledPicture =   "1_oikonomiki_kinisi_management.frx":E9C9
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
      Left            =   9600
      MaskColor       =   &H80000014&
      Picture         =   "1_oikonomiki_kinisi_management.frx":133CE
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Εκτύπωση ΤΡΕΧΟΥΣΑΣ Οικονομικής Κίνησης"
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H80000014&
      Caption         =   "Κλείσιμο"
      DisabledPicture =   "1_oikonomiki_kinisi_management.frx":17DD3
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
      Left            =   10920
      MaskColor       =   &H80000014&
      Picture         =   "1_oikonomiki_kinisi_management.frx":17F1B
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cancel 
      BackColor       =   &H80000014&
      Caption         =   "Ακύρωση"
      DisabledPicture =   "1_oikonomiki_kinisi_management.frx":1D993
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
      Left            =   8280
      MaskColor       =   &H80000014&
      Picture         =   "1_oikonomiki_kinisi_management.frx":1DADB
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Storage 
      Caption         =   "Αποθήκευση"
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
      Left            =   4440
      Picture         =   "1_oikonomiki_kinisi_management.frx":229F8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton update 
      Caption         =   "Ενημέρωση"
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
      Left            =   6960
      Picture         =   "1_oikonomiki_kinisi_management.frx":275F7
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Clean 
      BackColor       =   &H80000014&
      Caption         =   "Καθαρισμός"
      DisabledPicture =   "1_oikonomiki_kinisi_management.frx":2F0F1
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
      Left            =   5640
      MaskColor       =   &H80000014&
      Picture         =   "1_oikonomiki_kinisi_management.frx":2F239
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6840
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc ado_athlites 
      Height          =   375
      Left            =   8880
      Top             =   120
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
      RecordSource    =   "ΟνοματεπώνυμαΑθλητών"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Στοιχεία Οικονομικής Κίνησης"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   6615
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   12135
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "1_oikonomiki_kinisi_management.frx":33D7D
         Height          =   375
         Left            =   5760
         TabIndex        =   51
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid dt_tipoi_parastatikwn 
         Bindings        =   "1_oikonomiki_kinisi_management.frx":33D94
         Height          =   255
         Left            =   3600
         TabIndex        =   29
         Top             =   3240
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
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
      Begin VB.Frame Eksoda 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Στοιχεία Χρέωσης"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   6120
         TabIndex        =   34
         Top             =   5280
         Width           =   6015
         Begin VB.ComboBox co_raw_tipoi_eksodwn 
            Height          =   315
            Left            =   2520
            TabIndex        =   10
            Text            =   "co_raw_tipoi_eksodwn"
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox f12 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   2520
            TabIndex        =   11
            Top             =   720
            Width           =   1455
         End
         Begin MSDataGridLib.DataGrid dt_py_eksodwn 
            Bindings        =   "1_oikonomiki_kinisi_management.frx":33DB9
            Height          =   375
            Left            =   5040
            TabIndex        =   35
            Top             =   600
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
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
         Begin MSAdodcLib.Adodc ado_py_eksodwn 
            Height          =   375
            Left            =   5040
            Top             =   240
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
            RecordSource    =   "ΠροϋπολογισμοίΈξοδα"
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
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Τύπος Εξόδου:"
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   6
            Left            =   1200
            TabIndex        =   37
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ποσό Χρέωσης (σε €):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   4
            Left            =   0
            TabIndex        =   36
            Top             =   720
            Width           =   2415
         End
      End
      Begin VB.ComboBox raw_co_tipoi_parastatikwn 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Top             =   2760
         Width           =   3915
      End
      Begin MSAdodcLib.Adodc ado_organismoi 
         Height          =   375
         Left            =   10200
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
         RecordSource    =   "ΟνοματεπώνυμαΟργανισμών"
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
      Begin MSAdodcLib.Adodc ado_meli 
         Height          =   375
         Left            =   7200
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
         RecordSource    =   "ΟνοματεπώνυμαΜελών"
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
      Begin VB.Frame Frame6 
         Caption         =   "Η Κίνηση Αφορά ΤΟΝ/ΤΟΥΣ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   6120
         TabIndex        =   16
         Top             =   360
         Width           =   6015
         Begin MSAdodcLib.Adodc ado_paroysies_athliti 
            Height          =   375
            Left            =   5640
            Top             =   480
            Visible         =   0   'False
            Width           =   1740
            _ExtentX        =   3069
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
            RecordSource    =   "ΣύνολοΠαρουσιώνΑθλητήΑνάΜήνα"
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
         Begin VB.CommandButton bt_par_athl 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Π"
            DisabledPicture =   "1_oikonomiki_kinisi_management.frx":33DD6
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
            Height          =   290
            Left            =   5520
            MaskColor       =   &H00FFFFFF&
            Picture         =   "1_oikonomiki_kinisi_management.frx":340B6
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Ενημέρωση Αιτιολογίας"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   255
         End
         Begin VB.CommandButton bt_en_athl 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "V"
            DisabledPicture =   "1_oikonomiki_kinisi_management.frx":38AE4
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
            Height          =   290
            Left            =   5200
            MaskColor       =   &H00FFFFFF&
            Picture         =   "1_oikonomiki_kinisi_management.frx":3A209
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Ενημέρωση Αιτιολογίας"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   255
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Ο"
            Height          =   290
            Left            =   4920
            TabIndex        =   57
            ToolTipText     =   "Εμφανίζει ΟΛΟΥΣ τους Γονείς"
            Top             =   720
            Width           =   255
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Ο"
            Height          =   290
            Left            =   4920
            TabIndex        =   56
            ToolTipText     =   "Εμφανίζει ΟΛΑ τα Παιδιά"
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Π"
            Height          =   290
            Left            =   4635
            TabIndex        =   55
            ToolTipText     =   "Αναζητά τα Παιδιά του Επελεγμένου Γονέα"
            Top             =   720
            Width           =   255
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Γ"
            Height          =   290
            Left            =   4635
            TabIndex        =   54
            ToolTipText     =   "Αναζητά τους Γονείς του Επιλεγμένου Παιδιού"
            Top             =   360
            Width           =   255
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Bindings        =   "1_oikonomiki_kinisi_management.frx":3B92E
            Height          =   375
            Left            =   2640
            TabIndex        =   53
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
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
         Begin VB.ComboBox co_athlites 
            Height          =   315
            Left            =   1200
            TabIndex        =   52
            Top             =   360
            Width           =   3375
         End
         Begin VB.ComboBox co_meli 
            Height          =   315
            Left            =   1200
            TabIndex        =   50
            Top             =   720
            Width           =   3375
         End
         Begin MSDataListLib.DataCombo co_organismoi 
            Bindings        =   "1_oikonomiki_kinisi_management.frx":3B949
            Height          =   315
            Left            =   1200
            TabIndex        =   5
            Top             =   1080
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "Επωνυμία"
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Οργανισμός:"
            Height          =   315
            Left            =   0
            TabIndex        =   20
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Μέλος:"
            Height          =   315
            Left            =   0
            TabIndex        =   19
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Αθλητής:"
            Height          =   315
            Left            =   0
            TabIndex        =   18
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Esoda 
         BackColor       =   &H00008000&
         Caption         =   "Στοιχεία Πίστωσης"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   6120
         TabIndex        =   15
         Top             =   2040
         Width           =   6015
         Begin VB.CommandButton bt_kath_a 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "Χ"
            DisabledPicture =   "1_oikonomiki_kinisi_management.frx":3B966
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
            Left            =   4640
            MaskColor       =   &H00FFFFFF&
            Picture         =   "1_oikonomiki_kinisi_management.frx":40394
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Καθαρισμός Αιτιολογίας"
            Top             =   2220
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton bt_en_a 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "V"
            DisabledPicture =   "1_oikonomiki_kinisi_management.frx":44DC2
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
            Left            =   4640
            MaskColor       =   &H00FFFFFF&
            Picture         =   "1_oikonomiki_kinisi_management.frx":464E7
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Ενημέρωση Αιτιολογίας"
            Top             =   1840
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.ListBox lst_mn 
            Enabled         =   0   'False
            Height          =   1860
            ItemData        =   "1_oikonomiki_kinisi_management.frx":47C0C
            Left            =   2520
            List            =   "1_oikonomiki_kinisi_management.frx":47C34
            Style           =   1  'Checkbox
            TabIndex        =   62
            Top             =   720
            Width           =   2055
         End
         Begin MSDataGridLib.DataGrid dt_py_esodwn 
            Bindings        =   "1_oikonomiki_kinisi_management.frx":47CB7
            Height          =   375
            Left            =   5040
            TabIndex        =   44
            Top             =   600
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
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
         Begin VB.TextBox f15 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   2520
            TabIndex        =   7
            Top             =   2640
            Width           =   1455
         End
         Begin VB.ComboBox co_raw_tipoi_esodwn 
            Height          =   315
            Left            =   2520
            TabIndex        =   6
            Top             =   360
            Width           =   2055
         End
         Begin MSAdodcLib.Adodc ado_py_esodwn 
            Height          =   375
            Left            =   4440
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
            RecordSource    =   "ΠροϋπολογισμοίΈσοδα"
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
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Μήνας Συνδρομής Αθλητή:"
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   61
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ποσό Πίστωσης (σε €):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   33
            Top             =   2640
            Width           =   2295
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Τύπος Εσόδου:"
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   1
            Left            =   1200
            TabIndex        =   32
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Γενικά Στοιχεία"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6165
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   5895
         Begin VB.TextBox txt_aitiol2 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   600
            MultiLine       =   -1  'True
            TabIndex        =   66
            Top             =   4920
            Visible         =   0   'False
            Width           =   1245
         End
         Begin MSAdodcLib.Adodc ado_py 
            Height          =   375
            Left            =   480
            Top             =   4200
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
            RecordSource    =   "Προϋπολογισμός"
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
         Begin VB.TextBox f3 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox f16 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Left            =   1920
            MultiLine       =   -1  'True
            TabIndex        =   39
            Top             =   3840
            Width           =   3885
         End
         Begin VB.TextBox f9 
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox f8 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Frame Frame8 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1335
            Left            =   1920
            TabIndex        =   25
            Top             =   240
            Width           =   2775
            Begin VB.OptionButton opt_f2 
               Caption         =   "ΑΚΥΡΗ ΚΙΝΗΣΗ"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   161
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   28
               Top             =   120
               Width           =   2415
            End
            Begin VB.OptionButton opt_f2 
               Caption         =   "ΔΕΣΜΕΥΜΕΝΗ ΚΙΝΗΣΗ"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   161
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000080FF&
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   27
               Top             =   480
               Width           =   2295
            End
            Begin VB.OptionButton opt_f2 
               Caption         =   "ΕΝΕΡΓΗ ΚΙΝΗΣΗ"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   161
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   26
               Top             =   840
               Value           =   -1  'True
               Width           =   2175
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   360
            TabIndex        =   22
            Top             =   240
            Width           =   1455
            Begin VB.OptionButton opt_f1 
               Caption         =   "ΠΙΣΤΩΣΗ"
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
               Index           =   0
               Left            =   120
               TabIndex        =   24
               Top             =   120
               Width           =   1215
            End
            Begin VB.OptionButton opt_f1 
               Caption         =   "ΧΡΕΩΣΗ"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   161
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   23
               Top             =   480
               Width           =   1215
            End
         End
         Begin MSAdodcLib.Adodc ado_tipoi_parastatikwn 
            Height          =   375
            Left            =   3480
            Top             =   2520
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
            RecordSource    =   "ΤύποιΠαραστατικώνΕσόδωνΕξόδων"
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
         Begin VB.TextBox f0 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   0
            Top             =   1680
            Width           =   1335
         End
         Begin MSDataGridLib.DataGrid dt_py 
            Bindings        =   "1_oikonomiki_kinisi_management.frx":47CD3
            Height          =   375
            Left            =   1080
            TabIndex        =   43
            Top             =   4440
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
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
         Begin MSDataListLib.DataCombo co_py 
            Bindings        =   "1_oikonomiki_kinisi_management.frx":47CEF
            Height          =   315
            Left            =   1920
            TabIndex        =   45
            Top             =   3480
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "Περιγραφή"
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Αντιστοίχιση στον Προϋπολογισμό:"
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   7
            Left            =   360
            TabIndex        =   42
            Top             =   3400
            Width           =   1455
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "( ηη/μμ/εεεε )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Index           =   4
            Left            =   3120
            TabIndex        =   41
            Top             =   3195
            Width           =   1455
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "( ηη/μμ/εεεε )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Index           =   3
            Left            =   3120
            TabIndex        =   40
            Top             =   2100
            Width           =   1455
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Αιτιολογία:"
            Height          =   315
            Index           =   2
            Left            =   840
            TabIndex        =   38
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ημ/νία Παραστατικού:"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   3120
            Width           =   1695
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Αριθμός Παρ/κού:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   30
            Top             =   2760
            Width           =   1725
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Τύπος Παραστατικού:"
            Height          =   300
            Index           =   0
            Left            =   0
            TabIndex        =   21
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Κωδικός:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            TabIndex        =   17
            Top             =   1680
            Width           =   1005
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ημ/νία Κίνησης:"
            Height          =   315
            Index           =   0
            Left            =   360
            TabIndex        =   14
            Top             =   2040
            Width           =   1455
         End
      End
   End
   Begin MSAdodcLib.Adodc ado_anal_py_es 
      Height          =   375
      Left            =   0
      Top             =   7920
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
      RecordSource    =   "ΑνάλυσηΠροϋπολογισμούΕσόδων"
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
   Begin MSAdodcLib.Adodc ado_anal_py_eks 
      Height          =   375
      Left            =   3240
      Top             =   7920
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
      RecordSource    =   "ΑνάλυσηΠροϋπολογισμούΕξόδων"
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
End
Attribute VB_Name = "MIA_oikonomiki_kinisi_management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id_α, defined_col As Integer
Public rs_ado_dimoi As ADODB.Recordset
Public rs_ado_jobs As ADODB.Recordset
Public rs_ado_meli As ADODB.Recordset
Public rs_ado_pe As ADODB.Recordset
Public for_search As Integer
Public s_sort As String
Public kl As Integer

Private Sub bt_cl_dt_Click()

    Me.dt_paroysies_athliti.Visible = False
    Me.bt_cl_dt.Visible = False

End Sub

Private Sub bt_en_a_Click()
    
    Dim i As Integer
    Dim str_a As String

    'Ενημέρωση της Αιτιολογίας με τους ΜΗΝΕΣ Συνδρομής του Αθλητή
    str_a = ""
    For i = 0 To Me.lst_mn.ListCount - 1
        If Me.lst_mn.Selected(i) = True Then
            If str_a = "" Then
                str_a = "<<" & Me.lst_mn.List(i)
            Else
                str_a = str_a & ", " & Me.lst_mn.List(i)
            End If
        End If
    Next i
    ye = year(Me.ado_py.Recordset.Fields(2).Value)
    yl = year(Me.ado_py.Recordset.Fields(3).Value)
    str_a = str_a & " Αθλητικού Έτους " & ye & "-" & yl & ">>"
    If Me.f16.Text = "" Then
        Me.f16.Text = str_a
    Else
        Me.f16.Text = Me.f16.Text & " " & str_a
    End If
    
End Sub

Private Sub bt_en_athl_Click()
   
    'Ενημέρωση της Αιτιολογίας με το όνομα του Αθλητή
    If Me.co_athlites.Text <> "" Then
        If Me.f16.Text = "" Then
            Me.f16.Text = "<<" & Me.co_athlites.Text & ">>"
        Else
            Me.f16.Text = Me.f16.Text & " <<" & Me.co_athlites.Text & ">>"
        End If
        If Me.txt_aitiol2.Text = "" Then
            Me.txt_aitiol2.Text = "<<" & Me.co_athlites.Text & ">>"
        Else
            Me.txt_aitiol2.Text = Me.txt_aitiol2.Text & " <<" & Me.co_athlites.Text & ">>"
        End If
    End If
    
End Sub

Private Sub bt_kath_a_Click()
    
    Me.f16.Text = ""
    
End Sub

Private Sub bt_par_athl_Click()
    
    Me.ado_paroysies_athliti.Recordset.Filter = "[id_athliti] = " & ΕνημέρωσηΑθλητή
    Me.dt_paroysies_athliti.Columns(0).Width = 1500
    Me.dt_paroysies_athliti.Columns(0).Caption = "Αθλητικό Έτος: "
    Me.dt_paroysies_athliti.Columns(1).Visible = False
    Me.dt_paroysies_athliti.Columns(2).Width = 1200
    Me.dt_paroysies_athliti.Columns(2).Caption = "Μήνας: "
    Me.dt_paroysies_athliti.Columns(3).Visible = False
    Me.dt_paroysies_athliti.Columns(4).Width = 1800
    Me.dt_paroysies_athliti.Columns(4).Caption = "με ΠΑΡΟΥΣΙΕΣ: "
    Me.dt_paroysies_athliti.Columns(5).Visible = False
    Me.dt_paroysies_athliti.Columns(6).Visible = False
    Me.dt_paroysies_athliti.Visible = True
    Me.bt_cl_dt.Visible = True

End Sub

Private Sub bt_refr_esoda_Click()


    'ΑΥΤΗ Η ΑΝΑΝΕΩΣΗ ΠΡΟΣ ΤΟ ΠΑΡΟΝ ΑΝΑΝΕΩΝΕΙ ΜΟΝΟ ΜΕΛΗ ΚΑΙ ΑΘΛΗΤΕΣ
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Me.ado_meli.Recordset.Requery
    Me.ado_meli.Refresh
    Me.DataGrid1.Refresh
    Me.co_meli.Refresh
    'ΕΝΗΜΕΡΩΣΗ ΤΗΣ ΛΙΣΤΑΣ ΤΩΝ ΜΕΛΩΝ
    MIA_oikonomiki_kinisi_management.co_meli.Clear
    MIA_oikonomiki_kinisi_management.ado_meli.Recordset.Filter = ""
    If MIA_oikonomiki_kinisi_management.ado_meli.Recordset.RecordCount >= 1 Then
        MIA_oikonomiki_kinisi_management.ado_meli.Recordset.MoveFirst
        For i = 0 To MIA_oikonomiki_kinisi_management.ado_meli.Recordset.RecordCount - 1
            MIA_oikonomiki_kinisi_management.co_meli.AddItem (MIA_oikonomiki_kinisi_management.ado_meli.Recordset.Fields(1).Value)
            MIA_oikonomiki_kinisi_management.ado_meli.Recordset.MoveNext
        Next i
    End If
    '
    
    Me.ado_athlites.Recordset.Requery
    Me.ado_athlites.Refresh
    Me.DataGrid2.Refresh
    Me.co_athlites.Refresh
    'ΕΝΗΜΕΡΩΣΗ ΤΗΣ ΛΙΣΤΑΣ ΤΩΝ ΑΘΛΗΤΩΝ
    MIA_oikonomiki_kinisi_management.co_athlites.Clear
    MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.Filter = ""
    If MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.RecordCount >= 1 Then
        MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.MoveFirst
        For i = 0 To MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.RecordCount - 1
            MIA_oikonomiki_kinisi_management.co_athlites.AddItem (MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.Fields(1).Value)
            MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.MoveNext
        Next i
    End If
    
End Sub

Private Sub cancel_Click()

    'id
    MIA_oikonomiki_kinisi_management.f0.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(0).Value
    'Προϋπολογισμός
    'MIA_oikonomiki_kinisi_management.co_py.Text = ΕύρεσηΑΜΟΠ("Προϋπολογισμός", 1, oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(1).Value)
    MIA_oikonomiki_kinisi_management.co_py.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(2).Value
    'Τύπος Κίνησης
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(3).Value = "ΧΡΕΩΣΗ" Then
        MIA_oikonomiki_kinisi_management.opt_f1(0).Value = False
        MIA_oikonomiki_kinisi_management.opt_f1(1).Value = True
    Else 'ΕΙΝΑΙ ΠΙΣΤΩΣΗ
        MIA_oikonomiki_kinisi_management.opt_f1(0).Value = True
        MIA_oikonomiki_kinisi_management.opt_f1(1).Value = False
    End If
    'Κατάσταση Κίνησης
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(4).Value = "ΑΚΥΡΗ" Then
        MIA_oikonomiki_kinisi_management.opt_f2(0).Value = True
        MIA_oikonomiki_kinisi_management.opt_f2(1).Value = False
        MIA_oikonomiki_kinisi_management.opt_f2(2).Value = False
    Else
        If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(4).Value = "ΔΕΣΜΕΥΜΕΝΗ" Then
            MIA_oikonomiki_kinisi_management.opt_f2(0).Value = False
            MIA_oikonomiki_kinisi_management.opt_f2(1).Value = True
            MIA_oikonomiki_kinisi_management.opt_f2(2).Value = False
        Else
            MIA_oikonomiki_kinisi_management.opt_f2(0).Value = False
            MIA_oikonomiki_kinisi_management.opt_f2(1).Value = False
            MIA_oikonomiki_kinisi_management.opt_f2(2).Value = True
        End If
    End If
    'Ημερομηνία Κίνησης
    MIA_oikonomiki_kinisi_management.f3.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(5).Value
    'Αθλητής
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(6).Value <> "" Then
        MIA_oikonomiki_kinisi_management.co_athlites.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(6).Value
    Else
        MIA_oikonomiki_kinisi_management.co_athlites.Text = ""
    End If
    'Μέλος
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(7).Value <> "" Then
        MIA_oikonomiki_kinisi_management.co_meli.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(7).Value
    Else
        MIA_oikonomiki_kinisi_management.co_meli.Text = ""
    End If
    'Οργανισμός
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(8).Value <> "" Then
        MIA_oikonomiki_kinisi_management.co_organismoi.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(8).Value
    Else
        MIA_oikonomiki_kinisi_management.co_organismoi.Text = ""
    End If
    'Τύπος Παραστατικού
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(9).Value <> "" Then
        MIA_oikonomiki_kinisi_management.raw_co_tipoi_parastatikwn.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(9).Value
    Else
        MIA_oikonomiki_kinisi_management.raw_co_tipoi_parastatikwn.Text = ""
    End If
    'Αριθμός Παραστατικού
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(10).Value <> "" Then
        MIA_oikonomiki_kinisi_management.f8.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(10).Value
    Else
        MIA_oikonomiki_kinisi_management.f8.Text = ""
    End If
    'Ημερομηνία Παραστατικού
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(11).Value <> "" Then
        MIA_oikonomiki_kinisi_management.f9.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(11).Value
    Else
        MIA_oikonomiki_kinisi_management.f9.Text = ""
    End If
    'Τύπος ΠΥ Εξόδων
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(12).Value <> "" Then
        MIA_oikonomiki_kinisi_management.co_raw_tipoi_eksodwn.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(12).Value
    Else
        MIA_oikonomiki_kinisi_management.co_raw_tipoi_eksodwn.Text = ""
    End If
    'Ποσό Χρέωσης
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(13).Value <> "" And oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(13).Value <> 0 Then
        MIA_oikonomiki_kinisi_management.f12.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(13).Value & " €"
    Else
        MIA_oikonomiki_kinisi_management.f12.Text = ""
    End If
    'Τύπος ΠΥ Εσόδων
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(14).Value <> "" Then
        MIA_oikonomiki_kinisi_management.co_raw_tipoi_esodwn.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(14).Value
    Else
        MIA_oikonomiki_kinisi_management.co_raw_tipoi_esodwn.Text = ""
    End If
    'Μήνες Συνδρομής Αθλητή
    For i = 0 To 11
        If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(32 + i).Value = True Then
            MIA_oikonomiki_kinisi_management.lst_mn.Selected(i) = True
        Else
            MIA_oikonomiki_kinisi_management.lst_mn.Selected(i) = False
        End If
    Next i
    'Ποσό Πίστωσης
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(15).Value <> "" And oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(15).Value <> 0 Then
        MIA_oikonomiki_kinisi_management.f15.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(15).Value & " €"
    Else
        MIA_oikonomiki_kinisi_management.f15.Text = ""
    End If
    'Αιτιολογία
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(16).Value <> "" Then
        MIA_oikonomiki_kinisi_management.f16.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(16).Value
    Else
        MIA_oikonomiki_kinisi_management.f16.Text = ""
    End If
    '
    MIA_oikonomiki_kinisi_management.raw_co_tipoi_parastatikwn.Locked = True
    MIA_oikonomiki_kinisi_management.Frame1.Enabled = False
    '
    'MIA_oikonomiki_kinisi_management.update.Enabled = True
    'MIA_oikonomiki_kinisi_management.Command7.Enabled = True

End Sub

Private Sub Clean_Click()

    Dim id_par As Integer

    'id
    'Προϋπολογισμός
    'Me.co_py.Text = ""
    'ΤύποςΚίνησης
    Me.opt_f1(0).Value = 1
    Me.opt_f1(1).Value = 0
    'ΚατάστασηΚίνησης
    Me.opt_f2(0).Value = 0
    Me.opt_f2(1).Value = 0
    Me.opt_f2(2).Value = 1
    ' ΗΜΕΡΟΜΗΝΙΑ ΚΙΝΗΣΗΣ
    Me.f3.Text = DateValue(Now)
    ' ΑΘΛΗΤΗΣ
    Me.co_athlites.Text = ""
    ' ΜΕΛΟΣ
    Me.co_meli.Text = ""
    ' ΟΡΓΑΝΙΣΜΟΣ
    Me.co_organismoi.Text = ""
    ' ΤΥΠΟΣ ΠΑΡΑΣΤΑΤΙΚΟΎ
    Me.raw_co_tipoi_parastatikwn.Text = ""
    ' ΑΡΙΘΜΟΣ ΠΑΡΑΣΤΑΤΙΚΟΎ
    Me.f8.Text = ""
    'ΗμερομηνίαΠαραστατικού
    Me.f9.Text = DateValue(Now)
    'ΤύποςΠΥΕξόδων
    Me.co_raw_tipoi_eksodwn.Text = ""
    'ΠοσόΧρέωσης
    Me.f12.Text = ""
    'ΤύποςΠΥΕσόδων
    Me.co_raw_tipoi_esodwn.Text = ""
    'ΜήνεςΣυνδρομήςΑθλητή
    For i = 0 To 11
        Me.lst_mn.Selected(i) = False
    Next i
    Me.lst_mn.Enabled = False
    Me.bt_en_a.Enabled = False
    Me.bt_kath_a.Enabled = False
    'ΠοσόΠίστωσης
    Me.f15.Text = ""
    'Αιτιολογία
    Me.f16.Text = ""
    
End Sub

Private Sub co_athlites_Change()
    
    Me.bt_en_athl.Enabled = True
    Me.bt_par_athl.Enabled = True
    
End Sub

Private Sub co_athlites_Click()
    
    Me.bt_en_athl.Enabled = True
    Me.bt_par_athl.Enabled = True
    Me.dt_paroysies_athliti.Visible = False
    Me.bt_cl_dt.Visible = False
    
End Sub

Private Sub co_py_Change()

    If Me.co_py.Text <> "" Then
        Me.ado_py.Recordset.MoveFirst
        Me.ado_py.Recordset.Find "[Περιγραφή] like '" & Trim(Me.co_py.Text) & "'"
    End If
    Call Refresh_ΤύποιΕσόδων(1)
    Call Refresh_ΤύποιΕξόδων(1)

End Sub

Private Sub co_raw_tipoi_esodwn_Click()

    Dim mn As Integer
    
    If co_raw_tipoi_esodwn.Text = "Συνδρομές Αθλητών" Then
        Me.Label3(2).Enabled = True
        Me.lst_mn.Enabled = True
        Me.bt_en_a.Enabled = True
        Me.bt_kath_a.Enabled = True
        mn = Month(Now)
        Me.lst_mn.ListIndex = mn - 1
        Me.lst_mn.SetFocus
    Else
        Me.Label3(2).Enabled = False
        Me.lst_mn.Enabled = False
        Me.lst_mn.ListIndex = -1
        Me.bt_en_a.Enabled = False
        Me.bt_kath_a.Enabled = False
    End If
    
End Sub

Private Sub Command1_Click()
    
    tmp_id_a = ΕνημέρωσηΑθλητή
    If tmp_id_a >= 1 Then
        Me.co_meli.Clear
        Me.ado_athlites.Recordset.Filter = "[id] = " & tmp_id_a
        If Me.ado_athlites.Recordset.RecordCount >= 1 Then
            Me.ado_athlites.Recordset.MoveFirst
            For i = 0 To ado_athlites.Recordset.RecordCount - 1
                Me.co_meli.AddItem (Me.ado_athlites.Recordset.Fields(4).Value)
                Me.co_meli.AddItem (Me.ado_athlites.Recordset.Fields(5).Value)
                Me.ado_athlites.Recordset.MoveNext
            Next i
        End If
    Else
        Call Command3_Click
    End If

        
End Sub

Private Sub dt_meli_HeadClick(ByVal ColIndex As Integer)

    defined_col = ColIndex
    
End Sub

Private Sub met_st_Click()

    melos_id = rs_ado_meli.Fields(0).Value
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
        athlet_management.rs_ado_pateres.Filter = "[id] LIKE '" & str(melos_id) & "'"
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
        athlet_management.rs_ado_miteres.Filter = "[id] LIKE '" & str(melos_id) & "'"
        athlet_management.flag_mitera = 0
    End If
    
    ' ΜΕΤΑΦΟΡΑ ΥΠΟΨΗΦΙΩΝ ΙΔΙΩΝ ΣΤΟΙΧΕΙΩΝ ΑΠΟ ΓΟΝΕΑ ΣΕ ΠΑΙΔΙ
    If Trim(rs_ado_meli.Fields(4).Value) <> "" And Trim(athlet_management.tmp_odos.Text) = "" Then
        athlet_management.tmp_odos.Text = rs_ado_meli.Fields(4).Value
    End If
    If Trim(rs_ado_meli.Fields(5).Value) <> "" And Trim(athlet_management.tmp_arithmos.Text) = "" Then
        athlet_management.tmp_arithmos.Text = rs_ado_meli.Fields(5).Value
    End If
    If Trim(rs_ado_meli.Fields(6).Value) <> "" And Trim(athlet_management.tmp_perioxi.Text) = "" Then
        athlet_management.tmp_perioxi.Text = rs_ado_meli.Fields(6).Value
    End If
    If Trim(rs_ado_meli.Fields(8).Value) <> "" And Trim(athlet_management.co_pe.Text) = "" Then
        athlet_management.co_pe.Text = rs_ado_meli.Fields(8).Value
    End If
    If Trim(rs_ado_meli.Fields(7).Value) <> "" And Trim(athlet_management.co_dimoi.Text) = "" Then
        athlet_management.co_dimoi.Text = rs_ado_meli.Fields(7).Value
    End If
    If Trim(rs_ado_meli.Fields(9).Value) <> "" And Trim(athlet_management.tmp_tk.Text) = "" Then
        athlet_management.tmp_tk.Text = rs_ado_meli.Fields(9).Value
    End If
    If Trim(rs_ado_meli.Fields(10).Value) <> "" And Trim(athlet_management.tmp_til_oikias.Text) = "" Then
        athlet_management.tmp_til_oikias.Text = rs_ado_meli.Fields(10).Value
    End If
    If Trim(rs_ado_meli.Fields(11).Value) <> "" And Trim(athlet_management.tmp_kinito.Text) = "" Then
        athlet_management.tmp_kinito.Text = rs_ado_meli.Fields(11).Value
    End If
    If Trim(rs_ado_meli.Fields(12).Value) <> "" And Trim(athlet_management.tmp_fax.Text) = "" Then
        athlet_management.tmp_fax.Text = rs_ado_meli.Fields(12).Value
    End If
    If Trim(rs_ado_meli.Fields(13).Value) <> "" And Trim(athlet_management.tmp_email.Text) = "" Then
        athlet_management.tmp_email.Text = rs_ado_meli.Fields(13).Value
    End If
    '
    
    
    'meli_management.Hide
    Unload meli_management

End Sub

Private Sub del_bt_Click()
    
    Dim ms As String
    
    If Not rs_ado_meli.EOF And rs_ado_meli.AbsolutePosition > 1 Then
    ms = MsgBox("Είσαι σίγουρος; (ΝΑΙ ή ΟΧΙ)", vbYesNo, "Παράθυρο διαγραφής")
    If ms = 6 Then
        rs_ado_meli.Delete
    End If
    Else
        MsgBox "Δεν υπάρχει εγγραφή προς ΔΙΑΓΡΑΦΗ!", vbCritical, "Μήνυμα Λάθους"
    End If
    
End Sub

Private Sub co_raw_tipoi_eksodwn_GotFocus()
    
    Call Refresh_ΤύποιΕξόδων(1)
    
End Sub

Private Sub co_raw_tipoi_esodwn_GotFocus()
    
    Call Refresh_ΤύποιΕσόδων(1)

End Sub

Private Sub Command2_Click()
    
    tmp_id_m = ΕνημέρωσηΜέλους
    If tmp_id_m >= 1 Then
        Me.co_athlites.Clear
        Me.ado_athlites.Recordset.Filter = "[id_πατέρα] = " & tmp_id_m
        If Me.ado_athlites.Recordset.RecordCount >= 1 Then
            Me.ado_athlites.Recordset.MoveFirst
            For i = 0 To ado_athlites.Recordset.RecordCount - 1
                Me.co_athlites.AddItem (Me.ado_athlites.Recordset.Fields(1).Value)
                Me.ado_athlites.Recordset.MoveNext
            Next i
        End If
        Me.ado_athlites.Recordset.Filter = "[id_μητέρας] = " & tmp_id_m
        If Me.ado_athlites.Recordset.RecordCount >= 1 Then
            Me.ado_athlites.Recordset.MoveFirst
            For i = 0 To ado_athlites.Recordset.RecordCount - 1
                Me.co_athlites.AddItem (Me.ado_athlites.Recordset.Fields(1).Value)
                Me.ado_athlites.Recordset.MoveNext
            Next i
        End If
    Else
        Call Command4_Click
    End If

End Sub

Private Sub Command3_Click()
    
    Me.co_athlites.Clear
    Me.ado_athlites.Recordset.Filter = ""
    If Me.ado_athlites.Recordset.RecordCount >= 1 Then
        Me.ado_athlites.Recordset.MoveFirst
        For i = 0 To Me.ado_athlites.Recordset.RecordCount - 1
            Me.co_athlites.AddItem (Me.ado_athlites.Recordset.Fields(1).Value)
            Me.ado_athlites.Recordset.MoveNext
        Next i
    End If

End Sub

Private Sub Command4_Click()
    
    Me.co_meli.Clear
    Me.ado_meli.Recordset.Filter = ""
    If Me.ado_meli.Recordset.RecordCount >= 1 Then
        Me.ado_meli.Recordset.MoveFirst
        For i = 0 To Me.ado_meli.Recordset.RecordCount - 1
            Me.co_meli.AddItem (Me.ado_meli.Recordset.Fields(1).Value)
            Me.ado_meli.Recordset.MoveNext
        Next i
    End If
    
End Sub

Private Sub Command5_Click()
        
    Call oikonomikes_kiniseis_management.Command5_Click
        
End Sub

Private Sub Command9_Click()

    If Me.Storage.Enabled = False And Me.update.Enabled = True Then
        'ms = MsgBox("Έχετε αποθηκεύσει τις αλλαγές στη Β.Δ. της Εφαρμογής Διαχείρισης Ποσειδώνα;", vbYesNo, "Μήνυμα Προειδοποίησης")
        ms = MsgBox("Είσαι σίγουρος για το κλείσιμο;", vbYesNo, "Μήνυμα Προειδοποίησης")
        If ms = vbYes Then
            kl = 1
            Unload Me
        Else
            'kl = 1
            'Unload Me
        End If
    Else
        kl = 1
        Unload Me
    End If
    
End Sub

Private Sub f12_GotFocus()
    
    On Error GoTo l1
    
    With f12
        If .Text <> "" Then
            '.Text = Val(.Text)
            .Text = CDbl(.Text)
        End If
    End With

l1:
    
End Sub

Private Sub f12_LostFocus()

    Dim s As String
    Dim i As Integer

    'With f12
    '    If .Text <> "" Then
    '        s = .Text
    '        For i = 1 To Len(s)
    '            If (Mid$(s, i, 1) <> "0" And Mid$(s, i, 1) <> "1" And Mid$(s, i, 1) <> "2" And Mid$(s, i, 1) <> "3" And Mid$(s, i, 1) <> "4" And Mid$(s, i, 1) <> "5" And Mid$(s, i, 1) <> "6" And Mid$(s, i, 1) <> "7" And Mid$(s, i, 1) <> "8" And Mid$(s, i, 1) <> "9" And Mid$(s, i, 1) <> "," And Asc(Mid$(s, i, 1)) <> 46) Then 'το asc('.') είναι 46
    '                MsgBox "Λάθος τιμή ποσού!", vbCritical, "Μήνυμα λάθους"
    '                .SelStart = 0
    '                .SelLength = Len(.Text)
    '                .SetFocus
    '                i = 100
    '                Exit For
    '            End If
    '        Next i
    '        If i <> 100 Then
    '            '.Text = .Text & " €"
    '        End If
    '    End If
    'End With
    
    On Error GoTo f12_LostFocus_l3
    
    With f12
        If .Text <> "" Then
            Dim per As Double
            'per = CDbl(.Text)
            If InStr(CStr(Val(.Text)), ",") > 0 Then
                per = Val(.Text)
            Else
                per = CDbl(.Text)
            End If
            GoTo f12_LostFocus_l4
        End If
    End With

f12_LostFocus_l3:
            MsgBox "Λάθος τιμή ποσού!", vbCritical, "Μήνυμα λάθους"
            f12.SelStart = 0
            f12.SelLength = Len(f12.Text)
            f12.SetFocus
            GoTo f12_LostFocus_l5
            
f12_LostFocus_l4:
    f12.Text = per
    
f12_LostFocus_l5:
    
End Sub

Private Sub f15_GotFocus()


    On Error GoTo f15_GotFocus_l1

    With f15
        If .Text <> "" Then
            '.Text = Val(.Text)
            .Text = CDbl(.Text)
        End If
    End With
    
f15_GotFocus_l1:
    

End Sub

Private Sub f15_LostFocus()
    
    Dim s As String
    Dim i As Integer
    
'    With f15
'        If .Text <> "" Then
'            s = .Text
'            For i = 1 To Len(s)
'                If (Mid$(s, i, 1) <> "0" And Mid$(s, i, 1) <> "1" And Mid$(s, i, 1) <> "2" And Mid$(s, i, 1) <> "3" And Mid$(s, i, 1) <> "4" And Mid$(s, i, 1) <> "5" And Mid$(s, i, 1) <> "6" And Mid$(s, i, 1) <> "7" And Mid$(s, i, 1) <> "8" And Mid$(s, i, 1) <> "9" And Mid$(s, i, 1) <> "," And Mid$(s, i, 1) <> ".") Then
'                    MsgBox "Λάθος τιμή ποσού!", vbCritical, "Μήνυμα λάθους"
'                    .SelStart = 0
'                    .SelLength = Len(.Text)
'                    .SetFocus
'                    i = 100
'                    Exit For
'                End If
'            Next i
'            If i <> 100 Then
'                '.Text = .Text & " €"
'            End If
'        End If
'    End With

On Error GoTo f15_LostFocus_l1
    
    With f15
        If .Text <> "" Then
            Dim per As Double
            If InStr(CStr(Val(.Text)), ",") > 0 Then
                per = Val(.Text)
            Else
                per = CDbl(.Text)
            End If
            GoTo f15_LostFocus_l2
        End If
    End With

f15_LostFocus_l1:
            MsgBox "Λάθος τιμή ποσού!", vbCritical, "Μήνυμα λάθους"
            f15.SelStart = 0
            f15.SelLength = Len(f15.Text)
            f15.SetFocus
            GoTo f15_LostFocus_l3
            
f15_LostFocus_l2:
    f15.Text = per
    
f15_LostFocus_l3:
        
End Sub

Private Sub f3_GotFocus()
    
    With f3
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub f3_LostFocus()
        
        'If IsDate(.Text) = False And (.Text <> "__/__/____") And (.Text <> "  /  /    ") Then
    With f3
        If IsDate(.Text) = False Then
            MsgBox "Λάθος τιμή ημερομηνίας!", vbCritical, "Μήνυμα λάθους"
            .SelStart = 0
            .SelLength = 10
            .SetFocus
        End If
    End With

End Sub

Private Sub f8_GotFocus()
    
    With f8
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub f9_GotFocus()
    
    With f9
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub f9_LostFocus()
        
    With f9
        If IsDate(.Text) = False And .Text <> "" Then
            MsgBox "Λάθος τιμή ημερομηνίας!", vbCritical, "Μήνυμα λάθους"
            .SelStart = 0
            .SelLength = 10
            .SetFocus
        End If
    End With

End Sub

Private Sub kl_bt_Click()

    Unload Me

End Sub

Private Sub Form_Deactivate()


    Me.SetFocus

End Sub

Private Sub Form_Load()

    kl = 0

End Sub

Private Sub Form_LostFocus()

    'Call Command9_Click

End Sub

Private Sub Form_Unload(cancel As Integer)

    'If kl = 0 Then
    '    ms = MsgBox("Έχετε περάσει τις αλλαγές στη Β.Δ. της Εφαρμογής Διαχείρισης Ποσειδώνα;", vbYesNo, "Μήνυμα Προειδοποίησης")
    '    If ms = vbYes Then
    '        Unload Me
    '    Else
    '        cancel = 1
    '    End If
    'End If


    If Me.Storage.Enabled = True Or kl = 0 Then
        'ms = MsgBox("Θέλετε να αποθηκεύσετε τις αλλαγές στη Β.Δ. της Εφαρμογής Διαχείρισης Ποσειδώνα;", vbYesNo, "Μήνυμα Προειδοποίησης")
        ms = MsgBox("Είσαι σίγουρος για το κλείσιμο;", vbYesNo, "Μήνυμα Προειδοποίησης")
        If ms = vbYes Then
            kl = 1
            Unload Me
        Else
            cancel = 1
        End If
    Else
        kl = 1
        Unload Me
    End If

End Sub

Private Sub opt_f1_Click(Index As Integer)

    
    If oikonomikes_kiniseis_management.co_py.Text <> "" Then
        co_py.Text = oikonomikes_kiniseis_management.co_py.Text
        'co_py.Locked = True
    Else
        'co_py.Locked = False
    End If
    If Index = 0 And opt_f1(Index).Value = True Then 'ΕΙΝΑΙ ΠΙΣΤΩΣΗ
        '
        Call Refresh_ΤύποιΠαραστατικών(True, 1)
        Me.f8.Locked = True
        Me.f8.Text = ""
        Me.f8.Enabled = True
        Me.f8.BackColor = vbWhite
        '
        Me.Esoda.Enabled = True
        Me.Esoda.BackColor = &H8000&
        Call Refresh_ΤύποιΕσόδων(1)
        Me.Eksoda.BackColor = &HC0C0C0
        Me.co_raw_tipoi_eksodwn.Clear
        '
        Me.f12.Text = ""
        Me.f9.Locked = True
        Me.f9.Text = DateValue(Now)
    Else
        If Index = 1 And opt_f1(Index).Value = True Then 'ΕΙΝΑΙ ΧΡΕΩΣΗ
            '
            Call Refresh_ΤύποιΠαραστατικών(False, 1)
            Me.f8.Text = ""
            Me.f8.Locked = False
            Me.f9.Locked = False
            'Me.f8.Enabled = False
            'Me.f8.BackColor = &HC0C0C0
            '
            Me.Esoda.Enabled = False
            Me.Esoda.BackColor = &HC0C0C0
            Me.co_raw_tipoi_esodwn.Clear
            Me.Eksoda.Enabled = True
            Me.Eksoda.BackColor = &H800000
            Call Refresh_ΤύποιΕξόδων(1)
            '
            Me.f15.Text = ""
        Else
            Call Refresh_ΤύποιΠαραστατικών(False, 0) 'ΔΕΒ ΓΙΝΕΤΑΙ ΦΙΛΤΡΑΡΙΣΜΑ
        End If
    End If
    
End Sub

Private Sub raw_co_tipoi_parastatikwn_Click()
    
    'If Me.opt_f1(0).Value = True Then 'ΕΙΝΑΙ ΠΙΣΤΩΣΗ
    ar = ΕνημέρωσηΑριθμούΠαραστατικούΕσόδου
    If ar >= 1 Then
        Me.f8.Text = ΕνημέρωσηΑριθμούΠαραστατικούΕσόδου
    Else
        Me.f8.Text = ""
    End If
    'End If

End Sub

Private Sub Storage_Click()
    
    Dim id_par As Integer
    Dim ms As String

    ms = MsgBox("Είσαι σίγουρος; Μετά την αποθήκευση δεν επιτρέπεται η διαγραφή παρά μόνο η ακύρωση...(ΝΑΙ ή ΟΧΙ)", vbYesNo, "Παράθυρο επιβεβαίωσης")
    If ms = 6 Then
        
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.AddNew
        'id
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(0).Value = Val(Me.f0.Text)
        'Προϋπολογισμός
        If Me.co_py.SelectedItem <> Empty Then
            id_s = ΕύρεσηID_από_String("Προϋπολογισμός", Me.co_py.Text)
            oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(1).Value = id_s
        End If
        'ΤύποςΚίνησης
        If Me.opt_f1(0).Value = True Then 'EINAI ΕΣΟΔΟ
            oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(2).Value = 1
        Else 'ΕΝΑΙ ΕΞΟΔΟ
            oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(2).Value = 0
        End If
        'ΚατάστασηΚίνησης
        If Me.opt_f2(0).Value = True Then 'EINAI ΑΚΥΡΗ
            oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(3).Value = -1
        Else
            If Me.opt_f2(1).Value = True Then 'EINAI ΔΕΣΜΕΥΜΕΝΗ
                oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(3).Value = 0
            Else 'ΕΙΝΑΙ ΕΝΕΡΓΗ
                oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(3).Value = 1
            End If
        End If
        'ΗΜΕΡΟΜΗΝΙΑ ΚΙΝΗΣΗΣ
        If Me.f3.Text <> "" Then
            oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(4).Value = Me.f3.Text
        End If
        'ΑΘΛΗΤΗΣ
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(5).Value = ΕνημέρωσηΑθλητή
        'ΜΕΛΟΣ
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(6).Value = ΕνημέρωσηΜέλους
        'ΟΡΓΑΝΙΣΜΟΣ ή ΠΡΟΜΗΘΕΥΤΗΣ
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(7).Value = ΕνημέρωσηΟργανισμού
        'ΤΥΠΟΣ ΠΑΡΑΣΤΑΤΙΚΟΎ
        id_par = ΕνημέρωσηΠαραστατικού
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(8).Value = id_par
        'ΑΡΙΘΜΟΣ ΠΑΡΑΣΤΑΤΙΚΟΎ
        If Me.f8.Text <> "" Then
            oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(9).Value = Me.f8.Text
            Call ΕνημέρωσηΤρέχοντοςΑριθμούΠαραστατικούΕσόδου(id_par, Val(Me.f8.Text))
        End If
        'ΗμερομηνίαΠαραστατικού
        If Me.f9.Text <> "" Then
            oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(10).Value = Me.f9.Text
        End If
        'ΤύποςΠΥΕξόδων
        id_tip_eks = ΕνημέρωσηΤύπουΕξόδου
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(11).Value = id_tip_eks
        'ΠοσόΧρέωσης
        If Me.f12.Text <> "" Then
            oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(12).Value = CDbl(Me.f12.Text)
        End If
        'ΤύποςΠΥΕσόδων
        id_tip_es = ΕνημέρωσηΤύπουΕσόδου
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(13).Value = id_tip_es
        'Μήνες Συνδρομής Αθλητή
        For i = 0 To Me.lst_mn.ListCount - 1
            If Me.lst_mn.Selected(i) = True Then
                oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(16 + i).Value = True
            Else
                oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(16 + i).Value = False
            End If
        Next i
        'ΠοσόΠίστωσης
        If Me.f15.Text <> "" Then
            oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(14).Value = CDbl(Me.f15.Text)
        End If
        'Αιτιολογία
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(15).Value = Me.f16.Text
        'Αιτιολογία2
        'MsgBox "mikos = " & Len(Me.txt_aitiol2.Text)
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields("Αιτιολογία2").Value = Me.txt_aitiol2.Text
    
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.UpdateBatch adAffectCurrent
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Requery
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Refresh
        'xxxx
        oikonomikes_kiniseis_management.tmp_dt_oikon_kiniseis.Refresh
    
        oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Requery
        oikonomikes_kiniseis_management.ado_oikon_kiniseis.Refresh
        'xxxx
        oikonomikes_kiniseis_management.dt_oikon_kiniseis.Refresh
        oikonomikes_kiniseis_management.tmp2_ado_oik_kiniseis.Recordset.Requery
        oikonomikes_kiniseis_management.tmp2_ado_oik_kiniseis.Refresh
        'xxxx
        oikonomikes_kiniseis_management.tmp2_dt_oikon_kiniseis.Refresh
        oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Filter = MDIForm1.s_string
        oikonomikes_kiniseis_management.tmp2_ado_oik_kiniseis.Recordset.Filter = MDIForm1.s_string
        oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Sort = "[id]"
        oikonomikes_kiniseis_management.tmp2_ado_oik_kiniseis.Recordset.Sort = "[id]"
        MDIForm1.s_sort = "[id]"
    
        Call OikonomikesKiniseisRefresh
        oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.MoveLast
        oikonomikes_kiniseis_management.Command4.Enabled = True
        Me.Command5.Enabled = True
        
        
        'ON LINE ενημέρωση του Π/Υ
        Dim sin_p, sin_d As String
        If Me.opt_f1(0).Value = True Then 'EINAI ΕΣΟΔΟ
            'ON LINE ενημέρωση του Π/Υ ΕΣΟΔΩΝ
            '1. ΣΤΟ STARAGE A NEW RECORD, ΔΕΝ ΠΡΟΣΑΡΜΟΖΟΝΤΑΙ ΠΑΛΙΕΣ ΤΙΜΕΣ, ΔΙΟΤΙ ΔΕΝ ΥΠΑΡΧΟΥΝ
            '
            '2. ΠΡΟΣΑΡΜΟΖΩ ΤΙΣ ΝΕΕΣ ΤΙΜΕΣ ΣΤΑ ΕΣΟΔΑ
            Me.ado_anal_py_es.Recordset.Filter = "[id_προϋπολογισμού] = " & id_s & " AND [id_τύπου_εσόδου] = " & id_tip_es
            If Me.ado_anal_py_es.Recordset.RecordCount >= 1 Then
                If Me.opt_f2(0).Value = True Then 'EINAI ΑΚΥΡΗ
                    'NOTHING HERE
                Else
                    If Me.opt_f2(1).Value = True Then  'EINAI ΔΕΣΜΕΥΜΕΝΗ
                        sin_d = Me.ado_anal_py_es.Recordset.Fields(5).Value
                        'sin_d = sin_d + Val(Me.f15.Text)
                        sin_d = sin_d + CDbl(Me.f15.Text)
                        Me.ado_anal_py_es.Recordset.Fields(5).Value = sin_d
                    Else 'ΕΙΝΑΙ ΕΝΕΡΓΗ
                        sin_p = Me.ado_anal_py_es.Recordset.Fields(4).Value
                        'sin_p = sin_p + Val(Me.f15.Text)
                        If Me.f15.Text <> "" Then
                            sin_p = sin_p + CDbl(Me.f15.Text)
                        End If
                        Me.ado_anal_py_es.Recordset.Fields(4).Value = sin_p
                    End If
                End If
                Me.ado_anal_py_es.Recordset.UpdateBatch adAffectCurrent
                Me.ado_anal_py_es.Recordset.Requery
                Me.ado_anal_py_es.Refresh
            End If
            '
        Else 'ΕΝΑΙ ΕΞΟΔΟ
            '
            'ON LINE ενημέρωση του Π/Υ ΕΞΟΔΩΝ
            '1. ΣΤΟ STARAGE A NEW RECORD, ΔΕΝ ΠΡΟΣΑΡΜΟΖΟΝΤΑΙ ΠΑΛΙΕΣ ΤΙΜΕΣ, ΔΙΟΤΙ ΔΕΝ ΥΠΑΡΧΟΥΝ
            '
            '2. ΠΡΟΣΑΡΜΟΖΩ ΤΙΣ ΝΕΕΣ ΤΙΜΕΣ ΣΤΑ ΕΞΟΔΑ
            Me.ado_anal_py_eks.Recordset.Filter = "[id_προϋπολογισμού] = " & id_s & " AND [id_τύπου_εξόδου] = " & id_tip_eks
            If Me.ado_anal_py_eks.Recordset.RecordCount >= 1 Then
                If Me.opt_f2(0).Value = True Then 'EINAI ΑΚΥΡΗ
                    'NOTHING HERE
                Else
                    If Me.opt_f2(1).Value = True Then  'EINAI ΔΕΣΜΕΥΜΕΝΗ
                        sin_d = Me.ado_anal_py_eks.Recordset.Fields(5).Value
                        'sin_d = sin_d + Val(Me.f12.Text)
                        sin_d = sin_d + CDbl(Me.f12.Text)
                        Me.ado_anal_py_eks.Recordset.Fields(5).Value = sin_d
                    Else 'ΕΙΝΑΙ ΕΝΕΡΓΗ
                        sin_p = Me.ado_anal_py_eks.Recordset.Fields(4).Value
                        'sin_p = sin_p + Val(Me.f12.Text)
                        sin_p = sin_p + CDbl(Me.f12.Text)
                        Me.ado_anal_py_eks.Recordset.Fields(4).Value = sin_p
                    End If
                End If
                Me.ado_anal_py_eks.Recordset.UpdateBatch adAffectCurrent
                Me.ado_anal_py_eks.Recordset.Requery
                Me.ado_anal_py_eks.Refresh
            End If
            '
        End If
        
        'ΤΕΛΙΚΗ ΕΝΕΡΓΕΙΑ ΤΟ ΚΛΕΙΣΙΜΟ ΤΗΣ ΦΟΡΜΑΣ
        'Unload MIA_oikonomiki_kinisi_management
        Me.Storage.Enabled = False
        Me.Clean.Enabled = False
        Me.update.Enabled = True
        'Me.Command5.Enabled = False
    '
    Else
        MsgBox "Ακύρωση Αποθήκευσης ΝΕΑΣ Οικονομικής Κίνησης.", , "Μήνυμα Προειδοποίησης"
    End If

End Sub

Private Sub update_Click()
    
    Dim id_par As Integer

    oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.MoveFirst
    oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Find "[id] = " & Val(Me.f0.Text)
    'id
    '
    'ΠΡΟΥΠΟΛΟΓΙΣΜΟΣ
    If Me.co_py.SelectedItem <> Empty Then
        id_s = ΕύρεσηID_από_String("Προϋπολογισμός", Me.co_py.Text)
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(1).Value = id_s
    End If
    '
    'ΤύποςΚίνησης
    If Me.opt_f1(0).Value = True Then 'EINAI ΕΣΟΔΟ
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(2).Value = 1
    Else 'ΕΝΑΙ ΕΞΟΔΟ
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(2).Value = 0
    End If
    'ΚατάστασηΚίνησης
    If Me.opt_f2(0).Value = True Then 'EINAI ΑΚΥΡΗ
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(3).Value = -1
    Else
        If Me.opt_f2(1).Value = True Then 'EINAI ΔΕΣΜΕΥΜΕΝΗ
            oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(3).Value = 0
        Else 'ΕΙΝΑΙ ΕΝΕΡΓΗ
            oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(3).Value = 1
        End If
    End If
    ' ΗΜΕΡΟΜΗΝΙΑ ΚΙΝΗΣΗΣ
    If Me.f3.Text <> "" Then
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(4).Value = Me.f3.Text
    End If
    ' ΑΘΛΗΤΗΣ
    oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(5).Value = ΕνημέρωσηΑθλητή
    ' ΜΕΛΟΣ
    oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(6).Value = ΕνημέρωσηΜέλους
    ' ΟΡΓΑΝΙΣΜΟΣ
    oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(7).Value = ΕνημέρωσηΟργανισμού
    ' ΤΥΠΟΣ ΠΑΡΑΣΤΑΤΙΚΟΎ
    id_par = ΕνημέρωσηΠαραστατικού
    If id_par <> -1 Then
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(8).Value = id_par
    End If
    ' ΑΡΙΘΜΟΣ ΠΑΡΑΣΤΑΤΙΚΟΎ
    oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(9).Value = Val(Me.f8.Text)
    Call ΕνημέρωσηΤρέχοντοςΑριθμούΠαραστατικούΕσόδου(id_par, Val(Me.f8.Text))
    'ΗμερομηνίαΠαραστατικού
    If Me.f9.Text <> "" Then
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(10).Value = Me.f9.Text
    End If
    'ΤύποςΠΥΕξόδων
    id_py_eks = ΕνημέρωσηΤύπουΕξόδου
    If id_py_eks <> -1 Then
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(11).Value = ΕνημέρωσηΤύπουΕξόδου
    Else
        id_py_eks = oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(11).Value
    End If
    'ΠοσόΧρέωσης
    If Me.f12.Text <> "" Then
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(12).Value = CDbl(Me.f12.Text)
    End If
    'ΤύποςΠΥΕσόδων
    id_py_es = ΕνημέρωσηΤύπουΕσόδου
    If id_py_es <> -1 Then
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(13).Value = ΕνημέρωσηΤύπουΕσόδου
    Else
        id_py_es = oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(13).Value
    End If
    For i = 0 To Me.lst_mn.ListCount - 1
        If Me.lst_mn.Selected(i) = True Then
            oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(16 + i).Value = True
        Else
            oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(16 + i).Value = False
        End If
    Next i
    'ΠοσόΠίστωσης
    If Me.f15.Text <> "" Then
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(14).Value = CDbl(Me.f15.Text)
    End If
    'Αιτιολογία
    oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(15).Value = Me.f16.Text
    'Αιτιολογία2
    oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields("Αιτιολογία2").Value = Me.txt_aitiol2.Text
    
    oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.UpdateBatch adAffectCurrent
    oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Requery
    oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Refresh
    
    cur_row = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.AbsolutePosition
    oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Requery
    oikonomikes_kiniseis_management.ado_oikon_kiniseis.Refresh
    oikonomikes_kiniseis_management.tmp2_ado_oik_kiniseis.Recordset.Requery
    oikonomikes_kiniseis_management.tmp2_ado_oik_kiniseis.Refresh
    oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Filter = MDIForm1.s_string
    oikonomikes_kiniseis_management.tmp2_ado_oik_kiniseis.Recordset.Filter = MDIForm1.s_string
    oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Sort = MDIForm1.s_sort
    oikonomikes_kiniseis_management.tmp2_ado_oik_kiniseis.Recordset.Sort = MDIForm1.s_sort
    Call OikonomikesKiniseisRefresh
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.RecordCount >= 1 Then
        oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.MoveFirst
        oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Find "[id] = " & Val(Me.f0.Text)
    Else
        oikonomikes_kiniseis_management.ado_oikon_kiniseis.Caption = "Παραστατικό 0 από 0"
    End If
    
    'ON LINE ενημέρωση του Π/Υ
    Dim sin_p, sin_d As String
    If Me.opt_f1(0).Value = True Then 'EINAI ΕΣΟΔΟ
        'ON LINE ενημέρωση του Π/Υ ΣΤΑ ΕΣΟΔΑ
        '1. ΠΡΟΣΑΡΜΟΖΩ ΤΙΣ ΠΑΛΙΕΣ ΤΙΜΕΣ
        Me.ado_anal_py_es.Recordset.Filter = "[id_προϋπολογισμού] = " & oikonomikes_kiniseis_management.global_py & " AND [id_τύπου_εσόδου] = " & oikonomikes_kiniseis_management.global_tip_es
        If Me.ado_anal_py_es.Recordset.RecordCount >= 1 Then
            If oikonomikes_kiniseis_management.global_kk = -1 Then 'EINAI ΑΚΥΡΗ
                'ccccc
            Else
                If oikonomikes_kiniseis_management.global_kk = 0 Then 'EINAI ΔΕΣΜΕΥΜΕΝΗ
                    sin_d = Me.ado_anal_py_es.Recordset.Fields(5).Value
                    sin_d = sin_d - oikonomikes_kiniseis_management.global_poso_pistosis
                    Me.ado_anal_py_es.Recordset.Fields(5).Value = sin_d
                Else 'ΕΙΝΑΙ ΕΝΕΡΓΗ
                    sin_p = Me.ado_anal_py_es.Recordset.Fields(4).Value
                    sin_p = sin_p - oikonomikes_kiniseis_management.global_poso_pistosis
                    Me.ado_anal_py_es.Recordset.Fields(4).Value = sin_p
                End If
            End If
            Me.ado_anal_py_es.Recordset.UpdateBatch adAffectCurrent
            Me.ado_anal_py_es.Recordset.Requery
            Me.ado_anal_py_es.Refresh
        End If
        '2. ΠΡΟΣΑΡΜΟΖΩ ΤΙΣ ΝΕΕΣ ΤΙΜΕΣ
        Me.ado_anal_py_es.Recordset.Filter = "[id_προϋπολογισμού] = " & id_s & " AND [id_τύπου_εσόδου] = " & id_py_es
        If Me.ado_anal_py_es.Recordset.RecordCount >= 1 Then
            If Me.opt_f2(0).Value = True Then 'EINAI ΑΚΥΡΗ
                'bbb
            Else
                If Me.opt_f2(1).Value = True Then  'EINAI ΔΕΣΜΕΥΜΕΝΗ
                    sin_d = Me.ado_anal_py_es.Recordset.Fields(5).Value
                    'sin_d = sin_d + Val(Me.f15.Text)
                    sin_d = sin_d + CDbl(Me.f15.Text)
                    Me.ado_anal_py_es.Recordset.Fields(5).Value = sin_d
                Else 'ΕΙΝΑΙ ΕΝΕΡΓΗ
                    sin_p = Me.ado_anal_py_es.Recordset.Fields(4).Value
                    'sin_p = sin_p + Val(Me.f15.Text)
                    If Me.f15.Text <> "" Then
                        sin_p = sin_p + CDbl(Me.f15.Text)
                    End If
                    Me.ado_anal_py_es.Recordset.Fields(4).Value = sin_p
                End If
            End If
            Me.ado_anal_py_es.Recordset.UpdateBatch adAffectCurrent
            Me.ado_anal_py_es.Recordset.Requery
            Me.ado_anal_py_es.Refresh
        End If
    'ΤΕΛΕΙΩΝΕΙ Η ΕΠΕΞΕΡΓΑΣΙΑ ΜΕ έσοδο
    Else 'ΕΙΝΑΙ ΕΞΟΔΟ
        'ON LINE ενημέρωση του Π/Υ ΣΤΑ ΕΞΟΔΑ
        '1. ΠΡΟΣΑΡΜΟΖΩ ΤΙΣ ΠΑΛΙΕΣ ΤΙΜΕΣ
        Me.ado_anal_py_eks.Recordset.Filter = "[id_προϋπολογισμού] = " & oikonomikes_kiniseis_management.global_py & " AND [id_τύπου_εξόδου] = " & oikonomikes_kiniseis_management.global_tip_eks
        If Me.ado_anal_py_eks.Recordset.RecordCount >= 1 Then
            If oikonomikes_kiniseis_management.global_kk = -1 Then 'EINAI ΑΚΥΡΗ
                'ccccc
            Else
                If oikonomikes_kiniseis_management.global_kk = 0 Then 'EINAI ΔΕΣΜΕΥΜΕΝΗ
                    sin_d = Me.ado_anal_py_eks.Recordset.Fields(5).Value
                    sin_d = sin_d - oikonomikes_kiniseis_management.global_poso_xreosis
                    Me.ado_anal_py_eks.Recordset.Fields(5).Value = sin_d
                Else 'ΕΙΝΑΙ ΕΝΕΡΓΗ
                    sin_p = Me.ado_anal_py_eks.Recordset.Fields(4).Value
                    sin_p = sin_p - oikonomikes_kiniseis_management.global_poso_xreosis
                    Me.ado_anal_py_eks.Recordset.Fields(4).Value = sin_p
                End If
            End If
            Me.ado_anal_py_eks.Recordset.UpdateBatch adAffectCurrent
            Me.ado_anal_py_eks.Recordset.Requery
            Me.ado_anal_py_eks.Refresh
        End If
        '2. ΠΡΟΣΑΡΜΟΖΩ ΤΙΣ ΝΕΕΣ ΤΙΜΕΣ
        Me.ado_anal_py_eks.Recordset.Filter = "[id_προϋπολογισμού] = " & id_s & " AND [id_τύπου_εξόδου] = " & id_py_eks
        If Me.ado_anal_py_eks.Recordset.RecordCount >= 1 Then
            If Me.opt_f2(0).Value = True Then 'EINAI ΑΚΥΡΗ
                'bbb
            Else
                If Me.opt_f2(1).Value = True Then  'EINAI ΔΕΣΜΕΥΜΕΝΗ
                    sin_d = Me.ado_anal_py_eks.Recordset.Fields(5).Value
                    'sin_d = sin_d + Val(Me.f12.Text)
                    sin_d = sin_d + CDbl(Me.f12.Text)
                    Me.ado_anal_py_eks.Recordset.Fields(5).Value = sin_d
                Else 'ΕΙΝΑΙ ΕΝΕΡΓΗ
                    sin_p = Me.ado_anal_py_eks.Recordset.Fields(4).Value
                    'sin_p = sin_p + Val(Me.f12.Text)
                    sin_p = sin_p + CDbl(Me.f12.Text)
                    Me.ado_anal_py_eks.Recordset.Fields(4).Value = sin_p
                End If
            End If
            Me.ado_anal_py_eks.Recordset.UpdateBatch adAffectCurrent
            Me.ado_anal_py_eks.Recordset.Requery
            Me.ado_anal_py_eks.Refresh
        End If
        'ΤΕΛΕΙΩΝΕΙ Η ΕΠΕΞΕΡΓΑΣΙΑ ΜΕ έξοδο
    End If

    'ΤΕΛΙΚΗ ΕΝΕΡΓΕΙΑ ΤΟ ΚΛΕΙΣΙΜΟ ΤΗΣ ΦΟΡΜΑΣ
    'Unload MIA_oikonomiki_kinisi_management

End Sub
