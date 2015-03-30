VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form promitheytes_management 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Διαχείριση Προμηθευτών - Οργανισμών"
   ClientHeight    =   9930
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11775
   ForeColor       =   &H000000FF&
   LinkTopic       =   "entity_management"
   MDIChild        =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   11775
   Begin MSDataGridLib.DataGrid dt_dimoi 
      Bindings        =   "promitheytes_management.frx":0000
      Height          =   375
      Left            =   9600
      TabIndex        =   38
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin MSDataGridLib.DataGrid dt_doi 
      Bindings        =   "promitheytes_management.frx":0018
      Height          =   375
      Left            =   9600
      TabIndex        =   37
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin MSAdodcLib.Adodc ado_dimoi 
      Height          =   375
      Left            =   9600
      Top             =   1440
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
      RecordSource    =   "Δήμοι"
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
      Top             =   2400
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
      RecordSource    =   "ΠεριφερειακέςΕνότητες"
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
   Begin MSAdodcLib.Adodc ado_doi 
      Height          =   375
      Left            =   9600
      Top             =   480
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
      RecordSource    =   "ΔΟΥ"
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
      Caption         =   "Όλοι οι Προμηθευτές - Οργανισμοί"
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
      Height          =   5055
      Left            =   120
      TabIndex        =   14
      Top             =   4800
      Width           =   11535
      Begin VB.CommandButton taksin 
         BackColor       =   &H80000014&
         Caption         =   "Ταξινόμηση"
         DisabledPicture =   "promitheytes_management.frx":002E
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
         Picture         =   "promitheytes_management.frx":4D35
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4320
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dt_pr 
         Bindings        =   "promitheytes_management.frx":9A3C
         Height          =   3495
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6165
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
      Begin MSAdodcLib.Adodc ado_pr 
         Height          =   375
         Left            =   120
         Top             =   3840
         Width           =   11325
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
         RecordSource    =   "Προμηθευτές"
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
      Caption         =   "Στοιχεία Προμηθευτή - Οργανισμού"
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
      Height          =   5175
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   11535
      Begin VB.CommandButton insert_bt 
         BackColor       =   &H80000014&
         Caption         =   "Προσθήκη"
         DisabledPicture =   "promitheytes_management.frx":9A51
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
         Picture         =   "promitheytes_management.frx":E7E0
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton save_command 
         BackColor       =   &H80000014&
         Caption         =   "Αποθήκευση"
         DisabledPicture =   "promitheytes_management.frx":1356F
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
         Left            =   2880
         MaskColor       =   &H80000014&
         Picture         =   "promitheytes_management.frx":1816E
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton canc_bt 
         BackColor       =   &H80000014&
         Caption         =   "Ακύρωση"
         DisabledPicture =   "promitheytes_management.frx":1CD6D
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
         Left            =   8640
         MaskColor       =   &H80000014&
         Picture         =   "promitheytes_management.frx":21C8A
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton kl_bt 
         BackColor       =   &H80000014&
         Caption         =   "Κλείσιμο"
         DisabledPicture =   "promitheytes_management.frx":26BA7
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
         Left            =   10080
         MaskColor       =   &H80000014&
         Picture         =   "promitheytes_management.frx":2C61F
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton up_bt 
         BackColor       =   &H80000014&
         Caption         =   "Ενημέρωση"
         DisabledPicture =   "promitheytes_management.frx":32097
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
         Left            =   1440
         MaskColor       =   &H80000014&
         Picture         =   "promitheytes_management.frx":39B91
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000014&
         Caption         =   "Καθαρισμός"
         DisabledPicture =   "promitheytes_management.frx":4168B
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
         Left            =   5760
         MaskColor       =   &H80000014&
         Picture         =   "promitheytes_management.frx":461CF
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton del_bt 
         BackColor       =   &H80000014&
         Caption         =   "Διαγραφή"
         DisabledPicture =   "promitheytes_management.frx":4AD13
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
         Left            =   4320
         MaskColor       =   &H80000014&
         Picture         =   "promitheytes_management.frx":4F9C3
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton sear_bt 
         BackColor       =   &H80000014&
         Caption         =   "Αναζήτηση"
         DisabledPicture =   "promitheytes_management.frx":54673
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
         Left            =   7200
         MaskColor       =   &H80000014&
         Picture         =   "promitheytes_management.frx":59A27
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3960
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo co_doi 
         Bindings        =   "promitheytes_management.frx":5EDDB
         Height          =   345
         Left            =   6240
         TabIndex        =   2
         Top             =   1080
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
         Height          =   2325
         Left            =   4800
         TabIndex        =   25
         Top             =   1560
         Width           =   4695
         Begin VB.TextBox txt_tk 
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
            TabIndex        =   10
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txt_perioxi 
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
            TabIndex        =   11
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txt_arithmos 
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
            TabIndex        =   9
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txt_odos 
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
            TabIndex        =   8
            Top             =   360
            Width           =   3015
         End
         Begin MSDataListLib.DataCombo co_pe 
            Bindings        =   "promitheytes_management.frx":5EDF1
            Height          =   345
            Left            =   1200
            TabIndex        =   13
            Top             =   1800
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
            Bindings        =   "promitheytes_management.frx":5EE06
            Height          =   345
            Left            =   1200
            TabIndex        =   12
            Top             =   1440
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
            TabIndex        =   31
            Top             =   360
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
            TabIndex        =   30
            Top             =   720
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
            TabIndex        =   29
            Top             =   1200
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
            TabIndex        =   28
            Top             =   1560
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
            TabIndex        =   27
            Top             =   1920
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
            TabIndex        =   26
            Top             =   720
            Width           =   1065
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
         Height          =   2325
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   4695
         Begin VB.TextBox txt_web 
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
            Left            =   1440
            TabIndex        =   7
            Top             =   1800
            Width           =   3015
         End
         Begin VB.TextBox txt_email 
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
            Left            =   1440
            TabIndex        =   6
            Top             =   1440
            Width           =   3015
         End
         Begin VB.TextBox txt_fax 
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
            TabIndex        =   5
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txt_kinito 
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
            TabIndex        =   4
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txt_til 
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
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Web Site"
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
            TabIndex        =   36
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Τηλέφωνο"
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
            TabIndex        =   24
            Top             =   360
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
            TabIndex        =   23
            Top             =   720
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
            TabIndex        =   22
            Top             =   1080
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
            TabIndex        =   21
            Top             =   1440
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
         Height          =   1290
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   9375
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
            TabIndex        =   32
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txt_ep 
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
            TabIndex        =   0
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txt_afm 
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
            Left            =   6120
            TabIndex        =   1
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Δ.Ο.Υ."
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
            Height          =   375
            Left            =   5040
            TabIndex        =   35
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Α.Φ.Μ."
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
            Left            =   4920
            TabIndex        =   34
            Top             =   360
            Width           =   1125
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
            TabIndex        =   33
            Top             =   435
            Width           =   1005
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Επωνυμία"
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
            TabIndex        =   19
            Top             =   735
            Width           =   1005
         End
      End
      Begin MSDataGridLib.DataGrid dt_pe 
         Bindings        =   "promitheytes_management.frx":5EE1E
         Height          =   375
         Left            =   9480
         TabIndex        =   39
         Top             =   2640
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         AllowUpdate     =   0   'False
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
End
Attribute VB_Name = "promitheytes_management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id_α As Integer
Public rs_ado_dimoi As ADODB.Recordset
Public rs_ado_jobs As ADODB.Recordset
Public rs_ado_meli As ADODB.Recordset
Public rs_ado_pe As ADODB.Recordset
Public for_search As Integer
Public s_sort As String
'
Public c_rec, defined_col  As Integer

Private Sub ado_pr_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    If Not adReason = 10 And pRecordset.AbsolutePosition >= 1 Then
        txt_kod.Text = ado_pr.Recordset.Fields(0).Value
        If Trim(ado_pr.Recordset.Fields(1).Value) <> "" Then
            txt_ep.Text = ado_pr.Recordset.Fields(1).Value
        Else
            txt_ep.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(2).Value) <> "" Then
            txt_odos.Text = ado_pr.Recordset.Fields(2).Value
        Else
            txt_odos.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(3).Value) <> "" Then
            txt_arithmos.Text = ado_pr.Recordset.Fields(3).Value
        Else
            txt_arithmos.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(4).Value) <> "" Then
            txt_perioxi.Text = ado_pr.Recordset.Fields(4).Value
        Else
            txt_perioxi.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(5).Value) <> "" Then
            If ado_dimoi.Recordset.RecordCount >= 1 Then
                ado_dimoi.Recordset.MoveFirst
                ado_dimoi.Recordset.Find "[id_δήμου] = '" & ado_pr.Recordset.Fields(5).Value & "'"
                If Not ado_dimoi.Recordset.EOF Then
                    co_dimoi.Text = ado_dimoi.Recordset.Fields(1).Value
                End If
            End If
        Else
            co_dimoi.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(6).Value) <> "" Then
            If ado_pe.Recordset.RecordCount >= 1 Then
                ado_pe.Recordset.MoveFirst
                ado_pe.Recordset.Find "[id_πε] = '" & ado_pr.Recordset.Fields(6).Value & "'"
                If Not ado_pe.Recordset.EOF Then
                    co_pe.Text = ado_pe.Recordset.Fields(1).Value
                End If
            End If
        Else
            co_pe.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(7).Value) <> "" Then
            txt_tk.Text = ado_pr.Recordset.Fields(7).Value
        Else
            txt_tk.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(8).Value) <> "" Then
            txt_til.Text = ado_pr.Recordset.Fields(8).Value
        Else
            txt_til.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(9).Value) <> "" Then
            txt_kinito.Text = ado_pr.Recordset.Fields(9).Value
        Else
            txt_kinito.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(10).Value) <> "" Then
            txt_fax.Text = ado_pr.Recordset.Fields(10).Value
        Else
            txt_fax.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(11).Value) <> "" Then
            txt_email.Text = ado_pr.Recordset.Fields(11).Value
        Else
            txt_email.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(12).Value) <> "" Then
            txt_web.Text = ado_pr.Recordset.Fields(12).Value
        Else
            txt_web.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(13).Value) <> "" Then
            txt_afm.Text = ado_pr.Recordset.Fields(13).Value
        Else
            txt_afm.Text = ""
        End If
        If ado_pr.Recordset.Fields(14).Value <> "" Then
            If ado_doi.Recordset.RecordCount >= 1 Then
                ado_doi.Recordset.MoveFirst
                ado_doi.Recordset.Find "[id_ΔΟΥ] = '" & ado_pr.Recordset.Fields(14).Value & "'"
                If Not ado_doi.Recordset.EOF Then
                    co_doi.Text = ado_doi.Recordset.Fields(1).Value
                Else
                    co_doi.Text = ""
                End If
            End If
        Else
            co_doi.Text = ""
        End If
    End If
    If pRecordset.RecordCount > 0 Then
        ado_pr.Caption = "Προμηθευτής " & pRecordset.AbsolutePosition & " από " & pRecordset.RecordCount
    Else
        ado_pr.Caption = "Προμηθευτής " & 0 & " από " & 0
    End If
    
End Sub

Private Sub canc_bt_Click()

    insert_bt.Enabled = True
    up_bt.Enabled = True
    s_sort = ""
    for_search = 0
    
    dt_pr.Refresh
    ado_pr.Refresh
    
    If ado_pr.Recordset.RecordCount >= 1 Then
        ado_pr.Recordset.Sort = "[id]"
        ado_pr.Recordset.MoveFirst
        ado_pr.Recordset.Move c_rec - 1
    End If
    '***************************************************************

    dt_pr.Columns(0).Caption = "Κωδικός"
    dt_pr.Columns(0).Width = 1500
    dt_pr.Columns(1).Caption = "Επωνυμία"
    dt_pr.Columns(1).Width = 4000
    For i = 2 To 12
        dt_pr.Columns(i).Visible = False
    Next i
    dt_pr.Columns(13).Caption = "Α.Φ.Μ."
    dt_pr.Columns(13).Width = 2000
    dt_pr.Columns(14).Visible = False
    
    txt_ep.SetFocus

End Sub

Private Sub co_jobs_Change()
    'strSQL = "SELECT Name FROM Table1 WHERE Category = '" & DataCombo1.BoundText & "'"
    'Set rsCombo2 = New ADODB.Recordset
    'rsCombo2.Open strSQL, adoConn, adOpenForwardOnly
    'rsCombo2.Sort = rsCombo2.Fields("Name").Name
    'Set DataCombo2.RowSource = rsCombo2
    'DataCombo2.ListField = rsCombo2.Fields("Name").Name
    'DataCombo2.BoundColumn = rsCombo2.Fields("Name").Name
    
    'rs_ado_jobs.Sort = "[" & Trim(rs_ado_jobs.Fields(1).Name) & "]"
    'co_jobs.ListField = rs_ado_jobs.Fields(1).Name
    'Me.co_jobs.BoundColumn = rs_ado_jobs.Fields(1).Name
    
End Sub

Private Sub Command1_Click()

        for_search = 1
        txt_kod.Locked = False
        txt_kod.Text = ""
        txt_ep.Text = ""
        txt_odos.Text = ""
        txt_arithmos.Text = ""
        txt_perioxi.Text = ""
        co_dimoi.Text = ""
        co_pe.Text = ""
        txt_tk.Text = ""
        txt_til.Text = ""
        txt_kinito.Text = ""
        txt_fax.Text = ""
        txt_email.Text = ""
        txt_web.Text = ""
        txt_afm.Text = ""
        co_doi.Text = ""
    
        insert_bt.Enabled = False
        up_bt.Enabled = False
        save_command.Enabled = False
        del_bt.Enabled = False
        
        Me.sear_bt.Enabled = True
        
End Sub

Private Sub dt_pr_HeadClick(ByVal ColIndex As Integer)

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
    
    If Not ado_pr.Recordset.EOF And ado_pr.Recordset.AbsolutePosition > 1 Then
    ms = MsgBox("Είσαι σίγουρος; (ΝΑΙ ή ΟΧΙ)", vbYesNo, "Παράθυρο διαγραφής")
    If ms = 6 Then
        ado_pr.Recordset.Delete
    Else
        MsgBox "ΑΚΥΡΩΣΗ ΔΙΑΓΡΑΦΗΣ!", vbCritical, "Μήνυμα Προειδοποίησης"
    End If
    Else
        MsgBox "Δεν υπάρχει εγγραφή προς ΔΙΑΓΡΑΦΗ!", vbCritical, "Μήνυμα Λάθους"
    End If
    
End Sub

Private Sub dt_pr_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If ado_pr.Recordset.AbsolutePosition >= 1 And ado_pr.Recordset.AbsolutePosition <= ado_pr.Recordset.RecordCount Then
        txt_kod.Text = ado_pr.Recordset.Fields(0).Value
        If Trim(ado_pr.Recordset.Fields(1).Value) <> "" Then
            txt_ep.Text = ado_pr.Recordset.Fields(1).Value
        Else
            txt_ep.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(2).Value) <> "" Then
            txt_odos.Text = ado_pr.Recordset.Fields(2).Value
        Else
            txt_odos.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(3).Value) <> "" Then
            txt_arithmos.Text = ado_pr.Recordset.Fields(3).Value
        Else
            txt_arithmos.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(4).Value) <> "" Then
            txt_perioxi.Text = ado_pr.Recordset.Fields(4).Value
        Else
            txt_perioxi.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(5).Value) <> "" Then
            If ado_dimoi.Recordset.RecordCount >= 1 Then
                ado_dimoi.Recordset.MoveFirst
                ado_dimoi.Recordset.Find "[id_δήμου] = '" & ado_pr.Recordset.Fields(5).Value & "'"
                If Not ado_dimoi.Recordset.EOF Then
                    co_dimoi.Text = ado_dimoi.Recordset.Fields(1).Value
                End If
            End If
        Else
            co_dimoi.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(6).Value) <> "" Then
            If ado_pe.Recordset.RecordCount >= 1 Then
                ado_pe.Recordset.MoveFirst
                ado_pe.Recordset.Find "[id_πε] = '" & ado_pr.Recordset.Fields(6).Value & "'"
                If Not ado_pe.Recordset.EOF Then
                    co_pe.Text = ado_pe.Recordset.Fields(1).Value
                End If
            End If
        Else
            co_pe.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(7).Value) <> "" Then
            txt_tk.Text = ado_pr.Recordset.Fields(7).Value
        Else
            txt_tk.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(8).Value) <> "" Then
            txt_til.Text = ado_pr.Recordset.Fields(8).Value
        Else
            txt_til.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(9).Value) <> "" Then
            txt_kinito.Text = ado_pr.Recordset.Fields(9).Value
        Else
            txt_kinito.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(10).Value) <> "" Then
            txt_fax.Text = ado_pr.Recordset.Fields(10).Value
        Else
            txt_fax.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(11).Value) <> "" Then
            txt_email.Text = ado_pr.Recordset.Fields(11).Value
        Else
            txt_email.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(12).Value) <> "" Then
            txt_web.Text = ado_pr.Recordset.Fields(12).Value
        Else
            txt_web.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(13).Value) <> "" Then
            txt_afm.Text = ado_pr.Recordset.Fields(13).Value
        Else
            txt_afm.Text = ""
        End If
        If ado_pr.Recordset.Fields(14).Value <> "" Then
            If ado_doi.Recordset.RecordCount >= 1 Then
                ado_doi.Recordset.MoveFirst
                ado_doi.Recordset.Find "[id_ΔΟΥ] = '" & ado_pr.Recordset.Fields(14).Value & "'"
                If Not ado_doi.Recordset.EOF Then
                    co_doi.Text = ado_doi.Recordset.Fields(1).Value
                Else
                    co_doi.Text = ""
                End If
            End If
        Else
            co_doi.Text = ""
        End If
    End If

End Sub

Private Sub Form_Load()
 
    for_search = 0
    s_sort = ""
    txt_kod.Locked = True
      
    Me.Height = 10500
    Me.Width = 11900
 
    ado_pr.Recordset.Sort = "[" & Trim(ado_pr.Recordset.Fields(0).Name) & "]"
    If Not ado_pr.Recordset.EOF Then
        ado_pr.Recordset.MoveFirst
        txt_kod.Text = ado_pr.Recordset.Fields(0).Value
        If Trim(ado_pr.Recordset.Fields(1).Value) <> "" Then
            txt_ep.Text = ado_pr.Recordset.Fields(1).Value
        Else
            txt_ep.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(2).Value) <> "" Then
            txt_odos.Text = ado_pr.Recordset.Fields(2).Value
        Else
            txt_odos.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(3).Value) <> "" Then
            txt_arithmos.Text = ado_pr.Recordset.Fields(3).Value
        Else
            txt_arithmos.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(4).Value) <> "" Then
            txt_perioxi.Text = ado_pr.Recordset.Fields(4).Value
        Else
            txt_perioxi.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(5).Value) <> "" Then
            If ado_dimoi.Recordset.RecordCount >= 1 Then
                ado_dimoi.Recordset.MoveFirst
                ado_dimoi.Recordset.Find "[id_δήμου] = '" & ado_pr.Recordset.Fields(5).Value & "'"
                If Not ado_dimoi.Recordset.EOF Then
                    co_dimoi.Text = ado_dimoi.Recordset.Fields(1).Value
                End If
            End If
        Else
            co_dimoi.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(6).Value) <> "" Then
            If ado_pe.Recordset.RecordCount >= 1 Then
                ado_pe.Recordset.MoveFirst
                ado_pe.Recordset.Find "[id_πε] = '" & ado_pr.Recordset.Fields(6).Value & "'"
                If Not ado_pe.Recordset.EOF Then
                    co_pe.Text = ado_pe.Recordset.Fields(1).Value
                End If
            End If
        Else
            co_pe.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(7).Value) <> "" Then
            txt_tk.Text = ado_pr.Recordset.Fields(7).Value
        Else
            txt_tk.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(8).Value) <> "" Then
            txt_til.Text = ado_pr.Recordset.Fields(8).Value
        Else
            txt_til.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(9).Value) <> "" Then
            txt_kinito.Text = ado_pr.Recordset.Fields(9).Value
        Else
            txt_kinito.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(10).Value) <> "" Then
            txt_fax.Text = ado_pr.Recordset.Fields(10).Value
        Else
            txt_fax.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(11).Value) <> "" Then
            txt_email.Text = ado_pr.Recordset.Fields(11).Value
        Else
            txt_email.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(12).Value) <> "" Then
            txt_web.Text = ado_pr.Recordset.Fields(12).Value
        Else
            txt_web.Text = ""
        End If
        If Trim(ado_pr.Recordset.Fields(13).Value) <> "" Then
            txt_afm.Text = ado_pr.Recordset.Fields(13).Value
        Else
            txt_afm.Text = ""
        End If
        If ado_pr.Recordset.Fields(14).Value <> "" Then
            If ado_doi.Recordset.RecordCount >= 1 Then
                ado_doi.Recordset.MoveFirst
                ado_doi.Recordset.Find "[id_ΔΟΥ] = '" & ado_pr.Recordset.Fields(14).Value & "'"
                If Not ado_doi.Recordset.EOF Then
                    co_doi.Text = ado_doi.Recordset.Fields(1).Value
                Else
                    co_doi.Text = ""
                End If
            End If
        Else
            co_doi.Text = ""
        End If
    Else
        txt_kod.Text = ""
        txt_ep.Text = ""
        txt_odos.Text = ""
        txt_arithmos.Text = ""
        txt_perioxi.Text = ""
        co_dimoi.Text = ""
        co_pe.Text = ""
        txt_tk.Text = ""
        txt_til.Text = ""
        txt_kinito.Text = ""
        txt_fax.Text = ""
        txt_email.Text = ""
        txt_web.Text = ""
        txt_afm.Text = ""
        co_doi.Text = ""
        txt_ep.Locked = True
        txt_odos.Locked = True
        txt_arithmos.Locked = True
        txt_perioxi.Locked = True
        co_dimoi.Locked = True
        co_pe.Locked = True
        txt_tk.Locked = True
        txt_til.Locked = True
        txt_kinito.Locked = True
        txt_fax.Locked = True
        txt_email.Locked = True
        txt_web.Locked = True
        txt_afm.Locked = True
        co_doi.Locked = True
        up_bt.Enabled = False
        save_command.Enabled = False
        del_bt.Enabled = False
        Command1.Enabled = False
        sear_bt.Enabled = False
        canc_bt.Enabled = False
    End If
    '
    If ado_pr.Recordset.RecordCount > 0 Then
        dt_pr.Row = 0
        dt_pr.Col = 1
        ado_pr.Caption = "Προμηθευτής 1" & " από " & ado_pr.Recordset.RecordCount
    Else
        ado_pr.Caption = "Προμηθευτής 0 από 0"
    End If
    '
    dt_pr.Columns(0).Caption = "Κωδικός"
    dt_pr.Columns(0).Width = 1500
    dt_pr.Columns(1).Caption = "Επωνυμία"
    dt_pr.Columns(1).Width = 4000
    For i = 2 To 12
        dt_pr.Columns(i).Visible = False
    Next i
    dt_pr.Columns(13).Caption = "Α.Φ.Μ."
    dt_pr.Columns(13).Width = 2000
    dt_pr.Columns(14).Visible = False
    
End Sub

Private Sub insert_bt_Click()

    'Να βρω το υποψήφιο id_προμηθευτή
    If ado_pr.Recordset.RecordCount >= 1 Then
        c_rec = ado_pr.Recordset.AbsolutePosition
        ado_pr.Recordset.Sort = "[" & Trim(ado_pr.Recordset.Fields(0).Name) & "]"
        ado_pr.Recordset.MoveLast
        id = ado_pr.Recordset.Fields(0).Value
        id = id + 1
    Else
        id = 1
    End If
    'Καθαρισμός πεδίων και ενεργοποίηση κουμπιών
    txt_kod.Text = id
    txt_ep.Text = ""
    txt_odos.Text = ""
    txt_arithmos.Text = ""
    txt_perioxi.Text = ""
    co_dimoi.Text = ""
    co_pe.Text = ""
    txt_tk.Text = ""
    txt_til.Text = ""
    txt_kinito.Text = ""
    txt_fax.Text = ""
    txt_email.Text = ""
    txt_web.Text = ""
    txt_afm.Text = ""
    co_doi.Text = ""
    txt_ep.Locked = False
    txt_odos.Locked = False
    txt_arithmos.Locked = False
    txt_perioxi.Locked = False
    co_dimoi.Locked = False
    co_pe.Locked = False
    txt_tk.Locked = False
    txt_til.Locked = False
    txt_kinito.Locked = False
    txt_fax.Locked = False
    txt_email.Locked = False
    txt_web.Locked = False
    txt_afm.Locked = False
    co_doi.Locked = False
    up_bt.Enabled = False
    save_command.Enabled = True
    del_bt.Enabled = False
    Command1.Enabled = False
    sear_bt.Enabled = False
    canc_bt.Enabled = True
    
    txt_ep.SetFocus
    
End Sub

Private Sub kl_bt_Click()

    Unload Me

End Sub

Private Sub save_command_Click()

    'Αποθήκευση στους προμηθευτές
    ado_pr.Recordset.AddNew
    ado_pr.Recordset.Fields(0).Value = txt_kod.Text
    'Αποθήκευση ΛΟΙΠΩΝ ΣΤΟΙΧΕΙΑ
    If txt_ep.Text <> "" Then
        ado_pr.Recordset.Fields(1).Value = txt_ep.Text
    End If
    If txt_odos.Text <> "" Then
        ado_pr.Recordset.Fields(2).Value = txt_odos.Text
    End If
    If txt_arithmos.Text <> "" Then
        ado_pr.Recordset.Fields(3).Value = txt_arithmos.Text
    End If
    If txt_perioxi.Text <> "" Then
    ado_pr.Recordset.Fields(4).Value = txt_perioxi.Text
    End If
    If co_dimoi.Text <> "" Then
        If co_dimoi.MatchedWithList = False Then
            'ΕΙΣΑΓΩΓΗ ΝΕΟΥ ΔΗΜΟΥ
            If ado_dimoi.Recordset.RecordCount >= 1 Then
                ado_dimoi.Recordset.Sort = "[" & Trim(ado_dimoi.Recordset.Fields(0).Name) & "]"
                ado_dimoi.Recordset.MoveLast
                id = ado_dimoi.Recordset.Fields(0).Value
            Else
                id = 0
            End If
            ado_dimoi.Recordset.AddNew
            ado_dimoi.Recordset.Fields(0) = id + 1
            ado_dimoi.Recordset.Fields(1) = co_dimoi.Text
            ado_dimoi.Recordset.UpdateBatch adAffectCurrent
            ado_pr.Recordset.Fields(5).Value = id + 1
        Else
            ado_dimoi.Recordset.MoveFirst
            ado_dimoi.Recordset.Find "[περιγραφή] LIKE '" & Trim(co_dimoi.Text) & "'"
            If Not ado_dimoi.Recordset.EOF Then
                ado_pr.Recordset.Fields(5).Value = ado_dimoi.Recordset.Fields(0).Value
            End If
        End If
    End If
    If co_pe.Text <> "" Then
        If co_pe.MatchedWithList = False Then
            'ΕΙΣΑΓΩΓΗ ΝΕΑΣ ΠΕΡΙΦΕΡΕΙΑΚΗΣ ΕΝΟΤΗΤΑΣ
            If ado_pe.Recordset.RecordCount >= 1 Then
                ado_pe.Recordset.Sort = "[" & Trim(ado_pe.Recordset.Fields(0).Name) & "]"
                ado_pe.Recordset.MoveLast
                id = ado_pe.Recordset.Fields(0).Value
            Else
                id = 0
            End If
            ado_pe.Recordset.AddNew
            ado_pe.Recordset.Fields(0) = id + 1
            ado_pe.Recordset.Fields(1) = co_pe.Text
            ado_pe.Recordset.UpdateBatch adAffectCurrent
            ado_pr.Recordset.Fields(6).Value = id + 1
        Else
            ado_pe.Recordset.MoveFirst
            ado_pe.Recordset.Find "[περιγραφή] LIKE '" & Trim(co_pe.Text) & "'"
            If Not ado_pe.Recordset.EOF Then
                ado_pr.Recordset.Fields(6).Value = ado_pe.Recordset.Fields(0).Value
            End If
        End If
    End If
    If txt_tk.Text <> "" Then
            ado_pr.Recordset.Fields(7).Value = txt_tk.Text
    End If
    If txt_til.Text <> "" Then
        ado_pr.Recordset.Fields(8).Value = txt_til.Text
    End If
    If txt_kinito.Text <> "" Then
        ado_pr.Recordset.Fields(9).Value = txt_kinito.Text
    End If
    If txt_fax.Text <> "" Then
        ado_pr.Recordset.Fields(10).Value = txt_fax.Text
    End If
    If txt_email.Text <> "" Then
        ado_pr.Recordset.Fields(11).Value = txt_email.Text
    End If
    If txt_web.Text <> "" Then
        ado_pr.Recordset.Fields(12).Value = txt_web.Text
    End If
    If txt_afm.Text <> "" Then
        ado_pr.Recordset.Fields(13).Value = txt_afm.Text
    End If
    If co_doi.Text <> "" Then
        If co_doi.MatchedWithList = False Then
            'ΕΙΣΑΓΩΓΗ ΝΕΑΣ ΔΟΥ
            If ado_doi.Recordset.RecordCount >= 1 Then
                ado_doi.Recordset.Sort = "[" & Trim(ado_doi.Recordset.Fields(0).Name) & "]"
                ado_doi.Recordset.MoveLast
                id = ado_doi.Recordset.Fields(0).Value
            Else
                id = 0
            End If
            ado_doi.Recordset.AddNew
            ado_doi.Recordset.Fields(0) = id + 1
            ado_doi.Recordset.Fields(1) = co_doi.Text
            ado_doi.Recordset.UpdateBatch adAffectCurrent
            ado_pr.Recordset.Fields(14).Value = id + 1
        Else
            ado_doi.Recordset.MoveFirst
            ado_doi.Recordset.Find "[περιγραφή] LIKE '" & Trim(co_doi.Text) & "'"
            If Not ado_doi.Recordset.EOF Then
                ado_pr.Recordset.Fields(14).Value = ado_doi.Recordset.Fields(0).Value
            End If
        End If
    End If
    '
    ado_pr.Recordset.UpdateBatch adAffectCurrent
    ado_pr.Recordset.Sort = "[" & Trim(ado_pr.Recordset.Fields(0).Name) & "]"
    ado_pr.Recordset.MoveLast
    
    save_command.Enabled = False
    Me.canc_bt.Enabled = False
    Me.insert_bt.Enabled = True
    Me.up_bt.Enabled = True
    
End Sub

Private Sub sear_bt_Click()

    s_string = ""
    'ΚΡΙΤΗΡΙΟ ΚΩΔΙΚΟΥ
    If Trim(txt_kod.Text) <> "" Then
        s_string = "[id] LIKE " & Trim(txt_kod.Text)
    End If
    'ΚΡΙΤΗΡΙΟ ΕΠΩΝΥΜΙΑΣ
    If Trim(txt_ep.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Επωνυμία] LIKE '*" & Trim(txt_ep.Text) & "*'"
        Else
            s_string = "[Επωνυμία] LIKE '*" & Trim(txt_ep.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΟΔΟΥ
    If Trim(txt_odos.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Οδός] LIKE '*" & Trim(txt_odos.Text) & "*'"
        Else
            s_string = "[Οδός] LIKE '*" & Trim(txt_odos.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΑΡΙΘΜΟΥ ΟΔΟΥ
    If Trim(txt_arithmos.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Αριθμός] LIKE '*" & Trim(txt_arithmos.Text) & "*'"
        Else
            s_string = "[Αριθμός] LIKE '*" & Trim(txt_arithmos.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΠΕΡΙΟΧΗΣ
    If Trim(txt_perioxi.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Περιοχή] LIKE '*" & Trim(txt_perioxi.Text) & "*'"
        Else
            s_string = "[Περιοχή] LIKE '*" & Trim(txt_perioxi.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΔΗΜΟΥ
    If Trim(co_dimoi.Text) <> "" Then
        ado_dimoi.Recordset.MoveFirst
        ado_dimoi.Recordset.Find "[περιγραφή] LIKE '" & Trim(co_dimoi.Text) & "'"
        If Not ado_dimoi.Recordset.EOF Then
            If s_string <> "" Then
                s_string = s_string & " AND [Δήμος] LIKE '" & ado_dimoi.Recordset.Fields(0).Value & "'"
            Else
                s_string = "[Δήμος] LIKE '" & ado_dimoi.Recordset.Fields(0).Value & "'"
            End If
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΠΕΡΙΦΕΡΕΙΑΚΗΣ ΕΝΟΤΗΤΑΣ
    If Trim(co_pe.Text) <> "" Then
        ado_pe.Recordset.MoveFirst
        ado_pe.Recordset.Find "[περιγραφή] LIKE '" & Trim(co_pe.Text) & "'"
        If Not ado_pe.Recordset.EOF Then
            If s_string <> "" Then
                s_string = s_string & " AND [Περιφερειακή Ενότητα] LIKE '" & ado_pe.Recordset.Fields(0).Value & "'"
            Else
                s_string = "[Περιφερειακή Ενότητα] LIKE '" & ado_pe.Recordset.Fields(0).Value & "'"
            End If
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΤΑΧΥΔΡΟΜΙΚΟΥ ΚΩΔΙΚΑ
    If Trim(txt_tk.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Ταχυδρομικός Κώδικας] LIKE '*" & Trim(txt_tk.Text) & "*'"
        Else
            s_string = "[Ταχυδρομικός Κώδικας] LIKE '*" & Trim(txt_tk.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΣΤΑΘΕΡΟΥ ΤΗΛΕΦΩΝΟΥ
    If Trim(txt_til.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΣταθερόΤηλέφωνο] LIKE '*" & Trim(txt_til.Text) & "*'"
        Else
            s_string = "[ΣταθερόΤηλέφωνο] LIKE '*" & Trim(txt_til.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΚΙΝΗΤΟΥ ΤΗΛΕΦΩΝΟΥ
    If Trim(txt_kinito.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΚινητόΤηλέφωνο] LIKE '*" & Trim(txt_kinito.Text) & "*'"
        Else
            s_string = "[ΚινητόΤηλέφωνο] LIKE '*" & Trim(txt_kinito.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΑΡΙΘΜΟΥ FAX
    If Trim(txt_fax.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΑριθμόςΦαξ] LIKE '*" & Trim(txt_fax.Text) & "*'"
        Else
            s_string = "[ΑριθμόςΦαξ] LIKE '*" & Trim(txt_fax.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ EMAIL
    If Trim(txt_email.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΌνομαEmail] LIKE '*" & Trim(txt_email.Text) & "*'"
        Else
            s_string = "[ΌνομαEmail] LIKE '*" & Trim(txt_email.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ WEB SITE
    If Trim(txt_web.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [WebSite] LIKE '*" & Trim(txt_web.Text) & "*'"
        Else
            s_string = "[WebSite] LIKE '*" & Trim(txt_web.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΑΦΜ
    If Trim(txt_afm.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΑΦΜ] LIKE '*" & Trim(txt_afm.Text) & "*'"
        Else
            s_string = "[ΑΦΜ] LIKE '*" & Trim(txt_afm.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΔΟΥ
    If Trim(co_doi.Text) <> "" Then
        ado_doi.Recordset.MoveFirst
        ado_doi.Recordset.Find "[περιγραφή] LIKE '" & Trim(co_doi.Text) & "'"
        If Not ado_doi.Recordset.EOF Then
            If s_string <> "" Then
                s_string = s_string & " AND [ΔΟΥ] LIKE '" & ado_doi.Recordset.Fields(0).Value & "'"
            Else
                s_string = "[ΔΟΥ] LIKE '" & ado_doi.Recordset.Fields(0).Value & "'"
            End If
        End If
    End If
    '
    ado_pr.Recordset.Filter = s_string
    If s_sort <> "" Then
        ado_pr.Recordset.Sort = Trim(s_sort)
    End If
    '
    Me.canc_bt.Enabled = True
    '
    'for_search = 0
    c_rec = 1
    txt_kod.Locked = True
    
End Sub

Private Sub taksin_Click()

    If dt_pr.Col >= 0 Then
        ado_pr.Recordset.Sort = "[" & Trim(ado_pr.Recordset.Fields(dt_pr.Col).Name) & "]"
        s_sort = "[" & Trim(ado_pr.Recordset.Fields(dt_pr.Col).Name) & "]"
    Else
        ado_pr.Recordset.Sort = "[" & Trim(ado_pr.Recordset.Fields(defined_col).Name) & "]"
        s_sort = "[" & Trim(ado_pr.Recordset.Fields(defined_col).Name) & "]"
    End If
  
End Sub

Private Sub txt_afm_GotFocus()

    With txt_afm
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txt_arithmos_GotFocus()

    With txt_arithmos
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txt_email_GotFocus()

    With txt_email
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txt_ep_GotFocus()

    With txt_ep
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txt_fax_GotFocus()

    With txt_fax
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txt_kinito_GotFocus()

    With txt_kinito
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txt_odos_GotFocus()

    With txt_odos
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txt_perioxi_GotFocus()

    With txt_perioxi
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txt_til_GotFocus()

    With txt_til
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txt_tk_GotFocus()

    With txt_tk
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub txt_web_GotFocus()

    With txt_web
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub up_bt_Click()

    'Αποθήκευση ΣΤΟΙΧΕΙΩΝ (εκτός id) στους ΠΡΟΜΗΘΕΥΤΕΣ
    If ado_pr.Recordset.RecordCount >= 1 Then
        c_rec = ado_pr.Recordset.AbsolutePosition
    End If
    If txt_ep.Text <> "" Then
        ado_pr.Recordset.Fields(1).Value = txt_ep.Text
    End If
    If txt_odos.Text <> "" Then
        ado_pr.Recordset.Fields(2).Value = txt_odos.Text
    End If
    If txt_arithmos.Text <> "" Then
        ado_pr.Recordset.Fields(3).Value = txt_arithmos.Text
    End If
    If txt_perioxi.Text <> "" Then
    ado_pr.Recordset.Fields(4).Value = txt_perioxi.Text
    End If
    If co_dimoi.Text <> "" Then
        If co_dimoi.MatchedWithList = False Then
            'ΕΙΣΑΓΩΓΗ ΝΕΟΥ ΔΗΜΟΥ
            If ado_dimoi.Recordset.RecordCount >= 1 Then
                ado_dimoi.Recordset.Sort = "[" & Trim(ado_dimoi.Recordset.Fields(0).Name) & "]"
                ado_dimoi.Recordset.MoveLast
                id = ado_dimoi.Recordset.Fields(0).Value
            Else
                id = 0
            End If
            ado_dimoi.Recordset.AddNew
            ado_dimoi.Recordset.Fields(0) = id + 1
            ado_dimoi.Recordset.Fields(1) = co_dimoi.Text
            ado_dimoi.Recordset.UpdateBatch adAffectCurrent
            ado_pr.Recordset.Fields(5).Value = id + 1
        Else
            ado_dimoi.Recordset.MoveFirst
            ado_dimoi.Recordset.Find "[περιγραφή] LIKE '" & Trim(co_dimoi.Text) & "'"
            If Not ado_dimoi.Recordset.EOF Then
                ado_pr.Recordset.Fields(5).Value = ado_dimoi.Recordset.Fields(0).Value
            End If
        End If
    End If
    If co_pe.Text <> "" Then
        If co_pe.MatchedWithList = False Then
            'ΕΙΣΑΓΩΓΗ ΝΕΑΣ ΠΕΡΙΦΕΡΕΙΑΚΗΣ ΕΝΟΤΗΤΑΣ
            If ado_pe.Recordset.RecordCount >= 1 Then
                ado_pe.Recordset.Sort = "[" & Trim(ado_pe.Recordset.Fields(0).Name) & "]"
                ado_pe.Recordset.MoveLast
                id = ado_pe.Recordset.Fields(0).Value
            Else
                id = 0
            End If
            ado_pe.Recordset.AddNew
            ado_pe.Recordset.Fields(0) = id + 1
            ado_pe.Recordset.Fields(1) = co_pe.Text
            ado_pe.Recordset.UpdateBatch adAffectCurrent
            ado_pr.Recordset.Fields(6).Value = id + 1
        Else
            ado_pe.Recordset.MoveFirst
            ado_pe.Recordset.Find "[περιγραφή] LIKE '" & Trim(co_pe.Text) & "'"
            If Not ado_pe.Recordset.EOF Then
                ado_pr.Recordset.Fields(6).Value = ado_pe.Recordset.Fields(0).Value
            End If
        End If
    End If
    If txt_tk.Text <> "" Then
            ado_pr.Recordset.Fields(7).Value = txt_tk.Text
    End If
    If txt_til.Text <> "" Then
        ado_pr.Recordset.Fields(8).Value = txt_til.Text
    End If
    If txt_kinito.Text <> "" Then
        ado_pr.Recordset.Fields(9).Value = txt_kinito.Text
    End If
    If txt_fax.Text <> "" Then
        ado_pr.Recordset.Fields(10).Value = txt_fax.Text
    End If
    If txt_email.Text <> "" Then
        ado_pr.Recordset.Fields(11).Value = txt_email.Text
    End If
    If txt_web.Text <> "" Then
        ado_pr.Recordset.Fields(12).Value = txt_web.Text
    End If
    If txt_afm.Text <> "" Then
        ado_pr.Recordset.Fields(13).Value = txt_afm.Text
    End If
    If co_doi.Text <> "" Then
        If co_doi.MatchedWithList = False Then
            'ΕΙΣΑΓΩΓΗ ΝΕΑΣ ΔΟΥ
            If ado_doi.Recordset.RecordCount >= 1 Then
                ado_doi.Recordset.Sort = "[" & Trim(ado_doi.Recordset.Fields(0).Name) & "]"
                ado_doi.Recordset.MoveLast
                id = ado_doi.Recordset.Fields(0).Value
            Else
                id = 0
            End If
            ado_doi.Recordset.AddNew
            ado_doi.Recordset.Fields(0) = id + 1
            ado_doi.Recordset.Fields(1) = co_doi.Text
            ado_doi.Recordset.UpdateBatch adAffectCurrent
            ado_pr.Recordset.Fields(14).Value = id + 1
        Else
            ado_doi.Recordset.MoveFirst
            ado_doi.Recordset.Find "[περιγραφή] LIKE '" & Trim(co_doi.Text) & "'"
            If Not ado_doi.Recordset.EOF Then
                ado_pr.Recordset.Fields(14).Value = ado_doi.Recordset.Fields(0).Value
            End If
        End If
    End If
    '
    ado_pr.Recordset.UpdateBatch adAffectCurrent
    
End Sub
