VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form athlet_management 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Διαχείριση Αθλητών"
   ClientHeight    =   10740
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   12660
   ForeColor       =   &H00000000&
   LinkTopic       =   "entity_management"
   MDIChild        =   -1  'True
   ScaleHeight     =   10740
   ScaleWidth      =   12660
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Όλοι οι Αθλητές"
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
      Height          =   4695
      Left            =   120
      TabIndex        =   62
      Top             =   5880
      Width           =   11535
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ε&ξαγωγή Επιλεγμένων Αθλητών στο Excel"
         DisabledPicture =   "athlet_management.frx":0000
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
         Left            =   5520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "athlet_management.frx":268B
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   3960
         Width           =   3135
      End
      Begin VB.CommandButton taksin 
         BackColor       =   &H80000014&
         Caption         =   "&Ταξινόμηση"
         DisabledPicture =   "athlet_management.frx":4D16
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
         Picture         =   "athlet_management.frx":5689
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox image_path 
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
         Left            =   8760
         Locked          =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   64
         Top             =   3480
         Width           =   2655
      End
      Begin VB.CommandButton bt_print 
         BackColor       =   &H80000014&
         Caption         =   "Εκτύπω&ση"
         DisabledPicture =   "athlet_management.frx":A390
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
         Picture         =   "athlet_management.frx":AD03
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   3960
         Width           =   2655
      End
      Begin MSDataGridLib.DataGrid dt_athlites 
         Bindings        =   "athlet_management.frx":F708
         Height          =   3135
         Left            =   120
         TabIndex        =   66
         Top             =   360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5530
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
      Begin MSAdodcLib.Adodc ado_athlites 
         Height          =   375
         Left            =   120
         Top             =   3480
         Width           =   8565
         _ExtentX        =   15108
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
         RecordSource    =   "Αθλητές"
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
      Begin MSAdodcLib.Adodc tmp_ado_athlites 
         Height          =   375
         Left            =   1800
         Top             =   4080
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
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
         RecordSource    =   "Αθλητές"
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
      Begin MSDataGridLib.DataGrid tmp_dt_athlites 
         Bindings        =   "athlet_management.frx":F723
         Height          =   495
         Left            =   3600
         TabIndex        =   72
         Top             =   4080
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
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
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   3135
         Left            =   8760
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Φωτογραφία Αθλητή / Αθλήτριας"
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   9360
         TabIndex        =   67
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.CommandButton cancel_cur_rec 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   8760
      MaskColor       =   &H00000000&
      Picture         =   "athlet_management.frx":F742
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Ακύρωση Τρέχουσας Εγγραφής"
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Frame Frame10 
      Caption         =   "Παρατηρήσεις"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   56
      Top             =   3960
      Width           =   4095
      Begin VB.TextBox txt_parat 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   360
         Width           =   3855
      End
   End
   Begin MSAdodcLib.Adodc ado_dimoi 
      Height          =   375
      Left            =   11760
      Top             =   2880
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
      Left            =   11760
      Top             =   3240
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
   Begin MSAdodcLib.Adodc ado_pateres 
      Height          =   375
      Left            =   11760
      Top             =   3600
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
      RecordSource    =   "Μέλη"
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
   Begin MSAdodcLib.Adodc ado_miteres 
      Height          =   375
      Left            =   11760
      Top             =   3960
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
      RecordSource    =   "Μέλη"
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
   Begin MSAdodcLib.Adodc ado_sxolia 
      Height          =   375
      Left            =   11760
      Top             =   4320
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
      RecordSource    =   "Σχολεία_ταξινομημένα"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Στοιχεία Αθλητή"
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
      Height          =   6255
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   11535
      Begin VB.Frame Frame11 
         Caption         =   "Αδέρφια"
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
         Height          =   2295
         Left            =   10320
         TabIndex        =   69
         Top             =   2640
         Width           =   1095
         Begin VB.TextBox txt_aderf 
            Height          =   375
            Left            =   120
            TabIndex        =   71
            Text            =   "0"
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "Πλήθος αδερφιών που είναι ηδη εγγεγραμ- μένα:"
            Height          =   1335
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Στοιχεία Γονέων"
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
         Height          =   2295
         Left            =   4320
         TabIndex        =   37
         Top             =   2640
         Width           =   6015
         Begin VB.Frame Frame8 
            Caption         =   "Μητέρα"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   40
            Top             =   1200
            Width           =   5775
            Begin VB.CommandButton Command4 
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Caption         =   "x"
               DisabledPicture =   "athlet_management.frx":F901
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   161
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5280
               MaskColor       =   &H000000FF&
               Picture         =   "athlet_management.frx":FA1F
               Style           =   1  'Graphical
               TabIndex        =   74
               ToolTipText     =   "Διαγραφή Επιλεγμένης Μητέρας από τον Τρέχων Αθλητή"
               Top             =   100
               UseMaskColor    =   -1  'True
               Width           =   375
            End
            Begin VB.CommandButton sear_moth 
               BackColor       =   &H80000014&
               Caption         =   "Αναζήτηση(&μ)"
               DisabledPicture =   "athlet_management.frx":FFD6
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
               Left            =   4560
               MaskColor       =   &H80000014&
               Picture         =   "athlet_management.frx":10949
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   330
               Width           =   1095
            End
            Begin MSDataGridLib.DataGrid dt_mother 
               Bindings        =   "athlet_management.frx":15CFD
               Height          =   495
               Left            =   120
               TabIndex        =   41
               Top             =   210
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   873
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
         End
         Begin VB.Frame Frame6 
            Caption         =   "Πατέρας"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   5775
            Begin VB.CommandButton bt_del_athl 
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Caption         =   "x"
               DisabledPicture =   "athlet_management.frx":15D17
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   161
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5280
               MaskColor       =   &H000000FF&
               Picture         =   "athlet_management.frx":15E35
               Style           =   1  'Graphical
               TabIndex        =   73
               ToolTipText     =   "Διαγραφή Επιλεγμένου Πατέρα από τον Τρέχων Αθλητή"
               Top             =   100
               UseMaskColor    =   -1  'True
               Width           =   375
            End
            Begin VB.CommandButton sear_pat 
               BackColor       =   &H80000014&
               Caption         =   "Αναζήτηση(&π)"
               DisabledPicture =   "athlet_management.frx":163EC
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
               Left            =   4560
               MaskColor       =   &H80000014&
               Picture         =   "athlet_management.frx":16D5F
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   340
               Width           =   1095
            End
            Begin MSDataGridLib.DataGrid dt_father 
               Bindings        =   "athlet_management.frx":1C113
               Height          =   495
               Left            =   120
               TabIndex        =   39
               Top             =   220
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   873
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
         End
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
         Height          =   1575
         Left            =   4320
         TabIndex        =   30
         Top             =   1200
         Width           =   7095
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
            Height          =   315
            Left            =   5880
            TabIndex        =   15
            Top             =   650
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
            Height          =   315
            Left            =   1200
            TabIndex        =   14
            Top             =   650
            Width           =   3975
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
            Height          =   315
            Left            =   6240
            TabIndex        =   13
            Top             =   310
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
            Height          =   315
            Left            =   1200
            TabIndex        =   12
            Top             =   310
            Width           =   3975
         End
         Begin MSDataListLib.DataCombo co_pe 
            Bindings        =   "athlet_management.frx":1C12D
            Height          =   345
            Left            =   4920
            TabIndex        =   17
            Top             =   960
            Width           =   2055
            _ExtentX        =   3625
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
            Bindings        =   "athlet_management.frx":1C142
            Height          =   345
            Left            =   1200
            TabIndex        =   16
            Top             =   960
            Width           =   2055
            _ExtentX        =   3625
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
            TabIndex        =   36
            Top             =   300
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
            Left            =   5280
            TabIndex        =   35
            Top             =   300
            Width           =   825
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
            Left            =   240
            TabIndex        =   34
            Top             =   640
            Width           =   825
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
            TabIndex        =   33
            Top             =   960
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
            Left            =   3960
            TabIndex        =   32
            Top             =   960
            Width           =   825
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
            Height          =   375
            Left            =   4440
            TabIndex        =   31
            Top             =   640
            Width           =   1305
         End
      End
      Begin VB.CommandButton kl_bt 
         BackColor       =   &H80000014&
         Caption         =   "Κ&λείσιμο"
         DisabledPicture =   "athlet_management.frx":1C15A
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
         Picture         =   "athlet_management.frx":1CACD
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton insert_bt 
         BackColor       =   &H80000014&
         Caption         =   "Προσ&θήκη"
         DisabledPicture =   "athlet_management.frx":22545
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
         Picture         =   "athlet_management.frx":22663
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton save_command 
         BackColor       =   &H80000014&
         Caption         =   "&Αποθήκευση"
         DisabledPicture =   "athlet_management.frx":273F2
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
         Picture         =   "athlet_management.frx":2BFF1
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton canc_bt 
         BackColor       =   &H80000014&
         Caption         =   "Ακύ&ρωση"
         DisabledPicture =   "athlet_management.frx":30BF0
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
         Picture         =   "athlet_management.frx":35B0D
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Επαναφορά όλων των Αθλητών σε Προηγούμενη Σταθερή Κατάσταση"
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CommandButton up_bt 
         BackColor       =   &H80000014&
         Caption         =   "&Ενημέρωση"
         Default         =   -1  'True
         DisabledPicture =   "athlet_management.frx":3AA2A
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
         Picture         =   "athlet_management.frx":42524
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000014&
         Caption         =   "&Καθαρισμός"
         DisabledPicture =   "athlet_management.frx":4A01E
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
         Picture         =   "athlet_management.frx":4A166
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton del_bt 
         BackColor       =   &H80000014&
         Caption         =   "&Διαγραφή"
         DisabledPicture =   "athlet_management.frx":4ECAA
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
         Picture         =   "athlet_management.frx":53AAE
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton sear_bt 
         BackColor       =   &H80000014&
         Caption         =   "Α&ναζήτηση"
         DisabledPicture =   "athlet_management.frx":588B2
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
         Picture         =   "athlet_management.frx":5DC66
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   5040
         Width           =   975
      End
      Begin VB.Frame Frame9 
         Caption         =   "Στοιχεία ΚΟΕ"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4320
         TabIndex        =   45
         Top             =   360
         Width           =   7095
         Begin VB.TextBox tmp_am 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   9
            Top             =   330
            Width           =   1335
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   285
            Left            =   3600
            TabIndex        =   10
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
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
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   285
            Left            =   5880
            TabIndex        =   11
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
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
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ημ/νία Λήξης"
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
            Left            =   4860
            TabIndex        =   48
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ημ/νία Έναρξης"
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
            Left            =   2220
            TabIndex        =   47
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Αριθμός ΚΟΕ"
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
            Left            =   0
            TabIndex        =   46
            Top             =   290
            Width           =   885
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
         Height          =   1815
         Left            =   120
         TabIndex        =   25
         Top             =   2280
         Width           =   4095
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
            Height          =   285
            Left            =   1440
            TabIndex        =   8
            Top             =   1275
            Width           =   2535
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
            Height          =   285
            Left            =   1440
            TabIndex        =   7
            Top             =   975
            Width           =   2535
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
            Height          =   285
            Left            =   1440
            TabIndex        =   6
            Top             =   675
            Width           =   2535
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
            Height          =   285
            Left            =   1440
            TabIndex        =   5
            Top             =   360
            Width           =   2535
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
            TabIndex        =   29
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
            TabIndex        =   28
            Top             =   675
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
            TabIndex        =   27
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
            TabIndex        =   26
            Top             =   1245
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
         Height          =   2175
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   4095
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
            Height          =   285
            Left            =   1440
            TabIndex        =   2
            Top             =   880
            Width           =   2295
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
            Height          =   275
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   905
            Width           =   255
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
            Height          =   285
            Left            =   1440
            TabIndex        =   1
            Top             =   600
            Width           =   2535
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   285
            Left            =   1440
            TabIndex        =   3
            Top             =   1200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
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
         Begin MSDataListLib.DataCombo co_sxolia 
            Bindings        =   "athlet_management.frx":6301A
            Height          =   345
            Left            =   1440
            TabIndex        =   4
            Top             =   1500
            Width           =   2535
            _ExtentX        =   4471
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
         Begin VB.TextBox tmp_ar_kar 
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
            Left            =   3120
            TabIndex        =   0
            Top             =   310
            Width           =   850
         End
         Begin VB.TextBox tmp_kod 
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
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   310
            Width           =   735
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Αρ. Κάρτας"
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
            Left            =   2040
            TabIndex        =   59
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label Label16 
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
            Left            =   120
            TabIndex        =   44
            Top             =   300
            Width           =   1245
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Σχολείο"
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
            TabIndex        =   42
            Top             =   1510
            Width           =   1455
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
            Left            =   -600
            TabIndex        =   24
            Top             =   1200
            Width           =   1935
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
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   600
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
            TabIndex        =   22
            Top             =   880
            Width           =   1005
         End
      End
   End
End
Attribute VB_Name = "athlet_management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public initial_string As String
Public final_string As String
Public it_is_a_new_record As Integer
Public id_α, flag_pateras, flag_mitera, defined_col, for_search As Integer
Public rs_ado_dimoi As ADODB.Recordset
Public rs_ado_pe As ADODB.Recordset
Public rs_ado_sxolia As ADODB.Recordset
Public rs_ado_miteres As ADODB.Recordset
Public rs_ado_pateres As ADODB.Recordset
Dim bytData() As Byte
Public sFile As String
Public is_to_delete As Integer
    

Private Sub ado_athlites_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    If Not adReason = 7 Then ' adRsnRequery = 7
        If Not adReason = 10 And pRecordset.AbsolutePosition >= 1 Then
            If Trim(Me.ado_athlites.Recordset.Fields(0).Value) <> "" And Me.it_is_a_new_record = 0 Then
                tmp_kod.Text = Trim(str(Me.ado_athlites.Recordset.Fields(0).Value))
            Else
                If Me.it_is_a_new_record = 1 Then
                Else
                    tmp_kod.Text = ""
                End If
            End If
            If Trim(pRecordset.Fields(22).Value) <> "" And Me.it_is_a_new_record = 0 Then
                tmp_ar_kar.Text = Trim(pRecordset.Fields(22).Value)
            Else
                If Me.it_is_a_new_record = 1 Then
                Else
                    tmp_ar_kar.Text = ""
                End If
            End If
            If Trim(pRecordset.Fields(1).Value) <> "" Then
                tmp_am.Text = pRecordset.Fields(1).Value
            Else
                tmp_am.Text = ""
            End If
            If Trim(pRecordset.Fields(18).Value) <> "" Then
                MaskEdBox2.Text = pRecordset.Fields(18).Value
            Else
                MaskEdBox2.Text = "  /  /    "
            End If
            If Trim(pRecordset.Fields(19).Value) <> "" Then
                MaskEdBox3.Text = pRecordset.Fields(19).Value
            Else
                MaskEdBox3.Text = "  /  /    "
            End If
            If Trim(pRecordset.Fields(3).Value) <> "" And Me.it_is_a_new_record = 0 Then
                tmp_onoma.Text = pRecordset.Fields(3).Value
            Else
                If Me.it_is_a_new_record = 1 Then
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
                Me.MaskEdBox1.Text = "  /  /    "
            End If
            If Trim(Me.tmp_am.Text) <> "" Then
                pRecordset.Fields(1).Value = Me.tmp_am.Text
            End If
            If Trim(pRecordset.Fields(15).Value) <> 0 Then
                rs_ado_pateres.Filter = "[id] LIKE '" & str(pRecordset.Fields(15).Value) & "'"
            Else
                rs_ado_pateres.Filter = "[id] LIKE '" & str(-1) & "'"
            End If
            If Trim(pRecordset.Fields(16).Value) <> 0 Then
                rs_ado_miteres.Filter = "[id] LIKE '" & str(pRecordset.Fields(16).Value) & "'"
            Else
                rs_ado_miteres.Filter = "[id] LIKE '" & str(-1) & "'"
            End If
            If Trim(pRecordset.Fields(17).Value) <> "" Then
                Me.co_sxolia.Text = pRecordset.Fields(17).Value
            Else
                Me.co_sxolia.Text = ""
            End If
            'ΠΑΡΟΥΣΙΑΣΗ ΠΑΡΑΤΗΡΗΣΕΩΝ
            If Trim(pRecordset.Fields(21).Value) <> "" Then
                Me.txt_parat.Text = pRecordset.Fields(21).Value
            Else
                Me.txt_parat.Text = ""
            End If
            'ΠΛΗΘΟΣ ΑΔΕΡΦΙΩΝ ΠΟΥ ΕΙΝΑΙ ΗΔΗ ΕΓΓΕΓΡΑΜΜΕΝΑ
            If Trim(pRecordset.Fields(23).Value) <> "" Then
                Me.txt_aderf.Text = pRecordset.Fields(23).Value
            Else
                Me.txt_aderf.Text = 0
            End If
        End If
    End If
    If adReason <> 7 Then
        If pRecordset.AbsolutePosition >= 1 Then
            Me.ado_athlites.Caption = "Αθλητής " & pRecordset.AbsolutePosition & " από " & pRecordset.RecordCount
        End If
    End If
    
End Sub

Private Sub bt_del_athl_Click()

    'ΕΝΗΜΕΡΩΣΗ ΣΤΟΙΧΕΙΩΝ ΠΑΤΕΡΑ
    melos_id = 0
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
    athlet_management.ado_pateres.Recordset.Filter = "[id] LIKE '" & str(melos_id) & "'"
    
End Sub

Private Sub bt_print_Click()

    frm_pr_athl.Show

End Sub

Private Sub canc_bt_Click()


    it_is_a_new_record = 0
    MDIForm1.is_a_new_record_without_save = 0
    Me.Command2.Enabled = False
    If Me.tmp_ado_athlites.Recordset.RecordCount >= 1 Then
        
        for_search = 0
        s_string = ""
        s_string2 = ""
        s_sort = ""
        Me.tmp_kod.Locked = True
    
        Me.tmp_ado_athlites.Recordset.Requery
        Me.tmp_ado_athlites.Refresh
        Me.ado_athlites.Recordset.Requery
        Me.ado_athlites.Refresh

        Me.ado_athlites.Recordset.Filter = ""
        Me.ado_athlites.Recordset.Sort = "[id]"
        '***************************************************************
        If Not Me.ado_athlites.Recordset.EOF Then
            Me.ado_athlites.Recordset.MoveFirst
            If Trim(Me.ado_athlites.Recordset.Fields(0).Value) <> "" Then
                tmp_kod.Text = Trim(str(Me.ado_athlites.Recordset.Fields(0).Value))
            Else
                tmp_kod.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(22).Value) <> "" Then
                tmp_ar_kar.Text = Trim(str(Me.ado_athlites.Recordset.Fields(22).Value))
            Else
                tmp_ar_kar.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(1).Value) <> "" Then
                tmp_am.Text = str(Me.ado_athlites.Recordset.Fields(1).Value)
            Else
                tmp_am.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(18).Value) <> "" Then
                MaskEdBox2.Text = Me.ado_athlites.Recordset.Fields(18).Value
            Else
                MaskEdBox2.Text = "00/00/0000"
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(19).Value) <> "" Then
                MaskEdBox3.Text = Me.ado_athlites.Recordset.Fields(19).Value
            Else
                MaskEdBox3.Text = "00/00/0000"
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(3).Value) <> "" Then
                tmp_onoma.Text = Me.ado_athlites.Recordset.Fields(3).Value
            Else
                tmp_onoma.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(2).Value) <> "" Then
                tmp_eponimo.Text = Me.ado_athlites.Recordset.Fields(2).Value
            Else
                tmp_eponimo.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(4).Value) <> "" Then
                tmp_odos.Text = Me.ado_athlites.Recordset.Fields(4).Value
            Else
                tmp_odos.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(5).Value) <> "" Then
                tmp_arithmos.Text = Me.ado_athlites.Recordset.Fields(5).Value
            Else
                tmp_arithmos.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(6).Value) <> "" Then
                tmp_perioxi.Text = Me.ado_athlites.Recordset.Fields(6).Value
            Else
                tmp_perioxi.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(8).Value) <> "" Then
                Me.co_pe.Text = Me.ado_athlites.Recordset.Fields(8).Value
            Else
                Me.co_pe.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(7).Value) <> "" Then
                Me.co_dimoi.Text = Me.ado_athlites.Recordset.Fields(7).Value
            Else
                Me.co_dimoi.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(9).Value) <> "" Then
                tmp_tk.Text = Me.ado_athlites.Recordset.Fields(9).Value
            Else
                tmp_tk.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(10).Value) <> "" Then
                tmp_til_oikias.Text = Me.ado_athlites.Recordset.Fields(10).Value
            Else
                tmp_til_oikias.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(11).Value) <> "" Then
                tmp_kinito.Text = Me.ado_athlites.Recordset.Fields(11).Value
            Else
                tmp_kinito.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(12).Value) <> "" Then
                tmp_fax.Text = Me.ado_athlites.Recordset.Fields(12).Value
            Else
                tmp_fax.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(13).Value) <> "" Then
                tmp_email.Text = Me.ado_athlites.Recordset.Fields(13).Value
            Else
                tmp_email.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(14).Value) <> "" Then
                Me.MaskEdBox1.Text = Me.ado_athlites.Recordset.Fields(14).Value
            Else
                Me.MaskEdBox1.Text = "00/00/0000"
            End If
            'ΠΑΡΟΥΣΙΑΣΗ ΣΤΟΙΧΕΙΩΝ ΠΑΤΕΡΑ
            Me.dt_father.Columns(0).Visible = False
            Me.dt_father.Columns(1).Caption = "Α.Μ."
            Me.dt_father.Columns(1).Width = 1000
            Me.dt_father.Columns(2).Caption = "Επώνυμο"
            Me.dt_father.Columns(2).Width = 1500
            Me.dt_father.Columns(3).Caption = "Όνομα"
            Me.dt_father.Columns(3).Width = 1200
            For i = 4 To rs_ado_pateres.Fields.Count - 1
                Me.dt_father.Columns(i).Visible = False
            Next i
            If Trim(Me.ado_athlites.Recordset.Fields(15).Value) <> "" Then
                rs_ado_pateres.Filter = "[id] LIKE '" & str(Me.ado_athlites.Recordset.Fields(15).Value) & "'"
            Else
                rs_ado_pateres.Filter = "[id] LIKE '" & str(-1) & "'"
            End If
            '
            'ΠΑΡΟΥΣΙΑΣΗ ΣΤΟΙΧΕΙΩΝ ΜΗΤΕΡΑΣ
            Me.dt_mother.Columns(0).Visible = False
            Me.dt_mother.Columns(1).Caption = "A.Μ."
            Me.dt_mother.Columns(1).Width = 1000
            Me.dt_mother.Columns(2).Caption = "Επώνυμο"
            Me.dt_mother.Columns(2).Width = 1500
            Me.dt_mother.Columns(3).Caption = "Όνομα"
            Me.dt_mother.Columns(3).Width = 1200
            For i = 4 To rs_ado_miteres.Fields.Count - 1
                Me.dt_mother.Columns(i).Visible = False
            Next i
            If Trim(Me.ado_athlites.Recordset.Fields(16).Value) <> "" Then
                rs_ado_miteres.Filter = "[id] LIKE '" & str(Me.ado_athlites.Recordset.Fields(16).Value) & "'"
            Else
                rs_ado_miteres.Filter = "[id] LIKE '" & str(-1) & "'"
            End If
            '
            'ΠΑΡΟΥΣΙΑΣΗ ΣΤΟΙΧΕΩΝ ΣΧΟΛΕΙΟΥ
            If Trim(Me.ado_athlites.Recordset.Fields(17).Value) <> "" Then
                Me.co_sxolia.Text = Me.ado_athlites.Recordset.Fields(17).Value
            Else
                Me.co_sxolia.Text = ""
            End If
            Me.insert_bt.Enabled = True
            Me.up_bt.Enabled = True
            Me.sear_bt.Enabled = False
            Me.canc_bt.Enabled = False
            Me.del_bt.Enabled = True
            Me.Image1.Enabled = True
            Me.Label19.Visible = True
            Me.Command1.Enabled = True
            Me.cancel_cur_rec.Enabled = True
            Me.save_command.Enabled = False
        End If
        '***************************************************************

        Me.dt_athlites.Columns(0).Caption = "Κωδικός"
        Me.dt_athlites.Columns(0).Width = 1000
        Me.dt_athlites.Columns(1).Caption = "Αριθμός ΚΟΕ"
        Me.dt_athlites.Columns(1).Width = 1200
        Me.dt_athlites.Columns(2).Caption = "Επώνυμο"
        Me.dt_athlites.Columns(2).Width = 2300
        Me.dt_athlites.Columns(3).Caption = "Όνομα"
        Me.dt_athlites.Columns(3).Width = 1800
        For i = 4 To 13
            Me.dt_athlites.Columns(i).Visible = False
        Next i
        Me.dt_athlites.Columns(14).Caption = "Ημ/νία Γέννησης"
        Me.dt_athlites.Columns(14).Width = 1400
        For i = 15 To Me.ado_athlites.Recordset.Fields.Count - 1
            Me.dt_athlites.Columns(i).Visible = False
        Next i
        MDIForm1.s_string = ""
        MDIForm1.s_sort = ""
        MDIForm1.rep_lbl = ""
    
    Else
    
        Me.ado_athlites.Recordset.Requery
        Me.ado_athlites.Refresh
        Me.dt_athlites.Refresh
        If Me.ado_athlites.Recordset.RecordCount >= 1 Then
            Me.ado_athlites.Recordset.MoveFirst
            Me.dt_athlites.Columns(0).Caption = "Κωδικός"
            Me.dt_athlites.Columns(0).Width = 1000
            Me.dt_athlites.Columns(1).Caption = "Αριθμός ΚΟΕ"
            Me.dt_athlites.Columns(1).Width = 1200
            Me.dt_athlites.Columns(2).Caption = "Επώνυμο"
            Me.dt_athlites.Columns(2).Width = 2300
            Me.dt_athlites.Columns(3).Caption = "Όνομα"
            Me.dt_athlites.Columns(3).Width = 1800
            For i = 4 To 13
                Me.dt_athlites.Columns(i).Visible = False
            Next i
            Me.dt_athlites.Columns(14).Caption = "Ημ/νία Γέννησης"
            Me.dt_athlites.Columns(14).Width = 1400
            For i = 15 To Me.ado_athlites.Recordset.Fields.Count - 1
                Me.dt_athlites.Columns(i).Visible = False
            Next i
        Else
            Me.tmp_kod.Locked = True
            Me.tmp_kod = ""
            Me.tmp_ar_kar.Locked = True
            Me.tmp_ar_kar = ""
            Me.tmp_am.Locked = True
            Me.tmp_am.Text = ""
            Me.tmp_onoma.Locked = True
            Me.tmp_onoma.Text = ""
            Me.tmp_eponimo.Locked = True
            Me.tmp_eponimo.Text = ""
            Me.tmp_odos.Locked = True
            Me.tmp_odos.Text = ""
            Me.tmp_arithmos.Locked = True
            Me.tmp_arithmos.Text = ""
            Me.tmp_perioxi.Locked = True
            Me.tmp_perioxi.Text = ""
            Me.co_dimoi.Locked = True
            Me.co_dimoi.Text = ""
            Me.co_pe.Locked = True
            Me.co_pe.Text = ""
            Me.tmp_tk.Locked = True
            Me.tmp_tk.Text = ""
            Me.tmp_til_oikias.Locked = True
            Me.tmp_til_oikias.Text = ""
            Me.tmp_kinito.Locked = True
            Me.tmp_kinito.Text = ""
            Me.tmp_fax.Locked = True
            Me.tmp_fax.Text = ""
            Me.tmp_email.Locked = True
            Me.tmp_email.Text = ""
            Me.MaskEdBox1.Enabled = False
            Me.MaskEdBox1.Text = "  /  /    "
            Me.MaskEdBox2.Enabled = False
            Me.MaskEdBox2.Text = "  /  /    "
            Me.MaskEdBox3.Enabled = False
            Me.MaskEdBox3.Text = "  /  /    "
            Me.ado_pateres.Recordset.Filter = "[id] LIKE '" & str(-1) & "'"
            Me.ado_miteres.Recordset.Filter = "[id] LIKE '" & str(-1) & "'"
            Me.co_sxolia.Locked = True
            Me.co_sxolia.Text = ""
            Me.txt_parat.Locked = True
            Me.txt_parat.Text = ""
        End If
        '

        Me.save_command.Enabled = False
        Me.insert_bt.Enabled = True
        Me.Command1.Enabled = True
        Me.cancel_cur_rec.Enabled = True
        Me.up_bt.Enabled = True
        Me.del_bt.Enabled = True
        Me.Image1.Enabled = True
        Me.Label19.Visible = True
    
    End If

    Me.canc_bt.Enabled = False

End Sub

Private Sub cancel_cur_rec_Click()

it_is_a_new_record = 0

    Dim id_s As Integer
    Dim f_s As String
    
    If Me.ado_athlites.Recordset.RecordCount >= 1 Then
        'id_s = Me.ado_athlites.Recordset.Fields(0).Value
        id_s = Me.tmp_kod
        'f_s = "[id] LIKE " & id_s
        'Me.ado_athlites.Recordset.Filter = f_s
        'Me.ado_athlites.Recordset.Sort = "[id]"
        Me.ado_athlites.Recordset.MoveFirst
        Me.ado_athlites.Recordset.Find "[id] like '" & str(id_s) & "'", , adSearchForward
        If Not Me.ado_athlites.Recordset.EOF Then 'ΤΟ ΕΧΕΙΣ ΒΡΕΙ
            Me.canc_bt.Enabled = True
        Else 'ΔΕΝ ΤΟ ΒΡΕΙΣ
            Me.ado_athlites.Recordset.Requery
            Me.ado_athlites.Refresh
            Me.dt_athlites.Refresh
            Me.ado_athlites.Recordset.MoveFirst
            Me.dt_athlites.Columns(0).Caption = "Κωδικός"
            Me.dt_athlites.Columns(0).Width = 1000
            Me.dt_athlites.Columns(1).Caption = "Αριθμός ΚΟΕ"
            Me.dt_athlites.Columns(1).Width = 1200
            Me.dt_athlites.Columns(2).Caption = "Επώνυμο"
            Me.dt_athlites.Columns(2).Width = 2300
            Me.dt_athlites.Columns(3).Caption = "Όνομα"
            Me.dt_athlites.Columns(3).Width = 1800
            For i = 4 To 13
                Me.dt_athlites.Columns(i).Visible = False
            Next i
            Me.dt_athlites.Columns(14).Caption = "Ημ/νία Γέννησης"
            Me.dt_athlites.Columns(14).Width = 1400
            For i = 15 To Me.ado_athlites.Recordset.Fields.Count - 1
                Me.dt_athlites.Columns(i).Visible = False
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
        'Me.co_dimoi.BoundText = initial_string & final_string
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
        'Me.co_pe.BoundText = initial_string & final_string
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

Private Sub co_sxolia_Click(Area As Integer)

    initial_string = Left(Me.co_sxolia.Text, Me.co_sxolia.SelStart)
    final_string = Right(Me.co_sxolia.Text, Len(Me.co_sxolia.Text) - Me.co_sxolia.SelStart)

End Sub

Private Sub co_sxolia_GotFocus()
    
    initial_string = ""
    final_string = ""
    If Me.co_sxolia.Text <> "" Then
        initial_string = Left(Me.co_sxolia.Text, Me.co_sxolia.SelStart)
        final_string = Right(Me.co_sxolia.Text, Len(Me.co_sxolia.Text) - Me.co_sxolia.SelStart)
    End If
    
End Sub

Private Sub co_sxolia_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'If KeyCode = vbKeyTab And initial_string <> "" And final_string <> "" Then
    If KeyCode = vbKeyTab Then
        'Me.co_sxolia.BoundText = initial_string & final_string
        Me.co_sxolia.Text = initial_string & final_string
    End If
    
End Sub

Private Sub co_sxolia_KeyPress(KeyAscii As Integer)
        
    'SendKeys "{Esc}"
    'SendKeys "%{Up}"
    
    initial_string = Left(Me.co_sxolia.Text, Me.co_sxolia.SelStart)
    final_string = Right(Me.co_sxolia.Text, Len(Me.co_sxolia.Text) - Me.co_sxolia.SelStart)
    initial_string = initial_string & Chr(KeyAscii)
    
End Sub

Private Sub co_sxolia_KeyUp(KeyCode As Integer, Shift As Integer)
    
    initial_string = Left(Me.co_sxolia.Text, Me.co_sxolia.SelStart)
    final_string = Right(Me.co_sxolia.Text, Len(Me.co_sxolia.Text) - Me.co_sxolia.SelStart)
    
End Sub

Private Sub co_sxolia_LostFocus()
    
    'Me.co_sxolia.BoundText = initial_string & final_string
    Me.co_sxolia.Text = initial_string & final_string
    
End Sub

Private Sub Command1_Click()
        
        
    'Me.up_bt.Enabled = False
    'Me.del_bt.Enabled = False
    Me.cancel_cur_rec.Enabled = False
        
        
        for_search = 1
        Me.tmp_kod.Locked = False
        
        Me.tmp_kod = ""
        Me.tmp_ar_kar = ""
        tmp_am.Text = ""
        tmp_onoma.Text = ""
        tmp_eponimo.Text = ""
        tmp_odos.Text = ""
        tmp_arithmos.Text = ""
        tmp_perioxi.Text = ""
        Me.co_dimoi.Text = ""
        Me.co_pe.Text = ""
        tmp_tk.Text = ""
        tmp_til_oikias.Text = ""
        tmp_kinito.Text = ""
        tmp_fax.Text = ""
        tmp_email.Text = ""
        'Me.MaskEdBox1.Text = "00/00/0000"
        'Me.MaskEdBox2.Text = "00/00/0000"
        'Me.MaskEdBox3.Text = "00/00/0000"
        Me.MaskEdBox1.Text = "  /  /    "
        Me.MaskEdBox2.Text = "  /  /    "
        Me.MaskEdBox3.Text = "  /  /    "
        rs_ado_pateres.Filter = "[id] LIKE '" & str(-1) & "'"
        rs_ado_miteres.Filter = "[id] LIKE '" & str(-1) & "'"
        Me.co_sxolia.Text = ""
        Me.txt_parat.Text = ""
        Me.txt_aderf.Text = ""
        
        'Me.up_bt.Enabled = False
        'Me.save_command.Enabled = False
        Me.sear_bt.Enabled = True
        'Me.insert_bt.Enabled = False
        'Me.del_bt.Enabled = False
        Me.canc_bt.Enabled = True
                
End Sub

Private Sub Command2_Click()

    Dim f_s As String
    
    'If Me.ado_athlites.Recordset.RecordCount >= 1 And Me.tmp_eponimo.Text <> "" And it_is_a_new_record = 1 Then
    If Me.tmp_eponimo.Text <> "" And it_is_a_new_record = 1 Then
        f_s = "[Επώνυμο] LIKE '" & Trim(Me.tmp_eponimo.Text) & "*'"
        Me.ado_athlites.Recordset.Filter = f_s
        If Not Me.ado_athlites.Recordset.EOF Then 'ΤΟ ΕΧΕΙΣ ΒΡΕΙ
            Me.ado_athlites.Recordset.MoveFirst
            Me.dt_athlites.SetFocus
            Me.canc_bt.Enabled = True
            Me.Command1.Enabled = False
            'Me.del_bt.Enabled = True
        Else
            Me.dt_athlites.ReBind
            Me.dt_athlites.Columns(0).Caption = "Κωδικός"
            Me.dt_athlites.Columns(1).Width = 1500
            Me.dt_athlites.Columns(1).Caption = "Αριθμός ΚΟΕ"
            Me.dt_athlites.Columns(1).Width = 1500
            Me.dt_athlites.Columns(2).Caption = "Επώνυμο"
            Me.dt_athlites.Columns(2).Width = 3000
            Me.dt_athlites.Columns(3).Caption = "Όνομα"
            Me.dt_athlites.Columns(3).Width = 2500
            For i = 4 To Me.ado_athlites.Recordset.Fields.Count - 1
                Me.dt_athlites.Columns(i).Visible = False
            Next i
            Me.ado_athlites.Caption = "Αθλητής 0 από 0"
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
                        If athlet_management.ado_athlites.Recordset.RecordCount >= 1 Then
                            athlet_management.ado_athlites.Recordset.MoveFirst
                            sData = "ΑΜ Αθλητή" & vbTab & "Επώνυμο" & vbTab & "Όνομα" & vbTab & "Οδός" & vbTab & "Αριθμός" & vbTab & "Περιοχή" & vbTab & "Δήμος" & vbTab & "Περιφερειακή Ενότητα" & vbTab & "Ταχυδρομικός Κώδικας" & vbTab & "Τηλέφωνο Οικίας" & vbTab & "Κινητό Τηλέφωνο" & vbTab & "Αριθμός Fax" & vbTab & "Email" & vbTab & "Ημερομηνία Γέννησης" & vbTab & "Σχόλειο Φοίτησης" & vbTab & "Ημ. Έναρξης ΚΟΕ" & vbTab & "Ημ. Λήξης ΚΟΕ" & vbTab & "Αριθμός Ηλεκτρονικής Κάρτας" & vbTab & "Παρατηρήσεις" & vbCr
                            For i = 0 To athlet_management.ado_athlites.Recordset.RecordCount - 1
                                sData = sData & athlet_management.ado_athlites.Recordset.Fields(1) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(2) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(3) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(4) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(5) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(6) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(7) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(8) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(9) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(10) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(11) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(12) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(13) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(14) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(17) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(18) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(19) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(22) _
                                & vbTab & athlet_management.ado_athlites.Recordset.Fields(21) _
                                & vbCr
                                athlet_management.ado_athlites.Recordset.MoveNext
                            Next i
                        End If
                        Clipboard.SetText sData
                        oBook.Worksheets(1).Range("A1").Select
                        oBook.Worksheets(1).Paste
                        oSheet.Range("A1:Z1").Font.Bold = True
                        'oSheet.Range("A1:Z1").ColumnWidth = 20
                        oSheet.Range("A1:Z1").EntireColumn.AutoFit
                        oBook.Worksheets(1).Columns.HorizontalAlignment = -4131
                        oExcel.Visible = True
            
End Sub

Private Sub Command4_Click()

    'ΕΝΗΜΕΡΩΣΗ ΣΤΟΙΧΕΙΩΝ ΜΗΤΕΡΑΣ
    melos_id = 0
    athlet_management.ado_miteres.Refresh
    athlet_management.dt_mother.Columns(0).Visible = False
    athlet_management.dt_mother.Columns(1).Caption = "Α.Μ."
    athlet_management.dt_mother.Columns(1).Width = 1000
    athlet_management.dt_mother.Columns(2).Caption = "Επώνυμο"
    athlet_management.dt_mother.Columns(2).Width = 1500
    athlet_management.dt_mother.Columns(3).Caption = "Όνομα"
    athlet_management.dt_mother.Columns(3).Width = 1200
    Set athlet_management.rs_ado_miteres = athlet_management.ado_miteres.Recordset
    For i = 4 To athlet_management.rs_ado_miteres.Fields.Count - 1
        athlet_management.dt_mother.Columns(i).Visible = False
    Next i
    athlet_management.ado_miteres.Recordset.Filter = "[id] LIKE '" & str(melos_id) & "'"
    

End Sub

Private Sub del_bt_Click()
    
    Dim ms As String
    Dim rec_index As Integer

    If Me.ado_athlites.Recordset.RecordCount >= 1 Then
        ms = MsgBox("Είσαι σίγουρος; (ΝΑΙ ή ΟΧΙ)", vbYesNo, "Παράθυρο διαγραφής")
        If ms = 6 Then
        
        
            '
            Me.tmp_ado_athlites.Recordset.MoveFirst
            id_α = Me.ado_athlites.Recordset.Fields(0).Value
            Me.tmp_ado_athlites.Recordset.Find "[id] = " & id_α

            '
        
            is_to_delete = 1
            If Me.ado_athlites.Recordset.AbsolutePosition = Me.ado_athlites.Recordset.RecordCount Then
                rec_index = Me.ado_athlites.Recordset.AbsolutePosition - 1
            Else
                rec_index = Me.ado_athlites.Recordset.AbsolutePosition
            End If
            'ΔΙΑΓΡΑΦΗ ΚΑΙ ΤΗΣ ΦΩΤΟΓΡΑΦΙΑΣ, αν υπάρχει ΑΠΟ ΤΟ ΔΙΣΚΟ
            If Me.ado_athlites.Recordset.Fields(20).Value <> "" Then
                Kill Me.ado_athlites.Recordset.Fields(20).Value
            End If
            '
            'Me.ado_athlites.Recordset.Delete
            mv = Me.ado_athlites.Recordset.AbsolutePosition - 1
            If mv = Me.ado_athlites.Recordset.RecordCount - 1 Then 'ΔΗΛ. ΔΙΑΓΡΑΦΕΤΑΙ ΤΟ LAST
                mv = mv - 1
            End If
            Me.tmp_ado_athlites.Recordset.Delete
            Me.tmp_ado_athlites.Recordset.Requery
            Me.tmp_ado_athlites.Refresh
            Me.ado_athlites.Recordset.Requery
            Me.ado_athlites.Refresh
            
            
            '
            Me.ado_athlites.Recordset.Filter = MDIForm1.s_string
            If MDIForm1.s_sort <> "" Then
                Me.ado_athlites.Recordset.Sort = MDIForm1.s_sort
            Else
                Me.ado_athlites.Recordset.Sort = "[id]"
            End If
            If Me.ado_athlites.Recordset.RecordCount >= 1 Then
                
                    Me.ado_athlites.Recordset.MoveFirst
                    If mv - 1 >= 1 Then
                        Me.ado_athlites.Recordset.Move mv, 0
                    End If
                
            End If
    
            Me.dt_athlites.Columns(0).Caption = "Κωδικός"
            Me.dt_athlites.Columns(0).Width = 1000
            Me.dt_athlites.Columns(1).Caption = "Αριθμός ΚΟΕ"
            Me.dt_athlites.Columns(1).Width = 1200
            Me.dt_athlites.Columns(2).Caption = "Επώνυμο"
            Me.dt_athlites.Columns(2).Width = 2300
            Me.dt_athlites.Columns(3).Caption = "Όνομα"
            Me.dt_athlites.Columns(3).Width = 1800
            For i = 4 To 13
                Me.dt_athlites.Columns(i).Visible = False
            Next i
            Me.dt_athlites.Columns(14).Caption = "Ημ/νία Γέννησης"
            Me.dt_athlites.Columns(14).Width = 1400
            For i = 15 To Me.ado_athlites.Recordset.Fields.Count - 1
                Me.dt_athlites.Columns(i).Visible = False
            Next i
            '

            'Έχουν μείνει αθλητές μετά τη ΔΙΑΓΡΑΦΗ
            If Me.ado_athlites.Recordset.RecordCount >= 1 Then
            '
            'ΔΕΝ έχουν μείνει αθλητές μετά τη ΔΙΑΓΡΑΦΗ, άρα REQUERY
            Else
            '    'ΔΕΝ έχουν μείνει αθλητές μετά τη ΔΙΑΓΡΑΦΗ, άρα καθαρισμός όλων των πεδίων της φόρμας
                    Me.tmp_kod.Locked = True
                    Me.tmp_kod = ""
                    Me.tmp_ar_kar.Locked = True
                    Me.tmp_ar_kar = ""
                    Me.tmp_am.Locked = True
                    Me.tmp_am.Text = ""
                    Me.tmp_onoma.Locked = True
                    Me.tmp_onoma.Text = ""
                    Me.tmp_eponimo.Locked = True
                    Me.tmp_eponimo.Text = ""
                    Me.tmp_odos.Locked = True
                    Me.tmp_odos.Text = ""
                    Me.tmp_arithmos.Locked = True
                    Me.tmp_arithmos.Text = ""
                    Me.tmp_perioxi.Locked = True
                    Me.tmp_perioxi.Text = ""
                    Me.co_dimoi.Locked = True
                    Me.co_dimoi.Text = ""
                    Me.co_pe.Locked = True
                    Me.co_pe.Text = ""
                    Me.tmp_tk.Locked = True
                    Me.tmp_tk.Text = ""
                    Me.tmp_til_oikias.Locked = True
                    Me.tmp_til_oikias.Text = ""
                    Me.tmp_kinito.Locked = True
                    Me.tmp_kinito.Text = ""
                    Me.tmp_fax.Locked = True
                    Me.tmp_fax.Text = ""
                    Me.tmp_email.Locked = True
                    Me.tmp_email.Text = ""
                    Me.MaskEdBox1.Enabled = False
                    Me.MaskEdBox1.Text = "  /  /    "
                    Me.MaskEdBox2.Enabled = False
                    Me.MaskEdBox2.Text = "  /  /    "
                    Me.MaskEdBox3.Enabled = False
                    Me.MaskEdBox3.Text = "  /  /    "
                    Me.ado_pateres.Recordset.Filter = "[id] LIKE '" & str(-1) & "'"
                    Me.ado_miteres.Recordset.Filter = "[id] LIKE '" & str(-1) & "'"
                    Me.co_sxolia.Locked = True
                    Me.co_sxolia.Text = ""
                    Me.txt_parat.Locked = True
                    Me.txt_parat.Text = ""
                    Me.ado_athlites.Caption = "Αθλητής 0 από 0"
            End If
        End If
    Else
        MsgBox "Δεν υπάρχει αθλητής για ΔΙΑΓΡΑΦΗ!", vbOKOnly, "Μήνυμα Ενημέρωσης"
    End If
    
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub dt_athlites_HeadClick(ByVal ColIndex As Integer)
    
    defined_col = ColIndex
    
End Sub

Private Sub dt_athlites_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    
    If is_to_delete <> 1 Then
        If Me.ado_athlites.Recordset.AbsolutePosition >= 1 And Me.ado_athlites.Recordset.AbsolutePosition <= Me.ado_athlites.Recordset.RecordCount Then
            If Trim(Me.ado_athlites.Recordset.Fields(0).Value) <> "" And Me.it_is_a_new_record = 0 Then
                tmp_kod.Text = Trim(str(Me.ado_athlites.Recordset.Fields(0).Value))
            Else
                If Me.it_is_a_new_record = 1 Then
                Else
                    tmp_kod.Text = ""
                End If
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(22).Value) <> "" And Me.it_is_a_new_record = 0 Then
                tmp_ar_kar.Text = Trim(Me.ado_athlites.Recordset.Fields(22).Value)
            Else
                If Me.it_is_a_new_record = 1 Then
                Else
                    tmp_ar_kar.Text = ""
                End If
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(1).Value) <> "" Then
                tmp_am.Text = Me.ado_athlites.Recordset.Fields(1).Value
            Else
                tmp_am.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(18).Value) <> "" Then
                MaskEdBox2.Text = Me.ado_athlites.Recordset.Fields(18).Value
            Else
                MaskEdBox2.Text = "00/00/0000"
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(19).Value) <> "" Then
                MaskEdBox3.Text = Me.ado_athlites.Recordset.Fields(19).Value
            Else
                MaskEdBox3.Text = "00/00/0000"
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(3).Value) <> "" And Me.it_is_a_new_record = 0 Then
                tmp_onoma.Text = Me.ado_athlites.Recordset.Fields(3).Value
            Else
                If Me.it_is_a_new_record = 1 Then
                Else
                    tmp_onoma.Text = ""
                End If
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(2).Value) <> "" Then
                tmp_eponimo.Text = Me.ado_athlites.Recordset.Fields(2).Value
            Else
                tmp_eponimo.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(4).Value) <> "" Then
                tmp_odos.Text = Me.ado_athlites.Recordset.Fields(4).Value
            Else
                tmp_odos.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(5).Value) <> "" Then
                tmp_arithmos.Text = Me.ado_athlites.Recordset.Fields(5).Value
            Else
                tmp_arithmos.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(6).Value) <> "" Then
                tmp_perioxi.Text = Me.ado_athlites.Recordset.Fields(6).Value
            Else
                tmp_perioxi.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(8).Value) <> "" Then
                Me.co_pe.Text = Me.ado_athlites.Recordset.Fields(8).Value
            Else
                Me.co_pe.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(7).Value) <> "" Then
                Me.co_dimoi.Text = Me.ado_athlites.Recordset.Fields(7).Value
            Else
                Me.co_dimoi.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(9).Value) <> "" Then
                tmp_tk.Text = Me.ado_athlites.Recordset.Fields(9).Value
            Else
                tmp_tk.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(10).Value) <> "" Then
                tmp_til_oikias.Text = Me.ado_athlites.Recordset.Fields(10).Value
            Else
                tmp_til_oikias.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(11).Value) <> "" Then
                tmp_kinito.Text = Me.ado_athlites.Recordset.Fields(11).Value
            Else
                tmp_kinito.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(12).Value) <> "" Then
                tmp_fax.Text = Me.ado_athlites.Recordset.Fields(12).Value
            Else
                tmp_fax.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(13).Value) <> "" Then
                tmp_email.Text = Me.ado_athlites.Recordset.Fields(13).Value
            Else
                tmp_email.Text = ""
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(14).Value) <> "" Then
                Me.MaskEdBox1.Text = Me.ado_athlites.Recordset.Fields(14).Value
            Else
                Me.MaskEdBox1.Text = "00/00/0000"
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(15).Value) <> "" Then
                rs_ado_pateres.Filter = "[id] LIKE '" & str(Me.ado_athlites.Recordset.Fields(15).Value) & "'"
            Else
                rs_ado_pateres.Filter = "[id] LIKE '" & str(-1) & "'"
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(16).Value) <> "" Then
                rs_ado_miteres.Filter = "[id] LIKE '" & str(Me.ado_athlites.Recordset.Fields(16).Value) & "'"
            Else
                rs_ado_miteres.Filter = "[id] LIKE '" & str(-1) & "'"
            End If
            If Trim(Me.ado_athlites.Recordset.Fields(17).Value) <> "" Then
                Me.co_sxolia.Text = Me.ado_athlites.Recordset.Fields(17).Value
            Else
                Me.co_sxolia.Text = ""
            End If
            'ΠΑΡΟΥΣΙΑΣΗ ΦΩΤΟΓΡΑΦΙΑΣ
            Me.image_path.Text = ""
            Image1.Picture = LoadPicture()
            Label19.Visible = True
            'ΠΑΡΟΥΣΙΑΣΗ ΠΑΡΑΤΗΡΗΣΕΩΝ
            If Trim(Me.ado_athlites.Recordset.Fields(21).Value) <> "" Then
                Me.txt_parat.Text = Me.ado_athlites.Recordset.Fields(21).Value
            Else
                Me.txt_parat.Text = ""
            End If
            'ΠΛΗΘΟΣ ΑΔΕΡΦΙΩΝ ΠΟΥ ΕΙΝΑΙ ΗΔΗ ΕΓΓΕΓΡΑΜΜΕΝΑ
            If Trim(Me.ado_athlites.Recordset.Fields(23).Value) <> "" Then
                Me.txt_aderf.Text = Me.ado_athlites.Recordset.Fields(23).Value
            Else
                Me.txt_aderf.Text = 0
            End If
        End If
    End If

End Sub

Private Sub File1_Click()

    Dim img As String
    img = File1.Path & "\" & File1.FileName
    Image1.Picture = LoadPicture(img)
    Me.Label19.Visible = False

End Sub

Private Sub dt_father_DblClick()

    Dim id_s As Integer
    Dim f_s As String
    
    If Me.ado_pateres.Recordset.RecordCount >= 1 And Me.ado_athlites.Recordset.Fields(15).Value <> 0 Then
        id_s = Me.ado_pateres.Recordset.Fields(0).Value
        athlet_management.flag_pateras = 1
        athlet_management.flag_mitera = 0
        meli_management.met_st.Enabled = True
        meli_management.Show
        f_s = "[id] LIKE " & id_s
        meli_management.ado_meli.Recordset.Filter = f_s
        meli_management.canc_bt.Enabled = True
        meli_management.insert_bt.Enabled = False
        meli_management.del_bt.Enabled = False
        meli_management.Command1.Enabled = False
    Else
        athlet_management.flag_pateras = 1
        athlet_management.flag_mitera = 0
        meli_management.met_st.Enabled = True
        meli_management.Show
    End If

End Sub

Private Sub dt_mother_DblClick()

    Dim id_s As Integer
    Dim f_s As String
    
    If Me.ado_miteres.Recordset.RecordCount >= 1 And Me.ado_athlites.Recordset.Fields(16).Value <> 0 Then
        id_s = Me.ado_miteres.Recordset.Fields(0).Value
        athlet_management.flag_mitera = 1
        athlet_management.flag_pateras = 0
        meli_management.met_st.Enabled = True
        meli_management.Show
        f_s = "[id] LIKE " & id_s
        meli_management.ado_meli.Recordset.Filter = f_s
        meli_management.canc_bt.Enabled = True
        meli_management.insert_bt.Enabled = False
        meli_management.del_bt.Enabled = False
        meli_management.Command1.Enabled = False
    Else
        athlet_management.flag_mitera = 1
        athlet_management.flag_pateras = 0
        meli_management.met_st.Enabled = True
        meli_management.Show
    End If


End Sub

Private Sub Form_Load()
  
  
    it_is_a_new_record = 0
    is_to_delete = 0
    is_a_new_record_without_save = 0
    MDIForm1.s_string = ""
    MDIForm1.s_sort = ""
    MDIForm1.rep_lbl = ""
    
    Me.ado_dimoi.Refresh
    Set rs_ado_dimoi = ado_dimoi.Recordset
    Set Me.co_dimoi.RowSource = Me.ado_dimoi
    
    Me.ado_miteres.Refresh
    Set rs_ado_miteres = ado_miteres.Recordset
    Set Me.dt_mother.DataSource = Me.ado_miteres
    
    Me.ado_pateres.Refresh
    Set rs_ado_pateres = ado_pateres.Recordset
    Set Me.dt_father.DataSource = Me.ado_pateres
    
    Me.ado_pe.Refresh
    Set rs_ado_pe = ado_pe.Recordset
    Set Me.co_pe.RowSource = Me.ado_pe
    Me.ado_sxolia.Refresh
    Set rs_ado_sxolia = ado_sxolia.Recordset
    Set Me.co_sxolia.RowSource = Me.ado_sxolia
  
  
    Me.Top = 200
    Me.Left = 200
    Me.Height = 11300
    Me.Width = 11835
    flag_pateras = 0
    flag_mitera = 0
    
    for_search = 0
    Me.tmp_kod.Locked = True
    s_sort = ""
    
    Me.ado_athlites.Recordset.Sort = "[" & Trim(Me.ado_athlites.Recordset.Fields(0).Name) & "]"
    '******************************************************
    If Not Me.ado_athlites.Recordset.EOF Then
        Me.Label19.Visible = True
        Me.ado_athlites.Recordset.MoveFirst
        If Trim(Me.ado_athlites.Recordset.Fields(0).Value) <> "" Then
            tmp_kod.Text = Trim(str(Me.ado_athlites.Recordset.Fields(0).Value))
        Else
            tmp_kod.Text = ""
        End If
        'Αριθμός Ηλεκτρονικής Κάρτας
        If Trim(Me.ado_athlites.Recordset.Fields(22).Value) <> "" Then
            tmp_ar_kar.Text = Trim(str(Me.ado_athlites.Recordset.Fields(22).Value))
        Else
            tmp_ar_kar.Text = ""
        End If
        '
        'Πλήθος αδερφιών που είναι ήδη εγγεγραμμένα
        If Trim(Me.ado_athlites.Recordset.Fields(23).Value) <> "" Then
            txt_aderf.Text = Trim(str(Me.ado_athlites.Recordset.Fields(23).Value))
        Else
            txt_aderf.Text = 0
        End If
        '
        If Trim(Me.ado_athlites.Recordset.Fields(1).Value) <> "" Then
            tmp_am.Text = str(Me.ado_athlites.Recordset.Fields(1).Value)
        Else
            tmp_am.Text = ""
        End If
        If Trim(Me.ado_athlites.Recordset.Fields(18).Value) <> "" Then
            MaskEdBox2.Text = Me.ado_athlites.Recordset.Fields(18).Value
        Else
            MaskEdBox2.Text = "00/00/0000"
        End If
        If Trim(Me.ado_athlites.Recordset.Fields(19).Value) <> "" Then
            MaskEdBox3.Text = Me.ado_athlites.Recordset.Fields(19).Value
        Else
            MaskEdBox3.Text = "00/00/0000"
        End If
        If Trim(Me.ado_athlites.Recordset.Fields(3).Value) <> "" Then
            tmp_onoma.Text = Me.ado_athlites.Recordset.Fields(3).Value
        Else
            tmp_onoma.Text = ""
        End If
        If Trim(Me.ado_athlites.Recordset.Fields(2).Value) <> "" Then
            tmp_eponimo.Text = Me.ado_athlites.Recordset.Fields(2).Value
        Else
            tmp_eponimo.Text = ""
        End If
        If Trim(Me.ado_athlites.Recordset.Fields(4).Value) <> "" Then
            tmp_odos.Text = Me.ado_athlites.Recordset.Fields(4).Value
        Else
            tmp_odos.Text = ""
        End If
        If Trim(Me.ado_athlites.Recordset.Fields(5).Value) <> "" Then
            tmp_arithmos.Text = Me.ado_athlites.Recordset.Fields(5).Value
        Else
            tmp_arithmos.Text = ""
        End If
        If Trim(Me.ado_athlites.Recordset.Fields(6).Value) <> "" Then
            tmp_perioxi.Text = Me.ado_athlites.Recordset.Fields(6).Value
        Else
            tmp_perioxi.Text = ""
        End If
        If Trim(Me.ado_athlites.Recordset.Fields(8).Value) <> "" Then
            Me.co_pe.Text = Me.ado_athlites.Recordset.Fields(8).Value
        Else
            Me.co_pe.Text = ""
        End If
        If Trim(Me.ado_athlites.Recordset.Fields(7).Value) <> "" Then
            Me.co_dimoi.Text = Me.ado_athlites.Recordset.Fields(7).Value
        Else
            Me.co_dimoi.Text = ""
        End If
        If Trim(Me.ado_athlites.Recordset.Fields(9).Value) <> "" Then
            tmp_tk.Text = Me.ado_athlites.Recordset.Fields(9).Value
        Else
            tmp_tk.Text = ""
        End If
        If Trim(Me.ado_athlites.Recordset.Fields(10).Value) <> "" Then
            tmp_til_oikias.Text = Me.ado_athlites.Recordset.Fields(10).Value
        Else
            tmp_til_oikias.Text = ""
        End If
        If Trim(Me.ado_athlites.Recordset.Fields(11).Value) <> "" Then
            tmp_kinito.Text = Me.ado_athlites.Recordset.Fields(11).Value
        Else
            tmp_kinito.Text = ""
        End If
        If Trim(Me.ado_athlites.Recordset.Fields(12).Value) <> "" Then
            tmp_fax.Text = Me.ado_athlites.Recordset.Fields(12).Value
        Else
            tmp_fax.Text = ""
        End If
        If Trim(Me.ado_athlites.Recordset.Fields(13).Value) <> "" Then
            tmp_email.Text = Me.ado_athlites.Recordset.Fields(13).Value
        Else
            tmp_email.Text = ""
        End If
        If Trim(Me.ado_athlites.Recordset.Fields(14).Value) <> "" Then
            Me.MaskEdBox1.Text = Me.ado_athlites.Recordset.Fields(14).Value
        Else
            Me.MaskEdBox1.Text = "00/00/0000"
        End If
        'ΠΑΡΟΥΣΙΑΣΗ ΣΤΟΙΧΕΙΩΝ ΠΑΤΕΡΑ
        Me.dt_father.Columns(0).Visible = False
        Me.dt_father.Columns(1).Caption = "Α.Μ."
        Me.dt_father.Columns(1).Width = 1000
        Me.dt_father.Columns(2).Caption = "Επώνυμο"
        Me.dt_father.Columns(2).Width = 1500
        Me.dt_father.Columns(3).Caption = "Όνομα"
        Me.dt_father.Columns(3).Width = 1200
        For i = 4 To rs_ado_pateres.Fields.Count - 1
            Me.dt_father.Columns(i).Visible = False
        Next i
        If Trim(Me.ado_athlites.Recordset.Fields(15).Value) <> "" Then
            rs_ado_pateres.Filter = "[id] LIKE '" & str(Me.ado_athlites.Recordset.Fields(15).Value) & "'"
        Else
            rs_ado_pateres.Filter = "[id] LIKE '" & str(-1) & "'"
        End If
        '
        'ΠΑΡΟΥΣΙΑΣΗ ΣΤΟΙΧΕΙΩΝ ΜΗΤΕΡΑΣ
        Me.dt_mother.Columns(0).Visible = False
        Me.dt_mother.Columns(1).Caption = "A.Μ."
        Me.dt_mother.Columns(1).Width = 1000
        Me.dt_mother.Columns(2).Caption = "Επώνυμο"
        Me.dt_mother.Columns(2).Width = 1500
        Me.dt_mother.Columns(3).Caption = "Όνομα"
        Me.dt_mother.Columns(3).Width = 1200
        For i = 4 To rs_ado_miteres.Fields.Count - 1
            Me.dt_mother.Columns(i).Visible = False
        Next i
        If Trim(Me.ado_athlites.Recordset.Fields(16).Value) <> "" Then
            rs_ado_miteres.Filter = "[id] LIKE '" & str(Me.ado_athlites.Recordset.Fields(16).Value) & "'"
        Else
            rs_ado_miteres.Filter = "[id] LIKE '" & str(-1) & "'"
        End If
        '
        'ΠΑΡΟΥΣΙΑΣΗ ΣΤΟΙΧΕΩΝ ΣΧΟΛΕΙΟΥ
        If Trim(Me.ado_athlites.Recordset.Fields(17).Value) <> "" Then
            Me.co_sxolia.Text = Me.ado_athlites.Recordset.Fields(17).Value
        Else
            Me.co_sxolia.Text = ""
        End If
        'ΠΑΡΟΥΣΙΑΣΗ ΦΩΤΟΓΡΑΦΙΑΣ
        Me.Label19.Visible = True
        'ΠΑΡΟΥΣΙΑΣΗ ΠΑΡΑΤΗΡΗΣΕΩΝ
        If Trim(Me.ado_athlites.Recordset.Fields(21).Value) <> "" Then
            Me.txt_parat.Text = Me.ado_athlites.Recordset.Fields(21).Value
        Else
            Me.txt_parat.Text = ""
        End If
    'Δεν υπάρχουν ΑΘΛΗΤΕΣ για εμφάνιση
    Else
        Me.Label19.Visible = False
        Me.tmp_kod.Locked = True
        Me.tmp_ar_kar.Locked = True
        Me.tmp_am.Locked = True
        Me.tmp_onoma.Locked = True
        Me.tmp_eponimo.Locked = True
        Me.tmp_odos.Locked = True
        Me.tmp_arithmos.Locked = True
        Me.tmp_perioxi.Locked = True
        Me.co_dimoi.Locked = True
        Me.co_pe.Locked = True
        Me.tmp_tk.Locked = True
        Me.tmp_til_oikias.Locked = True
        Me.tmp_kinito.Locked = True
        Me.tmp_fax.Locked = True
        Me.tmp_email.Locked = True
        Me.MaskEdBox1.Enabled = False
        Me.MaskEdBox2.Enabled = False
        Me.MaskEdBox3.Enabled = False
        rs_ado_pateres.Filter = "[id] LIKE '" & str(-1) & "'"
        Me.sear_pat.Enabled = False
        rs_ado_miteres.Filter = "[id] LIKE '" & str(-1) & "'"
        Me.sear_moth.Enabled = False
        Me.co_sxolia.Locked = True
        Me.txt_parat.Locked = True
        Me.up_bt.Enabled = False
        Me.del_bt.Enabled = False
        Me.sear_bt.Enabled = False
        Me.Command1.Enabled = False
        Me.canc_bt.Enabled = False
    End If
    
    '***************************************************************
    If Me.ado_athlites.Recordset.RecordCount > 0 Then
        Me.dt_athlites.Row = 0
        Me.dt_athlites.Col = 1
    End If
    If Me.ado_athlites.Recordset.RecordCount > 0 Then
        Me.ado_athlites.Caption = "Αθλητής " & Me.dt_athlites.Row + 1 & " από " & Me.ado_athlites.Recordset.RecordCount
    End If
    Me.dt_athlites.Columns(0).Caption = "Κωδικός"
    Me.dt_athlites.Columns(0).Width = 1000
    Me.dt_athlites.Columns(1).Caption = "Αριθμός ΚΟΕ"
    Me.dt_athlites.Columns(1).Width = 1200
    Me.dt_athlites.Columns(2).Caption = "Επώνυμο"
    Me.dt_athlites.Columns(2).Width = 2300
    Me.dt_athlites.Columns(3).Caption = "Όνομα"
    Me.dt_athlites.Columns(3).Width = 1800
    For i = 4 To 13
        Me.dt_athlites.Columns(i).Visible = False
    Next i
    Me.dt_athlites.Columns(14).Caption = "Ημ/νία Γέννησης"
    Me.dt_athlites.Columns(14).Width = 1400
    For i = 15 To Me.ado_athlites.Recordset.Fields.Count - 1
        Me.dt_athlites.Columns(i).Visible = False
    Next i
    
End Sub

Private Sub Form_Unload(cancel As Integer)

    Set rs_ado_athlites = Nothing
    Set Me.dt_athlites.DataSource = Nothing
    Set rs_ado_dimoi = Nothing
    Set rs_ado_miteres = Nothing
    Set Me.dt_mother.DataSource = Nothing
    Set rs_ado_pateres = Nothing
    Set Me.dt_father.DataSource = Nothing
    Set rs_ado_pe = Nothing
    Set rs_ado_sxolia = Nothing
    Set Me.co_sxolia.RowSource = Nothing
    Set Me.co_dimoi.RowSource = Nothing
    Set Me.co_pe.RowSource = Nothing

End Sub

Private Sub fr_sav_but_Click()
    Frame10.Visible = False
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 And MDIForm1.is_a_new_record_without_save = 0 Then
        PopupMenu MDIForm1.mn1
    End If

End Sub

Private Sub insert_bt_Click()

    'Να βρω το υποψήφιο id_αθλητή
    Me.it_is_a_new_record = 1
    Me.Command2.Enabled = True
    Me.tmp_ado_athlites.Recordset.Sort = "[" & Trim(Me.ado_athlites.Recordset.Fields(0).Name) & "]"
    If Not Me.tmp_ado_athlites.Recordset.EOF Then
        Me.tmp_ado_athlites.Recordset.MoveLast
        id_α = Me.tmp_ado_athlites.Recordset.Fields(0).Value
        id_α = id_α + 1
    Else
        id_α = 1
    End If
    MDIForm1.is_a_new_record_without_save = 1
    Me.Label19.Visible = False
    'Καθαρισμός πεδίων
    Me.tmp_kod.Text = id_α
    Me.tmp_ar_kar.Text = ""
    Me.tmp_ar_kar.Locked = False
    Me.tmp_am.Text = ""
    Me.tmp_am.Locked = False
    Me.tmp_onoma.Text = ""
    Me.tmp_onoma.Locked = False
    Me.tmp_eponimo.Text = ""
    Me.tmp_eponimo.Locked = False
    Me.tmp_odos.Text = ""
    Me.tmp_odos.Locked = False
    Me.tmp_arithmos.Text = ""
    Me.tmp_arithmos.Locked = False
    Me.tmp_perioxi.Text = ""
    Me.tmp_perioxi.Locked = False
    Me.co_pe.Text = "Ιωαννίνων"
    Me.co_pe.Locked = False
    Me.co_dimoi.Text = "Ιωαννιτών"
    Me.co_dimoi.Locked = False
    Me.tmp_tk.Text = ""
    Me.tmp_tk.Locked = False
    Me.tmp_til_oikias.Text = ""
    Me.tmp_til_oikias.Locked = False
    Me.tmp_kinito.Text = ""
    Me.tmp_kinito.Locked = False
    Me.tmp_fax.Text = ""
    Me.tmp_fax.Locked = False
    Me.tmp_email.Text = ""
    Me.tmp_email.Locked = False
    Me.MaskEdBox1.Text = "  /  /    "
    Me.MaskEdBox1.Enabled = True
    Me.MaskEdBox2.Text = "  /  /    "
    Me.MaskEdBox2.Enabled = True
    Me.MaskEdBox3.Text = "  /  /    "
    Me.MaskEdBox3.Enabled = True
    rs_ado_pateres.Filter = "[id] LIKE '" & str(-1) & "'"
    Me.sear_pat.Enabled = True
    rs_ado_miteres.Filter = "[id] LIKE '" & str(-1) & "'"
    Me.sear_moth.Enabled = True
    Me.co_sxolia.Text = ""
    Me.co_sxolia.Locked = False
    Me.Image1.Picture = LoadPicture()
    Me.txt_aderf.Text = 0
    Me.txt_parat.Text = ""
    
    Me.tmp_ar_kar.SetFocus
    
    Me.insert_bt.Enabled = False
    Me.save_command.Enabled = True
    Me.canc_bt.Enabled = True
    Me.up_bt.Enabled = False
    Me.del_bt.Enabled = False
    Me.Command1.Enabled = False
    Me.cancel_cur_rec.Enabled = False
    
End Sub

Private Sub kl_bt_Click()

    it_is_a_new_record = 0
    Unload Me

End Sub



Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 And MDIForm1.is_a_new_record_without_save = 0 Then
        PopupMenu MDIForm1.mn1
    End If
    
End Sub

Private Sub MaskEdBox1_GotFocus()
    
    With MaskEdBox1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub MaskEdBox1_LostFocus()
    
    'If IsDate(Me.MaskEdBox1.Text) = False And for_search = 1 Then
    If IsDate(Me.MaskEdBox1.Text) = False Then
        Dim imera, year, minas As Variant
        Me.MaskEdBox1.SelStart = 0
        Me.MaskEdBox1.SelLength = 2
        imera = Me.MaskEdBox1.SelText
        If Not (imera >= 1 And imera <= 31) Then
            imera = 0
            Me.MaskEdBox1.SelText = "  "
        End If
        Me.MaskEdBox1.SelStart = 3
        Me.MaskEdBox1.SelLength = 2
        minas = Me.MaskEdBox1.SelText
        If Not (minas >= 1 And minas <= 12) Then
            minas = 0
            Me.MaskEdBox1.SelText = "  "
        End If
        Me.MaskEdBox1.SelStart = 6
        Me.MaskEdBox1.SelLength = 4
        year = Me.MaskEdBox1.SelText
        If Not (year >= 0) Then
            year = 0
            Me.MaskEdBox1.SelText = "  "
        End If
        Me.MaskEdBox1.SelStart = 0
        Me.MaskEdBox1.SelLength = 10
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
    'If IsDate(Me.MaskEdBox1.Text) = False And (Me.MaskEdBox1.Text <> "__/__/____") And (Me.MaskEdBox1.Text <> "  /  /    ") Then
    If st = Me.MaskEdBox1.Text Then
        If IsDate(Me.MaskEdBox1.Text) = False And (Me.MaskEdBox1.Text <> "__/__/____") And (Me.MaskEdBox1.Text <> "  /  /    ") And st <> "00/00/0000" Then
            MsgBox "Λάθος τιμή ημερομηνίας!", vbCritical, "Μήνυμα λάθους"
            Me.MaskEdBox1.SelStart = 0
            Me.MaskEdBox1.SelLength = 10
            Me.MaskEdBox1.SelText = "  /  /    "
            Me.MaskEdBox1.SetFocus
        End If
    End If
    
End Sub

Private Sub MaskEdBox2_GotFocus()

    With MaskEdBox2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub MaskEdBox2_LostFocus()
        
    With Me.MaskEdBox2
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

Private Sub MaskEdBox3_GotFocus()

    With MaskEdBox3
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub MaskEdBox3_LostFocus()
    
    With Me.MaskEdBox3
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

Private Sub popmn1_Click()
    
    f_foto_a.Show
    
End Sub

Private Sub popmn2_Click()

    'ΠΑΡΟΥΣΙΑΣΗ ΦΩΤΟΓΡΑΦΙΑΣ
    'If Me.ado_athlites.Recordset.Fields(20).ActualSize <> 0 Then
    If Me.ado_athlites.Recordset.Fields(20).Value <> "" Then
        'sFile = App.Path + "\" + "image1.jpg"
        'Dim bytes() As Byte
        'Dim num_blocks As Long
        'Dim left_over As Long
        'Dim block_num As Long
        'Open sFile For Binary As #1
        'file_length = LenB(Me.ado_athlites.Recordset.Fields(20))
        'num_blocks = 1
        'left_over = 0
        'Me.ado_athlites.Recordset.Move Me.ado_athlites.Recordset.AbsolutePosition - 1, 1
        'For block_num = 1 To num_blocks
        '    bytes = Me.ado_athlites.Recordset.Fields(20).GetChunk(file_length)
        '    Put #1, , bytes
        'Next block_num
        'If left_over > 0 Then
        '    bytes = Me.ado_athlites.Recordset.Fields(20).GetChunk(left_over)
        '    Put #1, , bytes
        'End If
        'Close #1
        Image1.Picture = LoadPicture(Me.ado_athlites.Recordset.Fields(20).Value)
        Me.image_path.Text = Me.ado_athlites.Recordset.Fields(20).Value
        Me.Label19.Visible = False
        'Kill sFile
    Else
        MsgBox "Δεν έχει αποθηκευτεί ΦΩΤΟΓΡΑΦΙΑ στον αθλητή / αθλήτρια!", , "Μήνυμα Ενημέρωσης!"
        'Image1.Picture = LoadPicture()
        Me.Label19.Visible = True
    End If

End Sub

Private Sub popmn3_Click()

    Dim ms As String
    
    ms = MsgBox("Είσαι σίγουρος; (ΝΑΙ ή ΟΧΙ)", vbYesNo, "Παράθυρο διαγραφής")
    If ms = 6 Then
        'ΔΙΑΓΡΑΦΗ ΦΩΤΟΓΡΑΦΙΑΣ
        If Me.ado_athlites.Recordset.Fields(20).ActualSize <> 0 Then
            rs_ado_athlites.Fields(20).AppendChunk ""
            rs_ado_athlites.UpdateBatch adAffectCurrent
            Me.Image1.Picture = LoadPicture()
            Me.Label19.Visible = True
        End If
    End If
    '
End Sub

Private Sub save_command_Click()

    it_is_a_new_record = 0
    Me.Command2.Enabled = False
    'Αποθήκευση στους αθλητές
    Me.tmp_ado_athlites.Recordset.AddNew
    Me.tmp_ado_athlites.Recordset.Fields(0).Value = id_α
    'Αποθήκευση ΛΟΙΠΩΝ ΣΤΟΙΧΕΙΑ
    If Trim(Me.tmp_ar_kar.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(22).Value = Me.tmp_ar_kar.Text
    End If
    If Trim(Me.tmp_onoma.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(3).Value = Me.tmp_onoma.Text
    End If
    If Trim(Me.tmp_eponimo.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(2).Value = Me.tmp_eponimo.Text
    End If
    If Trim(Me.tmp_odos.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(4).Value = Me.tmp_odos.Text
    End If
    If Trim(Me.tmp_arithmos.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(5).Value = Me.tmp_arithmos.Text
    End If
    If Trim(Me.tmp_perioxi.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(6).Value = Me.tmp_perioxi.Text
    End If
    If Trim(Me.co_dimoi.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(7).Value = Me.co_dimoi.Text
        Me.ado_dimoi.Recordset.Requery
        Me.ado_dimoi.Refresh
        Me.co_dimoi.ReFill
        Me.co_dimoi.Text = Me.tmp_ado_athlites.Recordset.Fields(7).Value
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
        Me.tmp_ado_athlites.Recordset.Fields(8).Value = Me.co_pe.Text
        Me.ado_pe.Recordset.Requery
        Me.ado_pe.Refresh
        Me.co_pe.ReFill
        Me.co_pe.Text = Me.tmp_ado_athlites.Recordset.Fields(8).Value
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
        Me.tmp_ado_athlites.Recordset.Fields(9).Value = Me.tmp_tk.Text
    End If
    If Trim(Me.tmp_til_oikias.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(10).Value = Me.tmp_til_oikias.Text
    End If
    If Trim(Me.tmp_kinito.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(11).Value = Me.tmp_kinito.Text
    End If
    If Trim(Me.tmp_fax.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(12).Value = Me.tmp_fax.Text
    End If
    If Trim(Me.tmp_email.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(13).Value = Me.tmp_email.Text
    End If
    If Trim(Me.MaskEdBox1.Text) <> "00/00/0000" Then
        Me.tmp_ado_athlites.Recordset.Fields(14).Value = Me.MaskEdBox1.Text
    End If
    If Trim(Me.tmp_am.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(1).Value = Me.tmp_am.Text
    End If
    If Trim(Me.MaskEdBox2.Text) <> "00/00/0000" Then
        Me.tmp_ado_athlites.Recordset.Fields(18).Value = Me.MaskEdBox2.Text
    End If
    If Trim(Me.MaskEdBox3.Text) <> "00/00/0000" Then
        Me.tmp_ado_athlites.Recordset.Fields(19).Value = Me.MaskEdBox3.Text
    End If
    If Me.dt_father.Row >= 0 Then
        Me.tmp_ado_athlites.Recordset.Fields(15).Value = rs_ado_pateres.Fields(0).Value
    End If
    If Me.dt_mother.Row >= 0 Then
        Me.tmp_ado_athlites.Recordset.Fields(16).Value = rs_ado_miteres.Fields(0).Value
    End If
    If Me.co_sxolia.Text <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(17).Value = Me.co_sxolia.Text
            If Me.co_sxolia.MatchedWithList = False Then
            'ΕΙΣΑΓΩΓΗ ΝΕΟΥ ΣΧΟΛΕΙΟΥ
                If Me.ado_sxolia.Recordset.RecordCount >= 1 Then
                    Me.ado_sxolia.Recordset.Sort = "[id_σχολείου]"
                    Me.ado_sxolia.Recordset.MoveLast
                    id = Me.ado_sxolia.Recordset![id_σχολείου]
                Else
                    id = 0
                End If
            Me.ado_sxolia.Recordset.AddNew
            Me.ado_sxolia.Recordset.Fields(0) = id + 1
            Me.ado_sxolia.Recordset.Fields(1) = Trim(Me.co_sxolia.Text)
            Me.ado_sxolia.Recordset.UpdateBatch adAffectCurrent
            Me.ado_sxolia.Recordset.Requery
        End If
    End If
    'ΑΠΟΘΗΚΕΥΣΗ ΦΩΤΟΓΡΑΦΙΑΣ
    If Me.image_path.Text <> "" Then
        Dim final_file, tp As String
        tp = Right$(Me.image_path.Text, 3)
        final_file = "ΦΩΤΟΓΡΑΦΙΕΣ\αθλητής" & Me.tmp_kod & "." & tp
        FileCopy Me.image_path.Text, final_file
        Me.tmp_ado_athlites.Recordset.Fields(20).Value = final_file
    End If
    'ΑΠΟΘΗΚΕΥΣΗ ΠΑΡΑΤΗΡΗΣΕΩΝ
    If Trim(Me.txt_parat.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(21).Value = Me.txt_parat.Text
    End If
    '
    'ΑΠΟΘΗΚΕΥΣΗ ΑΡΙΘΜΟΥ ΑΔΕΡΦΙΩΝ
    If Trim(Me.txt_aderf.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(23).Value = Me.txt_aderf.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(23).Value = 0
    End If
    '

    MDIForm1.is_a_new_record_without_save = 0
    Me.Label19.Visible = True
    Me.tmp_ado_athlites.Recordset.UpdateBatch adAffectCurrent
    Me.tmp_ado_athlites.Recordset.Requery
    Me.tmp_ado_athlites.Refresh
    Me.ado_athlites.Recordset.Requery
    Me.ado_athlites.Refresh
    'Me.ado_athlites.Recordset.Sort = "[" & Trim(ado_athlites.Recordset.Fields(0).Name) & "]"
    'Me.ado_athlites.Recordset.MoveLast
    Me.ado_athlites.Recordset.Filter = MDIForm1.s_string
    If MDIForm1.s_sort <> "" Then
        Me.ado_athlites.Recordset.Sort = MDIForm1.s_sort
    Else
        Me.ado_athlites.Recordset.Sort = "[id]"
    End If
    If Me.ado_athlites.Recordset.RecordCount >= 1 Then
        Me.ado_athlites.Recordset.Find "[id] = " & id_α
        If Not Me.ado_athlites.Recordset.EOF Then
            mv = Me.ado_athlites.Recordset.AbsolutePosition
            Me.ado_athlites.Recordset.MoveFirst
            Me.ado_athlites.Recordset.Move mv - 1
        Else
            Me.ado_athlites.Recordset.MoveFirst
        End If
    End If
    
    Me.dt_athlites.Columns(0).Caption = "Κωδικός"
    Me.dt_athlites.Columns(0).Width = 1000
    Me.dt_athlites.Columns(1).Caption = "Αριθμός ΚΟΕ"
    Me.dt_athlites.Columns(1).Width = 1200
    Me.dt_athlites.Columns(2).Caption = "Επώνυμο"
    Me.dt_athlites.Columns(2).Width = 2300
    Me.dt_athlites.Columns(3).Caption = "Όνομα"
    Me.dt_athlites.Columns(3).Width = 1800
    For i = 4 To 13
        Me.dt_athlites.Columns(i).Visible = False
    Next i
    Me.dt_athlites.Columns(14).Caption = "Ημ/νία Γέννησης"
    Me.dt_athlites.Columns(14).Width = 1400
    For i = 15 To Me.ado_athlites.Recordset.Fields.Count - 1
        Me.dt_athlites.Columns(i).Visible = False
    Next i
    
    Me.save_command.Enabled = False
    'Me.canc_bt.Enabled = False
    Me.insert_bt.Enabled = True
    Me.up_bt.Enabled = True
    Me.del_bt.Enabled = True
    Me.Command1.Enabled = True
    Me.cancel_cur_rec.Enabled = True
        
End Sub

Private Sub sear_bt_Click()

    Me.tmp_kod.Locked = True

    Me.up_bt.Enabled = True
    Me.del_bt.Enabled = True
    Me.cancel_cur_rec.Enabled = True
    
    Me.MaskEdBox1.Visible = True
    
    s_string = ""
    s_string2 = ""
    If Trim(Me.tmp_kod.Text) <> "" Then
        s_string = "[id] LIKE " & Trim(Me.tmp_kod.Text)
        s_string2 = "[id] LIKE " & Trim(Me.tmp_kod.Text)
    End If
    If Trim(Me.tmp_am.Text) <> "" Then
        s_string = "[ΑΜ_αθλητή] LIKE '*" & Trim(Me.tmp_am.Text) & "*'"
        s_string2 = "[ΑΜ_αθλητή] LIKE '%" & Trim(Me.tmp_am.Text) & "%'"
    End If
    'Αριθμός Ηλεκτρονικής Κάρτας
    If Trim(tmp_ar_kar.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΑριθμόςΗλεκτρονικήςΚάρτας] LIKE '*" & Trim(tmp_ar_kar.Text) & "*'"
            s_string2 = s_string2 & " AND [ΑριθμόςΗλεκτρονικήςΚάρτας] LIKE '%" & Trim(tmp_ar_kar.Text) & "%'"
        Else
            s_string = "[ΑριθμόςΗλεκτρονικήςΚάρτας] LIKE '*" & Trim(tmp_ar_kar.Text) & "*'"
            s_string2 = "[ΑριθμόςΗλεκτρονικήςΚάρτας] LIKE '%" & Trim(tmp_ar_kar.Text) & "%'"
        End If
    End If
    If Trim(tmp_eponimo.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Επώνυμο] LIKE '*" & Trim(tmp_eponimo.Text) & "*'"
            s_string2 = s_string2 & " AND [Επώνυμο] LIKE '%" & Trim(tmp_eponimo.Text) & "%'"
        Else
            s_string = "[Επώνυμο] LIKE '*" & Trim(tmp_eponimo.Text) & "*'"
            s_string2 = "[Επώνυμο] LIKE '%" & Trim(tmp_eponimo.Text) & "%'"
        End If
    End If
    If Trim(tmp_onoma.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Όνομα] LIKE '*" & Trim(tmp_onoma.Text) & "*'"
            s_string2 = s_string2 & " AND [Όνομα] LIKE '%" & Trim(tmp_onoma.Text) & "%'"
        Else
            s_string = "[Όνομα] LIKE '*" & Trim(tmp_onoma.Text) & "*'"
            s_string2 = "[Όνομα] LIKE '%" & Trim(tmp_onoma.Text) & "%'"
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
    '
    If Trim(tmp_til_oikias.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΤηλέφωνοΟικίας] LIKE '*" & Trim(tmp_til_oikias.Text) & "*'"
            s_string2 = s_string2 & " AND [ΤηλέφωνοΟικίας] LIKE '%" & Trim(tmp_til_oikias.Text) & "%'"
        Else
            s_string = "[ΤηλέφωνοΟικίας] LIKE '*" & Trim(tmp_til_oikias.Text) & Right(1, tmp_til_oikias.Text) & "*'"
            s_string2 = "[ΤηλέφωνοΟικίας] LIKE '%" & Trim(tmp_til_oikias.Text) & Right(1, tmp_til_oikias.Text) & "%'"
        End If
    End If
    If Trim(tmp_kinito.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΚινητόΤηλέφωνο] LIKE '*" & Trim(tmp_kinito.Text) & "*'"
            s_string2 = s_string2 & " AND [ΚινητόΤηλέφωνο] LIKE '%" & Trim(tmp_kinito.Text) & "%'"
        Else
            s_string = "[ΚινητόΤηλέφωνο] LIKE '*" & Trim(tmp_kinito.Text) & "*'"
            s_string2 = "[ΚινητόΤηλέφωνο] LIKE '%" & Trim(tmp_kinito.Text) & "%'"
        End If
    End If
    If Trim(tmp_fax.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΑριθμόςΦαξ] LIKE '*" & Trim(tmp_fax.Text) & "*'"
            s_string2 = s_string2 & " AND [ΑριθμόςΦαξ] LIKE '%" & Trim(tmp_fax.Text) & "%'"
        Else
            s_string = "[ΑριθμόςΦαξ] LIKE '*" & Trim(tmp_fax.Text) & "*'"
            s_string2 = "[ΑριθμόςΦαξ] LIKE '%" & Trim(tmp_fax.Text) & "%'"
        End If
    End If
    If Trim(tmp_email.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΌνομαEmail] LIKE '*" & Trim(tmp_email.Text) & "*'"
            s_string2 = s_string2 & " AND [ΌνομαEmail] LIKE '%" & Trim(tmp_email.Text) & "%'"
        Else
            s_string = "[ΌνομαEmail] LIKE '*" & Trim(tmp_email.Text) & "*'"
            s_string2 = "[ΌνομαEmail] LIKE '%" & Trim(tmp_email.Text) & "%'"
        End If
    End If
    If Trim(tmp_odos.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Οδός] LIKE '*" & Trim(tmp_odos.Text) & "*'"
            s_string2 = s_string2 & " AND [Οδός] LIKE '%" & Trim(tmp_odos.Text) & "%'"
        Else
            s_string = "[Οδός] LIKE '*" & Trim(tmp_odos.Text) & "*'"
            s_string2 = "[Οδός] LIKE '%" & Trim(tmp_odos.Text) & "%'"
        End If
    End If
    If Trim(tmp_arithmos.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Αριθμός] LIKE '*" & Trim(tmp_arithmos.Text) & "*'"
            s_string2 = s_string2 & " AND [Αριθμός] LIKE '%" & Trim(tmp_arithmos.Text) & "%'"
        Else
            s_string = "[Αριθμός] LIKE '*" & Trim(tmp_arithmos.Text) & "*'"
            s_string2 = "[Αριθμός] LIKE '%" & Trim(tmp_arithmos.Text) & "%'"
        End If
    End If
    If Trim(tmp_perioxi.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Περιοχή] LIKE '*" & Trim(tmp_perioxi.Text) & "*'"
            s_string2 = s_string2 & " AND [Περιοχή] LIKE '%" & Trim(tmp_perioxi.Text) & "%'"
        Else
            s_string = "[Περιοχή] LIKE '*" & Trim(tmp_perioxi.Text) & "*'"
            s_string2 = s_string2 & " AND [Περιοχή] LIKE '%" & Trim(tmp_perioxi.Text) & "%'"
        End If
    End If
    If Trim(co_dimoi.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Δήμος] LIKE '*" & Trim(co_dimoi.Text) & "*'"
            s_string2 = s_string2 & " AND [Δήμος] LIKE '%" & Trim(co_dimoi.Text) & "%'"
        Else
            s_string = "[Δήμος] LIKE '*" & Trim(co_dimoi.Text) & "*'"
            s_string2 = "[Δήμος] LIKE '%" & Trim(co_dimoi.Text) & "%'"
        End If
    End If
    If Trim(co_pe.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Περιφερειακή Ενότητα] LIKE '*" & Trim(co_pe.Text) & "*'"
            s_string2 = s_string2 & " AND [Περιφερειακή Ενότητα] LIKE '%" & Trim(co_pe.Text) & "%'"
        Else
            s_string = "[Περιφερειακή Ενότητα] LIKE '*" & Trim(co_pe.Text) & "*'"
            s_string2 = "[Περιφερειακή Ενότητα] LIKE '%" & Trim(co_pe.Text) & "%'"
        End If
    End If
    If Trim(tmp_tk.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Ταχυδρομικός Κώδικας] LIKE '*" & Trim(tmp_tk.Text) & "*'"
            s_string2 = s_string2 & " AND [Ταχυδρομικός Κώδικας] LIKE '%" & Trim(tmp_tk.Text) & "%'"
        Else
            s_string = "[Ταχυδρομικός Κώδικας] LIKE '*" & Trim(tmp_tk.Text) & "*'"
            s_string2 = "[Ταχυδρομικός Κώδικας] LIKE '%" & Trim(tmp_tk.Text) & "%'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΠΑΤΕΡΑ
    If Me.dt_father.Row >= 0 Then
        If s_string <> "" Then
            s_string = s_string & " AND [Id_Πατέρα] = " & Val(Me.dt_father.Columns(0).Value)
            s_string2 = s_string2 & " AND [Id_Πατέρα] = " & Val(Me.dt_father.Columns(0).Value)
        Else
            s_string = "[Id_Πατέρα] = " & Val(Me.dt_father.Columns(0).Value)
            s_string2 = "[Id_Πατέρα] = " & Val(Me.dt_father.Columns(0).Value)
        End If
    End If
    '
    'ΚΡΙΤΗΡΙΟ ΜΗΤΕΡΑΣ
    If Me.dt_mother.Row >= 0 Then
        If s_string <> "" Then
            s_string = s_string & " AND [Id_Μητέρας] = " & Val(Me.dt_mother.Columns(0).Value)
            s_string2 = s_string2 & " AND [Id_Μητέρας] = " & Val(Me.dt_mother.Columns(0).Value)
        Else
            s_string = "[Id_Μητέρας] = " & Val(Me.dt_mother.Columns(0).Value)
            s_string2 = "[Id_Μητέρας] = " & Val(Me.dt_mother.Columns(0).Value)
        End If
    End If
    ' ΚΡΙΤΗΡΙΟ ΣΧΟΛΕΙΟΥ
    If Trim(co_sxolia.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Σχολείο] LIKE '*" & Trim(co_sxolia.Text) & "*'"
            s_string2 = s_string2 & " AND [Σχολείο] LIKE '%" & Trim(co_sxolia.Text) & "%'"
        Else
            s_string = "[Σχολείο] LIKE '*" & Trim(co_sxolia.Text) & "*'"
            s_string2 = "[Σχολείο] LIKE '%" & Trim(co_sxolia.Text) & "%'"
        End If
    End If
    ' ΚΡΙΤΗΡΙΟ ΠΑΡΑΤΗΡΗΣΕΩΝ
    If Trim(txt_parat.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Παρατηρήσεις] LIKE '*" & Trim(txt_parat.Text) & "*'"
        Else
            s_string = "[Παρατηρήσεις] LIKE '*" & Trim(txt_parat.Text) & "*'"
        End If
    End If
    ' ΚΡΙΤΗΡΙΟ ΑΡΙΘΜΟΥ ΑΔΕΡΦΙΩΝ
    If Trim(txt_aderf.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Αδέρφια] = " & Trim(txt_aderf.Text)
        Else
            s_string = "[Αδέρφια] LIKE " & Trim(txt_aderf.Text)
        End If
    End If
    '

    If s_string <> "" Then
        Me.ado_athlites.Recordset.Filter = Trim(s_string)
        If Not Me.ado_athlites.Recordset.EOF Then
            Me.ado_athlites.Recordset.MoveFirst
        End If
    End If
    MDIForm1.s_string = s_string
    If s_sort <> "" Then
        Me.ado_athlites.Recordset.Sort = Trim(s_sort)
    End If
    
    Me.canc_bt.Enabled = True
    Me.up_bt.Enabled = True
    
End Sub

Private Sub sear_moth_Click()
    
    athlet_management.flag_mitera = 1
    athlet_management.flag_pateras = 0
    meli_management.met_st.Enabled = True
    meli_management.Show
    
End Sub

Private Sub sear_pat_Click()

    athlet_management.flag_pateras = 1
    athlet_management.flag_mitera = 0
    meli_management.met_st.Enabled = True
    meli_management.Show
    
End Sub

Private Sub TabStrip1_Click()

    If TabStrip1.SelectedItem.Index = 2 Then
        athl_tmima_management.Show
    End If

End Sub

Private Sub taksin_Click()

    If Me.ado_athlites.Recordset.RecordCount >= 1 Then
        If Me.dt_athlites.Col >= 0 Then
            Me.ado_athlites.Recordset.Sort = "[" & Trim(Me.ado_athlites.Recordset.Fields(dt_athlites.Col).Name) & "]"
            s_sort = "[" & Trim(Me.ado_athlites.Recordset.Fields(dt_athlites.Col).Name) & "]"
        Else
            Me.ado_athlites.Recordset.Sort = "[" & Trim(Me.ado_athlites.Recordset.Fields(defined_col).Name) & "]"
            s_sort = "[" & Trim(Me.ado_athlites.Recordset.Fields(defined_col).Name) & "]"
        End If
        Me.ado_athlites.Caption = "Αθλητής 1 από " & Me.ado_athlites.Recordset.RecordCount
        MDIForm1.s_sort = s_sort
        Me.canc_bt.Enabled = True
    End If

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

Private Sub up_bt_Click()

    Dim id As Integer
    Dim f_ole As Object
    
    
If Me.ado_athlites.Recordset.RecordCount >= 1 Then
Me.tmp_ado_athlites.Recordset.MoveFirst
id_α = Me.ado_athlites.Recordset.Fields(0).Value
Me.tmp_ado_athlites.Recordset.Find "[id] = " & id_α
If Not Me.tmp_ado_athlites.Recordset.EOF Then
    
    'ΕΝΗΜΕΡΩΣΗ ΑΡΙΘΜΟΥ ΚΑΡΤΑΣ
    If Trim(Me.tmp_ar_kar.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(22).Value = Me.tmp_ar_kar.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(22).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ ΟΝΟΜΑΤΟΣ
    If Trim(Me.tmp_onoma.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(3).Value = Me.tmp_onoma.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(3).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ ΕΠΩΝΥΜΟΥ
    If Trim(Me.tmp_eponimo.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(2).Value = Me.tmp_eponimo.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(2).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ ΟΔΟΥ ΔΙΕΥΘΥΝΣΗΣ
    If Trim(Me.tmp_odos.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(4).Value = Me.tmp_odos.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(4).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ ΑΡΙΘΜΟΥ ΔΙΕΥΘΥΝΣΗΣ
    If Trim(Me.tmp_arithmos.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(5).Value = Me.tmp_arithmos.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(5).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ ΠΕΡΙΟΧΗΣ ΔΙΕΥΘΥΝΣΗΣ
    If Trim(Me.tmp_perioxi.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(6).Value = Me.tmp_perioxi.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(6).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ ΔΗΜΟΥ
    If Me.co_dimoi.Text <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(7).Value = Me.co_dimoi.Text
            Me.ado_dimoi.Recordset.Requery
            Me.ado_dimoi.Refresh
            Me.co_dimoi.ReFill
            Me.co_dimoi.Text = Trim(Me.tmp_ado_athlites.Recordset.Fields(7).Value)
            If Me.co_dimoi.MatchedWithList = False Then
            'ΕΙΣΑΓΩΓΗ ΝΕΟΥ ΔΗΜΟΥ
                If Me.ado_dimoi.Recordset.RecordCount >= 1 Then
                    Me.ado_dimoi.Recordset.Sort = "id_δήμου"
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
        Me.tmp_ado_athlites.Recordset.Fields(7).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ ΠΕΡΙΦΕΡΕΙΑΚΗΣ ΕΝΟΤΗΤΑΣ
    If Trim(Me.co_pe.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(8).Value = Me.co_pe.Text
        Me.ado_pe.Recordset.Requery
        Me.ado_pe.Refresh
        Me.co_pe.ReFill
        Me.co_pe.Text = Me.ado_athlites.Recordset.Fields(8).Value
        If Me.co_pe.MatchedWithList = False Then
        'ΕΙΣΑΓΩΓΗ ΝΕΑΣ ΠΕΡΙΦΕΡΕΙΑΚΗΣ ΕΝΟΤΗΤΑΣ
            Me.ado_pe.Recordset.Sort = "[" & Trim(Me.ado_pe.Recordset.Fields(0).Name) & "]"
            If Me.ado_pe.Recordset.RecordCount >= 1 Then
                Me.ado_pe.Recordset.Sort = "id_πε"
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
        Me.tmp_ado_athlites.Recordset.Fields(8).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ ΤΑΧΥΔΡΟΜΙΚΟΥ ΚΩΔΙΚΑ
    If Trim(Me.tmp_tk.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(9).Value = Me.tmp_tk.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(9).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ ΤΗΛΕΦΩΝΟΥ ΟΙΚΙΑΣ
    If Trim(Me.tmp_til_oikias.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(10).Value = Me.tmp_til_oikias.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(10).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ ΚΙΝΙΤΟΥ ΤΗΛΕΦΩΝΟΥ
    If Trim(Me.tmp_kinito.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(11).Value = Me.tmp_kinito.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(11).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ FAX
    If Trim(Me.tmp_fax.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(12).Value = Me.tmp_fax.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(12).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ EMAIL
    If Trim(Me.tmp_email.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(13).Value = Me.tmp_email.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(13).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ ΗΜΕΡΟΜΗΝΙΑΣ ΓΕΝΝΗΣΗΣ
    If Trim(Me.MaskEdBox1.Text) <> "00/00/0000" Then
        Me.tmp_ado_athlites.Recordset.Fields(14).Value = Me.MaskEdBox1.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(14).Value = "  /  /    "
    End If
    'ΕΝΗΜΕΡΩΣΗ ΑΡΙΘΜΟΥ ΜΗΤΡΩΟΥ ΑΘΛΗΤΗ
    If Trim(Me.tmp_am.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(1).Value = Me.tmp_am.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(1).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ ΗΜΕΡΟΜΗΝΙΑΣ ΕΓΓΡΑΦΗΣ ΣΤΗΝ ΚΟΕ
    If Trim(Me.MaskEdBox2.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(18).Value = Me.MaskEdBox2.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(18).Value = "  /  /    "
    End If
    'ΕΝΗΜΕΡΩΣΗ ΗΜΕΡΟΜΗΝΙΑΣ ΛΗΞΗΣ ΣΤΗΝ ΚΟΕ
    If Trim(Me.MaskEdBox3.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(19).Value = Me.MaskEdBox3.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(19).Value = "  /  /    "
    End If
    'ΕΝΗΜΕΡΩΣΗ ΣΤΟΙΧΕΙΩΝ ΠΑΤΕΡΑ
    If Me.dt_father.Row >= 0 Then
        Me.tmp_ado_athlites.Recordset.Fields(15).Value = rs_ado_pateres.Fields(0).Value
    Else
        Me.tmp_ado_athlites.Recordset.Fields(15).Value = 0 'ΔΕΝ ΥΠΑΡΧΕΙ ΠΑΤΕΡΑΣ
    End If
    'ΕΝΗΜΕΡΩΣΗ ΣΤΟΙΧΕΙΩΝ ΜΗΤΕΡΑΣ
    If Me.dt_mother.Row >= 0 Then
        Me.tmp_ado_athlites.Recordset.Fields(16).Value = rs_ado_miteres.Fields(0).Value
    Else
        Me.tmp_ado_athlites.Recordset.Fields(16).Value = 0 'ΔΕΝ ΥΠΑΡΧΕΙ ΜΗΤΕΡΑ
    End If
    'ΕΝΗΜΕΡΩΣΗ ΣΧΟΛΕΙΟΥ ΑΘΛΗΤΗ
    If Me.co_sxolia.Text <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(17).Value = Me.co_sxolia.Text
            If Me.co_sxolia.MatchedWithList = False Then
            'ΕΙΣΑΓΩΓΗ ΝΕΟΥ ΣΧΟΛΕΙΟΥ
                If Me.ado_sxolia.Recordset.RecordCount >= 1 Then
                    Me.ado_sxolia.Recordset.Sort = "id_σχολείου"
                    Me.ado_sxolia.Recordset.MoveLast
                    id = Me.ado_sxolia.Recordset![id_σχολείου]
                Else
                    id = 0
                End If
            Me.ado_sxolia.Recordset.AddNew
            Me.ado_sxolia.Recordset.Fields(0) = id + 1
            Me.ado_sxolia.Recordset.Fields(1) = Trim(Me.co_sxolia.Text)
            Me.ado_sxolia.Recordset.UpdateBatch adAffectCurrent
            Me.ado_sxolia.Recordset.Requery
        End If
    Else
        Me.tmp_ado_athlites.Recordset.Fields(17).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ ΦΩΤΟΓΡΑΦΙΑΣ
    If Me.image_path.Text <> "" Then
        Dim final_file, tp As String
        tp = Right$(Me.image_path.Text, 3)
        final_file = "ΦΩΤΟΓΡΑΦΙΕΣ\αθλητής" & Me.tmp_kod & "." & tp
        If Trim(Me.image_path.Text) <> Trim(final_file) Then
            FileCopy Me.image_path.Text, final_file
            Me.tmp_ado_athlites.Recordset.Fields(20).Value = final_file
        End If
        Me.image_path.Text = final_file
    End If
    'ΕΝΗΜΕΡΩΣΗ ΠΑΡΑΤΗΡΗΣΕΩΝ
    If Trim(Me.txt_parat.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(21).Value = Me.txt_parat.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(21).Value = Null
    End If
    'ΕΝΗΜΕΡΩΣΗ ΑΡΙΘΜΟΥ ΑΔΕΡΦΙΩΝ
    If Trim(Me.txt_aderf.Text) <> "" Then
        Me.tmp_ado_athlites.Recordset.Fields(23).Value = Me.txt_aderf.Text
    Else
        Me.tmp_ado_athlites.Recordset.Fields(23).Value = 0
    End If
    '
    
    Me.tmp_ado_athlites.Recordset.UpdateBatch adAffectCurrent
    Me.tmp_ado_athlites.Recordset.Requery
    Me.tmp_ado_athlites.Refresh
    Me.ado_athlites.Recordset.Requery
    Me.ado_athlites.Refresh
    Me.ado_athlites.Recordset.Filter = MDIForm1.s_string
    Me.ado_athlites.Recordset.Sort = MDIForm1.s_sort
    If Me.ado_athlites.Recordset.RecordCount >= 1 Then
        Me.ado_athlites.Recordset.Find "[id] = " & id_α
        If Not Me.ado_athlites.Recordset.EOF Then
            mv = Me.ado_athlites.Recordset.AbsolutePosition
            Me.ado_athlites.Recordset.MoveFirst
            Me.ado_athlites.Recordset.Move mv - 1
        Else
            Me.ado_athlites.Recordset.MoveFirst
        End If
    End If
    
    Me.dt_athlites.Columns(0).Caption = "Κωδικός"
    Me.dt_athlites.Columns(0).Width = 1000
    Me.dt_athlites.Columns(1).Caption = "Αριθμός ΚΟΕ"
    Me.dt_athlites.Columns(1).Width = 1200
    Me.dt_athlites.Columns(2).Caption = "Επώνυμο"
    Me.dt_athlites.Columns(2).Width = 2300
    Me.dt_athlites.Columns(3).Caption = "Όνομα"
    Me.dt_athlites.Columns(3).Width = 1800
    For i = 4 To 13
        Me.dt_athlites.Columns(i).Visible = False
    Next i
    Me.dt_athlites.Columns(14).Caption = "Ημ/νία Γέννησης"
    Me.dt_athlites.Columns(14).Width = 1400
    For i = 15 To Me.ado_athlites.Recordset.Fields.Count - 1
        Me.dt_athlites.Columns(i).Visible = False
    Next i
    '
    'AYTO ΕΛΕΧΞΕ ΚΑΠΟΙΑ ΣΤΙΓΜΗ ΑΝ ΧΡΕΙΑΖΕΤΑΙ
    'tmima_management.dt_tmimata.Refresh
    
End If
End If
    
End Sub

