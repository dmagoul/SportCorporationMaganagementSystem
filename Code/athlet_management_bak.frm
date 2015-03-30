VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form athlet_management 
   BackColor       =   &H80000014&
   ClientHeight    =   9870
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11970
   LinkTopic       =   "entity_management"
   MDIChild        =   -1  'True
   ScaleHeight     =   9870
   ScaleWidth      =   11970
   Begin MSAdodcLib.Adodc ado_dimoi 
      Height          =   375
      Left            =   10560
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
      Connect         =   $"athlet_management.frx":0000
      OLEDBString     =   $"athlet_management.frx":00AD
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
      Left            =   10560
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
      Connect         =   $"athlet_management.frx":015A
      OLEDBString     =   $"athlet_management.frx":0207
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
   Begin MSAdodcLib.Adodc ado_pateres 
      Height          =   375
      Left            =   10560
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
      Connect         =   $"athlet_management.frx":02B4
      OLEDBString     =   $"athlet_management.frx":0361
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
      Left            =   10560
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
      Connect         =   $"athlet_management.frx":040E
      OLEDBString     =   $"athlet_management.frx":04BB
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
      Left            =   10560
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
      Connect         =   $"athlet_management.frx":0568
      OLEDBString     =   $"athlet_management.frx":0615
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Σχολεία"
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
      Caption         =   "Αναζήτηση"
      DisabledPicture =   "athlet_management.frx":06C2
      Enabled         =   0   'False
      Height          =   495
      Left            =   7320
      MaskColor       =   &H80000014&
      Picture         =   "athlet_management.frx":1035
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton del_bt 
      BackColor       =   &H80000014&
      Caption         =   "Διαγραφή"
      DisabledPicture =   "athlet_management.frx":114A
      Height          =   495
      Left            =   4440
      MaskColor       =   &H80000014&
      Picture         =   "athlet_management.frx":1ABD
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000014&
      Caption         =   "Καθαρισμός"
      DisabledPicture =   "athlet_management.frx":1B7A
      Height          =   495
      Left            =   5880
      MaskColor       =   &H80000014&
      Picture         =   "athlet_management.frx":1CC2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton up_bt 
      BackColor       =   &H80000014&
      Caption         =   "Ενημέρωση"
      DisabledPicture =   "athlet_management.frx":1D7B
      Height          =   495
      Left            =   1560
      MaskColor       =   &H80000014&
      Picture         =   "athlet_management.frx":26EE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton kl_bt 
      BackColor       =   &H80000014&
      Caption         =   "Κλείσιμο"
      DisabledPicture =   "athlet_management.frx":2B55
      Height          =   495
      Left            =   10200
      MaskColor       =   &H80000014&
      Picture         =   "athlet_management.frx":34C8
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton canc_bt 
      BackColor       =   &H80000014&
      Caption         =   "Ακύρωση"
      DisabledPicture =   "athlet_management.frx":3E3B
      Height          =   495
      Left            =   8760
      MaskColor       =   &H80000014&
      Picture         =   "athlet_management.frx":3F83
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton save_command 
      BackColor       =   &H80000014&
      Caption         =   "Αποθήκευση"
      DisabledPicture =   "athlet_management.frx":4226
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      MaskColor       =   &H80000014&
      Picture         =   "athlet_management.frx":436E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton insert_bt 
      BackColor       =   &H80000014&
      Caption         =   "Προσθήκη"
      DisabledPicture =   "athlet_management.frx":4433
      Height          =   495
      Left            =   120
      MaskColor       =   &H80000014&
      Picture         =   "athlet_management.frx":4551
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ΟΛΟΙ ΟΙ ΑΘΛΗΤΕΣ"
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   11655
      Begin VB.CommandButton taksin 
         BackColor       =   &H80000014&
         Caption         =   "Ταξινόμηση"
         DisabledPicture =   "athlet_management.frx":4677
         Height          =   495
         Left            =   120
         MaskColor       =   &H80000014&
         Picture         =   "athlet_management.frx":4FEA
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3720
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dt_athlites 
         Bindings        =   "athlet_management.frx":50A7
         Height          =   3015
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5318
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
      Begin MSAdodcLib.Adodc ado_athlites 
         Height          =   375
         Left            =   120
         Top             =   3240
         Width           =   8535
         _ExtentX        =   15055
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
         Connect         =   $"athlet_management.frx":50C2
         OLEDBString     =   $"athlet_management.frx":516F
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Αθλητές"
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
      Begin VB.Image Image1 
         Height          =   3375
         Left            =   8760
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Στοιχεία Αθλητή"
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
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   11775
      Begin VB.Frame Frame9 
         Caption         =   "Στοιχεία ΚΟΕ"
         Height          =   735
         Left            =   4440
         TabIndex        =   52
         Top             =   240
         Width           =   7215
         Begin VB.TextBox tmp_am 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   53
            Top             =   240
            Width           =   1575
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   375
            Left            =   3600
            TabIndex        =   55
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   375
            Left            =   5880
            TabIndex        =   57
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ημ/νία Λήξης"
            Height          =   375
            Left            =   4800
            TabIndex        =   58
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ημ/νία Έναρξης"
            Height          =   375
            Left            =   2400
            TabIndex        =   56
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Αριθμός ΚΟΕ"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   765
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Στοιχεία Γονέων"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1935
         Left            =   4440
         TabIndex        =   40
         Top             =   2880
         Width           =   5895
         Begin VB.Frame Frame8 
            Caption         =   "Μητέρα"
            Height          =   735
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   5655
            Begin VB.CommandButton sear_moth 
               BackColor       =   &H80000014&
               Caption         =   "Αναζήτηση"
               DisabledPicture =   "athlet_management.frx":521C
               Height          =   495
               Left            =   4560
               MaskColor       =   &H80000014&
               Picture         =   "athlet_management.frx":5B8F
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   240
               Width           =   975
            End
            Begin MSDataGridLib.DataGrid dt_mother 
               Bindings        =   "athlet_management.frx":5CA4
               Height          =   495
               Left            =   120
               TabIndex        =   46
               Top             =   240
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   873
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
         Begin VB.Frame Frame6 
            Caption         =   "Πατέρας"
            Height          =   735
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   5655
            Begin VB.CommandButton sear_pat 
               BackColor       =   &H80000014&
               Caption         =   "Αναζήτηση"
               DisabledPicture =   "athlet_management.frx":5CBE
               Height          =   495
               Left            =   4560
               MaskColor       =   &H80000014&
               Picture         =   "athlet_management.frx":6631
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   240
               Width           =   975
            End
            Begin MSDataGridLib.DataGrid dt_father 
               Bindings        =   "athlet_management.frx":6746
               Height          =   495
               Left            =   120
               TabIndex        =   42
               Top             =   240
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   873
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
      End
      Begin VB.Frame Frame5 
         Caption         =   "Στοιχεία Διεύθυνσης"
         Height          =   1695
         Left            =   4440
         TabIndex        =   27
         Top             =   1080
         Width           =   7215
         Begin VB.TextBox tmp_tk 
            Height          =   375
            Left            =   5880
            TabIndex        =   31
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox tmp_perioxi 
            Height          =   375
            Left            =   1200
            TabIndex        =   30
            Top             =   720
            Width           =   3975
         End
         Begin VB.TextBox tmp_arithmos 
            Height          =   375
            Left            =   6240
            TabIndex        =   29
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox tmp_odos 
            Height          =   375
            Left            =   1200
            TabIndex        =   28
            Top             =   240
            Width           =   3975
         End
         Begin MSDataListLib.DataCombo co_pe 
            Bindings        =   "athlet_management.frx":6760
            Height          =   315
            Left            =   4920
            TabIndex        =   38
            Top             =   1200
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "περιγραφή"
            BoundColumn     =   "id_πε"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo co_dimoi 
            Bindings        =   "athlet_management.frx":6775
            Height          =   315
            Left            =   1200
            TabIndex        =   39
            Top             =   1200
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "περιγραφή"
            BoundColumn     =   "id_δήμου"
            Text            =   ""
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Οδός"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   945
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Αριθμός"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5280
            TabIndex        =   36
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Περιοχή"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   35
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Δήμος"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   1200
            Width           =   825
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Π. Ε."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            TabIndex        =   33
            Top             =   1200
            Width           =   825
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Τ.Κ."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4440
            TabIndex        =   32
            Top             =   720
            Width           =   1305
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Στοιχεία Επικοινωνίας"
         Height          =   1935
         Left            =   240
         TabIndex        =   18
         Top             =   2880
         Width           =   4095
         Begin VB.TextBox tmp_email 
            Height          =   375
            Left            =   1800
            TabIndex        =   22
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox tmp_fax 
            Height          =   375
            Left            =   1800
            TabIndex        =   21
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox tmp_kinito 
            Height          =   375
            Left            =   1800
            TabIndex        =   20
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox tmp_til_oikias 
            Height          =   375
            Left            =   1800
            TabIndex        =   19
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Τηλ.οικίας"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Κινητό"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   25
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   24
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Top             =   1440
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Γενικά Στοιχεία"
         Height          =   2415
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   4095
         Begin VB.TextBox tmp_ap 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   50
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox tmp_onoma 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   13
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox tmp_eponimo 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   12
            Top             =   1080
            Width           =   2535
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   375
            Left            =   1440
            TabIndex        =   16
            Top             =   1560
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo co_sxolia 
            Bindings        =   "athlet_management.frx":678D
            Height          =   315
            Left            =   1440
            TabIndex        =   47
            Top             =   1920
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "περιγραφή"
            BoundColumn     =   "id_δήμου"
            Text            =   ""
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Αριθμός Ποσειδώνα"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Σχολείο"
            Height          =   375
            Left            =   -120
            TabIndex        =   48
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ημ/νία Γέν."
            Height          =   375
            Left            =   -120
            TabIndex        =   17
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Όνομα"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   15
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Επώνυμο"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   14
            Top             =   1080
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
Public id_α, flag_pateras, flag_mitera, defined_col As Integer

Private Sub ado_athlites_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    'On Error Resume Next

    If Not adReason = 10 And pRecordset.AbsolutePosition >= 1 Then

    If Trim(pRecordset.Fields(1).Value) <> "" Then
        tmp_am.Text = pRecordset.Fields(1).Value
    Else
        tmp_am.Text = ""
    End If
    If Trim(pRecordset.Fields(3).Value) <> "" Then
        tmp_onoma.Text = pRecordset.Fields(3).Value
    Else
        tmp_onoma.Text = ""
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
    If Trim(Me.tmp_am.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(1).Value = Me.tmp_am.Text
    End If
    If Trim(Me.ado_athlites.Recordset.Fields(15).Value) <> "" Then
        Me.ado_pateres.Recordset.Filter = "[id] LIKE '" & Str(Me.ado_athlites.Recordset.Fields(15).Value) & "'"
    Else
        Me.ado_pateres.Recordset.Filter = "[id] LIKE '" & Str(-1) & "'"
    End If
    If Trim(Me.ado_athlites.Recordset.Fields(16).Value) <> "" Then
        Me.ado_miteres.Recordset.Filter = "[id] LIKE '" & Str(Me.ado_athlites.Recordset.Fields(16).Value) & "'"
    Else
        Me.ado_miteres.Recordset.Filter = "[id] LIKE '" & Str(-1) & "'"
    End If
    If Trim(pRecordset.Fields(17).Value) <> "" Then
        Me.co_sxolia.Text = pRecordset.Fields(17).Value
    Else
        Me.co_sxolia.Text = ""
    End If
    
   End If
   
       
    Me.ado_athlites.Caption = "Αθλητής " & pRecordset.AbsolutePosition & " από " & pRecordset.RecordCount
    
    'Err.Clear

End Sub

Private Sub canc_bt_Click()

    Me.insert_bt.Enabled = True
    Me.up_bt.Enabled = True
    Me.sear_bt.Enabled = False
    
    Me.ado_athlites.Refresh
    Me.ado_athlites.Recordset.MoveFirst
    
    Me.dt_athlites.Columns(0).Visible = False
    Me.dt_athlites.Columns(1).Caption = "Αριθμός Μητρώου"
    Me.dt_athlites.Columns(1).Width = 2000
    Me.dt_athlites.Columns(2).Caption = "Επώνυμο"
    Me.dt_athlites.Columns(2).Width = 4000
    Me.dt_athlites.Columns(3).Caption = "Όνομα"
    Me.dt_athlites.Columns(3).Width = 4000
    For i = 4 To Me.ado_athlites.Recordset.Fields.Count - 1
        Me.dt_athlites.Columns(i).Visible = False
    Next i

End Sub

Private Sub Command1_Click()

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
        Me.MaskEdBox1.Text = "00/00/0000"
        Me.ado_pateres.Recordset.Filter = "[id] LIKE '" & Str(-1) & "'"
        Me.ado_miteres.Recordset.Filter = "[id] LIKE '" & Str(-1) & "'"
        Me.co_sxolia.Text = ""
        
        'Me.up_bt.Enabled = False
        Me.save_command.Enabled = False
        Me.sear_bt.Enabled = True
        Me.insert_bt.Enabled = False
        'Me.del_bt.Enabled = False
        
        Me.ado_athlites.Refresh
        Me.dt_athlites.Columns(0).Visible = False
        Me.dt_athlites.Columns(1).Caption = "Αριθμός Μητρώου"
        Me.dt_athlites.Columns(1).Width = 2000
        Me.dt_athlites.Columns(2).Caption = "Επώνυμο"
        Me.dt_athlites.Columns(2).Width = 4000
        Me.dt_athlites.Columns(3).Caption = "Όνομα"
        Me.dt_athlites.Columns(3).Width = 4000
        For i = 4 To Me.ado_athlites.Recordset.Fields.Count - 1
            Me.dt_athlites.Columns(i).Visible = False
        Next i
        
End Sub

Private Sub Command2_Click()
    meli_management.Show
End Sub

Private Sub del_bt_Click()
    
    Dim ms As String

    ms = MsgBox("Είσαι σίγουρος; (ΝΑΙ ή ΟΧΙ)", vbYesNo, "Παράθυρο διαγραφής")
    If ms = 6 Then
        Me.ado_athlites.Recordset.Delete
    End If
    
End Sub

Private Sub dt_athlites_HeadClick(ByVal ColIndex As Integer)
    
    defined_col = ColIndex
    
End Sub

Private Sub dt_athlites_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If Me.ado_athlites.Recordset.AbsolutePosition >= 1 And Me.ado_athlites.Recordset.AbsolutePosition <= Me.ado_athlites.Recordset.RecordCount Then
    
    If Trim(Me.ado_athlites.Recordset.Fields(1).Value) <> "" Then
        tmp_am.Text = Me.ado_athlites.Recordset.Fields(1).Value
    Else
        tmp_am.Text = ""
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
    If Trim(Me.ado_athlites.Recordset.Fields(15).Value) <> "" Then
        Me.ado_pateres.Recordset.Filter = "[id] LIKE '" & Str(Me.ado_athlites.Recordset.Fields(15).Value) & "'"
    Else
        Me.ado_pateres.Recordset.Filter = "[id] LIKE '" & Str(-1) & "'"
    End If
    If Trim(Me.ado_athlites.Recordset.Fields(16).Value) <> "" Then
        Me.ado_miteres.Recordset.Filter = "[id] LIKE '" & Str(Me.ado_athlites.Recordset.Fields(16).Value) & "'"
    Else
        Me.ado_miteres.Recordset.Filter = "[id] LIKE '" & Str(-1) & "'"
    End If
    If Trim(Me.ado_athlites.Recordset.Fields(17).Value) <> "" Then
        Me.co_sxolia.Text = Me.ado_athlites.Recordset.Fields(17).Value
    Else
        Me.co_sxolia.Text = ""
    End If
    
    End If

End Sub

Private Sub Form_Load()
  
    Me.Height = 10380
    Me.Width = 12075
    flag_pateras = 0
    flag_mitera = 0
    
    Me.ado_athlites.Recordset.Sort = "[" & Trim(Me.ado_athlites.Recordset.Fields(0).Name) & "]"
    '******************************************************
    If Not Me.ado_athlites.Recordset.EOF Then
        Me.ado_athlites.Recordset.MoveFirst
        If Trim(Me.ado_athlites.Recordset.Fields(1).Value) <> "" Then
            tmp_am.Text = Str(Me.ado_athlites.Recordset.Fields(1).Value)
        Else
            tmp_am.Text = ""
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
        For i = 4 To Me.ado_pateres.Recordset.Fields.Count - 1
            Me.dt_father.Columns(i).Visible = False
        Next i
        If Trim(Me.ado_athlites.Recordset.Fields(15).Value) <> "" Then
            Me.ado_pateres.Recordset.Filter = "[id] LIKE '" & Str(Me.ado_athlites.Recordset.Fields(15).Value) & "'"
        Else
            Me.ado_pateres.Recordset.Filter = "[id] LIKE '" & Str(-1) & "'"
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
        For i = 4 To Me.ado_miteres.Recordset.Fields.Count - 1
            Me.dt_mother.Columns(i).Visible = False
        Next i
        If Trim(Me.ado_athlites.Recordset.Fields(16).Value) <> "" Then
            Me.ado_miteres.Recordset.Filter = "[id] LIKE '" & Str(Me.ado_athlites.Recordset.Fields(16).Value) & "'"
        Else
            Me.ado_miteres.Recordset.Filter = "[id] LIKE '" & Str(-1) & "'"
        End If
        '
        'ΠΑΡΟΥΣΙΑΣΗ ΣΤΟΙΧΕΩΝ ΣΧΟΛΕΙΟΥ
        If Trim(Me.ado_athlites.Recordset.Fields(17).Value) <> "" Then
            Me.co_sxolia.Text = Me.ado_athlites.Recordset.Fields(17).Value
        Else
            Me.co_sxolia.Text = ""
        End If
    End If
    
    '***************************************************************
       
    If Me.ado_athlites.Recordset.RecordCount > 0 Then
        Me.dt_athlites.Row = 0
        Me.dt_athlites.Col = 1
    End If
    If Me.ado_athlites.Recordset.RecordCount > 0 Then
        Me.ado_athlites.Caption = "Αθλητής " & Me.dt_athlites.Row + 1 & " από " & Me.ado_athlites.Recordset.RecordCount
    End If
    
    Me.dt_athlites.Columns(0).Visible = False
    Me.dt_athlites.Columns(1).Caption = "Αριθμός Μητρώου"
    Me.dt_athlites.Columns(1).Width = 2000
    Me.dt_athlites.Columns(2).Caption = "Επώνυμο"
    Me.dt_athlites.Columns(2).Width = 4000
    Me.dt_athlites.Columns(3).Caption = "Όνομα"
    Me.dt_athlites.Columns(3).Width = 4000
    For i = 4 To Me.ado_athlites.Recordset.Fields.Count - 1
        Me.dt_athlites.Columns(i).Visible = False
    Next i
    
    'Err.Clear
    
End Sub

Private Sub insert_bt_Click()

    'Να βρω το υποψήφιο id_αθλητή
    Me.ado_athlites.Recordset.Sort = "[" & Trim(Me.ado_athlites.Recordset.Fields(0).Name) & "]"
    If Not Me.ado_athlites.Recordset.EOF Then
        Me.ado_athlites.Recordset.MoveLast
        id_α = Me.ado_athlites.Recordset.Fields(0).Value
        id_α = id_α + 1
    End If
    'Καθαρισμός πεδίων
    tmp_am.Text = ""
    tmp_onoma.Text = ""
    tmp_eponimo.Text = ""
    tmp_odos.Text = ""
    tmp_arithmos.Text = ""
    tmp_perioxi.Text = ""
    Me.co_pe.Text = ""
    Me.co_dimoi.Text = ""
    tmp_tk.Text = ""
    tmp_til_oikias.Text = ""
    tmp_kinito.Text = ""
    tmp_fax.Text = ""
    tmp_email.Text = ""
    Me.MaskEdBox1.Text = "00/00/0000"
    Me.ado_pateres.Recordset.Filter = "[id] LIKE '" & Str(-1) & "'"
    Me.ado_miteres.Recordset.Filter = "[id] LIKE '" & Str(-1) & "'"
    Me.co_sxolia.Text = ""
    
    Me.tmp_am.SetFocus
    
    Me.insert_bt.Enabled = False
    Me.save_command.Enabled = True
    Me.canc_bt.Enabled = True
    Me.up_bt.Enabled = False
    
End Sub

Private Sub kl_bt_Click()

    Unload Me

End Sub



Private Sub MaskEdBox1_LostFocus()
    
    If IsDate(Me.MaskEdBox1.Text) = False Then
        Dim year, minas As String
        Me.MaskEdBox1.SelStart = 3
        Me.MaskEdBox1.SelLength = 2
        minas = Me.MaskEdBox1.SelText
        Me.MaskEdBox1.SelStart = 6
        Me.MaskEdBox1.SelLength = 4
        year = Me.MaskEdBox1.SelText
        Me.MaskEdBox1.SelStart = 0
        Me.MaskEdBox1.SelLength = 10
        If Val(minas) = 0 Then
            If Val(year) <> 0 Then
                Me.MaskEdBox1.SelText = "00/00/" & Trim(Trim(year))
            Else
                Me.MaskEdBox1.SelText = "00/00/00"
            End If
        Else
            Me.MaskEdBox1.SelText = "00/" & Trim(minas) & "/" & Trim(year)
        End If
    End If
    
End Sub

Private Sub save_command_Click()

    'On Error Resume Next
    
    'Αποθήκευση στους αθλητές
    Me.ado_athlites.Recordset.AddNew
    Me.ado_athlites.Recordset.Fields(0).Value = id_α
    'Αποθήκευση ΛΟΙΠΩΝ ΣΤΟΙΧΕΙΑ
    If Trim(Me.tmp_onoma.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(3).Value = Me.tmp_onoma.Text
    End If
    If Trim(Me.tmp_eponimo.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(2).Value = Me.tmp_eponimo.Text
    End If
    If Trim(Me.tmp_odos.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(4).Value = Me.tmp_odos.Text
    End If
    If Trim(Me.tmp_arithmos.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(5).Value = Me.tmp_arithmos.Text
    End If
    If Trim(Me.tmp_perioxi.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(6).Value = Me.tmp_perioxi.Text
    End If
    If Trim(Me.co_dimoi.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(7).Value = Me.co_dimoi.Text
        If Me.co_dimoi.MatchedWithList = False Then
        'ΕΙΣΑΓΩΓΗ ΝΕΟΥ ΔΗΜΟΥ
            Me.ado_dimoi.Recordset.Sort = "[" & Trim(Me.ado_dimoi.Recordset.Fields(0).Name) & "]"
            If Not Me.ado_dimoi.Recordset.EOF Then
                Me.ado_dimoi.Recordset.MoveLast
                id = Me.ado_dimoi.Recordset![id_δήμου]
            End If
            Me.ado_dimoi.Recordset.AddNew
            Me.ado_dimoi.Recordset.Fields(0) = id + 1
            Me.ado_dimoi.Recordset.Fields(1) = Trim(Me.co_dimoi.Text)
            Me.ado_dimoi.Recordset.UpdateBatch adAffectCurrent
        End If
    End If
    If Trim(Me.co_pe.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(8).Value = Me.co_pe.Text
        If Me.co_pe.MatchedWithList = False Then
        'ΕΙΣΑΓΩΓΗ ΝΕΑΣ ΠΕΡΙΦΕΡΕΙΑΚΗΣ ΕΝΟΤΗΤΑΣ
            Me.ado_pe.Recordset.Sort = "[" & Trim(Me.ado_pe.Recordset.Fields(0).Name) & "]"
            If Not Me.ado_pe.Recordset.EOF Then
                Me.ado_pe.Recordset.MoveLast
                id = Me.ado_pe.Recordset![id_πε]
            End If
            Me.ado_pe.Recordset.AddNew
            Me.ado_pe.Recordset.Fields(0) = id + 1
            Me.ado_pe.Recordset.Fields(1) = Trim(Me.co_pe.Text)
            Me.ado_pe.Recordset.UpdateBatch adAffectCurrent
        End If
    End If
    If Trim(Me.tmp_tk.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(9).Value = Me.tmp_tk.Text
    End If
    If Trim(Me.tmp_til_oikias.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(10).Value = Me.tmp_til_oikias.Text
    End If
    If Trim(Me.tmp_kinito.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(11).Value = Me.tmp_kinito.Text
    End If
    If Trim(Me.tmp_fax.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(12).Value = Me.tmp_fax.Text
    End If
    If Trim(Me.tmp_email.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(13).Value = Me.tmp_email.Text
    End If
    If Trim(Me.MaskEdBox1.Text) <> "00/00/0000" Then
        Me.ado_athlites.Recordset.Fields(14).Value = Me.MaskEdBox1.Text
    End If
    If Trim(Me.tmp_am.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(1).Value = Me.tmp_am.Text
    End If
    If Me.dt_father.Row >= 0 Then
        Me.ado_athlites.Recordset.Fields(15).Value = Me.ado_pateres.Recordset.Fields(0).Value
    End If
    If Me.dt_mother.Row >= 0 Then
        Me.ado_athlites.Recordset.Fields(16).Value = Me.ado_miteres.Recordset.Fields(0).Value
    End If
    If Me.co_sxolia.Text <> "" Then
        Me.ado_athlites.Recordset.Fields(17).Value = Me.co_sxolia.Text
            If Me.co_sxolia.MatchedWithList = False Then
            'ΕΙΣΑΓΩΓΗ ΝΕΟΥ ΣΧΟΛΕΙΟΥ
                If Not Me.ado_sxolia.Recordset.EOF Then
                    Me.ado_sxolia.Recordset.MoveLast
                    id = Me.ado_sxolia.Recordset![id_σχολείου]
                End If
            Me.ado_sxolia.Recordset.AddNew
            Me.ado_sxolia.Recordset.Fields(0) = id + 1
            Me.ado_sxolia.Recordset.Fields(1) = Trim(Me.co_sxolia.Text)
            Me.ado_sxolia.Recordset.UpdateBatch adAffectCurrent
        End If
    End If
    '
    Me.ado_athlites.Recordset.UpdateBatch adAffectCurrent
    Me.ado_athlites.Recordset.Sort = "[" & Trim(Me.ado_athlites.Recordset.Fields(0).Name) & "]"
    Me.ado_athlites.Recordset.MoveLast
    
    Me.save_command.Enabled = False
    Me.canc_bt.Enabled = True
    Me.insert_bt.Enabled = True
    Me.up_bt.Enabled = True
    
    'Err.Clear
    
End Sub

Private Sub sear_bt_Click()

    Dim s_string As String
    
    s_string = ""
    If Trim(Me.tmp_am.Text) <> "" Then
        s_string = "[ΑΜ_αθλητή] LIKE '*" & Trim(Me.tmp_am.Text) & "*'"
    End If
    If Trim(tmp_eponimo.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Επώνυμο] LIKE '*" & Trim(tmp_eponimo.Text) & "*'"
        Else
            s_string = "[Επώνυμο] LIKE '*" & Trim(tmp_eponimo.Text) & "*'"
        End If
    End If
    If Trim(tmp_onoma.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Όνομα] LIKE '*" & Trim(tmp_onoma.Text) & "*'"
        Else
            s_string = "[Όνομα] LIKE '*" & Trim(tmp_onoma.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΗΜΕΡΟΜΗΝΙΑ ΓΕΝΝΗΣΗΣ
    If Trim(MaskEdBox1.Text) <> "00/00/0000" Then
        Dim st As String
        Dim imera, minas, etos As String
        st = Trim(MaskEdBox1.Text)
        imera = Mid(Trim(MaskEdBox1.Text), 1, 2)
        minas = Mid(Trim(MaskEdBox1.Text), 4, 2)
        etos = Mid(Trim(MaskEdBox1.Text), 7, 4)
        If Val(imera) = 0 Then
            If Val(minas) = 0 Then
                If Val(etos) = 0 Then
                    st = ""
                Else
                    st = Val(etos)
                End If
            Else
                st = Val(minas) & "/" & Val(etos)
            End If
        End If
        If s_string <> "" Then
            s_string = s_string & " AND [Γέννηση] LIKE '*" & st & "*'"
        Else
            s_string = "[Γέννηση] LIKE '*" & st & "*'"
        End If
    End If
    '
    If Trim(tmp_til_oikias.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΤηλέφωνοΟικίας] LIKE '*" & Trim(tmp_til_oikias.Text) & "*'"
        Else
            s_string = "[ΤηλέφωνοΟικίας] LIKE '*" & Trim(tmp_til_oikias.Text) & "*'"
        End If
    End If
    If Trim(tmp_kinito.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΚινητόΤηλέφωνο] LIKE '*" & Trim(tmp_kinito.Text) & "*'"
        Else
            s_string = "[ΚινητόΤηλέφωνο] LIKE '*" & Trim(tmp_kinito.Text) & "*'"
        End If
    End If
    If Trim(tmp_fax.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΑριθμόςΦαξ] LIKE '*" & Trim(tmp_fax.Text) & "*'"
        Else
            s_string = "[ΑριθμόςΦαξ] LIKE '*" & Trim(tmp_fax.Text) & "*'"
        End If
    End If
    If Trim(tmp_email.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΌνομαEmail] LIKE '*" & Trim(tmp_email.Text) & "*'"
        Else
            s_string = "[ΌνομαEmail] LIKE '*" & Trim(tmp_email.Text) & "*'"
        End If
    End If
    If Trim(tmp_odos.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Οδός] LIKE '*" & Trim(tmp_odos.Text) & "*'"
        Else
            s_string = "[Οδός] LIKE '*" & Trim(tmp_odos.Text) & "*'"
        End If
    End If
    If Trim(tmp_arithmos.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Αριθμός] LIKE '*" & Trim(tmp_arithmos.Text) & "*'"
        Else
            s_string = "[Αριθμός] LIKE '*" & Trim(tmp_arithmos.Text) & "*'"
        End If
    End If
    If Trim(tmp_perioxi.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Περιοχή] LIKE '*" & Trim(tmp_perioxi.Text) & "*'"
        Else
            s_string = "[Περιοχή] LIKE '*" & Trim(tmp_perioxi.Text) & "*'"
        End If
    End If
    If Trim(co_dimoi.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Δήμος] LIKE '*" & Trim(co_dimoi.Text) & "*'"
        Else
            s_string = "[Δήμος] LIKE '*" & Trim(co_dimoi.Text) & "*'"
        End If
    End If
    If Trim(co_pe.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Περιφερειακή Ενότητα] LIKE '*" & Trim(co_pe.Text) & "*'"
        Else
            s_string = "[Περιφερειακή Ενότητα] LIKE '*" & Trim(co_pe.Text) & "*'"
        End If
    End If
    If Trim(tmp_tk.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Ταχυδρομικός Κώδικας] LIKE '*" & Trim(tmp_tk.Text) & "*'"
        Else
            s_string = "[Ταχυδρομικός Κώδικας] LIKE '*" & Trim(tmp_tk.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΠΑΤΕΡΑ
    If Me.dt_father.Row >= 0 Then
        If s_string <> "" Then
            s_string = s_string & " AND [Id_Πατέρα] = '" & Val(Me.dt_father.Columns(0).Value) & "'"
        Else
            s_string = "[Id_Πατέρα] = '" & Val(Me.dt_father.Columns(0).Value) & "'"
        End If
    End If
    '
    'ΚΡΙΤΗΡΙΟ ΜΗΤΕΡΑΣ
    If Me.dt_mother.Row >= 0 Then
        If s_string <> "" Then
            s_string = s_string & " AND [Id_Μητέρας] = '" & Val(Me.dt_mother.Columns(0).Value) & "'"
        Else
            s_string = "[Id_Μητέρας] = '" & Val(Me.dt_mother.Columns(0).Value) & "'"
        End If
    End If
    '
    ' ΚΡΙΤΗΡΙΟ ΣΧΟΛΕΙΟΥ
    If Trim(co_sxolia.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Σχολείο] LIKE '*" & Trim(co_sxolia.Text) & "*'"
        Else
            s_string = "[Σχολείο] LIKE '*" & Trim(co_sxolia.Text) & "*'"
        End If
    End If
    '
    Me.ado_athlites.Recordset.Filter = s_string
    
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

    If Me.dt_athlites.Col >= 0 Then
        Me.ado_athlites.Recordset.Sort = "[" & Trim(Me.ado_athlites.Recordset.Fields(dt_athlites.Col).Name) & "]"
    Else
        Me.ado_athlites.Recordset.Sort = "[" & Trim(Me.ado_athlites.Recordset.Fields(defined_col).Name) & "]"
    End If
  

End Sub

Private Sub up_bt_Click()

    Dim id As Integer

    If Trim(Me.tmp_onoma.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(3).Value = Me.tmp_onoma.Text
    End If
    If Trim(Me.tmp_eponimo.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(2).Value = Me.tmp_eponimo.Text
    End If
    If Trim(Me.tmp_odos.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(4).Value = Me.tmp_odos.Text
    End If
    If Trim(Me.tmp_arithmos.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(5).Value = Me.tmp_arithmos.Text
    End If
    If Trim(Me.tmp_perioxi.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(6).Value = Me.tmp_perioxi.Text
    End If
    If Me.co_dimoi.Text <> "" Then
        Me.ado_athlites.Recordset.Fields(7).Value = Me.co_dimoi.Text
            If Me.co_dimoi.MatchedWithList = False Then
            'ΕΙΣΑΓΩΓΗ ΝΕΟΥ ΔΗΜΟΥ
                If Not Me.ado_dimoi.Recordset.EOF Then
                    Me.ado_dimoi.Recordset.MoveLast
                    id = Me.ado_dimoi.Recordset![id_δήμου]
                End If
            Me.ado_dimoi.Recordset.AddNew
            Me.ado_dimoi.Recordset.Fields(0) = id + 1
            Me.ado_dimoi.Recordset.Fields(1) = Trim(Me.co_dimoi.Text)
            Me.ado_dimoi.Recordset.UpdateBatch adAffectCurrent
        End If
    End If
    If Trim(Me.co_pe.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(8).Value = Me.co_pe.Text
        If Me.co_pe.MatchedWithList = False Then
        'ΕΙΣΑΓΩΓΗ ΝΕΑΣ ΠΕΡΙΦΕΡΕΙΑΚΗΣ ΕΝΟΤΗΤΑΣ
            Me.ado_pe.Recordset.Sort = "[" & Trim(Me.ado_pe.Recordset.Fields(0).Name) & "]"
            If Not Me.ado_pe.Recordset.EOF Then
                Me.ado_pe.Recordset.MoveLast
                id = Me.ado_pe.Recordset![id_πε]
            End If
            Me.ado_pe.Recordset.AddNew
            Me.ado_pe.Recordset.Fields(0) = id + 1
            Me.ado_pe.Recordset.Fields(1) = Trim(Me.co_pe.Text)
            Me.ado_pe.Recordset.UpdateBatch adAffectCurrent
        End If
    End If
    If Trim(Me.tmp_tk.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(9).Value = Me.tmp_tk.Text
    End If
    If Trim(Me.tmp_til_oikias.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(10).Value = Me.tmp_til_oikias.Text
    End If
    If Trim(Me.tmp_kinito.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(11).Value = Me.tmp_kinito.Text
    End If
    If Trim(Me.tmp_fax.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(12).Value = Me.tmp_fax.Text
    End If
    If Trim(Me.tmp_email.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(13).Value = Me.tmp_email.Text
    End If
    If Trim(Me.MaskEdBox1.Text) <> "00/00/0000" Then
        Me.ado_athlites.Recordset.Fields(14).Value = Me.MaskEdBox1.Text
    End If
    If Trim(Me.tmp_am.Text) <> "" Then
        Me.ado_athlites.Recordset.Fields(1).Value = Me.tmp_am.Text
    End If
    If Me.dt_father.Row >= 0 Then
        Me.ado_athlites.Recordset.Fields(15).Value = Me.ado_pateres.Recordset.Fields(0).Value
    End If
    If Me.dt_mother.Row >= 0 Then
        Me.ado_athlites.Recordset.Fields(16).Value = Me.ado_miteres.Recordset.Fields(0).Value
    End If
    If Me.co_sxolia.Text <> "" Then
        Me.ado_athlites.Recordset.Fields(17).Value = Me.co_sxolia.Text
            If Me.co_sxolia.MatchedWithList = False Then
            'ΕΙΣΑΓΩΓΗ ΝΕΟΥ ΣΧΟΛΕΙΟΥ
                If Not Me.ado_sxolia.Recordset.EOF Then
                    Me.ado_sxolia.Recordset.MoveLast
                    id = Me.ado_sxolia.Recordset![id_σχολείου]
                End If
            Me.ado_sxolia.Recordset.AddNew
            Me.ado_sxolia.Recordset.Fields(0) = id + 1
            Me.ado_sxolia.Recordset.Fields(1) = Trim(Me.co_sxolia.Text)
            Me.ado_sxolia.Recordset.UpdateBatch adAffectCurrent
        End If
    End If
    '
    
    Me.ado_athlites.Recordset.UpdateBatch adAffectCurrent
    
End Sub
