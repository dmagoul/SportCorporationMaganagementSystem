VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form athl_tmima_management 
   BackColor       =   &H80000014&
   ClientHeight    =   9870
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12480
   LinkTopic       =   "entity_management"
   MDIChild        =   -1  'True
   ScaleHeight     =   9870
   ScaleWidth      =   12480
   Begin VB.CommandButton sear_bt 
      BackColor       =   &H80000014&
      Caption         =   "Αναζήτηση"
      DisabledPicture =   "athl_tmima_management.frx":0000
      Enabled         =   0   'False
      Height          =   495
      Left            =   7320
      MaskColor       =   &H80000014&
      Picture         =   "athl_tmima_management.frx":0973
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton del_bt 
      BackColor       =   &H80000014&
      Caption         =   "Διαγραφή"
      DisabledPicture =   "athl_tmima_management.frx":0A88
      Height          =   495
      Left            =   4440
      MaskColor       =   &H80000014&
      Picture         =   "athl_tmima_management.frx":13FB
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000014&
      Caption         =   "Καθαρισμός"
      DisabledPicture =   "athl_tmima_management.frx":14B8
      Height          =   495
      Left            =   5880
      MaskColor       =   &H80000014&
      Picture         =   "athl_tmima_management.frx":1600
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton up_bt 
      BackColor       =   &H80000014&
      Caption         =   "Ενημέρωση"
      DisabledPicture =   "athl_tmima_management.frx":16B9
      Height          =   495
      Left            =   1560
      MaskColor       =   &H80000014&
      Picture         =   "athl_tmima_management.frx":202C
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton kl_bt 
      BackColor       =   &H80000014&
      Caption         =   "Κλείσιμο"
      DisabledPicture =   "athl_tmima_management.frx":2493
      Height          =   495
      Left            =   10200
      MaskColor       =   &H80000014&
      Picture         =   "athl_tmima_management.frx":2E06
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton canc_bt 
      BackColor       =   &H80000014&
      Caption         =   "Ακύρωση"
      DisabledPicture =   "athl_tmima_management.frx":3779
      Height          =   495
      Left            =   8760
      MaskColor       =   &H80000014&
      Picture         =   "athl_tmima_management.frx":38C1
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton save_command 
      BackColor       =   &H80000014&
      Caption         =   "Αποθήκευση"
      DisabledPicture =   "athl_tmima_management.frx":3B64
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      MaskColor       =   &H80000014&
      Picture         =   "athl_tmima_management.frx":3CAC
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton insert_bt 
      BackColor       =   &H80000014&
      Caption         =   "Εισαγωγή Τμήματος"
      DisabledPicture =   "athl_tmima_management.frx":3D71
      Height          =   495
      Left            =   120
      MaskColor       =   &H80000014&
      Picture         =   "athl_tmima_management.frx":3E8F
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Όλα τα Τμήματα"
      Height          =   4335
      Left            =   120
      TabIndex        =   14
      Top             =   5520
      Width           =   11655
      Begin VB.CommandButton taksin 
         BackColor       =   &H80000014&
         Caption         =   "Ταξινόμηση"
         DisabledPicture =   "athl_tmima_management.frx":3FB5
         Height          =   495
         Left            =   120
         MaskColor       =   &H80000014&
         Picture         =   "athl_tmima_management.frx":4928
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3600
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dt_all_tmimata 
         Bindings        =   "athl_tmima_management.frx":49E5
         Height          =   3015
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
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
      Begin MSAdodcLib.Adodc ado_all_tmimata 
         Height          =   375
         Left            =   120
         Top             =   3240
         Width           =   11415
         _ExtentX        =   20135
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
         Connect         =   $"athl_tmima_management.frx":4A03
         OLEDBString     =   $"athl_tmima_management.frx":4AB0
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Πλήρης_Στοιχεία_Τμημάτων"
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
      Begin MSAdodcLib.Adodc ado_tm_ids 
         Height          =   330
         Left            =   1920
         Top             =   3600
         Visible         =   0   'False
         Width           =   6375
         _ExtentX        =   11245
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
         Connect         =   $"athl_tmima_management.frx":4B5D
         OLEDBString     =   $"athl_tmima_management.frx":4C0A
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "athl_tmima_management.frx":4CB7
         Height          =   255
         Left            =   9240
         TabIndex        =   40
         Top             =   3840
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
      BackColor       =   &H00C0FFFF&
      Caption         =   "Στοιχεία Τμήματος"
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
      TabIndex        =   16
      Top             =   0
      Width           =   12255
      Begin VB.Frame Frame4 
         Caption         =   "Τμήματα του Αθλητή: <<"
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
         Height          =   3015
         Left            =   4920
         TabIndex        =   38
         Top             =   1800
         Width           =   6735
         Begin MSDataGridLib.DataGrid dt_tmimata 
            Bindings        =   "athl_tmima_management.frx":4CD0
            Height          =   2175
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   6345
            _ExtentX        =   11192
            _ExtentY        =   3836
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
         Begin MSAdodcLib.Adodc ado_tmimata 
            Height          =   375
            Left            =   120
            Top             =   2520
            Width           =   6360
            _ExtentX        =   11218
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
            Connect         =   $"athl_tmima_management.frx":4CEA
            OLEDBString     =   $"athl_tmima_management.frx":4D97
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "ΑθλητέςΤμήματα"
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
      Begin VB.Frame Frame7 
         Caption         =   "Εβδομαδιαίο Πρόγραμμα Τμήματος"
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
         Height          =   1455
         Left            =   5040
         TabIndex        =   26
         Top             =   240
         Width           =   7095
         Begin MSMask.MaskEdBox im_l 
            Height          =   300
            Index           =   0
            Left            =   840
            TabIndex        =   13
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_en 
            Height          =   300
            Index           =   0
            Left            =   840
            TabIndex        =   7
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_en 
            Height          =   300
            Index           =   1
            Left            =   1800
            TabIndex        =   6
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_en 
            Height          =   300
            Index           =   2
            Left            =   2760
            TabIndex        =   5
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_en 
            Height          =   300
            Index           =   3
            Left            =   3720
            TabIndex        =   4
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_en 
            Height          =   300
            Index           =   4
            Left            =   4680
            TabIndex        =   3
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_en 
            Height          =   300
            Index           =   5
            Left            =   5640
            TabIndex        =   2
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_l 
            Height          =   300
            Index           =   1
            Left            =   1800
            TabIndex        =   12
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
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
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_l 
            Height          =   300
            Index           =   3
            Left            =   3720
            TabIndex        =   10
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_l 
            Height          =   300
            Index           =   4
            Left            =   4680
            TabIndex        =   9
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox im_l 
            Height          =   300
            Index           =   5
            Left            =   5640
            TabIndex        =   8
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Λήξη "
            Height          =   300
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Έναρξη "
            Height          =   300
            Left            =   120
            TabIndex        =   47
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Δευτέρα"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   41
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Σάββατο"
            Height          =   255
            Index           =   5
            Left            =   5640
            TabIndex        =   46
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Παρασκευή"
            Height          =   255
            Index           =   4
            Left            =   4680
            TabIndex        =   45
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Πέμπτη"
            Height          =   255
            Index           =   3
            Left            =   3720
            TabIndex        =   44
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Τετάρτη"
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   43
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Τρίτη"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   42
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Γενικά Στοιχεία"
         Height          =   4575
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   4695
         Begin VB.TextBox tmp_im_eis 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   2
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   49
            Top             =   2160
            Width           =   1215
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
            Height          =   375
            Left            =   1560
            TabIndex        =   34
            Top             =   3120
            Width           =   1215
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
            Height          =   375
            Left            =   1560
            TabIndex        =   32
            Top             =   2640
            Width           =   1215
         End
         Begin MSDataListLib.DataCombo co_ae 
            Bindings        =   "athl_tmima_management.frx":4E44
            Height          =   315
            Left            =   1560
            TabIndex        =   0
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "περιγραφή"
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc ado_ae 
            Height          =   375
            Left            =   3960
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
            Connect         =   $"athl_tmima_management.frx":4E59
            OLEDBString     =   $"athl_tmima_management.frx":4F06
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
            Bindings        =   "athl_tmima_management.frx":4FB3
            Height          =   315
            Left            =   1560
            TabIndex        =   1
            Top             =   720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "περιγραφή"
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc ado_kt 
            Height          =   375
            Left            =   3960
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
            Connect         =   $"athl_tmima_management.frx":4FC8
            OLEDBString     =   $"athl_tmima_management.frx":5075
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Κατηγορίες_Τμημάτων"
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
            Bindings        =   "athl_tmima_management.frx":5122
            Height          =   315
            Left            =   1560
            TabIndex        =   30
            Top             =   1200
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "OE"
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin MSAdodcLib.Adodc ado_prop 
            Height          =   375
            Left            =   3960
            Top             =   1200
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
            Connect         =   $"athl_tmima_management.frx":5139
            OLEDBString     =   $"athl_tmima_management.frx":51E6
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Ονοματεπώνυμα_Μελών"
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
            Bindings        =   "athl_tmima_management.frx":5293
            Height          =   315
            Left            =   1560
            TabIndex        =   36
            Top             =   1680
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            ListField       =   "OE"
            BoundColumn     =   ""
            Text            =   ""
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ποσό Εγγραφής"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   50
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Προπονητής Β"
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ποσό Μηνιαίας Συνδρομής"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   480
            TabIndex        =   35
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ημερομηνία Εισαγωγής"
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
            Left            =   0
            TabIndex        =   33
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Προπονητής Α"
            Height          =   375
            Left            =   240
            TabIndex        =   31
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Κατηγορία Τμήματος"
            Height          =   375
            Left            =   480
            TabIndex        =   29
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Αθλητικό Έτος"
            Height          =   375
            Left            =   480
            TabIndex        =   27
            Top             =   240
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "athl_tmima_management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id_τμ, flag_pateras, flag_mitera, defined_col As Integer

Private Sub ado_full_tmimata_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    
End Sub

Private Sub ado_tmimata_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

'    If Not adReason = 10 And pRecordset.AbsolutePosition >= 1 Then
'        If Not pRecordset.EOF Then
'            If Trim(pRecordset.Fields(0).Value) <> "" Then
'                If Not Me.ado_ae.Recordset.EOF Then
'                    'Me.ado_ae.Recordset.MoveFirst
'                    Me.ado_ae.Recordset.Find "[id_αθλητικού_έτους] = '" & Me.ado_tmimata.Recordset.Fields(0).Value & "'"
'                    If Not Me.ado_ae.Recordset.EOF Then
'                        co_ae.Text = Me.ado_ae.Recordset.Fields(1).Value
'                    End If
'                End If
'            Else
'                co_ae.Text = ""
'            End If
'            If Trim(Me.ado_tmimata.Recordset.Fields(2).Value) <> "" Then
'                If Not Me.ado_kt.Recordset.EOF Then
'                    'Me.ado_kt.Recordset.MoveFirst
'                   Me.ado_kt.Recordset.Find "[id_κατηγορίας_τμήματος] = '" & Me.ado_tmimata.Recordset.Fields(2).Value & "'"
'                    If Not Me.ado_kt.Recordset.EOF Then
'                        co_kt.Text = Me.ado_kt.Recordset.Fields(1).Value
'                    End If
'                End If
'            Else
'                co_kt.Text = ""
'            End If
'            If Trim(Me.ado_tmimata.Recordset.Fields(4).Value) <> "" Then
'                If Not Me.ado_prop.Recordset.EOF Then
'                    'Me.ado_prop.Recordset.MoveFirst
'                    Me.ado_prop.Recordset.Find "[id] = '" & Me.ado_tmimata.Recordset.Fields(4).Value & "'"
'                    If Not Me.ado_prop.Recordset.EOF Then
'                        co_propa.Text = Me.ado_prop.Recordset.Fields(1).Value
'                    End If
'                End If
'            Else
'                tmp_propA.Text = ""
'            End If
'            If Trim(Me.ado_tmimata.Recordset.Fields(6).Value) <> "" Then
'                If Not Me.ado_prop.Recordset.EOF Then
'                    'Me.ado_prop.Recordset.MoveFirst
'                    Me.ado_prop.Recordset.Find "[id] = '" & Me.ado_tmimata.Recordset.Fields(6).Value & "'"
'                    If Not Me.ado_prop.Recordset.EOF Then
'                        co_propb.Text = Me.ado_prop.Recordset.Fields(1).Value
'                    End If
'                End If
'            Else
'                co_propb.Text = ""
'            End If
'            If Trim(Me.ado_tmimata.Recordset.Fields(8).Value) <> "" Then
'                tmp_poso_eggrafis.Text = Str(Me.ado_tmimata.Recordset.Fields(8).Value)
'            Else
'                tmp_poso_eggrafis.Text = ""
'            End If
'            If Trim(Me.ado_tmimata.Recordset.Fields(9).Value) <> "" Then
'                Me.tmp_poso_mina.Text = Str(Me.ado_tmimata.Recordset.Fields(9).Value)
'            Else
'                Me.tmp_poso_mina.Text = ""
'            End If
'            'ΕΝΗΜΕΡΩΣΗ ΤΟΥ ΕΒΔΟΜΑΔΙΑΙΟΥ ΠΡΟΓΡΑΜΜΑΤΟΣ
'            'i = 0
'            'For j = 0 To 5
'            '    If Trim(Me.ado_tmimata.Recordset.Fields(11 + i).Value) <> "" Then
'            '        Me.im_en(j).Text = Me.ado_tmimata.Recordset.Fields(11 + i).Value
'            '    Else
'            '        Me.im_en(j).Text = "00:00"
'            '    End If
'            '    If Trim(Me.ado_tmimata.Recordset.Fields(12 + i).Value) <> "" Then
'            '        Me.im_l(j).Text = Me.ado_tmimata.Recordset.Fields(12 + i).Value
'            '    Else
'            '        Me.im_l(j).Text = "00:00"
'            '    End If
'            '    i = i + 2
'            'Next j
'        End If
'    End If
'    '*******************
'    Me.ado_tmimata.Caption = "Τμήμα " & pRecordset.AbsolutePosition & " από " & pRecordset.RecordCount'

End Sub

Private Sub Adodc3_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    If Not adReason = 10 And pRecordset.AbsolutePosition >= 1 Then
        If Not pRecordset.EOF Then
            If Trim(pRecordset.Fields(0).Value) <> "" Then
                If Not Me.ado_ae.Recordset.EOF Then
                    Me.ado_ae.Recordset.MoveFirst
                    Me.ado_ae.Recordset.Find "[id_αθλητικού_έτους] = '" & Me.ado_tmimata.Recordset.Fields(0).Value & "'"
                    If Not Me.ado_ae.Recordset.EOF Then
                        co_ae.Text = Me.ado_ae.Recordset.Fields(0).Value
                    End If
                End If
            Else
                co_ae.Text = ""
            End If
'            If Trim(Me.ado_tmimata.Recordset.Fields(2).Value) <> "" Then
'                If Not Me.ado_kt.Recordset.EOF Then
'                    Me.ado_kt.Recordset.MoveFirst
'                   Me.ado_kt.Recordset.Find "[id_κατηγορίας_τμήματος] = '" & Me.ado_tmimata.Recordset.Fields(2).Value & "'"
'                    If Not Me.ado_kt.Recordset.EOF Then
'                        co_kt.Text = Me.ado_kt.Recordset.Fields(1).Value
'                    End If
'                End If
'            Else
'                co_kt.Text = ""
'            End If
'            If Trim(Me.ado_tmimata.Recordset.Fields(3).Value) <> "" Then
'                If Not Me.ado_prop.Recordset.EOF Then
'                    Me.ado_prop.Recordset.MoveFirst
'                    Me.ado_prop.Recordset.Find "[id] = '" & Me.ado_tmimata.Recordset.Fields(3).Value & "'"
'                    If Not Me.ado_prop.Recordset.EOF Then
'                        co_propa.Text = Me.ado_prop.Recordset.Fields(1).Value
'                    End If
'                End If
'            Else
'                tmp_propA.Text = ""
'            End If
'            If Trim(Me.ado_tmimata.Recordset.Fields(4).Value) <> "" Then
'                If Not Me.ado_prop.Recordset.EOF Then
'                    Me.ado_prop.Recordset.MoveFirst
'                    Me.ado_prop.Recordset.Find "[id] = '" & Me.ado_tmimata.Recordset.Fields(4).Value & "'"
'                    If Not Me.ado_prop.Recordset.EOF Then
'                        co_propb.Text = Me.ado_prop.Recordset.Fields(1).Value
'                    End If
'                End If
'            Else
'                co_propb.Text = ""
'            End If
'            If Trim(Me.ado_tmimata.Recordset.Fields(5).Value) <> "" Then
'                tmp_poso_eggrafis.Text = Str(Me.ado_tmimata.Recordset.Fields(5).Value)
'            Else
'                tmp_poso_eggrafis.Text = ""
'            End If
'            If Trim(Me.ado_tmimata.Recordset.Fields(6).Value) <> "" Then
'                Me.tmp_poso_mina.Text = Str(Me.ado_tmimata.Recordset.Fields(6).Value)
'            Else
'                Me.tmp_poso_mina.Text = ""
'            End If
        End If
    End If
    '*******************
    Me.Adodc3.Caption = "Τμήμα " & pRecordset.AbsolutePosition & " από " & pRecordset.RecordCount


End Sub

Private Sub canc_bt_Click()

    Me.insert_bt.Enabled = True
    Me.up_bt.Enabled = True
    Me.sear_bt.Enabled = False
    
    Me.ado_tmimata.Refresh
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
        
        Me.ado_tmimata.Refresh
        Me.dt_athlites.Columns(0).Visible = False
        Me.dt_athlites.Columns(1).Caption = "Αριθμός Μητρώου"
        Me.dt_athlites.Columns(1).Width = 2000
        Me.dt_athlites.Columns(2).Caption = "Επώνυμο"
        Me.dt_athlites.Columns(2).Width = 4000
        Me.dt_athlites.Columns(3).Caption = "Όνομα"
        Me.dt_athlites.Columns(3).Width = 4000
        For i = 4 To Me.ado_tmimata.Recordset.Fields.Count - 1
            Me.dt_athlites.Columns(i).Visible = False
        Next i
        
End Sub

Private Sub Command2_Click()
    meli_management.Show
End Sub

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    Me.ado_tmimata.Recordset.AbsolutePosition = Me.Adodc3.Recordset.AbsolutePosition

End Sub

Private Sub del_bt_Click()
    
    Dim ms As String

    If Not Me.ado_tmimata.Recordset.EOF Then
        ms = MsgBox("Είσαι σίγουρος; (ΝΑΙ ή ΟΧΙ)", vbYesNo, "Παράθυρο διαγραφής")
        If ms = 6 Then
            Me.ado_tm_ids.Recordset.Sort = "[" & Trim(Me.ado_tm_ids.Recordset.Fields(0).Name) & "]"
            col_affected = Me.ado_tmimata.Recordset.AbsolutePosition
            Me.ado_tm_ids.Recordset.MoveFirst
            If col_affected - 1 <> 0 Then
                Me.ado_tm_ids.Recordset.Move col_affected - 1
            End If
            If Me.ado_tmimata.Recordset.AbsolutePosition = Me.ado_tmimata.Recordset.RecordCount Then
                col_affected = col_affected - 1
            End If
            Me.ado_tm_ids.Recordset.Delete adAffectCurrent
            Me.ado_tm_ids.Recordset.Requery
            Me.ado_tm_ids.Refresh
            Me.DataGrid2.Refresh
            Me.ado_tmimata.Recordset.Requery
            Me.ado_tmimata.Refresh
            Me.dt_tmimata.Refresh
            If Not Me.ado_tmimata.Recordset.EOF Then
                Me.ado_tmimata.Recordset.MoveFirst
                Me.ado_tmimata.Recordset.Move col_affected - 1
                Me.dt_tmimata.Row = col_affected - 1
                Me.dt_tmimata.Col = 1
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
        End If
    Else
        MsgBox "Δεν υπάρχει εγγραφή προς ΔΙΑΓΡΑΦΗ!", vbCritical, "Μήνυμα Λάθους"
    End If
    
    
End Sub

Private Sub dt_athlites_HeadClick(ByVal ColIndex As Integer)
    
    defined_col = ColIndex
    
End Sub

Private Sub dt_athlites_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If Me.ado_tmimata.Recordset.AbsolutePosition >= 1 And Me.ado_tmimata.Recordset.AbsolutePosition <= Me.ado_tmimata.Recordset.RecordCount Then
    
    If Trim(Me.ado_tmimata.Recordset.Fields(1).Value) <> "" Then
        tmp_am.Text = Me.ado_tmimata.Recordset.Fields(1).Value
    Else
        tmp_am.Text = ""
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(3).Value) <> "" Then
        tmp_onoma.Text = Me.ado_tmimata.Recordset.Fields(3).Value
    Else
        tmp_onoma.Text = ""
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(2).Value) <> "" Then
        tmp_eponimo.Text = Me.ado_tmimata.Recordset.Fields(2).Value
    Else
        tmp_eponimo.Text = ""
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(4).Value) <> "" Then
        tmp_odos.Text = Me.ado_tmimata.Recordset.Fields(4).Value
    Else
        tmp_odos.Text = ""
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(5).Value) <> "" Then
        tmp_arithmos.Text = Me.ado_tmimata.Recordset.Fields(5).Value
    Else
        tmp_arithmos.Text = ""
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(6).Value) <> "" Then
        tmp_perioxi.Text = Me.ado_tmimata.Recordset.Fields(6).Value
    Else
        tmp_perioxi.Text = ""
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(8).Value) <> "" Then
        Me.co_pe.Text = Me.ado_tmimata.Recordset.Fields(8).Value
    Else
        Me.co_pe.Text = ""
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(7).Value) <> "" Then
        Me.co_dimoi.Text = Me.ado_tmimata.Recordset.Fields(7).Value
    Else
        Me.co_dimoi.Text = ""
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(9).Value) <> "" Then
        tmp_tk.Text = Me.ado_tmimata.Recordset.Fields(9).Value
    Else
        tmp_tk.Text = ""
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(10).Value) <> "" Then
        tmp_til_oikias.Text = Me.ado_tmimata.Recordset.Fields(10).Value
    Else
        tmp_til_oikias.Text = ""
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(11).Value) <> "" Then
        tmp_kinito.Text = Me.ado_tmimata.Recordset.Fields(11).Value
    Else
        tmp_kinito.Text = ""
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(12).Value) <> "" Then
        tmp_fax.Text = Me.ado_tmimata.Recordset.Fields(12).Value
    Else
        tmp_fax.Text = ""
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(13).Value) <> "" Then
        tmp_email.Text = Me.ado_tmimata.Recordset.Fields(13).Value
    Else
        tmp_email.Text = ""
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(14).Value) <> "" Then
        Me.MaskEdBox1.Text = Me.ado_tmimata.Recordset.Fields(14).Value
    Else
        Me.MaskEdBox1.Text = "00/00/0000"
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(15).Value) <> "" Then
        Me.ado_pateres.Recordset.Filter = "[id] LIKE '" & Str(Me.ado_tmimata.Recordset.Fields(15).Value) & "'"
    Else
        Me.ado_pateres.Recordset.Filter = "[id] LIKE '" & Str(-1) & "'"
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(16).Value) <> "" Then
        Me.ado_miteres.Recordset.Filter = "[id] LIKE '" & Str(Me.ado_tmimata.Recordset.Fields(16).Value) & "'"
    Else
        Me.ado_miteres.Recordset.Filter = "[id] LIKE '" & Str(-1) & "'"
    End If
    If Trim(Me.ado_tmimata.Recordset.Fields(17).Value) <> "" Then
        Me.co_sxolia.Text = Me.ado_tmimata.Recordset.Fields(17).Value
    Else
        Me.co_sxolia.Text = ""
    End If
    
    End If

End Sub

Private Sub dt_full_tmimata_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Me.ado_full_tmimata.Recordset.AbsolutePosition >= 1 Then
        Me.ado_tmimata.Recordset.AbsolutePosition = Me.ado_full_tmimata.Recordset.AbsolutePosition
    End If

End Sub

Private Sub dt_tmimata_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Me.ado_tmimata.Recordset.AbsolutePosition >= 1 And Me.ado_tmimata.Recordset.AbsolutePosition <= Me.ado_tmimata.Recordset.RecordCount Then
            If Trim(Me.ado_tmimata.Recordset.Fields(10).Value) <> "" Then
                If Not Me.ado_ae.Recordset.EOF Then
                    Me.ado_ae.Recordset.MoveFirst
                    Me.ado_ae.Recordset.Find "[id_αθλητικού_έτους] = '" & Me.ado_tmimata.Recordset.Fields(10).Value & "'"
                    If Not Me.ado_ae.Recordset.EOF Then
                        co_ae.Text = Me.ado_ae.Recordset.Fields(1).Value
                    End If
                End If
            Else
                co_ae.Text = ""
            End If
            If Trim(Me.ado_tmimata.Recordset.Fields(11).Value) <> "" Then
                If Not Me.ado_kt.Recordset.EOF Then
                    Me.ado_kt.Recordset.MoveFirst
                   Me.ado_kt.Recordset.Find "[id_κατηγορίας_τμήματος] = '" & Me.ado_tmimata.Recordset.Fields(11).Value & "'"
                    If Not Me.ado_kt.Recordset.EOF Then
                        co_kt.Text = Me.ado_kt.Recordset.Fields(1).Value
                    End If
                End If
            Else
                co_kt.Text = ""
            End If
            If Trim(Me.ado_tmimata.Recordset.Fields(12).Value) <> "" Then
                If Not Me.ado_prop.Recordset.EOF Then
                    Me.ado_prop.Recordset.MoveFirst
                    Me.ado_prop.Recordset.Find "[id] = '" & Me.ado_tmimata.Recordset.Fields(12).Value & "'"
                    If Not Me.ado_prop.Recordset.EOF Then
                        co_propa.Text = Me.ado_prop.Recordset.Fields(1).Value
                    End If
                End If
            Else
                tmp_propA.Text = ""
            End If
            If Trim(Me.ado_tmimata.Recordset.Fields(13).Value) <> "" Then
                If Not Me.ado_prop.Recordset.EOF Then
                    Me.ado_prop.Recordset.MoveFirst
                    Me.ado_prop.Recordset.Find "[id] = '" & Me.ado_tmimata.Recordset.Fields(13).Value & "'"
                    If Not Me.ado_prop.Recordset.EOF Then
                        co_propb.Text = Me.ado_prop.Recordset.Fields(1).Value
                    End If
                End If
            Else
                co_propb.Text = ""
            End If
            If Trim(Me.ado_tmimata.Recordset.Fields(7).Value) <> "00/00/0000" Then
                tmp_im_eis.Text = Str(Me.ado_tmimata.Recordset.Fields(7).Value)
            Else
                tmp_im_eis.Text = "00/00/0000"
            End If
            If Trim(Me.ado_tmimata.Recordset.Fields(8).Value) <> "" Then
                tmp_poso_eggrafis.Text = Str(Me.ado_tmimata.Recordset.Fields(8).Value)
            Else
                tmp_poso_eggrafis.Text = ""
            End If
            If Trim(Me.ado_tmimata.Recordset.Fields(9).Value) <> "" Then
                Me.tmp_poso_mina.Text = Str(Me.ado_tmimata.Recordset.Fields(9).Value)
            Else
                Me.tmp_poso_mina.Text = ""
            End If
            'ΕΝΗΜΕΡΩΣΗ ΤΟΥ ΕΒΔΟΜΑΔΙΑΙΟΥ ΠΡΟΓΡΑΜΜΑΤΟΣ
            'i = 0
            'For j = 0 To 5
            '    If Trim(Me.ado_tmimata.Recordset.Fields(11 + i).Value) <> "" Then
            '        Me.im_en(j).Text = Me.ado_tmimata.Recordset.Fields(11 + i).Value
            '    Else
            '        Me.im_en(j).Text = "00:00"
            '    End If
            '    If Trim(Me.ado_tmimata.Recordset.Fields(12 + i).Value) <> "" Then
            '        Me.im_l(j).Text = Me.ado_tmimata.Recordset.Fields(12 + i).Value
            '    Else
            '        Me.im_l(j).Text = "00:00"
            '    End If
            '    i = i + 2
            'Next j
        End If
    '*******************
    Me.ado_tmimata.Caption = "Τμήμα " & Me.ado_tmimata.Recordset.AbsolutePosition & " από " & Me.ado_tmimata.Recordset.RecordCount

End Sub

Private Sub Form_Load()
  
    Me.Height = 10380
    Me.Width = 12075
    
    
    Me.Frame4.Caption = Me.Frame4.Caption & athlet_management.tmp_onoma & " " & athlet_management.tmp_eponimo & ">>"
    
    Me.ado_tmimata.Recordset.Sort = "[" & Trim(Me.ado_tmimata.Recordset.Fields(6).Name) & "]"
    Me.ado_tmimata.Recordset.Filter = "[id_αθλητή] = '" & athlet_management.ado_athlites.Recordset.Fields(0).Value & "'"
    'Me.ado_tm_ids.Recordset.Sort = "[" & Trim(Me.ado_tm_ids.Recordset.Fields(0).Name) & "]"
    '******************************************************
    If Not Me.ado_tmimata.Recordset.EOF Then
        Me.ado_tmimata.Recordset.MoveLast
        If Trim(Me.ado_tmimata.Recordset.Fields(0).Value) <> "" Then
                    co_ae.Text = Trim(Me.ado_tmimata.Recordset.Fields(0).Value)
        Else
            co_ae.Text = ""
        End If
        If Trim(Me.ado_tmimata.Recordset.Fields(1).Value) <> "" Then
            co_kt.Text = Trim(Me.ado_tmimata.Recordset.Fields(1).Value)
        Else
            co_kt.Text = ""
        End If
        If Trim(Me.ado_tmimata.Recordset.Fields(2).Value) <> "" Then
            co_propa.Text = Trim(Me.ado_tmimata.Recordset.Fields(2).Value)
        Else
            tmp_propA.Text = ""
        End If
        If Trim(Me.ado_tmimata.Recordset.Fields(3).Value) <> "" Then
            co_propb.Text = Trim(Me.ado_tmimata.Recordset.Fields(3).Value)
        Else
            co_propb.Text = ""
        End If
         If Trim(Me.ado_tmimata.Recordset.Fields(7).Value) <> "" Then
            tmp_im_eis.Text = Trim(Me.ado_tmimata.Recordset.Fields(7).Value)
        Else
            tmp_im_eis.Text = ""
        End If
        If Trim(Me.ado_tmimata.Recordset.Fields(8).Value) <> "" Then
            tmp_poso_eggrafis.Text = Me.ado_tmimata.Recordset.Fields(8).Value
        Else
            tmp_poso_eggrafis.Text = ""
        End If
        If Trim(Me.ado_tmimata.Recordset.Fields(9).Value) <> "" Then
            Me.tmp_poso_mina.Text = Me.ado_tmimata.Recordset.Fields(9).Value
        Else
            Me.tmp_poso_mina.Text = ""
        End If
        'ΕΝΗΜΕΡΩΣΗ ΤΟΥ ΕΒΔΟΜΑΔΙΑΙΟΥ ΠΡΟΓΡΑΜΜΑΤΟΣ
        'i = 0
        'For j = 0 To 5
        '    If Trim(Me.ado_tmimata.Recordset.Fields(11 + i).Value) <> "" Then
        '        Me.im_en(j).Text = Me.ado_tmimata.Recordset.Fields(11 + i).Value
        '    Else
        '        Me.im_en(j).Text = "00:00"
        '    End If
        '    If Trim(Me.ado_tmimata.Recordset.Fields(12 + i).Value) <> "" Then
        '        Me.im_l(j).Text = Me.ado_tmimata.Recordset.Fields(12 + i).Value
        '    Else
        '        Me.im_l(j).Text = "00:00"
        '    End If
        '    i = i + 2
        'Next j
    End If
    '***************************************************************
       
    If Me.ado_tmimata.Recordset.RecordCount > 0 Then
        'Me.dt_tmimata.Row = 0
        Me.dt_tmimata.Col = 1
    End If
    If Me.ado_tmimata.Recordset.RecordCount > 0 Then
        Me.ado_tmimata.Caption = "Τμήμα " & Me.dt_tmimata.Row + 1 & " από " & Me.ado_tmimata.Recordset.RecordCount
    Else
        If Me.ado_tmimata.Recordset.RecordCount = 0 Then
            Me.ado_tmimata.Caption = "Τμήμα " & 0 & " από " & 0
        End If
    End If
    
    Me.dt_tmimata.Columns(0).Caption = "Αθλητικό Έτος"
    Me.dt_tmimata.Columns(0).Width = 1000
    Me.dt_tmimata.Columns(1).Caption = "Κατηγορία Τμήματος"
    Me.dt_tmimata.Columns(1).Width = 1500
    Me.dt_tmimata.Columns(2).Caption = "Α Προπονητής"
    Me.dt_tmimata.Columns(2).Width = 1500
    Me.dt_tmimata.Columns(3).Caption = "Β Προπονητής"
    Me.dt_tmimata.Columns(3).Width = 1500
    For i = 4 To Me.ado_tmimata.Recordset.Fields.Count - 1
        Me.dt_tmimata.Columns(i).Visible = False
    Next i
    'Me.dt_tmimata.Columns(Me.ado_tmimata.Recordset.Fields.Count - 1).Caption = "Β Προπονητής"
    'Me.dt_tmimata.Columns(Me.ado_tmimata.Recordset.Fields.Count - 1).Width = 2500
    
    'Err.Clear
    
    
    If Me.ado_all_tmimata.Recordset.RecordCount > 0 Then
        Me.dt_all_tmimata.Row = 0
        Me.dt_all_tmimata.Col = 1
    End If
    If Me.ado_all_tmimata.Recordset.RecordCount > 0 Then
        Me.ado_all_tmimata.Caption = "Τμήμα " & Me.dt_all_tmimata.Row + 1 & " από " & Me.ado_all_tmimata.Recordset.RecordCount
    Else
        If Me.ado_all_tmimata.Recordset.RecordCount = 0 Then
            Me.ado_all_tmimata.Caption = "Τμήμα " & 0 & " από " & 0
        End If
    End If
    
    Me.dt_all_tmimata.Columns(0).Visible = False
    Me.dt_all_tmimata.Columns(1).Caption = "Αθλητικό Έτος"
    Me.dt_all_tmimata.Columns(1).Width = 1500
    Me.dt_all_tmimata.Columns(2).Visible = False
    Me.dt_all_tmimata.Columns(3).Caption = "Κατηγορία Τμήματος"
    Me.dt_all_tmimata.Columns(3).Width = 2000
    Me.dt_all_tmimata.Columns(4).Visible = False
    Me.dt_all_tmimata.Columns(5).Caption = "Α Προπονητής"
    Me.dt_all_tmimata.Columns(5).Width = 2500
    Me.dt_all_tmimata.Columns(6).Visible = False
    Me.dt_all_tmimata.Columns(7).Caption = "Β Προπονητής"
    Me.dt_all_tmimata.Columns(7).Width = 2500
    For i = 8 To Me.ado_all_tmimata.Recordset.Fields.Count - 1
        Me.dt_all_tmimata.Columns(i).Visible = False
    Next i
    
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

    'Να βρω το υποψήφιο id_τμήματος
    Me.ado_tm_ids.Recordset.Sort = "[" & Trim(Me.ado_tm_ids.Recordset.Fields(0).Name) & "]"
    If Not Me.ado_tm_ids.Recordset.EOF Then
        Me.ado_tm_ids.Recordset.MoveLast
        id_τμ = Me.ado_tm_ids.Recordset.Fields(0).Value
        id_τμ = id_τμ + 1
    End If
    'Μεταφορά πεδίων
   If Not Me.ado_all_tmimata.Recordset.EOF Then
        'Me.ado_all_tmimata.Recordset.MoveFirst
        If Trim(Me.ado_all_tmimata.Recordset.Fields(0).Value) <> "" Then
            'If Not Me.ado_ae.Recordset.EOF Then
                Me.ado_ae.Recordset.MoveFirst
                Me.ado_ae.Recordset.Find "[id_αθλητικού_έτους] = '" & Me.ado_all_tmimata.Recordset.Fields(0).Value & "'"
                If Not Me.ado_ae.Recordset.EOF Then
                    co_ae.Text = Me.ado_ae.Recordset.Fields(1).Value
                End If
            'End If
        Else
            co_ae.Text = ""
        End If
        If Trim(Me.ado_all_tmimata.Recordset.Fields(2).Value) <> "" Then
            'If Not Me.ado_kt.Recordset.EOF Then
                Me.ado_kt.Recordset.MoveFirst
                Me.ado_kt.Recordset.Find "[id_κατηγορίας_τμήματος] = '" & Me.ado_all_tmimata.Recordset.Fields(2).Value & "'"
                If Not Me.ado_kt.Recordset.EOF Then
                    co_kt.Text = Me.ado_kt.Recordset.Fields(1).Value
                End If
            'End If
        Else
            co_kt.Text = ""
        End If
        If Trim(Me.ado_all_tmimata.Recordset.Fields(4).Value) <> "" Then
            'If Not Me.ado_prop.Recordset.EOF Then
                Me.ado_prop.Recordset.MoveFirst
                Me.ado_prop.Recordset.Find "[id] = '" & Me.ado_all_tmimata.Recordset.Fields(4).Value & "'"
                If Not Me.ado_prop.Recordset.EOF Then
                    co_propa.Text = Me.ado_prop.Recordset.Fields(1).Value
                End If
            'End If
        Else
            tmp_propA.Text = ""
        End If
        If Trim(Me.ado_all_tmimata.Recordset.Fields(6).Value) <> "" Then
            'If Not Me.ado_prop.Recordset.EOF Then
                Me.ado_prop.Recordset.MoveFirst
                Me.ado_prop.Recordset.Find "[id] = '" & Me.ado_all_tmimata.Recordset.Fields(6).Value & "'"
                If Not Me.ado_prop.Recordset.EOF Then
                    co_propb.Text = Me.ado_prop.Recordset.Fields(1).Value
                End If
            'End If
        Else
            co_propb.Text = ""
        End If
        
            tmp_im_eis.Text = "00/00/0000"
        
        If Trim(Me.ado_all_tmimata.Recordset.Fields(8).Value) <> "" Then
            tmp_poso_eggrafis.Text = Me.ado_all_tmimata.Recordset.Fields(8).Value
        Else
            tmp_poso_eggrafis.Text = ""
        End If
        If Trim(Me.ado_all_tmimata.Recordset.Fields(9).Value) <> "" Then
            Me.tmp_poso_mina.Text = Me.ado_all_tmimata.Recordset.Fields(9).Value
        Else
            Me.tmp_poso_mina.Text = ""
        End If
        'ΚΑΘΑΡΙΣΜΟΣ ΤΟΥ ΕΒΔΟΜΑΔΙΑΙΟΥ ΠΡΟΓΡΑΜΜΑΤΟΣ
    'For j = 0 To 5
    '    Me.im_en(j).Text = Me.im_en(j).Mask
    '    Me.im_l(j).Text = Me.im_l(j).Mask
    'Next j
    End If
    
    Me.co_ae.SetFocus
    
    'Me.insert_bt.Enabled = False
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

    'Αποθήκευση στα τμήματα
    Me.ado_tm_ids.Recordset.AddNew
    'Αποθήκευση id
    Me.ado_tm_ids.Recordset.Fields(0).Value = id_τμ
    'Αποθήκευση ΛΟΙΠΩΝ ΣΤΟΙΧΕΙΑ
    Me.ado_tm_ids.Recordset.Fields(1).Value = Me.ado_all_tmimata.Recordset.Fields(10).Value
    Me.ado_tm_ids.Recordset.Fields(2).Value = athlet_management.ado_athlites.Recordset.Fields(0).Value
    If Trim(Me.tmp_im_eis.Text) <> "" Then
        Me.ado_tm_ids.Recordset.Fields(3).Value = Me.tmp_im_eis.Text
    End If
    If Trim(Me.tmp_poso_eggrafis.Text) <> "" Then
        Me.ado_tm_ids.Recordset.Fields(4).Value = Me.tmp_poso_eggrafis.Text
    End If
    If Trim(Me.tmp_poso_mina.Text) <> "" Then
        Me.ado_tm_ids.Recordset.Fields(5).Value = Me.tmp_poso_mina.Text
    End If
    'ΕΝΗΜΕΡΩΣΗ ΤΟΥ ΕΒΔΟΜΑΔΙΑΙΟΥ ΠΡΟΓΡΑΜΜΑΤΟΣ
    'For j = 0 To 5
    '    If Trim(Me.im_en(j).Text) <> "" Then
    '        Me.ado_tm_ids.Recordset.Fields(7 + j).Value = Me.im_en(j).Text
    '    End If
    '    If Trim(Me.im_l(j).Text) <> "" Then
    '        Me.ado_tm_ids.Recordset.Fields(8 + j).Value = Me.im_l(j).Text
    '    End If
    'Next j
        
    Me.ado_tm_ids.Recordset.UpdateBatch adAffectCurrent
    Me.ado_tm_ids.Recordset.Requery
    Me.ado_tm_ids.Refresh
    Me.DataGrid2.Refresh
    Me.ado_tmimata.Recordset.Requery
    Me.ado_tmimata.Refresh
    Me.dt_tmimata.Refresh
    Me.ado_tmimata.Recordset.Filter = "[id_αθλητή] = '" & athlet_management.ado_athlites.Recordset.Fields(0).Value & "'"
    Me.ado_tmimata.Recordset.MoveLast
    Me.dt_tmimata.Col = 1
    If Me.ado_tmimata.Recordset.RecordCount > 0 Then
        Me.ado_tmimata.Caption = "Τμήμα " & Me.ado_tmimata.Recordset.AbsolutePosition & " από " & Me.ado_tmimata.Recordset.RecordCount
    End If
    '''
    
    Me.dt_tmimata.Columns(0).Caption = "Αθλητικό Έτος"
    Me.dt_tmimata.Columns(0).Width = 1000
    Me.dt_tmimata.Columns(1).Caption = "Κατηγορία Τμήματος"
    Me.dt_tmimata.Columns(1).Width = 1500
    Me.dt_tmimata.Columns(2).Caption = "Α Προπονητής"
    Me.dt_tmimata.Columns(2).Width = 1500
    Me.dt_tmimata.Columns(3).Caption = "Β Προπονητής"
    Me.dt_tmimata.Columns(3).Width = 1500
    For i = 4 To Me.ado_tmimata.Recordset.Fields.Count - 1
        Me.dt_tmimata.Columns(i).Visible = False
    Next i
    '''
    
        
    Me.save_command.Enabled = False
    Me.canc_bt.Enabled = True
    Me.insert_bt.Enabled = True
    Me.up_bt.Enabled = True
    
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
    Me.ado_tmimata.Recordset.Filter = s_string
    
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
        Me.ado_tmimata.Recordset.Sort = "[" & Trim(Me.ado_tmimata.Recordset.Fields(dt_athlites.Col).Name) & "]"
    Else
        Me.ado_tmimata.Recordset.Sort = "[" & Trim(Me.ado_tmimata.Recordset.Fields(defined_col).Name) & "]"
    End If
  

End Sub

Private Sub tmp_poso_eggrafis_LostFocus()
        
    If Me.tmp_poso_eggrafis.Text <> "0" Then
        If FormatCurrency(Val(Me.tmp_poso_eggrafis)) = False Then
            tmp_poso_eggrafis.Text = ""
            tmp_poso_eggrafis.SetFocus
        End If
    End If
    
End Sub

Private Sub tmp_poso_mina_LostFocus()
        
    If Me.tmp_poso_mina.Text <> "0" Then
        If FormatCurrency(Val(Me.tmp_poso_mina)) = False Then
            tmp_poso_mina.Text = ""
            tmp_poso_mina.SetFocus
        End If
    End If
    
End Sub

Private Sub up_bt_Click()

    Dim id As Integer
    
    If Not Me.ado_tmimata.Recordset.EOF And Not Me.ado_tm_ids.Recordset.EOF Then
        Me.ado_tm_ids.Recordset.MoveFirst
        If Me.ado_tmimata.Recordset.AbsolutePosition - 1 >= 1 Then
            Me.ado_tm_ids.Recordset.Move Me.ado_tmimata.Recordset.AbsolutePosition - 1
        End If
        If Me.co_ae.SelectedItem >= 1 Then
            Me.ado_ae.Recordset.MoveFirst
            Me.ado_ae.Recordset.Move Me.co_ae.SelectedItem - 1
            Me.ado_tm_ids.Recordset.Fields(1).Value = Me.ado_ae.Recordset.Fields(0).Value
        End If
        If Me.co_kt.SelectedItem >= 1 Then
            Me.ado_kt.Recordset.MoveFirst
            Me.ado_kt.Recordset.Move Me.co_kt.SelectedItem - 1
            Me.ado_tm_ids.Recordset.Fields(2).Value = Me.ado_kt.Recordset.Fields(0).Value
        End If
        If Me.co_propa.SelectedItem >= 1 Then
            Me.ado_prop.Recordset.MoveFirst
            Me.ado_prop.Recordset.Move Me.co_propa.SelectedItem - 1
            Me.ado_tm_ids.Recordset.Fields(3).Value = Me.ado_prop.Recordset.Fields(0).Value
        End If
        If Me.co_propb.SelectedItem >= 1 Then
            Me.ado_prop.Recordset.MoveFirst
            Me.ado_prop.Recordset.Move Me.co_propb.SelectedItem - 1
            Me.ado_tm_ids.Recordset.Fields(4).Value = Me.ado_prop.Recordset.Fields(0).Value
        End If
        If Trim(Me.tmp_poso_eggrafis.Text) <> "" Then
            Me.ado_tm_ids.Recordset.Fields(5).Value = Me.tmp_poso_eggrafis.Text
        End If
        If Trim(Me.tmp_poso_mina.Text) <> "" Then
            Me.ado_tm_ids.Recordset.Fields(6).Value = Me.tmp_poso_mina.Text
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
        'If Me.co_dimoi.Text <> "" Then
        '    Me.ado_tmimata.Recordset.Fields(7).Value = Me.co_dimoi.Text
        '        If Me.co_dimoi.MatchedWithList = False Then
        '        'ΕΙΣΑΓΩΓΗ ΝΕΟΥ ΔΗΜΟΥ
        '            If Not Me.ado_dimoi.Recordset.EOF Then
        '                Me.ado_dimoi.Recordset.MoveLast
        '                id = Me.ado_dimoi.Recordset![id_δήμου]
        '            End If
        '        Me.ado_dimoi.Recordset.AddNew
        '        Me.ado_dimoi.Recordset.Fields(0) = id + 1
        '        Me.ado_dimoi.Recordset.Fields(1) = Trim(Me.co_dimoi.Text)
        '        Me.ado_dimoi.Recordset.UpdateBatch adAffectCurrent
        '    End If
        'End If
        If Me.ado_tmimata.Recordset.AbsolutePosition >= 1 Then
            col_aff = Me.ado_tmimata.Recordset.AbsolutePosition
        Else
            col_aff = 1
        End If
        Me.ado_tm_ids.Recordset.UpdateBatch adAffectCurrent
        Me.ado_tm_ids.Recordset.Requery
        Me.ado_tm_ids.Refresh
        Me.DataGrid2.Refresh
        Me.ado_tmimata.Recordset.Requery
        Me.ado_tmimata.Refresh
        Me.dt_tmimata.Refresh
        If Not Me.ado_tmimata.Recordset.EOF Then
            Me.ado_tmimata.Recordset.Sort = "[" & Trim(Me.ado_tmimata.Recordset.Fields(10).Name) & "]"
            Me.ado_tm_ids.Recordset.Sort = "[" & Trim(Me.ado_tm_ids.Recordset.Fields(0).Name) & "]"
            Me.ado_tmimata.Recordset.MoveFirst
            Me.ado_tm_ids.Recordset.MoveFirst
            If Me.ado_tmimata.Recordset.RecordCount > 1 Then
                Me.ado_tmimata.Recordset.Move col_aff - 1
                'Me.ado_tm_ids.Recordset.Move col_aff - 1
            End If
            Me.dt_tmimata.Row = col_aff - 1
            Me.dt_tmimata.Col = 1
            Me.DataGrid2.Row = col_aff - 1
            Me.DataGrid2.Col = 1
        End If
        
        
        
        
        
        
        
        'Me.ado_tm_ids.Recordset.Close
        'Me.ado_tm_ids.Recordset.Open
        'Me.ado_tmimata.Recordset.Requery
        'Me.ado_tmimata.Recordset.Sort = "[" & Trim(Me.ado_tmimata.Recordset.Fields(10).Name) & "]"
        'Me.ado_tm_ids.Recordset.Sort = "[" & Trim(Me.ado_tm_ids.Recordset.Fields(0).Name) & "]"
        'If Me.ado_tmimata.Recordset.RecordCount > 0 Then
        '    Me.dt_tmimata.Row = col_aff - 1
        '    Me.dt_tmimata.Col = 1
        'End If
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
        MsgBox "Ακύρωση Ενημέρωσης! Απαιτείται πρώτσ προσθήκη εγγραφής!", vbCritical, "Μήνυμα Λάθους!"
        Me.co_ae.SetFocus
    End If
    '
    
End Sub
