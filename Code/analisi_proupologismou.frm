VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form anlisi_proupologismou 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Ανάλυση Προϋπολογισμού"
   ClientHeight    =   7905
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   16185
   LinkTopic       =   "entity_management"
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   16185
   Begin VB.TextBox txt_im_egkr 
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
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txt_im_liks 
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
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txt_im_en 
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
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Ανάλυση Προϋπολογισμού"
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
      Height          =   7575
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   15855
      Begin VB.CommandButton bt_refr_esoda 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ανανέωση"
         DisabledPicture =   "analisi_proupologismou.frx":0000
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   13320
         MaskColor       =   &H80000014&
         Picture         =   "analisi_proupologismou.frx":4E2E
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   6840
         Width           =   1095
      End
      Begin VB.CommandButton kl_bt 
         BackColor       =   &H80000014&
         Caption         =   "Κλείσιμο"
         DisabledPicture =   "analisi_proupologismou.frx":9C5C
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   14520
         MaskColor       =   &H80000014&
         Picture         =   "analisi_proupologismou.frx":F6D4
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   6840
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Off Line Ενημέρωση"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   0
         Picture         =   "analisi_proupologismou.frx":1514C
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   6840
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Τρέχων Προϋπολογισμός"
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
         Height          =   1215
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   15640
         Begin VB.TextBox txt_egkr_gs 
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
            Left            =   6600
            TabIndex        =   20
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txt_perigrafi 
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
            Left            =   1080
            TabIndex        =   19
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox txt_id 
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
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ημερομηνία:"
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
            Index           =   6
            Left            =   5520
            TabIndex        =   26
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Έγκριση ΓΣ:"
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
            Left            =   5520
            TabIndex        =   25
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Από:"
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
            Left            =   3600
            TabIndex        =   24
            Top             =   360
            Width           =   525
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Περιγραφή:"
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
            TabIndex        =   23
            Top             =   720
            Width           =   885
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Κωδικός:"
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
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Έως:"
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
            Left            =   3600
            TabIndex        =   21
            Top             =   720
            Width           =   525
         End
      End
      Begin MSDataGridLib.DataGrid dt_pr_eksod 
         Bindings        =   "analisi_proupologismou.frx":1548E
         Height          =   255
         Left            =   8760
         TabIndex        =   4
         Top             =   4440
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
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
      Begin MSDataGridLib.DataGrid dt_eksoda 
         Bindings        =   "analisi_proupologismou.frx":154A9
         Height          =   255
         Left            =   9000
         TabIndex        =   5
         Top             =   3120
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
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
      Begin MSDataGridLib.DataGrid dt_esoda 
         Bindings        =   "analisi_proupologismou.frx":154C2
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   2760
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
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
      Begin MSDataGridLib.DataGrid dt_pr_esod 
         Bindings        =   "analisi_proupologismou.frx":154DA
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   3960
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
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
      Begin MSAdodcLib.Adodc ado_pr_esod 
         Height          =   330
         Left            =   1920
         Top             =   3600
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         RecordSource    =   "ΑνάλυσηΠροϋπολογισμούΕσόδων"
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
      Begin MSAdodcLib.Adodc ado_pr_eksod 
         Height          =   330
         Left            =   8880
         Top             =   3960
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         RecordSource    =   "ΑνάλυσηΠροϋπολογισμούΕξόδων"
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
      Begin VB.Frame Frame4 
         Caption         =   "Προυπολογισμός Εσόδων"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   5535
         Left            =   0
         TabIndex        =   8
         Top             =   1200
         Width           =   7845
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Bodoni MT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   360
            Left            =   6390
            TabIndex        =   45
            Top             =   4580
            Width           =   930
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Bodoni MT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   360
            Left            =   4800
            TabIndex        =   44
            Top             =   4580
            Width           =   1590
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Bodoni MT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   360
            Left            =   3440
            TabIndex        =   43
            Top             =   4580
            Width           =   1360
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Bodoni MT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   345
            Left            =   2420
            TabIndex        =   42
            Top             =   4580
            Width           =   1015
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "analisi_proupologismou.frx":154F4
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
            Left            =   7380
            MaskColor       =   &H00008000&
            Picture         =   "analisi_proupologismou.frx":19EF9
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Εκτύπωση Προϋπολογισμού Εσόδων"
            Top             =   2040
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton bt_ins_esod 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "+"
            DisabledPicture =   "analisi_proupologismou.frx":1E8FE
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
            Left            =   7380
            MaskColor       =   &H000000FF&
            Picture         =   "analisi_proupologismou.frx":1EABD
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Προσθήκη Νέου Εσόδου στον Τρέχοντα Προϋπολογισμό"
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton bt_up_esoda 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "V"
            DisabledPicture =   "analisi_proupologismou.frx":1F148
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
            Left            =   7380
            MaskColor       =   &H000000FF&
            Picture         =   "analisi_proupologismou.frx":2086D
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Ενημέρωση στη ΒΔ των Αλλαγών"
            Top             =   960
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton bt_del_esoda 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Caption         =   "x"
            DisabledPicture =   "analisi_proupologismou.frx":21F92
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
            Left            =   7380
            MaskColor       =   &H000000FF&
            Picture         =   "analisi_proupologismou.frx":220B0
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Διαγραφή Επιλεγμένου Εσόδου από τον Τρέχοντα Προϋπολογισμό"
            Top             =   1320
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton bt_can_esoda 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            Caption         =   "<-"
            DisabledPicture =   "analisi_proupologismou.frx":22667
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
            Left            =   7380
            MaskColor       =   &H000000FF&
            Picture         =   "analisi_proupologismou.frx":22925
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Αναίρεση τελευταίας ενέργειας"
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin MSAdodcLib.Adodc ado_esoda 
            Height          =   330
            Left            =   2280
            Top             =   1200
            Visible         =   0   'False
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
            CommandType     =   8
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
            RecordSource    =   "SELECT ΤύποιΕσόδων.id_τύπου_εσόδου, ΤύποιΕσόδων.περιγραφή FROM ΤύποιΕσόδων ORDER BY ΤύποιΕσόδων.περιγραφή"
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
         Begin MSAdodcLib.Adodc ado_ap_esod 
            Height          =   330
            Left            =   120
            Top             =   5040
            Width           =   7245
            _ExtentX        =   12779
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   4
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
            RecordSource    =   "Για_την_ανάλυση_προϋπολογισμού_εσόδων"
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
         Begin MSDataGridLib.DataGrid dt_an_esod 
            Bindings        =   "analisi_proupologismou.frx":22BE3
            Height          =   3975
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
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
         Begin MSDataListLib.DataCombo co_esoda 
            Bindings        =   "analisi_proupologismou.frx":22BFD
            Height          =   345
            Left            =   435
            TabIndex        =   15
            Top             =   285
            Width           =   2045
            _ExtentX        =   3598
            _ExtentY        =   609
            _Version        =   393216
            IntegralHeight  =   0   'False
            Locked          =   -1  'True
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
         Begin VB.TextBox txt_poso_esod 
            Alignment       =   1  'Right Justify
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
            Left            =   2450
            TabIndex        =   14
            Top             =   285
            Width           =   1015
         End
         Begin VB.TextBox Text1 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
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
            Left            =   3440
            TabIndex        =   38
            Top             =   285
            Visible         =   0   'False
            Width           =   1480
         End
         Begin VB.TextBox Text2 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
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
            Left            =   4920
            TabIndex        =   39
            Top             =   285
            Visible         =   0   'False
            Width           =   1490
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Σύνολα:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   345
            Left            =   0
            TabIndex        =   16
            Top             =   4635
            Width           =   2325
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Προυπολογισμός Εξόδων"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   5535
         Left            =   7800
         TabIndex        =   27
         Top             =   1200
         Width           =   7845
         Begin VB.TextBox Text12 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Bodoni MT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   6360
            TabIndex        =   49
            Top             =   4560
            Width           =   970
         End
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Bodoni MT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   4880
            TabIndex        =   48
            Top             =   4560
            Width           =   1490
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Bodoni MT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   3420
            TabIndex        =   47
            Top             =   4560
            Width           =   1460
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Bodoni MT"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   2430
            TabIndex        =   46
            Top             =   4560
            Width           =   1000
         End
         Begin VB.CommandButton bt_can_eksoda 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            Caption         =   "<-"
            DisabledPicture =   "analisi_proupologismou.frx":22C15
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
            Left            =   7380
            MaskColor       =   &H000000FF&
            Picture         =   "analisi_proupologismou.frx":22ED3
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Αναίρεση τελευταίας ενέργειας"
            Top             =   1680
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "analisi_proupologismou.frx":23191
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
            Left            =   7380
            MaskColor       =   &H00008000&
            Picture         =   "analisi_proupologismou.frx":27B96
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Εκτύπωση Προϋπολογισμού Εξόδων"
            Top             =   2040
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton bt_del_eksoda 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Caption         =   "x"
            DisabledPicture =   "analisi_proupologismou.frx":2C59B
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
            Left            =   7380
            MaskColor       =   &H000000FF&
            Picture         =   "analisi_proupologismou.frx":2C6B9
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Διαγραφή Επιλεγμένου Εξόδου από τον Τρέχοντα Προϋπολογισμό"
            Top             =   1320
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton bt_up_eksoda 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "V"
            DisabledPicture =   "analisi_proupologismou.frx":2CC70
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
            Left            =   7380
            MaskColor       =   &H000000FF&
            Picture         =   "analisi_proupologismou.frx":2E395
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Ενημέρωση στη ΒΔ των Αλλαγών"
            Top             =   960
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.CommandButton bt_ins_eksod 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "+"
            DisabledPicture =   "analisi_proupologismou.frx":2FABA
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
            Left            =   7380
            MaskColor       =   &H000000FF&
            Picture         =   "analisi_proupologismou.frx":2FC79
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Προσθήκη Νέου Εξόδου στον Τρέχοντα Προϋπολογισμό"
            Top             =   600
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin MSAdodcLib.Adodc ado_eksoda 
            Height          =   330
            Left            =   2280
            Top             =   1320
            Visible         =   0   'False
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
            CommandType     =   8
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
            RecordSource    =   "SELECT ΤύποιΕξόδων.id_τύπου_εξόδου, ΤύποιΕξόδων.περιγραφή FROM ΤύποιΕξόδων ORDER BY ΤύποιΕξόδων.περιγραφή"
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
         Begin MSAdodcLib.Adodc ado_ap_eksod 
            Height          =   330
            Left            =   120
            Top             =   5040
            Width           =   7260
            _ExtentX        =   12806
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
            RecordSource    =   "Για_την_ανάλυση_προϋπολογισμού_εξόδων"
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
         Begin MSDataGridLib.DataGrid dt_an_eksod 
            Bindings        =   "analisi_proupologismou.frx":30304
            Height          =   3975
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   0   'False
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
         Begin MSDataListLib.DataCombo co_eksoda 
            Bindings        =   "analisi_proupologismou.frx":3031F
            Height          =   345
            Left            =   415
            TabIndex        =   33
            Top             =   285
            Width           =   2045
            _ExtentX        =   3598
            _ExtentY        =   609
            _Version        =   393216
            IntegralHeight  =   0   'False
            Locked          =   -1  'True
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
         Begin VB.TextBox txt_poso_eksod 
            Alignment       =   1  'Right Justify
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
            Left            =   2430
            TabIndex        =   28
            Top             =   285
            Width           =   1015
         End
         Begin VB.TextBox Text3 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
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
            Left            =   3440
            TabIndex        =   40
            Top             =   285
            Visible         =   0   'False
            Width           =   1480
         End
         Begin VB.TextBox Text4 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00 ""€"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1032
               SubFormatType   =   0
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
            TabIndex        =   41
            Top             =   285
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Σύνολα:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   1520
            TabIndex        =   34
            Top             =   4640
            Width           =   885
         End
      End
   End
End
Attribute VB_Name = "anlisi_proupologismou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id_τμ, defined_col As Integer
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
'
Public col_affected, new_addition, new_addition_eks, is_to_delete_eks, is_to_delete As Integer

Private Sub bt_ins_athl_Click()

    new_addition = 1
    co_esoda.Text = ""
    txt_poso_esod.Text = ""
    bt_del_athl.Enabled = False
    bt_up_esoda.Enabled = True
    bt_can_athl.Enabled = True
    co_esoda.SetFocus
    
End Sub

Private Sub ado_ap_eksod_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    If ado_ap_eksod.Recordset.AbsolutePosition >= 1 And ado_ap_eksod.Recordset.AbsolutePosition <= ado_ap_eksod.Recordset.RecordCount Then
        If Not ado_ap_eksod.Recordset.EOF Then
          If Trim(ado_ap_eksod.Recordset.Fields(2).Value) <> "" Then
                If ado_eksoda.Recordset.RecordCount >= 1 Then
                    ado_eksoda.Recordset.MoveFirst
                    ado_eksoda.Recordset.Find "[id_τύπου_εξόδου] = '" & ado_ap_eksod.Recordset.Fields(4).Value & "'"
                    If Not ado_eksoda.Recordset.EOF Then
                        co_eksoda.Text = ado_eksoda.Recordset.Fields(1).Value
                    End If
                End If
            Else
                co_eksoda.Text = ""
            End If
            If Trim(ado_ap_eksod.Recordset.Fields(3).Value) <> "" Then
                txt_poso_eksod.Text = ado_ap_eksod.Recordset.Fields(3).Value
                txt_poso_eksod.Text = FormatCurrency(txt_poso_eksod.Text, 2, vbTrue, , vbTrue)
            Else
                txt_poso_eksod.Text = ""
                txt_poso_eksod.Text = FormatCurrency(txt_poso_eksod.Text, 2, vbTrue, , vbTrue)
                txt_poso_eksod.Enabled = False
            End If
            ado_ap_eksod.Caption = "Έξοδα " & ado_ap_eksod.Recordset.AbsolutePosition & " από " & ado_ap_eksod.Recordset.RecordCount
            bt_up_eksoda.Enabled = True
            bt_del_eksoda.Enabled = True
        End If
    End If
    
End Sub

Private Sub ado_ap_esod_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    If ado_ap_esod.Recordset.AbsolutePosition >= 1 And ado_ap_esod.Recordset.AbsolutePosition <= ado_ap_esod.Recordset.RecordCount Then
        If Not ado_ap_esod.Recordset.EOF Then
          If Trim(ado_ap_esod.Recordset.Fields(2).Value) <> "" Then
                If Not ado_esoda.Recordset.EOF Then
                    ado_esoda.Recordset.MoveFirst
                    ado_esoda.Recordset.Find "[id_τύπου_εσόδου] = '" & ado_ap_esod.Recordset.Fields(4).Value & "'"
                    If Not ado_esoda.Recordset.EOF Then
                        co_esoda.Text = ado_esoda.Recordset.Fields(1).Value
                    End If
                End If
            Else
                co_esoda.Text = ""
            End If
            If Trim(ado_ap_esod.Recordset.Fields(3).Value) <> "" Then
                txt_poso_esod.Text = ado_ap_esod.Recordset.Fields(3).Value
                txt_poso_esod.Text = FormatCurrency(txt_poso_esod.Text, 2, vbTrue, , vbTrue)
            Else
                txt_poso_esod.Text = ""
                txt_poso_esod.Text = FormatCurrency(txt_poso_esod.Text, 2, vbTrue, , vbTrue)
            End If
            ado_ap_esod.Caption = "Έσοδα " & ado_ap_esod.Recordset.AbsolutePosition & " από " & ado_ap_esod.Recordset.RecordCount
            bt_up_esoda.Enabled = True
            bt_del_esoda.Enabled = True
        End If
    End If
    
End Sub

Private Sub bt_can_eksoda_Click()

    If Not ado_ap_eksod.Recordset.EOF Then
        pr_c = ado_ap_eksod.Recordset.AbsolutePosition
        ado_ap_eksod.Recordset.MoveFirst
        If pr_c >= 1 Then
            ado_ap_eksod.Recordset.Move pr_c - 1
        End If
        Me.bt_del_eksoda.Enabled = True
    Else
        co_eksoda.Text = ""
        txt_poso_eksod.Text = ""
    End If
    Me.bt_up_eksoda.Enabled = False
    Me.bt_ins_eksod.Enabled = True
    Me.bt_can_eksoda.Enabled = False
    new_addition_eks = 0

End Sub

Private Sub bt_can_esoda_Click()

    If Not ado_ap_esod.Recordset.EOF Then
        pr_c = ado_ap_esod.Recordset.AbsolutePosition
        ado_ap_esod.Recordset.MoveFirst
        If pr_c >= 1 Then
            ado_ap_esod.Recordset.Move pr_c - 1
        End If
        Me.bt_del_esoda.Enabled = True
    Else
        co_esoda.Text = ""
        txt_poso_esod.Text = ""
    End If
    Me.bt_up_esoda.Enabled = False
    Me.bt_ins_esod.Enabled = True
    Me.bt_can_esoda.Enabled = False
    new_addition = 0


End Sub

Private Sub bt_del_eksoda_Click()
    
    Dim ms As String
    Dim rec_index As Integer

    If Not ado_ap_eksod.Recordset.EOF Then
        ms = MsgBox("Είσαι σίγουρος; (ΝΑΙ ή ΟΧΙ)", vbYesNo, "Παράθυρο διαγραφής")
        If ms = 6 Then
            is_to_delete_eks = 1
            If ado_ap_eksod.Recordset.AbsolutePosition = ado_ap_eksod.Recordset.RecordCount Then
                rec_index = ado_ap_eksod.Recordset.AbsolutePosition - 1
            Else
                rec_index = ado_ap_eksod.Recordset.AbsolutePosition
            End If
            
            ado_pr_eksod.Recordset.Find "[id] = '" & ado_ap_eksod.Recordset.Fields(0).Value & "'"
            
            ado_pr_eksod.Recordset.Delete
            ado_pr_eksod.Recordset.Requery
            dt_pr_eksod.Refresh
            ado_pr_eksod.Refresh
            ado_ap_eksod.Recordset.Requery
            ado_ap_eksod.Refresh
            dt_an_eksod.Refresh
            
            ado_ap_eksod.Recordset.Filter = "[id_προϋπολογισμού] = '" & txt_id & "'"
            is_to_delete_esk = 0
            'Υπολογισμός Sum
            s = 0
            s2 = 0
            s3 = 0
            s4 = 0
            For i = 0 To ado_ap_eksod.Recordset.RecordCount - 1
                If i = 0 And Not ado_ap_eksod.Recordset.EOF Then
                    ado_ap_eksod.Recordset.MoveFirst
                End If
                s = s + ado_ap_eksod.Recordset.Fields(3).Value
                s2 = s2 + ado_ap_eksod.Recordset.Fields(5).Value
                s3 = s3 + ado_ap_eksod.Recordset.Fields(6).Value
                s4 = s4 + ado_ap_eksod.Recordset.Fields(7).Value
                ado_ap_eksod.Recordset.MoveNext
            Next i
            Text9.Text = FormatCurrency(s, 2, vbTrue, , vbTrue)
            Text10.Text = FormatCurrency(s2, 2, vbTrue, , vbTrue)
            Text11.Text = FormatCurrency(s3, 2, vbTrue, , vbTrue)
            Text12.Text = FormatCurrency(s4, 2, vbTrue, , vbTrue)
            '
            If ado_ap_eksod.Recordset.RecordCount >= 1 Then
                ado_ap_eksod.Recordset.MoveFirst
                ado_ap_eksod.Recordset.Move rec_index - 1, 0
                ado_ap_eksod.Recordset.MovePrevious
                ado_ap_eksod.Recordset.MoveNext
            Else
                co_eksoda.Text = ""
                txt_poso_eksod.Text = ""
                ado_ap_eksod.Caption = "Έξοδα 0 από 0"
                txt_poso_eksod.Enabled = False
                bt_up_eksoda.Enabled = False
                bt_del_eksoda.Enabled = False
            End If
            dt_an_eksod.Columns(0).Visible = False
            dt_an_eksod.Columns(1).Visible = False
            dt_an_eksod.Columns(2).Caption = "Περιγραφή Εξόδων"
            dt_an_eksod.Columns(2).Width = 2000
            dt_an_eksod.Columns(3).Alignment = dbgRight
            dt_an_eksod.Columns(3).Caption = "     Ποσό ΠΥ"
            dt_an_eksod.Columns(3).Width = 1000
            dt_an_eksod.Columns(3).NumberFormat = "currency"
            dt_an_eksod.Columns(4).Visible = False
            dt_an_eksod.Columns(5).Caption = "Ποσό Κινήσεων"
            dt_an_eksod.Columns(5).Width = 1335
            dt_an_eksod.Columns(5).NumberFormat = "currency"
            dt_an_eksod.Columns(5).Alignment = dbgRight
            dt_an_eksod.Columns(6).Caption = "Ποσό Δεσμεύσεων"
            dt_an_eksod.Columns(6).Width = 1575
            dt_an_eksod.Columns(6).NumberFormat = "currency"
            dt_an_eksod.Columns(6).Alignment = dbgRight
            dt_an_eksod.Columns(7).Caption = "  Υπόλοιπο"
            dt_an_eksod.Columns(7).Width = 975
            dt_an_eksod.Columns(7).NumberFormat = "currency"
            dt_an_eksod.Columns(7).Alignment = dbgRight
            If txt_poso_eksod.Text <> "" Then
                txt_poso_eksod.Text = FormatCurrency(txt_poso_eksod.Text, 2, vbTrue, , vbTrue)
            Else
                txt_poso_eksod.Text = ""
            End If
        Else
            MsgBox "ΑΚΥΡΩΣΗ ΔΙΑΓΡΑΦΗΣ", , "Μήνυμα Προειδοποίησης!"
        End If
    Else
        MsgBox "Δεν υπάρχει Εγγραφή για ΔΙΑΓΡΑΦΗ", , "Μήνυμα Προειδοποίησης!"
    End If
    
End Sub

Private Sub bt_del_esoda_Click()
    
    Dim ms As String
    Dim rec_index As Integer

    If Not ado_ap_esod.Recordset.EOF Then
        ms = MsgBox("Είσαι σίγουρος; (ΝΑΙ ή ΟΧΙ)", vbYesNo, "Παράθυρο διαγραφής")
        If ms = 6 Then
            is_to_delete = 1
            If ado_ap_esod.Recordset.AbsolutePosition = ado_ap_esod.Recordset.RecordCount Then
                rec_index = ado_ap_esod.Recordset.AbsolutePosition - 1
            Else
                rec_index = ado_ap_esod.Recordset.AbsolutePosition
            End If
            
            ado_pr_esod.Recordset.Find "[id] = '" & ado_ap_esod.Recordset.Fields(0).Value & "'"
            
            ado_pr_esod.Recordset.Delete
            ado_pr_esod.Recordset.Requery
            dt_pr_esod.Refresh
            ado_pr_esod.Refresh
            ado_ap_esod.Recordset.Requery
            ado_ap_esod.Refresh
            dt_an_esod.Refresh
            
            ado_ap_esod.Recordset.Filter = "[id_προϋπολογισμού] = '" & txt_id & "'"
            is_to_delete = 0
            'Υπολογισμός Sum
            s = 0
            s2 = 0
            s3 = 0
            s4 = 0
            For i = 0 To ado_ap_esod.Recordset.RecordCount - 1
                If i = 0 And Not ado_ap_esod.Recordset.EOF Then
                    ado_ap_esod.Recordset.MoveFirst
                End If
                s = s + ado_ap_esod.Recordset.Fields(3).Value
                s2 = s2 + ado_ap_esod.Recordset.Fields(5).Value
                s3 = s3 + ado_ap_esod.Recordset.Fields(6).Value
                s4 = s4 + ado_ap_esod.Recordset.Fields(7).Value
                ado_ap_esod.Recordset.MoveNext
            Next i
            Text5.Text = FormatCurrency(s, 2, vbTrue, , vbTrue)
            Text6.Text = FormatCurrency(s2, 2, vbTrue, , vbTrue)
            Text7.Text = FormatCurrency(s3, 2, vbTrue, , vbTrue)
            Text8.Text = FormatCurrency(s4, 2, vbTrue, , vbTrue)
            '
            If ado_ap_esod.Recordset.RecordCount >= 1 Then
                ado_ap_esod.Recordset.MoveFirst
                ado_ap_esod.Recordset.Move rec_index - 1, 0
                ado_ap_esod.Recordset.MovePrevious
                ado_ap_esod.Recordset.MoveNext
            Else
                co_esoda.Text = ""
                txt_poso_esod.Text = ""
                ado_ap_esod.Caption = "Έσοδα 0 από 0"
                txt_poso_esod.Enabled = False
                bt_up_esoda.Enabled = False
                bt_del_esoda.Enabled = False
            End If
            dt_an_esod.Columns(0).Visible = False
            dt_an_esod.Columns(1).Visible = False
            dt_an_esod.Columns(2).Caption = "Περιγραφή Εσόδων"
            dt_an_esod.Columns(2).Width = 2000
            dt_an_esod.Columns(3).Alignment = dbgRight
            dt_an_esod.Columns(3).Caption = "     Ποσό ΠΥ"
            dt_an_esod.Columns(3).Width = 1000
            dt_an_esod.Columns(3).NumberFormat = "currency"
            dt_an_esod.Columns(4).Visible = False
            dt_an_esod.Columns(5).Caption = "Ποσό Κινήσεων"
            dt_an_esod.Columns(5).Width = 1335
            dt_an_esod.Columns(5).NumberFormat = "currency"
            dt_an_esod.Columns(5).Alignment = dbgRight
            dt_an_esod.Columns(6).Caption = "Ποσό Δεσμεύσεων"
            dt_an_esod.Columns(6).Width = 1575
            dt_an_esod.Columns(6).NumberFormat = "currency"
            dt_an_esod.Columns(6).Alignment = dbgRight
            dt_an_esod.Columns(7).Caption = "  Υπόλοιπο"
            dt_an_esod.Columns(7).Width = 975
            dt_an_esod.Columns(7).NumberFormat = "currency"
            dt_an_esod.Columns(7).Alignment = dbgRight
            If txt_poso_esod.Text <> "" Then
                txt_poso_esod.Text = FormatCurrency(txt_poso_esod.Text, 2, vbTrue, , vbTrue)
            Else
                txt_poso_esod.Text = ""
            End If
        Else
            MsgBox "ΑΚΥΡΩΣΗ ΔΙΑΓΡΑΦΗΣ", , "Μήνυμα Προειδοποίησης!"
        End If
    Else
        MsgBox "Δεν υπάρχει Εγγραφή για ΔΙΑΓΡΑΦΗ", , "Μήνυμα Προειδοποίησης!"
    End If

End Sub

Private Sub bt_ins_eksod_Click()

    txt_poso_eksod.Enabled = True

    new_addition_eks = 1
    co_eksoda.Text = ""
    txt_poso_eksod.Text = ""
    bt_del_eksoda.Enabled = False
    bt_up_eksoda.Enabled = True
    bt_can_eksoda.Enabled = True
    co_eksoda.Locked = False
    co_eksoda.SetFocus
    
End Sub

Private Sub bt_ins_esod_Click()

    txt_poso_esod.Enabled = True
    
    new_addition = 1
    co_esoda.Text = ""
    txt_poso_esod.Text = ""
    bt_del_esoda.Enabled = False
    co_esoda.Locked = False
    bt_up_esoda.Enabled = True
    bt_can_esoda.Enabled = True
    co_esoda.SetFocus
    
End Sub

Private Sub bt_print_Click()
    
    f_n = Me.Name
    frm_pr_tm.Show

End Sub

Private Sub co_eksoda_Click(Area As Integer)

    bt_can_eksoda.Enabled = True

End Sub

Private Sub co_esoda_Click(Area As Integer)

    bt_can_esoda.Enabled = True

End Sub

Private Sub Command1_Click()

    Rep_analisi_proypologismou_esodwn.Orientation = rptOrientPortrait
    Rep_analisi_proypologismou_esodwn.Show
    
End Sub

Private Sub Command2_Click()
    
    Rep_analisi_proypologismou_eksodwn.Orientation = rptOrientPortrait
    Rep_analisi_proypologismou_eksodwn.Show
    
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

Private Sub dt_tmimata_HeadClick(ByVal ColIndex As Integer)

    defined_col = ColIndex

End Sub

Private Sub bt_refr_esoda_Click()

    Unload Me
    anlisi_proupologismou.Show
    
End Sub

Private Sub bt_up_eksoda_Click()
    
    'ΑΡΧΙΚΑ ΠΡΕΠΕΙ ΝΑ ΕΛΕΓΧΕΙ ΑΝ Η ΚΑΤΗΓΟΡΙΑ ΕΞΟΔΟΥ ΕΙΝΑΙ ΗΔΗ ΣΤΟΝ ΠΡΟΥΠΟΛΟΓΙΜΟ - ΑΝ ΕΙΝΑΙ ΔΕΝ ΣΥΝΕΧΙΖΕΤΑΙ Η ΕΝΗΜΕΡΩΣΗ
    col_affected = ado_ap_eksod.Recordset.AbsolutePosition
    'Εύρεση id_τύπου_εξόδου
    Dim id_eks As Integer
    id_eks = 0
    If Trim(Me.co_eksoda.Text) <> "" Then
        If Me.co_eksoda.SelectedItem >= 1 Then
            ado_eksoda.Recordset.MoveFirst
            ado_eksoda.Recordset.Move Me.co_eksoda.SelectedItem - 1
            id_eks = ado_eksoda.Recordset.Fields(0).Value
        End If
    End If
    ado_pr_eksod.Recordset.Filter = "[id_προϋπολογισμού] Like '" & txt_id.Text & "'"
    ado_pr_eksod.Recordset.Find "[id_τύπου_εξόδου] = " & id_eks
    If Not ado_pr_eksod.Recordset.EOF Then
        If new_addition_eks = 1 Then
            MsgBox "Η Κατηγορία <<" & Me.co_eksoda.Text & ">> είναι ήδη καταχωρημένη στον τρέχων προϋπολογισμό! Προσπαθήστε ξανά!", , "ΠΡΟΫΠΟΛΟΓΙΣΜΟΣ ΕΞΟΔΩΝ"
            If ado_ap_eksod.Recordset.RecordCount >= 1 Then
                ado_ap_eksod.Recordset.MoveFirst
                If col_affected > 1 Then
                    ado_ap_eksod.Recordset.Move col_affected - 1
                End If
                dt_an_eksod.Col = 1
                If ado_ap_eksod.Recordset.RecordCount > 0 Then
                    ado_ap_eksod.Caption = "Έξοδα " & ado_ap_eksod.Recordset.AbsolutePosition & " από " & ado_ap_eksod.Recordset.RecordCount
                End If
            End If
            co_eksoda.Locked = True
            new_addition_eks = 0
            Exit Sub
        End If
    End If
    ado_pr_eksod.Recordset.Requery
    ado_pr_eksod.Refresh
    dt_an_eksod.Refresh
    '
    '
    ''''''''''''''''ΣΥΝΕΧΙΖΕΙ ΟΜΑΛΑ Η ΕΝΗΜΕΡΩΣΗ''''''''''''''''
    ''
    'ΠΡΟΣΘΗΚΗ ΝΕΟΥ ΕΞΟΔΟΥ ΣΕ ΑΝΑΛΥΣΗ ΠΡΟΥΠΟΛΟΓΙΣΜΟΥ
    If new_addition_eks = 1 Then
        new_addition_eks = 0
        co_eksoda.Locked = True
        bt_up_eksoda.Enabled = True
        bt_ins_eksod.Enabled = True
        bt_del_eksoda.Enabled = True
        bt_can_eksoda.Enabled = False
        'Να βρω το υποψήφιο id για την αποθήκευση στον πίνακα ΑΝΑΛΥΣΗ_ΠΡΟΥΠΟΛΟΓΙΣΜΟΥ_ΕΞΟΔΩΝ
        If Not ado_pr_eksod.Recordset.EOF Then
            ado_pr_eksod.Recordset.Sort = "[id]"
            ado_pr_eksod.Recordset.MoveLast
            id_πε = ado_pr_eksod.Recordset.Fields(0).Value
            id_πε = id_πε + 1
        Else
            id_πε = 1
        End If
        'Αποθήκευση του ΝΕΟΥ ΕΞΟΔΟΥ για τον τρέχων προϋπολογισμό
        ado_pr_eksod.Recordset.AddNew
        'Αποθήκευση id
        ado_pr_eksod.Recordset.Fields(0).Value = id_πε
        'Αποθήκευση id_προϋπολογισμού
        ado_pr_eksod.Recordset.Fields(1).Value = txt_id
        'Αποθήκευση id_τύπου_εσόδου
        If Trim(Me.co_eksoda.Text) <> "" Then
            If Me.co_eksoda.SelectedItem >= 1 Then
                ado_eksoda.Recordset.MoveFirst
                ado_eksoda.Recordset.Move Me.co_eksoda.SelectedItem - 1
                ado_pr_eksod.Recordset.Fields(2).Value = ado_eksoda.Recordset.Fields(0).Value
            End If
        Else
            ado_pr_eksod.Recordset.Fields(2).Value = ""
        End If
        If Trim(txt_poso_eksod.Text) <> "" Then
            ado_pr_eksod.Recordset.Fields(3).Value = Val(txt_poso_eksod.Text)
        End If
        ado_pr_eksod.Recordset.UpdateBatch adAffectCurrent
        ado_pr_eksod.Recordset.Requery
        ado_pr_eksod.Refresh
        ado_ap_eksod.Recordset.Requery
        ado_ap_eksod.Refresh
        dt_an_eksod.Refresh
        '
        ado_ap_eksod.Recordset.Filter = "[id_προϋπολογισμού] = '" & txt_id & "'"
        'Υπολογισμός Sum
        s = 0
        s2 = 0
        s3 = 0
        s4 = 0
        For i = 0 To ado_ap_eksod.Recordset.RecordCount - 1
            If i = 0 And Not ado_ap_eksod.Recordset.EOF Then
                ado_ap_eksod.Recordset.MoveFirst
            End If
            s = s + ado_ap_eksod.Recordset.Fields(3).Value
            s2 = s2 + ado_ap_eksod.Recordset.Fields(5).Value
            s3 = s3 + ado_ap_eksod.Recordset.Fields(6).Value
            s4 = s4 + ado_ap_eksod.Recordset.Fields(7).Value
            ado_ap_eksod.Recordset.MoveNext
        Next i
        Text9.Text = FormatCurrency(s, 2, vbTrue, , vbTrue)
        Text10.Text = FormatCurrency(s2, 2, vbTrue, , vbTrue)
        Text11.Text = FormatCurrency(s3, 2, vbTrue, , vbTrue)
        Text12.Text = FormatCurrency(s4, 2, vbTrue, , vbTrue)
        '
        If ado_ap_eksod.Recordset.RecordCount >= 1 Then
            ado_ap_eksod.Recordset.MoveLast
            dt_an_eksod.Col = 1
            If ado_ap_eksod.Recordset.RecordCount > 0 Then
                ado_ap_eksod.Caption = "Έξοδα " & ado_ap_eksod.Recordset.AbsolutePosition & " από " & ado_ap_eksod.Recordset.RecordCount
            End If
        End If
        '''
        dt_an_eksod.Columns(0).Visible = False
        dt_an_eksod.Columns(1).Visible = False
        dt_an_eksod.Columns(2).Caption = "Περιγραφή Εξόδων"
        dt_an_eksod.Columns(2).Width = 2000
        dt_an_eksod.Columns(3).Alignment = dbgRight
        dt_an_eksod.Columns(3).Caption = "     Ποσό ΠΥ"
        dt_an_eksod.Columns(3).Width = 1000
        dt_an_eksod.Columns(3).NumberFormat = "currency"
        dt_an_eksod.Columns(4).Visible = False
        dt_an_eksod.Columns(5).Caption = "Ποσό Κινήσεων"
        dt_an_eksod.Columns(5).Width = 1335
        dt_an_eksod.Columns(5).NumberFormat = "currency"
        dt_an_eksod.Columns(5).Alignment = dbgRight
        dt_an_eksod.Columns(6).Caption = "Ποσό Δεσμεύσεων"
        dt_an_eksod.Columns(6).Width = 1575
        dt_an_eksod.Columns(6).NumberFormat = "currency"
        dt_an_eksod.Columns(6).Alignment = dbgRight
        dt_an_eksod.Columns(7).Caption = "  Υπόλοιπο"
        dt_an_eksod.Columns(7).Width = 975
        dt_an_eksod.Columns(7).NumberFormat = "currency"
        dt_an_eksod.Columns(7).Alignment = dbgRight
        txt_poso_eksod.Text = FormatCurrency(txt_poso_eksod.Text, 2, vbTrue, , vbTrue)
    'ΤΕΛΟΣ --- ΠΡΟΣΘΗΚΗ ΝΕΟΥ ΕΞΟΔΟΥ ΣΕ ΠΡΟΥΠΟΛΟΓΙΣΜΟ
    Else
        'ΑΡΧΗ - ΕΝΗΜΕΡΩΣΗ ΠΑΛΙΟΥ ΕΞΟΔΟΥ ΣΕ ΠΡΟΥΠΟΛΟΓΙΣΜΟ
        new_addition_eks = 0
        bt_up_eksoda.Enabled = False
        bt_ins_eksod.Enabled = True
        bt_del_eksoda.Enabled = True
        bt_can_eksoda.Enabled = False
        
        ado_pr_eksod.Recordset.Find "[id] = '" & ado_ap_eksod.Recordset.Fields(0).Value & "'"
        
        'Ενημέρωση id_τύπου_εξόδου
        If Trim(Me.co_eksoda.Text) <> "" Then
            If Me.co_eksoda.SelectedItem >= 1 Then
                ado_eksoda.Recordset.MoveFirst
                ado_eksoda.Recordset.Move Me.co_eksoda.SelectedItem - 1
                ado_pr_eksod.Recordset.Fields(2).Value = ado_eksoda.Recordset.Fields(0).Value
            End If
        Else
            ado_pr_eksod.Recordset.Fields(2).Value = ""
        End If
        'Ενημέρωση ποσού εξόδου
        If Trim(txt_poso_eksod.Text) <> "" Then
            ado_pr_eksod.Recordset.Fields(3).Value = Val(txt_poso_eksod.Text)
        End If
        '
        col_affected = ado_ap_eksod.Recordset.AbsolutePosition
        ado_pr_eksod.Recordset.UpdateBatch adAffectCurrent
        
        ado_pr_eksod.Recordset.Requery
        ado_pr_eksod.Refresh
        ado_ap_eksod.Recordset.Requery
        ado_ap_eksod.Refresh
        dt_an_eksod.Refresh
        '
        ado_ap_eksod.Recordset.Filter = "[id_προϋπολογισμού] = '" & txt_id & "'"
        'Υπολογισμός Sum
        s = 0
        s2 = 0
        s3 = 0
        s4 = 0
        For i = 0 To ado_ap_eksod.Recordset.RecordCount - 1
            If i = 0 And Not ado_ap_eksod.Recordset.EOF Then
                ado_ap_eksod.Recordset.MoveFirst
            End If
            s = s + ado_ap_eksod.Recordset.Fields(3).Value
            s2 = s2 + ado_ap_eksod.Recordset.Fields(5).Value
            s3 = s3 + ado_ap_eksod.Recordset.Fields(6).Value
            s4 = s4 + ado_ap_eksod.Recordset.Fields(7).Value
            ado_ap_eksod.Recordset.MoveNext
        Next i
        Text9.Text = FormatCurrency(s, 2, vbTrue, , vbTrue)
        Text10.Text = FormatCurrency(s2, 2, vbTrue, , vbTrue)
        Text11.Text = FormatCurrency(s3, 2, vbTrue, , vbTrue)
        Text12.Text = FormatCurrency(s4, 2, vbTrue, , vbTrue)
        '
        If ado_ap_eksod.Recordset.RecordCount >= 1 Then
            ado_ap_eksod.Recordset.MoveFirst
            If col_affected > 1 Then
                ado_ap_eksod.Recordset.Move col_affected - 1
            End If
            dt_an_eksod.Col = 1
            If ado_ap_eksod.Recordset.RecordCount > 0 Then
                ado_ap_eksod.Caption = "Έξοδα " & ado_ap_eksod.Recordset.AbsolutePosition & " από " & ado_ap_eksod.Recordset.RecordCount
            End If
        End If
        '''
        dt_an_eksod.Columns(0).Visible = False
        dt_an_eksod.Columns(1).Visible = False
        dt_an_eksod.Columns(2).Caption = "Περιγραφή Εξόδων"
        dt_an_eksod.Columns(2).Width = 2000
        dt_an_eksod.Columns(3).Alignment = dbgRight
        dt_an_eksod.Columns(3).Caption = "     Ποσό ΠΥ"
        dt_an_eksod.Columns(3).Width = 1000
        dt_an_eksod.Columns(3).NumberFormat = "currency"
        dt_an_eksod.Columns(4).Visible = False
        dt_an_eksod.Columns(5).Caption = "Ποσό Κινήσεων"
        dt_an_eksod.Columns(5).Width = 1335
        dt_an_eksod.Columns(5).NumberFormat = "currency"
        dt_an_eksod.Columns(5).Alignment = dbgRight
        dt_an_eksod.Columns(6).Caption = "Ποσό Δεσμεύσεων"
        dt_an_eksod.Columns(6).Width = 1575
        dt_an_eksod.Columns(6).NumberFormat = "currency"
        dt_an_eksod.Columns(6).Alignment = dbgRight
        dt_an_eksod.Columns(7).Caption = "  Υπόλοιπο"
        dt_an_eksod.Columns(7).Width = 975
        dt_an_eksod.Columns(7).NumberFormat = "currency"
        dt_an_eksod.Columns(7).Alignment = dbgRight
        txt_poso_eksod.Text = FormatCurrency(txt_poso_eksod.Text, 2, vbTrue, , vbTrue)
        'ΤΕΛΟΣ --- ΕΝΗΜΕΡΩΣΗ ΠΑΛΙΟΥ ΕΞΟΔΟΥ ΣΕ ΠΡΟΥΠΟΛΟΓΙΣΜΟ
    End If
    
End Sub

Private Sub bt_up_esoda_Click()
    
    
    'ΑΡΧΙΚΑ ΠΡΕΠΕΙ ΝΑ ΕΛΕΓΧΕΙ ΑΝ Η ΚΑΤΗΓΟΡΙΑ ΕΣΟΔΟΥ ΕΙΝΑΙ ΗΔΗ ΣΤΟΝ ΠΡΟΥΠΟΛΟΓΙΜΟ - ΑΝ ΕΙΝΑΙ ΔΕΝ ΣΥΝΕΧΙΖΕΤΑΙ Η ΕΝΗΜΕΡΩΣΗ
    col_affected = ado_ap_esod.Recordset.AbsolutePosition
    Dim old_per As String
    'Εύρεση id_τύπου_εσόδου
    Dim id_es As Integer
    id_es = 0
    If Trim(Me.co_esoda.Text) <> "" Then
        If Me.co_esoda.SelectedItem >= 1 Then
            ado_esoda.Recordset.MoveFirst
            ado_esoda.Recordset.Move Me.co_esoda.SelectedItem - 1
            id_es = ado_esoda.Recordset.Fields(0).Value
        End If
    End If
    ado_pr_esod.Recordset.Filter = "[id_προϋπολογισμού] Like '" & txt_id.Text & "'"
    ado_pr_esod.Recordset.Find "[id_τύπου_εσόδου] = " & id_es
    If Not ado_pr_esod.Recordset.EOF Then
        If new_addition = 1 Then
            MsgBox "Η Κατηγορία <<" & Me.co_esoda.Text & ">> είναι ήδη καταχωρημένη στον τρέχων προϋπολογισμό! Προσπαθήστε ξανά!", , "ΠΡΟΫΠΟΛΟΓΙΣΜΟΣ ΕΣΟΔΩΝ"
            If ado_ap_esod.Recordset.RecordCount >= 1 Then
                ado_ap_esod.Recordset.MoveFirst
                If col_affected > 1 Then
                    ado_ap_esod.Recordset.Move col_affected - 1
                End If
                dt_an_esod.Col = 1
                If ado_ap_esod.Recordset.RecordCount > 0 Then
                    ado_ap_esod.Caption = "Έσοδα " & ado_ap_esod.Recordset.AbsolutePosition & " από " & ado_ap_esod.Recordset.RecordCount
                End If
            End If
            co_esoda.Locked = True
            new_addition = 0
            Exit Sub
        End If
    End If
    ado_pr_esod.Recordset.Requery
    ado_pr_esod.Refresh
    dt_an_esod.Refresh
    '
    '
    ''''''''''''''''ΣΥΝΕΧΙΖΕΙ ΟΜΑΛΑ Η ΕΝΗΜΕΡΩΣΗ''''''''''''''''
    ''
    'ΠΡΟΣΘΗΚΗ ΝΕΟΥ ΕΣΟΔΟΥ ΣΕ ΑΝΑΛΥΣΗ ΠΡΟΥΠΟΛΟΓΙΣΜΟΥ
    If new_addition = 1 Then
        new_addition = 0
        co_esoda.Locked = True
        bt_up_esoda.Enabled = True
        bt_ins_esod.Enabled = True
        bt_del_esoda.Enabled = True
        bt_can_esoda.Enabled = False
        'Να βρω το υποψήφιο id για την αποθήκευση στον πίνακα ΑΝΑΛΥΣΗ_ΠΡΟΥΠΟΛΟΓΙΣΜΟΥ_ΕΣΟΔΩΝ
        If Not ado_pr_esod.Recordset.EOF Then
            ado_pr_esod.Recordset.Sort = "[id]"
            ado_pr_esod.Recordset.MoveLast
            id_πε = ado_pr_esod.Recordset.Fields(0).Value
            id_πε = id_πε + 1
        Else
            id_πε = 1
        End If
        'Αποθήκευση του ΝΕΟΥ ΕΣΟΔΟΥ για τον τρέχων προϋπολογισμό
        ado_pr_esod.Recordset.AddNew
        'Αποθήκευση id
        ado_pr_esod.Recordset.Fields(0).Value = id_πε
        'Αποθήκευση id_προϋπολογισμού
        ado_pr_esod.Recordset.Fields(1).Value = txt_id
        'Αποθήκευση id_τύπου_εσόδου
        If id_es >= 1 Then
            ado_pr_esod.Recordset.Fields(2).Value = id_es
        Else
            ado_pr_esod.Recordset.Fields(2).Value = ""
        End If
        'Αποθήκευση ποσού εσόδου
        If Trim(txt_poso_esod.Text) <> "" Then
            ado_pr_esod.Recordset.Fields(3).Value = Val(txt_poso_esod.Text)
        End If
        ado_pr_esod.Recordset.UpdateBatch adAffectCurrent
        ado_pr_esod.Recordset.Requery
        ado_pr_esod.Refresh
        ado_ap_esod.Recordset.Requery
        ado_ap_esod.Refresh
        dt_an_esod.Refresh
        '
        ado_ap_esod.Recordset.Filter = "[id_προϋπολογισμού] = '" & txt_id & "'"
        'Υπολογισμός Sum
        s = 0
        s2 = 0
        s3 = 0
        s4 = 0
        For i = 0 To ado_ap_esod.Recordset.RecordCount - 1
            If i = 0 And Not ado_ap_esod.Recordset.EOF Then
                ado_ap_esod.Recordset.MoveFirst
            End If
            s = s + ado_ap_esod.Recordset.Fields(3).Value
            s2 = s2 + ado_ap_esod.Recordset.Fields(5).Value
            s3 = s3 + ado_ap_esod.Recordset.Fields(6).Value
            s4 = s4 + ado_ap_esod.Recordset.Fields(7).Value
            ado_ap_esod.Recordset.MoveNext
        Next i
        Text5.Text = FormatCurrency(s, 2, vbTrue, , vbTrue)
        Text6.Text = FormatCurrency(s2, 2, vbTrue, , vbTrue)
        Text7.Text = FormatCurrency(s3, 2, vbTrue, , vbTrue)
        Text8.Text = FormatCurrency(s4, 2, vbTrue, , vbTrue)

        '
        If ado_ap_esod.Recordset.RecordCount >= 1 Then
            ado_ap_esod.Recordset.MoveLast
            dt_an_esod.Col = 1
            If ado_ap_esod.Recordset.RecordCount > 0 Then
                ado_ap_esod.Caption = "Έσοδα " & ado_ap_esod.Recordset.AbsolutePosition & " από " & ado_ap_esod.Recordset.RecordCount
            End If
        End If
        '''
        dt_an_esod.Columns(0).Visible = False
        dt_an_esod.Columns(1).Visible = False
        dt_an_esod.Columns(2).Caption = "Περιγραφή Εσόδων"
        dt_an_esod.Columns(2).Width = 2000
        dt_an_esod.Columns(3).Alignment = dbgRight
        dt_an_esod.Columns(3).Caption = "     Ποσό ΠΥ"
        dt_an_esod.Columns(3).Width = 1000
        dt_an_esod.Columns(3).NumberFormat = "currency"
        dt_an_esod.Columns(4).Visible = False
        dt_an_esod.Columns(5).Caption = "Ποσό Κινήσεων"
        dt_an_esod.Columns(5).Width = 1335
        dt_an_esod.Columns(5).NumberFormat = "currency"
        dt_an_esod.Columns(5).Alignment = dbgRight
        dt_an_esod.Columns(6).Caption = "Ποσό Δεσμεύσεων"
        dt_an_esod.Columns(6).Width = 1575
        dt_an_esod.Columns(6).NumberFormat = "currency"
        dt_an_esod.Columns(6).Alignment = dbgRight
        dt_an_esod.Columns(7).Caption = "  Υπόλοιπο"
        dt_an_esod.Columns(7).Width = 975
        dt_an_esod.Columns(7).NumberFormat = "currency"
        dt_an_esod.Columns(7).Alignment = dbgRight
        txt_poso_esod.Text = FormatCurrency(txt_poso_esod.Text, 2, vbTrue, , vbTrue)    'ΤΕΛΟΣ --- ΠΡΟΣΘΗΚΗ ΝΕΟΥ ΕΣΟΔΟΥ ΣΕ ΠΡΟΥΠΟΛΟΓΙΣΜΟ
    Else
        'ΑΡΧΗ - ΕΝΗΜΕΡΩΣΗ ΠΑΛΙΟΥ ΕΣΟΔΟΥ ΣΕ ΠΡΟΥΠΟΛΟΓΙΣΜΟ
        new_addition = 0
        bt_up_esoda.Enabled = False
        bt_ins_esod.Enabled = True
        bt_del_esoda.Enabled = True
        bt_can_esoda.Enabled = False
        
        ado_pr_esod.Recordset.Find "[id] = '" & ado_ap_esod.Recordset.Fields(0).Value & "'"
        
        'Ενημέρωση id_τύπου_εσόδου
        If id_es >= 1 Then
            ado_pr_esod.Recordset.Fields(2).Value = id_es
        Else
            ado_pr_esod.Recordset.Fields(2).Value = ""
        End If
        'Ενημέρωση ποσού εσόδου
        If Trim(txt_poso_esod.Text) <> "" Then
            ado_pr_esod.Recordset.Fields(3).Value = Val(txt_poso_esod.Text)
        End If
        '
        col_affected = ado_ap_esod.Recordset.AbsolutePosition
        ado_pr_esod.Recordset.UpdateBatch adAffectCurrent
        
        ado_pr_esod.Recordset.Requery
        ado_pr_esod.Refresh
        ado_ap_esod.Recordset.Requery
        ado_ap_esod.Refresh
        dt_an_esod.Refresh
        '
        ado_ap_esod.Recordset.Filter = "[id_προϋπολογισμού] = '" & txt_id & "'"
        'Υπολογισμός Sum
        s = 0
        s2 = 0
        s3 = 0
        s4 = 0
        For i = 0 To ado_ap_esod.Recordset.RecordCount - 1
            If i = 0 And Not ado_ap_esod.Recordset.EOF Then
                ado_ap_esod.Recordset.MoveFirst
            End If
            s = s + ado_ap_esod.Recordset.Fields(3).Value
            s2 = s2 + ado_ap_esod.Recordset.Fields(5).Value
            s3 = s3 + ado_ap_esod.Recordset.Fields(6).Value
            s4 = s4 + ado_ap_esod.Recordset.Fields(7).Value
            ado_ap_esod.Recordset.MoveNext
        Next i
        Text5.Text = FormatCurrency(s, 2, vbTrue, , vbTrue)
        Text6.Text = FormatCurrency(s2, 2, vbTrue, , vbTrue)
        Text7.Text = FormatCurrency(s3, 2, vbTrue, , vbTrue)
        Text8.Text = FormatCurrency(s4, 2, vbTrue, , vbTrue)
        '
        If ado_ap_esod.Recordset.RecordCount >= 1 Then
            ado_ap_esod.Recordset.MoveFirst
            If col_affected > 1 Then
                ado_ap_esod.Recordset.Move col_affected - 1
            End If
            dt_an_esod.Col = 1
            If ado_ap_esod.Recordset.RecordCount > 0 Then
                ado_ap_esod.Caption = "Έσοδα " & ado_ap_esod.Recordset.AbsolutePosition & " από " & ado_ap_esod.Recordset.RecordCount
            End If
        End If
        '''
        dt_an_esod.Columns(0).Visible = False
        dt_an_esod.Columns(1).Visible = False
        dt_an_esod.Columns(2).Caption = "Περιγραφή Εσόδων"
        dt_an_esod.Columns(2).Width = 2000
        dt_an_esod.Columns(3).Alignment = dbgRight
        dt_an_esod.Columns(3).Caption = "     Ποσό ΠΥ"
        dt_an_esod.Columns(3).Width = 1000
        dt_an_esod.Columns(3).NumberFormat = "currency"
        dt_an_esod.Columns(4).Visible = False
        dt_an_esod.Columns(5).Caption = "Ποσό Κινήσεων"
        dt_an_esod.Columns(5).Width = 1335
        dt_an_esod.Columns(5).NumberFormat = "currency"
        dt_an_esod.Columns(5).Alignment = dbgRight
        dt_an_esod.Columns(6).Caption = "Ποσό Δεσμεύσεων"
        dt_an_esod.Columns(6).Width = 1575
        dt_an_esod.Columns(6).NumberFormat = "currency"
        dt_an_esod.Columns(6).Alignment = dbgRight
        dt_an_esod.Columns(7).Caption = "  Υπόλοιπο"
        dt_an_esod.Columns(7).Width = 975
        dt_an_esod.Columns(7).NumberFormat = "currency"
        dt_an_esod.Columns(7).Alignment = dbgRight
        txt_poso_esod.Text = FormatCurrency(txt_poso_esod.Text, 2, vbTrue, , vbTrue)
        'ΤΕΛΟΣ --- ΕΝΗΜΕΡΩΣΗ ΠΑΛΙΟΥ ΕΣΟΔΟΥ ΣΕ ΠΡΟΥΠΟΛΟΓΙΣΜΟ
    End If
    
End Sub

Private Sub Command3_Click()

    Dim Sum_p, Sum_x, Sum_d As Double

    'Off Line Ενημέρωση
    'ΕΣΟΔΑ
    Dim oik_recs As New ADODB.Recordset
    oik_recs.Open "ΟικονομικέςΚινήσεις", Me.ado_ap_esod.ConnectionString, adOpenDynamic, adLockBatchOptimistic
    oik_recs.Filter = "[id_py] = '" & txt_id.Text & "'"
    If Not oik_recs.EOF Then
        oik_recs.MoveFirst
        ado_pr_esod.Recordset.Requery
        ado_pr_esod.Refresh
        dt_pr_esod.Refresh
        ado_pr_esod.Recordset.Filter = "[id_προϋπολογισμού] Like '" & txt_id.Text & "'"
        If Not ado_pr_esod.Recordset.EOF Then
            ado_pr_esod.Recordset.MoveFirst
            Do While Not ado_pr_esod.Recordset.EOF
                id_es = ado_pr_esod.Recordset.Fields("id_τύπου_εσόδου").Value
                oik_recs.Filter = "[id_ΚατηγορίαΠΥΕσόδων] Like '" & id_es & "'"
                Sum_p = 0
                Sum_d = 0
                If Not oik_recs.EOF Then
                    oik_recs.MoveFirst
                    Do While Not oik_recs.EOF
                        If ((oik_recs.Fields("ΤύποςΚίνησης").Value = 1) And (oik_recs.Fields("ΚατάστασηΚίνησης").Value = 1)) Then
                            Sum_p = Sum_p + CDbl(oik_recs.Fields("ΠοσόΠίστωσης").Value)
                        End If
                        If ((oik_recs.Fields("ΤύποςΚίνησης").Value = 1) And (oik_recs.Fields("ΚατάστασηΚίνησης").Value = 0)) Then
                            Sum_d = CDbl(Sum_d + oik_recs.Fields("ΠοσόΠίστωσης").Value)
                        End If
                        oik_recs.MoveNext
                    Loop
                    ado_pr_esod.Recordset.Fields(4).Value = Sum_p
                    ado_pr_esod.Recordset.Fields(5).Value = Sum_d
                    ado_pr_esod.Recordset.UpdateBatch adAffectCurrent
                End If
                ado_pr_esod.Recordset.MoveNext
            Loop
        End If
    End If
    oik_recs.Close
    ''''''ΤΕΛΟΣ ΕΣΟΔΑ''''''
    'ΕΞΟΔΑ
    oik_recs.Open "ΟικονομικέςΚινήσεις", Me.ado_ap_esod.ConnectionString, adOpenDynamic, adLockBatchOptimistic
    oik_recs.Filter = "[id_py] = '" & txt_id.Text & "'"
    If Not oik_recs.EOF Then
        oik_recs.MoveFirst
        ado_pr_eksod.Recordset.Requery
        ado_pr_eksod.Refresh
        dt_pr_eksod.Refresh
        ado_pr_eksod.Recordset.Filter = "[id_προϋπολογισμού] Like '" & txt_id.Text & "'"
        If Not ado_pr_eksod.Recordset.EOF Then
            ado_pr_eksod.Recordset.MoveFirst
            Do While Not ado_pr_eksod.Recordset.EOF
                id_eks = ado_pr_eksod.Recordset.Fields("id_τύπου_εξόδου").Value
                oik_recs.Filter = "[id_ΚατηγορίαΠΥΕξόδων] Like '" & id_eks & "'"
                Sum_x = 0
                Sum_d = 0
                If Not oik_recs.EOF Then
                    oik_recs.MoveFirst
                    Do While Not oik_recs.EOF
                        If ((oik_recs.Fields("ΤύποςΚίνησης").Value = 0) And (oik_recs.Fields("ΚατάστασηΚίνησης").Value = 1)) Then
                            Sum_x = Sum_x + CDbl(oik_recs.Fields("ΠοσόΧρέωσης").Value)
                        End If
                        If ((oik_recs.Fields("ΤύποςΚίνησης").Value = 0) And (oik_recs.Fields("ΚατάστασηΚίνησης").Value = 0)) Then
                            Sum_d = Sum_d + CDbl(oik_recs.Fields("ΠοσόΧρέωσης").Value)
                        End If
                        oik_recs.MoveNext
                    Loop
                    ado_pr_eksod.Recordset.Fields(4).Value = Sum_x
                    ado_pr_eksod.Recordset.Fields(5).Value = Sum_d
                    ado_pr_eksod.Recordset.UpdateBatch adAffectCurrent
                End If
                ado_pr_eksod.Recordset.MoveNext
            Loop
        End If
    End If
    oik_recs.Close
    'Unload Me
    'Load Me
    bt_refr_esoda_Click
    MsgBox "Ολοκλήρωση της Off Line Ενημέρωσης.", , "Off Line Ενημέρωση"
    
End Sub

Private Sub Command4_Click()
    
    Rep_analisi_proypologismou_esoda_eksoda.Show
    
End Sub

Private Sub dt_an_eksod_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If ado_ap_eksod.Recordset.AbsolutePosition >= 1 And ado_ap_eksod.Recordset.AbsolutePosition <= ado_ap_eksod.Recordset.RecordCount Then
        If Not ado_ap_eksod.Recordset.EOF Then
          If Trim(ado_ap_eksod.Recordset.Fields(2).Value) <> "" Then
                If Not ado_eksoda.Recordset.EOF Then
                    ado_eksoda.Recordset.MoveFirst
                    ado_eksoda.Recordset.Find "[id_τύπου_εξόδου] = '" & ado_ap_eksod.Recordset.Fields(4).Value & "'"
                    If Not ado_eksoda.Recordset.EOF Then
                        co_eksoda.Text = ado_eksoda.Recordset.Fields(1).Value
                    End If
                End If
            Else
                co_eksoda.Text = ""
            End If
            If Trim(ado_ap_eksod.Recordset.Fields(3).Value) <> "" Then
                txt_poso_eksod.Text = ado_ap_eksod.Recordset.Fields(3).Value
                txt_poso_eksod.Text = FormatCurrency(txt_poso_eksod.Text, 2, vbTrue, , vbTrue)
            Else
                txt_poso_eksod.Text = ""
                txt_poso_eksod.Text = FormatCurrency(txt_poso_eksod.Text, 2, vbTrue, , vbTrue)
                txt_poso_eksod.Enabled = False
            End If
        End If
    End If

End Sub

Private Sub dt_an_esod_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If ado_ap_esod.Recordset.AbsolutePosition >= 1 And ado_ap_esod.Recordset.AbsolutePosition <= ado_ap_esod.Recordset.RecordCount Then
        If Not ado_ap_esod.Recordset.EOF Then
          If Trim(ado_ap_esod.Recordset.Fields(2).Value) <> "" Then
                If Not ado_esoda.Recordset.EOF Then
                    ado_esoda.Recordset.MoveFirst
                    ado_esoda.Recordset.Find "[id_τύπου_εσόδου] = '" & ado_ap_esod.Recordset.Fields(4).Value & "'"
                    If Not ado_esoda.Recordset.EOF Then
                        co_esoda.Text = ado_esoda.Recordset.Fields(1).Value
                    End If
                End If
            Else
                co_esoda.Text = ""
            End If
            If Trim(ado_ap_esod.Recordset.Fields(3).Value) <> "" Then
                txt_poso_esod.Text = ado_ap_esod.Recordset.Fields(3).Value
                txt_poso_esod.Text = FormatCurrency(txt_poso_esod.Text, 2, vbTrue, , vbTrue)
            Else
                txt_poso_esod.Text = ""
                txt_poso_esod.Text = FormatCurrency(txt_poso_esod.Text, 2, vbTrue, , vbTrue)
            End If
        End If
    End If
    If txt_poso_esod.Text = "" Then
        txt_poso_esod.Enabled = False
    Else
        txt_poso_esod.Enabled = True
    End If

End Sub

Private Sub Form_Load()

    Dim s, s2, s3, s4 As Double

    Me.Width = 16000
    Me.Height = 8250
    new_addition = 0
    
    txt_id.Locked = True
    txt_perigrafi.Locked = True
    txt_im_en.Locked = True
    txt_im_liks.Locked = True
    txt_egkr_gs.Locked = True
    txt_im_egkr.Locked = True
    
    txt_id.Text = proupologismos_management.txt_id.Text
    txt_perigrafi.Text = proupologismos_management.txt_perigrafi.Text
    txt_im_en.Text = proupologismos_management.mb_im_en.Text
    txt_im_liks.Text = proupologismos_management.mb_im_liks.Text
    txt_egkr_gs.Text = proupologismos_management.txt_egkr_gs.Text
    txt_im_egkr.Text = proupologismos_management.mb_im_egkr.Text
    
    If Not ado_pr_esod.Recordset.EOF Then
        ado_ap_esod.Recordset.Filter = "[id_προϋπολογισμού] = '" & txt_id & "'"
        'Υπολογισμός Sum
        s = 0
        s2 = 0
        s3 = 0
        s4 = 0
        For i = 0 To ado_ap_esod.Recordset.RecordCount - 1
            If i = 0 And Not ado_ap_esod.Recordset.EOF Then
                ado_ap_esod.Recordset.MoveFirst
            End If
            s = s + ado_ap_esod.Recordset.Fields(3).Value
            s2 = s2 + ado_ap_esod.Recordset.Fields(5).Value
            s3 = s3 + ado_ap_esod.Recordset.Fields(6).Value
            s4 = s4 + ado_ap_esod.Recordset.Fields(7).Value
            ado_ap_esod.Recordset.MoveNext
        Next i
        Text5.Text = FormatCurrency(s, 2, vbTrue, , vbTrue)
        Text6.Text = FormatCurrency(s2, 2, vbTrue, , vbTrue)
        Text7.Text = FormatCurrency(s3, 2, vbTrue, , vbTrue)
        Text8.Text = FormatCurrency(s4, 2, vbTrue, , vbTrue)
        '
        If ado_ap_esod.Recordset.RecordCount >= 1 Then
            ado_ap_esod.Recordset.MoveFirst
        Else
            co_esoda.Text = ""
            txt_poso_esod = ""
            ado_ap_esod.Caption = "Έσοδα 0 από 0"
            bt_up_esoda.Enabled = False
            bt_del_esoda.Enabled = False
        End If
    End If
    
    dt_an_esod.Columns(0).Visible = False
    dt_an_esod.Columns(1).Visible = False
    dt_an_esod.Columns(2).Caption = "Περιγραφή Εσόδων"
    dt_an_esod.Columns(2).Width = 2000
    dt_an_esod.Columns(3).Alignment = dbgRight
    dt_an_esod.Columns(3).Caption = "     Ποσό ΠΥ"
    dt_an_esod.Columns(3).Width = 1000
    dt_an_esod.Columns(3).NumberFormat = "currency"
    dt_an_esod.Columns(4).Visible = False
    dt_an_esod.Columns(5).Caption = "Ποσό Κινήσεων"
    dt_an_esod.Columns(5).Width = 1335
    dt_an_esod.Columns(5).NumberFormat = "currency"
    dt_an_esod.Columns(5).Alignment = dbgRight
    dt_an_esod.Columns(6).Caption = "Ποσό Δεσμεύσεων"
    dt_an_esod.Columns(6).Width = 1575
    dt_an_esod.Columns(6).NumberFormat = "currency"
    dt_an_esod.Columns(6).Alignment = dbgRight
    dt_an_esod.Columns(7).Caption = "  Υπόλοιπο"
    dt_an_esod.Columns(7).Width = 975
    dt_an_esod.Columns(7).NumberFormat = "currency"
    dt_an_esod.Columns(7).Alignment = dbgRight
    If txt_poso_esod.Text <> "" Then
        txt_poso_esod.Text = FormatCurrency(txt_poso_esod.Text, 2, vbTrue, , vbTrue)
    Else
        txt_poso_esod.Text = ""
    End If

    '''ΕΝΗΜΕΡΩΣΗ ΕΞΟΔΩΝ ΚΑΤΑ ΤΟ LOADING ΤΗΣ ΦΟΡΜΑΣ
    new_addition_eks = 0
    
    If Not ado_pr_eksod.Recordset.EOF Then
        ado_ap_eksod.Recordset.Filter = "[id_προϋπολογισμού] = '" & txt_id & "'"
        'Υπολογισμός Sum
        s = 0
        s2 = 0
        s3 = 0
        s4 = 0
        For i = 0 To ado_ap_eksod.Recordset.RecordCount - 1
            If i = 0 And Not ado_ap_eksod.Recordset.EOF Then
                ado_ap_eksod.Recordset.MoveFirst
            End If
            s = s + CDbl(ado_ap_eksod.Recordset.Fields(3).Value)
            s2 = s2 + CDbl(ado_ap_eksod.Recordset.Fields(5).Value)
            s3 = s3 + CDbl(ado_ap_eksod.Recordset.Fields(6).Value)
            s4 = s4 + CDbl(ado_ap_eksod.Recordset.Fields(7).Value)
            ado_ap_eksod.Recordset.MoveNext
        Next i
        Text9.Text = FormatCurrency(s, 2, vbTrue, , vbTrue)
        Text10.Text = FormatCurrency(s2, 2, vbTrue, , vbTrue)
        Text11.Text = FormatCurrency(s3, 2, vbTrue, , vbTrue)
        Text12.Text = FormatCurrency(s4, 2, vbTrue, , vbTrue)
        '
        If ado_ap_eksod.Recordset.RecordCount >= 1 Then
            ado_ap_eksod.Recordset.MoveFirst
        Else
            co_eksoda.Text = ""
            txt_poso_eksod = ""
            ado_ap_eksod.Caption = "Έσοδα 0 από 0"
            bt_up_eksoda.Enabled = False
            bt_del_eksoda.Enabled = False
        End If
    End If
    
    dt_an_eksod.Columns(0).Visible = False
    dt_an_eksod.Columns(1).Visible = False
    dt_an_eksod.Columns(2).Caption = "Περιγραφή Εξόδων"
    dt_an_eksod.Columns(2).Width = 2000
    dt_an_eksod.Columns(3).Alignment = dbgRight
    dt_an_eksod.Columns(3).Caption = "     Ποσό ΠΥ"
    dt_an_eksod.Columns(3).Width = 1000
    dt_an_eksod.Columns(3).NumberFormat = "currency"
    dt_an_eksod.Columns(4).Visible = False
    dt_an_eksod.Columns(5).Caption = "Ποσό Κινήσεων"
    dt_an_eksod.Columns(5).Width = 1335
    dt_an_eksod.Columns(5).NumberFormat = "currency"
    dt_an_eksod.Columns(5).Alignment = dbgRight
    dt_an_eksod.Columns(6).Caption = "Ποσό Δεσμεύσεων"
    dt_an_eksod.Columns(6).Width = 1575
    dt_an_eksod.Columns(6).NumberFormat = "currency"
    dt_an_eksod.Columns(6).Alignment = dbgRight
    dt_an_eksod.Columns(7).Caption = "  Υπόλοιπο"
    dt_an_eksod.Columns(7).Width = 975
    dt_an_eksod.Columns(7).NumberFormat = "currency"
    dt_an_eksod.Columns(7).Alignment = dbgRight
    If txt_poso_eksod.Text <> "" Then
        txt_poso_eksod.Text = FormatCurrency(txt_poso_eksod.Text, 2, vbTrue, , vbTrue)
    Else
        txt_poso_eksod.Text = ""
    End If
    
    If txt_poso_esod.Text = "" Then
        txt_poso_esod.Enabled = False
    Else
        txt_poso_esod.Enabled = True
    End If
    If txt_poso_eksod.Text = "" Then
        txt_poso_eksod.Enabled = False
    Else
        txt_poso_eksod.Enabled = True
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

Private Sub txt_poso_eksod_GotFocus()
    
    bt_can_eksoda.Enabled = True
    
End Sub

Private Sub txt_poso_esod_GotFocus()
    
    bt_can_esoda.Enabled = True

End Sub

Private Sub txt_poso_esod_LostFocus()
    
    'txt_poso_esod.Text = FormatCurrency(Val(txt_poso_esod.Text), 2, vbTrue, , vbTrue)
    
End Sub
