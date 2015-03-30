VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form typoi_parastatikwn_esodwn_eksodwn_management 
   BackColor       =   &H00C0C0FF&
   Caption         =   "ƒÈ·˜ÂﬂÒÈÛÁ ‘˝˘Ì –·Ò·ÛÙ·ÙÈÍ˛Ì ≈Û¸‰˘Ì - ≈Ó¸‰˘Ì"
   ClientHeight    =   10560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10560
   ScaleWidth      =   11535
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000014&
      Caption         =   "¡ÔËﬁÍÂıÛÁ"
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
      Left            =   1320
      MaskColor       =   &H80000014&
      Picture         =   "tipoi_parastatikwn_eksodwn_management.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H80000014&
      Caption         =   " ÎÂﬂÛÈÏÔ"
      DisabledPicture =   "tipoi_parastatikwn_eksodwn_management.frx":4BFF
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
      Picture         =   "tipoi_parastatikwn_eksodwn_management.frx":4D47
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H80000014&
      Caption         =   "¡Ì·ÊﬁÙÁÛÁ"
      DisabledPicture =   "tipoi_parastatikwn_eksodwn_management.frx":A7BF
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
      Left            =   6240
      MaskColor       =   &H80000014&
      Picture         =   "tipoi_parastatikwn_eksodwn_management.frx":A907
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000014&
      Caption         =   "¡Í˝Ò˘ÛÁ"
      DisabledPicture =   "tipoi_parastatikwn_eksodwn_management.frx":FCBB
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
      Left            =   7560
      MaskColor       =   &H80000014&
      Picture         =   "tipoi_parastatikwn_eksodwn_management.frx":FE03
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000014&
      Caption         =   " ·Ë·ÒÈÛÏ¸Ú"
      DisabledPicture =   "tipoi_parastatikwn_eksodwn_management.frx":14D20
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
      Left            =   4920
      MaskColor       =   &H80000014&
      Picture         =   "tipoi_parastatikwn_eksodwn_management.frx":14E68
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2520
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "tipoi_parastatikwn_eksodwn_management.frx":199AC
      Height          =   495
      Left            =   8640
      TabIndex        =   7
      Top             =   6000
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   393216
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000014&
      Caption         =   "≈ÌÁÏ›Ò˘ÛÁ"
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
      Left            =   2520
      MaskColor       =   &H80000014&
      Picture         =   "tipoi_parastatikwn_eksodwn_management.frx":199C1
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "‘˝ÔÚ –·Ò·ÛÙ·ÙÈÍÔ˝"
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
      Height          =   2295
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   10695
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1935
         Begin VB.OptionButton opt 
            Caption         =   "≈Ó¸‰˘Ì"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   1
            Left            =   480
            TabIndex        =   29
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton opt 
            Caption         =   "≈Û¸‰˘Ì"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   28
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lb_f 
            Caption         =   "–·Ò·ÛÙ·ÙÈÍ¸"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   161
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   30
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.TextBox txt_f 
         BackColor       =   &H00FFFFFF&
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
         Index           =   5
         Left            =   8160
         TabIndex        =   6
         Top             =   1750
         Width           =   1455
      End
      Begin VB.TextBox txt_f 
         BackColor       =   &H00FFFFFF&
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
         Index           =   4
         Left            =   8160
         TabIndex        =   5
         Top             =   1150
         Width           =   1455
      End
      Begin VB.TextBox txt_f 
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
         Index           =   2
         Left            =   3840
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txt_f 
         BackColor       =   &H00FFFFFF&
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
         Index           =   3
         Left            =   8160
         TabIndex        =   4
         Top             =   520
         Width           =   1455
      End
      Begin VB.TextBox txt_f 
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
         Index           =   1
         Left            =   3840
         TabIndex        =   1
         Top             =   910
         Width           =   3135
      End
      Begin VB.TextBox txt_f 
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
         Index           =   0
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox Check1 
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
         Left            =   3840
         TabIndex        =   3
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lb_f 
         Alignment       =   1  'Right Justify
         Caption         =   "◊ÂÈÒ¸„Ò·ˆÁ ∏Í‰ÔÛÁ:"
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
         Index           =   1
         Left            =   2400
         TabIndex        =   20
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lb_f 
         Alignment       =   1  'Right Justify
         Caption         =   "‘Ò›˜˘Ì ¡ÒÈËÏ¸Ú:"
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
         Index           =   7
         Left            =   7080
         TabIndex        =   19
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lb_f 
         Alignment       =   1  'Right Justify
         Caption         =   "‘ÂÎÂıÙ·ﬂÔÚ ¡ÒÈËÏ¸Ú:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   5
         Left            =   7080
         TabIndex        =   18
         Top             =   1095
         Width           =   855
      End
      Begin VB.Label lb_f 
         Alignment       =   1  'Right Justify
         Caption         =   "”ıÌÙÔÏÔ„Ò·ˆﬂ·:"
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
         Index           =   4
         Left            =   2400
         TabIndex        =   17
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lb_f 
         Alignment       =   1  'Right Justify
         Caption         =   "–Ò˛ÙÔÚ ¡ÒÈËÏ¸Ú:"
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
         Index           =   3
         Left            =   7200
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lb_f 
         Alignment       =   1  'Right Justify
         Caption         =   "–ÂÒÈ„Ò·ˆﬁ:"
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
         Index           =   2
         Left            =   2520
         TabIndex        =   15
         Top             =   910
         Width           =   1215
      End
      Begin VB.Label lb_f 
         Alignment       =   1  'Right Justify
         Caption         =   " ˘‰ÈÍ¸Ú:"
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
         Index           =   0
         Left            =   2520
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton lb_arrow 
      Appearance      =   0  'Flat
      DisabledPicture =   "tipoi_parastatikwn_eksodwn_management.frx":214BB
      Enabled         =   0   'False
      Height          =   375
      Left            =   10440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tipoi_parastatikwn_eksodwn_management.frx":2157E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9480
      Width           =   490
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000014&
      Caption         =   "‘·ÓÈÌ¸ÏÁÛÁ"
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
      Left            =   240
      MaskColor       =   &H80000014&
      Picture         =   "tipoi_parastatikwn_eksodwn_management.frx":21641
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000014&
      Caption         =   "–ÒÔÛËﬁÍÁ"
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
      Picture         =   "tipoi_parastatikwn_eksodwn_management.frx":26348
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000014&
      Caption         =   "ƒÈ·„Ò·ˆﬁ"
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
      Left            =   3720
      MaskColor       =   &H80000014&
      Picture         =   "tipoi_parastatikwn_eksodwn_management.frx":2B0D7
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   9240
      Width           =   10695
      _ExtentX        =   18865
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
      MaxRecords      =   2
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=poseidon.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=poseidon.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "‘˝ÔÈ–·Ò·ÛÙ·ÙÈÍ˛Ì≈Û¸‰˘Ì≈Ó¸‰˘Ì"
      Caption         =   "Adodc1"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      CausesValidation=   0   'False
      Height          =   5655
      Left            =   240
      TabIndex        =   11
      Top             =   3600
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   9975
      _Version        =   393216
      BackColor       =   16777215
      WordWrap        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ºÎÔÈ ÔÈ ‘˝ÔÈ –·Ò·ÛÙ·ÙÈÍ˛Ì"
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
      Height          =   7215
      Left            =   120
      TabIndex        =   31
      Top             =   3240
      Width           =   10935
   End
End
Attribute VB_Name = "typoi_parastatikwn_esodwn_eksodwn_management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public defined_col, c_r, g_left, ADD_DONE As Integer
Public rs As ADODB.Recordset

Private Sub Adodc1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

    MsgBox "An error occured: Description=" & Destription & "Error Number=" & ErrorNumber
    ErrorNumber = 0

End Sub

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    'If adReason <> adRsnRequery And Me.Adodc1.Recordset.AbsolutePosition >= 1 Then
    If (adReason = adRsnMove Or adReason = adRsnMoveLast Or adReason = adRsnMoveFirst Or adReason = adRsnMoveNext Or adReason = adRsnMovePrevious) And adReason <> adRsnAddNew And Me.Adodc1.Recordset.AbsolutePosition >= 1 And ADD_DONE <> 1 Then
        Me.MSHFlexGrid1.Row = Me.Adodc1.Recordset.AbsolutePosition
        Me.MSHFlexGrid1.Col = Me.DataGrid1.Col + 1
        c_c = Me.MSHFlexGrid1.Col
        Me.MSHFlexGrid1.Col = 0
        lb_arrow.Left = Me.MSHFlexGrid1.CellLeft + Me.MSHFlexGrid1.Left
        lb_arrow.Top = Me.MSHFlexGrid1.CellTop + Me.MSHFlexGrid1.Top
        lb_arrow.Width = Me.MSHFlexGrid1.CellWidth
        'lb_arrow.Height = Me.MSHFlexGrid1.CellHeight
        lb_arrow.Height = 500
        Me.MSHFlexGrid1.Col = c_c
        '≈Õ«Ã≈—Ÿ”« ‘ŸÕ TEXT BOXES
        k = 0
        j = 0
        If Me.Adodc1.Recordset.Fields(j).Value <> "" Then
            txt_f(k).Text = Me.Adodc1.Recordset.Fields(j).Value
        Else
            txt_f(k).Text = ""
        End If
        k = k + 1
        j = j + 1
        If Me.Adodc1.Recordset.Fields(j).Value = True Then 'EINAI ≈”œƒœ
            Me.opt(0).Value = True
            clr = &H8000&
        Else '≈Õ¡… ≈Œœƒœ
            Me.opt(1).Value = True
            clr = &H800000
        End If
        txt_f(0).ForeColor = clr
        For j = j + 1 To 3
            If Me.Adodc1.Recordset.Fields(j).Value <> "" Then
                txt_f(k).Text = Me.Adodc1.Recordset.Fields(j).Value
            Else
                txt_f(k).Text = ""
            End If
            txt_f(k).ForeColor = clr
            k = k + 1
        Next j
        If Me.Adodc1.Recordset.Fields(j).Value = True Then 'EINAI ◊≈…—œ√—¡÷« ≈ ƒœ”«
            Me.Check1.Value = 1
        Else 'EINAI «À≈ ‘—œÕ… « ≈ ƒœ”«
            Me.Check1.Value = 0
        End If
        For j = j + 1 To 7
            If Me.Adodc1.Recordset.Fields(j).Value <> "" Then
                txt_f(k).Text = Me.Adodc1.Recordset.Fields(j).Value
            Else
                txt_f(k).Text = ""
            End If
            txt_f(k).ForeColor = clr
            k = k + 1
        Next j
        '
        'Me.Adodc1.Caption = "–·Ò·ÛÙ·ÙÈÍ¸ " & Me.MSHFlexGrid1.Row & " ·¸ " & Me.MSHFlexGrid1.Rows - 1
        Me.Adodc1.Caption = "–·Ò·ÛÙ·ÙÈÍ¸ " & Me.Adodc1.Recordset.AbsolutePosition & " ·¸ " & Me.MSHFlexGrid1.Rows - 1
    End If
    
End Sub

Private Sub Command_update_Click()

'    Dim id As Integer
'
'    If Me.Option1 = True Then
'        Me.Adodc1.Recordset.Fields(1).Value = 1
'        c_c = Me.MSHFlexGrid1.Col
'        For j = 1 To 3
'            Me.MSHFlexGrid1.Col = j
'            Me.MSHFlexGrid1.CellForeColor = &H8000&
'        Next j
'        Me.MSHFlexGrid1.Col = c_c
'    End If
'    If Me.Option2 = True Then
'        Me.Adodc1.Recordset.Fields(1).Value = 0
'        Me.MSHFlexGrid1.CellForeColor = &H800000
'        c_c = Me.MSHFlexGrid1.Col
'        For j = 1 To 3
'            Me.MSHFlexGrid1.Col = j
'            Me.MSHFlexGrid1.CellForeColor = &H800000
'        Next j
'        Me.MSHFlexGrid1.Col = c_c
'    End If
'    For j = 1 To 3
'        If Me.MSHFlexGrid1.TextMatrix(0, j) = "–ÂÒÈ„Ò·ˆﬁ" Then
'            Me.Adodc1.Recordset.Fields(2).Value = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, j)
'        End If
'        If Me.MSHFlexGrid1.TextMatrix(0, j) = "”ıÌÙÔÏÔ„Ò·ˆﬂ·" Then
'            Me.Adodc1.Recordset.Fields(3).Value = Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, j)
'        End If
'    Next j
'    Me.Adodc1.Recordset.UpdateBatch adAffectCurrent
'    'Me.Adodc1.Recordset.Sort = "id"
'    'Me.Adodc1.Recordset.MoveLast
'    'Me.DataGrid1.Col = 1
'    'Me.DataGrid1.SetFocus
    
End Sub

Private Sub Command1_Click()
    
    Dim ms As String
    
    If Me.Adodc1.Recordset.RecordCount >= 1 Then
        ms = MsgBox("≈ﬂÛ·È Ûﬂ„ÔıÒÔÚ; (Õ¡… ﬁ œ◊…)", vbYesNo, "–·Ò‹ËıÒÔ ‰È·„Ò·ˆﬁÚ")
        If ms = 6 Then
            If Not Me.Adodc1.Recordset.EOF Then
                l_r = Me.MSHFlexGrid1.Row
                Me.MSHFlexGrid1.RowSel = l_r
                If l_r > 1 Then
                    Me.MSHFlexGrid1.RemoveItem l_r
                Else
                    Me.Command1.Enabled = False 'DELETE
                    Me.Command4.Enabled = False 'UPDATE
                    Me.Command5.Enabled = False 'STORAGE
                End If
                With Me.Adodc1.Recordset
                    .Delete
                    .MoveNext
                    If .EOF And .RecordCount <> 0 Then
                        .MoveLast
                    Else
                        c_r = 0
                    End If
                End With
            End If
        End If
    End If

End Sub

Private Sub Command2_Click()

    Dim id As Integer
    
    ADD_DONE = 1
    Me.Adodc1.Recordset.Sort = "[" & Trim(Me.Adodc1.Recordset.Fields(0).Name) & "]"
    If Me.Adodc1.Recordset.RecordCount >= 1 Then
        Me.Adodc1.Recordset.MoveLast
        id = Me.Adodc1.Recordset![id]
        Me.Adodc1.Recordset.MoveLast
    Else
        id = 0
    End If
    
    'Ã«ƒ≈Õ…”Ãœ” ‘ŸÕ TEXT BOXES
    k = 0
    txt_f(k).Text = id + 1
    txt_f(k).ForeColor = &H8000&
    Me.opt(0).Value = True
    Me.opt(0).ForeColor = &H8000&
    For k = 1 To 5
            txt_f(k).Text = ""
            txt_f(k).ForeColor = &H8000&
    Next k
    Me.Check1.Value = 0
    Me.Check1.ForeColor = &H8000&
    '
    Me.txt_f(1).SetFocus
    Me.Command5.Enabled = True '¡–œ»« ≈’”«
    Me.Command2.Enabled = False
    Me.Command4.Enabled = False
    Me.Command1.Enabled = False 'DELETE
    Me.Command6.Enabled = False 'CLEAR
    
End Sub

Private Sub Command3_Click()

    If Me.MSHFlexGrid1.Col >= 1 Then
        Me.Adodc1.Recordset.Sort = "[" & Trim(Me.Adodc1.Recordset.Fields(Me.MSHFlexGrid1.Col - 1).Name) & "]"
    End If
    
    
    'ENHMERVSH TOY MSHFLEXGRID
    If Me.Adodc1.Recordset.RecordCount >= 1 Then
        Me.MSHFlexGrid1.Rows = Me.Adodc1.Recordset.RecordCount + 1
        Me.MSHFlexGrid1.Col = 0
        Me.MSHFlexGrid1.Row = 1
        lb_arrow.Left = Me.MSHFlexGrid1.CellLeft + Me.MSHFlexGrid1.Left
        lb_arrow.Top = Me.MSHFlexGrid1.CellTop + Me.MSHFlexGrid1.Top
        lb_arrow.Width = Me.MSHFlexGrid1.CellWidth
        lb_arrow.Height = Me.MSHFlexGrid1.CellHeight
    
        Me.Adodc1.Recordset.MoveFirst
        Me.DataGrid1.Row = 0
        For i = 1 To Me.MSHFlexGrid1.Rows - 1
            Me.MSHFlexGrid1.RowHeight(i) = 500
            For j = 1 To 1
                Me.MSHFlexGrid1.TextMatrix(i, j) = Me.Adodc1.Recordset.Fields(j - 1).Value
                Me.txt_f(0).Text = Me.Adodc1.Recordset.Fields(j - 1).Value
                k = 1
            Next j
            j = 2
            If Me.Adodc1.Recordset.Fields(j - 1).Value = False Then
                Me.MSHFlexGrid1.Row = i
                Me.MSHFlexGrid1.Col = j
                Me.MSHFlexGrid1.TextMatrix(i, j) = "≈ŒœƒŸÕ"
                Me.opt(1).Value = True
                Me.MSHFlexGrid1.CellFontBold = True
            Else
                Me.MSHFlexGrid1.Row = i
                Me.MSHFlexGrid1.Col = j
                Me.MSHFlexGrid1.TextMatrix(i, j) = "≈”œƒŸÕ"
                Me.opt(0).Value = True
                Me.MSHFlexGrid1.CellFontBold = True
            End If
            For j = 3 To 8
                If Me.Adodc1.Recordset.Fields(j - 1).Value <> "" Then
                    If j = 5 And Me.Adodc1.Recordset.Fields(j - 1).Value = True Then
                        Me.MSHFlexGrid1.Row = i
                        Me.MSHFlexGrid1.Col = j
                        If Me.Adodc1.Recordset.Fields(1).Value = True Then
                            Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/pic_cheched.jpg")
                        Else
                            Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/blue_cheched.jpg")
                        End If
                        Me.MSHFlexGrid1.CellPictureAlignment = 4
                        Me.Check1.Value = 1
                    Else
                        If j = 5 And Me.Adodc1.Recordset.Fields(j - 1).Value = False Then
                            Me.MSHFlexGrid1.Row = i
                            Me.MSHFlexGrid1.Col = j
                            If Me.Adodc1.Recordset.Fields(1).Value = True Then
                                Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/pic_uncheched.jpg")
                            Else
                                Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/blue_uncheched.jpg")
                            End If
                            Me.MSHFlexGrid1.CellPictureAlignment = 4
                            Me.Check1.Value = 0
                        Else
                            If j = 5 And Me.Adodc1.Recordset.Fields(1).Value = False Then
                            Else
                                Me.MSHFlexGrid1.TextMatrix(i, j) = Me.Adodc1.Recordset.Fields(j - 1).Value
                                'Me.txt_f(k).Text = Me.Adodc1.Recordset.Fields(j - 1).Value
                                Me.MSHFlexGrid1.CellAlignment = 9
                                k = k + 1
                            End If
                        End If
                    End If
                Else
                    Me.MSHFlexGrid1.TextMatrix(i, j) = ""
                End If
            Next j
            Me.MSHFlexGrid1.FillStyle = flexFillRepeat
            Me.MSHFlexGrid1.Row = i
            Me.MSHFlexGrid1.Col = 1
            Me.MSHFlexGrid1.ColSel = 8
            Me.MSHFlexGrid1.RowSel = i
            If Me.Adodc1.Recordset.Fields(1).Value = True Then 'EINAI ≈”œƒœ
                Me.MSHFlexGrid1.CellForeColor = &H8000&
                Me.MSHFlexGrid1.CellAlignment = flexAlignLeftCenter
            Else '≈…Õ¡… ≈Œœƒœ
                Me.MSHFlexGrid1.CellForeColor = &H800000
                Me.MSHFlexGrid1.CellAlignment = flexAlignLeftCenter
            End If
            If Me.Adodc1.Recordset.Fields(1).Value = True Then 'EINAI ESODO
                Me.MSHFlexGrid1.Row = i
                Me.MSHFlexGrid1.Col = 1
                Me.MSHFlexGrid1.ColSel = 8
                Me.MSHFlexGrid1.RowSel = i
                Me.MSHFlexGrid1.CellBackColor = vbWhite
            End If
            'If Me.Adodc1.Recordset.Fields(1).Value = False Then 'EINAI ≈Œœƒœ
            '    Me.MSHFlexGrid1.Row = i
            '    Me.MSHFlexGrid1.Col = 5
            '    Me.MSHFlexGrid1.ColSel = 8
            '    Me.MSHFlexGrid1.RowSel = i
            '    Me.MSHFlexGrid1.CellBackColor = &H800000
            'End If
            Me.DataGrid1.Col = 0
            If Not Me.Adodc1.Recordset.EOF Then
                Me.Adodc1.Recordset.MoveNext
            End If
        Next i
    
        Me.Adodc1.Recordset.MoveFirst
        Me.DataGrid1.Row = 0
        '≈Õ«Ã≈—Ÿ”« ‘ŸÕ TEXT BOXES
        k = 0
        j = 0
        txt_f(k).Text = Me.Adodc1.Recordset.Fields(j).Value
        k = k + 1
        j = j + 1
        If Me.Adodc1.Recordset.Fields(j).Value = True Then 'EINAI ≈”œƒœ
            Me.opt(0).Value = True
            clr = &H8000&
        Else '≈Õ¡… ≈Œœƒœ
            Me.opt(1).Value = True
            clr = &H800000
        End If
        txt_f(0).ForeColor = clr
        For j = j + 1 To 3
            If Me.Adodc1.Recordset.Fields(j).Value <> "" Then
                txt_f(k).Text = Me.Adodc1.Recordset.Fields(j).Value
            Else
                txt_f(k).Text = ""
            End If
            txt_f(k).ForeColor = clr
            k = k + 1
        Next j
        If Me.Adodc1.Recordset.Fields(j).Value = True Then 'EINAI ◊≈…—œ√—¡÷« ≈ ƒœ”«
            Me.Check1.Value = 1
        Else 'EINAI «À≈ ‘—œÕ… « ≈ ƒœ”«
            Me.Check1.Value = 0
        End If
        For j = j + 1 To 7
            If Me.Adodc1.Recordset.Fields(j).Value <> "" Then
                txt_f(k).Text = Me.Adodc1.Recordset.Fields(j).Value
            Else
                txt_f(k).Text = ""
            End If
            txt_f(k).ForeColor = clr
            k = k + 1
        Next j
        '
    
        Me.MSHFlexGrid1.Row = Me.Adodc1.Recordset.AbsolutePosition
        Me.MSHFlexGrid1.Col = 0
        lb_arrow.Left = Me.MSHFlexGrid1.CellLeft + Me.MSHFlexGrid1.Left
        lb_arrow.Top = Me.MSHFlexGrid1.CellTop + Me.MSHFlexGrid1.Top
        lb_arrow.Width = Me.MSHFlexGrid1.CellWidth
        lb_arrow.Height = Me.MSHFlexGrid1.CellHeight
        Me.MSHFlexGrid1.Col = 1
    
        Me.Command5.Enabled = False '¡–œ»« ≈’”«
    Else 'KENO TO GRID, xwris eggrafes
        Me.Command5.Enabled = False 'APOTHIKEYSH
        Me.Command4.Enabled = False ' UPDATE
        Me.Command1.Enabled = False 'DELETE
    End If
 
End Sub

Private Sub Command4_Click()

        k = 0
        j = 0
        k = k + 1
        j = j + 1
        If Me.opt(0).Value = True Then 'EINAI ≈”œƒœ
            Me.Adodc1.Recordset.Fields(j).Value = True
        Else '≈Õ¡… ≈Œœƒœ
            Me.Adodc1.Recordset.Fields(j).Value = False
        End If
        For j = j + 1 To 3
            If txt_f(k).Text <> "" Then
                Me.Adodc1.Recordset.Fields(j).Value = txt_f(k).Text
            Else
                Me.Adodc1.Recordset.Fields(j).Value = ""
            End If
            k = k + 1
        Next j
        If Me.Check1.Value = 1 Then 'EINAI ◊≈…—œ√—¡÷« ≈ ƒœ”«
            Me.Adodc1.Recordset.Fields(j).Value = True
        Else 'EINAI «À≈ ‘—œÕ… « ≈ ƒœ”«
            Me.Adodc1.Recordset.Fields(j).Value = False
        End If
        For j = j + 1 To 7
            If txt_f(k).Text <> "" Then
                Me.Adodc1.Recordset.Fields(j).Value = Val(txt_f(k).Text)
            Else
                Me.Adodc1.Recordset.Fields(j).Value = Null
            End If
            k = k + 1
        Next j
        '
        Me.Adodc1.Recordset.UpdateBatch adAffectCurrent
        
        '≈Õ«Ã≈—Ÿ”« ‘œ’ MSHFLEXGRID
        i = Me.MSHFlexGrid1.Row
        j = 2
        If Me.Adodc1.Recordset.Fields(j - 1).Value = False Then
            Me.MSHFlexGrid1.Col = j
            Me.MSHFlexGrid1.TextMatrix(i, Me.MSHFlexGrid1.Col) = "≈ŒœƒŸÕ"
            Me.MSHFlexGrid1.CellFontBold = True
        Else
            Me.MSHFlexGrid1.Col = j
            Me.MSHFlexGrid1.Text = "≈”œƒŸÕ"
            Me.MSHFlexGrid1.CellFontBold = True
        End If
        For j = 3 To 8
            If Me.Adodc1.Recordset.Fields(j - 1).Value <> "" Then
                If j = 5 And Me.Adodc1.Recordset.Fields(j - 1).Value = True Then
                        Me.MSHFlexGrid1.Row = i
                        Me.MSHFlexGrid1.Col = j
                        If Me.Adodc1.Recordset.Fields(1).Value = True Then
                            Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/pic_cheched.jpg")
                        Else
                            Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/blue_cheched.jpg")
                        End If
                        Me.MSHFlexGrid1.CellPictureAlignment = 4
                        Me.Check1.Value = 1
                    Else
                        If j = 5 And Me.Adodc1.Recordset.Fields(j - 1).Value = False Then
                            Me.MSHFlexGrid1.Row = i
                            Me.MSHFlexGrid1.Col = j
                            If Me.Adodc1.Recordset.Fields(1).Value = True Then
                                Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/pic_uncheched.jpg")
                            Else
                                Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/blue_uncheched.jpg")
                            End If
                            Me.MSHFlexGrid1.CellPictureAlignment = 4
                            Me.Check1.Value = 0
                        Else
                            If j = 5 And Me.Adodc1.Recordset.Fields(1).Value = False Then
                            Else
                                Me.MSHFlexGrid1.TextMatrix(i, j) = Me.Adodc1.Recordset.Fields(j - 1).Value
                                'Me.txt_f(k).Text = Me.Adodc1.Recordset.Fields(j - 1).Value
                                Me.MSHFlexGrid1.CellAlignment = 9
                                k = k + 1
                            End If
                        End If
                    End If
                Else
                    Me.MSHFlexGrid1.TextMatrix(i, j) = ""
                End If
        Next j
        Me.MSHFlexGrid1.FillStyle = flexFillRepeat
        Me.MSHFlexGrid1.Col = 1
        Me.MSHFlexGrid1.ColSel = 8
        Me.MSHFlexGrid1.RowSel = Me.MSHFlexGrid1.Row
        If Me.Adodc1.Recordset.Fields(1).Value = True Then 'EINAI ≈”œƒœ
            Me.MSHFlexGrid1.CellForeColor = &H8000&
            Me.MSHFlexGrid1.CellAlignment = flexAlignLeftCenter
            Me.MSHFlexGrid1.CellBackColor = vbWhite
        Else '≈…Õ¡… ≈Œœƒœ
            Me.MSHFlexGrid1.CellForeColor = &H800000
            Me.MSHFlexGrid1.CellAlignment = flexAlignLeftCenter
        End If
        'If Me.Adodc1.Recordset.Fields(1).Value = False Then 'EINAI ≈Œœƒœ
        '    Me.MSHFlexGrid1.Col = 5
        '    Me.MSHFlexGrid1.ColSel = 8
        '    Me.MSHFlexGrid1.RowSel = Me.MSHFlexGrid1.Row
        '    Me.MSHFlexGrid1.CellBackColor = &H800000
        'End If
        Me.MSHFlexGrid1.Col = 1
        'Me.MSHFlexGrid1.CellAlignment = 0
        
End Sub

Private Sub Command5_Click()

    k = 0
    j = 0
    '
    Me.Adodc1.Recordset.AddNew
    Me.Adodc1.Recordset.Fields(0).Value = Val(Me.txt_f(0).Text)
    Me.MSHFlexGrid1.AddItem ("" & vbTab & Trim(Me.txt_f(0).Text))
    '
    k = k + 1
    j = j + 1
    If Me.opt(0).Value = True Then 'EINAI ≈”œƒœ
        Me.Adodc1.Recordset.Fields(j).Value = True
    Else '≈Õ¡… ≈Œœƒœ
        Me.Adodc1.Recordset.Fields(j).Value = False
    End If
    For j = j + 1 To 3
        If txt_f(k).Text <> "" Then
            Me.Adodc1.Recordset.Fields(j).Value = txt_f(k).Text
        Else
            Me.Adodc1.Recordset.Fields(j).Value = ""
        End If
        k = k + 1
    Next j
    If Me.Check1.Value = 1 Then 'EINAI ◊≈…—œ√—¡÷« ≈ ƒœ”«
        Me.Adodc1.Recordset.Fields(j).Value = True
    Else 'EINAI «À≈ ‘—œÕ… « ≈ ƒœ”«
        Me.Adodc1.Recordset.Fields(j).Value = False
    End If
    For j = j + 1 To 7
        If txt_f(k).Text <> "" Then
            Me.Adodc1.Recordset.Fields(j).Value = Val(txt_f(k).Text)
        Else
            If j = 7 And txt_f(3).Text <> "" Then '≈…Õ¡… œ ‘—≈◊ŸÕ ¡—…»Ãœ”  ¡… ƒ≈Õ ƒœ»« ≈ « ¡—◊… « ‘œ’ ‘…Ã« ¡–œ ‘œ ◊—«”‘«
                Me.Adodc1.Recordset.Fields(j).Value = Val(txt_f(3).Text)
            Else
                'Me.Adodc1.Recordset.Fields(j).Value = Null
            End If
        End If
        k = k + 1
    Next j
    '
    Me.Adodc1.Recordset.UpdateBatch adAffectCurrent
        
    '≈Õ«Ã≈—Ÿ”« ‘œ’ MSHFLEXGRID
    ADD_DONE = 0
    Me.Adodc1.Recordset.MoveLast
    Me.MSHFlexGrid1.Row = Me.Adodc1.Recordset.RecordCount
    i = Me.MSHFlexGrid1.Row
    j = 1
    Me.MSHFlexGrid1.Col = j
    Me.MSHFlexGrid1.RowHeight(i) = 500
    Me.MSHFlexGrid1.TextMatrix(i, j) = Me.Adodc1.Recordset.Fields(j - 1).Value
    Me.MSHFlexGrid1.CellAlignment = 9
    j = 2
    If Me.Adodc1.Recordset.Fields(j - 1).Value = False Then
        Me.MSHFlexGrid1.Col = j
        Me.MSHFlexGrid1.TextMatrix(i, Me.MSHFlexGrid1.Col) = "≈ŒœƒŸÕ"
        Me.MSHFlexGrid1.CellFontBold = True
    Else
        Me.MSHFlexGrid1.Col = j
        Me.MSHFlexGrid1.Text = "≈”œƒŸÕ"
        Me.MSHFlexGrid1.CellFontBold = True
    End If
    For j = 3 To 8
        If Me.Adodc1.Recordset.Fields(j - 1).Value <> "" Then
            If j = 5 And Me.Adodc1.Recordset.Fields(j - 1).Value = True Then
                        Me.MSHFlexGrid1.Row = i
                        Me.MSHFlexGrid1.Col = j
                        If Me.Adodc1.Recordset.Fields(1).Value = True Then
                            Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/pic_cheched.jpg")
                        Else
                            Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/blue_cheched.jpg")
                        End If
                        Me.MSHFlexGrid1.CellPictureAlignment = 4
                        Me.Check1.Value = 1
                    Else
                        If j = 5 And Me.Adodc1.Recordset.Fields(j - 1).Value = False Then
                            Me.MSHFlexGrid1.Row = i
                            Me.MSHFlexGrid1.Col = j
                            If Me.Adodc1.Recordset.Fields(1).Value = True Then
                                Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/pic_uncheched.jpg")
                            Else
                                Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/blue_uncheched.jpg")
                            End If
                            Me.MSHFlexGrid1.CellPictureAlignment = 4
                            Me.Check1.Value = 0
                        Else
                            If j = 5 And Me.Adodc1.Recordset.Fields(1).Value = False Then
                            Else
                                Me.MSHFlexGrid1.TextMatrix(i, j) = Me.Adodc1.Recordset.Fields(j - 1).Value
                                'Me.txt_f(k).Text = Me.Adodc1.Recordset.Fields(j - 1).Value
                                Me.MSHFlexGrid1.CellAlignment = 9
                                k = k + 1
                            End If
                        End If
                    End If
                Else
                    Me.MSHFlexGrid1.TextMatrix(i, j) = ""
                End If
    Next j
    Me.MSHFlexGrid1.FillStyle = flexFillRepeat
    Me.MSHFlexGrid1.Col = 1
    Me.MSHFlexGrid1.ColSel = 8
    Me.MSHFlexGrid1.RowSel = Me.MSHFlexGrid1.Row
    If Me.Adodc1.Recordset.Fields(1).Value = True Then 'EINAI ≈”œƒœ
        Me.MSHFlexGrid1.CellForeColor = &H8000&
        Me.MSHFlexGrid1.CellAlignment = flexAlignLeftCenter
        Me.MSHFlexGrid1.CellBackColor = vbWhite
    Else '≈…Õ¡… ≈Œœƒœ
        Me.MSHFlexGrid1.CellForeColor = &H800000
        Me.MSHFlexGrid1.CellAlignment = flexAlignLeftCenter
    End If
    'If Me.Adodc1.Recordset.Fields(1).Value = False Then 'EINAI ≈Œœƒœ
    '    Me.MSHFlexGrid1.Col = 5
    '    Me.MSHFlexGrid1.ColSel = 8
    '    Me.MSHFlexGrid1.RowSel = Me.MSHFlexGrid1.Row
    '    Me.MSHFlexGrid1.CellBackColor = &H800000
    'End If
    Me.MSHFlexGrid1.Col = 1
        
    Me.Adodc1.Recordset.UpdateBatch adAffectCurrent
    Me.Adodc1.Recordset.Sort = "[" & Trim(Me.Adodc1.Recordset.Fields(0).Name) & "]"
    Me.Adodc1.Recordset.MoveLast
    Me.Adodc1.Recordset.MoveFirst
    For i = 1 To Me.Adodc1.Recordset.RecordCount - 1
        Me.Adodc1.Recordset.MoveNext
    Next i
    
    Me.Command2.Enabled = True '–—œ”»« «
    Me.Command4.Enabled = True 'UPDATE
    Me.Command5.Enabled = False '¡–œ»« ≈’”«
    Me.Command1.Enabled = True ' DELETE
    Me.Command6.Enabled = True ' CLEAR
    
End Sub

Private Sub Command6_Click()
    
    'Ã«ƒ≈Õ…”Ãœ” ‘ŸÕ TEXT BOXES
    k = 0
    txt_f(k).Text = ""
    txt_f(k).Locked = False
    'txt_f(k).ForeColor = &H8000&
    Me.opt(0).Value = False
    Me.opt(1).Value = False
    'Me.opt(0).ForeColor = &H8000&
    For k = 1 To 5
            txt_f(k).Text = ""
            'txt_f(k).ForeColor = &H8000&
    Next k
    Me.Check1.Value = 0
    'Me.Check1.ForeColor = &H8000&
    '
    Me.txt_f(0).SetFocus
    Me.Command5.Enabled = False '¡–œ»« ≈’”«
    Me.Command2.Enabled = False
    Me.Command4.Enabled = False
    Me.Command1.Enabled = False 'DELETE
    Me.Command8.Enabled = True 'SEARCH

End Sub

Private Sub Command7_Click()

    If Me.Adodc1.Recordset.RecordCount >= 1 Then
        Me.Adodc1.Refresh
        Me.Adodc1.Recordset.Sort = "[id]"
        Me.Adodc1.Recordset.MoveFirst
    End If
    
    'ENHMERVSH TOY MSHFLEXGRID
If Me.Adodc1.Recordset.RecordCount >= 1 Then
    Me.MSHFlexGrid1.Rows = Me.Adodc1.Recordset.RecordCount + 1
    Me.MSHFlexGrid1.Col = 0
    Me.MSHFlexGrid1.Row = 1
    lb_arrow.Left = Me.MSHFlexGrid1.CellLeft + Me.MSHFlexGrid1.Left
    lb_arrow.Top = Me.MSHFlexGrid1.CellTop + Me.MSHFlexGrid1.Top
    lb_arrow.Width = Me.MSHFlexGrid1.CellWidth
    lb_arrow.Height = Me.MSHFlexGrid1.CellHeight
    
    Me.Adodc1.Recordset.MoveFirst
    Me.DataGrid1.Row = 0
    For i = 1 To Me.MSHFlexGrid1.Rows - 1
        Me.MSHFlexGrid1.RowHeight(i) = 500
        For j = 1 To 1
            Me.MSHFlexGrid1.TextMatrix(i, j) = Me.Adodc1.Recordset.Fields(j - 1).Value
            Me.txt_f(0).Text = Me.Adodc1.Recordset.Fields(j - 1).Value
            k = 1
        Next j
        j = 2
        If Me.Adodc1.Recordset.Fields(j - 1).Value = False Then
            Me.MSHFlexGrid1.Row = i
            Me.MSHFlexGrid1.Col = j
            Me.MSHFlexGrid1.TextMatrix(i, j) = "≈ŒœƒŸÕ"
            Me.opt(1).Value = True
            Me.MSHFlexGrid1.CellFontBold = True
        Else
            Me.MSHFlexGrid1.Row = i
            Me.MSHFlexGrid1.Col = j
            Me.MSHFlexGrid1.TextMatrix(i, j) = "≈”œƒŸÕ"
            Me.opt(0).Value = True
            Me.MSHFlexGrid1.CellFontBold = True
        End If
        For j = 3 To 8
            If Me.Adodc1.Recordset.Fields(j - 1).Value <> "" Then
                If j = 5 And Me.Adodc1.Recordset.Fields(j - 1).Value = True Then
                        Me.MSHFlexGrid1.Row = i
                        Me.MSHFlexGrid1.Col = j
                        If Me.Adodc1.Recordset.Fields(1).Value = True Then
                            Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/pic_cheched.jpg")
                        Else
                            Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/blue_cheched.jpg")
                        End If
                        Me.MSHFlexGrid1.CellPictureAlignment = 4
                        Me.Check1.Value = 1
                    Else
                        If j = 5 And Me.Adodc1.Recordset.Fields(j - 1).Value = False Then
                            Me.MSHFlexGrid1.Row = i
                            Me.MSHFlexGrid1.Col = j
                            If Me.Adodc1.Recordset.Fields(1).Value = True Then
                                Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/pic_uncheched.jpg")
                            Else
                                Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/blue_uncheched.jpg")
                            End If
                            Me.MSHFlexGrid1.CellPictureAlignment = 4
                            Me.Check1.Value = 0
                        Else
                            If j = 5 And Me.Adodc1.Recordset.Fields(1).Value = False Then
                            Else
                                Me.MSHFlexGrid1.TextMatrix(i, j) = Me.Adodc1.Recordset.Fields(j - 1).Value
                                'Me.txt_f(k).Text = Me.Adodc1.Recordset.Fields(j - 1).Value
                                Me.MSHFlexGrid1.CellAlignment = 9
                                k = k + 1
                            End If
                        End If
                    End If
                Else
                    Me.MSHFlexGrid1.TextMatrix(i, j) = ""
                End If
        Next j
        Me.MSHFlexGrid1.FillStyle = flexFillRepeat
        Me.MSHFlexGrid1.Row = i
        Me.MSHFlexGrid1.Col = 1
        Me.MSHFlexGrid1.ColSel = 8
        Me.MSHFlexGrid1.RowSel = i
        If Me.Adodc1.Recordset.Fields(1).Value = True Then 'EINAI ≈”œƒœ
            Me.MSHFlexGrid1.CellForeColor = &H8000&
            Me.MSHFlexGrid1.CellAlignment = flexAlignLeftCenter
        Else '≈…Õ¡… ≈Œœƒœ
            Me.MSHFlexGrid1.CellForeColor = &H800000
            Me.MSHFlexGrid1.CellAlignment = flexAlignLeftCenter
        End If
        'If Me.Adodc1.Recordset.Fields(1).Value = True Then 'EINAI ≈Sœƒœ
        '    Me.MSHFlexGrid1.Row = i
        '    Me.MSHFlexGrid1.Col = 1
        '    Me.MSHFlexGrid1.ColSel = 8
        '    Me.MSHFlexGrid1.RowSel = i
        '    Me.MSHFlexGrid1.CellBackColor = vbWhite
        'End If
        'If Me.Adodc1.Recordset.Fields(1).Value = False Then 'EINAI ≈Œœƒœ
        '    Me.MSHFlexGrid1.Row = i
        '    Me.MSHFlexGrid1.Col = 5
        '    Me.MSHFlexGrid1.ColSel = 8
        '    Me.MSHFlexGrid1.RowSel = i
        '    Me.MSHFlexGrid1.CellBackColor = &H800000
        'End If
        Me.DataGrid1.Col = 0
        If Not Me.Adodc1.Recordset.EOF Then
            Me.Adodc1.Recordset.MoveNext
        End If
    Next i
    
    Me.Adodc1.Recordset.MoveLast
    'Me.Adodc1.Recordset.MoveFirst
    'Me.DataGrid1.Row = 0
    '≈Õ«Ã≈—Ÿ”« ‘ŸÕ TEXT BOXES
    k = 0
    j = 0
    txt_f(k).Text = Me.Adodc1.Recordset.Fields(j).Value
    k = k + 1
    j = j + 1
    If Me.Adodc1.Recordset.Fields(j).Value = True Then 'EINAI ≈”œƒœ
        Me.opt(0).Value = True
        clr = &H8000&
    Else '≈Õ¡… ≈Œœƒœ
        Me.opt(1).Value = True
        clr = &H800000
    End If
    txt_f(0).ForeColor = clr
    For j = j + 1 To 3
        If Me.Adodc1.Recordset.Fields(j).Value <> "" Then
            txt_f(k).Text = Me.Adodc1.Recordset.Fields(j).Value
        Else
            txt_f(k).Text = ""
        End If
        txt_f(k).ForeColor = clr
        k = k + 1
    Next j
    If Me.Adodc1.Recordset.Fields(j).Value = True Then 'EINAI ◊≈…—œ√—¡÷« ≈ ƒœ”«
        Me.Check1.Value = 1
    Else 'EINAI «À≈ ‘—œÕ… « ≈ ƒœ”«
        Me.Check1.Value = 0
    End If
    For j = j + 1 To 7
        If Me.Adodc1.Recordset.Fields(j).Value <> "" Then
            txt_f(k).Text = Me.Adodc1.Recordset.Fields(j).Value
        Else
            txt_f(k).Text = ""
        End If
        txt_f(k).ForeColor = clr
        k = k + 1
    Next j
    '
    
    Me.MSHFlexGrid1.Row = Me.Adodc1.Recordset.AbsolutePosition
    Me.MSHFlexGrid1.Col = 0
    lb_arrow.Left = Me.MSHFlexGrid1.CellLeft + Me.MSHFlexGrid1.Left
    lb_arrow.Top = Me.MSHFlexGrid1.CellTop + Me.MSHFlexGrid1.Top
    lb_arrow.Width = Me.MSHFlexGrid1.CellWidth
    lb_arrow.Height = Me.MSHFlexGrid1.CellHeight
    Me.MSHFlexGrid1.Col = 1
    Me.Command4.Enabled = True ' UPDATE
    Me.Command1.Enabled = True 'DELETE
    Me.Command6.Enabled = True 'CLEAR
Else 'KENO TO GRID, xwris eggrafes
    Me.Command4.Enabled = False ' UPDATE
    Me.Command1.Enabled = False 'DELETE
End If
 
    Me.Command5.Enabled = False 'STORAGE
    Me.Command8.Enabled = False 'DELETE
    Me.Command2.Enabled = True 'INSERT
    
End Sub

Private Sub Command8_Click()

    s_string = ""
    ' —…‘«—…œ  Ÿƒ… œ’
    If Trim(Me.txt_f(0).Text) <> "" Then
        s_string = "[id] LIKE " & Trim(Me.txt_f(0).Text)
    End If
    ' —…‘«—…œ –≈—…√—¡÷«”
    If Trim(Me.txt_f(1).Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [œÌÔÏ·Ûﬂ·] LIKE '*" & Trim(Me.txt_f(1).Text) & "*'"
        Else
            s_string = "[œÌÔÏ·Ûﬂ·] LIKE '*" & Trim(Me.txt_f(1).Text) & "*'"
        End If
    End If
    ' —…‘«—…œ ”’Õ‘œÃœ√—¡÷…¡”
    If Trim(Me.txt_f(2).Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [”ıÌÙÔÏÔ„Ò·ˆﬂ·] LIKE '*" & Trim(Me.txt_f(2).Text) & "*'"
        Else
            s_string = "[”ıÌÙÔÏÔ„Ò·ˆﬂ·] LIKE '*" & Trim(Me.txt_f(2).Text) & "*'"
        End If
    End If
    ' —…‘«—…œ –—Ÿ‘œ’ ¡—…»Ãœ’
    If Trim(Me.txt_f(3).Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [–Ò˛ÙÔÚ¡ÒÈËÏ¸Ú] LIKE " & Trim(Me.txt_f(3).Text)
        Else
            s_string = "[–Ò˛ÙÔÚ¡ÒÈËÏ¸Ú] LIKE " & Trim(Me.txt_f(3).Text)
        End If
    End If
    ' —…‘«—…œ ‘≈À≈’‘¡…œ’ ¡—…»Ãœ’
    If Trim(Me.txt_f(4).Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [‘ÂÎÂıÙ·ﬂÔÚ¡ÒÈËÏ¸Ú] LIKE " & Trim(Me.txt_f(4).Text)
        Else
            s_string = "[‘ÂÎÂıÙ·ﬂÔÚ¡ÒÈËÏ¸Ú] LIKE " & Trim(Me.txt_f(4).Text)
        End If
    End If
    ' —…‘«—…œ ‘—≈◊œÕ‘œ” ¡—…»Ãœ’
    If Trim(Me.txt_f(5).Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [‘Ò›˜˘Ì¡ÒÈËÏ¸Ú] LIKE " & Trim(Me.txt_f(5).Text)
        Else
            s_string = "[‘Ò›˜˘Ì¡ÒÈËÏ¸Ú] LIKE " & Trim(Me.txt_f(5).Text)
        End If
    End If
    ' —…‘«—…œ ‘’–œ’ –¡—¡”‘¡‘… œ’
    If Me.opt(0).Value = True And Me.opt(1).Value = False Then
        If s_string <> "" Then
            s_string = s_string & " AND [‘˝ÔÚ] LIKE TRUE"
        Else
            s_string = "[‘˝ÔÚ] LIKE TRUE"
        End If
    Else
        If Me.opt(0).Value = False And Me.opt(1).Value = True Then
            If s_string <> "" Then
                s_string = s_string & " AND [‘˝ÔÚ] LIKE FALSE"
            Else
                s_string = "[‘˝ÔÚ] LIKE FALSE"
            End If
        Else
            'If s_string <> "" Then
            '    s_string = s_string & " AND [‘˝ÔÚ] LIKE '*'"
            'Else
            '    s_string = "[‘˝ÔÚ] LIKE '*'"
            'End If
        End If
    End If
    ' —…‘«—…œ ‘—œ–œ’ ≈ ƒœ”«” (˜ÂÈÒ¸„Ò·ˆÔ ﬁ ÁÎÂÍÙÒÔÌÈÍ¸)
    If Me.Check1.Value = 1 Then
        If s_string <> "" Then
            s_string = s_string & " AND [◊ÂÈÒÔ„Ò_«ÎÂÍÙÒ] LIKE TRUE"
        Else
            s_string = "[◊ÂÈÒÔ„Ò_«ÎÂÍÙÒ] LIKE TRUE"
        End If
    End If
    '
    Me.Adodc1.Recordset.Filter = s_string
    'If s_sort <> "" Then
    '    Me.Adodc1.Recordset.Sort = Trim(s_sort)
    'End If
    '
    
    Me.Command7.Enabled = True 'CANCEL
    '
    'for_search = 0
    Me.txt_f(0).Locked = True
    
    'ENHMERVSH TOY MSHFLEXGRID
If Me.Adodc1.Recordset.RecordCount >= 1 Then
    Me.MSHFlexGrid1.Rows = Me.Adodc1.Recordset.RecordCount + 1
    Me.MSHFlexGrid1.Col = 0
    Me.MSHFlexGrid1.Row = 1
    lb_arrow.Left = Me.MSHFlexGrid1.CellLeft + Me.MSHFlexGrid1.Left
    lb_arrow.Top = Me.MSHFlexGrid1.CellTop + Me.MSHFlexGrid1.Top
    lb_arrow.Width = Me.MSHFlexGrid1.CellWidth
    lb_arrow.Height = Me.MSHFlexGrid1.CellHeight
    
    Me.Adodc1.Recordset.MoveFirst
    Me.DataGrid1.Row = 0
    For i = 1 To Me.MSHFlexGrid1.Rows - 1
        Me.MSHFlexGrid1.RowHeight(i) = 500
        For j = 1 To 1
            Me.MSHFlexGrid1.TextMatrix(i, j) = Me.Adodc1.Recordset.Fields(j - 1).Value
            Me.txt_f(0).Text = Me.Adodc1.Recordset.Fields(j - 1).Value
            k = 1
        Next j
        j = 2
        If Me.Adodc1.Recordset.Fields(j - 1).Value = False Then
            Me.MSHFlexGrid1.Row = i
            Me.MSHFlexGrid1.Col = j
            Me.MSHFlexGrid1.TextMatrix(i, j) = "≈ŒœƒŸÕ"
            Me.opt(1).Value = True
            Me.MSHFlexGrid1.CellFontBold = True
        Else
            Me.MSHFlexGrid1.Row = i
            Me.MSHFlexGrid1.Col = j
            Me.MSHFlexGrid1.TextMatrix(i, j) = "≈”œƒŸÕ"
            Me.opt(0).Value = True
            Me.MSHFlexGrid1.CellFontBold = True
        End If
        For j = 3 To 8
            If Me.Adodc1.Recordset.Fields(j - 1).Value <> "" Then
                If j = 5 And Me.Adodc1.Recordset.Fields(j - 1).Value = True Then
                        Me.MSHFlexGrid1.Row = i
                        Me.MSHFlexGrid1.Col = j
                        If Me.Adodc1.Recordset.Fields(1).Value = True Then
                            Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/pic_cheched.jpg")
                        Else
                            Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/blue_cheched.jpg")
                        End If
                        Me.MSHFlexGrid1.CellPictureAlignment = 4
                        Me.Check1.Value = 1
                    Else
                        If j = 5 And Me.Adodc1.Recordset.Fields(j - 1).Value = False Then
                            Me.MSHFlexGrid1.Row = i
                            Me.MSHFlexGrid1.Col = j
                            If Me.Adodc1.Recordset.Fields(1).Value = True Then
                                Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/pic_uncheched.jpg")
                            Else
                                Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/blue_uncheched.jpg")
                            End If
                            Me.MSHFlexGrid1.CellPictureAlignment = 4
                            Me.Check1.Value = 0
                        Else
                            If j = 5 And Me.Adodc1.Recordset.Fields(1).Value = False Then
                            Else
                                Me.MSHFlexGrid1.TextMatrix(i, j) = Me.Adodc1.Recordset.Fields(j - 1).Value
                                'Me.txt_f(k).Text = Me.Adodc1.Recordset.Fields(j - 1).Value
                                Me.MSHFlexGrid1.CellAlignment = 9
                                k = k + 1
                            End If
                        End If
                    End If
                Else
                    Me.MSHFlexGrid1.TextMatrix(i, j) = ""
                End If
        Next j
        Me.MSHFlexGrid1.FillStyle = flexFillRepeat
        Me.MSHFlexGrid1.Row = i
        Me.MSHFlexGrid1.Col = 1
        Me.MSHFlexGrid1.ColSel = 8
        Me.MSHFlexGrid1.RowSel = i
        If Me.Adodc1.Recordset.Fields(1).Value = True Then 'EINAI ≈”œƒœ
            Me.MSHFlexGrid1.CellForeColor = &H8000&
            Me.MSHFlexGrid1.CellAlignment = flexAlignLeftCenter
        Else '≈…Õ¡… ≈Œœƒœ
            Me.MSHFlexGrid1.CellForeColor = &H800000
            Me.MSHFlexGrid1.CellAlignment = flexAlignLeftCenter
        End If
        If Me.Adodc1.Recordset.Fields(1).Value = True Then 'EINAI ESODO
            Me.MSHFlexGrid1.Row = i
            Me.MSHFlexGrid1.Col = 1
            Me.MSHFlexGrid1.ColSel = 8
            Me.MSHFlexGrid1.RowSel = i
            Me.MSHFlexGrid1.CellBackColor = vbWhite
        End If
        'If Me.Adodc1.Recordset.Fields(1).Value = False Then 'EINAI ≈Œœƒœ
        '    Me.MSHFlexGrid1.Row = i
        '    Me.MSHFlexGrid1.Col = 5
        '    Me.MSHFlexGrid1.ColSel = 8
        '    Me.MSHFlexGrid1.RowSel = i
        '    Me.MSHFlexGrid1.CellBackColor = &H800000
        'End If
        Me.DataGrid1.Col = 0
        If Not Me.Adodc1.Recordset.EOF Then
            Me.Adodc1.Recordset.MoveNext
        End If
    Next i
    
    Me.Adodc1.Recordset.MoveFirst
    Me.DataGrid1.Row = 0
    '≈Õ«Ã≈—Ÿ”« ‘ŸÕ TEXT BOXES
    k = 0
    j = 0
    txt_f(k).Text = Me.Adodc1.Recordset.Fields(j).Value
    k = k + 1
    j = j + 1
    If Me.Adodc1.Recordset.Fields(j).Value = True Then 'EINAI ≈”œƒœ
        Me.opt(0).Value = True
        clr = &H8000&
    Else '≈Õ¡… ≈Œœƒœ
        Me.opt(1).Value = True
        clr = &H800000
    End If
    txt_f(0).ForeColor = clr
    For j = j + 1 To 3
        If Me.Adodc1.Recordset.Fields(j).Value <> "" Then
            txt_f(k).Text = Me.Adodc1.Recordset.Fields(j).Value
        Else
            txt_f(k).Text = ""
        End If
        txt_f(k).ForeColor = clr
        k = k + 1
    Next j
    If Me.Adodc1.Recordset.Fields(j).Value = True Then 'EINAI ◊≈…—œ√—¡÷« ≈ ƒœ”«
        Me.Check1.Value = 1
    Else 'EINAI «À≈ ‘—œÕ… « ≈ ƒœ”«
        Me.Check1.Value = 0
    End If
    For j = j + 1 To 7
        If Me.Adodc1.Recordset.Fields(j).Value <> "" Then
            txt_f(k).Text = Me.Adodc1.Recordset.Fields(j).Value
        Else
            txt_f(k).Text = ""
        End If
        txt_f(k).ForeColor = clr
        k = k + 1
    Next j
    '
    
    Me.MSHFlexGrid1.Row = Me.Adodc1.Recordset.AbsolutePosition
    Me.MSHFlexGrid1.Col = 0
    lb_arrow.Left = Me.MSHFlexGrid1.CellLeft + Me.MSHFlexGrid1.Left
    lb_arrow.Top = Me.MSHFlexGrid1.CellTop + Me.MSHFlexGrid1.Top
    lb_arrow.Width = Me.MSHFlexGrid1.CellWidth
    lb_arrow.Height = Me.MSHFlexGrid1.CellHeight
    Me.MSHFlexGrid1.Col = 1
    
    Me.Command5.Enabled = False '¡–œ»« ≈’”«
Else 'KENO TO GRID, xwris eggrafes
    Me.Command5.Enabled = False 'APOTHIKEYSH
    Me.Command4.Enabled = False ' UPDATE
    Me.Command1.Enabled = False 'DELETE
End If
 

End Sub

Private Sub Command9_Click()
    Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    defined_col = ColIndex
End Sub

Private Sub Form_Load()

    Dim rws As Integer
        
    ADD_DONE = 0
    Me.Width = 11300
    Me.Height = 11000
    
    rws = 0
    If Me.Adodc1.Recordset.RecordCount >= 1 Then
        Me.Adodc1.Recordset.Sort = "id"
        rws = Me.Adodc1.Recordset.RecordCount
    End If
        
    Me.MSHFlexGrid1.Rows = rws + 1
    Me.MSHFlexGrid1.Cols = 9
    Me.MSHFlexGrid1.FixedRows = 1
    Me.MSHFlexGrid1.FixedCols = 1
    Me.MSHFlexGrid1.RowHeight(0) = 700
    Me.MSHFlexGrid1.ColWidth(0) = 500
    Me.MSHFlexGrid1.ColWidth(1) = 700
    Me.MSHFlexGrid1.TextMatrix(0, 1) = " ˘‰ÈÍ¸Ú"
    Me.MSHFlexGrid1.ColWidth(2) = 1200
    Me.MSHFlexGrid1.TextMatrix(0, 2) = "‘˝ÔÚ –·Ò·ÛÙ·ÙÈÍÔ˝"
    Me.MSHFlexGrid1.TextMatrix(0, 3) = "–ÂÒÈ„Ò·ˆﬁ"
    Me.MSHFlexGrid1.ColWidth(3) = 3000
    Me.MSHFlexGrid1.TextMatrix(0, 4) = "”ıÌÙÔÏÔ„Ò·ˆﬂ·"
    Me.MSHFlexGrid1.ColWidth(4) = 1400
    Me.MSHFlexGrid1.TextMatrix(0, 5) = "◊ÂÈÒ¸„Ò·ˆÁ ∏Í‰ÔÛÁ"
    Me.MSHFlexGrid1.ColWidth(5) = 1050
    Me.MSHFlexGrid1.TextMatrix(0, 6) = "–Ò˛ÙÔÚ ¡ÒÈËÏ¸Ú ∏Í‰ÔÛÁÚ"
    Me.MSHFlexGrid1.ColWidth(6) = 850
    Me.MSHFlexGrid1.TextMatrix(0, 7) = "‘ÂÎÂıÙ·ﬂÔÚ ¡ÒÈËÏ¸Ú ∏Í‰ÔÛÁÚ"
    Me.MSHFlexGrid1.ColWidth(7) = 1000
    Me.MSHFlexGrid1.TextMatrix(0, 8) = "‘Ò›˜˘Ì ¡ÒÈËÏ¸Ú ∏Í‰ÔÛÁÚ"
    Me.MSHFlexGrid1.ColWidth(8) = 900

    If Me.Adodc1.Recordset.RecordCount >= 1 Then
        Me.MSHFlexGrid1.Col = 0
        Me.MSHFlexGrid1.Row = 1
        lb_arrow.Left = Me.MSHFlexGrid1.CellLeft + Me.MSHFlexGrid1.Left
        lb_arrow.Top = Me.MSHFlexGrid1.CellTop + Me.MSHFlexGrid1.Top
        lb_arrow.Width = Me.MSHFlexGrid1.CellWidth
        lb_arrow.Height = Me.MSHFlexGrid1.CellHeight
    
        Me.Adodc1.Recordset.MoveFirst
        Me.DataGrid1.Row = 0
        For i = 1 To Me.MSHFlexGrid1.Rows - 1
            Me.MSHFlexGrid1.RowHeight(i) = 500
            For j = 1 To 1
                Me.MSHFlexGrid1.TextMatrix(i, j) = Me.Adodc1.Recordset.Fields(j - 1).Value
                Me.MSHFlexGrid1.CellAlignment = 9
                'Me.txt_f(0).Text = Me.Adodc1.Recordset.Fields(j - 1).Value
                k = 1
            Next j
            j = 2
            If Me.Adodc1.Recordset.Fields(j - 1).Value = False Then
                Me.MSHFlexGrid1.Row = i
                Me.MSHFlexGrid1.Col = j
                Me.MSHFlexGrid1.TextMatrix(i, j) = "≈ŒœƒŸÕ"
                Me.MSHFlexGrid1.CellAlignment = 9
                Me.opt(1).Value = True
                Me.MSHFlexGrid1.CellFontBold = True
            Else
                Me.MSHFlexGrid1.Row = i
                Me.MSHFlexGrid1.Col = j
                Me.MSHFlexGrid1.TextMatrix(i, j) = "≈”œƒŸÕ"
                Me.MSHFlexGrid1.CellAlignment = 9
                Me.opt(0).Value = True
                Me.MSHFlexGrid1.CellFontBold = True
            End If
            For j = 3 To 8
                If Me.Adodc1.Recordset.Fields(j - 1).Value <> "" Then
                    If j = 5 And Me.Adodc1.Recordset.Fields(j - 1).Value = True Then
                        Me.MSHFlexGrid1.Row = i
                        Me.MSHFlexGrid1.Col = j
                        If Me.Adodc1.Recordset.Fields(1).Value = True Then
                            Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/pic_cheched.jpg")
                        Else
                            Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/blue_cheched.jpg")
                        End If
                        Me.MSHFlexGrid1.CellPictureAlignment = 4
                        Me.Check1.Value = 1
                    Else
                        If j = 5 And Me.Adodc1.Recordset.Fields(j - 1).Value = False Then
                            Me.MSHFlexGrid1.Row = i
                            Me.MSHFlexGrid1.Col = j
                            If Me.Adodc1.Recordset.Fields(1).Value = True Then
                                Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/pic_uncheched.jpg")
                            Else
                                Set Me.MSHFlexGrid1.CellPicture = LoadPicture("Btmps/blue_uncheched.jpg")
                            End If
                            Me.MSHFlexGrid1.CellPictureAlignment = 4
                            Me.Check1.Value = 0
                        Else
                            If j = 5 And Me.Adodc1.Recordset.Fields(1).Value = False Then
                            Else
                                Me.MSHFlexGrid1.TextMatrix(i, j) = Me.Adodc1.Recordset.Fields(j - 1).Value
                                'Me.txt_f(k).Text = Me.Adodc1.Recordset.Fields(j - 1).Value
                                Me.MSHFlexGrid1.CellAlignment = 9
                                k = k + 1
                            End If
                        End If
                    End If
                End If
            Next j
            Me.MSHFlexGrid1.FillStyle = flexFillRepeat
            Me.MSHFlexGrid1.Row = i
            Me.MSHFlexGrid1.Col = 1
            Me.MSHFlexGrid1.ColSel = 8
            Me.MSHFlexGrid1.RowSel = i
            If Me.Adodc1.Recordset.Fields(1).Value = True Then 'EINAI ≈”œƒœ
                Me.MSHFlexGrid1.CellForeColor = &H8000&
                Me.MSHFlexGrid1.CellAlignment = flexAlignLeftCenter
            Else '≈…Õ¡… ≈Œœƒœ
                Me.MSHFlexGrid1.CellForeColor = &H800000
                Me.MSHFlexGrid1.CellAlignment = flexAlignLeftCenter
            End If
            'If Me.Adodc1.Recordset.Fields(1).Value = False Then 'EINAI ≈Œœƒœ
                'Me.MSHFlexGrid1.Row = i
                'Me.MSHFlexGrid1.Col = 5
                'Me.MSHFlexGrid1.ColSel = 8
                'Me.MSHFlexGrid1.RowSel = i
                'Me.MSHFlexGrid1.CellBackColor = &H800000
            'End If
            'Me.DataGrid1.Col = 0
            If Not Me.Adodc1.Recordset.EOF Then
                Me.Adodc1.Recordset.MoveNext
            End If
        Next i
    
        'Me.Adodc1.Recordset.MoveFirst
        Me.Adodc1.Recordset.MoveLast
        'Me.DataGrid1.Row = 0
        '≈Õ«Ã≈—Ÿ”« ‘ŸÕ TEXT BOXES
        k = 0
        j = 0
        txt_f(k).Text = Me.Adodc1.Recordset.Fields(j).Value
        k = k + 1
        j = j + 1
        If Me.Adodc1.Recordset.Fields(j).Value = True Then 'EINAI ≈”œƒœ
            Me.opt(0).Value = True
            clr = &H8000&
        Else '≈Õ¡… ≈Œœƒœ
            Me.opt(1).Value = True
            clr = &H800000
        End If
        txt_f(0).ForeColor = clr
        For j = j + 1 To 3
            If Me.Adodc1.Recordset.Fields(j).Value <> "" Then
                txt_f(k).Text = Me.Adodc1.Recordset.Fields(j).Value
            Else
                txt_f(k).Text = ""
            End If
            txt_f(k).ForeColor = clr
            k = k + 1
        Next j
        If Me.Adodc1.Recordset.Fields(j).Value = True Then 'EINAI ◊≈…—œ√—¡÷« ≈ ƒœ”«
            Me.Check1.Value = 1
        Else 'EINAI «À≈ ‘—œÕ… « ≈ ƒœ”«
            Me.Check1.Value = 0
        End If
        For j = j + 1 To 7
            If Me.Adodc1.Recordset.Fields(j).Value <> "" Then
                txt_f(k).Text = Me.Adodc1.Recordset.Fields(j).Value
            Else
                txt_f(k).Text = ""
            End If
            txt_f(k).ForeColor = clr
            k = k + 1
        Next j
        '
    
        Me.MSHFlexGrid1.Row = Me.Adodc1.Recordset.AbsolutePosition
        Me.MSHFlexGrid1.Col = 0
        lb_arrow.Left = Me.MSHFlexGrid1.CellLeft + Me.MSHFlexGrid1.Left
        lb_arrow.Top = Me.MSHFlexGrid1.CellTop + Me.MSHFlexGrid1.Top
        lb_arrow.Width = Me.MSHFlexGrid1.CellWidth
        lb_arrow.Height = Me.MSHFlexGrid1.CellHeight
        Me.MSHFlexGrid1.Col = 1
        Me.Command5.Enabled = False '¡–œ»« ≈’”«
    Else 'KENO TO GRID, xwris eggrafes
        Me.Command5.Enabled = False 'APOTHIKEYSH
        Me.Command4.Enabled = False ' UPDATE
        Me.Command1.Enabled = False 'DELETE
    End If

    Me.Command8.Enabled = False 'SEARCH
    
End Sub

Private Sub Form_Unload(cancel As Integer)

    Set rs = Nothing
    Set Me.DataGrid1.DataSource = Nothing

End Sub

Private Sub MSHFlexGrid1_RowColChange()
    
    c_c = Me.MSHFlexGrid1.Col
    ar = Me.Adodc1.Recordset.AbsolutePosition
    r = Me.MSHFlexGrid1.Row
    t = Trim(Me.MSHFlexGrid1.Text)
    l = Me.MSHFlexGrid1.CellLeft + Me.MSHFlexGrid1.Left
    tp = Me.MSHFlexGrid1.CellTop + Me.MSHFlexGrid1.Top
    wd = Me.MSHFlexGrid1.CellWidth - 10
    hg = Me.MSHFlexGrid1.CellHeight - 10
    '
    'Me.Adodc1.Recordset.MoveFirst
    'For i = 1 To r - 1
    '    Me.Adodc1.Recordset.MoveNext
    'Next i
    Me.Adodc1.Recordset.Move r - ar
    '
    Me.MSHFlexGrid1.Col = 0
    Me.MSHFlexGrid1.Row = r
    lb_arrow.Left = Me.MSHFlexGrid1.CellLeft + Me.MSHFlexGrid1.Left
    lb_arrow.Top = Me.MSHFlexGrid1.CellTop + Me.MSHFlexGrid1.Top
    lb_arrow.Width = Me.MSHFlexGrid1.CellWidth
    lb_arrow.Height = Me.MSHFlexGrid1.CellHeight - 5
    Me.MSHFlexGrid1.Col = c_c
    '
    
End Sub

Private Sub opt_Click(Index As Integer)

    If Index = 1 And opt(Index).Value = True Then   'EINAI ≈Œœƒœ
        'Me.Check1.Enabled = False
        'Me.Check1.Value = 0
        For k = 0 To 2
            Me.txt_f(k).ForeColor = &H800000
        Next k
        For k = 3 To 5
            Me.txt_f(k).Text = ""
            Me.txt_f(k).ForeColor = &H800000
            'Me.txt_f(k).Locked = True
            'Me.txt_f(k).BackColor = &HE0E0E0
        Next k
    Else '≈…Õ¡… ≈”œƒœ
        Me.Check1.Enabled = True
        For k = 0 To 2
            Me.txt_f(k).ForeColor = &H8000&
        Next k
        For k = 3 To 5
            'Me.txt_f(k).Text = ""
            'Me.txt_f(k).Locked = False
            'Me.txt_f(k).BackColor = vbWhite
            Me.txt_f(k).ForeColor = &H8000&
        Next k
    End If

End Sub

Private Sub txt_f_LostFocus(Index As Integer)
    
    If Index >= 3 And Index <= 5 Then
        If Val(Me.txt_f(Index).Text) = 0 And Me.txt_f(Index).Text <> "0" Then
            Me.txt_f(Index).Text = ""
        End If
    End If
        
End Sub

Private Sub txt_up_Change()
    
    'Me.MSHFlexGrid1.TextMatrix(Me.MSHFlexGrid1.Row, Me.MSHFlexGrid1.Col) = Trim(Me.txt_up.Text)
    
    
    
End Sub

Private Sub txt_up_LostFocus()

'If Me.MSHFlexGrid1.Col >= 3 And Me.MSHFlexGrid1.Col <= 4 And Trim(Me.txt_up.Text) <> "" Then
'        If Me.DataGrid1.Row <> Me.MSHFlexGrid1.Row - 1 Then
'            r = Me.MSHFlexGrid1.Row
'            Me.DataGrid1.Row = r - 1
'            Me.Adodc1.Recordset.MoveFirst
'            Me.Adodc1.Recordset.Move r - 1
'        End If
'        Me.Adodc1.Recordset.Fields(Me.MSHFlexGrid1.Col - 1).Value = Trim(Me.txt_up.Text)
'        Me.Adodc1.Recordset.UpdateBatch adAffectCurrent
'        Me.MSHFlexGrid1.CellAlignment = flexAlignLeftCenter
'    End If
'    If Me.MSHFlexGrid1.Col >= 5 And Me.MSHFlexGrid1.Col <= 8 And Me.txt_up.Text <> "" Then
'        If Me.DataGrid1.Row <> Me.MSHFlexGrid1.Row - 1 Then
'            r = Me.MSHFlexGrid1.Row
'            Me.DataGrid1.Row = r - 1
'            Me.Adodc1.Recordset.MoveFirst
'            Me.Adodc1.Recordset.Move r - 1
'        End If
'        Me.Adodc1.Recordset.Fields(Me.MSHFlexGrid1.Col - 1).Value = Trim(Me.txt_up.Text)
'        Me.Adodc1.Recordset.UpdateBatch adAffectCurrent
'        Me.MSHFlexGrid1.CellAlignment = flexAlignLeftCenter
'    End If
    
End Sub
