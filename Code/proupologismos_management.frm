VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form proupologismos_management 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Διαχείριση Προϋπολογισμών"
   ClientHeight    =   6720
   ClientLeft      =   1050
   ClientTop       =   1440
   ClientWidth     =   8895
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6720
   ScaleLeft       =   1000
   ScaleMode       =   0  'User
   ScaleTop        =   1000
   ScaleWidth      =   8895
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000014&
      Caption         =   "Εκτύπωση Π/Υ"
      DisabledPicture =   "proupologismos_management.frx":0000
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
      MaskColor       =   &H80000014&
      Picture         =   "proupologismos_management.frx":4A05
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Συνοπτική Εκτύπωση όλων των Προϋπολογισμών"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton insert_bt 
      BackColor       =   &H80000014&
      Caption         =   "Προσθήκη"
      DisabledPicture =   "proupologismos_management.frx":940A
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
      Picture         =   "proupologismos_management.frx":E199
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton save_command 
      BackColor       =   &H80000014&
      Caption         =   "Αποθήκευση"
      DisabledPicture =   "proupologismos_management.frx":12F28
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
      Left            =   2280
      MaskColor       =   &H80000014&
      Picture         =   "proupologismos_management.frx":17B27
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton up_bt 
      BackColor       =   &H80000014&
      Caption         =   "Ενημέρωση"
      DisabledPicture =   "proupologismos_management.frx":1C726
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
      Left            =   1200
      MaskColor       =   &H80000014&
      Picture         =   "proupologismos_management.frx":24220
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton del_bt 
      BackColor       =   &H80000014&
      Caption         =   "Διαγραφή"
      DisabledPicture =   "proupologismos_management.frx":2BD1A
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
      Left            =   3360
      MaskColor       =   &H80000014&
      Picture         =   "proupologismos_management.frx":309CA
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton kl_bt 
      BackColor       =   &H80000014&
      Caption         =   "Κλείσιμο"
      DisabledPicture =   "proupologismos_management.frx":3567A
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
      Left            =   7680
      MaskColor       =   &H80000014&
      Picture         =   "proupologismos_management.frx":3B0F2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton anal_pr_bt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ανάλυση"
      DisabledPicture =   "proupologismos_management.frx":40B6A
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
      Left            =   4440
      MaskColor       =   &H80000014&
      Picture         =   "proupologismos_management.frx":45AB1
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton bt_print 
      BackColor       =   &H80000014&
      Caption         =   "Εκτύπωση Όλων"
      DisabledPicture =   "proupologismos_management.frx":4A9F8
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
      Left            =   6600
      MaskColor       =   &H80000014&
      Picture         =   "proupologismos_management.frx":4F3FD
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Συνοπτική Εκτύπωση όλων των Προϋπολογισμών"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Όλοι οι Προϋπολογισμοί"
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
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   8655
      Begin MSDataGridLib.DataGrid dt_proupol 
         Bindings        =   "proupologismos_management.frx":53E02
         Height          =   2895
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5106
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
      Begin MSAdodcLib.Adodc ado_proupol 
         Height          =   375
         Left            =   120
         Top             =   3240
         Width           =   8415
         _ExtentX        =   14843
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
      Caption         =   "Στοιχεία Προυπολογισμού"
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
      Height          =   1935
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   8655
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
         Height          =   420
         Left            =   1920
         TabIndex        =   5
         Top             =   1275
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
         Left            =   4920
         TabIndex        =   2
         Top             =   360
         Width           =   3615
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mb_im_en 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSMask.MaskEdBox mb_im_liks 
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Top             =   840
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
      Begin MSMask.MaskEdBox mb_im_egkr 
         Height          =   375
         Left            =   4920
         TabIndex        =   6
         Top             =   1320
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
      Begin VB.Label Label10 
         Caption         =   "(ηη/μμ/εεεε)"
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
         Height          =   255
         Left            =   6010
         TabIndex        =   26
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "(ηη/μμ/εεεε)"
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
         Height          =   255
         Left            =   3120
         TabIndex        =   25
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "(ηη/μμ/εεεε)"
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
         Height          =   255
         Left            =   6010
         TabIndex        =   24
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ημερομηνία Έγκρισης:"
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
         Left            =   3000
         TabIndex        =   17
         Top             =   1380
         Width           =   1845
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Έγκριση (Αρ. ΓΣ):"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1380
         Width           =   1605
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
         Left            =   4320
         TabIndex        =   15
         Top             =   900
         Width           =   525
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
         Left            =   1320
         TabIndex        =   14
         Top             =   900
         Width           =   525
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(Διάρκεια Π/Υ)"
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
         Left            =   240
         TabIndex        =   13
         Top             =   900
         Width           =   1245
      End
      Begin VB.Label Label6 
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
         Left            =   3840
         TabIndex        =   12
         Top             =   435
         Width           =   1005
      End
      Begin VB.Label Label1 
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
         Left            =   600
         TabIndex        =   11
         Top             =   435
         Width           =   1245
      End
   End
End
Attribute VB_Name = "proupologismos_management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id_α, flag_pateras, flag_mitera As Integer
Public rs_ado_athlites As ADODB.Recordset
Public rs_ado_dimoi As ADODB.Recordset
Public rs_ado_pe As ADODB.Recordset
Public rs_ado_sxolia As ADODB.Recordset
Public rs_ado_miteres As ADODB.Recordset
Public rs_ado_pateres As ADODB.Recordset
Dim bytData() As Byte
Public sFile As String
'
Public defined_col As Integer
Public s_sort As String
Public is_to_delete, for_search As Integer
    
Private Sub ado_proupol_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    If Not adReason = 7 Then ' adRsnRequery = 7
        If Not adReason = 10 And pRecordset.AbsolutePosition >= 1 Then
            If Trim(pRecordset.Fields(0).Value) <> "" Then
                txt_id.Text = pRecordset.Fields(0).Value
            Else
                txt_id.Text = ""
            End If
            If Trim(pRecordset.Fields(1).Value) <> "" Then
                txt_perigrafi.Text = pRecordset.Fields(1).Value
            Else
                txt_perigrafi.Text = ""
            End If
            If pRecordset.Fields(2).Value <> "" Then
                mb_im_en.Text = pRecordset.Fields(2).Value
            Else
                mb_im_en.Text = "  /  /    "
            End If
            If pRecordset.Fields(3).Value <> "" Then
                mb_im_liks.Text = pRecordset.Fields(3).Value
            Else
                mb_im_liks.Text = "  /  /    "
            End If
            If Trim(pRecordset.Fields(4).Value) <> "" Then
                txt_egkr_gs.Text = pRecordset.Fields(4).Value
            Else
                txt_egkr_gs.Text = ""
            End If
            If pRecordset.Fields(5).Value <> "" Then
                mb_im_egkr.Text = pRecordset.Fields(5).Value
            Else
                mb_im_egkr.Text = "  /  /    "
            End If
        End If
    End If
    If adReason <> 7 Then
        If pRecordset.AbsolutePosition >= 1 Then
            ado_proupol.Caption = "Προϋπολογισμός " & pRecordset.AbsolutePosition & " από " & pRecordset.RecordCount
'            clean_bt.Enabled = True
        End If
    End If
    
End Sub

Private Sub anal_pr_bt_Click()

    anlisi_proupologismou.Show

End Sub

Private Sub bt_print_Click()

    Dim s_str As String
    
    s_str = "id_προϋπολογισμού >= -1"
    Poseidon_DB.rsεκτύπωση_προϋπολογισμών.Filter = Trim(s_str)
    Rep_ektiposi_proypologismwn.Orientation = rptOrientPortrait
    Rep_ektiposi_proypologismwn.Show

End Sub

Private Sub canc_bt_Click()

    
    for_search = 0
    s_string = ""
    s_sort = ""
    txt_id.Locked = True

    Me.insert_bt.Enabled = True
    Me.up_bt.Enabled = True
    'Me.sear_bt.Enabled = False
    'Me.canc_bt.Enabled = False
    
    ado_proupol.Refresh
    ado_proupol.Recordset.Sort = "[id_προϋπολογισμού]"
    '***************************************************************
    If Not ado_proupol.Recordset.EOF Then
        ado_proupol.Recordset.MoveFirst
        If Trim(ado_proupol.Recordset.Fields(0).Value) <> "" Then
            txt_id.Text = ado_proupol.Recordset.Fields(0).Value
        Else
            txt_id.Text = ""
        End If
        If Trim(ado_proupol.Recordset.Fields(1).Value) <> "" Then
            txt_perigrafi.Text = ado_proupol.Recordset.Fields(1).Value
        Else
            txt_perigrafi.Text = ""
        End If
        If ado_proupol.Recordset.Fields(2).Value <> "" Then
            mb_im_en.Text = ado_proupol.Recordset.Fields(2).Value
        Else
            mb_im_en.Text = "  /  /    "
        End If
        If ado_proupol.Recordset.Fields(3).Value <> "" Then
            mb_im_liks.Text = ado_proupol.Recordset.Fields(3).Value
        Else
            mb_im_liks.Text = "  /  /    "
        End If
        If Trim(ado_proupol.Recordset.Fields(4).Value) <> "" Then
            txt_egkr_gs.Text = ado_proupol.Recordset.Fields(4).Value
        Else
            txt_egkr_gs.Text = ""
        End If
        If ado_proupol.Recordset.Fields(5).Value <> "" Then
            mb_im_egkr.Text = ado_proupol.Recordset.Fields(5).Value
        Else
            mb_im_egkr.Text = "  /  /    "
        End If
    End If
    '***************************************************************
    dt_proupol.Columns(0).Caption = "Κωδικός"
    dt_proupol.Columns(1).Width = 1500
    dt_proupol.Columns(1).Caption = "Περιγραφή"
    dt_proupol.Columns(1).Width = 2500
    dt_proupol.Columns(2).Caption = "Ημερομηνία Έναρξης"
    dt_proupol.Columns(2).Width = 2000
    dt_proupol.Columns(3).Caption = "Ημερομηνία Λήξης"
    dt_proupol.Columns(3).Width = 2000
    For i = 4 To ado_proupol.Recordset.Fields.Count - 1
        dt_proupol.Columns(i).Visible = False
    Next i

End Sub

Private Sub Command2_Click()
    meli_management.Show
End Sub

Private Sub clean_bt_Click()

    for_search = 1

    txt_id.Text = ""
    txt_perigrafi.Text = ""
    mb_im_en.Text = "  /  /    "
    mb_im_liks.Text = "  /  /    "
    txt_egkr_gs.Text = ""
    mb_im_egkr.Text = "  /  /    "
    txt_id.Locked = False
    txt_id.SetFocus
    insert_bt.Enabled = False
    save_command.Enabled = False
    up_bt.Enabled = False
    del_bt.Enabled = False
    sear_bt.Enabled = True
            
End Sub

Private Sub Command1_Click()

    Dim s_str As String
    
    s_str = "id_προϋπολογισμού = " & Trim(txt_id.Text)
    Poseidon_DB.rsεκτύπωση_προϋπολογισμών.Filter = Trim(s_str)
    Rep_ektiposi_proypologismwn.Orientation = rptOrientPortrait
    Rep_ektiposi_proypologismwn.Show

End Sub

Private Sub del_bt_Click()
    
    Dim ms As String
    Dim rec_index, vId_Proyp As Integer

    If Not ado_proupol.Recordset.EOF Then
        ms = MsgBox("Είσαι σίγουρος; Μαζί με τον Προϋπολογισμό θα διαγραφούν και τα αντίστοιχα Έσοδα - Έξοδα ... (ΝΑΙ ή ΟΧΙ) ...", vbYesNo, "Παράθυρο διαγραφής")
        If ms = 6 Then
            is_to_delete = 1
            vId_Proyp = Val(txt_id.Text)
            If ado_proupol.Recordset.AbsolutePosition = ado_proupol.Recordset.RecordCount Then
                rec_index = ado_proupol.Recordset.AbsolutePosition - 1
            Else
                rec_index = ado_proupol.Recordset.AbsolutePosition
            End If
            ado_proupol.Recordset.Delete
            ado_proupol.Recordset.Requery
            dt_proupol.Refresh
            ado_proupol.Refresh
            ado_proupol.Recordset.Sort = "[" & Trim(ado_proupol.Recordset.Fields(0).Name) & "]"
            is_to_delete = 0
            If Not ado_proupol.Recordset.EOF Then
                ado_proupol.Recordset.MoveFirst
                ado_proupol.Recordset.Move rec_index - 1, 0
                ado_proupol.Recordset.MovePrevious
                ado_proupol.Recordset.MoveNext
            Else
                txt_id.Text = ""
                txt_perigrafi.Text = ""
                mb_im_en.Text = "  /  /    "
                mb_im_liks.Text = "  /  /    "
                txt_egkr_gs.Text = ""
                mb_im_egkr.Text = "  /  /    "
                Me.ado_proupol.Caption = "Προϋπολογισμός 0 από 0"
                txt_perigrafi.Locked = True
                'mb_im_en.Enabled = False
                'mb_im_liks.Enabled = False
                txt_egkr_gs.Locked = True
                'mb_im_egkr.Enabled = False
                up_bt.Enabled = False
                clean_bt.Enabled = False
                sear_bt.Enabled = False
                insert_bt.Enabled = True
                save_command.Enabled = False
                del_bt.Enabled = False
            End If
            dt_proupol.Columns(0).Caption = "Κωδικός"
            dt_proupol.Columns(1).Width = 1500
            dt_proupol.Columns(1).Caption = "Περιγραφή"
            dt_proupol.Columns(1).Width = 2500
            dt_proupol.Columns(2).Caption = "Ημερομηνία Έναρξης"
            dt_proupol.Columns(2).Width = 2000
            dt_proupol.Columns(3).Caption = "Ημερομηνία Λήξης"
            dt_proupol.Columns(3).Width = 2000
            For i = 4 To ado_proupol.Recordset.Fields.Count - 1
                dt_proupol.Columns(i).Visible = False
            Next i
            'ΑΠΑΙΤΟΥΜΕΝΕΣ ΕΝΕΡΓΕΙΕΣ ΓΙΑ ΤΗ ΔΙΑΓΡΑΦΗ ΤΩΝ ΑΝΤΙΣΤΟΙΧΩΝ ΕΣΟΔΩΝ
            Dim d_recs As New ADODB.Recordset
            d_recs.Open "ΑνάλυσηΠροϋπολογισμούΕσόδων", Me.ado_proupol.ConnectionString, adOpenDynamic, adLockBatchOptimistic
            d_recs.MoveFirst
            Do While Not d_recs.EOF
                If d_recs.Fields(1).Value = vId_Proyp Then
                    d_recs.Delete
                    d_recs.UpdateBatch
                End If
                d_recs.MoveNext
            Loop
            d_recs.Close
            'ΑΠΑΙΤΟΥΜΕΝΕΣ ΕΝΕΡΓΕΙΕΣ ΓΙΑ ΤΗ ΔΙΑΓΡΑΦΗ ΤΩΝ ΑΝΤΙΣΤΟΙΧΩΝ ΕΞΟΔΩΝ
            d_recs.Open "ΑνάλυσηΠροϋπολογισμούΕξόδων", Me.ado_proupol.ConnectionString, adOpenDynamic, adLockBatchOptimistic
            d_recs.MoveFirst
            Do While Not d_recs.EOF
                If d_recs.Fields(1).Value = vId_Proyp Then
                    d_recs.Delete
                    d_recs.UpdateBatch
                End If
                d_recs.MoveNext
            Loop
            d_recs.Close
            '
        Else
            MsgBox "ΑΚΥΡΩΣΗ ΔΙΑΓΡΑΦΗΣ", , "Μήνυμα Προειδοποίησης!"
        End If
    Else
        MsgBox "Δεν υπάρχει Εγγραφή για ΔΙΑΓΡΑΦΗ", , "Μήνυμα Προειδοποίησης!"
    End If
    
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub dt_proupol_HeadClick(ByVal ColIndex As Integer)
    
    defined_col = ColIndex
    
End Sub

Private Sub dt_proupol_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If is_to_delete <> 1 Then
        If ado_proupol.Recordset.AbsolutePosition >= 1 And ado_proupol.Recordset.AbsolutePosition <= ado_proupol.Recordset.RecordCount Then
            If Trim(ado_proupol.Recordset.Fields(0).Value) <> "" Then
                txt_id.Text = ado_proupol.Recordset.Fields(0).Value
            Else
                txt_id.Text = ""
            End If
            If Trim(ado_proupol.Recordset.Fields(1).Value) <> "" Then
                txt_perigrafi.Text = ado_proupol.Recordset.Fields(1).Value
            Else
                txt_perigrafi.Text = ""
            End If
            If Trim(ado_proupol.Recordset.Fields(2).Value) <> "" Then
                mb_im_en.Text = ado_proupol.Recordset.Fields(2).Value
            Else
                mb_im_en.Text = "  /  /    "
            End If
            If Trim(ado_proupol.Recordset.Fields(3).Value) <> "" Then
                mb_im_liks.Text = ado_proupol.Recordset.Fields(3).Value
            Else
                mb_im_liks.Text = "  /  /    "
            End If
            If Trim(ado_proupol.Recordset.Fields(4).Value) <> "" Then
                txt_egkr_gs.Text = ado_proupol.Recordset.Fields(4).Value
            Else
                txt_egkr_gs.Text = ""
            End If
            If Trim(ado_proupol.Recordset.Fields(5).Value) <> "" Then
                mb_im_egkr.Text = ado_proupol.Recordset.Fields(5).Value
            Else
                mb_im_egkr.Text = "  /  /    "
            End If
        End If
    End If
    
End Sub

Private Sub Form_Load()
  
    'Me.Height = 10935
    'Me.Width = 11835
    Me.Top = 400
    Me.Left = 400
    
    'Me.ado_proupol.Refresh
  
    's_sort = ""
    is_to_delete = 0
    for_search = 0
    Me.txt_id.Locked = True
    
    ado_proupol.Recordset.Sort = "[" & Trim(ado_proupol.Recordset.Fields(0).Name) & "]"
    '******************************************************
    If Not ado_proupol.Recordset.EOF Then
        ado_proupol.Recordset.MoveFirst
        If Trim(ado_proupol.Recordset.Fields(0).Value) <> "" Then
            txt_id.Text = ado_proupol.Recordset.Fields(0).Value
        Else
            txt_id.Text = ""
        End If
        If Trim(ado_proupol.Recordset.Fields(1).Value) <> "" Then
            txt_perigrafi.Text = ado_proupol.Recordset.Fields(1).Value
        Else
            txt_perigrafi.Text = ""
        End If
        If Trim(ado_proupol.Recordset.Fields(2).Value) <> "" Then
            mb_im_en.Text = ado_proupol.Recordset.Fields(2).Value
        Else
            mb_im_en.Text = "  /  /    "
        End If
        If Trim(ado_proupol.Recordset.Fields(3).Value) <> "" Then
            mb_im_liks.Text = ado_proupol.Recordset.Fields(3).Value
        Else
            mb_im_liks.Text = "  /  /    "
        End If
        If Trim(ado_proupol.Recordset.Fields(4).Value) <> "" Then
            txt_egkr_gs.Text = ado_proupol.Recordset.Fields(4).Value
        Else
            txt_egkr_gs.Text = ""
        End If
        If Trim(ado_proupol.Recordset.Fields(5).Value) <> "" Then
            mb_im_egkr.Text = ado_proupol.Recordset.Fields(5).Value
        Else
            mb_im_egkr.Text = "  /  /    "
        End If
    End If
    
    '***************************************************************
    If ado_proupol.Recordset.RecordCount > 0 Then
        dt_proupol.Row = 0
        dt_proupol.Col = 1
        txt_perigrafi.Locked = False
        'mb_im_en.Enabled = True
        'mb_im_liks.Enabled = True
        txt_egkr_gs.Locked = False
        'mb_im_egkr.Enabled = True
        Me.ado_proupol.Caption = "Προϋπολογισμός " & dt_proupol.Row + 1 & " από " & ado_proupol.Recordset.RecordCount
        up_bt.Enabled = True
        del_bt.Enabled = True
'        clean_bt.Enabled = True
        anal_pr_bt.Enabled = True
    Else
        Me.ado_proupol.Caption = "Προϋπολογισμός 0 από 0"
        txt_perigrafi.Locked = True
        'mb_im_en.Enabled = False
        'mb_im_liks.Enabled = False
        txt_egkr_gs.Locked = True
        'mb_im_egkr.Enabled = False
    End If
    dt_proupol.Columns(0).Caption = "Κωδικός"
    dt_proupol.Columns(1).Width = 1500
    dt_proupol.Columns(1).Caption = "Περιγραφή"
    dt_proupol.Columns(1).Width = 2750
    dt_proupol.Columns(2).Caption = "Ημερομηνία Έναρξης"
    dt_proupol.Columns(2).Width = 2000
    dt_proupol.Columns(3).Caption = "Ημερομηνία Λήξης"
    dt_proupol.Columns(3).Width = 2000
    For i = 4 To ado_proupol.Recordset.Fields.Count - 1
        dt_proupol.Columns(i).Visible = False
    Next i
    
End Sub

Private Sub fr_sav_but_Click()
    Frame10.Visible = False
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        PopupMenu mn1
    End If

End Sub

Private Sub insert_bt_Click()

    Dim id_π As Integer

    'Να βρω το υποψήφιο id_προϋπολογισμού
    ado_proupol.Recordset.Sort = "[" & Trim(ado_proupol.Recordset.Fields(0).Name) & "]"
    If Not ado_proupol.Recordset.EOF Then
        ado_proupol.Recordset.MoveLast
        id_π = ado_proupol.Recordset.Fields(0).Value
        id_π = id_π + 1
    Else
        id_π = 1
    End If
    
    'Ενεργοποίηση και Καθαρισμός πεδίων
    txt_perigrafi.Locked = False
    'mb_im_en.Enabled = True
    'mb_im_liks.Enabled = True
    txt_egkr_gs.Locked = False
    'mb_im_egkr.Enabled = True
    txt_id.Text = id_π
    txt_perigrafi.Text = ""
    mb_im_en.Text = "  /  /    "
    mb_im_liks.Text = "  /  /    "
    txt_egkr_gs.Text = ""
    mb_im_egkr.Text = "  /  /    "
    
    txt_perigrafi.SetFocus
    
    'Το grid όλων των διαθέσιμων Π/Υ απενεργοποιείται, όταν προσθέτω ΝΕΟ
    dt_proupol.Enabled = False
    
    Frame2.Caption = Frame2.Caption & " (είστε στη διαδικασία ΠΡΟΣΘΗΚΗΣ ΝΕΟΥ Π/Υ)"
    
    insert_bt.Enabled = False
    save_command.Enabled = True
'    canc_bt.Enabled = True
    up_bt.Enabled = False
    del_bt.Enabled = False
'    clean_bt.Enabled = False
    anal_pr_bt.Enabled = False
    
End Sub

Private Sub kl_bt_Click()

    Unload Me

End Sub



Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mn1
    End If
    Label19.Visible = False
End Sub

Private Sub mb_im_egkr_GotFocus()
    
    mb_im_egkr.SelStart = 0
    mb_im_egkr.SelLength = 10

End Sub

Private Sub mb_im_egkr_LostFocus()

    If for_search = 0 And IsDate(mb_im_egkr.Text) = False And (mb_im_egkr.Text <> "__/__/____") And (mb_im_egkr.Text <> "  /  /    ") Then
        mb_im_egkr.SelStart = 0
        mb_im_egkr.SelLength = 10
        mb_im_egkr.SetFocus
    End If
End Sub

Private Sub mb_im_en_GotFocus()
    
    mb_im_en.SelStart = 0
    mb_im_en.SelLength = 10
    
End Sub

Private Sub mb_im_en_LostFocus()
    
    If for_search = 0 And IsDate(mb_im_en.Text) = False And (mb_im_en.Text <> "__/__/____") And (mb_im_en.Text <> "  /  /    ") Then
        mb_im_en.SelStart = 0
        mb_im_en.SelLength = 10
        mb_im_en.SetFocus
    End If
    
End Sub

Private Sub mb_im_liks_GotFocus()
    
    mb_im_liks.SelStart = 0
    mb_im_liks.SelLength = 10
    
End Sub

Private Sub mb_im_liks_LostFocus()

    If for_search = 0 And IsDate(mb_im_liks.Text) = False And (mb_im_liks.Text <> "__/__/____") And (mb_im_liks.Text <> "  /  /    ") Then
        mb_im_liks.SelStart = 0
        mb_im_liks.SelLength = 10
        mb_im_liks.SetFocus
    End If
    
End Sub

Private Sub popmn1_Click()
    
    f_foto_a.Show
    
End Sub

Private Sub save_command_Click()
    
    Dim tmp1, tmp2, tmp3, tmp4, tmp5, tmp6 As String
    
    MousePointer = 11
    
    tmp1 = ""
    tmp2 = ""
    tmp3 = "__/__/____"
    tmp4 = "__/__/____"
    tmp5 = ""
    tmp6 = "__/__/____"
    
    If Trim(Me.txt_id.Text) <> "" Then
        tmp1 = txt_id.Text
    End If
    If Trim(txt_perigrafi.Text) <> "" Then
        tmp2 = txt_perigrafi.Text
    End If
    If mb_im_en.Text <> "  /  /    " And mb_im_en.Text <> "__/__/____" Then
        tmp3 = mb_im_en.Text
    End If
    If mb_im_liks.Text <> "  /  /    " And mb_im_liks.Text <> "__/__/____" Then
        tmp4 = mb_im_liks.Text
    End If
    If Trim(txt_egkr_gs.Text) <> "" Then
        tmp5 = txt_egkr_gs.Text
    End If
    If mb_im_egkr.Text <> "  /  /    " And mb_im_egkr.Text <> "__/__/____" Then
        tmp6 = mb_im_egkr.Text
    End If
    
    'Αποθήκευση στους προϋπολογισμούς
    ado_proupol.Recordset.AddNew
    'Αποθήκευση ΛΟΙΠΩΝ ΣΤΟΙΧΕΙΑ
    If Trim(tmp1) <> "" Then
        ado_proupol.Recordset.Fields(0).Value = tmp1
    End If
    If Trim(tmp2) <> "" Then
        ado_proupol.Recordset.Fields(1).Value = tmp2
    End If
    If tmp3 <> "  /  /    " And tmp3 <> "__/__/____" Then
        ado_proupol.Recordset.Fields(2).Value = tmp3
    End If
    If tmp4 <> "  /  /    " And tmp4 <> "__/__/____" Then
        ado_proupol.Recordset.Fields(3).Value = tmp4
    End If
    If tmp5 <> "" Then
        ado_proupol.Recordset.Fields(4).Value = tmp5
    End If
    If tmp6 <> "  /  /    " And tmp6 <> "__/__/____" Then
        ado_proupol.Recordset.Fields(5).Value = tmp6
    End If
    
    ado_proupol.Recordset.UpdateBatch adAffectCurrent
    
    ado_proupol.Recordset.Requery
    dt_proupol.Refresh
    ado_proupol.Refresh
    dt_proupol.Columns(0).Caption = "Κωδικός"
    dt_proupol.Columns(1).Width = 1500
    dt_proupol.Columns(1).Caption = "Περιγραφή"
    dt_proupol.Columns(1).Width = 2500
    dt_proupol.Columns(2).Caption = "Ημερομηνία Έναρξης"
    dt_proupol.Columns(2).Width = 2000
    dt_proupol.Columns(3).Caption = "Ημερομηνία Λήξης"
    dt_proupol.Columns(3).Width = 2000
    For i = 4 To ado_proupol.Recordset.Fields.Count - 1
        dt_proupol.Columns(i).Visible = False
    Next i
    ado_proupol.Recordset.Sort = "[" & Trim(ado_proupol.Recordset.Fields(0).Name) & "]"
    ado_proupol.Recordset.MoveLast
    
    'Το grid όλων των διαθέσιμων Π/Υ εργοποιείται
    dt_proupol.Enabled = True
    
    Frame2.Caption = "Διαχείριση Προϋπολογισμών"
    
    save_command.Enabled = False
'    canc_bt.Enabled = True
    insert_bt.Enabled = True
    up_bt.Enabled = True
    del_bt.Enabled = True
'    clean_bt.Visible = True
    anal_pr_bt.Enabled = True
        
    Me.MousePointer = 0
    
End Sub

Private Sub sear_bt_Click()

    
    s_string = ""
    If Trim(txt_id.Text) <> "" Then
        s_string = "[id_προϋπολογισμού] LIKE " & Trim(txt_id.Text)
    End If
    If Trim(txt_perigrafi.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [Περιγραφή] LIKE '*" & Trim(txt_perigrafi.Text) & "*'"
        Else
            s_string = "[Περιγραφή] LIKE '*" & Trim(txt_perigrafi.Text) & "*'"
        End If
    End If
    'ΚΡΙΤΗΡΙΟ ΗΜΕΡΟΜΗΝΙΑ ΕΝΑΡΞΗΣ
    'If Trim(MaskEdBox1.Text) <> "00/00/0000" Then
        Dim imera, year, minas As String
        mb_im_en.SelStart = 0
        mb_im_en.SelLength = 2
        imera = mb_im_en.SelText
        mb_im_en.SelStart = 3
        mb_im_en.SelLength = 2
        minas = mb_im_en.SelText
        mb_im_en.SelStart = 6
        mb_im_en.SelLength = 4
        year = mb_im_en.SelText
        mb_im_en.SelStart = 0
        mb_im_en.SelLength = 10
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
                s_string = s_string & "AND [Έναρξη] LIKE '" & st1 & "'"
            End If
            If st2 <> "" Then
                s_string = s_string & " AND [Έναρξη] LIKE '" & st2 & "'"
            End If
            If st3 <> "" Then
                    s_string = s_string & " AND [Έναρξη] LIKE '" & st3 & "'"
            End If
        Else
            If st1 <> "" Then
                s_string = "[Έναρξη] LIKE '" & st1 & "'"
            End If
            If st2 <> "" Then
                If s_string <> "" Then
                    s_string = s_string & " AND [Έναρξη] LIKE '" & st2 & "'"
                Else
                    s_string = "[Έναρξη] LIKE '" & st2 & "'"
                End If
            End If
            If st3 <> "" Then
                If s_string <> "" Then
                    s_string = s_string & " AND [Έναρξη] LIKE '" & st3 & "'"
                Else
                    s_string = "[Έναρξη] LIKE '" & st3 & "'"
                End If
            End If
        End If
    'End If
    '
    'ΚΡΙΤΗΡΙΟ ΗΜΕΡΟΜΗΝΙΑ ΛΗΞΗΣ
    'If Trim(MaskEdBox1.Text) <> "00/00/0000" Then
        mb_im_liks.SelStart = 0
        mb_im_liks.SelLength = 2
        imera = mb_im_liks.SelText
        mb_im_liks.SelStart = 3
        mb_im_liks.SelLength = 2
        minas = mb_im_liks.SelText
        mb_im_liks.SelStart = 6
        mb_im_liks.SelLength = 4
        year = mb_im_liks.SelText
        mb_im_liks.SelStart = 0
        mb_im_liks.SelLength = 10
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
                s_string = s_string & "AND [Λήξη] LIKE '" & st1 & "'"
            End If
            If st2 <> "" Then
                s_string = s_string & " AND [Λήξη] LIKE '" & st2 & "'"
            End If
            If st3 <> "" Then
                    s_string = s_string & " AND [Λήξη] LIKE '" & st3 & "'"
            End If
        Else
            If st1 <> "" Then
                s_string = "[Λήξη] LIKE '" & st1 & "'"
            End If
            If st2 <> "" Then
                If s_string <> "" Then
                    s_string = s_string & " AND [Λήξη] LIKE '" & st2 & "'"
                Else
                    s_string = "[Λήξη] LIKE '" & st2 & "'"
                End If
            End If
            If st3 <> "" Then
                If s_string <> "" Then
                    s_string = s_string & " AND [Λήξη] LIKE '" & st3 & "'"
                Else
                    s_string = "[Λήξη] LIKE '" & st3 & "'"
                End If
            End If
        End If
    'End If
    '
    If Trim(txt_egkr_gs.Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [ΈγκρισηΓΣ] LIKE '*" & Trim(txt_egkr_gs.Text) & "*'"
        Else
            s_string = "[ΈγκρισηΓΣ] LIKE '*" & Trim(txt_egkr_gs.Text) & "*'"
        End If
    End If
    '
    '
    'ΚΡΙΤΗΡΙΟ ΗΜΕΡΟΜΗΝΙΑ ΈΓΚΡΙΣΗΣ
    'If Trim(MaskEdBox1.Text) <> "00/00/0000" Then
        mb_im_egkr.SelStart = 0
        mb_im_egkr.SelLength = 2
        imera = mb_im_egkr.SelText
        mb_im_egkr.SelStart = 3
        mb_im_egkr.SelLength = 2
        minas = mb_im_egkr.SelText
        mb_im_egkr.SelStart = 6
        mb_im_egkr.SelLength = 4
        year = mb_im_egkr.SelText
        mb_im_egkr.SelStart = 0
        mb_im_egkr.SelLength = 10
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
                s_string = s_string & "AND [ΗμερομηνίαΈγκρισης] LIKE '" & st1 & "'"
            End If
            If st2 <> "" Then
                s_string = s_string & " AND [ΗμερομηνίαΈγκρισης] LIKE '" & st2 & "'"
            End If
            If st3 <> "" Then
                    s_string = s_string & " AND [ΗμερομηνίαΈγκρισης] LIKE '" & st3 & "'"
            End If
        Else
            If st1 <> "" Then
                s_string = "[ΗμερομηνίαΈγκρισης] LIKE '" & st1 & "'"
            End If
            If st2 <> "" Then
                If s_string <> "" Then
                    s_string = s_string & " AND [ΗμερομηνίαΈγκρισης] LIKE '" & st2 & "'"
                Else
                    s_string = "[ΗμερομηνίαΈγκρισης] LIKE '" & st2 & "'"
                End If
            End If
            If st3 <> "" Then
                If s_string <> "" Then
                    s_string = s_string & " AND [ΗμερομηνίαΈγκρισης] LIKE '" & st3 & "'"
                Else
                    s_string = "[ΗμερομηνίαΈγκρισης] LIKE '" & st3 & "'"
                End If
            End If
        End If
    'End If
    '
    '
    If s_string <> "" Then
        ado_proupol.Recordset.Filter = Trim(s_string)
        If Not ado_proupol.Recordset.EOF Then
            ado_proupol.Recordset.MoveFirst
        End If
    End If
    If s_sort <> "" Then
        ado_proupol.Recordset.Sort = Trim(s_sort)
    End If
    
    'Me.canc_bt.Enabled = True
    
    
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

Private Sub TabStrip1_Click()

    If TabStrip1.SelectedItem.Index = 2 Then
        athl_tmima_management.Show
    End If

End Sub

Private Sub taksin_Click()

    If ado_proupol.Recordset.RecordCount > 0 Then
        If dt_proupol.Col >= 0 Then
            ado_proupol.Recordset.Sort = "[" & Trim(ado_proupol.Recordset.Fields(dt_proupol.Col).Name) & "]"
            s_sort = "[" & Trim(ado_proupol.Recordset.Fields(dt_proupol.Col).Name) & "]"
        Else
            ado_proupol.Recordset.Sort = "[" & Trim(ado_proupol.Recordset.Fields(defined_col).Name) & "]"
            s_sort = "[" & Trim(ado_proupol.Recordset.Fields(defined_col).Name) & "]"
        End If
    End If
  
End Sub

Private Sub Text2_Change()

End Sub

Private Sub up_bt_Click()

    Dim id As Integer
    Dim f_ole As Object
    
    MousePointer = 11
    
    If Trim(txt_id.Text) <> "" Then
        ado_proupol.Recordset.Fields(0).Value = txt_id.Text
    End If
    If Trim(txt_perigrafi.Text) <> "" Then
        ado_proupol.Recordset.Fields(1).Value = txt_perigrafi.Text
    End If
    If mb_im_en.Text <> "  /  /    " And mb_im_en.Text <> "__/__/____" Then
        ado_proupol.Recordset.Fields(2).Value = mb_im_en.Text
    End If
    If mb_im_liks.Text <> "  /  /    " And mb_im_liks.Text <> "__/__/____" Then
        ado_proupol.Recordset.Fields(3).Value = mb_im_liks.Text
    End If
    If Trim(txt_egkr_gs.Text) <> "" Then
        ado_proupol.Recordset.Fields(4).Value = txt_egkr_gs.Text
    End If
    If mb_im_egkr.Text <> "  /  /    " And mb_im_egkr.Text <> "__/__/____" Then
        ado_proupol.Recordset.Fields(5).Value = mb_im_egkr.Text
    End If
    '
    ado_proupol.Recordset.UpdateBatch adAffectCurrent
    
    MousePointer = 0
    
End Sub
