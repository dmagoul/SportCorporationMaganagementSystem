VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form oikonomikes_kiniseis_management 
   BackColor       =   &H00C0C0FF&
   Caption         =   "���������� ����������� ��������"
   ClientHeight    =   11445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11445
   ScaleMode       =   0  'User
   ScaleWidth      =   21759.04
   Visible         =   0   'False
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   161
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Left            =   9944
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   20
      Visible         =   0   'False
      Width           =   342
   End
   Begin MSDataGridLib.DataGrid tmp2_dt_oikon_kiniseis 
      Bindings        =   "oikonomikes_kiniseis_management.frx":0000
      Height          =   2055
      Left            =   3240
      TabIndex        =   34
      Top             =   3120
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
      Appearance      =   0
      ColumnHeaders   =   0   'False
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
   Begin MSAdodcLib.Adodc tmp2_ado_oik_kiniseis 
      Height          =   330
      Left            =   600
      Top             =   3480
      Visible         =   0   'False
      Width           =   6795
      _ExtentX        =   11986
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
      RecordSource    =   "Rep_��������_�����������_�����������_��������"
      Caption         =   "tmp2_ado_oik_kiniseis"
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
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000014&
      Caption         =   "�&������ ����������� ���. �������� ��� Excel"
      DisabledPicture =   "oikonomikes_kiniseis_management.frx":0024
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
      Height          =   855
      Left            =   12425
      MaskColor       =   &H80000014&
      Picture         =   "oikonomikes_kiniseis_management.frx":26AF
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   10440
      Width           =   1754
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000014&
      Caption         =   "�������� ��������� ���. �������"
      DisabledPicture =   "oikonomikes_kiniseis_management.frx":4D3A
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
      Height          =   855
      Left            =   14160
      MaskColor       =   &H80000014&
      Picture         =   "oikonomikes_kiniseis_management.frx":973F
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "�������� ��������� ����������� �������"
      Top             =   10440
      Width           =   1754
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000014&
      Caption         =   "����� ���� ��� �������� ���. ������"
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
      Height          =   855
      Left            =   5400
      MaskColor       =   &H80000014&
      Picture         =   "oikonomikes_kiniseis_management.frx":E144
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   10440
      Width           =   1754
   End
   Begin MSDataGridLib.DataGrid tmp_dt_oikon_kiniseis 
      Bindings        =   "oikonomikes_kiniseis_management.frx":E6FB
      Height          =   2055
      Left            =   480
      TabIndex        =   29
      Top             =   4680
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
      Appearance      =   0
      ColumnHeaders   =   0   'False
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
   Begin MSAdodcLib.Adodc tmp_ado_oik_kiniseis 
      Height          =   330
      Left            =   480
      Top             =   4200
      Visible         =   0   'False
      Width           =   6795
      _ExtentX        =   11986
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
      RecordSource    =   "�������������������"
      Caption         =   "tmp_ado_oik_kiniseis"
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
   Begin VB.TextBox s_c 
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
      Index           =   5
      Left            =   16854
      TabIndex        =   28
      Top             =   240
      Visible         =   0   'False
      Width           =   1052
   End
   Begin VB.CommandButton bt_print 
      BackColor       =   &H80000014&
      Caption         =   "�������� ����������� ���. ��������"
      DisabledPicture =   "oikonomikes_kiniseis_management.frx":E71E
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
      Height          =   855
      Left            =   15924
      MaskColor       =   &H80000014&
      Picture         =   "oikonomikes_kiniseis_management.frx":13123
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "�������� ����������� ����������� ��������"
      Top             =   10440
      Width           =   1754
   End
   Begin VB.TextBox s_c 
      Alignment       =   1  'Right Justify
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
      Index           =   1
      Left            =   4244
      TabIndex        =   26
      ToolTipText     =   "�������� ���������� ����� ������������� ����������� (�.�. 10/9/2013) � �������� ���������� ����� ���� (�.�. 9 ��� ��� ����������)"
      Top             =   240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox co_katastasi_kinisis 
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
      Left            =   3360
      TabIndex        =   25
      Top             =   240
      Visible         =   0   'False
      Width           =   903
   End
   Begin VB.TextBox s_xr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   14513
      TabIndex        =   20
      Top             =   9680
      Width           =   1052
   End
   Begin VB.TextBox s_pis 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   16902
      TabIndex        =   19
      Top             =   9680
      Width           =   1052
   End
   Begin VB.TextBox s_c 
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
      Index           =   6
      Left            =   17888
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   1070
   End
   Begin VB.TextBox s_c 
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
      Index           =   4
      Left            =   14468
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   1052
   End
   Begin VB.TextBox s_c 
      Alignment       =   1  'Right Justify
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
      Index           =   3
      Left            =   12013
      TabIndex        =   10
      ToolTipText     =   "�������� ���������� ����� ������������� ����������� (�.�. 10/9/2013) � �������� ���������� ����� ���� (�.�. 9 ��� ��� ����������)"
      Top             =   240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox s_c 
      Alignment       =   1  'Right Justify
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
      Index           =   2
      Left            =   11605
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   438
   End
   Begin VB.TextBox s_c 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   368
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   438
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H80000014&
      Caption         =   "��������"
      DisabledPicture =   "oikonomikes_kiniseis_management.frx":17B28
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   17640
      MaskColor       =   &H80000014&
      Picture         =   "oikonomikes_kiniseis_management.frx":17C70
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10440
      Width           =   1395
   End
   Begin VB.CommandButton search 
      BackColor       =   &H80000014&
      Caption         =   "���������"
      DisabledPicture =   "oikonomikes_kiniseis_management.frx":1D6E8
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
      Height          =   855
      Left            =   8909
      MaskColor       =   &H80000014&
      Picture         =   "oikonomikes_kiniseis_management.frx":22A9C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   10440
      Width           =   1754
   End
   Begin VB.CommandButton cancel 
      BackColor       =   &H80000014&
      Caption         =   "�������"
      DisabledPicture =   "oikonomikes_kiniseis_management.frx":27E50
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
      Height          =   855
      Left            =   10663
      MaskColor       =   &H80000014&
      Picture         =   "oikonomikes_kiniseis_management.frx":27F98
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10440
      Width           =   1754
   End
   Begin VB.CommandButton in_order_to_filter 
      BackColor       =   &H80000014&
      Caption         =   "�������� ����������"
      DisabledPicture =   "oikonomikes_kiniseis_management.frx":2CEB5
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
      Height          =   855
      Left            =   7155
      MaskColor       =   &H80000014&
      Picture         =   "oikonomikes_kiniseis_management.frx":2CFFD
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10440
      Width           =   1754
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000014&
      Caption         =   "����������� ��������� ���. �������"
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
      Height          =   855
      Left            =   3630
      MaskColor       =   &H80000014&
      Picture         =   "oikonomikes_kiniseis_management.frx":31B41
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   10440
      Width           =   1754
   End
   Begin VB.CommandButton lb_arrow 
      Appearance      =   0  'Flat
      DisabledPicture =   "oikonomikes_kiniseis_management.frx":3963B
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      MaskColor       =   &H00E0E0E0&
      Picture         =   "oikonomikes_kiniseis_management.frx":396FE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9720
      Visible         =   0   'False
      Width           =   490
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000014&
      Caption         =   "����������"
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
      Height          =   855
      Left            =   35
      MaskColor       =   &H80000014&
      Picture         =   "oikonomikes_kiniseis_management.frx":397C1
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10440
      Width           =   1859
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000014&
      Caption         =   "�������� ���� ���. �������"
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
      Height          =   855
      Left            =   1877
      MaskColor       =   &H80000014&
      Picture         =   "oikonomikes_kiniseis_management.frx":3E4C8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10440
      Width           =   1754
   End
   Begin MSAdodcLib.Adodc ado_oikon_kiniseis 
      Height          =   330
      Left            =   35
      Top             =   9960
      Width           =   19019
      _ExtentX        =   33549
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
      RecordSource    =   "Rep_��������_�����������_�����������_��������"
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
   Begin MSDataListLib.DataCombo co_athlites 
      Bindings        =   "oikonomikes_kiniseis_management.frx":43257
      Height          =   345
      Left            =   5370
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      _Version        =   393216
      IntegralHeight  =   0   'False
      ListField       =   "���"
      BoundColumn     =   "id_������������"
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
   Begin MSAdodcLib.Adodc ado_athlites 
      Height          =   375
      Left            =   5880
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
      RecordSource    =   "��������������������"
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
   Begin MSDataListLib.DataCombo co_meli 
      Bindings        =   "oikonomikes_kiniseis_management.frx":43272
      Height          =   345
      Left            =   7020
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      _Version        =   393216
      IntegralHeight  =   0   'False
      ListField       =   "OE_���"
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
   Begin MSAdodcLib.Adodc ado_meli 
      Height          =   375
      Left            =   7320
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
      RecordSource    =   "������������������"
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
   Begin MSDataListLib.DataCombo co_organismoi 
      Bindings        =   "oikonomikes_kiniseis_management.frx":43289
      Height          =   345
      Left            =   8700
      TabIndex        =   14
      Top             =   240
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   609
      _Version        =   393216
      IntegralHeight  =   0   'False
      ListField       =   "��������"
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
   Begin MSAdodcLib.Adodc ado_organismoi 
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
      RecordSource    =   "�����������������������"
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
   Begin MSDataListLib.DataCombo co_tipoi_parastatikwn 
      Bindings        =   "oikonomikes_kiniseis_management.frx":432A6
      Height          =   345
      Left            =   9930
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      _Version        =   393216
      IntegralHeight  =   0   'False
      ListField       =   "LBL"
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
   Begin MSAdodcLib.Adodc ado_tipoi_esodwn 
      Height          =   375
      Left            =   15720
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
      RecordSource    =   "�����������"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo co_tipoi_eksodwn 
      Bindings        =   "oikonomikes_kiniseis_management.frx":432CB
      Height          =   345
      Left            =   13155
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      IntegralHeight  =   0   'False
      ListField       =   "���������"
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
   Begin MSDataListLib.DataCombo co_tipoi_esodwn 
      Bindings        =   "oikonomikes_kiniseis_management.frx":432EB
      Height          =   315
      Left            =   15521
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   1337
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      IntegralHeight  =   0   'False
      ListField       =   "���������"
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc ado_tipoi_eksodwn 
      Height          =   375
      Left            =   13680
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
      RecordSource    =   "�����������"
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
   Begin MSDataListLib.DataCombo co_py 
      Bindings        =   "oikonomikes_kiniseis_management.frx":4330A
      Height          =   345
      Left            =   795
      TabIndex        =   23
      Top             =   240
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   609
      _Version        =   393216
      IntegralHeight  =   0   'False
      ListField       =   "���������"
      BoundColumn     =   "id_������������"
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
   Begin MSAdodcLib.Adodc ado_proipologismos 
      Height          =   375
      Left            =   720
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
      RecordSource    =   "��������������"
      Caption         =   "ado_proipologismos"
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
   Begin VB.ComboBox co_tipos_kinisis 
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
      ItemData        =   "oikonomikes_kiniseis_management.frx":4332B
      Left            =   2560
      List            =   "oikonomikes_kiniseis_management.frx":4332D
      TabIndex        =   24
      Top             =   240
      Visible         =   0   'False
      Width           =   789
   End
   Begin MSDataGridLib.DataGrid dt_oikon_kiniseis 
      Bindings        =   "oikonomikes_kiniseis_management.frx":4332F
      Height          =   9015
      Left            =   35
      TabIndex        =   35
      Top             =   600
      Visible         =   0   'False
      Width           =   18975
      _ExtentX        =   33470
      _ExtentY        =   15901
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   41
      TabAction       =   1
      RowDividerStyle =   3
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
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
   Begin VB.ComboBox lst_mn 
      Height          =   315
      ItemData        =   "oikonomikes_kiniseis_management.frx":43350
      Left            =   10286
      List            =   "oikonomikes_kiniseis_management.frx":43378
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc ado_tipoi_parastatikwn 
      Height          =   375
      Left            =   10440
      Top             =   120
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
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
      RecordSource    =   "������������������������������"
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
      BackStyle       =   0  'Transparent
      Caption         =   "������� �/�:"
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
      Left            =   840
      TabIndex        =   32
      Top             =   30
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���. ����.:"
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
      Height          =   615
      Left            =   15585
      TabIndex        =   22
      Top             =   9675
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "���. �����.:"
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
      Height          =   735
      Left            =   12420
      TabIndex        =   21
      Top             =   9675
      Width           =   2085
   End
End
Attribute VB_Name = "oikonomikes_kiniseis_management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public global_py, global_kk, global_tip_eks, global_tip_es As Integer
Dim global_minas_sindromis(12) As Boolean
Public global_poso_xreosis, global_poso_pistosis As Long
Public defined_col, c_r, g_left As Integer
Public cur_row As Integer
Public rs As ADODB.Recordset

Private Sub ado_oikon_kiniseis_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    If (adReason = adRsnMove Or adReason = adRsnMoveLast Or adReason = adRsnMoveFirst Or adReason = adRsnMoveNext Or adReason = adRsnMovePrevious) And adReason <> adRsnAddNew And Me.ado_oikon_kiniseis.Recordset.AbsolutePosition >= 1 Then
        Me.ado_oikon_kiniseis.Caption = "����������� " & Me.ado_oikon_kiniseis.Recordset.AbsolutePosition & " ��� " & Me.ado_oikon_kiniseis.Recordset.RecordCount
    End If
    
End Sub

Private Sub bt_print_Click()

    'Rep_�������������������.Show
    Rep_�������������������.Hide
    If MDIForm1.s_string <> "" Then
        'Poseidon_DB.rs��������_�����������_��������.Filter = MDIForm1.s_string
        Poseidon_DB.rs��������_�����������_��������.Filter = MDIForm1.s_string
    Else
        'Poseidon_DB.rs��������_�����������_��������.Filter = ""
        Poseidon_DB.rs��������_�����������_��������.Filter = ""
    End If
    If MDIForm1.s_sort <> "" Then
        'Poseidon_DB.rs��������_�����������_��������.Sort = MDIForm1.s_sort
        Poseidon_DB.rs��������_�����������_��������.Sort = MDIForm1.s_sort
    Else
        'Poseidon_DB.rs��������_�����������_��������.Sort = "[id]"
        Poseidon_DB.rs��������_�����������_��������.Sort = "[id]"
    End If
    Rep_�������������������.Sections("ReportHeader").Controls("label5").Caption = Me.co_py.Text
    MDIForm1.s_sort = "[id]"
    Rep_�������������������.Orientation = rptOrientPortrait
    Rep_�������������������.Show

End Sub

Private Sub cancel_Click()
    
    
    'Me.co_py.Locked = False
    
    'Me.co_py.Text = ""
    
    For i = 0 To 6
        Me.s_c(i).Visible = False
    Next i
    'Me.co_py.Visible = False
    Me.co_tipos_kinisis.Visible = False
    Me.co_katastasi_kinisis.Visible = False
    Me.co_athlites.Visible = False
    Me.co_meli.Visible = False
    Me.co_organismoi.Visible = False
    Me.co_tipoi_parastatikwn.Visible = False
    Me.Command7.Visible = False
    Me.lst_mn.Visible = False
    Me.lst_mn.ListIndex = -1
    Me.co_tipoi_eksodwn.Visible = False
    Me.co_tipoi_esodwn.Visible = False
    
    id_s = ������ID_���_String("��������������", Me.co_py.Text)
    MDIForm1.s_string = "[����������������] = 1 AND [id_py] LIKE " & id_s
    Me.co_py.Locked = False
    
    Me.ado_oikon_kiniseis.Recordset.Filter = MDIForm1.s_string
    Me.tmp2_ado_oik_kiniseis.Recordset.Filter = MDIForm1.s_string
    MDIForm1.s_sort = "id"
    Me.ado_oikon_kiniseis.Recordset.Sort = MDIForm1.s_sort
    Me.tmp2_ado_oik_kiniseis.Recordset.Sort = MDIForm1.s_sort

    'ENHMERVSH TOY DATAGRID
    If Me.ado_oikon_kiniseis.Recordset.RecordCount >= 1 Then
        Call OikonomikesKiniseisRefresh
        Me.ado_oikon_kiniseis.Recordset.MoveLast
        Me.Command3.Enabled = True
        Me.Command4.Enabled = True
        Me.Command1.Enabled = True
        Me.Command6.Enabled = True
        Me.Command5.Enabled = True
        Me.bt_print.Enabled = True
    End If
    
    Me.Command2.Enabled = True
    Me.search.Enabled = False

End Sub

Private Sub co_py_Change()
    
    
    '�������� ��������������
    If Me.co_py.Text <> "" Then
    
        Me.dt_oikon_kiniseis.Visible = True
        
        id_s = ������ID_���_String("��������������", Me.co_py.Text)
        
        MDIForm1.s_string = "[����������������] = 1 AND [id_py] LIKE " & id_s
        Me.ado_oikon_kiniseis.Recordset.Filter = MDIForm1.s_string
        Me.tmp2_ado_oik_kiniseis.Recordset.Filter = MDIForm1.s_string
        MDIForm1.s_sort = "id"
        Me.ado_oikon_kiniseis.Recordset.Sort = MDIForm1.s_sort
        Me.tmp2_ado_oik_kiniseis.Recordset.Sort = MDIForm1.s_sort
            
        Call OikonomikesKiniseisRefresh
    
        If Me.ado_oikon_kiniseis.Recordset.RecordCount >= 1 Then
            Me.ado_oikon_kiniseis.Recordset.MoveLast
        Else 'KENO TO GRID, xwris eggrafes
            Me.Command4.Enabled = False ' UPDATE
        End If
    
        Me.search.Enabled = False 'SEARCH
        Me.Command3.Enabled = True 'SORTING
        Me.Command2.Enabled = True 'STORAGE A NEW
        Me.Command4.Enabled = True 'PROCESS AN OLD
        Me.Command1.Enabled = True '���� AKYRH ��� ��������
        Me.in_order_to_filter.Enabled = True
        Me.search.Enabled = False
        Me.cancel.Enabled = True
        Me.Command5.Enabled = True 'PRINTING A CURRENT
        Me.bt_print.Enabled = True
        Me.Command6.Enabled = True 'MAKE AN EXCEL FILE
        
    End If

End Sub

Private Sub Command1_Click()
    
    Dim id_par As Integer

    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.RecordCount >= 1 Then
    
        Me.tmp_ado_oik_kiniseis.Recordset.MoveFirst
        Me.tmp_ado_oik_kiniseis.Recordset.Find "[id] = " & Me.ado_oikon_kiniseis.Recordset.Fields(0).Value
        cur_row = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.AbsolutePosition
    
        '��������� ����������������
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Fields(3).Value = -1
    
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.UpdateBatch adAffectCurrent
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Requery
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Refresh
    
        'cur_row = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.AbsolutePosition
        oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Requery
        oikonomikes_kiniseis_management.ado_oikon_kiniseis.Refresh
        oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Filter = MDIForm1.s_string
        oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Sort = "[id]"
    
        Call OikonomikesKiniseisRefresh
    
        If Me.ado_oikon_kiniseis.Recordset.RecordCount >= 1 Then
            If cur_row - 1 = Me.ado_oikon_kiniseis.Recordset.RecordCount Then
                oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.MoveLast
            Else
                If cur_row <> 1 Then
                    oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.MoveFirst
                    oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Move cur_row - 1
                Else
                    oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.MoveFirst
                    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.EOF Then
                        Me.ado_oikon_kiniseis.Caption = "����������� " & 0 & " ��� " & 0
                    End If
                End If
            End If
        Else 'KENO TO GRID, xwris eggrafes
            Me.ado_oikon_kiniseis.Caption = "����������� " & 0 & " ��� " & 0
        End If
    End If
    
End Sub

Private Sub Command2_Click()

    ' ��� ������������ �� �������������� ����������� �������� ���� �� ���� ��� �/�
    Me.ado_proipologismos.Recordset.Filter = "[id_��������������] = " & ������ID_���_String("��������������", Me.co_py.Text)
    If CDate(Me.ado_proipologismos.Recordset.Fields(3).Value) <= CDate(DateValue(Now)) Then
        MsgBox "���������� �� �������� ����������� ������� �/� ��� ���� ����� - � ����� ������������!", vbCritical, "���������� - �������� ����������� ���������"
        Exit Sub
    End If


    MIA_oikonomiki_kinisi_management.Show
    
    MIA_oikonomiki_kinisi_management.Height = 9000
    MIA_oikonomiki_kinisi_management.Width = 12500
    
    '��������� ��� ������ ��� �������
    MIA_oikonomiki_kinisi_management.co_athlites.Clear
    MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.Filter = ""
    If MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.RecordCount >= 1 Then
        MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.MoveFirst
        For i = 0 To MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.RecordCount - 1
            If MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.Fields(1).Value <> "" Then
                MIA_oikonomiki_kinisi_management.co_athlites.AddItem (MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.Fields(1).Value)
            Else
                '��� ������ ���� ��� ���� ������� ��� ����� � ������� - ���� ����� �������!!!
                'MIA_oikonomiki_kinisi_management.co_athlites.AddItem ("")
                'MsgBox "Mpika edo me to id = " & MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.Fields(0).Value
            End If
            MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.MoveNext
        Next i
    End If
    '��������� ��� ������ ��� �����
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

 
    MIA_oikonomiki_kinisi_management.f0.Text = ����������������
    MIA_oikonomiki_kinisi_management.f3.Text = DateValue(Now)
    'Call Refresh_�����������������(MIA_oikonomiki_kinisi_management.opt_f1(0).Value, 1)
    Call Refresh_�����������������(True, 1)
    MIA_oikonomiki_kinisi_management.f9.Text = DateValue(Now)
    '��������� ��� FRAME �� �� �������� ��������
    MIA_oikonomiki_kinisi_management.opt_f1(0).Value = True
    '
    
    '��������� ������ �����
    If Not oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.EOF Then
        global_py = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(1).Value
    End If
    '��������� ������ �����
    If Not oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.EOF Then
        global_kk = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(21).Value
    End If
    '��������� ������ �����
    If Not oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.EOF Then
        global_tip_eks = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(23).Value
    End If
    '��������� ������ �����
    If MIA_oikonomiki_kinisi_management.f12.Text <> "" Then
        global_poso_xreosis = Val(MIA_oikonomiki_kinisi_management.f12.Text)
    End If
    '��������� ������ �����
    If Not oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.EOF Then
        global_tip_es = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(24).Value
    End If
    '��������� ������ �����
    If MIA_oikonomiki_kinisi_management.f15.Text <> "" Then
        global_poso_pistosis = Val(MIA_oikonomiki_kinisi_management.f15.Text)
    End If
    
    MIA_oikonomiki_kinisi_management.raw_co_tipoi_parastatikwn.Locked = False
    MIA_oikonomiki_kinisi_management.Frame1.Enabled = True
    MIA_oikonomiki_kinisi_management.Storage.Enabled = True
    MIA_oikonomiki_kinisi_management.Clean.Enabled = True
    MIA_oikonomiki_kinisi_management.Command5.Enabled = False
    
End Sub

Private Sub Command3_Click()

    s_sort = ""
    If Me.ado_oikon_kiniseis.Recordset.RecordCount >= 1 Then
        If Me.dt_oikon_kiniseis.Col >= 0 Then
            Me.ado_oikon_kiniseis.Recordset.Sort = "[" & Trim(Me.ado_oikon_kiniseis.Recordset.Fields(Me.dt_oikon_kiniseis.Col).Name) & "]"
            s_sort = "[" & Trim(Poseidon_DB.rs��������_�����������_��������.Fields(Me.dt_oikon_kiniseis.Col).Name) & "]"
        Else
            Me.ado_oikon_kiniseis.Recordset.Sort = "[" & Trim(Me.ado_oikon_kiniseis.Recordset.Fields(defined_col).Name) & "]"
            s_sort = "[" & Trim(Me.ado_oikon_kiniseis.Recordset.Fields(defined_col).Name) & "]"
        End If
    End If
    MDIForm1.s_sort = s_sort
    
End Sub

Private Sub Command4_Click()

    On Error GoTo Command4_Click_l1

    MIA_oikonomiki_kinisi_management.Show
    
    MIA_oikonomiki_kinisi_management.Height = 8000
    MIA_oikonomiki_kinisi_management.Width = 12500
    
    Call Refresh_�����������������(MIA_oikonomiki_kinisi_management.opt_f1(0).Value, 1)
    
    
    '
    MIA_oikonomiki_kinisi_management.co_athlites.Clear
    MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.Filter = ""
    If MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.RecordCount >= 1 Then
        MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.MoveFirst
        For i = 0 To MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.RecordCount - 1
            MIA_oikonomiki_kinisi_management.co_athlites.AddItem (MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.Fields(1).Value)
            MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.MoveNext
        Next i
    End If

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
 
    'id
    MIA_oikonomiki_kinisi_management.f0.Text = Me.ado_oikon_kiniseis.Recordset.Fields(0).Value
    '��������������
    'MIA_oikonomiki_kinisi_management.co_py.Text = ����������("��������������", 1, oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(1).Value)
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(2).Value <> "" Then
        MIA_oikonomiki_kinisi_management.co_py.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(2).Value
        'MIA_oikonomiki_kinisi_management.ado_py.Recordset.MoveFirst
        'MIA_oikonomiki_kinisi_management.ado_py.Recordset.Find "[���������] like '" & Trim(oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(2).Value) & "'"
        Call Refresh_�����������(2)
    Else
        MIA_oikonomiki_kinisi_management.co_py.Text = ""
    End If
        '��������� ������ �����
        global_py = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(1).Value
    
    '����� �������
    If Me.ado_oikon_kiniseis.Recordset.Fields(3).Value = "�������" Then
        MIA_oikonomiki_kinisi_management.opt_f1(0).Value = True
        MIA_oikonomiki_kinisi_management.opt_f1(1).Value = False
    Else '����� ������
        MIA_oikonomiki_kinisi_management.opt_f1(0).Value = False
        MIA_oikonomiki_kinisi_management.opt_f1(1).Value = True
    End If
    
    '��������� �������
    If Me.ado_oikon_kiniseis.Recordset.Fields(4).Value = "�����" Then
        MIA_oikonomiki_kinisi_management.opt_f2(0).Value = True
        MIA_oikonomiki_kinisi_management.opt_f2(1).Value = False
        MIA_oikonomiki_kinisi_management.opt_f2(2).Value = False
    Else
        If Me.ado_oikon_kiniseis.Recordset.Fields(4).Value = "����������" Then
            MIA_oikonomiki_kinisi_management.opt_f2(0).Value = False
            MIA_oikonomiki_kinisi_management.opt_f2(1).Value = True
            MIA_oikonomiki_kinisi_management.opt_f2(2).Value = False
        Else
            MIA_oikonomiki_kinisi_management.opt_f2(0).Value = False
            MIA_oikonomiki_kinisi_management.opt_f2(1).Value = False
            MIA_oikonomiki_kinisi_management.opt_f2(2).Value = True
        End If
    End If
        '��������� ������ �����
        global_kk = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(21).Value
    
    '���������� �������
    If Me.ado_oikon_kiniseis.Recordset.Fields(5).Value <> "" Then
        MIA_oikonomiki_kinisi_management.f3.Text = Me.ado_oikon_kiniseis.Recordset.Fields(5).Value
    End If
    '�������
    'MIA_oikonomiki_kinisi_management.co_athlites.Text = ����������("��������������������", 1, oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(6).Value)
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(6).Value <> "" Then
        MIA_oikonomiki_kinisi_management.co_athlites.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(6).Value
    Else
        MIA_oikonomiki_kinisi_management.co_athlites.Text = ""
    End If
    '�����
    'MIA_oikonomiki_kinisi_management.co_meli.Text = ����������("������������������", 1, oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(7).Value)
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(7).Value <> "" Then
        MIA_oikonomiki_kinisi_management.co_meli.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(7).Value
    Else
        MIA_oikonomiki_kinisi_management.co_meli.Text = ""
    End If
    '����������
    'MIA_oikonomiki_kinisi_management.co_organismoi.Text = ����������("�����������", 1, oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(8).Value)
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(8).Value <> "" Then
        MIA_oikonomiki_kinisi_management.co_organismoi.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(8).Value
    Else
        MIA_oikonomiki_kinisi_management.co_organismoi.Text = ""
    End If
    '����� ������������
    'MIA_oikonomiki_kinisi_management.raw_co_tipoi_parastatikwn.Text = ����������("�����������������������������", 3, oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(9).Value)
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(9).Value <> "" Then
        MIA_oikonomiki_kinisi_management.raw_co_tipoi_parastatikwn.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(9).Value
    Else
        MIA_oikonomiki_kinisi_management.raw_co_tipoi_parastatikwn.Text = ""
    End If
    '������� ������������
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(10).Value <> "" Then
        MIA_oikonomiki_kinisi_management.f8.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(10).Value
    Else
        MIA_oikonomiki_kinisi_management.f8.Text = ""
    End If
    '���������� ������������
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(11).Value <> "" Then
        MIA_oikonomiki_kinisi_management.f9.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(11).Value
    Else
        MIA_oikonomiki_kinisi_management.f9.Text = ""
    End If
    '����� �� ������
    'MIA_oikonomiki_kinisi_management.co_raw_tipoi_eksodwn.Text = ����������("�����������", 1, oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(12).Value)
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(12).Value <> "" Then
        MIA_oikonomiki_kinisi_management.co_raw_tipoi_eksodwn.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(12).Value
    Else
        MIA_oikonomiki_kinisi_management.co_raw_tipoi_eksodwn.Text = ""
    End If
        '��������� ������ �����
        global_tip_eks = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(23).Value
    
    '���� �������
    'If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(12).Value <> "" And oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(13).Value <> 0 Then
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(13).Value <> 0 Then
        'MIA_oikonomiki_kinisi_management.f12.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(13).Value & " �"
        MIA_oikonomiki_kinisi_management.f12.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(13).Value
    Else
        MIA_oikonomiki_kinisi_management.f12.Text = ""
    End If
        '��������� ������ �����
        global_poso_xreosis = Val(MIA_oikonomiki_kinisi_management.f12.Text)
        
    '����� �� ������
    'Call Refresh_�����������(1)
    'MIA_oikonomiki_kinisi_management.co_raw_tipoi_esodwn.Text = ����������("�����������", 1, oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(14).Value)
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(14).Value <> "" Then
        MIA_oikonomiki_kinisi_management.co_raw_tipoi_esodwn.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(14).Value
    Else
        MIA_oikonomiki_kinisi_management.co_raw_tipoi_esodwn.Text = ""
    End If
        '��������� ������ �����
        global_tip_es = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(24).Value
    
    '����� ��������� ������
    For i = 0 To MIA_oikonomiki_kinisi_management.lst_mn.ListCount - 1
        If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(32 + i).Value = True Then
            MIA_oikonomiki_kinisi_management.lst_mn.Selected(i) = True
            global_minas_sindromis(i) = True
            MIA_oikonomiki_kinisi_management.Label3(2).Enabled = True
            MIA_oikonomiki_kinisi_management.lst_mn.Enabled = True
            MIA_oikonomiki_kinisi_management.bt_en_a.Enabled = True
            MIA_oikonomiki_kinisi_management.bt_kath_a.Enabled = True
        Else
            MIA_oikonomiki_kinisi_management.lst_mn.Selected(i) = False
            global_minas_sindromis(i) = False
        End If
    Next i
    
    '���� ��������
    'If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(14).Value <> "" And oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(15).Value <> 0 Then
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(15).Value <> 0 Then
        'MIA_oikonomiki_kinisi_management.f15.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(15).Value & " �"
        MIA_oikonomiki_kinisi_management.f15.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(15).Value
    Else
        MIA_oikonomiki_kinisi_management.f15.Text = ""
    End If
        '��������� ������ �����
        global_poso_pistosis = Val(MIA_oikonomiki_kinisi_management.f15.Text)
        
    '����������
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(16).Value <> "" Then
        MIA_oikonomiki_kinisi_management.f16.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(16).Value
    Else
        MIA_oikonomiki_kinisi_management.f16.Text = ""
    End If
    
    '����������2
    If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields("����������2").Value <> "" Then
        MIA_oikonomiki_kinisi_management.txt_aitiol2.Text = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(����������2).Value
    Else
        MIA_oikonomiki_kinisi_management.txt_aitiol2.Text = ""
    End If
    '
    MIA_oikonomiki_kinisi_management.raw_co_tipoi_parastatikwn.Locked = True
    MIA_oikonomiki_kinisi_management.Frame1.Enabled = False
    '
    MIA_oikonomiki_kinisi_management.update.Enabled = True
    MIA_oikonomiki_kinisi_management.cancel.Enabled = True
    'MIA_oikonomiki_kinisi_management.Command5.Enabled = False
    
Command4_Click_l1:
    
End Sub

Public Sub Command5_Click()

    On Error GoTo Command5_Click_l1

    If Not oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.EOF Then
        Rep_ektiposi_1_oikonomikis_kinisis.Hide
        
        Dim tmp_string As String
        'MsgBox oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(0).Value
        tmp_string = "[id] LIKE '" & str(oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(0).Value) & "'"
        Poseidon_DB.rs��������_�����������_��������.Filter = tmp_string
        If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(3).Value = "�������" Then
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label1").Caption = "����� ��� ��� / ���"
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label4").Caption = "� �����"
        Else
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label1").Caption = "������� ��� / ���"
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label4").Caption = "� �������"
        End If
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label7").Caption = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(9).Value & "        No " & oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(10).Value
        If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(18).Value >= 1 Then
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label10").Caption = "�� ��������� "
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label46").Caption = "�� ��������� "
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label26").Caption = "�� ��������� "
            If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(7).Value <> "" Then
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label5").Caption = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(7).Value & ","
            Else
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label5").Caption = "______________ ,"
            End If
            If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(26).Value <> "" Then
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label11").Caption = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(26).Value
            Else
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label11").Caption = ""
            End If
            If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(27).Value <> "" Then
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label13").Caption = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(27).Value & ","
            Else
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label13").Caption = "______________ ,"
            End If
        Else
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label5").Caption = ""
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label11").Caption = ""
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label13").Caption = "______________ ,"
        End If
        '����������� - ����������
        If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(19).Value >= 1 Then
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label10").Caption = "�� �.�.�. "
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label46").Caption = "�� �.�.�. "
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label26").Caption = "�� �.�.�. "
            If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(8).Value <> "" Then
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label5").Caption = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(8).Value & ","
            Else
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label5").Caption = "______________ ,"
            End If
            If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(28).Value <> "" Then
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label11").Caption = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(28).Value
            Else
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label11").Caption = ""
            End If
            If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(29).Value <> "" Then
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label13").Caption = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(29).Value & ","
            Else
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label13").Caption = "______________ ,"
            End If
        End If
        '�������, ������ ������� �����
        If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(18).Value >= 1 Then
            If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields("����������2").Value <> "" Then
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label8").Caption = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields("����������2").Value & ","
            ElseIf oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(6).Value <> "" Then
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label8").Caption = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(6).Value & ","
            Else
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label8").Caption = "______________ ,"
            End If
        Else '�������, ������ ��� ������� �����
            If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields("����������2").Value <> "" Then
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label8").Caption = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields("����������2").Value & "," & " (�����-��/���� ��� ���������),"
            ElseIf oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(6).Value <> "" Then
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label8").Caption = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(6).Value & "," & " (�����-��/���� ��� ���������),"
            Else
                Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label8").Caption = "______________ ,"
            End If
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label6").Visible = False
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label44").Visible = False
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label24").Visible = False
        End If
        '����
        Dim ts As String
        ts = ""
        If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(3).Value = "�������" Then
            If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(15).Value <> "" Then
                ts = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(15).Value
            End If
        Else
            If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(13).Value <> "" Then
                ts = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(13).Value
            End If
        End If
        If ts <> "" Then
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label15").Caption = ts & "."
        Else
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label15").Caption = "______________ ."
        End If
        '����������
        ts = ""
        If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(3).Value = "�������" Then
            If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(14).Value <> "" Then
                ts = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(14).Value
            End If
        Else
            If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(12).Value <> "" Then
                ts = oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(12).Value
            End If
        End If
        If ts <> "" Then
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label3").Caption = "����������: " & ts
        Else
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label3").Caption = "����������: "
        End If
        If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(16).Value <> "" Then
            Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label3").Caption = "����������: " & ts & " (" & oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(16).Value & ")."
        End If
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label40").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label1").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label20").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label1").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label42").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label4").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label22").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label4").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label32").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label7").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label9").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label7").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label43").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label5").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label23").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label5").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label45").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label8").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label25").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label8").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label47").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label11").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label27").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label11").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label49").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label13").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label29").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label13").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label51").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label15").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label31").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label15").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label41").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label3").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label21").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label3").Caption
    
        If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(11).Value <> "" Then
            Rep_ektiposi_1_oikonomikis_kinisis.Sections(2).Controls("Label2").Caption = "����������: " & oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(11).Value
        Else
            Rep_ektiposi_1_oikonomikis_kinisis.Sections(2).Controls("Label2").Caption = "����������: "
        End If
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label39").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label2").Caption
        Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section1").Controls("Label19").Caption = Rep_ektiposi_1_oikonomikis_kinisis.Sections("Section2").Controls("Label2").Caption
        
        'rep.Orientation = rptOrientLandscape
        'Rep_ektiposi_1_oikonomikis_kinisis.Orientation = rptOrientPortait
        Rep_ektiposi_1_oikonomikis_kinisis.Show
        Poseidon_DB.rs��������_�����������_��������.Close
    '______________________________________________________________________________________________________
    Else '��� �������� �������� - ���� datagrid
        MsgBox "��� ���� �������� ���������� ������ ���� ��������.", , "������ ������"
    End If
    
Command5_Click_l1:
    
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
   If Me.ado_oikon_kiniseis.Recordset.RecordCount >= 1 Then
        Me.ado_oikon_kiniseis.Recordset.MoveFirst
        sData = "��������������" & vbTab & "����� �������" & vbTab & "��������� �������" & vbTab & "���������� �������" & vbTab & "�������" & vbTab & "�����" & vbTab & "���������� - �����������" & vbTab & "����� ������������" & vbTab & "������� ������������" & vbTab & "���������� ������������" & vbTab & "����� ������" & vbTab & "���� �������" & vbTab & "����� ������" & vbTab & "���� ��������" & vbTab & "����������" & vbCr
        For i = 0 To Me.ado_oikon_kiniseis.Recordset.RecordCount - 1
            sData = sData & Me.ado_oikon_kiniseis.Recordset.Fields(2) & vbTab & Me.ado_oikon_kiniseis.Recordset.Fields(3) & vbTab & Me.ado_oikon_kiniseis.Recordset.Fields(4) & vbTab & Me.ado_oikon_kiniseis.Recordset.Fields(5) & vbTab & Me.ado_oikon_kiniseis.Recordset.Fields(6) & vbTab & Me.ado_oikon_kiniseis.Recordset.Fields(7) & vbTab & Me.ado_oikon_kiniseis.Recordset.Fields(8) & vbTab & Me.ado_oikon_kiniseis.Recordset.Fields(9) & vbTab & Me.ado_oikon_kiniseis.Recordset.Fields(10) & vbTab & Me.ado_oikon_kiniseis.Recordset.Fields(11) & vbTab & Me.ado_oikon_kiniseis.Recordset.Fields(12) & vbTab & Me.ado_oikon_kiniseis.Recordset.Fields(13) & vbTab & Me.ado_oikon_kiniseis.Recordset.Fields(14) & vbTab & Me.ado_oikon_kiniseis.Recordset.Fields(15) & vbTab & Me.ado_oikon_kiniseis.Recordset.Fields(16) & vbCr
            Me.ado_oikon_kiniseis.Recordset.MoveNext
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

Private Sub Command7_Click()

    Me.lst_mn.Visible = True
    Me.lst_mn.Enabled = True
    'Me.lst_mn.ZOrder (1)
    
End Sub

Private Sub Command9_Click()

    Unload Me
    
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    
    defined_col = ColIndex
    
End Sub

Private Sub delete_Click()
        
    Dim ms As String
    
    ms = MsgBox("����� ��������; (��� � ���)", vbYesNo, "�������� ���������")
    If ms = 6 Then
        If Not Me.ado_oikon_kiniseis.Recordset.EOF Then
            With Me.ado_oikon_kiniseis.Recordset
                .Delete
                .MoveNext
                If .EOF And .RecordCount <> 0 Then
                    .MoveLast
                Else
                    c_r = 0
                End If
            End With
            oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Requery
            oikonomikes_kiniseis_management.ado_oikon_kiniseis.Refresh
        End If
    Else
        '
    End If

End Sub

Private Sub dt_oikon_kiniseis_HeadClick(ByVal ColIndex As Integer)
    
    defined_col = ColIndex
    
End Sub

Private Sub Form_GotFocus()

    Me.Refresh

End Sub

Private Sub Form_Load()
        
    
    Me.Top = 20
    Me.Left = 10
    Me.Width = 19200
    Me.Height = 12000
    
    Me.s_xr.Text = FormatCurrency(0, 2, vbTrue, , vbTrue)
    Me.s_pis.Text = FormatCurrency(0, 2, vbTrue, , vbTrue)
    Me.ado_oikon_kiniseis.Caption = "����������� " & 0 & " ��� " & 0
    
    Me.co_tipos_kinisis.AddItem "�������"
    Me.co_tipos_kinisis.AddItem "������"
    Me.co_tipos_kinisis.AddItem "����"
    Me.co_katastasi_kinisis.AddItem "�����"
    Me.co_katastasi_kinisis.AddItem "����������"
    Me.co_katastasi_kinisis.AddItem "������"
    Me.co_katastasi_kinisis.AddItem "����"
    
    
    
    '������� ���� ��������� �/�
    If Not ado_proipologismos.Recordset.EOF Then
        ado_proipologismos.Recordset.MoveLast
        id_s = ado_proipologismos.Recordset.Fields(0).Value
        MDIForm1.s_string = "[����������������] = 1 AND [id_py] LIKE " & id_s
        Me.ado_oikon_kiniseis.Recordset.Filter = MDIForm1.s_string
        Me.tmp2_ado_oik_kiniseis.Recordset.Filter = MDIForm1.s_string
        MDIForm1.s_sort = "id"
        Me.ado_oikon_kiniseis.Recordset.Sort = MDIForm1.s_sort
        Me.tmp2_ado_oik_kiniseis.Recordset.Sort = MDIForm1.s_sort
    
        Call OikonomikesKiniseisRefresh
        dt_oikon_kiniseis.Visible = True
        'co_py.Index = ado_proipologismos.Recordset.RecordCount - 1
        co_py.Text = ado_proipologismos.Recordset.Fields(1).Value
    
   
        If Me.ado_oikon_kiniseis.Recordset.RecordCount >= 1 Then
            Me.ado_oikon_kiniseis.Recordset.MoveLast
        Else 'KENO TO GRID, xwris eggrafes
            Me.Command4.Enabled = False ' UPDATE
        End If
   
        Me.search.Enabled = False 'SEARCH
        Me.Command3.Enabled = True 'SORTING
        Me.Command2.Enabled = True 'STORAGE A NEW
        Me.Command4.Enabled = True 'PROCESS AN OLD
        Me.Command1.Enabled = True '���� AKYRH ��� ��������
        Me.in_order_to_filter.Enabled = True
        Me.search.Enabled = False
        Me.cancel.Enabled = True
        Me.Command5.Enabled = True 'PRINTING A CURRENT
        Me.bt_print.Enabled = True
        Me.Command6.Enabled = True 'MAKE AN EXCEL FILE
   
    End If

End Sub

Private Sub in_order_to_filter_Click()


    Me.co_py.Locked = True

    For i = 0 To 0 '�������
        Me.s_c(i).Visible = True
        Me.s_c(i).Text = ""
    Next i
    Me.co_py.Visible = True
    'Me.co_py.Text = ""
    Me.co_tipos_kinisis.Visible = True
    Me.co_tipos_kinisis.Text = ""
    Me.co_katastasi_kinisis.Visible = True
    Me.co_katastasi_kinisis.Text = ""
    For i = 1 To 1 '���������� �������
        Me.s_c(i).Visible = True
        Me.s_c(i).Text = ""
    Next i
    For i = 2 To 2 '������� ������������
        Me.s_c(i).Visible = True
        Me.s_c(i).Text = ""
    Next i
    For i = 3 To 3 '���������� ������������
        Me.s_c(i).Visible = True
        Me.s_c(i).Text = ""
    Next i
    For i = 4 To 4 '���� �������
        Me.s_c(i).Visible = True
        Me.s_c(i).Text = ""
    Next i
    For i = 5 To 5 '���� ��������
        Me.s_c(i).Visible = True
        Me.s_c(i).Text = ""
    Next i
    For i = 6 To 6 '����������
        Me.s_c(i).Visible = True
        Me.s_c(i).Text = ""
    Next i
    Me.co_athlites.Visible = True
    Me.co_athlites.Text = ""
    Me.co_meli.Visible = True
    Me.co_meli.Text = ""
    Me.co_organismoi.Visible = True
    Me.co_organismoi.Text = ""
    Me.co_tipoi_parastatikwn.Visible = True
    Me.co_tipoi_parastatikwn.Text = ""
    Me.Command7.Visible = True
    Me.co_tipoi_eksodwn.Visible = True
    Me.co_tipoi_eksodwn.Text = ""
    If ado_tipoi_eksodwn.Recordset.RecordCount >= 1 Then
        Me.ado_tipoi_eksodwn.Recordset.Sort = "[���������]"
    End If
    Me.co_tipoi_esodwn.Visible = True
    Me.co_tipoi_esodwn.Text = ""
    If ado_tipoi_esodwn.Recordset.RecordCount >= 1 Then
        Me.ado_tipoi_esodwn.Recordset.Sort = "[���������]"
    End If

    Me.search.Enabled = True
    'Me.Command2.Enabled = False
    'Me.Command4.Enabled = False
    
End Sub

Private Sub search_Click()

    Dim id_s As Integer
    
    s_string = ""
    
    MDIForm1.s_string = "[����������������] = 1"
    MDIForm1.s_sort = "id"
    '
    '�������� �������
    If Trim(Me.s_c(0).Text) <> "" Then
        s_string = "[id] LIKE " & Trim(Me.s_c(0).Text)
    End If
    '
    '�������� ��������������
    If Me.co_py.Text <> "" Then
        id_s = ������ID_���_String("��������������", Me.co_py.Text)
        If s_string <> "" Then
            s_string = s_string & " AND [id_py] LIKE " & id_s
        Else
            s_string = "[id_py] LIKE " & id_s
        End If
    End If
    '
    '�������� ����� �������
    If Me.co_tipos_kinisis <> "" Then
        If Me.co_tipos_kinisis.ListIndex = 0 Then
            If s_string <> "" Then
                s_string = s_string & " AND [������������] LIKE 1"
            Else
                s_string = "[������������] LIKE 1"
            End If
        Else
            If Me.co_tipos_kinisis.ListIndex = 1 Then
                If s_string <> "" Then
                    s_string = s_string & " AND [������������] LIKE 0"
                Else
                    s_string = "[������������] LIKE 0"
                End If
            Else
                If Me.co_tipos_kinisis.ListIndex = 2 Then
                    'If s_string <> "" Then
                    '    s_string = s_string & " AND [������������] LIKE 0"
                    'Else
                    '    s_string = "[������������] LIKE 0"
                    'End If
                End If
            End If
        End If
    End If
    '
    '�������� ���������� �������
    If Me.co_katastasi_kinisis <> "" Then
        If Me.co_katastasi_kinisis.ListIndex = 0 Then
            MDIForm1.s_string = ""
            If s_string <> "" Then
                s_string = s_string & " AND [����������������] LIKE -1"
            Else
                s_string = "[����������������] LIKE -1"
            End If
        Else
            If Me.co_katastasi_kinisis.ListIndex = 1 Then
                MDIForm1.s_string = ""
                If s_string <> "" Then
                    s_string = s_string & " AND [����������������] LIKE 0"
                Else
                    s_string = "[����������������] LIKE 0"
                End If
            Else
                If Me.co_katastasi_kinisis.ListIndex = 2 Then
                    If s_string <> "" Then
                        s_string = s_string & " AND [����������������] LIKE 1"
                    Else
                        s_string = "[����������������] LIKE 1"
                    End If
                Else
                    If Me.co_katastasi_kinisis.ListIndex = 3 Then
                        MDIForm1.s_string = ""
                    End If
                End If
            End If
        End If
    End If
    '
    '�������� ����������� �������
    's_string = GetDateFilter(s_string, Trim(Me.s_c(1).Text), "�����������������")
    If Trim(Me.s_c(1).Text) <> "" And IsDate(Trim(Me.s_c(1).Text)) = True Then
        If s_string <> "" Then
            s_string = s_string & " AND [�����������������] LIKE '" & Trim(Me.s_c(1).Text) & "'"
        Else
            s_string = "[�����������������] LIKE '" & Trim(Me.s_c(1).Text) & "'"
        End If
    Else
        If Trim(Me.s_c(1).Text) <> "" And Trim(Me.s_c(1).Text) >= 1 And Trim(Me.s_c(1).Text) <= 12 Then
            If s_string <> "" Then
                s_string = s_string & " AND [���] LIKE " & Trim(Me.s_c(1).Text)
            Else
                s_string = "[���] LIKE " & Trim(Me.s_c(1).Text)
            End If
        Else
            Me.s_c(1).Text = ""
        End If
    End If
    '
    '�������� ������
    If Me.co_athlites.Text <> "" Then
        id_s = ������ID_���_String("��������������������", Me.co_athlites.Text)
        If s_string <> "" Then
            s_string = s_string & " AND [id_������] LIKE " & id_s
        Else
            s_string = "[id_������] LIKE " & id_s
        End If
    End If
    '
    '�������� ������
    If Me.co_meli.Text <> "" Then
        id_s = ������ID_���_String("������������������", Me.co_meli.Text)
        If s_string <> "" Then
            s_string = s_string & " AND [id_������] LIKE " & id_s
        Else
            s_string = "[id_������] LIKE " & id_s
        End If
    End If
    '
    '�������� ����������
    If Me.co_organismoi.Text <> "" Then
        id_s = ������ID_���_String("�����������������������", Me.co_organismoi.Text)
        If s_string <> "" Then
            s_string = s_string & " AND [id_����������] LIKE " & id_s
        Else
            s_string = "[id_����������] LIKE " & id_s
        End If
    End If
    '
    '�������� ����� ������������
    If Me.co_tipoi_parastatikwn.Text <> "" Then
        id_s = ������ID_���_String("������������������������������", Me.co_tipoi_parastatikwn.Text)
        If s_string <> "" Then
            s_string = s_string & " AND [id_�����������������] LIKE " & id_s
        Else
            s_string = "[id_�����������������] LIKE " & id_s
        End If
    End If
    '
    '�������� ���A
    Dim str_m As String
    If Me.lst_mn.ListIndex >= 0 And Me.lst_mn.Text <> "" Then
        str_m = "�����" & Me.lst_mn.ListIndex + 1
        If s_string <> "" Then
            s_string = s_string & " AND [" & str_m & "] = True"
        Else
            s_string = "[" & str_m & "] = True"
        End If
    End If
    '
    '�������� ������� ������������
    If Trim(Me.s_c(2).Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [������� ������������] LIKE '" & Trim(Me.s_c(2).Text) & "'"
        Else
            s_string = "[������� ������������] LIKE '" & Trim(Me.s_c(2).Text) & "'"
        End If
    End If
    '
    '�������� ����������� ������������
    's_string = GetDateFilter(s_string, Trim(Me.s_c(3).Text), "����������������������")
    If Trim(Me.s_c(3).Text) <> "" And IsDate(Trim(Me.s_c(3).Text)) = True Then
        If s_string <> "" Then
            s_string = s_string & " AND [����������������������] LIKE '" & Trim(Me.s_c(3).Text) & "'"
        Else
            s_string = "[����������������������] LIKE '" & Trim(Me.s_c(3).Text) & "'"
        End If
    Else
        If Trim(Me.s_c(3).Text) <> "" And Trim(Me.s_c(3).Text) >= 1 And Trim(Me.s_c(3).Text) <= 12 Then
            If s_string <> "" Then
                s_string = s_string & " AND [��P] LIKE " & Trim(Me.s_c(3).Text)
            Else
                s_string = "[��P] LIKE " & Trim(Me.s_c(3).Text)
            End If
        Else
            Me.s_c(3).Text = ""
        End If
    End If
    '
    '�������� �������������� ������
    If Me.co_tipoi_eksodwn.Text <> "" Then
        id_s = ������ID_���_String("�����������", Me.co_tipoi_eksodwn.Text)
        If s_string <> "" Then
            s_string = s_string & " AND [id_�����������������] LIKE " & id_s
        Else
            s_string = "[id_�����������������] LIKE " & id_s
        End If
    End If
    '
    '�������� ����� �������
    If Trim(Me.s_c(4).Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [�����������] LIKE '" & Trim(Me.s_c(4).Text) & "'"
            Else
                s_string = "[�����������] LIKE '" & Trim(Me.s_c(4).Text) & "'"
            End If
    End If
    '
    '�������� �������������� ������
    If Me.co_tipoi_esodwn.Text <> "" Then
        id_s = ������ID_���_String("�����������", Me.co_tipoi_esodwn.Text)
        If s_string <> "" Then
            s_string = s_string & " AND [id_�����������������] LIKE " & id_s
        Else
            s_string = "[id_�����������������] LIKE " & id_s
        End If
    End If
    '
    '�������� ����� ��������
    If Trim(Me.s_c(5).Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [������������] LIKE '" & Trim(Me.s_c(5).Text) & "'"
            Else
                s_string = "[������������] LIKE '" & Trim(Me.s_c(5).Text) & "'"
            End If
    End If
    '
    '�������� �����������
    If Trim(Me.s_c(6).Text) <> "" Then
        If s_string <> "" Then
            s_string = s_string & " AND [����������] LIKE *" & Me.s_c(6).Text & "*"
            Else
                s_string = "[����������] LIKE *" & Me.s_c(6).Text & "*"
            End If
    End If
    '
    '�������� FILTER ����� ��� s_string
    If MDIForm1.s_string <> "" Then
        If s_string <> "" Then
            MDIForm1.s_string = MDIForm1.s_string & " AND " & s_string
        End If
    Else
        MDIForm1.s_string = s_string
    End If
    Me.ado_oikon_kiniseis.Recordset.Filter = MDIForm1.s_string
    Me.tmp2_ado_oik_kiniseis.Recordset.Filter = MDIForm1.s_string
    '
    'ENHMERVSH TOY DATAGRID
    If Me.ado_oikon_kiniseis.Recordset.RecordCount >= 1 Then
        Call OikonomikesKiniseisRefresh
        Me.ado_oikon_kiniseis.Recordset.MoveLast
        Command4.Enabled = True
    Else 'KENO TO GRID, xwris eggrafes
        Me.ado_oikon_kiniseis.Caption = "����������� " & 0 & " ��� " & 0
        Me.s_xr.Text = FormatCurrency(0, 2, vbTrue, , vbTrue)
        Me.s_pis.Text = FormatCurrency(0, 2, vbTrue, , vbTrue)
        Me.Command4.Enabled = False ' UPDATE
        Me.Command3.Enabled = False 'SORT
        Me.Command1.Enabled = False 'MAKE AKYRH thn TREXOYSA
        Me.Command6.Enabled = False 'MAKE AN EXCEL FILE
        Me.bt_print.Enabled = False 'PRINT ALL
    End If
 
End Sub
