VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form ektiposi_tmimatwn_ana_ae 
   Caption         =   "Εκτύπωση Τμημάτων ανά Αθλητικό Έτος"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3570
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   3570
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Caption         =   "Εκτύπωση"
      Height          =   615
      Left            =   1680
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "Form2.frx":4A05
      Height          =   3570
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6297
      _Version        =   393216
      ListField       =   "Περιγραφή"
   End
   Begin MSAdodcLib.Adodc ado_ae 
      Height          =   375
      Left            =   2160
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "Αθλητικά_Έτη"
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
End
Attribute VB_Name = "ektiposi_tmimatwn_ana_ae"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    If Not Me.ado_ae.Recordset.EOF And Me.ado_ae.Recordset.Fields("id_αθλητικού_έτους").Value <> "" Then
        s_string = "[id_ΑθλητικούΈτους] = " & Me.ado_ae.Recordset.Fields("id_αθλητικού_έτους").Value
    Else
        s_string = ""
    End If
    Rep_ΚαρτέλαΤμήματοςΑνάΑΕ.Show

End Sub

Private Sub DataList1_Click()

    If Me.ado_ae.Recordset.RecordCount >= 1 Then
        If Me.DataList1.Text <> "" Then
            'Me.ado_ae.Recordset.MoveFirst
            'Me.ado_ae.Recordset.Find "[Περιγραφή] LIKE *'" & Trim(Me.DataList1.Text) & "*'"
            'Me.ado_ae.Recordset.Find "[id_αθλητικού_έτους] LIKE *'" & 2 & "*'"
            Me.ado_ae.Recordset.Move Me.DataList1.SelectedItem - Me.ado_ae.Recordset.AbsolutePosition
        End If
    End If

End Sub

Private Sub Form_Load()

    Me.Width = 5000
    Me.Height = 7000
    
    If Me.ado_ae.Recordset.RecordCount >= 1 Then
        Me.ado_ae.Recordset.MoveLast
        Me.DataList1.Text = Me.ado_ae.Recordset.Fields("Περιγραφή").Value
    End If
    
End Sub
