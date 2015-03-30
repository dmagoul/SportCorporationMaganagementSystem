VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form doi_management 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Διαχείριση ΔΟΥ"
   ClientHeight    =   9615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   6405
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Κλείσιμο"
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
      Left            =   5040
      Picture         =   "doi_management.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ακύρωση"
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
      Left            =   3840
      Picture         =   "doi_management.frx":5A78
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ταξινόμηση"
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
      Picture         =   "doi_management.frx":A995
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Προσθήκη"
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
      Picture         =   "doi_management.frx":F69C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Διαγραφή"
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
      Left            =   2640
      Picture         =   "doi_management.frx":1442B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8760
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   8280
      Width           =   6015
      _ExtentX        =   10610
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
      RecordSource    =   "ΔΟΥ"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "doi_management.frx":1922F
      Height          =   8175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   14420
      _Version        =   393216
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   18
      AllowDelete     =   -1  'True
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
Attribute VB_Name = "doi_management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim defined_col, c_r As Integer
Public rs As ADODB.Recordset

Private Sub Adodc1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

    MsgBox "An error occured: Description=" & Destription & "Error Number=" & ErrorNumber
    ErrorNumber = 0

End Sub

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    If adReason <> 10 Then
        If Not rs.EOF Then
            If rs.AbsolutePosition > 0 Then
                Me.Adodc1.Caption = "Εγγραφή " & rs.AbsolutePosition & " από " & rs.RecordCount
            Else
                Me.Adodc1.Caption = "Εγγραφή " & 0 & " από " & rs.RecordCount
            End If
        End If
    End If
    
End Sub

Private Sub Command1_Click()
    
    'On Error Resume Next
    If Not Me.Adodc1.Recordset.EOF Then
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
    'Err.Clear

End Sub

Private Sub Command2_Click()

    Dim id As Integer
    
    Me.Adodc1.Recordset.Sort = "[" & Trim(Me.Adodc1.Recordset.Fields(0).Name) & "]"
    
    If Not Me.Adodc1.Recordset.EOF Then
        Me.Adodc1.Recordset.MoveLast
        id = Me.Adodc1.Recordset![id_ΔΟΥ]
    End If
    
    Me.Adodc1.Recordset.AddNew
    Me.Adodc1.Recordset.Fields(0) = id + 1
    Me.Adodc1.Recordset.UpdateBatch adAffectCurrent
    Me.Adodc1.Recordset.Sort = "id_ΔΟΥ"
    Me.Adodc1.Recordset.MoveLast
    Me.DataGrid1.Col = 1
    Me.DataGrid1.SetFocus
    
End Sub

Private Sub Command3_Click()

    If Me.DataGrid1.Col >= 0 Then
        Me.Adodc1.Recordset.Sort = "[" & Trim(Me.Adodc1.Recordset.Fields(Me.DataGrid1.Col).Name) & "]"
    Else
        Me.Adodc1.Recordset.Sort = "[" & Trim(Me.Adodc1.Recordset.Fields(defined_col).Name) & "]"
    End If
    
End Sub

Private Sub Command4_Click()

    Me.Adodc1.Recordset.Sort = "[" & Trim(Me.Adodc1.Recordset.Fields(0).Name) & "]"
    

End Sub

Private Sub Command5_Click()

    Unload Me

End Sub

Private Sub Command6_Click()

    If Me.DataGrid1.Col >= 0 Then
        Me.Adodc1.Recordset.Sort = "[" & Trim(Me.Adodc1.Recordset.Fields(Me.DataGrid1.Col).Name) & "]"
    Else
        If defined_col <> Empty Then
            Me.Adodc1.Recordset.Sort = "[" & Trim(Me.Adodc1.Recordset.Fields(defined_col).Name) & "]"
        End If
    End If
    

End Sub

Private Sub Command7_Click()

    Dim id As Integer
    
    Me.Adodc1.Recordset.Sort = "[" & Trim(Me.Adodc1.Recordset.Fields(0).Name) & "]"
    
    If Not Me.Adodc1.Recordset.EOF Then
        Me.Adodc1.Recordset.MoveLast
        id = Me.Adodc1.Recordset![id_ΔΟΥ]
    End If
    
    Me.Adodc1.Recordset.AddNew
    Me.Adodc1.Recordset.Fields(0) = id + 1
    Me.Adodc1.Recordset.UpdateBatch adAffectCurrent
    Me.Adodc1.Recordset.Sort = "id_ΔΟΥ"
    Me.Adodc1.Recordset.MoveLast
    Me.DataGrid1.Col = 1
    Me.DataGrid1.SetFocus
    

End Sub

Private Sub Command8_Click()
   
    Dim ms As String
    If Me.Adodc1.Recordset.RecordCount >= 1 Then
        ms = MsgBox("Είσαι σίγουρος; (ΝΑΙ ή ΟΧΙ)", vbYesNo, "Παράθυρο διαγραφής")
        If ms = 6 Then
            If Not Me.Adodc1.Recordset.EOF Then
                With Me.Adodc1.Recordset
                    .Delete
                    .MoveNext
                    If .EOF And .RecordCount <> 0 Then
                        .MoveLast
                    Else
                        c_r = 0
                    End If
                    cr = .AbsolutePosition
                    .Requery
                    Me.DataGrid1.Columns(0).Visible = False
                    Me.DataGrid1.Columns(0).Caption = "Κωδικός ΔΟΥ"
                    Me.DataGrid1.Columns(1).Caption = "Περιγραφή ΔΟΥ"
                    Me.DataGrid1.Columns(0).Width = 2500
                    Me.DataGrid1.Columns(1).Width = 5500
                    .MoveFirst
                    .Move cr - 1
                End With
            End If
        End If
    End If


End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    defined_col = ColIndex
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = 0 And KeyCode = 9 Then
        rs.MoveNext
        Me.DataGrid1.Col = 1
        If rs.EOF = True Then
            rs.MoveFirst
            Me.DataGrid1.Col = 1
        End If
    ElseIf Shift = 1 And KeyCode = 9 Then
        If Me.DataGrid1.Row = 0 Then
            rs.MoveLast
        Else
            rs.MovePrevious
        End If
        Me.DataGrid1.Col = 1
    End If

End Sub

Private Sub Form_Load()
        
    
    Me.Adodc1.Refresh
    Set rs = Adodc1.Recordset
    Set Me.DataGrid1.DataSource = Me.Adodc1
        
    Me.Width = 6600
    Me.Height = 10100
        
    rs.Sort = "id_ΔΟΥ"
    If rs.RecordCount > 0 Then
        Me.DataGrid1.Row = 0
        Me.DataGrid1.Col = 1
    End If
    'Me.DataGrid1.Columns(0).Locked = True
    Me.DataGrid1.Columns(0).Visible = False
    
    Me.Adodc1.Caption = "Εγγραφή " & Me.DataGrid1.Row + 1 & " από " & rs.RecordCount
        
    Me.DataGrid1.Columns(0).Caption = "Κωδικός ΔΟΥ"
    Me.DataGrid1.Columns(1).Caption = "Περιγραφή ΔΟΥ"
    
    Me.DataGrid1.Columns(0).Width = 2500
    Me.DataGrid1.Columns(1).Width = 5500
    
    
        
End Sub

Private Sub Form_Unload(cancel As Integer)

    Set rs = Nothing
    Set Me.DataGrid1.DataSource = Nothing

End Sub
