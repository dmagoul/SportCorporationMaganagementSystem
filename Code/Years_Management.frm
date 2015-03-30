VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Years_Management 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Διαχείριση Αθλητικών Ετών"
   ClientHeight    =   9645
   ClientLeft      =   300
   ClientTop       =   570
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   6450
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
      Left            =   4920
      Picture         =   "Years_Management.frx":0000
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
      Left            =   3720
      Picture         =   "Years_Management.frx":5A78
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
      Picture         =   "Years_Management.frx":A995
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
      Picture         =   "Years_Management.frx":F69C
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
      Picture         =   "Years_Management.frx":1442B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8760
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Years_Management.frx":1922F
      Height          =   8175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   14420
      _Version        =   393216
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
      TabAcrossSplits =   -1  'True
      TabAction       =   2
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   8280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      ConnectMode     =   3
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
End
Attribute VB_Name = "Years_Management"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public defined_col As Integer
Public rs As ADODB.Recordset

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
    If Not Me.Adodc1.Recordset.EOF And Me.Adodc1.Recordset.AbsolutePosition > 1 Then
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
        id = Me.Adodc1.Recordset![id_αθλητικού_έτους]
    End If
    
    Me.Adodc1.Recordset.AddNew
    Me.Adodc1.Recordset.Fields(0).Value = id + 1
    Me.Adodc1.Recordset.UpdateBatch adAffectCurrent
    Me.Adodc1.Recordset.Sort = "id_αθλητικού_έτους"
    Me.Adodc1.Recordset.MoveLast
    Me.DataGrid1.Col = 2
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
        id = Me.Adodc1.Recordset![id_αθλητικού_έτους]
    End If
    
    Me.Adodc1.Recordset.AddNew
    Me.Adodc1.Recordset.Fields(0).Value = id + 1
    Me.Adodc1.Recordset.UpdateBatch adAffectCurrent
    Me.Adodc1.Recordset.Sort = "id_αθλητικού_έτους"
    Me.Adodc1.Recordset.MoveLast
    Me.DataGrid1.Col = 2
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
                    Me.DataGrid1.Columns(1).Caption = "Περιγραφή Αθλητικού Έτους"
                    Me.DataGrid1.Columns(2).Caption = "Έναρξη"
                    Me.DataGrid1.Columns(3).Caption = "Λήξη"
                    Me.DataGrid1.Columns(1).Width = 2800
                    Me.DataGrid1.Columns(2).Width = 1300
                    Me.DataGrid1.Columns(3).Width = 1300
                    .MoveFirst
                    .Move cr - 1
                End With
            End If
        End If
    End If


End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
        
    If Me.DataGrid1.Columns(2).Text <> "" And Me.DataGrid1.Columns(3).Text <> "" Then
        Me.DataGrid1.Columns(1).Text = year(Me.DataGrid1.Columns(2).Text) & "-" & year(Me.DataGrid1.Columns(3).Text)
    Else
        If Me.DataGrid1.Columns(2).Text <> "" And Me.DataGrid1.Columns(3).Text = "" Then
            Me.DataGrid1.Columns(1).Text = year(Me.DataGrid1.Columns(2).Text)
        Else
            If Me.DataGrid1.Columns(2).Text = "" And Me.DataGrid1.Columns(3).Text <> "" Then
                Me.DataGrid1.Columns(1).Text = year(Me.DataGrid1.Columns(3).Text)
            Else
                Me.DataGrid1.Columns(1).Text = "ΑΟΡΙΣΤΟ"
                'Me.Adodc1.Recordset.Delete adAffectCurrent
                'Me.Adodc1.Recordset.MoveNext
            End If
        End If
    End If
    
End Sub

Private Sub DataGrid1_ColEdit(ByVal ColIndex As Integer)
    
    If Me.DataGrid1.Columns(2).Text = "" And Me.DataGrid1.Columns(3).Text = "" Then
                Me.DataGrid1.Columns(1).Text = "ΑΟΡΙΣΤΟ"
                'Me.Adodc1.Recordset.Delete adAffectCurrent
                'Me.Adodc1.Recordset.MoveNext
    End If

End Sub

Private Sub DataGrid1_Error(ByVal DataError As Integer, Response As Integer)
    If DataError = 7007 Then
        Response = 0
    End If
End Sub

Private Sub DataGrid1_GotFocus()
    
    If Me.DataGrid1.Columns(2).Text = "" And Me.DataGrid1.Columns(3).Text = "" Then
                Me.DataGrid1.Columns(1).Text = "ΑΟΡΙΣΤΟ"
                'Me.Adodc1.Recordset.Delete adAffectCurrent
                'Me.Adodc1.Recordset.MoveNext
    End If

End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    
    defined_col = ColIndex
    
End Sub

Private Sub DataGrid1_LostFocus()

    If Me.DataGrid1.Columns(2).Text = "" And Me.DataGrid1.Columns(3).Text = "" Then
                Me.DataGrid1.Columns(1).Text = "ΑΟΡΙΣΤΟ"
                'Me.Adodc1.Recordset.Delete adAffectCurrent
                'Me.Adodc1.Recordset.MoveNext
    End If

End Sub

Private Sub Form_Load()
          
    Me.Adodc1.Refresh
    Set rs = Adodc1.Recordset
    Set Me.DataGrid1.DataSource = Me.Adodc1
                
    Me.Width = 6500
    Me.Height = 10100
      
    rs.Sort = "id_αθλητικού_έτους"
    If rs.RecordCount > 0 Then
        Me.DataGrid1.Row = 0
        Me.DataGrid1.Col = 1
    End If
    
    Me.Adodc1.Caption = "Εγγραφή " & Me.DataGrid1.Row + 1 & " από " & rs.RecordCount
        
    Me.DataGrid1.Columns(0).Visible = False
    Me.DataGrid1.Columns(1).Caption = "Περιγραφή Αθλητικού Έτους"
    'Me.DataGrid1.Columns(1).Locked = True
    Me.DataGrid1.Columns(2).Caption = "Έναρξη"
    Me.DataGrid1.Columns(3).Caption = "Λήξη"
    
    Me.DataGrid1.Columns(1).Width = 2800
    Me.DataGrid1.Columns(2).Width = 1300
    Me.DataGrid1.Columns(3).Width = 1300
    
    If rs.RecordCount > 0 Then
        Me.DataGrid1.Row = rs.RecordCount - 1
        Me.DataGrid1.Col = 2
    End If
    
End Sub

Private Sub refr_bt_Click()

    Me.Adodc1.Recordset.Fields(1).Value = Me.Adodc1.Recordset.Fields(2).Value & Me.Adodc1.Recordset.Fields(2).Value
    Me.Adodc1.Recordset.UpdateBatch adAffectCurrent
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    
    Set rs = Nothing
    Set Me.DataGrid1.DataSource = Nothing
        
End Sub
