VERSION 5.00
Begin VB.Form frm_pr_athl 
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   7995
   Begin VB.CommandButton bt_epib 
      Caption         =   "... �������� ..."
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "frm_pr_athl.frx":0000
      Left            =   1200
      List            =   "frm_pr_athl.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frm_pr_athl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id_tmimatos_report As Integer

Private Sub bt_epib_Click()

    If List1.ListIndex = 0 Then
        If athlet_management.ado_athlites.Recordset.RecordCount >= 1 And athlet_management.tmp_kod <> "" Then
            Rep_���������_�������_1_������.Show
            Poseidon_DB.rs���������_�������_����_���_�������.Filter = "[id] = " & athlet_management.tmp_kod
            '���������� �����������
            If athlet_management.ado_athlites.Recordset.RecordCount >= 1 Then
                If athlet_management.ado_athlites.Recordset.Fields(20).Value <> Empty Then
                    Dim s As String
                    s = athlet_management.ado_athlites.Recordset.Fields(20).Value
                    Set Rep_���������_�������_1_������.Sections(1).Controls("Image1").Picture = LoadPicture(s)
                Else
                    Set Rep_���������_�������_1_������.Sections(1).Controls("Image1").Picture = LoadPicture("")
                End If
            End If
            Rep_���������_�������_1_������.Refresh
        Else
            MsgBox "��� ������� ����������� �������! ����������� ����...", vbOKOnly, "������ ������"
        End If
    Else
        If List1.ListIndex = 1 Then
            Rep_���������_�������_�����������_�������.Hide
            If MDIForm1.s_string <> "" Then
                Poseidon_DB.rs���������_�������_����_���_�������.Filter = MDIForm1.s_string
            Else
                Poseidon_DB.rs���������_�������_����_���_�������.Filter = ""
            End If
            If MDIForm1.s_sort <> "" Then
                Poseidon_DB.rs���������_�������_����_���_�������.Sort = MDIForm1.s_sort
            Else
                Poseidon_DB.rs���������_�������_����_���_�������.Sort = "[id]"
            End If
            Rep_���������_�������_�����������_�������.Refresh
            Rep_���������_�������_�����������_�������.Show
        Else
            If List1.ListIndex = 2 Then
                Rep_���������_�������_����_���_�������.Hide
                Poseidon_DB.rs���������_�������_����_���_�������.Filter = ""
                Rep_���������_�������_����_���_�������.Refresh
                Rep_���������_�������_����_���_�������.Show
            Else
                If List1.ListIndex = 3 Then
                    Rep_���������_��������_����_���_�������.Hide
                    Poseidon_DB.rs���������_�������_����_���_�������.Filter = ""
                    Rep_���������_��������_����_���_�������.Refresh
                    Rep_���������_��������_����_���_�������.Show
                Else
                    If List1.ListIndex = 4 Then
                        Rep_���������_��������_�����������_�������.Hide
                        If MDIForm1.s_string <> "" Then
                            Poseidon_DB.rs���������_�������_����_���_�������.Filter = MDIForm1.s_string
                        Else
                            Poseidon_DB.rs���������_�������_����_���_�������.Filter = ""
                        End If
                        If MDIForm1.s_sort <> "" Then
                            Poseidon_DB.rs���������_�������_����_���_�������.Sort = MDIForm1.s_sort
                        Else
                            Poseidon_DB.rs���������_�������_����_���_�������.Sort = "[id]"
                        End If
                        Rep_���������_��������_�����������_�������.Refresh
                        Rep_���������_��������_�����������_�������.Show
                    Else
                    '
                    '
                    End If
                End If
            End If
        End If
    End If
    frm_pr_athl.Hide
    
End Sub

Private Sub Form_Load()
    
    List1.AddItem "�������� ���������� �������� ��������� ������"
    List1.AddItem "�������� ���������� �������� ����������� �������"
    List1.AddItem "�������� ���������� �������� ���� ��� �������"
    List1.AddItem "��������� �������� ���� ��� �������"
    List1.AddItem "��������� �������� ����������� �������"

End Sub

