VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000F&
   Caption         =   $"MDIForm1.frx":0000
   ClientHeight    =   10155
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   17160
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu parameters_man 
      Caption         =   "���������� ���������� ��� ����� ��������"
      Begin VB.Menu dimoi_man 
         Caption         =   "���������� �����"
      End
      Begin VB.Menu pe_man 
         Caption         =   "���������� ������������� ��������"
      End
      Begin VB.Menu d0 
         Caption         =   "-"
      End
      Begin VB.Menu years_man 
         Caption         =   "���������� ��������� ����"
      End
      Begin VB.Menu class_man 
         Caption         =   "���������� ���������� �������� �������"
      End
      Begin VB.Menu d05 
         Caption         =   "-"
      End
      Begin VB.Menu doi_man 
         Caption         =   "���������� ���"
      End
      Begin VB.Menu jobs_man 
         Caption         =   "���������� ������������"
      End
      Begin VB.Menu d6 
         Caption         =   "-"
      End
      Begin VB.Menu schools_man 
         Caption         =   "���������� ��������"
      End
      Begin VB.Menu d1 
         Caption         =   "-"
      End
      Begin VB.Menu t_esoda 
         Caption         =   "���������� ����� ������"
      End
      Begin VB.Menu t_eksoda 
         Caption         =   "���������� ����� ������"
      End
      Begin VB.Menu tip_par_eksod 
         Caption         =   "���������� ����� ������������ ������ - ������"
      End
      Begin VB.Menu d7 
         Caption         =   "-"
      End
      Begin VB.Menu for_backup 
         Caption         =   "���� ���������� ���������"
      End
   End
   Begin VB.Menu entity_man 
      Caption         =   "���������� ������� && �����"
      Begin VB.Menu athlet_man 
         Caption         =   "���������� �������"
      End
      Begin VB.Menu mel_man 
         Caption         =   "���������� �����"
      End
   End
   Begin VB.Menu promith_man 
      Caption         =   "���������� ����������� && ����������"
   End
   Begin VB.Menu man_tm 
      Caption         =   "���������� �������� && ���������"
   End
   Begin VB.Menu oikonomika_manag 
      Caption         =   "���������� �����������"
      Begin VB.Menu proupol_man 
         Caption         =   "���������� ��������������"
      End
      Begin VB.Menu oik_kiniseis 
         Caption         =   "���������� ����������� �������� (����� - �����)"
      End
   End
   Begin VB.Menu rep_man 
      Caption         =   "����������"
      Begin VB.Menu rep_man_tmim 
         Caption         =   "��������"
         Begin VB.Menu rep_man_tmim_all 
            Caption         =   "��� �� �������"
         End
         Begin VB.Menu rep_man_tmim_ae 
            Caption         =   "��� �������� ����"
         End
      End
      Begin VB.Menu rep_man_mel 
         Caption         =   "�����"
         Begin VB.Menu rep_man_mel_all 
            Caption         =   "��� �� ����"
         End
         Begin VB.Menu rep_man_mel_energa 
            Caption         =   "������� �����"
         End
         Begin VB.Menu rep_man_mel_ethelodes 
            Caption         =   "��������� �����"
         End
         Begin VB.Menu rep_man_mel_prop 
            Caption         =   "����������"
         End
         Begin VB.Menu rep_man_mel_gs 
            Caption         =   "����� ������� ����������"
         End
         Begin VB.Menu rep_man_mel_ds 
            Caption         =   "����� ����������� ����������"
         End
      End
   End
   Begin VB.Menu eksodos 
      Caption         =   "������"
   End
   Begin VB.Menu mn1 
      Caption         =   "mn1"
      Visible         =   0   'False
      Begin VB.Menu popmn1 
         Caption         =   "�������� ����������� ������"
      End
      Begin VB.Menu popmn2 
         Caption         =   "�������� ����������� ������"
      End
      Begin VB.Menu popmn3 
         Caption         =   "�������� ����������� ������"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public melos_id As Integer
Public s_sort As String
Public s_string As String
Public rep_lbl As String
Public is_a_new_record_without_save As Integer



Private Sub athlet_man_Click()
    
    athlet_management.Show
    
End Sub

Private Sub class_man_Click()
    
    class_management.Show
    
End Sub

Private Sub dimoi_man_Click()

    dimoi_management.Show

End Sub

Private Sub doi_man_Click()

    doi_management.Show

End Sub

Private Sub eksodos_Click()

    Unload MDIForm1

End Sub

Private Sub for_backup_Click()

    f_forBackup.Show

End Sub

Private Sub jobs_man_Click()

    job_management.Show

End Sub

Private Sub man_tm_Click()
    tmima_management.Show
End Sub

Private Sub MDIForm_Load()
    Static defined_col As Integer
    Me.Caption = "�������� ����������� ��������� ���������                                                                                       � ����� ��������-���� ����������"
End Sub

Private Sub mel_man_Click()

    meli_management.Show

End Sub

Private Sub oik_kiniseis_Click()

    oikonomikes_kiniseis_management.Show

End Sub

Private Sub pe_man_Click()

    pe_management.Show

End Sub

Private Sub popmn1_Click()

    f_foto_a.Show

End Sub

Private Sub popmn2_Click()

    '���������� �����������
    'If Me.ado_athlites.Recordset.Fields(20).ActualSize <> 0 Then
    If athlet_management.ado_athlites.Recordset.Fields(20).Value <> "" Then
        'sFile = App.Path + "\" + "image1.jpg"
        'Dim bytes() As Byte
        'Dim num_blocks As Long
        'Dim left_over As Long
        'Dim block_num As Long
        'Open sFile For Binary As #1
        'file_length = LenB(Me.ado_athlites.Recordset.Fields(20))
        'num_blocks = 1
        'left_over = 0
        'Me.ado_athlites.Recordset.Move Me.ado_athlites.Recordset.AbsolutePosition - 1, 1
        'For block_num = 1 To num_blocks
        '    bytes = Me.ado_athlites.Recordset.Fields(20).GetChunk(file_length)
        '    Put #1, , bytes
        'Next block_num
        'If left_over > 0 Then
        '    bytes = Me.ado_athlites.Recordset.Fields(20).GetChunk(left_over)
        '    Put #1, , bytes
        'End If
        'Close #1
        athlet_management.Image1.Picture = LoadPicture(athlet_management.ado_athlites.Recordset.Fields(20).Value)
        athlet_management.image_path.Text = athlet_management.ado_athlites.Recordset.Fields(20).Value
        athlet_management.Label19.Visible = False
        'Kill sFile
    Else
        MsgBox "��� ���� ����������� ���������� ���� ������ / ��������!", , "������ ����������!"
        'Image1.Picture = LoadPicture()
        athlet_management.Label19.Visible = True
    End If

End Sub

Private Sub popmn3_Click()

    Dim ms As String
    
    If athlet_management.ado_athlites.Recordset.Fields(20).Value <> "" Then
        ms = MsgBox("����� ��������; (��� � ���)", vbYesNo, "�������� ���������")
        If ms = 6 Then
            '�������� �����������
            If athlet_management.ado_athlites.Recordset.Fields(20).Value <> "" Then
                Kill athlet_management.ado_athlites.Recordset.Fields(20).Value
                athlet_management.image_path.Text = ""
                athlet_management.ado_athlites.Recordset.Fields(20).Value = ""
                athlet_management.ado_athlites.Recordset.UpdateBatch adAffectCurrent
                athlet_management.Image1.Picture = LoadPicture()
                athlet_management.Label19.Visible = True
            End If
        End If
        '
    Else
        MsgBox "��� ������� ���������� ��� �������� ...", vbOKOnly, "�������� ����������"
    End If
    
End Sub

Private Sub promith_man_Click()

    promitheytes_management.Show

End Sub

Private Sub proupol_man_Click()

    proupologismos_management.Show

End Sub

Private Sub rep_man_mel_all_Click()
    
    Poseidon_DB.rs���������_��������_�����������_�����.Filter = ""
    s_sort = "[�������]"
    Poseidon_DB.rs���������_��������_�����������_�����.Sort = s_sort
    Rep_���������_��������_�����������_�����.Sections("ReportHeader").Controls("Label12").Caption = MDIForm1.rep_lbl
    Rep_���������_��������_�����������_�����.Show

End Sub

Private Sub rep_man_mel_ds_Click()
    
    Rep_���������_��������_�������.Show
    
End Sub

Private Sub rep_man_mel_energa_Click()

    Rep_���������_��������_�������_�����.Show

End Sub

Private Sub rep_man_mel_ethelodes_Click()

    Rep_���������_��������_���������_�����.Show
    
End Sub

Private Sub rep_man_mel_gs_Click()

    Rep_���������_��������_�������.Show

End Sub

Private Sub rep_man_mel_prop_Click()

    Rep_���������_��������_����������_�����.Show

End Sub

Private Sub rep_man_tmim_ae_Click()

    ektiposi_tmimatwn_ana_ae.Show

End Sub

Private Sub rep_man_tmim_all_Click()

    Rep_���������������.Show

End Sub

Private Sub schools_man_Click()

    school_management.Show

End Sub

Private Sub t_eksoda_Click()

    typoi_eksodwn_management.Show

End Sub

Private Sub t_esoda_Click()

    typoi_esodwn_management.Show

End Sub

Private Sub tip_par_eksod_Click()

    typoi_parastatikwn_esodwn_eksodwn_management.Show
      
End Sub

Private Sub years_man_Click()

    Years_Management.Show

End Sub
