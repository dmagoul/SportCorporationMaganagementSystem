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
      Caption         =   "Διαχείριση Παραμέτρων και Λοιπά Εργαλεία"
      Begin VB.Menu dimoi_man 
         Caption         =   "Διαχείριση Δήμων"
      End
      Begin VB.Menu pe_man 
         Caption         =   "Διαχείριση Περιφερειακών Ενοτήτων"
      End
      Begin VB.Menu d0 
         Caption         =   "-"
      End
      Begin VB.Menu years_man 
         Caption         =   "Διαχείριση Αθλητικών Ετών"
      End
      Begin VB.Menu class_man 
         Caption         =   "Διαχείριση Κατηγοριών Τμημάτων Πισίνας"
      End
      Begin VB.Menu d05 
         Caption         =   "-"
      End
      Begin VB.Menu doi_man 
         Caption         =   "Διαχείριση ΔΟΥ"
      End
      Begin VB.Menu jobs_man 
         Caption         =   "Διαχείριση Επαγγελμάτων"
      End
      Begin VB.Menu d6 
         Caption         =   "-"
      End
      Begin VB.Menu schools_man 
         Caption         =   "Διαχείριση Σχολείων"
      End
      Begin VB.Menu d1 
         Caption         =   "-"
      End
      Begin VB.Menu t_esoda 
         Caption         =   "Διαχείριση Τύπων Εσόδων"
      End
      Begin VB.Menu t_eksoda 
         Caption         =   "Διαχείριση Τύπων Εξόδων"
      End
      Begin VB.Menu tip_par_eksod 
         Caption         =   "Διαχείριση Τύπων Παραστατικών Εσόδων - Εξόδων"
      End
      Begin VB.Menu d7 
         Caption         =   "-"
      End
      Begin VB.Menu for_backup 
         Caption         =   "Λήψη Αντιγράφου Ασφαλείας"
      End
   End
   Begin VB.Menu entity_man 
      Caption         =   "Διαχείριση Αθλητών && Μελών"
      Begin VB.Menu athlet_man 
         Caption         =   "Διαχείριση Αθλητών"
      End
      Begin VB.Menu mel_man 
         Caption         =   "Διαχείριση Μελών"
      End
   End
   Begin VB.Menu promith_man 
      Caption         =   "Διαχείριση Προμηθευτών && Οργανισμών"
   End
   Begin VB.Menu man_tm 
      Caption         =   "Διαχείριση Τμημάτων && Παρουσιών"
   End
   Begin VB.Menu oikonomika_manag 
      Caption         =   "Διαχείριση Οικονομικών"
      Begin VB.Menu proupol_man 
         Caption         =   "Διαχείριση Προϋπολογισμών"
      End
      Begin VB.Menu oik_kiniseis 
         Caption         =   "Διαχείριση Οικονομικών Κινήσεων (Έοοδα - Έξοδα)"
      End
   End
   Begin VB.Menu rep_man 
      Caption         =   "Εκτυπώσεις"
      Begin VB.Menu rep_man_tmim 
         Caption         =   "Τμημάτων"
         Begin VB.Menu rep_man_tmim_all 
            Caption         =   "Όλα τα Τμήματα"
         End
         Begin VB.Menu rep_man_tmim_ae 
            Caption         =   "Ανά Αθλητικό Έτος"
         End
      End
      Begin VB.Menu rep_man_mel 
         Caption         =   "Μελών"
         Begin VB.Menu rep_man_mel_all 
            Caption         =   "Όλα τα Μέλη"
         End
         Begin VB.Menu rep_man_mel_energa 
            Caption         =   "Ενεργών Μελών"
         End
         Begin VB.Menu rep_man_mel_ethelodes 
            Caption         =   "Εθελοντών Μελών"
         End
         Begin VB.Menu rep_man_mel_prop 
            Caption         =   "Προπονητών"
         End
         Begin VB.Menu rep_man_mel_gs 
            Caption         =   "Μελών Γενικής Συνέλευσης"
         End
         Begin VB.Menu rep_man_mel_ds 
            Caption         =   "Μελών Διοικητικού Συμβουλίου"
         End
      End
   End
   Begin VB.Menu eksodos 
      Caption         =   "Έξοδος"
   End
   Begin VB.Menu mn1 
      Caption         =   "mn1"
      Visible         =   0   'False
      Begin VB.Menu popmn1 
         Caption         =   "Εισαγωγή Φωτογραφίας Αθλητή"
      End
      Begin VB.Menu popmn2 
         Caption         =   "Εμφάνιση Φωτογραφίας Αθλητή"
      End
      Begin VB.Menu popmn3 
         Caption         =   "Διαγραφή Φωτογραφίας Αθλητή"
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
    Me.Caption = "Εφαρμογή Διαχείρισης Σωματείου ΠΟΣΕΙΔΩΝΑ                                                                                       © Άλκης Σερβετάς-Δώρα Μαγουλιώτη"
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

    'ΠΑΡΟΥΣΙΑΣΗ ΦΩΤΟΓΡΑΦΙΑΣ
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
        MsgBox "Δεν έχει αποθηκευτεί ΦΩΤΟΓΡΑΦΙΑ στον αθλητή / αθλήτρια!", , "Μήνυμα Ενημέρωσης!"
        'Image1.Picture = LoadPicture()
        athlet_management.Label19.Visible = True
    End If

End Sub

Private Sub popmn3_Click()

    Dim ms As String
    
    If athlet_management.ado_athlites.Recordset.Fields(20).Value <> "" Then
        ms = MsgBox("Είσαι σίγουρος; (ΝΑΙ ή ΟΧΙ)", vbYesNo, "Παράθυρο διαγραφής")
        If ms = 6 Then
            'ΔΙΑΓΡΑΦΗ ΦΩΤΟΓΡΑΦΙΑΣ
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
        MsgBox "Δεν υπάρχει Φωτογραφία για ΔΙΑΓΡΑΦΗ ...", vbOKOnly, "Παράθυρο ενημέρωσης"
    End If
    
End Sub

Private Sub promith_man_Click()

    promitheytes_management.Show

End Sub

Private Sub proupol_man_Click()

    proupologismos_management.Show

End Sub

Private Sub rep_man_mel_all_Click()
    
    Poseidon_DB.rsσυνοπτική_εκτύπωση_επιλεγμένων_μελών.Filter = ""
    s_sort = "[Επώνυμο]"
    Poseidon_DB.rsσυνοπτική_εκτύπωση_επιλεγμένων_μελών.Sort = s_sort
    Rep_συνοπτική_εκτύπωση_επιλεγμένων_μελών.Sections("ReportHeader").Controls("Label12").Caption = MDIForm1.rep_lbl
    Rep_συνοπτική_εκτύπωση_επιλεγμένων_μελών.Show

End Sub

Private Sub rep_man_mel_ds_Click()
    
    Rep_συνοπτική_εκτύπωση_μελώνΔΣ.Show
    
End Sub

Private Sub rep_man_mel_energa_Click()

    Rep_συνοπτική_εκτύπωση_ΕΝΕΡΓΩΝ_μελών.Show

End Sub

Private Sub rep_man_mel_ethelodes_Click()

    Rep_συνοπτική_εκτύπωση_ΕΘΕΛΟΝΤΩΝ_μελών.Show
    
End Sub

Private Sub rep_man_mel_gs_Click()

    Rep_συνοπτική_εκτύπωση_μελώνΓΣ.Show

End Sub

Private Sub rep_man_mel_prop_Click()

    Rep_συνοπτική_εκτύπωση_ΠΡΟΠΟΝΗΤΩΝ_μελών.Show

End Sub

Private Sub rep_man_tmim_ae_Click()

    ektiposi_tmimatwn_ana_ae.Show

End Sub

Private Sub rep_man_tmim_all_Click()

    Rep_ΚαρτέλαΤμήματος.Show

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
