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
      Caption         =   "... συνέχισε ..."
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
            Rep_Αναλυτική_Καρτέλα_1_Αθλητή.Show
            Poseidon_DB.rsαναλυτική_καρτέλα_όλων_των_αθλητών.Filter = "[id] = " & athlet_management.tmp_kod
            'ΠΑΡΟΥΣΙΑΣΗ ΦΩΤΟΓΡΑΦΙΑΣ
            If athlet_management.ado_athlites.Recordset.RecordCount >= 1 Then
                If athlet_management.ado_athlites.Recordset.Fields(20).Value <> Empty Then
                    Dim s As String
                    s = athlet_management.ado_athlites.Recordset.Fields(20).Value
                    Set Rep_Αναλυτική_Καρτέλα_1_Αθλητή.Sections(1).Controls("Image1").Picture = LoadPicture(s)
                Else
                    Set Rep_Αναλυτική_Καρτέλα_1_Αθλητή.Sections(1).Controls("Image1").Picture = LoadPicture("")
                End If
            End If
            Rep_Αναλυτική_Καρτέλα_1_Αθλητή.Refresh
        Else
            MsgBox "Δεν Υπάρχει Επιλεγμένος Αθλητής! Προσπαθήστε Ξανά...", vbOKOnly, "Μήνυμα Λάθους"
        End If
    Else
        If List1.ListIndex = 1 Then
            Rep_Αναλυτική_Καρτέλα_επιλεγμένων_Μαθητών.Hide
            If MDIForm1.s_string <> "" Then
                Poseidon_DB.rsαναλυτική_καρτέλα_όλων_των_αθλητών.Filter = MDIForm1.s_string
            Else
                Poseidon_DB.rsαναλυτική_καρτέλα_όλων_των_αθλητών.Filter = ""
            End If
            If MDIForm1.s_sort <> "" Then
                Poseidon_DB.rsαναλυτική_καρτέλα_όλων_των_αθλητών.Sort = MDIForm1.s_sort
            Else
                Poseidon_DB.rsαναλυτική_καρτέλα_όλων_των_αθλητών.Sort = "[id]"
            End If
            Rep_Αναλυτική_Καρτέλα_επιλεγμένων_Μαθητών.Refresh
            Rep_Αναλυτική_Καρτέλα_επιλεγμένων_Μαθητών.Show
        Else
            If List1.ListIndex = 2 Then
                Rep_Αναλυτική_Καρτέλα_όλων_των_Μαθητών.Hide
                Poseidon_DB.rsαναλυτική_καρτέλα_όλων_των_αθλητών.Filter = ""
                Rep_Αναλυτική_Καρτέλα_όλων_των_Μαθητών.Refresh
                Rep_Αναλυτική_Καρτέλα_όλων_των_Μαθητών.Show
            Else
                If List1.ListIndex = 3 Then
                    Rep_Συνοπτική_Εκτύπωση_Όλων_των_Αθλητών.Hide
                    Poseidon_DB.rsαναλυτική_καρτέλα_όλων_των_αθλητών.Filter = ""
                    Rep_Συνοπτική_Εκτύπωση_Όλων_των_Αθλητών.Refresh
                    Rep_Συνοπτική_Εκτύπωση_Όλων_των_Αθλητών.Show
                Else
                    If List1.ListIndex = 4 Then
                        Rep_συνοπτική_εκτύπωση_επιλεγμένων_αθλητών.Hide
                        If MDIForm1.s_string <> "" Then
                            Poseidon_DB.rsαναλυτική_καρτέλα_όλων_των_αθλητών.Filter = MDIForm1.s_string
                        Else
                            Poseidon_DB.rsαναλυτική_καρτέλα_όλων_των_αθλητών.Filter = ""
                        End If
                        If MDIForm1.s_sort <> "" Then
                            Poseidon_DB.rsαναλυτική_καρτέλα_όλων_των_αθλητών.Sort = MDIForm1.s_sort
                        Else
                            Poseidon_DB.rsαναλυτική_καρτέλα_όλων_των_αθλητών.Sort = "[id]"
                        End If
                        Rep_συνοπτική_εκτύπωση_επιλεγμένων_αθλητών.Refresh
                        Rep_συνοπτική_εκτύπωση_επιλεγμένων_αθλητών.Show
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
    
    List1.AddItem "ΕΚΤΥΠΩΣΗ Αναλυτικής Καρτέλας Τρέχοντος Αθλητή"
    List1.AddItem "ΕΚΤΥΠΩΣΗ Αναλυτικής Καρτέλας Επιλεγμένων Αθλητών"
    List1.AddItem "ΕΚΤΥΠΩΣΗ Αναλυτικής Καρτέλας Όλων των Αθλητών"
    List1.AddItem "Συνοπτική ΕΚΤΥΠΩΣΗ Όλων των Αθλητών"
    List1.AddItem "Συνοπτική ΕΚΤΥΠΩΣΗ Επιλεγμένων Αθλητών"

End Sub

