Attribute VB_Name = "Module1"
Public s_string As String
Public s_string2 As String
Public s_sort As String
Const strChecked = "ώ"
Const strUnChecked = "q"

Function Refresh_ΤύποιΠαραστατικών(flg As Boolean, done_search As Integer)


    If done_search = 1 Then

        Dim i As Integer
        Dim f_string, s2 As String
    
        If flg = True Then  'ΕΙΝΑΙ ΠΙΣΤΩΣΗ, update με τους ΤΥΠΟΥΣ ΠΑΡΑΣΤΑΤΙΚΩΝ ΕΣΟΔΩΝ
            f_string = "[Τύπος] LIKE TRUE"
            MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Filter = f_string
            If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.RecordCount >= 1 Then
                f_string = ""
                MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.MoveFirst
                If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(5).Value <> Empty Then
                    pr = MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(5).Value
                End If
                If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(6).Value <> Empty Then
                    te = MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(6).Value
                End If
                If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(7).Value <> Empty Then
                    tr = MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(7).Value
                End If
                For i = 0 To MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.RecordCount - 1
                    'If tr >= pr And tr < te And te >= pr Then
                    If tr < te Then
                        If f_string = "" Then
                            f_string = "[Ονομασία] LIKE '" & MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(2).Value & "'"
                        Else
                            f_string = f_string & " OR [Ονομασία] LIKE '" & MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(2).Value & "'"
                        End If
                    End If
                    MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.MoveNext
                    If Not MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.EOF Then
                        If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(5).Value <> Empty Then
                            pr = MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(5).Value
                        End If
                        If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(6).Value <> Empty Then
                            te = MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(6).Value
                        End If
                        If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(7).Value <> Empty Then
                            tr = MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(7).Value
                        End If
                    End If
                Next i
                MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Filter = f_string
            End If
        Else 'ΕΙΝΑΙ ΧΡΕΩΣΗ, update με τους ΤΥΠΟΥΣ ΠΑΡΑΣΤΑΤΙΚΩΝ ΕΞΟΔΩΝ
            f_string = "[Τύπος] LIKE FALSE"
            MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Filter = f_string
            If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.RecordCount >= 1 Then
                f_string = ""
                MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.MoveFirst
                If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(5).Value <> Empty Then
                    pr = MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(5).Value
                End If
                If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(6).Value <> Empty Then
                    te = MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(6).Value
                End If
                If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(7).Value <> Empty Then
                    tr = MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(7).Value
                End If
                For i = 0 To MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.RecordCount - 1
                    'If tr >= pr And tr < te And te >= pr Then
                    If tr < te Then
                        If f_string = "" Then
                            f_string = "[Ονομασία] LIKE '" & MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(2).Value & "'"
                        Else
                            f_string = f_string & " OR [Ονομασία] LIKE '" & MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(2).Value & "'"
                        End If
                    End If
                    MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.MoveNext
                    If Not MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.EOF Then
                        If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(5).Value <> Empty Then
                            pr = MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(5).Value
                        Else
                            pr = -2
                        End If
                        If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(6).Value <> Empty Then
                            te = MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(6).Value
                        Else
                            pr = -1
                        End If
                        If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(7).Value <> Empty Then
                            tr = MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(7).Value
                        Else
                            pr = -3
                        End If
                    End If
                Next i
                MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Filter = f_string
            End If
        End If
        '
        MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Sort = "[Ονομασία]"
        MIA_oikonomiki_kinisi_management.raw_co_tipoi_parastatikwn.Clear
        If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.RecordCount >= 1 Then
            For i = 0 To MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.RecordCount - 1
                MIA_oikonomiki_kinisi_management.raw_co_tipoi_parastatikwn.AddItem (MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(3).Value & " <" & MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(2).Value & ">")
                MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.MoveNext
            Next i
        End If
    Else 'NOT DONE SEARCH
        MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Filter = ""
        MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Sort = "[Ονομασία]"
        MIA_oikonomiki_kinisi_management.raw_co_tipoi_parastatikwn.Clear
        If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.RecordCount >= 1 Then
            For i = 0 To MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.RecordCount - 1
                MIA_oikonomiki_kinisi_management.raw_co_tipoi_parastatikwn.AddItem (MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(3).Value & " <" & MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(2).Value & ">")
                MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.MoveNext
            Next i
        End If
    End If
    
End Function

Function Refresh_ΤύποιΕσόδων(done_search As Integer)

    Dim s As String

    s = MIA_oikonomiki_kinisi_management.co_raw_tipoi_esodwn.Text
    If done_search = 1 And MIA_oikonomiki_kinisi_management.co_py.Text <> "" Then
        Dim i As Integer
        MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.Filter = "[id_προϋπολογισμού] = " & MIA_oikonomiki_kinisi_management.ado_py.Recordset.Fields(0).Value
        MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.Sort = "[περιγραφή]"
        MIA_oikonomiki_kinisi_management.co_raw_tipoi_esodwn.Clear
        If MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.RecordCount >= 1 Then
            For i = 0 To MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.RecordCount - 1
                MIA_oikonomiki_kinisi_management.co_raw_tipoi_esodwn.AddItem (MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.Fields(2).Value)
                MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.MoveNext
            Next i
        End If
        MIA_oikonomiki_kinisi_management.co_raw_tipoi_esodwn.Text = s
    Else
        If done_search = 2 Then
            MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.Filter = "[id_προϋπολογισμού] = " & oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(1).Value
            MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.Sort = "[περιγραφή]"
            MIA_oikonomiki_kinisi_management.co_raw_tipoi_esodwn.Clear
            If MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.RecordCount >= 1 Then
                For i = 0 To MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.RecordCount - 1
                    MIA_oikonomiki_kinisi_management.co_raw_tipoi_esodwn.AddItem (MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.Fields(2).Value)
                    MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.MoveNext
                Next i
            End If
            MIA_oikonomiki_kinisi_management.co_raw_tipoi_esodwn.Text = s
        Else
            MIA_oikonomiki_kinisi_management.co_raw_tipoi_esodwn.Clear
        End If
    End If
    
End Function

Function Refresh_ΠΥ_Εσόδων(done_search As Integer)


    If done_search = 1 Then

        Dim i As Integer
        Dim f_string As String
    
        f_string = "[ΤύποιΕσόδων.περιγραφή] LIKE '" & MIA_oikonomiki_kinisi_management.co_raw_tipoi_esodwn.Text & "'"
        MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.Filter = f_string
        MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.Sort = "[Προϋπολογισμός.Περιγραφή]"
        'MIA_oikonomiki_kinisi_management.co_raw_py_esodwn.Clear
        If MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.RecordCount >= 1 Then
            For i = 0 To MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.RecordCount - 1
                'MIA_oikonomiki_kinisi_management.co_raw_py_esodwn.AddItem (MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.Fields(2).Value & " <" & MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.Fields(3).Value & ">")
                MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.MoveNext
            Next i
        End If
    Else 'NOT DONE SEARCH
        '
    End If
    
End Function

Function Refresh_ΤύποιΕξόδων(done_search As Integer)

On Error GoTo Refresh_ΤύποιΕξόδων_l1

    If done_search = 1 And MIA_oikonomiki_kinisi_management.co_py.Text <> "" Then
        Dim i As Integer
        Dim s As String
        
        s = MIA_oikonomiki_kinisi_management.co_raw_tipoi_eksodwn.Text
        MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.Filter = "[id_προϋπολογισμού] = " & MIA_oikonomiki_kinisi_management.ado_py.Recordset.Fields(0).Value
        MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.Sort = "[περιγραφή]"
        MIA_oikonomiki_kinisi_management.co_raw_tipoi_eksodwn.Clear
        If MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.RecordCount >= 1 Then
            For i = 0 To MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.RecordCount - 1
                MIA_oikonomiki_kinisi_management.co_raw_tipoi_eksodwn.AddItem (MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.Fields(3).Value)
                MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.MoveNext
            Next i
        End If
        MIA_oikonomiki_kinisi_management.co_raw_tipoi_eksodwn.Text = s
    Else
        MIA_oikonomiki_kinisi_management.co_raw_tipoi_eksodwn.Clear
    End If
    
Refresh_ΤύποιΕξόδων_l1:
    
End Function

Function Refresh_ΠΥ_Εξόδων(done_search As Integer)


    If done_search = 1 Then

        Dim i As Integer
        Dim f_string As String
    
        f_string = "[ΤύποιΕξόδων.περιγραφή] LIKE '" & MIA_oikonomiki_kinisi_management.co_raw_tipoi_eksodwn.Text & "'"
        MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.Filter = f_string
        MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.Sort = "[Προϋπολογισμός.Περιγραφή]"
        'MIA_oikonomiki_kinisi_management.co_raw_py_eksodwn.Clear
        If MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.RecordCount >= 1 Then
            For i = 0 To MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.RecordCount - 1
                'MIA_oikonomiki_kinisi_management.co_raw_py_eksodwn.AddItem (MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.Fields(5).Value & " <" & MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.Fields(2).Value & ">")
                MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.MoveNext
            Next i
        End If
    Else 'NOT DONE SEARCH
        '
    End If
    
End Function

Function ΕνημέρωσηΚωδικού() As Integer

    Dim id As Integer
    
    oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Filter = ""
    oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.Sort = "[id]"
    
    If oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.RecordCount >= 1 Then
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.MoveLast
        ΕνημέρωσηΚωδικού = oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset![id] + 1
        oikonomikes_kiniseis_management.tmp_ado_oik_kiniseis.Recordset.MoveLast
    Else
        ΕνημέρωσηΚωδικού = 1
    End If
    
End Function

Function ΕνημέρωσηΑριθμούΠαραστατικούΕσόδου() As Integer

    Dim s As String
    '
    i = MIA_oikonomiki_kinisi_management.raw_co_tipoi_parastatikwn.ListIndex
    If i >= 0 Then
        MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.MoveFirst
        For j = 1 To i
            MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.MoveNext
        Next j
        If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(7).Value <> "" Then
            id = MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(7).Value
        Else ' ΕΙΝΑΙ EMPTY Ο ΤΡΕΧΩΝ ΑΡΙΘΜΟΣ
            id = -1
        End If
    Else ' ΔΕΝ ΕΧΕΙ ΕΠΙΛΕΓΕΙ ΤΥΠΟΣ ΠΑΡΑΣΤΑΤΙΚΟΥ
        id = -1
    End If
    '
    ΕνημέρωσηΑριθμούΠαραστατικούΕσόδου = id + 1
    
End Function

Function ΕνημέρωσηΑθλητή() As Integer

    Dim i, j As Integer
    
    'i = MIA_oikonomiki_kinisi_management.co_athlites.SelectedItem
    'If i >= 1 Then
    '    MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.MoveFirst
    '    For j = 1 To i - 1
    '        MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.MoveNext
    '    Next j
    '    ΕνημέρωσηΑθλητή = MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.Fields(0).Value
    'Else
    '    ΕνημέρωσηΑθλητή = 0
    'End If
   '
    MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.Filter = ""
    If MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.RecordCount >= 1 And MIA_oikonomiki_kinisi_management.co_athlites.Text <> "" Then
        MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.MoveFirst
        MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.Find "[ΟΕΑ] LIKE '" & MIA_oikonomiki_kinisi_management.co_athlites.Text & "'"
        If Not MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.EOF Then
            ΕνημέρωσηΑθλητή = MIA_oikonomiki_kinisi_management.ado_athlites.Recordset.Fields("id").Value
        Else
            ΕνημέρωσηΑθλητή = 0
        End If
    Else
        ΕνημέρωσηΑθλητή = 0
    End If
    
    
End Function

Function ΕνημέρωσηΜέλους() As Integer

    'Dim i, j As Integer
    
    'i = MIA_oikonomiki_kinisi_management.co_meli.SelectedItem
    'If i >= 1 Then
    '    MIA_oikonomiki_kinisi_management.ado_meli.Recordset.MoveFirst
    '    For j = 1 To i - 1
    '        MIA_oikonomiki_kinisi_management.ado_meli.Recordset.MoveNext
    '    Next j
    '    ΕνημέρωσηΜέλους = MIA_oikonomiki_kinisi_management.ado_meli.Recordset.Fields(0).Value
    'Else
    '    ΕνημέρωσηΜέλους = -1
    'End If
    
    
    MIA_oikonomiki_kinisi_management.ado_meli.Recordset.Filter = ""
    If MIA_oikonomiki_kinisi_management.ado_meli.Recordset.RecordCount >= 1 And MIA_oikonomiki_kinisi_management.co_meli.Text <> "" Then
        MIA_oikonomiki_kinisi_management.ado_meli.Recordset.MoveFirst
        MIA_oikonomiki_kinisi_management.ado_meli.Recordset.Find "[OE_ΑΔΤ] LIKE '" & MIA_oikonomiki_kinisi_management.co_meli.Text & "'"
        If Not MIA_oikonomiki_kinisi_management.ado_meli.Recordset.EOF Then
            ΕνημέρωσηΜέλους = MIA_oikonomiki_kinisi_management.ado_meli.Recordset.Fields("id").Value
        Else
            ΕνημέρωσηΜέλους = 0
        End If
    Else
        ΕνημέρωσηΜέλους = 0
    End If
    
    
    
End Function

Function ΕνημέρωσηΟργανισμού() As Integer

    Dim i, j As Integer
    
    i = MIA_oikonomiki_kinisi_management.co_organismoi.SelectedItem
    If i >= 1 Then
        MIA_oikonomiki_kinisi_management.ado_organismoi.Recordset.MoveFirst
        For j = 1 To i - 1
            MIA_oikonomiki_kinisi_management.ado_organismoi.Recordset.MoveNext
        Next j
        ΕνημέρωσηΟργανισμού = MIA_oikonomiki_kinisi_management.ado_organismoi.Recordset.Fields(0).Value
    Else
        ΕνημέρωσηΟργανισμού = -1
    End If
    
End Function

Function ΕνημέρωσηΠαραστατικού() As Integer

    Dim i, j As Integer
    
    i = MIA_oikonomiki_kinisi_management.raw_co_tipoi_parastatikwn.ListIndex
    If i >= 0 Then
        MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.MoveFirst
        For j = 1 To i
            MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.MoveNext
        Next j
        If Not MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.EOF Then
            ΕνημέρωσηΠαραστατικού = MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(0).Value
        Else
            ΕνημέρωσηΠαραστατικού = -1
        End If
    Else
        ΕνημέρωσηΠαραστατικού = -1
    End If
    
End Function

Function ΕνημέρωσηΤρέχοντοςΑριθμούΠαραστατικούΕσόδου(id_par As Integer, tr_ar As Long)

    If MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.RecordCount >= 1 Then
        MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Find "[id] = " & id_par
        If Not MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.EOF Then
            MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.Fields(7).Value = tr_ar
            MIA_oikonomiki_kinisi_management.ado_tipoi_parastatikwn.Recordset.UpdateBatch adAffectCurrent
        End If
    End If

End Function

Function ΕνημέρωσηΤύπουΕξόδου() As Integer

    Dim i, j As Integer
    
    
    i = MIA_oikonomiki_kinisi_management.co_raw_tipoi_eksodwn.ListIndex
    If i >= 0 Then
        MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.MoveFirst
        For j = 1 To i
            MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.MoveNext
        Next j
        If Not MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.EOF Then
            ΕνημέρωσηΤύπουΕξόδου = MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.Fields(5).Value
        Else
            ΕνημέρωσηΤύπουΕξόδου = -1
        End If
    Else
        ΕνημέρωσηΤύπουΕξόδου = -1
    End If
    
End Function

Function ΕνημέρωσηΤύπουΠΥΕξόδων() As Integer

    Dim i, j As Integer
    
    'i = MIA_oikonomiki_kinisi_management.co_raw_py_eksodwn.ListIndex
    If i >= 0 Then
        MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.MoveFirst
        For j = 1 To i
            MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.MoveNext
        Next j
        If Not MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.EOF Then
            ΕνημέρωσηΤύπουΠΥΕξόδων = MIA_oikonomiki_kinisi_management.ado_py_eksodwn.Recordset.Fields(1).Value
        Else
            ΕνημέρωσηΤύπουΠΥΕξόδων = 0
        End If
    Else
        ΕνημέρωσηΤύπουΠΥΕξόδων = 0
    End If
    
End Function

Function ΕνημέρωσηΤύπουΕσόδου() As Integer

    Dim i, j As Integer
    
    i = MIA_oikonomiki_kinisi_management.co_raw_tipoi_esodwn.ListIndex
    If i >= 0 Then
        MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.MoveFirst
        For j = 1 To i
            MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.MoveNext
        Next j
        If Not MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.EOF Then
            ΕνημέρωσηΤύπουΕσόδου = MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.Fields(3).Value
        Else
            ΕνημέρωσηΤύπουΕσόδου = -1
        End If
    Else
        ΕνημέρωσηΤύπουΕσόδου = -1
    End If
    
End Function

Function ΕνημέρωσηΤύπουΠΥΕσόδων() As Integer

    Dim i, j As Integer
    
    'i = MIA_oikonomiki_kinisi_management.co_raw_py_esodwn.ListIndex
    'If i >= 0 Then
    '    MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.MoveFirst
    '    For j = 1 To i
    '        MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.MoveNext
    '    Next j
    '    If Not MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.EOF Then
    '        ΕνημέρωσηΤύπουΠΥΕσόδων = MIA_oikonomiki_kinisi_management.ado_py_esodwn.Recordset.Fields(1).Value
    '    Else
    '        ΕνημέρωσηΤύπουΠΥΕσόδων = 0
    '    End If
    'Else
    '    ΕνημέρωσηΤύπουΠΥΕσόδων = 0
    'End If
    
End Function

Function ΕύρεσηΑΜΟΠ(p As String, f As Integer, f_id As Integer) As String

    Dim cn As ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim s As String
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Open "poseidon.mdb"
    
    rs.Open p, cn, adOpenKeyset, adLockOptimistic
    rs.MoveFirst
    s = rs.Fields(0).Name & " = " & f_id
    rs.Find s
    If Not rs.EOF Then
        If f = 2 Then
            ΕύρεσηΑΜΟΠ = rs.Fields(f).Value & " " & rs.Fields(f + 1).Value
        Else
            If f = 3 Then 'ΤύποιΠαραστατικώνΕσόδωνΕξόδων
                ΕύρεσηΑΜΟΠ = rs.Fields(f).Value & " <" & rs.Fields(f - 1).Value & ">"
            Else
                ΕύρεσηΑΜΟΠ = rs.Fields(f).Value
            End If
        End If
    Else
        ΕύρεσηΑΜΟΠ = ""
    End If
    
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

End Function

Function GetDateFilter(s_cr As String, inpt As String, fld As String) As String

    If inpt <> "  /  /    " And inpt <> "" Then
        Dim ch, imera, year, minas As String
        imera = ""
        For i = 1 To 2
            ch = Mid$(inpt, i, 1)
            If ch <> "/" Then
                If ch <> "0" Then
                    imera = imera & ch
                End If
            Else
                Exit For
            End If
        Next i
        minas = ""
        For i = i + 1 To 5
            ch = Mid$(inpt, i, 1)
            If ch <> "/" Then
                If ch <> "0" Then
                    minas = minas & ch
                End If
            Else
                Exit For
            End If
        Next i
        year = ""
        For i = i + 1 To 10
            ch = Mid$(inpt, i, 1)
            If ch <> "/" Then
                year = year & ch
            Else
                Exit For
            End If
        Next i
        'imera = Mid$(inpt, 1, 2)
        'minas = Mid$(inpt, 4, 2)
        'year = Mid$(inpt, 7, 4)
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
        If s_cr <> "" Then
            If st1 <> "" Then
                's_cr = s_cr & "AND [Γέννηση] LIKE '" & st1 & "'"
                s_cr = s_cr & " AND [" & fld & "] LIKE '" & st1 & "'"
            End If
            If st2 <> "" Then
                s_cr = s_cr & " AND [" & fld & "] LIKE '" & st2 & "'"
            End If
            If st3 <> "" Then
                    s_cr = s_cr & " AND [" & fld & "] LIKE '" & st3 & "'"
            End If
        Else
            If st1 <> "" Then
                s_cr = "[" & fld & "] LIKE '" & st1 & "'"
            End If
            If st2 <> "" Then
                If s_cr <> "" Then
                    s_cr = s_cr & " AND [" & fld & "] LIKE '" & st2 & "'"
                Else
                    s_cr = "[" & fld & "] LIKE '" & st2 & "'"
                End If
            End If
            If st3 <> "" Then
                If s_cr <> "" Then
                    s_cr = s_cr & " AND [" & fld & "] LIKE '" & st3 & "'"
                Else
                    s_cr = "[" & fld & "] LIKE '" & st3 & "'"
                End If
            End If
        End If
    End If
    
    GetDateFilter = s_cr

End Function

Function ΕύρεσηID_από_String(p As String, str As String) As Integer

    Dim cn As ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim s As String
    
    Set cn = New ADODB.Connection
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Open "poseidon.mdb"
    
    rs.Open p, cn, adOpenKeyset, adLockOptimistic
    rs.MoveFirst
    s = rs.Fields(1).Name & " LIKE '*" & str & "*'"
    rs.Find s
    If Not rs.EOF Then
        ΕύρεσηID_από_String = rs.Fields(0).Value
    Else
        ΕύρεσηID_από_String = 0
    End If
    
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing

End Function

Function OikonomikesKiniseisRefresh()

    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(0).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(0).Width = 500
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(0).Alignment = dbgCenter
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(0).Caption = "Κωδ."
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(1).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(2).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(2).Width = 2000
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(2).Alignment = dbgLeft
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(2).Caption = "Προϋπολογισμός"
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(3).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(3).Width = 900
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(3).Alignment = dbgCenter
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(3).Caption = "Τύπος"
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(4).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(4).Width = 1000
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(4).Alignment = dbgCenter
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(4).Caption = "Κατάστ."
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(5).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(5).Width = 1300
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(5).Alignment = dbgRight
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(5).Caption = "Ημ/νία Κίν."
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(6).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(6).Caption = "Αθλητής"
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(6).Width = 1900
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(6).Alignment = dbgLeft
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(7).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(7).Caption = "Μέλος"
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(7).Width = 1900
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(7).Alignment = dbgLeft
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(8).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(8).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(8).Caption = "Οργανισμός"
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(8).Width = 1400
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(8).Alignment = dbgCenter
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(9).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(9).Caption = "Παραστατ."
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(9).Width = 1900
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(9).Alignment = dbgLeft
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(10).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(10).Caption = "Αρ."
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(10).Width = 500
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(10).Alignment = dbgCenter
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(11).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(11).Caption = "Ημ/νία Παρ."
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(11).Width = 1300
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(11).Alignment = dbgRight
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(12).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(12).Caption = "Κατηγορία Εξ."
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(12).Width = 1500
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(12).Alignment = dbgLeft
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(13).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(13).Caption = "Ποσό Χρ."
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(13).Width = 1200
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(13).Alignment = dbgRight
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(13).NumberFormat = "currency"
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(14).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(14).Caption = "Κατηγορία Εσ."
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(14).Width = 1500
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(14).Alignment = dbgLeft
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(15).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(15).Caption = "Ποσό Πίστ."
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(15).Width = 1200
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(15).Alignment = dbgRight
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(15).NumberFormat = "currency"
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(16).WrapText = True
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(16).Caption = "Αιτιολογία"
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(16).Width = 7000
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(16).Alignment = dbgLeft
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(17).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(18).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(19).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(20).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(21).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(22).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(23).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(24).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(25).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(26).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(27).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(28).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(29).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(30).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(31).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(32).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(33).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(34).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(35).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(36).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(37).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(38).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(39).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(40).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(41).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(42).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns(43).Visible = False
    oikonomikes_kiniseis_management.dt_oikon_kiniseis.Columns("Αιτιολογία2").Visible = False
    
    'Υπολογισμός Sum        tmp_ado_oik_kiniseis
    '    s1 = 0
    '    s2 = 0
    '    For i = 0 To oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.RecordCount - 1
    '        If i = 0 And Not oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.EOF Then
    '            oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.MoveFirst
    '        End If
    '        If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(13).Value <> "" Then
    '            s1 = s1 + Val(oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(13).Value)
    '        End If
    '        If oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(15).Value <> "" Then
    '            s2 = s2 + Val(oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.Fields(15).Value)
    '        End If
    '        oikonomikes_kiniseis_management.ado_oikon_kiniseis.Recordset.MoveNext
    '    Next i
    '    oikonomikes_kiniseis_management.s_xr.Text = FormatCurrency(s1, 2, vbTrue, , vbTrue)
    '    oikonomikes_kiniseis_management.s_pis.Text = FormatCurrency(s2, 2, vbTrue, , vbTrue)
    '
    
    'Υπολογισμός Sum        tmp_ado_oik_kiniseis
        Dim rs As Recordset
        Set rs = oikonomikes_kiniseis_management.tmp2_ado_oik_kiniseis.Recordset
        s1 = 0
        s2 = 0
        For i = 0 To rs.RecordCount - 1
            'If i = 0 And Not rs.EOF Then
            If i = 0 Then
                rs.MoveFirst
            End If
            If rs.Fields("ΠοσόΧρέωσης").Value <> "" Then
                's1 = s1 + Val(rs.Fields("ΠοσόΧρέωσης").Value)
                s1 = s1 + CDbl(rs.Fields("ΠοσόΧρέωσης").Value)
            End If
            If rs.Fields("ΠοσόΠίστωσης").Value <> "" Then
                's2 = s2 + Val(rs.Fields("ΠοσόΠίστωσης").Value)
                s2 = s2 + CDbl(rs.Fields("ΠοσόΠίστωσης").Value)
            End If
            rs.MoveNext
        Next i
        oikonomikes_kiniseis_management.s_xr.Text = FormatCurrency(s1, 2, vbTrue, , vbTrue)
        oikonomikes_kiniseis_management.s_pis.Text = FormatCurrency(s2, 2, vbTrue, , vbTrue)
    '
        
End Function

Sub CallRefreshParousiologio()

    parousies_management.dt_analytiko_par.Columns(0).Visible = False
    parousies_management.dt_analytiko_par.Columns(1).Visible = False
    parousies_management.dt_analytiko_par.Columns(2).Width = 2000
    parousies_management.dt_analytiko_par.Columns(2).Caption = "Ονοματεπώνυμο"
    parousies_management.dt_analytiko_par.Columns(3).Width = 800
    parousies_management.dt_analytiko_par.Columns(3).Caption = "Γέννηση"
    parousies_management.dt_analytiko_par.Columns(4).Visible = False
    For i = 1 To 31
        ind = 4 + i
        'parousies_management.dt_analytiko_par.Columns(ind).Width = 250
        'parousies_management.dt_analytiko_par.Columns(ind).Width = 500
        'parousies_management.dt_analytiko_par.Columns(ind).Caption = i
        'parousies_management.dt_analytiko_par.Columns(ind).Locked = False
        'parousies_management.dt_analytiko_par.Columns(ind).Text = strUnChecked
    Next i
    
End Sub
