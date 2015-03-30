VERSION 5.00
Begin VB.Form f_foto_a 
   Caption         =   "Εισαγωγή Φωτογραφίας Αθλητή"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   3990
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      Height          =   3915
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "f_foto_a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()

    File1.Path = Dir1.Path

End Sub

Private Sub File1_Click()

    Dim img As String
    
    On Error Resume Next   ' Set up error handling.
    
    img = File1.Path & "\" & File1.FileName
    athlet_management.Image1.Picture = LoadPicture(img)
    If Err Then
      Msg = "Δεν είναι δυνατή η εμφάνιση του επιλεγόμενου αρχείου ως ΦΩΤΟΓΡΑΦΙΑ! Επιλέξτε άλλο αρχείο!"
      MsgBox Msg, , "Μήνυμα Λάθους!" ' Display error message.
      'athlet_management.Image1.Picture = LoadPicture()   'Clear the picturebox.
      Err.Clear
      'Exit Sub   ' Quit if error occurs.
   Else
      athlet_management.image_path.Text = img
   End If

End Sub
