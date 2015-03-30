VERSION 5.00
Begin VB.Form f_forBackup 
   Caption         =   "���� ���������� ���������"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "..."
      DisabledPicture =   "ForBackup.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5880
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "���� ���������� ���������"
      Top             =   360
      Width           =   375
   End
   Begin VB.FileListBox File1 
      Height          =   3600
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "..."
      DisabledPicture =   "ForBackup.frx":00C3
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2520
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "���������� ���� �������"
      Top             =   360
      Width           =   375
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
Attribute VB_Name = "f_forBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    On Error GoTo l1

    Dim fold, fullpath As String
    
    fold = InputBox("����� �������", "���������� ���� �������")
    If Len(fold) <> 0 Then
        fullpath = Dir1.Path & fold
        MkDir fullpath
        Dir1.Path = fullpath
        File1.Path = fullpath
    End If
    
    
l1:     If Err.Number = 75 Then '� ������� ��� �������
            MsgBox "� ������� ��� �������, ������������ ���...", vbOKOnly, "������ ������"
            Call Command1_Click
        End If

End Sub

Private Sub Command2_Click()


    Dim tmpf As String
    
    tmpf = Dir1.Path & "\POSEIDON_BACKUP_" & Day(Date) & "_" & Month(Date) & "_" & year(Date) & "_" & Hour(Time) & "_" & Minute(Time) & "_" & Second(Time)
    MkDir tmpf
    
    ap = MsgBox("����� �������� ��� �� �������� ���������� ��� ���������� ��������� - " & Dir1.Path, vbOKCancel, "�������� ����������� ���������")
    
    If ap = 1 And (App.Path & "\poseidon.mdb" <> tmpf & "\poseidon.mdb") Then

        FileCopy App.Path & "\poseidon.mdb", tmpf & "\poseidon.mdb"
        FileCopy App.Path & "\SETUP\setup.exe", tmpf & "\setup.exe"
    
        Dim objfso
        Dim strFile, strSource, strDestination As String
    
        strDestination = tmpf & "\�����������\"
        strSource = App.Path & "\�����������\"
        strFile = Dir(strSource & "*.*")
        Set objfso = CreateObject("Scripting.FileSystemObject")
        Do While Len(strFile)
            With objfso
               If Not .FolderExists(strDestination) Then .CreateFolder (strDestination)
                    .CopyFile strSource & strFile, strDestination & strFile
            End With
            strFile = Dir
        Loop
    
    
        strDestination = tmpf & "\Btmps\"
        strSource = App.Path & "\Btmps\"
        strFile = Dir(strSource & "*.*")
        Set objfso = CreateObject("Scripting.FileSystemObject")
        Do While Len(strFile)
            With objfso
               If Not .FolderExists(strDestination) Then .CreateFolder (strDestination)
                    .CopyFile strSource & strFile, strDestination & strFile
            End With
            strFile = Dir
        Loop
    

        File1.Refresh
        Dir1.Refresh
        MsgBox "��������� ���� ��� ������� �� ��������!"
        Set objfso = Nothing
    ElseIf ap = 2 Then
    '
    Else
        MsgBox "����� ���������� � ���������� ������� ���������� ��� ���������� ���������!", vbOKOnly, "������ ������"
    End If

End Sub

Private Sub Dir1_Change()

    File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()

    Dir1.Path = Drive1.Drive
    'Dir1.Refresh
    'Me.File1.Refresh

End Sub

Private Sub File1_Click()

    Dim img As String
    
    On Error Resume Next   ' Set up error handling.
    
    img = File1.Path & "\" & File1.FileName
    athlet_management.Image1.Picture = LoadPicture(img)
    If Err Then
      Msg = "��� ����� ������ � �������� ��� ������������ ������� �� ����������! �������� ���� ������!"
      MsgBox Msg, , "������ ������!" ' Display error message.
      'athlet_management.Image1.Picture = LoadPicture()   'Clear the picturebox.
      Err.Clear
      'Exit Sub   ' Quit if error occurs.
   Else
      athlet_management.image_path.Text = img
   End If

End Sub

