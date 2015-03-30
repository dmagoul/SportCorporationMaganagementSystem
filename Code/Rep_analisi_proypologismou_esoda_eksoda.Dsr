VERSION 5.00
Begin {78E93846-85FD-11D0-8487-00A0C90DC8A9} Rep_analisi_proypologismou_esoda_eksoda 
   Bindings        =   "Rep_analisi_proypologismou_esoda_eksoda.dsx":0000
   Caption         =   "Ανάλυση Προϋπολογισμού Εσόδων - Εξόδων"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26247
   _ExtentY        =   16140
   _Version        =   393216
   _DesignerVersion=   100684101
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   GridX           =   1
   GridY           =   1
   LeftMargin      =   1000
   TopMargin       =   500
   BottomMargin    =   500
   _Settings       =   7
   DataMember      =   "ανάλυση_γενικού_προϋπολογισμού"
   NumSections     =   7
   SectionCode0    =   1
   BeginProperty Section0 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "ReportHeader"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
   SectionCode1    =   2
   BeginProperty Section1 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "PageHeader"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
   SectionCode2    =   3
   BeginProperty Section2 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "ανάλυση_γενικού_προϋπολογισμού_Header"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
   SectionCode3    =   4
   BeginProperty Section3 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "ΜόνοΈσοδα_Detail"
      Object.Height          =   1440
      NumControls     =   0
   EndProperty
   SectionCode4    =   5
   BeginProperty Section4 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "ανάλυση_γενικού_προϋπολογισμού_Footer"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
   SectionCode5    =   7
   BeginProperty Section5 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "PageFooter"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
   SectionCode6    =   8
   BeginProperty Section6 {1C13A8E0-A0B6-11D0-848E-00A0C90DC8A9} 
      _Version        =   393216
      Name            =   "ReportFooter"
      Object.Height          =   360
      NumControls     =   0
   EndProperty
End
Attribute VB_Name = "Rep_analisi_proypologismou_esoda_eksoda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DataReport_Initialize()

    Poseidon_DB.Commands("ΔιαγραφήΈσοδαΈξοδα").Execute
    Poseidon_DB.Commands("ΈσοδαΈξοδα1").Execute
    Poseidon_DB.Commands("ΈσοδαΈξοδα2").Execute
    'Poseidon_DB.Commands("Έσοδα_Έξοδα2").Execute
    ''Poseidon_DB.rsαναλυτική_καρτέλα_1_αθλητή.Open
    ''Poseidon_DB.rsαναλυτική_καρτέλα_1_αθλητή.Requery
    
    Poseidon_DB.rsανάλυση_γενικού_προϋπολογισμού.Filter = "[id_προϋπολογισμού] = '" & anlisi_proupologismou.txt_id.Text & "'"
    
End Sub
