VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDiscos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discos"
   ClientHeight    =   5445
   ClientLeft      =   1095
   ClientTop       =   405
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   9030
   Begin VB.Data datPrimaryRS 
      Caption         =   " "
      Connect         =   "Access"
      DatabaseName    =   "Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5145
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from [Discos]"
      Top             =   120
      Width           =   2760
   End
   Begin TabDlg.SSTab sstGeral 
      Height          =   4680
      Left            =   90
      TabIndex        =   8
      Top             =   600
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8255
      _Version        =   327680
      TabHeight       =   520
      TabCaption(0)   =   "Discos"
      TabPicture(0)   =   "frmDisco.frx":0000
      Tab(0).ControlCount=   22
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLabels(8)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabels(7)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabels(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLabels(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLabels(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblLabels(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblLabels(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblLabels(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dlgCapa"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdAddFoto"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Figura"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtComent�rio"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtConserva��o"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtNota"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtDescri��o"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Data2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Data1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "dbcM�dia"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "updNota"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "updConserva��o"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "dbcGravadora"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtData"
      Tab(0).Control(21).Enabled=   0   'False
      TabCaption(1)   =   "Empr�stimos"
      TabPicture(1)   =   "frmDisco.frx":001C
      Tab(1).ControlCount=   12
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Shape1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblLabels(13)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblLabels(12)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblLabels(11)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblLabels(10)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblLabels(9)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtEmpEndere�o"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtEmpNome"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "chkEmprestado"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtEmpData"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtEmpTelefone"
      Tab(1).Control(11).Enabled=   0   'False
      TabCaption(2)   =   "M�sicas"
      TabPicture(2)   =   "frmDisco.frx":0038
      Tab(2).ControlCount=   12
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label5"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "grdDataGrid"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "dbcBanda"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "dbcEstilo"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Text1"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "Data4"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Data3"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "txtDura��o"
      Tab(2).Control(11).Enabled=   0   'False
      Begin MSMask.MaskEdBox txtData 
         DataField       =   "Data"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   1320
         TabIndex        =   47
         Top             =   1920
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   327680
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin MSDBCtls.DBCombo dbcGravadora 
         Bindings        =   "frmDisco.frx":0054
         DataField       =   "Gravadora"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   1320
         TabIndex        =   46
         Top             =   1125
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   327680
         Style           =   2
         ListField       =   "Descri��o"
         BoundColumn     =   "C�digo"
         Text            =   ""
      End
      Begin MSMask.MaskEdBox txtDura��o 
         DataField       =   "Dura��o"
         DataSource      =   "datSecondaryRS"
         Height          =   315
         Left            =   -74760
         TabIndex        =   45
         Top             =   3165
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   327680
         ForeColor       =   16512
         MaxLength       =   8
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtEmpTelefone 
         DataField       =   "EmpTelefone"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   -73305
         TabIndex        =   43
         Top             =   2910
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   327680
         ForeColor       =   16512
         MaxLength       =   14
         Mask            =   "(###) ### ####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtEmpData 
         DataField       =   "EmpData"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   -73305
         TabIndex        =   42
         Top             =   1710
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   327680
         ForeColor       =   16512
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin ComCtl2.UpDown updConserva��o 
         Height          =   315
         Left            =   1830
         TabIndex        =   41
         Top             =   2760
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   327680
         BuddyControl    =   "txtConserva��o"
         BuddyDispid     =   196625
         OrigLeft        =   3120
         OrigTop         =   3000
         OrigRight       =   3315
         OrigBottom      =   3375
         Max             =   5
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown updNota 
         Height          =   315
         Left            =   1830
         TabIndex        =   40
         Top             =   2340
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   327680
         BuddyControl    =   "txtNota"
         BuddyDispid     =   196624
         OrigLeft        =   3120
         OrigTop         =   2640
         OrigRight       =   3315
         OrigBottom      =   2895
         Max             =   5
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSDBCtls.DBCombo dbcM�dia 
         Bindings        =   "frmDisco.frx":0072
         DataField       =   "M�dia"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   1320
         TabIndex        =   39
         Top             =   1530
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   327680
         MatchEntry      =   -1  'True
         Style           =   2
         ForeColor       =   16512
         ListField       =   "Descri��o"
         BoundColumn     =   "C�digo"
         Text            =   "DBCombo3"
      End
      Begin VB.CheckBox chkEmprestado 
         DataField       =   "Emprestado"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Left            =   -73305
         TabIndex        =   33
         Top             =   645
         Width           =   3375
      End
      Begin VB.TextBox txtEmpNome 
         DataField       =   "EmpNome"
         DataSource      =   "datPrimaryRS"
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   -73305
         MaxLength       =   50
         TabIndex        =   32
         Top             =   2100
         Width           =   4395
      End
      Begin VB.TextBox txtEmpEndere�o 
         DataField       =   "EmpEndere�o"
         DataSource      =   "datPrimaryRS"
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   -73305
         MaxLength       =   50
         TabIndex        =   31
         Top             =   2505
         Width           =   4395
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "Discos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4140
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from Gravadoras order by Descri��o"
         Top             =   1125
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "Discos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4125
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from M�dias order by Descri��o"
         Top             =   1530
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtDescri��o 
         DataField       =   "Descri��o"
         DataSource      =   "datPrimaryRS"
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   22
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox txtNota 
         DataField       =   "Nota"
         DataSource      =   "datPrimaryRS"
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   1320
         TabIndex        =   21
         Top             =   2340
         Width           =   510
      End
      Begin VB.TextBox txtConserva��o 
         DataField       =   "Conserva��o"
         DataSource      =   "datPrimaryRS"
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   1320
         TabIndex        =   20
         Top             =   2760
         Width           =   510
      End
      Begin VB.TextBox txtComent�rio 
         DataField       =   "Coment�rio"
         DataSource      =   "datPrimaryRS"
         ForeColor       =   &H00004080&
         Height          =   990
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   3480
         Width           =   4725
      End
      Begin VB.PictureBox Figura 
         BorderStyle     =   0  'None
         DataField       =   "Capa"
         DataSource      =   "datPrimaryRS"
         Height          =   2295
         Left            =   4215
         ScaleHeight     =   2295
         ScaleWidth      =   3255
         TabIndex        =   18
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CommandButton cmdAddFoto 
         Caption         =   "&Inserir foto ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   17
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   "Discos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -73365
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from Bandas order by Descri��o"
         Top             =   2370
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data4 
         Caption         =   "Data4"
         Connect         =   "Access"
         DatabaseName    =   "Discos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -69210
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select * from Estilos order by Descri��o"
         Top             =   2385
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         DataField       =   "Coment�rio"
         DataSource      =   "datSecondaryRS"
         ForeColor       =   &H00004080&
         Height          =   645
         Left            =   -74745
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3855
         Width           =   7305
      End
      Begin MSDBCtls.DBCombo dbcEstilo 
         Bindings        =   "frmDisco.frx":008C
         DataField       =   "Estilo"
         DataSource      =   "datSecondaryRS"
         Height          =   315
         Left            =   -71025
         TabIndex        =   10
         Top             =   2520
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   556
         _Version        =   327680
         Style           =   2
         ForeColor       =   16512
         ListField       =   "Descri��o"
         BoundColumn     =   "C�digo"
         Text            =   "DBCombo1"
      End
      Begin MSDBCtls.DBCombo dbcBanda 
         Bindings        =   "frmDisco.frx":009C
         DataField       =   "Banda"
         DataSource      =   "datSecondaryRS"
         Height          =   315
         Left            =   -74760
         TabIndex        =   11
         Top             =   2520
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   327680
         Style           =   2
         ForeColor       =   16512
         ListField       =   "Descri��o"
         BoundColumn     =   "C�digo"
         Text            =   ""
      End
      Begin MSDBGrid.DBGrid grdDataGrid 
         Bindings        =   "frmDisco.frx":00AC
         Height          =   1305
         Left            =   -74760
         OleObjectBlob   =   "frmDisco.frx":022E
         TabIndex        =   12
         Top             =   840
         Width           =   7245
      End
      Begin MSComDlg.CommonDialog dlgCapa 
         Left            =   6555
         Top             =   3945
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   327680
      End
      Begin VB.Label Label5 
         Caption         =   "Listagem de can��es:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74730
         TabIndex        =   44
         Top             =   630
         Width           =   3315
      End
      Begin VB.Label lblLabels 
         Caption         =   "Emprestado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -74790
         TabIndex        =   38
         Top             =   660
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   -74625
         TabIndex        =   37
         Top             =   1770
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Para:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   -74625
         TabIndex        =   36
         Top             =   2145
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Endere�o:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   -74625
         TabIndex        =   35
         Top             =   2550
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Telefone:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   -74625
         TabIndex        =   34
         Top             =   2940
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Descri��o:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   30
         Top             =   765
         Width           =   960
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Capa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3090
         TabIndex        =   29
         Top             =   1965
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Gravadora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   28
         Top             =   1155
         Width           =   960
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "M�dia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   435
         TabIndex        =   27
         Top             =   1575
         Width           =   765
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt. Prod.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   75
         TabIndex        =   26
         Top             =   1965
         Width           =   1110
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Nota:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   25
         Top             =   2370
         Width           =   960
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Conserva��o:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   45
         TabIndex        =   24
         Top             =   2775
         Width           =   1185
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Coment�rio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   195
         TabIndex        =   23
         Top             =   3480
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Banda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   16
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Dura��o:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   15
         Top             =   2940
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Estilo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71025
         TabIndex        =   14
         Top             =   2295
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Coment�rio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74745
         TabIndex        =   13
         Top             =   3615
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         Height          =   2505
         Left            =   -74835
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   6525
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000010&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   2505
         Left            =   -74745
         Shape           =   4  'Rounded Rectangle
         Top             =   1275
         Width           =   6525
      End
   End
   Begin VB.CommandButton cmdEditar 
      Height          =   630
      Left            =   7995
      Picture         =   "frmDisco.frx":0F44
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1305
      Width           =   960
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   630
      Left            =   7995
      Picture         =   "frmDisco.frx":124E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   630
      Width           =   960
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   630
      Left            =   7995
      Picture         =   "frmDisco.frx":1558
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1995
      Width           =   960
   End
   Begin VB.CommandButton cmdCancela 
      Height          =   630
      Left            =   7995
      Picture         =   "frmDisco.frx":1862
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3375
      Width           =   960
   End
   Begin VB.CommandButton cmdUpdate 
      Height          =   630
      Left            =   7995
      Picture         =   "frmDisco.frx":1B6C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2685
      Width           =   960
   End
   Begin VB.CommandButton cmdClose 
      Height          =   630
      Left            =   7995
      Picture         =   "frmDisco.frx":1E76
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4065
      Width           =   960
   End
   Begin VB.Data datSecondaryRS 
      Connect         =   "Access"
      DatabaseName    =   "Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "M�sicas"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtC�digo 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      DataField       =   "C�digo"
      DataSource      =   "datPrimaryRS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   285
      Left            =   3225
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   210
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "C�digo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2550
      TabIndex        =   0
      Top             =   210
      Width           =   1815
   End
End
Attribute VB_Name = "frmDiscos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Editando As Boolean

Private Sub chkEmprestado_Click()
    On Error GoTo Fim
    If Screen.ActiveControl.Name <> Me.chkEmprestado.Name Then Exit Sub
    If Me.chkEmprestado Then
        Me.txtEmpData.Enabled = True
        Me.txtEmpEndere�o.Locked = False
        Me.txtEmpNome.Locked = False
        Me.txtEmpTelefone.Enabled = True
        Me.txtEmpData.SetFocus
    Else
        Me.txtEmpData.PromptInclude = False
        Me.txtEmpTelefone.PromptInclude = False
        Me.txtEmpData.Text = ""
        Me.txtEmpEndere�o.Text = ""
        Me.txtEmpNome.Text = ""
        Me.txtEmpTelefone.Text = ""
        Me.txtEmpData.PromptInclude = True
        Me.txtEmpTelefone.PromptInclude = True

        Me.txtEmpData.Enabled = False
        Me.txtEmpEndere�o.Locked = True
        Me.txtEmpNome.Locked = True
        Me.txtEmpTelefone.Enabled = False
    End If
Fim:
End Sub

Private Sub cmdAdd_Click()
    Trava False
    datPrimaryRS.Recordset.AddNew
    Me.cmdAdd.Enabled = False
    Me.cmdCancela.Enabled = True
    Me.cmdClose.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.cmdUpdate.Enabled = True
    Me.cmdEditar.Enabled = False
    Me.sstGeral.Tab = 0
    Me.txtDescri��o.SetFocus
    Editando = True
    Me.datPrimaryRS.Visible = False
End Sub

Private Sub cmdAddFoto_Click()
    On Error Resume Next
    With dlgCapa
        .CancelError = True
        .DialogTitle = "Capa do disco ..."
        .Filter = "Arquivos Gr�ficos (*.bmp;*.jpg;*.wmf;*.gif)|*.bmp;*.jpg;*.wmf;*.gif"
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames
        .ShowOpen
    End With
    If Err = 32755 Then Exit Sub
    Me.Figura.Picture = LoadPicture(Me.dlgCapa.filename)
End Sub

Private Sub cmdCancela_Click()
    Dim MSG As String
    MSG = "Deseja desfazer as altera��es ?"
    If MsgBox(MSG, vbYesNo, Caption) = vbNo Then Exit Sub
    Trava True
    datPrimaryRS.Recordset.CancelUpdate
    If Me.datPrimaryRS.Recordset.RecordCount = 0 Then
        Me.cmdAdd.Enabled = True
        Me.cmdCancela.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = False
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = False
    Else
        Me.cmdAdd.Enabled = True
        Me.cmdCancela.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = True
    End If
    Me.sstGeral.Tab = 0
    Me.txtDescri��o.SetFocus
    Editando = False
    Me.datPrimaryRS.Visible = True
End Sub

Private Sub cmdDelete_Click()
    If Me.datPrimaryRS.Recordset.RecordCount = 0 Then Exit Sub
    Dim MSG As String
    Dim Cod As Long
    MSG = "Deseja excluir o disco " & _
        Me.datPrimaryRS.Recordset("C�digo") & " " & _
        "e todas as m�sicas que pertencem a ele ?"
    If MsgBox(MSG, vbYesNo, Caption) = vbNo Then Exit Sub
  With datPrimaryRS.Recordset
    Cod = CLng(.Fields("C�digo").Value)
    .Delete
    If .RecordCount <> 0 Then .MoveNext
    If .EOF Then .MoveLast
    If .RecordCount = 0 Then Me.datPrimaryRS.Refresh
  End With
  Me.sstGeral.Tab = 0
  Me.txtDescri��o.SetFocus
End Sub

Private Sub cmdEditar_Click()
    Trava False
    datPrimaryRS.Recordset.Edit
    Me.cmdAdd.Enabled = False
    Me.cmdCancela.Enabled = True
    Me.cmdClose.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.cmdUpdate.Enabled = True
    Me.cmdEditar.Enabled = False
    Me.sstGeral.Tab = 0
    Me.txtDescri��o.SetFocus
    Editando = True
    Me.datPrimaryRS.Visible = False
End Sub

Private Sub cmdUpdate_Click()
    If Not Valida Then Exit Sub
    Me.datPrimaryRS.UpdateRecord
    datPrimaryRS.Recordset.Bookmark = datPrimaryRS.Recordset.LastModified
    Trava True
    If Me.datPrimaryRS.Recordset.RecordCount = 0 Then
        Me.cmdAdd.Enabled = True
        Me.cmdCancela.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = False
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = False
    Else
        Me.cmdAdd.Enabled = True
        Me.cmdCancela.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = True
    End If
    Me.sstGeral.Tab = 0
    Me.txtDescri��o.SetFocus
    Editando = False
    Me.datPrimaryRS.Visible = True
End Sub

Private Sub cmdClose_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub datPrimaryRS_Error(DataErr As Integer, Response As Integer)
  MsgBox "Ocorreu o erro n� " & DataErr & ": " & Error$(DataErr) & _
    vbCrLf & vbCrLf & "Entre em contato com o fornecedor."
  Response = 0
End Sub

Private Sub datPrimaryRS_Reposition()
    If Me.datPrimaryRS.Recordset.RecordCount = 0 Then
        Me.cmdAdd.Enabled = True
        Me.cmdCancela.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = False
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = False
    Else
        Me.cmdAdd.Enabled = True
        Me.cmdCancela.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = True
     End If
    On Error Resume Next
    datSecondaryRS.RecordSource = "select * from [M�sicas] where [Disco]=" & datPrimaryRS.Recordset![C�digo] & " Order by [N�mero]"
    datSecondaryRS.Refresh
    datPrimaryRS.Caption = "Registro: " & Format((datPrimaryRS.Recordset.AbsolutePosition + 1), "000000")
End Sub

''Private Sub datSecondaryRS_Validate(Action As Integer, Save As Integer)
''  Select Case Action
''    Case vbDataActionAddNew
''    Case vbDataActionUpdate
''        Me.datSecondaryRS.Recordset("Disco") = _
''            Me.datPrimaryRS.Recordset("C�digo")
''    Case vbDataActionDelete
''        If MsgBox("Deletar a can��o " & Chr(34) & _
''            Me.datSecondaryRS.Recordset("Descri��o") & _
''            Chr(34) & " ?", vbYesNo, Caption) = vbNo Then
''                Action = vbDataActionCancel
''        End If
''    Case vbDataActionFind
''    Case vbDataActionBookmark
''    Case vbDataActionClose
''      Screen.MousePointer = vbDefault
''  End Select
''
''End Sub


Private Sub datSecondaryRS_Validate(Action As Integer, Save As Integer)
    Select Case Action
        Case vbDataActionAddNew
            If Me.datPrimaryRS.Recordset.RecordCount = 0 Then
                MsgBox "N�o existem disco cadastrados.", , Caption
                Action = vbDataActionCancel
            End If
        Case vbDataActionUpdate
            Me.datSecondaryRS.Recordset("Disco") = CLng(Trim(Me.txtC�digo.Text))
    End Select
        
End Sub

Private Sub dbcBanda_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbLeftButton Then Exit Sub
    Me.dbcBanda.BoundText = ""
    If Area = 0 Or Area = 1 Then
        Me.Data3.Refresh
        Me.dbcBanda.ReFill
    End If
End Sub

Private Sub dbcEstilo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbLeftButton Then Exit Sub
    Me.dbcEstilo.BoundText = ""
    If Area = 0 Or Area = 1 Then
        Me.Data4.Refresh
        Me.dbcEstilo.ReFill
    End If
End Sub

Private Sub dbcGravadora_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbLeftButton Then Exit Sub
    Me.dbcGravadora.BoundText = ""
    If Area = 0 Or Area = 1 Then
        Me.Data1.Refresh
        Me.dbcGravadora.ReFill
    End If
End Sub

Private Sub dbcM�dia_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbLeftButton Then Exit Sub
    Me.dbcM�dia.BoundText = ""
    If Area = 0 Or Area = 1 Then
        Me.Data2.Refresh
        Me.dbcM�dia.ReFill
    End If
End Sub

Private Sub Form_Load()
    Trava True
    datPrimaryRS.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Editando Then
        MsgBox "Existem transa��es de dados na tela de discos " & _
            "que ainda est�o pendentes." & vbCrLf & _
            "Salve-as primeiro.", vbCritical, Caption
            Cancel = True
    End If
End Sub

Sub Trava(Travar As Boolean)
    Me.txtComent�rio.Locked = Travar
    Me.txtConserva��o.Locked = Travar
    Me.txtData.Enabled = Not Travar
    Me.txtDescri��o.Locked = Travar
    If Me.chkEmprestado Then
        Me.txtEmpData.Enabled = Not Travar
        Me.txtEmpEndere�o.Locked = Travar
        Me.txtEmpNome.Locked = Travar
        Me.txtEmpTelefone.Enabled = Not Travar
    Else
        Me.txtEmpData.Enabled = False
        Me.txtEmpEndere�o.Locked = True
        Me.txtEmpNome.Locked = True
        Me.txtEmpTelefone.Enabled = False
    End If
    Me.txtNota.Locked = Travar
    Me.updConserva��o.Enabled = Not Travar
    Me.updNota.Enabled = Not Travar
    Me.cmdAddFoto.Enabled = Not Travar
    Me.dbcGravadora.Enabled = Not Travar
    Me.dbcM�dia.Enabled = Not Travar
    Me.chkEmprestado.Enabled = Not Travar
    Me.sstGeral.TabVisible(2) = Travar
End Sub


Function Valida() As Boolean
    Valida = True
    If Trim(Me.txtDescri��o.Text) = "" Then
        MsgBox "A descri��o do disco � requerida.", vbCritical, Caption
        Valida = False
        Me.sstGeral.Tab = 0
        Me.txtDescri��o.SetFocus
        Exit Function
    End If
    If Trim(Me.dbcGravadora.BoundText) = "" Then
        MsgBox "A gravadora do disco � requerida.", vbCritical, Caption
        Valida = False
        Me.sstGeral.Tab = 0
        Me.dbcGravadora.SetFocus
        Exit Function
    End If
    If Trim(Me.dbcM�dia.BoundText) = "" Then
        MsgBox "A m�dia do disco � requerida.", vbCritical, Caption
        Valida = False
        Me.sstGeral.Tab = 0
        Me.dbcM�dia.SetFocus
        Exit Function
    End If
    If Me.chkEmprestado And Trim(Me.txtEmpNome.Text) = "" Then
        MsgBox "A pessoa para quem a m�dia foi emprestada � requerida.", vbCritical, Caption
        Valida = False
        Me.sstGeral.Tab = 1
        Me.txtEmpNome.SetFocus
        Exit Function
    End If
    If Not IsDate(Me.txtData.Text) And Trim(Me.txtData.ClipText) <> "" Then
        MsgBox "A data de produ��o n�o foi digitada corretamente.", vbCritical, Caption
        Valida = False
        Me.sstGeral.Tab = 0
        Me.txtData.SetFocus
        Exit Function
    End If
    If Not IsDate(Me.txtEmpData.Text) And Me.txtEmpData.Enabled And Trim(Me.txtEmpData.ClipText) <> "" Then
        MsgBox "A data de empr�stimo da m�dia n�o foi digitada corretamente.", vbCritical, Caption
        Valida = False
        Me.sstGeral.Tab = 1
        Me.txtEmpData.SetFocus
        Exit Function
    End If
    Dim RC As Recordset
    Set RC = Me.datPrimaryRS.Recordset.Clone
    RC.FindFirst "Descri��o = '" & Trim(Me.txtDescri��o.Text) & "'"
    If Not RC.NoMatch Then
        If Me.datPrimaryRS.EditMode = dbEditInProgress Then
            If CLng(RC("C�digo")) <> CLng(Me.datPrimaryRS.Recordset("C�digo")) Then
                MsgBox "Banda j� cadastrada.", vbInformation, Caption
                Valida = False
                RC.Close
                Exit Function
            Else
                RC.Close
            End If
        Else
            MsgBox "Banda j� cadastrada.", vbInformation, Caption
            Valida = False
            RC.Close
            Exit Function
        End If
    Else
        RC.Close
    End If
End Function


Private Sub grdDataGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Select Case ColIndex
        Case 0
            If Trim(Me.grdDataGrid.Columns(0)) = "" Then
                    MsgBox "O nome da can��o n�o foi informado."
                    Cancel = True
                    Exit Sub
            End If
            If Trim(Me.txtC�digo.Text) = "" Then
                Cancel = True
                Exit Sub
            End If
        Case 2
            If Not IsNumeric(Me.grdDataGrid.Columns(2)) Then
                    MsgBox "Para a nota s�o apenas valores num�ricos de 0 a 5."
                    Cancel = True
                    Exit Sub
            End If
            If Me.grdDataGrid.Columns(2) > 5 Or _
                Me.grdDataGrid.Columns(2) < 0 Then
                    MsgBox "Para a nota s�o aceito valores de 0 a 5."
                    Cancel = True
                    Exit Sub
            End If
            If Trim(Me.txtC�digo.Text) = "" Then
                Cancel = True
                Exit Sub
            End If
    End Select
End Sub

Private Sub grdDataGrid_BeforeDelete(Cancel As Integer)
    If MsgBox("Deletar a can��o  ?", vbYesNo _
        , Caption) = vbNo Then Cancel = True
End Sub

Private Sub grdDataGrid_BeforeUpdate(Cancel As Integer)
    If Trim(Me.grdDataGrid.Columns(0)) = "" Then
        MsgBox "O nome da can��o n�o foi informado."
        Cancel = True
        Exit Sub
    End If
    If Trim(Me.dbcBanda.BoundText) = "" Then
        MsgBox "A banda n�o foi informada.", , Caption
        Cancel = True
        Exit Sub
    End If
    If Trim(Me.dbcEstilo.BoundText) = "" Then
        MsgBox "O estilo n�o foi informado.", , Caption
        Cancel = True
        Exit Sub
    End If
    If Trim(Me.txtDura��o.ClipText) <> "" And Not IsDate(Me.txtDura��o.Text) Then
        MsgBox "A dura��o da can��o n�o est� correta.", , Caption
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub grdDataGrid_Error(ByVal DataError As Integer, Response As Integer)
    Response = 0
End Sub

