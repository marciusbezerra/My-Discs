VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmNacionalidades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nacionalidades"
   ClientHeight    =   5130
   ClientLeft      =   1095
   ClientTop       =   405
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Fechar"
      Height          =   810
      Left            =   5100
      Picture         =   "frmNacio.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2370
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Salvar"
      Height          =   810
      Left            =   5100
      Picture         =   "frmNacio.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   630
      Width           =   975
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "&Cancelar"
      Height          =   810
      Left            =   5100
      Picture         =   "frmNacio.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1500
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Deletar"
      Height          =   810
      Left            =   4065
      Picture         =   "frmNacio.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2370
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Adicionar"
      Height          =   810
      Left            =   4065
      Picture         =   "frmNacio.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   645
      Width           =   975
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   810
      Left            =   4065
      Picture         =   "frmNacio.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1500
      Width           =   975
   End
   Begin VB.TextBox txtFields 
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
      Index           =   0
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Descri��o"
      DataSource      =   "datPrimaryRS"
      ForeColor       =   &H00004080&
      Height          =   315
      Index           =   1
      Left            =   975
      MaxLength       =   50
      TabIndex        =   1
      Top             =   3285
      Width           =   5130
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Coment�rio"
      DataSource      =   "datPrimaryRS"
      ForeColor       =   &H00004080&
      Height          =   915
      Index           =   2
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4035
      Width           =   5970
   End
   Begin VB.Data datPrimaryRS 
      Caption         =   " "
      Connect         =   "Access"
      DatabaseName    =   "Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from [Nacionalidades] Order by [Descri��o]"
      Top             =   180
      Visible         =   0   'False
      Width           =   1320
   End
   Begin MSDBCtls.DBList DBList1 
      Bindings        =   "frmNacio.frx":123C
      Height          =   2400
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4233
      _Version        =   327680
      ForeColor       =   16512
      ListField       =   "Descri��o"
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
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Edi��o:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   3315
      Width           =   1815
   End
   Begin VB.Label lblLabels 
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
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   3780
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lista de estilos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   540
      Width           =   1350
   End
End
Attribute VB_Name = "frmNacionalidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Editando As Boolean


Private Sub cmdAdd_Click()
    Trava False
    datPrimaryRS.Recordset.AddNew
    Me.cmdAdd.Enabled = False
    Me.cmdCancela.Enabled = True
    Me.cmdClose.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.cmdUpdate.Enabled = True
    Me.cmdEditar.Enabled = False
    Me.txtFields(1).SetFocus
    Editando = True
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
    Me.txtFields(1).SetFocus
    Editando = False
End Sub

Private Sub cmdDelete_Click()
    If Me.datPrimaryRS.Recordset.RecordCount = 0 Then Exit Sub
    Dim MSG As String
    Dim Cod As Long
    MSG = "Deseja excluir a nacionalidade " & _
        Me.datPrimaryRS.Recordset("C�digo") & " " & _
        "e todas as bandas que pertencem a ela ?"
    If MsgBox(MSG, vbYesNo, Caption) = vbNo Then Exit Sub
  With datPrimaryRS.Recordset
    Cod = CLng(.Fields("C�digo").Value)
    .Delete
    If .RecordCount <> 0 Then .MoveNext
    If .EOF Then .MoveLast
    If .RecordCount = 0 Then Me.datPrimaryRS.Refresh
  End With
  Me.DBList1.ReFill
  Me.txtFields(1).SetFocus
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
    Me.txtFields(1).SetFocus
    Editando = True
End Sub

Private Sub cmdUpdate_Click()
    If Not Valida Then Exit Sub
    datPrimaryRS.UpdateRecord
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
    Me.DBList1.ReFill
    Me.DBList1.Refresh
    Me.txtFields(1).SetFocus
    Editando = False
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

  'datPrimaryRS.Caption = "R: " & (datPrimaryRS.Recordset.AbsolutePosition + 1)
End Sub


Private Sub DBList1_Click()
    If Me.datPrimaryRS.Recordset.RecordCount = 0 Then Exit Sub
    Me.datPrimaryRS.Recordset.Bookmark = Me.DBList1.SelectedItem
End Sub

Private Sub Form_Load()
    Trava True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Editando Then
        MsgBox "Existem transa��es de dados na tela de nacionalidades " & _
            "que ainda est�o pendentes." & vbCrLf & _
            "Salve-as primeiro.", vbCritical, Caption
            Cancel = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Sub Trava(Travar As Boolean)
    Dim I As Integer
    For I = 1 To txtFields.Count - 1
        txtFields(I).Locked = Travar
    Next
End Sub


Function Valida() As Boolean
    Valida = True
    If Trim(Me.txtFields(1).Text) = "" Then
        MsgBox "A descri��o da nacionalidade � requerida.", vbCritical, Caption
        Valida = False
        Me.txtFields(1).SetFocus
        Exit Function
    End If
    Dim RC As Recordset
    Set RC = Me.datPrimaryRS.Recordset.Clone
    RC.FindFirst "Descri��o = '" & Trim(Me.txtFields(1).Text) & "'"
    If Not RC.NoMatch Then
        If Me.datPrimaryRS.EditMode = dbEditInProgress Then
            If CLng(RC("C�digo")) <> CLng(Me.datPrimaryRS.Recordset("C�digo")) Then
                MsgBox "Nacionalidade j� cadastrada.", vbInformation, Caption
                Valida = False
                RC.Close
                Exit Function
            Else
                RC.Close
            End If
        Else
            MsgBox "Nacionalidade j� cadastrada.", vbInformation, Caption
            Valida = False
            RC.Close
            Exit Function
        End If
    Else
        RC.Close
    End If
End Function


