VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmEstilos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estilos"
   ClientHeight    =   5160
   ClientLeft      =   1095
   ClientTop       =   405
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   810
      Left            =   4080
      Picture         =   "frmEstil.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Adicionar"
      Height          =   810
      Left            =   4080
      Picture         =   "frmEstil.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   585
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Deletar"
      Height          =   810
      Left            =   4080
      Picture         =   "frmEstil.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2310
      Width           =   975
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "&Cancelar"
      Height          =   810
      Left            =   5115
      Picture         =   "frmEstil.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Salvar"
      Height          =   810
      Left            =   5115
      Picture         =   "frmEstil.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   570
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Fechar"
      Height          =   810
      Left            =   5115
      Picture         =   "frmEstil.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2310
      Width           =   975
   End
   Begin MSDBCtls.DBList DBList1 
      Bindings        =   "frmEstil.frx":123C
      Height          =   2400
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4233
      _Version        =   327680
      ForeColor       =   16512
      ListField       =   "Descrição"
   End
   Begin VB.Data datPrimaryRS 
      Caption         =   " "
      Connect         =   "Access"
      DatabaseName    =   "C:\Dados\Prog\VBasic\Discos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from [Estilos] Order by [Descrição]"
      Top             =   120
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Comentário"
      DataSource      =   "datPrimaryRS"
      ForeColor       =   &H00004080&
      Height          =   915
      Index           =   2
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3960
      Width           =   6000
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Descrição"
      DataSource      =   "datPrimaryRS"
      ForeColor       =   &H00004080&
      Height          =   315
      Index           =   1
      Left            =   960
      MaxLength       =   50
      TabIndex        =   3
      Top             =   3225
      Width           =   5145
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      DataField       =   "Código"
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   1935
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
      TabIndex        =   6
      Top             =   480
      Width           =   1350
   End
   Begin VB.Label lblLabels 
      Caption         =   "Comentário:"
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
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Edição:"
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
      TabIndex        =   2
      Top             =   3255
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Código:"
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
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmEstilos"
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
    MSG = "Deseja desfazer as alterações ?"
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
    MSG = "Deseja excluir o estilo " & _
        Me.datPrimaryRS.Recordset("Código") & " " & _
        "e todas as músicas que pertencem a ele ?"
    If MsgBox(MSG, vbYesNo, Caption) = vbNo Then Exit Sub
  With datPrimaryRS.Recordset
    Cod = CLng(.Fields("Código").Value)
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
  MsgBox "Ocorreu o erro nº " & DataErr & ": " & Error$(DataErr) & _
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

  datPrimaryRS.Caption = "R: " & (datPrimaryRS.Recordset.AbsolutePosition + 1)
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
        MsgBox "Existem transações de dados na tela de estilos " & _
            "que ainda estão pendentes." & vbCrLf & _
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
        MsgBox "A descrição do estilo é requerida.", vbCritical, Caption
        Valida = False
        Me.txtFields(1).SetFocus
        Exit Function
    End If
    Dim RC As Recordset
    Set RC = Me.datPrimaryRS.Recordset.Clone
    RC.FindFirst "Descrição = '" & Trim(Me.txtFields(1).Text) & "'"
    If Not RC.NoMatch Then
        If Me.datPrimaryRS.EditMode = dbEditInProgress Then
            If CLng(RC("Código")) <> CLng(Me.datPrimaryRS.Recordset("Código")) Then
                MsgBox "Estilo já cadastrado.", vbInformation, Caption
                Valida = False
                RC.Close
                Exit Function
            Else
                RC.Close
            End If
        Else
            MsgBox "Estilo já cadastrado.", vbInformation, Caption
            Valida = False
            RC.Close
            Exit Function
        End If
    Else
        RC.Close
    End If
End Function
