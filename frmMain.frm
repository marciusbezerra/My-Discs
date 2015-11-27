VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   ClientHeight    =   6105
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   7290
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuCadastros 
      Caption         =   "&Cadastros"
      Begin VB.Menu mnuDiscosEMusicas 
         Caption         =   "&Disco e Músicas"
      End
      Begin VB.Menu mnuBandas 
         Caption         =   "&Bandas"
      End
      Begin VB.Menu mnuEstlios 
         Caption         =   "&Estilos"
      End
      Begin VB.Menu mnuGravadoras 
         Caption         =   "&Gravadoras"
      End
      Begin VB.Menu mnuMidias 
         Caption         =   "&Mídias"
      End
      Begin VB.Menu mnuNacionalidades 
         Caption         =   "&Nacionalidades"
      End
      Begin VB.Menu mnuRel 
         Caption         =   "&Relatório"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub mnuBandas_Click()
    frmBandas.Show
End Sub

Private Sub mnuDiscosEMusicas_Click()
    frmDiscos.Show
End Sub

Private Sub mnuEstlios_Click()
    frmEstilos.Show
End Sub

Private Sub mnuGravadoras_Click()
    frmGravadoras.Show
End Sub

Private Sub mnuMidias_Click()
    frmMídias.Show
End Sub

Private Sub mnuNacionalidades_Click()
    frmNacionalidades.Show
End Sub

