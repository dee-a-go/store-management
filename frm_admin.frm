VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_admin 
   Caption         =   "Menu do Administrador"
   ClientHeight    =   1980
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3360
   OleObjectBlob   =   "frm_admin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_cadastrar_usuario_Click()
    frm_cadastro.Show
End Sub

Private Sub btn_gerenciar_Click()
    frm_gerenciamento.Show
End Sub

