VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_cadastro 
   Caption         =   "Cadastro de usuários"
   ClientHeight    =   3800
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7440
   OleObjectBlob   =   "frm_cadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_cadastrar_Click()
    Sheets("contas_login").Select
    
    If txt_usuario = "" Or txt_email = "" Or txt_senha = "" Or txt_rsenha = "" Then
        Sheets("home").Select
        alerta = MsgBox("Preencha todos os campos", vbExclamation + vbOKOnly, "ATENÇÃO")
        
    ElseIf txt_senha <> txt_rsenha Then
        Sheets("home").Select
        alerta = MsgBox("Senhas não conferem", vbExclamation + vbOKOnly, "ATENÇÃO")
        txt_senha.Value = ""
        txt_rsenha.Value = ""
        txt_senha.SetFocus
        
    ElseIf Range("B20").Value = "" Then
    
        cont_linhas = 10
        
        While Range("B" & cont_linhas).Value <> ""
            cont_linhas = cont_linhas + 1
        Wend
        
        If cont_linhas > 10 Then
            id_usuario = CInt(Range("B" & cont_linhas - 1).Value) + 1
        Else
            id_usuario = 1
        End If
        
        Range("B" & cont_linhas).Value = id_usuario
        Range("C" & cont_linhas).Value = txt_usuario
        Range("D" & cont_linhas).Value = txt_email
        Range("E" & cont_linhas).Value = txt_senha
        Range("F" & cont_linhas).Value = "ATIVO"
        
        Sheets("home").Select
        alerta = MsgBox("Usuário cadastrado" & vbNewLine & "Deseja cadastrar outro usuário?", vbQuestion + vbYesNo, "ATENÇÃO")
        If alerta = vbYes Then
            txt_usuario = ""
            txt_email = ""
            txt_senha = ""
            txt_rsenha = ""
            txt_usuario.SetFocus
        Else
            Unload Me
        End If
        
    Else
        Sheets("home").Select
        alerta = MsgBox("Banco de cadastros cheio.", vbExclamation + vbOKOnly, "ATENÇÃO")
        
    End If
End Sub

Private Sub chk_visualizar_Click()
    If chk_visualizar.Value Then
        txt_senha.PasswordChar = ""
        txt_rsenha.PasswordChar = ""
    Else
        txt_senha.PasswordChar = "•"
        txt_rsenha.PasswordChar = "•"
    End If
End Sub
