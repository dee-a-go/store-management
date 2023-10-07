VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_gerenciamento 
   Caption         =   "Gerenciar contas de usuários"
   ClientHeight    =   3470
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5376
   OleObjectBlob   =   "frm_gerenciamento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_gerenciamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    cont_linhas = 10
    Sheets("contas_login").Select
    cmb_usuario.Clear
    
    While Range("B" & cont_linhas).Value <> ""
        usuario = Range("C" & cont_linhas).Value
        cmb_usuario.AddItem (usuario)
        cont_linhas = cont_linhas + 1
    Wend
    
    checkbox_status.Value = False
    lbl_status.Caption = "Status do usuário:"
    lbl_status.ForeColor = &H0&
    Sheets("home").Select
End Sub

Private Sub cmb_usuario_Change()
    cont_linhas = 10
    Sheets("contas_login").Select
    
    While Range("B" & cont_linhas).Value <> ""
        usuario = Range("C" & cont_linhas).Value

        If usuario = cmb_usuario.Value Then
            status_usuario = Range("F" & cont_linhas).Value
            email = Range("D" & cont_linhas).Value
            senha = CStr(Range("E" & cont_linhas).Value)
            
            lbl_email.Caption = "Email: " & email
            lbl_senha.Caption = "Senha: " & senha
            
            If status_usuario = "ATIVO" Then
                checkbox_status.Value = True
                lbl_status.Caption = "Status do usuário: Ativo"
                lbl_status.ForeColor = &H8000&
            ElseIf status_usuario = "INATIVO" Then
                checkbox_status.Value = False
                lbl_status.Caption = "Status do usuário: Inativo"
                lbl_status.ForeColor = &HFF&
            End If
            
            Sheets("home").Select
            Exit Sub
        End If
        
        cont_linhas = cont_linhas + 1
    Wend
    
End Sub

Private Sub checkbox_status_Click()
    If cmb_usuario.Value = "" Then
        alerta = MsgBox("Nenhum usuário selecionado.", vbExclamation + vbOKOnly, "ALERTA")
        Exit Sub
    End If
End Sub

Private Sub btn_salvar_Click()
    If cmb_usuario.Value = "" Then
        alerta = MsgBox("Nenhum usuário selecionado.", vbExclamation + vbOKOnly, "ALERTA")
        Exit Sub
    End If
    
    cont_linhas = 10
    Sheets("contas_login").Select
    
    While Range("B" & cont_linhas).Value <> ""
        usuario = Range("C" & cont_linhas).Value
        status_usuario = Range("F" & cont_linhas).Value
        
        If usuario = cmb_usuario.Value Then
            If checkbox_status.Value Then
                Range("F" & cont_linhas).Value = "ATIVO"
                lbl_status.Caption = "Status do usuário: Ativo"
                lbl_status.ForeColor = &H8000&
                Sheets("home").Select
                Exit Sub
            Else
                Range("F" & cont_linhas).Value = "INATIVO"
                lbl_status.Caption = "Status do usuário: Inativo"
                lbl_status.ForeColor = &HFF&
                Sheets("home").Select
                Exit Sub
            End If
        End If
        cont_linhas = cont_linhas + 1
    Wend
End Sub

Private Sub btn_apagar_conta_Click()
    If cmb_usuario.Value = "" Then
        alerta = MsgBox("Nenhum usuário selecionado.", vbExclamation + vbOKOnly, "ALERTA")
        Exit Sub
    End If
    
    alerta = MsgBox("Tem certeza que deseja apagar permanentemente os dados do usuário selecionado?", vbQuestion + vbYesNo, "AVISO")
    If alerta = vbYes Then
        cont_linhas = 10
        Sheets("contas_login").Select
        
        While Range("B" & cont_linhas).Value <> ""
            usuario = Range("C" & cont_linhas).Value
            If usuario = cmb_usuario.Value Then
                Range("B" & cont_linhas & ":F" & cont_linhas).Value = ""
                
                While Range("B" & (cont_linhas + 1)).Value <> ""
                    Range("B" & cont_linhas & ":F" & cont_linhas).Value = Range("B" & cont_linhas + 1 & ":F" & cont_linhas + 1).Value
                    cont_linhas = cont_linhas + 1
                Wend
                
                Range("B" & cont_linhas & ":F" & cont_linhas).Value = ""
                Sheets("home").Select
                alerta = MsgBox("Conta apagada com sucesso.", vbExclamation + vbOKOnly, "AVISO")
                Unload Me
                frm_gerenciamento.Show
                Exit Sub
            End If
            cont_linhas = cont_linhas + 1
        Wend
    End If
End Sub

