VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_login 
   Caption         =   "ATIVIDADE ADS (P1) - PROG. MICROINFORMÁTICA"
   ClientHeight    =   4884
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   7860
   OleObjectBlob   =   "frm_login.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_entrar_Click()
    Planilha2.Visible = xlSheetVisible
    Worksheets("CONTAS_LOGIN").Select
    cont_linhas = 10
    
    If txt_usuario = "admin" And txt_senha = "admin" Then
        Call mostrar_planilhas_adm
        Worksheets("HOME").Select
        Unload Me
        alerta = MsgBox("Administrador logado com sucesso.", vbOKOnly + vbExclamation, "BEM-VINDO!")
        Exit Sub
    End If
    
    While Range("B" & cont_linhas) <> ""
    
        usuario = Range("C" & cont_linhas).Value
        email = Range("D" & cont_linhas).Value
        senha = CStr(Range("E" & cont_linhas).Value)
        
        If (txt_usuario = usuario Or txt_usuario = email) And txt_senha = senha Then
        
            status_usuario = Range("F" & cont_linhas).Value
            
            If status_usuario = "INATIVO" Then
                Worksheets("HOME").Select
                Planilha2.Visible = xlSheetVeryHidden
                alerta = MsgBox("Usuário não autorizado." & vbNewLine & "Por favor, entre em contato com um de nossos atendentes.", vbOKOnly + vbExclamation, "ERRO.")
                Exit Sub
                
            Else
                
                Worksheets("HOME").Select
                Call mostrar_planilhas_usuario
                Unload Me
                alerta = MsgBox("Usuário logado com sucesso.", vbOKOnly + vbExclamation, "BEM-VINDO!")
                Exit Sub
                
            End If
            
        End If
        
        cont_linhas = cont_linhas + 1
        
    Wend

    Planilha2.Visible = xlSheetVeryHidden
    alerta = MsgBox("Login e/ou senha incorreto(s).", vbOKOnly + vbExclamation, "ERRO.")
End Sub
