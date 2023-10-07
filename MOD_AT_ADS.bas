Attribute VB_Name = "MOD_AT_ADS"
Public cont_linhas, id_usuario, linha_anterior, num_socio, total As Integer
Public usuario, email, senha, status_usuario, genero_filme, filme As String

Function ocultar_planilhas()
    Application.Visible = True
    Plan2.Visible = xlSheetVeryHidden
    Plan3.Visible = xlSheetVeryHidden
    Plan4.Visible = xlSheetVeryHidden
    Plan5.Visible = xlSheetVeryHidden
    Planilha2.Visible = xlSheetVeryHidden
    Application.StatusBar = "N�o logado"
End Function

Function mostrar_planilhas_usuario()
    Application.Visible = True
    Plan2.Visible = xlSheetVeryHidden
    Plan3.Visible = xlSheetVeryHidden
    Plan4.Visible = xlSheetVeryHidden
    Plan5.Visible = xlSheetVeryHidden
    Planilha2.Visible = xlSheetVeryHidden
    Application.StatusBar = "Usu�rio logado"
End Function

Function mostrar_planilhas_adm()
    Application.Visible = True
    Plan2.Visible = xlSheetVisible
    Plan3.Visible = xlSheetVisible
    Plan4.Visible = xlSheetVisible
    Plan5.Visible = xlSheetVisible
    Planilha2.Visible = xlSheetVisible
    Application.StatusBar = "Admin logado"
End Function

Function sair_planilha()
    alerta1 = MsgBox("Deseja sair?", vbQuestion + vbYesNo, "ATEN��O!")
    If alerta1 = vbYes Then
        alerta2 = MsgBox("Deseja salvar?", vbQuestion + vbYesNo, "ATEN��O!")
        If alerta2 = vbYes Then
            ActiveWorkbook.Save
        End If
        Application.Quit
    End If
End Function

Function chamar_locacao()
    If Application.StatusBar = "N�o logado" Then
        alerta = MsgBox("� necess�rio logar para usar esta fun��o.", vbExclamation + vbOKOnly, "ATEN��O!")
    ElseIf Application.StatusBar = "Admin logado" Then
        alerta = MsgBox("� necess�rio logar como usu�rio para usar esta fun��o.", vbExclamation + vbOKOnly, "ATEN��O!")
    Else
        frm_locacao.Show
    End If
End Function

Function chamar_login()
    If Application.StatusBar = "N�o logado" Then
        frm_login.Show
    Else
        alerta = MsgBox("Voc� j� est� logado." & vbNewLine & "Deseja sair?", vbQuestion + vbYesNo, "ATEN��O!")
        If alerta <> vbNo Then
            Call ocultar_planilhas
            alerta = MsgBox("At� logo!", vbExclamation + vbOKOnly, "USU�RIO DESLOGADO.")
        End If
    End If
End Function

Function chamar_menu_adm()
    If Application.StatusBar = "Admin logado" Then
        frm_admin.Show
    Else
        alerta = MsgBox("Voc� precisa logar como administrador para usar esta funcionalidade!", vbExclamation + vbOKOnly, "ATEN��O!")
    End If
End Function
