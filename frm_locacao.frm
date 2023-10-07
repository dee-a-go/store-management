VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_locacao 
   Caption         =   "ATIVIDADE ADS (P1) - PROG. MICROINFORMÁTICA"
   ClientHeight    =   8220.001
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   10968
   OleObjectBlob   =   "frm_locacao.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frm_locacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    txt_nome.Value = usuario
    txt_nome.SetFocus
End Sub

Private Sub txt_fone_Change()
    txt_fone.MaxLength = 14
    Select Case txt_fone.SelStart
        Case Is = 1
            txt_fone.SelStart = 0
            txt_fone.SelText = "("
            txt_fone.SelStart = 2
        Case Is = 3
            txt_fone.SelText = ")"
        Case Is = 9
            txt_fone.SelText = "-"
    End Select
End Sub

Private Sub op_acao_Click()
    cont_linhas = 4
    cmb_filmes.Clear
    lbl_valor = ""
    lbl_total.Caption = ""
    lbl_avaliacao = ""
    Plan3.Visible = xlSheetVisible
    Sheets("lista de filmes").Select
    
    While Range("B" & cont_linhas).Value <> ""
        genero_filme = Range("F" & cont_linhas).Value
        If genero_filme = "Ação" Then
            cmb_filmes.AddItem (Range("C" & cont_linhas).Value)
        End If
        cont_linhas = cont_linhas + 1
    Wend
    
    Sheets("home").Select
    Plan3.Visible = xlSheetVeryHidden
End Sub

Private Sub op_aventura_Click()
    cont_linhas = 4
    cmb_filmes.Clear
    lbl_valor = ""
    lbl_total.Caption = ""
    lbl_avaliacao = ""
    Plan3.Visible = xlSheetVisible
    Sheets("lista de filmes").Select
    
    While Range("B" & cont_linhas).Value <> ""
        genero_filme = Range("F" & cont_linhas).Value
        If genero_filme = "Aventura" Then
            cmb_filmes.AddItem (Range("C" & cont_linhas).Value)
        End If
        cont_linhas = cont_linhas + 1
    Wend
    
    Sheets("home").Select
    Plan3.Visible = xlSheetVeryHidden
End Sub

Private Sub op_comedia_Click()
    cont_linhas = 4
    cmb_filmes.Clear
    lbl_valor = ""
    lbl_total.Caption = ""
    lbl_avaliacao = ""
    Plan3.Visible = xlSheetVisible
    Sheets("lista de filmes").Select
    
    While Range("B" & cont_linhas).Value <> ""
        genero_filme = Range("F" & cont_linhas).Value
        If genero_filme = "Comédia" Then
            cmb_filmes.AddItem (Range("C" & cont_linhas).Value)
        End If
        cont_linhas = cont_linhas + 1
    Wend
    
    Sheets("home").Select
    Plan3.Visible = xlSheetVeryHidden
End Sub

Private Sub op_drama_Click()
    cont_linhas = 4
    cmb_filmes.Clear
    lbl_valor = ""
    lbl_total.Caption = ""
    lbl_avaliacao = ""
    Plan3.Visible = xlSheetVisible
    Sheets("lista de filmes").Select
    
    While Range("B" & cont_linhas).Value <> ""
        genero_filme = Range("F" & cont_linhas).Value
        If genero_filme = "Drama" Then
            cmb_filmes.AddItem (Range("C" & cont_linhas).Value)
        End If
        cont_linhas = cont_linhas + 1
    Wend
    
    Sheets("home").Select
    Plan3.Visible = xlSheetVeryHidden
End Sub

Private Sub op_romance_Click()
    cont_linhas = 4
    cmb_filmes.Clear
    lbl_valor = ""
    lbl_total.Caption = ""
    lbl_avaliacao = ""
    Plan3.Visible = xlSheetVisible
    Sheets("lista de filmes").Select
    
    While Range("B" & cont_linhas).Value <> ""
        genero_filme = Range("F" & cont_linhas).Value
        If genero_filme = "Romance" Then
            cmb_filmes.AddItem (Range("C" & cont_linhas).Value)
        End If
        cont_linhas = cont_linhas + 1
    Wend
    
    Sheets("home").Select
    Plan3.Visible = xlSheetVeryHidden
End Sub

Private Sub op_suspense_Click()
    cont_linhas = 4
    cmb_filmes.Clear
    lbl_valor = ""
    lbl_total.Caption = ""
    lbl_avaliacao = ""
    Plan3.Visible = xlSheetVisible
    Sheets("lista de filmes").Select
    
    While Range("B" & cont_linhas).Value <> ""
        genero_filme = Range("F" & cont_linhas).Value
        If genero_filme = "Suspense" Then
            cmb_filmes.AddItem (Range("C" & cont_linhas).Value)
        End If
        cont_linhas = cont_linhas + 1
    Wend
    
    Sheets("home").Select
    Plan3.Visible = xlSheetVeryHidden
End Sub

Private Sub op_terror_Click()
    cont_linhas = 4
    cmb_filmes.Clear
    lbl_valor = ""
    lbl_total.Caption = ""
    lbl_avaliacao = ""
    Plan3.Visible = xlSheetVisible
    Sheets("lista de filmes").Select
    
    While Range("B" & cont_linhas).Value <> ""
        genero_filme = Range("F" & cont_linhas).Value
        If genero_filme = "terror" Then
            cmb_filmes.AddItem (Range("C" & cont_linhas).Value)
        End If
        cont_linhas = cont_linhas + 1
    Wend
    
    Sheets("home").Select
    Plan3.Visible = xlSheetVeryHidden
End Sub

Private Sub cmb_filmes_Click()
    cont_linhas = 4
    Plan3.Visible = xlSheetVisible
    Sheets("lista de filmes").Select
    
    While Range("B" & cont_linhas).Value <> ""
        filme = Range("C" & cont_linhas).Value
        If filme = cmb_filmes.Value Then
            lbl_valor.Caption = Format(Range("G" & cont_linhas).Value, "R$ 0.00")
            lbl_avaliacao.Caption = Range("E" & cont_linhas).Value
            Call txt_qtde_Change
            Sheets("home").Select
            Plan3.Visible = xlSheetVeryHidden
            Exit Sub
        End If
        cont_linhas = cont_linhas + 1
    Wend
    Sheets("home").Select
    Plan3.Visible = xlSheetVeryHidden
End Sub

Private Sub op_infantil_Click()
    cont_linhas = 4
    cmb_filmes.Clear
    lbl_valor = ""
    lbl_total.Caption = ""
    lbl_avaliacao = ""
    Plan3.Visible = xlSheetVisible
    Sheets("lista de filmes").Select
    
    While Range("B" & cont_linhas).Value <> ""
        genero_filme = Range("F" & cont_linhas).Value
        If genero_filme = "Infantil" Then
            cmb_filmes.AddItem (Range("C" & cont_linhas).Value)
        End If
        cont_linhas = cont_linhas + 1
    Wend
    
    Sheets("home").Select
    Plan3.Visible = xlSheetVeryHidden
End Sub

Private Sub txt_qtde_Change()
    txt_qtde.MaxLength = 3
    If filme <> "" Then
        If txt_qtde <> "" And lbl_valor.Caption <> "" Then
            total = CInt(txt_qtde) * lbl_valor.Caption
            lbl_total.Caption = Format(total, "R$ 0.00")
        Else
            lbl_total.Caption = ""
        End If
    End If
End Sub

Private Sub btn_confirmar_Click()
    
    If txt_nome = "" Or txt_fone = "" Or cmb_filmes.Value = "" Or txt_qtde = "" Then
        alerta = MsgBox("Preencha todos os campos", vbExclamation + vbOKOnly, "ATENÇÃO")
    ElseIf CInt(txt_qtde) <= 0 Then
        alerta = MsgBox("Por favor, insira um número maior que 0 no campo de quantidade.", vbExclamation + vbOKOnly, "ATENÇÃO")
    Else
        alerta = MsgBox("Confirmar locação?", vbQuestion + vbYesNo, "ATENÇÃO")
        If alerta = vbYes Then
            Plan3.Visible = xlSheetVisible
            Sheets("lista de filmes").Select
            cont_linhas = 4
            num_socio = 1
            
            While Range("I" & cont_linhas) <> ""
                cont_linhas = cont_linhas + 1
                num_socio = num_socio + 1
            Wend
            
            Range("I" & cont_linhas) = num_socio
            Range("J" & cont_linhas) = UCase(txt_nome)
            Range("N" & cont_linhas) = txt_fone
            Range("O" & cont_linhas) = cmb_filmes.Value
            Range("P" & cont_linhas) = genero_filme
            Range("Q" & cont_linhas) = txt_qtde
            Range("R" & cont_linhas) = CInt(txt_qtde) * lbl_valor.Caption
            Range("S" & cont_linhas) = lbl_avaliacao.Caption
            Range("T" & cont_linhas) = Date
            Range("U" & cont_linhas) = Time
            
            Sheets("home").Select
            Plan3.Visible = xlSheetVeryHidden
            alerta = MsgBox("Locação confirmada." & vbNewLine & "Deseja fazer outra locação?", vbQuestion + vbYesNo, "ATENÇÃO")
            If alerta = vbYes Then
                Unload Me
                frm_locacao.Show
            Else
                Unload Me
            End If
        End If

    End If
End Sub

