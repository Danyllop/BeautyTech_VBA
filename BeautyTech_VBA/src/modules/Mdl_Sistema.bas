Attribute VB_Name = "Mdl_Sistema"
' ==============================================================================
' NOME DO MÓDULO: Mdl_Sistema
' OBJETIVO:       Gerenciar Regras de Negócio, Navegação e Segurança do Sistema
' AUTOR:          Danyllo Jonathas - LogicUp Solutions
' DATA:           Fevereiro/2026
' ==============================================================================
Option Explicit

' -------------------------------------------------------------------------
' Propósito: Roteamento de Páginas (Navegação do Menu Lateral)
' Parâmetros: Frm - O Formulário atual
'             IndexPagina - O número da aba na MultiPage (0 = Home, 1 = Agenda, etc)
'             NovoTitulo - O texto que vai aparecer na barra superior
' -------------------------------------------------------------------------
Public Sub NavegarPara(ByVal Frm As Object, ByVal IndexPagina As Integer, ByVal NovoTitulo As String)
    On Error Resume Next

    ' 1. Reset visual de todos os botões
    Call Mdl_UI_Efeitos.ResetarDestaque

    ' 2. Muda a página
    Frm.MultiPagMain.Value = IndexPagina

    ' 3. Atualiza o Título com proteção de largura
    With Frm.LbTitulo
        .AutoSize = False                   ' Trava para não bugar
        .WordWrap = False                   ' Impede de jogar letra para baixo
        .Caption = (NovoTitulo)
        .Width = Frm.FrmTitulo.Width - 20   ' Ocupa a barra quase toda
        .TextAlign = 2                      ' Centraliza via propriedade nativa (fmTextAlignCenter)
        .Top = (Frm.FrmTitulo.Height - .Height) / 2
        .Left = 10
    End With

    ' 4. Força a data a se reposicionar caso a janela tenha mudado
    Call RedimensionarBarraSuperior(Frm)
End Sub

' -------------------------------------------------------------------------
' Propósito: Encerra o sistema com segurança e gerencia o processo do Excel
' -------------------------------------------------------------------------
Public Sub EncerrarSistemaBeautyTech()
    On Error Resume Next
    
    ' 1. Salva o banco de dados/planilha silenciosamente
    ThisWorkbook.Save
    
    ' 2. Lógica Inteligente de Encerramento
    If Workbooks.Count > 1 Then
        ' Se o usuário tem outras planilhas abertas (ex: controle pessoal dele),
        ' devolvemos a visibilidade do Excel e fechamos SÓ o nosso sistema.
        Application.Visible = True
        ThisWorkbook.Close SaveChanges:=False ' Já salvamos acima
    Else
        ' Se só o BeautyTech estiver aberto, encerramos o Excel por completo.
        Application.Quit
    End If
End Sub

Private Sub IcoModoDev_Click(ByVal Frm As Object)
    ' Confirmação de segurança para evitar cliques acidentais
    If MsgBox("Atenção: Você está prestes a acessar o código fonte e a base de dados." & vbCrLf & _
              "Deseja ativar o Modo Desenvolvedor?", vbExclamation + vbYesNo, "Área Restrita") = vbYes Then
        
        ' 1. Devolve a visibilidade do programa Excel
        Application.Visible = True
        
        ' 2. Devolve a visibilidade específica da pasta de trabalho
        Windows(ThisWorkbook.Name).Visible = True
        
        ' 3. Descarrega o UserForm para liberar a edição do VBE
        Unload Frm
        
        ' 4. Opcional: Vai direto para a aba onde os dados brutos ficam
        ' Planilha1.Activate
    End If
End Sub

' -------------------------------------------------------------------------
' Propósito: Aplica as restrições visuais de acordo com o nível do usuário
' Parâmetros: Frm - O formulário principal (Usf_MenuPrincipal)
' -------------------------------------------------------------------------
Public Sub AplicarPermissoes(ByVal Frm As Object)
    
    ' Consulta a variável global protegida para saber o nível
    If Mdl_VariaveisGlobais.UsuarioNivel = "Admin" Then
        ' Se for Admin, libera o botão de desenvolvimento
        Frm.IcoModoDev.Visible = True
        
        ' [FUTURO] Aqui podemos colocar: Frm.LblMenuFinanceiro.Visible = True
    Else
        ' Se for Profissional/Comum, oculta o botão e protege o código
        Frm.IcoModoDev.Visible = False
        
        ' [FUTURO] Aqui podemos colocar: Frm.LblMenuFinanceiro.Visible = False
    End If
    
    ' Aproveita a variável global para preencher o nome e nível na tela!
    Frm.LbpUsuarioLogado.Caption = Mdl_VariaveisGlobais.UsuarioNome
    Frm.LbpUsuarioNivel.Caption = Mdl_VariaveisGlobais.UsuarioNivel

End Sub

' -------------------------------------------------------------------------
' Propósito: Ativa o ambiente de desenvolvimento e devolve o Excel
' Parâmetros: Frm - O formulário que chamou a ação (para ser fechado)
' -------------------------------------------------------------------------
Public Sub AtivarModoDesenvolvedor(ByVal Frm As Object)
    
    ' Confirmação de segurança
    If MsgBox("Atenção: Você está prestes a acessar o código-fonte e a base de dados." & vbCrLf & _
              "Deseja ativar o Modo Desenvolvedor?", vbExclamation + vbYesNo, "LogicUp Solutions - Área Restrita") = vbYes Then
        
        ' 1. Devolve a visibilidade da aplicação Excel inteira
        Application.Visible = True
        
        ' 2. Devolve a visibilidade específica da pasta de trabalho
        Windows(ThisWorkbook.Name).Visible = True
        
        ' 3. Descarrega o UserForm passado como parâmetro (em vez de Unload Me)
        Unload Frm
        
        ' 4. [Opcional] Seleciona uma planilha específica para não cair em uma tela em branco
        ' Planilha1.Activate
    End If
    
End Sub

