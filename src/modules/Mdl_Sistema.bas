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
'             IndexPagina - O número da aba na MultiPage (0 = Home, 1 = Agenda)
'             NovoTitulo - O texto que vai aparecer na barra superior
' -------------------------------------------------------------------------
Public Sub NavegarPara(ByVal Frm As Object, ByVal IndexPagina As Integer, ByVal NovoTitulo As String)
    On Error Resume Next

    ' 1. Muda a página
    Frm.MultiPagMain.Value = IndexPagina
    Frm.LbTitulo.Caption = NovoTitulo
    Frm.LbData.Caption = StrConv(Format(Date, "dddd, dd/mm/yyyy"), vbProperCase)

    ' 2. GATILHO DE CARGA (Apenas quando a página é ativada)
    Select Case IndexPagina
        Case 5 ' Gestão de Usuários
            Mdl_Gestao_Usuarios.CarregarDadosUsuarios Frm
    End Select

    ' 3. Chama o Designer (Apenas para ajustar a "moldura" visual)
    Mdl_UI_Designer.RedimensionarBarraSuperior Frm
    Mdl_UI_Designer.GerenciarRenderizacaoPaginas Frm
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
' Propósito: Garante a segurança do sistema aplicando permissões de acesso.
' -------------------------------------------------------------------------
'Public Sub AplicarPermissoes(ByVal Frm As Object)
'    Dim TemAcessoGestao As Boolean
'
'    ' 1. VALIDAÇÃO LÓGICA (Case-Insensitive)
'    ' Centralizamos a regra: ADMIN e SUPERVISOR têm acesso à gestão.
'    Select Case UCase(Mdl_VariaveisGlobais.UsuarioNivel)
'        Case "ADMIN", "SUPERVISOR"
'            TemAcessoGestao = True
'        Case Else
'            TemAcessoGestao = False
'    End Select
'
'    ' 2. SEGURANÇA DA INTERFACE (Controles confirmados: IcoUsuarios, IcoModoDev)
'    On Error Resume Next ' Proteção contra erro 438 se um controle for renomeado
'
'    ' Somente ADMIN visualiza o Modo Desenvolvedor
'    Frm.IcoModoDev.Visible = (UCase(Mdl_VariaveisGlobais.UsuarioNivel) = "ADMIN")
'
'    ' ADMIN/SUPERVISOR visualizam o botão de Gestão
'    Frm.IcoUsuarios.Visible = TemAcessoGestao
'
'    ' 3. PROTEÇÃO ESTRUTURAL (MultiPage)
'    ' Desabilita a aba de gestão para impedir acesso via teclado ou código
'    With Frm.MultiPagMain
'        .Pages(5).Enabled = TemAcessoGestao
'
'        ' Se o usuário estiver na página de gestão sem permissão, força o retorno à Dashboard
'        If .Value = 5 And Not TemAcessoGestao Then .Value = 0
'    End With
'
'    ' 4. DADOS DA SESSÃO (Rodapé)
'    Frm.LbpUsuarioLogado.Caption = UCase(Mdl_VariaveisGlobais.UsuarioLogin)
'    Frm.LbpUsuarioNivel.Caption = UCase(Mdl_VariaveisGlobais.UsuarioNivel)
'
'    On Error GoTo 0
'End Sub

Public Sub AplicarPermissoes(ByVal Frm As Object)
    Dim TemAcessoGestao As Boolean
    Dim NivelAtual As String
    
    ' Captura os dados do erro imediatamente se algo falhar na lógica
    On Error GoTo ErroPermissoes

    ' 1. CAPTURA E NORMALIZAÇÃO
    ' Usamos UCase para evitar conflitos entre "Admin", "admin" ou "ADMIN"
    NivelAtual = UCase(Mdl_VariaveisGlobais.UsuarioNivel)

    ' 2. VALIDAÇÃO LÓGICA DE NÍVEIS
    ' Definimos quem possui acesso às abas de gestão e usuários
    Select Case NivelAtual
        Case "ADMIN", "GERENTE"
            TemAcessoGestao = True
        Case Else
            ' Nível PADRAO ou qualquer outro não terá acesso administrativo
            TemAcessoGestao = False
    End Select

    ' 3. SEGURANÇA DA INTERFACE (Controles Visuais)
    ' On Error Resume Next protege contra erros de controles ausentes no formulário
    On Error Resume Next

    ' Exclusividade técnica: Somente o ADMIN visualiza ferramentas de desenvolvedor
    Frm.IcoModoDev.Visible = (NivelAtual = "ADMIN")

    ' Acesso Gerencial: ADMIN e GERENTE visualizam o botão de gestão de usuários
    Frm.IcoUsuarios.Visible = TemAcessoGestao

    ' 4. PROTEÇÃO ESTRUTURAL (MultiPage)
    ' Desabilita fisicamente a aba de gestão para impedir navegação forçada
    With Frm.MultiPagMain
        ' Assume-se que a página de Gestão é o índice 5
        .Pages(5).Enabled = TemAcessoGestao

        ' INTERCEPTADOR DE ROTA: Se o usuário estiver na aba proibida, volta para a Dashboard
        If .Value = 5 And Not TemAcessoGestao Then .Value = 0
    End With

    ' 5. DADOS DA SESSÃO (Rodapé do Sistema)
    Frm.LbpUsuarioLogado.Caption = UCase(Mdl_VariaveisGlobais.UsuarioLogin)
    Frm.LbpUsuarioNivel.Caption = NivelAtual

    ' Retorna o tratamento de erro para o fluxo normal
    On Error GoTo 0
    Exit Sub

ErroPermissoes:
    ' Registro técnico da falha no banco de dados para auditoria
    Mdl_Utilitarios.GravarLogErro "Mdl_Seguranca.AplicarPermissoes", Err.Number, Err.Description
    Mdl_Utilitarios.msgErro "Falha ao aplicar as restrições de acesso do usuário."
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

