Attribute VB_Name = "Mdl_Login"
' ==============================================================================
' Módulo: Mdl_Login
' Objetivo: Controlar a lógica e navegação do Formulário de Login (Controller)
' ==============================================================================
Option Explicit

' ------------------------------------------------------------------------------
' 1. CONFIGURAÇÃO INICIAL E VISUAL DO FORMULÁRIO DE LOGIN
' ------------------------------------------------------------------------------
Public Sub ConfigurarFormulario(ByVal Frm As Object)
    
    ' 1. Configurações visuais da MultiPage (Esconder abas e ajustar tamanho)
    With Frm.MultiPagLogin
        .Value = 0              ' Começa na aba Login
        .Style = 2              ' 2 = fmTabStyleNone (Esconde as abas superiores)
        .Left = -5
        .Top = -5
        .Height = 550           ' Ajuste conforme necessário para cobrir o form
        .Width = 410
    End With
        
    ' 2. Centraliza na tela do Excel
    Frm.StartUpPosition = 1 ' 1 = CenterOwner
    
    ' 3. ESCONDER SENHAS (PasswordChar)
    ' Definimos o caractere de máscara. O "•" (Chr(149)) dá um visual web super moderno!
    Dim Mascara As String
    Mascara = Chr(149) ' Se preferir o clássico, basta trocar para "*"
    
    ' Aplicamos a máscara em todos os campos de senha do sistema
    Frm.TxPass.PasswordChar = Mascara             ' Tela de Login
    Frm.TxRegPass.PasswordChar = Mascara          ' Tela de Cadastro
    Frm.TxRegPassConfirm.PasswordChar = Mascara   ' Confirmação de Cadastro
    
End Sub

' ------------------------------------------------------------------------------
' 2. NAVEGAÇÃO ENTRE ABAS (Login / Cadastro / Reset)
' ------------------------------------------------------------------------------
Public Sub IrParaLogin(ByVal Frm As Object)
    ' Limpa campos de Cadastro
    Frm.TxtRegName.Value = ""
    Frm.TxtRegUser.Value = ""
    Frm.TxtRegEmail.Value = ""
    Frm.TxRegPass.Value = ""
    Frm.TxRegPassConfirm.Value = ""
       
    ' Vai para aba 0 (Login)
    Frm.MultiPagLogin.Value = 0
End Sub

Public Sub IrParaCadastro(ByVal Frm As Object)
    ' Prepara a tela de cadastro
    Frm.TxtUser.Value = ""
    Frm.TxPass.Value = ""
    
    ' Vai para aba 1 (Cadastro)
    Frm.MultiPagLogin.Value = 1
    
    ' Foca no primeiro campo
    On Error Resume Next
    Frm.TxtRegName.SetFocus
End Sub

' ------------------------------------------------------------------------------
' 3. AÇÕES DOS BOTÕES (Clicar em Entrar, Cadastrar, etc)
' ------------------------------------------------------------------------------
Public Sub ExecutarLogin(ByVal Frm As Object)
    Dim Sucesso As Boolean
    
    On Error GoTo ErroLogin
    
    ' =========================================================================
    ' 1. VALIDAÇÃO VISUAL
    ' =========================================================================
    If Mdl_Utilitarios.CampoVazio(Frm.TxtUser, "Digite seu usuário.") Then Exit Sub
    If Mdl_Utilitarios.CampoVazio(Frm.TxPass, "Digite sua senha.") Then Exit Sub
    
    Application.Cursor = xlWait
    
    ' =========================================================================
    ' 2. MOTOR DE AUTENTICAÇÃO
    ' =========================================================================
    Mdl_Conexao.ConectarBD
    ' O Mdl_Autenticacao levanta a bandeira Mdl_VariaveisGlobais.RequerTrocaSenha se a senha for a padrão
    Sucesso = Mdl_Autenticacao.ValidarUsuario(UCase(Frm.TxtUser.Value), Frm.TxPass.Value)
    Mdl_Conexao.DesconectarBD
    
    Application.Cursor = xlDefault
    
    ' =========================================================================
    ' 3. INTERCEPTADOR DE ROTA E FLUXO DE ACESSO
    ' =========================================================================
    If Sucesso Then
        Mdl_Utilitarios.RegistrarLogAcesso Frm.TxtUser.Value, "SUCESSO"
        
        ' Verifica a bandeira global: É o primeiro acesso com a senha padrão?
        If Mdl_VariaveisGlobais.RequerTrocaSenha = True Then
            
            ' --- ROTA A: PRIMEIRO ACESSO (Obrigatório Trocar Senha) ---
            Mdl_Utilitarios.MsgInfo "Bem-vindo(a), " & Mdl_VariaveisGlobais.UsuarioNome & "!" & vbCrLf & vbCrLf & _
                                    "Por questões de segurança, é necessário cadastrar uma senha pessoal definitiva para continuar.", "Primeiro Acesso"
            
            ' Esconde e descarrega o login, abrindo o modal obrigatório de troca
            Frm.Hide
            Usf_TrocarSenhaProvisoria.Show
            Unload Frm
            
        Else
            
            ' --- ROTA B: ACESSO NORMAL ---
            Mdl_Utilitarios.MsgInfo "Bem-vindo(a), " & Mdl_VariaveisGlobais.UsuarioNome & "!", "Login"
            Frm.Hide
            Usf_MenuPrincipal.Show
            Unload Frm
            
        End If
        
    Else
        ' --- FALHA NO LOGIN ---
        Mdl_Utilitarios.RegistrarLogAcesso Frm.TxUser.Value, "FALHA_SENHA"
        Mdl_Utilitarios.MsgAviso "Usuário ou senha incorretos."
        
        ' Limpa a senha e devolve o foco para tentar novamente
        Frm.TxPass.Value = ""
        Frm.TxPass.SetFocus
    End If
    
    Exit Sub

ErroLogin:
    Application.Cursor = xlDefault
    Mdl_Utilitarios.GravarLogErro "Mdl_Login.ExecutarLogin", Err.Number, Err.Description
    Mdl_Utilitarios.msgErro "Falha crítica ao tentar processar o login. O erro foi registrado."
    Mdl_Conexao.DesconectarBD
End Sub

' ------------------------------------------------------------------------------
' Módulo: Mdl_Login
' Rotina: ExecutarCadastro (Versão Validada e Auditada)
' ------------------------------------------------------------------------------
Public Sub ExecutarCadastro(ByVal Frm As Object)
    Dim SQL As String
    Dim Rs As Object
    Dim NomeLimpo As String, UserLimpo As String, EmailLimpo As String
    
    On Error GoTo ErroCadastro
    
    ' =========================================================================
    ' 0. SANITIZAÇÃO E PREPARAÇÃO
    ' =========================================================================
    ' Limpa espaços e prepara strings para evitar quebra por aspas simples (')
    Mdl_Utilitarios.TrimTodosCampos Frm.TxtRegUser, Frm.TxtRegEmail
    
    NomeLimpo = Replace(Application.WorksheetFunction.Trim(Frm.TxtRegName.Value), "'", "''")
    UserLimpo = Replace(UCase(Frm.TxtRegUser.Value), "'", "''")
    EmailLimpo = Replace(UCase(Frm.TxtRegEmail.Value), "'", "''")
    
    ' =========================================================================
    ' 1. VALIDAÇÕES VISUAIS (Feedback em Rosa)
    ' =========================================================================
    If Mdl_Utilitarios.CampoVazio(Frm.TxtRegName, "Preencha o nome completo.") Then Exit Sub
    If Mdl_Utilitarios.CampoVazio(Frm.TxtRegUser, "Preencha o nome de usuário.") Then Exit Sub
    If Mdl_Utilitarios.CampoVazio(Frm.TxtRegEmail, "Preencha o e-mail.") Then Exit Sub
    If Mdl_Utilitarios.CampoVazio(Frm.TxRegPass, "Digite uma senha.") Then Exit Sub
    If Mdl_Utilitarios.CampoVazio(Frm.TxRegPassConfirm, "Confirme a sua senha.") Then Exit Sub
    
    ' =========================================================================
    ' 2. REGRAS DE NEGÓCIO E FORMATO
    ' =========================================================================
    ' Valida Nome Completo
    If InStr(NomeLimpo, " ") = 0 Then
        Mdl_Utilitarios.MsgAviso "Por favor, digite seu nome e sobrenome.", "Cadastro"
        Frm.TxtRegName.SetFocus: Exit Sub
    End If
       
    ' Valida Senha Forte
    If Not Mdl_Seguranca.ValidarSenhaForte(Frm.TxRegPass.Value) Then
        Mdl_Utilitarios.MsgAviso "A senha deve conter no mínimo 8 caracteres (A-z, 0-9 e símbolos).", "Segurança"
        Frm.TxRegPass.SetFocus: Exit Sub
    End If
    
    ' Confirmação de Senha
    If Frm.TxRegPass.Value <> Frm.TxRegPassConfirm.Value Then
        Mdl_Utilitarios.MsgAviso "As senhas digitadas não são iguais!", "Cadastro"
        Frm.TxRegPassConfirm.Value = "": Frm.TxRegPass.SetFocus: Exit Sub
    End If
    
    ' Valida E-mail
    If Not Mdl_Seguranca.ValidarEmail(EmailLimpo) Then
        Mdl_Utilitarios.MsgAviso "O formato do e-mail é inválido!", "Cadastro"
        Frm.TxtRegEmail.SetFocus: Exit Sub
    End If

    ' =========================================================================
    ' 3. DUPLICIDADE E PERSISTÊNCIA
    ' =========================================================================
    Mdl_Conexao.ConectarBD
    
    ' Verifica se o usuário ou e-mail já existem
    SQL = "SELECT Usuario FROM Tbl_Usuarios WHERE Usuario = '" & UserLimpo & "' OR Email = '" & EmailLimpo & "'"
    Set Rs = Mdl_Conexao.ObterRecordset(SQL)
    
    If Not Rs.EOF Then
        Mdl_Utilitarios.MsgAviso "Usuário ou E-mail já cadastrados.", "Duplicidade"
        Rs.Close: Mdl_Conexao.DesconectarBD: Exit Sub
    End If
    Rs.Close
    
    ' =========================================================================
    ' 4. GRAVAÇÃO E AUDITORIA
    ' =========================================================================
    If MsgBox("Deseja confirmar seu pedido de cadastro?", vbQuestion + vbYesNo, "Confirmar") = vbYes Then
        
        ' Inserção com Status 0 (Inativo) para aprovação posterior
        SQL = "INSERT INTO Tbl_Usuarios (Nome, Usuario, Email, Senha, Nivel, Status, DataCadastro) VALUES (" & _
              "'" & UCase(NomeLimpo) & "', " & _
              "'" & UserLimpo & "', " & _
              "'" & EmailLimpo & "', " & _
              "'" & Mdl_Seguranca.GerarHashSHA256(Frm.TxRegPass.Value) & "', " & _
              "'PADRAO', 0, #" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "#)"
        
        Mdl_Conexao.ExecutarSQL SQL
        
        ' --- CAMADA DE AUDITORIA ---
        Mdl_Utilitarios.RegistrarAuditoria "SOLICITACAO_ACESSO", "Tbl_Usuarios", 0, "Novo auto-cadastro pendente: " & UserLimpo
        
        Mdl_Utilitarios.MsgInfo "Cadastro realizado! Aguarde a ativação pela gerência.", "Sucesso"
        Mdl_Login.IrParaLogin Frm
    End If
    
    Mdl_Conexao.DesconectarBD
    Exit Sub

ErroCadastro:
    ' --- CAMADA DE LOG DE ERRO ---
    Mdl_Utilitarios.GravarLogErro "Mdl_Login.ExecutarCadastro", Err.Number, Err.Description
    Mdl_Utilitarios.msgErro "Falha crítica ao processar cadastro. O erro foi registrado no sistema."
    Mdl_Conexao.DesconectarBD
End Sub

' ==============================================================================
' Objetivo: Alternar exibição da senha trocando a imagem (Picture) do controle
' ==============================================================================
Public Sub AlternarVisualizacaoSenha(ByRef Txt As Object, ByRef LblIconeClique As Object, ByRef LblFonteVer As Object, ByRef LblFonteEsconder As Object)
    
    ' Define qual é a máscara oficial do sistema (Bullet Moderno)
    Dim Mascara As String
    Mascara = Chr(149) ' Bolinha "•"
    
    ' LÓGICA BLINDADA: Se PasswordChar for diferente de Vazio, significa que está escondida
    If Txt.PasswordChar <> "" Then
        ' Limpa a máscara para mostrar o texto e troca o ícone
        Txt.PasswordChar = ""
        Set LblIconeClique.Picture = LblFonteEsconder.Picture
    Else
        ' Se estiver visível (vazia), aplicamos a máscara oficial e voltamos o ícone
        Txt.PasswordChar = Mascara
        Set LblIconeClique.Picture = LblFonteVer.Picture
    End If
    
    ' UX: Mantém o foco e posiciona o cursor no final do texto
    Txt.SetFocus
    Txt.SelStart = Len(Txt.Text)
    
End Sub

Public Sub PaginaCadastro(ByVal Frm As Object)
    On Error Resume Next
    
    With Frm.LbTituloPaginaCadastro
        .BackStyle = 0 ' Fundo Transparente
        
        ' A MÁGICA DO DESIGN: Usando a cor do Menu como cor da Letra
        .ForeColor = RGB(33, 47, 61)
        
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 24
                
        .Top = 80
        .Left = (Frm.MultiPagLogin.Width / 2) - (.Width / 2)
    End With
End Sub

Public Sub TerminarSessao(ByVal Frm As Object)
    On Error Resume Next
    
    ' 1. Registro de Log (Se houver sessão ativa)
    If Mdl_VariaveisGlobais.UsuarioLogado Then
        Mdl_Conexao.ConectarBD
        Mdl_Utilitarios.RegistrarLogAcesso Mdl_VariaveisGlobais.UsuarioNome, "LOGOUT"
        Mdl_Conexao.DesconectarBD
    End If
    
    ' 2. Limpeza de Memória
    Mdl_VariaveisGlobais.LimparSessao
    Unload Frm
    
    ' 3. DECISÃO DE ENCERRAMENTO (Crítico para não deixar o Excel oculto)
    ' Se o login for fechado sem que o Menu Principal tenha sido aberto,
    ' precisamos encerrar a aplicação.
    If Workbooks.Count > 1 Then
        Application.Visible = True
        ThisWorkbook.Close SaveChanges:=True
    Else
        ThisWorkbook.Save
        Application.Quit
    End If
End Sub





