Attribute VB_Name = "Mdl_Login"
Option Explicit
' ==============================================================================
' Módulo: Mdl_Login
' Objetivo: Controlar a lógica e navegação do Formulário de Login (Controller)
' ==============================================================================

' ------------------------------------------------------------------------------
' 1. CONFIGURAÇÃO INICIAL
' ------------------------------------------------------------------------------
Public Sub ConfigurarFormulario(ByVal Frm As Object)
    ' Configurações visuais da MultiPage
    With Frm.MultiPagLogin
        .Value = 0              ' Começa na aba Login
        .Style = fmTabStyleNone ' Esconde as abas superiores
        .Left = -5
        .Top = -5
        .Height = 585           ' Ajuste conforme necessário para cobrir o form
        .Width = 410
    End With
        
    ' Centraliza na tela do Excel
    Frm.StartUpPosition = 1 ' CenterOwner
    
    ' 3. ESCONDER SENHAS (PasswordChar)
    ' Aplicamos o asterisco em todos os campos de senha do sistema
    Frm.TxPass.PasswordChar = "*"            ' Tela de Login
    Frm.TxRegPass.PasswordChar = "*"         ' Tela de Cadastro
    Frm.TxRegPassConfirm.PasswordChar = "*"  ' Confirmação de Cadastro
    Frm.TxResetPass.PasswordChar = "*"       ' Senha Atual (Reset)
    Frm.TxResetNewPass.PasswordChar = "*"    ' Nova Senha (Reset)
    
End Sub

' ------------------------------------------------------------------------------
' 2. NAVEGAÇÃO ENTRE ABAS (Login / Cadastro / Reset)
' ------------------------------------------------------------------------------
Public Sub IrParaLogin(ByVal Frm As Object)
    ' Limpa campos de Cadastro
    Frm.TxtRegName = ""
    Frm.TxtRegUser = ""
    Frm.TxRegPass = ""
    Frm.TxRegPassConfirm = ""
    
    ' Limpa campos de Reset
    Frm.TxtResetName = ""
    Frm.TxtResetUser = ""
    Frm.TxResetPass = ""
    Frm.TxResetNewPass = ""
    Frm.TxtResetStatus = ""
    
    ' Vai para aba 0 (Login)
    Frm.MultiPagLogin.Value = 0
End Sub

Public Sub IrParaCadastro(ByVal Frm As Object)
    ' Prepara a tela de cadastro
    Frm.TxtRegName = ""
    Frm.TxtRegUser = ""
    
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
    
    ' Validação Visual
    If Mdl_Utilitarios.CampoVazio(Frm.TxUser, "Digite seu usuário.") Then Exit Sub
    If Mdl_Utilitarios.CampoVazio(Frm.TxPass, "Digite sua senha.") Then Exit Sub
    
    Application.Cursor = xlWait
    
    Mdl_Conexao.ConectarBD
    ' Usa o Hash e valida
    Sucesso = Mdl_Autenticacao.ValidarUsuario(UCase(Frm.TxUser.Value), Frm.TxPass.Value)
    Mdl_Conexao.DesconectarBD
    
    Application.Cursor = xlDefault
    
    If Sucesso Then
        Mdl_Utilitarios.RegistrarLogAcesso Frm.TxUser.Value, "SUCESSO"
        Mdl_Utilitarios.MsgInfo "Bem-vindo(a), " & Mdl_VariaveisGlobais.UsuarioNome & "!", "Login"
        Unload Frm
        
        ' Aqui chamaria o Call Usf_MenuPrincipal.Show
'        MsgBox "Abrindo Menu Principal...", vbSystemModal
        Usf_MenuPrincipal.Show
        
    Else
        Mdl_Utilitarios.RegistrarLogAcesso Frm.TxUser.Value, "FALHA_SENHA"
        Mdl_Utilitarios.MsgAviso "Usuário ou senha incorretos."
        Frm.TxPass.Value = ""
        Frm.TxPass.SetFocus
    End If
End Sub

Public Sub ExecutarCadastro(ByVal Frm As Object)
    On Error GoTo ErroCadastro
    
    ' 1. VALIDAÇÃO DE CAMPOS VAZIOS
    If Mdl_Utilitarios.CampoVazio(Frm.TxtRegName, "Preencha o nome completo.") Then Exit Sub
    If Mdl_Utilitarios.CampoVazio(Frm.TxtRegUser, "Preencha o nome de usuário.") Then Exit Sub
    If Mdl_Utilitarios.CampoVazio(Frm.TxtRegEmail, "Preencha o e-mail.") Then Exit Sub ' Assumindo que o nome é TxtRegEmail
    If Mdl_Utilitarios.CampoVazio(Frm.TxRegPass, "Digite uma senha.") Then Exit Sub
    If Mdl_Utilitarios.CampoVazio(Frm.TxRegPassConfirm, "Confirme a sua senha.") Then Exit Sub
    
    ' 2. VALIDAÇÃO DE FORMATO (REGEX E REGRAS)
    
    ' Nome Completo
    If InStr(Trim(Frm.TxtRegName.Value), " ") = 0 Then
        Mdl_Utilitarios.MsgAviso "Por favor, digite seu nome e sobrenome.", "Cadastro"
        Frm.TxtRegName.SetFocus
        Exit Sub
    End If
       
    ' Senha Forte (NOVO!)
    If Not Mdl_Seguranca.ValidarSenhaForte(Frm.TxRegPass.Value) Then
        Mdl_Utilitarios.MsgAviso "A senha deve conter no mínimo 8 caracteres, incluindo letras maiúsculas, minúsculas, números e caracteres especiais (@$!%*?&).", "Segurança da Senha"
        Frm.TxRegPass.SetFocus
        Exit Sub
    End If
    
    ' Senhas Batem?
    If Frm.TxRegPass.Value <> Frm.TxRegPassConfirm.Value Then
        Mdl_Utilitarios.MsgAviso "As senhas digitadas não são iguais!", "Cadastro"
        Frm.TxRegPassConfirm.Value = ""
        Frm.TxRegPass.SetFocus
        Exit Sub
    End If
    
    ' E-mail Válido
    If Not Mdl_Seguranca.ValidarEmail(Frm.TxtRegEmail.Value) Then
        Mdl_Utilitarios.MsgAviso "O formato do e-mail é inválido!", "Cadastro"
        Frm.TxtRegEmail.SetFocus
        Exit Sub
    End If

    ' 3. VERIFICAR SE USUÁRIO OU E-MAIL JÁ EXISTEM
    Dim Rs As Object
    Dim SQL As String
    
    Mdl_Conexao.ConectarBD
    
    SQL = "SELECT Usuario FROM Tbl_Usuarios WHERE Usuario = '" & UCase(Frm.TxtRegUser.Value) & "' OR Email = '" & UCase(Frm.TxtRegEmail.Value) & "'"
    Set Rs = Mdl_Conexao.ObterRecordset(SQL)
    
    If Not Rs.EOF Then
        Mdl_Utilitarios.MsgAviso "Este Usuário ou E-mail já está cadastrado no sistema.", "Duplicidade"
        Mdl_Conexao.DesconectarBD
        Exit Sub
    End If
    Rs.Close
    
    ' 4. SALVAMENTO COM SEGURANÇA (SHA-256)
    If MsgBox("Deseja confirmar seu cadastro?", vbQuestion + vbYesNo, "Confirmar") = vbYes Then
        
        ' Note que usamos o Hash da senha para salvar
        SQL = "INSERT INTO Tbl_Usuarios (Nome, Usuario, Email, Senha, Nivel, Status, DataCadastro) VALUES (" & _
              "'" & UCase(Frm.TxtRegName.Value) & "', " & _
              "'" & UCase(Frm.TxtRegUser.Value) & "', " & _
              "'" & UCase(Frm.TxtRegEmail.Value) & "', " & _
              "'" & Mdl_Seguranca.GerarHashSHA256(Frm.TxRegPass.Value) & "', " & _
              "'PADRAO', 0, #" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "#)"
        
        Mdl_Conexao.ExecutarSQL SQL
        
        Mdl_Utilitarios.MsgInfo "Cadastro realizado com sucesso! Aguarde ativação pelo administrador.", "Sucesso"
        Mdl_Login.IrParaLogin Frm
    End If
    
    Mdl_Conexao.DesconectarBD
    Exit Sub

ErroCadastro:
    Mdl_Conexao.DesconectarBD
    Mdl_Utilitarios.GravarLogErro "Mdl_Login.ExecutarCadastro", Err.Number, Err.Description, Erl
    Mdl_Utilitarios.msgErro "Falha no cadastro: " & Err.Description
End Sub

Public Sub IrParaEsqueciSenha(ByVal Frm As Object)
    Dim Rs As Object
    Dim SQL As String
    Dim UsuarioInformado As String
    
    ' 1. Verifica se o usuário digitou o login na tela inicial
    UsuarioInformado = Trim(Frm.TxUser.Value)
    
    If UsuarioInformado = "" Then
        Mdl_Utilitarios.MsgAviso "Por favor, digite seu Usuário antes de clicar em 'Esqueci Senha'.", "Identificação"
        Frm.TxUser.SetFocus
        Exit Sub
    End If
    
    ' 2. Busca informações no banco
    Mdl_Conexao.ConectarBD
    SQL = "SELECT Nome, Status FROM Tbl_Usuarios WHERE Usuario = '" & UCase(UsuarioInformado) & "'"
    Set Rs = Mdl_Conexao.ObterRecordset(SQL)
    
    ' 3. Validações de existência e Status
    If Rs.EOF Then
        Mdl_Utilitarios.MsgAviso "Usuário não encontrado em nossa base de dados.", "Erro de Identificação"
        Frm.TxUser.SetFocus
        Mdl_Conexao.DesconectarBD
        Exit Sub
    End If
    
    If Rs("Status") <> 1 Then
        Mdl_Utilitarios.MsgAviso "Este usuário está inativo. Entre em contato com o administrador.", "Acesso Negado"
        Mdl_Conexao.DesconectarBD
        Exit Sub
    End If
    
    ' 4. Carrega os dados na tela de Reset (Aba 2)
    Frm.TxtResetName.Value = Rs("Nome")
    Frm.TxtResetUser.Value = UCase(UsuarioInformado)
    Frm.TxtResetStatus.Value = "ATIVO"
    
    ' Limpa campos de senha por segurança
    Frm.TxResetPass.Value = ""
    Frm.TxResetNewPass.Value = ""
    
    ' 5. Navega para a aba de Reset
    Frm.MultiPagLogin.Value = 2
    
    Mdl_Conexao.DesconectarBD
    
    ' Foca no campo de Senha Atual
    On Error Resume Next
    Frm.TxResetPass.SetFocus
End Sub

Public Sub ExecutarResetSenha(ByVal Frm As Object)
    On Error GoTo ErroReset
    
    ' 1. Validação de preenchimento das senhas
    If Mdl_Utilitarios.CampoVazio(Frm.TxResetPass, "Informe sua senha atual para validar.") Then Exit Sub
    If Mdl_Utilitarios.CampoVazio(Frm.TxResetNewPass, "Digite a nova senha desejada.") Then Exit Sub
    
    ' 2. Validar se a nova senha é forte
    If Not Mdl_Seguranca.ValidarSenhaForte(Frm.TxResetNewPass.Value) Then
        Mdl_Utilitarios.MsgAviso "A nova senha não atende aos requisitos de segurança.", "Senha Fraca"
        Frm.TxResetNewPass.SetFocus
        Exit Sub
    End If
    
    ' 3. Verificar se a senha ATUAL informada confere com o banco
    Dim SenhaAtualHash As String
    Dim SQL As String
    Dim Rs As Object
    
    SenhaAtualHash = Mdl_Seguranca.GerarHashSHA256(Frm.TxResetPass.Value)
    
    Mdl_Conexao.ConectarBD
    SQL = "SELECT ID FROM Tbl_Usuarios WHERE Usuario = '" & UCase(Frm.TxtResetUser.Value) & "' " & _
          "AND Senha = '" & SenhaAtualHash & "'"
    Set Rs = Mdl_Conexao.ObterRecordset(SQL)
    
    If Rs.EOF Then
        Mdl_Utilitarios.MsgAviso "A 'Senha Atual' informada está incorreta.", "Falha de Segurança"
        Frm.TxResetPass.Value = ""
        Frm.TxResetPass.SetFocus
        Mdl_Conexao.DesconectarBD
        Exit Sub
    End If
    
    ' 4. Atualizar para a Nova Senha
    If MsgBox("Confirma a alteração de senha para o usuário " & Frm.TxtResetUser.Value & "?", _
              vbQuestion + vbYesNo, "Confirmar Alteração") = vbYes Then
              
        Dim NovaSenhaHash As String
        NovaSenhaHash = Mdl_Seguranca.GerarHashSHA256(Frm.TxResetNewPass.Value)
        
        SQL = "UPDATE Tbl_Usuarios SET Senha = '" & NovaSenhaHash & "' " & _
              "WHERE Usuario = '" & UCase(Frm.TxtResetUser.Value) & "'"
              
        Mdl_Conexao.ExecutarSQL SQL
        
        Mdl_Utilitarios.MsgInfo "Senha alterada com sucesso! Utilize suas novas credenciais.", "Sucesso"
        Mdl_Login.IrParaLogin Frm
    End If
    
    Mdl_Conexao.DesconectarBD
    Exit Sub

ErroReset:
    Mdl_Conexao.DesconectarBD
    Mdl_Utilitarios.GravarLogErro "Mdl_Login.ExecutarResetSenha", Err.Number, Err.Description, Erl
    Mdl_Utilitarios.msgErro "Erro ao processar reset: " & Err.Description
End Sub

' ==============================================================================
' Objetivo: Alternar exibição da senha trocando a imagem (Picture) do controle
' ==============================================================================
Public Sub AlternarVisualizacaoSenha(ByRef Txt As Object, ByRef LblIconeClique As Object, ByRef LblFonteVer As Object, ByRef LblFonteEsconder As Object)
    ' Se a senha estiver escondida (*), mostramos o texto e o olho aberto
    If Txt.PasswordChar = "*" Then
        Txt.PasswordChar = ""
        Set LblIconeClique.Picture = LblFonteEsconder.Picture
    Else
        ' Se estiver visível, escondemos o texto e voltamos para o olho fechado
        Txt.PasswordChar = "*"
        Set LblIconeClique.Picture = LblFonteVer.Picture
    End If
    
    ' UX: Mantém o foco e posiciona o cursor no final do texto
    Txt.SetFocus
    Txt.SelStart = Len(Txt.Text)
    
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

