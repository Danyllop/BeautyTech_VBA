VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Usf_TrocarSenhaProvisoria 
   Caption         =   "BeautyTech - Segurança"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   OleObjectBlob   =   "Usf_TrocarSenhaProvisoria.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Usf_TrocarSenhaProvisoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================
' NOME DO ARQUIVO: Usf_TrocarSenhaProvisoria
' PROJETO:         Sistema BeautyTech - Gestão Integrada
' DESCRIÇÃO:       Interface de Segurança para Definição de Senha Definitiva
' AUTOR:           LogicUp Solutions
' DATA:            Fevereiro/2026
' ==============================================================================
Option Explicit

Private EfeitoCurso   As Collection
Private SimpleButton  As Collection

Private Sub UserForm_Initialize()
    ' 1. Personaliza a barra de título (Padrão do BeautyTech)
    Mdl_UI_Efeitos.PersonalizarBarraTitulo Me, RGB(33, 47, 61), RGB(255, 255, 255)
    
    ' 2. Carrega os efeitos de hover (Mãozinha e Negrito nos botões)
    Set EfeitoCurso = Mdl_UI_Efeitos.CriarEfeitosMaozinha(Me)
    Set SimpleButton = Mdl_UI_Efeitos.CriarSimpleButton(Me)
    
    ' 3. UX: Preenche os dados do usuário logado automaticamente
    Me.TxtNome.Text = Mdl_VariaveisGlobais.UsuarioNome
    Me.TxtUser.Text = Mdl_VariaveisGlobais.UsuarioLogin
    
    ' Bloqueia a edição dessas caixas e tira o "Tab" delas
    Me.TxtNome.Locked = True
    Me.TxtUser.Locked = True
    Me.TxtNome.TabStop = False
    Me.TxtUser.TabStop = False
    
    ' 4. Foco direto na caixa de digitar a nova senha
    Me.TxtNewPass.SetFocus
End Sub

' -------------------------------------------------------------------------
' EVENTO: Botão Salvar e Entrar (Troca de Senha Obrigatória)
' -------------------------------------------------------------------------
Private Sub BtnSalvar_Click()
    Dim Senha1 As String
    Dim Senha2 As String
    Dim NovaSenhaHash As String
    Dim SQL As String
    
    On Error GoTo ErroTrocaSenha
    
    ' =========================================================================
    ' CAMADA 1: CAPTURA E VALIDAÇÃO VISUAL (Corrigido para TxtNewPass e TxtConfirmPass)
    ' =========================================================================
    If Mdl_Utilitarios.CampoVazio(Me.TxtNewPass, "Digite a sua nova senha.") Then Exit Sub
    If Mdl_Utilitarios.CampoVazio(Me.TxtConfirmPass, "Confirme a sua nova senha.") Then Exit Sub
    
    Senha1 = Trim(Me.TxtNewPass.Text)
    Senha2 = Trim(Me.TxtConfirmPass.Text)
    
    ' =========================================================================
    ' CAMADA 2: REGRAS DE NEGÓCIO DA SENHA
    ' =========================================================================
    ' 1. Verifica se as senhas coincidem
    If Senha1 <> Senha2 Then
        Mdl_Utilitarios.MsgAviso "As senhas digitadas não conferem. Tente novamente.", "Divergência"
        Me.TxtNewPass.Text = ""
        Me.TxtConfirmPass.Text = ""
        Me.TxtNewPass.SetFocus
        Exit Sub
    End If
    
    ' 2. Verifica se a senha tem o tamanho mínimo exigido (Ex: 6 caracteres)
    If Len(Senha1) < 6 Then
        Mdl_Utilitarios.MsgAviso "A sua nova senha deve ter pelo menos 6 caracteres por motivos de segurança.", "Senha Fraca"
        Me.TxtNewPass.SetFocus
        Exit Sub
    End If
    
    ' 3. Impede que o utilizador tente usar a mesma senha padrão como "nova senha"
    If Senha1 = "Senh@1234" Then
        Mdl_Utilitarios.MsgAviso "Você não pode usar a senha provisória como sua senha definitiva. Escolha uma nova senha.", "Senha Inválida"
        Me.TxtNewPass.SetFocus
        Exit Sub
    End If

    ' =========================================================================
    ' CAMADA 3: CRIPTOGRAFIA E ATUALIZAÇÃO NO BANCO DE DADOS
    ' =========================================================================
    Application.Cursor = xlWait
    
    NovaSenhaHash = Mdl_Seguranca.GerarHashSHA256(Senha1)
    
    ' O Mdl_Autenticacao já guardou o ID do usuário nesta variável quando ele logou
    SQL = "UPDATE Tbl_Usuarios SET Senha = '" & NovaSenhaHash & "' WHERE ID = " & Mdl_VariaveisGlobais.UsuarioID
    
    Mdl_Conexao.ConectarBD
    Mdl_Conexao.ExecutarSQL SQL
    
    ' =========================================================================
    ' CAMADA 4: AUDITORIA E DESBLOQUEIO DO SISTEMA
    ' =========================================================================
    Mdl_Utilitarios.RegistrarAuditoria "TROCA_SENHA_OBRIGATORIA", "Tbl_Usuarios", Mdl_VariaveisGlobais.UsuarioID, "O usuário definiu a sua senha definitiva no primeiro acesso."
    Mdl_Conexao.DesconectarBD
    
    Application.Cursor = xlDefault
    
    ' Baixa a bandeira de restrição (Destranca o acesso geral)
    Mdl_VariaveisGlobais.RequerTrocaSenha = False
    
    Mdl_Utilitarios.MsgInfo "Senha cadastrada com sucesso! O seu sistema está liberado.", "Acesso Permitido"
    
    ' Remove o modal de segurança e abre finalmente o BeautyTech!
    Unload Me
    Usf_MenuPrincipal.Show
    
    Exit Sub

ErroTrocaSenha:
    Application.Cursor = xlDefault
    Mdl_Utilitarios.GravarLogErro "Usf_TrocarSenhaProvisoria.BtnSalvar_Click", Err.Number, Err.Description
    Mdl_Utilitarios.msgErro "Ocorreu um erro ao salvar a nova senha. Tente novamente."
    Mdl_Conexao.DesconectarBD
End Sub

' -------------------------------------------------------------------------
' EVENTOS DE UX E NAVEGAÇÃO
' -------------------------------------------------------------------------
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Mdl_UI_Efeitos.LimparFoco
End Sub

' Bloqueia o fechamento pelo "X" do Windows
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        Mdl_Utilitarios.MsgAviso "Por favor, defina a nova senha ou cancele o login para sair.", "Segurança"
    End If
End Sub

' O Botão de Fuga Seguro (Cancelar)
Private Sub BtnCancelar_Click()
    Mdl_VariaveisGlobais.UsuarioLogado = False
    Unload Me
    Usf_Login.Show
End Sub

