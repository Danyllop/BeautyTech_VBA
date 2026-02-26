VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Usf_EditarUsuario 
   Caption         =   "BeautyTech - Editar Usuário"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "Usf_EditarUsuario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Usf_EditarUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private EfeitoCurso   As Collection ' Cursor Mãozinha
Private SimpleButton  As Collection ' Efeito Visual (Negrito/Tamanho)
Private ColMascaras   As Collection ' Máscaras de Texto (Data, CPF, etc)
Private ColMaiusculas As Collection ' Força caixa alta em TextBoxes

' Variáveis Globais do Formulário para o "Dirty Check" (Verificação de Alteração)
Private NomeOriginal    As String
Private EmailOriginal   As String
Private PerfilOriginal  As String
Private StatusOriginal  As Integer

' --- Inicialização ---
Private Sub UserForm_Initialize()
    ' 1. Configurações Físicas
    Me.Height = 350
    Me.Width = 400
                
    ' 2. Ativação dos Efeitos (Chamadas Individuais Organizadas)
    Set EfeitoCurso = Mdl_UI_Efeitos.CriarEfeitosMaozinha(Me)
    Set SimpleButton = Mdl_UI_Efeitos.CriarSimpleButton(Me)
    Set ColMaiusculas = Mdl_UI_Efeitos.AtivarMaiusculas(Me)
    Set ColMascaras = Mdl_UI_Efeitos.AtivarMascaras(Me)

    Mdl_UI_Efeitos.PersonalizarBarraTitulo Me, RGB(33, 95, 152), RGB(255, 255, 255)
    
    ' 3. Foco
    Me.TxtNome.SetFocus
    
    ' 4. Proteção para os campos.
    Me.TxtId.Locked = True
    Me.TxtUsuario.Locked = True
    Me.TxtDataCadastro.Locked = True
    
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Mdl_UI_Efeitos.LimparFoco
End Sub

' -------------------------------------------------------------------------
' EVENTO: Interceptação do fechamento da janela (Bloqueia o "X")
' -------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' CloseMode = 0 (ou vbFormControlMenu) significa que o usuário clicou no "X"
    If CloseMode = 0 Then
        ' Cancela a ação de fechar
        Cancel = True
        
        ' Exibe um aviso padronizado orientando o uso correto da interface
        Mdl_Utilitarios.MsgAviso "Por favor, utilize os botões 'Salvar' ou 'Cancelar' para fechar o formulário.", "Ação Bloqueada"
    End If
End Sub

' -------------------------------------------------------------------------
' EVENTO: Botão Cancelar
' -------------------------------------------------------------------------
Private Sub BtnCancelar_Click()
    ' Fecha o formulário sem fazer nenhuma alteração no banco
    Unload Me
End Sub

' -------------------------------------------------------------------------
' MÉTODOS PÚBLICOS (Encapsulamento com Snapshot e Auditoria)
' -------------------------------------------------------------------------
Public Sub CarregarDados(ByVal IDUsuario As String)
    Dim Rs As Object
    Dim SQL As String
    
    On Error GoTo ErroCarga
    
    Mdl_Conexao.ConectarBD
    
    ' Busca todos os dados do usuário específico
    SQL = "SELECT * FROM Tbl_Usuarios WHERE ID = " & IDUsuario
    Set Rs = Mdl_Conexao.ObterRecordset(SQL)
    
    If Not Rs.EOF Then
        ' 1. Preenche as Caixas de Texto
        Me.TxtId.Text = Rs("ID")
        Me.TxtNome.Text = Rs("Nome")
        Me.TxtUsuario.Text = Rs("Usuario")
        Me.TxtEmail.Text = Rs("Email")
        
        ' Garante o formato de data brasileiro
        Me.TxtDataCadastro.Text = Format(Rs("DataCadastro"), "dd/mm/yyyy")
        
        ' 2. Lógica do Perfil (OptionButtons)
        Select Case UCase(Rs("Nivel"))
            Case "ADMIN": Me.OptAdmin.Value = True
            Case "GERENTE": Me.OptGerente.Value = True
            Case "PADRAO": Me.OptPadrao.Value = True
        End Select
        
        ' 3. Lógica do Status (1 = Ativo, 0 = Inativo)
        If Rs("Status") = 1 Then
            Me.OptAtivo.Value = True
        Else
            Me.OptInativo.Value = True
        End If
        
        ' =====================================================================
        ' 4. SNAPSHOT: Tira a "foto" do estado original para o Dirty Check
        ' =====================================================================
        NomeOriginal = Trim(Rs("Nome"))
        EmailOriginal = Trim(Rs("Email"))
        PerfilOriginal = UCase(Rs("Nivel"))
        StatusOriginal = Rs("Status")
        ' =====================================================================
        
        ' =====================================================================
        ' 5. AUDITORIA: Registra o acesso (leitura) ao cadastro
        ' =====================================================================
        Mdl_Utilitarios.RegistrarAuditoria "LEITURA_USUARIO", "Tbl_Usuarios", CLng(IDUsuario), _
                                           "Acesso ao cadastro do usuário '" & Me.TxtUsuario.Text & "' para visualização/edição."
        
    End If
    
    Rs.Close
    Mdl_Conexao.DesconectarBD
    Exit Sub

ErroCarga:
    ' Registra a falha técnica no banco para análise do desenvolvedor
    Mdl_Utilitarios.GravarLogErro "Usf_EditarUsuario.CarregarDados", Err.Number, Err.Description
    ' Exibe mensagem amigável e padronizada ao usuário
    Mdl_Utilitarios.msgErro "Falha crítica ao tentar carregar os dados. O erro foi registrado no log do sistema." & vbCrLf & "Detalhe: " & Err.Description
    Mdl_Conexao.DesconectarBD
End Sub

' -------------------------------------------------------------------------
' EVENTO: Botão Salvar (Com Validação Visual, Dirty Check e Auditoria)
' -------------------------------------------------------------------------
Private Sub BtnSalvar_Click()
    Dim SQL As String
    Dim PerfilSelecionado As String
    Dim StatusSelecionado As Integer
    Dim NomeTratado As String
    Dim EmailTratado As String
    Dim IDUsuarioAlvo As Long
    Dim DescricaoAuditoria As String
    
    On Error GoTo ErroSalvar
    
    ' =========================================================================
    ' CAMADA 1: SANITIZAÇÃO (Limpeza na Fonte)
    ' =========================================================================
    ' Remove espaços das pontas e espaços duplos do meio
    Me.TxtNome.Text = Application.WorksheetFunction.Trim(Me.TxtNome.Text)
    Me.TxtEmail.Text = Trim(Me.TxtEmail.Text)
    
    ' =========================================================================
    ' CAMADA 2: VALIDAÇÃO DE CAMPOS VAZIOS (Feedback Visual)
    ' =========================================================================
    If Mdl_Utilitarios.CampoVazio(Me.TxtNome, "Preencha o nome completo.") Then Exit Sub
    If Mdl_Utilitarios.CampoVazio(Me.TxtEmail, "Preencha o e-mail.") Then Exit Sub

    ' =========================================================================
    ' CAMADA 3: REGRAS DE NEGÓCIO E FORMATOS
    ' =========================================================================
    ' Regra 1: Exige Nome e Sobrenome (pelo menos um espaço)
    If InStr(Me.TxtNome.Text, " ") = 0 Then
        Mdl_Utilitarios.MsgAviso "Por favor, digite seu nome e sobrenome.", "Cadastro Incompleto"
        Me.TxtNome.SetFocus
        Exit Sub
    End If
    
    ' Regra 2: E-mail Válido (Usa a sua função de segurança)
    If Not Mdl_Seguranca.ValidarEmail(Me.TxtEmail.Text) Then
        Mdl_Utilitarios.MsgAviso "O formato do e-mail é inválido!", "Formato Incorreto"
        Me.TxtEmail.SetFocus
        Exit Sub
    End If

    ' =========================================================================
    ' CAMADA 4: TRADUÇÃO DAS OPÇÕES (UI para Banco)
    ' =========================================================================
    ' Nível / Perfil
    If Me.OptAdmin.Value = True Then
        PerfilSelecionado = "ADMIN"
    ElseIf Me.OptGerente.Value = True Then
        PerfilSelecionado = "GERENTE"
    Else
        PerfilSelecionado = "PADRAO"
    End If
    
    ' Status
    If Me.OptAtivo.Value = True Then StatusSelecionado = 1 Else StatusSelecionado = 0

    ' =========================================================================
    ' CAMADA 5: DIRTY CHECK (Otimização e Bloqueio)
    ' =========================================================================
    ' Compara se houve alguma mudança real em relação aos dados originais carregados
    If Me.TxtNome.Text = NomeOriginal And _
       Me.TxtEmail.Text = EmailOriginal And _
       PerfilSelecionado = PerfilOriginal And _
       StatusSelecionado = StatusOriginal Then
       
        Mdl_Utilitarios.MsgInfo "Nenhuma alteração foi realizada. Os dados permanecem os mesmos.", "Sem Alterações"
        Unload Me
        Exit Sub
    End If

    ' =========================================================================
    ' CAMADA 6: TRATAMENTO DE STRINGS (Prevenção de Injeção SQL)
    ' =========================================================================
    NomeTratado = Replace(Me.TxtNome.Text, "'", "''")
    EmailTratado = Replace(Me.TxtEmail.Text, "'", "''")
    IDUsuarioAlvo = CLng(Me.TxtId.Text)

    ' =========================================================================
    ' CAMADA 7: EXECUÇÃO NO BANCO DE DADOS
    ' =========================================================================
    SQL = "UPDATE Tbl_Usuarios SET " & _
          "Nome = '" & UCase(NomeTratado) & "', " & _
          "Email = '" & UCase(EmailTratado) & "', " & _
          "Nivel = '" & PerfilSelecionado & "', " & _
          "Status = " & StatusSelecionado & " " & _
          "WHERE ID = " & IDUsuarioAlvo
          
    Mdl_Conexao.ConectarBD
    Mdl_Conexao.ExecutarSQL SQL
    
    ' =========================================================================
    ' CAMADA 8: RASTREABILIDADE (Auditoria Corporativa)
    ' =========================================================================
    ' Monta um texto explicando como o cadastro ficou após a edição
    DescricaoAuditoria = "Perfil atualizado para: " & PerfilSelecionado & " | Status: " & IIf(StatusSelecionado = 1, "ATIVO", "INATIVO") & " | E-mail: " & UCase(EmailTratado)
    
    Mdl_Utilitarios.RegistrarAuditoria "UPDATE_USUARIO", "Tbl_Usuarios", IDUsuarioAlvo, DescricaoAuditoria
    
    Mdl_Conexao.DesconectarBD
    
    ' Feedback final e fechamento
    Mdl_Utilitarios.MsgInfo "Dados do usuário atualizados com sucesso!", "Sucesso"
    Unload Me
    Exit Sub

ErroSalvar:
    ' =========================================================================
    ' CAMADA 9: ROTA DE FUGA E LOG DE ERRO
    ' =========================================================================
    ' Registra a falha técnica no banco para análise do desenvolvedor
    Mdl_Utilitarios.GravarLogErro "Usf_EditarUsuario.BtnSalvar_Click", Err.Number, Err.Description
    ' Exibe mensagem amigável ao usuário
    Mdl_Utilitarios.msgErro "Falha crítica ao tentar salvar as alterações. O erro foi registrado no log do sistema." & vbCrLf & "Detalhe: " & Err.Description
    Mdl_Conexao.DesconectarBD
End Sub

' -------------------------------------------------------------------------
' EVENTO: Botão Reset de Senha (Criptografia e Auditoria)
' -------------------------------------------------------------------------
Private Sub BtnResetSenha_Click()
    Dim SQL As String
    Dim SenhaPadraoHashed As String
    Dim IDUsuarioAlvo As Long
    
    ' =========================================================================
    ' 1. CONFIRMAÇÃO DE SEGURANÇA (Prevenção de clique acidental)
    ' =========================================================================
    If MsgBox("Atenção: Esta ação substituirá a senha atual do usuário " & UCase(Me.TxtUsuario.Text) & " pela senha padrão (Senh@1234)." & vbCrLf & vbCrLf & _
              "Deseja realmente continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Reset de Senha") = vbNo Then
        Exit Sub
    End If
    
    ' Captura o ID para usar no SQL e na Auditoria
    IDUsuarioAlvo = CLng(Me.TxtId.Text)
    
    ' =========================================================================
    ' 2. INÍCIO DO PROCESSO COM TRATAMENTO DE ERROS
    ' =========================================================================
    On Error GoTo ErroResetSenha
    
    ' Gera o Hash SHA-256 da senha padrão para não expor texto puro no banco
    SenhaPadraoHashed = Mdl_Seguranca.GerarHashSHA256("Senh@1234")
    
    ' =========================================================================
    ' 3. EXECUÇÃO NO BANCO DE DADOS
    ' =========================================================================
    Mdl_Conexao.ConectarBD
    
    SQL = "UPDATE Tbl_Usuarios SET Senha = '" & SenhaPadraoHashed & "' WHERE ID = " & IDUsuarioAlvo
    Mdl_Conexao.ExecutarSQL SQL
    
    ' =========================================================================
    ' 4. RASTREABILIDADE (Auditoria Corporativa)
    ' =========================================================================
    ' Registra a operação na tabela de auditoria informando o que foi feito
    Mdl_Utilitarios.RegistrarAuditoria "UPDATE_SENHA", "Tbl_Usuarios", IDUsuarioAlvo, _
                                       "Reset de senha realizado pelo Administrador. Senha alterada para o padrão do sistema."
    
    Mdl_Conexao.DesconectarBD
    
    ' Feedback visual para o administrador
    Mdl_Utilitarios.MsgInfo "Senha resetada com sucesso!" & vbCrLf & "O usuário deverá acessar usando: Senha@1234", "Reset Concluído"
    
    ' FECHA A TELA APÓS O RESET (A sua correção de UX!)
    Unload Me
    
    Exit Sub

ErroResetSenha:
    ' =========================================================================
    ' 5. ROTA DE FUGA E LOG DE ERRO (Tratamento Silencioso e Rastreável)
    ' =========================================================================
    ' Grava o erro na tabela do banco usando a sua função utilitária
    ' (O 'Erl' capturará a linha se você usar um numeradores de linha no VBA, senão enviará 0)
    Mdl_Utilitarios.GravarLogErro "Usf_EditarUsuario.BtnResetSenha_Click", Err.Number, Err.Description
    ' Avisa o usuário que algo deu errado, sem estourar a tela amarela do VBA
    Mdl_Utilitarios.msgErro "Falha crítica ao tentar resetar a senha. O erro foi registrado no log do sistema." & vbCrLf & "Detalhe: " & Err.Description
    Mdl_Conexao.DesconectarBD
End Sub
