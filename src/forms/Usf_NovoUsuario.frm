VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Usf_NovoUsuario 
   Caption         =   "BeautyTech - Novo Usuário"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "Usf_NovoUsuario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Usf_NovoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================
' NOME DO ARQUIVO: Usf_NovoUsuario
' PROJETO:         Sistema BeautyTech - Gestão Integrada
' DESCRIÇÃO:       Tela de Cadastro Rápido de Novos Colaboradores
' ==============================================================================
Option Explicit

Private EfeitoCurso   As Collection ' Cursor Mãozinha
Private SimpleButton  As Collection ' Efeito Visual (Negrito/Tamanho)
Private ColMaiusculas As Collection ' Força caixa alta em TextBoxes
Private ColMascaras   As Collection ' Máscaras de Texto

' ==============================================================================
' 1. INICIALIZAÇÃO E UX VISUAL
' ==============================================================================
Private Sub UserForm_Initialize()
    ' Configurações Físicas (Ajuste a altura conforme a sua tela limpa)
    Me.Height = 280
    Me.Width = 400
    
    ' Personalização da Barra de Título
    Mdl_UI_Efeitos.PersonalizarBarraTitulo Me, RGB(33, 95, 152), RGB(255, 255, 255)
    
    ' Ativação dos Efeitos
    Set EfeitoCurso = Mdl_UI_Efeitos.CriarEfeitosMaozinha(Me)
    Set SimpleButton = Mdl_UI_Efeitos.CriarSimpleButton(Me)
    Set ColMaiusculas = Mdl_UI_Efeitos.AtivarMaiusculas(Me)
    Set ColMascaras = Mdl_UI_Efeitos.AtivarMascaras(Me)
    
    ' Regra de Negócio: Seleciona o perfil PADRAO por default para agilizar
    Me.OptPadrao.Value = True
    
    ' Foco inicial no primeiro campo
    Me.TxtNome.SetFocus
End Sub

' Bloqueia o fechamento pelo "X" do Windows
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        Mdl_Utilitarios.MsgAviso "Por favor, utilize os botões 'Salvar' ou 'Cancelar' para fechar o formulário.", "Ação Bloqueada"
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Mdl_UI_Efeitos.LimparFoco
End Sub

' ==============================================================================
' 2. AÇÕES DOS BOTÕES
' ==============================================================================
Private Sub BtnCancelar_Click()
    Unload Me
End Sub

' -------------------------------------------------------------------------
' EVENTO: Botão Salvar (Com Validação Visual e Regras de Negócio)
' -------------------------------------------------------------------------
Private Sub BtnSalvar_Click()
    Dim SQL As String
    Dim PerfilSelecionado As String
    Dim NomeTratado As String, EmailTratado As String, UsuarioTratado As String
    Dim SenhaPadraoHashed As String
    Dim Rs As Object
    
    On Error GoTo ErroCadastro
    
    ' =========================================================================
    ' CAMADA 1: SANITIZAÇÃO (Limpeza na Fonte)
    ' =========================================================================
    ' Limpa espaços das pontas e duplos no meio do nome
    Me.TxtNome.Text = Application.WorksheetFunction.Trim(Me.TxtNome.Text)
    Me.TxtUsuario.Text = Trim(Me.TxtUsuario.Text)
    Me.TxtEmail.Text = Trim(Me.TxtEmail.Text)
    
    ' =========================================================================
    ' CAMADA 2: VALIDAÇÃO DE CAMPOS VAZIOS (Feedback Visual Rosa/Branco)
    ' =========================================================================
    ' Se a função retornar True (vazio), ela mesma avisa, pinta a caixa e aborta o código.
    If Mdl_Utilitarios.CampoVazio(Me.TxtNome, "Preencha o nome completo.") Then Exit Sub
    If Mdl_Utilitarios.CampoVazio(Me.TxtUsuario, "Preencha o nome de usuário.") Then Exit Sub
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
    ' CAMADA 4: PREPARAÇÃO DE DADOS E PREVENÇÃO SQL INJECTION
    ' =========================================================================
    ' Perfil
    If Me.OptAdmin.Value = True Then
        PerfilSelecionado = "ADMIN"
    ElseIf Me.OptGerente.Value = True Then
        PerfilSelecionado = "GERENTE"
    Else
        PerfilSelecionado = "PADRAO"
    End If
    
    ' Tratamento de aspas simples para o Access não quebrar
    NomeTratado = Replace(Me.TxtNome.Text, "'", "''")
    EmailTratado = Replace(Me.TxtEmail.Text, "'", "''")
    UsuarioTratado = Replace(Me.TxtUsuario.Text, "'", "''")

    ' =========================================================================
    ' CAMADA 5: VERIFICAÇÃO DE DUPLICIDADE NO BANCO
    ' =========================================================================
    Mdl_Conexao.ConectarBD
    
    ' Procura se o Login ou E-mail já existem (ativos ou inativos)
    SQL = "SELECT ID FROM Tbl_Usuarios WHERE Usuario = '" & UCase(UsuarioTratado) & "' OR Email = '" & UCase(EmailTratado) & "'"
    Set Rs = Mdl_Conexao.ObterRecordset(SQL)
    
    If Not Rs.EOF Then
        Mdl_Utilitarios.MsgAviso "Este Nome de Usuário ou E-mail já está em uso por outro cadastro.", "Duplicidade Encontrada"
        Rs.Close
        Mdl_Conexao.DesconectarBD
        Exit Sub
    End If
    Rs.Close

    ' =========================================================================
    ' CAMADA 6: INSERÇÃO NO BANCO DE DADOS (INSERT INTO)
    ' =========================================================================
    ' Gera a senha padrão já criptografada para o banco
    SenhaPadraoHashed = Mdl_Seguranca.GerarHashSHA256("Senh@1234")
    
    ' O Status entra hardcoded como 1 (Ativo) e a DataCadastro usa a função Now()
    SQL = "INSERT INTO Tbl_Usuarios (Nome, Usuario, Email, Senha, Nivel, Status, DataCadastro) VALUES (" & _
          "'" & UCase(NomeTratado) & "', " & _
          "'" & UCase(UsuarioTratado) & "', " & _
          "'" & UCase(EmailTratado) & "', " & _
          "'" & SenhaPadraoHashed & "', " & _
          "'" & PerfilSelecionado & "', " & _
          "1, " & _
          "#" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "#)"
          
    Mdl_Conexao.ExecutarSQL SQL
    
    ' =========================================================================
    ' CAMADA 7: AUDITORIA E ENCERRAMENTO
    ' =========================================================================
    Mdl_Utilitarios.RegistrarAuditoria "NOVO_USUARIO", "Tbl_Usuarios", 0, "Novo utilizador criado: " & UCase(UsuarioTratado) & " | Perfil: " & PerfilSelecionado
    
    Mdl_Conexao.DesconectarBD
    
    Mdl_Utilitarios.MsgInfo "Utilizador cadastrado com sucesso!" & vbCrLf & "Login: " & UCase(UsuarioTratado) & vbCrLf & "Senha Provisória: Senh@1234", "Sucesso"
    
    Unload Me
    Exit Sub

ErroCadastro:
    Mdl_Utilitarios.GravarLogErro "Usf_NovoUsuario.BtnSalvar_Click", Err.Number, Err.Description
    Mdl_Utilitarios.msgErro "Falha crítica ao tentar cadastrar o usuário. O erro foi registrado no log do sistema." & vbCrLf & "Detalhe: " & Err.Description
    Mdl_Conexao.DesconectarBD
End Sub

