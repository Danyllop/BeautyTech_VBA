Attribute VB_Name = "Mdl_InstalacaoBD"
' ==============================================================================
' NOME DO ARQUIVO: Mdl_InstalacaoBD
' PROJETO:         Sistema BeautyTech - Gestão Integrada
' DESCRIÇÃO:       Garante a Integridade e Auto-Cura da Estrutura do Banco
' AUTOR:           LogicUp Solutions
' DATA:            26/02/2026
' ==============================================================================
Option Explicit

' ------------------------------------------------------------------------------
' 1. ROTINA MESTRE: VERIFICAR ESTRUTURA
' ------------------------------------------------------------------------------
Public Sub VerificarEstruturaBanco()
    On Error GoTo ErroGeral
    
    ' Garante a conexão física ativa antes de validar tabelas
    Mdl_Conexao.ConectarBD
    
    ' --- 1. TABELAS DE INFRAESTRUTURA (Logs) ---
    ' Criamos estas primeiro para que as outras possam registrar falhas, se houver
    If Not TabelaExiste("Tbl_LogErro") Then CriarTabelaLogErro
    If Not TabelaExiste("Tbl_LogAcesso") Then CriarTabelaLogAcesso
    If Not TabelaExiste("Tbl_Auditoria") Then CriarTabelaAuditoria
    
    ' --- 2. GESTÃO DE ACESSOS ---
    If Not TabelaExiste("Tbl_Usuarios") Then
        CriarTabelaUsuarios
        CriarAdminPadrao ' Garante o primeiro acesso ao sistema
    End If
    
    ' --- 3. TABELAS DE NEGÓCIO (Core) ---
    If Not TabelaExiste("Tbl_Clientes") Then CriarTabelaClientes
    If Not TabelaExiste("Tbl_Servicos") Then CriarTabelaServicos
    If Not TabelaExiste("Tbl_Agendamentos") Then CriarTabelaAgendamentos
    If Not TabelaExiste("Tbl_Movimentacao") Then CriarTabelaMovimentacao
    
    Exit Sub
    
ErroGeral:
    ' Falha crítica antes das tabelas de log existirem
    MsgBox "Erro crítico ao inicializar estrutura de dados: " & Err.Description, vbCritical, "BeautyTech - Setup"
End Sub

' ------------------------------------------------------------------------------
' 2. AUXILIAR: VERIFICA SE TABELA EXISTE NO ACCESS
' ------------------------------------------------------------------------------
Private Function TabelaExiste(ByVal NomeTabela As String) As Boolean
    Dim Rs As Object
    On Error Resume Next
    
    ' Teste de baixo custo de processamento
    Set Rs = Mdl_Conexao.Conexao.Execute("SELECT TOP 1 * FROM " & NomeTabela)
    
    If Err.Number = 0 Then
        TabelaExiste = True
        Rs.Close
    Else
        TabelaExiste = False
        Err.Clear
    End If
    Set Rs = Nothing
End Function

' ------------------------------------------------------------------------------
' 3. CRIAÇÃO DE TABELAS DE INFRAESTRUTURA (Logs e Auditoria)
' ------------------------------------------------------------------------------

Private Sub CriarTabelaLogErro()
    Dim SQL As String
    ' Removida a coluna [LinhaErro] para simplificação do banco
    SQL = "CREATE TABLE Tbl_LogErro (" & _
          "[ID] AUTOINCREMENT PRIMARY KEY, " & _
          "[DataHora] DATETIME, " & _
          "[Usuario] TEXT(150), " & _
          "[NomeMaquina] TEXT(100), " & _
          "[Modulo] TEXT(100), " & _
          "[NumeroErro] LONG, " & _
          "[DescricaoErro] MEMO)"
    Mdl_Conexao.Conexao.Execute SQL
End Sub

Private Sub CriarTabelaLogAcesso()
    Dim SQL As String
    SQL = "CREATE TABLE Tbl_LogAcesso (" & _
          "[ID] AUTOINCREMENT PRIMARY KEY, " & _
          "[DataHora] DATETIME, " & _
          "[Usuario] TEXT(100), " & _
          "[NomeMaquina] TEXT(100), " & _
          "[Status] TEXT(50))"
    Mdl_Conexao.Conexao.Execute SQL
End Sub

Private Sub CriarTabelaAuditoria()
    Dim SQL As String
    ' Corrigido: TipoOperacao expandido para evitar erro de truncamento
    SQL = "CREATE TABLE Tbl_Auditoria (" & _
          "[ID] AUTOINCREMENT PRIMARY KEY, " & _
          "[DataHora] DATETIME, " & _
          "[TipoOperacao] TEXT(200), " & _
          "[Tabela] TEXT(150), " & _
          "[RegistroID] LONG, " & _
          "[Usuario] TEXT(150), " & _
          "[Maquina] TEXT(100), " & _
          "[Descricao] MEMO)" ' MEMO para logs de auditoria detalhados
    Mdl_Conexao.Conexao.Execute SQL
End Sub

' ------------------------------------------------------------------------------
' 4. CRIAÇÃO DE TABELAS DE USUÁRIOS E ADMIN PADRÃO
' ------------------------------------------------------------------------------

Private Sub CriarTabelaUsuarios()
    Dim SQL As String
    ' Estrutura validada para senhas SHA-256 e Status Integer
    SQL = "CREATE TABLE Tbl_Usuarios (" & _
          "[ID] AUTOINCREMENT PRIMARY KEY, " & _
          "[Nome] TEXT(150), " & _
          "[Usuario] TEXT(50) NOT NULL UNIQUE, " & _
          "[Email] TEXT(100), " & _
          "[Senha] TEXT(100), " & _
          "[Nivel] TEXT(20), " & _
          "[Status] INTEGER DEFAULT 0, " & _
          "[DataCadastro] DATETIME)"
    Mdl_Conexao.Conexao.Execute SQL
End Sub

Private Sub CriarAdminPadrao()
    Dim SQL As String
    Dim SenhaHash As String
    
    ' Senha master encriptada
    SenhaHash = Mdl_Seguranca.GerarHashSHA256("Admin@123")
    
    ' Status 1 = Ativo para o Administrador inicial
    SQL = "INSERT INTO Tbl_Usuarios (Nome, Usuario, Email, Senha, Nivel, Status, DataCadastro) " & _
          "VALUES ('Administrador do Sistema', 'ADMINISTRADOR', 'admin@beautytech.com', " & _
          "'" & SenhaHash & "', 'ADMIN', 1, #" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "#)"
          
    Mdl_Conexao.Conexao.Execute SQL
End Sub

' ------------------------------------------------------------------------------
' 5. TABELAS DE NEGÓCIO (Core do Sistema)
' ------------------------------------------------------------------------------

Private Sub CriarTabelaClientes()
    Dim SQL As String
    SQL = "CREATE TABLE Tbl_Clientes (" & _
          "[ID] AUTOINCREMENT PRIMARY KEY, " & _
          "[Nome] TEXT(160) NOT NULL, " & _
          "[Telefone] TEXT(50), " & _
          "[Email] TEXT(100), " & _
          "[DataNascimento] DATETIME, " & _
          "[Observacoes] MEMO, " & _
          "[DataCadastro] DATETIME DEFAULT Now(), " & _
          "[Ativo] BIT DEFAULT -1)" ' -1 representa TRUE no Access
    Mdl_Conexao.Conexao.Execute SQL
End Sub

Private Sub CriarTabelaServicos()
    Dim SQL As String
    SQL = "CREATE TABLE Tbl_Servicos (" & _
          "[ID] AUTOINCREMENT PRIMARY KEY, " & _
          "[Nome] TEXT(100) NOT NULL, " & _
          "[Descricao] MEMO, " & _
          "[Valor] CURRENCY DEFAULT 0, " & _
          "[DuracaoMinutos] INTEGER DEFAULT 30, " & _
          "[ComissaoPercentual] DOUBLE DEFAULT 0, " & _
          "[Ativo] BIT DEFAULT -1)"
    Mdl_Conexao.Conexao.Execute SQL
End Sub

Private Sub CriarTabelaAgendamentos()
    Dim SQL As String
    ' Status: 0=Pendente, 1=Confirmado, 2=Cancelado, 3=Concluido
    SQL = "CREATE TABLE Tbl_Agendamentos (" & _
          "[ID] AUTOINCREMENT PRIMARY KEY, " & _
          "[ClienteID] LONG, " & _
          "[ServicoID] LONG, " & _
          "[ProfissionalID] LONG, " & _
          "[DataHoraInicio] DATETIME, " & _
          "[DataHoraFim] DATETIME, " & _
          "[Status] INTEGER DEFAULT 0, " & _
          "[Observacoes] MEMO, " & _
          "[ValorCobrado] CURRENCY DEFAULT 0, " & _
          "[ComissaoValor] CURRENCY DEFAULT 0)"
    Mdl_Conexao.Conexao.Execute SQL
End Sub

Private Sub CriarTabelaMovimentacao()
    Dim SQL As String
    ' Tipo: 1=Receita, 2=Despesa
    SQL = "CREATE TABLE Tbl_Movimentacao (" & _
          "[ID] AUTOINCREMENT PRIMARY KEY, " & _
          "[AgendamentoID] LONG, " & _
          "[Tipo] INTEGER, " & _
          "[Valor] CURRENCY, " & _
          "[DataMovimento] DATETIME DEFAULT Now(), " & _
          "[Descricao] TEXT(255), " & _
          "[Categoria] TEXT(50))"
    Mdl_Conexao.Conexao.Execute SQL
End Sub

