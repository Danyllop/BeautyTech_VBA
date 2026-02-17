Attribute VB_Name = "Mdl_InstalacaoBD"
Option Explicit
' ==============================================================================
' Modulo: Mdl_InstalacaoBD
' Objetivo: Garantir estrutura do Banco (Agora com nomes em PascalCase: Tbl_...)
' ==============================================================================

Public Sub VerificarEstruturaBanco()
    On Error GoTo ErroGeral
    
    Mdl_Conexao.ConectarBD
    
    ' Nota: O Access nao diferencia maiusculas na verificacao,
    ' mas na criacao (CREATE TABLE) ele respeita o que escrevemos.
    
    ' --- 1. Tabela de Usuarios ---
    If Not TabelaExiste("Tbl_Usuarios") Then
        CriarTabelaUsuarios
        CriarAdminPadrao
    End If
    
    ' --- 2. Logs de Sistema ---
    If Not TabelaExiste("Tbl_LogErro") Then CriarTabelaLogErro
    If Not TabelaExiste("Tbl_LogAcesso") Then CriarTabelaLogAcesso
    If Not TabelaExiste("Tbl_Auditoria") Then CriarTabelaAuditoria
    
    ' --- 3. Tabelas de Negocio (Core) ---
    If Not TabelaExiste("Tbl_Clientes") Then CriarTabelaClientes
    If Not TabelaExiste("Tbl_Servicos") Then CriarTabelaServicos
    If Not TabelaExiste("Tbl_Agendamentos") Then CriarTabelaAgendamentos
    If Not TabelaExiste("Tbl_Movimentacao") Then CriarTabelaMovimentacao
    
    Exit Sub
    
ErroGeral:
    MsgBox "Erro ao verificar estrutura: " & Err.Description, vbCritical
End Sub

Private Function TabelaExiste(ByVal NomeTabela As String) As Boolean
    Dim Rs As Object
    On Error Resume Next
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

Private Sub CriarTabelaUsuarios()
    Dim SQL As String
    
    ' Atualizacoes importantes:
    ' 1. Email TEXT(100) como solicitado.
    ' 2. Senha TEXT(100) para comportar os 64 caracteres do Hash SHA-256.
    ' 3. Status como INTEGER (0 = Inativo, 1 = Ativo) para maior flexibilidade.
    ' 4. DataCadastro DATETIME para auditoria.
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
    Dim SQL         As String
    Dim SenhaHash   As String
    
    ' 1. Geramos o Hash da senha padrao para que o login funcione
    SenhaHash = Mdl_Seguranca.GerarHashSHA256("Admin@123")
    
    ' 2. Montamos o SQL com todos os novos campos (Email, Status=1, DataCadastro)
    ' Nota: O Status deve ser 1 para o Admin ja nascer liberado.
    SQL = "INSERT INTO Tbl_Usuarios (Nome, Usuario, Email, Senha, Nivel, Status, DataCadastro) " & _
          "VALUES (" & _
          "'Administrador do Sistema', " & _
          "'ADMIN', " & _
          "'admin@exemplo.com', " & _
          "'" & SenhaHash & "', " & _
          "'ADMIN', " & _
          "1, " & _
          "#" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "#)"
          
    Mdl_Conexao.Conexao.Execute SQL
    
    Debug.Print "Admin padrao criado com sucesso!"
End Sub

Private Sub CriarTabelaLogErro()
    Dim SQL As String
    SQL = "CREATE TABLE Tbl_LogErro (" & _
          "[ID] AUTOINCREMENT PRIMARY KEY, " & _
          "[DataHora] DATETIME, " & _
          "[Usuario] TEXT(100), " & _
          "[NomeMaquina] TEXT(50), " & _
          "[Modulo] TEXT(100), " & _
          "[NumeroErro] LONG, " & _
          "[LinhaErro] LONG, " & _
          "[DescricaoErro] MEMO)"
    Mdl_Conexao.Conexao.Execute SQL
End Sub

Private Sub CriarTabelaLogAcesso()
    Dim SQL As String
    SQL = "CREATE TABLE Tbl_LogAcesso (" & _
          "[ID] AUTOINCREMENT PRIMARY KEY, " & _
          "[DataHora] DATETIME, " & _
          "[Usuario] TEXT(50), " & _
          "[NomeMaquina] TEXT(50), " & _
          "[Status] TEXT(50))"
    Mdl_Conexao.Conexao.Execute SQL
End Sub

Private Sub CriarTabelaAuditoria()
    Dim SQL As String
    SQL = "CREATE TABLE Tbl_Auditoria (" & _
          "[ID] AUTOINCREMENT PRIMARY KEY, " & _
          "[DataHora] DATETIME, " & _
          "[TipoOperacao] TEXT(20), " & _
          "[Tabela] TEXT(50), " & _
          "[RegistroID] LONG, " & _
          "[Usuario] TEXT(50), " & _
          "[Maquina] TEXT(50), " & _
          "[Descricao] MEMO)"
    Mdl_Conexao.Conexao.Execute SQL
End Sub

' --- Tabelas de Negocio ---
Private Sub CriarTabelaClientes()
    Dim SQL As String
    SQL = "CREATE TABLE Tbl_Clientes (" & _
          "[ID] AUTOINCREMENT PRIMARY KEY, " & _
          "[Nome] TEXT(160) NOT NULL, " & _
          "[Telefone] TEXT(20), " & _
          "[Email] TEXT(100), " & _
          "[DataNascimento] DATETIME, " & _
          "[Observacoes] MEMO, " & _
          "[DataCadastro] DATETIME DEFAULT Now(), " & _
          "[Ativo] BIT DEFAULT -1)" ' Access TRUE is -1
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
    ' Status pode ser: 0=Pendente, 1=Confirmado, 2=Cancelado, 3=Concluido
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
