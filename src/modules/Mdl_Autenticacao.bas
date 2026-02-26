Attribute VB_Name = "Mdl_Autenticacao"
' ==============================================================================
' Módulo: Mdl_Autenticacao
' Objetivo: Validar credenciais, auditar acessos e verificar senha provisória
' ==============================================================================
Option Explicit

Public Function ValidarUsuario(ByVal Usuario As String, ByVal Senha As String) As Boolean
    Dim Rs               As ADODB.Recordset
    Dim SQL              As String
    Dim UsuarioTratado   As String
    Dim SenhaHash        As String
    Dim intStatus        As Integer
    Dim UserIDTemporario As Long
    
    On Error GoTo ErroValidacao
    
    ' 1. Tratamento do Usuário (Evita SQL Injection)
    UsuarioTratado = Replace(Trim(Usuario), "'", "''")
    
    ' 2. SEGURANÇA: Gerar Hash da Senha para comparar com o banco
    SenhaHash = Mdl_Seguranca.GerarHashSHA256(Senha)
    
    ' 3. Monta a consulta
    SQL = "SELECT ID, Nome, Usuario, Nivel, Status FROM Tbl_Usuarios " & _
          "WHERE Usuario = '" & UsuarioTratado & "' AND Senha = '" & SenhaHash & "'"
    
    Set Rs = Mdl_Conexao.ObterRecordset(SQL)
    
    ' Verifica se houve erro na conexão
    If Rs Is Nothing Then
        ValidarUsuario = False
        Exit Function
    End If
    
    ' 4. LÓGICA DE VALIDAÇÃO
    If Not Rs.EOF Then
        ' --- A SENHA ESTÁ CORRETA ---
        intStatus = Val(Mdl_Utilitarios.SafeValue(Rs("Status"), "0"))
        UserIDTemporario = CLng(Mdl_Utilitarios.SafeValue(Rs("ID"), "0"))
                
        If intStatus = 1 Then
            ' --- PREENCHIMENTO CORRETO DAS VARIÁVEIS GLOBAIS ---
            Mdl_VariaveisGlobais.UsuarioID = UserIDTemporario
            Mdl_VariaveisGlobais.UsuarioNome = Mdl_Utilitarios.SafeValue(Rs("Nome"))
            Mdl_VariaveisGlobais.UsuarioLogin = Mdl_Utilitarios.SafeValue(Rs("Usuario"))
            Mdl_VariaveisGlobais.UsuarioNivel = Mdl_Utilitarios.SafeValue(Rs("Nivel"))
            Mdl_VariaveisGlobais.UsuarioLogado = True
            
            ' =================================================================
            ' CHECK DA SENHA PADRÃO (Força a troca no primeiro acesso)
            ' =================================================================
            If Senha = "Senh@1234" Then
                Mdl_VariaveisGlobais.RequerTrocaSenha = True
            Else
                Mdl_VariaveisGlobais.RequerTrocaSenha = False
            End If
            
            ' AUDITORIA: Regista o acesso bem-sucedido
            Mdl_Utilitarios.RegistrarAuditoria "LOGIN_SUCESSO", "Sistema", UserIDTemporario, "Acesso autorizado."
            
            ValidarUsuario = True
        Else
            ' AUDITORIA: Regista a tentativa de uma conta inativa
            Mdl_Utilitarios.RegistrarAuditoria "LOGIN_BLOQUEADO", "Sistema", UserIDTemporario, "Tentativa de login em conta inativa/pendente."
            
            Mdl_Utilitarios.MsgAviso "Seu acesso está pendente ou inativo. Contate o administrador.", "Acesso Negado"
            ValidarUsuario = False
        End If
        
    Else
        ' --- FALHA: Usuário inexistente ou Senha incorreta ---
        ' AUDITORIA: Regista a tentativa de intrusão ou erro de digitação
        Mdl_Utilitarios.RegistrarAuditoria "LOGIN_FALHA", "Sistema", 0, "Falha de credenciais para o usuário digitado: " & UsuarioTratado
        
        ValidarUsuario = False
    End If
    
    ' 5. Limpeza e fechamento seguro
    If Not Rs Is Nothing Then
        If Rs.State = 1 Then Rs.Close ' 1 = adStateOpen
        Set Rs = Nothing
    End If
    Exit Function

ErroValidacao:
    ' LOG DE ERRO: Captura qualquer queda de rede ou falha de leitura
    Mdl_Utilitarios.GravarLogErro "Mdl_Autenticacao.ValidarUsuario", Err.Number, Err.Description
    ValidarUsuario = False
    
    ' Garante que o Recordset é fechado mesmo em caso de erro
    If Not Rs Is Nothing Then
        If Rs.State = 1 Then Rs.Close
        Set Rs = Nothing
    End If
End Function

