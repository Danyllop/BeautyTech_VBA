Attribute VB_Name = "Mdl_Autenticacao"
Option Explicit
' ==============================================================================
' Módulo: Mdl_Autenticacao
' Objetivo: Validar credenciais usando SafeValue para proteção
' Dependências: Mdl_Conexao, Mdl_VariaveisGlobais, Mdl_Utilitarios
' ==============================================================================

Public Function ValidarUsuario(ByVal Usuario As String, ByVal Senha As String) As Boolean
    Dim Rs               As ADODB.Recordset
    Dim SQL              As String
    Dim UsuarioTratado   As String
    Dim SenhaHash        As String
    Dim intStatus        As Integer
    
    ' 1. Tratamento do Usuário (Evita SQL Injection)
    UsuarioTratado = Replace(Usuario, "'", "''")
    
    ' 2. SEGURANÇA: Gerar Hash da Senha para comparar com o banco
    SenhaHash = Mdl_Seguranca.GerarHashSHA256(Senha)
    
    ' 3. Monta a consulta (Note que agora incluímos o campo Status)
    SQL = "SELECT ID, Nome,Usuario, Nivel, Status FROM Tbl_Usuarios " & _
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
        ' Verificamos se a conta está ativa (Status = 1)
        intStatus = val(Mdl_Utilitarios.SafeValue(Rs("Status"), "0"))
               
        If intStatus = 1 Then
            ' --- PREENCHIMENTO CORRETO DAS VARIÁVEIS ---
            Mdl_VariaveisGlobais.UsuarioID = CLng(Mdl_Utilitarios.SafeValue(Rs("ID"), "0"))
            
            ' Guarda o Nome Completo (Para relatórios futuros)
            Mdl_VariaveisGlobais.UsuarioNome = Mdl_Utilitarios.SafeValue(Rs("Nome"))
            
            ' NOVA VARIÁVEL: Guarda o Login (Para a Label do Menu)
            Mdl_VariaveisGlobais.UsuarioLogin = Mdl_Utilitarios.SafeValue(Rs("Usuario"))
            
            Mdl_VariaveisGlobais.UsuarioNivel = Mdl_Utilitarios.SafeValue(Rs("Nivel"))
            
            ' Mantém True apenas para controle de acesso
            Mdl_VariaveisGlobais.UsuarioLogado = True
            
            ValidarUsuario = True
        Else
            Mdl_Utilitarios.MsgAviso "Acesso pendente.", "Aviso"
            ValidarUsuario = False
        End If
        
    Else
        ' --- FALHA: Usuário inexistente ou Senha incorreta ---
        ' Para segurança, não dizemos qual dos dois está errado aqui.
        ValidarUsuario = False
    End If
    
    ' 5. Limpeza e fechamento seguro
    If Not Rs Is Nothing Then
        If Rs.State = 1 Then Rs.Close ' 1 = adStateOpen
        Set Rs = Nothing
    End If

End Function

