Attribute VB_Name = "Mdl_Utilitarios"
Option Explicit
' ==============================================================================
' Módulo: Mdl_Utilitarios
' Objetivo: Kit de ferramentas completo (Mensagens, Validações, Logs e Conversões)
' Dependências: Mdl_Conexao, Mdl_VariaveisGlobais
' ==============================================================================

' ==============================================================================
' 1) MENSAGENS PADRONIZADAS
' ==============================================================================
Public Sub MsgInfo(ByVal Texto As String, Optional ByVal Titulo As String = "Informação")
    MsgBox Texto, vbInformation, Titulo
End Sub

Public Sub msgErro(ByVal Texto As String, Optional ByVal Titulo As String = "Erro")
    MsgBox Texto, vbCritical, Titulo
End Sub

Public Sub MsgAviso(ByVal Texto As String, Optional ByVal Titulo As String = "Atenção")
    MsgBox Texto, vbExclamation, Titulo
End Sub

' ==============================================================================
' 2) VALIDAÇÃO DE CAMPOS (VISUAL)
' ==============================================================================

' -----------------------------------------------------------
' CampoVazio: Verifica se está vazio, avisa e pinta de rosa
' Retorna: TRUE se houver erro (vazio)
' -----------------------------------------------------------
Public Function CampoVazio(ByVal Campo As Object, ByVal Mensagem As String) As Boolean
    On Error GoTo TrataErro
    
    If Trim(SafeValue(Campo)) = "" Then
        PintarErro Campo, True
        MsgAviso Mensagem, "Campo Obrigatório"
        Campo.SetFocus
        CampoVazio = True
    Else
        PintarErro Campo, False
        CampoVazio = False
    End If
    Exit Function

TrataErro:
    Call GravarLogErro("CampoVazio", Err.Number, Err.Description, Erl)
    CampoVazio = True
End Function

' -----------------------------------------------------------
' CampoData: Verifica se é uma data válida (dd/mm/yyyy), avisa e pinta
' Retorna: TRUE se houver erro (data inválida)
' -----------------------------------------------------------
Public Function CampoData(ByVal Campo As Object, ByVal Mensagem As String) As Boolean
    On Error GoTo TrataErro
    
    Dim Valor As String
    Valor = Trim(SafeValue(Campo))
    
    ' Se estiver vazio, consideramos erro ou não?
    ' Geralmente data é obrigatória. Se for opcional, teríamos que adaptar.
    ' Aqui assumo que se chamou CampoData, é porque precisa ter data.
    
    If Not IsDataValida(Valor) Then
        PintarErro Campo, True
        MsgAviso Mensagem, "Data Inválida"
        Campo.SetFocus
        CampoData = True
    Else
        PintarErro Campo, False
        CampoData = False
    End If
    Exit Function

TrataErro:
    Call GravarLogErro("CampoData", Err.Number, Err.Description, Erl)
    CampoData = True
End Function

' -----------------------------------------------------------
' Auxiliar Privada: Pinta o fundo do controle
' -----------------------------------------------------------
Private Sub PintarErro(ByVal Campo As Object, ByVal TemErro As Boolean)
    On Error Resume Next ' Previne erro se o controle não tiver BackColor
    If TemErro Then
        Campo.BackColor = RGB(255, 220, 220) ' Rosa Claro
    Else
        Campo.BackColor = vbWhite
    End If
End Sub

' ==============================================================================
' 3) LÓGICA DE VALIDAÇÃO(BACKEND) IsDataValida: Validação estrita de datas (dd/mm/yyyy)
' ==============================================================================
Public Function IsDataValida(ByVal DataTexto As String) As Boolean
    On Error GoTo TrataErro

    Dim Dia As Integer, Mes As Integer, Ano As Integer
    Dim s As String

    IsDataValida = False
    s = Trim(DataTexto)

    ' 1. Formato Básico
    If Len(s) <> 10 Then Exit Function
    If Mid$(s, 3, 1) <> "/" Or Mid$(s, 6, 1) <> "/" Then Exit Function
    If Not s Like "##/##/####" Then Exit Function

    ' 2. Quebra
    Dia = val(Left$(s, 2))
    Mes = val(Mid$(s, 4, 2))
    Ano = val(Right$(s, 4))

    ' 3. Limites Básicos
    If Dia < 1 Or Dia > 31 Then Exit Function
    If Mes < 1 Or Mes > 12 Then Exit Function
    If Ano < 1900 Then Exit Function ' Regra de negócio: Nada antes de 1900

    ' 4. Meses com 30 dias
    Select Case Mes
        Case 4, 6, 9, 11
            If Dia > 30 Then Exit Function
    End Select

    ' 5. Fevereiro e Bissexto
    If Mes = 2 Then
        If ((Ano Mod 4 = 0 And Ano Mod 100 <> 0) Or (Ano Mod 400 = 0)) Then
            If Dia > 29 Then Exit Function
        Else
            If Dia > 28 Then Exit Function
        End If
    End If

    IsDataValida = True
    Exit Function

TrataErro:
    Call GravarLogErro("IsDataValida", Err.Number, Err.Description, Erl)
    IsDataValida = False
End Function

' ==============================================================================
' 4) CONVERSÃO E SEGURANÇA
' ==============================================================================

Public Function SafeValue(ByVal v As Variant, Optional ByVal defaultValue As String = "") As String
    On Error GoTo ErrHandler
    
    If IsObject(v) Then
        If v Is Nothing Then SafeValue = defaultValue: Exit Function
        On Error Resume Next
        SafeValue = Trim$(CStr(v.Value))
        If Len(SafeValue) = 0 Then SafeValue = defaultValue
        Exit Function
    End If
    
    If IsNull(v) Or IsEmpty(v) Then SafeValue = defaultValue: Exit Function
    
    If IsDate(v) Then
        SafeValue = Format(v, "dd/mm/yyyy hh:nn:ss")
    Else
        SafeValue = Trim$(CStr(v))
    End If
    
    If Len(SafeValue) = 0 Then SafeValue = defaultValue
    Exit Function

ErrHandler:
    SafeValue = defaultValue
End Function

Public Sub TrimTodosCampos(ParamArray Campos() As Variant)
    On Error Resume Next ' Evita erro se o controle estiver vazio ou nulo
    Dim i As Integer
    For i = LBound(Campos) To UBound(Campos)
        If TypeName(Campos(i)) = "String" Then
            Campos(i) = Trim(Campos(i))
        ElseIf TypeName(Campos(i)) Like "*TextBox" Or TypeName(Campos(i)) Like "*ComboBox" Then
            Campos(i).Text = Trim(Campos(i).Text)
        End If
    Next i
    On Error GoTo 0
End Sub

' ==============================================================================
' 5) LOGS E AUDITORIA (BANCO DE DADOS)
' ==============================================================================

Public Sub GravarLogErro(Optional ByVal Modulo As String = "", Optional ByVal NumeroErro As Long = 0, Optional ByVal DescricaoErro As String = "", Optional ByVal LinhaErro As Long = 0)
    On Error Resume Next
    Dim SQL As String
    ' Sanitização básica
    DescricaoErro = Replace(DescricaoErro, "'", "''")
    
    SQL = "INSERT INTO Tbl_LogErro (DataHora, Usuario, NomeMaquina, Modulo, NumeroErro, LinhaErro, DescricaoErro) " & _
          "VALUES (#" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "#, '" & _
          ObterUsuarioAtual() & "', '" & ObterMaquinaAtual() & "', '" & _
          Modulo & "', " & NumeroErro & ", " & LinhaErro & ", '" & DescricaoErro & "')"
    
    Mdl_Conexao.ExecutarSQL SQL
End Sub

Public Sub RegistrarAuditoria(ByVal TipoOperacao As String, ByVal Tabela As String, ByVal RegistroID As Long, ByVal Descricao As String)
    On Error Resume Next
    Dim SQL As String
    Descricao = Replace(Descricao, "'", "''")
    
    SQL = "INSERT INTO Tbl_Auditoria (DataHora, TipoOperacao, Tabela, RegistroID, Descricao, Usuario, Maquina) " & _
          "VALUES (#" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "#, '" & _
          TipoOperacao & "', '" & Tabela & "', " & RegistroID & ", " & _
          "'" & Descricao & "', '" & ObterUsuarioAtual() & "', '" & ObterMaquinaAtual() & "')"
          
    Mdl_Conexao.ExecutarSQL SQL
End Sub

Public Sub RegistrarLogAcesso(ByVal UsuarioTentativa As String, ByVal Status As String)
    On Error Resume Next
    Dim SQL As String
    UsuarioTentativa = Replace(UsuarioTentativa, "'", "''")
    
    SQL = "INSERT INTO Tbl_LogAcesso (DataHora, Usuario, NomeMaquina, Status) " & _
          "VALUES (#" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "#, '" & _
          UsuarioTentativa & "', '" & ObterMaquinaAtual() & "', '" & Status & "')"
          
    Mdl_Conexao.ExecutarSQL SQL
End Sub

' --- Helpers Privados ---
Private Function ObterUsuarioAtual() As String
    If Mdl_VariaveisGlobais.UsuarioNome <> "" Then
        ObterUsuarioAtual = Replace(Mdl_VariaveisGlobais.UsuarioNome, "'", "''")
    Else
        ObterUsuarioAtual = "Sistema"
    End If
End Function

Private Function ObterMaquinaAtual() As String
    ObterMaquinaAtual = Replace(Environ$("COMPUTERNAME"), "'", "''")
End Function

