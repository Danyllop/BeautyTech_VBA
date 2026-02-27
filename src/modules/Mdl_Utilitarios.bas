Attribute VB_Name = "Mdl_Utilitarios"
' ==============================================================================
' NOME DO ARQUIVO: Mdl_Utilitarios
' PROJETO:         Sistema BeautyTech - Gestão Integrada
' DESCRIÇÃO:       Kit de ferramentas (Mensagens, Validações, Logs e Metadados)
' DEPENDÊNCIAS:    Mdl_Conexao, Mdl_VariaveisGlobais
' AUTOR:           LogicUp Solutions
' DATA:            26/02/2026
' ==============================================================================
Option Explicit

Private CoresOriginais As New Collection

' ==============================================================================
' SEÇÃO 1: COMUNICAÇÃO COM O USUÁRIO (Mensagens Padronizadas)
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
' SEÇÃO 2: VALIDAÇÃO DE INTERFACE (Feedback Visual e Regras)
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
    Call GravarLogErro("CampoVazio", Err.Number, Err.Description)
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
    Call GravarLogErro("CampoData", Err.Number, Err.Description)
    CampoData = True
End Function

' ------------------------------------------------------------------------------
' Auxiliar Privada: Gerencia feedback visual usando Coleção Interna
' ------------------------------------------------------------------------------
Private Sub PintarErro(ByVal Campo As Object, ByVal TemErro As Boolean)
    Dim Chave As String
    ' Criamos uma chave única combinando o nome do formulário e do campo
    Chave = Campo.Parent.Name & "_" & Campo.Name
    
    On Error Resume Next ' Previne erros em controles sem BackColor
    
    If TemErro Then
        ' 1. Se o erro apareceu, tentamos salvar a cor original na Coleção
        ' Usamos a tentativa de Add; se já existir, o erro é ignorado pelo Resume Next
        CoresOriginais.Add Campo.BackColor, Chave
        
        ' 2. Aplica a cor de Alerta (Rosa Claro)
        Campo.BackColor = RGB(255, 220, 220)
    Else
        ' 3. Se o erro foi corrigido, recuperamos a cor da Coleção
        Dim CorSalva As Long
        CorSalva = CoresOriginais(Chave)
        
        If Err.Number = 0 Then
            Campo.BackColor = CorSalva
            ' 4. Removemos da coleção para manter a memória limpa
            CoresOriginais.Remove Chave
        Else
            ' Fallback: Se por algum motivo a coleção falhar, volta para o Dark Padrão
            Campo.BackColor = RGB(33, 47, 61)
            Err.Clear
        End If
    End If
End Sub

' ==============================================================================
' SEÇÃO 3: LÓGICA DE PROCESSAMENTO (Backend e Conversões)
' ==============================================================================

' -----------------------------------------------------------
' IsDataValida: Validação estrita de datas (dd/mm/yyyy)
' -----------------------------------------------------------
Public Function IsDataValida(ByVal DataTexto As String) As Boolean
    On Error GoTo TrataErro

    Dim Dia As Integer, Mes As Integer, Ano As Integer
    Dim s As String

    IsDataValida = False
    s = Trim(DataTexto)

    ' 1. Validação de Formato e Máscara
    If Len(s) <> 10 Then Exit Function
    If Mid$(s, 3, 1) <> "/" Or Mid$(s, 6, 1) <> "/" Then Exit Function
    If Not s Like "##/##/####" Then Exit Function

    ' 2. Extração de Componentes
    Dia = Val(Left$(s, 2))
    Mes = Val(Mid$(s, 4, 2))
    Ano = Val(Right$(s, 4))

    ' 3. Limites Lógicos
    If Dia < 1 Or Dia > 31 Then Exit Function
    If Mes < 1 Or Mes > 12 Then Exit Function
    If Ano < 1900 Then Exit Function ' Regra: Dados históricos aceitáveis apenas pós-1900

    ' 4. Validação de Meses com 30 dias
    Select Case Mes
        Case 4, 6, 9, 11
            If Dia > 30 Then Exit Function
    End Select

    ' 5. Tratamento de Ano Bissexto (Fevereiro)
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
    Call GravarLogErro("IsDataValida", Err.Number, Err.Description)
    IsDataValida = False
End Function

' -----------------------------------------------------------
' SafeValue: Captura valores de objetos ou variáveis sem gerar erro de Nulo/Empty
' -----------------------------------------------------------
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

' -----------------------------------------------------------
' TrimTodosCampos: Aplica Trim em lote em múltiplos controles/strings
' -----------------------------------------------------------
Public Sub TrimTodosCampos(ParamArray Campos() As Variant)
    On Error Resume Next
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
' SEÇÃO 4: PERSISTÊNCIA DE LOGS E AUDITORIA (Blindagem SHA e SQL)
' ==============================================================================
Public Sub GravarLogErro(Optional ByVal Modulo As String = "", Optional ByVal NumeroErro As Long = 0, Optional ByVal DescricaoErro As String = "")
    
    On Error Resume Next
    Dim SQL As String
    Dim Usuario As String, Maquina As String
    
    ' Captura e sanitiza metadados
    Usuario = Left(Replace(ObterUsuarioAtual(), "'", "''"), 150)
    Maquina = Left(Replace(ObterMaquinaAtual(), "'", "''"), 100)
    
    ' Limpa parâmetros
    DescricaoErro = Replace(DescricaoErro, "'", "''")
    Modulo = Left(Replace(Modulo, "'", "''"), 100)
    
    ' SQL sem o campo LinhaErro
    SQL = "INSERT INTO Tbl_LogErro (DataHora, Usuario, NomeMaquina, Modulo, NumeroErro, DescricaoErro) " & _
          "VALUES (#" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "#, '" & _
          Usuario & "', '" & Maquina & "', '" & Modulo & "', " & NumeroErro & ", '" & DescricaoErro & "')"
    
    Mdl_Conexao.ExecutarSQL SQL
End Sub

Public Sub RegistrarAuditoria(ByVal TipoOperacao As String, ByVal Tabela As String, ByVal RegistroID As Long, ByVal Descricao As String)
    On Error Resume Next
    Dim SQL As String
    Dim Usuario As String, Maquina As String
    
    Usuario = Left(ObterUsuarioAtual(), 150)
    Maquina = Left(ObterMaquinaAtual(), 100)
    
    TipoOperacao = Left(Replace(TipoOperacao, "'", "''"), 100)
    Tabela = Left(Replace(Tabela, "'", "''"), 150)
    Descricao = Replace(Descricao, "'", "''")
    
    SQL = "INSERT INTO Tbl_Auditoria (DataHora, TipoOperacao, Tabela, RegistroID, Descricao, Usuario, Maquina) " & _
          "VALUES (#" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "#, '" & _
          TipoOperacao & "', '" & Tabela & "', " & RegistroID & ", '" & Descricao & "', '" & Usuario & "', '" & Maquina & "')"
          
    Mdl_Conexao.ExecutarSQL SQL
End Sub

Public Sub RegistrarLogAcesso(ByVal UsuarioTentativa As String, ByVal Status As String)
    On Error Resume Next
    Dim SQL As String
    Dim Maquina As String
    
    Maquina = Left(ObterMaquinaAtual(), 100)
    UsuarioTentativa = Left(Replace(UsuarioTentativa, "'", "''"), 100)
    Status = Left(Replace(Status, "'", "''"), 50)
    
    SQL = "INSERT INTO Tbl_LogAcesso (DataHora, Usuario, NomeMaquina, Status) " & _
          "VALUES (#" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "#, '" & _
          UsuarioTentativa & "', '" & Maquina & "', '" & Status & "')"
          
    Mdl_Conexao.ExecutarSQL SQL
End Sub

' ==============================================================================
' SEÇÃO 5: METADADOS E AMBIENTE (Hardware e Sessão)
' ==============================================================================

Public Function ObterUsuarioAtual() As String
    If Mdl_VariaveisGlobais.UsuarioNome <> "" Then
        ObterUsuarioAtual = Mdl_VariaveisGlobais.UsuarioNome
    Else
        ObterUsuarioAtual = Environ("Username")
    End If
End Function

Public Function ObterMaquinaAtual() As String
    ObterMaquinaAtual = Environ("Computername")
End Function

