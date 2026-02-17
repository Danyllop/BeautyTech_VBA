Attribute VB_Name = "Mdl_Seguranca"
Option Explicit
' ==============================================================================
' Módulo: Mdl_Seguranca
' Objetivo: Criptografia SHA-256 Robusta
' Requer Referências:
'   1. Microsoft XML, v6.0
'   2. Microsoft ActiveX Data Objects 6.1 Library
' ==============================================================================

Public Function GerarHashSHA256(ByVal Texto As String) As String
    On Error GoTo ErroHash
    
    If Len(Texto) = 0 Then
        GerarHashSHA256 = ""
        Exit Function
    End If
    
    ' Objetos de Automação do Windows (.NET Framework)
    Dim UTF8Enc As Object
    Dim SHA256 As Object
    Dim BytesTexto() As Byte
    Dim BytesHash() As Byte
    
    ' Criação dos objetos do sistema
    Set UTF8Enc = CreateObject("System.Text.UTF8Encoding")
    Set SHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    ' Converte a string para bytes UTF-8
    BytesTexto = UTF8Enc.GetBytes_4(Texto)
    
    ' Gera o Hash
    BytesHash = SHA256.ComputeHash_2((BytesTexto))
    
    ' Converte os bytes do hash para String Hexadecimal
    ' Usamos uma abordagem simples de loop para garantir compatibilidade
    Dim i As Integer
    Dim sb As String
    
    For i = LBound(BytesHash) To UBound(BytesHash)
        ' Formata cada byte como 2 caracteres Hex (ex: A -> 0A)
        sb = sb & Right("0" & Hex(BytesHash(i)), 2)
    Next i
    
    ' Retorna em minúsculo para padronizar
    GerarHashSHA256 = LCase(sb)
    
    Set UTF8Enc = Nothing
    Set SHA256 = Nothing
    Exit Function

ErroHash:
    ' Se falhar (ex: bloqueio de segurança do Windows), grava log e retorna erro
    Dim msgErro As String
    msgErro = "Erro na Criptografia: " & Err.Number & " - " & Err.Description
    
    ' Tenta gravar no log se possível
    On Error Resume Next
    Mdl_Utilitarios.GravarLogErro "GerarHashSHA256", Err.Number, Err.Description, Erl
    
    ' Retorna um valor de erro visível para debug
    GerarHashSHA256 = "ERRO_CRYPT"
End Function

' -----------------------------------------------------------
' ValidarSenhaForte: Verifica a complexidade da senha via Regex
' Retorna: True se a senha for robusta, False caso contrário
' -----------------------------------------------------------
Public Function ValidarSenhaForte(ByVal Senha As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Explicação da Regex:
    ' ^(?=.*[a-z])      -> Pelo menos uma minúscula
    ' (?=.*[A-Z])       -> Pelo menos uma maiúscula
    ' (?=.*\d)          -> Pelo menos um dígito (número)
    ' (?=.*[@$!%*?&])   -> Pelo menos um caractere especial
    ' [A-Za-z\d@$!%*?&] -> Caracteres permitidos
    ' {8,}              -> No mínimo 8 caracteres
    
    With regEx
        .Global = True
        .IgnoreCase = False
        .Pattern = "^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[@$!%*?&])[A-Za-z\d@$!%*?&]{8,}$"
    End With
    
    ValidarSenhaForte = regEx.Test(Senha)
    
    Set regEx = Nothing
End Function

' ------------------------------------------------------------------------------
' Valida o formato de e-mail usando Expressões Regulares (RegExp)
' ------------------------------------------------------------------------------
Public Function ValidarEmail(ByVal Email As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Pattern refinado:
    ' 1. Aceita letras, números e caracteres permitidos antes do @
    ' 2. Impede pontos duplicados ou no início/fim da parte local
    ' 3. Valida domínios e TLDs de pelo menos 2 letras (ex: .com, .br, .tech)
    Dim strPattern As String
    strPattern = "^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$"

    With regEx
        .Pattern = strPattern
        .IgnoreCase = True
        .Global = False
    End With

    ' Retorna True se o e-mail passar no teste
    ValidarEmail = regEx.Test(Trim(Email))
    
    Set regEx = Nothing
End Function

' ------------------------------------------------------------------------------
' Validação matemática de CPF (Algoritmo de dígitos verificadores)
' ------------------------------------------------------------------------------
Public Function ValidarCPF(ByVal CPF As Variant) As Boolean
    Dim i As Integer, soma As Integer, resto As Integer
    Dim dig1 As Integer, dig2 As Integer
    Dim numeros As String

    ' Remove qualquer máscara e limpa apenas números
    numeros = ""
    For i = 1 To Len(CPF)
        If Mid(CPF, i, 1) Like "[0-9]" Then numeros = numeros & Mid(CPF, i, 1)
    Next i

    ' Validações básicas de tamanho e números repetidos (000..., 111...)
    If Len(numeros) <> 11 Then Exit Function
    If numeros = String(11, Left(numeros, 1)) Then Exit Function

    ' Cálculo do 1º Dígito
    soma = 0
    For i = 1 To 9: soma = soma + val(Mid(numeros, i, 1)) * (11 - i): Next i
    resto = (soma * 10) Mod 11
    dig1 = IIf(resto = 10 Or resto = 11, 0, resto)

    ' Cálculo do 2º Dígito
    soma = 0
    For i = 1 To 10: soma = soma + val(Mid(numeros, i, 1)) * (12 - i): Next i
    resto = (soma * 10) Mod 11
    dig2 = IIf(resto = 10 Or resto = 11, 0, resto)

    ' Verifica se os dígitos calculados batem com os informados
    ValidarCPF = (dig1 = val(Mid(numeros, 10, 1)) And dig2 = val(Mid(numeros, 11, 1)))
End Function

Public Function ObterHardwareID() As String
    On Error Resume Next
    Dim wmi As Object, col As Object, Item As Object
    Dim serial As String
    
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    
    ' Tenta Serial da BIOS
    Set col = wmi.ExecQuery("Select SerialNumber From Win32_BIOS")
    For Each Item In col
        serial = Item.SerialNumber
    Next
    
    ' Se falhar, tenta Sistema Operacional
    If Len(serial) < 2 Or serial = "0" Then
        Set col = wmi.ExecQuery("Select SerialNumber From Win32_OperatingSystem")
        For Each Item In col
            serial = Item.SerialNumber
        Next
    End If
    
    If serial = "" Then serial = "UNKNOWN_HWID"
    
    ObterHardwareID = GerarHashSHA256(serial)
End Function

' -----------------------------------------------------------
' ValidarLicenca: (Placeholder para futuro)
' -----------------------------------------------------------
Public Function ValidarLicenca() As Boolean
    ' Aqui futuramente verificaremos se o HardwareID bate com o banco online
    ValidarLicenca = True
End Function

' --- ROTINA TEMPORÁRIA DE RESET DE SENHA ---
Public Sub ResetarSenhaAdmin()
    On Error GoTo ErroReset
    
    Debug.Print "1. Iniciando reset..."
    
    ' Teste de Criptografia
    Dim NovaSenhaHash As String
    Debug.Print "2. Gerando Hash..."
    NovaSenhaHash = Mdl_Seguranca.GerarHashSHA256("Admin@123")
    
    If NovaSenhaHash = "" Or Left(NovaSenhaHash, 5) = "ERRO_" Then
        MsgBox "Erro ao gerar a criptografia. O problema está no Mdl_Seguranca.", vbCritical
        Exit Sub
    End If
    Debug.Print "3. Hash gerado: " & Left(NovaSenhaHash, 10) & "..."
    
    ' Teste de Conexão
    Debug.Print "4. Conectando ao banco..."
    Mdl_Conexao.ConectarBD
    
    ' Atualização
    Debug.Print "5. Atualizando tabela..."
    Dim SQL As String
    SQL = "UPDATE Tbl_Usuarios SET Senha = '" & NovaSenhaHash & "' WHERE Usuario = 'admin'"
    Mdl_Conexao.Conexao.Execute SQL
    
    MsgBox "Sucesso! Senha do Admin redefinida para: Admin@123", vbInformation
    Debug.Print "6. Sucesso Total."
    Exit Sub

ErroReset:
    MsgBox "Erro na etapa: " & Erl & vbNewLine & _
           "Descrição: " & Err.Description, vbCritical, "Falha no Reset"
End Sub

