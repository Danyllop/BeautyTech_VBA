Attribute VB_Name = "Mdl_Conexao"
Option Explicit
' ==============================================================================
' Módulo: Mdl_Conexao
' Objetivo: Gerenciar conexão com Access, suportando OneDrive e Rede.
' Dependência: Microsoft ActiveX Data Objects X.X Library (Ferramentas > Referências)
' ==============================================================================
'
' Variável global para a conexão
Public Conexao As ADODB.Connection
'
' Constante com o nome do banco (ajuste se necessário)
Public Const NOME_BANCO As String = "BeautyTech_DB.accdb"

' ------------------------------------------------------------------------------
' Abre a conexão e CRIA o arquivo se ele não existir
' ------------------------------------------------------------------------------
Public Sub ConectarBD()
    On Error GoTo ErroConexao

    ' 1. Se o objeto de conexão não existe na memória, instanciamos agora
    If Conexao Is Nothing Then Set Conexao = New ADODB.Connection

    ' 2. Se já estiver conectado, não faz nada
    If Conexao.State = adStateOpen Then Exit Sub

    ' 3. Define o caminho real (Sua lógica de OneDrive/Local)
    Dim CaminhoCompleto As String
    CaminhoCompleto = ObterCaminhoLocal(ThisWorkbook.Path) & "\" & NOME_BANCO

    ' 4. AUTO-CURA: Se o arquivo físico não existe, criamos via ADOX (Late Binding)
    If Dir(CaminhoCompleto) = "" Then
        Dim Catalog As Object
        Set Catalog = CreateObject("ADOX.Catalog")
        ' Cria o arquivo .accdb vazio
        Catalog.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CaminhoCompleto
        Set Catalog = Nothing
        Debug.Print "Arquivo de banco criado em: " & CaminhoCompleto
    End If

    ' 5. Configura e abre a conexão ADODB
    With Conexao
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                           "Data Source=" & CaminhoCompleto & ";" & _
                           "Persist Security Info=False;"
        .Open
    End With

    Exit Sub

ErroConexao:
    MsgBox "Falha crítica ao conectar/criar o banco de dados." & vbNewLine & _
           "Erro: " & Err.Description, vbCritical, "Erro de Sistema"
End Sub

' ------------------------------------------------------------------------------
' Fecha a conexão (Importante para liberar o arquivo do Access)
' ------------------------------------------------------------------------------
'Public Sub DesconectarBD()
'    On Error Resume Next
'    If Not Conexao Is Nothing Then
'        If Conexao.State = adStateOpen Then Conexao.Close
'        Set Conexao = Nothing
'    End If
'End Sub

Public Sub DesconectarBD()
    On Error Resume Next
    If Not Conexao Is Nothing Then
        ' Se estiver aberta (1), fecha
        If Conexao.State = 1 Then Conexao.Close
        ' Remove o objeto da memória (Crucial para deletar o .laccdb)
        Set Conexao = Nothing
    End If
End Sub

' ------------------------------------------------------------------------------
' Executa comandos de ação (INSERT, UPDATE, DELETE)
' ------------------------------------------------------------------------------
Public Sub ExecutarSQL(ByVal SQL As String)
    On Error GoTo ErroSQL
    
    ' Garante que está conectado
    ConectarBD
    If Conexao.State <> adStateOpen Then Exit Sub
    
    Conexao.Execute SQL
    Exit Sub

ErroSQL:
    MsgBox "Erro ao executar comando SQL." & vbNewLine & _
           "SQL: " & SQL & vbNewLine & _
           "Erro: " & Err.Description, vbCritical
End Sub

' ------------------------------------------------------------------------------
' Retorna dados de consulta (SELECT)
' ------------------------------------------------------------------------------
Public Function ObterRecordset(ByVal SQL As String) As ADODB.Recordset
    On Error GoTo ErroConsulta
    
    ' Garante que está conectado
    ConectarBD
    If Conexao.State <> adStateOpen Then
        Set ObterRecordset = Nothing
        Exit Function
    End If
    
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset
    
    Rs.CursorLocation = adUseClient ' Permite contar registros e navegar livremente
    Rs.Open SQL, Conexao, adOpenStatic, adLockReadOnly
    
    Set ObterRecordset = Rs
    Exit Function

ErroConsulta:
    MsgBox "Erro na consulta de dados." & vbNewLine & _
           "SQL: " & SQL & vbNewLine & _
           "Erro: " & Err.Description, vbCritical
    Set ObterRecordset = Nothing
End Function

' ------------------------------------------------------------------------------
' Função Auxiliar: Corrige caminhos do OneDrive (HTTPS -> Local)
' ------------------------------------------------------------------------------
Private Function ObterCaminhoLocal(ByVal Path As String) As String
    Dim i As Integer
    Dim CaminhoTemp As String
    
    ' Se não for caminho web (https), retorna o próprio caminho
    If Left(Path, 4) <> "http" Then
        ObterCaminhoLocal = Path
        Exit Function
    End If
    
    ' Lógica para converter URL do OneDrive em caminho local
    Path = Replace(Path, "https://", "")
    Path = Replace(Path, "/", "\")
    
    ' Tenta encontrar o ponto de montagem do OneDrive
    ' Esta lógica varre a string procurando onde começa a pasta local
    For i = 1 To Len(Path)
        If Mid(Path, i, 1) = "\" Then
            ' Tenta reconstruir usando a variável de ambiente OneDrive
            CaminhoTemp = Environ("onedrive") & Mid(Path, i)
            If Dir(CaminhoTemp, vbDirectory) <> "" Then
                ObterCaminhoLocal = CaminhoTemp
                Exit Function
            End If
            
            ' Tenta OneDrive Commercial (Empresas)
            CaminhoTemp = Environ("onedrivecommercial") & Mid(Path, i)
            If Dir(CaminhoTemp, vbDirectory) <> "" Then
                ObterCaminhoLocal = CaminhoTemp
                Exit Function
            End If
        End If
    Next i
    
    ' Se falhar a conversão, retorna o original (pode gerar erro, mas é o fallback)
    ObterCaminhoLocal = Path
End Function

