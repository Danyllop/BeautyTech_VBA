Attribute VB_Name = "Mdl_Gestao_Usuarios"
' ==============================================================================
' NOME DO ARQUIVO: Mdl_Gestao_Usuarios
' PROJETO:         Sistema BeautyTech - Gestão Integrada
' DESCRIÇÃO:       Motor de Carga e Pesquisa em Memória (High Performance)
' ==============================================================================
Option Explicit

' Variável na Memória (O Segredo da Velocidade)
Private ArrayUsuariosTodos As Variant

' -------------------------------------------------------------------------
' 1. CARGA DE DADOS (Com Log de Erros)
' -------------------------------------------------------------------------
Public Sub CarregarDadosUsuarios(ByVal Frm As Object)
    Dim Rs As Object
    Dim SQL As String
    
    On Error GoTo ErroCarga
    
    Mdl_Conexao.ConectarBD
    ' Carrega os campos principais para a ListBox
    SQL = "SELECT ID, Nome, Usuario, Email, Nivel, Format(DataCadastro, 'dd/mm/yyyy') " & _
          "FROM Tbl_Usuarios ORDER BY Nome ASC"
          
    Set Rs = Mdl_Conexao.ObterRecordset(SQL)
    
    With Frm.ListUsuarios
        .Clear
        If Not Rs.EOF Then
            ArrayUsuariosTodos = Rs.GetRows ' Salva TUDO na memória (Array Bidimensional)
            .Column = ArrayUsuariosTodos    ' Joga para a tela instantaneamente
        Else
            ArrayUsuariosTodos = Empty
        End If
    End With
    
    Rs.Close
    Mdl_Conexao.DesconectarBD
    Exit Sub

ErroCarga:
    ' Registro técnico do erro para o desenvolvedor
    Mdl_Utilitarios.GravarLogErro "Mdl_Gestao_Usuarios.CarregarDadosUsuarios", Err.Number, Err.Description
    Mdl_Utilitarios.msgErro "Erro ao carregar a lista de usuários. O incidente foi registrado."
    Mdl_Conexao.DesconectarBD
End Sub

' -------------------------------------------------------------------------
' 2. MOTOR DE FILTRO INSTANTÂNEO (Com Auditoria Inteligente)
' -------------------------------------------------------------------------
Public Sub FiltrarUsuarios(ByVal Frm As Object, ByVal TextoBusca As String)
    If IsEmpty(ArrayUsuariosTodos) Then Exit Sub
    
    Dim TotalCols As Long, TotalRows As Long
    Dim i As Long, j As Long
    Dim Termo As String
    Dim Achou As Boolean
    
    On Error GoTo ErroFiltro
    
    TotalCols = UBound(ArrayUsuariosTodos, 1)
    TotalRows = UBound(ArrayUsuariosTodos, 2)
    
    ' Se a busca estiver vazia, restaura a lista completa
    If Trim(TextoBusca) = "" Then
        Frm.ListUsuarios.Column = ArrayUsuariosTodos
        Exit Sub
    End If
    
    ' --- CAMADA DE AUDITORIA ---
    ' Recomendação: Auditar apenas buscas significativas (ex: > 2 caracteres)
    ' para não sobrecarregar o banco no evento 'Change'
    If Len(TextoBusca) >= 3 Then
        Mdl_Utilitarios.RegistrarAuditoria "PESQUISA_USUARIO", "Tbl_Usuarios", 0, "Pesquisa realizada: " & TextoBusca
    End If
    
    ' Prepara Array temporário
    Dim ArrFiltrado() As Variant
    Dim LinhaFiltro As Long
    ReDim ArrFiltrado(TotalCols, TotalRows)
    
    Termo = UCase(Trim(TextoBusca))
    LinhaFiltro = 0
    
    ' Varredura de Alta Velocidade em Memória
    For i = 0 To TotalRows
        Achou = False
        ' Pesquisa em Nome, Usuário ou Email
        If InStr(1, UCase(ArrayUsuariosTodos(1, i)), Termo) > 0 Or _
           InStr(1, UCase(ArrayUsuariosTodos(2, i)), Termo) > 0 Or _
           InStr(1, UCase(ArrayUsuariosTodos(3, i)), Termo) > 0 Then
            Achou = True
        End If
        
        If Achou Then
            For j = 0 To TotalCols
                ArrFiltrado(j, LinhaFiltro) = ArrayUsuariosTodos(j, i)
            Next j
            LinhaFiltro = LinhaFiltro + 1
        End If
    Next i
    
    If LinhaFiltro > 0 Then
        ReDim Preserve ArrFiltrado(TotalCols, LinhaFiltro - 1)
        Frm.ListUsuarios.Column = ArrFiltrado
    Else
        Frm.ListUsuarios.Clear
    End If
    Exit Sub

ErroFiltro:
    Mdl_Utilitarios.GravarLogErro "Mdl_Gestao_Usuarios.FiltrarUsuarios", Err.Number, Err.Description
End Sub

