Attribute VB_Name = "Mdl_VariaveisGlobais"
Option Explicit
' ==============================================================================
' Módulo: Mdl_VariaveisGlobais
' Objetivo: Armazena os dados do usuário logado.
' ==============================================================================

' Variáveis Privadas (Internas)
Private pUsuarioID      As Long
Private pUsuarioNome    As String
Private pUsuarioNivel   As String
Private pUsuarioLogado  As Boolean

' --- Propriedades Públicas (Acesso Seguro) ---

' 1. ID do Usuário
Public Property Let UsuarioID(ByVal Valor As Long)
    pUsuarioID = Valor
End Property
Public Property Get UsuarioID() As Long
    UsuarioID = pUsuarioID
End Property

' 2. Nome do Usuário
Public Property Let UsuarioNome(ByVal Valor As String)
    pUsuarioNome = Valor
End Property
Public Property Get UsuarioNome() As String
    UsuarioNome = pUsuarioNome
End Property

' 3. Nível de Acesso
Public Property Let UsuarioNivel(ByVal Valor As String)
    pUsuarioNivel = Valor
End Property
Public Property Get UsuarioNivel() As String
    UsuarioNivel = pUsuarioNivel
End Property

' 4. Status de Login
Public Property Let UsuarioLogado(ByVal Valor As Boolean)
    pUsuarioLogado = Valor
End Property
Public Property Get UsuarioLogado() As Boolean
    UsuarioLogado = pUsuarioLogado
End Property

' --- Função de Limpeza ---
Public Sub LimparSessao()
    pUsuarioID = 0
    pUsuarioNome = ""
    pUsuarioNivel = ""
    pUsuarioLogado = False
End Sub
