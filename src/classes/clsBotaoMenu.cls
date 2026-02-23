VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBotaoMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================
' CLASSE: clsBotaoMenu
' OBJETIVO: Gerenciar Menu Lateral (Normal, Hover, Ativo) sem ícones
' ==============================================================================
Option Explicit

' Evento para capturar o mouse e o clique
Private WithEvents mLabel As MSForms.Label
Attribute mLabel.VB_VarHelpID = -1

' Variáveis de Armazenamento (Imagens e Referências)
Private mPicNormal As StdPicture ' Guarda a IMAGEM (Leve), não o controle
Private mPicHover  As StdPicture
Private mPicAtivo  As StdPicture
Private mFormPai   As Object     ' Para chamar a rotina de clique

' Estado do Botão
Private mIsActive  As Boolean

' Cores (Configuradas internamente para padronização)
Private Const COR_NORMAL As Long = 11513740 ' RGB(140, 155, 175) - Cinza Azulado
Private Const COR_BRANCO As Long = 16777215 ' RGB(255, 255, 255)

' ------------------------------------------------------------------------------
' MÉTODOS DE INICIALIZAÇÃO
' ------------------------------------------------------------------------------
Public Sub Inicializar(ByRef LabelAlvo As MSForms.Label, _
                       ByRef FrmPai As Object, _
                       ByRef ImgNorm As MSForms.Image, _
                       ByRef ImgHov As MSForms.Image, _
                       ByRef ImgAtv As MSForms.Image)
                       
    Set mLabel = LabelAlvo
    Set mFormPai = FrmPai
    
    ' Carrega as imagens na memória da classe
    Set mPicNormal = ImgNorm.Picture
    Set mPicHover = ImgHov.Picture
    Set mPicAtivo = ImgAtv.Picture
    
    ' Define estado inicial
    mIsActive = False
    Me.Renderizar ' Aplica o visual inicial
End Sub

' ------------------------------------------------------------------------------
' PROPRIEDADE: ATIVO
' Define se este é o botão da página atual
' ------------------------------------------------------------------------------
Public Property Let Ativo(ByVal Valor As Boolean)
    mIsActive = Valor
    Renderizar ' Redesenha o botão imediatamente ao mudar o estado
End Property

Public Property Get Ativo() As Boolean
    Ativo = mIsActive
End Property

' ------------------------------------------------------------------------------
' PROPRIEDADE LEITURA: Nome do Botão (Útil para o Módulo saber quem é quem)
' ------------------------------------------------------------------------------
Public Property Get Nome() As String
    Nome = mLabel.Name
End Property

' ------------------------------------------------------------------------------
' RENDERIZAÇÃO (O "Cérebro" Visual)
' Aplica a imagem e cor corretas baseadas no estado ATIVO
' ------------------------------------------------------------------------------
Public Sub Renderizar()
    If mIsActive Then
        ' ESTADO ATIVO (Selecionado)
        mLabel.Picture = mPicAtivo
        mLabel.ForeColor = COR_BRANCO
    Else
        ' ESTADO NORMAL (Inativo)
        mLabel.Picture = mPicNormal
        mLabel.ForeColor = COR_NORMAL
    End If
End Sub

' ------------------------------------------------------------------------------
' EVENTO: MOUSE MOVE (Efeito Hover Temporário)
' ------------------------------------------------------------------------------
Private Sub mLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Se já estiver ativo, não precisa fazer efeito hover
    If mIsActive Then Exit Sub
    
    ' Se já estiver com a imagem hover (verificação visual para não piscar), sai
    If mLabel.ForeColor = COR_BRANCO Then Exit Sub
    
    ' Aplica visual de Hover
    mLabel.Picture = mPicHover
    mLabel.ForeColor = COR_BRANCO
    
    ' DICA: Aqui você pode chamar um ResetarHoverMenu no form se o mouse sair muito rápido,
    ' mas geralmente o controle de movimento do frame resolve.
End Sub

' ------------------------------------------------------------------------------
' EVENTO: CLIQUE
' ------------------------------------------------------------------------------
Private Sub mLabel_Click()
    ' 1. Dispara a ação no formulário pai
    If Not mFormPai Is Nothing Then
        On Error Resume Next
        ' Exemplo: Se o botão chama BtnDashboard, procura a Sub "Click_BtnDashboard"
        CallByName mFormPai, "Click_" & mLabel.Name, VbMethod
        On Error GoTo 0
    End If
    
    ' 2. Avisa o módulo para atualizar quem está ativo (Visual)
    ' Isso será feito pelo próprio formulário chamando o módulo, veja abaixo.
End Sub

' Limpeza
Private Sub Class_Terminate()
    Set mLabel = Nothing
    Set mFormPai = Nothing
    Set mPicNormal = Nothing
    Set mPicHover = Nothing
    Set mPicAtivo = Nothing
End Sub

