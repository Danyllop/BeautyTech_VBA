VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLabelComCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================
' ARQUIVO: Classe clsLabelComCursor
' OBJETIVO: Gerenciar o evento de mouse e API do cursor Mãozinha
' AUTOR: Danyllo (Otimizado para Multi-Formulários)
' ==============================================================================
Option Explicit

' --- Declarações da API do Windows (Compatível com 32 e 64 bits) ---
#If VBA7 Then
    Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As Long) As LongPtr
    Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
    Private hCursorMao As LongPtr
#Else
    Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
    Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
    Private hCursorMao As Long
#End If

' Constante do Windows para a Mãozinha (Hand)
Private Const IDC_HAND As Long = 32649

' Variável de Evento (A "alma" da classe)
Private WithEvents mLabel As MSForms.Label
Attribute mLabel.VB_VarHelpID = -1

' ------------------------------------------------------------------------------
' INICIALIZAÇÃO DA CLASSE
' Carrega o cursor na memória apenas uma vez quando a classe nasce.
' Isso economiza processamento durante o movimento do mouse.
' ------------------------------------------------------------------------------
Private Sub Class_Initialize()
    hCursorMao = LoadCursor(0, IDC_HAND)
End Sub

' ------------------------------------------------------------------------------
' MÉTODO PÚBLICO: CONFIGURAR
' Recebe o controle do formulário e conecta os eventos.
' ------------------------------------------------------------------------------
Public Sub Inicializar(ByRef LabelControle As MSForms.Label)
    Set mLabel = LabelControle
    
    ' Define o ponteiro como Customizado (necessário para a API funcionar bem)
    mLabel.MousePointer = fmMousePointerCustom
    
    ' DICA: Se no futuro quiser adicionar ToolTip ou Cor, faça aqui.
End Sub

' ------------------------------------------------------------------------------
' EVENTO: MOUSE MOVE
' Dispara quando o mouse passa sobre o rótulo.
' ------------------------------------------------------------------------------
Private Sub mLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Aplica o cursor carregado na memória
    If hCursorMao <> 0 Then
        SetCursor hCursorMao
    End If
End Sub

' ------------------------------------------------------------------------------
' LIMPEZA DE MEMÓRIA (Destructor)
' Garante que a referência ao Label seja solta quando a classe morrer.
' ------------------------------------------------------------------------------
Private Sub Class_Terminate()
    Set mLabel = Nothing
End Sub
