VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSimpleButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================
' CLASSE: clsSimpleButton
' OBJETIVO: Controlar o visual (Negrito/Tamanho) de Labels e Botões de Texto
' ==============================================================================
Option Explicit

' Evento para capturar o mouse
Private WithEvents mLabel As MSForms.Label
Attribute mLabel.VB_VarHelpID = -1

' Propriedades internas
Private mTipo As String ' Guarda se é "BTN" ou "LBL"

' ------------------------------------------------------------------------------
' Inicializa o controle e define seu tipo
' ------------------------------------------------------------------------------
Public Sub Inicializar(ByRef LabelControle As MSForms.Label, ByVal sTipo As String)
    Set mLabel = LabelControle
    mTipo = sTipo
End Sub

' ------------------------------------------------------------------------------
' APLICA O DESTAQUE (Chamado pelo Módulo)
' ------------------------------------------------------------------------------
Public Sub Destacar()
    With mLabel
        If mTipo = "BTN" Then
            ' Botão: Fonte Semibold (Negrito leve) e Tamanho 13
            .Font.Name = "Segoe UI Semibold"
            .Font.Size = 13
        ElseIf mTipo = "LBL" Then
            ' Label: Apenas Fonte Semibold
            .Font.Name = "Segoe UI Semibold"
        End If
    End With
End Sub

' ------------------------------------------------------------------------------
' REMOVE O DESTAQUE (Chamado pelo Módulo)
' ------------------------------------------------------------------------------
Public Sub Resetar()
    ' On Error Resume Next previne falhas se o form estiver fechando
    On Error Resume Next
    With mLabel
        If mTipo = "BTN" Then
            ' Botão: Volta para fonte normal
            .Font.Name = "Segoe UI"
            .Font.Size = 12
        ElseIf mTipo = "LBL" Then
            ' Label: Volta para fonte normal
            .Font.Name = "Segoe UI"
        End If
    End With
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' EVENTO: MOUSE MOVE
' Ao passar o mouse, avisa o módulo que ESTE botão é o novo foco.
' ------------------------------------------------------------------------------
Private Sub mLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Envia a própria instância (Me) para o gerenciador de foco
    Mdl_UI_Efeitos.DefinirFocoVisual Me
End Sub

' Limpeza de memória ao destruir a classe
Private Sub Class_Terminate()
    Set mLabel = Nothing
End Sub

