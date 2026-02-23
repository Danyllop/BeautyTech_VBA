VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIconeHover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==============================================================================
' CLASSE: clsIconeHover (Versão Flexível)
' ==============================================================================
Option Explicit

Private WithEvents mLabel As MSForms.Label
Attribute mLabel.VB_VarHelpID = -1

' Cores
Private Const COR_NORMAL As Long = 11513740 ' Cinza Azulado (Padrão do seu tema)
Private mCorHover As Long ' Variável para guardar a cor de destaque escolhida

' ------------------------------------------------------------------------------
' INICIALIZAÇÃO (Agora aceita a cor desejada!)
' ------------------------------------------------------------------------------
' Se não informar a cor, ele usa Branco (vbWhite) como padrão
Public Sub Inicializar(ByRef LabelControle As MSForms.Label, Optional ByVal CorDestaque As Long = vbWhite)
    Set mLabel = LabelControle
    mCorHover = CorDestaque
End Sub

' ------------------------------------------------------------------------------
' APLICA O DESTAQUE
' ------------------------------------------------------------------------------
Public Sub Destacar()
    ' Agora usa a variável definida na inicialização, não mais uma constante
    mLabel.ForeColor = mCorHover
End Sub

' ------------------------------------------------------------------------------
' RESETA
' ------------------------------------------------------------------------------
Public Sub Resetar()
    On Error Resume Next
    mLabel.ForeColor = COR_NORMAL
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' EVENTOS
' ------------------------------------------------------------------------------
Private Sub mLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Mantém sua lógica de gerenciador central
    Mdl_UI_Efeitos.DefinirFocoIcone Me
End Sub

Private Sub Class_Terminate()
    Set mLabel = Nothing
End Sub
