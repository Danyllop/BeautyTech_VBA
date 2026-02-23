VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBotaoEfeito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents GrupoLabel As MSForms.Label
Attribute GrupoLabel.VB_VarHelpID = -1
Public Tipo                  As String ' "BTN", "LBL" ou "ICON"

Private Sub GrupoLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ' Se for botão, aumenta a fonte para 14
    If Tipo = "BTN" Then
        If GrupoLabel.Font.Size <> 14 Then
            Mdl_UI_Efeitos.ResetarDestaque
            With GrupoLabel
                .Font.Size = 14
                .Font.Name = "Segoe UI Semibold"
            End With
        End If
        
    ' Se for ícone, aumenta o tamanho e muda a cor, mas MANTÉM a fonte de ativos
    ElseIf Tipo = "ICO" Then
        If GrupoLabel.ForeColor <> RGB(255, 255, 255) Then
            Mdl_UI_Efeitos.ResetarDestaque
            With GrupoLabel
                .ForeColor = RGB(255, 255, 255) ' Branco
                .Font.Size = 20                 ' Efeito de leve aumento
                .Font.Name = "Segoe MDL2 Assets"
            End With
        End If
        
    ElseIf Tipo = "LBL" Then
        If GrupoLabel.Font.Name <> "Segoe UI Semibold" Then
            Mdl_UI_Efeitos.ResetarDestaque
            GrupoLabel.Font.Name = "Segoe UI Semibold"
        End If
    End If
End Sub
