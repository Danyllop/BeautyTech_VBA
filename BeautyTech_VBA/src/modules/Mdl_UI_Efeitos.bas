Attribute VB_Name = "Mdl_UI_Efeitos"
' ==============================================================================
' NOME DO MÓDULO: Mdl_UI_Efeitos
' OBJETIVO:       Centralizar e otimizar a interatividade visual (Hover, Cursor,
'                 Máscaras e Maiúsculas) do sistema BeautyTech.
' AUTOR:          Danyllo Pereira - LogicUp Solutions
' DATA:           Fevereiro/2026
' ==============================================================================
Option Explicit

' ==============================================================================
' SEÇÃO 1: APIs DO WINDOWS (Barra de Título)
' ==============================================================================
#If VBA7 Then
    Private Declare PtrSafe Function DwmSetWindowAttribute Lib "dwmapi.dll" (ByVal hWnd As LongPtr, ByVal dwAttribute As Long, ByRef pvAttribute As Any, ByVal cbAttribute As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Private Declare Function DwmSetWindowAttribute Lib "dwmapi.dll" (ByVal hwnd As Long, ByVal dwAttribute As Long, ByRef pvAttribute As Any, ByVal cbAttribute As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

Private Const DWMWA_CAPTION_COLOR As Long = 35
Private Const DWMWA_TEXT_COLOR    As Long = 36
Private Const DWMWA_BORDER_COLOR  As Long = 34

' ==============================================================================
' SEÇÃO 2: COLEÇÕES DE MEMÓRIA (Persistência de Objetos)
' ==============================================================================
Public colMascaras    As Collection
Public colMaiusculas  As Collection
Public colBotoesMao   As Collection
Public colBotoes      As Collection

' ==============================================================================
' SEÇÃO 3: PROCEDIMENTOS DE INICIALIZAÇÃO E OTIMIZAÇÃO
' ==============================================================================

' ------------------------------------------------------------------------------
' A. BARRA DE TÍTULO NATIVA
' ------------------------------------------------------------------------------
Public Sub PersonalizarBarraTitulo(ByVal Frm As Object, ByVal CorFundo As Long, ByVal CorTexto As Long)
    #If VBA7 Then
        Dim hWnd As LongPtr
    #Else
        Dim hWnd As Long
    #End If
    hWnd = FindWindow("ThunderDFrame", Frm.Caption)
    If hWnd <> 0 Then
        Call DwmSetWindowAttribute(hWnd, DWMWA_CAPTION_COLOR, CorFundo, 4)
        Call DwmSetWindowAttribute(hWnd, DWMWA_TEXT_COLOR, CorTexto, 4)
        Call DwmSetWindowAttribute(hWnd, DWMWA_BORDER_COLOR, CorFundo, 4)
    End If
End Sub

' ==============================================================================
' SEÇÃO: INICIALIZAÇÃO INDEPENDENTE DE INTERATIVIDADE
' ==============================================================================

' --- C. MÃOZINHA E HOVER (BTN/LBL) ---
Public Sub AtivarCursorEMao(ByVal Frm As Object)
    Dim ctrl As Control
    Dim obj As clsLabelComCursor
    Set colBotoesMao = New Collection

    For Each ctrl In Frm.Controls
        If TypeName(ctrl) = "Label" Then
            Dim prefix As String: prefix = UCase(Left(ctrl.Name, 3))
            If prefix = "BTN" Or prefix = "LBL" Then
                Set obj = New clsLabelComCursor
                obj.Inicializar ctrl
                colBotoesMao.Add obj
            End If
        End If
    Next
End Sub

' Rotina de auxílio para o Form resetar o efeito
Public Sub ResetarEfeitosCombo()
    If colBotoesMao Is Nothing Then Exit Sub
    Dim obj As clsLabelComCursor
    For Each obj In colBotoesMao: obj.Reset: Next
End Sub

' ------------------------------------------------------------------------------
' C. MAIÚSCULAS E MÁSCARAS (Opcionais por formulário)
' ------------------------------------------------------------------------------
Public Sub AtivarMaiusculas(ByVal Frm As Object)
    Dim ctrl As Control, obj As clsTxtMaiuscula
    Set colMaiusculas = New Collection
    For Each ctrl In Frm.Controls
        If TypeName(ctrl) = "TextBox" And UCase(Left(ctrl.Name, 3)) = "TXT" Then
            Set obj = New clsTxtMaiuscula: Set obj.txtGroup = ctrl
            colMaiusculas.Add obj
        End If
    Next
End Sub

Public Sub AtivarMascaras(ByVal Frm As Object)
    Dim ctrl As Control, obj As clsMascara
    Set colMascaras = New Collection
    For Each ctrl In Frm.Controls
        If TypeName(ctrl) = "TextBox" And ctrl.Tag <> "" Then
            Set obj = New clsMascara: Set obj.CampoTexto = ctrl
            colMascaras.Add obj
        End If
    Next
End Sub

' ------------------------------------------------------------------------------
' Propósito: Inicializa Hover e Cursor via prefixos (BTN, LBL, ICO)
' ------------------------------------------------------------------------------
Public Sub AplicarDestaque(ByRef Frm As Object)
    Dim ctrl       As Control
    Dim ClasseEf   As clsBotaoEfeito
    Dim ClasseMao  As clsLabelComCursor
    Dim Prefixo    As String
    
    Set colBotoes = New Collection
    Set colBotoesMao = New Collection
    
    For Each ctrl In Frm.Controls
        If TypeOf ctrl Is MSForms.Label Then
            Prefixo = UCase(Left(ctrl.Name, 3))
            
            ' Verifica se o prefixo é um dos três permitidos
            If Prefixo = "BTN" Or Prefixo = "LBL" Or Prefixo = "ICO" Then
                
                ' 1. EFEITO VISUAL
                Set ClasseEf = New clsBotaoEfeito
                Set ClasseEf.GrupoLabel = ctrl
                ClasseEf.Tipo = Prefixo
                colBotoes.Add ClasseEf
                
                ' 2. CURSOR MÃOZINHA
                Set ClasseMao = New clsLabelComCursor
                ClasseMao.Inicializar ctrl
                colBotoesMao.Add ClasseMao
                
            End If
        End If
    Next ctrl
End Sub

Public Sub ResetarDestaque()
    On Error Resume Next
    Dim obj As clsBotaoEfeito
    
    For Each obj In colBotoes
        With obj.GrupoLabel
            Select Case obj.Tipo
                Case "BTN"
                    .Font.Size = 12
                    .Font.Name = "Segoe UI"
                
                Case "ICO"
                    ' Reset específico para ícones: volta cor e tamanho original
                    .ForeColor = RGB(140, 155, 175) ' Cinza Azulado
                    .Font.Size = 20
                    .Font.Name = "Segoe MDL2 Assets"
                    
                Case "LBL"
                    .Font.Name = "Segoe UI"
            End Select
        End With
    Next obj
End Sub























