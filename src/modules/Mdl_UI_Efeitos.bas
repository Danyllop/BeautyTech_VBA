Attribute VB_Name = "Mdl_UI_Efeitos"
' ==============================================================================
' NOME DO MÓDULO: Mdl_UI_Efeitos
' OBJETIVO:       Fábrica Central de Efeitos (Cursor, Menu, Hover)
' ARQUITETURA:    Factory Pattern (O Form segura a memória, o Módulo apenas cria)
' ==============================================================================
Option Explicit

' ==============================================================================
' 1. APIS DO WINDOWS (Barra de Título)
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
' 2. VARIÁVEIS DE CONTROLE INTERNO (Singleton de Foco)
' ==============================================================================
' Rastreia APENAS o último botão simples focado para evitar loops de reset.
Private UltimoFoco  As clsSimpleButton
Private UltimoIcone As clsIconeHover

' ==============================================================================
' 3. FÁBRICAS DE EFEITOS (Retornam Collections para o Formulário)
' ==============================================================================

' --- A. CURSOR MÃOZINHA (BTN, LBL, ICO) ---
Public Function CriarEfeitosMaozinha(ByVal Frm As Object) As Collection
    Dim colRetorno As New Collection
    Dim Ctrl As Control
    Dim Obj As clsLabelComCursor
    Dim sPrefixo As String
    
    For Each Ctrl In Frm.Controls
        If TypeName(Ctrl) = "Label" Then
            If Len(Ctrl.Name) >= 3 Then
                sPrefixo = UCase(Left(Ctrl.Name, 3))
                ' Aceita BTN, LBL e ICO (Ícones)
                If sPrefixo = "BTN" Or sPrefixo = "LBL" Or sPrefixo = "ICO" Then
                    Set Obj = New clsLabelComCursor
                    Obj.Inicializar Ctrl
                    colRetorno.Add Obj
                End If
            End If
        End If
    Next Ctrl
    Set CriarEfeitosMaozinha = colRetorno
End Function

' --- B. BOTÕES SIMPLES - VISUAL (BTN, LBL) ---
Public Function CriarSimpleButton(ByVal Frm As Object) As Collection
    Dim colRetorno As New Collection
    Dim Ctrl As Control
    Dim Obj As clsSimpleButton
    Dim sPrefixo As String
    
    For Each Ctrl In Frm.Controls
        If TypeName(Ctrl) = "Label" Then
            If Len(Ctrl.Name) >= 3 Then
                sPrefixo = UCase(Left(Ctrl.Name, 3))
                ' Aceita apenas BTN e LBL (Texto)
                If sPrefixo = "BTN" Or sPrefixo = "LBL" Then
                    Set Obj = New clsSimpleButton
                    Obj.Inicializar Ctrl, sPrefixo
                    colRetorno.Add Obj
                End If
            End If
        End If
    Next Ctrl
    Set CriarSimpleButton = colRetorno
End Function

' --- C. MENU LATERAL (BTN com Imagens) ---
Public Function CriarMenuLateral(ByVal Frm As Object, ByRef ImgNorm As MSForms.Image, ByRef ImgHov As MSForms.Image, ByRef ImgAtv As MSForms.Image) As Collection
                                 
    Dim colRetorno As New Collection
    Dim Ctrl As Control
    Dim Obj As clsBotaoMenu
    
    For Each Ctrl In Frm.Controls
        If TypeName(Ctrl) = "Label" Then
            ' Filtra apenas BTN para o menu principal
            If UCase(Left(Ctrl.Name, 3)) = "BTN" Then
                Set Obj = New clsBotaoMenu
                ' Passa as imagens de referência uma única vez
                Obj.Inicializar Ctrl, Frm, ImgNorm, ImgHov, ImgAtv
                ' Adiciona com CHAVE (Nome do controle) para busca rápida
                colRetorno.Add Obj, Ctrl.Name
            End If
        End If
    Next Ctrl
    Set CriarMenuLateral = colRetorno
End Function

' ==============================================================================
' 4. GERENCIADORES DE ESTADO (Lógica de Negócio Visual)
' ==============================================================================

' --- GERENCIADOR DE FOCO (Botão Simples) ---
' O(1) Performance - Sem Loops
Public Sub DefinirFocoVisual(ByRef NovoObjeto As clsSimpleButton)
    If NovoObjeto Is UltimoFoco Then Exit Sub
    
    ' Reseta o anterior se existir
    If Not UltimoFoco Is Nothing Then UltimoFoco.Resetar
    
    ' Define e destaca o novo
    Set UltimoFoco = NovoObjeto
    UltimoFoco.Destacar
End Sub

' Limpeza forçada (usar no Terminate do Form)
Public Sub LimparFoco()
    If Not UltimoFoco Is Nothing Then
        UltimoFoco.Resetar
        Set UltimoFoco = Nothing
    End If
End Sub

' --- GERENCIADOR DE MENU (Troca de Página) ---
Public Sub SelecionarBotao(ByVal ColMenu As Collection, ByVal NomeBotaoAtivo As String)
    Dim Obj As clsBotaoMenu
    ' Varre o menu para garantir exclusividade (um aceso, resto apagado)
    For Each Obj In ColMenu
        If Obj.Nome = NomeBotaoAtivo Then
            Obj.Ativo = True
        Else
            Obj.Ativo = False
        End If
    Next Obj
End Sub

' Reset visual do Menu (usar no MouseMove do Frame/Fundo)
Public Sub ResetarHoverMenu(ByVal ColMenu As Collection)
    If ColMenu Is Nothing Then Exit Sub
    Dim Obj As clsBotaoMenu
    For Each Obj In ColMenu: Obj.Renderizar: Next Obj
End Sub

' ==============================================================================
' 5. UTILITÁRIOS EXTRAS (Barra de Título e Máscaras)
' ==============================================================================
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

' (Mantive Máscaras e Maiúsculas pois são úteis, mas tirei as globais)
Public Function AtivarMaiusculas(ByVal Frm As Object) As Collection
    Dim col As New Collection, Ctrl As Control, Obj As clsTxtMaiuscula
    For Each Ctrl In Frm.Controls
        If TypeName(Ctrl) = "TextBox" And UCase(Left(Ctrl.Name, 3)) = "TXT" Then
            Set Obj = New clsTxtMaiuscula: Set Obj.txtGroup = Ctrl
            col.Add Obj
        End If
    Next
    Set AtivarMaiusculas = col
End Function

Public Function AtivarMascaras(ByVal Frm As Object) As Collection
    Dim col As New Collection, Ctrl As Control, Obj As clsMascara
    For Each Ctrl In Frm.Controls
        If TypeName(Ctrl) = "TextBox" And Ctrl.Tag <> "" Then
            Set Obj = New clsMascara: Set Obj.CampoTexto = Ctrl
            col.Add Obj
        End If
    Next
    Set AtivarMascaras = col
End Function

' --- FACTORY: Cria a coleção de efeitos para ícones (Prefixo ICO) ---
Public Function CriarEfeitosIcone(ByVal Frm As Object) As Collection
    Dim colRetorno As New Collection
    Dim Ctrl As Control
    Dim Obj As clsIconeHover
    
    For Each Ctrl In Frm.Controls
        If TypeName(Ctrl) = "Label" Then
            ' Verifica prefixo ICO
            If UCase(Left(Ctrl.Name, 3)) = "ICO" Then
                Set Obj = New clsIconeHover
                Obj.Inicializar Ctrl
                colRetorno.Add Obj
            End If
        End If
    Next Ctrl
    
    Set CriarEfeitosIcone = colRetorno
End Function

' --- GERENCIADOR DE FOCO (ÍCONES) ---
Public Sub DefinirFocoIcone(ByRef NovoObjeto As clsIconeHover)
    ' Se o mouse ainda está no mesmo ícone, não faz nada
    If NovoObjeto Is UltimoIcone Then Exit Sub
    
    ' 1. Reseta o anterior (Volta para Cinza)
    If Not UltimoIcone Is Nothing Then
        UltimoIcone.Resetar
    End If
    
    ' 2. Define o novo
    Set UltimoIcone = NovoObjeto
    
    ' 3. Destaca o novo (Vira Branco)
    UltimoIcone.Destacar
End Sub

' --- LIMPEZA (Usar no MouseMove do Frame/Fundo) ---
Public Sub LimparFocoIcone()
    If Not UltimoIcone Is Nothing Then
        UltimoIcone.Resetar
        Set UltimoIcone = Nothing
    End If
End Sub

' --- LIMPEZA GERAL (Atualize seu LimparFoco antigo ou use este no Terminate) ---
Public Sub LimparTudosFocos()
    ' Limpa Botões de Texto
    If Not UltimoFoco Is Nothing Then
        UltimoFoco.Resetar
        Set UltimoFoco = Nothing
    End If
    
    ' Limpa Ícones
    If Not UltimoIcone Is Nothing Then
        UltimoIcone.Resetar
        Set UltimoIcone = Nothing
    End If
End Sub









