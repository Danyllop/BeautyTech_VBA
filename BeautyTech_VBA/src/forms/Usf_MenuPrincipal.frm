VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Usf_MenuPrincipal 
   Caption         =   "LogicUp Solutions - BeautyTech"
   ClientHeight    =   12135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19365
   OleObjectBlob   =   "Usf_MenuPrincipal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Usf_MenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    ' =======================================================================
    ' [MOCK DE DADOS - TEMPORÁRIO PARA TESTES SEM O BANCO ACCESS]
    ' Captura os textos de teste que você deixou nas Labels do UserForm
    ' =======================================================================
    Mdl_VariaveisGlobais.UsuarioID = 1
    Mdl_VariaveisGlobais.UsuarioNome = Me.LbpUsuarioLogado.Caption
    Mdl_VariaveisGlobais.UsuarioNivel = Me.LbpUsuarioNivel.Caption
    Mdl_VariaveisGlobais.UsuarioLogado = True
    ' =======================================================================
    
    ' 1. Configurações Físicas Base
    Me.Height = 540
    Me.Width = 980
    
    ' 2. Configurações Visuais (Módulo de UI)
    Mdl_UI_Designer.ConfigIcoSair Me
    Mdl_UI_Designer.ConfigIcoUsuarios Me
    Mdl_UI_Designer.ConfigInfoUsuario Me
    Mdl_UI_Designer.ConfigIcoModoDev Me
    
    Mdl_UI_Efeitos.AtivarMaiusculas Me
    Mdl_UI_Efeitos.AtivarMascaras Me
    Mdl_UI_Efeitos.AplicarDestaque Me
    Mdl_UI_Efeitos.PersonalizarBarraTitulo Me, RGB(33, 95, 152), RGB(255, 255, 255)
        
    ' 3. Redimensionamento Responsivo (90% da tela)
    Mdl_UI_Designer.AjustarTamanhoFormulario Me
    
    ' 4. Aplica Permissões do Sistema (Exibe/Oculta botão Dev com base no Mock)
    Mdl_Sistema.AplicarPermissoes Me
            
End Sub

Private Sub UserForm_Resize()
    If Me.InsideHeight > 100 Then
        ' 1. Redimensiona o Menu Lateral
        Mdl_UI_Designer.RedimensionarMenuLateral Me
        
        ' 2. Redimensiona a Barra Superior
        Mdl_UI_Designer.RedimensionarBarraSuperior Me
        
        ' 3. Redimensiona a MultiPage (Conteúdo)
        Mdl_UI_Designer.RedimensionarConteudoPrincipal Me
        
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' CloseMode 0 (vbFormControlMenu) significa que o usuário clicou no "X"
    If CloseMode = 0 Then
        ' 1. Bloqueia o fechamento
        Cancel = True
        
        ' 2. Exibe a mensagem de instrução
        MsgBox "Acesso Restrito: Por favor, utilize o botão 'Sair' ou o menu para encerrar o sistema com segurança.", _
               vbExclamation, "Segurança LogicUp Solutions"
    End If
    
    ' Se for 1 (Unload Me disparado pelo seu botão Sair), ele passa direto e fecha.
End Sub

Private Sub FrmMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Mdl_UI_Efeitos.ResetarDestaque
End Sub

Private Sub FraRodapeMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Mdl_UI_Efeitos.ResetarDestaque
End Sub

Private Sub ICOSair_Click()
    If MsgBox("Deseja realmente encerrar o sistema e sair?", vbQuestion + vbYesNo, "Sair") = vbYes Then
        ' O Unload Me ativa o CloseMode = 1 e permite o formulário sumir
        Unload Me
        EncerrarSistemaBeautyTech
    End If
End Sub

Private Sub IcoModoDev_Click()
    ' Repassa o comando e o próprio formulário (Me) para o Módulo de Sistema
    Mdl_Sistema.AtivarModoDesenvolvedor Me
End Sub

' ==============================================================================
' EVENTO: Terminate
' OBJETIVO: Limpeza de memória e destruição de instâncias de classes
' ==============================================================================
Private Sub UserForm_Terminate()
    On Error Resume Next
    
    ' 1. Limpa as coleções de interatividade (Hover e Cursores)
    Set Mdl_UI_Efeitos.colBotoes = Nothing
    Set Mdl_UI_Efeitos.colBotoesMao = Nothing
        
    ' 2. Limpa as ferramentas de texto (Máscaras e Maiúsculas)
    Set Mdl_UI_Efeitos.colMascaras = Nothing
    Set Mdl_UI_Efeitos.colMaiusculas = Nothing
    
    ' 3. Força o processamento de eventos pendentes para liberar a UI
    DoEvents
End Sub



