VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Usf_Login 
   Caption         =   "LogicUp Solutions"
   ClientHeight    =   11190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7860
   OleObjectBlob   =   "Usf_Login.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Usf_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ==============================================================================
' Formulário: Usf_Login
' Objetivo: Interface visual (View). Toda a lógica está no Mdl_Login.
' ==============================================================================

' --- 1. Inicialização ---
Private Sub UserForm_Initialize()
    ' 1. Configurações Físicas
    Me.Height = 570
    Me.Width = 400
    
    ' 2. Infraestrutura
    Mdl_InstalacaoBD.VerificarEstruturaBanco
    Mdl_Login.ConfigurarFormulario Me
    
    ' 3. Ativação dos Efeitos (Chamadas Individuais Organizadas)
    Mdl_UI_Efeitos.AtivarMaiusculas Me
    Mdl_UI_Efeitos.AtivarCursorEMao Me
    Mdl_UI_Efeitos.AtivarMascaras Me
    Mdl_UI_Efeitos.AplicarDestaque Me
    Mdl_UI_Efeitos.PersonalizarBarraTitulo Me, RGB(33, 95, 152), RGB(255, 255, 255)
    
    ' 5. Foco
    Me.TxUser.SetFocus
End Sub

' --- 2. Botões de Ação (Login, Registrar, Resetar) ---
' Botão de Confirmar Login
Private Sub BtnLogin_Click()
    Mdl_Login.ExecutarLogin Me
End Sub

' Botão de Confirmar Cadastro
 Private Sub BtnCadastrar_Click()
    Mdl_Login.ExecutarCadastro Me
 End Sub
 
' Botão de Confirmar Reset
Private Sub BtnReset_Click()
    Mdl_Login.ExecutarResetSenha Me
End Sub

' --- 3. Navegação (Labels de Link) ---

' Ir para Cadastro
Private Sub LblRegister_Click()
    Mdl_Login.IrParaCadastro Me
End Sub

' Ir para Esqueci Senha
Private Sub LblForgot_Click()
    Mdl_Login.IrParaEsqueciSenha Me
End Sub

' Voltar para Login (Vindo do Cadastro)
Private Sub LblBackToLogin_Click()
    Mdl_Login.IrParaLogin Me
End Sub

' Voltar para Login (Vindo do Reset)
Private Sub LblBackToLogin2_Click()
    Mdl_Login.IrParaLogin Me
End Sub

' --- Ícone na aba de Login ---
Private Sub LblIconeLogin_Click()
    Mdl_Login.AlternarVisualizacaoSenha Me.TxPass, Me.LblIconeLogin, Me.LblVer, Me.LblEsconder
End Sub

' --- Ícone na aba de Cadastro ---
Private Sub LblIconeReg_Click()
    Mdl_Login.AlternarVisualizacaoSenha Me.TxRegPass, Me.LblIconeReg, Me.LblVer, Me.LblEsconder
    Mdl_Login.AlternarVisualizacaoSenha Me.TxRegPassConfirm, Me.LblIconeReg, Me.LblVer, Me.LblEsconder
End Sub

' --- Ícone na aba de Esqueci Senha ---
Private Sub LblIconeReset_Click()
    Mdl_Login.AlternarVisualizacaoSenha Me.TxResetPass, Me.LblIconeReset, Me.LblVer, Me.LblEsconder
    Mdl_Login.AlternarVisualizacaoSenha Me.TxResetNewPass, Me.LblIconeReset, Me.LblVer, Me.LblEsconder
End Sub

Private Sub MultiPagLogin_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Mdl_UI_Efeitos.ResetarDestaque
End Sub

Private Sub BtnSair_Click()
    ' Passa o próprio form de login para a rotina de encerramento
    Mdl_Login.TerminarSessao Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' CloseMode 0 (vbFormControlMenu) é o clique no "X"
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "Por favor, utilize o botão 'Sair' para encerrar o sistema.", _
               vbExclamation, "BeautyTech"
    End If
End Sub


