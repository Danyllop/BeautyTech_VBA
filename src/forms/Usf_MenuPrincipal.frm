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
' ==============================================================================
' NOME DO ARQUIVO: Usf_MenuPrincipal
' PROJETO:         Sistema BeautyTech - Gestão Integrada
' DESCRIÇÃO:       Controlador Mestre da Interface (Dashboard e Navegação)
' AUTOR:           LogicUp Solutions
' DATA:            Fevereiro/2026
' ==============================================================================
Option Explicit
' ==============================================================================
' 1. GERENCIAMENTO DE MEMÓRIA E ESTADO (COLEÇÕES PERSISTENTES)
' Estas variáveis mantêm as classes de eventos "vivas" durante a execução.
' ==============================================================================
Private ColEfeitos      As Collection   ' Gerencia Cursor Mãozinha e Efeitos Gerais
Private ColBotoes       As Collection   ' Gerencia Animações de Botões (Zoom/Negrito)
Private ColMenuLateral  As Collection   ' Gerencia Lógica Visual do Menu
Private ColInputs       As Collection   ' Gerencia Máscaras e Validações de Texto
Private ColIconesSys    As Collection   ' Gerencia Ícones Especiais (Sair, Dev, Users)

' ==============================================================================
' 2. CICLO DE VIDA DO FORMULÁRIO (Eventos de Sistema)
' ==============================================================================

Private Sub UserForm_Initialize()
    ' A. Sincroniza a interface com os dados do Login (Global)
    AtualizarDadosSessao
    
    ' B. Configura a estrutura visual estática (Cores, Tamanhos, Barra Título)
    ConfigurarInterfaceVisual
    
    ' C. Inicializa os motores de interatividade (Hover, Menu, Ícones Dinâmicos)
    CarregarMotoresDeInteratividade
    
    ' D. Aplica regras de negócio e define o estado inicial
    AplicarEstadoInicial
End Sub

Private Sub UserForm_Terminate()
    ' GARBAGE COLLECTION: Limpeza explícita para evitar vazamento de memória
    ' (Fundamental para estabilidade em sessões longas)
    Mdl_UI_Efeitos.LimparFoco
    Mdl_UI_Efeitos.LimparTudosFocos
    
    Set ColEfeitos = Nothing
    Set ColBotoes = Nothing
    Set ColMenuLateral = Nothing
    Set ColInputs = Nothing
    Set ColIconesSys = Nothing
End Sub

' Previne o fechamento acidental pelo "X" do Windows
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then ' 0 = Usuário clicou no X
        Cancel = True
        MsgBox "Para manter a integridade dos dados, por favor utilize o botão 'Sair' no menu lateral.", _
               vbExclamation, "Sistema BeautyTech"
    End If
End Sub

' Responsividade: Recalcula layout ao redimensionar a janela
Private Sub UserForm_Resize()
    ' Bloqueia erro ao minimizar a janela (InsideHeight fica muito pequeno)
    If Me.InsideHeight > 100 Then
        Mdl_UI_Designer.RedimensionarMenuLateral Me
        Mdl_UI_Designer.RedimensionarBarraSuperior Me
        Mdl_UI_Designer.RedimensionarConteudoPrincipal Me
    End If
End Sub

' ==============================================================================
' 3. SUB-ROTINAS DE CONFIGURAÇÃO (Helpers Privados)
' ==============================================================================

Private Sub AtualizarDadosSessao()
    ' Verifica integridade do Login
    If Mdl_VariaveisGlobais.UsuarioLogado = False Then
        Me.LbpUsuarioLogado.Caption = "Visitante"
        Me.LbpUsuarioNivel.Caption = "---"
        Exit Sub
    End If

    ' Preenche a UI com os dados reais carregados no Login
    Me.LbpUsuarioLogado.Caption = Mdl_VariaveisGlobais.UsuarioLogin
    Me.LbpUsuarioNivel.Caption = Mdl_VariaveisGlobais.UsuarioNivel
End Sub

Private Sub ConfigurarInterfaceVisual()
    ' 1. Definições de Geometria Base
    Me.Height = 540
    Me.Width = 980
    
    ' 2. Personalização Avançada da Janela (API Windows)
    '    Cor: Azul Profundo (RGB 33, 95, 152) | Texto: Branco
    Call Mdl_UI_Efeitos.PersonalizarBarraTitulo(Me, RGB(33, 95, 152), RGB(255, 255, 255))
    
    ' 3. Ajuste inicial de layout responsivo
    Mdl_UI_Designer.AjustarTamanhoFormulario Me
    
    ' 3. Ajuste inicial de layout do Usuário logado
    Mdl_UI_Designer.ConfigInfoUsuario Me
    
End Sub

Private Sub CarregarMotoresDeInteratividade()
    ' Instancia as coleções (Containers vazios)
    Set ColEfeitos = New Collection
    Set ColIconesSys = New Collection
    Set ColInputs = New Collection
    
    ' A. Fábrica de Efeitos Visuais (Cursor e Botões)
    Set ColEfeitos = Mdl_UI_Efeitos.CriarEfeitosMaozinha(Me)
    Set ColBotoes = Mdl_UI_Efeitos.CriarSimpleButton(Me)
    
    ' B. Fábrica do Menu Lateral (Imagens de Estado)
    Set ColMenuLateral = Mdl_UI_Efeitos.CriarMenuLateral(Me, Me.ImgStart, Me.ImgHover, Me.ImgAtivo)
    
    ' C. Fábrica de Ícones do Sistema (Sair, Dev, Users)
    '    O Módulo Designer cria as classes clsIconeHover e as guarda em ColIconesSys
    Mdl_UI_Designer.ConfigurarIconesDoSistema Me, ColIconesSys
    
    ' D. Fábrica de Inputs (Agrupa Máscaras e Maiúsculas na mesma coleção)
    Dim TempColl As Collection
    
    Set TempColl = Mdl_UI_Efeitos.AtivarMaiusculas(Me)
    AdicionarColecao ColInputs, TempColl
    
    Set TempColl = Mdl_UI_Efeitos.AtivarMascaras(Me)
    AdicionarColecao ColInputs, TempColl
End Sub

Private Sub AplicarEstadoInicial()
    ' 1. Aplica Permissões de Acesso (Ex: Ocultar Financeiro para Recepcionista)
    Mdl_Sistema.AplicarPermissoes Me
    
    ' 2. Define Dashboard como tela inicial se o menu carregou corretamente
    If Not ColMenuLateral Is Nothing Then
        NavegarMenu "BtnDashboard", 0, "Dashboard"
    End If
End Sub

' Helper utilitário para fundir coleções (Mantém o código limpo)
Private Sub AdicionarColecao(ByRef Destino As Collection, ByVal Origem As Collection)
    Dim item As Object
    If Not Origem Is Nothing Then
        For Each item In Origem
            Destino.Add item
        Next item
    End If
End Sub

' ==============================================================================
' 4. NAVEGAÇÃO E AÇÕES DO USUÁRIO
' ==============================================================================

' --- CONTROLADOR CENTRAL DE NAVEGAÇÃO ---
' Evita repetição de código nos botões. Centraliza a lógica de troca de tela.
Private Sub NavegarMenu(ByVal NomeBotao As String, ByVal IndexPagina As Integer, ByVal TituloPagina As String)
    ' 1. Atualiza visual do Menu Lateral (Quem está ativo/iluminado)
    Mdl_UI_Efeitos.SelecionarBotao ColMenuLateral, NomeBotao
    
    ' 2. Executa a troca física da MultiPage e atualização de títulos
    Mdl_Sistema.NavegarPara Me, IndexPagina, TituloPagina
End Sub

' --- BOTÕES DO MENU PRINCIPAL ---

Private Sub BtnDashBoard_Click()
    NavegarMenu "BtnDashboard", 0, "Dashboard"
End Sub

Private Sub BtnAgenda_Click()
    NavegarMenu "BtnAgenda", 1, "Agenda"
End Sub

Private Sub BtnClientes_Click()
    NavegarMenu "BtnClientes", 2, "Gestão de Clientes"
End Sub

Private Sub BtnFinanceiro_Click()
    NavegarMenu "BtnFinanceiro", 3, "Gestão de Finanças"
End Sub

Private Sub BtnEstoqueTécnico_Click()
    NavegarMenu "BtnEstoqueTécnico", 3, "Gestão de Estoque"
End Sub

Private Sub IcoUsuarios_Click()
    ' Acesso rápido à gestão de usuários via ícone
    NavegarMenu "IcoUsuarios", 4, "Gestão de Usuários"
End Sub

' --- ÍCONES DE AÇÕES DE SISTEMA ---
Private Sub IcoSair_Click()
    Dim Resposta As VbMsgBoxResult
    Resposta = MsgBox("Deseja realmente encerrar sua sessão no BeautyTech?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Encerrar Sistema")
                      
    If Resposta = vbYes Then
        ' Descarrega o formulário e chama rotina de encerramento global
        Unload Me
        EncerrarSistemaBeautyTech
    End If
End Sub

Private Sub IcoModoDev_Click()
    ' Ativa ferramentas de desenvolvedor (Logs, Console, Debug)
    Mdl_Sistema.AtivarModoDesenvolvedor Me
End Sub

' ==============================================================================
' 5. EVENTOS DE UX (Mouse Move / Hover Reset)
' ==============================================================================

Private Sub FrmMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Se o mouse sair de um botão e tocar no fundo do menu, reseta o brilho
    Mdl_UI_Efeitos.ResetarHoverMenu ColMenuLateral
End Sub

Private Sub FraRodapeMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Garante que os ícones do rodapé percam o foco (cor) ao sair deles
    Mdl_UI_Efeitos.LimparFocoIcone
End Sub

