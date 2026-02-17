Attribute VB_Name = "Mdl_UI_Designer"
' ==============================================================================
' NOME DO MÓDULO: Mdl_UI_Designer
' OBJETIVO:       Construir e Gerenciar a Interface SaaS
'                 (Ícones, Interatividade, Redimensionamento e Widgets)
' AUTOR:          Danyllo Pereira - LogicUp Solutions
' DATA:           Fevereiro/2026
' ==============================================================================
Option Explicit

' ==============================================================================
' SEÇÃO 1: CONFIGURAÇÕES VISUAIS E TIPOGRAFIA
' ==============================================================================

' -------------------------------------------------------------------------
' Propósito: Configura o design do botão (Label) de saída (SignOut)
' Parâmetros: Frm - O formulário onde o controle está localizado
' -------------------------------------------------------------------------
Public Sub ConfigIcoSair(ByVal Frm As Object)
    
    With Frm.IcoSair
        ' Define a fonte específica do Windows para exibir o ícone
        .Font.Name = "Segoe MDL2 Assets"
        
        ' Insere o ícone de "Sair" (seta apontando para a direita)
        .Caption = ChrW(&HF3B1)
        
        ' Ajusta o tamanho da fonte para o ícone
        .Font.Size = 20
        
        ' Aplica a cor cinza azulada (estado de repouso/padrão)
        .ForeColor = RGB(140, 155, 175)
        
        ' Configura o fundo como transparente (0 = fmBackStyleTransparent)
        .BackStyle = 0
        
        ' Remove a borda do controle (0 = fmBorderStyleNone)
        .BorderStyle = 0
    End With
    
End Sub

Public Sub ConfigIcoUsuarios(ByVal Frm As Object)
    
    With Frm.IcoUsuarios
        ' Define a fonte específica do Windows para exibir o ícone
        .Font.Name = "Segoe MDL2 Assets"
        
        ' Insere o ícone de "Sair" (seta apontando para a direita)
        .Caption = ChrW(&HE716)
        
        ' Ajusta o tamanho da fonte para o ícone
        .Font.Size = 20
        
        ' Aplica a cor cinza azulada (estado de repouso/padrão)
        .ForeColor = RGB(140, 155, 175)
        
        ' Configura o fundo como transparente (0 = fmBackStyleTransparent)
        .BackStyle = 0
        
        ' Remove a borda do controle (0 = fmBorderStyleNone)
        .BorderStyle = 0
    End With
    
End Sub

Public Sub ConfigIcoModoDev(ByVal Frm As Object)
    
    With Frm.IcoModoDev
        ' Define a fonte específica do Windows para exibir o ícone
        .Font.Name = "Segoe MDL2 Assets"
        
        ' Insere o ícone de "Sair" (seta apontando para a direita)
        .Caption = ChrW(&HE756)
        
        ' Ajusta o tamanho da fonte para o ícone
        .Font.Size = 20
        
        ' Aplica a cor cinza azulada (estado de repouso/padrão)
        .ForeColor = RGB(140, 155, 175)
        
        ' Configura o fundo como transparente (0 = fmBackStyleTransparent)
        .BackStyle = 0
        
        ' Remove a borda do controle (0 = fmBorderStyleNone)
        .BorderStyle = 0
    End With
    
End Sub

' -------------------------------------------------------------------------
' Propósito: Configura a tipografia e as cores das Labels de perfil do usuário
' Parâmetros: Frm - O formulário onde os controles estão localizados
' -------------------------------------------------------------------------
Public Sub ConfigInfoUsuario(ByVal Frm As Object)
    
    ' --- 1. Label do Nome do Usuário Logado ("Administrador") ---
    With Frm.LbpUsuarioLogado
        ' Utiliza a variante Semibold nativa para destacar o título
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 14
        
        ' Define a cor branca para o texto principal
        .ForeColor = RGB(255, 255, 255)
        
        ' Define alinhamento à esquerda (1 = fmTextAlignLeft)
        .TextAlign = 1
        
        ' Remove fundo e bordas para integrar suavemente com o menu
        .BackStyle = 0
        .BorderStyle = 0
    End With
    
    ' --- 2. Label do Nível de Acesso ("Admin") ---
    With Frm.LbpUsuarioNivel
        ' Utiliza a fonte padrão sem negrito para texto secundário
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .Font.Bold = False
        
        ' Aplica a cor cinza azulada para criar hierarquia visual
        .ForeColor = RGB(140, 155, 175)
        
        ' Define alinhamento à esquerda (1 = fmTextAlignLeft)
        .TextAlign = 1
        
        ' Remove fundo e bordas
        .BackStyle = 0
        .BorderStyle = 0
    End With
    
End Sub


' ==============================================================================
' SEÇÃO 2: LAYOUT E REDIMENSIONAMENTO DE TELA
' ==============================================================================

' -------------------------------------------------------------------------
' Propósito: Ajusta o Form principal para ocupar 90% da tela e centraliza
' Parâmetros: Frm - O formulário principal
' -------------------------------------------------------------------------
Public Sub AjustarTamanhoFormulario(ByVal Frm As Object)
    Dim LarguraTela As Double
    Dim AlturaTela As Double

    ' Pega a resolução baseada na janela do Excel
    LarguraTela = Application.Width
    AlturaTela = Application.Height

    ' Define o tamanho do formulário para 90%
    Frm.Width = LarguraTela * 0.9
    Frm.Height = AlturaTela * 0.9

    ' Centraliza na tela matematicamente
    Frm.Left = (LarguraTela - Frm.Width) / 2
    Frm.Top = (AlturaTela - Frm.Height) / 2
End Sub

' -------------------------------------------------------------------------
' Propósito: Ajusta responsivamente o layout do menu lateral e seus componentes
' Parâmetros: Frm - O formulário principal
' -------------------------------------------------------------------------
Public Sub RedimensionarMenuLateral(ByVal Frm As Object)
    On Error Resume Next
    
    ' --- 1. Ajuste do Fundo do Menu (FrmMenu) ---
    ' Usa -2 para esconder as bordas nativas do Windows e InsideHeight para a área útil
    With Frm.FrmMenu
        .Left = -2
        .Top = -2
        .Width = 200
        ' Soma +4 para compensar o Top -2 e garantir que esconda a borda inferior também
        .Height = Frm.InsideHeight + 4
        .BackColor = RGB(34, 45, 60)
        .BorderStyle = 0
    End With
    
    ' --- 2. Ajuste da Logo (ImgLogo) ---
    ' Alinhada no topo do formulário (ou dentro do Frame, se estiver contida nele)
    With Frm.ImgLogo
        .Left = 0
        .Top = 0
        .Width = 200 ' Usa os 200 para alinhar perfeitamente com a largura do menu
        .Height = 86 ' Altura fixa para não distorcer a imagem
    End With
    
    ' --- 3. Área dos Botões do Menu (Espaço Reservado) ---
    ' [COMENTÁRIO] Aqui entrará a lógica de distribuição das Labels/Botões.
    ' Como a ImgLogo ocupa 86 de altura, o Topo Inicial para os botões deverá
    ' ser a partir de Top = 90 ou 100, criando um respiro abaixo da logo.
    
    ' --- 4. Ajuste do Rodapé com Usuário e Botão Sair (FraRodapeMenu) ---
    With Frm.FraRodapeMenu
        .Left = 0
        .Width = 200
        ' Ancorado no limite inferior da tela visível, com uma margem de 10 para respirar
        .Top = Frm.InsideHeight - .Height - 10
        .BackColor = RGB(34, 45, 60)
        .BorderStyle = 0
    End With
    
End Sub

' -------------------------------------------------------------------------
' Propósito: Ajusta a barra superior (FrmTitulo) e seus elementos internos
' Parâmetros: Frm - O formulário principal
' -------------------------------------------------------------------------
Public Sub RedimensionarBarraSuperior(ByVal Frm As Object)
    On Error Resume Next
    
    ' --- 1. Ajuste do Frame da Barra Superior (FrmTitulo) ---
    With Frm.FrmTitulo
        ' Começa exatamente onde o menu lateral termina
        .Left = Frm.FrmMenu.Left + Frm.FrmMenu.Width - 2
        .Top = -2 ' Esconde a borda superior
        
        ' Altura FIXA em 60 para um visual mais limpo e moderno
        .Height = 60
        
        ' Estica até o final da tela direita (+4 para esconder bordas nativas)
        .Width = Frm.InsideWidth - .Left + 4
        
        ' Aplica a cor Azul elegante
        .BackColor = RGB(45, 60, 80)
        .BorderStyle = 0
    End With
    
    ' --- 2. Ajuste do Título da Tela (LbTitulo) ---
    With Frm.LbTitulo
        .AutoSize = False ' Desliga temporariamente para aplicar a fonte
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 20
        .ForeColor = RGB(255, 255, 255) ' Branco
        .BackStyle = 0 ' Transparente
        .BorderStyle = 0
        .AutoSize = True
        
        ' Centraliza VERTICALMENTE na barra
        .Top = (Frm.FrmTitulo.Height - .Height) / 2
        
        ' Centraliza HORIZONTALMENTE na barra superior de forma dinâmica
        .Left = (Frm.FrmTitulo.Width - .Width) / 2
    End With
    
    ' --- 3. Ajuste da Data (LbData) ---
    With Frm.LbData
        .AutoSize = False
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .ForeColor = RGB(140, 155, 175) ' Cinza azulado
        .BackStyle = 0 ' Transparente
        .BorderStyle = 0
        .TextAlign = 3 ' 3 = fmTextAlignRight
        
        ' Gera a data formatada
        .Caption = StrConv(Format(Date, "dddd, dd/mm/yyyy"), vbProperCase)
        .AutoSize = True
        
        ' Centraliza verticalmente
        .Top = (Frm.FrmTitulo.Height - .Height) / 2
        
        ' Cola no lado direito da tela
        .Left = Frm.FrmTitulo.Width - .Width - 30
    End With
    
End Sub

' -------------------------------------------------------------------------
' Propósito: Ajusta a MultiPage principal para ocupar o resto da tela
' Parâmetros: Frm - O formulário principal
' -------------------------------------------------------------------------
Public Sub RedimensionarConteudoPrincipal(ByVal Frm As Object)
    On Error Resume Next
    
    With Frm.MultiPagMain
        ' 1. ESTILO DA MULTIPAGE: 2 = fmTabStyleNone (Esconde abas nativas)
        .Style = 2
        
        ' 2. POSIÇÃO HORIZONTAL: Segue o mesmo alinhamento exato da barra de título
        .Left = Frm.FrmMenu.Left + Frm.FrmMenu.Width - 2
        
        ' 3. POSIÇÃO VERTICAL: Mantido em 40 para junção perfeita com a barra de título
        .Top = 40
        
        ' 4. LARGURA: Estica até o limite direito da tela (+4 para esconder borda)
        .Width = Frm.InsideWidth - .Left + 4
        
        ' 5. ALTURA: Estica até o limite inferior da tela (+4 para esconder borda)
        .Height = Frm.InsideHeight - .Top + 4
        
        ' Remove o fundo padrão
        .BackStyle = 0
    End With
    
End Sub

