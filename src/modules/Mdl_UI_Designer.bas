Attribute VB_Name = "Mdl_UI_Designer"
' ==============================================================================
' NOME DO M√ìDULO: Mdl_UI_Designer
' OBJETIVO:       Construir e Gerenciar a Interface SaaS
'                 (√çcones, Interatividade, Redimensionamento e Widgets)
' AUTOR:          Danyllo Pereira - LogicUp Solutions
' DATA:           Fevereiro/2026
' ==============================================================================
Option Explicit

' ==============================================================================
' SE√á√ÉO 1: CONFIGURA√á√ïES VISUAIS E TIPOGRAFIA
' ==============================================================================
'
' -------------------------------------------------------------------------
' CONTROLADOR MESTRE: Chama todas as configura√ß√µes de uma vez
' -------------------------------------------------------------------------
Public Sub ConfigurarIconesDoSistema(ByVal Frm As Object, ByVal ColStorage As Collection)
    ' Chama as rotinas individuais passando a Cole√ß√£o para guardar a classe
    ConfigIcoSair Frm, ColStorage
    ConfigIcoModoDev Frm, ColStorage
    ConfigIcoUsuarios Frm, ColStorage
    
    ' Se tiver outros no futuro, adicione aqui...
End Sub

' -------------------------------------------------------------------------
' CONFIGURA√á√ÉO: √çcone Sair (Vermelho)
' -------------------------------------------------------------------------
Public Sub ConfigIcoSair(ByVal Frm As Object, ByVal ColStorage As Collection)
    Dim cls As clsIconeHover
    
    With Frm.IcoSair
        ' 1. Configura√ß√£o Visual (Est√°tica)
        .Font.Name = "Segoe MDL2 Assets"
        .Caption = ChrW(&HF3B1) ' Seta Sair
        .Font.Size = 20
        .ForeColor = RGB(140, 155, 175)
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .ControlTipText = "Sair do Sistema"
        
        ' 2. Configura√ß√£o de Comportamento (Classe Hover)
        Set cls = New clsIconeHover
        ' Define a cor Vermelho Suave (RGB 255, 80, 80)
        cls.Inicializar Frm.IcoSair, RGB(255, 80, 80)
        
        ' Guarda na cole√ß√£o para manter o efeito funcionando
        ColStorage.Add cls
    End With
End Sub

' -------------------------------------------------------------------------
' CONFIGURA√á√ÉO: √çcone Modo Dev (Verde Tech)
' -------------------------------------------------------------------------
Public Sub ConfigIcoModoDev(ByVal Frm As Object, ByVal ColStorage As Collection)
    Dim cls As clsIconeHover
    
    With Frm.IcoModoDev
        ' 1. Configura√ß√£o Visual
        .Font.Name = "Segoe MDL2 Assets"
        .Caption = ChrW(&HE756) ' Prompt de Comando
        .Font.Size = 20
        .ForeColor = RGB(140, 155, 175)
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .ControlTipText = "Modo Desenvolvedor"
        
        ' 2. Configura√ß√£o de Comportamento
        Set cls = New clsIconeHover
        ' Define a cor Verde Matrix (RGB 0, 200, 100)
        cls.Inicializar Frm.IcoModoDev, RGB(0, 200, 100)
        
        ColStorage.Add cls
    End With
End Sub

' -------------------------------------------------------------------------
' CONFIGURA√á√ÉO: √çcone Usu√°rios (Padr√£o Branco)
' -------------------------------------------------------------------------
Public Sub ConfigIcoUsuarios(ByVal Frm As Object, ByVal ColStorage As Collection)
    Dim cls As clsIconeHover
    
    With Frm.IcoUsuarios
        ' 1. Configura√ß√£o Visual
        .Font.Name = "Segoe MDL2 Assets"
        .Caption = ChrW(&HE716) ' Pessoas/User
        .Font.Size = 20
        .ForeColor = RGB(140, 155, 175)
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        
        ' 2. Configura√ß√£o de Comportamento
        Set cls = New clsIconeHover
        ' Sem cor definida = usa o padr√£o Branco (definido na classe)
        cls.Inicializar Frm.IcoUsuarios
        
        ColStorage.Add cls
    End With
End Sub

' -------------------------------------------------------------------------
' Prop√≥sito: Configura a tipografia e as cores das Labels de perfil do usu√°rio
' Par√¢metros: Frm - O formul√°rio onde os controles est√£o localizados
' -------------------------------------------------------------------------
Public Sub ConfigInfoUsuario(ByVal Frm As Object)

    ' --- 1. Label do Nome do Usu√°rio Logado ("Administrador") ---
    With Frm.LbpUsuarioLogado
        ' Utiliza a variante Semibold nativa para destacar o t√≠tulo
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 14

        ' Define a cor branca para o texto principal
        .ForeColor = RGB(255, 255, 255)

        ' Define alinhamento √† esquerda (1 = fmTextAlignLeft)
        .TextAlign = 1

        ' Remove fundo e bordas para integrar suavemente com o menu
        .BackStyle = 0
        .BorderStyle = 0
    End With

    ' --- 2. Label do N√≠vel de Acesso ("Admin") ---
    With Frm.LbpUsuarioNivel
        ' Utiliza a fonte padr√£o sem negrito para texto secund√°rio
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .Font.Bold = False

        ' Aplica a cor cinza azulada para criar hierarquia visual
        .ForeColor = RGB(140, 155, 175)

        ' Define alinhamento √† esquerda (1 = fmTextAlignLeft)
        .TextAlign = 1

        ' Remove fundo e bordas
        .BackStyle = 0
        .BorderStyle = 0
    End With

End Sub


' ==============================================================================
' SE√á√ÉO 2: LAYOUT E REDIMENSIONAMENTO DE TELA
' ==============================================================================

' -------------------------------------------------------------------------
' Prop√≥sito: Ajusta o Form principal para ocupar 90% da tela e centraliza
' Par√¢metros: Frm - O formul√°rio principal
' -------------------------------------------------------------------------
Public Sub AjustarTamanhoFormulario(ByVal Frm As Object)
    Dim LarguraTela As Double
    Dim AlturaTela As Double

    ' Pega a resolu√ß√£o baseada na janela do Excel
    LarguraTela = Application.Width
    AlturaTela = Application.Height

    ' Define o tamanho do formul√°rio para 90%
    Frm.Width = LarguraTela * 0.9
    Frm.Height = AlturaTela * 0.9

    ' Centraliza na tela matematicamente
    Frm.Left = (LarguraTela - Frm.Width) / 2
    Frm.Top = (AlturaTela - Frm.Height) / 2
End Sub

' -------------------------------------------------------------------------
' Prop√≥sito: Ajusta responsivamente o layout do menu lateral e seus componentes
' Par√¢metros: Frm - O formul√°rio principal
' -------------------------------------------------------------------------
Public Sub RedimensionarMenuLateral(ByVal Frm As Object)
    On Error Resume Next

    ' --- 1. Ajuste do Fundo do Menu (FrmMenu) ---
    ' Usa -2 para esconder as bordas nativas do Windows e InsideHeight para a √°rea √∫til
    With Frm.FrmMenu
        .Left = -2
        .Top = -2
        .Width = 200
        ' Soma +4 para compensar o Top -2 e garantir que esconda a borda inferior tamb√©m
        .Height = Frm.InsideHeight + 4
        .BackColor = RGB(34, 45, 60)
        .BorderStyle = 0
    End With

    ' --- 2. Ajuste da Logo (ImgLogo) ---
    With Frm.ImgLogo
        .Left = 0
        .Top = 0
        .Width = Frm.FrmMenu.Width ' Amarra dinamicamente ‡ largura exata do menu
        .Height = 86
        .PictureSizeMode = 3 ' fmPictureSizeModeZoom (Garante proporÁ„o sem distorcer)
        
        .BackColor = RGB(34, 45, 60)
        .BackStyle = 0
        .BorderStyle = 0
    End With

    ' --- 3. √Årea dos Bot√µes do Menu (Espa√ßo Reservado) ---
    ' [COMENT√ÅRIO] Aqui entrar√° a l√≥gica de distribui√ß√£o das Labels/Bot√µes.
    ' Como a ImgLogo ocupa 86 de altura, o Topo Inicial para os bot√µes dever√°
    ' ser a partir de Top = 90 ou 100, criando um respiro abaixo da logo.

    ' --- 4. Ajuste do Rodap√© com Usu√°rio e Bot√£o Sair (FraRodapeMenu) ---
    With Frm.FraRodapeMenu
        .Left = 0
        .Width = 200
        ' Ancorado no limite inferior da tela vis√≠vel, com uma margem de 10 para respirar
        .Top = Frm.InsideHeight - .Height - 10
        .BackColor = RGB(34, 45, 60)
        .BorderStyle = 0
    End With

End Sub

' -------------------------------------------------------------------------
' PropÛsito: Ajusta a MultiPage principal para ocupar o resto da tela
' Par‚metros: Frm - O formul·rio principal
' -------------------------------------------------------------------------
Public Sub RedimensionarBarraSuperior(ByVal Frm As Object)
    On Error Resume Next

    ' --- 1. Ajuste do Frame da Barra Superior (FrmTitulo) ---
    With Frm.FrmTitulo
        ' Come√ßa exatamente onde o menu lateral termina
        .Left = Frm.FrmMenu.Left + Frm.FrmMenu.Width - 2
        .Top = -2 ' Esconde a borda superior

        ' Altura FIXA em 60 para um visual mais limpo e moderno
        .Height = 60

        ' Estica at√© o final da tela direita (+4 para esconder bordas nativas)
        .Width = Frm.InsideWidth - .Left + 4

        ' Aplica a cor Azul elegante
        .BackColor = RGB(45, 60, 80)
        .BorderStyle = 0
    End With

    ' --- 2. Ajuste do T√≠tulo da Tela (LbTitulo) ---
    With Frm.LbTitulo
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 20
        .ForeColor = RGB(255, 255, 255) ' Branco
        .BackStyle = 0 ' Transparente
        .BorderStyle = 0
        .Height = 40
        
        ' Centraliza VERTICALMENTE na barra
        .Top = (Frm.FrmTitulo.Height - .Height) / 2

        ' Centraliza HORIZONTALMENTE na barra superior de forma din√¢mica
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
' PropÛsito: Ajusta a MultiPage principal para ocupar o resto da tela
' Par‚metros: Frm - O formul·rio principal
' -------------------------------------------------------------------------
Public Sub RedimensionarConteudoPrincipal(ByVal Frm As Object)
    On Error Resume Next

    With Frm.MultiPagMain
        ' 1. ESTILO DA MULTIPAGE: 2 = fmTabStyleNone (Esconde abas nativas)
        .Style = 2

        ' 2. POSI«√O HORIZONTAL: Segue o mesmo alinhamento exato da barra de tÌtulo
        .Left = Frm.FrmMenu.Left + Frm.FrmMenu.Width - 2

        ' 3. POSI«√O VERTICAL: Mantido em 40 para junÁ„o perfeita com a barra de tÌtulo
        .Top = 40

        ' 4. LARGURA: Estica atÈ o limite direito da tela (+4 para esconder borda)
        .Width = Frm.InsideWidth - .Left + 4

        ' 5. ALTURA: Estica atÈ o limite inferior da tela (+4 para esconder borda)
        .Height = Frm.InsideHeight - .Top + 4

        ' Remove o fundo padr„o
        .BackStyle = 0
    End With

End Sub

' -------------------------------------------------------------------------
' PropÛsito: Renderiza a Grid Responsiva (Sem barras extras e sem foco visual)
' -------------------------------------------------------------------------
Public Sub RenderizarGridResponsiva(ByVal Frm As Object, ByVal ItemPagina As Variant, ByVal LstAlvo As Object, ByVal ArrLabels As Variant, ByVal ArrPorcentagens As Variant)
    Dim Pagina As Object
    Dim Margem As Double: Margem = 20
    Dim TopoInicial As Double: TopoInicial = 70
    Dim AlturaCabecalho As Double: AlturaCabecalho = 28
    
    Dim LarguraTotal As Double
    Dim LarguraUtilColunas As Double
    Dim i As Integer
    Dim lbl As Object
    
    Dim PorcentagemAcumulada As Double
    Dim PontoAtualX As Long
    Dim ProximoPontoX As Long
    Dim LarguraColuna As Long
    Dim LarguraHeader As Long
    Dim ColWidthsStr As String
    Dim TextoLimpo As String
    
    ' Paleta de Cores
    Dim CorFundoPadrao As Long: CorFundoPadrao = RGB(33, 47, 61)
    Dim CorFundoCabecalho As Long: CorFundoCabecalho = RGB(26, 37, 49)
    Dim CorTextoPadrao As Long: CorTextoPadrao = RGB(255, 255, 255)
    
    On Error Resume Next
    
    If IsObject(ItemPagina) Then Set Pagina = ItemPagina Else Set Pagina = Frm.MultiPagMain.Pages(ItemPagina)
    If Pagina Is Nothing Then Exit Sub

    ' 1. C¡LCULO DE ¡REA
    LarguraTotal = Frm.MultiPagMain.Width - (Margem * 2) - 8
    If LarguraTotal < 100 Then Exit Sub
    LarguraUtilColunas = LarguraTotal - 15 ' Limite seguro para n„o gerar a barra horizontal

    PorcentagemAcumulada = 0
    ColWidthsStr = ""

    ' 2. CABE«ALHOS E COLUNAS
    For i = LBound(ArrLabels) To UBound(ArrLabels)
        Set lbl = Pagina.Controls(ArrLabels(i))
        If Not lbl Is Nothing Then
            
            PontoAtualX = Int(LarguraUtilColunas * PorcentagemAcumulada)
            
            If i = UBound(ArrLabels) Then
                ' M¡GICA 1: O CabeÁalho vai atÈ o fim (LarguraTotal) para selar a tela
                LarguraHeader = LarguraTotal - PontoAtualX
                ' M¡GICA 2: A Coluna da ListBox para no limite ˙til, matando a barra de rolagem
                LarguraColuna = LarguraUtilColunas - PontoAtualX
            Else
                PorcentagemAcumulada = PorcentagemAcumulada + ArrPorcentagens(i)
                ProximoPontoX = Int(LarguraUtilColunas * PorcentagemAcumulada)
                LarguraColuna = ProximoPontoX - PontoAtualX
                LarguraHeader = LarguraColuna
            End If
            
            With lbl
                .Left = Margem + PontoAtualX
                .Top = TopoInicial
                .Width = LarguraHeader
                .Height = AlturaCabecalho
                
                .BorderStyle = 0
                .SpecialEffect = 0
                .BackStyle = 1
                .BackColor = CorFundoCabecalho
                .ForeColor = CorTextoPadrao
                .Font.Name = "Segoe UI Semibold"
                .Font.Size = 10
                .TextAlign = 1
                
                TextoLimpo = Application.WorksheetFunction.Trim(.Caption)
                .Caption = "  " & UCase(TextoLimpo)
            End With
            
            ColWidthsStr = ColWidthsStr & LarguraColuna & " pt;"
        End If
    Next i

    ' 3. LISTBOX LIMPA
    With LstAlvo
        .Left = Margem
        .Top = TopoInicial + AlturaCabecalho - 1
        .Width = LarguraTotal
        .Height = Frm.MultiPagMain.Height - .Top - Margem - 10
        
        .IntegralHeight = False
        .BorderStyle = 0
        .SpecialEffect = 0
        .BackColor = CorFundoPadrao
        .ForeColor = CorTextoPadrao
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .TextAlign = 1
        .ColumnCount = UBound(ArrLabels) + 1
        
        If Len(ColWidthsStr) > 0 Then
            .ColumnWidths = Left(ColWidthsStr, Len(ColWidthsStr) - 1)
        End If
        
        ' Garante que nenhum item comece selecionado por acidente
        .ListIndex = -1
    End With
    
    ' 4. M¡GICA 3: O REMOVEDOR DE FOCO (Sume com o pontilhado)
    ' Ao jogar o foco de volta para a MultiPage, a ListBox perde o contorno tracejado
    Frm.TxtPesquisaUser.SetFocus
       
    On Error GoTo 0
End Sub

' -------------------------------------------------------------------------
' PropÛsito: Centraliza a renderizaÁ„o de componentes especÌficos de cada aba
' Par‚metros: Frm - O formul·rio principal
' -------------------------------------------------------------------------
'Public Sub GerenciarRenderizacaoPaginas(ByVal Frm As Object)
'    On Error Resume Next
'
'    ' Usa Select Case para organizar de forma limpa as chamadas de cada tela
'    Select Case Frm.MultiPagMain.Value
'
'        Case 0 ' Aba 0: DashBoard
'            ' (EspaÁo reservado para as futuras rotinas da Dashboard)
'
'        Case 1 ' Aba 1: Agenda
'            ' (EspaÁo reservado para as futuras rotinas da Agenda)
'
'        Case 5 ' Aba 5: Gest„o de Usu·rios
'            ' 1. Configura a barra de pesquisa visualmente
'            Mdl_UI_Designer.ConfigurarPesquisaUsuario Frm
'
'            ' Chama a construÁ„o da grid passando os controles pelo Form (Frm)
'            RenderizarGridResponsiva Frm, 5, Frm.ListUsuarios, _
'                Array("LbID", "LbNome", "LbUsu·rio", "LbEmail", "LbPerfil", "LbDataCadastro"), _
'                Array(0.05, 0.25, 0.15, 0.25, 0.15, 0.15)
'
'        ' Case 6 ' PrÛxima tela... etc.
'
'    End Select
'End Sub

' -------------------------------------------------------------------------
' PropÛsito: Centraliza a renderizaÁ„o de componentes especÌficos de cada aba
' Par‚metros: Frm - O formul·rio principal
' -------------------------------------------------------------------------
Public Sub GerenciarRenderizacaoPaginas(ByVal Frm As Object)
    On Error Resume Next
    
    ' Usa Select Case para organizar de forma limpa as chamadas de cada tela
    Select Case Frm.MultiPagMain.Value
        
        Case 0 ' Aba 0: DashBoard
            ' (EspaÁo reservado para as futuras rotinas da Dashboard)
            
        Case 1 ' Aba 1: Agenda
            ' (EspaÁo reservado para as futuras rotinas da Agenda)
            
        Case 5 ' Aba 5: Gest„o de Usu·rios
            
            ' 1. Configura a barra de pesquisa visualmente (Posiciona ‡ esquerda)
            Mdl_UI_Designer.ConfigurarPesquisaUsuario Frm
            
            ' 2. Estica e renderiza a Tabela/Grid de acordo com a tela
            RenderizarGridResponsiva Frm, 5, Frm.ListUsuarios, _
                Array("LbID", "LbNome", "LbUsu·rio", "LbEmail", "LbPerfil", "LbDataCadastro"), _
                Array(0.05, 0.25, 0.15, 0.25, 0.15, 0.15)
                
            ' 3. ALINHAMENTO DIN¬MICO: Ancora o bot„o de Novo Usu·rio ‡ direita da Tabela
            Mdl_UI_Designer.PosicionarBotaoNovoUsuario Frm
                
        ' Case 6 ' PrÛxima tela... etc.
        
    End Select
End Sub

' -------------------------------------------------------------------------
' PropÛsito: Configura a Barra de Pesquisa (Efeito Overlay Moderno)
' -------------------------------------------------------------------------
Public Sub ConfigurarPesquisaUsuario(ByVal Frm As Object)
    Dim Margem As Double: Margem = 20
    Dim TopoPesquisa As Double: TopoPesquisa = 40
    Dim AlturaPesquisa As Double: AlturaPesquisa = 18 ' Ligeiramente maior para eleg‚ncia
    
    Dim LarguraTotalBusca As Double: LarguraTotalBusca = 280 ' Largura do bloco inteiro
    Dim LarguraIcone As Double: LarguraIcone = 18
    
    Dim CorFundoBusca As Long: CorFundoBusca = RGB(26, 37, 49)
    Dim CorTextoBusca As Long: CorTextoBusca = RGB(255, 255, 255)
    
    On Error Resume Next
    
    ' 1. A Caixa de Texto (O Fundo MaciÁo)
    With Frm.TxtPesquisaUser
        .Left = Margem
        .Top = TopoPesquisa
        .Width = LarguraTotalBusca ' Pega toda a ·rea
        .Height = AlturaPesquisa
        
        .BorderStyle = 0
        .SpecialEffect = 0
        .BackColor = CorFundoBusca
        .ForeColor = CorTextoBusca
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        
    End With
    
    ' 2. O Õcone (Transparente e Flutuante)
    With Frm.IcoPesquisaUser
        ' Posiciona o Ìcone dentro da extremidade direita da caixa de texto
        .Left = (Margem + LarguraTotalBusca) - LarguraIcone - 2
        
        ' A M¡GICA: Empurramos 4 pixels para baixo para centralizar a lupa perfeitamente
        .Top = TopoPesquisa + 3
        
        .Width = LarguraIcone
        .Height = AlturaPesquisa
        
        .Font.Name = "Segoe MDL2 Assets"
        .Caption = ChrW(&HE721)
        .Font.Size = 11 ' Tamanho ideal do Ìcone
        .ForeColor = RGB(140, 155, 175)
        
        ' Fundo Transparente! Ele assume a cor da TextBox que est· atr·s.
        .BackStyle = 0 ' 0 = fmBackStyleTransparent
        .BorderStyle = 0
        .TextAlign = 2 ' Centro horizontal
        
        ' ForÁa o Ìcone a ficar na frente do texto
        .ZOrder 0
    End With
    
    On Error GoTo 0
End Sub

' -------------------------------------------------------------------------
' PropÛsito: Posiciona o Bot„o "Novo Usu·rio" alinhado ‡ direita da Grid
' -------------------------------------------------------------------------
Public Sub PosicionarBotaoNovoUsuario(ByVal Frm As Object)
    Dim TopoPesquisa As Double
    Dim AlturaPesquisa As Double
    
    On Error Resume Next
    
    ' Pega as referÍncias da barra de pesquisa para alinhar verticalmente
    TopoPesquisa = Frm.TxtPesquisaUser.Top
    AlturaPesquisa = Frm.TxtPesquisaUser.Height
    
    With Frm.BtnNovoUsuario
        ' 1. ALINHAMENTO HORIZONTAL (A M·gica da Ancoragem)
        ' Left da Grid + Largura da Grid = Fim da Grid na direita.
        ' SubtraÌmos a largura do bot„o para ele n„o passar para fora.
        .Left = (Frm.ListUsuarios.Left + Frm.ListUsuarios.Width) - .Width
        
        ' 2. ALINHAMENTO VERTICAL (CentralizaÁ„o fina)
        ' Se o bot„o for mais gordinho que a barra de pesquisa, ele sobe uns pixels para ficar no centro exato do eixo Y
        .Top = TopoPesquisa - ((.Height - AlturaPesquisa) / 2)
        
        ' Traz para a frente para n„o ser engolido por nenhum outro controle
        .ZOrder 0
    End With
    
    On Error GoTo 0
End Sub
