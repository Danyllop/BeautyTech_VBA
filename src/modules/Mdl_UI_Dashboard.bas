Attribute VB_Name = "Mdl_UI_Dashboard"
Option Explicit

'' ==============================================================================
'' Modulo: Mdl_UI_Dashboard
'' Objetivo: Construir e Gerenciar a Interface SaaS (Com Icones, Interatividade e Widgets)
'' ==============================================================================
'
'' Cores do Tema
'Private Const COR_SIDEBAR As Long = 4010017    ' Azul Escuro (RGB 33, 47, 61)
'Private Const COR_HEADER As Long = 16777215    ' Branco
'Private Const COR_BACKGROUND As Long = 15790320 ' Cinza Claro (RGB 240, 240, 240)
'Private Const COR_TEXTO_MENU As Long = 14474460 ' Cinza claro
'Private Const COR_TEXTO_HOVER As Long = 16777215 ' Branco
'Private Const COR_ICONE As Long = 16777215      ' Branco
'
'' Dimensoes e Fontes
'Private Const LARGURA_SIDEBAR As Double = 220
'Private Const ALTURA_HEADER As Double = 60
'Private Const FONTE_ICONES As String = "Segoe MDL2 Assets"
'
'' Colecao para manter as classes de botoes vivas
'Private colBotoes As Collection
'
'' ==============================================================================
'' 1. Configurar Layout Inicial
'' ==============================================================================
'Public Sub SetupDashboard(Form As Object)
'    On Error Resume Next
'
'    Set colBotoes = New Collection
'
'    ' Configura Form
'    With Form
'        .BackColor = COR_BACKGROUND
'        .Caption = "LogicUp SolutionsBeautyTech"
'        .ScrollBars = 0
'    End With
'
'    ' --- Camada 1: Estrutura Base ---
'
'    ' Sidebar
'    Dim fraSide As Object
'    Set fraSide = Form.Controls.Add("Forms.Frame.1", "fraSidebar", True)
'    With fraSide
'        .BackColor = COR_SIDEBAR
'        .BorderStyle = 0
'        .SpecialEffect = 0 ' Flat (Sem borda 3D)
'        .Width = LARGURA_SIDEBAR
'        .Left = 0
'        .Top = 0
'    End With
'
'    ' Header
'    Dim fraHead As Object
'    Set fraHead = Form.Controls.Add("Forms.Frame.1", "fraHeader", True)
'    With fraHead
'        .BackColor = COR_HEADER
'        .BorderStyle = 0
'        .SpecialEffect = 0 ' Flat (Sem borda 3D)
'        .Height = ALTURA_HEADER
'        .Top = 0
'        .Left = LARGURA_SIDEBAR
'    End With
'
'    ' MultiPage (Conteudo)
'    Dim mpContent As Object
'    Set mpContent = Form.Controls.Add("Forms.MultiPage.1", "mpContent", True)
'    With mpContent
'        .Style = 2 ' Sem abas
'        .Top = ALTURA_HEADER + 20
'        .Left = LARGURA_SIDEBAR + 20
'        .BackColor = COR_BACKGROUND
'    End With
'
'    ' Cria Paginas
'    Dim pgDash As Object
'    Set pgDash = mpContent.Pages.Add("pgDashboard")
'    mpContent.Pages.Add "pgAgenda"
'    mpContent.Pages.Add "pgClientes"
'    mpContent.Pages.Add "pgServicos"
'    mpContent.Pages.Add "pgFinanceiro"
'
'    ' --- Camada 2: Elementos Visuais ---
'
'    ' Logo Area
'    Dim strLogoPath As String
'    Dim caminhoBase As String
'    caminhoBase = ThisWorkbook.Path & "\assets\"
'
'    ' Tenta encontrar formatos compativeis com VBA (JPG ou BMP)
'    ' PNG muitas vezes falha no LoadPicture padrao
'    If Dir(caminhoBase & "BeautyTech_Logo.jpg") <> "" Then
'        strLogoPath = caminhoBase & "BeautyTech_Logo.jpg"
'    ElseIf Dir(caminhoBase & "BeautyTech_Logo.bmp") <> "" Then
'        strLogoPath = caminhoBase & "BeautyTech_Logo.bmp"
'    Else
'        strLogoPath = ""
'    End If
'
'    If strLogoPath <> "" Then
'        ' Se a imagem existir, usa ela
'        Dim imgLogo As Object
'        Set imgLogo = fraSide.Controls.Add("Forms.Image.1", "imgLogo", True)
'        With imgLogo
'            .Picture = LoadPicture(strLogoPath)
'            .PictureSizeMode = 3 ' fmPictureSizeModeZoom
'            .BackStyle = 0 ' Transparente
'            .BorderStyle = 0
'            .Top = 20
'            .Left = 20
'            .Width = 180
'            .Height = 60
'        End With
'    Else
'        ' Fallback para texto/icone se imagem nao existir ou for incompativel
'        CreateLabel fraSide, "lblLogoIcon", ChrW(&HEC06), FONTE_ICONES, 24, COR_ICONE, 25, 25, 40, 40
'        CreateLabel fraSide, "lblLogoTxt", "BeautyTech", "Segoe UI", 16, COR_ICONE, 25, 70, 150, 30, True
'        CreateLabel fraSide, "lblSubLogo", "MVP Management", "Segoe UI", 9, COR_TEXTO_MENU, 55, 72, 120, 20
'    End If
'
'    ' Header Title
'    CreateLabel fraHead, "lblPageTitle", "Dashboard", "Segoe UI", 18, COR_SIDEBAR, 15, 30, 300, 35, True
'
'    ' --- Camada 3: Botoes do Menu ---
'    Dim TopoMenu As Double: TopoMenu = 120
'
'    CriarBotaoMenu Form, fraSide, "dash", "Dashboard", ChrW(&HE80F), 0, TopoMenu
'    CriarBotaoMenu Form, fraSide, "agend", "Agenda", ChrW(&HE787), 1, TopoMenu + 50
'    CriarBotaoMenu Form, fraSide, "cli", "Clientes", ChrW(&HE77B), 2, TopoMenu + 100
'    CriarBotaoMenu Form, fraSide, "serv", "Servicos", ChrW(&HE79C), 3, TopoMenu + 150
'    CriarBotaoMenu Form, fraSide, "fin", "Financeiro", ChrW(&HE8C7), 4, TopoMenu + 200
'
'    ' --- Camada 4: Widgets do Dashboard (Migrado do Modulo Antigo) ---
'    ' Cria Cards na pagina pgDashboard
'    ' Top, Left, Width, Height, Icone, Valor, Titulo, CorBarra
'
'    CriarCard pgDash, "cardAgend", 20, 0, 250, 100, ChrW(&HE787), "12", "Agendamentos Hoje", 16750848 ' Azul
'    CriarCard pgDash, "cardFat", 20, 270, 250, 100, ChrW(&HE8C7), "R$ 450", "Faturamento Dia", 5025616 ' Verde
'    CriarCard pgDash, "cardCli", 20, 540, 250, 100, ChrW(&HE77B), "3", "Novos Clientes", 10040319 ' Roxo
'
'End Sub
'
'' ==============================================================================
'' 2. Logica de Redimensionamento Responsivo
'' ==============================================================================
'Public Sub ResizeDashboard(Form As Object)
'    On Error Resume Next
'    Dim W As Double, H As Double
'    W = Form.InsideWidth
'    H = Form.InsideHeight
'
'    If W < 100 Then Exit Sub
'
'    ' Sidebar preenche toda a altura
'    With Form.Controls("fraSidebar")
'        .Height = H + 20 ' Garante que passe um pouco para evitar borda inferior branca
'        .Top = 0
'    End With
'
'    ' Header preenche a largura restante
'    With Form.Controls("fraHeader")
'        .Width = W - LARGURA_SIDEBAR + 20 ' Garante que passe um pouco para evitar borda direita branca
'        .Left = LARGURA_SIDEBAR
'    End With
'
'    Dim LarguraConteudo As Double
'    Dim MargemEsquerda As Double
'
'    ' Centraliza conteudo se a tela for muito larga
'    If (W - LARGURA_SIDEBAR) > 1250 Then
'        LarguraConteudo = 1200
'        MargemEsquerda = LARGURA_SIDEBAR + ((W - LARGURA_SIDEBAR - 1200) / 2)
'    Else
'        LarguraConteudo = W - LARGURA_SIDEBAR - 40
'        MargemEsquerda = LARGURA_SIDEBAR + 20
'    End If
'
'    ' Ajusta area de conteudo
'    With Form.Controls("mpContent")
'        .Top = ALTURA_HEADER + 20
'        .Height = H - ALTURA_HEADER - 40
'        .Width = LarguraConteudo
'        .Left = MargemEsquerda
'    End With
'
'    ' Reposiciona widgets se necessario (Ex: se a largura permitir 4 colunas, etc)
'    ' Por enquanto mantemos fixo, mas a logica poderia ser expandida aqui.
'
'End Sub
'
'' ==============================================================================
'' 3. Auxiliares e Components (Cards/Botoes)
'' ==============================================================================
'
'Private Sub CreateLabel(Parent As Object, Name As String, Caption As String, FontName As String, FontSize As Double, Color As Long, Top As Double, Left As Double, Width As Double, Height As Double, Optional Bold As Boolean = False)
'    Dim lbl As Object
'    Set lbl = Parent.Controls.Add("Forms.Label.1", Name, True)
'    With lbl
'        .Caption = Caption
'        .Font.Name = FontName
'        .Font.Size = FontSize
'        .Font.Bold = Bold
'        .ForeColor = Color
'        .BackStyle = 0
'        .Top = Top
'        .Left = Left
'        .Width = Width
'        .Height = Height
'    End With
'End Sub
'
'Private Sub CriarBotaoMenu(Form As Object, Parent As Object, Suffix As String, Caption As String, IconChar As String, PageIndex As Integer, Top As Double)
'    Dim bg As Object, ico As Object, txt As Object
'
'    ' Fundo do botao (Hover area)
'    Set bg = Parent.Controls.Add("Forms.Label.1", "btnBg_" & Suffix, True)
'    With bg
'        .Caption = ""
'        .BackColor = COR_SIDEBAR
'        .Top = Top
'        .Left = 0
'        .Width = LARGURA_SIDEBAR
'        .Height = 45
'        .Tag = PageIndex
'    End With
'
'    ' Icone
'    Set ico = Parent.Controls.Add("Forms.Label.1", "btnIco_" & Suffix, True)
'    With ico
'        .Caption = IconChar
'        .Font.Name = FONTE_ICONES
'        .Font.Size = 14
'        .ForeColor = COR_TEXTO_MENU
'        .BackStyle = 0
'        .Top = Top + 12
'        .Left = 25
'        .Width = 30
'        .Height = 30
'    End With
'
'    ' Texto
'    Set txt = Parent.Controls.Add("Forms.Label.1", "btnTxt_" & Suffix, True)
'    With txt
'        .Caption = Caption
'        .Font.Name = "Segoe UI"
'        .Font.Size = 11
'        .ForeColor = COR_TEXTO_MENU
'        .BackStyle = 0
'        .Top = Top + 10
'        .Left = 60
'        .Width = LARGURA_SIDEBAR - 70
'        .Height = 25
'    End With
'
'    ' Registra na colecao para eventos
'    Dim cls As clsBotaoMenu
'    Set cls = New clsBotaoMenu
'    Set cls.BgLabel = bg
'    Set cls.IconLabel = ico
'    Set cls.TxtLabel = txt
'    Set cls.ParentForm = Form
'    cls.IndexPagina = PageIndex
'    colBotoes.Add cls
'End Sub
'
'Public Sub ResetarBotoesMenu()
'    Dim item As clsBotaoMenu
'    For Each item In colBotoes
'        item.ResetarCor
'    Next item
'End Sub
'
'' --- Novo: Componente Card (Widget) ---
'Private Sub CriarCard(Parent As Object, Name As String, Top As Double, Left As Double, Width As Double, Height As Double, IconChar As String, Valor As String, Titulo As String, CorBarra As Long)
'    ' Frame do Card
'    Dim fr As Object
'    Set fr = Parent.Controls.Add("Forms.Frame.1", Name, True)
'    With fr
'        .Top = Top
'        .Left = Left
'        .Width = Width
'        .Height = Height
'        .BackColor = 16777215  ' Branco
'        .BorderStyle = 0
'        .SpecialEffect = 2 ' Etched effect para dar profundidade ao card (manter aqui, remover so da estrutura principal)
'    End With
'
'    ' Barra Colorida Lateral
'    Dim bar As Object
'    Set bar = fr.Controls.Add("Forms.Label.1", Name & "_Bar", True)
'    With bar
'        .Top = 0
'        .Left = 0
'        .Width = 5
'        .Height = Height
'        .BackColor = CorBarra
'    End With
'
'    ' Icone Grande (Fundo)
'    Dim ico As Object
'    Set ico = fr.Controls.Add("Forms.Label.1", Name & "_Ico", True)
'    With ico
'        .Caption = IconChar
'        .Font.Name = FONTE_ICONES
'        .Font.Size = 32
'        .ForeColor = 15790320 ' Cinza bem claro
'        .BackStyle = 0
'        .Top = 15
'        .Left = Width - 60
'        .Width = 50
'        .Height = 50
'    End With
'
'    ' Valor (Grande)
'    Dim val As Object
'    Set val = fr.Controls.Add("Forms.Label.1", Name & "_Val", True)
'    With val
'        .Caption = Valor
'        .Font.Name = "Segoe UI"
'        .Font.Size = 22
'        .Font.Bold = True
'        .ForeColor = COR_SIDEBAR
'        .BackStyle = 0
'        .Top = 15
'        .Left = 20
'        .Width = Width - 70
'        .Height = 35
'    End With
'
'    ' Titulo (Pequeno)
'    Dim tit As Object
'    Set tit = fr.Controls.Add("Forms.Label.1", Name & "_Tit", True)
'    With tit
'        .Caption = Titulo
'        .Font.Name = "Segoe UI"
'        .Font.Size = 10
'        .ForeColor = 8421504 ' Cinza medio
'        .BackStyle = 0
'        .Top = 55
'        .Left = 20
'        .Width = Width - 30
'        .Height = 20
'    End With
'
'End Sub



