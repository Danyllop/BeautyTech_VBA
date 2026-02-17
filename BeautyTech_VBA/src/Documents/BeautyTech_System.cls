VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BeautyTech_System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
'    Dim AppWasVisible As Boolean
'
'    'Armazena o estado anterior para decisão (boas práticas)
'    AppWasVisible = Application.Visible
'
'    On Error GoTo TrataErro
'
'    'Otimização de tela para evitar "piscas"
'    Application.ScreenUpdating = False
'
'    'LÓGICA DE OCULTAÇÃO INTELIGENTE
'    If Workbooks.Count > 1 Then
'        'Se houver outras planilhas, ocultamos APENAS a janela deste projeto
'        Windows(ThisWorkbook.Name).Visible = False
'    Else
'        'Se for a única, ocultamos a aplicação inteira (efeito sistema)
'        Application.Visible = False
'    End If
'
'    'RESTAURA ATUALIZAÇÃO DE TELA (necessário para o Form desenhar corretamente)
'    Application.ScreenUpdating = True
'
'    'CHAMADA DO FORMULÁRIO (MODAL É OBRIGATÓRIO)
'    'Isso pausa o código aqui até o Unload do Form ou Me.Hide
'    Usf_Login.Show vbModal
'
'    '--- O CÓDIGO SÓ CONTINUA AQUI APÓS O LOGIN/FECHAMENTO DO FORM ---
'
'    'Verificação Pós-Login (Opcional: varia conforme sua lógica no botão logar)
'    'Normalmente, o botão de login define uma variável pública ou Tag
'
'    Exit Sub
'
'TrataErro:
'    'Safety Net: Em caso de qualquer falha catastrófica, DEVOLVA o Excel ao usuário
'    Application.Visible = True
'    Windows(ThisWorkbook.Name).Visible = True
'    MsgBox "Ocorreu um erro crítico ao iniciar o sistema: " & Err.Description, vbCritical, "Erro de Inicialização"
End Sub
