Attribute VB_Name = "Mdl_GitTools"
Option Explicit

''' <summary>Exporta todos os componentes do VBA para uma pasta 'src' para controle de versão (Git).</summary>
Public Sub ExportarProjetoParaGit()
    Dim vbc As Object
    Dim folderPath As String
    Dim fileName As String
    Dim extension As String
    Dim fso As Object
    Dim subFolder As String
    
    ' 1. Define o caminho da pasta raiz (Mesmo local da planilha)
    folderPath = ThisWorkbook.Path & "\src"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 2. Cria a pasta 'src' se não existir
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder (folderPath)
    
    Debug.Print "--- Iniciando Exportação para: " & folderPath & " ---"
    
    ' 3. Varre todos os componentes do projeto
    For Each vbc In ThisWorkbook.VBProject.VBComponents
        
        ' Define a extensão e subpasta baseada no tipo de componente
        Select Case vbc.Type
            Case 1 ' msoext_ct_StdModule (Módulo Padrão)
                extension = ".bas"
                subFolder = "\Modules"
            Case 2 ' msoext_ct_ClassModule (Módulo de Classe)
                extension = ".cls"
                subFolder = "\Classes"
            Case 3 ' msoext_ct_MSForm (UserForm)
                extension = ".frm"
                subFolder = "\Forms"
            Case 100 ' msoext_ct_Document (Planilhas e ThisWorkbook)
                extension = ".cls"
                subFolder = "\Documents"
            Case Else
                extension = ".txt"
                subFolder = "\Others"
        End Select
        
        ' Cria a subpasta se não existir
        If Not fso.FolderExists(folderPath & subFolder) Then
            fso.CreateFolder (folderPath & subFolder)
        End If
        
        ' Caminho final do arquivo
        fileName = folderPath & subFolder & "\" & vbc.Name & extension
        
        ' Exporta o componente
        On Error Resume Next
        vbc.Export fileName
        If Err.Number <> 0 Then
            Debug.Print "Erro ao exportar: " & vbc.Name & " - " & Err.Description
            Err.Clear
        Else
            Debug.Print "Exportado: " & vbc.Name & extension
        End If
        On Error GoTo 0
    Next vbc
    
    MsgBox "Projeto exportado com sucesso para a pasta 'src'!" & vbCrLf & _
           "Agora você pode subir os arquivos para o GitHub.", vbInformation, "LogicUp Solutions - Git Tool"
End Sub
