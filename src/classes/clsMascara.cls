VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMascara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Define que esta classe responde a eventos de uma TextBox
Public WithEvents CampoTexto As MSForms.TextBox
Attribute CampoTexto.VB_VarHelpID = -1

' -----------------------------------------------------------
' Evento KeyPress: Onde a "mágica" da máscara acontece
' -----------------------------------------------------------
Private Sub CampoTexto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim strValidar As String
    
    ' Bloqueia caracteres que não sejam números (exceto para tags de letras)
    If Not (CampoTexto.Tag = "ll" Or CampoTexto.Tag = "el" Or CampoTexto.Tag = "ln") Then
        If CampoTexto.Tag <> "nv" Then ' nv permite vírgula
            If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
        End If
    End If

    Select Case Me.CampoTexto.Tag
        Case "dd" ' DATA (00/00/0000)
            CampoTexto.MaxLength = 10
            If Len(CampoTexto) = 2 Or Len(CampoTexto) = 5 Then CampoTexto = CampoTexto & "/"

        Case "hh" ' HORA (00:00)
            CampoTexto.MaxLength = 5
            If Len(CampoTexto) = 2 Then CampoTexto = CampoTexto & ":"

        Case "tt" ' TELEFONE ((00) 00000-0000)
            CampoTexto.MaxLength = 15
            If Len(CampoTexto) = 0 Then CampoTexto = "(" & CampoTexto
            If Len(CampoTexto) = 3 Then CampoTexto = CampoTexto & ") "
            If Len(CampoTexto) = 10 Then CampoTexto = CampoTexto & "-"

        Case "cpf" ' CPF (000.000.000-00)
            CampoTexto.MaxLength = 14
            If Len(CampoTexto) = 3 Or Len(CampoTexto) = 7 Then CampoTexto = CampoTexto & "."
            If Len(CampoTexto) = 11 Then CampoTexto = CampoTexto & "-"

        Case "cep" ' CEP (00000-000)
            CampoTexto.MaxLength = 9
            If Len(CampoTexto) = 5 Then CampoTexto = CampoTexto & "-"

        Case "nv" ' NÚMEROS COM VÍRGULA
            If KeyAscii = 44 Then ' Se for vírgula
                If InStr(CampoTexto, ",") > 0 Or CampoTexto = "" Then KeyAscii = 0
            ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
                KeyAscii = 0
            End If
            
        Case "ll" ' APENAS LETRAS (Sem acento)
            If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32) Then
                KeyAscii = 0
            End If
    End Select
End Sub

