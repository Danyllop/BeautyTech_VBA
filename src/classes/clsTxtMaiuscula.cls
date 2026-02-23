VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTxtMaiuscula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents txtGroup As MSForms.TextBox
Attribute txtGroup.VB_VarHelpID = -1

' Ao digitar, converte para maiúsculo
Private Sub txtGroup_Change()
    ' Evita loop infinito desativando eventos ou verificando se já está maiúsculo
    Dim pos As Long
    If txtGroup.Value <> UCase(txtGroup.Value) Then
        pos = txtGroup.SelStart
        txtGroup.Value = UCase(txtGroup.Value)
        txtGroup.SelStart = pos
    End If
End Sub
