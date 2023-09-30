VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Comissão RH"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8100
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim iSetor As Integer, cc As String, uiCalculo As Long, setor As String
Dim i As Integer, iRH As Integer, key As String, ii As Integer
Dim comissMaster As Double
Application.ScreenUpdating = False
ActiveSheet.Unprotect
Range("A5:F20").ClearContents

'--------------------------------------VERIFICAÇÃO PARA INICIO DO PROCESSO ---------------------------
'-----------------------Escolhendo Setor
If opcNovos Then
    iSetor = 8
    setor = "NV"
ElseIf opcUsados Then
    iSetor = 44
    setor = "SN"
ElseIf opcVD Then
    iSetor = 70
    setor = "VD"
ElseIf opcAvaliadores Then
    iSetor = 88
    setor = "AV"
Else
    MsgBox "Você precisa escolher um setor", vbExclamation, "Me ajude!"
    Exit Sub
End If
setorGlobal = setor
'-----------------------Escolhendo CC
If opcE20 Then
cc = "E20"
ElseIf opcN53 Then
cc = "N53"
ElseIf opcS46 Then
cc = "S46"
ElseIf opcT08 Then
cc = "T08"
Else
MsgBox "Preciso que você escolha uma concessinária", vbExclamation, "Me ajude!"
Exit Sub
End If
ccGlobal = cc
'--------------------------------------FIM DA VERIFICAÇÃO DE INICIALIZAÇÃO --------------------
Application.DisplayStatusBar = False

'----------------Atualizando RH
Planilha1.Select
Range("A4") = cc
If Range("A1") = 0 Then 'Ajustar visão dos Calculos
    Call ocultar_Exibir
    Call ocultar_Exibir
Else
    Call ocultar_Exibir
End If

Application.ScreenUpdating = False
If opcAvaliadores Then
    uiCalculo = 94
Else
    uiCalculo = Range("A" & iSetor).End(xlDown).Row
End If
'--------------Adicionar nome de vendedores de acordo com o cadastro
Planilha7.Select
Range("A2") = Planilha1.Range("E4") & " de " & Planilha1.Range("B4") & ". Setor: " & setor
iRH = 5
For i = iSetor + 1 To uiCalculo
    If Planilha1.Range("A" & i) <> "Nada encontrado!" Then
        key = Planilha1.Range("B" & i) & cc & setor
        For ii = 3 To 105
            If CadVendedores.Range("H" & ii).Value = key Then 'H é a coluna onde está a chave no cadastro dos vendedores
                Range("A" & iRH) = CadVendedores.Range("A" & ii) 'nome completo
                Range("B" & iRH) = CadVendedores.Range("B" & ii) 'Cód FUNCIONARIO
                Range("C" & iRH) = CadVendedores.Range("G" & ii) 'Função
                Range("D" & iRH) = Planilha1.Range("AA" & i) 'RH = coluna de (COMISSÃO) da planilha calculo
                Range("E" & iRH) = Planilha1.Range("X" & i) 'RH = coluna de(GRATIFICAÇÃO ISC)
                Range("F" & iRH) = Planilha1.Range("W" & i) 'RH = coluna de(GRATIFICAÇÃO SOBRE CAPTAÇÃO)
                
                iRH = iRH + 1
                Exit For
            End If
        Next ii
    End If
Next i

'--------------Adicinoar comissão de Vendedores Master
i = 0
On Error Resume Next
i = Cells.Find("Vend Master").Row
On Error GoTo -1
On Error GoTo 0
If i <> 0 Then 'Se i<>0 significa que tem vendedor Master na Folha do RH
'Calculando comissaoMaster
    comissMaster = 0
    i = 38 'Primeira linha da tabela de comissão Master
    key = cc & setor
    Do While Planilha2.Range("K" & i) <> ""
        If Planilha2.Range("k" & i) = key Then
            If Planilha2.Range("J" & i) >= 1.1 Then
                comissMaster = comissMaster + Planilha2.Range("F" & i)
            ElseIf Planilha2.Range("J" & i) >= 1 Then
                comissMaster = comissMaster + Planilha2.Range("E" & i)
            Else
                comissMaster = comissMaster + Planilha2.Range("D" & i)
            End If
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
    'Adicionando comissao Master
    For i = 5 To iRH
        If Range("C" & i) = "Vend Master" Then
            Range("D" & i) = Range("D" & i) + comissMaster
        End If
    Next i
End If
'--------------------------------------------------
'Add comissão das aceleradoras
If opcNovos Or opcUsados Then
    Call comissDigitalAcelerador
End If

Call comissAcessoriosDemaisFuncoes
Call corrigirFolhaRH

Application.CutCopyMode = False
Application.DisplayStatusBar = True
ActiveSheet.Protect
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton2_Click()
Unload Me
'ActiveSheet.Protect
End Sub

Private Sub txtMetaMB_AfterUpdate()
txtMetaMB = Format(txtMetaMB.Value, "R$ #,###.00")
End Sub


