Attribute VB_Name = "Enviar_Whatsapp"
Sub enviar()
Dim opc As Integer
Dim strURL As String, texto As String
Dim num As String, i As Integer, iNum As Integer, col As String
Dim key As String
opc = MsgBox("Posso iniciar o envio?", vbYesNo, "")
If opc = vbNo Then
    MsgBox "Ok, saindo.", , ""
    Exit Sub
End If

fEspere 1
ui = Planilha10.Range("L1000").End(xlUp).Row
With Application
    'ABRINDO NAVEGADOR
    .SendKeys "^{ESC}"
    fEspere (1)
    .SendKeys "Google"
    fEspere (1)
    .SendKeys "~"
    fEspere (5)

    For i = 8 To 100 'PASSAR POR CADA LINHA DA PLANILHA CALCULO
        Planilha10.Range("a1").Value = i
        key = i & Planilha10.Range("E1").Value
        If Not Planilha10.Range("L1:L1000").Find(key, , , xlWhole) Is Nothing And Planilha10.Range("C1").Value <> "Nada encontrado!" Then
            iNum = Planilha10.Range("L1:L1000").Find(key, , , xlWhole).Row 'PROCURA NUMERO DA LINHA NO CADASTRO DE CONTATOS
            num = Planilha10.Range("M" & iNum).Value
            
            If num <> "" Then 'VERIFICA SE O NÚMERO É DIFERENTE DE VAZIO
                texto = Planilha10.Range("F12").Value
                strURL = "https://wa.me/" & num & " "
                .SendKeys strURL, True 'ESCREVER URL
                fEspere (2)
                .SendKeys "~" 'DÁ ENTER NA URL
                fEspere (5)
                
                'Mid(texto, posição_inicial, comprimento)
                For ii = 1 To Len(texto)
                    digito = Mid(texto, ii, 1)
                    .SendKeys digito, True
                Next ii
                fEspere 2
                .SendKeys "~" 'ENVIAR MENSAGEM
                fEspere (2)
                Planilha10.Range("A1:C30").Copy
                .SendKeys "^V" 'COLAR IMAGEM
                fEspere (1)
                .SendKeys "^+{HOME}" 'SELECIONAR TEXTO QUE VEM COM A IMAGEM
                fEspere (1)
                .SendKeys "{DEL}" 'EXCLUIR TEXTO
                fEspere (1)
                .SendKeys "~" 'ENVIAR IMAGEM
                fEspere (1)
                .SendKeys "^+{HOME}" 'SELECIONAR TEXTO QUE VEM COM A IMAGEM
                fEspere (1)
                .SendKeys "{DEL}" 'EXCLUIR TEXTO
                fEspere (1)
                .SendKeys "%{TAB}" 'VOLTAR AO NAVEGADOR
                fEspere (1)
                .SendKeys "{F6}" 'SELECIONAR BARRA DE ENDEREÇO
                fEspere (1)
                .SendKeys "{DEL}"
                fEspere (1)
                .CutCopyMode = False
            End If
        End If
    Next i
    .SendKeys "^W"
    
End With


End Sub

Function fEspere(Segundos As Integer)
'MsgBox "AGUARDAR AGORA"
Application.Wait Now + TimeValue(Format(Segundos, "00:00:00"))
'MsgBox "PRONTO"
End Function
