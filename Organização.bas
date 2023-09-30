Attribute VB_Name = "Organização"
'Ocutar ou exibir"Nada encontrado!" na planilha "Calculos"
Sub ocultar_Exibir()
Dim i As Integer, ui As Long
Application.ScreenUpdating = False
ActiveSheet.Unprotect
ui = 94 'Ultima linha a ser considerada
If Range("a1").Value = 1 Then
    For i = 8 To ui
        If Range("A" & i).Value = "Nada encontrado!" Then
            Range("A" & i).Rows.Hidden = True
        End If
    Next i
    Range("a1").Value = 0
Else
    For i = 8 To ui
       Range("A" & i).Rows.Hidden = False
    Next i
    Range("a1").Value = 1
End If
ActiveSheet.Protect
Application.ScreenUpdating = True
End Sub

Sub limpar()
Dim opc As Integer
opc = MsgBox("Deseja apagar todas as informações de venda de acessórios desta aba?", vbYesNo, "Limpar dados")
If opc = 6 Then
    ActiveSheet.Unprotect
    Range("B5:C1000").ClearContents
    Range("G5:H1000").ClearContents
    ActiveSheet.Protect
End If
End Sub

Sub AtualizarRH_Novos()
ActiveSheet.Unprotect
Planilha1.Select
Range(Cells(8, 1), Range("A8").End(xlDown)).Copy ' Nome dos vendedores
Planilha7.Select
Range("A4").PasteSpecial xlPasteValues
End Sub

'Esse modulo serve para remover plalavras lixo dos nomes dos vendedores
Sub corrigirFolhaRH()
Dim i As Long, novoNome As String, cont As Integer
Dim palavrasParaRemover As Variant

palavrasParaRemover = Array("Acessorios", "-") 'Palavras que desejo remover
On Error Resume Next
For cont = o To UBound(palavrasParaRemover) 'procurar cada palavra do Araay, usando cont como variável index
    i = 0
    Do While True
    i = Range("A:A").Find(palavrasParaRemover(cont)).Row
    If i = 0 Then
        Exit Do ' Se i = 0 siginifica que a palavra da vez não foi encontrada, então sair do loop
    Else
        novoNome = Replace(Range("A" & i), palavrasParaRemover(cont), "") 'Substituindo palavra da vez por ""
        Range("A" & i) = novoNome
        i = 0 'reiniciando variável para proxima procura
    End If
    Loop
Next cont

' Formatação como moeda
Columns("D:E").Style = "Currency"
End Sub
