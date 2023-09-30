Attribute VB_Name = "ComissãoDemaisFunções"
Public ccGlobal As String
Public setorGlobal As String

'--------------------------------------COMISSÃO DIGITAL POR VENDA DE CARRO
Sub comissDigitalAcelerador()
Dim i As Long, ui As Long, icad As Long, iRH As Long, ii As Long
Dim nome As String, nomeAnterior As String
Dim tabelaDigital As String, colunaNome As String, colunaMatricula As String, colunaComissao As String


If setorGlobal = "NV" Then
    tabelaDigital = "_Varejo"
    ui = baseDigital.Range("C1000000").End(xlUp).Row
    colunaNome = "F"
    colunaMatricula = "E"
    colunaComissao = "L"
    
ElseIf setorGlobal = "SN" Then
    tabelaDigital = "_Usados"
    ui = baseDigital.Range("R1000000").End(xlUp).Row
    colunaNome = "U"
    colunaMatricula = "T"
    colunaComissao = "Z"
    
End If
On Error GoTo ErroClassificacao
baseDigital.Select '------ Classificar MACRO DE GRAVAÇÃO
ActiveWorkbook.Worksheets("Base Digital").ListObjects("Dealer_Calc_Comissao" & tabelaDigital). _
        Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Base Digital").ListObjects("Dealer_Calc_Comissao" & tabelaDigital). _
        Sort.SortFields.Add2 key:=Range("Dealer_Calc_Comissao" & tabelaDigital & "[[#All],[Nome Dealer]]") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Base Digital").ListObjects( _
        "Dealer_Calc_Comissao" & tabelaDigital).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
Planilha7.Select

'-------------------------------------------------------
For i = 2 To ui
    iRH = Range("A1000000").End(xlUp).Offset(1, 0).Row 'linha livre no RH
    icad = 0
    nome = baseDigital.Range(colunaNome & i).Value 'Coluna do nome do Dealer
    On Error Resume Next 'Procurar nome no cadastro
    icad = CadVendedores.Range("A1:A106").Find(nome, , , xlWhole).Row
    If nomeAnterior = CadVendedores.Range("A" & icad) Then 'Evitar somar comissão do mesmo funcionario
        GoTo pularNome
    Else
        If icad <> 0 Then
            If ccGlobal = CadVendedores.Range("D" & icad) Then 'verificar se o funcionário é da CC selecionada
            '------------ FAZER GANHAR PELAS 2 CC (E20 e N53)----------------------------
                For cont = 1 To 2
                    If cont = 1 Then
                        Planilha1.Range("A4") = "E20"
                    Else
                        Planilha1.Range("A4") = "N53"
                    End If
                    Range("A" & iRH) = nome
                    Range("B" & iRH) = baseDigital.Range(colunaMatricula & i).Value 'Matricula
                    Range("C" & iRH) = CadVendedores.Range("G" & icad).Value 'Função
                    ii = i
                    Do While baseDigital.Range(colunaComissao & ii) <> ""
                        If baseDigital.Range(colunaNome & ii) = nome Then
                        Range("D" & iRH) = Range("D" & iRH) + baseDigital.Range(colunaComissao & ii).Value 'Coluna de comissão
                        ii = ii + 1
                        Else
                            ii = ii + 1
                        End If
                        
                    Loop
                Next cont
                nomeAnterior = nome
                
            '------------ FIM DE FAZER GANHAR PELAS 2 CC ----------------------------
            End If
        End If
    End If
pularNome:
    On Error GoTo -1
    On Error GoTo 0
Next i
Planilha1.Range("A4") = ccGlobal

Exit Sub
ErroClassificacao:
    MsgBox "Erro ao tentar classificar nomes dos funcionários na planilha 'Base Digital'", vbInformation, ""
End Sub

'--------------------------------------------------------------------------ACESSÓRIOS---------------------------------------
'Adicionar comissão de ACESSÓRIOS de que não for vendedor
'O funcionário deve está cadastrado no setor, função e CC correto
Sub comissAcessoriosDemaisFuncoes()
Dim iCadastro As Long, uiCadastro As Long, iAcess As Long, iRH As Long
Dim nomeFuncionario As String
Dim vlrComissao As Double

uiCadastro = 105 'Ultima linha na planilha cadastro
For iCadastro = 3 To uiCadastro
    'Verificar se o setor, CC e Função são válidas para iniciar o procedimento
    If (CadVendedores.Range("A" & iCadastro) <> Empty And CadVendedores.Range("D" & iCadastro) = ccGlobal And _
    CadVendedores.Range("E" & iCadastro) = setorGlobal And CadVendedores.Range("G" & iCadastro) <> "Vendedor" And CadVendedores.Range("G" & iCadastro) <> "Vend Master") Then
        
        nomeFuncionario = CadVendedores.Range("A" & iCadastro) 'Pegando nome do funcionário
        iAcess = 0
        iRH = 0
        On Error Resume Next
        iAcess = Planilha9.Range("B:B").Find(nomeFuncionario, , , xlWhole).Row 'Procurando funcionario em acessórios
        On Error GoTo -1
        On Error GoTo 0
        If iAcess <> 0 Then 'Se iAcess <> 0 significa que tem venda do funcionário
            vlrComissao = Planilha9.Range("C" & iAcess).Value * 0.02 '------- 0.02 = 2% de comissão
            
            On Error Resume Next 'Verificar se o funcionário já está na folha do RH
            iRH = Planilha7.Range("A4:A64").Find(nomeFuncionario, , , xlWhole).Row
            On Error GoTo -1
            On Error GoTo 0
            If iRH = 0 Then 'Se iRH <> 0 o funcinário já está na folha do RH
                iRH = Planilha7.Range("A64").End(xlUp).Offset(1, 0).Row
                Planilha7.Range("A" & iRH) = nomeFuncionario 'Nome funcinoário
                Planilha7.Range("B" & iRH) = CadVendedores.Range("B" & iCadastro)  'Código
                Planilha7.Range("C" & iRH) = CadVendedores.Range("G" & iCadastro) 'Função
                Planilha7.Range("E" & iRH) = Planilha7.Range("E" & iRH) + vlrComissao 'Gratificação
            Else
                Planilha7.Range("E" & iRH) = Planilha7.Range("E" & iRH) + vlrComissao  'Gratificação
            End If
        End If
        
    End If

Next iCadastro

End Sub
