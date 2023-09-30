VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Informa��es para c�lculos"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6615
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Eu adicionei a soma de vendas no final do bloco de seminovos mas deixei oculto
'tbm adicionei a soma de capta��o de cada setor no fim de cada coluna e tbm deixei oculto
'essas somas servem para facilitar o trabalho dessa macro

Private Sub CommandButton1_Click()
Dim cc As String
Dim i As Integer, ui As Long
Dim mes As Integer, ano As Integer
Application.ScreenUpdating = False


cc = Range("A4")
ui = 94
ano = Range("b4")
mes = Range("c4")
'LIMPAR
Range("C89:C" & ui).ClearContents
Range("E89:E" & ui).ClearContents
Range("I89:I" & ui).ClearContents

'---E20
If CheckE20 Then
    Range("A4") = "E20"
    For i = 89 To ui
        If Range("A" & i) <> "" Then
            Range("C" & i) = Range("C" & i) + Range("C65") 'C65 = total de vendas USADOS
            Range("E" & i) = Range("E" & i) + Range("E39") 'SOMAR CAPTA��O NOVOS
            Range("E" & i) = Range("E" & i) + Range("E65") 'SOMAR CAPTA��O USADOS
            Range("E" & i) = Range("E" & i) + Range("E83") 'SOMAR CAPTA��O VD
            Range("I" & i) = Range("I" & i) + Range("I65") '% da margem de seminovos
        End If
    Next i
End If
'-----N53
If CheckN53 Then
    Range("A4") = "N53"
    For i = 89 To ui
        If Range("A" & i) <> "" Then
            Range("C" & i) = Range("C" & i) + Range("C65") 'C65 = total de vendas USADOS
            Range("E" & i) = Range("E" & i) + Range("E39") 'SOMAR CAPTA��O NOVOS
            Range("E" & i) = Range("E" & i) + Range("E65") 'SOMAR CAPTA��O USADOS
            Range("E" & i) = Range("E" & i) + Range("E83") 'SOMAR CAPTA��O VD
            Range("I" & i) = Range("I" & i) + Range("I66") '% da margem de seminovos
        End If
    Next i
End If
'-----S46
If CheckS46 Then
    Range("A4") = "S46"
    For i = 89 To ui
        If Range("A" & i) <> "" Then
            Range("C" & i) = Range("C" & i) + Range("C65") 'C65 = total de vendas USADOS
            Range("E" & i) = Range("E" & i) + Range("E39") 'SOMAR CAPTA��O NOVOS
            Range("E" & i) = Range("E" & i) + Range("E65") 'SOMAR CAPTA��O USADOS
            Range("E" & i) = Range("E" & i) + Range("E83") 'SOMAR CAPTA��O VD
            Range("I" & i) = Range("I" & i) + Range("I65") '% da margem de seminovos
        End If
    Next i
End If
'-----T08
If CheckT08 Then
    Range("A4") = "T08"
    For i = 89 To ui
        If Range("A" & i) <> "" Then
            Range("C" & i) = Range("C" & i) + Range("C65") 'C65 = total de vendas USADOS
            Range("E" & i) = Range("E" & i) + Range("E39") 'SOMAR CAPTA��O NOVOS
            Range("E" & i) = Range("E" & i) + Range("E65") 'SOMAR CAPTA��O USADOS
            Range("E" & i) = Range("E" & i) + Range("E83") 'SOMAR CAPTA��O VD
            Range("I" & i) = Range("I" & i) + Range("I65") '% da margem de seminovos
        End If
    Next i
End If

'Adicionar f�rias

For i = 89 To ui
    If Range("Z" & i) <> "N�o" Then 'Se + Periodo for diferente de "N�o", fazer todo calculos ap�nas na linha atual
                                     'para cada empresa selecionada
    
    Range("C4") = Range("Z" & i).Value 'adicionar m�s de f�rias (Z � coluna de m�s)
    Range("B4") = Range("AA" & i).Value 'adicionar ano de f�rias (AA � coluna de ano)
'-------------------------------------------------------------------------
        If CheckE20 Then
            Range("A4") = "E20"
            If Range("A" & i) <> "" Then
                Range("C" & i) = Range("C" & i) + Range("C65") 'C65 = total de vendas USADOS
                Range("E" & i) = Range("E" & i) + Range("E39") 'SOMAR CAPTA��O NOVOS
                Range("E" & i) = Range("E" & i) + Range("E65") 'SOMAR CAPTA��O USADOS
                Range("E" & i) = Range("E" & i) + Range("E83") 'SOMAR CAPTA��O VD
                Range("I" & i) = Range("I" & i) + Range("I65") '% da margem de seminovos
            End If
        End If
        
        If CheckN53 Then
            Range("A4") = "N53"
            If Range("A" & i) <> "" Then
                Range("C" & i) = Range("C" & i) + Range("C65") 'C65 = total de vendas USADOS
                Range("E" & i) = Range("E" & i) + Range("E39") 'SOMAR CAPTA��O NOVOS
                Range("E" & i) = Range("E" & i) + Range("E65") 'SOMAR CAPTA��O USADOS
                Range("E" & i) = Range("E" & i) + Range("E83") 'SOMAR CAPTA��O VD
                Range("I" & i) = Range("I" & i) + Range("I65") '% da margem de seminovos
            End If
        End If
        
        If CheckS46 Then
            Range("A4") = "S46"
            If Range("A" & i) <> "" Then
                Range("C" & i) = Range("C" & i) + Range("C65") 'C65 = total de vendas USADOS
                Range("E" & i) = Range("E" & i) + Range("E39") 'SOMAR CAPTA��O NOVOS
                Range("E" & i) = Range("E" & i) + Range("E65") 'SOMAR CAPTA��O USADOS
                Range("E" & i) = Range("E" & i) + Range("E83") 'SOMAR CAPTA��O VD
                Range("I" & i) = Range("I" & i) + Range("I65") '% da margem de seminovos
            End If
        End If
        
        If CheckT08 Then
            Range("A4") = "T08"
            If Range("A" & i) <> "" Then
                Range("C" & i) = Range("C" & i) + Range("C65") 'C65 = total de vendas USADOS
                Range("E" & i) = Range("E" & i) + Range("E39") 'SOMAR CAPTA��O NOVOS
                Range("E" & i) = Range("E" & i) + Range("E65") 'SOMAR CAPTA��O USADOS
                Range("E" & i) = Range("E" & i) + Range("E83") 'SOMAR CAPTA��O VD
                Range("I" & i) = Range("I" & i) + Range("I65") '% da margem de seminovos
            End If
        End If
'-----------------------------------------------------------------------------------------------
        
    End If
Next i

Range("A4") = cc
Range("c4") = mes
Range("b4") = ano

Application.ScreenUpdating = True
End Sub

Private Sub CommandButton2_Click()
Unload Me
ActiveSheet.Protect
End Sub

Private Sub CommandButton3_Click()
Range("C89:C94").ClearContents
Range("E89:E94").ClearContents
Range("I89:I94").ClearContents
End Sub


Private Sub UserForm_Initialize()
ActiveSheet.Unprotect
End Sub
