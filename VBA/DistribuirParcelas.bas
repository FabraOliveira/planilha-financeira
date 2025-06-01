Attribute VB_Name = "Módulo1"
Sub DistribuirParcelas()
    Dim ws As Worksheet
    Dim wsMesAtual As Worksheet
    Dim mes As Integer
    Dim meses(1 To 24) As String

    ' Array de meses para 2025 (em português) e 2026 (em inglês)
    meses(1) = "Janeiro"
    meses(2) = "Fevereiro"
    meses(3) = "Março"
    meses(4) = "Abril"
    meses(5) = "Maio"
    meses(6) = "Junho"
    meses(7) = "Julho"
    meses(8) = "Agosto"
    meses(9) = "Setembro"
    meses(10) = "Outubro"
    meses(11) = "Novembro"
    meses(12) = "Dezembro"

    meses(13) = "January"
    meses(14) = "February"
    meses(15) = "March"
    meses(16) = "April"
    meses(17) = "May"
    meses(18) = "June"
    meses(19) = "July"
    meses(20) = "August"
    meses(21) = "September"
    meses(22) = "October"
    meses(23) = "November"
    meses(24) = "December"

    For mes = 1 To 24
        Set ws = ThisWorkbook.Sheets(meses(mes))
        ProcessarIntervalo ws, "D62:J1059", "C62:C1059", meses, mes
    Next mes
End Sub

Sub ProcessarIntervalo(ws As Worksheet, intervalo As String, intervaloAdiantada As String, meses() As String, mes As Integer)
    Dim wsMesAtual As Worksheet
    Dim linha As Integer, parcelaAtual As Long, totalParcelas As Long
    Dim nomeMes As String, i As Integer, proximaLinhaVazia As Long
    Dim linhaExistente As Range, produto As String, adiantada As String
    Dim valorParcela As Variant, valorTotalParcelas As Variant

    For linha = ws.Range(intervalo).Rows(1).Row To ws.Range(intervalo).Rows(ws.Range(intervalo).Rows.Count).Row
        If Not IsEmpty(ws.Cells(linha, ws.Range(intervalo).Columns(3).Column).Value) Then ' Nome do produto (coluna F)
            valorParcela = ws.Cells(linha, ws.Range(intervalo).Columns(6).Column).Value ' Parcela atual (coluna I)
            valorTotalParcelas = ws.Cells(linha, ws.Range(intervalo).Columns(7).Column).Value ' Total de parcelas (coluna J)

            If IsNumeric(valorParcela) And IsNumeric(valorTotalParcelas) Then
                parcelaAtual = CLng(valorParcela)
                totalParcelas = CLng(valorTotalParcelas)
                If totalParcelas > 1 Then
                    produto = ws.Cells(linha, ws.Range(intervalo).Columns(3).Column).Value ' Nome do produto (coluna F)
                    adiantada = ws.Cells(linha, ws.Range(intervaloAdiantada).Column).Value

                    If adiantada <> "Adiantada" Then
                        For i = 1 To totalParcelas - parcelaAtual
                            If (mes + i) <= 24 Then
                                nomeMes = meses(mes + i)
                                Set wsMesAtual = ThisWorkbook.Sheets(nomeMes)

                                If Not ParcelaJaAdiantada(mes, produto, parcelaAtual + i, meses, intervalo, intervaloAdiantada) Then
                                    Set linhaExistente = wsMesAtual.Range(wsMesAtual.Range(intervalo).Cells(1, 3), _
                                        wsMesAtual.Range(intervalo).Cells(wsMesAtual.Range(intervalo).Rows.Count, 3)).Find(produto, LookIn:=xlValues)

                                    If linhaExistente Is Nothing Then
                                        proximaLinhaVazia = EncontrarProximaLinhaVazia(wsMesAtual, intervalo)
                                        wsMesAtual.Cells(proximaLinhaVazia, wsMesAtual.Range(intervalo).Columns(1).Column).Resize(1, 7).Value = _
                                            ws.Cells(linha, ws.Range(intervalo).Columns(1).Column).Resize(1, 7).Value
                                        wsMesAtual.Cells(proximaLinhaVazia, wsMesAtual.Range(intervalo).Columns(6).Column).Value = parcelaAtual + i
                                        wsMesAtual.Cells(proximaLinhaVazia, wsMesAtual.Range(intervalo).Columns(7).Column).Value = totalParcelas
                                    End If
                                Else
                                    Call RemoverParcelaAdiantada(wsMesAtual, produto, parcelaAtual + i, intervalo)
                                End If
                            End If
                        Next i
                    End If
                End If
            End If
        End If
ProximaLinha:
    Next linha
End Sub

Function EncontrarProximaLinhaVazia(ws As Worksheet, intervalo As String) As Long
    Dim linhaAtual As Long
    For linhaAtual = ws.Range(intervalo).Rows(1).Row To ws.Range(intervalo).Rows(ws.Range(intervalo).Rows.Count).Row
        If IsEmpty(ws.Cells(linhaAtual, ws.Range(intervalo).Columns(3).Column).Value) Then
            EncontrarProximaLinhaVazia = linhaAtual
            Exit Function
        End If
    Next linhaAtual
    EncontrarProximaLinhaVazia = ws.Range(intervalo).Rows(ws.Range(intervalo).Rows.Count).Row + 1
End Function

Function ParcelaJaAdiantada(mesAtual As Integer, produto As String, parcela As Integer, meses() As String, intervalo As String, intervaloAdiantada As String) As Boolean
    Dim wsAnterior As Worksheet, linha As Integer
    For m = 1 To mesAtual
        Set wsAnterior = ThisWorkbook.Sheets(meses(m))
        For linha = wsAnterior.Range(intervalo).Rows(1).Row To wsAnterior.Range(intervalo).Rows(wsAnterior.Range(intervalo).Rows.Count).Row
            If wsAnterior.Cells(linha, wsAnterior.Range(intervalo).Columns(3).Column).Value = produto And _
               wsAnterior.Cells(linha, wsAnterior.Range(intervalo).Columns(6).Column).Value = parcela Then
                If wsAnterior.Cells(linha, wsAnterior.Range(intervaloAdiantada).Column).Value = "Adiantada" Then
                    ParcelaJaAdiantada = True
                    Exit Function
                End If
            End If
        Next linha
    Next m
    ParcelaJaAdiantada = False
End Function

Sub RemoverParcelaAdiantada(wsMes As Worksheet, produto As String, parcela As Integer, intervalo As String)
    Dim linha As Integer
    For linha = wsMes.Range(intervalo).Rows(1).Row To wsMes.Range(intervalo).Rows(wsMes.Range(intervalo).Rows.Count).Row
        If wsMes.Cells(linha, wsMes.Range(intervalo).Columns(3).Column).Value = produto And _
           wsMes.Cells(linha, wsMes.Range(intervalo).Columns(6).Column).Value = parcela Then
            wsMes.Range(wsMes.Cells(linha, wsMes.Range(intervalo).Columns(1).Column), _
                        wsMes.Cells(linha, wsMes.Range(intervalo).Columns(7).Column)).ClearContents
            Exit For
        End If
    Next linha
End Sub


