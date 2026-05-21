Attribute VB_Name = "Calendario"
Sub calendario_prog_icones()
    Dim wsCalendario As Worksheet
    Dim wsProgramacao As Worksheet
    Dim dataBusca As Date
    Dim ultimaLinha As Long
    Dim i As Long, j As Integer
    Dim resultado As String
    Dim colunasDestino As Variant

    ' Define as planilhas
    Set wsCalendario = ThisWorkbook.Sheets("Calendário")
    Set wsProgramacao = ThisWorkbook.Sheets("Programações")

    ' Obtém os filtros da aba "calendario" uma vez
    Dim filtroObra As String: filtroObra = Trim(wsCalendario.Range("B2").value) 'obra
    Dim filtroEquipe As String: filtroEquipe = Trim(wsCalendario.Range("C2").value) 'equipe
    Dim filtroStatus As String: filtroStatus = Trim(wsCalendario.Range("D2").value) 'status
    Dim filtromes As String: filtromes = Trim(wsCalendario.Range("E2").value) 'mes
    Dim filtroano As String: filtroano = Trim(wsCalendario.Range("F2").value) 'ano
    
    If filtromes = Empty Or filtroano = Empty Then
        MsgBox "Necessário preencher ano/mês para busca", vbCritical, "Data inválida"
        Exit Sub
    End If
    
    
    ' Define as colunas de destino
    colunasDestino = Array("A5", "B5", "C5", "D5", "E5", "F5", "G5", _
                           "A7", "B7", "C7", "D7", "E7", "F7", "G7", _
                           "A9", "B9", "C9", "D9", "E9", "F9", "G9", _
                           "A11", "B11", "C11", "D11", "E11", "F11", "G11", _
                           "A13", "B13", "C13", "D13", "E13", "F13", "G13")

    ' Otimização: desativa atualizações de tela, eventos e cálculo automático
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    ' Percorre as células de destino
    For j = LBound(colunasDestino) To UBound(colunasDestino)
        dataBusca = wsCalendario.Range(colunasDestino(j)).Offset(-1, 0).value
        resultado = ""

        ultimaLinha = wsProgramacao.Cells(wsProgramacao.Rows.Count, "A").End(xlUp).Row

        ' Constrói o resultado
        For i = 3 To ultimaLinha
            If wsProgramacao.Cells(i, 1).value = dataBusca Then
                Dim status As String: status = Trim(wsProgramacao.Cells(i, 14).value) 'status
                Dim equipeAtual As String: equipeAtual = Trim(wsProgramacao.Cells(i, 2).value) 'equipe
                Dim obraAtual As String: obraAtual = Trim(wsProgramacao.Cells(i, 3).value) 'obra

                If (filtroObra = "" Or obraAtual = filtroObra) And _
                   (filtroEquipe = "" Or equipeAtual = filtroEquipe) And _
                   (filtroStatus = "" Or status = filtroStatus) Then
                   
                    Select Case UCase(status)
                        Case "PROGRAMADO": resultado = resultado & equipeAtual & " - " & obraAtual & " " & ChrW(9679) & vbCrLf
                        Case "CANCELADO": resultado = resultado & equipeAtual & " - " & obraAtual & " " & ChrW(10008) & vbCrLf
                        Case "EXECUTADO": resultado = resultado & equipeAtual & " - " & obraAtual & " " & ChrW(10004) & vbCrLf
                        Case "SEM PROGRAMAÇÃO": resultado = resultado & equipeAtual & " - " & obraAtual & " " & ChrW(9675) & vbCrLf
                        Case "REPROGRAMADO": resultado = resultado & equipeAtual & " - " & obraAtual & " " & ChrW(8594) & vbCrLf
                        Case Else: resultado = resultado & equipeAtual & " - " & obraAtual & " " & ChrW(36) & vbCrLf
                    End Select
                End If
            End If
        Next i

        ' Escreve e formata o resultado
        With wsCalendario.Range(colunasDestino(j))
            If resultado <> "" Then
                .value = Left(resultado, Len(resultado) - 1)
                Dim texto As String: texto = .value
                Dim pos As Integer: pos = 1
                
                While pos <= Len(texto)
                    Select Case AscW(Mid(texto, pos, 1))
                        Case 9679: .Characters(pos, 1).Font.Color = vbBlue
                        Case 10008: .Characters(pos, 1).Font.Color = vbRed
                        Case 10004: .Characters(pos, 1).Font.Color = vbGreen
                        Case 9675: .Characters(pos, 1).Font.Color = RGB(255, 165, 0)
                        Case 8594: .Characters(pos, 1).Font.Color = vbMagenta
                        Case 36: .Characters(pos, 1).Font.Color = vbBlack
                    End Select
                    pos = pos + 1
                Wend
            Else
                .value = ""
            End If
        End With
    Next j

    ' Restaura configurações padrão
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With

    MsgBox "Calendário atualizado", vbInformation
    wsCalendario.Range("E2").Select
End Sub





