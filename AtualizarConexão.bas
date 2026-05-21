Attribute VB_Name = "AtualizarConexão"
Sub AtualizarConexaoTabProgramacoes()
    Dim conn As WorkbookConnection
    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim connectionName As String
    Dim foundConn As Boolean
    Dim foundQt As Boolean
    
    connectionName = "Consulta - tab_programacao"
    foundConn = False
    foundQt = False
    
    Application.ScreenUpdating = False
    
    For Each conn In ThisWorkbook.Connections
        If LCase(conn.Name) = LCase(connectionName) Then
            foundConn = True
            On Error Resume Next
            
            Set ws = ThisWorkbook.Worksheets("Programações")
            
            For Each qt In ws.QueryTables
                If qt.Connection = conn.OLEDBConnection.Connection Then
                    foundQt = True
                    qt.BackgroundQuery = True
                    qt.Refresh
                    If Err.Number <> 0 Then
                        MsgBox "Erro ao atualizar a conexão '" & connectionName & "': " & Err.Description, vbCritical
                        Err.Clear
                        Application.ScreenUpdating = True
                        Exit Sub
                    End If
                    Exit For
                End If
            Next qt
            
            If Not foundQt Then
                conn.Refresh
                If Err.Number <> 0 Then
                    MsgBox "Erro ao atualizar a conexão '" & connectionName & "': " & Err.Description, vbCritical
                    Err.Clear
                    Application.ScreenUpdating = True
                    Exit Sub
                End If
            End If
            
            On Error GoTo 0
            Exit For
        End If
    Next conn
    
    If Not foundConn Then
        MsgBox connectionName & " não encontrada.", vbExclamation
    Else
        MsgBox connectionName & " atualizada com sucesso!", vbInformation
    End If
    
    Application.ScreenUpdating = True
End Sub

' Função auxiliar para concatenar itens do lote
Public Function JoinCollection(coll As Collection, delimiter As String) As String
    Dim result As String
    Dim i As Long
    result = ""
    For i = 1 To coll.Count
        result = result & coll(i)
        If i < coll.Count Then
            result = result & delimiter
        End If
    Next i
    JoinCollection = result
End Function

