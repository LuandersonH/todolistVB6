Attribute VB_Name = "modFuncoes"
Public Function ConsultarTasks(frm As Object)
    Dim queryInputHistory As String
    
    ' Verifica se o campo de pesquisa está preenchido
    If frm.inputHistoryFilter.Text = "" Then
        MsgBox "Preencha o campo de pesquisa antes de consultar!", vbExclamation, "Aviso"
        Exit Function
    End If

    ' Monta a query SQL, evitando erros com aspas simples
    queryInputHistory = "SELECT * FROM Tasks WHERE Descricao = '" & Replace(frm.inputHistoryFilter.Text, "'", "''") & "'"

    ' Fecha o recordBD antes de abrir uma nova consulta
    If recordBD.State = adStateOpen Then recordBD.Close

    ' Executa a consulta
    recordBD.Open queryInputHistory, connectBD, adOpenStatic, adLockReadOnly

    ' Limpa o campo de pesquisa
    frm.inputHistoryFilter.Text = ""

    ' Verifica se encontrou resultados antes de acessar os campos
    If Not recordBD.EOF Then
        MsgBox "Tarefa encontrada: " & recordBD.Fields("Descricao").Value, vbInformation, "Resultado"
    Else
        MsgBox "Nenhuma tarefa encontrada!", vbExclamation, "Aviso"
    End If

    ' Fecha o Recordset
    recordBD.Close
End Function


Public Function addTasks(frm As Object)
    Dim queryAddTask As String
    Dim tarefaClicadaDesc As String

    ' Verifica se uma tarefa foi selecionada
    If listTasks.ListIndex <> -1 Then
        tarefaClicadaDesc = listTasks.List(listTasks.ListIndex)
        tarefaClicadaDesc = "[CHECK!] " & tarefaClicadaDesc

        ' Query correta para o INSERT
        queryAddTask = "INSERT INTO Tasks (descricao, status) VALUES ('" & Replace(tarefaClicadaDesc, "'", "''") & "', 'CONCLUIDA')"

        ' Executa a query com Execute (não usa recordBD.Open para INSERT)
        connectBD.Execute queryAddTask

        MsgBox "Tarefa adicionada com sucesso!", vbInformation, "Sucesso"
    Else
        MsgBox "Selecione uma tarefa a ser concluída!", vbExclamation, "Aviso"
    End If
End Function


//A FAZER
Private Sub btnDeleteTask_Click()
If listTasks.ListIndex <> -1 Then
listTasks.RemoveItem listTasks.ListIndex

Else
MsgBox "Selecione uma tarefa a ser removida!", vbExclamation, "Aviso"
End If
End Sub