Attribute VB_Name = "modFuncoes"
Public Function ConsultarTasks(frm As Object)
    Dim queryInputHistory As String

    ' Verifica se o campo de pesquisa est� preenchido
    If frm.inputHistoryFilter.Text = "" Then
        MsgBox "Preencha o campo de pesquisa antes de consultar!", vbExclamation, "Aviso"
        Exit Function
    End If

    ' Monta a query SQL
    queryInputHistory = "SELECT * FROM Tasks WHERE Descricao = '" & frm.inputHistoryFilter.Text & "'"

    ' Executa a consulta
    recordBD.Open queryInputHistory, connectBD, adOpenStatic, adLockReadOnly

    ' Limpa o campo de pesquisa
    frm.inputHistoryFilter.Text = ""

    ' Verifica se encontrou resultados antes de tentar acessar os campos
    If Not recordBD.EOF Then
        MsgBox "Tarefa encontrada: " & recordBD.Fields("Descricao").Value, vbInformation, "Resultado"
    Else
        MsgBox "Nenhuma tarefa encontrada!", vbExclamation, "Aviso"
    End If

    ' Fecha o Recordset
    recordBD.Close
End Function

Public Function addTasks(frm as Object)
    If listTasks.ListIndex <> -1 Then
        tarefaClicadaDesc = listTasks.List(listTasks.ListIndex)
        tarefaClicadaDesc = "[CHECK!] " & tarefaClicadaDesc
        INSERT INTO Tasks (descricao, status) VALUES ("'" & tarefaClicadaDesc & "'", "CONCLUIDA")
    Else
        MsgBox "Selecione uma tarefa a ser conclu�da!", vbExclamation, "Aviso"
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