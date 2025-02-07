Attribute VB_Name = "modConectarBD"
Public connectBD As ADODB.Connection
Public recordBD As ADODB.Recordset
Public myBD As String

Public Sub InitConexao()

Set connectBD = New ADODB.Connection
Set recordBD = New ADODB.Recordset

myBD = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\todolist_db.mdb"

MsgBox myBD
connectBD.Open myBD

If connectBD.State = adStateOpen Then
MsgBox "Conexão aberta com sucesso"
Else
MsgBox "Deu errado"
End If

End Sub
