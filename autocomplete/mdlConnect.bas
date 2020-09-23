Attribute VB_Name = "mdlConnect"
Public adoc As New ADODB.Connection


Public Sub OpenConnection()

Set adoc = New ADODB.Connection
With adoc
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbase.mdb;Persist Security Info=False"
    .CommandTimeout = 0
    .CursorLocation = adUseClient
    .Open
End With
End Sub




