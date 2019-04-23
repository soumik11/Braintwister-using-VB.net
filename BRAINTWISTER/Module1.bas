Attribute VB_Name = "Module1"
Option Explicit

Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public sql As String

Public Sub connect()
sql = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ab.mdb;Persist Security Info=False"
If con.State = adStateClosed Then
con.Open (sql)
End If
MsgBox ("user connected")
End Sub
