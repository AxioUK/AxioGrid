Attribute VB_Name = "mdlSQLFunctions"
Option Explicit

Public Cnn        As ADODB.Connection

Public sqlConn As String

Public Sub Connect(iOpen As Boolean)

sqlConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                App.Path & "\DBMAINCB.accdb;Persist Security Info=False;"

If iOpen = True Then
  Set Cnn = New ADODB.Connection
  Cnn.Open sqlConn
Else
  Cnn.Close
  Set Cnn = Nothing
End If
End Sub

