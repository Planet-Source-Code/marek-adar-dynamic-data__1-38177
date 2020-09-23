Attribute VB_Name = "modDatabase"
Public CN As New ADODB.Connection   'Connection
Public RSMetaData As New ADODB.Recordset 'Recordset for Mete-Data => Columnames and Datatypes
Public RSTableData As New ADODB.Recordset   'Recordset for tabledata
Public strDatabase As String    '
Public strTable As String
Public strServer As String

Sub SubCreateConnection(strServer As String, Optional strDatabase As String = "MASTER")
    'Opens the Connection to SQL-Server, Uses Trusted Connections
    If CN.State = adStateOpen Then CN.Close
    CN.Open "PROVIDER=SQLOLEDB;SERVER=" & strServer & ";DATABASE=" & strDatabase & ";INTEGRATED SECURITY=SSPI"
End Sub

Sub SubCreateRecordset(RS As ADODB.Recordset, Optional strSQL As String = "")
    'Creates the recordsets
    If RS.State = adStateOpen Then RS.Close
    RS.CursorLocation = adUseClient
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    Set RS.ActiveConnection = CN
    RS.Open strSQL
End Sub


