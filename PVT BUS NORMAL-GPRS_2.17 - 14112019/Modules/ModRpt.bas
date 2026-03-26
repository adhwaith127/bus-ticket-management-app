Attribute VB_Name = "ModRpt"
Public CNN As DAO.Database
Public Function CONNECTDB() As Boolean
On Error GoTo err
CONNECTDB = True
DBPWD = "silbus"
 
 'CON.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Pvt.mdb; Jet OLEDB:Database Password =" & DBPWD
 'CON.Open

 Set CNN = DAO.OpenDatabase(App.Path & "\Pvt.mdb", dbDriverComplete, False, ";UID=;PWD=" & DBPWD)
Exit Function
err:
    CONNECT_DB = False
    MsgBox err.Description & vbCrLf & err.Number, vbCritical, "Distribution"
    Exit Function
End Function
Public Function getvalueQuery(strSql As String) As String
'Author Mubeena
On Error GoTo lblErr
Dim oRS As New ADODB.Recordset
Dim QryValue As String
  Set oRS = gbladoCon.Execute(strSql)
  If oRS.EOF = False Then
        QryValue = oRS.Fields(0)
    Else
        QryValue = ""
    End If
    Set oRS = Nothing
    getvalueQuery = QryValue
    Exit Function
lblErr:
'MsgBox err.Description, vbOKOnly, prjTitle
QryValue = ""
End Function

