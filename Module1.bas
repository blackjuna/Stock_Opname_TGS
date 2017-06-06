Attribute VB_Name = "Module1"
Public conn As New ADODB.Connection
Public conn_excel As New ADODB.Connection
Public rs_so As ADODB.Recordset
Public rs_export As ADODB.Recordset
Public rs_import As ADODB.Recordset

Public Sub db()
Set conn = New ADODB.Connection
Set conn_excel = New ADODB.Connection
Set rs_so = New ADODB.Recordset

strnamadatabase = "stock_opname_db"
strnamaserver = "proliant\sqlexpress"
'strnamaserver = "HDGNGIT002\SQLEXPRESS"
'strnamaserver = "IT-JUN-PC"
strnamapemakai = "sa"
strpassword = "admin123"
intconnectiontimeout = 60
strprovider = "SQLOLEDB.1"

strconstr = "Provider=" & strprovider & "; Data Source=" & "192.168.10.250, 1433" & "; Network Library=DBMSSOCN; Initial Catalog=" & strnamadatabase & ";User ID=" & strnamapemakai & ";Password=" & strpassword
'strconstr = "Provider=" & strprovider & "; Data Source=" & " 115.85.74.130,8795 " & "; Network Library=DBMSSOCN; Initial Catalog=" & strnamadatabase & ";User ID=" & strnamapemakai & ";Password=" & strpassword
'strconstr = "Provider=" & strprovider & "; Data Source=" & "192.168.0.108, 1433" & "; Network Library=DBMSSOCN; Initial Catalog=" & strnamadatabase & ";User ID=" & strnamapemakai & ";Password=" & strpassword
'strconstr = "Provider=" & strprovider & ";Data Source=" & strnamaserver & ";Initial Catalog=" & strnamadatabase & ";User ID=" & strnamapemakai & ";Password=" & strpassword
conn.ConnectionString = strconstr
conn.CursorLocation = adUseClient
conn.Open
'conn.Open "driver={mysql odbc 5.1 driver};server=localhost;uid=root;pwd=;db=stock_opname_db;"
'conn.Open "provider=" & strprovider & ";Data Source=" & "192.168.0.108, 1433" & ";"

End Sub

Public Function CheckCharacter(ByVal xParam As String) As String
If InStr(1, xParam, "'") Then
    retValue = Replace(xParam, "'", "''", , , vbTextCompare)
Else
    retValue = xParam
End If
    CheckCharacter = retValue
End Function

