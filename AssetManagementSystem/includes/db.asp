<%
' ============================================================
' Asset Management System - MySQL Database Connection
' Classic ASP + ADO + MySQL ODBC
' ============================================================

Dim conn, connStr, dbError
Set conn = Server.CreateObject("ADODB.Connection")

' Change ONLY the DRIVER name if your installed driver is different.
' Common names:
' MySQL ODBC 8.0 Unicode Driver
' MySQL ODBC 9.5 Unicode Driver
' MySQL ODBC 9.6 Unicode Driver
' MySQL ODBC 26.7 Unicode Driver

connStr = "DRIVER={MySQL ODBC 8.0 Unicode Driver};" & _
          "SERVER=localhost;" & _
          "PORT=3306;" & _
          "DATABASE=AssetManagementDB;" & _
          "USER=root;" & _
          "PASSWORD=;" & _
          "OPTION=3;"

On Error Resume Next
conn.Open connStr
dbError = Err.Description
On Error GoTo 0

If conn.State <> 1 Then
    Response.Write "<h2>Database connection failed</h2>"
    Response.Write "<p>Check MySQL Server, ODBC Driver, database name, username/password, and db.asp.</p>"
    Response.Write "<p>Error: " & Server.HTMLEncode(dbError) & "</p>"
    Response.End
End If
%>
