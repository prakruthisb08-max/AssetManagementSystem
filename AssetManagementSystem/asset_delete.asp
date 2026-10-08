<%@ Language="VBScript" %>
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/db.asp"-->
<%
Dim mac, cmd, sql
mac = LCase(Replace(Replace(Trim(Request.QueryString("mac")), "-", ":"), " ", ""))

If mac <> "" Then
    sql = "DELETE FROM assets WHERE mac_address = ?"
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = 1
    cmd.Parameters.Append cmd.CreateParameter("p1", 200, 1, 50, mac)
    cmd.Execute
    Set cmd = Nothing
End If

conn.Close
Set conn = Nothing
Response.Redirect "asset_report.asp"
Response.End
%>
