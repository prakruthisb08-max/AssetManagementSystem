<%@ Language="VBScript" %>
<%
Option Explicit
Dim errorMessage, staffIdInput, passwordInput

If Session("staffId") <> "" Then
    Response.Redirect "index.asp"
    Response.End
End If

errorMessage = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    staffIdInput = Trim(Request.Form("staff_id"))
    passwordInput = Trim(Request.Form("password"))

    If staffIdInput = "" Or passwordInput = "" Then
        errorMessage = "Please enter Staff ID and password."
    Else
        %><!--#include file="includes/db.asp"--><%
        Dim rs, cmd, sql
        sql = "SELECT staff_id, full_name FROM users WHERE staff_id = ? AND password_value = ? AND is_active = 1"

        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = conn
        cmd.CommandText = sql
        cmd.CommandType = 1
        cmd.Parameters.Append cmd.CreateParameter("p1", 200, 1, 50, staffIdInput)
        cmd.Parameters.Append cmd.CreateParameter("p2", 200, 1, 50, passwordInput)

        Set rs = cmd.Execute

        If Not rs.EOF Then
            Session("staffId") = rs("staff_id")
            Session("fullName") = rs("full_name")
            rs.Close
            Set rs = Nothing
            Set cmd = Nothing
            conn.Close
            Set conn = Nothing
            Response.Redirect "index.asp"
            Response.End
        Else
            errorMessage = "Invalid Staff ID or password."
        End If

        rs.Close
        Set rs = Nothing
        Set cmd = Nothing
        conn.Close
        Set conn = Nothing
    End If
End If
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Asset Management - Login</title>
<link rel="stylesheet" href="wwwroot/css/style.css">
</head>
<body class="login-page">
<div class="login-card">
    <h1>IT Asset Management</h1>
    <p class="subtitle">Staff Login</p>

    <% If errorMessage <> "" Then %>
        <div class="alert error"><%=Server.HTMLEncode(errorMessage)%></div>
    <% End If %>

    <form method="post" action="login.asp">
        <label>Staff ID</label>
        <input type="text" name="staff_id" maxlength="50" required>

        <label>Password</label>
        <input type="password" name="password" maxlength="50" required>

        <button type="submit" class="btn primary full">Login</button>
    </form>

    <p class="hint">Password format: DDMMYYYY</p>
</div>
</body>
</html>
