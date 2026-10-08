<%
If Session("staffId") = "" Then
    Response.Redirect "login.asp"
    Response.End
End If
%>
