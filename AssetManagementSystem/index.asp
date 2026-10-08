<%@ Language="VBScript" %>
<!--#include file="includes/auth.asp"-->
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Asset Management Dashboard</title>
<link rel="stylesheet" href="wwwroot/css/style.css">
</head>
<body>
<div class="topbar">
    <div><strong>IT Asset Management</strong></div>
    <div>Welcome, <%=Server.HTMLEncode(Session("fullName"))%> |
        <a href="logout.asp">Logout</a></div>
</div>

<div class="container">
    <h1>Asset Management</h1>
    <p>Select an operation.</p>

    <div class="cards">
        <a class="menu-card" href="asset_declaration.asp">
            <h2>Asset Declaration</h2>
            <p>Declare/register a new IT asset.</p>
        </a>
        <a class="menu-card" href="asset_report.asp">
            <h2>Asset Report</h2>
            <p>View, search, edit, delete and export assets.</p>
        </a>
    </div>
</div>
</body>
</html>
