<%@ Language="VBScript" %>
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/db.asp"-->
<%
Option Explicit
Dim searchText, rs, cmd, sql, exportCsv
searchText = Trim(Request.QueryString("search"))
exportCsv = (LCase(Request.QueryString("export")) = "csv")

If exportCsv Then
    Response.Clear
    Response.ContentType = "text/csv"
    Response.AddHeader "Content-Disposition", "attachment; filename=asset_report.csv"
End If

If searchText <> "" Then
    sql = "SELECT * FROM assets WHERE " & _
          "mac_address LIKE ? OR asset_number LIKE ? OR asset_type LIKE ? OR make LIKE ? OR model LIKE ? OR serial_no LIKE ? OR ip_address LIKE ? OR zone LIKE ? OR location LIKE ? OR hostname LIKE ? OR domainname LIKE ? OR operating_system LIKE ? OR year_of_manufacture LIKE ? OR network_type LIKE ? OR warranty LIKE ? OR staff_id LIKE ? OR cost_center LIKE ? " & _
          "ORDER BY created_at DESC"
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandText = sql
    cmd.CommandType = 1
    Dim s
    s = "%" & searchText & "%"
    cmd.Parameters.Append cmd.CreateParameter("p1", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p2", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p3", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p4", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p5", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p6", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p7", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p8", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p9", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p10", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p11", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p12", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p13", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p14", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p15", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p16", 200, 1, 100, s)
    cmd.Parameters.Append cmd.CreateParameter("p17", 200, 1, 100, s)
    Set rs = cmd.Execute
Else
    sql = "SELECT * FROM assets ORDER BY created_at DESC"
    Set rs = conn.Execute(sql)
End If

If exportCsv Then
    Response.Write "Staff ID,Cost Center,Asset Number,Asset Type,Make,Model,Serial Number,MAC Address,IP Address,Hostname,Domain Name,RAM (GB),Hard Disk (GB),Hard Disk Used (GB),Operating System,Year of Manufacture,Type of Network,Warranty,Zone,Location,Asset Used By,Declared By,Created At" & vbCrLf
    Do Until rs.EOF
        Response.Write Csv(rs("staff_id")) & "," & Csv(rs("cost_center")) & "," & Csv(rs("asset_number")) & "," & Csv(rs("asset_type")) & "," & _
                       Csv(rs("make")) & "," & Csv(rs("model")) & "," & Csv(rs("serial_no")) & "," & _
                       Csv(rs("mac_address")) & "," & Csv(rs("ip_address")) & "," & Csv(rs("hostname")) & "," & Csv(rs("domainname")) & "," & _
                       Csv(rs("ram")) & "," & Csv(rs("hard_disk")) & "," & Csv(rs("hard_disk_used")) & "," & _
                       Csv(rs("operating_system")) & "," & Csv(rs("year_of_manufacture")) & "," & Csv(rs("network_type")) & "," & _
                       Csv(rs("warranty")) & "," & Csv(rs("zone")) & "," & Csv(rs("location")) & "," & Csv(rs("asset_used_by")) & "," & _
                       Csv(rs("declared_by")) & "," & Csv(rs("created_at")) & vbCrLf
        rs.MoveNext
    Loop
    rs.Close
    If Not cmd Is Nothing Then Set cmd = Nothing
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    Response.End
End If

Function Csv(v)
    If IsNull(v) Then
        Csv = """"""
    Else
        Csv = """" & Replace(CStr(v), """", """""") & """"
    End If
End Function
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Asset Report</title>
<link rel="stylesheet" href="wwwroot/css/style.css">
</head>
<body>
<div class="topbar">
    <strong>IT Asset Management</strong>
    <div><a href="index.asp">Home</a> | <a href="asset_declaration.asp">Declare Asset</a> | <a href="logout.asp">Logout</a></div>
</div>

<div class="container">
<h1>Asset Report</h1>

<form method="get" action="asset_report.asp" class="search-bar">
    <input type="text" name="search" value="<%=Server.HTMLEncode(searchText)%>" placeholder="Search MAC, asset number, type, make, model, serial or IP">
    <button type="submit" class="btn primary">Search</button>
    <a href="asset_report.asp" class="btn secondary">Clear</a>
    <a href="asset_report.asp?export=csv&search=<%=Server.URLEncode(searchText)%>" class="btn success">Export CSV</a>
</form>

<div class="table-wrap">
<table>
<thead>
<tr>
    <th>Staff ID</th>
    <th>Cost Center</th>
    <th>Asset Number</th>
    <th>Asset Type</th>
    <th>Make</th>
    <th>Model</th>
    <th>Serial Number</th>
    <th>MAC Address</th>
    <th>IP Address</th>
    <th>Hostname</th>
    <th>Domain Name</th>
    <th>RAM (GB)</th>
    <th>Hard Disk (GB)</th>
    <th>Hard Disk Used (GB)</th>
    <th>Operating System</th>
    <th>Year of Manufacture</th>
    <th>Type of Network</th>
    <th>Warranty</th>
    <th>Zone</th>
    <th>Location</th>
    <th>Asset Used By</th>
    <th>Action</th>
</tr>
</thead>
<tbody>
<%
If rs.EOF Then
    Response.Write "<tr><td colspan='23' class='empty'>No assets found.</td></tr>"
Else
    Do Until rs.EOF
%>
<tr>
    <td><%=Server.HTMLEncode(rs("staff_id"))%></td>
    <td><%=Server.HTMLEncode(rs("cost_center"))%></td>
    <td><%=Server.HTMLEncode(rs("asset_number"))%></td>
    <td><%=Server.HTMLEncode(rs("asset_type"))%></td>
    <td><%=Server.HTMLEncode(rs("make"))%></td>
    <td><%=Server.HTMLEncode(rs("model"))%></td>
    <td><%=Server.HTMLEncode(rs("serial_no"))%></td>
    <td><%=Server.HTMLEncode(rs("mac_address"))%></td>
    <td><%=Server.HTMLEncode(rs("ip_address"))%></td>
    <td><%=Server.HTMLEncode(rs("hostname"))%></td>
    <td><%=Server.HTMLEncode(rs("domainname"))%></td>
    <td><%=Server.HTMLEncode(rs("ram"))%></td>
    <td><%=Server.HTMLEncode(rs("hard_disk"))%></td>
    <td><%=Server.HTMLEncode(rs("hard_disk_used"))%></td>
    <td><%=Server.HTMLEncode(rs("operating_system"))%></td>
    <td><%=Server.HTMLEncode(rs("year_of_manufacture"))%></td>
    <td><%=Server.HTMLEncode(rs("network_type"))%></td>
    <td><%=Server.HTMLEncode(rs("warranty"))%></td>
    <td><%=Server.HTMLEncode(rs("zone"))%></td>
    <td><%=Server.HTMLEncode(rs("location"))%></td>
    <td><%=Server.HTMLEncode(rs("asset_used_by"))%></td>
    <td class="nowrap">
        <a class="btn small" href="asset_edit.asp?mac=<%=Server.URLEncode(rs("mac_address"))%>">Edit</a>
        <a class="btn small danger" href="asset_delete.asp?mac=<%=Server.URLEncode(rs("mac_address"))%>" onclick="return confirm('Delete this asset?');">Delete</a>
    </td>
</tr>
<%
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing
If Not cmd Is Nothing Then Set cmd = Nothing
conn.Close
Set conn = Nothing
%>
</tbody>
</table>
</div>
</div>
</body>
</html>
