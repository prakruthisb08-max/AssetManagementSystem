<%@ Language="VBScript" %>
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/db.asp"-->
<%
Option Explicit
Dim mac, rs, cmd, sql, message, errorMessage
mac = LCase(Replace(Replace(Trim(Request.QueryString("mac")), "-", ":"), " ", ""))
message = ""
errorMessage = ""

If mac = "" Then
    Response.Redirect "asset_report.asp"
    Response.End
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    mac = LCase(Replace(Replace(Trim(Request.Form("mac_address")), "-", ":"), " ", ""))

    Dim assetNumber, assetType, makeName, modelName, serialNo, ipAddress
    Dim ramValue, hardDiskValue, hardDiskUsedValue, assetUsedBy
    Dim zoneValue, locationValue, hostnameValue, domainNameValue, operatingSystemValue
    Dim manufactureYearValue, networkTypeValue, warrantyValue, staffIdValue, costCenterValue
    assetNumber = Trim(Request.Form("asset_number"))
    assetType = Trim(Request.Form("asset_type"))
    makeName = Trim(Request.Form("make"))
    modelName = Trim(Request.Form("model"))
    serialNo = Trim(Request.Form("serial_no"))
    ipAddress = Trim(Request.Form("ip_address"))
    ramValue = Trim(Request.Form("ram"))
    hardDiskValue = Trim(Request.Form("hard_disk"))
    hardDiskUsedValue = Trim(Request.Form("hard_disk_used"))
    assetUsedBy = Trim(Request.Form("asset_used_by"))
    zoneValue = Trim(Request.Form("zone"))
    locationValue = Trim(Request.Form("location"))
    hostnameValue = Trim(Request.Form("hostname"))
    domainNameValue = Trim(Request.Form("domainname"))
    operatingSystemValue = Trim(Request.Form("operating_system"))
    manufactureYearValue = Trim(Request.Form("year_of_manufacture"))
    networkTypeValue = Trim(Request.Form("network_type"))
    warrantyValue = Trim(Request.Form("warranty"))
    staffIdValue = Trim(Request.Form("staff_id"))
    costCenterValue = Trim(Request.Form("cost_center"))

    If mac = "" Then
        errorMessage = "MAC Address is required."
    ElseIf assetNumber = "" Then
        errorMessage = "Asset Number is required."
    ElseIf assetType = "" Then
        errorMessage = "Asset Type is required."
    ElseIf zoneValue = "" Or locationValue = "" Or hostnameValue = "" Or domainNameValue = "" Then
        errorMessage = "Zone, Location, Hostname and Domain Name are required."
    ElseIf operatingSystemValue = "" Or manufactureYearValue = "" Or networkTypeValue = "" Or warrantyValue = "" Then
        errorMessage = "Operating System, Manufacture Year, Network Type and Warranty are required."
    ElseIf staffIdValue = "" Or costCenterValue = "" Then
        errorMessage = "Staff ID and Cost Center are required."
    ElseIf Not IsValidMac(mac) Then
        errorMessage = "Enter a valid MAC Address."
    ElseIf IsNumeric(hardDiskValue) And IsNumeric(hardDiskUsedValue) Then
        If CDbl(hardDiskUsedValue) > CDbl(hardDiskValue) Then
            errorMessage = "Hard Disk Used cannot be greater than Hard Disk capacity."
        End If
    End If

    If errorMessage = "" Then
        sql = "UPDATE assets SET staff_id=?, cost_center=?, asset_number=?, asset_type=?, make=?, model=?, serial_no=?, ip_address=?, hostname=?, domainname=?, ram=?, hard_disk=?, hard_disk_used=?, operating_system=?, year_of_manufacture=?, network_type=?, warranty=?, zone=?, location=?, asset_used_by=?, updated_by=? WHERE mac_address=?"
        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = conn
        cmd.CommandText = sql
        cmd.CommandType = 1

        cmd.Parameters.Append cmd.CreateParameter("p1", 200, 1, 50, staffIdValue)
        cmd.Parameters.Append cmd.CreateParameter("p2", 200, 1, 50, costCenterValue)
        cmd.Parameters.Append cmd.CreateParameter("p3", 200, 1, 50, assetNumber)
        cmd.Parameters.Append cmd.CreateParameter("p4", 200, 1, 100, assetType)
        cmd.Parameters.Append cmd.CreateParameter("p5", 200, 1, 100, makeName)
        cmd.Parameters.Append cmd.CreateParameter("p6", 200, 1, 100, modelName)
        cmd.Parameters.Append cmd.CreateParameter("p7", 200, 1, 100, serialNo)
        cmd.Parameters.Append cmd.CreateParameter("p8", 200, 1, 45, ipAddress)
        cmd.Parameters.Append cmd.CreateParameter("p9", 200, 1, 100, hostnameValue)
        cmd.Parameters.Append cmd.CreateParameter("p10", 200, 1, 150, domainNameValue)
        cmd.Parameters.Append cmd.CreateParameter("p11", 200, 1, 50, ramValue)
        cmd.Parameters.Append cmd.CreateParameter("p12", 200, 1, 50, hardDiskValue)
        cmd.Parameters.Append cmd.CreateParameter("p13", 200, 1, 50, hardDiskUsedValue)
        cmd.Parameters.Append cmd.CreateParameter("p14", 200, 1, 100, operatingSystemValue)
        cmd.Parameters.Append cmd.CreateParameter("p15", 3, 1, 0, CLng(manufactureYearValue))
        cmd.Parameters.Append cmd.CreateParameter("p16", 200, 1, 30, networkTypeValue)
        cmd.Parameters.Append cmd.CreateParameter("p17", 200, 1, 3, warrantyValue)
        cmd.Parameters.Append cmd.CreateParameter("p18", 200, 1, 100, zoneValue)
        cmd.Parameters.Append cmd.CreateParameter("p19", 200, 1, 150, locationValue)
        cmd.Parameters.Append cmd.CreateParameter("p20", 200, 1, 20, assetUsedBy)
        cmd.Parameters.Append cmd.CreateParameter("p21", 200, 1, 50, Session("staffId"))
        cmd.Parameters.Append cmd.CreateParameter("p22", 200, 1, 50, mac)

        On Error Resume Next
        cmd.Execute
        If Err.Number <> 0 Then
            errorMessage = "Could not update asset: " & Err.Description
            Err.Clear
        Else
            message = "Asset updated successfully."
        End If
        On Error GoTo 0
        Set cmd = Nothing
    End If
End If

sql = "SELECT * FROM assets WHERE mac_address = ?"
Set cmd = Server.CreateObject("ADODB.Command")
Set cmd.ActiveConnection = conn
cmd.CommandText = sql
cmd.CommandType = 1
cmd.Parameters.Append cmd.CreateParameter("p1", 200, 1, 50, mac)
Set rs = cmd.Execute

If rs.EOF Then
    rs.Close
    Set rs = Nothing
    Set cmd = Nothing
    conn.Close
    Set conn = Nothing
    Response.Redirect "asset_report.asp"
    Response.End
End If

Function IsValidMac(macValue)
    Dim re
    Set re = New RegExp
    re.Pattern = "^[0-9a-fA-F]{2}(:[0-9a-fA-F]{2}){5}$"
    IsValidMac = re.Test(macValue)
    Set re = Nothing
End Function
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Edit Asset</title>
<link rel="stylesheet" href="wwwroot/css/style.css">
<script src="wwwroot/js/script.js"></script>
</head>
<body>
<div class="topbar">
    <strong>IT Asset Management</strong>
    <div><a href="index.asp">Home</a> | <a href="asset_report.asp">Asset Report</a> | <a href="logout.asp">Logout</a></div>
</div>

<div class="container form-container">
<h1>Edit Asset</h1>

<% If message <> "" Then %><div class="alert success"><%=Server.HTMLEncode(message)%></div><% End If %>
<% If errorMessage <> "" Then %><div class="alert error"><%=Server.HTMLEncode(errorMessage)%></div><% End If %>

<form method="post" action="asset_edit.asp?mac=<%=Server.URLEncode(mac)%>" onsubmit="return validateAssetForm();">
<div class="form-grid">
    <div><label>Staff ID *</label><input type="text" name="staff_id" value="<%=Server.HTMLEncode(rs("staff_id"))%>" maxlength="50" required></div>
    <div><label>Cost Center *</label><input type="text" name="cost_center" value="<%=Server.HTMLEncode(rs("cost_center"))%>" maxlength="50" required></div>

    <div><label>Asset Number *</label><input type="text" name="asset_number" maxlength="50" value="<%=Server.HTMLEncode(rs("asset_number"))%>" required></div>
    <div>
        <label>Asset Type *</label>
        <select name="asset_type" required><option value="">Select Asset Type</option>
        <% Dim types, i : types=Array("Desktop PC - High End","Desktop PC - Standard","Laptop - High End","Laptop - Standard","Laptop - Medium","Workstation - High","Workstation - Standard","Workstation - Medium","Printer","Scanner","Server","Other")
        For i=0 To UBound(types)
            Response.Write "<option value="" & Server.HTMLEncode(types(i)) & """
            If CStr(rs("asset_type"))=types(i) Then Response.Write " selected"
            Response.Write ">" & Server.HTMLEncode(types(i)) & "</option>"
        Next %></select>
    </div>
    <div><label>Make</label><input type="text" name="make" value="<%=Server.HTMLEncode(rs("make"))%>" maxlength="100"></div>
    <div><label>Model</label><input type="text" name="model" value="<%=Server.HTMLEncode(rs("model"))%>" maxlength="100"></div>
    <div><label>Serial Number</label><input type="text" name="serial_no" value="<%=Server.HTMLEncode(rs("serial_no"))%>" maxlength="100"></div>
    <div><label>MAC Address *</label><input type="text" name="mac_address" value="<%=Server.HTMLEncode(rs("mac_address"))%>" readonly><small>MAC Address is the PRIMARY KEY and cannot be changed.</small></div>
    <div><label>IP Address</label><input type="text" name="ip_address" value="<%=Server.HTMLEncode(rs("ip_address"))%>" maxlength="45"></div>
    <div><label>Hostname *</label><input type="text" name="hostname" value="<%=Server.HTMLEncode(rs("hostname"))%>" maxlength="100" required></div>
    <div><label>Domain Name *</label><input type="text" name="domainname" value="<%=Server.HTMLEncode(rs("domainname"))%>" maxlength="150" required></div>
    <div><label>RAM (GB)</label><input type="number" name="ram" value="<%=Server.HTMLEncode(rs("ram"))%>" min="0" step="0.01" maxlength="50"></div>
    <div><label>Hard Disk (GB)</label><input type="number" name="hard_disk" value="<%=Server.HTMLEncode(rs("hard_disk"))%>" min="0" step="0.01" maxlength="50"></div>
    <div><label>Hard Disk Used (GB)</label><input type="number" name="hard_disk_used" value="<%=Server.HTMLEncode(rs("hard_disk_used"))%>" min="0" step="0.01" maxlength="50"></div>

    <div>
        <label>Operating System *</label>
        <select name="operating_system" required><option value="">Select Operating System</option>
        <% Dim osList, oi : osList=Array("Windows 11","Windows 12","Windows Server 2012","Windows Server 2016","Windows Server 2019","Windows Server 2022","Thin OS")
        For oi=0 To UBound(osList)
            Response.Write "<option value="" & Server.HTMLEncode(osList(oi)) & """
            If CStr(rs("operating_system"))=osList(oi) Then Response.Write " selected"
            Response.Write ">" & Server.HTMLEncode(osList(oi)) & "</option>"
        Next %></select>
    </div>
    <div>
        <label>Year of Manufacture *</label>
        <select name="year_of_manufacture" required><option value="">Select Year</option>
        <% Dim yy : For yy=Year(Date) To 1990 Step -1
            Response.Write "<option value="" & yy & """
            If Not IsNull(rs("year_of_manufacture")) And CInt(rs("year_of_manufacture"))=yy Then Response.Write " selected"
            Response.Write ">" & yy & "</option>"
        Next %></select>
    </div>
    <div>
        <label>Type of Network *</label>
        <select name="network_type" required><option value="">Select Network Type</option>
            <option <% If CStr(rs("network_type"))="Internet" Then Response.Write "selected" %>>Internet</option>
            <option <% If CStr(rs("network_type"))="Intranet" Then Response.Write "selected" %>>Intranet</option>
            <option <% If CStr(rs("network_type"))="Standalone" Then Response.Write "selected" %>>Standalone</option>
        </select>
    </div>
    <div>
        <label>Warranty *</label><select name="warranty" required><option value="">Select</option>
            <option <% If CStr(rs("warranty"))="Yes" Then Response.Write "selected" %>>Yes</option>
            <option <% If CStr(rs("warranty"))="No" Then Response.Write "selected" %>>No</option>
        </select>
    </div>
    <div><label>Zone *</label><input type="text" name="zone" value="<%=Server.HTMLEncode(rs("zone"))%>" maxlength="100" required></div>
    <div><label>Location *</label><input type="text" name="location" value="<%=Server.HTMLEncode(rs("location"))%>" maxlength="150" required></div>
    <div>
        <label>Asset Used By *</label><select name="asset_used_by" required><option value="">Select</option>
            <option value="Self" <% If CStr(rs("asset_used_by"))="Self" Then Response.Write "selected" %>>Self</option>
            <option value="Others" <% If CStr(rs("asset_used_by"))="Others" Then Response.Write "selected" %>>Others</option>
        </select>
    </div>
</div>

<div class="actions">
    <button type="submit" class="btn primary">Update Asset</button>
    <a href="asset_report.asp" class="btn secondary">Cancel</a>
</div>
</form>
</div>
<%
rs.Close
Set rs = Nothing
Set cmd = Nothing
conn.Close
Set conn = Nothing
%>
</body>
</html>
