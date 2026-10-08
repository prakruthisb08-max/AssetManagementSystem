<%@ Language="VBScript" %>
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/db.asp"-->
<%
Option Explicit
Dim message, errorMessage
message = ""
errorMessage = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim assetNumber, assetType, makeName, modelName, serialNo
    Dim macAddress, ipAddress, ramValue, hardDiskValue, hardDiskUsedValue, assetUsedBy
    Dim zoneValue, locationValue, hostnameValue, domainNameValue, operatingSystemValue
    Dim manufactureYearValue, networkTypeValue, warrantyValue, staffIdValue, costCenterValue
    Dim cmd, rs, sql, diskTotal, diskUsed

    assetNumber = Trim(Request.Form("asset_number"))
    assetType = Trim(Request.Form("asset_type"))
    makeName = Trim(Request.Form("make"))
    modelName = Trim(Request.Form("model"))
    serialNo = Trim(Request.Form("serial_no"))
    macAddress = LCase(Replace(Replace(Trim(Request.Form("mac_address")), "-", ":"), " ", ""))
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
    If staffIdValue = "" Then staffIdValue = Session("staffId")

    If assetNumber = "" Then
        errorMessage = "Asset Number is required."
    ElseIf assetType = "" Then
        errorMessage = "Asset Type is required."
    ElseIf zoneValue = "" Then
        errorMessage = "Zone is required."
    ElseIf locationValue = "" Then
        errorMessage = "Location is required."
    ElseIf hostnameValue = "" Then
        errorMessage = "Hostname is required."
    ElseIf domainNameValue = "" Then
        errorMessage = "Domain Name is required."
    ElseIf operatingSystemValue = "" Then
        errorMessage = "Operating System is required."
    ElseIf manufactureYearValue = "" Then
        errorMessage = "Year of Manufacture is required."
    ElseIf networkTypeValue = "" Then
        errorMessage = "Type of Network is required."
    ElseIf warrantyValue = "" Then
        errorMessage = "Warranty is required."
    ElseIf staffIdValue = "" Then
        errorMessage = "Staff ID is required."
    ElseIf costCenterValue = "" Then
        errorMessage = "Cost Center is required."
    ElseIf macAddress = "" Then
        errorMessage = "MAC Address is required."
    ElseIf Not IsValidMac(macAddress) Then
        errorMessage = "Enter a valid MAC Address, for example 00:11:22:33:44:55."
    ElseIf hardDiskValue <> "" And hardDiskUsedValue <> "" And IsNumeric(hardDiskValue) And IsNumeric(hardDiskUsedValue) Then
        diskTotal = CDbl(hardDiskValue)
        diskUsed = CDbl(hardDiskUsedValue)
        If diskUsed > diskTotal Then errorMessage = "Hard Disk Used cannot be greater than Hard Disk capacity."
    End If

    If errorMessage = "" Then
        sql = "SELECT mac_address FROM assets WHERE mac_address = ?"
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
        cmd.Parameters.Append cmd.CreateParameter("p8", 200, 1, 50, macAddress)
        cmd.Parameters.Append cmd.CreateParameter("p9", 200, 1, 45, ipAddress)
        cmd.Parameters.Append cmd.CreateParameter("p10", 200, 1, 100, hostnameValue)
        cmd.Parameters.Append cmd.CreateParameter("p11", 200, 1, 150, domainNameValue)
        cmd.Parameters.Append cmd.CreateParameter("p12", 200, 1, 50, ramValue)
        cmd.Parameters.Append cmd.CreateParameter("p13", 200, 1, 50, hardDiskValue)
        cmd.Parameters.Append cmd.CreateParameter("p14", 200, 1, 50, hardDiskUsedValue)
        cmd.Parameters.Append cmd.CreateParameter("p15", 200, 1, 100, operatingSystemValue)
        cmd.Parameters.Append cmd.CreateParameter("p16", 3, 1, 0, CLng(manufactureYearValue))
        cmd.Parameters.Append cmd.CreateParameter("p17", 200, 1, 30, networkTypeValue)
        cmd.Parameters.Append cmd.CreateParameter("p18", 200, 1, 3, warrantyValue)
        cmd.Parameters.Append cmd.CreateParameter("p19", 200, 1, 100, zoneValue)
        cmd.Parameters.Append cmd.CreateParameter("p20", 200, 1, 150, locationValue)
        cmd.Parameters.Append cmd.CreateParameter("p21", 200, 1, 20, assetUsedBy)
        cmd.Parameters.Append cmd.CreateParameter("p22", 200, 1, 50, Session("staffId"))

        On Error Resume Next
        cmd.Execute
        If Err.Number <> 0 Then
            errorMessage = "Could not save the asset: " & Err.Description
            Err.Clear
        Else
            message = "Asset declared successfully."
        End If
        On Error GoTo 0

        Set cmd = Nothing
    End If
End If

Function IsValidMac(mac)
    Dim re
    Set re = New RegExp
    re.Pattern = "^[0-9a-fA-F]{2}(:[0-9a-fA-F]{2}){5}$"
    re.IgnoreCase = True
    IsValidMac = re.Test(mac)
    Set re = Nothing
End Function
%>
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Asset Declaration</title>
<link rel="stylesheet" href="wwwroot/css/style.css">
<script src="wwwroot/js/script.js"></script>
</head>
<body>
<div class="topbar">
    <strong>IT Asset Management</strong>
    <div><a href="index.asp">Home</a> | <a href="asset_report.asp">Asset Report</a> | <a href="logout.asp">Logout</a></div>
</div>

<div class="container form-container">
<h1>Asset Declaration</h1>

<% If message <> "" Then %><div class="alert success"><%=Server.HTMLEncode(message)%></div><% End If %>
<% If errorMessage <> "" Then %><div class="alert error"><%=Server.HTMLEncode(errorMessage)%></div><% End If %>

<form method="post" action="asset_declaration.asp" onsubmit="return validateAssetForm();">
<div class="form-grid">
    <div><label>Staff ID *</label><input type="text" name="staff_id" maxlength="50" value="<%=Server.HTMLEncode(staffIdValue)%>" required></div>
    <div><label>Cost Center *</label><input type="text" name="cost_center" maxlength="50" value="<%=Server.HTMLEncode(costCenterValue)%>" required></div>

    <div><label>Asset Number *</label><input type="text" name="asset_number" maxlength="50" required></div>
    <div>
        <label>Asset Type *</label>
        <select name="asset_type" id="asset_type" required>
            <option value="">Select Asset Type</option>
            <option>Desktop PC - High End</option><option>Desktop PC - Standard</option>
            <option>Laptop - High End</option><option>Laptop - Standard</option><option>Laptop - Medium</option>
            <option>Workstation - High</option><option>Workstation - Standard</option><option>Workstation - Medium</option>
            <option>Printer</option><option>Scanner</option><option>Server</option><option>Other</option>
        </select>
    </div>
    <div><label>Make</label><input type="text" name="make" maxlength="100"></div>
    <div><label>Model</label><input type="text" name="model" maxlength="100"></div>
    <div><label>Serial Number</label><input type="text" name="serial_no" maxlength="100"></div>
    <div><label>MAC Address *</label><input type="text" name="mac_address" maxlength="17" placeholder="00:11:22:33:44:55" required><small>MAC Address is the PRIMARY KEY.</small></div>
    <div><label>IP Address</label><input type="text" name="ip_address" maxlength="45"></div>
    <div><label>Hostname *</label><input type="text" name="hostname" maxlength="100" required></div>
    <div><label>Domain Name *</label><input type="text" name="domainname" maxlength="150" required></div>
    <div><label>RAM (GB)</label><input type="number" name="ram" min="0" step="0.01" maxlength="50" placeholder="e.g. 16"></div>
    <div><label>Hard Disk (GB)</label><input type="number" name="hard_disk" min="0" step="0.01" maxlength="50" placeholder="e.g. 512"></div>
    <div><label>Hard Disk Used (GB)</label><input type="number" name="hard_disk_used" min="0" step="0.01" maxlength="50" placeholder="e.g. 200"></div>

    <div>
        <label>Operating System *</label>
        <select name="operating_system" required>
            <option value="">Select Operating System</option>
            <option>Windows 11</option><option>Windows 12</option>
            <option>Windows Server 2012</option><option>Windows Server 2016</option>
            <option>Windows Server 2019</option><option>Windows Server 2022</option><option>Thin OS</option>
        </select>
    </div>
    <div>
        <label>Year of Manufacture *</label>
        <select name="year_of_manufacture" required><option value="">Select Year</option>
        <% Dim yy : For yy=Year(Date) To 1990 Step -1
            Response.Write "<option value="" & yy & "">" & yy & "</option>"
        Next %></select>
    </div>
    <div>
        <label>Type of Network *</label>
        <select name="network_type" required>
            <option value="">Select Network Type</option><option>Internet</option><option>Intranet</option><option>Standalone</option>
        </select>
    </div>
    <div>
        <label>Warranty *</label>
        <select name="warranty" required><option value="">Select</option><option>Yes</option><option>No</option></select>
    </div>
    <div><label>Zone *</label><input type="text" name="zone" maxlength="100" required></div>
    <div><label>Location *</label><input type="text" name="location" maxlength="150" required></div>
    <div>
        <label>Asset Used By *</label>
        <select name="asset_used_by" required><option value="">Select</option><option value="Self">Self</option><option value="Others">Others</option></select>
    </div>
</div>

<div class="actions">
    <button type="submit" class="btn primary">Save Asset</button>
    <a href="index.asp" class="btn secondary">Cancel</a>
</div>
</form>
</div>
</body>
</html>
