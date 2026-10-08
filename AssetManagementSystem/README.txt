IT ASSET MANAGEMENT SYSTEM
Classic ASP + IIS + MySQL + ODBC + SQLyog

1. DATABASE
-----------
Open SQLyog.
Connect to your MySQL Server.
Open schema_mysql.sql.
Execute the complete file.

This creates:
    AssetManagementDB
    users
    assets

Important:
MAC Address is the PRIMARY KEY of assets.

Demo login:
    Staff ID: BEL001
    Password: 01011990

2. ODBC
--------
Install a 64-bit MySQL ODBC driver if IIS is 64-bit.
Open:
    ODBC Data Sources (64-bit)

The application uses a DRIVER connection string, so a DSN is not required.

3. DB CONNECTION
----------------
Open:
    includes\db.asp

If your driver is not:
    MySQL ODBC 8.0 Unicode Driver

change it to the exact driver name installed on the computer.

Also enter your MySQL root password if one is configured.

Example:
    DRIVER={MySQL ODBC 8.0 Unicode Driver};
    SERVER=localhost;
    PORT=3306;
    DATABASE=AssetManagementDB;
    USER=root;
    PASSWORD=your_password;
    OPTION=3;

4. IIS
-------
Enable:
    Internet Information Services
    World Wide Web Services
    Application Development Features
    ASP

Open IIS Manager.

Create an Application/Virtual Directory pointing to this project folder.

Example:
    C:\inetpub\wwwroot\AssetManagementSystem

Open:
    ASP

Set:
    Enable Parent Paths = True

Browse:
    http://localhost/AssetManagementSystem/login.asp

5. PAGES
--------
login.asp
    Staff login.

index.asp
    Main menu.

asset_declaration.asp
    Declare/register an asset.

asset_report.asp
    Search, view, edit, delete and export CSV.

asset_edit.asp
    Edit an existing asset.

asset_delete.asp
    Delete an asset.

6. IMPORTANT
------------
This project is Classic ASP, not ASP.NET.

The website does not require SQLyog to be running.
SQLyog is only used to manage the MySQL database.

The runtime path is:

Browser
  -> IIS
  -> Classic ASP
  -> ADO
  -> MySQL ODBC Driver
  -> MySQL Server

MAC Address is the asset primary key and cannot be changed from the edit page.

ADDED ASSET FIELDS
------------------
Zone, Location, Hostname, Domain Name, Operating System, Year of Manufacture,
Type of Network (Internet/Intranet/Standalone), Warranty (Yes/No), Staff ID,
and Cost Center.

Operating System options:
Windows 11, Windows 12, Windows Server 2012, Windows Server 2016,
Windows Server 2019, Windows Server 2022.

After updating an existing database, run the updated schema_mysql.sql from a
fresh database (the schema drops/recreates the tables), or apply equivalent
ALTER TABLE statements to preserve existing records.


Asset field order:
Staff ID -> Cost Center -> Asset Number -> Asset Type -> Make -> Model -> Serial Number -> MAC Address -> IP Address -> Hostname -> Domain Name -> RAM (GB) -> Hard Disk (GB) -> Hard Disk Used (GB) -> Operating System -> Year of Manufacture -> Type of Network -> Warranty -> Zone -> Location -> Asset Used By
Operating System options: Windows 11, Windows 12, Windows Server 2012, Windows Server 2016, Windows Server 2019, Windows Server 2022, Thin OS.
Network options: Internet, Intranet, Standalone. Warranty options: Yes, No.
