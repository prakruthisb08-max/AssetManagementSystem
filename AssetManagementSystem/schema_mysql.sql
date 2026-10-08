-- ============================================================
-- IT Asset Management System
-- MySQL database for SQLyog
-- ============================================================

-- Create database with matching character set and collation
CREATE DATABASE IF NOT EXISTS AssetManagementDB
CHARACTER SET utf8mb4
COLLATE utf8mb4_unicode_ci;

USE AssetManagementDB;

-- ============================================================
-- DROP OLD TABLES
-- ============================================================

DROP TABLE IF EXISTS assets;
DROP TABLE IF EXISTS users;

-- ============================================================
-- USERS TABLE
-- ============================================================

CREATE TABLE users (
    staff_id VARCHAR(50) NOT NULL,
    full_name VARCHAR(100) NOT NULL,
    password_value VARCHAR(50) NOT NULL,
    is_active TINYINT(1) NOT NULL DEFAULT 1,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,

    PRIMARY KEY (staff_id)

) ENGINE=INNODB
DEFAULT CHARSET=utf8mb4
COLLATE=utf8mb4_unicode_ci;

-- ============================================================
-- DEMO LOGIN
-- Staff ID: BEL001
-- Password: 01011990
-- ============================================================

INSERT INTO users
(
    staff_id,
    full_name,
    password_value
)
VALUES
(
    'BEL001',
    'Demo Staff',
    '01011990'
);

-- ============================================================
-- ASSETS TABLE
-- ============================================================

CREATE TABLE assets (

    -- 1. Staff ID
    staff_id VARCHAR(50) NOT NULL,

    -- 2. Cost Center
    cost_center VARCHAR(50) NOT NULL,

    -- 3. Asset Number
    asset_number VARCHAR(50) NOT NULL,

    -- 4. Asset Type
    asset_type VARCHAR(100) NOT NULL,

    -- 5. Make
    make VARCHAR(100) DEFAULT NULL,

    -- 6. Model
    model VARCHAR(100) DEFAULT NULL,

    -- 7. Serial Number
    serial_no VARCHAR(100) DEFAULT NULL,

    -- 8. MAC Address
    mac_address VARCHAR(17) NOT NULL,

    -- 9. IP Address
    ip_address VARCHAR(45) DEFAULT NULL,

    -- 10. Hostname
    hostname VARCHAR(100) NOT NULL,

    -- 11. Domain Name
    domainname VARCHAR(150) NOT NULL,

    -- 12. RAM
    ram VARCHAR(50) DEFAULT NULL,

    -- 13. Hard Disk
    hard_disk VARCHAR(50) DEFAULT NULL,

    -- 14. Hard Disk Used
    hard_disk_used VARCHAR(50) DEFAULT NULL,

    -- 15. Operating System
    -- Application values:
    -- Windows 11
    -- Windows 12
    -- Windows Server 2012
    -- Windows Server 2016
    -- Windows Server 2019
    -- Windows Server 2022
    -- Thin OS
    operating_system VARCHAR(100) NOT NULL,

    -- 16. Year of Manufacture
    year_of_manufacture YEAR NOT NULL,

    -- 17. Type of Network
    -- Internet / Intranet / Standalone
    network_type VARCHAR(30) NOT NULL,

    -- 18. Warranty
    -- Yes / No
    warranty ENUM('Yes','No') NOT NULL,

    -- 19. Zone
    zone VARCHAR(100) NOT NULL,

    -- 20. Location
    location VARCHAR(150) NOT NULL,

    -- 21. Asset Used By
    asset_used_by VARCHAR(50) NOT NULL,

    -- System fields
    declared_by VARCHAR(50) DEFAULT NULL,
    updated_by VARCHAR(50) DEFAULT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,

    -- Primary key
    PRIMARY KEY (mac_address),

    -- Asset number must be unique
    UNIQUE KEY uq_asset_number (asset_number),

    -- Indexes
    INDEX idx_staff_id (staff_id),
    INDEX idx_cost_center (cost_center),
    INDEX idx_asset_type (asset_type),
    INDEX idx_serial_no (serial_no),
    INDEX idx_hostname (hostname),
    INDEX idx_ip_address (ip_address),
    INDEX idx_zone (zone),
    INDEX idx_location (location),

    -- Staff relationship
    CONSTRAINT fk_assets_staff_id
        FOREIGN KEY (staff_id)
        REFERENCES users(staff_id)
        ON UPDATE CASCADE
        ON DELETE RESTRICT,

    -- Declared by relationship
    CONSTRAINT fk_assets_declared_by
        FOREIGN KEY (declared_by)
        REFERENCES users(staff_id)
        ON UPDATE CASCADE
        ON DELETE SET NULL,

    -- Updated by relationship
    CONSTRAINT fk_assets_updated_by
        FOREIGN KEY (updated_by)
        REFERENCES users(staff_id)
        ON UPDATE CASCADE
        ON DELETE SET NULL

) ENGINE=INNODB
DEFAULT CHARSET=utf8mb4
COLLATE=utf8mb4_unicode_ci;

-- ============================================================
-- SAMPLE ASSET
-- ============================================================

INSERT INTO assets
(
    staff_id,
    cost_center,
    asset_number,
    asset_type,
    make,
    model,
    serial_no,
    mac_address,
    ip_address,
    hostname,
    domainname,
    ram,
    hard_disk,
    hard_disk_used,
    operating_system,
    year_of_manufacture,
    network_type,
    warranty,
    zone,
    location,
    asset_used_by,
    declared_by
)
VALUES
(
    'BEL001',
    'CC001',
    'AST001',
    'Desktop PC - Standard',
    'Dell',
    'OptiPlex',
    'DEMO123',
    '00:11:22:33:44:55',
    '192.168.1.10',
    'BEL-PC-001',
    'bel.local',
    '16',
    '512',
    '200',
    'Windows 11',
    2024,
    'Intranet',
    'Yes',
    'Zone A',
    'BEL Office',
    'Self',
    'BEL001'
);

-- ============================================================
-- CHECK DATABASE
-- ============================================================

SELECT DATABASE();

SHOW TABLES;

-- Check users
SELECT * FROM users;

-- Check assets
SELECT * FROM assets;

-- Check asset columns
DESCRIBE assets;