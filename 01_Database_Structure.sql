-- ============================================================================
-- AZANI ISP MANAGEMENT SYSTEM - DATABASE STRUCTURE
-- ============================================================================
-- This script creates all necessary tables for the Azani ISP system
-- Import into MS Access using File > External Data > Import

-- ============================================================================
-- TABLE 1: tbl_Institutions
-- ============================================================================
CREATE TABLE tbl_Institutions (
    InstitutionID AutoIncrement PRIMARY KEY,
    InstitutionName Text(100) NOT NULL,
    RegistrationNumber Text(50) UNIQUE,
    InstitutionType Text(20),
    City Text(50),
    Region Text(50),
    PhoneNumber Text(15),
    Email Text(100),
    PhysicalAddress Text(200),
    RegistrationFeeAmount Currency DEFAULT 8500,
    RegistrationFeePaid YesNo DEFAULT No,
    RegistrationFeeDate DateTime,
    InstallationFeeAmount Currency DEFAULT 10000,
    InstallationFeePaid YesNo DEFAULT No,
    InstallationFeeDate DateTime,
    InstallationStatus Text(20),
    DateRegistered DateTime DEFAULT Now(),
    DateInstalled DateTime,
    AccountStatus Text(20) DEFAULT 'Active',
    Notes Memo,
    CreatedDate DateTime DEFAULT Now(),
    ModifiedDate DateTime
);

-- ============================================================================
-- TABLE 2: tbl_ContactPersons
-- ============================================================================
CREATE TABLE tbl_ContactPersons (
    ContactPersonID AutoIncrement PRIMARY KEY,
    InstitutionID Long NOT NULL,
    FirstName Text(50) NOT NULL,
    LastName Text(50) NOT NULL,
    Title Text(50),
    PhoneNumber Text(15),
    Email Text(100),
    Department Text(50),
    IsPrimary YesNo DEFAULT Yes,
    CreatedDate DateTime DEFAULT Now(),
    ModifiedDate DateTime,
    CONSTRAINT FK_ContactPerson_Institution FOREIGN KEY (InstitutionID) REFERENCES tbl_Institutions(InstitutionID)
);

-- ============================================================================
-- TABLE 3: tbl_Bandwidth_Plans
-- ============================================================================
CREATE TABLE tbl_Bandwidth_Plans (
    BandwidthID AutoIncrement PRIMARY KEY,
    BandwidthMBPS Long NOT NULL UNIQUE,
    MonthlyCostKSh Currency NOT NULL,
    Description Text(200),
    IsActive YesNo DEFAULT Yes,
    CreatedDate DateTime DEFAULT Now()
);

-- ============================================================================
-- TABLE 4: tbl_InfrastructurePricing
-- ============================================================================
CREATE TABLE tbl_InfrastructurePricing (
    InfraID AutoIncrement PRIMARY KEY,
    ItemType Text(50) NOT NULL,
    MinQuantity Long,
    MaxQuantity Long,
    CostPerUnitKSh Currency,
    Description Text(200),
    CreatedDate DateTime DEFAULT Now()
);

-- ============================================================================
-- TABLE 5: tbl_Institution_Infrastructure
-- ============================================================================
CREATE TABLE tbl_Institution_Infrastructure (
    InfraDetailsID AutoIncrement PRIMARY KEY,
    InstitutionID Long NOT NULL,
    ItemType Text(50),
    Quantity Long,
    UnitCostKSh Currency,
    TotalCostKSh Currency,
    DateRequired DateTime,
    IsPaid YesNo DEFAULT No,
    PaymentDate DateTime,
    Notes Text(200),
    CreatedDate DateTime DEFAULT Now(),
    CONSTRAINT FK_InfraDetails_Institution FOREIGN KEY (InstitutionID) REFERENCES tbl_Institutions(InstitutionID)
);

-- ============================================================================
-- TABLE 6: tbl_Monthly_Services
-- ============================================================================
CREATE TABLE tbl_Monthly_Services (
    ServiceID AutoIncrement PRIMARY KEY,
    InstitutionID Long NOT NULL,
    BandwidthID Long NOT NULL,
    MonthYear Text(7),
    BandwidthMBPS Long,
    MonthlyCostKSh Currency,
    IsUpgrade YesNo DEFAULT No,
    UpgradeDiscountPercent Currency DEFAULT 0,
    UpgradeDiscountAmount Currency DEFAULT 0,
    FinalCostKSh Currency,
    IsPaid YesNo DEFAULT No,
    PaymentDate DateTime,
    PreviousBandwidthID Long,
    ActivationDate DateTime,
    CreatedDate DateTime DEFAULT Now(),
    CONSTRAINT FK_Service_Institution FOREIGN KEY (InstitutionID) REFERENCES tbl_Institutions(InstitutionID),
    CONSTRAINT FK_Service_Bandwidth FOREIGN KEY (BandwidthID) REFERENCES tbl_Bandwidth_Plans(BandwidthID)
);

-- ============================================================================
-- TABLE 7: tbl_Payments
-- ============================================================================
CREATE TABLE tbl_Payments (
    PaymentID AutoIncrement PRIMARY KEY,
    InstitutionID Long NOT NULL,
    PaymentType Text(50) NOT NULL,
    Amount Currency NOT NULL,
    AmountPaid Currency,
    AmountOutstanding Currency,
    PaymentDate DateTime,
    DueDate DateTime,
    OverdueFineAmount Currency DEFAULT 0,
    OverdueFineRate Currency DEFAULT 0.15,
    FineAppliedDate DateTime,
    FinePaidDate DateTime,
    ReconnectionFeeAmount Currency DEFAULT 0,
    ReconnectionFeeCharged YesNo DEFAULT No,
    ReconnectionFeePaidDate DateTime,
    PaymentStatus Text(20),
    PaymentMethod Text(50),
    TransactionReference Text(100),
    ServiceID Long,
    Notes Memo,
    CreatedDate DateTime DEFAULT Now(),
    ModifiedDate DateTime,
    CONSTRAINT FK_Payment_Institution FOREIGN KEY (InstitutionID) REFERENCES tbl_Institutions(InstitutionID)
);

-- ============================================================================
-- TABLE 8: tbl_Disconnections
-- ============================================================================
CREATE TABLE tbl_Disconnections (
    DisconnectionID AutoIncrement PRIMARY KEY,
    InstitutionID Long NOT NULL,
    DisconnectionDate DateTime,
    DisconnectionReason Text(200),
    OverdueDays Long,
    AmountDueKSh Currency,
    OverdueFinePaidDate DateTime,
    ReconnectionDate DateTime,
    ReconnectionFeePaidDate DateTime,
    DisconnectionStatus Text(20),
    Notes Memo,
    CreatedDate DateTime DEFAULT Now(),
    ModifiedDate DateTime,
    CONSTRAINT FK_Disconnection_Institution FOREIGN KEY (InstitutionID) REFERENCES tbl_Institutions(InstitutionID)
);

-- ============================================================================
-- TABLE 9: tbl_SystemSettings
-- ============================================================================
CREATE TABLE tbl_SystemSettings (
    SettingID AutoIncrement PRIMARY KEY,
    SettingName Text(100) UNIQUE NOT NULL,
    SettingValue Text(255),
    SettingDescription Memo,
    SettingType Text(20),
    ModifiedDate DateTime DEFAULT Now()
);

-- ============================================================================
-- INSERT DEFAULT BANDWIDTH PLANS
-- ============================================================================
INSERT INTO tbl_Bandwidth_Plans (BandwidthMBPS, MonthlyCostKSh, Description, IsActive)
VALUES 
(4, 1200, '4 MBPS Internet Service', Yes),
(10, 1200, '10 MBPS Internet Service', Yes),
(20, 2000, '20 MBPS Internet Service', Yes),
(25, 3500, '25 MBPS Internet Service', Yes),
(50, 4000, '50 MBPS Internet Service', Yes);

-- ============================================================================
-- INSERT INFRASTRUCTURE PRICING
-- ============================================================================
INSERT INTO tbl_InfrastructurePricing (ItemType, MinQuantity, MaxQuantity, CostPerUnitKSh, Description)
VALUES 
('PersonalComputer', 1, 999, 40000, 'Desktop/Laptop Computer'),
('LANNode', 2, 10, 10000, 'LAN Nodes (2-10)'),
('LANNode', 11, 20, 20000, 'LAN Nodes (11-20)'),
('LANNode', 21, 40, 30000, 'LAN Nodes (21-40)'),
('LANNode', 41, 100, 40000, 'LAN Nodes (41-100)');

-- ============================================================================
-- INSERT DEFAULT SYSTEM SETTINGS
-- ============================================================================
INSERT INTO tbl_SystemSettings (SettingName, SettingValue, SettingDescription, SettingType)
VALUES 
('REGISTRATION_FEE', '8500', 'Standard registration fee in KSh', 'Currency'),
('INSTALLATION_FEE', '10000', 'Standard installation fee in KSh', 'Currency'),
('OVERDUE_FINE_RATE', '0.15', 'Overdue fine rate (15%)', 'Currency'),
('RECONNECTION_FEE', '1000', 'Service reconnection fee in KSh', 'Currency'),
('PC_COST', '40000', 'Cost per Personal Computer in KSh', 'Currency'),
('PAYMENT_DUE_DAY', '25', 'Day of month when payment is due', 'Integer'),
('DISCONNECTION_DAY', '10', 'Day of month for service disconnection if unpaid', 'Integer'),
('COMPANY_NAME', 'AZANI INTERNET SERVICE PROVIDER', 'Company name', 'Text'),
('COMPANY_EMAIL', 'info@azani.co.ke', 'Company contact email', 'Text'),
('COMPANY_PHONE', '+254712345678', 'Company contact phone', 'Text');
