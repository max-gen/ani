-- ========================================
-- AZANI ISP MANAGEMENT SYSTEM
-- Database Schema Definition
-- ========================================

-- ========================================
-- TABLE 1: tbl_System_Settings
-- ========================================
CREATE TABLE tbl_System_Settings (
    SettingID AUTOINCREMENT PRIMARY KEY,
    SettingName TEXT NOT NULL,
    SettingValue TEXT,
    Description TEXT,
    DataType TEXT,
    DateCreated DATETIME DEFAULT NOW(),
    DateModified DATETIME DEFAULT NOW()
);

-- INSERT DEFAULT SETTINGS
INSERT INTO tbl_System_Settings (SettingName, SettingValue, Description, DataType)
VALUES 
    ('REGISTRATION_FEE', '8500', 'Institution Registration Fee (KSh)', 'Currency'),
    ('INSTALLATION_FEE', '10000', 'Installation Fee (KSh)', 'Currency'),
    ('PC_COST', '40000', 'Cost per Personal Computer (KSh)', 'Currency'),
    ('OVERDUE_FINE_PERCENT', '15', 'Overdue Fine Percentage (%)', 'Percentage'),
    ('UPGRADE_DISCOUNT_PERCENT', '10', 'Upgrade Discount Percentage (%)', 'Percentage'),
    ('RECONNECTION_FEE', '1000', 'Service Reconnection Fee (KSh)', 'Currency'),
    ('DISCONNECTION_DAY', '10', 'Disconnection Day in Next Month', 'Number'),
    ('COMPANY_NAME', 'Azani ISP', 'Company Name', 'Text'),
    ('COMPANY_ADDRESS', 'Nairobi, Kenya', 'Company Address', 'Text'),
    ('COMPANY_PHONE', '+254 700 000 000', 'Company Phone', 'Text'),
    ('COMPANY_EMAIL', 'info@azani.co.ke', 'Company Email', 'Text');

-- ========================================
-- TABLE 2: tbl_Bandwidth_Plans
-- ========================================
CREATE TABLE tbl_Bandwidth_Plans (
    BandwidthID AUTOINCREMENT PRIMARY KEY,
    BandwidthMBPS INT NOT NULL UNIQUE,
    MonthlyCost CURRENCY NOT NULL,
    Description TEXT,
    Status TEXT DEFAULT 'Active',
    DateCreated DATETIME DEFAULT NOW(),
    DateModified DATETIME DEFAULT NOW()
);

-- INSERT BANDWIDTH PLANS
INSERT INTO tbl_Bandwidth_Plans (BandwidthMBPS, MonthlyCost, Description, Status)
VALUES 
    (4, 1200, 'Basic - 4 MBPS', 'Active'),
    (10, 2000, 'Standard - 10 MBPS', 'Active'),
    (20, 2000, 'Standard - 20 MBPS', 'Active'),
    (25, 3500, 'Premium - 25 MBPS', 'Active'),
    (50, 4000, 'Premium Plus - 50 MBPS', 'Active'),
    (100, 7000, 'Enterprise - 100 MBPS', 'Active');

-- ========================================
-- TABLE 3: tbl_Infrastructure
-- ========================================
CREATE TABLE tbl_Infrastructure (
    InfrastructureID AUTOINCREMENT PRIMARY KEY,
    InfrastructureType TEXT NOT NULL,
    MinUnits INT,
    MaxUnits INT,
    Cost CURRENCY NOT NULL,
    DateCreated DATETIME DEFAULT NOW(),
    DateModified DATETIME DEFAULT NOW()
);

-- INSERT LAN NODE PRICING
INSERT INTO tbl_Infrastructure (InfrastructureType, MinUnits, MaxUnits, Cost)
VALUES 
    ('LAN Nodes', 2, 10, 10000),
    ('LAN Nodes', 11, 20, 20000),
    ('LAN Nodes', 21, 40, 30000),
    ('LAN Nodes', 41, 100, 40000);

-- ========================================
-- TABLE 4: tbl_ContactPersons
-- ========================================
CREATE TABLE tbl_ContactPersons (
    ContactPersonID AUTOINCREMENT PRIMARY KEY,
    FirstName TEXT NOT NULL,
    LastName TEXT NOT NULL,
    JobTitle TEXT,
    Email TEXT,
    PhoneNumber TEXT NOT NULL,
    AlternatePhone TEXT,
    Department TEXT,
    DateCreated DATETIME DEFAULT NOW(),
    DateModified DATETIME DEFAULT NOW()
);

-- ========================================
-- TABLE 5: tbl_Institutions
-- ========================================
CREATE TABLE tbl_Institutions (
    InstitutionID AUTOINCREMENT PRIMARY KEY,
    InstitutionName TEXT NOT NULL,
    InstitutionType TEXT NOT NULL,
    RegistrationNumber TEXT UNIQUE,
    ContactPersonID INT NOT NULL,
    PhysicalAddress TEXT NOT NULL,
    PostalAddress TEXT,
    County TEXT NOT NULL,
    PhoneNumber TEXT NOT NULL,
    Email TEXT,
    Website TEXT,
    InfrastructureReady YES/NO DEFAULT FALSE,
    NumComputers INT,
    NumLANNodes INT,
    CurrentBandwidthID INT,
    InstallationDate DATETIME,
    Status TEXT DEFAULT 'Active',
    RegistrationFeeID INT,
    InstallationFeeID INT,
    DateRegistered DATETIME DEFAULT NOW(),
    DateModified DATETIME DEFAULT NOW(),
    FOREIGN KEY (ContactPersonID) REFERENCES tbl_ContactPersons(ContactPersonID),
    FOREIGN KEY (CurrentBandwidthID) REFERENCES tbl_Bandwidth_Plans(BandwidthID)
);

-- ========================================
-- TABLE 6: tbl_Institution_Infrastructure
-- ========================================
CREATE TABLE tbl_Institution_Infrastructure (
    InstitutionInfraID AUTOINCREMENT PRIMARY KEY,
    InstitutionID INT NOT NULL,
    InfrastructureType TEXT NOT NULL,
    Quantity INT NOT NULL,
    UnitCost CURRENCY NOT NULL,
    TotalCost CURRENCY NOT NULL,
    DatePurchased DATETIME,
    Status TEXT DEFAULT 'Active',
    DateCreated DATETIME DEFAULT NOW(),
    DateModified DATETIME DEFAULT NOW(),
    FOREIGN KEY (InstitutionID) REFERENCES tbl_Institutions(InstitutionID)
);

-- ========================================
-- TABLE 7: tbl_Payments
-- ========================================
CREATE TABLE tbl_Payments (
    PaymentID AUTOINCREMENT PRIMARY KEY,
    InstitutionID INT NOT NULL,
    PaymentType TEXT NOT NULL,
    AmountDue CURRENCY NOT NULL,
    AmountPaid CURRENCY NOT NULL,
    PaymentDate DATETIME,
    PaymentMethod TEXT,
    TransactionReference TEXT,
    ReceiptNumber TEXT UNIQUE,
    Notes TEXT,
    Status TEXT DEFAULT 'Pending',
    DueDate DATETIME,
    OverdueFine CURRENCY DEFAULT 0,
    FineAppliedDate DATETIME,
    ReconnectionFee CURRENCY DEFAULT 0,
    ReconnectionDate DATETIME,
    VerifiedBy TEXT,
    DateCreated DATETIME DEFAULT NOW(),
    DateModified DATETIME DEFAULT NOW(),
    FOREIGN KEY (InstitutionID) REFERENCES tbl_Institutions(InstitutionID)
);

-- ========================================
-- TABLE 8: tbl_Monthly_Services
-- ========================================
CREATE TABLE tbl_Monthly_Services (
    MonthlyServiceID AUTOINCREMENT PRIMARY KEY,
    InstitutionID INT NOT NULL,
    BillingMonth INT NOT NULL,
    BillingYear INT NOT NULL,
    CurrentBandwidthID INT NOT NULL,
    PreviousBandwidthID INT,
    IsUpgrade YES/NO DEFAULT FALSE,
    BaseMonthlyCharge CURRENCY NOT NULL,
    UpgradeDiscount CURRENCY DEFAULT 0,
    FinalMonthlyCharge CURRENCY NOT NULL,
    BillingDate DATETIME,
    DueDate DATETIME,
    PaymentID INT,
    Status TEXT DEFAULT 'Pending',
    DateCreated DATETIME DEFAULT NOW(),
    DateModified DATETIME DEFAULT NOW(),
    FOREIGN KEY (InstitutionID) REFERENCES tbl_Institutions(InstitutionID),
    FOREIGN KEY (CurrentBandwidthID) REFERENCES tbl_Bandwidth_Plans(BandwidthID),
    FOREIGN KEY (PreviousBandwidthID) REFERENCES tbl_Bandwidth_Plans(BandwidthID),
    FOREIGN KEY (PaymentID) REFERENCES tbl_Payments(PaymentID)
);

-- ========================================
-- TABLE 9: tbl_Disconnections
-- ========================================
CREATE TABLE tbl_Disconnections (
    DisconnectionID AUTOINCREMENT PRIMARY KEY,
    InstitutionID INT NOT NULL,
    DisconnectionDate DATETIME,
    DisconnectionReason TEXT,
    AmountOutstanding CURRENCY NOT NULL,
    OverdueFineAmount CURRENCY NOT NULL,
    ReconnectionFeeAmount CURRENCY DEFAULT 1000,
    TotalAmountDue CURRENCY NOT NULL,
    ReconnectionDate DATETIME,
    ReconnectionPaymentID INT,
    Status TEXT DEFAULT 'Disconnected',
    Notes TEXT,
    DateCreated DATETIME DEFAULT NOW(),
    DateModified DATETIME DEFAULT NOW(),
    FOREIGN KEY (InstitutionID) REFERENCES tbl_Institutions(InstitutionID),
    FOREIGN KEY (ReconnectionPaymentID) REFERENCES tbl_Payments(PaymentID)
);

-- ========================================
-- INDEXES FOR PERFORMANCE
-- ========================================
CREATE INDEX idx_Institution_Status ON tbl_Institutions(Status);
CREATE INDEX idx_Institution_County ON tbl_Institutions(County);
CREATE INDEX idx_Payment_Institution ON tbl_Payments(InstitutionID);
CREATE INDEX idx_Payment_Status ON tbl_Payments(Status);
CREATE INDEX idx_Payment_Date ON tbl_Payments(PaymentDate);
CREATE INDEX idx_Monthly_Status ON tbl_Monthly_Services(Status);
CREATE INDEX idx_Monthly_Date ON tbl_Monthly_Services(BillingDate);
CREATE INDEX idx_Disconnection_Status ON tbl_Disconnections(Status);
