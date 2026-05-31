# Azani ISP Management System

## Overview
A comprehensive Microsoft Access VBA-based database management system for Azani Internet Service Provider, handling institution registration, service management, billing, and reporting.

## Features
- Institution registration and contact person details
- Bandwidth and service package management
- Infrastructure requirements tracking (PCs, LAN nodes)
- Payment tracking (registration, installation, monthly)
- Automatic overdue fine calculation (15%)
- Service disconnection management
- Comprehensive reporting system
- Modern forms with validation
- Complete audit trail and transaction history

## Database Structure

### Tables
1. **tbl_Institutions** - Institution details and service information
2. **tbl_ContactPersons** - Contact person information
3. **tbl_Bandwidth_Plans** - Internet service bandwidth plans
4. **tbl_Infrastructure** - PC and LAN node pricing
5. **tbl_Institution_Infrastructure** - Infrastructure per institution
6. **tbl_Payments** - Payment tracking (Registration, Installation, Monthly)
7. **tbl_Monthly_Services** - Monthly service subscriptions
8. **tbl_Disconnections** - Disconnection and reconnection records
9. **tbl_System_Settings** - System configuration

### Forms
- frm_InstitutionRegistration
- frm_PaymentEntry
- frm_MonthlyBilling
- frm_Infrastructure
- frm_Dashboard

### Reports
- rpt_RegisteredInstitutions
- rpt_Defaulters
- rpt_DisconnectionList
- rpt_InfrastructureRequirements
- rpt_BillingAnalysis
- rpt_MonthlyCharges

## Test Data
20 Kenyan educational institutions with complete sample data included.

## Installation
1. Open MS Access
2. Import the database structure from `01_Database_Schema.sql`
3. Run `02_Setup_Database.bas` to initialize
4. Run `03_Load_TestData.bas` to populate sample data
5. Configure system settings in frm_Settings

## File Structure
```
ani/
├── README.md
├── 01_Database_Schema.sql          # Table definitions and relationships
├── 02_Setup_Database.bas            # Database initialization module
├── 03_Load_TestData.bas             # Test data population
├── 04_BusinessLogic.bas             # Core calculations and validations
├── 05_PaymentProcessing.bas         # Payment handling module
├── 06_BillingEngine.bas             # Monthly billing automation
├── 07_ReportingFunctions.bas        # Report generation functions
├── 08_Forms/                        # All form designs and code
│   ├── frm_Dashboard.frm
│   ├── frm_InstitutionRegistration.frm
│   ├── frm_PaymentEntry.frm
│   ├── frm_MonthlyBilling.frm
│   └── frm_Infrastructure.frm
└── 09_Reports/                      # All report definitions
    ├── rpt_RegisteredInstitutions.rpt
    ├── rpt_Defaulters.rpt
    ├── rpt_DisconnectionList.rpt
    ├── rpt_InfrastructureRequirements.rpt
    └── rpt_BillingAnalysis.rpt
```

## Pricing Information

### Registration Fee
- KSh 8,500 per institution

### Installation Fee
- KSh 10,000 (if infrastructure ready)

### Internet Service (Monthly)
| Bandwidth | Cost (KSh) |
|-----------|-----------|
| 4 MBPS    | 1,200     |
| 10 MBPS   | 2,000     |
| 20 MBPS   | 2,000     |
| 25 MBPS   | 3,500     |
| 50 MBPS   | 4,000     |
| 100 MBPS  | 7,000     |

### Infrastructure Costs
#### Personal Computers
- KSh 40,000 per unit

#### LAN Nodes
| Quantity | Cost (KSh) |
|----------|-----------|
| 2-10     | 10,000    |
| 11-20    | 20,000    |
| 21-40    | 30,000    |
| 41-100   | 40,000    |

### Penalties & Fees
- Upgrade Discount: 10% off new bandwidth cost
- Overdue Fine: 15% of unpaid amount
- Reconnection Fee: KSh 1,000 (after payment of fine)
- Disconnection Trigger: Non-payment by 10th of subsequent month

## System Requirements
- Microsoft Access 2016 or later
- Windows OS
- VBA enabled
- Minimum 100 MB free disk space

## Version
1.0.0 - Initial Release

## Author
Azani ISP Management Team
