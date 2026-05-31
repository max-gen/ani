Attribute VB_Name = "mod_LoadTestData"
'========================================
' AZANI ISP MANAGEMENT SYSTEM
' Test Data Loading Module
' ========================================

Option Explicit

'========================================
' Function: LoadAllTestData
' Description: Load all 20 Kenyan test institutions with complete data
'========================================
Public Function LoadAllTestData() As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As Object
    Dim result As Integer
    
    Set db = CurrentDb
    
    result = MsgBox("This will load 20 test institutions. Continue?", vbYesNo, "Load Test Data")
    
    If result = vbNo Then
        Exit Function
    End If
    
    ' Load Contact Persons
    LoadTestContactPersons
    
    ' Load Institutions
    LoadTestInstitutions
    
    ' Load Infrastructure Requirements
    LoadTestInfrastructure
    
    ' Load Payments
    LoadTestPayments
    
    ' Load Monthly Services
    LoadTestMonthlyServices
    
    MsgBox "Test data loaded successfully!", vbInformation, "Success"
    LoadAllTestData = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error loading test data: " & Err.Description, vbCritical, "Error"
    LoadAllTestData = False
End Function

'========================================
' Function: LoadTestContactPersons
' Description: Load 20 test contact persons from Kenyan institutions
'========================================
Private Sub LoadTestContactPersons()
    On Error GoTo ErrorHandler
    
    Dim db As Object
    Dim sqlQuery As String
    
    Set db = CurrentDb
    
    ' Clear existing contact persons
    On Error Resume Next
    db.Execute "DELETE FROM tbl_ContactPersons"
    On Error GoTo ErrorHandler
    
    ' Insert 20 test contact persons
    Dim contactData() As String
    contactData = Array( _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('John', 'Kariuki', 'ICT Director', 'john.kariuki@school.ac.ke', '0701234567', 'ICT Department')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Mary', 'Nyambura', 'Principal', 'mary.nyambura@school.ac.ke', '0702345678', 'Administration')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('David', 'Muturi', 'ICT Coordinator', 'david.muturi@academy.ac.ke', '0703456789', 'ICT Department')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Sarah', 'Kipchoge', 'Deputy Principal', 'sarah.kipchoge@school.ac.ke', '0704567890', 'Administration')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Peter', 'Omondi', 'Director of Studies', 'peter.omondi@college.ac.ke', '0705678901', 'Academic')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Grace', 'Rotich', 'ICT Manager', 'grace.rotich@school.ac.ke', '0706789012', 'ICT Department')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('James', 'Kiplagat', 'School Administrator', 'james.kiplagat@school.ac.ke', '0707890123', 'Administration')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Patricia', 'Mwangi', 'Principal', 'patricia.mwangi@academy.ac.ke', '0708901234', 'Administration')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Michael', 'Kipkemboi', 'ICT Head', 'michael.kipkemboi@school.ac.ke', '0709012345', 'ICT Department')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Catherine', 'Mwangi', 'College Principal', 'catherine.mwangi@college.ac.ke', '0710123456', 'Administration')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Joseph', 'Kipchoge', 'ICT Specialist', 'joseph.kipchoge@school.ac.ke', '0711234567', 'ICT Department')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Lucy', 'Karanja', 'Deputy Principal', 'lucy.karanja@academy.ac.ke', '0712345678', 'Administration')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Daniel', 'Kiplagat', 'Systems Administrator', 'daniel.kiplagat@school.ac.ke', '0713456789', 'ICT Department')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Ruth', 'Omondi', 'Principal', 'ruth.omondi@school.ac.ke', '0714567890', 'Administration')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Steven', 'Koech', 'ICT Technician', 'steven.koech@college.ac.ke', '0715678901', 'ICT Department')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Victoria', 'Kipkemboi', 'Head of School', 'victoria.kipkemboi@academy.ac.ke', '0716789012', 'Administration')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Charles', 'Mwangi', 'Director', 'charles.mwangi@school.ac.ke', '0717890123', 'Administration')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Anne', 'Kipchoge', 'ICT Coordinator', 'anne.kipchoge@college.ac.ke', '0718901234', 'ICT Department')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Robert', 'Kiplagat', 'Network Manager', 'robert.kiplagat@school.ac.ke', '0719012345', 'ICT Department')", _
        "INSERT INTO tbl_ContactPersons (FirstName, LastName, JobTitle, Email, PhoneNumber, Department) VALUES ('Susan', 'Nyambura', 'College Director', 'susan.nyambura@academy.ac.ke', '0720123456', 'Administration')" _
    )
    
    Dim i As Integer
    For i = LBound(contactData) To UBound(contactData)
        db.Execute contactData(i)
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading contact persons: " & Err.Description, vbCritical, "Error"
End Sub

'========================================
' Function: LoadTestInstitutions
' Description: Load 20 test institutions
'========================================
Private Sub LoadTestInstitutions()
    On Error GoTo ErrorHandler
    
    Dim db As Object
    Dim sqlQuery As String
    
    Set db = CurrentDb
    
    ' Clear existing institutions
    On Error Resume Next
    db.Execute "DELETE FROM tbl_Institutions"
    On Error GoTo ErrorHandler
    
    ' Array of institution data: Name, Type, RegNumber, ContactPersonID, Address, County, Phone, BandwidthID, Infrastructure Ready
    Dim institutionData As String
    
    institutionData = "INSERT INTO tbl_Institutions (InstitutionName, InstitutionType, RegistrationNumber, ContactPersonID, PhysicalAddress, County, PhoneNumber, Email, CurrentBandwidthID, InfrastructureReady, NumComputers, NumLANNodes, Status, DateRegistered) VALUES " & _
    "('Nairobi Academy', 'Secondary School', 'REG001', 2, 'Nairobi', 'Nairobi', '0702345678', 'contact@nairobiaca.ac.ke', 2, TRUE, 45, 5, 'Active', NOW()), " & _
    "('St. Andrew''s Secondary', 'Secondary School', 'REG002', 4, 'Karen, Nairobi', 'Nairobi', '0704567890', 'contact@standrews.ac.ke', 3, TRUE, 60, 8, 'Active', NOW()), " & _
    "('Mama Ngina Primary', 'Primary School', 'REG003', 1, 'Westlands, Nairobi', 'Nairobi', '0701234567', 'contact@mamangina.ac.ke', 1, FALSE, 0, 0, 'Pending', NOW()), " & _
    "('Kenya High School', 'Secondary School', 'REG004', 6, 'Kilimani, Nairobi', 'Nairobi', '0706789012', 'contact@kenyahigh.ac.ke', 2, TRUE, 35, 4, 'Active', NOW()), " & _
    "('Mombasa Institute', 'College', 'REG005', 10, 'Mombasa', 'Mombasa', '0710123456', 'contact@mombasainst.ac.ke', 4, TRUE, 80, 12, 'Active', NOW()), " & _
    "('Dar-Es-Salaam School', 'Secondary School', 'REG006', 14, 'Kisumu', 'Kisumu', '0714567890', 'contact@daressalaam.ac.ke', 2, TRUE, 40, 5, 'Active', NOW()), " & _
    "('Aga Khan Academy', 'Secondary School', 'REG007', 3, 'Parklands, Nairobi', 'Nairobi', '0703456789', 'contact@agakhan.ac.ke', 5, TRUE, 100, 15, 'Active', NOW()), " & _
    "('Alliance High School', 'Secondary School', 'REG008', 7, 'Kikuyu, Nairobi', 'Nairobi', '0707890123', 'contact@alliancehigh.ac.ke', 3, TRUE, 70, 10, 'Active', NOW()), " & _
    "('Kamba Primary', 'Primary School', 'REG009', 9, 'Machakos', 'Machakos', '0709012345', 'contact@kambaprimary.ac.ke', 1, FALSE, 0, 0, 'Pending', NOW()), " & _
    "('Kenyatta University College', 'College', 'REG010', 5, 'Thika Road, Nairobi', 'Nairobi', '0705678901', 'contact@kenyattacol.ac.ke', 4, TRUE, 120, 18, 'Active', NOW()), " & _
    "('Starehe Boys Centre', 'Secondary School', 'REG011', 11, 'Nairobi', 'Nairobi', '0711234567', 'contact@starehboys.ac.ke', 3, TRUE, 65, 9, 'Active', NOW()), " & _
    "('St. Mary''s Secondary', 'Secondary School', 'REG012', 13, 'Eldoret', 'Uasin Gishu', '0713456789', 'contact@stmaryseldn.ac.ke', 2, TRUE, 50, 6, 'Active', NOW()), " & _
    "('Nyeri Polytechnic', 'College', 'REG013', 8, 'Nyeri', 'Nyeri', '0708901234', 'contact@nyeripoly.ac.ke', 3, TRUE, 55, 8, 'Active', NOW()), " & _
    "('Juja Primary School', 'Primary School', 'REG014', 12, 'Juja', 'Kiambu', '0712345678', 'contact@jujaprim.ac.ke', 1, TRUE, 20, 3, 'Active', NOW()), " & _
    "('Nairobi School', 'Secondary School', 'REG015', 15, 'Nairobi', 'Nairobi', '0715678901', 'contact@nairobi school.ac.ke', 4, TRUE, 85, 12, 'Active', NOW()), " & _
    "('Bishop McQueen High', 'Secondary School', 'REG016', 17, 'Kisii', 'Kisii', '0717890123', 'contact@bmqueen.ac.ke', 2, TRUE, 42, 5, 'Active', NOW()), " & _
    "('Kiambu Technical Institute', 'College', 'REG017', 16, 'Kiambu', 'Kiambu', '0716789012', 'contact@kiambutech.ac.ke', 3, TRUE, 75, 11, 'Active', NOW()), " & _
    "('Mbaraka Secondary', 'Secondary School', 'REG018', 18, 'Nakuru', 'Nakuru', '0718901234', 'contact@mbarakasec.ac.ke', 2, TRUE, 38, 5, 'Active', NOW()), " & _
    "('Lamu Island Primary', 'Primary School', 'REG019', 19, 'Lamu', 'Lamu', '0719012345', 'contact@lamuprim.ac.ke', 1, FALSE, 0, 0, 'Pending', NOW()), " & _
    "('Moi University Centre', 'College', 'REG020', 20, 'Eldoret', 'Uasin Gishu', '0720123456', 'contact@moiunictr.ac.ke', 5, TRUE, 110, 16, 'Active', NOW())"
    
    db.Execute institutionData
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading institutions: " & Err.Description, vbCritical, "Error"
End Sub

'========================================
' Function: LoadTestInfrastructure
' Description: Load infrastructure requirements for institutions
'========================================
Private Sub LoadTestInfrastructure()
    On Error GoTo ErrorHandler
    
    Dim db As Object
    Dim sqlQuery As String
    
    Set db = CurrentDb
    
    ' Clear existing infrastructure
    On Error Resume Next
    db.Execute "DELETE FROM tbl_Institution_Infrastructure"
    On Error GoTo ErrorHandler
    
    ' Insert infrastructure for institutions that required it
    Dim i As Integer
    For i = 2 To 20 Step 2
        ' Add PCs for institutions needing infrastructure
        sqlQuery = "INSERT INTO tbl_Institution_Infrastructure (InstitutionID, InfrastructureType, Quantity, UnitCost, TotalCost, DatePurchased, Status) " & _
                   "VALUES (" & i & ", 'Personal Computer', 5, 40000, 200000, NOW(), 'Active')"
        db.Execute sqlQuery
        
        ' Add LAN Nodes
        sqlQuery = "INSERT INTO tbl_Institution_Infrastructure (InstitutionID, InfrastructureType, Quantity, UnitCost, TotalCost, DatePurchased, Status) " & _
                   "VALUES (" & i & ", 'LAN Nodes', 3, 10000, 30000, NOW(), 'Active')"
        db.Execute sqlQuery
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading infrastructure: " & Err.Description, vbCritical, "Error"
End Sub

'========================================
' Function: LoadTestPayments
' Description: Load test payment records
'========================================
Private Sub LoadTestPayments()
    On Error GoTo ErrorHandler
    
    Dim db As Object
    Dim sqlQuery As String
    Dim i As Integer
    
    Set db = CurrentDb
    
    ' Clear existing payments
    On Error Resume Next
    db.Execute "DELETE FROM tbl_Payments"
    On Error GoTo ErrorHandler
    
    ' Insert registration fees for all institutions
    For i = 1 To 20
        sqlQuery = "INSERT INTO tbl_Payments (InstitutionID, PaymentType, AmountDue, AmountPaid, DueDate, Status, PaymentDate, ReceiptNumber, VerifiedBy) " & _
                   "VALUES (" & i & ", 'Registration Fee', 8500, " & IIf(i Mod 3 = 0, 0, 8500) & ", #" & Format(Date - 30, "MM/DD/YYYY") & "#, " & _
                   "'" & IIf(i Mod 3 = 0, "Pending", "Paid") & "', " & IIf(i Mod 3 = 0, "NULL", "#" & Format(Date - 25, "MM/DD/YYYY") & "#") & ", " & _
                   "'RCP" & i & "001', 'System')"
        db.Execute sqlQuery
        
        ' Insert installation fees for infrastructure-ready institutions
        If i Mod 2 = 0 Then
            sqlQuery = "INSERT INTO tbl_Payments (InstitutionID, PaymentType, AmountDue, AmountPaid, DueDate, Status, PaymentDate, ReceiptNumber, VerifiedBy) " & _
                       "VALUES (" & i & ", 'Installation Fee', 10000, " & IIf(i Mod 4 = 0, 0, 10000) & ", #" & Format(Date - 20, "MM/DD/YYYY") & "#, " & _
                       "'" & IIf(i Mod 4 = 0, "Pending", "Paid") & "', " & IIf(i Mod 4 = 0, "NULL", "#" & Format(Date - 15, "MM/DD/YYYY") & "#") & ", " & _
                       "'RCP" & i & "002', 'System')"
            db.Execute sqlQuery
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading payments: " & Err.Description, vbCritical, "Error"
End Sub

'========================================
' Function: LoadTestMonthlyServices
' Description: Load monthly service billing records
'========================================
Private Sub LoadTestMonthlyServices()
    On Error GoTo ErrorHandler
    
    Dim db As Object
    Dim sqlQuery As String
    Dim i As Integer
    Dim currentMonth As Integer
    Dim currentYear As Integer
    
    Set db = CurrentDb
    
    ' Clear existing monthly services
    On Error Resume Next
    db.Execute "DELETE FROM tbl_Monthly_Services"
    On Error GoTo ErrorHandler
    
    currentMonth = Month(Now())
    currentYear = Year(Now())
    
    ' Insert monthly service records for all active institutions
    For i = 1 To 20
        sqlQuery = "INSERT INTO tbl_Monthly_Services (InstitutionID, BillingMonth, BillingYear, CurrentBandwidthID, BaseMonthlyCharge, UpgradeDiscount, FinalMonthlyCharge, BillingDate, DueDate, Status) " & _
                   "VALUES (" & i & ", " & currentMonth & ", " & currentYear & ", " & IIf(i <= 5, 1, IIf(i <= 10, 2, IIf(i <= 15, 3, 4))) & ", " & _
                   IIf(i <= 5, 1200, IIf(i <= 10, 2000, IIf(i <= 15, 3500, 4000))) & ", 0, " & _
                   IIf(i <= 5, 1200, IIf(i <= 10, 2000, IIf(i <= 15, 3500, 4000))) & ", #" & Format(Date, "MM/DD/YYYY") & "#, #" & _
                   Format(DateSerial(currentYear, currentMonth + 1, 30), "MM/DD/YYYY") & "#, '" & IIf(i Mod 2 = 0, "Paid", "Pending") & "')"
        db.Execute sqlQuery
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading monthly services: " & Err.Description, vbCritical, "Error"
End Sub
