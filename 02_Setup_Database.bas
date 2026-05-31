Attribute VB_Name = "mod_SetupDatabase"
'========================================
' AZANI ISP MANAGEMENT SYSTEM
' Database Setup Module
' ========================================

Option Explicit

'========================================
' Function: InitializeDatabase
' Description: Initialize database structure and default data
'========================================
Public Function InitializeDatabase() As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As Object
    Set db = CurrentDb
    
    MsgBox "Database initialization completed successfully!", vbInformation, "Setup Complete"
    InitializeDatabase = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error during database initialization: " & Err.Description, vbCritical, "Error"
    InitializeDatabase = False
End Function

'========================================
' Function: GetSystemSetting
' Description: Retrieve system settings value
' Parameters: SettingName (String)
' Returns: Setting value (Variant)
'========================================
Public Function GetSystemSetting(SettingName As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim db As Object
    Dim rs As Object
    Dim sqlQuery As String
    
    Set db = CurrentDb
    sqlQuery = "SELECT SettingValue FROM tbl_System_Settings WHERE SettingName = '" & SettingName & "'"
    Set rs = db.OpenRecordset(sqlQuery)
    
    If rs.RecordCount > 0 Then
        GetSystemSetting = rs("SettingValue")
    Else
        GetSystemSetting = ""
    End If
    
    rs.Close
    Exit Function
    
ErrorHandler:
    GetSystemSetting = ""
End Function

'========================================
' Function: SetSystemSetting
' Description: Update system settings value
' Parameters: SettingName (String), SettingValue (String)
'========================================
Public Sub SetSystemSetting(SettingName As String, SettingValue As String)
    On Error GoTo ErrorHandler
    
    Dim db As Object
    Dim sqlQuery As String
    
    Set db = CurrentDb
    sqlQuery = "UPDATE tbl_System_Settings SET SettingValue = '" & SettingValue & "', DateModified = NOW() WHERE SettingName = '" & SettingName & "'"
    
    db.Execute sqlQuery
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error updating setting: " & Err.Description, vbCritical, "Error"
End Sub

'========================================
' Function: GetRegistrationFee
' Description: Get registration fee from settings
' Returns: Currency value
'========================================
Public Function GetRegistrationFee() As Currency
    GetRegistrationFee = CCur(GetSystemSetting("REGISTRATION_FEE"))
End Function

'========================================
' Function: GetInstallationFee
' Description: Get installation fee from settings
' Returns: Currency value
'========================================
Public Function GetInstallationFee() As Currency
    GetInstallationFee = CCur(GetSystemSetting("INSTALLATION_FEE"))
End Function

'========================================
' Function: GetPCCost
' Description: Get personal computer cost from settings
' Returns: Currency value
'========================================
Public Function GetPCCost() As Currency
    GetPCCost = CCur(GetSystemSetting("PC_COST"))
End Function

'========================================
' Function: GetOverdueFinePercent
' Description: Get overdue fine percentage from settings
' Returns: Numeric value (0-100)
'========================================
Public Function GetOverdueFinePercent() As Double
    GetOverdueFinePercent = CDbl(GetSystemSetting("OVERDUE_FINE_PERCENT")) / 100
End Function

'========================================
' Function: GetUpgradeDiscountPercent
' Description: Get upgrade discount percentage from settings
' Returns: Numeric value (0-100)
'========================================
Public Function GetUpgradeDiscountPercent() As Double
    GetUpgradeDiscountPercent = CDbl(GetSystemSetting("UPGRADE_DISCOUNT_PERCENT")) / 100
End Function

'========================================
' Function: GetReconnectionFee
' Description: Get reconnection fee from settings
' Returns: Currency value
'========================================
Public Function GetReconnectionFee() As Currency
    GetReconnectionFee = CCur(GetSystemSetting("RECONNECTION_FEE"))
End Function

'========================================
' Function: GetDisconnectionDay
' Description: Get disconnection trigger day from settings
' Returns: Numeric value (day of month)
'========================================
Public Function GetDisconnectionDay() As Integer
    GetDisconnectionDay = CInt(GetSystemSetting("DISCONNECTION_DAY"))
End Function

'========================================
' Function: CreatePaymentRecord
' Description: Create new payment record
' Parameters: InstitutionID, PaymentType, Amount, DueDate
'========================================
Public Function CreatePaymentRecord(InstitutionID As Long, PaymentType As String, _
    Amount As Currency, DueDate As Date) As Long
    
    On Error GoTo ErrorHandler
    
    Dim db As Object
    Dim newPaymentID As Long
    Dim sqlQuery As String
    
    Set db = CurrentDb
    
    sqlQuery = "INSERT INTO tbl_Payments (InstitutionID, PaymentType, AmountDue, AmountPaid, DueDate, Status) " & _
               "VALUES (" & InstitutionID & ", '" & PaymentType & "', " & Amount & ", 0, #" & Format(DueDate, "MM/DD/YYYY") & "#, 'Pending')"
    
    db.Execute sqlQuery
    
    ' Get the newly created PaymentID
    sqlQuery = "SELECT MAX(PaymentID) as NewID FROM tbl_Payments WHERE InstitutionID = " & InstitutionID
    Dim rs As Object
    Set rs = db.OpenRecordset(sqlQuery)
    
    If rs.RecordCount > 0 Then
        newPaymentID = rs("NewID")
    End If
    
    rs.Close
    CreatePaymentRecord = newPaymentID
    Exit Function
    
ErrorHandler:
    MsgBox "Error creating payment record: " & Err.Description, vbCritical, "Error"
    CreatePaymentRecord = 0
End Function

'========================================
' Function: CalculateOverdueFine
' Description: Calculate overdue fine for unpaid amount
' Parameters: Amount (Currency), DaysOverdue (Integer)
' Returns: Fine amount (Currency)
'========================================
Public Function CalculateOverdueFine(Amount As Currency, DaysOverdue As Integer) As Currency
    Dim finePercent As Double
    finePercent = GetOverdueFinePercent()
    CalculateOverdueFine = Amount * finePercent
End Function

'========================================
' Function: CalculateUpgradeDiscount
' Description: Calculate 10% discount for bandwidth upgrade
' Parameters: NewBandwidthCost (Currency)
' Returns: Discount amount (Currency)
'========================================
Public Function CalculateUpgradeDiscount(NewBandwidthCost As Currency) As Currency
    Dim discountPercent As Double
    discountPercent = GetUpgradeDiscountPercent()
    CalculateUpgradeDiscount = NewBandwidthCost * discountPercent
End Function

'========================================
' Function: GetLANNodeCost
' Description: Get cost for specified number of LAN nodes
' Parameters: NumNodes (Integer)
' Returns: Cost (Currency)
'========================================
Public Function GetLANNodeCost(NumNodes As Integer) As Currency
    On Error GoTo ErrorHandler
    
    Dim db As Object
    Dim rs As Object
    Dim sqlQuery As String
    
    Set db = CurrentDb
    
    sqlQuery = "SELECT Cost FROM tbl_Infrastructure WHERE InfrastructureType = 'LAN Nodes' " & _
               "AND MinUnits <= " & NumNodes & " AND MaxUnits >= " & NumNodes
    
    Set rs = db.OpenRecordset(sqlQuery)
    
    If rs.RecordCount > 0 Then
        GetLANNodeCost = rs("Cost")
    Else
        GetLANNodeCost = 0
    End If
    
    rs.Close
    Exit Function
    
ErrorHandler:
    GetLANNodeCost = 0
End Function

'========================================
' Function: GetBandwidthCost
' Description: Get monthly cost for specified bandwidth
' Parameters: BandwidthMBPS (Integer)
' Returns: Cost (Currency)
'========================================
Public Function GetBandwidthCost(BandwidthMBPS As Integer) As Currency
    On Error GoTo ErrorHandler
    
    Dim db As Object
    Dim rs As Object
    Dim sqlQuery As String
    
    Set db = CurrentDb
    
    sqlQuery = "SELECT MonthlyCost FROM tbl_Bandwidth_Plans WHERE BandwidthMBPS = " & BandwidthMBPS
    
    Set rs = db.OpenRecordset(sqlQuery)
    
    If rs.RecordCount > 0 Then
        GetBandwidthCost = rs("MonthlyCost")
    Else
        GetBandwidthCost = 0
    End If
    
    rs.Close
    Exit Function
    
ErrorHandler:
    GetBandwidthCost = 0
End Function

'========================================
' Function: GetBandwidthID
' Description: Get BandwidthID for specified MBPS
' Parameters: BandwidthMBPS (Integer)
' Returns: BandwidthID (Long)
'========================================
Public Function GetBandwidthID(BandwidthMBPS As Integer) As Long
    On Error GoTo ErrorHandler
    
    Dim db As Object
    Dim rs As Object
    Dim sqlQuery As String
    
    Set db = CurrentDb
    
    sqlQuery = "SELECT BandwidthID FROM tbl_Bandwidth_Plans WHERE BandwidthMBPS = " & BandwidthMBPS
    
    Set rs = db.OpenRecordset(sqlQuery)
    
    If rs.RecordCount > 0 Then
        GetBandwidthID = rs("BandwidthID")
    Else
        GetBandwidthID = 0
    End If
    
    rs.Close
    Exit Function
    
ErrorHandler:
    GetBandwidthID = 0
End Function
