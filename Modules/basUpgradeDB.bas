Attribute VB_Name = "basUpgradeDB"
Option Explicit


Public Sub UpgradeDB()
On Error GoTo ErrorTrap

    
    '
    'We've already determined that DB version is prior to App version...
    '
    If gstrDBVersion = "0005.0017.0000" Then
        DB_Upgrade_5_17_0_To_5_18_0
    End If

    If gstrDBVersion <= "0005.0018.0000" Then
        DB_Upgrade_5_18_0_To_5_19_0
    End If

    If gstrDBVersion <= "0005.0019.0000" Then
        DB_Upgrade_5_19_0_To_5_20_0
    End If

    If gstrDBVersion <= "0005.0020.0000" Then
        DB_Upgrade_5_20_0_To_5_20_1
    End If

    If gstrDBVersion <= "0005.0020.0001" Then
        DB_Upgrade_5_20_1_To_5_20_2
    End If

    If gstrDBVersion <= "0005.0020.0002" Then
        DB_Upgrade_5_20_2_To_5_20_3
    End If

    If gstrDBVersion <= "0005.0020.0003" Then
        DB_Upgrade_5_20_3_To_5_20_4
    End If

    If gstrDBVersion <= "0005.0020.0004" Then
        DB_Upgrade_5_20_4_To_5_20_5
    End If

    If gstrDBVersion <= "0005.0020.0005" Then
        DB_Upgrade_5_20_5_To_5_20_6
    End If

    If gstrDBVersion <= "0005.0020.0006" Then
        DB_Upgrade_5_20_6_To_5_20_7
    End If

    If gstrDBVersion <= "0005.0020.0007" Then
        DB_Upgrade_5_20_7_To_5_20_8
    End If

    If gstrDBVersion <= "0005.0020.0008" Then
        DB_Upgrade_5_20_8_To_5_20_9
    End If

    If gstrDBVersion <= "0005.0020.0009" Then
        DB_Upgrade_5_20_9_To_5_21_00
    End If

    If gstrDBVersion <= "0005.0021.0000" Then
        DB_Upgrade_5_21_00_To_5_22_00
    End If

    If gstrDBVersion <= "0005.0022.0000" Then
        DB_Upgrade_5_22_00_To_5_23_00
    End If

    If gstrDBVersion <= "0005.0023.0000" Then
        DB_Upgrade_5_23_00_To_5_24_00
    End If

    If gstrDBVersion <= "0005.0024.0000" Then
        DB_Upgrade_5_24_00_To_5_25_00
    End If

    If gstrDBVersion <= "0005.0025.0000" Then
        DB_Upgrade_5_25_00_To_5_26_00
    End If

    If gstrDBVersion <= "0005.0026.0000" Then
        DB_Upgrade_5_26_00_To_5_27_00
    End If

    If gstrDBVersion <= "0005.0027.0000" Then
        DB_Upgrade_5_27_00_To_5_28_00
    End If

    If gstrDBVersion <= "0005.0028.0000" Then
        DB_Upgrade_5_28_00_To_5_29_00
    End If

    If gstrDBVersion <= "0005.0029.0000" Then
        DB_Upgrade_5_29_00_To_5_30_00
    End If

    If gstrDBVersion <= "0005.0030.0000" Then
        DB_Upgrade_5_30_00_To_5_31_00
    End If

    If gstrDBVersion <= "0005.0031.0000" Then
        DB_Upgrade_5_31_00_To_5_32_00
    End If

    If gstrDBVersion <= "0005.0032.0000" Then
        DB_Upgrade_5_32_00_To_5_33_00
    End If

    If gstrDBVersion <= "0005.0033.0000" Then
        DB_Upgrade_5_33_00_To_5_34_00
    End If

    If gstrDBVersion <= "0005.0034.0000" Then
        DB_Upgrade_5_34_00_To_5_35_00
    End If

    If gstrDBVersion <= "0005.0035.0000" Then
        '
        'Virgin DB is at v5.36
        '
        DB_Upgrade_5_35_00_To_5_36_00
    End If

    If gstrDBVersion <= "0005.0036.0000" Then
        DB_Upgrade_5_36_00_To_5_37_00
    End If

    If gstrDBVersion <= "0005.0037.0000" Then
        DB_Upgrade_5_37_00_To_5_38_00
    End If

    If gstrDBVersion <= "0005.0038.0000" Then
        DB_Upgrade_5_38_00_To_5_39_00
    End If

    If gstrDBVersion <= "0005.0039.0000" Then
        DB_Upgrade_5_39_00_To_5_40_00
    End If

    If gstrDBVersion <= "0005.0040.0000" Then
        DB_Upgrade_5_40_00_To_5_41_00
    End If

    If gstrDBVersion <= "0005.0041.0000" Then
        DB_Upgrade_5_41_00_To_5_42_00
    End If

    If gstrDBVersion <= "0005.0042.0000" Then
        DB_Upgrade_5_42_00_To_5_43_00
    End If

    If gstrDBVersion <= "0005.0043.0000" Then
        DB_Upgrade_5_43_00_To_5_44_00
    End If

    If gstrDBVersion <= "0005.0044.0000" Then
        DB_Upgrade_5_44_00_To_5_45_00
    End If

    If gstrDBVersion <= "0005.0045.0000" Then
        DB_Upgrade_5_45_00_To_5_46_00
    End If

    If gstrDBVersion <= "0005.0046.0000" Then
        DB_Upgrade_5_46_00_To_5_47_00
    End If

    If gstrDBVersion <= "0005.0047.0000" Then
        DB_Upgrade_5_47_00_To_5_48_00
    End If

    If gstrDBVersion <= "0005.0048.0000" Then
        DB_Upgrade_5_48_00_To_5_49_00
    End If

    If gstrDBVersion <= "0005.0049.0000" Then
        DB_Upgrade_5_49_00_To_5_50_00
    End If

    If gstrDBVersion <= "0005.0050.0000" Then
        DB_Upgrade_5_50_00_To_5_51_00
    End If

    If gstrDBVersion <= "0005.0051.0000" Then
        DB_Upgrade_5_51_00_To_5_52_00
    End If

    If gstrDBVersion <= "0005.0052.0000" Then
        DB_Upgrade_5_52_00_To_5_53_00
    End If

    If gstrDBVersion <= "0005.0053.0000" Then
        DB_Upgrade_5_53_00_To_5_54_00
    End If

    If gstrDBVersion <= "0005.0054.0000" Then
        DB_Upgrade_5_54_00_To_5_55_00
    End If

    If gstrDBVersion <= "0005.0055.0000" Then
        DB_Upgrade_5_55_00_To_5_56_00
    End If

    If gstrDBVersion <= "0005.0056.0000" Then
        DB_Upgrade_5_56_00_To_5_57_00
    End If

    If gstrDBVersion <= "0005.0057.0000" Then
        DB_Upgrade_5_57_00_To_5_58_00
    End If

    If gstrDBVersion <= "0005.0058.0000" Then
        DB_Upgrade_5_58_00_To_5_59_00
    End If

    If gstrDBVersion <= "0005.0059.0000" Then
        DB_Upgrade_5_59_00_To_5_60_00
    End If

    If gstrDBVersion <= "0005.0060.0000" Then
        DB_Upgrade_5_60_00_To_5_61_00
    End If

    If gstrDBVersion <= "0005.0061.0000" Then
        DB_Upgrade_5_61_00_To_5_62_00
    End If
    
    Go_to_UpgradeDB2_module
    
    Exit Sub
    
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_17_0_To_5_18_0()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    '
    'Add new ExportItem6 and ImportItem6 to tblConstants
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('ExportItem6', " & _
                          " False, " & _
                          " 'Initial value = FALSE (TMS Roles only)')"

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('ImportItem6', " & _
                          " False, " & _
                          " 'Initial value = FALSE (TMS Roles only)')"

    '
    'Add new Export chkbox to security table for school overseer and CMS admin
    '
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmExportDB', " & _
                          " 'chkExportItem(6)', " & _
                          "1)"

    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmExportDB', " & _
                          " 'chkExportItem(6)', " & _
                          "4)"
   
    '
    'Now build new tblMinReports table
    '
    CreateTable ErrorCode, "tblMinReports", "PersonID", "LONG", , , False
    CreateField ErrorCode, "tblMinReports", "SocietyReportingMonth", "LONG"
    CreateField ErrorCode, "tblMinReports", "SocietyReportingYear", "LONG"
    CreateField ErrorCode, "tblMinReports", "MinistryDoneInMonth", "LONG"
    CreateField ErrorCode, "tblMinReports", "MinistryDoneInYear", "LONG"
    CreateField ErrorCode, "tblMinReports", "NoBooks", "LONG"
    CreateField ErrorCode, "tblMinReports", "NoBooklets", "LONG"
    CreateField ErrorCode, "tblMinReports", "NoHours", "LONG"
    CreateField ErrorCode, "tblMinReports", "NoMagazines", "LONG"
    CreateField ErrorCode, "tblMinReports", "NoReturnVisits", "LONG"
    CreateField ErrorCode, "tblMinReports", "NoStudies", "LONG"
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblMinReports " & _
                  "   (PersonID, " & _
                  "    SocietyReportingMonth, " & _
                  "    SocietyReportingYear, " & _
                  "    MinistryDoneInMonth, " & _
                  "    MinistryDoneInYear) " & _
                  "WITH PRIMARY"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.18.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_18_0_To_5_19_0()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add new Aux Pio Task
    '
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description) " & _
                  " VALUES (5, " & _
                          " 8, " & _
                          " 89, " & _
                          " 'Auxiliary Pioneer')"

    '
    'Create tblBaptismDates
    '
    CreateTable ErrorCode, "tblBaptismDates", "PersonID", "LONG", , , False
    CreateField ErrorCode, "tblBaptismDates", "BaptismDate", "DATE"
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblBaptismDates " & _
                  "   (PersonID) " & _
                  "WITH PRIMARY"
    
    '
    'Create tblMinTypeDesc
    '
    CreateTable ErrorCode, "tblMinTypeDesc", "MinType", "LONG", , , False
    CreateField ErrorCode, "tblMinTypeDesc", "MinTypeDesc", "TEXT"
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblMinTypeDesc " & _
                  "   (MinType) " & _
                  "WITH PRIMARY"
    
    
    '
    'Add lookup values to tblMinTypeDesc
    '
    CMSDB.Execute "INSERT INTO tblMinTypeDesc " & _
                  "(MinType, " & _
                  " MinTypeDesc) " & _
                  " VALUES (1, " & _
                          " 'Publisher')"
    CMSDB.Execute "INSERT INTO tblMinTypeDesc " & _
                  "(MinType, " & _
                  " MinTypeDesc) " & _
                  " VALUES (2, " & _
                          " 'Auxiliary Pioneer')"
    CMSDB.Execute "INSERT INTO tblMinTypeDesc " & _
                  "(MinType, " & _
                  " MinTypeDesc) " & _
                  " VALUES (3, " & _
                          " 'Regular Pioneer')"
    
    '
    'Create tblPublisherDates
    '
    CreateTable ErrorCode, "tblPublisherDates", "PersonID", "LONG"
    CreateField ErrorCode, "tblPublisherDates", "MinType", "LONG"
    CreateField ErrorCode, "tblPublisherDates", "StartDate", "DATE"
    CreateField ErrorCode, "tblPublisherDates", "EndDate", "DATE"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.19.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_19_0_To_5_20_0()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Create tblEldersAndServants
    '
    CreateTable ErrorCode, "tblEldersAndServants", "PersonID", "LONG", , , False
    CreateField ErrorCode, "tblEldersAndServants", "ElderOrServant", "TEXT", "2"
    CreateField ErrorCode, "tblEldersAndServants", "AppointmentDate", "DATE"
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblEldersAndServants " & _
                  "   (PersonID) " & _
                  "WITH PRIMARY"
    
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.20.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_20_0_To_5_20_1()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Amend tasks
    '
    CMSDB.Execute "UPDATE tblTasks " & _
                  "SET Description = 'School Assistant' " & _
                  "WHERE Task = 31"
                  
    CMSDB.Execute "DELETE FROM tblTasks " & _
                  "WHERE Task = 32"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.20.1"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_20_1_To_5_20_2()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add new Field Service button to security table for General Admin and CMS admin
    '
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdFieldMinistry', " & _
                          "1)"

    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdFieldMinistry', " & _
                          "5)"
        
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.20.2"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_20_2_To_5_20_3()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add new remarks column to tblMinReports
    '
    CreateField ErrorCode, "tblMinReports", "Remarks", "TEXT", "100"
        
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.20.3"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_20_3_To_5_20_4()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Day by which Min Reports must be sent to Society - add to tblConstants
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('DayOfMonthForReportToSociety', " & _
                          " 6, " & _
                          " 'Initial value = 6')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.20.4"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_20_4_To_5_20_5()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Infirmity level above which publisher can enter report in 15 minute increments
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('ThresholdForReportIn15MinInc', " & _
                          "2, " & _
                          " 'Initial value = 2')"
        
    '
    'Allow decimal values in Hours field of tblMinReport
    '
    CMSDB.Execute "ALTER TABLE tblMinReports " & _
                  "ALTER COLUMN NoHours SINGLE" & ";"

    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.20.5"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_20_5_To_5_20_6()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Allow decimal values in Hours field of tblMinReport
    '
    CMSDB.Execute "ALTER TABLE tblMinReports " & _
                  "ALTER COLUMN NoHours SINGLE;"

    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.20.6"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_20_6_To_5_20_7()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add column to tblMonths enabling months to be ordered for Service Year
    '
    CreateField ErrorCode, "tblMonthName", "OrderForServiceYear", "LONG"

    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForServiceYear = 1 " & _
                  "WHERE MonthNum = 9"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForServiceYear = 2 " & _
                  "WHERE MonthNum = 10"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForServiceYear = 3 " & _
                  "WHERE MonthNum = 11"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForServiceYear = 4 " & _
                  "WHERE MonthNum = 12"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForServiceYear = 5 " & _
                  "WHERE MonthNum = 1"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForServiceYear = 6 " & _
                  "WHERE MonthNum = 2"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForServiceYear = 7 " & _
                  "WHERE MonthNum = 3"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForServiceYear = 8 " & _
                  "WHERE MonthNum = 4"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForServiceYear = 9 " & _
                  "WHERE MonthNum = 5"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForServiceYear = 10 " & _
                  "WHERE MonthNum = 6"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForServiceYear = 11 " & _
                  "WHERE MonthNum = 7"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForServiceYear = 12 " & _
                  "WHERE MonthNum = 8"

    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.20.7"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_20_7_To_5_20_8()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add new date column to tblMinReports showing normal calendar date for
    ' actual ministry
    '
    CreateField ErrorCode, "tblMinReports", "ActualMinPeriod", "DATE"
        
    '
    'Add new date column to tblMinReports showing normal calendar date for
    ' society reporting period
    '
    CreateField ErrorCode, "tblMinReports", "SocietyReportingPeriod", "DATE"
        
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.20.8"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_20_8_To_5_20_9()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add new Book Group button to security table for CMS admin
    '
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdBookGroups', " & _
                          "1)"

    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdBookGroups', " & _
                          "5)"
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.20.9"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_20_9_To_5_21_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Number of months over which we determine if a publisher is irregular/inactive
    ' This will be 6 months.
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('NoMonthsForMinReporting', " & _
                          "6, " & _
                          " 'Initial value = 6')"
           
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.21.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_21_00_To_5_22_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rstTemp As Recordset, TheString As String

    '
    'Split out tblPublisherDates to separate tables for aux and reg pio
    '
    
    '
    'First create new tables
    '
    '
    'tblAuxPioDates
    '
    CreateTable ErrorCode, "tblAuxPioDates", "PersonID", "LONG"
    CreateField ErrorCode, "tblAuxPioDates", "StartDate", "DATE"
    CreateField ErrorCode, "tblAuxPioDates", "EndDate", "DATE"
    '
    'tblRegPioDates
    '
    CreateTable ErrorCode, "tblRegPioDates", "PersonID", "LONG"
    CreateField ErrorCode, "tblRegPioDates", "StartDate", "DATE"
    CreateField ErrorCode, "tblRegPioDates", "EndDate", "DATE"
    
    '
    'Now go through tblPublisherDates and copy any aux/reg data to new tables.
    ' Delete pio stuff from tblPublisherDates, then drop the MinType column
    '
    TheString = "SELECT PersonID, " & _
                "       StartDate, " & _
                "       EndDate, " & _
                "       MinType " & _
                "FROM tblPublisherDates"
    
    Set rstTemp = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)

    With rstTemp
    Do Until .EOF Or .BOF
        Select Case !MinType
        Case 2
            CMSDB.Execute "INSERT INTO tblAuxPioDates " & _
                              "(PersonID, StartDate, EndDate) " & _
                              "VALUES (" & !PersonID & ", #" & _
                                           Format(!StartDate, "mm/dd/yyyy") & "#, #" & _
                                           Format(!EndDate, "mm/dd/yyyy") & "#)"
        Case 3
            CMSDB.Execute "INSERT INTO tblRegPioDates " & _
                              "(PersonID, StartDate, EndDate) " & _
                              "VALUES (" & !PersonID & ", #" & _
                                           Format(!StartDate, "mm/dd/yyyy") & "#, #" & _
                                           Format(!EndDate, "mm/dd/yyyy") & "#)"
        End Select
        .MoveNext
    Loop
    End With
    
    Set rstTemp = Nothing
           
    CMSDB.Execute "DELETE FROM tblPublisherDates " & _
                  "WHERE MinType IN (2, 3)"
           
    DropField ErrorCode, "tblPublisherDates", "MinType"
            
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.22.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_22_00_To_5_23_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add reason codes to tblPublisherDates
    '
    CreateField ErrorCode, "tblPublisherDates", "StartReason", "LONG"
    CreateField ErrorCode, "tblPublisherDates", "EndReason", "LONG"
           
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.23.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_23_00_To_5_24_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rstTemp As Recordset, TheString As String
    
    '
    'Create new table to hold missing reports. Updated dynamically.
    '
    
    CreateTable ErrorCode, "tblMissingReports", "PersonID", "LONG"
    CreateField ErrorCode, "tblMissingReports", "ServiceYear", "LONG"
    CreateField ErrorCode, "tblMissingReports", "ServiceMonth", "LONG"
    CreateField ErrorCode, "tblMissingReports", "ActualMinDate", "DATE"
            
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.24.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_24_00_To_5_25_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rstTemp As Recordset, TheString As String
                   
    '
    'Store date on which the app was installed and first run
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " DateVal, " & _
                  " Comment) " & _
                  " VALUES ('AppFirstRunDate', " & _
                          "# " & Format(Now, "mm/dd/yyyy") & "#, " & _
                          " 'Initial value = BLANK')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.25.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_25_00_To_5_26_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rstTemp As Recordset, TheString As String
                   
    '
    'Store last date on which the app was run
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " DateVal, " & _
                  " Comment) " & _
                  " VALUES ('LastRunDate', " & _
                          "# " & Format(Now, "mm/dd/yyyy hh:mm:ss") & "#, " & _
                          " 'Initial value = BLANK')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.26.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_5_26_00_To_5_27_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rstTemp As Recordset, TheString As String
    
    '
    'Add field to hold seq-no for groups of missing reports.
    '
    
    CreateField ErrorCode, "tblMissingReports", "MissingReportGroupID", "LONG"
            
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.27.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_27_00_To_5_28_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    '
    'Create new table to hold inactive pubs. Updated dynamically.
    '
    
    CreateTable ErrorCode, "tblInactivePubs", "MissingReportGroupID", "LONG"
    CreateField ErrorCode, "tblInactivePubs", "StartDate", "DATE"
    CreateField ErrorCode, "tblInactivePubs", "EndDate", "DATE"
            
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.28.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_28_00_To_5_29_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    '
    'Create new table to hold irregular pubs. Updated dynamically.
    '
    
    CreateTable ErrorCode, "tblIrregularPubs", "PersonID", "LONG"
    CreateField ErrorCode, "tblIrregularPubs", "MinistryDate", "DATE"
                            
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.29.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_29_00_To_5_30_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    '
    'Modify key on tblMissingReports
    '
    '
    'Must delete data BEFORE removing key and creating new one,
    ' otherwise get duplicate probs
    '
    DelAllRows "tblMissingReports"
    DelAllRows "tblInactivePubs"
    DelAllRows "tblIrregularPubs"
    
    DropIndex ErrorCode, "tblMissingReports", "Constr"
    
    DropField ErrorCode, "tblMissingReports", "SeqNum"
    
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblMissingReports " & _
                  "   (PersonID, " & _
                  "    ActualMinDate)" & _
                  "WITH PRIMARY"
                   
            
'    PutAllMissingReportsIntoTable "01/09/2002", "01/12/9999"
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.30.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_30_00_To_5_31_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    '
    'Modify tblIrregularPubs
    '
    '
    'First remove all missing report-type data
    '
    DelAllRows "tblMissingReports"
    DelAllRows "tblInactivePubs"
    DelAllRows "tblIrregularPubs"
        
    DropField ErrorCode, "tblIrregularPubs", "MinistryDate"
    CreateField ErrorCode, "tblIrregularPubs", "IrregStartDate", "DATE"
    CreateField ErrorCode, "tblIrregularPubs", "IrregEndDate", "DATE"
            
'    PutAllMissingReportsIntoTable "01/09/2002", "01/12/9999"
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.31.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub DB_Upgrade_5_31_00_To_5_32_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    '
    'Modify tblIrregularPubs
    '
    '
    'First remove all missing report-type data
    '
    DelAllRows "tblMissingReports"
    DelAllRows "tblInactivePubs"
    DelAllRows "tblIrregularPubs"
        
    DropIndex ErrorCode, "tblIrregularPubs", "Constr"
    DropField ErrorCode, "tblIrregularPubs", "SeqNum"
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblIrregularPubs " & _
                  "   (PersonID, " & _
                  "    IrregStartDate)" & _
                  "WITH PRIMARY"
            
'    PutAllMissingReportsIntoTable "01/09/2002", "01/12/9999"
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.32.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_32_00_To_5_33_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    '
    'First remove all missing report-type data
    '
    DelAllRows "tblMissingReports"
    DelAllRows "tblInactivePubs"
    DelAllRows "tblIrregularPubs"
    
    '
    'Modify tblMissingReports
    '
    CreateField ErrorCode, "tblMissingReports", "ZeroReport", "YESNO"
            
    PutAllMissingReportsIntoTable "01/09/2002", "01/12/9999"
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.33.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_33_00_To_5_34_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    '
    'Advanced Field Min figures - allow access to function
    '
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmOptionsMenu', " & _
                          " 'cmdFieldMinAdvanced', " & _
                          "1)"
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.34.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_34_00_To_5_35_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    '
    'Export field min stuff
    '
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmExportDB', " & _
                          " 'chkExportItem(7)', " & _
                          "1)"
                          
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmExportDB', " & _
                          " 'chkExportItem(7)', " & _
                          "5)"
                
    CMSDB.Execute "INSERT INTO tblExportDetails " & _
                  "(ExportDataType, " & _
                  " OrderingForSQL, " & _
                  " IncludeForExport, " & _
                  " Description) " & _
                  " VALUES (8, " & _
                          " 800, " & _
                          "FALSE, " & _
                          "'All ministry related info')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.35.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_35_00_To_5_36_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
'
'Virgin DB is at v5.36
'
    '
    'Export field min stuff
    '
                
    CMSDB.Execute "INSERT INTO tblExportDetails " & _
                  "(ExportDataType, " & _
                  " OrderingForSQL, " & _
                  " IncludeForExport, " & _
                  " Description) " & _
                  " VALUES (7, " & _
                          " 700, " & _
                          "FALSE, " & _
                          "'TMS Roles ONLY')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.36.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Private Sub DB_Upgrade_5_36_00_To_5_37_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    '
    'Add new ExportItem7 and ImportItem7 to tblConstants
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('ExportItem7', " & _
                          " False, " & _
                          " 'Initial value = FALSE (Field ministry stuff)')"

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('ImportItem7', " & _
                          " False, " & _
                          " 'Initial value = FALSE (Field ministry stuff)')"

   
    '
    'Now build new meeting tables...
    '
    CreateTable ErrorCode, "tblMeetingTypes", "MeetingTypeID", "LONG", , , False
    CreateField ErrorCode, "tblMeetingTypes", "MeetingType", "TEXT"
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblMeetingTypes " & _
                  "   (MeetingTypeID) " & _
                  "WITH PRIMARY"
    
    CreateTable ErrorCode, "tblMeetingAttendance", "MeetingTypeID", "LONG", , , False
    CreateField ErrorCode, "tblMeetingAttendance", "WeekBeginning", "DATE"
    CreateField ErrorCode, "tblMeetingAttendance", "Attendance", "LONG"
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblMeetingAttendance " & _
                  "   (MeetingTypeID, " & _
                  "    WeekBeginning) " & _
                  "WITH PRIMARY"
    
    CreateTable ErrorCode, "tblGroupAttendance", "GroupNo", "LONG", , , False
    CreateField ErrorCode, "tblGroupAttendance", "WeekBeginning", "DATE"
    CreateField ErrorCode, "tblGroupAttendance", "Attendance", "LONG"
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblGroupAttendance " & _
                  "   (GroupNo, " & _
                  "    WeekBeginning) " & _
                  "WITH PRIMARY"
    
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.37.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_37_00_To_5_38_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                
    CMSDB.Execute "INSERT INTO tblMeetingTypes " & _
                  "(MeetingTypeID, " & _
                  " MeetingType) " & _
                  " VALUES (0, " & _
                          " 'Public Talk')"
    CMSDB.Execute "INSERT INTO tblMeetingTypes " & _
                  "(MeetingTypeID, " & _
                  " MeetingType) " & _
                  " VALUES (1, " & _
                          " 'Watchtower Study')"
    CMSDB.Execute "INSERT INTO tblMeetingTypes " & _
                  "(MeetingTypeID, " & _
                  " MeetingType) " & _
                  " VALUES (2, " & _
                          " 'Theocratic Ministry School')"
    CMSDB.Execute "INSERT INTO tblMeetingTypes " & _
                  "(MeetingTypeID, " & _
                  " MeetingType) " & _
                  " VALUES (3, " & _
                          " 'Service Meeting')"
    CMSDB.Execute "INSERT INTO tblMeetingTypes " & _
                  "(MeetingTypeID, " & _
                  " MeetingType) " & _
                  " VALUES (4, " & _
                          " 'Book Study')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.38.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_38_00_To_5_39_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdMeetingAttendance', " & _
                          "1)"
                          
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdMeetingAttendance', " & _
                          "5)"
                
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.39.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Private Sub DB_Upgrade_5_39_00_To_5_40_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rstTemp As Recordset, TheString As String
    
    '
    'Create new table for printing ministry related data by book-group.
    '
    
    CreateTable ErrorCode, "tblPrintCongMinByGroup", "GroupName", "TEXT", _
                , "", True
    CreateField ErrorCode, "tblPrintCongMinByGroup", "MonthAndServiceYear", "TEXT"
    CreateField ErrorCode, "tblPrintCongMinByGroup", "PersonName", "TEXT"

    CMSDB.TableDefs.Refresh

    'whoops.... rename a field.....
    CMSDB.TableDefs("tblPrintCongMinByGroup").Fields("MonthAndServiceYear").Name = "CalendarMonthAndYear"
    CMSDB.TableDefs.Refresh
            
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.40.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_40_00_To_5_41_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rstTemp As Recordset, TheString As String
    
    
    CreateField ErrorCode, "tblPrintCongMinByGroup", "GroupNo", "LONG"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.41.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_41_00_To_5_42_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rstTemp As Recordset, TheString As String
    
    
    DropField ErrorCode, "tblPrintCongMinByGroup", "GroupNo"
    DropField ErrorCode, "tblPrintCongMinByGroup", "GroupName"
    CreateField ErrorCode, "tblPrintCongMinByGroup", "PersonID", "LONG"
    DropField ErrorCode, "tblPrintCongMinByGroup", "PersonName"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.42.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_42_00_To_5_43_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rstTemp As Recordset, TheString As String
    
    
    '
    'Add new 'Person gets a km' Task
    '
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description) " & _
                  " VALUES (4, " & _
                          " 4, " & _
                          " 90, " & _
                          " 'Person gets a Kingdom Ministry')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.43.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_43_00_To_5_44_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rstTemp As Recordset, TheString As String
    
    '
    'Create new table for printing ministry related data by book-group.
    '
    
    CreateTable ErrorCode, "tblPrintBookGroups", "GroupID", "LONG", , "", True
    CreateField ErrorCode, "tblPrintBookGroups", "GroupName", "TEXT"
    CreateField ErrorCode, "tblPrintBookGroups", "PersonID", "LONG"
    CreateField ErrorCode, "tblPrintBookGroups", "PersonName", "TEXT"
    CreateField ErrorCode, "tblPrintBookGroups", "IsOverseer", "TEXT", "1"
    CreateField ErrorCode, "tblPrintBookGroups", "IsAssistant", "TEXT", "1"
    CreateField ErrorCode, "tblPrintBookGroups", "IsReader", "TEXT", "1"
    CreateField ErrorCode, "tblPrintBookGroups", "IsPrayer", "TEXT", "1"
    CreateField ErrorCode, "tblPrintBookGroups", "HasKM", "TEXT", "1"
    CreateField ErrorCode, "tblPrintBookGroups", "HasKM_Num", "LONG"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.44.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_44_00_To_5_45_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rstTemp As Recordset, TheString As String
    
    DestroyGlobalObjects
    
    CreateField ErrorCode, "tblNameAddress", "LinkedAddressPerson", "LONG"
    
    SetUpGlobalObjects
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.45.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_45_00_To_5_46_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rstTemp As Recordset, TheString As String
    
    DestroyGlobalObjects

    CreateField ErrorCode, "tblNameAddress", "Anointed", "YESNO"

    SetUpGlobalObjects
    
    CreateField ErrorCode, "tblRegPioDates", "PioneerNumber", "TEXT"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.46.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_5_46_00_To_5_47_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add new publisher record card print parameters to tblConstants
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPublisherNameXPos', " & _
                          " 0.8, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPublisherNameYPos', " & _
                          " 1.25, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAddressXPos', " & _
                          " 1.0, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAddressYPos', " & _
                          " 1.7, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPublisherNameMaxWidth', " & _
                          " 13.1, " & _
                          " 'In cm. ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAddressMaxWidth', " & _
                          " 12.9, " & _
                          " 'In cm.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTelNoXPos', " & _
                          " 1.25, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTelNoYPos', " & _
                          " 2.15, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTelNoMaxWidth', " & _
                          " 5.6, " & _
                          " 'In cm. ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardElderXPos', " & _
                          " 11.23, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardElderYPos', " & _
                          " 2.59, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardServantXPos', " & _
                          " 12.59, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardServantYPos', " & _
                          " 2.59, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardRegPioXPos', " & _
                          " 11.23, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardRegPioYPos', " & _
                          " 2.85, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardBaptDateXPos', " & _
                          " 2.8, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardBaptDateYPos', " & _
                          " 2.8, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardBaptDateMaxWidth', " & _
                          " 4, " & _
                          " 'In cm. ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPioNoXPos', " & _
                          " 11.2, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPioNoYPos', " & _
                          " 0.85, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPioNoMaxWidth', " & _
                          " 2.6, " & _
                          " 'In cm. ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAnointedXPos', " & _
                          " 9.7, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAnointedYPos', " & _
                          " 2.8, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAnointedMaxWidth', " & _
                          " 1, " & _
                          " 'In cm. ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTopMargin', " & _
                          " 0.65, " & _
                          " 'In cm..')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardBottomMargin', " & _
                          " 0.8, " & _
                          " 'In cm..')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardLeftMargin', " & _
                          " 0.48, " & _
                          " 'In cm..')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardRightMargin', " & _
                          " 0.48, " & _
                          " 'In cm..')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardFontSize', " & _
                          " 8, " & _
                          " ' ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES('PubCardFontName', " & _
                          " 'Arial', " & _
                          " '.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPaperHeight', " & _
                          " 10.5, " & _
                          " 'In cm.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPaperWidth', " & _
                          " 14.8, " & _
                          " 'In cm.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTweakX', " & _
                          " 0, " & _
                          " ' ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTweakY', " & _
                          " 0, " & _
                          " ' ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardDOBXPos', " & _
                          " 9, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardDOBYPos', " & _
                          " 2.15, " & _
                          " 'In cm. Does not include margins.')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.47.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_47_00_To_5_48_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rstTemp As Recordset
    
    '
    'Create new table for monitoring whether pub rec card rows have been printed.
    '
    
'    DeleteTable "tblPubRecCardRowPrinted"
    
    CreateTable ErrorCode, "tblPubRecCardRowPrinted", "PersonID", "LONG", , "", False
    CreateField ErrorCode, "tblPubRecCardRowPrinted", "ActualMinPeriod", "DATE"
    CreateField ErrorCode, "tblPubRecCardRowPrinted", "Printed", "YESNO"
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblPubRecCardRowPrinted " & _
                  "   (PersonID, " & _
                  "    ActualMinPeriod) " & _
                  "WITH PRIMARY"
                  
    'initial load
    Set rstTemp = CMSDB.OpenRecordset("SELECT PersonID, " & _
                                         "       ActualMinPeriod " & _
                                         "FROM tblMinReports " & _
                                        " ORDER BY PersonID, ActualMinPeriod", dbOpenDynaset)

    With rstTemp
    If Not .BOF Then
        Do Until .EOF
            CMSDB.Execute "INSERT INTO tblPubRecCardRowPrinted " & _
                          "(PersonID, " & _
                          " ActualMinPeriod, " & _
                          " Printed) " & _
                          " VALUES(" & !PersonID & ", #" & _
                                        Format(!ActualMinPeriod, "mm/dd/yyyy") & "#, " & _
                                  " FALSE)"
        
            .MoveNext
        Loop
    End If
    
    .Close
    
    End With
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.48.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_48_00_To_5_49_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add more new publisher record card print parameters to tblConstants
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardServiceYearXPos', " & _
                          " 0.4, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardServiceYearYPos', " & _
                          " 3.6, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardBooksXPos', " & _
                          " 1.6, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardBrochuresXPos', " & _
                          " 3.0, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardHoursXPos', " & _
                          " 4.0, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardSubscripXPos', " & _
                          " 5.6, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardMagsXPos', " & _
                          " 5.5, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardRVsXPos', " & _
                          " 6.8, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardStudiesXPos', " & _
                          " 8.0, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardRemarksXPos', " & _
                          " 8.96, " & _
                          " 'In cm. Does not include margins.')"
    
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardSeptYPos', " & _
                          " 3.95, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardOctYPos', " & _
                          " 4.3, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardNovYPos', " & _
                          " 4.66, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardDecYPos', " & _
                          " 5.02, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardJanYPos', " & _
                          " 5.39, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardFebYPos', " & _
                          " 5.75, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardMarYPos', " & _
                          " 6.07, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAprYPos', " & _
                          " 6.45, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardMayYPos', " & _
                          " 6.82, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardJunYPos', " & _
                          " 7.15, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardJulYPos', " & _
                          " 7.5, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAugYPos', " & _
                          " 7.85, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTotalYPos', " & _
                          " 8.21, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardRemarksMaxWidth', " & _
                          " 4.0, " & _
                          " 'In cm. ')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.49.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_49_00_To_5_50_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    'Set Pub Rec-card printed flag en-mass.....
    
    CMSDB.Execute "UPDATE tblPubRecCardRowPrinted " & _
                  "SET Printed = TRUE " & _
                  "WHERE ActualMinPeriod BETWEEN #09/01/2004# AND #03/01/2005#"
    
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.50.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_50_00_To_5_51_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumFloat = 5 " & _
                  "WHERE FldName = 'PubCardRemarksMaxWidth' "
    
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.51.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_51_00_To_5_52_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.TableDefs.Refresh
    CMSDB.TableDefs("tblPublisherDates").Fields("StartReason").Required = False
    CMSDB.TableDefs("tblPublisherDates").Fields("EndReason").Required = False
    CMSDB.TableDefs.Refresh
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.52.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_52_00_To_5_53_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    
    '
    'tblSpecPioDates
    '
    CreateTable ErrorCode, "tblSpecPioDates", "PersonID", "LONG"
    CreateField ErrorCode, "tblSpecPioDates", "StartDate", "DATE"
    CreateField ErrorCode, "tblSpecPioDates", "EndDate", "DATE"
    CreateField ErrorCode, "tblSpecPioDates", "PioneerNumber", "TEXT"
    
    '
    'Add new 'Special Pio' Task
    '
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description) " & _
                  " VALUES (5, " & _
                          " 8, " & _
                          " 91, " & _
                          " 'Special Pioneer')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.53.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_5_53_00_To_5_54_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.TableDefs.Refresh
    CMSDB.TableDefs("tblConstants").Fields("Comment").AllowZeroLength = True
    CMSDB.TableDefs.Refresh
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES('IncludeSpecPioInTots', " & _
                          " 1, " & _
                          " '')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.54.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_54_00_To_5_55_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    DestroyGlobalObjects
    
    CMSDB.TableDefs.Refresh
    CMSDB.TableDefs("tblNameAddress").Fields("LinkedAddressPerson").Required = False
    CMSDB.TableDefs("tblNameAddress").Fields("Anointed").Required = False
    CMSDB.TableDefs.Refresh
    
    CMSDB.Execute "UPDATE tblNameAddress " & _
                  "SET Anointed = FALSE " & _
                  "WHERE Anointed IS NULL "
    
    CMSDB.Execute "UPDATE tblNameAddress " & _
                  "SET LinkedAddressPerson = 0 " & _
                  "WHERE LinkedAddressPerson IS NULL "
                  
    SetUpGlobalObjects
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.55.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_55_00_To_5_56_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

'    DeleteTable "tblAdvancedMinReporting"
    
    '
    'tblAdvancedMinReporting
    '
    CreateTable ErrorCode, "tblAdvancedMinReporting", "PersonID", "LONG", , , False
    CreateField ErrorCode, "tblAdvancedMinReporting", "ActualMinDate", "DATE"
    CreateField ErrorCode, "tblAdvancedMinReporting", "MinType", "LONG"
    CreateField ErrorCode, "tblAdvancedMinReporting", "ElderMS", "TEXT", "2"
    CreateField ErrorCode, "tblAdvancedMinReporting", "IsBaptised", "YESNO"
    CreateField ErrorCode, "tblAdvancedMinReporting", "NoBooks", "LONG", , ""
    CreateField ErrorCode, "tblAdvancedMinReporting", "NoBooklets", "LONG", , ""
    CreateField ErrorCode, "tblAdvancedMinReporting", "NoHours", "LONG", , ""
    CreateField ErrorCode, "tblAdvancedMinReporting", "NoMagazines", "LONG", , ""
    CreateField ErrorCode, "tblAdvancedMinReporting", "NoReturnVisits", "LONG", , ""
    CreateField ErrorCode, "tblAdvancedMinReporting", "NoStudies", "LONG", , ""
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblAdvancedMinReporting " & _
                  "   (PersonID, " & _
                  "    ActualMinDate) " & _
                  "WITH PRIMARY"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.56.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_56_00_To_5_57_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Allow decimal values
    '
    CMSDB.Execute "ALTER TABLE tblAdvancedMinReporting " & _
                  "ALTER COLUMN NoHours DOUBLE; "
    CMSDB.Execute "ALTER TABLE tblAdvancedMinReporting " & _
                  "ALTER COLUMN NoBooks DOUBLE; "
    CMSDB.Execute "ALTER TABLE tblAdvancedMinReporting " & _
                  "ALTER COLUMN NoBooklets DOUBLE; "
    CMSDB.Execute "ALTER TABLE tblAdvancedMinReporting " & _
                  "ALTER COLUMN NoMagazines DOUBLE; "
    CMSDB.Execute "ALTER TABLE tblAdvancedMinReporting " & _
                  "ALTER COLUMN NoReturnVisits DOUBLE; "
    CMSDB.Execute "ALTER TABLE tblAdvancedMinReporting " & _
                  "ALTER COLUMN NoStudies DOUBLE; "
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.57.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_57_00_To_5_58_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    DeleteTable "tblAdvancedMinReportingPrint"
    '
    'tblAdvancedMinReportingPrint
    '
    CreateTable ErrorCode, "tblAdvancedMinReportingPrint", "PersonID", "LONG", , , True
    CreateField ErrorCode, "tblAdvancedMinReportingPrint", "PersonName", "TEXT", "100"
    CreateField ErrorCode, "tblAdvancedMinReportingPrint", "AvgBooks", "DOUBLE"
    CreateField ErrorCode, "tblAdvancedMinReportingPrint", "AvgBooklets", "DOUBLE"
    CreateField ErrorCode, "tblAdvancedMinReportingPrint", "AvgHours", "DOUBLE"
    CreateField ErrorCode, "tblAdvancedMinReportingPrint", "AvgMagazines", "DOUBLE"
    CreateField ErrorCode, "tblAdvancedMinReportingPrint", "AvgReturnVisits", "DOUBLE"
    CreateField ErrorCode, "tblAdvancedMinReportingPrint", "AvgStudies", "DOUBLE"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.58.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_58_00_To_5_59_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CreateField ErrorCode, "tblAdvancedMinReportingPrint", "BookGroupName", "TEXT"
    CreateField ErrorCode, "tblAdvancedMinReporting", "BookGroupID", "LONG"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.59.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_59_00_To_5_60_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CreateField ErrorCode, "tblAdvancedMinReporting", "ArtificialValue", "YESNO"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.60.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_60_00_To_5_61_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES('UseWordForReports', " & _
                          " 1, " & _
                          " '')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.61.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_61_00_To_5_62_00()
On Error GoTo ErrorTrap
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.62.0"
    
    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

