Attribute VB_Name = "basUpgradeDB4"
Option Explicit

Public Sub Go_to_UpgradeDB4_module()
On Error GoTo ErrorTrap
    
    If gstrDBVersion <= "0005.9980.0000" Then
        DB_Upgrade_0005_9979_0157_To_0005_9980_0000
    End If
    If gstrDBVersion <= "0005.9980.0002" Then
        DB_Upgrade_0005_9980_0000_To_0005_9980_0003
    End If
    If gstrDBVersion <= "0005.9980.0003" Then
        DB_Upgrade_0005_9980_0003_To_0005_9980_0004
    End If
    If gstrDBVersion <= "0005.9980.0004" Then
        DB_Upgrade_0005_9980_0004_To_0005_9980_0005
    End If
    If gstrDBVersion <= "0005.9980.0005" Then
        DB_Upgrade_0005_9980_0005_To_0005_9980_0006
    End If
    If gstrDBVersion <= "0005.9980.0006" Then
        DB_Upgrade_0005_9980_0006_To_0005_9980_0007
    End If
    If gstrDBVersion <= "0005.9980.0007" Then
        DB_Upgrade_0005_9980_0007_To_0005_9980_0008
    End If
    If gstrDBVersion <= "0005.9980.0008" Then
        DB_Upgrade_0005_9980_0008_To_0005_9980_0009
    End If
    If gstrDBVersion <= "0005.9980.0009" Then
        DB_Upgrade_0005_9980_0009_To_0005_9980_0010
    End If
    If gstrDBVersion <= "0005.9980.0011" Then
        DB_Upgrade_0005_9980_0010_To_0005_9980_0012
    End If
    If gstrDBVersion <= "0005.9980.0012" Then
        DB_Upgrade_0005_9980_0012_To_0005_9980_0013
    End If
    If gstrDBVersion <= "0005.9980.0013" Then
        DB_Upgrade_0005_9980_0013_To_0005_9980_0014
    End If
    If gstrDBVersion <= "0005.9980.0014" Then
        DB_Upgrade_0005_9980_0014_To_0005_9980_0015
    End If
    If gstrDBVersion <= "0005.9980.0015" Then
        DB_Upgrade_0005_9980_0015_To_0005_9980_0016
    End If
    If gstrDBVersion <= "0005.9980.0016" Then
        DB_Upgrade_0005_9980_0016_To_0005_9980_0017
    End If
    If gstrDBVersion <= "0005.9980.0017" Then
        DB_Upgrade_0005_9980_0017_To_0005_9980_0018
    End If
    If gstrDBVersion <= "0005.9980.0018" Then
        DB_Upgrade_0005_9980_0018_To_0005_9980_0019
    End If
    If gstrDBVersion <= "0005.9980.0019" Then
        DB_Upgrade_0005_9980_0019_To_0005_9980_0020
    End If
    If gstrDBVersion <= "0005.9980.0020" Then
        DB_Upgrade_0005_9980_0020_To_0005_9980_0021
    End If
    If gstrDBVersion <= "0005.9980.0021" Then
        DB_Upgrade_0005_9980_0021_To_0005_9980_0022
    End If
    If gstrDBVersion <= "0005.9980.0023" Then
        DB_Upgrade_0005_9980_0022_To_0005_9980_0024
    End If
    If gstrDBVersion <= "0005.9980.0024" Then
        DB_Upgrade_0005_9980_0024_To_0005_9980_0025
    End If
    If gstrDBVersion <= "0005.9980.0025" Then
        DB_Upgrade_0005_9980_0025_To_0005_9980_0026
    End If
    If gstrDBVersion <= "0005.9980.0026" Then
        DB_Upgrade_0005_9980_0026_To_0005_9980_0027
    End If
    If gstrDBVersion <= "0005.9980.0027" Then
        DB_Upgrade_0005_9980_0027_To_0005_9980_0028
    End If
    If gstrDBVersion <= "0005.9980.0028" Then
        DB_Upgrade_0005_9980_0028_To_0005_9980_0029
    End If
    If gstrDBVersion <= "0005.9980.0029" Then
        DB_Upgrade_0005_9980_0029_To_0005_9980_0030
    End If

    
    
    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub


Private Sub DB_Upgrade_0005_9979_0157_To_0005_9980_0000()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0000"
    
    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9980_0000_To_0005_9980_0003()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
      
    CMSDB.Execute "DELETE FROM tblConstants " & _
              " WHERE FldName = 'TMSHighlightAsstDates' "
                                                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0003"
    
    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9980_0003_To_0005_9980_0004()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0004"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9980_0004_To_0005_9980_0005()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0005"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9980_0005_To_0005_9980_0006()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer


    WriteToLogFile "Applying DB_Upgrade_0005_9980_0005_To_0005_9980_0006"

    'update the Speech Quality descriptions from spreadsheet...
    
    If Not TMS_UpdateSQDescriptionsFromXLS(True) Then
        Err.Raise vbObjectError + 270, "basUpgrade4.DB_Upgrade_0005_9980_0005_To_0005_9980_0006", _
                   "Could not update SQ points from spreadsheet"
    End If

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0006"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9980_0006_To_0005_9980_0007()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0007"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9980_0007_To_0005_9980_0008()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    'goodbye to special assemblies...
    
        CMSDB.Execute "UPDATE tblEventLookup SET ShowInCalendar = FALSE " & _
              " WHERE EventID = 2 "

        CMSDB.Execute "DELETE FROM tblConstants  " & _
              " WHERE FldName = 'SpecialAssemblyOnCalendar' "

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0008"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9980_0008_To_0005_9980_0009()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    'TMS bros do no 3 demos only when using nwt...
    
    CMSDB.Execute "INSERT INTO tblConstants " & _
         "(FldName, " & _
         " AlphaVal, " & _
         " Comment) " & _
         " VALUES ('TMS_SourceTextForBroNo3Demo', " & _
                   "'nwt ', " & _
                 " 'Initial value = nwt ')"
                 
    CMSDB.Execute "INSERT INTO tblConstants " & _
         "(FldName, " & _
         " AlphaVal, " & _
         " Comment) " & _
         " VALUES ('TMS_SourceTextForBroNo3Talk', " & _
                   "'it- ', " & _
                 " 'Initial value = it- ')"

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0009"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9980_0009_To_0005_9980_0010()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
         "(FldName, " & _
         " TrueFalse, " & _
         " Comment) " & _
         " VALUES ('TMSLoadNo3BroOnlyWarning', " & _
                   "FALSE, " & _
                 " 'Initial value = FALSE ')"
                 

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0010"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9980_0010_To_0005_9980_0012()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
         "(FldName, " & _
         " AlphaVal, " & _
         " Comment) " & _
         " VALUES ('TMSSlipNote_ElderChooseAsst', " & _
                   "'Please select an assistant for field service or family worship setting', " & _
                 " 'Initial value = Please select an assistant for field service or family worship setting')"
                 
    CMSDB.Execute "INSERT INTO tblConstants " & _
         "(FldName, " & _
         " AlphaVal, " & _
         " Comment) " & _
         " VALUES ('TMSSlipNote_DoNo3AsTalk', " & _
                   "'Please handle as a talk', " & _
                 " 'Please handle as a talk')"

    CMSDB.Execute "INSERT INTO tblConstants " & _
         "(FldName, " & _
         " AlphaVal, " & _
         " Comment) " & _
         " VALUES ('TMSSlipNote_DoNo3AsFieldMinOrFamWorship', " & _
                   "'Please use family worship or field service setting', " & _
                 " 'Please use family worship or field service setting')"
                 
    CMSDB.Execute "INSERT INTO tblConstants " & _
         "(FldName, " & _
         " AlphaVal, " & _
         " Comment) " & _
         " VALUES ('TMSSlipNote_DoNo3AsFieldMin', " & _
                   "'Please use field service setting', " & _
                 " 'Please use field service setting')"

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0012"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9980_0012_To_0005_9980_0013()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer


    WriteToLogFile "Applying DB_Upgrade_0005_9980_0012_To_0005_9980_0013"

    'update the Speech Quality descriptions from spreadsheet...
    
    If Not TMS_UpdateSQDescriptionsFromXLS(True) Then
        Err.Raise vbObjectError + 270, "basUpgrade4.DB_Upgrade_0005_9980_0012_To_0005_9980_0013", _
                   "Could not update SQ points from spreadsheet"
    End If

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0013"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9980_0013_To_0005_9980_0014()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer


    WriteToLogFile "Applying DB_Upgrade_0005_9980_0013_To_0005_9980_0014"

    CMSDB.Execute "INSERT INTO tblConstants " & _
         "(FldName, " & _
         " TrueFalse, " & _
         " Comment) " & _
         " VALUES ('CBSAutoGen_IncludePrayer', " & _
                   "FALSE, " & _
                 " 'Initial value = FALSE ')"

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0014"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9980_0014_To_0005_9980_0015()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    WriteToLogFile "Applying DB_Upgrade_0005_9980_0014_To_0005_9980_0015"

    CMSDB.Execute "UPDATE tblTasks " & _
                  "SET AllowSuspend = FALSE " & _
                  "WHERE Task IN (19)"
                                 
    CMSDB.Execute "UPDATE tblEventLookup " & _
                  "SET EventName = 'Regional Convention' " & _
                  "WHERE EventID =3 "

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0015"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9980_0015_To_0005_9980_0016()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    WriteToLogFile "Applying DB_Upgrade_0005_9980_0015_To_0005_9980_0016"


    DeleteTable "tblTMSSchoolOverseerDates"

    CreateTable ErrorCode, "tblTMSSchoolOverseerDates", "PersonID", "LONG", , , True, "SeqID"
    CreateField ErrorCode, "tblTMSSchoolOverseerDates", "SchoolNo", "LONG"
    CreateField ErrorCode, "tblTMSSchoolOverseerDates", "SchoolDate", "DATE"
    
    CMSDB.TableDefs.Refresh



    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0016"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9980_0016_To_0005_9980_0017()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    WriteToLogFile "Applying DB_Upgrade_0005_9980_0016_To_0005_9980_0017"


    DelAllRows "tblStoredTMSSchedules"

    CreateField ErrorCode, "tblStoredTMSSchedules", "StartDate", "DATE"
    CreateField ErrorCode, "tblStoredTMSSchedules", "EndDate", "DATE"
    
    CMSDB.TableDefs.Refresh



    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0017"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9980_0017_To_0005_9980_0018()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    WriteToLogFile "Applying DB_Upgrade_0005_9980_0017_To_0005_9980_0018"

    CMSDB.Execute "INSERT INTO tblConstants " & _
         "(FldName, " & _
         " NumVal, " & _
         " Comment) " & _
         " VALUES ('TMSSlipNoteMaxLength', " & _
                   "130, " & _
                 " 'Initial value = 130 ');"
                 
                     
     DestroyGlobalObjects
     
     CMSDB.Execute "ALTER TABLE tblTMSSchedule " & _
                  "ALTER COLUMN StudentNote MEMO; "
                 
     SetUpGlobalObjects

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0018"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9980_0018_To_0005_9980_0019()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    WriteToLogFile "Applying DB_Upgrade_0005_9980_0018_To_0005_9980_0019"

    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumVal = 4, Comment = 'Initial value = 4' " & _
                  " WHERE FldName = 'TMSMinAge' "

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0019"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9980_0019_To_0005_9980_0020()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    WriteToLogFile "Applying DB_Upgrade_0005_9980_0019_To_0005_9980_0020"
                 
    CMSDB.Execute "UPDATE tblTasks SET Description = 'Bible Reading' WHERE Task = 99"

    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description, " & _
                  " AllowSuspend, " & _
                  " RequiresExemplaryBro, " & _
                  " TaskComment) " & _
                  " VALUES (4, " & _
                          " 6, " & _
                          " 105, " & _
                          " 'Initial Call',  TRUE, FALSE, " & _
                          " 'For AYM in 2016 onwards') "
                          
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description, " & _
                  " AllowSuspend, " & _
                  " RequiresExemplaryBro, " & _
                  " TaskComment) " & _
                  " VALUES (4, " & _
                          " 6, " & _
                          " 106, " & _
                          " 'Return Visit',  TRUE, FALSE, " & _
                          " 'For AYM in 2016 onwards') "
                          
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description, " & _
                  " AllowSuspend, " & _
                  " RequiresExemplaryBro, " & _
                  " TaskComment) " & _
                  " VALUES (4, " & _
                          " 6, " & _
                          " 107, " & _
                          " 'Bible Study',  TRUE, FALSE, " & _
                          " 'For AYM in 2016 onwards') "
                          
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description, " & _
                  " AllowSuspend, " & _
                  " RequiresExemplaryBro, " & _
                  " TaskComment) " & _
                  " VALUES (4, " & _
                          " 6, " & _
                          " 108, " & _
                          " 'Other',  TRUE, FALSE, " & _
                          " 'For AYM in 2016 onwards') "
                          
    CreateField ErrorCode, "tblTMSTalkNoDesc", "Order2016", "LONG"
    
       
                        
    
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0020"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9980_0020_To_0005_9980_0021()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    WriteToLogFile "Applying DB_Upgrade_0005_9980_0020_To_0005_9980_0021"

    'add the AYFM 'other' talk type
     CMSDB.Execute "INSERT INTO tblTMSTalkNoDesc " & _
                   " (TalkNo, TalkDesc, Order2016) " & _
                  "VALUES ('O','Other', 25)"


    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0021"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9980_0021_To_0005_9980_0022()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    WriteToLogFile "Applying DB_Upgrade_0005_9980_0021_To_0005_9980_0022"

    'new AYFM assignment weightings
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES ('TMSBibleReadingWeighting_2016', " & _
                          " 40, " & _
                          " 'Initial value = 40')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES ('TMSInitialCallWeighting_2016', " & _
                          " 60, " & _
                          " 'Initial value = 60')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES ('TMSReturnVisitWeighting_2016', " & _
                          " 80, " & _
                          " 'Initial value = 80')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES ('TMSBibleStudyWeighting_2016', " & _
                          " 100, " & _
                          " 'Initial value = 100')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES ('TMSOtherWeighting_2016', " & _
                          " 80, " & _
                          " 'Initial value = 80')"


    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0022"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub



Private Sub DB_Upgrade_0005_9980_0022_To_0005_9980_0024()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    WriteToLogFile "Applying DB_Upgrade_0005_9980_0022_To_0005_9980_0024"

    CMSDB.Execute "INSERT INTO tblTMSTalkNoDesc " & _
                  "(TalkNo, TalkDesc, Order2009, Order2016) " & _
                  " VALUES ('BR', 'Bible Reading', 0, 5) "
                          
    CMSDB.Execute "INSERT INTO tblTMSTalkNoDesc " & _
                  "(TalkNo, TalkDesc, Order2009, Order2016) " & _
                  " VALUES ('IC', 'Initial Call',0, 10) "
                          
    CMSDB.Execute "INSERT INTO tblTMSTalkNoDesc " & _
                  "(TalkNo, TalkDesc, Order2009, Order2016) " & _
                  " VALUES ('RV', 'Return Visit',0, 15) "
                          
    CMSDB.Execute "INSERT INTO tblTMSTalkNoDesc " & _
                  "(TalkNo, TalkDesc, Order2009, Order2016) " & _
                  " VALUES ('BS', 'Bible Study',0, 20) "



    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0024"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Private Sub DB_Upgrade_0005_9980_0024_To_0005_9980_0025()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    WriteToLogFile "Applying DB_Upgrade_0005_9980_0024_To_0005_9980_0025"

    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item1TalkNo", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item1StudentName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item1AssistantName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item1SettingName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item1Theme", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item1Source", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item1CounselPoint", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item1CounselSubPoint1", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item1CounselSubPoint2", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item1CounselSubPoint3", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item1CounselSubPoint4", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item1CounselSubPoint5", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item1Comment", "TEXT", "255"

    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item2TalkNo", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item2StudentName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item2AssistantName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item2SettingName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item2Theme", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item2Source", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item2CounselPoint", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item2CounselSubPoint1", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item2CounselSubPoint2", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item2CounselSubPoint3", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item2CounselSubPoint4", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item2CounselSubPoint5", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item2Comment", "TEXT", "255"

    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item3TalkNo", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item3StudentName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item3AssistantName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item3SettingName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item3Theme", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item3Source", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item3CounselPoint", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item3CounselSubPoint1", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item3CounselSubPoint2", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item3CounselSubPoint3", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item3CounselSubPoint4", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item3CounselSubPoint5", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item3Comment", "TEXT", "255"

    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item4TalkNo", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item4StudentName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item4AssistantName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item4SettingName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item4Theme", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item4Source", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item4CounselPoint", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item4CounselSubPoint1", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item4CounselSubPoint2", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item4CounselSubPoint3", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item4CounselSubPoint4", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item4CounselSubPoint5", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item4Comment", "TEXT", "255"

    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item5TalkNo", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item5StudentName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item5AssistantName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item5SettingName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item5Theme", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item5Source", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item5CounselPoint", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item5CounselSubPoint1", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item5CounselSubPoint2", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item5CounselSubPoint3", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item5CounselSubPoint4", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item5CounselSubPoint5", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item5Comment", "TEXT", "255"


    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0025"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9980_0025_To_0005_9980_0026()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset, rs2 As Recordset
                
    WriteToLogFile "Applying DB_Upgrade_0005_9980_0025_To_0005_9980_0026"

     DestroyGlobalObjects

    CreateField ErrorCode, "tblTMSSchedule", "ItemsSeqNum", "LONG"

   SetUpGlobalObjects
   
    Set rs = CMSDB.OpenRecordset("tblTMSSchedule", dbOpenDynaset)
    Set rs2 = CMSDB.OpenRecordset("tblTMSItems", dbOpenDynaset)
    
    With rs
    
        Do Until .EOF Or .BOF
        
            If year(!AssignmentDate) < "2016" Then
            
                rs2.FindFirst "AssignmentDate = #" & Format$(!AssignmentDate, "mm/dd/yyyy") & "# AND TalkNo = '" & !TalkNo & "'"
                
                .Edit
                
                If Not rs2.NoMatch Then
                    !ItemsSeqNum = rs2!ItemsSeqNum
                Else
                    !ItemsSeqNum = -1
                End If
                
                .Update
                
            
            
            End If
            
            .MoveNext
        
        
        Loop
    
    
    End With
    
    
    rs.Close
    Set rs = Nothing
    rs2.Close
    Set rs2 = Nothing
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0026"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9980_0026_To_0005_9980_0027()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                
    WriteToLogFile "Applying DB_Upgrade_0005_9980_0026_To_0005_9980_0027"
    
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item1SettingTitle", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item2SettingTitle", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item3SettingTitle", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item4SettingTitle", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "Item5SettingTitle", "TEXT", "255"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0027"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9980_0027_To_0005_9980_0028()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset, rs2 As Recordset
                
    WriteToLogFile "Applying DB_Upgrade_0005_9980_0027_To_0005_9980_0028"
    
  
    Set rs = CMSDB.OpenRecordset("tblTMSSchedule", dbOpenDynaset)
    Set rs2 = CMSDB.OpenRecordset("tblTMSItems", dbOpenDynaset)
    
    With rs
    
        Do Until .EOF Or .BOF
        
            If year(!AssignmentDate) >= "2016" Then
            
                rs2.FindFirst "AssignmentDate = #" & Format$(!AssignmentDate, "mm/dd/yyyy") & "# AND TalkNo = '" & !TalkNo & "'"
                
                .Edit
                
                If Not rs2.NoMatch Then
                    !ItemsSeqNum = rs2!ItemsSeqNum
                Else
                    !ItemsSeqNum = -1
                End If
                
                .Update
                
            
            
            End If
            
            .MoveNext
        
        
        Loop
    
    
    End With
    
    
    rs.Close
    Set rs = Nothing
    rs2.Close
    Set rs2 = Nothing
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0028"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub



Private Sub DB_Upgrade_0005_9980_0028_To_0005_9980_0029()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    WriteToLogFile "Applying DB_Upgrade_0005_9980_0028_To_0005_9980_0029"

    CMSDB.Execute "UPDATE tblConstants " & _
                  " SET AlphaVal = '~45~46~47~', " & _
                  "     Comment = 'Initial values = ~45~46~47~, TRUE (each SQ must be wrapped with ~nn~)'" & _
                  " WHERE FldName = 'TMS_AwkwardCounselPoints'"
                          
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0029"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9980_0029_To_0005_9980_0030()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    WriteToLogFile "Applying DB_Upgrade_0005_9980_0029_To_0005_9980_0030"


    CMSDB.Execute "UPDATE tblConstants " & _
                  " SET NumVal = 80, " & _
                  "     Comment = 'Initial values = 80'" & _
                  " WHERE FldName = 'TMSSlipNoteMaxLength'"
                          
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9980.0030"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub



