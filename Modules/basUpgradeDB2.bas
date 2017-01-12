Attribute VB_Name = "basUpgradeDB2"
Option Explicit

Public Sub Go_to_UpgradeDB2_module()
On Error GoTo ErrorTrap
    
    If gstrDBVersion <= "0005.0062.0000" Then
        DB_Upgrade_5_62_00_To_5_63_00
    End If

    If gstrDBVersion <= "0005.0063.0000" Then
        DB_Upgrade_5_63_00_To_5_64_00
    End If

    If gstrDBVersion <= "0005.0064.0000" Then
        DB_Upgrade_5_64_00_To_5_65_00
    End If

    If gstrDBVersion <= "0005.0065.0000" Then
        DB_Upgrade_5_65_00_To_5_66_00
    End If

    If gstrDBVersion <= "0005.0066.0000" Then
        DB_Upgrade_5_66_00_To_5_67_00
    End If

    If gstrDBVersion <= "0005.0067.0000" Then
        DB_Upgrade_5_67_00_To_5_68_00
    End If

    If gstrDBVersion <= "0005.0068.0000" Then
        DB_Upgrade_5_68_00_To_5_69_00
    End If

    If gstrDBVersion <= "0005.0069.0000" Then
        DB_Upgrade_5_69_00_To_5_70_00
    End If

    If gstrDBVersion <= "0005.0070.0000" Then
        DB_Upgrade_5_70_00_To_5_71_00
    End If

    If gstrDBVersion <= "0005.0071.0000" Then
        DB_Upgrade_5_71_00_To_5_72_00
    End If

    If gstrDBVersion <= "0005.0072.0000" Then
        DB_Upgrade_5_72_00_To_5_73_00
    End If

    If gstrDBVersion <= "0005.0073.0000" Then
        DB_Upgrade_5_73_00_To_5_74_00
    End If

    If gstrDBVersion <= "0005.0074.0000" Then
        DB_Upgrade_5_74_00_To_5_75_00
    End If

    If gstrDBVersion <= "0005.0075.0000" Then
        DB_Upgrade_5_75_00_To_5_76_00
    End If

    If gstrDBVersion <= "0005.0076.0000" Then
        DB_Upgrade_5_76_00_To_5_77_00
    End If

    If gstrDBVersion <= "0005.0077.0000" Then
        DB_Upgrade_5_77_00_To_5_78_00
    End If

    If gstrDBVersion <= "0005.0078.0000" Then
        DB_Upgrade_5_78_00_To_5_79_00
    End If

    If gstrDBVersion <= "0005.0079.0000" Then
        DB_Upgrade_5_79_00_To_5_80_00
    End If

    If gstrDBVersion <= "0005.0080.0000" Then
        DB_Upgrade_5_80_00_To_5_81_00
    End If

    If gstrDBVersion <= "0005.0081.0000" Then
        DB_Upgrade_5_81_00_To_5_82_00
    End If

    If gstrDBVersion <= "0005.0082.0000" Then
        DB_Upgrade_5_82_00_To_5_83_00
    End If

    If gstrDBVersion <= "0005.0083.0000" Then
        DB_Upgrade_5_83_00_To_5_84_00
    End If

    If gstrDBVersion <= "0005.0084.0000" Then
        DB_Upgrade_5_84_00_To_5_85_00
    End If

    If gstrDBVersion <= "0005.0085.0000" Then
        DB_Upgrade_5_85_00_To_5_86_00
    End If

    If gstrDBVersion <= "0005.0086.0000" Then
        DB_Upgrade_5_86_00_To_5_87_00
    End If

    If gstrDBVersion <= "0005.0087.0000" Then
        DB_Upgrade_5_87_00_To_5_88_00
    End If

    If gstrDBVersion <= "0005.0088.0000" Then
        DB_Upgrade_5_88_00_To_5_89_00
    End If

    If gstrDBVersion <= "0005.0089.0000" Then
        DB_Upgrade_5_89_00_To_5_90_00
    End If

    If gstrDBVersion <= "0005.0090.0000" Then
        DB_Upgrade_5_90_00_To_5_91_00
    End If

    If gstrDBVersion <= "0005.0091.0000" Then
        DB_Upgrade_5_91_00_To_5_92_00
    End If

    If gstrDBVersion <= "0005.0092.0000" Then
        DB_Upgrade_5_92_00_To_5_93_00
    End If

    If gstrDBVersion <= "0005.0093.0000" Then
        DB_Upgrade_5_93_00_To_5_94_00
    End If

    If gstrDBVersion <= "0005.0094.0000" Then
        DB_Upgrade_5_94_00_To_5_95_00
    End If

    If gstrDBVersion <= "0005.0095.0000" Then
        DB_Upgrade_5_95_00_To_5_96_00
    End If

    If gstrDBVersion <= "0005.0096.0000" Then
        DB_Upgrade_5_96_00_To_5_97_00
    End If

    If gstrDBVersion <= "0005.0097.0000" Then
        DB_Upgrade_5_97_00_To_5_98_00
    End If

    If gstrDBVersion <= "0005.0098.0000" Then
        DB_Upgrade_5_98_00_To_5_99_00
    End If

    If gstrDBVersion <= "0005.0099.0000" Then
        DB_Upgrade_5_99_00_To_5_9900_00
    End If

    If gstrDBVersion <= "0005.9900.0000" Then
        DB_Upgrade_5_9900_00_To_5_9901_00
    End If

    If gstrDBVersion <= "0005.9901.0000" Then
        DB_Upgrade_5_9901_00_To_5_9902_00
    End If

    If gstrDBVersion <= "0005.9902.0000" Then
        DB_Upgrade_5_9902_00_To_5_9903_00
    End If

    If gstrDBVersion <= "0005.9903.0000" Then
        DB_Upgrade_5_9903_00_To_5_9904_00
    End If

    If gstrDBVersion <= "0005.9904.0000" Then
        DB_Upgrade_5_9904_00_To_5_9905_00
    End If

    If gstrDBVersion <= "0005.9905.0000" Then
        DB_Upgrade_5_9905_00_To_5_9906_00
    End If

    If gstrDBVersion <= "0005.9906.0000" Then
        DB_Upgrade_5_9906_00_To_5_9907_00
    End If

    If gstrDBVersion <= "0005.9907.0000" Then
        DB_Upgrade_5_9907_00_To_5_9908_00
    End If

    If gstrDBVersion <= "0005.9908.0000" Then
        DB_Upgrade_5_9908_00_To_5_9909_00
    End If

    If gstrDBVersion <= "0005.9909.0000" Then
        DB_Upgrade_5_9909_00_To_5_9910_00
    End If

    If gstrDBVersion <= "0005.9910.0000" Then
        DB_Upgrade_5_9910_00_To_5_9911_00
    End If

    If gstrDBVersion <= "0005.9911.0000" Then
        DB_Upgrade_5_9911_00_To_5_9912_00
    End If

    If gstrDBVersion <= "0005.9912.0000" Then
        DB_Upgrade_5_9912_00_To_5_9913_00
    End If

    If gstrDBVersion <= "0005.9913.0000" Then
        DB_Upgrade_5_9913_00_To_5_9914_00
    End If

    If gstrDBVersion <= "0005.9914.0000" Then
        DB_Upgrade_5_9914_00_To_5_9915_00
    End If

    If gstrDBVersion <= "0005.9915.0000" Then
        DB_Upgrade_5_9915_00_To_5_9916_00
    End If

    If gstrDBVersion <= "0005.9916.0000" Then
        DB_Upgrade_5_9916_00_To_5_9917_00
    End If

    If gstrDBVersion <= "0005.9917.0000" Then
        DB_Upgrade_5_9917_00_To_5_9918_00
    End If

    If gstrDBVersion <= "0005.9918.0000" Then
        DB_Upgrade_5_9918_00_To_5_9919_00
    End If

    If gstrDBVersion <= "0005.9919.0000" Then
        DB_Upgrade_5_9919_00_To_5_9920_00
    End If

    If gstrDBVersion <= "0005.9920.0000" Then
        DB_Upgrade_5_9920_00_To_5_9921_00
    End If

    If gstrDBVersion <= "0005.9921.0000" Then
        DB_Upgrade_5_9921_00_To_5_9922_00
    End If

    If gstrDBVersion <= "0005.9922.0000" Then
        DB_Upgrade_5_9922_00_To_5_9923_00
    End If

    If gstrDBVersion <= "0005.9923.0000" Then
        DB_Upgrade_5_9923_00_To_5_9924_00
    End If

    If gstrDBVersion <= "0005.9924.0000" Then
        DB_Upgrade_5_9924_00_To_5_9925_00
    End If

    If gstrDBVersion <= "0005.9925.0000" Then
        DB_Upgrade_5_9925_00_To_5_9926_00
    End If

    If gstrDBVersion <= "0005.9926.0000" Then
        DB_Upgrade_5_9926_00_To_5_9927_00
    End If

    If gstrDBVersion <= "0005.9927.0000" Then
        DB_Upgrade_5_9927_00_To_5_9928_00
    End If

    If gstrDBVersion <= "0005.9928.0000" Then
        DB_Upgrade_5_9928_00_To_5_9929_00
    End If

    If gstrDBVersion <= "0005.9929.0000" Then
        DB_Upgrade_5_9929_00_To_5_9930_00
    End If

    If gstrDBVersion <= "0005.9930.0000" Then
        DB_Upgrade_5_9930_00_To_5_9931_00
    End If

    If gstrDBVersion <= "0005.9931.0000" Then
        DB_Upgrade_5_9931_00_To_5_9932_00
    End If

    If gstrDBVersion <= "0005.9932.0000" Then
        DB_Upgrade_5_9932_00_To_5_9933_00
    End If

    If gstrDBVersion <= "0005.9933.0000" Then
        DB_Upgrade_5_9933_00_To_5_9934_00
    End If

    If gstrDBVersion <= "0005.9934.0000" Then
        DB_Upgrade_5_9934_00_To_5_9935_00
    End If

    If gstrDBVersion <= "0005.9935.0000" Then
        DB_Upgrade_5_9935_00_To_5_9936_00
    End If

    If gstrDBVersion <= "0005.9936.0000" Then
        DB_Upgrade_5_9936_00_To_5_9937_00
    End If

    If gstrDBVersion <= "0005.9937.0000" Then
        DB_Upgrade_5_9937_00_To_5_9938_00
    End If

    If gstrDBVersion <= "0005.9938.0000" Then
        DB_Upgrade_5_9938_00_To_5_9939_00
    End If

    If gstrDBVersion <= "0005.9939.0000" Then
        DB_Upgrade_5_9939_00_To_5_9940_00
    End If

    If gstrDBVersion <= "0005.9940.0000" Then
        DB_Upgrade_5_9940_00_To_5_9941_00
    End If

    If gstrDBVersion <= "0005.9941.0000" Then
        DB_Upgrade_5_9941_00_To_5_9942_00
    End If

    If gstrDBVersion <= "0005.9942.0000" Then
        DB_Upgrade_5_9942_00_To_5_9943_00
    End If

    If gstrDBVersion <= "0005.9943.0000" Then
        DB_Upgrade_5_9943_00_To_5_9944_00
    End If

    If gstrDBVersion <= "0005.9944.0000" Then
        DB_Upgrade_5_9944_00_To_5_9945_00
    End If

    If gstrDBVersion <= "0005.9945.0000" Then
        DB_Upgrade_5_9945_00_To_5_9946_00
    End If

    If gstrDBVersion <= "0005.9946.0000" Then
        DB_Upgrade_5_9946_00_To_5_9947_00
    End If

    If gstrDBVersion <= "0005.9947.0000" Then
        DB_Upgrade_5_9947_00_To_5_9948_00
    End If

    If gstrDBVersion <= "0005.9948.0000" Then
        DB_Upgrade_5_9948_00_To_5_9949_00
    End If

    If gstrDBVersion <= "0005.9949.0000" Then
        DB_Upgrade_5_9949_00_To_5_9950_00
    End If

    If gstrDBVersion <= "0005.9950.0000" Then
        DB_Upgrade_5_9950_00_To_5_9951_00
    End If

    If gstrDBVersion <= "0005.9951.0000" Then
        DB_Upgrade_5_9951_00_To_5_9952_00
    End If

    If gstrDBVersion <= "0005.9952.0000" Then
        DB_Upgrade_5_9952_00_To_5_9953_00
    End If

    If gstrDBVersion <= "0005.9953.0000" Then
        DB_Upgrade_5_9953_00_To_5_9954_00
    End If

    If gstrDBVersion <= "0005.9954.0000" Then
        DB_Upgrade_5_9954_00_To_5_9955_00
    End If

    If gstrDBVersion <= "0005.9955.0000" Then
        DB_Upgrade_5_9955_00_To_5_9956_00
    End If

    If gstrDBVersion <= "0005.9956.0000" Then
        DB_Upgrade_5_9956_00_To_5_9957_00
    End If

    If gstrDBVersion <= "0005.9957.0000" Then
        DB_Upgrade_5_9957_00_To_5_9958_00
    End If

    If gstrDBVersion <= "0005.9958.0000" Then
        DB_Upgrade_5_9958_00_To_5_9959_00
    End If

    If gstrDBVersion <= "0005.9959.0000" Then
        DB_Upgrade_5_9959_00_To_5_9960_00
    End If

    If gstrDBVersion <= "0005.9960.0000" Then
        DB_Upgrade_5_9960_00_To_5_9961_00
    End If

    If gstrDBVersion <= "0005.9961.0000" Then
        DB_Upgrade_5_9961_00_To_5_9962_00
    End If

    If gstrDBVersion <= "0005.9962.0000" Then
        DB_Upgrade_5_9962_00_To_5_9963_00
    End If

    If gstrDBVersion <= "0005.9963.0000" Then
        DB_Upgrade_5_9963_00_To_5_9964_00
    End If

    If gstrDBVersion <= "0005.9964.0000" Then
        DB_Upgrade_5_9964_00_To_5_9965_00
    End If

    If gstrDBVersion <= "0005.9965.0000" Then
        DB_Upgrade_5_9965_00_To_5_9966_00
    End If

    If gstrDBVersion <= "0005.9966.0000" Then
        DB_Upgrade_5_9966_00_To_5_9967_00
    End If

    If gstrDBVersion <= "0005.9967.0000" Then
        DB_Upgrade_5_9967_00_To_5_9968_00
    End If

    If gstrDBVersion <= "0005.9968.0000" Then
        DB_Upgrade_5_9968_00_To_5_9969_00
    End If

    If gstrDBVersion <= "0005.9969.0001" Then
        DB_Upgrade_5_9969_00_To_5_9973_00
    End If

    If gstrDBVersion <= "0005.9973.0000" Then
        DB_Upgrade_5_9973_00_To_5_9974_00
    End If
    
    Go_to_UpgradeDB3_module
    
    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_62_00_To_5_63_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, TheString As String, rstTemp As Recordset
   
    DeleteTable "tblIndividualSPAMWeightings"
    
    '
    'Now build new tblIndividualSPAMWeightings table
    '
    CreateTable ErrorCode, "tblIndividualSPAMWeightings", "PersonID", "LONG", , , False
    CreateField ErrorCode, "tblIndividualSPAMWeightings", "Att_Wtg", "DOUBLE"
    CreateField ErrorCode, "tblIndividualSPAMWeightings", "Mic_Wtg", "DOUBLE"
    CreateField ErrorCode, "tblIndividualSPAMWeightings", "Snd_Wtg", "DOUBLE"
    CreateField ErrorCode, "tblIndividualSPAMWeightings", "Plt_Wtg", "DOUBLE"
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblIndividualSPAMWeightings " & _
                  "   (PersonID) " & _
                  "WITH PRIMARY"
                  
    'initial load
    TheString = "SELECT Person " & _
                "FROM tblTaskAndPerson " & _
                "WHERE Task IN (57, 58, 59, 60)"
    
    Set rstTemp = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)

    With rstTemp
    Do Until .EOF Or .BOF
        CMSDB.Execute "INSERT INTO tblIndividualSPAMWeightings " & _
                          "(PersonID, Att_Wtg, Mic_Wtg, Snd_Wtg, Plt_Wtg) " & _
                          "VALUES (" & !Person & ", 1, 1, 1, 1)"
        .MoveNext
    Loop
    End With
    
    Set rstTemp = Nothing
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.63.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_63_00_To_5_64_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, TheString As String, rstTemp As Recordset
   
    CMSDB.Execute "INSERT INTO tblEventLookup " & _
                  "(EventID, " & _
                  " EventName) " & _
                  " VALUES (14, " & _
                        " 'Field Service Report Entry Reminder')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.64.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_64_00_To_5_65_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, NewField As DAO.Field

    'new TimeVal column
    CMSDB.TableDefs.Refresh
    Set NewField = CMSDB.TableDefs("tblConstants").CreateField("TimeVal", dbDate)
    CMSDB.TableDefs("tblConstants").Fields.Append NewField
    NewField.Required = False
    
    CMSDB.TableDefs.Refresh
 
    '
    'meeting start times
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TimeVal, " & _
                  " Comment) " & _
                  " VALUES ('SundayMeetingStartTime', " & _
                          "#10:00#, " & _
                          " '')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TimeVal, " & _
                  " Comment) " & _
                  " VALUES ('MidWeekMeetingStartTime', " & _
                          "#19:30#, " & _
                          " '')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.65.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_65_00_To_5_66_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, NewField As DAO.Field

    'new "Show in Calendar" column
    CMSDB.TableDefs.Refresh
    Set NewField = CMSDB.TableDefs("tblEventLookup").CreateField("ShowInCalendar", dbBoolean)
    CMSDB.TableDefs("tblEventLookup").Fields.Append NewField
    NewField.Required = True
    
    CMSDB.TableDefs.Refresh
 
    
    CMSDB.Execute "UPDATE tblEventLookup " & _
                 "SET ShowInCalendar = TRUE "
    CMSDB.Execute "UPDATE tblEventLookup " & _
                 "SET ShowInCalendar = FALSE " & _
                 "WHERE EventID = 14"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.66.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_66_00_To_5_67_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add new Public Meeting button to security table for General Admin and CMS admin
    '
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdPublicMeeting', " & _
                          "1)"

    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdPublicMeeting', " & _
                          "5)"
        
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.67.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_67_00_To_5_68_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    DeleteTable "tblVisitingSpeakers"
    
    '
    'Now build new tblVisitingSpeakers table
    '
    CreateTable ErrorCode, "tblVisitingSpeakers", "PersonID", "LONG", , , False
    CreateField ErrorCode, "tblVisitingSpeakers", "CongNo", "LONG"
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblVisitingSpeakers " & _
                  "   (PersonID) " & _
                  "WITH PRIMARY"
                  
    '
    'Add new Public Meeting button to security table for General Admin and CMS admin
    '
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmPersonalDetails', " & _
                          " 'cmdPublicMeetingPersonnel', " & _
                          "1)"

    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmPersonalDetails', " & _
                          " 'cmdPublicMeetingPersonnel', " & _
                          "5)"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.68.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_5_68_00_To_5_69_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, fso As New Scripting.FileSystemObject
Dim fsoTextStream As Scripting.TextStream, fsoFile As Scripting.File
Dim TheFileRec As String, lTalkNo As Long, sTalkName As String, StrArr() As String
Dim lSubjectID As Long, sSubjectName As String

    DeleteTable "tblPublicTalkSubjectGroups"
    DeleteTable "tblPublicTalkOutlines"
    
    '
    'tblPublicTalkOutlines
    '
    CreateTable ErrorCode, "tblPublicTalkOutlines", "TalkNo", "LONG", , , False
    CreateField ErrorCode, "tblPublicTalkOutlines", "SubjectGroupID", "LONG"
    CreateField ErrorCode, "tblPublicTalkOutlines", "TalkTitle", "TEXT", "100"
    
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblPublicTalkOutlines " & _
                  "   (TalkNo,SubjectGroupID) " & _
                  "WITH PRIMARY"
    
    '
    'tblPublicTalkSubjectGroups
    '
    CreateTable ErrorCode, "tblPublicTalkSubjectGroups", "SubjectGroupID", "LONG", , , False
    CreateField ErrorCode, "tblPublicTalkSubjectGroups", "SubjectGroupName", "TEXT", "100"
    
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblPublicTalkSubjectGroups " & _
                  "   (SubjectGroupID) " & _
                  "WITH PRIMARY"
                  
    'load tables
    If Not fso.FileExists(JustTheDirectory & "\Public Talk Subjects.txt") Then
        MsgBox "Cannot find 'Public Talk Subjects.txt'", vbOKOnly + vbCritical, AppName
        EndProgram
    End If
    
    Set fsoFile = fso.GetFile(JustTheDirectory & "\Public Talk Subjects.txt")
    Set fsoTextStream = fsoFile.OpenAsTextStream
    
    With fsoTextStream
    
    Do Until .AtEndOfStream
        TheFileRec = Trim(RemoveNonPrintingChars(.ReadLine()))
        
        StrArr() = Split(TheFileRec, ",")
        
        lSubjectID = StrArr(0)
        sSubjectName = Trim(DoubleUpSingleQuotes(StrArr(1)))
        
        CMSDB.Execute "INSERT INTO tblPublicTalkSubjectGroups " & _
                      "(SubjectGroupID, " & _
                      " SubjectGroupName) " & _
                      " VALUES(" & lSubjectID & ", '" & _
                                   sSubjectName & "')"
    Loop
    
    .Close
    
    End With
    
    'load tables
    If Not fso.FileExists(JustTheDirectory & "\Public Talks Outlines.txt") Then
        MsgBox "Cannot find 'Public Talks Outlines.txt'", vbOKOnly + vbCritical, AppName
        EndProgram
    End If
    
    Set fsoFile = fso.GetFile(JustTheDirectory & "\Public Talks Outlines.txt")
    Set fsoTextStream = fsoFile.OpenAsTextStream
    
    With fsoTextStream
    
    ReDim StrArr(0, 0, 0)
    
    Do Until .AtEndOfStream
        TheFileRec = Trim(RemoveNonPrintingChars(.ReadLine()))
        
        StrArr() = Split(TheFileRec, ",", 3)
        
        lTalkNo = StrArr(1)
        lSubjectID = StrArr(0)
        sTalkName = Trim(DoubleUpSingleQuotes(StrArr(2)))
        
        CMSDB.Execute "INSERT INTO tblPublicTalkOutlines " & _
                      "(TalkNo, " & _
                      " SubjectGroupID, " & _
                      " TalkTitle) " & _
                      " VALUES(" & lTalkNo & ", " & lSubjectID & ", '" & _
                                   sTalkName & "')"
    Loop
    
    .Close
    
    End With
    
    fso.DeleteFile JustTheDirectory & "\Public Talk Subjects.txt", True
    fso.DeleteFile JustTheDirectory & "\Public Talks Outlines.txt", True
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.69.0"
    
    Set fso = Nothing
    Set fsoTextStream = Nothing
    Set fsoFile = Nothing

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_69_00_To_5_70_00()
On Error GoTo ErrorTrap
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.70.0"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_5_70_00_To_5_71_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    DeleteTable "tblChairmanNotes"
    
    '
    'Now build new tblChairmanNotes table
    '
    CreateTable ErrorCode, "tblChairmanNotes", "MeetingDate", "DATE", , , False
    CreateField ErrorCode, "tblChairmanNotes", "Chairman", "LONG"
    CreateField ErrorCode, "tblChairmanNotes", "SongForTalk", "LONG"
    CreateField ErrorCode, "tblChairmanNotes", "ThisWeekTalkTitle", "TEXT", "100"
    CreateField ErrorCode, "tblChairmanNotes", "ThisWeekSpeaker", "LONG"
    CreateField ErrorCode, "tblChairmanNotes", "ThisWeekCong", "LONG"
    CreateField ErrorCode, "tblChairmanNotes", "NextWeekTalkTitle", "TEXT", "100"
    CreateField ErrorCode, "tblChairmanNotes", "NextWeekSpeaker", "LONG"
    CreateField ErrorCode, "tblChairmanNotes", "NextWeekCong", "LONG"
    CreateField ErrorCode, "tblChairmanNotes", "Announcements", "MEMO"
    CreateField ErrorCode, "tblChairmanNotes", "SongForWT", "LONG"
    
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblChairmanNotes " & _
                  "   (MeetingDate) " & _
                  "WITH PRIMARY"
                      
    CMSDB.TableDefs.Refresh
    CMSDB.TableDefs("tblChairmanNotes").Fields("ThisWeekTalkTitle").AllowZeroLength = True
    CMSDB.TableDefs("tblChairmanNotes").Fields("ThisWeekTalkTitle").AllowZeroLength = True
    CMSDB.TableDefs("tblChairmanNotes").Fields("Announcements").AllowZeroLength = True
    
    CMSDB.TableDefs.Refresh
                      
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.71.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_71_00_To_5_72_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, fso As New Scripting.FileSystemObject
Dim fsoTextStream As Scripting.TextStream, fsoFile As Scripting.File
Dim TheFileRec As String, SongNo As Long, SongName As String

    DeleteTable "tblSongs"
    '
    'tblSongs
    '
    CreateTable ErrorCode, "tblSongs", "SongNo", "LONG", , , False
    CreateField ErrorCode, "tblSongs", "SongTitle", "TEXT", "100"
    
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblSongs " & _
                  "   (SongNo) " & _
                  "WITH PRIMARY"
    
'    If Not fso.FileExists(App.Path & "\All Songs.txt") Then
'        MsgBox "Cannot find 'All Songs.txt'", vbOKOnly + vbCritical, AppName
'        EndProgram
'    End If
'
'    Set fsoFile = fso.GetFile(App.Path & "\All Songs.txt")
'    Set fsoTextStream = fsoFile.OpenAsTextStream
'
'    With fsoTextStream
'
'    Do Until .AtEndOfStream
'        TheFileRec = Trim(RemoveNonPrintingChars(.ReadLine()))
'        Select Case True
'        Case IsNumeric(Right(TheFileRec, 3))
'            SongNo = CLng(Right(TheFileRec, 3))
'            SongName = Trim$(DoubleUpSingleQuotes(Left(TheFileRec, Len(TheFileRec) - 3)))
'        Case IsNumeric(Right(TheFileRec, 2))
'            SongNo = CLng(Right(TheFileRec, 2))
'            SongName = Trim$(DoubleUpSingleQuotes(Left(TheFileRec, Len(TheFileRec) - 2)))
'        Case IsNumeric(Right(TheFileRec, 1))
'            SongNo = CLng(Right(TheFileRec, 1))
'            SongName = Trim$(DoubleUpSingleQuotes(Left(TheFileRec, Len(TheFileRec) - 1)))
'        Case Else
'            SongNo = 0
'            SongName = ""
'        End Select
'
'        If SongNo > 0 Then
'            CMSDB.Execute "INSERT INTO tblSongs " & _
'                          "(SongNo, " & _
'                          " SongTitle) " & _
'                          " VALUES(" & SongNo & ", '" & _
'                                       SongName & "')"
'        End If
'    Loop
'
'    .Close
'
'    End With
'
'    fso.DeleteFile App.Path & "\All Songs.txt", True
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.72.0"
    
'    Set fso = Nothing
'    Set fsoTextStream = Nothing
'    Set fsoFile = Nothing

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_72_00_To_5_73_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Support for SMS
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('ModemCOMPort', " & _
                          " 1, " & _
                          " 'Initial value = 1')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('SMSCentreNo', " & _
                          " '07785 499993', " & _
                          " 'Initial value = 9,07785 499993 for Vodaphone')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.73.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_73_00_To_5_74_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Support for SMS
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('ClosingSalutation', " & _
                          " 'Thanks, ', " & _
                          " 'Initial value = Thanks')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.74.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_74_00_To_5_75_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add new CO Task
    '
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description) " & _
                  " VALUES (5, " & _
                          " 8, " & _
                          " 92, " & _
                          " 'Circuit Overseer')"
                          
    '
    'Now build new tblServiceMtgs table
    '
    DeleteTable "tblServiceMtgs"
    CreateTable ErrorCode, "tblServiceMtgs", "MeetingDate", "DATE", , , True
    CreateField ErrorCode, "tblServiceMtgs", "ItemTypeID", "LONG"
    CreateField ErrorCode, "tblServiceMtgs", "ItemName", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgs", "ItemLength", "LONG"
    CreateField ErrorCode, "tblServiceMtgs", "PersonID", "LONG"
    CMSDB.TableDefs.Refresh
    CMSDB.TableDefs("tblServiceMtgs").Fields("ItemName").AllowZeroLength = True
    
    CMSDB.TableDefs.Refresh
       
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.75.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_75_00_To_5_76_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                          
    '
    'Now build new tblPublicMtgSchedule table
    '
    DeleteTable "tblPublicMtgSchedule"
    CreateTable ErrorCode, "tblPublicMtgSchedule", "MeetingDate", "DATE", , , False
    CreateField ErrorCode, "tblPublicMtgSchedule", "SpeakerID", "LONG"
    CreateField ErrorCode, "tblPublicMtgSchedule", "TalkNo", "LONG"
    CreateField ErrorCode, "tblPublicMtgSchedule", "ChairmanID", "LONG"
    CreateField ErrorCode, "tblPublicMtgSchedule", "WTReaderID", "LONG"
       
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblPublicMtgSchedule " & _
                  "   (MeetingDate) " & _
                  "WITH PRIMARY"
                  
                  
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('ServiceMeetingDuration', " & _
                          " 45, " & _
                          " 'Initial value = 45. In minutes.')"
                  
       
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.76.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_76_00_To_5_77_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                          
    CreateField ErrorCode, "tblServiceMtgs", "Announcements", "MEMO"
    CMSDB.TableDefs.Refresh
    CMSDB.TableDefs("tblServiceMtgs").Fields("Announcements").AllowZeroLength = True
    
    CMSDB.TableDefs.Refresh
       
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.77.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_77_00_To_5_78_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Support for Email to SMS
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('EmailToSMSType', " & _
                          " 0, " & _
                          " 'Initial value = 0.')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.78.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_78_00_To_5_79_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, fld As DAO.Field
                          
    '
    'Now build new tblServiceMtgSchedulePrint table
    '
    DeleteTable "tblServiceMtgSchedulePrint"
    CreateTable ErrorCode, "tblServiceMtgSchedulePrint", "MeetingDate", "TEXT"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "OpeningSong", "TEXT"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item1Len", "TEXT"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item2Len", "TEXT"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item3Len", "TEXT"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item4Len", "TEXT"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item5Len", "TEXT"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item1Bro", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item2Bro", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item3Bro", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item4Bro", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item5Bro", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item1Name", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item2Name", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item3Name", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item4Name", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item5Name", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item1End", "TEXT"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item2End", "TEXT"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item3End", "TEXT"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item4End", "TEXT"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "Item5End", "TEXT"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "ConcSong", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "ConcPrayerBro", "TEXT", "100"
    
    CMSDB.TableDefs.Refresh
    
    For Each fld In CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields
        If fld.Name <> "SeqNum" Then
            fld.AllowZeroLength = True
            fld.Required = False
        End If
    Next
    
    CMSDB.TableDefs.Refresh
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.79.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_79_00_To_5_80_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " Comment) " & _
                  " VALUES ('NextSchedulePrintStartDate', " & _
                          " 'Initial value = Null')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.80.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_80_00_To_5_81_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer


    '
    'Support for Email to SMS
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('ShowEmailForSMS', " & _
                          " False, " & _
                          " 'Initial value = False.')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.81.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_81_00_To_5_82_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                          
    '
    'Now build new tblPublicMtgWtg table
    '
    DeleteTable "tblPublicMtgWtg"
    CreateTable ErrorCode, "tblPublicMtgWtg", "PersonID", "LONG", , , False
    CreateField ErrorCode, "tblPublicMtgWtg", "Weighting", "LONG"
    CreateField ErrorCode, "tblPublicMtgWtg", "ReaderLastTime", "YESNO"
    CreateField ErrorCode, "tblPublicMtgWtg", "ChairmanLastTime", "YESNO"
       
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblPublicMtgWtg " & _
                  "   (PersonID) " & _
                  "WITH PRIMARY"
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.82.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_82_00_To_5_83_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, NewField As DAO.Field

    DelAllRows "tblPublicMtgWtg"
    CMSDB.TableDefs.Refresh
    CMSDB.TableDefs("tblPublicMtgWtg").Fields("ReaderLastTime").Name = "ReaderCount"
    CMSDB.TableDefs("tblPublicMtgWtg").Fields("ChairmanLastTime").Name = "ChairmanCount"
    CMSDB.TableDefs.Refresh
    
    CMSDB.Execute "ALTER TABLE tblPublicMtgWtg " & _
                  "ALTER COLUMN ReaderCount LONG; "
    
    CMSDB.Execute "ALTER TABLE tblPublicMtgWtg " & _
                  "ALTER COLUMN ChairmanCount LONG; "
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.83.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_83_00_To_5_84_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('ReaderWeighting', " & _
                          " 1, " & _
                          " 'Initial value = 1.')"
        
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('ChairmanWeighting', " & _
                          " 2, " & _
                          " 'Initial value = 2.')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.84.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_84_00_To_5_85_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                          
    '
    'Now build new tblCongMinCardRowPrinted table
    '
    DeleteTable "tblCongMinCardRowPrinted"
    CreateTable ErrorCode, "tblCongMinCardRowPrinted", "MinServiceYear", "LONG", , , False
    CreateField ErrorCode, "tblCongMinCardRowPrinted", "MinMonth", "LONG"
    CreateField ErrorCode, "tblCongMinCardRowPrinted", "MinType", "LONG"
       
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblCongMinCardRowPrinted " & _
                  "   (MinServiceYear, MinMonth, MinType ) " & _
                  "WITH PRIMARY"
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.85.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_85_00_To_5_86_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    '
    'Export Public Mtg and Serv mtg schedules
    '
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmExportDB', " & _
                          " 'chkExportItem(8)', " & _
                          "1)"
                          
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmExportDB', " & _
                          " 'chkExportItem(8)', " & _
                          "5)"
                
    CMSDB.Execute "INSERT INTO tblExportDetails " & _
                  "(ExportDataType, " & _
                  " OrderingForSQL, " & _
                  " IncludeForExport, " & _
                  " Description) " & _
                  " VALUES (9, " & _
                          " 900, " & _
                          "FALSE, " & _
                          "'Public and Service mtg schedules')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('ExportItem8', " & _
                          " False, " & _
                          " 'Initial value = FALSE (Public and Service Mtg Schedules)')"

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('ImportItem8', " & _
                          " False, " & _
                          " 'Initial value = FALSE (Public and Service Mtg Schedules)')"
                          
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.86.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_86_00_To_5_87_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add new publisher record card print parameters to tblConstants
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepCongNameXPos', " & _
                          " 1.3, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepCongNameYPos', " & _
                          " 0.85, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepMonthYearXPos', " & _
                          " 1.7, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepMonthYearYPos', " & _
                          " 1.65, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepCongNoXPos', " & _
                          " 11.3, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepCongNoYPos', " & _
                          " 1.65, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepPubFiguresYPos', " & _
                          " 3.3, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepRegFiguresYPos', " & _
                          " 4.7, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepAuxFiguresYPos', " & _
                          " 4.0, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepNoReportingXPos', " & _
                          " 2.1, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepBooksXPos', " & _
                          " 3.8, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepBrochuresXPos', " & _
                          " 5.3, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepHoursXPos', " & _
                          " 6.7, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepMagsXPos', " & _
                          " 8.3, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepRVsXPos', " & _
                          " 9.9, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepStudiesXPos', " & _
                          " 11.5, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepTopMargin', " & _
                          " 5.8, " & _
                          " 'In cm..')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepLeftMargin', " & _
                          " 0.50, " & _
                          " 'In cm..')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepFontSize', " & _
                          " 8, " & _
                          " ' ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES('CongRepFontName', " & _
                          " 'Arial', " & _
                          " '.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepPaperHeight', " & _
                          " 8.0, " & _
                          " 'In cm.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepPaperWidth', " & _
                          " 14.0, " & _
                          " 'In cm.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepTweakX', " & _
                          " 0, " & _
                          " ' ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('CongRepTweakY', " & _
                          " 0, " & _
                          " ' ')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.87.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_87_00_To_5_88_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add new publisher record card print parameters to tblConstants
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES('EmailSMSWebpage', " & _
                          " 'www.sms2email.com', " & _
                          " ' ')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.88.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_88_00_To_5_89_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
Dim rs As Recordset, dtMaxDate As Date, dtTempDate As Date
Dim str As String
                          
    Set rs = CMSDB.OpenRecordset("SELECT MAX(ActualMinPeriod) AS MaxDate " & _
                                 "FROM tblMinreports ", dbOpenForwardOnly)
    
    If Not IsNull(rs!MaxDate) Then
        dtMaxDate = rs!MaxDate
        
        Set rs = CMSDB.OpenRecordset("SELECT PersonID, StartDate " & _
                                 "FROM tblSpecPioDates ", dbOpenForwardOnly)
        
        
        If Not rs.BOF Then
            On Error Resume Next
            Do Until rs.EOF
                dtTempDate = rs!StartDate
                Do Until dtTempDate > dtMaxDate
                    str = "INSERT INTO tblPubRecCardRowPrinted " & _
                            "VALUES (" & rs!PersonID & ", #" & _
                                        Format(dtTempDate, "mm/dd/yyyy") & "#, FALSE)"
                    
                    CMSDB.Execute (str)
                                                
                    CMSDB.Execute ("INSERT INTO tblMinReports " & _
                                    "VALUES (" & rs!PersonID & ", " & _
                                                Month(dtTempDate) & ", " & _
                                                year(ConvertNormalDateToServiceDate(dtTempDate)) & ", " & _
                                                Month(dtTempDate) & ", " & _
                                                year(ConvertNormalDateToServiceDate(dtTempDate)) & ", " & _
                                                "0, 0, 120, 0, 0, 0, '', #" & _
                                                Format(dtTempDate, "mm/dd/yyyy") & "#, #" & _
                                                Format(dtTempDate, "mm/dd/yyyy") & "#)")
                                                
                    
                    dtTempDate = DateAdd("m", 1, dtTempDate)
                Loop
                rs.MoveNext
            Loop
            On Error GoTo ErrorTrap
        End If
        
    End If
    
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES('SpecialPioHours', " & _
                          " 120, " & _
                          " 'Number of hours Special Pios do each month ')"
    
    DelAllRows "tblMissingReports"
    DelAllRows "tblInactivePubs"
    DelAllRows "tblIrregularPubs"
    PutAllMissingReportsIntoTable "01/09/2003", "01/12/9999"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.89.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_5_89_00_To_5_90_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                          
    CMSDB.TableDefs.Refresh
    CreateField ErrorCode, "tblEvents", "PrivateEntry", "YESNO"
    CreateField ErrorCode, "tblEvents", "AlarmDateTime", "DATE"
    CreateField ErrorCode, "tblEvents", "AlarmAcknowledged", "YESNO"
    CMSDB.TableDefs.Refresh
       
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.90.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_90_00_To_5_91_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                          
    CMSDB.TableDefs.Refresh
    
    CreateField ErrorCode, "tblCong", "SundayMtgDay", "LONG"
    CreateField ErrorCode, "tblCong", "SundayMtgTime", "DATE"
    CMSDB.Execute ("UPDATE tblCong " & _
                  "SET SundayMtgDay = 1, SundayMtgTime = #13:00#")
    
    CreateField ErrorCode, "tblPublicMtgSchedule", "CongNoWhereMtgIs", "LONG"
    CMSDB.Execute ("UPDATE tblPublicMtgSchedule " & _
                  "SET CongNoWhereMtgIs = " & GlobalParms.GetValue("DefaultCong", "NumVal"))
                  
    DropIndex ErrorCode, "tblPublicMtgSchedule", "IX1"
    
    CMSDB.Execute "CREATE INDEX IX1 " & _
              "ON tblPublicMtgSchedule " & _
              "   (MeetingDate, CongNoWhereMtgIs) " & _
              "WITH PRIMARY"

    
    CMSDB.TableDefs.Refresh
    
       
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.91.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_91_00_To_5_92_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " DateVal, " & _
                  " Comment) " & _
                  " VALUES('LastVisitingScheduleDate', " & _
                          " #01/01/2000#, " & _
                          " ' ')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.92.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_5_92_00_To_5_93_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add new visiting speaker
    '
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description) " & _
                  " VALUES (4, " & _
                          " 7, " & _
                          " 93, " & _
                          " 'Public Talks - Outbound')"
                          
    '
    'Now build new  table
    '
    DeleteTable "tblSpeakersTalks"
    CreateTable ErrorCode, "tblSpeakersTalks", "PersonID", "LONG", , , False
    CreateField ErrorCode, "tblSpeakersTalks", "TalkNo", "LONG"
        
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblSpeakersTalks " & _
                  "   (PersonID, TalkNo) " & _
                  "WITH PRIMARY"
                              
    CMSDB.TableDefs.Refresh
       
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.93.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_93_00_To_5_94_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES('2ndBackupLocation', " & _
                          " 'C:\', " & _
                          " ' ')"
    
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('Do2ndBackup', " & _
                          " False, " & _
                          " 'Initial value = FALSE ')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.94.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_94_00_To_5_95_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, NewField As DAO.Field
                          
    CMSDB.Execute ("INSERT INTO tblCong " & _
                  "VALUES (32767, 'Travelling Overseers', 1, #00:00#)")
           
    'new  column
    CMSDB.TableDefs.Refresh
    Set NewField = CMSDB.TableDefs("tblPublicMtgSchedule").CreateField("Info", dbText)
    CMSDB.TableDefs("tblPublicMtgSchedule").Fields.Append NewField
    NewField.Required = False
    NewField.AllowZeroLength = True
    
    CMSDB.TableDefs.Refresh
    
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.95.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_95_00_To_5_96_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, NewField As DAO.Field
                          
    'new  column
    CMSDB.TableDefs.Refresh
    Set NewField = CMSDB.TableDefs("tblTasks").CreateField("AllowSuspend", dbBoolean)
    CMSDB.TableDefs("tblTasks").Fields.Append NewField
    NewField.DefaultValue = False
    
    CMSDB.TableDefs.Refresh
    
    CMSDB.Execute ("UPDATE tblTasks " & _
                  "SET AllowSuspend = TRUE " & _
                  "WHERE Task IN (10,19,20,21,47,48,49,93) " & _
                  "OR Task BETWEEN 57 AND 60 " & _
                  "OR Task BETWEEN 33 AND 43 ")
                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.96.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_5_96_00_To_5_97_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, fso As New Scripting.FileSystemObject
Dim fsoTextStream As Scripting.TextStream, fsoFile As Scripting.File
Dim TheFileRec As String, lSongNo As Long, sSongName As String, StrArr() As String
Dim lSubjectID As Long, sSubjectName As String
'
'    If Not fso.FileExists(App.Path & "\Song Subjects.txt") Then
'        MsgBox "Cannot find 'Song Subjects.txt'", vbOKOnly + vbCritical, AppName
'        EndProgram
'    End If
'    If Not fso.FileExists(App.Path & "\SongsAndSubjects.txt") Then
'        MsgBox "Cannot find 'SongsAndSubjects.txt'", vbOKOnly + vbCritical, AppName
'        EndProgram
'    End If
'
    DeleteTable "tblSongSubjects"
    DeleteTable "tblSongNoAndSubject"

    CreateTable ErrorCode, "tblSongSubjects", "SubjectGroupID", "LONG", , , False
    CreateField ErrorCode, "tblSongSubjects", "SubjectGroupName", "TEXT", "100"

    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblSongSubjects " & _
                  "   (SubjectGroupID) " & _
                  "WITH PRIMARY"
'
'    'load tables
'
'    Set fsoFile = fso.GetFile(App.Path & "\Song Subjects.txt")
'    Set fsoTextStream = fsoFile.OpenAsTextStream
'
'    With fsoTextStream
'
'    Do Until .AtEndOfStream
'        TheFileRec = Trim(RemoveNonPrintingChars(.ReadLine()))
'
'        StrArr() = Split(TheFileRec, ",")
'
'        lSubjectID = StrArr(0)
'        sSubjectName = Trim(DoubleUpSingleQuotes(StrArr(1)))
'
'        CMSDB.Execute "INSERT INTO tblSongSubjects " & _
'                      "(SubjectGroupID, " & _
'                      " SubjectGroupName) " & _
'                      " VALUES(" & lSubjectID & ", '" & _
'                                   sSubjectName & "')"
'    Loop
'
'    .Close
'
'    End With
'
'
    CreateTable ErrorCode, "tblSongNoAndSubject", "SongNo", "LONG", , , False
    CreateField ErrorCode, "tblSongNoAndSubject", "SubjectGroupID", "LONG"

    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblSongNoAndSubject " & _
                  "   (SongNo,SubjectGroupID) " & _
                  "WITH PRIMARY"
'
'    'load tables
'
'    Set fsoFile = fso.GetFile(App.Path & "\SongsAndSubjects.txt")
'    Set fsoTextStream = fsoFile.OpenAsTextStream
'
'    With fsoTextStream
'
'    ReDim StrArr(0, 0, 0)
'
'    Do Until .AtEndOfStream
'        TheFileRec = Trim(RemoveNonPrintingChars(.ReadLine()))
'
'        StrArr() = Split(TheFileRec, ",", 2)
'
'        lSongNo = StrArr(1)
'        lSubjectID = StrArr(0)
'
'        CMSDB.Execute "INSERT INTO tblSongNoAndSubject " & _
'                      "(SongNo, " & _
'                      " SubjectGroupID) " & _
'                      " VALUES(" & lSongNo & ", " & lSubjectID & ")"
'    Loop
'
'    .Close
'
'    End With
'
    If fso.FileExists(JustTheDirectory & "\SongsAndSubjects.txt") Then
        fso.DeleteFile JustTheDirectory & "\SongsAndSubjects.txt", True
    End If
    If fso.FileExists(JustTheDirectory & "\Song Subjects.txt") Then
        fso.DeleteFile JustTheDirectory & "\Song Subjects.txt", True
    End If
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.97.0"
    
'    Set fso = Nothing
'    Set fsoTextStream = Nothing
'    Set fsoFile = Nothing

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_97_00_To_5_98_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES('AutoFillSpecialPioHrs', " & _
                          " False, " & _
                          " ' ')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.98.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_98_00_To_5_99_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES('CongMinRotaFrequency', " & _
                          " 0, " & _
                          " 'Weekly = 0, Monthly = 1')"
        
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES('CongMinRotaDuration', " & _
                          " 15, " & _
                          " 'No weeks/months')"
        
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " DateVal, " & _
                  " Comment) " & _
                  " VALUES('CongMinRotaLastRotaDate', " & _
                          " #01/01/2000#, " & _
                          " '')"
                          
    '
    'Add new role
    '
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description, " & _
                  " AllowSuspend) " & _
                  " VALUES (6, " & _
                          " 13, " & _
                          " 94, " & _
                          " 'Congregation Field Service Group Leader', " & _
                          " TRUE)"
                          
    CreateTable ErrorCode, "tblCongMinRota", "RotaDate", "DATE", , , True
    CreateField ErrorCode, "tblCongMinRota", "DateForPrint", "TEXT"
    CreateField ErrorCode, "tblCongMinRota", "PersonID", "LONG"
    CreateField ErrorCode, "tblCongMinRota", "PersonName", "TEXT"
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.99.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_99_00_To_5_9900_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, NewField As DAO.Field
                              
    CMSDB.Execute ("UPDATE tblPublisherDates " & _
                  "SET StartReason = 3 " & _
                  "WHERE StartReason IS NULL ")
                  
    CMSDB.Execute ("UPDATE tblPublisherDates " & _
                  "SET EndReason = 2 " & _
                  "WHERE EndReason IS NULL ")
                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9900.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_9900_00_To_5_9901_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmEditPublicMtgSchedule', " & _
                          " 'cmdPersonnel', " & _
                          "1)"

    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmEditPublicMtgSchedule', " & _
                          " 'cmdPersonnel', " & _
                          "5)"
        
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmEditVisitingSchedule', " & _
                          " 'cmdPersonnel', " & _
                          "1)"

    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmEditVisitingSchedule', " & _
                          " 'cmdPersonnel', " & _
                          "5)"
                          
    CMSDB.Execute "INSERT INTO tblAccessLevelDescriptions " & _
                  "(OrderingKey, AccessLevel, AccessDesc) " & _
                  "VALUES " & _
                  "(70, 8, 'Public Talks')"
                  
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdPublicMeeting', " & _
                          "8)"
        
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdServiceMtg', " & _
                          "1)"
        
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdServiceMtg', " & _
                          "5)"
        
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmPublicMeetingMenu', " & _
                          " 'cmdChairmansNotes', " & _
                          "1)"
        
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmPublicMeetingMenu', " & _
                          " 'cmdChairmansNotes', " & _
                          "5)"
        
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmPublicMeetingMenu', " & _
                          " 'cmdPersonnel', " & _
                          "1)"
        
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmPublicMeetingMenu', " & _
                          " 'cmdPersonnel', " & _
                          "5)"
        
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmPublicMeetingMenu', " & _
                          " 'cmdTalkOutlines', " & _
                          "1)"
        
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmPublicMeetingMenu', " & _
                          " 'cmdTalkOutlines', " & _
                          "5)"
        
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9901.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Private Sub DB_Upgrade_5_9901_00_To_5_9902_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('PublicMtgScheduleAlert', " & _
                          " True, " & _
                          " 14, " & _
                          " 'Initial value = True; 14 days.')"
        
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('SPAMRotaAlert', " & _
                          " True, " & _
                          " 14, " & _
                          " 'Initial value = True; 14 days.')"
        
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('CleaningRotaAlert', " & _
                          " True, " & _
                          " 14, " & _
                          " 'Initial value = True; 14 days.')"
        
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9902.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_9902_00_To_5_9903_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSSchedulingShowQuickNames', " & _
                          " True, " & _
                          " 'Initial value = True')"
        
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9903.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_9903_00_To_5_9904_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('AnnualRegPioHours', " & _
                          " 840, " & _
                          " 'Initial value = 840')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('AnnualSpecPioHours', " & _
                          " 1440, " & _
                          " 'Initial value = 1440')"
        
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9904.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_9904_00_To_5_9905_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    DeleteTable "tblPioHourCredit"
    
    CreateTable ErrorCode, "tblPioHourCredit", "PersonID", "LONG", , , True
    CreateField ErrorCode, "tblPioHourCredit", "MinDate", "DATE"
    CreateField ErrorCode, "tblPioHourCredit", "NoHours", "LONG"
        
    CMSDB.TableDefs.Refresh
                      
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9905.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_9905_00_To_5_9906_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('MonthlyAuxPioHours', " & _
                          " 50, " & _
                          " 'Initial value = 50')"
                   
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9906.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_9906_00_To_5_9907_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, NewField As DAO.Field

    'new  column
    CMSDB.TableDefs.Refresh
    Set NewField = CMSDB.TableDefs("tblMinReports").CreateField("OtherComments", dbText)
    CMSDB.TableDefs("tblMinReports").Fields.Append NewField
    NewField.Required = False
    NewField.AllowZeroLength = True
    
    CMSDB.TableDefs.Refresh
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9907.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_9907_00_To_5_9908_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
Dim str As String
Dim tdf As TableDef
Dim prp As DAO.Property, bPropertyExists As Boolean
    
    '
    'Add Description property to each table
    ' First, must add Description property to each table
    '
    For Each tdf In CMSDB.TableDefs
    
        'scan all properties for each table. If no Description property exists,
        ' then add one.
        
        For Each prp In tdf.Properties
        
            If prp.Name = "Description" Then
                bPropertyExists = True
                Exit For
            End If
            
        Next
        
        If Not bPropertyExists Then
            If Left(tdf.Name, 3) = "tbl" Then
                tdf.Properties.Append tdf.CreateProperty("Description", _
                                                           dbText, "CMS")
            End If
        End If
    
        bPropertyExists = False
            
    Next
    
    str = "Text description of each CMS access level. Access levels " & _
            "define what privileges are allowed. Access levels are assigned " & _
            "to users in tblAccessLevels"
    CMSDB.TableDefs("tblAccessLevelDescriptions").Properties("Description") = str
    
    str = "Access levels assigned to users"
    CMSDB.TableDefs("tblAccessLevels").Properties("Description") = str
    
    str = "Date ranges in which publishers aux pio"
    CMSDB.TableDefs("tblAuxPioDates").Properties("Description") = str
    
    str = "Baptism dates"
    CMSDB.TableDefs("tblBaptismDates").Properties("Description") = str
    CMSDB.TableDefs.Refresh
    
    str = "Who assigned to each group"
    CMSDB.TableDefs("tblBookGroupMembers").Properties("Description") = str
    
    str = "List of Book Groups"
    CMSDB.TableDefs("tblBookGroups").Properties("Description") = str
    
    str = "Notes for public meeting chairman, by date"
    CMSDB.TableDefs("tblChairmanNotes").Properties("Description") = str
    
    str = "Link parents to children"
    CMSDB.TableDefs("tblChildren").Properties("Description") = str
    
    str = "Latest cleaning rota"
    CMSDB.TableDefs("tblCleaningRota").Properties("Description") = str
    
    str = "List all CMS DB backups"
    CMSDB.TableDefs("tblCMSBackups").Properties("Description") = str
    
    str = "Email addresses of CMS users"
    CMSDB.TableDefs("tblCMSUsersEmailAddresses").Properties("Description") = str
    
    str = "All congs and their Sunday meeting times"
    CMSDB.TableDefs("tblCong").Properties("Description") = str
    
    str = "Shows whether cong month report (for Branch) has been printed"
    CMSDB.TableDefs("tblCongMinCardRowPrinted").Properties("Description") = str
    
    str = "Rota for weekly cong min arrangement leader"
    CMSDB.TableDefs("tblCongMinRota").Properties("Description") = str
    
    str = "CMS system parameters"
    CMSDB.TableDefs("tblConstants").Properties("Description") = str
    
    str = "7 days of week listed"
    CMSDB.TableDefs("tblDayLookup").Properties("Description") = str
    
    str = "Shows all appointed men with E/MS code and appointment date"
    CMSDB.TableDefs("tblEldersAndServants").Properties("Description") = str
    
    str = "Shows lookup of all calendar events on system"
    CMSDB.TableDefs("tblEventLookup").Properties("Description") = str
    
    str = "Shows lookup of all calendar recurring types"
    CMSDB.TableDefs("tblEventRecurringlookup").Properties("Description") = str
    
     str = "Calendar stored here"
     CMSDB.TableDefs("tblEvents").Properties("Description") = str
    
     str = "Shows attendance at each group for each week"
     CMSDB.TableDefs("tblGroupAttendance").Properties("Description") = str
    
     str = "Shows inactive pubs and date from/to. Also has MissingReportGroup - this links to a contiguous group of months on tblMissingReport"
     CMSDB.TableDefs("tblInactivePubs").Properties("Description") = str
    
     str = "SPAM weightings assigned to individual brothers"
     CMSDB.TableDefs("tblIndividualSPAMWeightings").Properties("Description") = str
    
     str = "Irregular pubs, along with start and end of irregular period"
     CMSDB.TableDefs("tblIrregularPubs").Properties("Description") = str
    
     str = "Whos married to who"
     CMSDB.TableDefs("tblMarriage").Properties("Description") = str
    
     str = "Shows meeting Type (see tblMeetingTypes), week and figure "
     CMSDB.TableDefs("tblMeetingAttendance").Properties("Description") = str
    
     str = "Lookup of meeting types"
     CMSDB.TableDefs("tblMeetingTypes").Properties("Description") = str
    
     str = "Show monthly reports for all pubs. SocietyReportingMonth & Year and MinistryDoneInMonth & Year are Service-Year format. ActualMinPeriod and SocietyReportingPeriod are normal calendar dates."
     CMSDB.TableDefs("tblMinReports").Properties("Description") = str
    
     str = "Ministry Type (pub/aux/reg)"
     CMSDB.TableDefs("tblMinTypeDesc").Properties("Description") = str
    
     str = "Missing reports by person and date. ActualMinDate is in normal calendar format. If ZeroReport, then a zero report has been submitted"
     CMSDB.TableDefs("tblMissingReports").Properties("Description") = str
    
     str = "List of months"
     CMSDB.TableDefs("tblMonthName").Properties("Description") = str
    
     str = "Name and address details"
     CMSDB.TableDefs("tblNameAddress").Properties("Description") = str
    
     str = "Shows forms and objects on forms. Level of security required to access the control."
     CMSDB.TableDefs("tblObjectSecurity").Properties("Description") = str
    
     str = "Shows month (normal calendar format) and number of hours credit person has been given"
     CMSDB.TableDefs("tblPioHourCredit").Properties("Description") = str
    
     str = "schedule for public meeting"
     CMSDB.TableDefs("tblPublicMtgSchedule").Properties("Description") = str
    
    str = "All outlines and groupings"
     CMSDB.TableDefs("tblPublicTalkOutlines").Properties("Description") = str
    
     str = "Lookup of talk groups"
     CMSDB.TableDefs("tblPublicTalkSubjectGroups").Properties("Description") = str
    
     str = "Dates a publisher starts/stops publishing (normal calendar format), along with reason codes"
     CMSDB.TableDefs("tblPublisherDates").Properties("Description") = str
    
     str = "For each pub and month, shows whether card row printed."
     CMSDB.TableDefs("tblPubRecCardRowPrinted").Properties("Description") = str
    
     str = "Dates reg pio starts/stops (normal calendar format)"
     CMSDB.TableDefs("tblRegPioDates").Properties("Description") = str
    
     str = "Latest SPAM rota"
     CMSDB.TableDefs("tblRota").Properties("Description") = str
    
     str = "All users, their names, userid and code, active dates"
     CMSDB.TableDefs("tblSecurity").Properties("Description") = str
    
     str = "Service meeting speaker schedule"
     CMSDB.TableDefs("tblServiceMtgs").Properties("Description") = str
    
     str = "All song nos and subject group"
     CMSDB.TableDefs("tblSongNoAndSubject").Properties("Description") = str
    
     str = "Lookup of all songs and titles"
     CMSDB.TableDefs("tblSongs").Properties("Description") = str
    
     str = "Lookup of all songs subject groups"
     CMSDB.TableDefs("tblSongSubjects").Properties("Description") = str
    
     str = "Shows public talk outlines in each speakers collection"
     CMSDB.TableDefs("tblSpeakersTalks").Properties("Description") = str
    
     str = "Dates of spec pios"
     CMSDB.TableDefs("tblSpecPioDates").Properties("Description") = str
    
     str = "Lookup of reasons for persons task suspensions"
     CMSDB.TableDefs("tblSuspendReasons").Properties("Description") = str
    
     str = "Link Task category to person"
     CMSDB.TableDefs("tblTaskAndPerson").Properties("Description") = str
    
     str = "Lookup of Task Category and description"
     CMSDB.TableDefs("tblTaskCategories").Properties("Description") = str
    
     str = "Person, their tasks and suspend dates"
     CMSDB.TableDefs("tblTaskPersonSuspendDates").Properties("Description") = str
    
     str = " Lookup of Tasks and description "
     CMSDB.TableDefs("tblTasks").Properties("Description") = str
    
     str = " Lookup of Task Sub-Category and description "
     CMSDB.TableDefs("tblTaskSubCategories").Properties("Description") = str
    
     str = "Weightings assigned to tasks used in SPAM rota generation"
     CMSDB.TableDefs("tblTaskWeightings").Properties("Description") = str
    
     str = "Lookup of TMS assignments"
     CMSDB.TableDefs("tblTMSAssignmentsForSearch").Properties("Description") = str
    
     str = "Links Speech Qualities to their components"
     CMSDB.TableDefs("tblTMSCounselPointComponents").Properties("Description") = str
    
     str = "List of all speech qualities and description"
     CMSDB.TableDefs("tblTMSCounselPointList").Properties("Description") = str
    
     str = "Holds counsel points for each students next talk"
     CMSDB.TableDefs("tblTMSCounselPoints").Properties("Description") = str
    
     str = "All TMS items loaded into system"
     CMSDB.TableDefs("tblTMSItems").Properties("Description") = str
    
     str = "Student TMS schedule, along with their counsel point"
     CMSDB.TableDefs("tblTMSSchedule").Properties("Description") = str
    
     str = "All TMS settings"
     CMSDB.TableDefs("tblTMSSettings").Properties("Description") = str
    
     str = "TMS talk type and description"
     CMSDB.TableDefs("tblTMSTalkNoDesc").Properties("Description") = str
    
     str = "All visiting speakers and their cong"
     CMSDB.TableDefs("tblVisitingSpeakers").Properties("Description") = str
    
     str = "Week of month and description"
     CMSDB.TableDefs("tblWeekOfMonth").Properties("Description") = str
    
    CMSDB.TableDefs.Refresh
    
    '
    'Add new table to hold user queries. Set QueryID key field as AutoNumber
    '
    DeleteTable "tblUserQueries"

    CreateTable ErrorCode, "tblUserQueries", "QueryName", "TEXT", 255, , True
    CreateField ErrorCode, "tblUserQueries", "QueryString", "MEMO"
    CMSDB.TableDefs.Refresh
    
    Set tdf = CMSDB.TableDefs("tblUserQueries")
    tdf.Properties.Append tdf.CreateProperty("Description", dbText, "CMS")
    str = "User queries that can be submitted within CMS"
    CMSDB.TableDefs("tblUserQueries").Properties("Description") = str

    
    '
    'Add new Advanced Reporting button to security table for General Admin and CMS admin
    '
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdAdvancedReporting', " & _
                          "1)"

    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdAdvancedReporting', " & _
                          "5)"


    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9908.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_9908_00_To_5_9909_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, NewField As DAO.Field

    CreateField ErrorCode, "tblUserQueries", "QueryDescription", "TEXT"
    CMSDB.TableDefs.Refresh
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9909.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_9909_00_To_5_9910_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('AdvancedReportMaxRows', " & _
                          " 2000, " & _
                          " 'Initial value = 2000')"
                   
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9910.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_5_9910_00_To_5_9911_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "ALTER TABLE tblUserQueries " & _
                  "ALTER COLUMN QueryDescription TEXT(255)" & ";"
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9911.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_5_9911_00_To_5_9912_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
Dim tdf As TableDef, prp As DAO.Property


    DeleteTable "tblIndividualPioTarget"
    

    CreateTable ErrorCode, "tblIndividualPioTarget", "PersonID", "LONG", , , False
    CreateField ErrorCode, "tblIndividualPioTarget", "ServiceYear", "LONG"
    CreateField ErrorCode, "tblIndividualPioTarget", "TargetHours", "LONG"
    CreateIndex ErrorCode, "tblIndividualPioTarget", "PersonID, ServiceYear", _
                "IX1", True, False
    CMSDB.TableDefs.Refresh
    
    Set tdf = CMSDB.TableDefs("tblIndividualPioTarget")
    
    On Error Resume Next
    
    Set prp = tdf.Properties("Description")
    
    If Err.number <> 0 Then
        tdf.Properties.Append tdf.CreateProperty("Description", _
            dbText, "Holds individual annual hour targets for reg and spec pioneers")
    Else
        prp.value = "Holds individual annual hour targets for reg and spec pioneers"
    End If
    
    CMSDB.TableDefs.Refresh
    
    On Error Resume Next
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9912.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_5_9912_00_To_5_9913_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, NewField As DAO.Field

    CreateField ErrorCode, "tblPublicMtgSchedule", "SpeakerID2", "LONG"
    CreateField ErrorCode, "tblChairmanNotes", "ThisWeekSpeaker2", "LONG"
    CreateField ErrorCode, "tblChairmanNotes", "NextWeekSpeaker2", "LONG"
    CMSDB.TableDefs.Refresh
    
    CMSDB.Execute "UPDATE tblPublicMtgSchedule " & _
                  "SET SpeakerID2 = 0"
    CMSDB.Execute "UPDATE tblChairmanNotes " & _
                  "SET ThisWeekSpeaker2 = 0"
    CMSDB.Execute "UPDATE tblChairmanNotes " & _
                  "SET NextWeekSpeaker2 = 0"
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9913.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_9913_00_To_5_9914_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer


    DeleteTable "tblPrintAddresses"

    CreateTable ErrorCode, "tblPrintAddresses", "PersonName", "TEXT", "100", , True
    CreateField ErrorCode, "tblPrintAddresses", "Address", "TEXT", "255"
    CreateField ErrorCode, "tblPrintAddresses", "HomePhone", "TEXT"
    CreateField ErrorCode, "tblPrintAddresses", "MobilePhone", "TEXT"
    CMSDB.TableDefs.Refresh
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9914.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_5_9914_00_To_5_9915_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer


    CMSDB.Execute "INSERT INTO tblAccessLevelDescriptions " & _
                  "(OrderingKey, AccessLevel, AccessDesc) " & _
                  "VALUES " & _
                  "(80, 9, 'Accounts')"
    
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdAccounts', " & _
                          "1)"

    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdAccounts', " & _
                          "5)"
                          
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdAccounts', " & _
                          "9)"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9915.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_5_9915_00_To_5_9916_00()
On Error GoTo ErrorTrap
Dim tdf As TableDef, prp As DAO.Property
Dim ErrorCode As Integer

'Accounts

    DeleteTable "tblGiftAidPayers"

    CreateTable ErrorCode, "tblGiftAidPayers", "PersonID", "LONG", , , False
    CreateField ErrorCode, "tblGiftAidPayers", "GiftAidNo", "LONG"
    CreateIndex ErrorCode, "tblGiftAidPayers", "PersonID, GiftAidNo", _
                "IX1", True, False
    CMSDB.TableDefs.Refresh
    
    Set tdf = CMSDB.TableDefs("tblGiftAidPayers")
    
    On Error Resume Next
    
    Set prp = tdf.Properties("Description")
    
    If Err.number <> 0 Then
        tdf.Properties.Append tdf.CreateProperty("Description", _
            dbText, "Holds CMS PersonID and Gift Aid No")
    Else
        prp.value = "Holds CMS PersonID and Gift Aid No"
    End If
    
    CMSDB.TableDefs.Refresh
    
    On Error GoTo ErrorTrap
    
'============================================================================

    DeleteTable "tblAccInOut"

    CreateTable ErrorCode, "tblAccInOut", "Description", "TEXT", "255", , True, "InOutID"

    CMSDB.TableDefs.Refresh
    
    Set tdf = CMSDB.TableDefs("tblAccInOut")
    
    On Error Resume Next
    
    Set prp = tdf.Properties("Description")
    
    If Err.number <> 0 Then
        tdf.Properties.Append tdf.CreateProperty("Description", _
            dbText, "Describes income/expenditure.")
    Else
        prp.value = "Describes income/expenditure"
    End If
    
    CMSDB.TableDefs.Refresh
    
    On Error GoTo ErrorTrap
    
    CMSDB.Execute "INSERT INTO tblAccInOut (Description) VALUES ('Income')"
    CMSDB.Execute "INSERT INTO tblAccInOut (Description) VALUES ('Expenditure')"
    
'============================================================================
     DeleteTable "tblAccInOutTypes"

    CreateTable ErrorCode, "tblAccInOutTypes", "Description", "TEXT", "255", , True, "InOutTypeID"
    CreateField ErrorCode, "tblAccInOutTypes", "InOutID", "LONG"
    CreateField ErrorCode, "tblAccInOutTypes", "IsAuto", "YESNO"

    CMSDB.TableDefs.Refresh
    
    Set tdf = CMSDB.TableDefs("tblAccInOutTypes")
    
    On Error Resume Next
    
    Set prp = tdf.Properties("Description")
    
    If Err.number <> 0 Then
        tdf.Properties.Append tdf.CreateProperty("Description", _
            dbText, "Describes type of income/expenditure.")
    Else
        prp.value = "Describes type of income/expenditure"
    End If
    
    CMSDB.TableDefs.Refresh
    
    On Error GoTo ErrorTrap
    
    CMSDB.Execute "INSERT INTO tblAccInOutTypes (Description,InOutID,IsAuto) VALUES ('Normal Income',1,False)"
    CMSDB.Execute "INSERT INTO tblAccInOutTypes (Description,InOutID,IsAuto) VALUES ('Normal Expenditure',2,False)"
    CMSDB.Execute "INSERT INTO tblAccInOutTypes (Description,InOutID,IsAuto) VALUES ('Regular Monthly Income',1,True)"
    CMSDB.Execute "INSERT INTO tblAccInOutTypes (Description,InOutID,IsAuto) VALUES ('Regular Monthly Expenditure',2,True)"
    CMSDB.Execute "INSERT INTO tblAccInOutTypes (Description,InOutID,IsAuto) VALUES ('Temporary Income - Auto Monthly outgoing',1,False)"
    CMSDB.Execute "INSERT INTO tblAccInOutTypes (Description,InOutID,IsAuto) VALUES ('Temporary Expenditure - Auto Monthly income',2,False)"
'============================================================================
    
    DeleteTable "tblTransactionTypes"

    CreateTable ErrorCode, "tblTransactionTypes", "TranCode", "TEXT", "2", , True, "TranCodeID"
    CreateField ErrorCode, "tblTransactionTypes", "Description", "TEXT", "255"
    CreateField ErrorCode, "tblTransactionTypes", "InOutTypeID", "LONG"
    CreateField ErrorCode, "tblTransactionTypes", "AutoDayOfMonth", "LONG"
    
    CMSDB.TableDefs.Refresh
    
    Set tdf = CMSDB.TableDefs("tblTransactionTypes")
    
    On Error Resume Next
    
    Set prp = tdf.Properties("Description")
    
    If Err.number <> 0 Then
        tdf.Properties.Append tdf.CreateProperty("Description", _
            dbText, "Describes various types of transaction.")
    Else
        prp.value = "Describes various types of transaction"
    End If
    
    CMSDB.TableDefs.Refresh
    
    On Error GoTo ErrorTrap
    
'============================================================================
    DeleteTable "tblTransactionDates"

    CreateTable ErrorCode, "tblTransactionDates", "TranCodeID", "LONG", , , True, "TranID"
    CreateField ErrorCode, "tblTransactionDates", "TranDate", "DATE"
    CreateField ErrorCode, "tblTransactionDates", "FinancialYear", "LONG"
    CreateField ErrorCode, "tblTransactionDates", "FinancialMonth", "LONG"
    CreateField ErrorCode, "tblTransactionDates", "FinancialQuarter", "LONG"
    CreateField ErrorCode, "tblTransactionDates", "Amount", "DOUBLE"
    CreateField ErrorCode, "tblTransactionDates", "TranDescription", "TEXT", "255"
    CreateField ErrorCode, "tblTransactionDates", "RefNo", "LONG"
    
    CMSDB.TableDefs.Refresh
    
    Set tdf = CMSDB.TableDefs("tblTransactionDates")
    
    On Error Resume Next
    
    Set prp = tdf.Properties("Description")
    
    If Err.number <> 0 Then
        tdf.Properties.Append tdf.CreateProperty("Description", _
            dbText, "Lists all transactions by date.")
    Else
        prp.value = "Lists all transactions by date"
    End If
    
    CMSDB.TableDefs.Refresh
    
    On Error GoTo ErrorTrap
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9916.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_5_9916_00_To_5_9917_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add column to tblMonths enabling months to be ordered for Financial Year
    '
    CreateField ErrorCode, "tblMonthName", "OrderForFiscalYear", "LONG"

    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForFiscalYear = 6 " & _
                  "WHERE MonthNum = 9"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForFiscalYear = 7 " & _
                  "WHERE MonthNum = 10"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForFiscalYear = 8 " & _
                  "WHERE MonthNum = 11"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForFiscalYear = 9 " & _
                  "WHERE MonthNum = 12"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForFiscalYear = 10 " & _
                  "WHERE MonthNum = 1"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForFiscalYear = 11 " & _
                  "WHERE MonthNum = 2"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForFiscalYear = 12 " & _
                  "WHERE MonthNum = 3"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForFiscalYear = 1 " & _
                  "WHERE MonthNum = 4"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForFiscalYear = 2 " & _
                  "WHERE MonthNum = 5"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForFiscalYear = 3 " & _
                  "WHERE MonthNum = 6"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForFiscalYear = 4 " & _
                  "WHERE MonthNum = 7"
    CMSDB.Execute "UPDATE tblMonthName " & _
                  "SET OrderForFiscalYear = 5 " & _
                  "WHERE MonthNum = 8"

    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9917.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_5_9917_00_To_5_9918_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('DayInAprForStartOfTaxYear', " & _
                          " 6, " & _
                          " 'Initial value = 6')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('DayInMarForEndOfTaxYear', " & _
                          " 5, " & _
                          " 'Initial value = 5')"
                   
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9918.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_5_9918_00_To_5_9919_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
Dim tdf As TableDef, prp As DAO.Property

'Accounts

    DeleteTable "tblExtOrgs"

    CreateTable ErrorCode, "tblExtOrgs", "OrgName", "TEXT", "255", , True, "OrgID"
    CMSDB.TableDefs.Refresh
    
    Set tdf = CMSDB.TableDefs("tblExtOrgs")
    
    On Error Resume Next
    
    Set prp = tdf.Properties("Description")
    
    If Err.number <> 0 Then
        tdf.Properties.Append tdf.CreateProperty("Description", _
            dbText, "Holds all external organisations")
    Else
        prp.value = "Holds all external organisations"
    End If
    
    CMSDB.TableDefs.Refresh
    
    On Error GoTo ErrorTrap
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9919.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_5_9919_00_To_5_9920_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('GiftAidTransactionCode', " & _
                          " 'G', " & _
                          " 'Initial value = G')"
                   
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9920.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_9920_00_To_5_9921_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " DateVal, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES ('AccountBalanceAtStartOfMonth', " & _
                          " #01/01/2000#, " & _
                          " 0, " & _
                          " 'Initial values: Date = 01/01/2000; Balance = 0')"
                   
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9921.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub DB_Upgrade_5_9921_00_To_5_9922_00()
On Error Resume Next
                          
Dim iErrorCode As Integer
    
                   
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9922.0"

End Sub
Private Sub DB_Upgrade_5_9922_00_To_5_9923_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
Dim tdf As TableDef, prp As DAO.Property
                          
    DeleteTable "tblGiftAidPayerActiveDates"

    CreateTable ErrorCode, "tblGiftAidPayerActiveDates", "GiftAidNo", "LONG", , , True
    CreateField ErrorCode, "tblGiftAidPayerActiveDates", "StartDate", "DATE"
    CreateField ErrorCode, "tblGiftAidPayerActiveDates", "EndDate", "DATE"
    CMSDB.TableDefs.Refresh
    
    Set tdf = CMSDB.TableDefs("tblGiftAidPayerActiveDates")
    
    On Error Resume Next
    
    Set prp = tdf.Properties("Description")
    
    If Err.number <> 0 Then
        tdf.Properties.Append tdf.CreateProperty("Description", _
            dbText, "Periods during which each Gift Aid payer is active")
    Else
        prp.value = "Periods during which each Gift Aid payer is active"
    End If
    
    CMSDB.TableDefs.Refresh
    
    On Error GoTo ErrorTrap
                     
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9923.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub DB_Upgrade_5_9923_00_To_5_9924_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                          
    DeleteTable "tblMonthlyBalance"
                     
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9924.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_9924_00_To_5_9925_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                     
    CreateField ErrorCode, "tblTransactionTypes", "Amount", "DOUBLE"
                     
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9925.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_9925_00_To_5_9926_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('BankInterestTransactionCode', " & _
                          " 'I', " & _
                          " 'Initial value = I')"
                   
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9926.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_9926_00_To_5_9927_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('TransactionCodesForCongContribs', " & _
                          " 'W,C,G,K', " & _
                          " 'Initial value = W,C,G,K')"
                   
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9927.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_5_9927_00_To_5_9928_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    '
    'Export accounts
    '
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmExportDB', " & _
                          " 'chkExportItem(9)', " & _
                          "1)"
                          
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmExportDB', " & _
                          " 'chkExportItem(9)', " & _
                          "5)"
                
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmExportDB', " & _
                          " 'chkExportItem(9)', " & _
                          "9)"
                
    CMSDB.Execute "INSERT INTO tblExportDetails " & _
                  "(ExportDataType, " & _
                  " OrderingForSQL, " & _
                  " IncludeForExport, " & _
                  " Description) " & _
                  " VALUES (10, " & _
                          " 1000, " & _
                          "FALSE, " & _
                          "'Accounts')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('ExportItem9', " & _
                          " False, " & _
                          " 'Initial value = FALSE (Accounts)')"

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('ImportItem9', " & _
                          " False, " & _
                          " 'Initial value = FALSE (Accounts)')"
                          
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9928.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_9928_00_To_5_9929_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                     
    CreateField ErrorCode, "tblTransactionTypes", "OnReceipt", "YESNO"
                     
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9929.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub DB_Upgrade_5_9929_00_To_5_9930_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                     
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('SpecialAssemblyOnCalendar', " & _
                          " True, " & _
                          " 6, " & _
                          " 'Initial value = True; 6 months.')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('CircuitAssemblyOnCalendar', " & _
                          " True, " & _
                          " 6, " & _
                          " 'Initial value = True; 6 months.')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('DistrictConventionOnCalendar', " & _
                          " True, " & _
                          " 6, " & _
                          " 'Initial value = True; 6 months.')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('COVisitOnCalendar', " & _
                          " True, " & _
                          " 6, " & _
                          " 'Initial value = True; 6 months.')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('MemorialOnCalendar', " & _
                          " True, " & _
                          " 6, " & _
                          " 'Initial value = True; 6 months.')"
                     
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9930.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_9930_00_To_5_9931_00()

On Error GoTo ErrorTrap
Dim ErrorCode As Integer

                  
    CMSDB.Execute "ALTER TABLE tblConstants " & _
                  "ALTER COLUMN AlphaVal MEMO; "
                  
    CMSDB.TableDefs("tblConstants").Fields("AlphaVal").AllowZeroLength = True
                  
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('ServMtgAnnouncementsTemplate', " & _
                          " '', " & _
                          " 'Initial value = zero-length string')"
                          

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9931.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub DB_Upgrade_5_9931_00_To_5_9932_00()

On Error GoTo ErrorTrap
Dim ErrorCode As Integer

                  
     CMSDB.Execute "ALTER TABLE tblServiceMtgs " & _
                  "ALTER COLUMN ItemName MEMO; "
                 
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9932.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


'Private Sub DB_Upgrade_5_9929_00_To_5_9930_00()
'On Error GoTo ErrorTrap
'Dim ErrorCode As Integer
'
'    'allow Announcements field to hold binary data (for RTF in this case.
'            'could also be used to store pictures etc...)
'    'In Access, field will be type 'OLE OBJECT'
'
'    CMSDB.Execute "ALTER TABLE tblServiceMtgs " & _
'                  "ALTER COLUMN Announcements LONGBINARY; "
'
'    '
'    'Update DB version
'    '
'    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9930.0"
'
'    Exit Sub
'ErrorTrap:
'    EndProgram
'
'End Sub

Private Sub DB_Upgrade_5_9932_00_To_5_9933_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
Dim tdf As TableDef, prp As DAO.Property
                          
    DeleteTable "tblTerritoryMaps"

    CreateTable ErrorCode, "tblTerritoryMaps", "MapNo", "LONG", , , False
    CreateField ErrorCode, "tblTerritoryMaps", "MapName", "TEXT", "255"
    CreateField ErrorCode, "tblTerritoryMaps", "MapImage", "LONGBINARY"
    
    CreateIndex ErrorCode, "tblTerritoryMaps", "MapNo", "IX1", True, False
        
    CMSDB.TableDefs.Refresh
    
    Set tdf = CMSDB.TableDefs("tblTerritoryMaps")
    
    On Error Resume Next
    
    Set prp = tdf.Properties("Description")
    
    If Err.number <> 0 Then
        tdf.Properties.Append tdf.CreateProperty("Description", _
            dbText, "Holds Territory Maps")
    Else
        prp.value = "Holds Territory Maps"
    End If
    
    CMSDB.TableDefs.Refresh
    
    On Error GoTo ErrorTrap
'--------------------------------------

    DeleteTable "tblTerritoryMapDates"

    CreateTable ErrorCode, "tblTerritoryMapDates", "MapNo", "LONG", , , True
    CreateField ErrorCode, "tblTerritoryMapDates", "StartDate", "DATE"
    CreateField ErrorCode, "tblTerritoryMapDates", "EndDate", "DATE"
    
    CMSDB.TableDefs.Refresh
    
    Set tdf = CMSDB.TableDefs("tblTerritoryMapDates")
    
    On Error Resume Next
    
    Set prp = tdf.Properties("Description")
    
    If Err.number <> 0 Then
        tdf.Properties.Append tdf.CreateProperty("Description", _
            dbText, "Holds Dates Territory Maps are worked")
    Else
        prp.value = "Holds Dates Territory Maps are worked"
    End If
    
    CMSDB.TableDefs.Refresh
    
    On Error GoTo ErrorTrap

'--------------------------------------

    DeleteTable "tblTerritoryDNCs"

    CreateTable ErrorCode, "tblTerritoryDNCs", "MapNo", "LONG", , , True
    CreateField ErrorCode, "tblTerritoryDNCs", "Street", "TEXT", "255"
    CreateField ErrorCode, "tblTerritoryDNCs", "HouseNo", "TEXT", "255"
    CreateField ErrorCode, "tblTerritoryDNCs", "DateAdded", "DATE"
    
    CMSDB.TableDefs.Refresh
    
    Set tdf = CMSDB.TableDefs("tblTerritoryDNCs")
    
    On Error Resume Next
    
    Set prp = tdf.Properties("Description")
    
    If Err.number <> 0 Then
        tdf.Properties.Append tdf.CreateProperty("Description", _
            dbText, "Holds Do Not Calls for Territory Maps")
    Else
        prp.value = "Holds Do Not Calls for Territory Maps"
    End If
    
    CMSDB.TableDefs.Refresh
    
    On Error GoTo ErrorTrap
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9933.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Private Sub DB_Upgrade_5_9933_00_To_5_9934_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdFieldMinistry', " & _
                          "10)"
                          
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmFieldMinistryMenu', " & _
                          " 'cmdTerritory', " & _
                          "1)"
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmFieldMinistryMenu', " & _
                          " 'cmdTerritory', " & _
                          "5)"
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmFieldMinistryMenu', " & _
                          " 'cmdTerritory', " & _
                          "10)"
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmFieldMinistryMenu', " & _
                          " 'cmdFieldServiceReports', " & _
                          "1)"
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmFieldMinistryMenu', " & _
                          " 'cmdFieldServiceReports', " & _
                          "5)"
    
    CMSDB.Execute "INSERT INTO tblAccessLevelDescriptions " & _
                  "(OrderingKey, AccessLevel, AccessDesc) " & _
                  "VALUES " & _
                  "(90, 10, 'Territory')"
    
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmExportDB', " & _
                          " 'chkExportItem(10)', " & _
                          "1)"

    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmExportDB', " & _
                          " 'chkExportItem(10)', " & _
                          "5)"

    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmExportDB', " & _
                          " 'chkExportItem(10)', " & _
                          "10)"

    CMSDB.Execute "INSERT INTO tblExportDetails " & _
                  "(ExportDataType, " & _
                  " OrderingForSQL, " & _
                  " IncludeForExport, " & _
                  " Description) " & _
                  " VALUES (11, " & _
                          " 1100, " & _
                          "FALSE, " & _
                          "'Territory')"

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('ExportItem10', " & _
                          " False, " & _
                          " 'Initial value = FALSE (Territory)')"

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('ImportItem10', " & _
                          " False, " & _
                          " 'Initial value = FALSE (Territory)')"

    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9934.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Private Sub DB_Upgrade_5_9934_00_To_5_9935_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    CMSDB.TableDefs("tblTerritoryMaps").Fields("MapImage").Required = False
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9935.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub DB_Upgrade_5_9935_00_To_5_9936_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    DropField ErrorCode, "tblTerritoryDNCs", "Street"
    DropField ErrorCode, "tblTerritoryDNCs", "MapNo"
    CreateField ErrorCode, "tblTerritoryDNCs", "StreetSeqNum", "LONG"
    CMSDB.TableDefs.Refresh
            
    DeleteTable "tblTerritoryStreets"

    CreateTable ErrorCode, "tblTerritoryStreets", "StreetName", "TEXT", "255", , True, "StreetSeqNum"
    CreateField ErrorCode, "tblTerritoryStreets", "MapNo", "LONG"
                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9936.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Private Sub DB_Upgrade_5_9936_00_To_5_9937_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset, lMyCong As Long
    
    DeleteTable "tblGeneralNotesForPerson"

    CreateTable ErrorCode, "tblGeneralNotesForPerson", "PersonID", "LONG", , , False
    CreateField ErrorCode, "tblGeneralNotesForPerson", "Notes", "MEMO"
    
    CreateIndex ErrorCode, "tblGeneralNotesForPerson", "PersonID", "IX1", True, False

    CMSDB.TableDefs.Refresh

    CMSDB.TableDefs("tblGeneralNotesForPerson").Fields("Notes").AllowZeroLength = True
    
    
    'now give all pubs a km.... at last....
    Set rs = CMSDB.OpenRecordset("SELECT ID FROM tblNameAddress " & _
                                " WHERE Active = TRUE ", dbOpenForwardOnly)
                                
    With rs
    
    lMyCong = GlobalParms.GetValue("DefaultCong", "NumVal")
    
    Do Until .BOF Or .EOF
        If CongregationMember.IsPublisher(!ID, Now) Then
            CongregationMember.AddPersonToRole lMyCong, 4, 4, 90, !ID
        End If
        .MoveNext
    Loop
    
    End With
    
    rs.Close
    Set rs = Nothing

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9937.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_9937_00_To_5_9938_00()

On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                  
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('SendTextEmailUsingWinsock', " & _
                          " FALSE, " & _
                          " 'Initial value = false')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('WinsockSMTPServer', " & _
                          " '', " & _
                          " 'Initial value = zerolengthstring')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('WinsockFromAddress', " & _
                          " '', " & _
                          " 'Initial value = zerolengthstring')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('WinsockFromName', " & _
                          " '', " & _
                          " 'Initial value = zerolengthstring')"
                          

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9938.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_9938_00_To_5_9939_00()

On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                  
    CreateField ErrorCode, "tblTransactionTypes", "Ref", "LONG"
    
    CMSDB.TableDefs.Refresh
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9939.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub DB_Upgrade_5_9939_00_To_5_9940_00()

On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    DestroyGlobalObjects
                  
    DropField ErrorCode, "tblNameAddress", "OfficialFirstName"
    CreateField ErrorCode, "tblNameAddress", "OfficialFirstName", "TEXT", "100"
    
    SetUpGlobalObjects
    CMSDB.TableDefs.Refresh
    
    CMSDB.Execute "UPDATE tblNameAddress SET OfficialFirstName = FirstName"
    
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9940.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub DB_Upgrade_5_9940_00_To_5_9941_00()

On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  "AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('PubRecCardVersion', " & _
                          " 1, " & _
                          "'5/02', " & _
                          " 'Initial value = 1 (ie S-21 5/02)')"

    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9941.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Private Sub DB_Upgrade_5_9941_00_To_5_9942_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add new publisher record card print parameters to tblConstants
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPublisherNameXPos_4_05', " & _
                          " 0.6, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPublisherNameYPos_4_05', " & _
                          " 1.2, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAddressXPos_4_05', " & _
                          " 1.55, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAddressYPos_4_05', " & _
                          " 1.65, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPublisherNameMaxWidth_4_05', " & _
                          " 13.1, " & _
                          " 'In cm. ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAddressMaxWidth_4_05', " & _
                          " 12.15, " & _
                          " 'In cm.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTelNoXPos_4_05', " & _
                          " 1.8, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTelNoYPos_4_05', " & _
                          " 2.1, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTelNoMaxWidth_4_05', " & _
                          " 2.85, " & _
                          " 'In cm. ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardElderXPos_4_05', " & _
                          " 10.36, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardElderYPos_4_05', " & _
                          " 2.54, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardServantXPos_4_05', " & _
                          " 11.45, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardServantYPos_4_05', " & _
                          " 2.54, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardRegPioXPos_4_05', " & _
                          " 10.36, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardRegPioYPos_4_05', " & _
                          " 2.8, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardBaptDateXPos_4_05', " & _
                          " 2.7, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardBaptDateYPos_4_05', " & _
                          " 2.75, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardBaptDateMaxWidth_4_05', " & _
                          " 3.7, " & _
                          " 'In cm. ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPioNoXPos_4_05', " & _
                          " 0.1, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPioNoYPos_4_05', " & _
                          " 0.85, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPioNoMaxWidth_4_05', " & _
                          " 2.7, " & _
                          " 'In cm. ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAnointedXPos_4_05', " & _
                          " 9, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAnointedYPos_4_05', " & _
                          " 2.75, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAnointedMaxWidth_4_05', " & _
                          " 1, " & _
                          " 'In cm. ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTopMargin_4_05', " & _
                          " 0.65, " & _
                          " 'In cm..')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardBottomMargin_4_05', " & _
                          " 0.8, " & _
                          " 'In cm..')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardLeftMargin_4_05', " & _
                          " 0.4, " & _
                          " 'In cm..')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardRightMargin_4_05', " & _
                          " 0.4, " & _
                          " 'In cm..')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardFontSize_4_05', " & _
                          " 8, " & _
                          " ' ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES('PubCardFontName_4_05', " & _
                          " 'Arial', " & _
                          " '.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPaperHeight_4_05', " & _
                          " 10.5, " & _
                          " 'In cm.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardPaperWidth_4_05', " & _
                          " 15.2, " & _
                          " 'In cm.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTweakX_4_05', " & _
                          " 0, " & _
                          " ' ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTweakY_4_05', " & _
                          " 0, " & _
                          " ' ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardDOBXPos_4_05', " & _
                          " 11.1, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardDOBYPos_4_05', " & _
                          " 2.1, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardMaleXPos_4_05', " & _
                          " 13.61, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardMaleYPos_4_05', " & _
                          " 0.72, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardFemaleXPos_4_05', " & _
                          " 13.61, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardFemaleYPos_4_05', " & _
                          " 1.03, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardMobileNoXPos_4_05', " & _
                          " 6.65, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardMobileNoYPos_4_05', " & _
                          " 2.15, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardMobileNoMaxWidth_4_05', " & _
                          " 2.8, " & _
                          " 'In cm. ')"
    
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardServiceYearXPos_4_05', " & _
                          " 0.3, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardServiceYearYPos_4_05', " & _
                          " 3.6, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardBooksXPos_4_05', " & _
                          " 1.55, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardBrochuresXPos_4_05', " & _
                          " 2.55, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardHoursXPos_4_05', " & _
                          " 3.95, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardSubscripXPos_4_05', " & _
                          " 5.55, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardMagsXPos_4_05', " & _
                          " 5.45, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardRVsXPos_4_05', " & _
                          " 6.55, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardStudiesXPos_4_05', " & _
                          " 7.95, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardRemarksXPos_4_05', " & _
                          " 8.91, " & _
                          " 'In cm. Does not include margins.')"
    
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardSeptYPos_4_05', " & _
                          " 3.9, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardOctYPos_4_05', " & _
                          " 4.25, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardNovYPos_4_05', " & _
                          " 4.61, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardDecYPos_4_05', " & _
                          " 4.97, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardJanYPos_4_05', " & _
                          " 5.34, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardFebYPos_4_05', " & _
                          " 5.7, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardMarYPos_4_05', " & _
                          " 6.02, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAprYPos_4_05', " & _
                          " 6.4, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardMayYPos_4_05', " & _
                          " 6.77, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardJunYPos_4_05', " & _
                          " 7.1, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardJulYPos_4_05', " & _
                          " 7.45, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardAugYPos_4_05', " & _
                          " 7.8, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTotalYPos_4_05', " & _
                          " 8.16, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardRemarksMaxWidth_4_05', " & _
                          " 5.0, " & _
                          " 'In cm. ')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9942.0"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_5_9942_00_To_5_9943_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " Comment) " & _
                  " VALUES ('NextScheduleSlipPrintStartDate', " & _
                          " 'Initial value = Null')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9943.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_9943_00_To_5_9944_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('COServiceMtgItemLength', " & _
                            30 & ", " & _
                          " 'Initial value = 30')"
        
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('TextForBroToIntroCO', " & _
                            "'Introduce Song and invite CO to platform.'" & ", " & _
                          " 'Initial value = Introduce Song and invite CO to platform.')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9944.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub DB_Upgrade_5_9944_00_To_5_9945_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('ServMtgStartMinsAfterTMS', " & _
                            55 & ", " & _
                          " 'Initial value = 55')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('ServMtgCOVisitStartMinsAfterTMS', " & _
                            40 & ", " & _
                          " 'Initial value = 40')"
            
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9945.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub DB_Upgrade_5_9945_00_To_5_9946_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    DeleteTable "tblCustomRotaDetails"

    CreateTable ErrorCode, "tblCustomRotaDetails", "RotaName", "TEXT", "100", , True
    CreateField ErrorCode, "tblCustomRotaDetails", "QueryString", "MEMO"
    CreateField ErrorCode, "tblCustomRotaDetails", "FrequencyID", "LONG"
    CreateField ErrorCode, "tblCustomRotaDetails", "DateFormatID", "LONG"
    CreateField ErrorCode, "tblCustomRotaDetails", "DaysToInclude", "TEXT"
    CreateField ErrorCode, "tblCustomRotaDetails", "Column1Name", "TEXT"
    CreateField ErrorCode, "tblCustomRotaDetails", "Column2Name", "TEXT"
    CreateField ErrorCode, "tblCustomRotaDetails", "EventsToSkip", "TEXT"
    CreateField ErrorCode, "tblCustomRotaDetails", "PrevRotaLastDate", "DATE"
    CreateField ErrorCode, "tblCustomRotaDetails", "PrevRotaLastValue", "TEXT", "255"
   
    CreateIndex ErrorCode, "tblCustomRotaDetails", "SeqNum", "IX1", True, False

    CMSDB.TableDefs.Refresh

    CMSDB.TableDefs("tblCustomRotaDetails").Fields("QueryString").AllowZeroLength = True
    CMSDB.TableDefs("tblCustomRotaDetails").Fields("PrevRotaLastValue").AllowZeroLength = True

'-----------------------------------

    DeleteTable "tblCustomRotaPrint"

    CreateTable ErrorCode, "tblCustomRotaPrint", "DisplayDate", "TEXT", "100", , True
    CreateField ErrorCode, "tblCustomRotaPrint", "DisplayItem", "TEXT", "255"
            
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9946.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_9946_00_To_5_9947_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('MinsAfterMtgStartForCOvisit', " & _
                            30 & ", " & _
                          " 'Initial value = 30')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9947.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_9947_00_To_5_9948_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('MultiUserMode', " & _
                            True & ", " & _
                          " 'Initial value = TRUE')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9948.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub DB_Upgrade_5_9948_00_To_5_9949_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add new role
    '
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description, " & _
                  " AllowSuspend) " & _
                  " VALUES (5, " & _
                          " 8, " & _
                          " 95, " & _
                          " 'Long-term Inactive Publisher', " & _
                          " FALSE)"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9949.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_9949_00_To_5_9950_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    DestroyGlobalObjects
                      
    CreateField ErrorCode, "tblPrintAddresses", "MobilePhone2", "TEXT"
    CreateField ErrorCode, "tblNameAddress", "MobilePhone2", "TEXT"
    
    CMSDB.TableDefs.Refresh
    
    CMSDB.TableDefs("tblNameAddress").Fields("MobilePhone2").AllowZeroLength = True
    
    CMSDB.Execute "UPDATE tblNameAddress SET MobilePhone2 = ''"
    
    SetUpGlobalObjects
    
    CMSDB.TableDefs.Refresh
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9950.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_5_9950_00_To_5_9951_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdFieldMinistry', " & _
                          "11)"
                          
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmMainMenu', " & _
                          " 'cmdMeetingAttendance', " & _
                          "11)"
                          
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmFieldMinistryMenu', " & _
                          " 'cmdFieldServiceReports', " & _
                          "11)"
                          
    CMSDB.Execute "INSERT INTO tblAccessLevelDescriptions " & _
                  "(OrderingKey, AccessLevel, AccessDesc) " & _
                  "VALUES " & _
                  "(100, 11, 'Field Ministry and Meeting Attendance')"
                          
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9951.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub DB_Upgrade_5_9951_00_To_5_9952_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    DestroyGlobalObjects
    
    CMSDB.TableDefs("tblNameAddress").Fields("MobilePhone2").Required = False
    
    SetUpGlobalObjects
    
    CMSDB.TableDefs.Refresh
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9952.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub DB_Upgrade_5_9952_00_To_5_9953_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmPublicMeetingMenu', " & _
                          " 'cmdChairmansNotes', " & _
                          "8)"
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmPublicMeetingMenu', " & _
                          " 'cmdTalkOutlines', " & _
                          "8)"
                          
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmSetUpMenu', " & _
                          " 'cmdExportDB', " & _
                          "8)"
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmSetUpMenu', " & _
                          " 'cmdExportDB', " & _
                          "9)"
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmSetUpMenu', " & _
                          " 'cmdExportDB', " & _
                          "10)"
    CMSDB.Execute "INSERT INTO tblObjectSecurity " & _
                  "(FormNameProperty, " & _
                  " ControlNameProperty, " & _
                  " SecurityLevel) " & _
                  " VALUES ('frmSetUpMenu', " & _
                          " 'cmdExportDB', " & _
                          "11)"
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9953.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_9953_00_To_5_9954_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblAccInOutTypes (Description,InOutID,IsAuto) VALUES ('Temporary Income - Auto Monthly outgoing (Non-WBTS)',1,False)"
    CMSDB.Execute "UPDATE tblAccInOutTypes SET Description = 'Temporary Income - Auto Monthly outgoing (WBTS)' WHERE InOutTypeID = 5"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9954.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub DB_Upgrade_5_9954_00_To_5_9955_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.TableDefs("tblTransactionTypes").Fields("Ref").Required = False
    
    CMSDB.TableDefs.Refresh

        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9955.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub DB_Upgrade_5_9955_00_To_5_9956_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    
    CMSDB.TableDefs("tblMinReports").Fields("Remarks").Required = False
    CMSDB.TableDefs("tblMinReports").Fields("Remarks").AllowZeroLength = True
    
    CMSDB.TableDefs.Refresh
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9956.0"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub DB_Upgrade_5_9956_00_To_5_9957_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    DestroyGlobalObjects
                      
    CreateField ErrorCode, "tblMeetingTypes", "AltOrder", "LONG"
    
    CMSDB.TableDefs.Refresh
    
    CMSDB.Execute "UPDATE tblMeetingTypes SET AltOrder = 2 WHERE MeetingTypeID = 0"
    CMSDB.Execute "UPDATE tblMeetingTypes SET AltOrder = 3 WHERE MeetingTypeID = 1"
    CMSDB.Execute "UPDATE tblMeetingTypes SET AltOrder = 0 WHERE MeetingTypeID = 2"
    CMSDB.Execute "UPDATE tblMeetingTypes SET AltOrder = 1 WHERE MeetingTypeID = 3"
    CMSDB.Execute "UPDATE tblMeetingTypes SET AltOrder = 4 WHERE MeetingTypeID = 4"
    
    SetUpGlobalObjects
    
    CMSDB.TableDefs.Refresh
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9957.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_5_9957_00_To_5_9958_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    DeleteTable "tblEventAutoAlarms"

    CreateTable ErrorCode, "tblEventAutoAlarms", "EventTypeID", "LONG", , , True, "AutoAlarmID"
    CreateField ErrorCode, "tblEventAutoAlarms", "AlarmDays", "LONG"
    CreateField ErrorCode, "tblEventAutoAlarms", "AlarmMemo", "MEMO"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9958.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_5_9958_00_To_5_9959_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "ALTER TABLE tblEvents " & _
                  "ALTER COLUMN LinkEventID LONG; "
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9959.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_5_9959_00_To_5_9960_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('MaxDBSizeMB', " & _
                          " 1000, " & _
                          " 'Initial value = 1000')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES ('WarnAboutDBSizeThreshold', " & _
                          " 0.8, " & _
                          " 'Initial value = 0.8')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9960.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_5_9960_00_To_5_9961_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('TMSSlipStudentNote2XPos', " & _
                          " 2.4, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('TMSSlipStudentNote2YPos', " & _
                          " 4.35, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('TMSSlipStudentNote2XPos_Sub', " & _
                          " 6.4, " & _
                          " 'In cm. Does not include margins.')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('TMSSlipStudentNote2YPos_Sub', " & _
                          " 5.9, " & _
                          " 'In cm. Does not include margins.')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSPrintOtherSpeakersForS1B', " & _
                            True & ", " & _
                          " 'Initial value = TRUE')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9961.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_5_9961_00_To_5_9962_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('DocumentLocation', " & _
                          " '', " & _
                          " 'Initial value = blank')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9962.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_5_9962_00_To_5_9963_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('DocsDeleteDays', " & _
                          " '750', " & _
                          " 'Days before docs deleted. 0 for no delete. Initial value = 750')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9963.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_5_9963_00_To_5_9964_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('ServiceMtgScheduleAlert', " & _
                          " True, " & _
                          " 14, " & _
                          " 'Initial value = True; 14 days.')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9964.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_5_9964_00_To_5_9965_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('DNCsToPrint', " & _
                          " '', " & _
                          " 'Initial value = blank')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9965.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_5_9965_00_To_5_9966_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('MaxSMSLength', " & _
                          " 160, " & _
                          " 'Initial value = 160')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9966.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_5_9966_00_To_5_9967_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('AlertForMissingBankInterest', " & _
                          " TRUE, " & _
                          " 'Initial value = TRUE')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9967.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_5_9967_00_To_5_9968_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    DeleteTable "tblPubCardTypes"
    CreateTable ErrorCode, "tblPubCardTypes", "CardTypeID", "LONG", , , False
    CreateField ErrorCode, "tblPubCardTypes", "CardTypeDesc", "TEXT"
    
    CMSDB.Execute "INSERT INTO tblPubCardTypes " & _
                  "(CardTypeID, " & _
                  " CardTypeDesc) " & _
                  " VALUES (0, 'S-21 5/02')"
    CMSDB.Execute "INSERT INTO tblPubCardTypes " & _
                  "(CardTypeID, " & _
                  " CardTypeDesc) " & _
                  " VALUES (1, 'S-21 4/05')"
    
    DeleteTable "tblPubCardTypeForPerson"
    CreateTable ErrorCode, "tblPubCardTypeForPerson", "PersonID", "LONG", , , True
    CreateField ErrorCode, "tblPubCardTypeForPerson", "CardTypeID", "LONG"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9968.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub DB_Upgrade_5_9968_00_To_5_9969_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "ALTER TABLE tblServiceMtgSchedulePrint " & _
                  "ALTER COLUMN MeetingDate TEXT(255)" & ";"
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9969.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub DB_Upgrade_5_9969_00_To_5_9973_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9973.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_5_9973_00_To_5_9974_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('YearRangeCheckPublicSpeaker', " & _
                          " 3, " & _
                          " 'Initial value = 3')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9974.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


