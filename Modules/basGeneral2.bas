Attribute VB_Name = "basGeneral2"
Option Explicit

Public Function ServiceMtgConflict(PersonID As Long, ItemDate As Date, Optional Prompt As Boolean = True) As Boolean

On Error GoTo ErrorTrap
Dim i As Long, colItemDtl As New Collection, str As String
    
    ServiceMtgConflict = False
            
    Set colItemDtl = CongregationMember.ServiceMtgItemThisWeek(PersonID, ItemDate)
       
    If colItemDtl.Count > 0 Then
        If Prompt Then
            Select Case CLng(colItemDtl.Item(1))
            Case 0
                str = GetServiceMtgItemType(0) & " (" & colItemDtl.Item(3) & " mins)"
            Case 1
                str = GetServiceMtgItemType(1) & " (" & colItemDtl.Item(3) & " mins)"
            Case 2
                str = GetServiceMtgItemType(2)
            Case Else
                str = ""
            End Select
            
            If MsgBox(CongregationMember.NameWithMiddleInitial(PersonID) & _
                      " has the " & str & _
                      " on the Service Meeting this week. Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                ServiceMtgConflict = True
            End If
        Else
            ServiceMtgConflict = True
        End If
            
    End If
    
    Set colItemDtl = Nothing
    
    Exit Function
ErrorTrap:
    EndProgram

End Function
Public Function TMSAssignedTooSoon(PersonID As Long, _
                                    CurrentDate As Date, _
                                    TalkNo As String) As Date

On Error GoTo ErrorTrap
    
Dim NoWeeks As Long, TempDateA As Date, TempDateB As Date, l1 As Long, l2 As Long

    TMSAssignedTooSoon = 0
    
    If Not GlobalParms.GetValue("TMSWarnAboutCloseAssignments", "TrueFalse") Then
        Exit Function
    End If
    
    Select Case True
    Case CongregationMember.IsAppointedMan(PersonID) Or TalkNo = "B"
        NoWeeks = GlobalParms.GetValue("TMSCloseAssignments_For_E_And_MS", "NumVal")
    Case CongregationMember.IsMale(PersonID)
        NoWeeks = GlobalParms.GetValue("TMSCloseAssignments_For_Bros_1and3", "NumVal")
    Case Else
        NoWeeks = GlobalParms.GetValue("TMSCloseAssignments_For_Sisters_2and3", "NumVal")
    End Select
    
    TempDateA = #1/1/1900#
    TempDateB = #1/1/1900#
    
    If CongregationMember.TMSPreviousAssignmentForPerson(PersonID, CurrentDate, False, False, , , True, , True) Then
        If CongregationMember.GetTMSAssignmentDate >= CurrentDate - 7 * NoWeeks And _
            CongregationMember.GetTMSAssignmentDate <= CurrentDate Then
            
            TempDateA = CongregationMember.GetTMSAssignmentDate
        End If
    End If
        
    If CongregationMember.TMSNextAssignmentForPerson(PersonID, CurrentDate, False, False, , , True, , , True) Then
        If CongregationMember.GetTMSAssignmentDate <= CurrentDate + 7 * NoWeeks And _
            CongregationMember.GetTMSAssignmentDate >= CurrentDate Then
            
            TempDateB = CongregationMember.GetTMSAssignmentDate
        End If
    End If
    
    l1 = Abs(DateDiff("d", CurrentDate, TempDateA))
    l2 = Abs(DateDiff("d", CurrentDate, TempDateB))
    
    If TempDateA = #1/1/1900# And TempDateB = #1/1/1900# Then
        TMSAssignedTooSoon = 0
    Else
        If l1 > l2 Then
            TMSAssignedTooSoon = TempDateB
        Else
            TMSAssignedTooSoon = TempDateA
        End If
    End If
        
    Exit Function
ErrorTrap:
    EndProgram

End Function





Public Function IsCircuitOrDistrictAssemblyWeek(TheDate As Date) As Boolean
                                 
Dim TheCong As Long, rstQuery As Recordset
                                 
On Error GoTo ErrorTrap

    Dim StartDate As Date, EndDate As Date
    
    StartDate = GetDateOfGivenDay(TheDate, vbMonday, False) 'monday
    EndDate = StartDate + 6 'Sunday
                                            
    '
    'Now check Calendar for Circuit/District assembly this week...
    '
    Set rstQuery = CMSDB.OpenRecordset("SELECT EventStartDate " & _
                                    "FROM tblEvents " & _
                                    "WHERE EventStartDate BETWEEN " & GetDateStringForSQLWhere(CStr(StartDate)) & _
                                        " AND " & GetDateStringForSQLWhere(CStr(EndDate)) & _
                                     " AND EventID IN (1, 3)  " & _
                                     " AND CongNo = " & GlobalDefaultCong, _
                                    , dbOpenSnapshot)
                                       
    If rstQuery.BOF Then
        IsCircuitOrDistrictAssemblyWeek = False
    Else
        IsCircuitOrDistrictAssemblyWeek = True
    End If
                                       

    Exit Function
ErrorTrap:
    EndProgram


End Function

Public Function IsHostVisitWeek(TheDate As Date) As Boolean
                                 
Dim TheCong As Long, rstQuery As Recordset
                                 
On Error GoTo ErrorTrap
                                            
    Dim StartDate As Date, EndDate As Date
    
    StartDate = GetDateOfGivenDay(TheDate, vbMonday, False) 'monday
    EndDate = StartDate + 6 'Sunday
                                            
                                            
    '
    'Now check Calendar for host week...
    '
    Set rstQuery = CMSDB.OpenRecordset("SELECT EventStartDate " & _
                                    "FROM tblEvents " & _
                                    "WHERE EventStartDate BETWEEN " & GetDateStringForSQLWhere(CStr(StartDate)) & _
                                        " AND " & GetDateStringForSQLWhere(CStr(EndDate)) & _
                                     " AND EventID IN (5)  " & _
                                     " AND CongNo = " & GlobalDefaultCong, _
                                    , dbOpenSnapshot)

                                       
    If rstQuery.BOF Then
        IsHostVisitWeek = False
    Else
        IsHostVisitWeek = True
    End If
                                       

    Exit Function
ErrorTrap:
    EndProgram


End Function


Public Function IsCircuitOrDistrictAssemblyDay(AssignmentDate As Date) As Boolean
                                 
Dim TheCong As Long, rstQuery As Recordset
                                 
On Error GoTo ErrorTrap
                                            
    '
    'Now check Calendar for Circuit/District assembly this week...
    '
    Set rstQuery = CMSDB.OpenRecordset("SELECT EventStartDate " & _
                                    "FROM tblEvents " & _
                                    "WHERE EventStartDate <= #" & Format(AssignmentDate, "mm/dd/yyyy") & "# " & _
                                    " AND EventEndDate >= #" & Format(AssignmentDate, "mm/dd/yyyy") & "# " & _
                                     " AND EventID IN (1, 3) " & _
                                     " AND CongNo = " & GlobalDefaultCong, _
                                    , dbOpenSnapshot)
                                       
    If rstQuery.BOF Then
        IsCircuitOrDistrictAssemblyDay = False
    Else
        IsCircuitOrDistrictAssemblyDay = True
    End If
                                       

    Exit Function
ErrorTrap:
    EndProgram


End Function
Public Function GetNextMonthNo(TheMonth As Long) As Long
                                 
    If TheMonth < 12 Then
        GetNextMonthNo = TheMonth + 1
    Else
        GetNextMonthNo = 1
    End If

    Exit Function
ErrorTrap:
    EndProgram


End Function

Public Function IsMemorialWeek(TheDate As Date) As Boolean
                                 
Dim TheCong As Long, rstQuery As Recordset
                                 
On Error GoTo ErrorTrap
                                            
    Dim StartDate As Date, EndDate As Date
    
    StartDate = GetDateOfGivenDay(TheDate, vbMonday, False) 'monday
    EndDate = StartDate + 6 'Sunday
                                                
    
    '
    'Now check Calendar for memorial this week...
    '
    Set rstQuery = CMSDB.OpenRecordset("SELECT EventStartDate " & _
                                    "FROM tblEvents " & _
                                    "WHERE EventStartDate BETWEEN " & GetDateStringForSQLWhere(CStr(StartDate)) & _
                                        " AND " & GetDateStringForSQLWhere(CStr(EndDate)) & _
                                     " AND EventID IN (6)  " & _
                                     " AND CongNo = " & GlobalDefaultCong, _
                                    , dbOpenSnapshot)
                                       
    If rstQuery.BOF Then
        IsMemorialWeek = False
    Else
        IsMemorialWeek = True
    End If
                                       

    Exit Function
ErrorTrap:
    EndProgram


End Function
Public Function IsCOVisitWeek(TheDate As Date) As Boolean
                                 
Dim TheCong As Long, rstQuery As Recordset
                                 
On Error GoTo ErrorTrap

    Dim StartDate As Date, EndDate As Date
    
    StartDate = GetDateOfGivenDay(TheDate, vbMonday, False) 'monday
    EndDate = StartDate + 6 'Sunday
                                            

    '
    'Now check Calendar for CO Visit this week...
    '
    Set rstQuery = CMSDB.OpenRecordset("SELECT EventStartDate " & _
                                    "FROM tblEvents " & _
                                    "WHERE EventStartDate BETWEEN " & GetDateStringForSQLWhere(CStr(StartDate)) & _
                                        " AND " & GetDateStringForSQLWhere(CStr(EndDate)) & _
                                     " AND EventID IN (4)  " & _
                                     " AND CongNo = " & GlobalDefaultCong, _
                                    , dbOpenSnapshot)
                                       
    If rstQuery.BOF Then
        IsCOVisitWeek = False
    Else
        IsCOVisitWeek = True
    End If
                                       

    Exit Function
ErrorTrap:
    EndProgram


End Function
Public Function IsCOVisitDay(TheDate As Date) As Boolean
                                 
Dim TheCong As Long, rstQuery As Recordset
                                 
On Error GoTo ErrorTrap

    '
    'Now check Calendar for CO Visit today...
    '
    Set rstQuery = CMSDB.OpenRecordset("SELECT EventStartDate " & _
                                    "FROM tblEvents " & _
                                    "WHERE EventStartDate <= #" & Format(TheDate, "mm/dd/yyyy") & "# " & _
                                    " AND EventEndDate >= #" & Format(TheDate, "mm/dd/yyyy") & "# " & _
                                     " AND EventID IN (4) " & _
                                     " AND CongNo = " & GlobalDefaultCong, _
                                    , dbOpenSnapshot)
                                       
    If rstQuery.BOF Then
        IsCOVisitDay = False
    Else
        IsCOVisitDay = True
    End If
                                       
    Exit Function
ErrorTrap:
    EndProgram


End Function



Public Function IsCOVisitThisMonth(AssignmentDate As Date) As Boolean
                                 
Dim Matched As Boolean, FirstDate As Date, TempDate As Date
                                 
On Error GoTo ErrorTrap

    TempDate = GetDateOfFirstWeekDayOfMonth(AssignmentDate, Weekday(AssignmentDate))
    
    Do Until Month(TempDate) <> Month(AssignmentDate)
        If IsCOVisitWeek(TempDate) Then
            Matched = True
            Exit Do
        End If
        TempDate = DateAdd("ww", 1, TempDate)
    Loop
    
                                       
    If Matched Then
        IsCOVisitThisMonth = True
    Else
        IsCOVisitThisMonth = False
    End If
                                       

    Exit Function
ErrorTrap:
    EndProgram


End Function


Public Function IsMemorialDay(AssignmentDate As Date) As Boolean
                                 
Dim TheCong As Long, rstQuery As Recordset, sDate As String, lMidWkMtgDay As Long
                                 
On Error GoTo ErrorTrap
                                          
    '
    'check Calendar for memorial today...
    '
   
    sDate = CStr(GetDateOfGivenDay(AssignmentDate, glMidWkMtgDay, True))
    
    Set rstQuery = CMSDB.OpenRecordset("SELECT EventStartDate " & _
                                    "FROM tblEvents " & _
                                    "WHERE EventStartDate = " & GetDateStringForSQLWhere(sDate) & _
                                    "AND EventID = 6 " & _
                                     " AND CongNo = " & GlobalDefaultCong, _
                                    , dbOpenSnapshot)
                                       
    If rstQuery.BOF Then
        IsMemorialDay = False
    Else
        IsMemorialDay = True
    End If
                                       

    Exit Function
ErrorTrap:
    EndProgram


End Function

' insert a pause of a given duration (rounded to nearest integer)

Sub Pause(Seconds As Integer)
    Const SECS_INDAY = 24! * 60 * 60    ' seconds per day
    Dim start As Single
    start = Timer
    Do: Loop Until (Timer + SECS_INDAY - start) Mod SECS_INDAY >= Seconds
End Sub

Sub PauseHundredthSecs(HundredthSeconds As Integer)
' insert a pause of a given duration (rounded to nearest integer)
    Const SECS_INDAY = 24! * 60 * 60 * 100   ' 1/100 seconds per day
    Dim start As Single
    start = Timer
    Do: Loop Until 100 * (Timer + SECS_INDAY - start) Mod (100 * SECS_INDAY) >= HundredthSeconds
End Sub

' returns True if an year is a leap year

Function IsLeapYear(year As Integer) As Boolean
    ' does February 29 coincides with March 1 ?
    IsLeapYear = DateSerial(year, 2, 29) <> DateSerial(year, 3, 1)
End Function
' returns True if an year is a leap year

Function GetNextInSequence(Reset As Boolean, _
                           StartValue As Long, _
                           UpperLimit As Long, _
                           LowerLimit As Long, _
                           Step As Long) As Long
                           

Static CurrentValue As Long

    If Reset Then
        CurrentValue = StartValue
    Else
        CurrentValue = CurrentValue + Step
    End If
    
    If CurrentValue > UpperLimit Then
        CurrentValue = LowerLimit
    End If
    If CurrentValue < LowerLimit Then
        CurrentValue = UpperLimit
    End If
    
    GetNextInSequence = CurrentValue

End Function

Public Sub CompactTheDB(Optional NewDBName As String = "")
On Error GoTo ErrorTrap
Dim s2ndBackupLocn As String
Dim fso As New FileSystemObject
    
    If Dir(JustTheDirectory & "\" & TheMDBFile & "_TEMP.cms") <> "" Then
        Kill JustTheDirectory & "\" & TheMDBFile & "_TEMP.cms"
    End If

    DBEngine.CompactDatabase JustTheDirectory & "\" & TheMDBFile & ".cms", JustTheDirectory & "\" & TheMDBFile & "_TEMP.cms", , dbEncrypt, ";pwd=" & TheDBPassword

    Kill JustTheDirectory & "\" & TheMDBFile & ".cms"

    DBEngine.CompactDatabase JustTheDirectory & "\" & TheMDBFile & "_TEMP.cms", JustTheDirectory & "\" & TheMDBFile & ".cms", , dbEncrypt, ";pwd=" & TheDBPassword

    If NewDBName = "" Then
        If Dir(JustTheDirectory & "\" & TheMDBFile & "_Backup" & CStr(NextCMSDBSeqNo) & ".cms") <> "" Then
            Kill JustTheDirectory & "\" & TheMDBFile & "_Backup" & CStr(NextCMSDBSeqNo) & ".cms"
        End If
    
        DBEngine.CompactDatabase JustTheDirectory & "\" & TheMDBFile & "_TEMP.cms", JustTheDirectory & "\" & TheMDBFile & "_Backup" & CStr(NextCMSDBSeqNo) & ".cms", , , ";pwd=" & TheDBPassword
    Else
        
        If Dir(JustTheDirectory & "\" & NewDBName & ".cms") <> "" Then
            Kill JustTheDirectory & "\" & NewDBName & ".cms"
        End If
    
        DBEngine.CompactDatabase JustTheDirectory & "\" & TheMDBFile & "_TEMP.cms", JustTheDirectory & "\" & NewDBName & ".cms", , , ";pwd=" & TheDBPassword
    End If
    
    Kill JustTheDirectory & "\" & TheMDBFile & "_TEMP.cms"
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Public Sub BackUpCMSDB(Optional ShowHourglass As Boolean = True, _
                       Optional CloseAllForms As Boolean = True)
On Error GoTo ErrorTrap
Dim NumberOfCMSDBsToKeep As Integer

    If ShowHourglass Then
        Screen.MousePointer = 11
    End If
    
    NumberOfCMSDBsToKeep = GlobalParms.GetValue("NumberOfCMSDBBackupsToKeep", "NumVal")
    NextCMSDBSeqNo = GlobalParms.GetValue("CurrentCMSDBBackupSeqNo", "NumVal") + 1

    If NextCMSDBSeqNo > NumberOfCMSDBsToKeep Then
        NextCMSDBSeqNo = 0
    End If

    GlobalParms.Save "CurrentCMSDBBackupSeqNo", "NumVal", NextCMSDBSeqNo

    DestroyGlobalObjects

    On Error Resume Next
    CMSDB.Close
    On Error GoTo ErrorTrap

    CompactTheDB

    If CloseAllForms Then
        CloseAllOpenForms False
    End If

    If ShowHourglass Then
        Screen.MousePointer = 0
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Sub Do2ndBackup(s2ndBackupLocn As String)
On Error GoTo ErrorTrap
Dim fso As New FileSystemObject
        
    If Trim(s2ndBackupLocn) <> "" Then
        If fso.DriveExists(Left(s2ndBackupLocn, 2)) Then
            If fso.GetDrive(Left(s2ndBackupLocn, 2)).IsReady Then
                If fso.FolderExists(s2ndBackupLocn) Then
                    On Error Resume Next
                    fso.CopyFile JustTheDirectory & "\" & TheMDBFile & "_Backup" & CStr(NextCMSDBSeqNo) & ".cms", _
                                 s2ndBackupLocn & "\" & TheMDBFile & "_Backup" & CStr(NextCMSDBSeqNo) & ".cms", _
                                 True
                    If Err.number <> 0 Then
                        MsgBox "Backup failed. Check that " & _
                                s2ndBackupLocn & " is writeable " & _
                                "and that you have " & _
                                "necessary permissions to write to it.", vbOKOnly + vbExclamation, AppName
                    End If
                    On Error GoTo ErrorTrap
                Else
                    MsgBox "Backup folder (" & s2ndBackupLocn & ") " & _
                            "specified in General Settings " & _
                           "does not exist", vbOKOnly + vbExclamation, AppName
                    
                End If
            Else
                 MsgBox "Backup drive (" & Left(s2ndBackupLocn, 2) & ") " & _
                         "specified in General Settings " & _
                        "is not ready/available", vbOKOnly + vbExclamation, AppName
            End If
        Else
             MsgBox "Backup drive (" & Left(s2ndBackupLocn, 2) & ") " & _
                     "specified in General Settings " & _
                    "is not ready/available", vbOKOnly + vbExclamation, AppName
    
        End If
    Else
        MsgBox "Invalid 2nd backup drive " & _
                "specified in General Settings. 2nd backup not done." _
               , vbOKOnly + vbExclamation, AppName
    End If
        
    Set fso = Nothing

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Sub Do3rdBackup(s2ndBackupLocn As String)
On Error GoTo ErrorTrap
Dim fso As New FileSystemObject
        
    If Trim(s2ndBackupLocn) <> "" Then
        If fso.DriveExists(Left(s2ndBackupLocn, 2)) Then
            If fso.GetDrive(Left(s2ndBackupLocn, 2)).IsReady Then
                If fso.FolderExists(s2ndBackupLocn) Then
                    On Error Resume Next
                    fso.CopyFile JustTheDirectory & "\" & TheMDBFile & ".cms", _
                                 s2ndBackupLocn & "\" & TheMDBFile & ".cms", _
                                 True
                    If Err.number <> 0 Then
                        MsgBox "Backup failed. Check that " & _
                                s2ndBackupLocn & " is writeable " & _
                                "and that you have " & _
                                "necessary permissions to write to it.", vbOKOnly + vbExclamation, AppName
                    End If
                    On Error GoTo ErrorTrap
                Else
                    MsgBox "Backup folder (" & s2ndBackupLocn & ") " & _
                            "specified in General Settings " & _
                           "does not exist", vbOKOnly + vbExclamation, AppName
                    
                End If
            Else
                 MsgBox "Backup drive (" & Left(s2ndBackupLocn, 2) & ") " & _
                         "specified in General Settings " & _
                        "is not ready/available", vbOKOnly + vbExclamation, AppName
            End If
        Else
             MsgBox "Backup drive (" & Left(s2ndBackupLocn, 2) & ") " & _
                     "specified in General Settings " & _
                    "is not ready/available", vbOKOnly + vbExclamation, AppName
    
        End If
    Else
        MsgBox "Invalid 3rd backup drive " & _
                "specified in General Settings. 3rd backup not done." _
               , vbOKOnly + vbExclamation, AppName
    End If
        
    Set fso = Nothing

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Function GetUserID(ByVal UserCode As Long, TheDB As DAO.Database) As String
On Error GoTo ErrorTrap
    Dim rstUser As Recordset

    Set rstUser = TheDB.OpenRecordset("SELECT DISTINCTROW TheUserID " & _
                                        "FROM tblSecurity " & _
                                        "WHERE UserCode = " & UserCode, dbOpenForwardOnly)
                                
    With rstUser
    
    If Not .BOF Then
        GetUserID = !TheUserID
    Else
        GetUserID = ""
    End If
    
    .Close
    
    End With
    
    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Function IsWindowsNT() As Boolean
On Error GoTo ErrorTrap
    
    If Environ$("OS") <> "" Then
        IsWindowsNT = True
    Else
        IsWindowsNT = False
    End If
        
    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Function GetWindowsUserName() As String

Dim l As Long
Dim sUser As String

sUser = Space$(255)
l = GetUserName(sUser, 255)

'strip null terminator

If l <> 0 Then
   GetWindowsUserName = Left(sUser, InStr(sUser, Chr(0)) - 1)
Else
   Err.Raise Err.LastDllError, , _
     "A system call returned an error code of " _
      & Err.LastDllError
    EndProgram
End If

End Function
Public Function ServiceMtgOpeningSong(MeetingDate As String) As Long
Dim rst As Recordset, str As String, i As Long
On Error GoTo ErrorTrap

    
    If Not ValidDate(MeetingDate) Then
        ServiceMtgOpeningSong = 0
        Exit Function
    End If
    
    str = "SELECT tblServiceMtgs.SeqNum, " & _
          "       tblServiceMtgs.MeetingDate, " & _
          "       tblServiceMtgs.ItemTypeID, " & _
          "       tblServiceMtgs.ItemName, " & _
          "       tblServiceMtgs.ItemLength, " & _
          "       tblServiceMtgs.PersonID, " & _
          "       tblServiceMtgs.Announcements " & _
          "FROM tblServiceMtgs " & _
          "WHERE MeetingDate = #" & Format(MeetingDate, "mm/dd/yyyy") & "# " & _
          " AND ItemTypeID = 3 "
          
    Set rst = CMSDB.OpenRecordset(str, dbOpenForwardOnly)
    
    With rst
    
    If .BOF Then
        ServiceMtgOpeningSong = 0
    Else
        ServiceMtgOpeningSong = CLng(!ItemName)
    End If
    
    .Close
    
    End With
    
    Set rst = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function GetAccountStartAmount(AccountID) As Double
Dim rst As Recordset, str As String, i As Long
On Error GoTo ErrorTrap

    If AccountID >= 0 Then
        str = "SELECT StartAmount " & _
              "FROM tblBankAccounts " & _
              "WHERE AccountID = " & AccountID
    Else
        str = "SELECT SUM(StartAmount) AS StartAmount " & _
              "FROM tblBankAccounts "
    End If
    
    Set rst = CMSDB.OpenRecordset(str, dbOpenForwardOnly)
    
    With rst
    
    If .BOF Then
        GetAccountStartAmount = 0
    Else
        GetAccountStartAmount = HandleNull(!StartAmount, 0)
    End If
    
    .Close
    
    End With
    
    Set rst = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function GetAccountStartDate(AccountID) As Date
Dim rst As Recordset, str As String, i As Long
On Error GoTo ErrorTrap

        
    str = "SELECT StartDate " & _
          "FROM tblBankAccounts " & _
          "WHERE AccountID = " & AccountID
          
    Set rst = CMSDB.OpenRecordset(str, dbOpenForwardOnly)
    
    With rst
    
    If .BOF Then
        GetAccountStartDate = 0
    Else
        GetAccountStartDate = !StartDate
    End If
    
    .Close
    
    End With
    
    Set rst = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function GetAccountName(AccountID) As String
Dim rst As Recordset, str As String, i As Long
On Error GoTo ErrorTrap

        
    str = "SELECT AccountName " & _
          "FROM tblBankAccounts " & _
          "WHERE AccountID = " & AccountID
          
    Set rst = CMSDB.OpenRecordset(str, dbOpenForwardOnly)
    
    With rst
    
    If .BOF Then
        GetAccountName = ""
    Else
        GetAccountName = !AccountName
    End If
    
    .Close
    
    End With
    
    Set rst = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function


Public Sub GetLastExportDate(ExporterUserCode As Long, ExportDate As Date, TheDB As DAO.Database)
On Error GoTo ErrorTrap
    Dim rstTemp As Recordset

    Set rstTemp = TheDB.OpenRecordset("SELECT UserCode, " & _
                                      "       LastExportDate " & _
                                      "FROM tblLastExportDate ", dbOpenForwardOnly)
                                            
    With rstTemp
    
    ExportDate = Format(!LastExportDate, "mm/dd/yyyy hh:mm:ss")
    ExporterUserCode = !UserCode
    
    .Close
    
    End With
    
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Public Sub ReplaceTableFromExternalDB(ByVal TableName As String, _
                                      ByVal SourceDB As DAO.Database, _
                                      ByVal DestinationDB As DAO.Database, _
                                      ByVal DestinationDBPath As String, _
                                      Optional ByVal WhereSQL As String = "")

Dim TheSQL As String
On Error GoTo ErrorTrap

    DestinationDB.Execute "DELETE * FROM " & TableName
    
    TheSQL = "INSERT INTO " & TableName & _
            " IN '" & DestinationDBPath & _
            "' SELECT * " & _
            " FROM " & TableName & " " & WhereSQL & ";"

    SourceDB.Execute TheSQL
    
'    DestinationDB.Close
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Sub CopyTableFromExternalDB(ByVal TableName As String, _
                                      ByVal SourceDB As DAO.Database, _
                                      ByVal DestinationDB As DAO.Database, _
                                      ByVal DestinationDBPath As String, _
                                      Optional ByVal DropTableFirst As Boolean = True)

Dim TheSQL As String
On Error GoTo ErrorTrap

    If DropTableFirst Then
        On Error Resume Next
        DestinationDB.Execute "DROP TABLE [" & TableName & "]"
        On Error GoTo ErrorTrap
    End If
    
    TheSQL = "SELECT * INTO " & TableName & _
            " IN '" & DestinationDBPath & _
            "' FROM " & TableName & ";"

    SourceDB.Execute TheSQL
    
'    DestinationDB.Close
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Public Sub SwitchOffDAO()

On Error GoTo ErrorTrap

    DestroyGlobalObjects
    CMSDB.Close
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Public Sub SwitchOnDAO()

On Error GoTo ErrorTrap
    
    ConnectToDAO (CompletePathToTheMDBFileAndExt)
    InstantiateGlobalObjects
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub


