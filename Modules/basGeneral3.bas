Attribute VB_Name = "basGeneral3"
Option Explicit


Public Sub ImportTMSScheduleTable()
On Error GoTo ErrorTrap
Dim rstCMS_Update As Recordset, rstCMS_LIVE As Recordset
'
'This requires that CMSDB and CMS_UPDATE_DB connections are open
'
    Set rstCMS_Update = CMS_Update_DB.OpenRecordset("SELECT ScheduleSeqNum, " & _
                                                    "       AssignmentDate, " & _
                                                    "       TalkNo, " & _
                                                    "       SchoolNo, " & _
                                                    "       TalkSeqNum, " & _
                                                    "       PersonID, " & _
                                                    "       Assistant1ID, " & _
                                                    "       Assistant2ID, " & _
                                                    "       Setting, " & _
                                                    "       CounselPoint, " & _
                                                    "       CounselPointAssignedDate, " & _
                                                    "       CounselPointCompletedDate, " & _
                                                    "       TalkCompleted, " & _
                                                    "       TalkDefaulted, " & _
                                                    "       ExerciseComplete, " & _
                                                    "       DiscussedWithStudent, " & _
                                                    "       IsVolunteer, " & _
                                                    "       Comment, " & _
                                                    "       SlipPrinted, " & _
                                                    "       SlipHandedOver, " & _
                                                    "       StudentNote " & _
                                                    "FROM tblTMSSchedule", _
                                                    dbOpenDynaset)
                                                    
    If rstCMS_Update.BOF Then
        Exit Sub
    Else
        Set rstCMS_LIVE = CMSDB.OpenRecordset("tblTMSSchedule", dbOpenDynaset)
    End If

    With rstCMS_Update
    '
    'For each row on CMS_UPDATE.tbleTMSSchedule, see if matching row exists on
    ' CMS_LIVE.tbleTMSSchedule. If so, update it (except SlipPrinted flag). Otherwise
    ' insert it.
    '
    Do Until .EOF
        rstCMS_LIVE.FindFirst "ScheduleSeqNum =  " & !ScheduleSeqNum & _
                             " AND AssignmentDate = #" & Format(!AssignmentDate, "mm/dd/yyyy") & _
                             "# AND TalkNo = '" & !TalkNo & _
                             "' AND SchoolNo = " & !SchoolNo
    
        If rstCMS_LIVE.NoMatch Then
        'This is a new row, so insert it on CMS_LIVE
            rstCMS_LIVE.AddNew
                rstCMS_LIVE!ScheduleSeqNum = !ScheduleSeqNum
                rstCMS_LIVE!AssignmentDate = !AssignmentDate
                rstCMS_LIVE!TalkNo = !TalkNo
                rstCMS_LIVE!SchoolNo = !SchoolNo
                rstCMS_LIVE!TalkSeqNum = !TalkSeqNum
                rstCMS_LIVE!PersonID = !PersonID
                rstCMS_LIVE!Assistant1ID = !Assistant1ID
                rstCMS_LIVE!Assistant2ID = !Assistant2ID
                rstCMS_LIVE!Setting = !Setting
                rstCMS_LIVE!CounselPoint = !CounselPoint
                rstCMS_LIVE!CounselPointAssignedDate = !CounselPointAssignedDate
                rstCMS_LIVE!CounselPointCompletedDate = !CounselPointCompletedDate
                rstCMS_LIVE!TalkCompleted = !TalkCompleted
                rstCMS_LIVE!talkdefaulted = !talkdefaulted
                rstCMS_LIVE!ExerciseComplete = !ExerciseComplete
                rstCMS_LIVE!DiscussedWithStudent = !DiscussedWithStudent
                rstCMS_LIVE!IsVolunteer = !IsVolunteer
                rstCMS_LIVE!Comment = !Comment
                rstCMS_LIVE!SlipPrinted = !SlipPrinted 'May have been printed by exporter
                rstCMS_LIVE!SlipHandedOver = !SlipHandedOver
                rstCMS_LIVE!StudentNote = !StudentNote
            rstCMS_LIVE.Update
        Else
        'This is an existing row, so update it on CMS_LIVE (except for SlipPrinted flag)
            rstCMS_LIVE.Edit
                rstCMS_LIVE!TalkSeqNum = !TalkSeqNum
                rstCMS_LIVE!PersonID = !PersonID
                rstCMS_LIVE!Assistant1ID = !Assistant1ID
                rstCMS_LIVE!Assistant2ID = !Assistant2ID
                rstCMS_LIVE!Setting = !Setting
                rstCMS_LIVE!CounselPoint = !CounselPoint
                rstCMS_LIVE!CounselPointAssignedDate = !CounselPointAssignedDate
                rstCMS_LIVE!CounselPointCompletedDate = !CounselPointCompletedDate
                rstCMS_LIVE!TalkCompleted = !TalkCompleted
                rstCMS_LIVE!talkdefaulted = !talkdefaulted
                rstCMS_LIVE!ExerciseComplete = !ExerciseComplete
                rstCMS_LIVE!DiscussedWithStudent = !DiscussedWithStudent
                rstCMS_LIVE!IsVolunteer = !IsVolunteer
                rstCMS_LIVE!Comment = !Comment
                rstCMS_LIVE!SlipHandedOver = !SlipHandedOver
                rstCMS_LIVE!StudentNote = !StudentNote
            rstCMS_LIVE.Update
        End If
        
        .MoveNext
    Loop
    
    '
    'For each row on CMS_LIVE.tblTMSSchedule, see if matching row exists on
    ' CMS_UPDATE.tblTMSSchedule. If not delete it.
    'There shouldn't be any rows that fall into this category since rows
    ' aren't physically deleted when a student is removed from the table
    ' .... but just in case.....
    '
    rstCMS_LIVE.MoveFirst
    
    Do Until rstCMS_LIVE.EOF
        .FindFirst "ScheduleSeqNum =  " & rstCMS_LIVE!ScheduleSeqNum & _
                  " AND AssignmentDate = #" & Format(rstCMS_LIVE!AssignmentDate, "mm/dd/yyyy") & _
                 "# AND TalkNo = '" & rstCMS_LIVE!TalkNo & _
                 "' AND SchoolNo = " & rstCMS_LIVE!SchoolNo
    
        If .NoMatch Then
        'This row is on CMS_LIVE but not on CMS_UPDATE, so delete it.
            rstCMS_LIVE.Delete
        End If
        
        rstCMS_LIVE.MoveNext
    Loop
    
    End With

    rstCMS_LIVE.Close
    rstCMS_Update.Close
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Public Sub ImportNonTMSRoles()
On Error GoTo ErrorTrap
Dim rstCMS_Update As Recordset, rstCMS_LIVE As Recordset
'
'This requires that CMSDB and CMS_UPDATE_DB connections are open
'

'
'First, get all non-TMS rows from tblTaskAndPerson....
'
    Set rstCMS_Update = CMS_Update_DB.OpenRecordset("SELECT CongNo, " & _
                                                    "       TaskCategory, " & _
                                                    "       TaskSubCategory, " & _
                                                    "       Task, " & _
                                                    "       Person, " & _
                                                    "       OnSunday, " & _
                                                    "       OnMidweek " & _
                                                    "FROM tblTaskAndPerson " & _
                                                    "WHERE TaskSubCategory <> 6", _
                                                    dbOpenDynaset)
                                                    
    If rstCMS_Update.BOF Then
        Exit Sub
    Else
        Set rstCMS_LIVE = CMSDB.OpenRecordset("SELECT CongNo, " & _
                                                "       TaskCategory, " & _
                                                "       TaskSubCategory, " & _
                                                "       Task, " & _
                                                "       Person, " & _
                                                "       OnSunday, " & _
                                                "       OnMidweek " & _
                                                "FROM tblTaskAndPerson " & _
                                                "WHERE TaskSubCategory <> 6", _
                                                dbOpenDynaset)
    End If

    With rstCMS_Update
    '
    'For each row on CMS_UPDATE.tblTaskAndPerson, see if matching row exists on
    ' CMS_LIVE.tblTaskAndPerson. If so, update it (except SlipPrinted flag). Otherwise
    ' insert it.
    '
    Do Until .EOF
        rstCMS_LIVE.FindFirst "CongNo =  " & !CongNo & _
                             " AND TaskCategory = " & !TaskCategory & _
                             " AND TaskSubCategory = " & !TaskSubcategory & _
                             " AND Task = " & !Task & _
                             " AND Person = " & !Person

        If rstCMS_LIVE.NoMatch Then
        'This is a new row, so insert it on CMS_LIVE
            rstCMS_LIVE.AddNew
                rstCMS_LIVE!CongNo = !CongNo
                rstCMS_LIVE!TaskCategory = !TaskCategory
                rstCMS_LIVE!TaskSubcategory = !TaskSubcategory
                rstCMS_LIVE!Task = !Task
                rstCMS_LIVE!Person = !Person
                rstCMS_LIVE!OnSunday = !OnSunday
                rstCMS_LIVE!OnMidweek = !OnMidweek
            rstCMS_LIVE.Update
        Else
        'This is an existing row, so update it on CMS_LIVE
            rstCMS_LIVE.Edit
                rstCMS_LIVE!OnSunday = !OnSunday
                rstCMS_LIVE!OnMidweek = !OnMidweek
            rstCMS_LIVE.Update
        End If
        
        .MoveNext
    Loop
    
    '
    'For each row on CMS_LIVE.tblTaskAndPerson, see if matching row exists on
    ' CMS_UPDATE.tblTaskAndPerson. If not delete it.
    '
    rstCMS_LIVE.MoveFirst
    
    Do Until rstCMS_LIVE.EOF
        .FindFirst "CongNo =  " & rstCMS_LIVE!CongNo & _
                " AND TaskCategory = " & rstCMS_LIVE!TaskCategory & _
                " AND TaskSubCategory = " & rstCMS_LIVE!TaskSubcategory & _
                " AND Task = " & rstCMS_LIVE!Task & _
                " AND Person = " & rstCMS_LIVE!Person
    
        If .NoMatch Then
        'This row is on CMS_LIVE but not on CMS_UPDATE, so delete it.
            rstCMS_LIVE.Delete
        End If
        
        rstCMS_LIVE.MoveNext
    Loop
    
    End With

'
'Next: tblTaskPersonSuspendDates....
'
    Set rstCMS_Update = CMS_Update_DB.OpenRecordset("SELECT CongNo, " & _
                                                    "       TaskCategory, " & _
                                                    "       TaskSubCategory, " & _
                                                    "       Task, " & _
                                                    "       Person, " & _
                                                    "       SuspendStartDate, " & _
                                                    "       SuspendEndDate, " & _
                                                    "       SuspendReason " & _
                                                    "FROM tblTaskPersonSuspendDates " & _
                                                    "WHERE TaskSubCategory <> 6", _
                                                    dbOpenDynaset)
                                                    
    If rstCMS_Update.BOF Then
        Exit Sub
    Else
        Set rstCMS_LIVE = CMSDB.OpenRecordset("SELECT CongNo, " & _
                                                "       TaskCategory, " & _
                                                "       TaskSubCategory, " & _
                                                "       Task, " & _
                                                "       Person, " & _
                                                "       SuspendStartDate, " & _
                                                "       SuspendEndDate, " & _
                                                "       SuspendReason " & _
                                                "FROM tblTaskPersonSuspendDates " & _
                                                "WHERE TaskSubCategory <> 6", _
                                                dbOpenDynaset)
    End If

    With rstCMS_Update
    '
    'For each row on CMS_UPDATE.tblTaskAndPerson, see if matching row exists on
    ' CMS_LIVE.tblTaskAndPerson. If so, leave it. Otherwise insert it.
    '
    Do Until .EOF
        rstCMS_LIVE.FindFirst "CongNo =  " & !CongNo & _
                             " AND TaskCategory = " & !TaskCategory & _
                             " AND TaskSubCategory = " & !TaskSubcategory & _
                             " AND Task = " & !Task & _
                             " AND Person = " & !Person

        If rstCMS_LIVE.NoMatch Then
        'This is a new row, so insert it on CMS_LIVE
            rstCMS_LIVE.AddNew
                rstCMS_LIVE!CongNo = !CongNo
                rstCMS_LIVE!TaskCategory = !TaskCategory
                rstCMS_LIVE!TaskSubcategory = !TaskSubcategory
                rstCMS_LIVE!Task = !Task
                rstCMS_LIVE!Person = !Person
            rstCMS_LIVE.Update
        Else
            rstCMS_LIVE.Edit
                rstCMS_LIVE!SuspendStartDate = !SuspendStartDate
                rstCMS_LIVE!SuspendEndDate = !SuspendEndDate
                rstCMS_LIVE!SuspendReason = !SuspendReason
            rstCMS_LIVE.Update
        End If
        
        .MoveNext
    Loop
    
    '
    'For each row on CMS_LIVE.tblTaskPersonSuspendDates, see if matching row exists on
    ' CMS_UPDATE.tblTaskPersonSuspendDates. If not delete it.
    '
    rstCMS_LIVE.MoveFirst
    
    Do Until rstCMS_LIVE.EOF
        .FindFirst "CongNo =  " & rstCMS_LIVE!CongNo & _
                " AND TaskCategory = " & rstCMS_LIVE!TaskCategory & _
                " AND TaskSubCategory = " & rstCMS_LIVE!TaskSubcategory & _
                " AND Task = " & rstCMS_LIVE!Task & _
                " AND Person = " & rstCMS_LIVE!Person
    
        If .NoMatch Then
        'This row is on CMS_LIVE but not on CMS_UPDATE, so delete it.
            rstCMS_LIVE.Delete
        End If
        
        rstCMS_LIVE.MoveNext
    Loop
    
    End With
    rstCMS_LIVE.Close
    rstCMS_Update.Close
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub
    
    

Public Sub ImportTMSRoles()
On Error GoTo ErrorTrap
Dim rstCMS_Update As Recordset, rstCMS_LIVE As Recordset
'
'This requires that CMSDB and CMS_UPDATE_DB connections are open
'

'
'First: tblTaskAndPerson....
'
    Set rstCMS_Update = CMS_Update_DB.OpenRecordset("SELECT CongNo, " & _
                                                    "       TaskCategory, " & _
                                                    "       TaskSubCategory, " & _
                                                    "       Task, " & _
                                                    "       Person, " & _
                                                    "       OnSunday, " & _
                                                    "       OnMidweek " & _
                                                    "FROM tblTaskAndPerson " & _
                                                    "WHERE TaskSubCategory = 6", _
                                                    dbOpenDynaset)
                                                    
    If rstCMS_Update.BOF Then
        Exit Sub
    Else
        Set rstCMS_LIVE = CMSDB.OpenRecordset("SELECT CongNo, " & _
                                                "       TaskCategory, " & _
                                                "       TaskSubCategory, " & _
                                                "       Task, " & _
                                                "       Person, " & _
                                                "       OnSunday, " & _
                                                "       OnMidweek " & _
                                                "FROM tblTaskAndPerson " & _
                                                "WHERE TaskSubCategory = 6", _
                                                dbOpenDynaset)
    End If

    With rstCMS_Update
    '
    'For each row on CMS_UPDATE.tblTaskAndPerson, see if matching row exists on
    ' CMS_LIVE.tblTaskAndPerson. If so, leave it. Otherwise insert it.
    '
    Do Until .EOF
        rstCMS_LIVE.FindFirst "CongNo =  " & !CongNo & _
                             " AND TaskCategory = " & !TaskCategory & _
                             " AND TaskSubCategory = " & !TaskSubcategory & _
                             " AND Task = " & !Task & _
                             " AND Person = " & !Person

        If rstCMS_LIVE.NoMatch Then
        'This is a new row, so insert it on CMS_LIVE
            rstCMS_LIVE.AddNew
                rstCMS_LIVE!CongNo = !CongNo
                rstCMS_LIVE!TaskCategory = !TaskCategory
                rstCMS_LIVE!TaskSubcategory = !TaskSubcategory
                rstCMS_LIVE!Task = !Task
                rstCMS_LIVE!Person = !Person
                rstCMS_LIVE!OnSunday = !OnSunday
                rstCMS_LIVE!OnMidweek = !OnMidweek
            rstCMS_LIVE.Update
        End If
        
        .MoveNext
    Loop
    
    '
    'For each row on CMS_LIVE.tblTaskAndPerson, see if matching row exists on
    ' CMS_UPDATE.tblTaskAndPerson. If not delete it.
    '
    rstCMS_LIVE.MoveFirst
    
    Do Until rstCMS_LIVE.EOF
        .FindFirst "CongNo =  " & rstCMS_LIVE!CongNo & _
                " AND TaskCategory = " & rstCMS_LIVE!TaskCategory & _
                " AND TaskSubCategory = " & rstCMS_LIVE!TaskSubcategory & _
                " AND Task = " & rstCMS_LIVE!Task & _
                " AND Person = " & rstCMS_LIVE!Person
    
        If .NoMatch Then
        'This row is on CMS_LIVE but not on CMS_UPDATE, so delete it.
            rstCMS_LIVE.Delete
        End If
        
        rstCMS_LIVE.MoveNext
    Loop
    
    End With

'
'Next: tblTaskPersonSuspendDates....
'
    Set rstCMS_Update = CMS_Update_DB.OpenRecordset("SELECT CongNo, " & _
                                                    "       TaskCategory, " & _
                                                    "       TaskSubCategory, " & _
                                                    "       Task, " & _
                                                    "       Person, " & _
                                                    "       SuspendStartDate, " & _
                                                    "       SuspendEndDate, " & _
                                                    "       SuspendReason " & _
                                                    "FROM tblTaskPersonSuspendDates " & _
                                                    "WHERE TaskSubCategory = 6", _
                                                    dbOpenDynaset)
                                                    
    If rstCMS_Update.BOF Then
        Exit Sub
    Else
        Set rstCMS_LIVE = CMSDB.OpenRecordset("SELECT CongNo, " & _
                                                "       TaskCategory, " & _
                                                "       TaskSubCategory, " & _
                                                "       Task, " & _
                                                "       Person, " & _
                                                "       SuspendStartDate, " & _
                                                "       SuspendEndDate, " & _
                                                "       SuspendReason " & _
                                                "FROM tblTaskPersonSuspendDates " & _
                                                "WHERE TaskSubCategory = 6", _
                                                dbOpenDynaset)
    End If

    With rstCMS_Update
    '
    'For each row on CMS_UPDATE.tblTaskAndPerson, see if matching row exists on
    ' CMS_LIVE.tblTaskAndPerson. If so, leave it. Otherwise insert it.
    '
    Do Until .EOF
        rstCMS_LIVE.FindFirst "CongNo =  " & !CongNo & _
                             " AND TaskCategory = " & !TaskCategory & _
                             " AND TaskSubCategory = " & !TaskSubcategory & _
                             " AND Task = " & !Task & _
                             " AND Person = " & !Person

        If rstCMS_LIVE.NoMatch Then
        'This is a new row, so insert it on CMS_LIVE
            rstCMS_LIVE.AddNew
                rstCMS_LIVE!CongNo = !CongNo
                rstCMS_LIVE!TaskCategory = !TaskCategory
                rstCMS_LIVE!TaskSubcategory = !TaskSubcategory
                rstCMS_LIVE!Task = !Task
                rstCMS_LIVE!Person = !Person
            rstCMS_LIVE.Update
        Else
            rstCMS_LIVE.Edit
                rstCMS_LIVE!SuspendStartDate = !SuspendStartDate
                rstCMS_LIVE!SuspendEndDate = !SuspendEndDate
                rstCMS_LIVE!SuspendReason = !SuspendReason
            rstCMS_LIVE.Update
        End If
        
        .MoveNext
    Loop
    
    '
    'For each row on CMS_LIVE.tblTaskPersonSuspendDates, see if matching row exists on
    ' CMS_UPDATE.tblTaskPersonSuspendDates. If not delete it.
    '
    rstCMS_LIVE.MoveFirst
    
    Do Until rstCMS_LIVE.EOF
        .FindFirst "CongNo =  " & rstCMS_LIVE!CongNo & _
                " AND TaskCategory = " & rstCMS_LIVE!TaskCategory & _
                " AND TaskSubCategory = " & rstCMS_LIVE!TaskSubcategory & _
                " AND Task = " & rstCMS_LIVE!Task & _
                " AND Person = " & rstCMS_LIVE!Person
    
        If .NoMatch Then
        'This row is on CMS_LIVE but not on CMS_UPDATE, so delete it.
            rstCMS_LIVE.Delete
        End If
        
        rstCMS_LIVE.MoveNext
    Loop
    
    End With
    rstCMS_LIVE.Close
    rstCMS_Update.Close
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Public Sub ImportTMSSlipPrintFlag()

On Error GoTo ErrorTrap
Dim rstCMS_Update As Recordset, rstCMS_LIVE As Recordset
'
'This requires that CMSDB and CMS_UPDATE_DB connections are open
'
    Set rstCMS_Update = CMS_Update_DB.OpenRecordset("SELECT ScheduleSeqNum, " & _
                                                    "       AssignmentDate, " & _
                                                    "       TalkNo, " & _
                                                    "       SchoolNo, " & _
                                                    "       TalkSeqNum, " & _
                                                    "       PersonID, " & _
                                                    "       Assistant1ID, " & _
                                                    "       Assistant2ID, " & _
                                                    "       Setting, " & _
                                                    "       CounselPoint, " & _
                                                    "       CounselPointAssignedDate, " & _
                                                    "       CounselPointCompletedDate, " & _
                                                    "       TalkCompleted, " & _
                                                    "       TalkDefaulted, " & _
                                                    "       ExerciseComplete, " & _
                                                    "       DiscussedWithStudent, " & _
                                                    "       IsVolunteer, " & _
                                                    "       Comment, " & _
                                                    "       SlipPrinted, " & _
                                                    "       SlipHandedOver, " & _
                                                    "       StudentNote " & _
                                                    "FROM tblTMSSchedule", _
                                                    dbOpenDynaset)
                                                    
    If rstCMS_Update.BOF Then
        Exit Sub
    Else
        Set rstCMS_LIVE = CMSDB.OpenRecordset("tblTMSSchedule", dbOpenDynaset)
    End If

    With rstCMS_Update
    '
    'For each row on CMS_UPDATE.tblTMSSchedule, see if matching row exists on
    ' CMS_LIVE.tblTMSSchedule. If so, update the SlipPrinted flag
    '
    Do Until .EOF
        rstCMS_LIVE.FindFirst "ScheduleSeqNum =  " & !ScheduleSeqNum & _
                             " AND AssignmentDate = #" & Format(!AssignmentDate, "mm/dd/yyyy") & _
                             "# AND TalkNo = '" & !TalkNo & _
                             "' AND SchoolNo = " & !SchoolNo
    
        If rstCMS_LIVE.NoMatch Then
        Else
        'This is an existing row, so update the SlipPrinted flag - if the slip
        ' has been printed by the exporter.
            If !SlipPrinted = True Then
                rstCMS_LIVE.Edit
                    rstCMS_LIVE!SlipPrinted = !SlipPrinted
                rstCMS_LIVE.Update
            End If
        End If
        
        .MoveNext
    Loop
        
    End With

    rstCMS_LIVE.Close
    rstCMS_Update.Close
    
    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Public Sub ImportUserDefinedQueries()

On Error GoTo ErrorTrap
Dim rstCMS_Update As Recordset, rstCMS_LIVE As Recordset
'
'This requires that CMSDB and CMS_UPDATE_DB connections are open
'

'First - tblUserQueries

    Set rstCMS_Update = CMS_Update_DB.OpenRecordset("SELECT SeqNum, " & _
                                                    "       QueryName, " & _
                                                    "       QueryString, " & _
                                                    "       QueryDescription " & _
                                                    "FROM tblUserQueries " & _
                                                    "WHERE Private = FALSE ", _
                                                    dbOpenDynaset)
                                                    
    If rstCMS_Update.BOF Then
        Exit Sub
    Else
        Set rstCMS_LIVE = CMSDB.OpenRecordset("tblUserQueries", dbOpenDynaset)
    End If

    With rstCMS_Update
    '
    'For each row on CMS_UPDATE.tblUserQueries, see if matching row exists on
    ' CMS_LIVE.tblUserQueries.
    '
    Do Until .EOF
        rstCMS_LIVE.FindFirst "QueryName =  '" & DoubleUpSingleQuotes(!QueryName) & "'"
    
        If rstCMS_LIVE.NoMatch Then
            rstCMS_LIVE.AddNew
            rstCMS_LIVE!QueryName = !QueryName
            rstCMS_LIVE!QueryString = !QueryString
            rstCMS_LIVE!QueryDescription = !QueryDescription
            rstCMS_LIVE.Update
        Else
            If (rstCMS_LIVE!QueryString <> !QueryString Or _
                rstCMS_LIVE!QueryDescription <> !QueryDescription) Then
                If MsgBox("Overwrite query '" & !QueryName & "'?", vbYesNo + vbQuestion, AppName) = vbYes Then
                    rstCMS_LIVE.Edit
                    rstCMS_LIVE!QueryString = !QueryString
                    rstCMS_LIVE!QueryDescription = !QueryDescription
                    rstCMS_LIVE.Update
                End If
            End If
        End If
        
        .MoveNext
    Loop
        
    End With


'Next - tblCustomRotaDetails

    Set rstCMS_Update = CMS_Update_DB.OpenRecordset("SELECT SeqNum, " & _
                                                    "       RotaName, " & _
                                                    "       QueryString, " & _
                                                    "       FrequencyID, " & _
                                                    "       DateFormatID, " & _
                                                    "       DaysToInclude, " & _
                                                    "       Column1Name, " & _
                                                    "       Column2Name, " & _
                                                    "       EventsToSkip, " & _
                                                    "       PrevRotaLastDate, " & _
                                                    "       PrevRotaLastValue " & _
                                                    "FROM tblCustomRotaDetails " & _
                                                    "WHERE Private = FALSE ", _
                                                    dbOpenDynaset)
                                                    
    If rstCMS_Update.BOF Then
        Exit Sub
    Else
        Set rstCMS_LIVE = CMSDB.OpenRecordset("tblCustomRotaDetails", dbOpenDynaset)
    End If

    With rstCMS_Update
    '
    'For each row on CMS_UPDATE.tblUserQueries, see if matching row exists on
    ' CMS_LIVE.tblUserQueries.
    '
    Do Until .EOF
        rstCMS_LIVE.FindFirst "RotaName =  '" & DoubleUpSingleQuotes(!RotaName) & "'"
    
        If rstCMS_LIVE.NoMatch Then
            rstCMS_LIVE.AddNew
            rstCMS_LIVE!RotaName = !RotaName
            rstCMS_LIVE!QueryString = !QueryString
            rstCMS_LIVE!FrequencyID = !FrequencyID
            rstCMS_LIVE!DateFormatID = !DateFormatID
            rstCMS_LIVE!DaysToInclude = !DaysToInclude
            rstCMS_LIVE!Column1Name = !Column1Name
            rstCMS_LIVE!Column2Name = !Column2Name
            rstCMS_LIVE!EventsToSkip = !EventsToSkip
            rstCMS_LIVE!PrevRotaLastDate = !PrevRotaLastDate
            rstCMS_LIVE!PrevRotaLastValue = !PrevRotaLastValue
            rstCMS_LIVE.Update
        Else
            If (rstCMS_LIVE!QueryString <> !QueryString Or _
                rstCMS_LIVE!FrequencyID <> !FrequencyID Or _
                rstCMS_LIVE!DateFormatID <> !DateFormatID Or _
                rstCMS_LIVE!DaysToInclude <> !DaysToInclude Or _
                rstCMS_LIVE!Column1Name <> !Column1Name Or _
                rstCMS_LIVE!Column2Name <> !Column2Name Or _
                rstCMS_LIVE!EventsToSkip <> !EventsToSkip Or _
                rstCMS_LIVE!PrevRotaLastDate <> !PrevRotaLastDate Or _
                rstCMS_LIVE!PrevRotaLastValue <> !PrevRotaLastValue) Then
                If MsgBox("Overwrite custom rota '" & !RotaName & "'?", vbYesNo + vbQuestion, AppName) = vbYes Then
                    rstCMS_LIVE.Edit
                    rstCMS_LIVE!QueryString = !QueryString
                    rstCMS_LIVE!FrequencyID = !FrequencyID
                    rstCMS_LIVE!DateFormatID = !DateFormatID
                    rstCMS_LIVE!DaysToInclude = !DaysToInclude
                    rstCMS_LIVE!Column1Name = !Column1Name
                    rstCMS_LIVE!Column2Name = !Column2Name
                    rstCMS_LIVE!EventsToSkip = !EventsToSkip
                    rstCMS_LIVE!PrevRotaLastDate = !PrevRotaLastDate
                    rstCMS_LIVE!PrevRotaLastValue = !PrevRotaLastValue
                    rstCMS_LIVE.Update
                End If
            End If
        End If
        
        .MoveNext
    Loop
        
    End With


    rstCMS_LIVE.Close
    rstCMS_Update.Close
    
    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Public Sub UpdateSecurityTable()

On Error GoTo ErrorTrap
Dim rstCMS_Update As Recordset, rstCMS_LIVE As Recordset
'
'This requires that CMSDB and CMS_UPDATE_DB connections are open
'
    Set rstCMS_Update = CMS_Update_DB.OpenRecordset("SELECT UserCode, " & _
                                                    "       TheUserID, " & _
                                                    "       ThePassword, " & _
                                                    "       ActiveFromDate, " & _
                                                    "       ActiveToDate " & _
                                                    "FROM tblSecurity", _
                                                    dbOpenDynaset)
                                                    
    If rstCMS_Update.BOF Then
        Exit Sub
    Else
        Set rstCMS_LIVE = CMSDB.OpenRecordset("tblSecurity", dbOpenDynaset)
    End If

    With rstCMS_Update
    '
    'For each row on CMS_UPDATE.tblSecurity, see if matching row exists on
    ' CMS_LIVE.tblSecurity. If so, update the UserID and (if Password is CMSPASSWORD)
    ' update the password too. Else insert.
    '
    Do Until .EOF
        rstCMS_LIVE.FindFirst "UserCode =  " & !UserCode
    
        If rstCMS_LIVE.NoMatch Then
            rstCMS_LIVE.AddNew
                rstCMS_LIVE!UserCode = !UserCode
                rstCMS_LIVE!TheUserID = !TheUserID
                rstCMS_LIVE!ThePassword = !ThePassword
                rstCMS_LIVE!ActiveFromDate = !ActiveFromDate
                rstCMS_LIVE!ActiveToDate = !ActiveToDate
            rstCMS_LIVE.Update
        Else
        'This is an existing row, so update...
            rstCMS_LIVE.Edit
                rstCMS_LIVE!TheUserID = !TheUserID
                rstCMS_LIVE!ActiveFromDate = !ActiveFromDate
                rstCMS_LIVE!ActiveToDate = !ActiveToDate
                If !ThePassword = "CMSPASSWORD" Then
                    If !UserCode = gCurrentUserCode Then
                      'This UserCode on rstCMS_Update is that of Current user
                        If MsgBox("Do you want your password reset?", _
                                vbYesNo + vbQuestion + vbDefaultButton1, AppName) = vbYes Then
                                    
                            rstCMS_LIVE!ThePassword = "CMSPASSWORD"
                            gbImportedResetPassword = True
                        End If
                    End If
                Else
                    If !UserCode = 1 Then
                        rstCMS_LIVE!ThePassword = "12Cm5Adm1N34"
                    End If
                End If
            rstCMS_LIVE.Update
        End If
        
        .MoveNext
    Loop
        
    End With

    rstCMS_LIVE.Close
    rstCMS_Update.Close
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Function SendMail(EmailAddresses As Collection, _
                         SubjectText As String, _
                         BodyText As String, _
                         AttachmentPaths As Collection) As Boolean
             
On Error GoTo ErrorTrap
Dim EmailAddress As Variant, AttachmentPath As Variant
Dim myOutlook As Object
Dim myMailItem As Object
'Dim myOutlook As Outlook.Application
'Dim myMailItem As Outlook.MailItem
    
    On Error GoTo ErrorTrap
    
    ' Instanciate MS Outlook Object....
'    Set myOutlook = New Outlook.Application
    Set myOutlook = CreateObject("Outlook.Application")
    
    
    ' Make mail item
    Set myMailItem = myOutlook.CreateItem(0)
    
    ' Set recipients
    For Each EmailAddress In EmailAddresses
        myMailItem.Recipients.Add EmailAddress
    Next
    
    ' Set subject
    myMailItem.Subject = SubjectText
    
    ' Set body
    myMailItem.Body = BodyText
    
    ' Add attachments
    For Each AttachmentPath In AttachmentPaths
        myMailItem.Attachments.Add AttachmentPath
    Next
    
    ' And send it!
    myMailItem.Send
    
    SendMail = True
               
    Exit Function
ErrorTrap:
    MsgBox "Could not send email. " & vbCrLf & vbCrLf & Err.Description, vbOKOnly + vbExclamation, AppName
    SendMail = False
    
End Function


Public Function SendEmailMAPI(SendTo As String, _
                         Subject As String, _
                         EmailText As String, _
                         TheMAPISessionControl As MAPISession, _
                         TheMAPIMessageControl As MAPIMessages, _
                         Optional AttachmentPath As String, _
                         Optional Attachment As String) As Boolean

' Sends an email to the appropriate person(s).
    
' SendTo = List of email addresses separated by a semicolon.  Example:
'                 sm@xyz.com; steve@work.com; jane@home.com
' Subject = Text that summarizes what the email is about
' EmailText = Body of text that is the email
' TheMAPISessionControl = fully qualified control eg frmTheForm.MAPISession1
' TheMAPIMessageControl = fully qualified control eg frmTheForm.MAPIMessages1
' AttachmentPath = Directory in which the attachment resides
' Attachment = File to send with the email
   
   Dim intStart As Integer
   Dim strSendTo As String
   Dim intEnd As Integer
   Dim i As Integer

   On Error GoTo ErrorTrap
   
   If TheMAPISessionControl.SessionID = 0 Then
      TheMAPISessionControl.SignOn
   End If

   If SendTo = "" Then Exit Function

   With TheMAPIMessageControl
      .SessionID = TheMAPISessionControl.SessionID
      .Compose

      'Make sure that the SendTo always has a trailing semi-colon (makes it
      ' easier below)
      'Strip out any spaces between names for consistency
      For i = 1 To Len(SendTo)
         If Mid$(SendTo, i, 1) <> " " Then
            strSendTo = strSendTo & Mid$(SendTo, i, 1)
         End If
      Next i

      SendTo = strSendTo
      If Right$(SendTo, 1) <> ";" Then
         SendTo = SendTo & ";"
      End If

      'Format each recipient, each are separated by a semi-colon, like this:
      '  steve.miller@aol.com;sm@psc.com; sm@teletech.com;
      intEnd = InStr(1, SendTo, ";")
      .RecipAddress = Mid$(SendTo, 1, intEnd - 1)
      .ResolveName

      intStart = intEnd + 1
      Do
         intEnd = InStr(intStart, SendTo, ";")
         If intEnd = 0 Then
            Exit Do
         Else
            .RecipIndex = .RecipIndex + 1
            .RecipAddress = Mid$(SendTo, intStart, intEnd - intStart)
            .ResolveName
         End If
         intStart = intEnd + 1
      Loop

      .MsgSubject = Subject
      .MsgNoteText = EmailText
      If Left$(Attachment, 1) = "\" Then
         Attachment = Mid$(Attachment, 2, Len(Attachment))
      End If

      If Attachment <> "" Then
         If Right$(AttachmentPath, 1) = "\" Then
            .AttachmentPathName = AttachmentPath & Attachment
         Else
            .AttachmentPathName = AttachmentPath & "\" & Attachment
         End If
        .AttachmentName = Attachment
      End If
      .Send False
   End With
   
   SendEmailMAPI = True

ExitMe:
   Exit Function

ErrorTrap:
    MsgBox "Could not send email. " & vbCrLf & vbCrLf & Err.Description, vbOKOnly + vbExclamation, AppName
   SendEmailMAPI = False
    

End Function

Public Function SendEmailMAPI2(SendTo As Collection, _
                         Subject As String, _
                         EmailText As String, _
                         TheMAPISessionControl As MAPISession, _
                         TheMAPIMessageControl As MAPIMessages, _
                         Optional AttachmentPath As String, _
                         Optional Attachment As String) As Boolean

' Sends an email to the appropriate person(s).
    
' SendTo = List of email addresses in a collection
' Subject = Text that summarizes what the email is about
' EmailText = Body of text that is the email
' TheMAPISessionControl = fully qualified control eg frmTheForm.MAPISession1
' TheMAPIMessageControl = fully qualified control eg frmTheForm.MAPIMessages1
' AttachmentPath = Directory in which the attachment resides
' Attachment = File to send with the email
   
   Dim i As Integer
   Dim TheAddress As Variant

   On Error GoTo ErrorTrap
   
   If TheMAPISessionControl.SessionID = 0 Then
'      TheMAPISessionControl.LogonUI = False
'      TheMAPISessionControl.UserName = "Michael J Thompson"
      TheMAPISessionControl.SignOn
   End If

   If SendTo.Count = 0 Then Exit Function

   With TheMAPIMessageControl
        .SessionID = TheMAPISessionControl.SessionID
        .Compose
        
        '
        'Get all addressees from collection into MAPI...
        '
        .RecipAddress = CStr(SendTo.Item(1))
        .ResolveName
        
        For i = 2 To SendTo.Count
            .RecipIndex = .RecipIndex + 1
            .RecipAddress = CStr(SendTo.Item(i))
            .ResolveName
        Next

        'Subject and Body text...
        '
        .MsgSubject = Subject
        .MsgNoteText = EmailText
        
        '
        'Now deal with the attachment....
        '
        If Left$(Attachment, 1) = "\" Then
           Attachment = Mid$(Attachment, 2, Len(Attachment))
        End If
      
        If Attachment <> "" Then
           If Right$(AttachmentPath, 1) = "\" Then
              .AttachmentPathName = AttachmentPath & Attachment
           Else
              .AttachmentPathName = AttachmentPath & "\" & Attachment
           End If
          .AttachmentName = Attachment
        End If
        .Send False
    
   End With

   SendEmailMAPI2 = True

ExitMe:
   Exit Function

ErrorTrap:
    MsgBox "Could not send email. " & vbCrLf & vbCrLf & Err.Description, vbOKOnly + vbExclamation, AppName
   SendEmailMAPI2 = False

End Function




Public Function BuiltInEventType(EventTypeID As Long) As Boolean
   On Error GoTo ErrorTrap

    Select Case EventTypeID
    Case 1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15
        BuiltInEventType = True
    Case Else
        BuiltInEventType = False
    End Select
    
   Exit Function

ErrorTrap:
    EndProgram
    
End Function
Public Function RecurringEventType(EventTypeID As Long) As Boolean
   On Error GoTo ErrorTrap

    Select Case EventTypeID
    Case 15
        RecurringEventType = True
    Case Else
        RecurringEventType = False
    End Select
    
   Exit Function

ErrorTrap:
    EndProgram
    
End Function

Public Function ActiveLogon() As Boolean
   On Error GoTo ErrorTrap

    If gCurrentUserCode <= 2 Then
        ActiveLogon = True
        Exit Function
    ElseIf gdActiveFromDate = "" Or gdActiveToDate = "" Then
        ActiveLogon = False
    ElseIf CDate(gdActiveFromDate) > CDate(Format(Now, "dd/mm/yyyy")) Then
        ActiveLogon = False
    ElseIf CDate(gdActiveToDate) < CDate(Format(Now, "dd/mm/yyyy")) Then
        ActiveLogon = False
    Else
        ActiveLogon = True
    End If
    
    
   Exit Function

ErrorTrap:
    EndProgram
    
End Function
Public Function ServiceYear(TheDate As Date) As Long
   On Error GoTo ErrorTrap

    If Month(TheDate) >= 9 And Month(TheDate) <= 12 Then
        ServiceYear = CLng(year(TheDate)) + 1
    Else
        ServiceYear = CLng(year(TheDate))
    End If
    
   Exit Function

ErrorTrap:
    EndProgram
    
End Function
Public Function ConvertMonthNoToServiceMonthNo(MonthNo As Long) As Long
   On Error GoTo ErrorTrap

    Select Case MonthNo
    Case 9 To 12
        ConvertMonthNoToServiceMonthNo = MonthNo - 8
    Case 1 To 8
        ConvertMonthNoToServiceMonthNo = MonthNo + 4
    End Select
    
   Exit Function

ErrorTrap:
    EndProgram
    
End Function
Public Function ConvertServiceMonthNoToMonthNo(ServiceMonthNo As Long) As Long
   On Error GoTo ErrorTrap

    Select Case ServiceMonthNo
    Case 5 To 12
        ConvertServiceMonthNoToMonthNo = ServiceMonthNo - 4
    Case 1 To 4
        ConvertServiceMonthNoToMonthNo = ServiceMonthNo + 8
    End Select
    
   Exit Function

ErrorTrap:
    EndProgram
    
End Function
Public Function ConvertServiceYearToNormalYear(TheDate As Date) As Long
   On Error GoTo ErrorTrap

    If Month(TheDate) >= 9 And Month(TheDate) <= 12 Then
        ConvertServiceYearToNormalYear = CLng(year(TheDate)) - 1
    Else
        ConvertServiceYearToNormalYear = CLng(year(TheDate))
    End If
    
   Exit Function

ErrorTrap:
    EndProgram
    
End Function
Public Function ConvertServiceDateToNormalDate(TheDate As Date) As Date
   On Error GoTo ErrorTrap
Dim TheMonth As Long, TheYear As String, TheDay As String
    
    ConvertServiceDateToNormalDate = CDate(CStr(Day(TheDate)) & "/" & _
                                           CStr(Month(TheDate)) & "/" & _
                                           CStr(ConvertServiceYearToNormalYear(TheDate)))
    
    
   Exit Function

ErrorTrap:
    EndProgram
    
End Function
Public Function ConvertNormalDateToServiceDate(TheDate As Date) As Date
   On Error GoTo ErrorTrap
Dim TheMonth As Long, TheYear As String, TheDay As String
    
    
    ConvertNormalDateToServiceDate = CDate(CStr(Day(TheDate)) & "/" & _
                                           CStr(Month(TheDate)) & "/" & _
                                           CStr(ServiceYear(TheDate)))
    
    
   Exit Function

ErrorTrap:
    EndProgram
    
End Function

Public Function GetTextWidthInTwips(TheText As String, TheForm As Form) As Long
   On Error GoTo ErrorTrap

    '
    'Form's units must be in twips. Form font attributes must be same as
    ' any control that you're comparing TextWidth with.
    '
    GetTextWidthInTwips = TheForm.TextWidth(TheText)
    
    
    
   Exit Function

ErrorTrap:
    EndProgram
    
End Function

Public Function FindPortionOfStringWithinSpecifiedTwips(TheText As String, _
                                                          MaxLengthTwips As Long, _
                                                          TheForm As Form) As Long
   Dim i As Long, Matched As Boolean
   On Error GoTo ErrorTrap

    '
    'Returns position in string TheText beyond which the actual length in twips
    ' would be greater than MaxLengthTwips
    '
    ' eg: If TheText is "abcdefghijklmnop", MaxLengthTwips is 1000, and TheForm has
    '     font attributes identical to that of the control from which TheText is
    '     derived, and TheForm's scalemode is twips, then this function would
    '     return, say, 9. ie width of string "abcdefghi" is <= 1000 twips
    '
    'Form's units must be in twips. Form font attributes must be same as
    ' any control that you're comparing TextWidth with.
    '
    
    For i = 1 To Len(TheText)
        If TheForm.TextWidth(Left(TheText, i)) > MaxLengthTwips Then
            FindPortionOfStringWithinSpecifiedTwips = i - 1
            Matched = True
            Exit Function
        End If
    Next i
    
    If Not Matched Then
        FindPortionOfStringWithinSpecifiedTwips = Len(TheText)
    End If
    
   Exit Function

ErrorTrap:
    EndProgram
    
End Function

Public Function NoDaysInMonth(ByVal TheMonth As Long, ByVal TheYear As Long) As Long
   On Error GoTo ErrorTrap
    
    Select Case TheMonth
    Case 1
        NoDaysInMonth = 31
    Case 2
        If IsLeapYear(CInt(TheYear)) Then
            NoDaysInMonth = 29
        Else
            NoDaysInMonth = 28
        End If
    Case 3
        NoDaysInMonth = 31
    Case 4
        NoDaysInMonth = 30
    Case 5
        NoDaysInMonth = 31
    Case 6
        NoDaysInMonth = 30
    Case 7
        NoDaysInMonth = 31
    Case 8
        NoDaysInMonth = 31
    Case 9
        NoDaysInMonth = 30
    Case 10
        NoDaysInMonth = 31
    Case 11
        NoDaysInMonth = 30
    Case 12
        NoDaysInMonth = 31
    End Select
        
   Exit Function

ErrorTrap:
    EndProgram
    
End Function

Public Function LastDateOfMonth(ByVal TheDate_UK As Date) As Date
   On Error GoTo ErrorTrap
    
    LastDateOfMonth = DateSerial(year(TheDate_UK), Month(TheDate_UK), _
                        CInt(NoDaysInMonth(Month(TheDate_UK), year(TheDate_UK))))
        
   Exit Function

ErrorTrap:
    EndProgram
    
End Function


Public Sub AddPublicSpeakerOutline(SpeakerID As Long, SpeakerID_B As Long, TalkNo As Long, SourceForm As Form)
Dim bOK As Boolean, rs As Recordset
On Error GoTo ErrorTrap
    
    If SpeakerID_B <> 0 Then Exit Sub 'don't bother for symposiums
    
    If SpeakerID = 0 Or TalkNo = 0 Then Exit Sub
    
    Set rs = GetGeneralRecordset("SELECT 1 " & _
                                " FROM tblSpeakersTalks " & _
                                " WHERE PersonID = " & SpeakerID & _
                                " AND TalkNo = " & TalkNo)
    
    If GlobalParms.GetValue("AddPublicSpeakerOutline", "AlphaVal") = "ASK" Then
        If rs.BOF Then
            If MsgBox("Add talk " & TalkNo & " ('" & GetTalkTitle(TalkNo) & "') to " & _
                    AddApostropheToPersonName(CongregationMember.FirstAndLastName(SpeakerID)) & _
                    " outline list?", vbYesNo + vbQuestion, AppName) = vbYes Then
                    
                bOK = True
            Else
                bOK = False
            End If
        Else
            bOK = False
        End If
    ElseIf GlobalParms.GetValue("AddPublicSpeakerOutline", "AlphaVal") = "AUTO" Then
        bOK = True
    Else
        bOK = False
    End If
    
    If bOK Then
        If rs.BOF Then
            Set rs = GetGeneralRecordset("tblSpeakersTalks")
            rs.AddNew
            rs!TalkNo = TalkNo
            rs!PersonID = SpeakerID
            rs.Update
            ShowMessage "Talk " & TalkNo & " added to " & _
                        AddApostropheToPersonName(CongregationMember.FirstAndLastName(SpeakerID)) & _
                        " outline list", 1500, SourceForm
        End If
    End If
    
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub


