Attribute VB_Name = "basYetMoreTMSCode"
Option Explicit

Public Function TMSScheduleMoveValid() As Boolean
On Error GoTo ErrorTrap

Dim ThePerson As Long, TheOtherPerson As Long, bNewMtg_Dest As Boolean, bNewMtg_Curr As Boolean
Dim lPers As Long, lTask As Long


Dim lCurrentStudent_New As Long
Dim lDestStudent_New As Long
Dim lCurrentAsst_New As Long
Dim lDestAsst_New As Long

Dim lCurrentStudent_Existing As Long
Dim lDestStudent_Existing As Long
Dim lCurrentAsst_Existing As Long
Dim lDestAsst_Existing As Long

Dim lStudent As Long, lAssistant As Long

Dim sTmp As String


    Screen.MousePointer = vbHourglass
    
    'some general validation first
    If Not TMSMoveGeneralValidation Then
        Screen.MousePointer = vbNormal
        TMSScheduleMoveValid = False
        Exit Function
    End If
    
    With frmTMSScheduling
       
    If .SwapStudents And (.CurrentAssistant > 0 Or .DestAssistant > 0) And _
        (.StudentSelected And .DestStudentSelected) Then
        
        'only allow auto-swap of assistants when current and dest selected are both _students_
    
        If MsgBox("Do you want to swap assistants too?", vbYesNo + vbQuestion, AppName) = vbYes Then
            .SwapAssistants = True
        Else
            .SwapAssistants = False
        End If
        
    Else
    
        .SwapAssistants = False
    
    End If
    
    
    If .StudentSelected And .DestStudentSelected And (Not .SwapAssistants) Then
        
        If .DestAssistant <> 0 Then
            sTmp = " This would overwrite " & _
                        CongregationMember.FullName(.DestAssistant) & " on " & .DestAssignmentDate & "."
        Else
            sTmp = ""
        End If
        
        If .CurrentAssistant <> 0 Then
        
            If MsgBox("Do you want to move the assistant?" & sTmp, vbYesNo + vbQuestion, AppName) = vbYes Then
    
                .MoveAssistants = True
            
            Else
            
                .MoveAssistants = False
                
            End If
        
        Else
        
            .MoveAssistants = False
            
        End If
        
    Else
    
        .MoveAssistants = True
        
    End If
    
                        
    
    lCurrentAsst_New = 0
    lCurrentStudent_New = 0
    lDestAsst_New = 0
    lDestStudent_New = 0
    
    
    
    'so this is the proposed schedule:
           
    If .StudentSelected Then
    
        If .DestStudentSelected Then
        
            lDestStudent_New = .CurrentStudent
            
            If .SwapStudents Then
                lCurrentStudent_New = .DestStudent
            Else
                lCurrentStudent_New = 0
            End If
                                      
            
        Else
        
            'Dest Asst selected. Moving current student to dest asst slot
            
            lDestAsst_New = .CurrentStudent
            
            If .SwapStudents Then
                lCurrentStudent_New = .DestAssistant
            Else
                lCurrentStudent_New = 0
            End If
            
        End If
    
    End If
    
    If .AssistantSelected Then
    
        If .DestStudentSelected Then
        
            lDestStudent_New = .CurrentAssistant
            
            If .SwapStudents Then
                lCurrentAsst_New = .DestStudent
            Else
                lCurrentAsst_New = 0
            End If
            
        Else
        
            'Dest Asst selected. Moving current student to dest asst slot
            
            lDestAsst_New = .CurrentAssistant
            
            If .SwapStudents Then
                lCurrentAsst_New = .DestAssistant
            Else
                lCurrentAsst_New = 0
            End If
            
        End If
    
    End If
    
    If .SwapAssistants Then
        lCurrentAsst_New = .DestAssistant
    End If
    
    If .MoveAssistants Then
        lDestAsst_New = .CurrentAssistant
    End If
    
    'OK lets check each of the (up to) 4 individuals can be in the proposed slots.
    
    '**********************
    'Check new dest student
    '**********************
    If lDestStudent_New > 0 Then
    
        'assigned to this talk-type?
        If Not PersonCanDoThisTalk(lDestStudent_New, .DestTalkNum, CStr(.DestAssignmentDate), .DestSchool) Then
            TMSScheduleMoveValid = False
            Screen.MousePointer = vbNormal
            Exit Function
        End If
        
        
        'suspended?
        lTask = CongregationMember.TMS_GetTaskForTMSTalkNo(.DestTalkNum)
        If CongregationMember.PersonIsSuspended(lDestStudent_New, .DestAssignmentDate, 4, 6, lTask) Then
                MsgBox CongregationMember.NameWithMiddleInitial(lDestStudent_New) & _
                        " has been suspended for " & .DestAssignmentDate & _
                        ". The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
        End If
        
        '
        'Can person do sourceless talk?
        '
        If TheTMS.IsItSourceless(.DestAssignmentDate, .DestTalkNum) And _
            Not CongregationMember.TMS_DoesSourcelessTalks(lDestStudent_New) Then
            
            If MsgBox(CongregationMember.NameWithMiddleInitial(lDestStudent_New) & _
                      " doesn't handle assignments with no source material. Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                Screen.MousePointer = vbNormal
                TMSScheduleMoveValid = False
                Exit Function
            End If
            
        End If
        
        '
        'Can person do talk with this assistant?
        '
        lStudent = lDestStudent_New 'already know lDestStudent_New is > 0
        If lDestAsst_New > 0 Then
            'we have a new dest asst
            lAssistant = lDestAsst_New
        Else
            'existing asst
            lAssistant = .DestAssistant
        End If
        If Not TheTMS.StudentAndAssistantValid(lStudent, .DestAssignmentDate, lAssistant, .DestSchool, .DestTalkNum, .DestItemsSeqNum) Then
            TMSScheduleMoveValid = False
            Screen.MousePointer = vbNormal
            Exit Function
        End If
        
        '
        'How many times is this person on this week already?
        '
        If CongregationMember.TMS_NoTimesOnSchoolThisWeek(lDestStudent_New, _
                                                          .DestAssignmentDate) > 0 _
            And _
           .DestAssignmentDate <> .CurrentAssignmentDate Then
                       
            With frmTMSMoveHelp
            .cmdDeleteDest.Visible = False
            .cmdSwap.Visible = False
            .cmdCancel.Visible = True
            .cmdOK.Visible = True
            .cmdDeleteAssistant.Visible = False
            .cmdLeaveAssistant.Visible = False
            .lblExplanation.Caption = CongregationMember.NameWithMiddleInitial(lDestStudent_New) & _
                                      " already has an assignment on " & _
                                      Format(frmTMSScheduling.DestAssignmentDate, "dd/mm/yyyy") & _
                                      ". What do you want to do?"
            .Show vbModal, frmTMSScheduling
            End With
            
            If .InMoveMode Then
            Else
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        End If

        '
        'Check if move would put the person on the school on the same week as a family member
        '
        If TheTMS.OnSchoolWithFamilyMemberThisWeek(lDestStudent_New, .DestAssignmentDate, lDestAsst_New, .DestSchool, .DestTalkNum) Then
            If MsgBox(CongregationMember.NameWithMiddleInitial(lDestStudent_New) & " has a family member also with a " & _
                        "school assignment on week " & Format(.DestAssignmentDate, "dd/mm/yyyy") & ". " & _
                        "Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                Screen.MousePointer = vbNormal
                TMSScheduleMoveValid = False
                Exit Function
            End If
            
        End If

        '
        'Check against Service Meeting and Cong Bible Study schedules...
        '
        If ServiceMtgConflict(lDestStudent_New, .DestAssignmentDate) Then
            TMSScheduleMoveValid = False
            Screen.MousePointer = vbNormal
            Exit Function
        End If
        
        If CongregationMember.CongBibleStudyConductorThisWeek(lDestStudent_New, .DestAssignmentDate) Then
            If MsgBox(CongregationMember.NameWithMiddleInitial(lDestStudent_New) & _
                      " is Congregation Bible Study Conductor on week " & _
                      Format(.DestAssignmentDate, "dd/mm/yyyy") & _
                      ". Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        End If
        
         If TheTMS.SchoolsBetweenDates(.DestAssignmentDate, .DestAssignmentDate) > 1 _
            And CongregationMember.IsSchoolAssistant(lDestStudent_New) Then
            If MsgBox(CongregationMember.NameWithMiddleInitial(.DestStudent) & _
                      " is a School assistant, and there is more than one school class for " & frmTMSScheduling.CurrentAssignmentDate & ". Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        End If
    
        
    End If
    
  
    '**********************
    'Check new dest assistant
    '**********************
    If lDestAsst_New > 0 Then
    
        'assigned to this talk-type?
        If Not PersonCanDoThisTalk(lDestAsst_New, "Asst", CStr(.DestAssignmentDate), .DestSchool) Then
            TMSScheduleMoveValid = False
            Screen.MousePointer = vbNormal
            Exit Function
        End If
        
        
        'suspended?
        lTask = CongregationMember.TMS_GetTaskForTMSTalkNo("Asst")
        If CongregationMember.PersonIsSuspended(lDestAsst_New, .DestAssignmentDate, 4, 6, lTask) Then
                MsgBox CongregationMember.NameWithMiddleInitial(lDestAsst_New) & _
                        " has been suspended for " & .DestAssignmentDate & _
                        ". The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
        End If
        
        '
        'Can assistant do talk with this student?
        '
        lAssistant = lDestAsst_New 'already know lDestAsst_New is > 0
        If lDestStudent_New > 0 Then
            'we have a new dest student
            lStudent = lDestStudent_New
        Else
            'existing student
            lStudent = .DestStudent
        End If
        If Not TheTMS.StudentAndAssistantValid(lStudent, .DestAssignmentDate, lAssistant, .DestSchool, .DestTalkNum, .DestItemsSeqNum) Then
            TMSScheduleMoveValid = False
            Screen.MousePointer = vbNormal
            Exit Function
        End If

               
        '
        'How many times is this person on this week already?
        '
        If CongregationMember.TMS_NoTimesOnSchoolThisWeek(lDestAsst_New, _
                                                          .DestAssignmentDate) > 0 _
            And _
           .DestAssignmentDate <> .CurrentAssignmentDate Then
                       
            With frmTMSMoveHelp
            .cmdDeleteDest.Visible = False
            .cmdSwap.Visible = False
            .cmdCancel.Visible = True
            .cmdOK.Visible = True
            .cmdDeleteAssistant.Visible = False
            .cmdLeaveAssistant.Visible = False
            .lblExplanation.Caption = CongregationMember.NameWithMiddleInitial(lDestAsst_New) & _
                                      " already has an assignment on " & _
                                      Format(frmTMSScheduling.DestAssignmentDate, "dd/mm/yyyy") & _
                                      ". What do you want to do?"
            .Show vbModal, frmTMSScheduling
            End With
            
            If .InMoveMode Then
            Else
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        End If

        '
        'Check if move would put the person on the school on the same week as a family member
        ' Only check if there was no dest student selected - will already have been checked there
        '
        If lDestStudent_New = 0 Then
            If TheTMS.OnSchoolWithFamilyMemberThisWeek(lDestAsst_New, .DestAssignmentDate, lDestStudent_New, .DestSchool, .DestTalkNum) Then
                If MsgBox(CongregationMember.NameWithMiddleInitial(lDestAsst_New) & " has a family member also with a " & _
                            "school assignment on week " & Format(.DestAssignmentDate, "dd/mm/yyyy") & ". " & _
                            "Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                    Screen.MousePointer = vbNormal
                    TMSScheduleMoveValid = False
                    Exit Function
                End If
            End If
        End If

        '
        'Check against Service Meeting and Cong Bible Study schedules...
        '
        If ServiceMtgConflict(lDestAsst_New, .DestAssignmentDate) Then
            TMSScheduleMoveValid = False
            Screen.MousePointer = vbNormal
            Exit Function
        End If
        
        If CongregationMember.CongBibleStudyConductorThisWeek(lDestAsst_New, .DestAssignmentDate) Then
            If MsgBox(CongregationMember.NameWithMiddleInitial(lDestAsst_New) & _
                      " is Congregation Bible Study Conductor on week " & _
                      Format(.DestAssignmentDate, "dd/mm/yyyy") & _
                      ". Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        End If
        
         If TheTMS.SchoolsBetweenDates(.DestAssignmentDate, .DestAssignmentDate) > 1 _
            And CongregationMember.IsSchoolAssistant(lDestAsst_New) Then
            If MsgBox(CongregationMember.NameWithMiddleInitial(lDestAsst_New) & _
                      " is a School assistant, and there is more than one school class for " & frmTMSScheduling.CurrentAssignmentDate & ". Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        End If
    
        
    End If

    
    '**********************
    'Check new current student
    '**********************
    If lCurrentStudent_New > 0 Then
    
        'assigned to this talk-type?
        If Not PersonCanDoThisTalk(lCurrentStudent_New, .CurrentTalkNum, CStr(.CurrentAssignmentDate), .CurrentSchool) Then
            TMSScheduleMoveValid = False
            Screen.MousePointer = vbNormal
            Exit Function
        End If
        
        
        'suspended?
        lTask = CongregationMember.TMS_GetTaskForTMSTalkNo(.CurrentTalkNum)
        If CongregationMember.PersonIsSuspended(lCurrentStudent_New, .CurrentAssignmentDate, 4, 6, lTask) Then
                MsgBox CongregationMember.NameWithMiddleInitial(lCurrentStudent_New) & _
                        " has been suspended for " & .CurrentAssignmentDate & _
                        ". The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
        End If
        
        '
        'Can person do sourceless talk?
        '
        If TheTMS.IsItSourceless(.CurrentAssignmentDate, .CurrentTalkNum) And _
            Not CongregationMember.TMS_DoesSourcelessTalks(lCurrentStudent_New) Then
            
            If MsgBox(CongregationMember.NameWithMiddleInitial(lCurrentStudent_New) & _
                      " doesn't handle assignments with no source material. Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                Screen.MousePointer = vbNormal
                TMSScheduleMoveValid = False
                Exit Function
            End If
            
        End If
        
        '
        'Can person do talk with this assistant?
        '
        lAssistant = lCurrentAsst_New 'already know lCurrentAsst_New is > 0
        If lCurrentStudent_New > 0 Then
            'we have a new current student
            lStudent = lCurrentStudent_New
        Else
            'existing student
            lStudent = .CurrentStudent
        End If
        
        If Not TheTMS.StudentAndAssistantValid(lStudent, .CurrentAssignmentDate, lAssistant, .CurrentSchool, .CurrentTalkNum, .currentItemsSeqNum) Then
            TMSScheduleMoveValid = False
            Screen.MousePointer = vbNormal
            Exit Function
        End If
        
        '
        'How many times is this person on this week already?
        '
        If CongregationMember.TMS_NoTimesOnSchoolThisWeek(lCurrentStudent_New, _
                                                          .CurrentAssignmentDate) > 0 _
            And _
           .CurrentAssignmentDate <> .DestAssignmentDate Then
                       
            With frmTMSMoveHelp
            .cmdDeleteDest.Visible = False
            .cmdSwap.Visible = False
            .cmdCancel.Visible = True
            .cmdOK.Visible = True
            .cmdDeleteAssistant.Visible = False
            .cmdLeaveAssistant.Visible = False
            .lblExplanation.Caption = CongregationMember.NameWithMiddleInitial(lCurrentStudent_New) & _
                                      " already has an assignment on " & _
                                      Format(frmTMSScheduling.CurrentAssignmentDate, "dd/mm/yyyy") & _
                                      ". What do you want to do?"
            .Show vbModal, frmTMSScheduling
            End With
            
            If .InMoveMode Then
            Else
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        End If

        '
        'Check if move would put the person on the school on the same week as a family member
        '
        If TheTMS.OnSchoolWithFamilyMemberThisWeek(lCurrentStudent_New, .CurrentAssignmentDate, lCurrentAsst_New, .CurrentSchool, .CurrentTalkNum) Then
            If MsgBox(CongregationMember.NameWithMiddleInitial(lCurrentStudent_New) & " has a family member also with a " & _
                        "school assignment on week " & Format(.CurrentAssignmentDate, "dd/mm/yyyy") & ". " & _
                        "Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                Screen.MousePointer = vbNormal
                TMSScheduleMoveValid = False
                Exit Function
            End If
            
        End If

        '
        'Check against Service Meeting and Cong Bible Study schedules...
        '
        If ServiceMtgConflict(lCurrentStudent_New, .CurrentAssignmentDate) Then
            TMSScheduleMoveValid = False
            Screen.MousePointer = vbNormal
            Exit Function
        End If
        
        If CongregationMember.CongBibleStudyConductorThisWeek(lCurrentStudent_New, .CurrentAssignmentDate) Then
            If MsgBox(CongregationMember.NameWithMiddleInitial(lCurrentStudent_New) & _
                      " is Congregation Bible Study Conductor on week " & _
                      Format(.CurrentAssignmentDate, "dd/mm/yyyy") & _
                      ". Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        End If
        
         If TheTMS.SchoolsBetweenDates(.CurrentAssignmentDate, .CurrentAssignmentDate) > 1 _
            And CongregationMember.IsSchoolAssistant(lCurrentStudent_New) Then
            If MsgBox(CongregationMember.NameWithMiddleInitial(.CurrentStudent) & _
                      " is a School assistant, and there is more than one school class for " & frmTMSScheduling.CurrentAssignmentDate & ". Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        End If
    
        
    End If
        
        
    '**********************
    'Check new current assistant
    '**********************
    If lCurrentAsst_New > 0 Then
    
        'assigned to this talk-type?
        If Not PersonCanDoThisTalk(lCurrentAsst_New, "Asst", CStr(.CurrentAssignmentDate), .CurrentSchool) Then
            TMSScheduleMoveValid = False
            Screen.MousePointer = vbNormal
            Exit Function
        End If
        
        
        'suspended?
        lTask = CongregationMember.TMS_GetTaskForTMSTalkNo("Asst")
        If CongregationMember.PersonIsSuspended(lCurrentAsst_New, .CurrentAssignmentDate, 4, 6, lTask) Then
                MsgBox CongregationMember.NameWithMiddleInitial(lCurrentAsst_New) & _
                        " has been suspended for " & .CurrentAssignmentDate & _
                        ". The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
        End If
        
               
        '
        'How many times is this person on this week already?
        '
        If CongregationMember.TMS_NoTimesOnSchoolThisWeek(lCurrentAsst_New, _
                                                          .CurrentAssignmentDate) > 0 _
            And _
           .CurrentAssignmentDate <> .CurrentAssignmentDate Then
                       
            With frmTMSMoveHelp
            .cmdDeleteDest.Visible = False
            .cmdSwap.Visible = False
            .cmdCancel.Visible = True
            .cmdOK.Visible = True
            .cmdDeleteAssistant.Visible = False
            .cmdLeaveAssistant.Visible = False
            .lblExplanation.Caption = CongregationMember.NameWithMiddleInitial(lCurrentAsst_New) & _
                                      " already has an assignment on " & _
                                      Format(frmTMSScheduling.CurrentAssignmentDate, "dd/mm/yyyy") & _
                                      ". What do you want to do?"
            .Show vbModal, frmTMSScheduling
            End With
            
            If .InMoveMode Then
            Else
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        End If

        '
        'Check if move would put the person on the school on the same week as a family member
        ' Only check if there was no current student selected - will already have been checked there
        '
        If lCurrentAsst_New = 0 Then
            If TheTMS.OnSchoolWithFamilyMemberThisWeek(lCurrentAsst_New, .CurrentAssignmentDate, lCurrentStudent_New, .CurrentSchool, .CurrentTalkNum) Then
                If MsgBox(CongregationMember.NameWithMiddleInitial(lCurrentAsst_New) & " has a family member also with a " & _
                            "school assignment on week " & Format(.CurrentAssignmentDate, "dd/mm/yyyy") & ". " & _
                            "Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                    Screen.MousePointer = vbNormal
                    TMSScheduleMoveValid = False
                    Exit Function
                End If
                
            End If
        End If

        '
        'Check against Service Meeting and Cong Bible Study schedules...
        '
        If ServiceMtgConflict(lCurrentAsst_New, .CurrentAssignmentDate) Then
            TMSScheduleMoveValid = False
            Screen.MousePointer = vbNormal
            Exit Function
        End If
        
        If CongregationMember.CongBibleStudyConductorThisWeek(lCurrentAsst_New, .CurrentAssignmentDate) Then
            If MsgBox(CongregationMember.NameWithMiddleInitial(lCurrentAsst_New) & _
                      " is Congregation Bible Study Conductor on week " & _
                      Format(.CurrentAssignmentDate, "dd/mm/yyyy") & _
                      ". Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        End If
        
         If TheTMS.SchoolsBetweenDates(.CurrentAssignmentDate, .CurrentAssignmentDate) > 1 _
            And CongregationMember.IsSchoolAssistant(lCurrentAsst_New) Then
            If MsgBox(CongregationMember.NameWithMiddleInitial(lCurrentAsst_New) & _
                      " is a School assistant, and there is more than one school class for " & frmTMSScheduling.CurrentAssignmentDate & ". Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                TMSScheduleMoveValid = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        End If
    
        
    End If



    TMSScheduleMoveValid = True
    Screen.MousePointer = vbNormal
    
    End With
    
    Exit Function
ErrorTrap:
    EndProgram
    

End Function

Private Function TMSMoveGeneralValidation() As Boolean
On Error GoTo ErrorTrap
Dim ThePerson As Long, TheOtherPerson As Long, bNewMtg_Dest As Boolean, bNewMtg_Curr As Boolean
Dim lPers As Long, lTask As Long

    Screen.MousePointer = vbHourglass
'
'Ensure that student/assistant moves are possible
' NOTE: If moving person from Slot A to B, slot A is CURRENT, Slot B is DESTINATION
'
    
    With frmTMSScheduling
    If .StudentSelected Then
        ThePerson = .CurrentStudent
    Else
        ThePerson = .CurrentAssistant
    End If
    
    If .DestStudentSelected Then
        TheOtherPerson = .DestStudent
    Else
        TheOtherPerson = .DestAssistant
    End If
    
    
    '
    'Any unexpected problem? Quit move! (THIS SHOULD BE THE FIRST VALIDATION)
    '
    If ThePerson = 0 Or _
       CDate(.CurrentAssignmentDate) < #1/1/2003# Or _
       .CurrentSchoolNo < 1 Or _
       .CurrentSchoolNo > CLng(GlobalParms.GetValue("NumberOfSchools", "NumVal")) Or _
       .CurrentTalkNum = "" Or _
       Not IsNumeric(.CurrentSchoolNo) Or _
       Not IsDate(.CurrentAssignmentDate) Or _
       CDate(.DestAssignmentDate) < #1/1/2003# Or _
       .DestSchool < 1 Or _
       .DestSchool > CLng(GlobalParms.GetValue("NumberOfSchools", "NumVal")) Or _
       .DestTalkNum = "" Or _
       Not IsNumeric(.DestSchool) Or _
       Not IsDate(.DestAssignmentDate) Then

        MsgBox "Unable to carry out the Move. Please check and try again." _
                , vbOKOnly + vbExclamation, AppName
        TMSMoveGeneralValidation = False
        Screen.MousePointer = vbNormal
        Exit Function
    End If
    
    'cater for different meeting formats (pre and post 2009)
    
    Select Case NewMtgArrangementStarted(CStr(.CurrentAssignmentDate))
    Case CLM2016
        Select Case .CurrentTalkNum
        Case "BR", "IC", "RV", "BS", "O"
        Case Else
            MsgBox "Unable to carry out the Move - Invalid Assignment Type '" & .CurrentTalkNum & "'. Please check and try again." _
                    , vbOKOnly + vbExclamation, AppName
            TMSMoveGeneralValidation = False
            Screen.MousePointer = vbNormal
            Exit Function
        End Select
    Case TMS2009
        Select Case .CurrentTalkNum
        Case "P", "S", "B", "1", "2", "3", "4", "MR", "R"
        Case Else
            MsgBox "Unable to carry out the Move - Invalid Talk Number. Please check and try again." _
                    , vbOKOnly + vbExclamation, AppName
            TMSMoveGeneralValidation = False
            Screen.MousePointer = vbNormal
            Exit Function
        End Select
    Case Else
        Select Case .CurrentTalkNum
        Case "P", "B", "1", "2", "3"
        Case Else
            MsgBox "Unable to carry out the Move - Invalid Talk Number. Please check and try again." _
                    , vbOKOnly + vbExclamation, AppName
            TMSMoveGeneralValidation = False
            Screen.MousePointer = vbNormal
            Exit Function
        End Select
    End Select
    
    Select Case NewMtgArrangementStarted(CStr(.DestAssignmentDate))
    Case CLM2016
        Select Case .DestTalkNum
        Case "BR", "IC", "RV", "BS", "O"
        Case Else
            MsgBox "Unable to carry out the Move - Invalid Assignment Type '" & .DestTalkNum & "'. Please check and try again." _
                    , vbOKOnly + vbExclamation, AppName
            TMSMoveGeneralValidation = False
            Screen.MousePointer = vbNormal
            Exit Function
        End Select
    Case TMS2009
        Select Case .DestTalkNum
        Case "P", "B", "1", "2", "3", "MR", "R"
        Case Else
            MsgBox "Unable to carry out the Move - Invalid Talk Number. Please check and try again." _
                    , vbOKOnly + vbExclamation, AppName
            TMSMoveGeneralValidation = False
            Screen.MousePointer = vbNormal
            Exit Function
        End Select
    Case Else
        Select Case .DestTalkNum
        Case "P", "S", "B", "1", "2", "3", "4"
        Case Else
            MsgBox "Unable to carry out the Move - Invalid Talk Number. Please check and try again." _
                    , vbOKOnly + vbExclamation, AppName
            TMSMoveGeneralValidation = False
            Screen.MousePointer = vbNormal
            Exit Function
        End Select
    End Select
    
    If TMSSchedulePrintedForDate(.CurrentAssignmentDate) Then
        If MsgBox("A schedule has been printed for " & .CurrentAssignmentDate & _
                    ". Continue with the move?", vbYesNo + vbQuestion, AppName) = vbNo Then
            TMSMoveGeneralValidation = False
            Screen.MousePointer = vbNormal
            Exit Function
        End If
    Else
        If TMSSchedulePrintedForDate(.DestAssignmentDate) Then
            If MsgBox("A schedule has been printed for " & .DestAssignmentDate & _
                        ". Continue with the move?", vbYesNo + vbQuestion, AppName) = vbNo Then
                TMSMoveGeneralValidation = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        End If
    End If
    
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    'END INITIAL VALIDATION
    '
       
    '
    'Trying to swap same person?
    '
    If ThePerson = TheOtherPerson Then
        MsgBox "You have selected the same person for both the source and destination." _
                , vbOKOnly + vbExclamation, AppName
        TMSMoveGeneralValidation = False
        Screen.MousePointer = vbNormal
        Exit Function
    End If
        
    '
    'Is Current Talk Completed?
    '
    If CongregationMember.TMSTalkIsComplete(.CurrentAssignmentDate, _
                                            .CurrentStudent, _
                                            .CurrentTalkNum, _
                                            .CurrentSchoolNo) Or _
       CongregationMember.TMSTalkIsDefaulted(.CurrentAssignmentDate, _
                                            .CurrentStudent, _
                                            .CurrentTalkNum, _
                                            .CurrentSchoolNo) Then
        
        If MsgBox("The assignment on " & .CurrentAssignmentDate & _
                    " in School " & .CurrentSchoolNo & _
                  " has been completed or defaulted. Do you want to proceed " & _
                  "with the move?", vbYesNo + vbQuestion, AppName) = vbNo Then
            
            TMSMoveGeneralValidation = False
            Screen.MousePointer = vbNormal
            Exit Function
        End If
    End If

    
    '
    'Is Destination Talk Completed?
    '
    If .CurrentAssignmentDate <> .DestAssignmentDate Or _
        .CurrentSchoolNo <> .DestSchool Then
    
        If CongregationMember.TMSTalkIsComplete(.DestAssignmentDate, _
                                                .DestStudent, _
                                                .DestTalkNum, _
                                                .DestSchool) Or _
           CongregationMember.TMSTalkIsDefaulted(.DestAssignmentDate, _
                                                 .DestStudent, _
                                                .DestTalkNum, _
                                                .DestSchool) Then
            If MsgBox("The assignment on " & .DestAssignmentDate & _
                    " in School " & .DestSchool & _
                      " has been completed or defaulted. Do you want to proceed " & _
                      "with the move?", vbYesNo + vbQuestion, AppName) = vbNo Then
                
                TMSMoveGeneralValidation = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
                
        End If
    End If
                                              
    '
    'Is Current Talk-slip printed?
    '
    If CongregationMember.TMSTalkSlipPrinted(.CurrentAssignmentDate, _
                                            .CurrentStudent, _
                                            .CurrentTalkNum, _
                                            .CurrentSchoolNo) Then
        
        If MsgBox("The talk-slip for the assignment on " & .CurrentAssignmentDate & _
                    " in School " & .CurrentSchoolNo & _
                  " has been printed. Do you want to proceed " & _
                  "with the move?", vbYesNo + vbQuestion, AppName) = vbNo Then
            
            TMSMoveGeneralValidation = False
            Screen.MousePointer = vbNormal
            Exit Function
        End If
    End If

    
    If .CurrentAssignmentDate <> .DestAssignmentDate Or _
        .CurrentSchoolNo <> .DestSchool Then
        '
        'Is destination Talk-slip printed?
        '
        If CongregationMember.TMSTalkSlipPrinted(.DestAssignmentDate, _
                                                 .DestStudent, _
                                                .DestTalkNum, _
                                                .DestSchool) Then
            
            If MsgBox("The talk-slip for the assignment on " & .DestAssignmentDate & _
                    " in School " & .DestSchool & _
                      " has been printed. Do you want to proceed " & _
                      "with the move?", vbYesNo + vbQuestion, AppName) = vbNo Then
                
                TMSMoveGeneralValidation = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        End If
    End If
    
    End With
    
    TMSMoveGeneralValidation = True
    Screen.MousePointer = vbNormal
    
    Exit Function
ErrorTrap:
    EndProgram
    

            
End Function




Public Function PersonCanDoThisTalk(ThePerson As Long, _
                                    TheTalkNo As String, _
                                    AssignmentDate As String, _
                                    SchoolNo As Long) As Boolean
On Error GoTo ErrorTrap

    Select Case NewMtgArrangementStarted(AssignmentDate)
    Case CLM2016
        PersonCanDoThisTalk = PersonCanDoThisTalk_2016(ThePerson, TheTalkNo, AssignmentDate, SchoolNo)
    Case TMS2009
        PersonCanDoThisTalk = PersonCanDoThisTalk_2009(ThePerson, TheTalkNo, AssignmentDate, SchoolNo)
    Case Else
       PersonCanDoThisTalk = PersonCanDoThisTalk_2002(ThePerson, TheTalkNo)
    End Select

    Exit Function
ErrorTrap:
    EndProgram
    
End Function
Public Function PersonCanDoThisTalk_2002(ThePerson As Long, TheTalkNo As String) As Boolean
On Error GoTo ErrorTrap

    Select Case TheTalkNo
    Case "P"
        If Not CongregationMember.DoesTMS_Prayer(ThePerson) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to Prayers. " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2002 = False
            Exit Function
        End If
    Case "S"
        If Not CongregationMember.DoesTMS_SQ(ThePerson) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to Speech Quality talks. " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2002 = False
            Exit Function
        End If
    Case "B"
        If Not CongregationMember.DoesTMS_BH(ThePerson) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to Bible Highlights. " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2002 = False
            Exit Function
        End If
    Case "1"
        If Not CongregationMember.DoesTMS_No1(ThePerson) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to No 1 talks. " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2002 = False
            Exit Function
        End If
    Case "2"
        If Not CongregationMember.DoesTMS_No2_BibleReading(ThePerson) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to Reading Assignments. " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2002 = False
            Exit Function
        End If
    Case "3"
        If Not CongregationMember.DoesTMS_No3_NoSourceMaterial(ThePerson) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to No 3 talks. " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2002 = False
            Exit Function
        End If
    Case "4"
        If Not CongregationMember.DoesTMS_No4_NoSourceMaterial(ThePerson) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to No 4 talks. " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2002 = False
            Exit Function
        End If
    End Select

    PersonCanDoThisTalk_2002 = True


    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Function PersonCanDoThisTalk_2016(ThePerson As Long, TheTalkNo As String, AssignmentDate As String, SchoolNo As Long) As Boolean
On Error GoTo ErrorTrap

    Select Case TheTalkNo
    Case "BR"
        If Not CongregationMember.DoesTMS_BibleReading_2016(ThePerson, SchoolNo) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to Bible Readings in School " & SchoolNo & ". " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2016 = False
            Exit Function
        End If
    Case "IC"
        If Not CongregationMember.DoesTMS_InitialCall_2016(ThePerson, SchoolNo) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to 'Initial Calls' in School " & SchoolNo & ". " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2016 = False
            Exit Function
        End If
    Case "RV"
        If Not CongregationMember.DoesTMS_RV_2016(ThePerson, SchoolNo) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to 'Return Visits' in School " & SchoolNo & ". " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2016 = False
            Exit Function
        End If
        
    Case "BS"
        If Not CongregationMember.DoesTMS_Study_2016(ThePerson, SchoolNo) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to 'Bible Studies' in School " & SchoolNo & ". " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2016 = False
            Exit Function
        End If
    Case "O"
        If Not CongregationMember.DoesTMS_Study_2016(ThePerson, SchoolNo) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to 'Other' in School " & SchoolNo & ". " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2016 = False
            Exit Function
        End If
        
    Case "Asst"
        If Not CongregationMember.DoesTMS_No3_Assistant(ThePerson) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to assistant in School " & SchoolNo & ". " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2016 = False
            Exit Function
        End If
    End Select

    PersonCanDoThisTalk_2016 = True


    Exit Function
ErrorTrap:
    EndProgram
    
End Function


Public Function PersonCanDoThisTalk_2009(ThePerson As Long, TheTalkNo As String, AssignmentDate As String, SchoolNo As Long) As Boolean
On Error GoTo ErrorTrap

    Select Case TheTalkNo
    Case "P"
        If Not CongregationMember.DoesTMS_Prayer(ThePerson) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to Prayers. " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2009 = False
            Exit Function
        End If
    Case "B"
        If Not CongregationMember.DoesTMS_BH(ThePerson) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to Bible Highlights in School " & SchoolNo & ". " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2009 = False
            Exit Function
        End If
    Case "1"
        If Not CongregationMember.DoesTMS_No1_2009(ThePerson) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to No 1 talks in School " & SchoolNo & ". " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2009 = False
            Exit Function
        End If
    Case "2"
        If CongregationMember.GetGender(ThePerson) = Female Then
            If Not CongregationMember.DoesTMS_No2_2009(ThePerson) Then
                MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                        " has not been assigned to No 2 talks in School " & SchoolNo & ". " & _
                        "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
                PersonCanDoThisTalk_2009 = False
                Exit Function
            End If
        Else
            'male...
            If CongregationMember.DoesTMS_No3_2009(ThePerson) Then
                
                If MsgBox("Assign " & CongregationMember.FirstAndLastName(ThePerson) & _
                            " to Talk No 2?", vbYesNo + vbQuestion, AppName) = vbNo Then
                    PersonCanDoThisTalk_2009 = False
                    ShowMessage "Operation cancelled", 1500, frmTMSScheduling
                    Exit Function
                End If
            
            Else
                
                MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                        " has not been assigned to No 3 talks in School " & SchoolNo & _
                        ", so cannot be assigned to No 2 talks. " & _
                        "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
                PersonCanDoThisTalk_2009 = False
                Exit Function
                
            End If
        End If
    Case "3"
        If Not CongregationMember.DoesTMS_No3_2009(ThePerson) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to No 3 talks in School " & SchoolNo & ". " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2009 = False
            Exit Function
        Else
            If CongregationMember.GetGender(ThePerson) = Female Then
                If TheTMS.IsItBroOnly(CDate(AssignmentDate)) Then
                    MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                            " cannot been assigned to a 'brothers only' part. " & SchoolNo & ". " & _
                            "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
                    PersonCanDoThisTalk_2009 = False
                    Exit Function
                End If
            
            End If
            
        End If
    Case "R", "MR"
        If Not CongregationMember.DoesTMS_OralReviewReading(ThePerson) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to read the scriptures in the Oral Review. " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2009 = False
            Exit Function
        End If
        
    Case "Asst"
        If Not CongregationMember.DoesTMS_No3_Assistant(ThePerson) Then
            MsgBox CongregationMember.NameWithMiddleInitial(ThePerson) & _
                    " has not been assigned to assistant in School " & SchoolNo & ". " & _
                    "The Move has been cancelled.", vbOKOnly + vbExclamation, AppName
            PersonCanDoThisTalk_2009 = False
            Exit Function
        End If
    End Select

    PersonCanDoThisTalk_2009 = True


    Exit Function
ErrorTrap:
    EndProgram
    
End Function


Public Sub DoTMSScheduleMove()
On Error GoTo ErrorTrap

    Dim lCurrentSQ As Long, dtCurrentSQDate As Date
    Dim lDestSQ As Long, dtDestSQDate As Date
    

    Screen.MousePointer = vbHourglass
'
'
'
    With frmTMSScheduling
    
    If Not .InMoveMode Then Exit Sub
    
    'Get SQs for current and destination slots...
    lCurrentSQ = TheTMS.GetTMSCounselPoint(.CurrentAssignmentDate, .CurrentTalkNum, .CurrentSchoolNo, .currentItemsSeqNum)
    dtCurrentSQDate = TheTMS.GetTMSCounselAssignedDate(.CurrentAssignmentDate, .CurrentTalkNum, .CurrentSchoolNo, .currentItemsSeqNum)
    lDestSQ = TheTMS.GetTMSCounselPoint(.DestAssignmentDate, .DestTalkNum, .DestSchool, .DestItemsSeqNum)
    dtDestSQDate = TheTMS.GetTMSCounselAssignedDate(.DestAssignmentDate, .DestTalkNum, .DestSchool, .DestItemsSeqNum)
                                                   
    
    If .StudentSelected Then
        '
        'Current STUDENT selected, so add to Destination then delete from
        ' Current slot.
        '
        CongregationMember.AddStudentToTMSSchedule 0, _
                                                   .CurrentAssignmentDate, _
                                                   .CurrentTalkNum, _
                                                   .CurrentSchoolNo, .currentItemsSeqNum, _
                                                   True, True, True
                                                   
        If .DestStudentSelected Then
            '
            'Move Current Student to Dest Student slot
            '
            CongregationMember.AddStudentToTMSSchedule .CurrentStudent, _
                                                       .DestAssignmentDate, _
                                                       .DestTalkNum, _
                                                       .DestSchool, .DestItemsSeqNum, _
                                                       True, True, True, , , _
                                                        CStr(lCurrentSQ), dtCurrentSQDate
                                                       
'            AlertIfNoCounselPointSet .CurrentStudent, .DestAssignmentDate, .DestTalkNum, .DestSchool
                                                       
        Else
            '
            'Move Current Student to Dest Assistant slot
            '
            CongregationMember.AddAssistantToTMSSchedule .CurrentStudent, _
                                                         .DestAssignmentDate, _
                                                         .DestTalkNum, _
                                                         .DestSchool, .DestItemsSeqNum, True, True
        End If
                                                   
    Else
        '
        'Current ASSISTANT selected, so add to Destination then delete from
        ' Current slot.
        '
        CongregationMember.AddAssistantToTMSSchedule 0, _
                                                     .CurrentAssignmentDate, _
                                                     .CurrentTalkNum, _
                                                     .CurrentSchoolNo, .currentItemsSeqNum, True, True
        If .DestStudentSelected Then
            '
            'Move Current Assistant to Dest Student slot
            '
            CongregationMember.AddStudentToTMSSchedule .CurrentAssistant, _
                                                       .DestAssignmentDate, _
                                                       .DestTalkNum, _
                                                       .DestSchool, .DestItemsSeqNum, _
                                                       True, True, True
'            AlertIfNoCounselPointSet .CurrentAssistant, .DestAssignmentDate, .DestTalkNum, .DestSchool
        
        Else
            '
            'Move Current Assistant to Dest Assistant slot
            '
            CongregationMember.AddAssistantToTMSSchedule .CurrentAssistant, _
                                                         .DestAssignmentDate, _
                                                         .DestTalkNum, _
                                                         .DestSchool, .DestItemsSeqNum, True, True
        End If
        
    End If
    
    If .SwapStudents Then
        '
        'Now move Dest person to Current slot...
        '
                
        If .DestStudentSelected Then
            '
            'Dest STUDENT selected, so add to Current slot.
            '
            If .StudentSelected Then
                '
                'Move Dest Student to Current Student slot
                '
                CongregationMember.AddStudentToTMSSchedule .DestStudent, _
                                                           .CurrentAssignmentDate, _
                                                           .CurrentTalkNum, _
                                                           .CurrentSchoolNo, .currentItemsSeqNum, _
                                                           True, True, True, , , _
                                                            CStr(lDestSQ), dtDestSQDate
                
                
            Else
                '
                'Move Dest Student to Current Assistant slot
                '
                CongregationMember.AddAssistantToTMSSchedule .DestStudent, _
                                                             .CurrentAssignmentDate, _
                                                             .CurrentTalkNum, _
                                                             .CurrentSchoolNo, .currentItemsSeqNum, True, True
            End If
                                                       
        Else
            '
            'Dest ASSISTANT selected, so add to Current.
            '
            If .StudentSelected Then
                '
                'Move Dest Assistant to Current Student slot
                '
                CongregationMember.AddStudentToTMSSchedule .DestAssistant, _
                                                           .CurrentAssignmentDate, _
                                                           .CurrentTalkNum, _
                                                           .CurrentSchoolNo, .currentItemsSeqNum, _
                                                           True, True, True
            Else
                '
                'Move Dest Assistant to Current Assistant slot
                '
                CongregationMember.AddAssistantToTMSSchedule .DestAssistant, _
                                                             .CurrentAssignmentDate, _
                                                             .CurrentTalkNum, _
                                                             .CurrentSchoolNo, .currentItemsSeqNum, True, True
            End If
            
        End If
        
    End If
    
    
    '
    'Has user elected to delete assistant? This is where moving a brother to a slot
    ' where there's an assistant....
    '
    If .DeleteCurrentAssistant Then
        CongregationMember.AddAssistantToTMSSchedule 0, _
                                                     .CurrentAssignmentDate, _
                                                     .CurrentTalkNum, _
                                                     .CurrentSchoolNo, .currentItemsSeqNum, True, True
    End If
    
    If .DeleteDestAssistant Then
        CongregationMember.AddAssistantToTMSSchedule 0, _
                                                     .DestAssignmentDate, _
                                                     .DestTalkNum, _
                                                     .DestSchool, .DestItemsSeqNum, True, True
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Auto-move assistant(s) as well... if selected earlier...

    If .StudentSelected And .DestStudentSelected Then
    
        
        If .MoveAssistants Then
        
            .AssistantSelected = True
            .DestAssistantSelected = True
            .StudentSelected = False
            .DestStudentSelected = False
            
                
            CongregationMember.AddAssistantToTMSSchedule 0, _
                                                         .CurrentAssignmentDate, _
                                                         .CurrentTalkNum, _
                                                         .CurrentSchoolNo, .currentItemsSeqNum, True, True
            
            CongregationMember.AddAssistantToTMSSchedule .CurrentAssistant, _
                                                         .DestAssignmentDate, _
                                                         .DestTalkNum, _
                                                         .DestSchool, .DestItemsSeqNum, True, True
                
                
            
        End If
           
    
    
        If .SwapAssistants Then
        
            .AssistantSelected = True
            .DestAssistantSelected = True
            .StudentSelected = False
            .DestStudentSelected = False
            
               
            CongregationMember.AddAssistantToTMSSchedule .DestAssistant, _
                                                         .CurrentAssignmentDate, _
                                                         .CurrentTalkNum, _
                                                         .CurrentSchoolNo, .currentItemsSeqNum, True, True
            
            CongregationMember.AddAssistantToTMSSchedule .CurrentAssistant, _
                                                         .DestAssignmentDate, _
                                                         .DestTalkNum, _
                                                         .DestSchool, .DestItemsSeqNum, True, True
            
        
        End If
            
        
        
    End If
    
    'reset these to default
    .AssistantSelected = False
    .DestAssistantSelected = False
    .StudentSelected = True
    .DestStudentSelected = True

    
       
    End With

    Screen.MousePointer = vbNormal

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub AlertIfNoCounselPointSet(ThePerson As Long, _
                                     AssignmentDate As Date, _
                                     TalkNo As String, _
                                     School As Long, _
                                     ItemsSeqNum As Long)

On Error GoTo ErrorTrap

Dim str As String
    
    'if no counsel point selected and there should be one, user may want to select one now...
    Select Case NewMtgArrangementStarted(CStr(AssignmentDate))
    Case CLM2016
        If TheTMS.GetTMSCounselPoint(AssignmentDate, TalkNo, School, ItemsSeqNum) <= 0 Then
            MsgBox "There is no counsel point set for " & _
                    AddApostropheToPersonName(CongregationMember.NameWithMiddleInitial(ThePerson)) & _
                    " assignment.", vbOKOnly + vbInformation, AppName
        End If
    Case TMS2009
        Select Case TalkNo
        Case "B"
            If CongregationMember.TMS_AllowCounselOnNo1AndBH(ThePerson) Then
                If TheTMS.GetTMSCounselPoint(AssignmentDate, TalkNo, School, ItemsSeqNum) <= 0 Then
                    MsgBox "There is no counsel point set for " & _
                            AddApostropheToPersonName(CongregationMember.NameWithMiddleInitial(ThePerson)) & _
                            " assignment.", vbOKOnly + vbInformation, AppName
                End If
            End If
        Case "3" And (CongregationMember.ElderDate(ThePerson) > 0 Or CongregationMember.ServantDate(ThePerson) > 0)
            If CongregationMember.TMS_AllowCounselOnNo1AndBH(ThePerson) Then
                If TheTMS.GetTMSCounselPoint(AssignmentDate, TalkNo, School, ItemsSeqNum) <= 0 Then
                    MsgBox "There is no counsel point set for " & _
                            AddApostropheToPersonName(CongregationMember.NameWithMiddleInitial(ThePerson)) & _
                            " assignment.", vbOKOnly + vbInformation, AppName
                End If
            End If
        Case "1", "3", "4"
            If TheTMS.GetTMSCounselPoint(AssignmentDate, TalkNo, School, ItemsSeqNum) <= 0 Then
                MsgBox "There is no counsel point set for " & _
                        AddApostropheToPersonName(CongregationMember.NameWithMiddleInitial(ThePerson)) & _
                        " assignment.", vbOKOnly + vbInformation, AppName
            End If
        Case Else
            'do nothing
        End Select
    Case Else
        Select Case TalkNo
        Case "P", "S"
        Case "1", "B"
            If CongregationMember.TMS_AllowCounselOnNo1AndBH(ThePerson) Then
                If TheTMS.GetTMSCounselPoint(AssignmentDate, TalkNo, School, ItemsSeqNum) <= 0 Then
                    MsgBox "There is no counsel point set for " & _
                            AddApostropheToPersonName(CongregationMember.NameWithMiddleInitial(ThePerson)) & _
                            " assignment.", vbOKOnly + vbInformation, AppName
                End If
            End If
        Case "2", "3", "4"
            If TheTMS.GetTMSCounselPoint(AssignmentDate, TalkNo, School, ItemsSeqNum) <= 0 Then
                MsgBox "There is no counsel point set for " & _
                        AddApostropheToPersonName(CongregationMember.NameWithMiddleInitial(ThePerson)) & _
                        " assignment.", vbOKOnly + vbInformation, AppName
            End If
        Case Else
            'do nothing
        End Select
    End Select

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Public Function TMSSchedulePrintedForDate(AssignmentDate As Date) As Boolean
On Error GoTo ErrorTrap
    
    TMSSchedulePrintedForDate = (AssignmentDate < CDate(GlobalParms.GetValue("NextTMSSchedulePrintStartDate", "DateVal", 0)))

    Exit Function
ErrorTrap:
    EndProgram
    
End Function

