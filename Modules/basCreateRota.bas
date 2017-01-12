Attribute VB_Name = "basCreateRota"
Option Explicit

Public BroCount As Integer 'Number of brothers
Public JobCount As Integer    'Number of Jobs
Public SlotCount   As Integer     'Number of slots to fill per week
Public NoOnMics As Integer, NoAttending As Integer, MaxWkWting As Double, WkWtingDiff As Double
Public NoOnPlatform As Integer, NoOnSound As Integer, LoneParentFactor As Double
Public rstConstants As Recordset, rstNameAddress As Recordset
Public rstWeightLkup As Recordset, rstIDJobs As Recordset, rstIDWeight As Recordset
Public rstRota As Recordset, rstBroList As Recordset, rstPersWtgsLkup As Recordset
Public Pos1stJobOnRota As Integer, NoOfWeeks As Integer, RotaStartDate As Date
Public DelUpToDate As Date, RotaEndDate As Date, MainForm As Object, OptionsForm As Object
Public Att_Weight As Integer, Mic_Weight As Integer, Plat_Weight As Integer, Snd_Weight As Integer
Public MaxConsecutiveRovingMic As Integer, MaxConsecutivePlatform As Integer
Public MaxConsecutiveAttendant As Integer, MaxConsecutiveSound As Integer, MaxConsecutiveUpperLimit As Integer
Public WksToCheckBack As Integer, ParentCoeff As Double, ZeroWtgAge As Integer
Public Marriage_Wtg As Double, ContFromLastRota As Boolean, RespCoeff As Double
Public ReturnValue As Integer, AttTimesPerWk As Integer, MicTimesPerWk As Integer, SoundTimesPerWk As Integer
Public PlatformTimesPerWk As Integer, DefaultCong As Integer, CurrentCong As Object, objGetWkWtg As clsWeekWeighting
Public SPAMFormControl As Control
Public PersonalWeightCoeff As Double, WeekIncrement As Single, SundayMeetingDay As String, MidWeekMeetingDay As String
Public RotaStartDay As String, MinOnPlat As Integer, MinOnMics As Integer, MinOnSound As Integer
Public MinOnAtt As Integer, SaveRotaPeriod As Boolean, ProgressBarIncValue As Double
Public RotaTimesPerWeek As Long, HandleListBox As New clsListCombo
Public PTObserved As Boolean
Public RotaEdited As Boolean, StopRotaGeneration As Boolean, CMSDB As DAO.Database
Public CMS_Update_DB As DAO.Database, TotalMicHandlers As Long, TotalAttendants As Long
Public TotalPlatformAttendants As Long, TotalSoundAttendants As Long
Public SundayMeetingDayNo As Long, MidWeekMeetingDayNo As Long
Public gbPrintExistingRota As Boolean

Dim arrJobNames() As String

Dim mbObservePartTimers As Boolean

Public Type RECT       'For window position determination
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type
Public Declare Sub GetWindowRect Lib "user32" (ByVal hWnd As Integer, lpRect As RECT)

Public Const NewRota As Integer = 1
Public Const ContinueRota As Integer = 2
Public Const AppName As String = "C.M.S."
Public Const gsAppNameFull As String = "Congregation Management System"

Public Function MainProc() As Boolean
Dim AttendantsNo As Integer, RovingMicsNo As Integer, SoundNo As Integer, PlatformNo As Integer, PosOf1stJob As Integer
'
'Main controlling procedure
'
Dim JobName As String ', TheConstants As clsApplicationConstants


    On Error GoTo ErrorTrap
    
    'don't allow user to stop rota generation until after tblRota has been reconstructed
    Screen.MousePointer = 11
    MainForm!cmdCancel.Enabled = False

    Set rstRotaForEdit = Nothing
    
    '
    'Calculate Number-of-Weeks on Rota (if an end date has been specified on form).
    'Also find RotaStartDate if continuing from previous rota.
    '
    If Not CalcNumOfWks Then
        MainProc = False
        Screen.MousePointer = 0
        MainForm!cmdCancel.Enabled = True
        Exit Function
    End If
    
    '
    'Get number of brothers on each job
    '
    GetNumberOfBrothers
    
    MainForm!lblProgress.Visible = True
    
    DisableForm
    
    'initialize the progress bar and set the increment value according to number of weeks on rota
    MainForm.afxProgressBar.value = 0
    ProgressBarIncValue = (100 / NoOfWeeks)
    
    
    Set objGetWkWtg = New clsWeekWeighting
    
    objGetWkWtg.InitialiseMaxWtg = MaxWkWting
    objGetWkWtg.InitialiseCoeff = WkWtingDiff
    
    '
    'Set up main recordsets
    '
    If Not BuildRecsets Then
        MainProc = False
        Exit Function
    End If
    
    
    '
    'Analyse current tblRota to ascertain structure...
    '
    AcquireRotaStructure AttendantsNo, RovingMicsNo, SoundNo, PlatformNo, PosOf1stJob
    
   
    If GlobalParms.GetValue("RotaTimesPerWeek", "NumVal") = 2 Then
        WeekIncrement = 0.5
    Else
        WeekIncrement = 1
    End If
       
    
    '
    'Initialise existing rota to contain only Number of Weeks = WksCheckBack
    '
    If Not InitialiseRota Then
        MainProc = False
        Exit Function
    End If
    
    '
    'Build weightings table, then compile Responsibility-weightings for each bro.
    '
    If Not CreateWeightingsTable Then
        MainProc = False
        Exit Function
    End If
    If Not CalcInitWtngs Then
        MainProc = False
        Exit Function
    End If
    
    '
    'Add initial week-weightings for all existing weeks
    '
    If GlobalParms.GetValue("ContFromLastRota", "TrueFalse") Then
        If Not InitWkWt(AttendantsNo, RovingMicsNo, SoundNo, PlatformNo, PosOf1stJob) Then
            MainProc = False
            Exit Function
        End If
    End If
    
    mbObservePartTimers = GlobalParms.GetValue("ObservePartTimers", "TrueFalse")
    
    '
    'Recreate tblRota  - with new structure if necessary
    '
    CreateNewRotaTables
    
    '
    'Analyse new tblRota to ascertain structure...
    '
    AcquireRotaStructure AttendantsNo, RovingMicsNo, SoundNo, PlatformNo, PosOf1stJob
    
    Screen.MousePointer = 0 'Normal
    MainForm!cmdCancel.Enabled = True
    MainForm!cmdCancel.Caption = "STOP"
    MainForm.cmdCancel.SetFocus
    
    '*********************************************
    '* MAIN ROUTINE - Fill tblRota with brothers *
    '                                            *
    If Not FillRota Then
        MainProc = False
        Exit Function
    End If
    '*********************************************
    
    If Not StopRotaGeneration Then
        '
        'Remove old rows from tblRota
        '
        If Not DeleteOld Then
            MainProc = False
            Exit Function
        End If
        
        '
        'Do any dates in new rota PARTIALLY overlap those on a previous rota?
        '
        If Not RotaDatesOverlap Then
    
            '
            'Create copy of rota just built
            '
            If Not CopyRota Then
                MainProc = False
                Exit Function
            End If
        
            EnableForm
            
            MainForm.afxProgressBar.value = 0
            
            '
            'Allow new rota to be manually edited
            '
            GlobalParms.Save "SPAMOptionsChanged", "TrueFalse", False
        Else
            StopRotaGeneration = True
        End If
                
    End If
        
    '
    'Close recordsets
    '
    If Not CloseRecordsets Then
        MainProc = False
        Exit Function
    End If
    
    Erase arrJobNames
        
    MainForm!cmdCancel.Caption = "&Close"
    MainForm!lblProgress.Visible = False
    
    Screen.MousePointer = 0 'Normal
    
'    MsgBox "The new rota has been created. ", _
'           vbOKOnly + vbInformation, AppName
           
    If Not StopRotaGeneration Then
        gbPrintExistingRota = False
        Select Case PrintUsingWord
        Case cmsUseWord
            frmSoundAndPlatformRota.PrintSPAMRotaUsingMSWord 'immediately print rota
        Case cmsUseMSDatareport
            frmSoundAndPlatformRota.PrintSPAMRotaToMSDataReport 'immediately print rota
        End Select
    End If
        
    MainProc = True
    
    Exit Function
    
ErrorTrap:

    Screen.MousePointer = 0 'Normal

    MainProc = False
    
    EndProgram
    
    

End Function
    
    
Public Function FillRota() As Boolean
Dim j As Integer, k As Double

    On Error GoTo ErrorTrap
    
    k = 0
    
    StopRotaGeneration = False
    
    DoEvents
    
    Do
        '
        'Add new row to tblRota with new date
        '
        NewRotaRow
        
        '
        'Now adjust bros' week-weightings for previous weeks
        '
        If Not AdjWkWting Then
            FillRota = False
            MainForm!cmdCancel.Caption = "&Close"
            Exit Function
        End If
                
        '
        'Check whether any brothers are suspended this week. If so zero-ize the appropriate job_wtg
        '
        AnyoneSuspended
                
        '
        'Insert bros into rota from first available rota slot (Pos1stJobOnRota) to
        'final slot (Pos1stJobOnRota + (SlotCount-1))
        '
        'Alternate the flow from R-L then L-R across tblRota to help even out the spread
        '
        If j = Pos1stJobOnRota Or j = 0 Then
            For j = Pos1stJobOnRota To (Pos1stJobOnRota + (SlotCount - 1))
            
                If Not InsertBrothers(j, rstRota!RotaDate) Then
                    FillRota = False
                    MainForm!cmdCancel.Caption = "&Close"
                    EnableForm
                    Exit Function
                End If
                
                If PTObserved Then 'are we observing part-timers?
                    If GetJobName(j) <> "RovingMic" Then 'Forget this for Roving Mics
                        If j < (Pos1stJobOnRota + (SlotCount - 1)) Then 'ensure don't overrun rota columns in look-ahead
                            If GetJobName(j) = GetJobName(j + 1) Then 'Only put brother on again if same job
                                If CongregationMember.CanBeInThisSlot(giGlobalDefaultCong, rstRota.Fields(j), rstRota!RotaDate, GetJobName(j)) = "Y" Then  'Special provision if School or Watchtower overseer
                                    With rstRota
                                    If Not CongregationMember.IsPartTimer(giGlobalDefaultCong, .Fields(j)) Then
                                        .Edit
                                        .Fields(j + 1) = .Fields(j)
                                        .Update
                                                                
                                        j = j + 1
                                    
                                    End If
                                    End With
                                End If
                            End If
                        End If
                    End If
                End If
            Next j
        Else
            For j = (Pos1stJobOnRota + (SlotCount - 1)) To Pos1stJobOnRota Step -1
            
                If Not InsertBrothers(j, rstRota!RotaDate) Then
                    FillRota = False
                    EnableForm
                    Exit Function
                End If
                
                'if brother is part-timer, don't put him on two halves of meeting. Doesn't apply to Mics since
                ' this generally is only done for half meeting anyway.
                If PTObserved Then 'are we observing part-timers?
                    If GetJobName(j) <> "RovingMic" Then 'Forget this for Roving Mics
                        If j > Pos1stJobOnRota Then 'ensure don't overrun rota columns in look-ahead
                            If GetJobName(j) = GetJobName(j - 1) Then 'Only put brother on again if same job
                                If CongregationMember.CanBeInThisSlot(giGlobalDefaultCong, rstRota.Fields(j), rstRota!RotaDate, GetJobName(j)) = "Y" Then 'Special provision if School or Watchtower overseer
                                    With rstRota
                                    If Not CongregationMember.IsPartTimer(giGlobalDefaultCong, .Fields(j)) Then
                                        .Edit
                                        .Fields(j - 1) = .Fields(j)
                                        .Update
                                                                
                                        j = j - 1
                                    
                                    End If
                                    End With
                                End If
                            End If
                        End If
                    End If
                End If
            Next j
        End If
               
        '
        'Make all brothers available for selection for next week
        '
        CMSDB.Execute ("UPDATE tblIDWeightings " & _
                           "SET OnThisWeek = " & 0 & _
                           ", Weighting = " & 0)
        
        rstRota.MoveLast
        
        k = k + ProgressBarIncValue * WeekIncrement
        
        If k > 100 Then 'In case of rounding errors, prevent ProgressBar value going above max value (100)
            k = 100
        End If
        
        MainForm.afxProgressBar.value = k
        
        DoEvents
        
        If StopRotaGeneration Then
            MainForm.afxProgressBar.value = 0
            Exit Function
        End If
        
    Loop Until rstRota!RotaDate >= RotaEndDate
    
    FillRota = True

    Exit Function


ErrorTrap:
    EndProgram
    
End Function

Public Sub NewRotaRow()
Dim NewDate As Date

'
'Add new row to tblRota with new date
'
    
    On Error GoTo ErrorTrap
    
    NewDate = NextRotaDate
    
    With rstRota
    .AddNew
    !RotaDate = NewDate
    .Update
    End With
    
    
    CMSDB.TableDefs.Refresh
    
    
    Exit Sub
    
ErrorTrap:

    EndProgram
    
End Sub
Public Function NextRotaDate() As Date
Dim CurrentRotaDate As Date, NewDate As Date
    
    On Error GoTo ErrorTrap
    
    With rstRota
    
    If Not .BOF Then
        .MoveLast
        NewDate = CalculateNextRotaDate(!RotaDate)
    Else
        NewDate = MainForm!txtRotaStart
    End If
    
    End With
    
    Do Until Not IsCircuitOrDistrictAssemblyWeek(NewDate)
        NewDate = CalculateNextRotaDate(NewDate)
    Loop
    
    If RotaTimesPerWeek = 2 Then
        Do Until Not IsMemorialDay(NewDate)
           NewDate = CalculateNextRotaDate(NewDate)
        Loop
    End If
    
    NextRotaDate = NewDate
          
    Exit Function
    
ErrorTrap:

    EndProgram
    
End Function

Public Function CalculateNextRotaDate(CurrentRotaDate As Date) As Date
Static plCOMeetingCount As Long
    
    On Error GoTo ErrorTrap
    
    If RotaTimesPerWeek = 2 Then
        If Not IsCOVisitWeek(CurrentRotaDate) Then
            If Format(CurrentRotaDate, "dddd") <> SundayMeetingDay And Format(CurrentRotaDate, "dddd") <> MidWeekMeetingDay Then
                CalculateNextRotaDate = CurrentRotaDate + 7
            Else
                If Format(CurrentRotaDate, "dddd") = SundayMeetingDay Then
                    If GetWeekDay(SundayMeetingDay) < GetWeekDay(MidWeekMeetingDay) Then
                        CalculateNextRotaDate = CurrentRotaDate + (GetWeekDay(MidWeekMeetingDay) - GetWeekDay(SundayMeetingDay))
                    Else
                        CalculateNextRotaDate = CurrentRotaDate + (7 - (GetWeekDay(SundayMeetingDay) - GetWeekDay(MidWeekMeetingDay)))
                    End If
                Else
                    If GetWeekDay(MidWeekMeetingDay) < GetWeekDay(SundayMeetingDay) Then
                        CalculateNextRotaDate = CurrentRotaDate + (GetWeekDay(SundayMeetingDay) - GetWeekDay(MidWeekMeetingDay))
                    Else
                        CalculateNextRotaDate = CurrentRotaDate + (7 - (GetWeekDay(MidWeekMeetingDay) - GetWeekDay(SundayMeetingDay)))
                    End If
                End If
            End If
            
            If IsCOVisitWeek(CalculateNextRotaDate) Then
                CalculateNextRotaDate = GetDateOfFirstWeekDayOfMonth(CalculateNextRotaDate, vbMonday) + 1 'tuesday
                plCOMeetingCount = 1
            End If
            
        Else 'is co week...
            Select Case plCOMeetingCount
            Case 0
                If GetWeekDay(SundayMeetingDay) < GetWeekDay(MidWeekMeetingDay) Then
                    CalculateNextRotaDate = CurrentRotaDate + (GetWeekDay(MidWeekMeetingDay) - 1)
                Else
                    CalculateNextRotaDate = CurrentRotaDate + (GetWeekDay(SundayMeetingDay) - 1)
                End If
            Case 1
                CalculateNextRotaDate = GetDateOfFirstWeekDayOfMonth(CurrentRotaDate, vbMonday) + 3 'thursday
                plCOMeetingCount = 2
            Case 2
                CalculateNextRotaDate = GetDateOfFirstWeekDayOfMonth(CurrentRotaDate, vbMonday) + 6 'sunday
                plCOMeetingCount = 0
            End Select
        End If
    Else
        CalculateNextRotaDate = CurrentRotaDate + 7
    End If
    
    Exit Function
    
ErrorTrap:

    EndProgram
    
End Function


Public Function AdjWkWting() As Boolean
'
'Move back from previous week (-1) and adjust each bro's weighting
'

Dim WeekNum As Integer
      
    On Error GoTo ErrorTrap
   
    With rstRota
    
    .MoveLast
    .Move (-1) 'Back a week
    
    WeekNum = -1
    
    If Not .BOF Then
        
        Do
            '
            'Add new week-weighting for each brother on all prior weeks
            '
            If Not AddWkWt(WeekNum) Then           'WeekNum = WeekIndex. Current wk = 0
                AdjWkWting = False                    'prev week = -1 etc...
                Exit Function
            End If
            
            '
            'Now check all brothers suspended this week - add a weighting for them
            '
            AddWkWtForSuspendedBros (WeekNum)
            
            
            WeekNum = WeekNum - 1
            
            .Move (-1) 'Back a week
            
        Loop Until .BOF
    
    End If
    
    End With
    
    
    
    AdjWkWting = True
    
    Exit Function
    
ErrorTrap:

    EndProgram
    
End Function
Public Function AddWkWt(WkIndex As Integer) As Boolean
Dim n As Byte, PrevBro As Integer, TheWtg As Double

    On Error GoTo ErrorTrap
    
    PrevBro = 0 'initialise
'
'For each job-slot on week-being-processed add on new week weighting, BUT only
' if bro changed. This is because if we're observing part-timers, full-timers
' will appear in two consecutive slots. Don't want to double up their weighting.
'
    With rstIDWeight
    For n = Pos1stJobOnRota To (Pos1stJobOnRota + (SlotCount - 1)) 'ie for each job-slot
        
        .FindFirst "ID = " & rstRota.Fields(n)
        
        If !ID <> PrevBro Then
            .Edit
            TheWtg = (GetWkWt(WkIndex) + 1) * 10 * _
                            ((!ResponsibilityWeighting + 1) * (!Personal_Wtg + 1))
            
            Select Case GetJobName(CInt(n))
            Case "Sound"
                !Weighting = !Weighting + TheWtg * CongregationMember.GetIndividualSPAMWeighting(!ID, cmsSound)
            Case "Platform"
                !Weighting = !Weighting + TheWtg * CongregationMember.GetIndividualSPAMWeighting(!ID, cmsPlatform)
            Case "Attendant"
                !Weighting = !Weighting + TheWtg * CongregationMember.GetIndividualSPAMWeighting(!ID, cmsAttendant)
            Case "RovingMic"
                !Weighting = !Weighting + TheWtg * CongregationMember.GetIndividualSPAMWeighting(!ID, cmsMicrophones)
            End Select
            
            .Update
            
            PrevBro = !ID
        End If
        
    Next n
    End With

    AddWkWt = True
    
    Exit Function
    
ErrorTrap:

    EndProgram
    

End Function

Public Sub AddWkWtForSuspendedBros(WkIndex As Integer)
Dim RotaDateUS As String, rstSuspendedBros As Recordset

    On Error GoTo ErrorTrap
    
    RotaDateUS = Format(rstRota!RotaDate, "mm/dd/yyyy")
    
    '
    'All suspended brothers this week...
    '
    Set rstSuspendedBros = CMSDB.OpenRecordset("SELECT DISTINCT Person " & _
                                            " FROM tblTaskPersonSuspendDates " & _
                                            " WHERE TaskCategory = 6 " & _
                                            "  AND TaskSubCategory IN (10, 11) " & _
                                            "  AND Task IN (57, 58, 59, 60)" & _
                                            "  AND (SuspendStartDate <= #" & RotaDateUS & _
                                                    "# AND SuspendEndDate >= #" & RotaDateUS & "#)", _
                                            dbOpenForwardOnly)
                
    If Not rstSuspendedBros.BOF Then
        Do Until rstSuspendedBros.EOF
            With rstIDWeight
            
            .FindFirst "ID = " & rstSuspendedBros!Person
            
            .Edit
            
            Select Case True
            Case CongregationMember.IsSound(rstSuspendedBros!Person, giGlobalDefaultCong)
                !Weighting = !Weighting + ((GetWkWt(WkIndex) + 1) * 10 * _
                                    ((!ResponsibilityWeighting + 1) * (!Personal_Wtg + 1))) / TotalSoundAttendants
            Case CongregationMember.IsPlatform(rstSuspendedBros!Person, giGlobalDefaultCong)
                !Weighting = !Weighting + ((GetWkWt(WkIndex) + 1) * 10 * _
                                    ((!ResponsibilityWeighting + 1) * (!Personal_Wtg + 1))) / TotalPlatformAttendants
            Case CongregationMember.IsAttendant(rstSuspendedBros!Person, giGlobalDefaultCong)
                !Weighting = !Weighting + ((GetWkWt(WkIndex) + 1) * 10 * _
                                    ((!ResponsibilityWeighting + 1) * (!Personal_Wtg + 1))) / TotalAttendants
            Case CongregationMember.IsRovingMic(rstSuspendedBros!Person, giGlobalDefaultCong)
                !Weighting = !Weighting + ((GetWkWt(WkIndex) + 1) * 10 * _
                                    ((!ResponsibilityWeighting + 1) * (!Personal_Wtg + 1))) / TotalMicHandlers
            End Select
            
            .Update
            
            End With
            
            rstSuspendedBros.MoveNext
        Loop
    
    End If
        
    Exit Sub
    
ErrorTrap:

    EndProgram
    

End Sub
Public Function SubtractWkWt(WkIndex As Integer) As Boolean
Dim n As Byte, PrevBro As Integer, tempvar

    On Error GoTo ErrorTrap
    
    PrevBro = 0 'initialise

'
'For each job-slot on week-being-processed subtract previous
' week weighting. BUT only if bro changed. This is because
' if we're observing part-timers, full-timers will appear in two consecutive slots. Don't want to double up their
' weighting subtraction.
'
'
    
    With rstIDWeight
    For n = Pos1stJobOnRota To (Pos1stJobOnRota + (SlotCount - 1)) 'ie for each job-slot
        
        .FindFirst "ID = " & rstRota.Fields(n)
        
        If !ID <> PrevBro Then
            .Edit
            tempvar = (GetWkWt(WkIndex) + 1) * 10 * ((!ResponsibilityWeighting + 1) * (!Personal_Wtg + 1))
            !Weighting = !Weighting - tempvar
            
            .Update
            
            PrevBro = !ID
        End If
        
    Next n
    End With

    SubtractWkWt = True
    
    Exit Function
    
ErrorTrap:

    EndProgram
    
End Function

Public Function InsertBrothers(CurrentRotaSlot As Integer, _
                                CurrentRotaDate As Date) As Boolean
Dim FindBroResult As String, TempMaxAtt As Integer, TempMaxPlat As Integer, TempMaxSnd As Integer, TempMaxRov As Integer
Dim MaxConsecutive As Integer, SelectedBro As Integer
Dim bCheckOtherMtgs As Boolean, lTimesPerWk As Long
Dim RotaMondayDate As Date, bOK As Boolean

    On Error GoTo ErrorTrap
    
    TempMaxAtt = MaxConsecutiveAttendant
    TempMaxPlat = MaxConsecutivePlatform
    TempMaxRov = MaxConsecutiveRovingMic
    TempMaxSnd = MaxConsecutiveSound
    
    'only check for conflict with other mtgs if RotaTimesPerWeek = 2
    bCheckOtherMtgs = (RotaTimesPerWeek = 2)
    
    RotaMondayDate = GetDateOfGivenDay(CurrentRotaDate, vbMonday, False)
        
    With rstRota
    
    .MoveLast
    '
    'Get list of brothers, in order of weighting, to fill slot
    '
    Do
        FindBroResult = FindBro(GetJobName(CurrentRotaSlot))
        
        If FindBroResult = "ERROR" Then
            InsertBrothers = False
            Exit Function
        End If
        
        
        '
        'If can't find any brother for slot, increase the MaxConsecutive value and try again
        '
        If FindBroResult = "BROTHERS FOUND" Then
            rstBroList.MoveFirst
            '
            '1st eligible bro in list goes into slot. Ensure that Watchtower COnductor
            ' does not do Roving Mic on Sunday!
            'Also check against other mtgs for conflicts - but only if
            ' RotaTimesPerWeek is 2
            '
            Do
                SelectedBro = rstBroList!Person
                If CongregationMember.CanBeInThisSlot(giGlobalDefaultCong, SelectedBro, !RotaDate, GetJobName(CurrentRotaSlot)) <> "N" Then
                    If bCheckOtherMtgs Then
                        bOK = CongregationMember.OKforSpamAgainstOtherMtgs( _
                                CLng(SelectedBro), CurrentRotaDate, RotaMondayDate, _
                                SundayMeetingDayNo, MidWeekMeetingDayNo)
                    Else
                        bOK = True
                    End If
                    
                    If bOK Then
                        Exit Do
                    End If
                End If
                rstBroList.MoveNext
                
                If rstBroList.EOF And bCheckOtherMtgs Then
                    'if reached this point, then we've not found a bro yet
                    ' try again, this time without checking against other mtgs
                    bCheckOtherMtgs = False
                    rstBroList.MoveFirst
                End If
                
            Loop Until rstBroList.EOF
            
            If Not rstBroList.EOF Then
                Exit Do 'SUCCESS!!!!!!!!!
            End If
        End If
        
        '
        'If this point is reached, no brother has been found.... increase
        ' MaxConsecutive and try again.....
        '
        Select Case GetJobName(CurrentRotaSlot)
        Case "Attendant"
            MaxConsecutiveAttendant = MaxConsecutiveAttendant + 1
        Case "RovingMic"
            MaxConsecutiveRovingMic = MaxConsecutiveRovingMic + 1
        Case "Sound"
            MaxConsecutiveSound = MaxConsecutiveSound + 1
        Case "Platform"
            MaxConsecutivePlatform = MaxConsecutivePlatform + 1
        End Select
        
    Loop Until FindBroResult = "IMPOSSIBLE"
    
    MaxConsecutiveAttendant = TempMaxAtt
    MaxConsecutivePlatform = TempMaxPlat
    MaxConsecutiveRovingMic = TempMaxRov
    MaxConsecutiveSound = TempMaxSnd
    
    If FindBroResult = "IMPOSSIBLE" Or rstBroList.EOF Then
        frmSoundAndPlatformRota.cmdCancel.Caption = "&Close"
        NotEnoughBrosMessage
        InsertBrothers = False
        Exit Function
    End If
    
    .Edit
    
    .Fields(CurrentRotaSlot) = SelectedBro
    
    .Update
    
    End With
    
    '
    'Now adjust brother's SPAM-weighting according to job he's just been entered
    ' for - done to stop him doing the same thing again next time
    
    If Not AdjJobWtg(CurrentRotaSlot) Then
        InsertBrothers = False
       Exit Function
    End If
    
    
    '
    'Now adjust various settings for brother just added to rota.
    '
    With rstIDWeight
    .FindFirst ("ID = " & rstBroList!Person)
    .Edit
                                                     
    '
    'indicate that brother has been inserted for this week
    '
    !OnThisWeek = !OnThisWeek + 1
    
    '
    'Increment counter saying how many times brother has been on Rota.
    '
    !NoOfTimesOnRota = !NoOfTimesOnRota + 1
    
    .Update
    
    Select Case GetJobName(CurrentRotaSlot)
    Case "Attendant"
        MaxConsecutive = MaxConsecutiveAttendant
    Case "RovingMic"
        MaxConsecutive = MaxConsecutiveRovingMic
    Case "Sound"
        MaxConsecutive = MaxConsecutiveSound
    Case "Platform"
        MaxConsecutive = MaxConsecutivePlatform
    End Select
    
    .Edit
        
    .Update
      
    
    End With
    
    InsertBrothers = True
    
    Exit Function


ErrorTrap:

    EndProgram
       
End Function
Public Function AdjJobWtg(CurrentRotaSlot As Integer) As Boolean
    '
    'Now adjust Job-weighting to reflect whether bro has just been put on mics or att.
    'Value is derived from tblConstants.
    'This is used to help put each bro evenly on Mic & Att and Sound & Platform.
    'Job-weighting not adjusted if bro ONLY does Mic OR Att, or does Sound OR Platform
    '
    
    On Error GoTo ErrorTrap
    
    With rstIDWeight
    
    .FindFirst ("ID = " & rstBroList!Person)
       
    Select Case GetJobName(CurrentRotaSlot)
    Case "Attendant"
        If CongregationMember.IsRovingMic(!ID, giGlobalDefaultCong) Or _
            CongregationMember.IsPlatform(!ID, giGlobalDefaultCong) Or _
            CongregationMember.IsSound(!ID, giGlobalDefaultCong) Then
            .Edit
            !Attendant_Wtg = !Attendant_Wtg + Att_Weight
            !RovingMic_Wtg = 0
            !Sound_Wtg = 0
            !Platform_Wtg = 0
            .Update
        End If
    Case "RovingMic"
        If CongregationMember.IsAttendant(!ID, giGlobalDefaultCong) Or _
            CongregationMember.IsPlatform(!ID, giGlobalDefaultCong) Or _
            CongregationMember.IsSound(!ID, giGlobalDefaultCong) Then
            .Edit
            !RovingMic_Wtg = !RovingMic_Wtg + Mic_Weight
            !Attendant_Wtg = 0
            !Sound_Wtg = 0
            !Platform_Wtg = 0
            .Update
        End If
    Case "Sound"
        If CongregationMember.IsPlatform(!ID, giGlobalDefaultCong) Or _
            CongregationMember.IsAttendant(!ID, giGlobalDefaultCong) Or _
            CongregationMember.IsRovingMic(!ID, giGlobalDefaultCong) Then
            .Edit
            !Sound_Wtg = !Sound_Wtg + Snd_Weight
            !Attendant_Wtg = 0
            !RovingMic_Wtg = 0
            !Platform_Wtg = 0
            .Update
        End If
    Case "Platform"
        If CongregationMember.IsSound(!ID, giGlobalDefaultCong) Or _
            CongregationMember.IsAttendant(!ID, giGlobalDefaultCong) Or _
            CongregationMember.IsRovingMic(!ID, giGlobalDefaultCong) Then
            .Edit
            !Platform_Wtg = !Platform_Wtg + Plat_Weight
            !Attendant_Wtg = 0
            !RovingMic_Wtg = 0
            !Sound_Wtg = 0
            .Update
        End If
    End Select
    
    End With
    
    AdjJobWtg = True
    
    Exit Function


ErrorTrap:

    EndProgram
           
End Function


Public Function FindBro(Job As String) As String
Dim RotaDateUS As Date, SuspendStartUS As Date, SuspendEndUS As Date
Dim MaxConsecutive As Integer, TaskCat As Integer, TaskSubCat As Integer, Task As Integer
Dim FindBroSQL As String, DayQuery As String

    On Error GoTo ErrorTrap
    
'
'SQL used to select list of bros eligible for the slot. Dates used in SQL must be US format.
'
    
    rstRota.MoveLast
    
    RotaDateUS = Format(rstRota!RotaDate, "mm/dd/yy")
    
    Select Case Job
    Case "Attendant"
        MaxConsecutive = MaxConsecutiveAttendant
        Task = 60
    Case "RovingMic"
        MaxConsecutive = MaxConsecutiveRovingMic
        Task = 59
    Case "Sound"
        MaxConsecutive = MaxConsecutiveSound
        Task = 57
    Case "Platform"
        MaxConsecutive = MaxConsecutivePlatform
        Task = 58
    End Select
    
    If RotaTimesPerWeek = 2 Then
        Select Case Format(rstRota!RotaDate, "dddd")
        Case SundayMeetingDay:
            DayQuery = " AND ((b.OnSunday = TRUE AND b.OnMidweek = FALSE) OR (b.OnSunday = TRUE AND b.OnMidweek = TRUE) OR (b.OnSunday = FALSE AND b.OnMidweek = FALSE))"
        Case MidWeekMeetingDay:
            DayQuery = " AND ((b.OnSunday = FALSE AND b.OnMidweek = TRUE) OR (b.OnSunday = TRUE AND b.OnMidweek = TRUE) OR (b.OnSunday = FALSE AND b.OnMidweek = FALSE))"
        End Select
    Else
        DayQuery = ""
    End If
        
    FindBroSQL = "SELECT b.Person, " & _
                "        a.Weighting FROM (tblIDWeightings a " & _
                "INNER JOIN tblTaskAndPerson b ON a.ID = b.Person) " & _
                " INNER JOIN tblNameAddress c ON b.Person = c.ID " & _
                "WHERE b.Task = " & Task & _
                " AND c.Active = TRUE" & _
                DayQuery & " " & _
                "AND a.OnThisWeek = " & 0 & _
                " AND a.[" & Job & "_Wtg] < " & MaxConsecutive & " " & _
                " AND NOT EXISTS (SELECT 1 " & _
                                    "FROM tblTaskPersonSuspendDates d " & _
                                    "WHERE Task = " & Task & _
                                    " AND d.Person = b.Person " & _
                                    " AND (SuspendStartDate <= #" & RotaDateUS & "# AND SuspendEndDate >= #" & RotaDateUS & "#)) " & _
                " ORDER BY a.Weighting, a.[" & Job & "_Wtg], a.NoOfTimesOnRota"
                
    
    Set rstBroList = CMSDB.OpenRecordset(FindBroSQL, dbOpenDynaset)
    
    With rstBroList
    If .BOF Then
        If MaxConsecutive = MaxConsecutiveUpperLimit Then
            FindBro = "IMPOSSIBLE"
        Else
            FindBro = "NO BROTHERS FOUND"
        End If
    Else
        FindBro = "BROTHERS FOUND"
    End If
    End With
    
    Exit Function


ErrorTrap:

    EndProgram


End Function


Public Function GetWkWt(WkNum As Integer) As Double
    

On Error GoTo ErrorTrap

    
    objGetWkWtg.TheWeek = WkNum
    
    GetWkWt = objGetWkWtg.WeekWeighting
    
    Exit Function

ErrorTrap:
    EndProgram
    
End Function

Public Function GetJobName(RotaFldPos As Integer) As String
Dim Pos As Long, RotaFldName As String


On Error GoTo ErrorTrap

    GetJobName = arrJobNames(RotaFldPos - 2)

''
''Derive jobname from fieldnames on tblRota. This jobname is used to reference jobs on
''tblResponsibilities.
''
''All job-slots on tblRota are of format <Jobname>_nn (nn=integer).
''
'    '
'    'Find position of "_" in Rota fieldname
'    '
'    RotaFldName = rstRota.Fields(RotaFldPos).Name
'    Pos = InStr(1, RotaFldName, "_")
'
'    '
'    'JobName is left part of Rota Fieldname. eg, if Rota Fieldname is Attendant_1,
'    'then pos = 10. Job name is Left 9 chars.
'    '
'    GetJobName = Left$(RotaFldName, Pos - 1)
    
    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Function BuildRecsets() As Boolean


On Error GoTo ErrorTrap

    
    Set rstWeightLkup = CMSDB.OpenRecordset("tblTaskWeightings", dbOpenDynaset)
    
    Set rstIDJobs = CMSDB.OpenRecordset("tblTasks", dbOpenDynaset)
                                                                                                                
    BuildRecsets = True
    
    Exit Function
    
ErrorTrap:

    EndProgram
                                                                                                                
End Function

Public Function InitWkWt(AttendantsNo As Integer, RovingMicsNo As Integer, SoundNo As Integer, PlatformNo As Integer, PosOf1stJob As Integer) As Boolean
Dim RotaRow As Integer, Slot As Integer

On Error GoTo ErrorTrap

'
'Move back from previous Rota Row (-1 ie last week of previous rota) and adjust each bro's week-weighting
'
'
    RotaRow = -1

   
    With rstRota
    
    
    If Not .BOF Then
    
    .MoveLast
        
        Do
            If Not AddInitWkWt(RotaRow, AttendantsNo, RovingMicsNo, SoundNo, PlatformNo, PosOf1stJob) Then              'Week = WeekIndex. Current wk = 0
                InitWkWt = False
                Exit Function
            End If
                                                            
            RotaRow = RotaRow - 1
            
            .Move (-1) 'Back a row
            
        Loop Until .BOF
    
    End If
    
'
'Now move forward through existing rota to determine initial Job-Weightings
'
        
    If Not .BOF Then
    
    .MoveFirst
    
        Do
            For Slot = PosOf1stJob To (PosOf1stJob + (AttendantsNo + RovingMicsNo + SoundNo + PlatformNo - 1))
                
                rstIDWeight.FindFirst "ID = " & .Fields(Slot)
                rstIDJobs.FindFirst "ID = " & .Fields(Slot)

                If Not AdjJobWtg(Slot) Then
                    InitWkWt = False
                    Exit Function
                End If
                
            Next Slot
            
            .MoveNext
            
        Loop Until .EOF
    
    End If
    
    End With
    
    InitWkWt = True
    
    Exit Function


ErrorTrap:

    EndProgram
    
End Function

Public Function AddInitWkWt(RowIX As Integer, AttendantsNo As Integer, RovingMicsNo As Integer, SoundNo As Integer, PlatformNo As Integer, PosOf1stJob As Integer) As Boolean
Dim m As Byte

    On Error GoTo ErrorTrap
    
'
'For each job-slot on week-being-processed add on new week weighting.
'

    For m = PosOf1stJob To (PosOf1stJob + (AttendantsNo + RovingMicsNo + SoundNo + PlatformNo - 1)) 'ie for each job-slot
        
        If Not (IsNull(rstRota.Fields(m))) Or Len(Trim(rstRota.Fields(m))) > 0 Then
            rstIDWeight.FindFirst "ID = " & rstRota.Fields(m)
            
            rstIDWeight.Edit
            rstIDWeight!Weighting = rstIDWeight!Weighting + GetWkWt(RowIX)
            rstIDWeight.Update
        End If
        
    Next m

    AddInitWkWt = True
    
    Exit Function


ErrorTrap:

    EndProgram

End Function

Public Function CopyRota() As Boolean
'
'Create copy of tblRota just produced using DoCmd.
'Name of table will be tblRota||<StartDate>||<EndDate>
'
Dim BeginDate As Date, StrBeginDate As String
Dim EndDate As Date, StrEndDate As String
Dim NewTable As String, i As Integer, temp, rstSPAMRotas As Recordset
Dim bOverwrite As Boolean

    On Error GoTo ErrorTrap
    
    i = 1
    
    With rstRota
    
    .MoveFirst
    BeginDate = !RotaDate
    StrBeginDate = CStr(BeginDate)
    
    .MoveLast
    EndDate = !RotaDate
    StrEndDate = CStr(EndDate)
        
    End With
    
    NewTable = "tblRota: " & StrBeginDate & " TO " & StrEndDate & " (" & Format(i, "0000") & ")"
            
    '
    'Does this tablename already exist? If so, keep incrementing suffix (i) and trying again until unique name
    ' found.
    '
    If Not TableExists(NewTable) Then 'Table doesn't exist. Fine.
        CopyTable NewTable, "tblRota", CMSDB
    Else
        bOverwrite = False 'init
        If MsgBox("A rota for this period already exists. Do you want to overwrite it?", vbYesNo + vbQuestion, AppName) = vbYes Then
            DelAllRows "[" & NewTable & "]"
            CopyTable NewTable, "tblRota", CMSDB
            CMSDB.Execute "DELETE FROM tblStoredSPAMRotas " & _
                          "WHERE RotaTableName = '" & NewTable & "'"
            
            bOverwrite = True
        Else
            Do
                i = i + 1
                Err.Clear
                NewTable = "tblRota: " & StrBeginDate & " TO " & StrEndDate & " (" & Format(i, "0000") & ")"
            Loop Until Not TableExists(NewTable) Or i = 9999
            
            If i = 9999 Then
                MsgBox "Could not copy tblRota. Delete some tables.", vbOKOnly + vbCritical, AppName
                CopyRota = False
                Exit Function
            Else
                CopyTable NewTable, "tblRota", CMSDB
            End If
        End If
    End If
            
    '
    'Insert new table name on tblStoredSPAMRotas - for display in combo
    '
             
    Set rstSPAMRotas = CMSDB.OpenRecordset("SELECT * FROM tblStoredSPAMRotas  ORDER BY ModifiedDateTime DESC", dbOpenDynaset)

    With rstSPAMRotas
    .AddNew
    !RotaTableName = NewTable
    !DisplayForCombo = "Rota dates " & StrBeginDate & " TO " & StrEndDate & " (" & Format(i, "0000") & ")"
    !CreatedDateTime = Now
    !ModifiedDateTime = Now
    .Update
    
    '
    'Now check if we need to delete old rotas
    '
    .Requery
    .MoveLast
    
    If GlobalParms.GetValue("NumberOfSPAMRotasToKeep", "NumVal") < .RecordCount Then
        DeleteTable !RotaTableName
        
        .Delete
    End If
    
    End With
    
    
    HandleListBox.Requery MainForm!cmbStoredRotas, False, CMSDB
    SetCmbSPAMRotas
    
    CopyRota = True
    
    Exit Function


ErrorTrap:

    EndProgram


End Function

Public Function DeleteOld() As Boolean
'
'Delete all rows from tblRota where date is prior to Start-date of current rota.
'Must format date to mmddyy for use in SQL.
'
    On Error GoTo TableInUse

    CMSDB.Execute ("DELETE * FROM tblRota " & _
                       "WHERE RotaDate < #" & Format(RotaStartDate, "mm/dd/yy") & "#")
    
    '
    'Update rstRota to reflect changes just made to tblRota
    '
    rstRota.Requery

    DeleteOld = True
    
    Exit Function
    
TableInUse:
    Select Case Err.number
    Case 0, 2580, 3265
        DeleteOld = True
    Case 2008, 3211, 3008, 3260, 3262
        DeleteOld = False
        MsgBox " tblRota is in use - close it down", vbExclamation, AppName
        EndProgram
    Case Else
        DeleteOld = False
        MsgBox "Error " & Err.number & " occured while deleting from tblRota." _
           , vbExclamation + vbOKOnly, AppName
        EndProgram
    End Select

End Function
Public Function RotaDatesOverlap() As Boolean
Dim rs As Recordset, rs2 As Recordset, str As String, rs3 As Recordset, bFound As Boolean
Dim rs4 As Recordset

'
'
'
'
    Set rs = CMSDB.OpenRecordset("SELECT MAX(RotaDate) AS MaxDate, MIN(RotaDate) AS MinDate from tblRota")
    
    With rs
    
    If IsNull(!MaxDate) Or IsNull(!MinDate) Then
        RotaDatesOverlap = False
        Exit Function
    End If
        
    str = "SELECT RotaTableName, DisplayForCombo " & _
          "FROM tblStoredSPAMRotas "
    
    Set rs2 = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    Do Until rs2.BOF Or rs2.EOF
    
        Set rs3 = CMSDB.OpenRecordset("SELECT MAX(RotaDate) AS MaxDate, " & _
                                      "MIN(RotaDate) AS MinDate " & _
                                      "FROM [" & rs2!RotaTableName & "]")
        
        
        If !MaxDate = rs3!MaxDate And !MinDate = rs3!MinDate Then
            'this scenario is covered by the appending of '(nnnn)' to the rota name
        Else
            If (DateDiff("d", !MinDate, rs3!MinDate) >= 0 And DateDiff("d", rs3!MinDate, !MaxDate) >= 0) Or _
                (DateDiff("d", !MinDate, rs3!MaxDate) >= 0 And DateDiff("d", rs3!MaxDate, !MaxDate) >= 0) Then
                
                
                If MsgBox("The new rota partially overlaps with an existing rota (" & _
                            rs2!DisplayForCombo & ")." & vbCrLf & vbCrLf & _
                            "Do you want to delete the overlapping dates from the old rota?", _
                            vbYesNo + vbQuestion, AppName) = vbYes Then
                      CMSDB.Execute "DELETE FROM [" & rs2!RotaTableName & "] " & _
                                "WHERE RotaDate BETWEEN " & GetDateStringForSQLWhere(!MinDate) & _
                                "                   AND " & GetDateStringForSQLWhere(!MaxDate)
                                
                Else
                    RotaDatesOverlap = True
                    GoTo GetOut
                End If
                
            End If
        End If
        
        rs2.MoveNext
        
    Loop


    End With
    
GetOut:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    rs2.Close
    Set rs2 = Nothing
    rs3.Close
    Set rs3 = Nothing
    On Error GoTo ErrorTrap
    
    Exit Function
    
ErrorTrap:

    EndProgram

End Function

Public Function CloseRecordsets() As Boolean

    On Error GoTo ErrorTrap
    
    Set rstIDWeight = Nothing
    Set rstBroList = Nothing
    
    CloseRecordsets = True
    
    Exit Function


ErrorTrap:

    EndProgram
    
    
End Function


Public Function InitialiseRota() As Boolean
On Error GoTo TableInUse

Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer
Dim PrevRotaStartDate As Date, rs As Recordset

    'if structure of current tblRota differs from that about to be produced, then we can't
    ' really generate tblPrevRota, used for the Service Meeting announcements' auto-text feature
    ' See functions AttendantsAfterDate and AttendantsToday
    
    AcquireRotaStructure a, b, c, d, e
    
    If a = NoAttending And _
       b = NoOnMics And _
       c = NoOnSound And _
       d = NoOnPlatform Then
       
        Set rs = CMSDB.OpenRecordset("SELECT RotaDate FROM tblRota ORDER BY 1 DESC", dbOpenForwardOnly)
        
        With rs
        
        If Not .BOF Then
            If CDate(!RotaDate) < RotaStartDate Then
                CopyTable "tblPrevRota", "tblRota", CMSDB
            Else
                DeleteTable "tblPrevRota"
            End If
        Else
            DeleteTable "tblPrevRota"
        End If
        
        End With
       
    Else
        DeleteTable "tblPrevRota"
    End If
    
'
'Delete all rows from tblRota where date is prior to Two weeks before Start-date of current rota.
'Must format date to mmddyy for use in SQL.
'

    CMSDB.Execute ("DELETE * FROM tblRota " & _
                       "WHERE RotaDate < #" & Format(RotaStartDate, "mm/dd/yy") & "# - " & (7 * WksToCheckBack))
    
    '
    'Update rstRota to reflect changes just made to tblRota
    '
    rstRota.Requery
    

    InitialiseRota = True

    Exit Function
    
TableInUse:
    Select Case Err.number
    Case 0, 2580, 3265
        InitialiseRota = True
    Case 2008, 3211, 3008, 3260, 3262
        InitialiseRota = False
        MsgBox "Function: InitialiseRota. tblRota is in use - close it down", vbExclamation, AppName
        EndProgram
    Case Else
        InitialiseRota = False
        EndProgram
    End Select


End Function

Public Function SetContinueNewOpt() As Boolean
    
    On Error GoTo ErrorTrap
    
    '
    'Set Continue/New option
    '
    If ContFromLastRota Then
        MainForm!optContinue = True
        MainForm!txtRotaStart.Enabled = False
        MainForm!cmdShowCalendar1.Enabled = False
        'SetUptblRota
        Call CalcRotaStartDate
        MainForm!txtRotaStart = RotaStartDate
    Else
        MainForm!optNew = True
        MainForm!txtRotaStart = ""
        MainForm!txtRotaPeriod = ""
        MainForm!txtRotaStart.Enabled = True
        MainForm!cmdShowCalendar1.Enabled = True
    End If
    
    MainForm!txtRotaPeriod = GlobalParms.GetValue("DefaultRotaPeriod", "NumVal")
    MainForm!txtRotaEnd = ""
    'MainForm!txtRotaPeriod = ""
    
    SetContinueNewOpt = True

    Exit Function

ErrorTrap:

    EndProgram
    
End Function


Private Sub DisableForm()
Dim SPAMFormControl As Control


On Error GoTo ErrorTrap
   
    For Each SPAMFormControl In MainForm.Controls
        If (TypeOf SPAMFormControl Is TextBox Or _
            TypeOf SPAMFormControl Is ComboBox Or _
            TypeOf SPAMFormControl Is CommandButton Or _
            TypeOf SPAMFormControl Is OptionButton) And _
            SPAMFormControl.Name <> "cmdCancel" Then
                SPAMFormControl.Enabled = False
        End If
    Next

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Public Sub EnableForm()

Dim SPAMFormControl As Control


On Error GoTo ErrorTrap

    Screen.MousePointer = 0
    
    For Each SPAMFormControl In MainForm.Controls
        If (TypeOf SPAMFormControl Is TextBox Or _
            TypeOf SPAMFormControl Is ComboBox Or _
            TypeOf SPAMFormControl Is CommandButton Or _
            TypeOf SPAMFormControl Is OptionButton) Then
                
                If SPAMFormControl.Name <> "cmbCongregation" Then
                    SPAMFormControl.Enabled = True
                End If
                
        End If
    Next
    
    
    With MainForm
    Select Case True
    Case !optContinue
        !txtRotaStart.Enabled = False
        !cmdShowCalendar1.Enabled = False
        ContFromLastRota = True
        CalcRotaStartDate
    Case NewRota
        !txtRotaStart.Enabled = True
        !cmdShowCalendar1.Enabled = True
        ContFromLastRota = False
    End Select
    End With
    
    MainForm.afxProgressBar.value = 0
    MainForm!lblProgress.Visible = False
    
    Screen.MousePointer = 0 'Normal
    

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Public Sub AcquireRotaStructure(TempNoAttending As Integer, TempNoOnMics As Integer, TempNoOnSound As Integer, TempNoOnPlatform As Integer, FirstJobPos As Integer)
Dim i As Integer, FirstJobFound As Boolean

On Error GoTo ErrorTrap
    
'
'Structure of current tblRota may be different to that set on tblConstants (eg NoOnMics etc). Therefore,
' acquire the ACTUAL structure from tblRota itself. These values are then used later in the program.
'
    
    FirstJobFound = False 'init

    'get jobnames into array for fast access in 'GetJobNames' proc
    ReDim arrJobNames(rstRota.Fields.Count - 3)
    
    For i = 0 To rstRota.Fields.Count - 1
        If InStr(1, rstRota.Fields(i).Name, "Attendant") Then
            TempNoAttending = TempNoAttending + 1
            arrJobNames(i - 2) = "Attendant"
            FirstJobFound = True
        ElseIf InStr(1, rstRota.Fields(i).Name, "RovingMic") Then
            TempNoOnMics = TempNoOnMics + 1
            arrJobNames(i - 2) = "RovingMic"
            FirstJobFound = True
        ElseIf InStr(1, rstRota.Fields(i).Name, "Platform") Then
            TempNoOnPlatform = TempNoOnPlatform + 1
            arrJobNames(i - 2) = "Platform"
            FirstJobFound = True
        ElseIf InStr(1, rstRota.Fields(i).Name, "Sound") Then
            TempNoOnSound = TempNoOnSound + 1
            arrJobNames(i - 2) = "Sound"
            FirstJobFound = True
        End If
        If FirstJobFound = False Then
            FirstJobPos = i
        End If
    Next i
        
    If i > 0 Then
        FirstJobPos = FirstJobPos + 1
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    
    
    
End Sub

Public Sub AnyoneSuspended()
On Error GoTo ErrorTrap

    rstRota.MoveLast
    
    With rstIDWeight
    .MoveFirst
    
    Do Until .EOF
        If CongregationMember.SPAMBroSuspended2(!ID, rstRota!RotaDate) Then
            .Edit
            !Attendant_Wtg = 0
            !RovingMic_Wtg = 0
            !Sound_Wtg = 0
            !Platform_Wtg = 0
            .Update
        End If
        .MoveNext
    Loop
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub



Public Sub WhatRotaDates()

On Error GoTo ErrorTrap

    If MainForm!cmbStoredRotas.ListIndex = -1 Then
        MainForm!optNew = True
        SetContinueNewOpt
    End If

    Select Case True
    Case MainForm!optContinue
        MainForm!txtRotaStart.Enabled = False
        MainForm!cmdShowCalendar1.Enabled = False
        ContFromLastRota = True
        GlobalParms.Save "ContFromLastRota", "TrueFalse", True
        If MainForm!cmbStoredRotas.ListIndex > -1 Then
            SetUptblRota
            If CalcRotaStartDate Then
                MainForm!txtRotaStart = RotaStartDate
            Else
                MsgBox "Cannot continue from last rota. Please enter a new Start Date.", vbOKOnly + vbExclamation
                MainForm!txtRotaStart.Enabled = True
                MainForm!cmdShowCalendar1.Enabled = True
                TextFieldGotFocus MainForm!txtRotaStart, True
                ContFromLastRota = False
                MainForm!optNew.value = True
            End If
        Else
            MsgBox "Select the Rota you want to continue from", vbOKOnly + vbExclamation, AppName
            MainForm!txtRotaStart.Enabled = True
            MainForm!cmdShowCalendar1.Enabled = True
            'MainForm!cmbStoredRotas.SetFocus
            ContFromLastRota = False
            MainForm!optNew = True
        End If
           
    Case MainForm!optNew
        GlobalParms.Save "ContFromLastRota", "TrueFalse", False
        MainForm!txtRotaStart.Enabled = True
        MainForm!cmdShowCalendar1.Enabled = True
        ContFromLastRota = False
    End Select

    MainForm!txtRotaPeriod = GlobalParms.GetValue("DefaultRotaPeriod", "NumVal")

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Public Sub NotEnoughBrosMessage()

On Error GoTo ErrorTrap

    MsgBox "You don't have enough brothers to fill the current rota structure. " & _
           "If possible, reduce the number of brothers working each week. " & _
           "Also ensure that no brothers are suspended unnecessarily.", _
           vbOKOnly + vbExclamation, AppName

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Public Function InitialSetUp() As Boolean
Dim RotaStructureChanged As Boolean, ModifyConstants As clsApplicationConstants

On Error GoTo ErrorTrap

    '
    'Set Main form qualifier
    '
    Set MainForm = frmSoundAndPlatformRota
'    Set globalparms = New clsApplicationConstants
        
    MainForm!lblProgress.Visible = False
    
'    MainForm!cmdPrintForBros.Enabled = PrintUsingWord(False)
    
    '
    'Get all constants from tblConstants and put into host variables
    '
    GetConstants
    
    SaveRotaPeriod = False
    
    RotaStructureChanged = False
    
    RotaEdited = False
    
    Set ModifyConstants = New clsApplicationConstants
    
    '
    'Build Name&Address recordset
    '
    If Not BuildMainRecordSets Then
        InitialSetUp = False
        Exit Function
    End If
       
    If Not gbPrintExistingRota Then
        HandleListBox.PopulateListBox MainForm!cmbStoredRotas, "SELECT SeqNum, DisplayForCombo FROM tblStoredSPAMRotas ORDER BY ModifiedDateTime DESC" _
                                , CMSDB, 0, "", False, 1
    End If
    
    SetCmbSPAMRotas
    
    SetContinueNewOpt
    
    HandleListBox.PopulateListBox MainForm!cmbCongregation, "SELECT CongName, CongNo FROM tblCong", CMSDB, 1, "", False, 0
    HandleListBox.SelectItem MainForm!cmbCongregation, CLng(DefaultCong)
        
    InitialSetUp = True

    Exit Function

ErrorTrap:

    InitialSetUp = False

        
End Function




Private Sub GetNumberOfBrothers()
On Error GoTo ErrorTrap
Dim rstCounter As Recordset

    Set rstCounter = CMSDB.OpenRecordset("SELECT COUNT(*) AS CountSound " & _
                                         "FROM tblTaskAndPerson " & _
                                        " WHERE Task = 57 ", dbOpenForwardOnly)

    TotalSoundAttendants = rstCounter!CountSound

    Set rstCounter = CMSDB.OpenRecordset("SELECT COUNT(*) AS CountPlatform " & _
                                         "FROM tblTaskAndPerson " & _
                                        " WHERE Task = 58 ", dbOpenForwardOnly)

    TotalPlatformAttendants = rstCounter!CountPlatform
    
    Set rstCounter = CMSDB.OpenRecordset("SELECT COUNT(*) AS CountAttendants " & _
                                         "FROM tblTaskAndPerson " & _
                                        " WHERE Task = 60 ", dbOpenForwardOnly)

    TotalAttendants = rstCounter!CountAttendants
    
    Set rstCounter = CMSDB.OpenRecordset("SELECT COUNT(*) AS CountMics " & _
                                         "FROM tblTaskAndPerson " & _
                                        " WHERE Task = 59 ", dbOpenForwardOnly)

    TotalMicHandlers = rstCounter!CountMics
    
    rstCounter.Close
    
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
