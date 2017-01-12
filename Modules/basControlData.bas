Attribute VB_Name = "basControlData"
Option Explicit

Public Function GetCongNoFromSeqNo(SeqNo As Long) As Long
Dim rs As Recordset

    On Error GoTo ErrorTrap
    
    Set rs = CMSDB.OpenRecordset("SELECT CongNo FROM tblCong WHERE SeqNo = " & SeqNo, dbOpenForwardOnly)
    
    With rs
    If Not .BOF Then
        GetCongNoFromSeqNo = !CongNo
    Else
        GetCongNoFromSeqNo = 0
    End If
    End With
    
    rs.Close
    Set rs = Nothing
    Exit Function
ErrorTrap:
    EndProgram
    
    
End Function
Public Function GetCongSeqNoFromCongNo(CongNo As Long) As Long
Dim rs As Recordset

    On Error GoTo ErrorTrap
    
    Set rs = CMSDB.OpenRecordset("SELECT SeqNo FROM tblCong WHERE CongNo = " & CongNo, dbOpenForwardOnly)
    
    With rs
    If Not .BOF Then
        GetCongSeqNoFromCongNo = !SeqNo
    Else
        GetCongSeqNoFromCongNo = 0
    End If
    End With
    
    rs.Close
    Set rs = Nothing
    Exit Function
ErrorTrap:
    EndProgram
    
    
End Function
Public Function CreateWeightingsTable() As Boolean
Dim ErrCode As Integer

    On Error GoTo ErrorTrap
    
    Set rstBroList = Nothing
    
    DelAllRows "tblIDWeightings"
    
    If rstIDWeight Is Nothing Then
        Set rstIDWeight = CMSDB.OpenRecordset("tblIDWeightings", dbOpenDynaset)
    Else
        rstIDWeight.Requery
    End If
        
    CreateWeightingsTable = True

    Exit Function
ErrorTrap:
    EndProgram
    
    
End Function

Public Function CalcInitWtngs() As Boolean
Dim rstAllSPAMBros As Recordset, rstSumWeightings As Recordset
Dim str As String, n As Integer, TempCount As Integer, strSQL As String

    On Error GoTo ErrorTrap
        
    '
    'Populate Bro's-Weightings table with IDs and zero weighting for each bro
    '

    
    strSQL = "SELECT * " & _
             "FROM tblNameAddress " & _
             "WHERE ID IN (SELECT Person " & _
                          "FROM tblTaskAndPerson " & _
                          "WHERE CongNo = " & GlobalDefaultCong & _
                          " AND TaskCategory = 6 " & _
                          " AND ((TaskSubCategory = 10 AND Task IN (57, 58, 59)) " & _
                          " OR (TaskSubCategory = 11 AND Task = 60))) "

             
    Set rstAllSPAMBros = CMSDB.OpenRecordset(strSQL, dbOpenDynaset)
    
    If Not rstAllSPAMBros.BOF Then
        rstAllSPAMBros.MoveFirst
        
        With rstIDWeight
        Do Until rstAllSPAMBros.EOF
        
            .AddNew
            !ID = rstAllSPAMBros!ID
            !Weighting = 0
            !ResponsibilityWeighting = 0
            !Attendant_Wtg = 0
            !RovingMic_Wtg = 0
            !Sound_Wtg = 0
            !Platform_Wtg = 0
            !Personal_Wtg = 0
            !OnThisWeek = 0
            !NoOfTimesOnRota = 0
            .Update
            
            rstAllSPAMBros.MoveNext
            
        Loop
        End With
    Else
        MsgBox "You must enter details of brothers in congregation", vbOKOnly + vbExclamation, AppName
        CalcInitWtngs = False
        Exit Function
    End If
                
        
    'Calculate each Bro's personal weighting (Based on family circumstances etc)
    '
    If Not CalcPersonalWtgs Then
        CalcInitWtngs = False
        Exit Function
    End If
        
        rstAllSPAMBros.MoveFirst
        
   'Acquire all responsibility weightings for brother, sum them and update tblIDWeightings.
   '
        Do
            strSQL = "SELECT SUM(tblTaskWeightings.SPAMRotaWeighting) AS SumOfWeighting " & _
                     "FROM tblTaskAndPerson " & _
                     "INNER JOIN tblTaskWeightings ON " & _
                     "(tblTaskAndPerson.TaskSubCategory = tblTaskWeightings.TaskSubCategory) " & _
                     "AND (tblTaskAndPerson.TaskCategory = tblTaskWeightings.TaskCategory) " & _
                     "AND (tblTaskAndPerson.Task = tblTaskWeightings.Task) " & _
                     "WHERE Person = " & rstAllSPAMBros!ID
            
            Set rstSumWeightings = CMSDB.OpenRecordset(strSQL, dbOpenDynaset)
    
            With rstIDWeight
            .FindFirst ("ID = " & rstAllSPAMBros!ID)
            If Not .BOF Then
                .Edit
                !ResponsibilityWeighting = (RespCoeff * rstSumWeightings!SumOfWeighting)
                .Update
            End If
            End With
            rstAllSPAMBros.MoveNext
        Loop Until rstAllSPAMBros.EOF
            
    

    Set rstAllSPAMBros = Nothing
    
    CalcInitWtngs = True
    
    Exit Function
    

ErrorTrap:
    EndProgram
    
   
End Function




Public Function GetConstants() As Boolean
Dim rstTemp As Recordset, GetConst As clsApplicationConstants

    On Error GoTo ErrorTrap
      
    Set rstConstants = CMSDB.OpenRecordset("tblConstants", dbOpenDynaset)
    
    With rstConstants
    'TO DO: Must convert all these FindFirsts to use GetConst class instead.... Idiot....
    
    .FindFirst "FldName = 'NoOnSound'"
    NoOnSound = !NumVal
    
    .FindFirst "FldName = 'NoOnPlatform'"
    NoOnPlatform = !NumVal
    
    .FindFirst "FldName = 'NoOnMics'"
    NoOnMics = !NumVal
    
    .FindFirst "FldName = 'NoAttending'"
    NoAttending = !NumVal
    
        
        
    .FindFirst "FldName = 'MaxWkWting'"
    MaxWkWting = !NumFloat
    
    .FindFirst "FldName = 'WkWtingDiff'"
    WkWtingDiff = !NumFloat
    
    .FindFirst "FldName = 'Pos1stJobOnRota'"
    Pos1stJobOnRota = !NumVal
    
    .FindFirst "FldName = 'NoOfWeeks'"
    NoOfWeeks = !NumVal
    
    .FindFirst "FldName = 'Att_Weight'"
    Att_Weight = !NumVal
    
    .FindFirst "FldName = 'Mic_Weight'"
    Mic_Weight = !NumVal
    
    .FindFirst "FldName = 'Snd_Weight'"
    Snd_Weight = !NumVal
    
    .FindFirst "FldName = 'Plat_Weight'"
    Plat_Weight = !NumVal

    .FindFirst "FldName = 'MaxConsecutiveRovingMic'"
    MaxConsecutiveRovingMic = !NumVal
    
    .FindFirst "FldName = 'MaxConsecutiveAttendant'"
    MaxConsecutiveAttendant = !NumVal
   
    .FindFirst "FldName = 'MaxConsecutiveSound'"
    MaxConsecutiveSound = !NumVal
   
    .FindFirst "FldName = 'MaxConsecutivePlatform'"
    MaxConsecutivePlatform = !NumVal
   
    .FindFirst "FldName = 'MaxConsecutiveUpperLimit'"
    MaxConsecutiveUpperLimit = !NumVal
   
    .FindFirst "FldName = 'WksToCheckBack'"
    WksToCheckBack = !NumVal
     
    .FindFirst "FldName = 'ParentCoeff'"
    ParentCoeff = !NumFloat
    
    .FindFirst "FldName = 'ZeroWtgAge'"
    ZeroWtgAge = !NumVal
    
    .FindFirst "FldName = 'LoneParentFactor'"
    LoneParentFactor = !NumFloat
     
    .FindFirst "FldName = 'Marriage_Wtg'"
    Marriage_Wtg = !NumFloat
     
    .FindFirst "FldName = 'ContFromLastRota'"
    ContFromLastRota = !TrueFalse
     
    .FindFirst "FldName = 'RespCoeff'"
    RespCoeff = !NumFloat
    
     
    If CongregationIsSetUp Then
        .FindFirst "FldName = 'DefaultCong'"
        DefaultCong = !NumVal
    End If
     
    .FindFirst "FldName = 'PersonalWeightCoeff'"
    PersonalWeightCoeff = !NumFloat
     
    End With
    
    SundayMeetingDay = GlobalParms.GetValue("SundayMeetingDay", "AlphaVal")
    MidWeekMeetingDay = GlobalParms.GetValue("MidWeekMeetingDay", "AlphaVal")
    SundayMeetingDayNo = GlobalParms.GetValue("SundayMeetingDay", "NumVal")
    MidWeekMeetingDayNo = GlobalParms.GetValue("MidWeekMeetingDay", "NumVal")
    
    MinOnMics = GlobalParms.GetValue("MinOnMics", "NumVal")
    MinOnSound = GlobalParms.GetValue("MinOnSound", "NumVal")
    MinOnPlat = GlobalParms.GetValue("MinOnPlat", "NumVal")
    MinOnAtt = GlobalParms.GetValue("MinOnAtt", "NumVal")
    RotaTimesPerWeek = GlobalParms.GetValue("RotaTimesPerWeek", "NumVal")
    PTObserved = GlobalParms.GetValue("ObservePartTimers", "TrueFalse")
    
    '
    'Calculate number of slots on Rota.
    '
    SlotCount = NoOnMics + NoAttending + NoOnSound + NoOnPlatform
    
    
    
    '
    'Find out how many brothers there are
    '
    'Set rstTemp = CMSDB.OpenRecordset("SELECT * " & _
     '                                       "FROM tblResponsibilities " _
      '                                         , dbOpenDynaset)

    
    'With rstTemp
    '.MoveLast    'Done to acquire RecordCount
    '.MoveFirst   '
    'BroCount = .RecordCount
    'End With
    
    Set rstConstants = Nothing
    
    
    GetConstants = True

    Exit Function
ErrorTrap:
    EndProgram
    
   
End Function

Public Function CalcNumOfWks() As Boolean
'
'Calculate Number-of-Weeks on Rota (if an end date has been specified on form). This
'will then override value set in tblConstants.
'
'Also validates form fields
'
    
    On Error GoTo ErrorTrap
    
    '
    'First check Rota Start field has been entered
    '
    If Len(MainForm!txtRotaStart) = 0 Or IsNull(MainForm!txtRotaStart) Then
        Beep
        MsgBox "You must enter a Rota Start Date", vbOKOnly + vbExclamation, AppName
        CalcNumOfWks = False
        TextFieldGotFocus MainForm!txtRotaStart, True
        Screen.MousePointer = 0
        SaveRotaPeriod = False
        Exit Function
    Else
        MainForm!txtRotaStart = Trim(MainForm!txtRotaStart)
        MainForm!txtRotaStart = Format(MainForm!txtRotaStart, "dd/mm/yyyy")
                
        If ValidDate(MainForm!txtRotaStart) Then
            RotaStartDate = Format(MainForm!txtRotaStart, "dd/mm/yy")
        Else
            MsgBox "The Rota Start Date is invalid", vbExclamation + vbOKOnly, AppName
            If MainForm!txtRotaStart.Enabled Then
               TextFieldGotFocus MainForm!txtRotaStart, True
            End If
            Screen.MousePointer = 0
            CalcNumOfWks = False
            SaveRotaPeriod = False
            Exit Function
        End If
    End If
    
    If Len(MainForm!txtRotaEnd) <> 0 _
       And Not IsNull(MainForm!txtRotaEnd) Then      'Rota End date field populated
       
        If Len(MainForm!txtRotaPeriod) <> 0 _
         And Not IsNull(MainForm!txtRotaPeriod) Then
         '
         'Allow the population of only End date OR Period
         '
            MsgBox "Do not enter End Date AND Period", vbExclamation + vbOKOnly, AppName
            TextFieldGotFocus MainForm!txtRotaEnd, True
            CalcNumOfWks = False
            Screen.MousePointer = 0
            SaveRotaPeriod = False
            Exit Function
        End If
       
        MainForm!txtRotaEnd = Trim(MainForm!txtRotaEnd)
        MainForm!txtRotaEnd = Format(MainForm!txtRotaEnd, "dd/mm/yyyy")
        
        '
        'Check txtRotaEnd contains valid date
        '
        If IsDate(MainForm!txtRotaEnd) And ValidDate(MainForm!txtRotaEnd) Then
            '
            'Calculate number of weeks
            '
            RotaEndDate = Format(MainForm!txtRotaEnd, "dd/mm/yy")
            
            If RotaEndDate <= RotaStartDate Then
                MsgBox "The Rota End Date must be greater than the Start Date", vbExclamation + vbOKOnly, AppName
                TextFieldGotFocus MainForm!txtRotaEnd, True
                CalcNumOfWks = False
                Screen.MousePointer = 0
                SaveRotaPeriod = False
                Exit Function
            Else
                MainForm!txtRotaEnd = RotaEndDate
                
                NoOfWeeks = (RotaEndDate - RotaStartDate) / 7
                '
                'Force txtRotaPeriod to display calculated value
                ' Add 1 to display no-of-wks actually printed on Rota
                '
                MainForm!txtRotaPeriod = NoOfWeeks + 1
            End If

        Else
            MsgBox "The Rota End Date is invalid", vbExclamation + vbOKOnly, AppName
            TextFieldGotFocus MainForm!txtRotaEnd, True
            CalcNumOfWks = False
            Screen.MousePointer = 0
            SaveRotaPeriod = False
            Exit Function
        End If
        
    ElseIf Len(MainForm!txtRotaPeriod) <> 0 _
         And Not IsNull(MainForm!txtRotaPeriod) Then
        '
        'If txtRotaPeriod populated (but txtRotaEnd isn't)
        '
        MainForm!txtRotaPeriod = Trim(MainForm!txtRotaPeriod)
        
        If IsNumeric(MainForm!txtRotaPeriod) Then
            If CInt(MainForm!txtRotaPeriod) = 0 Then
                MsgBox "The Rota Period must be greater than zero", vbExclamation + vbOKOnly, AppName
                TextFieldGotFocus MainForm!txtRotaPeriod, True
                CalcNumOfWks = False
                Screen.MousePointer = 0
                SaveRotaPeriod = False
                Exit Function
            End If
            
            NoOfWeeks = MainForm!txtRotaPeriod
            
            RotaEndDate = RotaStartDate + (7 * (NoOfWeeks - 1))
            MainForm!txtRotaEnd = RotaEndDate  'update form
        Else
            MsgBox "The Rota Period must be numeric", vbExclamation + vbOKOnly, AppName
            TextFieldGotFocus MainForm!txtRotaPeriod, True
            CalcNumOfWks = False
            Screen.MousePointer = 0
            SaveRotaPeriod = False
            Exit Function
        End If
    Else
        MsgBox "You must enter either a Rota End Date or Number of Weeks", vbExclamation + vbOKOnly, AppName
        CalcNumOfWks = False
        TextFieldGotFocus MainForm!txtRotaPeriod, True
        Screen.MousePointer = 0
        SaveRotaPeriod = False
        Exit Function
    End If

    MainForm!txtRotaEnd = ""  'update form
    
    If RotaTimesPerWeek = 2 Then
        If Format(RotaStartDate, "DDDD") <> SundayMeetingDay And Format(RotaStartDate, "DDDD") <> MidWeekMeetingDay Then
        'Work out which meeting day corresponds to RotaStartDate
            MsgBox "The Start Date does not correspond with either of your meeting days!", vbExclamation + vbOKOnly, AppName
            CalcNumOfWks = False
            TextFieldGotFocus MainForm!txtRotaStart, True
            Screen.MousePointer = 0
            SaveRotaPeriod = False
            Exit Function
        ElseIf Format(RotaStartDate, "dddd") = SundayMeetingDay Then
            RotaStartDay = SundayMeetingDay
        Else
    
        End If
    Else
        MainForm!txtRotaStart = CStr(GetDateOfGivenDay(CDate(MainForm!txtRotaStart), vbMonday))
        RotaStartDate = Format(MainForm!txtRotaStart, "dd/mm/yy")
        RotaStartDay = "Monday"
    End If
        
    If SaveRotaPeriod Then
        GlobalParms.Save "DefaultRotaPeriod", "NumVal", MainForm!txtRotaPeriod
        SaveRotaPeriod = False
    End If
    
    DoEvents
    
    CalcNumOfWks = True
    
    Exit Function
    

ErrorTrap:
    EndProgram
    
End Function

Public Function CalcPersonalWtgs() As Boolean
'
'Calculate each Bro's personal weighting (Based on family circumstances etc)
'
        
        On Error GoTo ErrorTrap

        With rstIDWeight
        
        .MoveFirst
        '
        'Move through each IDWeighting
        '
        Do Until .EOF
            .Edit
            !Personal_Wtg = CompiledWtg(!ID)    'Call function
            .Update
            .MoveNext
        Loop

        End With

    CalcPersonalWtgs = True
    
    Exit Function

ErrorTrap:
    EndProgram
    

End Function

Public Function CompiledWtg(BroId As Integer) As Double
Dim rstTemp As Recordset, SQLStr As String, AgeOfChild As Single
Dim ParentWtg As Single, MarriageWtg As Single, tempvar As Single, InfirmityWtg As Double
Dim Person As clsCongregationMember


On Error GoTo ErrorTrap

    Set Person = New clsCongregationMember

    CompiledWtg = 0
    ParentWtg = 0

'
'Calculate weighting due to Children - uses an inverted quadratic
'
    
    SQLStr = "SELECT * " & _
             "FROM tblChildren " & _
             "WHERE Parent = " & BroId

    Set rstTemp = CMSDB.OpenRecordset(SQLStr, dbOpenDynaset)
    
    With rstTemp
    
    If Not .BOF Then    'Find Parent-Wtg if bro is a dad.
        .MoveFirst
        '
        'Look at age of each child of this BroID. Use quadratic to compute wtg.
        '
        Do Until .EOF
            If Not CongregationMember.IsActive(!Child) Then
                tempvar = 0
            Else
                AgeOfChild = (date - CongregationMember.DateOfBirth(!Child)) / 365
                tempvar = -1 * ParentCoeff * (AgeOfChild ^ 2 - ZeroWtgAge ^ 2)
                '
                'stop quadratic going negative
    
                If tempvar < 0 Then
                    tempvar = 0
                End If
            End If
            
            ParentWtg = ParentWtg + tempvar
            .MoveNext
        Loop
    Else
        ParentWtg = 0
    End If
       
    .Close
       
    End With
    
    
'
'Calculate Weighting due to Marriage
'
    SQLStr = "SELECT * " & _
             "FROM tblMarriage " & _
             "WHERE ID = " & BroId

    Set rstTemp = CMSDB.OpenRecordset(SQLStr, dbOpenDynaset)

    If Not rstTemp.BOF Then    'Bro is married
        If CongregationMember.IsActive(rstTemp!Spouse) Then
            MarriageWtg = Marriage_Wtg
            If rstTemp.RecordCount > 1 Then   'Polygamy???
                MsgBox "Brother number " & BroId & " has more than one wife!", vbOKOnly, "How odd...."
            End If
        Else
            MarriageWtg = 0
            '
            'Bro is single - multiply his Parent-Wtg by the LoneParentFactor
            '
            ParentWtg = (LoneParentFactor + 1) * ParentWtg
        End If
    Else
        MarriageWtg = 0
        '
        'Bro is single - multiply his Parent-Wtg by the LoneParentFactor
        '
        ParentWtg = (LoneParentFactor + 1) * ParentWtg
    End If
                

    InfirmityWtg = GlobalParms.GetValue("InfirmityLevelWtg", "NumFloat") * Person.InfirmityLevel(BroId)

    '
    'Produce final Personal Wtg to return to calling function
    '
    CompiledWtg = PersonalWeightCoeff * (ParentWtg + MarriageWtg + InfirmityWtg)
    
    rstTemp.Close
       
    Exit Function


ErrorTrap:
    EndProgram
    
End Function

Public Function BuildMainRecordSets() As Boolean

    On Error GoTo ErrorTrap

    'Set rstRota = CMSDB.OpenRecordset("tblRota", dbOpenDynaset)
    

    Set rstNameAddress = CMSDB.OpenRecordset("tblNameAddress", dbOpenDynaset)

    BuildMainRecordSets = True
    
    Exit Function
    

ErrorTrap:
    EndProgram
    
End Function


Public Function CalcRotaStartDate() As Boolean
    
    On Error GoTo ErrorTrap
    
    If TableExists("tblRota") Then
        With rstRota
        If Not .BOF Then
            .MoveLast
            RotaStartDate = CalculateNextRotaDate(!RotaDate)
'            RotaStartDate = !RotaDate + 7
            CalcRotaStartDate = True
        Else
            CalcRotaStartDate = False
        End If
        End With
    Else
        CalcRotaStartDate = False
    End If
        
    Exit Function
    

ErrorTrap:
    EndProgram
    
End Function


Public Function CreateNewRotaTables() As Boolean
Dim ErrCode As Integer, i As Integer

    On Error GoTo ErrorTrap
    
    Set rstRota = Nothing
    
    '
    'Recreate rota tables' structure & fields
    '
    If Not DeleteTable("tblRota") Then
        CreateNewRotaTables = False
        Exit Function
    End If
    

    If Not CreateTable(ErrCode, "tblRota", "RotaDate", "DATE", , "") Then
        CreateNewRotaTables = False
        Exit Function
    End If
    
    
    
    For i = 1 To NoAttending
        If Not CreateField(ErrCode, "tblRota", "Attendant_" & Format(i, "00"), "SHORT", , "") Then
            CreateNewRotaTables = False
            Exit Function
        End If
    Next i
        
    For i = 1 To NoOnMics
        If Not CreateField(ErrCode, "tblRota", "RovingMic_" & Format(i, "00"), "SHORT", , "") Then
            CreateNewRotaTables = False
            Exit Function
        End If
    Next i
    
    For i = 1 To NoOnSound
        If Not CreateField(ErrCode, "tblRota", "Sound_" & Format(i, "00"), "SHORT", , "") Then
            CreateNewRotaTables = False
            Exit Function
        End If
    Next i
    
    For i = 1 To NoOnPlatform
        If Not CreateField(ErrCode, "tblRota", "Platform_" & Format(i, "00"), "SHORT", , "") Then
            CreateNewRotaTables = False
            Exit Function
        End If
    Next i
    
    '
    'Now create associated recordset
    '
    Set rstRota = CMSDB.OpenRecordset("tblRota", dbOpenDynaset)
    
    
    CreateNewRotaTables = True
    
    Exit Function
    

ErrorTrap:
    EndProgram
    
    
End Function

Public Function GetWeekDay(TheDay As String) As Long
    Select Case TheDay
    Case "Sunday"
        GetWeekDay = 1
    Case "Monday"
        GetWeekDay = 2
    Case "Tuesday"
        GetWeekDay = 3
    Case "Wednesday"
        GetWeekDay = 4
    Case "Thursday"
        GetWeekDay = 5
    Case "Friday"
        GetWeekDay = 6
    Case "Saturday"
        GetWeekDay = 7
    End Select
            
End Function
Public Function GetDayName(DayNo As Long) As String

    Select Case DayNo
    Case 1
        GetDayName = "Sunday"
    Case 2
        GetDayName = "Monday"
    Case 3
        GetDayName = "Tuesday"
    Case 4
        GetDayName = "Wednesday"
    Case 5
        GetDayName = "Thursday"
    Case 6
        GetDayName = "Friday"
    Case 7
        GetDayName = "Saturday"
    End Select
            
End Function

Public Sub SetCmbSPAMRotas()
Dim rstTemp As Recordset, strSQL As String, TempDate As Date

On Error GoTo ErrorTrap

    If Not gbPrintExistingRota Then
    
        strSQL = "SELECT max(ModifiedDateTime) as MaxDate " & _
                 "FROM tblStoredSPAMRotas "
                 
        Set rstTemp = CMSDB.OpenRecordset(strSQL, dbOpenDynaset)
        
        If Not rstTemp.BOF And Not IsNull(rstTemp!MaxDate) Then
            TempDate = rstTemp!MaxDate
            strSQL = "SELECT * " & _
                     "FROM tblStoredSPAMRotas "
                     
            Set rstTemp = CMSDB.OpenRecordset(strSQL, dbOpenDynaset)
            
            rstTemp.FindFirst "ModifiedDateTime = #" & Format(TempDate, "mm/dd/yy hh:mm:ss") & "#"
            'rstTemp.FindFirst "ModifiedDateTime = " & TempDate
            
            HandleListBox.SelectItem MainForm!cmbStoredRotas, rstTemp!SeqNum
                
        Else
            CreateNewRotaTables
        End If
        
    Else
        MainForm.cmbStoredRotas_Click
    End If
    
    On Error Resume Next
    rstTemp.Close
    Set rstTemp = Nothing

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Public Sub SetUptblRota()
Dim rstSPAMRotas As Recordset

On Error GoTo ErrorTrap


    If MainForm!cmbStoredRotas.ListIndex > -1 Then
        Set rstRota = Nothing
        Set rstRotaForEdit = Nothing
           
        Set rstSPAMRotas = CMSDB.OpenRecordset("SELECT * FROM tblStoredSPAMRotas", dbOpenDynaset)
    
        rstSPAMRotas.FindFirst "SeqNum = " & MainForm!cmbStoredRotas.ItemData(MainForm!cmbStoredRotas.ListIndex)
        
        If TableExists(rstSPAMRotas!RotaTableName) Then
            CopyTable "tblRota", rstSPAMRotas!RotaTableName, CMSDB
            'DoCmd.CopyObject , "tblRota", acTable, rstSPAMRotas!RotaTableName
            Set rstRota = CMSDB.OpenRecordset("tblRota", dbOpenDynaset)
        Else
            MsgBox MainForm!cmbStoredRotas.text & " do not exist. You should now select another rota to " & _
            "edit/print, or create a new one.", vbExclamation + vbOKOnly, AppName
            CreateNewRotaTables
            ContFromLastRota = False
            SetContinueNewOpt
        End If
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Public Function ConnectToDAO(PathToMDB As String) As Boolean
    
    On Error Resume Next
        
    Set CMSDB = OpenDatabase(PathToMDB, False, False, ";pwd=" & TheDBPassword) ', dbLangGeneral)
    
    If Err.number = 0 Then
        ConnectToDAO = True
    Else
        ConnectToDAO = False
    End If

End Function
Public Function ConnectToDAO_ForUpdateDB(PathToMDB As String) As Boolean
    
    On Error Resume Next
    
    Set CMS_Update_DB = OpenDatabase(PathToMDB, False, False, ";pwd=" & TheDBPassword) ', dbLangGeneral)
    
    If Err.number = 0 Then
        ConnectToDAO_ForUpdateDB = True
    Else
        ConnectToDAO_ForUpdateDB = False
    End If

End Function
Public Function CreateAccessDatabase(DBObjName As Database, _
                                     DBFilePath As String, _
                                     Password As String, _
                                     Encrypt As Boolean) As Boolean
    
    On Error Resume Next
    
    If Encrypt Then
        Set DBObjName = CreateDatabase(DBFilePath, dbLangGeneral & ";pwd=" & Password, dbEncrypt)
    Else
        Set DBObjName = CreateDatabase(DBFilePath, dbLangGeneral & ";pwd=" & Password)
    End If
    
    If Err.number = 0 Then
        CreateAccessDatabase = True
    Else
        CreateAccessDatabase = False
    End If

End Function
Public Sub SetUpGlobalObjects(Optional bExcludeFSO As Boolean = False)
On Error GoTo ErrorTrap
    
    Set GlobalParms = Nothing
    Set GlobalCalendar = Nothing
    Set CongregationMember = Nothing
    Set TheTMS = Nothing
    
    Set GlobalParms = New clsApplicationConstants
    Set GlobalCalendar = New clsCalendar
    Set CongregationMember = New clsCongregationMember
    Set TheTMS = New clsTMS
    
    If Not bExcludeFSO Then
        Set gFSO = New FileSystemObject
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram


End Sub


Public Sub CheckIfBrothersInDB()
    On Error GoTo ErrorTrap
    '
    'Are there any brothers held in the DB? If so, enable certain functions
    '
    SystemContainsPeople = GlobalParms.GetValue("PeopleAreInTheSystem", "TrueFalse")
    
    With frmMainMenu
    
    .cmdOpenSPAMRota.Enabled = SystemContainsPeople
    .cmdOpenTMS.Enabled = SystemContainsPeople
    .cmdOpenCalendar.Enabled = SystemContainsPeople
    .cmdServiceMtg.Enabled = SystemContainsPeople
    .cmdAccounts.Enabled = SystemContainsPeople
    .cmdBookGroups.Enabled = SystemContainsPeople
    .cmdFieldMinistry.Enabled = SystemContainsPeople
    .cmdMeetingAttendance.Enabled = SystemContainsPeople
    .cmdOpenHallCleaning.Enabled = SystemContainsPeople
    .cmdPublicMeeting.Enabled = SystemContainsPeople

    .cmdOpenPersonalDetails.Enabled = CongregationIsSetUp
    
    End With
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Sub CheckIfCongSetUp()
    On Error GoTo ErrorTrap
Dim ctl As Control
    '
    'Has cong been set up? If so, enable certain functions
    '
    CongregationIsSetUp = GlobalParms.GetValue("CongregationIsSetUp", "TrueFalse")
    SystemContainsPeople = GlobalParms.GetValue("PeopleAreInTheSystem", "TrueFalse")
        
    frmMainMenu.cmdSetUp.Enabled = True
    
    For Each ctl In frmMainMenu.Controls
        If TypeOf ctl Is CommandButton Then
            If ctl.Name <> "cmdSetUp" Then
                ctl.Enabled = False
            End If
        End If
    Next
        
    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Public Sub LoadApp()
On Error GoTo ErrorTrap
    
Dim FreeFileNum As Integer, FileIsOpen As Boolean, arr() As String
Dim lDBSize As Long, dDBSizeMB As Double
Dim sOldPath As String
Dim sVer1 As String, sVer2 As String, sVer3 As String, sVerDB As String
Dim sVerApp As String, str As String
Dim Pos1stDot As Long, Pos2ndDot As Long
Dim bConnOK As Boolean, sDocLocation As String
Dim sCMSInitFilePath As String

'
'Acquire mdb name from CMS_Init.txt, then open the specified DB
'
'File record format is "<mdb name>|<Password>"
'

    Set gFSO = New FileSystemObject
    
    gsLogPath = SpecialFolder(SpecialFolder_AppData) & "\" & gsAppNameFull & "\Logs"
    
    JustTheDirectory = SpecialFolder(SpecialFolder_AppData) & "\" & gsAppNameFull
'    If Not gFSO.FolderExists(JustTheDirectory) Then
'        gFSO.CreateFolder JustTheDirectory
'    End If
        
    
    
    If Not gFSO.FolderExists(gsLogPath) Then
        gFSO.CreateFolder gsLogPath
    End If
    
    'now set up a new log file
    str = gsLogPath & "\CMS - " & Replace(Replace(Now, "/", "-"), ":", "-") & ".log"
    Set tsLogFileTextStream = gFSO.OpenTextFile(str, ForAppending, True)
    
    
    WriteToLogFile "Initialising CMS..."
    
    WriteToLogFile "*** "
    WriteToLogFile "**********************************************"
    WriteToLogFile "*** "
    WriteToLogFile "*** CMS version " & App.major & "." & App.minor & "." & App.Revision
    WriteToLogFile "*** "
    WriteToLogFile "**********************************************"
    WriteToLogFile "*** "




    'delete temp folder now before any processes
    ' can get hold of files within it!
    DeleteTempFolder
    
    WriteToLogFile "Deleted Temp folder"
    
    '
    'First check that another instance of the app isn't already running...
    '
    If App.PrevInstance Then
        MsgBox "C.M.S. is already running on this computer.", vbOKOnly + vbExclamation, AppName
        End
    End If
    
    FreeFileNum = FreeFile() 'Get next available file number
    
    On Error Resume Next
    Err.Clear
    If gFSO.FileExists(App.Path & "\CMS_Init.txt") Then
        Open App.Path & "\CMS_Init.txt" For Input As #FreeFileNum
        JustTheDirectory = App.Path
    Else
        Open JustTheDirectory & "\CMS_Init.txt" For Input As #FreeFileNum
    End If
    
    
    If Err.number <> 0 Then
        WriteToLogFile "Could not open CMS_Init.txt"
        MsgBox "Could not open CMS_Init.txt", vbOKOnly + vbCritical, "C.M.S. Cannot Load"
        End
    Else
        FileIsOpen = True
    End If
    
    WriteToLogFile "Opened init file"
    
    On Error GoTo ErrorTrap
    
'    '
'    'Load entire file into string variable
'    '
'    TheMDBFile = Input(LOF(FreeFileNum), FreeFileNum) 'LOF = Length Of File (ie, number of chars)
'    TheMDBFileAndExt = TheMDBFile & ".mdb"
    
    '
    'Read through CMS_Init.txt until FIRST record found not beginning with "--"
    '
    Line Input #FreeFileNum, TheMDBFile
    
    WriteToLogFile "Read first record - '" & TheMDBFile & "'"
    
    Do While (Not EOF(FreeFileNum)) And Left$(TheMDBFile, 2) = "--"
        Line Input #FreeFileNum, TheMDBFile
        WriteToLogFile "Read next record - '" & TheMDBFile & "'"
        If TheMDBFile = "" Then
            MsgBox "CMS_Init.txt is not valid.", vbOKOnly + vbCritical, "C.M.S. Cannot Load"
            End
        End If
    Loop
    
    If TheMDBFile = "" Then
'    If EOF(FreeFileNum) And TheMDBFile = "" Then
        MsgBox "CMS_Init.txt is not valid.", vbOKOnly + vbCritical, "C.M.S. Cannot Load"
        End
    End If
    
    Close FreeFileNum
    
    WriteToLogFile "Read ini file - complete"
    
    arr() = Split(TheMDBFile, "|")
    
    WriteToLogFile "Split ini file rec"
    
    TheMDBFileAndExt = arr(0) & ".cms"
    
    '
    'Derive DB password from CMS_Init file
    '
    If arr(0) = "CMS_DEV" Then
        TheDBPassword = arr(1)
    Else
        TheDBPassword = "cmsLIVEdb01"
    End If
    
    WriteToLogFile "Got password"
    
    TheMDBFile = arr(0)
    
    
    WriteToLogFile "App Data path: '" & JustTheDirectory & "'"
    
    CompletePathToTheMDBFileAndExt = JustTheDirectory & "\" & TheMDBFileAndExt
    
           
    WriteToLogFile "CompletePathToTheMDBFileAndExt: '" & CompletePathToTheMDBFileAndExt & "'"
    WriteToLogFile "JustTheDirectory: '" & JustTheDirectory & "'"
    
    bConnOK = ConnectToDAO(CompletePathToTheMDBFileAndExt)
    
    If Not bConnOK Then
        MsgBox "Could not connect to database. CMS will close."
        End
    End If
    
    
    WriteToLogFile "Opened database " & CompletePathToTheMDBFileAndExt
    
    '
    'Now set up global objects
    '
    SetUpGlobalObjects
    
    WriteToLogFile "Global objects set up"
                    
    '
    'Check that DB version matches exe version....
    '
    gstrDBVersion = GlobalParms.GetValue("CMS_Version", "AlphaVal")
    
       
    Pos1stDot = InStr(1, gstrDBVersion, ".")
    Pos2ndDot = InStr(Pos1stDot + 1, gstrDBVersion, ".")
    sVer1 = Format(Left(gstrDBVersion, Pos1stDot - 1), "0000")
    sVer2 = Format(Mid(gstrDBVersion, Pos1stDot + 1, (Pos2ndDot - 1 - Pos1stDot)), "0000")
    sVer3 = Format(Right(gstrDBVersion, Len(gstrDBVersion) - Pos2ndDot), "0000")
    
    sVerDB = sVer1 & "." & sVer2 & "." & sVer3
    gstrDBVersion = sVerDB
    sVerApp = Format(App.major, "0000") & "." & _
              Format(App.minor, "0000") & "." & _
              Format(App.Revision, "0000")
              
    gstrAppVersion = sVerApp
    
    Select Case True
    Case sVerDB = sVerApp
    Case sVerDB < sVerApp
        UpgradeDB
    Case Else
        MsgBox "Database version is newer than the application. " & _
        "The application will now close.", vbOKOnly + vbCritical, "C.M.S. Cannot Load"
        End
    End Select
    
        '
    'Update TMS SQs...
    '
    TMS_UpdateSQDescriptionsFromXLS True, False

    
    

                                                    
    '
    'Now request log-on...
    '
    gbWindowsLogon = False
    If GlobalParms.GetValue("RequirePassword", "TrueFalse") Then
        frmLogin.Show vbModal
    Else
        If Not IsWindowsNT Then
            frmLogin.Show vbModal
        Else
            If Not PerformAutoLogon Then
                frmLogin.Show vbModal
            End If
        End If
    End If
    
    'check DB's not getting too big!
    lDBSize = FileLen(CompletePathToTheMDBFileAndExt)
    dDBSizeMB = lDBSize / 1000000
    If CLng(GlobalParms.GetValue("MaxDBSizeMB", "NumVal")) * _
        CDbl(GlobalParms.GetValue("WarnAboutDBSizeThreshold", "NumFloat")) <= dDBSizeMB Then
        
        MsgBox "The CMS database is approaching the maximum recommended size." & vbCrLf & vbCrLf & _
                "Please take steps " & _
                 "to address this, otherwise data corruption could occur.", _
                    vbOKOnly + vbExclamation, AppName & " - WARNING. DO NOT IGNORE!"
    End If
    
    WriteToLogFile "Checked DB size"
    
    gsDocsDirectory = GlobalParms.GetValue("DocumentLocation", "AlphaVal")
    If gsDocsDirectory = "" Then
        SetDefaultDocLocation
    Else
        If Not gFSO.FolderExists(gsDocsDirectory) Then
            SetDefaultDocLocation
        End If
    End If
    
    WriteToLogFile "Got docs directory"
    
    glMaxResultRows = GlobalParms.GetValue("MaxResultRows", "NumVal")
    gbShowMsgBox = GlobalParms.GetValue("UseMsgBoxForMessages", "TrueFalse")
    gbSuppressMsg = GlobalParms.GetValue("SuppressMessages", "TrueFalse")
    gbAutoSelectPersonIfOnlyMatch = GlobalParms.GetValue("AutoSelectPersonIfOnlyMatch", "TrueFalse")
    gbHandleForeignChars = GlobalParms.GetValue("HandleForeignChars", "TrueFalse")
    glMidWkMtgDay = GlobalParms.GetValue("MidWeekMeetingDay", "NumVal")
    glSundayMtgDay = GlobalParms.GetValue("SundayMeetingDay", "NumVal")
    
    
    Exit Sub
ErrorTrap:
    
    If FileIsOpen Then
        Close #FreeFileNum
    End If
    
    Set gFSO = Nothing
    
    EndProgram

End Sub
Private Sub SetDefaultDocLocation()

    On Error Resume Next
    
    gsDocsDirectory = SpecialFolder(SpecialFolder_Documents) & "\Congregation Management System\Documents"
    MkDir gsDocsDirectory
    GlobalParms.Save "DocumentLocation", "AlphaVal", gsDocsDirectory

End Sub

Public Function PerformAutoLogon() As Boolean
On Error GoTo ErrorTrap
    
Dim rstPassword As Recordset, SQLStr As String, ThePassword As String, TheUserName As String

    gWindowsUsername = GetWindowsUserName
    
    SQLStr = "SELECT ThePassword, UserCode, ActiveFromDate, ActiveToDate " & _
             "FROM tblSecurity WHERE TheUserID = '" & gWindowsUsername & "'"

    Set rstPassword = CMSDB.OpenRecordset(SQLStr, dbOpenSnapshot)
    
    If Not rstPassword.BOF Then
        gCurrentUserID = gWindowsUsername
        gCurrentPassword = rstPassword!ThePassword
        gCurrentUserCode = rstPassword!UserCode
        gdActiveFromDate = IIf(IsNull(rstPassword!ActiveFromDate), "", Format((rstPassword!ActiveFromDate), "dd/mm/yyyy"))
        gdActiveToDate = IIf(IsNull(rstPassword!ActiveToDate), "", Format((rstPassword!ActiveToDate), "dd/mm/yyyy"))
        If Not ActiveLogon Then
            MsgBox "Your logon is either not yet active or has expired. " & _
                    "Log on as Guest, then import new security settings.", vbOKOnly + vbExclamation, AppName
            LoginSucceeded = True
            Exit Function
        End If
        gbResetPassword = False
        PerformAutoLogon = True
        gbWindowsLogon = True
        Exit Function
    Else
        MsgBox "Cannot log on automatically using your Windows username. Please logon using your CMS password.", vbOKOnly + vbInformation, AppName
        PerformAutoLogon = False
        gbWindowsLogon = False
        Exit Function
    End If
    
    Exit Function
ErrorTrap:
    EndProgram
End Function

Private Sub DeleteTempFolder()
On Error Resume Next
Dim fso As New FileSystemObject
    fso.DeleteFolder JustTheDirectory & "\Temp", True
    Set fso = Nothing
End Sub
Public Sub DealWithCongSetup()

On Error GoTo ErrorTrap

    If CongregationIsSetUp Then
        '
        'Are there any brothers held in the DB? If so, enable certain functions
        '
        CheckIfBrothersInDB
        
        If SystemContainsPeople Then
            frmMainMenu.EnforceSecurity
        End If
        
        GlobalDefaultCong = GlobalParms.GetValue("DefaultCong", "NumVal")
        giGlobalDefaultCong = GlobalDefaultCong
        
    End If

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Public Sub RemoveOldLogFiles()

On Error GoTo ErrorTrap

Dim fsoFolder As Scripting.Folder
Dim fsoFile As Scripting.File, str As String, bFound As Boolean

    'remove log files more than 21 days old
    WriteToLogFile "***** Deleting any old log files *****"
    
    Set fsoFolder = gFSO.GetFolder(gsLogPath)
    
    For Each fsoFile In fsoFolder.Files
        If gFSO.GetExtensionName(fsoFile.Path) = "log" Then
            If DateDiff("d", fsoFile.DateCreated, Now) > 21 Then
                str = fsoFile.Name
                fsoFile.Delete True
                WriteToLogFile str & " deleted"
                bFound = True
            End If
        End If
    Next
    
    If Not bFound Then
        WriteToLogFile "No log files deleted"
    End If
    
    Set fsoFolder = Nothing
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Public Sub KeepCongNoInLine(NewCongNo As Long, OldCongNo As Long)

On Error Resume Next
    
    'keep other tables containing congno in line...
    CMSDB.Execute ("UPDATE tblEvents SET CongNo = " & NewCongNo & " WHERE CongNo = " & OldCongNo)
    CMSDB.Execute ("UPDATE tblVisitingSpeakers SET CongNo = " & NewCongNo & " WHERE CongNo = " & OldCongNo)
    CMSDB.Execute ("UPDATE tblPublicMtgSchedule SET CongNoWhereMtgIs = " & NewCongNo & " WHERE CongNoWhereMtgIs = " & OldCongNo)
    CMSDB.Execute ("UPDATE tblTaskAndPerson SET CongNo = " & NewCongNo & " WHERE CongNo = " & OldCongNo)
    CMSDB.Execute ("UPDATE tblTaskPersonSuspendDates SET CongNo = " & NewCongNo & " WHERE CongNo = " & OldCongNo)

End Sub

