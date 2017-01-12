Attribute VB_Name = "basTMSSchedule"
Option Explicit


Public Function PrintTMSSchedule(NewSchedule As Boolean, ByVal sPrintMode As String, bRegenStoredRota As Boolean, Optional bNewRota As Boolean = True) As Boolean
Dim WorkSheetTopMargin As Single
Dim WorkSheetBottomMargin As Single
Dim WorkSheetLeftMargin As Single
Dim WorkSheetRightMargin As Single
Dim bShowTheme As Boolean
Dim bSetNextScheduleDate As Boolean
Dim rs As Recordset
Dim sStartDate As String
Dim sEndDate As String

On Error GoTo ErrorTrap

    Screen.MousePointer = vbHourglass
    '
    'Build the table used for the report....
    '
    bSetNextScheduleDate = (sPrintMode = "FINAL")
    sStartDate = frmTMSPrinting.cmbStartDate
    sEndDate = frmTMSPrinting.cmbEndDate
    If Not BuildPrintTable Then
        PrintTMSSchedule = False
        Screen.MousePointer = vbNormal
        Exit Function
    End If
    
    
    If frmTMSPrinting.chkExportToFile = vbUnchecked Then
        Select Case PrintUsingWord(False)
        Case cmsUseWord
            Screen.MousePointer = vbNormal
            PrintTMSScheduleUsingWord_2016
        Case Else
            bSetNextScheduleDate = False
        End Select
        
        If bSetNextScheduleDate Then
            Set rs = GetGeneralRecordset("SELECT MAX(AssignmentDate) AS MaxDate " & _
                                         "FROM tblTMSPrintSchedule ")
                                         
            If Not (IsNull(rs!MaxDate)) Then
                If ValidDate(HandleNull(GlobalParms.GetValue("NextTMSSchedulePrintStartDate", "DateVal"), 0)) Then
                    If CDate(DateAdd("ww", 1, rs!MaxDate)) > CDate(GlobalParms.GetValue("NextTMSSchedulePrintStartDate", "DateVal")) Then
                        GlobalParms.Save "NextTMSSchedulePrintStartDate", "DateVal", DateAdd("ww", 1, rs!MaxDate)
                    End If
                Else
                    GlobalParms.Save "NextTMSSchedulePrintStartDate", "DateVal", DateAdd("ww", 1, rs!MaxDate)
                End If
            End If
        End If
        
    Else
        Select Case NewMtgArrangementStarted(frmTMSPrinting.cmbStartDate.text)
        Case CLM2016
            ExportScheduleToFile_2016 CDate(sStartDate), CDate(sEndDate)
        Case TMS2009
            ExportScheduleToFile_2009
        Case Else
            ExportScheduleToFile
        End Select
        Screen.MousePointer = vbNormal
    End If

    PrintTMSSchedule = True
    Screen.MousePointer = vbNormal

    Exit Function
ErrorTrap:
    EndProgram
    
End Function
Private Function BuildPrintTable(Optional pStartDate As Date, Optional pEndDate As Date) As Boolean
Dim rstNewPrintedSchedule As Recordset
Dim rstSchedule As Recordset, ScheduleSQL As String
Dim TheStartDate As String, TheEndDate As String
Dim TheStartDateUS As String, TheEndDateUS As String
Dim PrevAssignmentDate As Date
Dim PersonName As String
Dim AssistantName As String
Dim bNewArr_Start As MidweekMtgVersion, bNewArr_End As MidweekMtgVersion
Dim i As Long, sThemeAndSource As String
Dim dte As Date, ErrorCode As Integer
Dim CurrItemsSeqNum As Long
Dim PrevItemsSeqNum As Long
Dim lSQ As Long, sSQ As String
Dim sTheme As String



On Error GoTo ErrorTrap

    DelAllRows "tblTMSPrintSchedule"
    
    
    'build the print table
     DeleteTable "tblTMSPrintSchedule"
     CreateTable ErrorCode, "tblTMSPrintSchedule", "ItemsSeqNum", "LONG", , , True, "SeqNum"
     CreateField ErrorCode, "tblTMSPrintSchedule", "TalkSeqNum", "LONG"
     CreateField ErrorCode, "tblTMSPrintSchedule", "AssignmentDate", "DATE"
     CreateField ErrorCode, "tblTMSPrintSchedule", "AssignmentDateStr", "TEXT"
     CreateField ErrorCode, "tblTMSPrintSchedule", "TalkType", "TEXT"
     CreateField ErrorCode, "tblTMSPrintSchedule", "Theme", "TEXT", "255"
     CreateField ErrorCode, "tblTMSPrintSchedule", "Student1", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "Assistant1", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "SQ1", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "Student2", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "Assistant2", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "SQ2", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "Student3", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "Assistant3", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "SQ3", "TEXT", "100"
    
    
    With frmTMSPrinting
    
    If Not ValidDate(CStr(pStartDate)) Then
        If DateDiff("d", CDate(.cmbStartDate.text), CDate(.cmbEndDate.text)) > 366 Then
            MsgBox "Schedule should be no longer than one year", vbExclamation + vbOKOnly, AppName
            BuildPrintTable = False
            Exit Function
        End If
    
        bNewArr_Start = NewMtgArrangementStarted(.cmbStartDate.text)
        bNewArr_End = NewMtgArrangementStarted(.cmbEndDate.text)
        
        If bNewArr_Start <> bNewArr_End Then
            MsgBox "You cannot print a schedule spanning the old and new meeting arrangement " _
            , vbInformation + vbOKOnly, AppName
            BuildPrintTable = False
            Exit Function
        End If
        
        TheStartDateUS = Format(.cmbStartDate.text, "mm/dd/yyyy")
        TheEndDateUS = Format(.cmbEndDate.text, "mm/dd/yyyy")
        
    Else
        TheStartDateUS = Format(CStr(pStartDate), "mm/dd/yyyy")
        TheEndDateUS = Format(CStr(pEndDate), "mm/dd/yyyy")
    End If
    
    
    End With
    
    '
    'Driving Recset to get all assignments in date range
    '
    ScheduleSQL = "SELECT * FROM tblTMSSchedule " & _
                " WHERE AssignmentDate BETWEEN #" & TheStartDateUS & "# AND #" & _
                TheEndDateUS & "# " & _
                " AND (PersonID > 0 OR TalkNo IN ('A')) " & _
                " AND SchoolNo <= " & GlobalParms.GetValue("TMSNoSchoolsForSchedulePrint", "NumVal", 1) & " " & _
                " ORDER BY AssignmentDate, TalkSeqNum, ItemsSeqNum,  SchoolNo "

    Set rstSchedule = CMSDB.OpenRecordset(ScheduleSQL, dbOpenDynaset)
    
    
    If rstSchedule.BOF Then
        BuildPrintTable = False
        Exit Function
    Else
        PrevItemsSeqNum = rstSchedule!ItemsSeqNum
    End If
    
        
    Set rstNewPrintedSchedule = CMSDB.OpenRecordset("Select * FROM tblTMSPrintSchedule " & _
                                               " ORDER By AssignmentDate, TalkSeqNum, ItemsSeqNum", dbOpenDynaset)
        
    With rstNewPrintedSchedule
    Do Until rstSchedule.EOF
            
        .AddNew
        !AssignmentDateStr = CStr(Format(rstSchedule!AssignmentDate, "dd/mm/yyyy"))
        !AssignmentDate = rstSchedule!AssignmentDate
        !TalkSeqNum = rstSchedule!TalkSeqNum
        CurrItemsSeqNum = HandleNull(rstSchedule!ItemsSeqNum, -1)
        !ItemsSeqNum = CurrItemsSeqNum
        !TalkType = TheTMS.GetTMSTalkDescription(rstSchedule!TalkNo, CStr(Format(rstSchedule!AssignmentDate, "dd/mm/yyyy")))
         
        If TheTMS.GetTMSItemThemeAndSource(rstSchedule!AssignmentDate, _
                                           rstSchedule!TalkNo, CurrItemsSeqNum) = TMSOK Then
            sThemeAndSource = TheTMS.TMSThemeAndSource
        Else
            sThemeAndSource = ""
        End If
        
        !Theme = sThemeAndSource
        
        .Update
            
        Do While Not rstSchedule.EOF

            AssistantName = CongregationMember.NameWithMiddleInitial(rstSchedule!Assistant1ID)
            If Left(AssistantName, 1) = "?" Then
                AssistantName = ""
            End If
            
            PersonName = CongregationMember.NameWithMiddleInitial(rstSchedule!PersonID)
            If Left(PersonName, 1) = "?" Then
                PersonName = ""
            End If
            
            If rstSchedule!TalkNo <> "A" Then
                lSQ = rstSchedule!CounselPoint
                If lSQ > 0 Then
                    sSQ = CStr(lSQ) & "-" & TheTMS.GetTMSCounselDescription(lSQ)
                Else
                    sSQ = ""
                End If
            Else
                sSQ = ""
            End If
            
            .Requery
            .MoveLast
            .Edit
            .Fields("Student" & rstSchedule!SchoolNo) = PersonName
            .Fields("Assistant" & rstSchedule!SchoolNo) = AssistantName
            .Fields("SQ" & rstSchedule!SchoolNo) = sSQ
            
            .Update
            
            If Not IsNull(rstSchedule!ItemsSeqNum) Then
                PrevItemsSeqNum = rstSchedule!ItemsSeqNum
            Else
                PrevItemsSeqNum = -1
            End If
            
            rstSchedule.MoveNext
            
            
            If Not rstSchedule.EOF Then
                CurrItemsSeqNum = HandleNull(rstSchedule!ItemsSeqNum, -1)
                If PrevItemsSeqNum <> CurrItemsSeqNum Then
                    Exit Do
                End If
            End If
            
        Loop
            
        
        If rstSchedule.EOF Then
            Exit Do
        End If
            
    Loop
    End With
    
    If CopyPrintSchedule Then
        BuildPrintTable = True
    Else
        BuildPrintTable = False
        Exit Function
    End If
    
    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Function SetUpSchedulesForPrint(bRegen As Boolean) As Boolean
Dim rstPrintSchedules As Recordset

On Error GoTo ErrorTrap


    If frmTMSPrinting!cmbStoredSchedules.ListIndex > -1 Then
        Set rstPrintSchedules = CMSDB.OpenRecordset("SELECT * FROM tblStoredTMSSchedules", dbOpenDynaset)
    
        rstPrintSchedules.FindFirst "SeqNum = " & frmTMSPrinting!cmbStoredSchedules.ItemData(frmTMSPrinting!cmbStoredSchedules.ListIndex)
        
        If bRegen Then
            If Not BuildPrintTable(rstPrintSchedules!StartDate, rstPrintSchedules!EndDate) Then
                SetUpSchedulesForPrint = False
                Exit Function
            End If
        Else
            If TableExists(rstPrintSchedules!ScheduleTableName) Then
                CopyTable "tblTMSPrintSchedule", rstPrintSchedules!ScheduleTableName, CMSDB
                lnkTMSScheduleName = rstPrintSchedules!DisplayForCombo
            Else
                MsgBox frmTMSPrinting!cmbStoredSchedules.text & " do not exist. You should now select another schedule to " & _
                "print, or create a new one.", vbExclamation + vbOKOnly, AppName
                SetUpSchedulesForPrint = False
                Exit Function
            End If
        End If
    Else
        SetUpSchedulesForPrint = False
        Exit Function
    End If
    
    SetUpSchedulesForPrint = True

    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Function CopyPrintSchedule() As Boolean
'
'Create copy of tblTMSPrintSchedule just produced using CopyTable.
'Name of table will be tblTMSPrintSchedule||<StartDate>||<EndDate>
'
Dim BeginDate As Date, StrBeginDate As String
Dim EndDate As Date, StrEndDate As String
Dim NewTable As String, i As Integer, temp, rstTMSSchedules As Recordset
Dim rstLatestTMSSchedule As Recordset, NewName As String

    On Error GoTo ErrorTrap
    
    i = 1
    
    Set rstLatestTMSSchedule = CMSDB.OpenRecordset("SELECT * FROM tblTMSPrintSchedule", dbOpenDynaset)
    
    With rstLatestTMSSchedule
    
    .MoveFirst
    BeginDate = !AssignmentDate
    StrBeginDate = CStr(BeginDate)
    
    .MoveLast
    EndDate = !AssignmentDate
    StrEndDate = CStr(EndDate)
        
    End With
    
    NewTable = "tblTMSPrintSchedule: " & StrBeginDate & " TO " & StrEndDate & " (" & Format(i, "0000") & ")"
            
    '
    'Does this tablename already exist? If so, keep incrementing suffix (i) and trying again until unique name
    ' found.
    '
    If Not TableExists(NewTable) Then 'Table doesn't exist. Fine.
        CopyTable NewTable, "tblTMSPrintSchedule", CMSDB
    Else
        Do
            i = i + 1
            Err.Clear
            NewTable = "tblTMSPrintSchedule: " & StrBeginDate & " TO " & StrEndDate & " (" & Format(i, "0000") & ")"
        Loop Until Not TableExists(NewTable) Or i = 9999
        
        If i = 9999 Then
            MsgBox "Could not copy tblTMSPrintSchedule. Delete some tables.", vbOKOnly + vbCritical, AppName
            CopyPrintSchedule = False
            Exit Function
        Else
            CopyTable NewTable, "tblTMSPrintSchedule", CMSDB
        End If
    End If
            
    '
    'Insert new table name on tblStoredTMSSchedules - for display in combo
    '
             
    Set rstTMSSchedules = CMSDB.OpenRecordset("SELECT * FROM tblStoredTMSSchedules  ORDER BY ModifiedDateTime DESC", dbOpenDynaset)

    With rstTMSSchedules
    .AddNew
    !ScheduleTableName = NewTable
    NewName = "Schedule dates " & StrBeginDate & " TO " & StrEndDate & " (" & Format(i, "0000") & ")"
    !DisplayForCombo = NewName
    lnkTMSScheduleName = NewName
    !StartDate = BeginDate
    !EndDate = EndDate
    !CreatedDateTime = Now
    !ModifiedDateTime = Now
    .Update
    
    '
    'Now check if we need to delete old rotas
    '
    .Requery
    .MoveLast
    If GlobalParms.GetValue("NumberOfTMSSchedulesToKeep", "NumVal") < .RecordCount Then
        DeleteTable !ScheduleTableName
        
        .Delete
    End If
    
    End With
    
    HandleListBox.Requery frmTMSPrinting!cmbStoredSchedules, False, CMSDB
    SetCmbStoredSchedules
    
    CopyPrintSchedule = True
    
    Exit Function


ErrorTrap:

    EndProgram


End Function

Public Sub SetCmbStoredSchedules()
Dim rstTemp As Recordset, strSQL As String, TempDate As Date

On Error GoTo ErrorTrap


    strSQL = "SELECT max(ModifiedDateTime) as MaxDate " & _
             "FROM tblStoredTMSSchedules "
             
    Set rstTemp = CMSDB.OpenRecordset(strSQL, dbOpenDynaset)
    
    If Not rstTemp.BOF And Not IsNull(rstTemp!MaxDate) Then
        TempDate = rstTemp!MaxDate
        strSQL = "SELECT * " & _
                 "FROM tblStoredTMSSchedules "
                 
        Set rstTemp = CMSDB.OpenRecordset(strSQL, dbOpenDynaset)
        
        rstTemp.FindFirst "ModifiedDateTime = #" & Format(TempDate, "mm/dd/yy hh:mm:ss") & "#"
        
        HandleListBox.SelectItem frmTMSPrinting!cmbStoredSchedules, rstTemp!SeqNum
    End If

    rstTemp.Close

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Function GetNoOfSchools() As Integer
Dim rstTemp As Recordset, strSQL As String, NoOfSchools As Integer, MaxNoOfSchools As Integer

On Error GoTo ErrorTrap

'
'Scan the Schedule Printtable to find the number of schools.....
'

    GetNoOfSchools = 1
    MaxNoOfSchools = 1

    strSQL = "SELECT * " & _
             "FROM tblTMSPrintSchedule "
             
    Set rstTemp = CMSDB.OpenRecordset(strSQL, dbOpenSnapshot)
    
    With rstTemp
    
    If Not .BOF Then
        Do While Not .EOF
            If NewMtgArrangementStarted(!AssignmentDateStr) Then
                If IsNull(!No1BroSchool2) And _
                   IsNull(!No2BroSchool2) And _
                   IsNull(!No3BroSchool2) And _
                   IsNull(!No3AsstSchool2) And _
                   IsNull(!No2AsstSchool2) Then
                Else
                    If MaxNoOfSchools < 2 Then
                        MaxNoOfSchools = 2
                    End If
                End If
                    
                If IsNull(!No1BroSchool3) And _
                   IsNull(!No2BroSchool3) And _
                   IsNull(!No3BroSchool3) And _
                   IsNull(!No3AsstSchool3) And _
                   IsNull(!No2AsstSchool3) Then
                Else
                    If MaxNoOfSchools < 3 Then
                        MaxNoOfSchools = 3
                    End If
                End If
            Else
                If IsNull(!No2BroSchool2) And _
                   IsNull(!No3BroSchool2) And _
                   IsNull(!No4BroSchool2) And _
                   IsNull(!No3AsstSchool2) And _
                   IsNull(!No4AsstSchool2) Then
                Else
                    If MaxNoOfSchools < 2 Then
                        MaxNoOfSchools = 2
                    End If
                End If
                    
                If IsNull(!No2BroSchool3) And _
                   IsNull(!No3BroSchool3) And _
                   IsNull(!No4BroSchool3) And _
                   IsNull(!No3AsstSchool3) And _
                   IsNull(!No4AsstSchool3) Then
                Else
                    If MaxNoOfSchools < 3 Then
                        MaxNoOfSchools = 3
                    End If
                End If
            End If
                
            .MoveNext
        Loop
    Else
        GetNoOfSchools = 3
    End If
    End With
        
    GetNoOfSchools = MaxNoOfSchools
    
    rstTemp.Close

    Exit Function
ErrorTrap:
    EndProgram
    

End Function

Private Function ExportScheduleToFile() As Boolean
Dim FilePath As String, FileNum As Integer, rstPrintScheduleRecSet As Recordset
Dim StringToPrint As String, FileIsOpen As Boolean

On Error GoTo ErrorTrap

    '
    'Create a text file in which to store the schedule. The name of the file is
    ' derived from the name of the schedule as stored on tblStoredTMSSchedules.
    ' However, we must remove the '/' from the dates and replace with '-' since
    ' Windows doesn't like '/' in filenames.
    '
    FilePath = Replace(gsDocsDirectory & "\" & lnkTMSScheduleName & ".csv", "/", "-")
    
    '
    'Opens filesave dialogue
    '
    If Not SaveCSVFile(FilePath, _
           "C.M.S. Export Schedule to File", _
           frmTMSPrinting.CommonDialog1) Then
        ExportScheduleToFile = False
        Exit Function
    End If
    
    FileNum = FreeFile()
    Open FilePath For Output As #FileNum
    
    Set rstPrintScheduleRecSet = CMSDB.OpenRecordset("tblTMSPrintSchedule", dbOpenForwardOnly)
    
    With rstPrintScheduleRecSet
    If .BOF Then
        MsgBox "Nothing to export.", vbOKOnly + vbExclamation, AppName
        Exit Function
    End If
    
    '
    'Print School 1, 2, 3 heading as appropriate
    '
    StringToPrint = ",," & "School 1"
    If lnkNoOfSchools > 1 Then
        StringToPrint = StringToPrint & "," & "School 2"
    End If
    If lnkNoOfSchools = 3 Then
        StringToPrint = StringToPrint & "," & "School 3"
    End If
    
    StringToPrint = StringToPrint & vbCrLf
    Print #FileNum, StringToPrint;
    
    Do While Not .EOF 'For each row on tblTMSPrintSchedule
        '
        'Print Assignment Date, Prayer bro. Theme derived on-the-fly.
        '
        StringToPrint = !AssignmentDateStr & "," & "Opening Prayer" & "," & !PrayerBro
        
        TheTMS.GetTMSItemThemeAndSource !AssignmentDate, "P", 0 'TODO: need to derive and set 3rd parm
        
        'Align theme and source to appropriate column
        Select Case lnkNoOfSchools
        Case 1
            StringToPrint = StringToPrint & ","
        Case 2
            StringToPrint = StringToPrint & ",,"
        Case 3
            StringToPrint = StringToPrint & ",,,"
        End Select
        
        'Theme and source supplied as properties after execution of method
        ' GetTMSItemThemeAndSource. Strip out any commas to prevent extra field
        ' being created.
        StringToPrint = StringToPrint & Replace(TheTMS.TMSTheme, ",", " ") & "," & _
                                        Replace(TheTMS.TMSSourceMaterial, ",", " ")
        
        StringToPrint = vbCrLf & StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
        '
        'Print SQ bro
        '
        StringToPrint = "," & "Speech Quality" & "," & !SQBro
        
        TheTMS.GetTMSItemThemeAndSource !AssignmentDate, "S", 0
        
        'Align theme and source to appropriate column
        Select Case lnkNoOfSchools
        Case 1
            StringToPrint = StringToPrint & ","
        Case 2
            StringToPrint = StringToPrint & ",,"
        Case 3
            StringToPrint = StringToPrint & ",,,"
        End Select
        
        StringToPrint = StringToPrint & Replace(TheTMS.TMSTheme, ",", " ") & "," & _
                                        Replace(TheTMS.TMSSourceMaterial, ",", " ")
        
        StringToPrint = StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
        '
        'Print No1 bro
        '
        StringToPrint = "," & "Talk No 1" & "," & !No1Bro
        
        TheTMS.GetTMSItemThemeAndSource !AssignmentDate, "1", 0
        
        'Align theme and source to appropriate column
        Select Case lnkNoOfSchools
        Case 1
            StringToPrint = StringToPrint & ","
        Case 2
            StringToPrint = StringToPrint & ",,"
        Case 3
            StringToPrint = StringToPrint & ",,,"
        End Select
        
        StringToPrint = StringToPrint & Replace(TheTMS.TMSTheme, ",", " ") & "," & _
                                        Replace(TheTMS.TMSSourceMaterial, ",", " ")
        
        StringToPrint = StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
        '
        'Print BH bro
        '
        StringToPrint = "," & "Bible Highlights" & "," & !BHBro
        
        TheTMS.GetTMSItemThemeAndSource !AssignmentDate, "B", 0
        
        'Align theme and source to appropriate column
        Select Case lnkNoOfSchools
        Case 1
            StringToPrint = StringToPrint & ","
        Case 2
            StringToPrint = StringToPrint & ",,"
        Case 3
            StringToPrint = StringToPrint & ",,,"
        End Select
        
        StringToPrint = StringToPrint & Replace(TheTMS.TMSTheme, ",", " ") & "," & _
                                        Replace(TheTMS.TMSSourceMaterial, ",", " ")
        
        StringToPrint = StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
        '
        'Print No2 bro
        '
        StringToPrint = "," & "Talk No 2" & "," & !No2BroSchool1
        If lnkNoOfSchools > 1 Then
            StringToPrint = StringToPrint & "," & !No2BroSchool2
        End If
        If lnkNoOfSchools = 3 Then
            StringToPrint = StringToPrint & "," & !No2BroSchool3
        End If
        
        TheTMS.GetTMSItemThemeAndSource !AssignmentDate, "2", 0
        StringToPrint = StringToPrint & "," & Replace(TheTMS.TMSTheme, ",", " ") & "," & _
                                              Replace(TheTMS.TMSSourceMaterial, ",", " ")
        
        StringToPrint = StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
        '
        'Print No3 bro
        '
        StringToPrint = "," & "Talk No 3" & "," & !No3BroSchool1
        If lnkNoOfSchools > 1 Then
            StringToPrint = StringToPrint & "," & !No3BroSchool2
        End If
        If lnkNoOfSchools = 3 Then
            StringToPrint = StringToPrint & "," & !No3BroSchool3
        End If
        
        TheTMS.GetTMSItemThemeAndSource !AssignmentDate, "3", 0
        StringToPrint = StringToPrint & "," & Replace(TheTMS.TMSTheme, ",", " ") & "," & _
                                              Replace(TheTMS.TMSSourceMaterial, ",", " ")
        
        StringToPrint = StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
        '
        'Print No3 Assistant
        '
        StringToPrint = ",," & !No3AsstSchool1
        If lnkNoOfSchools > 1 Then
            StringToPrint = StringToPrint & "," & !No3AsstSchool2
        End If
        If lnkNoOfSchools = 3 Then
            StringToPrint = StringToPrint & "," & !No3AsstSchool3
        End If
        
        StringToPrint = StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
        '
        'Print No4 Bro
        '
        StringToPrint = "," & "Talk No 4" & "," & !No4BroSchool1
        If lnkNoOfSchools > 1 Then
            StringToPrint = StringToPrint & "," & !No4BroSchool2
        End If
        If lnkNoOfSchools = 3 Then
            StringToPrint = StringToPrint & "," & !No4BroSchool3
        End If
        
        TheTMS.GetTMSItemThemeAndSource !AssignmentDate, "4", 0
        StringToPrint = StringToPrint & "," & Replace(TheTMS.TMSTheme, ",", " ") & "," & _
                                              Replace(TheTMS.TMSSourceMaterial, ",", " ")
        
        StringToPrint = StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
        '
        'Print No4 Assistant
        '
        StringToPrint = ",," & !No4AsstSchool1
        If lnkNoOfSchools > 1 Then
            StringToPrint = StringToPrint & "," & !No4AsstSchool2
        End If
        If lnkNoOfSchools = 3 Then
            StringToPrint = StringToPrint & "," & !No4AsstSchool3
        End If
        
        StringToPrint = StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
        .MoveNext
    Loop
    
    End With
    
    Close #FileNum
        
    If MsgBox("Schedule successfully exported as '" & FilePath & _
           "'. You may open this file in a spreadsheet. Do you want to " & _
           "work with the file now?", vbYesNo + vbQuestion, AppName) = vbYes Then
           
           On Error Resume Next 'Any prob opening Explorer - don't abend rog.
           
           OpenWindowsExplorer FilePath, True, True
           
           If Err.number > 0 Then
                MsgBox "Error opening explorer.", vbOKOnly + vbExclamation, AppName
           End If
           
           On Error GoTo ErrorTrap
    
    End If
    
    Exit Function
ErrorTrap:
    If FileIsOpen Then
        Close #FileNum
    End If
    EndProgram
End Function
Private Function ExportScheduleToFile_2009() As Boolean
Dim FilePath As String, FileNum As Integer, rstPrintScheduleRecSet As Recordset
Dim StringToPrint As String, FileIsOpen As Boolean

On Error GoTo ErrorTrap

    '
    'Create a text file in which to store the schedule. The name of the file is
    ' derived from the name of the schedule as stored on tblStoredTMSSchedules.
    ' However, we must remove the '/' from the dates and replace with '-' since
    ' Windows doesn't like '/' in filenames.
    '
    FilePath = Replace(gsDocsDirectory & "\" & lnkTMSScheduleName & ".csv", "/", "-")
    
    '
    'Opens filesave dialogue
    '
    If Not SaveCSVFile(FilePath, _
           "C.M.S. Export Schedule to File", _
           frmTMSPrinting.CommonDialog1) Then
        ExportScheduleToFile_2009 = False
        Exit Function
    End If
    
    FileNum = FreeFile()
    Open FilePath For Output As #FileNum
    
    Set rstPrintScheduleRecSet = CMSDB.OpenRecordset("tblTMSPrintSchedule", dbOpenForwardOnly)
    
    With rstPrintScheduleRecSet
    If .BOF Then
        MsgBox "Nothing to export.", vbOKOnly + vbExclamation, AppName
        Exit Function
    End If
    
    '
    'Print School 1, 2, 3 heading as appropriate
    '
    StringToPrint = ",," & "School 1"
    If lnkNoOfSchools > 1 Then
        StringToPrint = StringToPrint & "," & "School 2"
    End If
    If lnkNoOfSchools = 3 Then
        StringToPrint = StringToPrint & "," & "School 3"
    End If
    
    StringToPrint = StringToPrint & vbCrLf
    Print #FileNum, StringToPrint;
    
    Do While Not .EOF 'For each row on tblTMSPrintSchedule
        '
        'Print Assignment Date, Prayer bro. Theme derived on-the-fly.
        '
        StringToPrint = !AssignmentDateStr & "," & "Opening Prayer" & "," & !PrayerBro
        
        TheTMS.GetTMSItemThemeAndSource !AssignmentDate, "P", 0
        
        'Align theme and source to appropriate column
        Select Case lnkNoOfSchools
        Case 1
            StringToPrint = StringToPrint & ","
        Case 2
            StringToPrint = StringToPrint & ",,"
        Case 3
            StringToPrint = StringToPrint & ",,,"
        End Select
        
        'Theme and source supplied as properties after execution of method
        ' GetTMSItemThemeAndSource. Strip out any commas to prevent extra field
        ' being created.
        StringToPrint = StringToPrint & Replace(TheTMS.TMSTheme, ",", " ") & "," & _
                                        Replace(TheTMS.TMSSourceMaterial, ",", " ")
        
        StringToPrint = vbCrLf & StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
             
        '
        'Print BH bro
        '
        StringToPrint = "," & "Bible Highlights" & "," & !BHBro
        
        TheTMS.GetTMSItemThemeAndSource !AssignmentDate, "B", 0
        
        'Align theme and source to appropriate column
        Select Case lnkNoOfSchools
        Case 1
            StringToPrint = StringToPrint & ","
        Case 2
            StringToPrint = StringToPrint & ",,"
        Case 3
            StringToPrint = StringToPrint & ",,,"
        End Select
        
        StringToPrint = StringToPrint & Replace(TheTMS.TMSTheme, ",", " ") & "," & _
                                        Replace(TheTMS.TMSSourceMaterial, ",", " ")
        
        StringToPrint = StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
        '
        'Print No1 bro
        '
        StringToPrint = "," & "Talk No 1" & "," & !No1Bro
        If lnkNoOfSchools > 1 Then
            StringToPrint = StringToPrint & "," & !No1BroSchool2
        End If
        If lnkNoOfSchools = 3 Then
            StringToPrint = StringToPrint & "," & !No1BroSchool3
        End If
        
        TheTMS.GetTMSItemThemeAndSource !AssignmentDate, "1", 0
        StringToPrint = StringToPrint & "," & Replace(TheTMS.TMSTheme, ",", " ") & "," & _
                                              Replace(TheTMS.TMSSourceMaterial, ",", " ")
        
        StringToPrint = StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
        '
        'Print No2 bro
        '
        StringToPrint = "," & "Talk No 2" & "," & !No2BroSchool1
        If lnkNoOfSchools > 1 Then
            StringToPrint = StringToPrint & "," & !No2BroSchool2
        End If
        If lnkNoOfSchools = 3 Then
            StringToPrint = StringToPrint & "," & !No2BroSchool3
        End If
        
        TheTMS.GetTMSItemThemeAndSource !AssignmentDate, "2", 0
        StringToPrint = StringToPrint & "," & Replace(TheTMS.TMSTheme, ",", " ") & "," & _
                                              Replace(TheTMS.TMSSourceMaterial, ",", " ")
        
        StringToPrint = StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
        '
        'Print No2 Assistant
        '
        StringToPrint = ",," & !No2AsstSchool1
        If lnkNoOfSchools > 1 Then
            StringToPrint = StringToPrint & "," & !No2AsstSchool2
        End If
        If lnkNoOfSchools = 3 Then
            StringToPrint = StringToPrint & "," & !No2AsstSchool3
        End If
        
        StringToPrint = StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
        '
        'Print No3 Bro
        '
        StringToPrint = "," & "Talk No 3" & "," & !No3BroSchool1
        If lnkNoOfSchools > 1 Then
            StringToPrint = StringToPrint & "," & !No3BroSchool2
        End If
        If lnkNoOfSchools = 3 Then
            StringToPrint = StringToPrint & "," & !No3BroSchool3
        End If
        
        TheTMS.GetTMSItemThemeAndSource !AssignmentDate, "3", 0
        StringToPrint = StringToPrint & "," & Replace(TheTMS.TMSTheme, ",", " ") & "," & _
                                              Replace(TheTMS.TMSSourceMaterial, ",", " ")
        
        StringToPrint = StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
        '
        'Print No3 Assistant
        '
        StringToPrint = ",," & !No3AsstSchool1
        If lnkNoOfSchools > 1 Then
            StringToPrint = StringToPrint & "," & !No3AsstSchool2
        End If
        If lnkNoOfSchools = 3 Then
            StringToPrint = StringToPrint & "," & !No3AsstSchool3
        End If
        
        StringToPrint = StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
        .MoveNext
    Loop
    
    End With
    
    Close #FileNum
        
    If MsgBox("Schedule successfully exported as '" & FilePath & _
           "'. You may open this file in a spreadsheet. Do you want to " & _
           "work with the file now?", vbYesNo + vbQuestion, AppName) = vbYes Then
           
           On Error Resume Next 'Any prob opening Explorer - don't abend rog.
           
           OpenWindowsExplorer FilePath, True, True
           
           If Err.number > 0 Then
                MsgBox "Error opening explorer.", vbOKOnly + vbExclamation, AppName
           End If
           
           On Error GoTo ErrorTrap
    
    End If
    
    Exit Function
ErrorTrap:
    If FileIsOpen Then
        Close #FileNum
    End If
    EndProgram
End Function

Private Function ExportScheduleToFile_2016(StartDate As Date, EndDate As Date) As Boolean

On Error GoTo ErrorTrap
    
    GenerateExcelStudentTalkSchedule StartDate, EndDate
    
    Exit Function
ErrorTrap:
    EndProgram
End Function



Public Sub AcquireTMSWeightingParameters()
On Error GoTo ErrorTrap

    With GlobalParms
    TMSBibleReadingWeighting_2016 = .GetValue("TMSBibleReadingWeighting_2016", "NumFloat")
    TMSInitialCallWeighting_2016 = .GetValue("TMSInitialCallWeighting_2016", "NumFloat")
    TMSReturnVisitWeighting_2016 = .GetValue("TMSReturnVisitWeighting_2016", "NumFloat")
    TMSBibleStudyWeighting_2016 = .GetValue("TMSBibleStudyWeighting_2016", "NumFloat")
    TMSOtherWeighting_2016 = .GetValue("TMSOtherWeighting_2016", "NumFloat")
    TMSPrayerWeighting = .GetValue("TMSPrayerWeighting", "NumFloat")
    TMSSQWeighting = .GetValue("TMSSQWeighting", "NumFloat")
    TMSNo1Weighting = .GetValue("TMSNo1Weighting", "NumFloat")
    TMSBHWeighting = .GetValue("TMSBHWeighting", "NumFloat")
    TMSReviewReaderWeighting = .GetValue("TMS_ReviewReaderWeighting", "NumFloat")
    TMSNo2Weighting = .GetValue("TMSNo2Weighting", "NumFloat")
    TMSNo3Weighting = .GetValue("TMSNo3Weighting", "NumFloat")
    TMSNo4Weighting = .GetValue("TMSNo4Weighting", "NumFloat")
    TMSAsstWeighting = .GetValue("TMSAsstWeighting", "NumFloat")
    TMSNo1Weighting_2009 = .GetValue("TMSTMSNo1Weighting_2009", "NumFloat")
    TMSNo2Weighting_2009 = .GetValue("TMSTMSNo2Weighting_2009", "NumFloat")
    TMSNo3Weighting_2009 = .GetValue("TMSTMSNo3Weighting_2009", "NumFloat")
    TMSWeightingIfAssistantOnly = .GetValue("TMSWeightingIfAssistantOnly", "NumFloat")
    End With
    

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Public Sub TransformPrintTable(NoSchools As Long)
On Error GoTo ErrorTrap
Dim rs1 As Recordset, str1 As String, ErrorCode As Integer, i As Long
Dim rs2 As Recordset, str2 As String
    
    DeleteTable "tblTMSTempSchedulePrint"
    CreateTable ErrorCode, "tblTMSTempSchedulePrint", "AssignmentDate", "DATE", , , True
    CreateField ErrorCode, "tblTMSTempSchedulePrint", "Assignment", "TEXT"
    CreateField ErrorCode, "tblTMSTempSchedulePrint", "School1", "TEXT"
    
    If NoSchools > 1 Then
        CreateField ErrorCode, "tblTMSTempSchedulePrint", "School2", "TEXT"
    End If
    If NoSchools > 2 Then
        CreateField ErrorCode, "tblTMSTempSchedulePrint", "School3", "TEXT"
    End If
    
    
    Set rs1 = CMSDB.OpenRecordset("tblTMSPrintSchedule", dbOpenDynaset)
    Set rs2 = CMSDB.OpenRecordset("tblTMSTempSchedulePrint", dbOpenDynaset)
    
    With rs1
    
    Do Until .EOF Or .BOF
        
        rs2.AddNew
        rs2!AssignmentDate = !AssignmentDate
        rs2!Assignment = "Opening Prayer"
        rs2!School1 = !PrayerBro
        If NoSchools > 1 Then
            rs2!School2 = ""
        End If
        If NoSchools > 2 Then
            rs2!School3 = ""
        End If
        rs2.Update
    
        rs2.AddNew
        rs2!AssignmentDate = !AssignmentDate
        rs2!Assignment = "Speech Quality"
        rs2!School1 = !SQBro
        If NoSchools > 1 Then
            rs2!School2 = ""
        End If
        If NoSchools > 2 Then
            rs2!School3 = ""
        End If
        rs2.Update
    
        rs2.AddNew
        rs2!AssignmentDate = !AssignmentDate
        rs2!Assignment = "Talk No 1"
        rs2!School1 = !No1Bro
        If NoSchools > 1 Then
            rs2!School2 = ""
        End If
        If NoSchools > 2 Then
            rs2!School3 = ""
        End If
        rs2.Update
    
        rs2.AddNew
        rs2!AssignmentDate = !AssignmentDate
        rs2!Assignment = "Talk No 2"
        rs2!School1 = !No2BroSchool1
        If NoSchools > 1 Then
            rs2!School2 = !No2BroSchool2
        End If
        If NoSchools > 2 Then
            rs2!School3 = !No2BroSchool3
        End If
        rs2.Update
    
        rs2.AddNew
        rs2!AssignmentDate = !AssignmentDate
        rs2!Assignment = "Talk No 3"
        rs2!School1 = !No3BroSchool1 & IIf(!No3AsstSchool1 <> "", vbCrLf & "   " & !No3AsstSchool1, "")
        If NoSchools > 1 Then
            rs2!School2 = !No3BroSchool2 & IIf(!No3AsstSchool2 <> "", vbCrLf & "   " & !No3AsstSchool2, "")
        End If
        If NoSchools > 2 Then
            rs2!School3 = !No3BroSchool3 & IIf(!No3AsstSchool3 <> "", vbCrLf & "   " & !No3AsstSchool3, "")
        End If
        rs2.Update
    
        rs2.AddNew
        rs2!AssignmentDate = !AssignmentDate
        rs2!Assignment = "Talk No 4"
        rs2!School1 = !No4BroSchool1 & IIf(!No4AsstSchool1 <> "", vbCrLf & "   " & !No4AsstSchool1, "")
        If NoSchools > 1 Then
            rs2!School2 = !No4BroSchool2 & IIf(!No4AsstSchool2 <> "", vbCrLf & "   " & !No4AsstSchool2, "")
        End If
        If NoSchools > 2 Then
            rs2!School3 = !No4BroSchool3 & IIf(!No4AsstSchool3 <> "", vbCrLf & "   " & !No4AsstSchool3, "")
        End If
        rs2.Update
    
        .MoveNext
        
    Loop
    
    End With
    
    rs1.Close
    Set rs1 = Nothing
    rs2.Close
    Set rs2 = Nothing

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Public Function PrintTMSScheduleUsingWord(NoSchools As Long, bDraft As Boolean) As Boolean

Dim reporter As MSWordReportingTool2.RptTool
Dim lNoCols As Long

On Error GoTo ErrorTrap

    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .ReportSQL = "SELECT AssignmentDate, " & _
                 "       PrayerBro, " & _
                 "       SQBro, " & _
                 "       No1Bro, " & _
                 "       BHBro, " & _
                 "       No2BroSchool1, " & _
                 "       No3BroSchool1 & CHR(13) & '  ' & No3AsstSchool1, " & _
                 "       No4BroSchool1 & CHR(13) & '  ' & No4AsstSchool1 "
                 
    If NoSchools > 1 Then
        .ReportSQL = .ReportSQL & ", BHBroSch2, " & _
                             "       No2BroSchool2, " & _
                             "       No3BroSchool2 & CHR(13) & '  ' & No3AsstSchool2, " & _
                             "       No4BroSchool2 & CHR(13) & '  ' & No4AsstSchool2 "
                            
    End If
    
    If NoSchools > 2 Then
        .ReportSQL = .ReportSQL & ", BHBroSch3, " & _
                             "       No2BroSchool3, " & _
                             "       No3BroSchool3 & CHR(13) & '  ' & No3AsstSchool3, " & _
                             "       No4BroSchool3 & CHR(13) & '  ' & No4AsstSchool3 "

    End If
    
    .ReportSQL = .ReportSQL & "FROM tblTMSPrintSchedule ORDER BY 1 "

    .ReportTitle = "Theocratic Ministry School Schedule" & IIf(bDraft, " - DRAFT", "")
    
    .SaveDoc = True
    .DocPath = gsDocsDirectory & "\" & "School Schedule " & IIf(bDraft, "(DRAFT) ", "") & _
                                Replace(Replace(Now, ":", "-"), "/", "-")
    
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 10
    .RightMargin = 10
    .ReportFooterFontName = "Arial"
    .ReportFooterFontSize = 10
    .ReportTitleFontName = "Times New Roman"
    .ReportTitleFontSize = 16
    .ApplyTableFormatting = True
    .ClientName = AppName
    .ShowPageNumber = False
    .HideWordWhileBuilding = True
    .DocumentType = cmsTablePerDBRow
    .NewPageAfterNumberOfTables = 5
    
    .ShowProgress = True
    
    .PageFormat = cmsPortrait
    
    Select Case NoSchools
    Case 1
        .TableCols_Adv = 2
        lNoCols = 2
    Case 2
        .TableCols_Adv = 3
        lNoCols = 3
    Case 3
        .TableCols_Adv = 4
        lNoCols = 4
    End Select
    
    .TableRows_Adv = 8
    
    'apply font attributes to whole table
    .AddCellAttributes_Adv 1, 1, 8, lNoCols, cmsleftTop, "Times New Roman", 10, cmsOptionfalse, _
                                                         cmsOptionfalse, , , , True
        
    'more formatting
    .AddCellAttributes_Adv 1, 1, 1, lNoCols, , , , cmsOptionTrue
    .AddCellAttributes_Adv 1, 1, 8, 1, , , , cmsOptionTrue
    .AddCellAttributes_Adv 1, 1, 1, 1, , , , , cmsOptionTrue
    
    .AddCellAttributes_Adv 1, 1, 8, 1, , , , , , 30
    .AddCellAttributes_Adv 1, 2, 8, 2, , , , , , 45
    
    If NoSchools > 1 Then
        .AddCellAttributes_Adv 1, 3, 8, 3, , , , , , 45
    End If
    If NoSchools > 2 Then
        .AddCellAttributes_Adv 1, 4, 8, 4, , , , , , 45
    End If
        
    'fixed text
    .AddFixedCellText_Adv 1, 2, "School 1", True
    
    If NoSchools > 1 Then
        .AddFixedCellText_Adv 1, 3, "School 2"
    End If
    If NoSchools > 2 Then
        .AddFixedCellText_Adv 1, 4, "School 3"
    End If
    
    .AddFixedCellText_Adv 2, 1, "Opening Prayer"
    .AddFixedCellText_Adv 3, 1, "Speech Quality"
    .AddFixedCellText_Adv 4, 1, "Talk No 1"
    .AddFixedCellText_Adv 5, 1, "Bible Highlights"
    .AddFixedCellText_Adv 6, 1, "Talk No 2"
    .AddFixedCellText_Adv 7, 1, "Talk No 3"
    .AddFixedCellText_Adv 8, 1, "Talk No 4"
     
    'db to ms word table mappings - in order of actual db fields
    .AddFieldMapping_Adv 1, 1, True 'date
    .AddFieldMapping_Adv 2, 2       'Prayer
    .AddFieldMapping_Adv 3, 2       'SQ
    .AddFieldMapping_Adv 4, 2       'No1
    .AddFieldMapping_Adv 5, 2       'BH
    .AddFieldMapping_Adv 6, 2       'No2 school 1
    .AddFieldMapping_Adv 7, 2       'No3 school 1
    .AddFieldMapping_Adv 8, 2       'No4 school 1
    If NoSchools > 1 Then
        .AddFieldMapping_Adv 5, 3   'BH school 2
    End If
    If NoSchools > 1 Then
        .AddFieldMapping_Adv 6, 3   'No2 school 2
    End If
    If NoSchools > 1 Then
        .AddFieldMapping_Adv 7, 3   'No3 school 2
    End If
    If NoSchools > 1 Then
        .AddFieldMapping_Adv 8, 3   'No4 school 2
    End If
    If NoSchools > 2 Then
        .AddFieldMapping_Adv 5, 4   'BH school 3
    End If
    If NoSchools > 2 Then
        .AddFieldMapping_Adv 6, 4   'No2 school 3
    End If
    If NoSchools > 2 Then
        .AddFieldMapping_Adv 7, 4   'No3 school 3
    End If
    If NoSchools > 2 Then
        .AddFieldMapping_Adv 8, 4   'No4 school 3
    End If
                                
    .GenerateReport
       
    End With
    
    SwitchOnDAO
    
    Screen.MousePointer = vbNormal

    Exit Function
ErrorTrap:
    EndProgram
    
    
End Function

Public Function PrintTMSScheduleUsingWord_2016() As Boolean

Dim reporter As MSWordReportingTool2.RptTool
Dim lNoCols As Long
Dim rs As Recordset, sSQL As String
Dim sStartDate As String, sEndDate As String, bSch2 As Boolean, bSch3 As Boolean
Dim sStudent1HdrName As String
Dim sAsst1HdrName As String
Dim lColWidth As Long

On Error GoTo ErrorTrap

    sSQL = "SELECT MAX(AssignmentDate) AS MaxDate, MIN(AssignmentDate) AS MinDate, " & _
            " MAX(Student2) AS MaxStudent2, MAX(Student3) AS MaxStudent3 " & _
            "FROM tblTMSPrintSchedule"
            
    Set rs = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    sStartDate = Format(rs!MinDate, "dd-mm-yyyy")
    sEndDate = Format(rs!MaxDate, "dd-mm-yyyy")
    bSch2 = Not IsNull(rs!MaxStudent2)
    bSch3 = Not IsNull(rs!MaxStudent3)
       
    rs.Close
    Set rs = Nothing


    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .SaveDoc = True
    .DocPath = gsDocsDirectory & "\" & "Student Talks Schedule " & _
                                sStartDate & " to " & sEndDate & " (" & _
                                Replace(Replace(Now, ":", "-"), "/", "-") & ")"

    sSQL = "SELECT AssignmentDateStr, TalkType, Student1, Assistant1 "
    
    If bSch2 Then
        sSQL = sSQL & ", Student2, Assistant2 "
    End If
    
    If bSch3 Then
        sSQL = sSQL & ", Student3, Assistant3 "
    End If
    
    If bSch2 Or bSch3 Then
        sStudent1HdrName = "Student - 1"
        sAsst1HdrName = "Assistant - 1"
    Else
        sStudent1HdrName = "Student"
        sAsst1HdrName = "Assistant"
    End If
    
    If bSch2 Then
        If bSch3 Then
            lColWidth = 24
        Else
            lColWidth = 35
        End If
    Else
        lColWidth = 45
    End If
    
    sSQL = sSQL & " FROM tblTMSPrintSchedule ORDER BY SeqNum "
    
    .ReportSQL = sSQL

    .ReportTitle = "Student Talks Schedule"
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 10
    .RightMargin = 10
    .ReportFooterFontName = "Arial"
    .ReportFooterFontSize = 8
    .ReportTitleFontName = "Times New Roman"
    .ReportTitleFontSize = 18
    .ApplyTableFormatting = True
    .ClientName = AppName
    .AdditionalReportHeading = ""
    .GroupingColumn = 1
    .HideWordWhileBuilding = True
    
    
    .AddTableColumnAttribute "Date", 20, , , , , 9, 9, True, True, , , True
    .AddTableColumnAttribute "Talk", 36, , , , , 9, 9, True, True
    .AddTableColumnAttribute sStudent1HdrName, lColWidth, , , , , 9, 9, True, True
    .AddTableColumnAttribute sAsst1HdrName, lColWidth, , , , , 9, 9, True, True
    
    If bSch2 Then
        .AddTableColumnAttribute "Student - 2", lColWidth, , , , , 9, 9, True, True
        .AddTableColumnAttribute "Assistant - 2", lColWidth, , , , , 9, 9, True, True
    End If
    
    If bSch3 Then
        .AddTableColumnAttribute "Student - 3", lColWidth, , , , , 9, 9, True, True
        .AddTableColumnAttribute "Assistant - 3", lColWidth, , , , , 9, 9, True, True
    End If
    
    .PageFormat = cmsPortrait
    
    .GenerateReport

    End With
    
    SwitchOnDAO
    
    Set reporter = Nothing
    Screen.MousePointer = vbNormal

    
    

    Exit Function
ErrorTrap:
    EndProgram
End Function
    

Public Function PrintTMSScheduleUsingWord_2009(NoSchools As Long, bDraft As Boolean, Optional bNewRota As Boolean = True) As Boolean

Dim reporter As MSWordReportingTool2.RptTool
Dim lNoCols As Long

On Error GoTo ErrorTrap

    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .ReportSQL = "SELECT AssignmentDate, " & _
                 "       BHBro, " & _
                 "       No1Bro, " & _
                 "       No1Title_2009, " & _
                 "       No2BroSchool1 & CHR(13) & '  ' & No2AsstSchool1, " & _
                 "       No3BroSchool1 & CHR(13) & '  ' & No3AsstSchool1 "
                 
    If NoSchools > 1 Then
        .ReportSQL = .ReportSQL & ", BHBroSch2, " & _
                             "       No1BroSchool2, " & _
                             "       No2BroSchool2 & CHR(13) & '  ' & No2AsstSchool2, " & _
                             "       No3BroSchool2 & CHR(13) & '  ' & No3AsstSchool2 "
                            
    End If
    
    If NoSchools > 2 Then
        .ReportSQL = .ReportSQL & ", BHBroSch3, " & _
                             "       No1BroSchool3, " & _
                             "       No2BroSchool3 & CHR(13) & '  ' & No2AsstSchool3, " & _
                             "       No3BroSchool3 & CHR(13) & '  ' & No3AsstSchool3 "

    End If
    
    .ReportSQL = .ReportSQL & "FROM tblTMSPrintSchedule ORDER BY 1 "

    .ReportTitle = "Theocratic Ministry School Schedule" & IIf(bDraft, " - DRAFT", "")
    
    .SaveDoc = True
    
    If bNewRota Then
        .DocPath = gsDocsDirectory & "\" & "School Schedule " & IIf(bDraft, "(DRAFT) ", "") & _
                                " for Weeks " & Replace(frmTMSPrinting.cmbStartDate.text & " to " & _
                                     frmTMSPrinting.cmbEndDate.text, "/", "-") & " (" & _
                                    Replace(Replace(Now, ":", "-"), "/", "-") & ")"
    Else
        .DocPath = gsDocsDirectory & "\" & "School Schedule " & IIf(bDraft, "(DRAFT) ", "") & _
                                " for Weeks " & Replace(frmTMSPrinting.cmbStoredSchedules.text, "/", "-") & " (" & _
                                    Replace(Replace(Now, ":", "-"), "/", "-") & ")"
    End If
    
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 10
    .RightMargin = 10
    .ReportFooterFontName = "Arial"
    .ReportFooterFontSize = 10
    .ReportTitleFontName = "Times New Roman"
    .ReportTitleFontSize = 16
    .ApplyTableFormatting = True
    .ClientName = AppName
    .ShowPageNumber = False
    .HideWordWhileBuilding = True
    .DocumentType = cmsTablePerDBRow
    .NewPageAfterNumberOfTables = 7
    
    .ShowProgress = True
    
    .PageFormat = cmsPortrait
    
    Select Case NoSchools
    Case 1
        .TableCols_Adv = 2
        lNoCols = 2
    Case 2
        .TableCols_Adv = 3
        lNoCols = 3
    Case 3
        .TableCols_Adv = 4
        lNoCols = 4
    End Select
    
    .TableRows_Adv = 5
    
    'apply font attributes to whole table
    .AddCellAttributes_Adv 1, 1, 5, lNoCols, cmsleftTop, "Times New Roman", 10, cmsOptionfalse, _
                                                         cmsOptionfalse, , , , True
        
    'more formatting
    .AddCellAttributes_Adv 1, 1, 1, lNoCols, , , , cmsOptionTrue
    .AddCellAttributes_Adv 1, 1, 5, 1, , , , cmsOptionTrue
    .AddCellAttributes_Adv 1, 1, 1, 1, , , , , cmsOptionTrue
    
    .AddCellAttributes_Adv 1, 1, 5, 1, , , , , , 30
    .AddCellAttributes_Adv 1, 2, 5, 2, , , , , , 45
    
    If NoSchools > 1 Then
        .AddCellAttributes_Adv 1, 3, 5, 3, , , , , , 45
    End If
    If NoSchools > 2 Then
        .AddCellAttributes_Adv 1, 4, 5, 4, , , , , , 45
    End If
        
    'fixed text
    .AddFixedCellText_Adv 1, 2, "School 1", True
    
    If NoSchools > 1 Then
        .AddFixedCellText_Adv 1, 3, "School 2"
    End If
    If NoSchools > 2 Then
        .AddFixedCellText_Adv 1, 4, "School 3"
    End If
    
    .AddFixedCellText_Adv 2, 1, "Bible Highlights"
'    .AddFixedCellText_Adv 3, 1, "Talk No 1"
    .AddFixedCellText_Adv 4, 1, "Talk No 2"
    .AddFixedCellText_Adv 5, 1, "Talk No 3"
     
    'db to ms word table mappings - in order of actual db fields
    .AddFieldMapping_Adv 1, 1, True 'date
    .AddFieldMapping_Adv 2, 2       'BH
    .AddFieldMapping_Adv 3, 2       'No1 school 1
    .AddFieldMapping_Adv 3, 1       'No1 title
    .AddFieldMapping_Adv 4, 2       'No2 school 1
    .AddFieldMapping_Adv 5, 2       'No3 school 1
    If NoSchools > 1 Then
        .AddFieldMapping_Adv 2, 3   'BH school 2
    End If
    If NoSchools > 1 Then
        .AddFieldMapping_Adv 3, 3   'No1 school 2
    End If
    If NoSchools > 1 Then
        .AddFieldMapping_Adv 4, 3   'No2 school 2
    End If
    If NoSchools > 1 Then
        .AddFieldMapping_Adv 5, 3   'No3 school 2
    End If
    If NoSchools > 2 Then
        .AddFieldMapping_Adv 2, 4   'BH school 3
    End If
    If NoSchools > 2 Then
        .AddFieldMapping_Adv 3, 4   'No1 school 3
    End If
    If NoSchools > 2 Then
        .AddFieldMapping_Adv 4, 4   'No2 school 3
    End If
    If NoSchools > 2 Then
        .AddFieldMapping_Adv 5, 4   'No3 school 3
    End If
    
    .AddRowToDeleteIfBlank_Adv 5, 2, 5, True
    .AddRowToDeleteIfBlank_Adv 4, 2, 4
                                
    .GenerateReport
       
    End With
    
    SwitchOnDAO
    
    Screen.MousePointer = vbNormal

    Exit Function
ErrorTrap:
    EndProgram
    
    
End Function

Public Function PrintTMSScheduleUsingWord_wTheme_2009(bDraft As Boolean, Optional bNewRota As Boolean = True) As Boolean

Dim reporter As MSWordReportingTool2.RptTool
Dim lNoCols As Long

On Error GoTo ErrorTrap

    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .ReportSQL = "SELECT AssignmentDate, " & _
                 "       BHBro, " & _
                 "       No1Bro, " & _
                 "       No1Title_2009, " & _
                 "       No2BroSchool1 & CHR(13) & '  ' & No2AsstSchool1, " & _
                 "       No3BroSchool1 & CHR(13) & '  ' & No3AsstSchool1, " & _
                 "       BHTheme, No1Theme, No2Theme, No3Theme "
                 
    
    .ReportSQL = .ReportSQL & "FROM tblTMSPrintSchedule ORDER BY 1 "

    .ReportTitle = "Theocratic Ministry School Schedule" & IIf(bDraft, " - DRAFT", "")
    
    .SaveDoc = True
    
    If bNewRota Then
        .DocPath = gsDocsDirectory & "\" & "School Schedule " & IIf(bDraft, "(DRAFT) ", "") & _
                                " for Weeks " & Replace(frmTMSPrinting.cmbStartDate.text & " to " & _
                                     frmTMSPrinting.cmbEndDate.text, "/", "-") & " (" & _
                                    Replace(Replace(Now, ":", "-"), "/", "-") & ")"
    Else
        .DocPath = gsDocsDirectory & "\" & "School Schedule " & IIf(bDraft, "(DRAFT) ", "") & _
                                " for Weeks " & Replace(frmTMSPrinting.cmbStoredSchedules.text, "/", "-") & " (" & _
                                    Replace(Replace(Now, ":", "-"), "/", "-") & ")"
    End If
    
    
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 10
    .RightMargin = 10
    .ReportFooterFontName = "Arial"
    .ReportFooterFontSize = 10
    .ReportTitleFontName = "Times New Roman"
    .ReportTitleFontSize = 16
    .ApplyTableFormatting = True
    .ClientName = AppName
    .ShowPageNumber = False
    .HideWordWhileBuilding = True
    .DocumentType = cmsTablePerDBRow
    .NewPageAfterNumberOfTables = 7
    .ShowProgress = True
    .PageFormat = cmsPortrait
    
    .TableCols_Adv = 3
    lNoCols = 3
    
    .TableRows_Adv = 5
    
    'apply font attributes to whole table
    .AddCellAttributes_Adv 1, 1, 5, lNoCols, cmsleftTop, "Times New Roman", 10, cmsOptionfalse, _
                                                         cmsOptionfalse, , , , True
        
    'more formatting
    .AddCellAttributes_Adv 1, 1, 1, lNoCols, , , , cmsOptionTrue
    .AddCellAttributes_Adv 1, 1, 5, 1, , , , cmsOptionTrue
    .AddCellAttributes_Adv 1, 1, 1, 1, , , , , cmsOptionTrue
    
    .AddCellAttributes_Adv 1, 1, 5, 1, , , , , , 30
    .AddCellAttributes_Adv 1, 2, 5, 2, , , , , , 45
    
    .AddCellAttributes_Adv 1, 3, 5, 3, , , , , , 100
        
    'fixed text
    .AddFixedCellText_Adv 2, 1, "Bible Highlights"
'    .AddFixedCellText_Adv 3, 1, "Talk No 1"
    .AddFixedCellText_Adv 4, 1, "Talk No 2"
    .AddFixedCellText_Adv 5, 1, "Talk No 3"
     
    'db to ms word table mappings - in order of actual db fields
    .AddFieldMapping_Adv 1, 1, True 'date
    .AddFieldMapping_Adv 2, 2       'BH
    .AddFieldMapping_Adv 3, 2       'No1 school 1
    .AddFieldMapping_Adv 3, 1       'no1 title
    .AddFieldMapping_Adv 4, 2       'No2 school 1
    .AddFieldMapping_Adv 5, 2       'No3 school 1
    .AddFieldMapping_Adv 2, 3
    .AddFieldMapping_Adv 3, 3
    .AddFieldMapping_Adv 4, 3
    .AddFieldMapping_Adv 5, 3

    .AddJoinedCellRange_Adv 1, 1, 1, lNoCols, True
    
    .AddRowToDeleteIfBlank_Adv 5, 2, 5, True
    .AddRowToDeleteIfBlank_Adv 4, 2, 4
    
    .GenerateReport
       
    End With
    
    SwitchOnDAO
    
    Screen.MousePointer = vbNormal

    Exit Function
ErrorTrap:
    EndProgram
    
    
End Function



