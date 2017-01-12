Attribute VB_Name = "basTMSCode"
Option Explicit

Dim theitems As String, FilePath As String
Public NextPrevDates() As TMSPersonAndNextPrevDates
Public NextPrevDates_2009() As TMSPersonAndNextPrevDates_2009
Public NextPrevDates_2016() As TMSPersonAndNextPrevDates_2016

Dim mbBroOnlyTalkFound As Boolean

Public Type TMSPersonAndNextPrevDates
    ThePersonID As Long
    NextPrevInfo(23) As String
End Type

Public Type TMSPersonAndNextPrevDates_2009
    ThePersonID As Long
    NextPrevInfo(19) As String
End Type
Public Type TMSPersonAndNextPrevDates_2016
    ThePersonID As Long
    NextPrevInfo(25) As String
End Type

Public Function GetTMSItemsRecordset(TMSYear As Integer) As Recordset
Dim SQLStr As String
'
'Put all TMS items for supplied year into a recordset.
'

On Error GoTo ErrorTrap

    Select Case NewMtgArrangementStarted("01/01/" & TMSYear)
    Case CLM2016
        SQLStr = "SELECT AssignmentDate, " & _
                "TalkNo, " & _
                "TalkSeqNum, " & _
                "TalkTheme, " & _
                "SourceMaterial, " & _
                "DifficultyRating0to5, " & _
                "BrotherOnly, " & _
                "ItemsSeqNum, " & _
                "Sourceless " & _
                "FROM tblTMSItems " & _
                "WHERE Year(AssignmentDate) = " & TMSYear & _
                " ORDER BY AssignmentDate, TalkSeqNum, ItemsSeqNum"
    
    Case TMS2009
        SQLStr = "SELECT AssignmentDate, " & _
                "TalkNo, " & _
                "TalkSeqNum, " & _
                "TalkTheme, " & _
                "SourceMaterial, " & _
                "DifficultyRating0to5, " & _
                "BrotherOnly, " & _
                "ItemsSeqNum, " & _
                "Sourceless " & _
                "FROM tblTMSItems " & _
                "WHERE Year(AssignmentDate) = " & TMSYear & _
                " AND TalkNo <> 'P'" & _
                " ORDER BY AssignmentDate, TalkSeqNum"
    Case Else
        SQLStr = "SELECT AssignmentDate, " & _
                "TalkNo, " & _
                "TalkSeqNum, " & _
                "TalkTheme, " & _
                "SourceMaterial, " & _
                "DifficultyRating0to5, " & _
                "BrotherOnly, " & _
                "ItemsSeqNum " & _
                "FROM tblTMSItems " & _
                "WHERE Year(AssignmentDate) = " & TMSYear & _
                " ORDER BY AssignmentDate, TalkSeqNum"
    End Select
    
    Set GetTMSItemsRecordset = CMSDB.OpenRecordset(SQLStr, dbOpenDynaset)

    Exit Function
ErrorTrap:
    EndProgram
    
End Function


Public Sub GetTMSScheduleRecordset(TMSYear As Long, TMSMonth As Long, SchoolNum As Long, rstItems As Recordset, rstStudents As Recordset)
Dim SQLStr As String, StartDate As String, EndDate As String, PieceOfSQL As String
'
'Put all items for supplied year and month into a recordset.
' Include first week of the next month to tie in with the Service Meeting Schedule
'
'Then put all students assigned to items into another recset
'

On Error GoTo ErrorTrap

    Select Case True
    Case frmTMSScheduling.opt1Month
        StartDate = Format$(CStr(DateOfNthDay(vbMonday, TMSYear, TMSMonth, 1)), "mm/dd/yyyy")
    
        If TMSMonth < 12 Then
            EndDate = Format$(CStr(DateOfNthDay(vbMonday, TMSYear, TMSMonth + 1, 1)), "mm/dd/yyyy")
        Else
            EndDate = Format$(CStr(DateOfNthDay(vbMonday, TMSYear + 1, 1, 1)), "mm/dd/yyyy")
        End If
    Case frmTMSScheduling.opt1Year
        '
        'Use TMSYear and TMSMonth to find dates 6 months ago/ahead
        '
        StartDate = DateAdd("m", -6, DateSerial(TMSYear, TMSMonth, 1))
        StartDate = Format$(CStr(DateOfNthDay(vbMonday, year(CDate(StartDate)), Month(CDate(StartDate)), 1)), "mm/dd/yyyy")
        EndDate = DateAdd("m", 6, DateSerial(TMSYear, TMSMonth, 1))
        EndDate = Format$(CStr(DateOfNthDay(vbMonday, year(CDate(EndDate)), Month(CDate(EndDate)), 1)), "mm/dd/yyyy")
    End Select

    Select Case SchoolNum
    Case 1:
        If NewMtgArrangementStarted(StartDate) Then
            PieceOfSQL = " AND TalkNo NOT IN ('P') "
        Else
            PieceOfSQL = " "
        End If
    Case 2:
        If NewMtgArrangementStarted(StartDate) Then
            If gbAllowHighlightsInSch2 Then
                PieceOfSQL = " AND TalkNo NOT IN ('P','MR','R') "
            Else
                PieceOfSQL = " AND TalkNo NOT IN ('B', 'P','MR','R') "
            End If
        Else
            PieceOfSQL = " AND TalkNo NOT IN ('S', '1', 'B', 'P','MR','R') "
        End If
    Case 3:
        If NewMtgArrangementStarted(StartDate) Then
            If gbAllowHighlightsInSch3 Then
                PieceOfSQL = " AND TalkNo NOT IN ('P','MR','R') "
            Else
                PieceOfSQL = " AND TalkNo NOT IN ('B', 'P','MR','R') "
            End If
        Else
            PieceOfSQL = " AND TalkNo NOT IN ('S', '1', 'B', 'P','MR','R') "
        End If
    End Select
    
    SQLStr = "SELECT tblTMSItems.AssignmentDate, " & _
            "tblTMSItems.TalkNo, " & _
            "tblTMSItems.TaskNo, " & _
            "tblTMSItems.DifficultyRating0to5, " & _
            "tblTMSItems.BrotherOnly, " & _
            "tblTMSItems.SourceMaterial, " & _
            "tblTMSItems.TalkTheme, " & _
            "tblTMSItems.ItemsSeqNum, " & _
            "tblTMSItems.TalkSeqNum, " & _
            "tblTMSItems.Sourceless " & _
            "FROM tblTMSItems " & _
            "WHERE tblTMSItems.AssignmentDate BETWEEN #" & StartDate & "# AND #" & EndDate & "# " & _
            PieceOfSQL & _
            " ORDER BY tblTMSItems.AssignmentDate, tblTMSItems.TalkSeqNum, ItemsSeqNum"
    

    Set rstItems = CMSDB.OpenRecordset(SQLStr, dbOpenDynaset)


    SQLStr = "SELECT AssignmentDate, " & _
            "TalkNo, " & _
            "PersonID, " & _
            "Assistant1ID, " & _
            "SchoolNo, " & _
            "Setting, " & _
            "ScheduleSeqNum, " & _
            "SlipPrinted,  " & _
            "ItemsSeqNum  " & _
            "FROM tblTMSSchedule " & _
            "WHERE tblTMSSchedule.AssignmentDate BETWEEN #" & StartDate & "# AND #" & EndDate & "# " & _
            PieceOfSQL & _
            " AND SchoolNo = " & SchoolNum & _
            " ORDER BY AssignmentDate, TalkSeqNum, ItemsSeqNum "

            

    Set rstStudents = CMSDB.OpenRecordset(SQLStr, dbOpenDynaset)

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Public Function ItemAlreadyExists(AssignmentDate As Date, TalkNo As String) As Boolean
Dim rstFindItem As Recordset, SQLStr As String

On Error GoTo ErrorTrap

    SQLStr = "SELECT AssignmentDate, " & _
            "TalkNo, " & _
            "TalkTheme, " & _
            "SourceMaterial, " & _
            "DifficultyRating0to5, " & _
            "BrotherOnly, " & _
            "ItemsSeqNum " & _
            "FROM tblTMSItems " & _
            "WHERE AssignmentDate = #" & Format(AssignmentDate, "mm/dd/yyyy") & "# " & _
            "AND TalkNo = '" & TalkNo & "'"
    
    Set rstFindItem = CMSDB.OpenRecordset(SQLStr, dbOpenDynaset)
    
    If rstFindItem.BOF Then
        ItemAlreadyExists = False
    Else
        ItemAlreadyExists = True
    End If
    
    rstFindItem.Close
    



    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Sub FillTMSItemsGrid(Optional SavePos As Boolean = False, Optional RefreshRecSet As Boolean = False)
'
'"Items" are just the talk details. Assigned students are on the "Schedule"
'
Dim i As Byte, j As Integer, lSaveRow As Long

On Error GoTo ErrorTrap

'
'Populate grid from recordset
'
    If SavePos Then lSaveRow = frmTMSItems!flxTMSItems.TopRow
    
    If RefreshRecSet Then
        Set rstTMSItems = GetTMSItemsRecordset(frmTMSItems!cmbYear.ItemData(frmTMSItems!cmbYear.ListIndex))
    End If
    
     'clear grid's non-fixed rows
    frmTMSItems!flxTMSItems.Rows = 1
    
    With rstTMSItems
            
    If Not .BOF Then
        .MoveFirst
        j = 1
        Do Until .EOF
            frmTMSItems!flxTMSItems.Rows = j + 1
            frmTMSItems!flxTMSItems.TextMatrix(j, 0) = !AssignmentDate
            frmTMSItems!flxTMSItems.TextMatrix(j, 1) = !TalkNo
            frmTMSItems!flxTMSItems.TextMatrix(j, 2) = !TalkTheme
            frmTMSItems!flxTMSItems.TextMatrix(j, 3) = !SourceMaterial
            
            If !BrotherOnly = False Then
                frmTMSItems!flxTMSItems.TextMatrix(j, 4) = ""
            Else
                frmTMSItems!flxTMSItems.TextMatrix(j, 4) = "Yes"
            End If
               
            frmTMSItems!flxTMSItems.TextMatrix(j, 5) = !ItemsSeqNum
                      
            j = j + 1
            .MoveNext
        Loop
    End If
        
    End With
    
    With frmTMSItems!flxTMSItems
    For j = 0 To .Rows - 1
        .Row = j
        For i = 0 To .Cols - 1
            .col = i
            .CellForeColor = QBColor(0) 'set all cells to black text
        Next i
    Next j
    
    If .Rows > 1 Then 'ie more than just the 2 fixed rows
        .Row = 1
        .col = 0
        frmTMSItems.PreviousRow = .Row  'ready for click event when we mess about with row colours
    End If
    End With

    With frmTMSItems
    '
    'Allow all cells in date column to merge
    '
    !flxTMSItems.MergeCol(0) = True
    
    !cmdDeleteItem.Enabled = False
    !cmdEditItem.Enabled = False
    !chkDeleteAll = vbUnchecked
    
    If FormIsOpen("frmTMSAddItems") Then
        SetTopRowOfGridToWeek .flxTMSItems, 5, 0, CDate(frmTMSAddItems.txtAssignmentDate)
    Else
        If SavePos Then
            If lSaveRow <= frmTMSItems.flxTMSItems.Rows - 1 And lSaveRow > 0 Then
                frmTMSItems.flxTMSItems.TopRow = lSaveRow
            Else
                SetTopRowOfGridToWeek .flxTMSItems, 5, 0
            End If
        Else
            SetTopRowOfGridToWeek .flxTMSItems, 5, 0
        End If
    End If
    
    End With
    

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Public Sub FillTMSScheduleGrid()
'
'Put student names to the items....
'
Dim j As Long, SchoolNum As Long, PersonStore As String, PrevDate As Date
Dim IsAssemblyWeek As Boolean, IsMovedOralReview As Boolean, bIsCOVisitWeek As Boolean
Dim IsMemorialNight As Boolean, rsttmsquery As Recordset, TheCong As Long, OverrideWritten As Boolean
Dim NothingWritten As Boolean, IsOralReview As Boolean, NextWeekIsMovedOralReview As Boolean
Dim SQLStr As String, bNewMtgArrangement As MidweekMtgVersion, dFirstDate As Date, dLastDate As Date
Dim lNoSchools As Long, lColour1 As Long, lColour2 As Long, lColour As Long, bSwitch As Boolean
Dim lColourSpecial As Long, bIsMemorialWk As Boolean

On Error GoTo ErrorTrap

    lColour1 = vbWhite
    lColour2 = RGB(245, 245, 245)        'light grey
    lColourSpecial = RGB(252, 254, 211)  'yellow
    bSwitch = False


'
'Populate grid from recordsets
'
     'clear grid's non-fixed rows
    frmTMSScheduling!flxTMSSchedule.Rows = 1
    
    Select Case True
    Case frmTMSScheduling!opt1stSchool:
        SchoolNum = 1
    Case frmTMSScheduling!opt2ndSchool:
        SchoolNum = 2
    Case frmTMSScheduling!opt3rdSchool:
        SchoolNum = 3
    End Select
    
    TheCong = GlobalParms.GetValue("DefaultCong", "NumVal")
    
    With rstTMSSchedule
        
    PrevDate = 0
        
    If Not .BOF Then
        .MoveFirst
        j = 1
        Do Until .EOF
        
            If PrevDate <> !AssignmentDate Then
            
                bNewMtgArrangement = NewMtgArrangementStarted(CStr(!AssignmentDate))
                
                
            '
            'On date change, see if any assembly/moved oral review/CO visit will
            ' change the school scheduling for that week
            '
                OverrideWritten = False
                '
                'First check for moved Oral Review. An Oral Review will ONLY appear in
                ' tblTMSSchedule if it has been moved from it's normal week due to eg
                ' a CO visit....
                '
                Set rsttmsquery = CMSDB.OpenRecordset("SELECT AssignmentDate " & _
                                                   "FROM tblTMSSchedule " & _
                                                   "WHERE AssignmentDate = #" & Format(!AssignmentDate, "mm/dd/yyyy") & _
                                                    "# AND TalkNo = 'MR'" _
                                                   , dbOpenSnapshot)
                                                   
                If rsttmsquery.BOF Then
                    IsMovedOralReview = False
                Else
                    IsMovedOralReview = True
                End If
                
                '
                'Now check if NEXT week has Moved oral Review...
                '
                Set rsttmsquery = CMSDB.OpenRecordset("SELECT AssignmentDate " & _
                                                   "FROM tblTMSSchedule " & _
                                                    "WHERE AssignmentDate = #" & Format(!AssignmentDate + 7, "mm/dd/yyyy") & _
                                                    "# AND TalkNo = 'MR'" _
                                                   , dbOpenSnapshot)
                                                   
                If rsttmsquery.BOF Then
                    NextWeekIsMovedOralReview = False
                Else
                    NextWeekIsMovedOralReview = True
                End If
                
                '
                'Is this Oral Review week?
                '
                Set rsttmsquery = CMSDB.OpenRecordset("SELECT AssignmentDate " & _
                                                   "FROM tblTMSItems " & _
                                                   "WHERE AssignmentDate = #" & Format(!AssignmentDate, "mm/dd/yyyy") & _
                                                   "# AND TalkNo = 'R'" _
                                                   , dbOpenSnapshot)
                                                   
                If rsttmsquery.BOF Then
                    IsOralReview = False
                Else
                    IsOralReview = True
                End If
                                                   
                '
                'Now check Calendar for assembly this week...
                '
                IsAssemblyWeek = IsCircuitOrDistrictAssemblyWeek(!AssignmentDate)
                                                   
                '
                'Now check Calendar for CO Visit this week...
                '
                bIsCOVisitWeek = IsCOVisitWeek(!AssignmentDate)
                                                                
            
                '
                'Now check tblTMSSchedule for Memorial on meeting night.
                '

                IsMemorialNight = IsMemorialDay(GetDateOfGivenDay(!AssignmentDate, GlobalParms.GetValue("MidWeekMeetingDay", "NumVal")))
                bIsMemorialWk = IsMemorialWeek(!AssignmentDate)
                
            
                bSwitch = Not bSwitch
                
                If bSwitch Then
                    If Not (IsMemorialNight Or bIsMemorialWk Or IsAssemblyWeek Or bIsCOVisitWeek) Then
                        lColour = lColour1
                    Else
                        lColour = lColourSpecial
                    End If
                Else
                    If Not (IsMemorialNight Or bIsMemorialWk Or IsAssemblyWeek Or bIsCOVisitWeek) Then
                        lColour = lColour2
                    Else
                        lColour = lColourSpecial
                    End If
                End If
                
            End If
            
            
            
            
    '
    'Get all items for specified period from rstTMSSchedule
    '
            
            Select Case True
            Case IsAssemblyWeek
                If Not OverrideWritten Then
                    FillCells j, "A", "Assembly", SchoolNum, lColour
                    NothingWritten = False
                    OverrideWritten = True
                Else
                    NothingWritten = True
                End If
'            Case IsMemorialNight
'                If Not OverrideWritten Then
'                    FillCells j, "M", "Memorial", SchoolNum
'                    NothingWritten = False
'                    OverrideWritten = True
'                Else
'                    NothingWritten = True
'                End If
            Case bIsCOVisitWeek
                Select Case bNewMtgArrangement
                Case CLM2016
                
                    FillCells j, "", "", SchoolNum, lColour
                    NothingWritten = False
                
                Case TMS2009
                    Select Case !TalkNo
                    Case "P"
                        FillCells j, "", "", SchoolNum, lColour
                    Case "B"
                        'put 'CO Visit' row prior to Prayer, just so it's visible
                        If Not OverrideWritten Then
                            FillCells j, "CO", "Circuit Visit", SchoolNum, lColour
                            OverrideWritten = True
                        End If
                        'All items remain for CO visit
                        FillCells j, "", "", SchoolNum, lColour
                        NothingWritten = False
                        If NextWeekIsMovedOralReview Then
                            'Force in 1,2,3 which won't exist on schedule
                            j = j + 1
                            FillCells j, "1", "", SchoolNum, lColour
                            NothingWritten = False
                            j = j + 1
                            FillCells j, "2", "", SchoolNum, lColour
                            NothingWritten = False
                            j = j + 1
                            FillCells j, "3", "", SchoolNum, lColour
                            NothingWritten = False
                        End If
                    Case "R"
                        If Not NextWeekIsMovedOralReview Then
                            FillCells j, "", "", SchoolNum, lColour
                            NothingWritten = False
                        Else
                            NothingWritten = True
                        End If
                    Case Else
                        FillCells j, "", "", SchoolNum, lColour
                        NothingWritten = False
                    End Select
                Case Else 'pre-2009
                    Select Case !TalkNo
                    Case "P", "B", "1"
                        FillCells j, "", "", SchoolNum, lColour
                        NothingWritten = False
                    Case "S"
                        FillCells j, "", "", SchoolNum, lColour
                        NothingWritten = False
                        
                        If NextWeekIsMovedOralReview Then
                            'Put next week's #1 this week
                            j = j + 1
                            FillCells j, "1", "", SchoolNum, lColour
                            NothingWritten = False
                        End If
                    Case Else
                        If Not OverrideWritten Then
                            FillCells j, "CO", "Circuit Visit", SchoolNum, lColour
                            NothingWritten = False
                            OverrideWritten = True
                        Else
                            NothingWritten = True
                        End If
                    End Select
                End Select
            Case IsMovedOralReview
                Select Case !TalkNo
                Case "P", "S", "B"
                    FillCells j, "", "", SchoolNum, lColour
                    NothingWritten = False
                Case "1"
                    NothingWritten = True
                Case Else
                    If Not OverrideWritten Then
                        FillCells j, "MR", "Oral Review (Moved)", SchoolNum, lColour
                        NothingWritten = False
                        OverrideWritten = True
                    Else
                        NothingWritten = True
                    End If
                End Select
            Case Else
                FillCells j, "", "", SchoolNum, lColour
                NothingWritten = False
            End Select
            
            If Not NothingWritten Then
                j = j + 1
            End If
            PrevDate = !AssignmentDate
            .MoveNext
        Loop
    End If
        
    End With
    
    With frmTMSScheduling!flxTMSSchedule
    
    If .Rows > 1 Then 'ie more than just the 1 fixed row
        .Row = 1
        .col = 0
        frmTMSScheduling.PreviousRow = .Row  'ready for click event when we mess about with row colours
    End If
    
    End With
    
    'now we want to highlight which schools are used during selected period
    On Error Resume Next
    With rstTMSSchedule
    
    .MoveFirst
    If Err.number = 0 Then
        dFirstDate = !AssignmentDate
        .MoveLast
        dLastDate = !AssignmentDate
        
        If TheTMS.SchoolNoActiveBetweenDates(1, dFirstDate, dLastDate) Then
            frmTMSScheduling.lblSch1.ForeColor = vbRed
        Else
            frmTMSScheduling.lblSch1.ForeColor = vbBlack
        End If
        If TheTMS.SchoolNoActiveBetweenDates(2, dFirstDate, dLastDate) Then
            frmTMSScheduling.lblSch2.ForeColor = vbRed
        Else
            frmTMSScheduling.lblSch2.ForeColor = vbBlack
        End If
        If TheTMS.SchoolNoActiveBetweenDates(3, dFirstDate, dLastDate) Then
            frmTMSScheduling.lblSch3.ForeColor = vbRed
        Else
            frmTMSScheduling.lblSch3.ForeColor = vbBlack
        End If
        
    End If
    
    End With
    On Error GoTo ErrorTrap
    
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
            
Public Sub FillCells(TheRow As Long, _
                     OverrideTalkNo As String, _
                     OverrideEventName As String, _
                     SchoolNum As Long, _
                     RowColour As Long)
Dim j As Long, PersonStore As String, NewArrVer As MidweekMtgVersion, PersonStore2 As String
Dim i As Long

On Error GoTo ErrorTrap
            
    j = TheRow
    
    
    With rstTMSSchedule
    
    frmTMSScheduling!flxTMSSchedule.Rows = j + 1
    
    frmTMSScheduling!flxTMSSchedule.Row = j
    For i = 0 To frmTMSScheduling!flxTMSSchedule.Cols - 1
        frmTMSScheduling!flxTMSSchedule.col = i
        frmTMSScheduling!flxTMSSchedule.CellBackColor = RowColour
    Next i
    
    
    frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 0) = Format$(!AssignmentDate, "dd/mm/yyyy")
    
    NewArrVer = NewMtgArrangementStarted(CStr(!AssignmentDate))
    
    Select Case OverrideTalkNo
    Case ""
        frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 1) = !TalkNo
        frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 4) = !ItemsSeqNum
        '
        'Now search rstTMSStudents to determine whether a student has been assigned
        ' to rstTMSSchedule's current item. If so, put them on grid.
        '
        rstTMSStudents.FindFirst "ItemsSeqNum = " & !ItemsSeqNum
'        rstTMSStudents.FindFirst "Assignmentdate = #" & Format$(!AssignmentDate, "mm/dd/yyyy") & _
'                    "# AND TalkNo = '" & !TalkNo & "' AND SchoolNo = " & SchoolNum & _
'                    " AND ItemsSeqNum = " & !ItemsSeqNum
                  
        With rstTMSStudents
        If .NoMatch Then
            frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 2) = ""
            frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 3) = ""
        Else
            PersonStore = CongregationMember.FullName(!PersonID, False)
            If Left$(PersonStore, 1) = "?" Then
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 2) = ""
            Else
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 2) = PersonStore
                '
                'Highlight name in blue if it matches Search criteria for student
                '
                If lnkTMSScheduleSearchStarted Then
                    If lnkTMSScheduleSearch(frmTMSScheduling.CurrentSearchIndex).TheAssignmentDate = !AssignmentDate And _
                       lnkTMSScheduleSearch(frmTMSScheduling.CurrentSearchIndex).TheSchool = SchoolNum And _
                       lnkTMSScheduleSearch(frmTMSScheduling.CurrentSearchIndex).ThePersonID = !PersonID Then
                            frmTMSScheduling!flxTMSSchedule.Row = j
                            frmTMSScheduling!flxTMSSchedule.col = 2
                            frmTMSScheduling!flxTMSSchedule.CellForeColor = vbBlue
'                            If j > 10 Then
'                                frmTMSScheduling.TheTopRow = j - 10
                            If j > 31 Then
                                frmTMSScheduling.TheTopRow = j - 27
                            Else
                                frmTMSScheduling.TheTopRow = j
                            End If
                    End If
                End If
            End If
            frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 6) = !ScheduleSeqNum
            
            PersonStore = CongregationMember.FullName(!Assistant1ID, False)
            If Left$(PersonStore, 1) = "?" Then
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 3) = ""
            Else
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 3) = PersonStore
                '
                'Highlight name in red if it matches Search criteria for assistant
                '
                If lnkTMSScheduleSearchStarted Then
                    If lnkTMSScheduleSearch(frmTMSScheduling.CurrentSearchIndex).TheAssignmentDate = !AssignmentDate And _
                       lnkTMSScheduleSearch(frmTMSScheduling.CurrentSearchIndex).TheSchool = SchoolNum And _
                       lnkTMSScheduleSearch(frmTMSScheduling.CurrentSearchIndex).ThePersonID = !Assistant1ID Then
                            frmTMSScheduling!flxTMSSchedule.Row = j
                            frmTMSScheduling!flxTMSSchedule.col = 3
                            frmTMSScheduling!flxTMSSchedule.CellForeColor = vbRed
'                            If j > 10 Then
'                                frmTMSScheduling.TheTopRow = j - 10
                            If j > 31 Then
                                frmTMSScheduling.TheTopRow = j - 27
                            Else
                                frmTMSScheduling.TheTopRow = j
                            End If
                    End If
                End If
            End If
            
        End If
        
        End With
    Case Else 'This is an override...
        If NewArrVer = CLM2016 Then
            frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 1) = OverrideTalkNo
            frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 4) = -1
            frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 2) = OverrideEventName
        ElseIf NewArrVer = TMS2009 Then
            rstTMSStudents.FindFirst "Assignmentdate = #" & Format$(!AssignmentDate, "mm/dd/yyyy") & _
                        "# AND TalkNo = '" & OverrideTalkNo & "' AND SchoolNo = " & SchoolNum
            If rstTMSStudents.NoMatch Then
                PersonStore = "?"
                PersonStore2 = "?"
            Else
                PersonStore = CongregationMember.FullName(rstTMSStudents!PersonID)
                PersonStore2 = CongregationMember.FullName(rstTMSStudents!Assistant1ID)
            End If
            Select Case OverrideTalkNo
            Case "1", "2", "3" 'CO Visit week on what should be Oral Review week (Review moved to next week)
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 1) = OverrideTalkNo
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 4) = -1
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 6) = rstTMSStudents!ScheduleSeqNum
                If Left$(PersonStore, 1) = "?" Then
                    frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 2) = ""
                Else
                    frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 2) = PersonStore
                End If
                If Left$(PersonStore2, 1) = "?" Then
                    frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 3) = ""
                Else
                    frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 3) = PersonStore2
                End If
            Case "MR"
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 1) = OverrideTalkNo
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 4) = -1
                
                If rstTMSStudents!PersonID > 0 Then
                    frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 2) = PersonStore
                Else
                    frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 2) = OverrideEventName
                End If
            Case Else
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 1) = OverrideTalkNo
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 4) = -1
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 2) = OverrideEventName
            End Select
        ElseIf NewArrVer = Pre2009 Then
            Select Case OverrideTalkNo
            Case "1" 'CO Visit week on what should be Oral Review week (Review moved to next week)
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 1) = OverrideTalkNo
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 4) = -1
                rstTMSStudents.FindFirst "Assignmentdate = #" & Format$(!AssignmentDate, "mm/dd/yyyy") & _
                            "# AND TalkNo = '1' AND SchoolNo = " & SchoolNum
                If rstTMSStudents.NoMatch Then
                    PersonStore = "?"
                Else
                    PersonStore = CongregationMember.FullName(rstTMSStudents!PersonID)
                End If
                If Left$(PersonStore, 1) = "?" Then
                    frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 2) = ""
                Else
                    frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 2) = PersonStore
                End If
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 3) = ""
            Case Else
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 1) = OverrideTalkNo
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 4) = -1
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 2) = OverrideEventName
            End Select
        End If
    End Select
    
    '
    'Now show whether assignments slip has been printed. "Y" is yes (obviously),
    ' "N" is no where the slip SHOULD be printed. Blank indicates that the
    ' assignment isn't one for which a slip is required...eg "P" or "R" etc
    '
    Select Case NewArrVer
    Case CLM2016
    
        If Not rstTMSStudents.NoMatch Then
            If rstTMSStudents!PersonID > 0 Then
                If rstTMSStudents!SlipPrinted = True Then
                    frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = "Y"
                Else
                    frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = "N"
                End If
            Else
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = ""
            End If
        Else
            frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = ""
        End If
    
    Case TMS2009
        If Not rstTMSStudents.NoMatch Then
            Select Case !TalkNo
            Case "1", "B", "2", "3"
                If (OverrideTalkNo = "" Or _
                    OverrideTalkNo = "1" Or _
                    OverrideTalkNo = "2" Or _
                    OverrideTalkNo = "3") _
                    And rstTMSStudents!PersonID > 0 Then
                    
                    If rstTMSStudents!SlipPrinted = True Then
                        frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = "Y"
                    Else
                        frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = "N"
                    End If
                    
                Else
                    frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = ""
                End If
            Case "R", "MR"
                If (gbPrintSlipsForOralReviewReader And (!TalkNo = "R" Or !TalkNo = "MR")) Then
                    If rstTMSStudents!SlipPrinted = True Then
                        frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = "Y"
                    Else
                        frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = "N"
                    End If
                Else
                    frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = ""
                End If
            Case Else
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = ""
            End Select
        Else
            frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = ""
        End If
    Case Else
        If Not rstTMSStudents.NoMatch Then
            Select Case !TalkNo
            Case "S", "1", "B", "2", "3", "4"
                If (OverrideTalkNo = "" Or OverrideTalkNo = "1") _
                    And rstTMSStudents!PersonID > 0 Then
                    If rstTMSStudents!SlipPrinted = True Then
                        frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = "Y"
                    Else
                        frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = "N"
                    End If
                Else
                    frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = ""
                End If
            Case Else
                frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = ""
            End Select
        Else
            frmTMSScheduling!flxTMSSchedule.TextMatrix(j, 5) = ""
        End If
    End Select
    
    End With
    
    
    ''
    'Highlight row for current week
    '
    With frmTMSScheduling!flxTMSSchedule
    If frmTMSScheduling.HighltCurrWk Then
        If CDate(rstTMSSchedule!AssignmentDate) = GetDateOfGivenDay(Format(Now, "dd/mm/yyyy"), vbMonday, False) Then
            .Row = TheRow
            .col = 0
            .CellFontBold = True
            .CellFontSize = .CellFontSize - 5
            
            'added this line 21/06/2011 so SQ form opens on correct date
            frmTMSScheduling.FormDate = CDate(rstTMSSchedule!AssignmentDate)
        End If
    End If
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Public Function FillTMSStudentGrid(TheGrid As MSFlexGrid) As Boolean
    Select Case NewMtgArrangementStarted(frmTMSInsertStudent.GetAssignmentDate)
    Case CLM2016
        FillTMSStudentGrid = FillTMSStudentGrid_2016(TheGrid)
    Case TMS2009
        FillTMSStudentGrid_2009 TheGrid
        FillTMSStudentGrid = True
    Case Else
        FillTMSStudentGrid_2002 TheGrid
        FillTMSStudentGrid = True
    End Select
End Function
           

Public Sub FillTMSStudentGrid_2002(TheGrid As MSFlexGrid)
'
'This grid allows replacement student to be picked.
'
'This proc is executed for both flxInsertStudent and flxInsertAssistant, depending on
' the parm. When executed for flsInsertStudent, Prev/Next dates are acquired from the DB]
' via the TMSPrevDateFAST and TMSNextDateFAST methods. These are stored in array NextPrevDates().
' This array has elements of Type TMSPersonAndNextPrevDates, one for each Student to
' be dislayed on the grid. Each student has a set of 23 elements in the UDT, each element storing
' a Next/Prev date/school. This stored data is then used to populate flxInsertAssistant.
'
Dim CurrentAssignmentDate As Date, CurrentTalkNo As String
Dim rstStudentList As Recordset, SQLStr As String, BroOnlyFlag As Boolean, SchoolNum As Long
Dim SchoolNumForSQL As Long, CurrentTalkNoForSQL As Long, i As Long, PersonStore As String
Dim ActualTalkDate As Date, OrderingSQL As String, PrevDates() As Variant, n As Integer
Dim a As String, BroOnlySQL As String, SuspendDateSQL As String, NextDates() As Variant
Dim AverageWeighting As Long, rstAvWtg As Recordset, NoPersons As Long, j As Long
Dim TheGridName As String, TaskSQL As String

On Error GoTo ErrorTrap

    Screen.MousePointer = vbHourglass
    
    TheGridName = TheGrid.Name

    With frmTMSScheduling
    '
    'Need to use the actual TaskNo for the school-no a student is able to have talks in
    ' Required for the SQL later...
    '
    Select Case True
    Case !opt1stSchool:
        SchoolNum = 1
        SchoolNumForSQL = 44
    Case !opt2ndSchool:
        SchoolNum = 2
        SchoolNumForSQL = 45
    Case !opt3rdSchool:
        SchoolNum = 3
        SchoolNumForSQL = 46
    End Select
    
    Select Case !lblBroOnly.Caption
    Case "Yes":
        BroOnlyFlag = True
    Case "":
        BroOnlyFlag = False
    End Select
    
    CurrentTalkNo = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 1)
    CurrentAssignmentDate = CDate(!flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 0))
    
    If TheGridName = "flxInsertStudent" Then 'only need to do this bit once
        frmTMSInsertStudent.FillAssignmentListBox CurrentAssignmentDate
    End If
    
    If TheGridName = "flxInsertStudent" Then
        frmTMSInsertStudent!lblCurrentStudent = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 2)
        frmTMSInsertStudent!lblCurrentAssistant = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 3)
        frmTMSInsertStudent!lblCurrentStudent2 = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 2)
        frmTMSInsertStudent!lblCurrentAssistant2 = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 3)
        frmTMSInsertStudent!lblAssignmentDate = Format$(CurrentAssignmentDate, "dd/mm/yyyy")
        frmTMSInsertStudent!lblAssignmentDate2 = Format$(CurrentAssignmentDate, "dd/mm/yyyy")
        frmTMSInsertStudent!lblSchool1 = SchoolNum
        frmTMSInsertStudent!lblSchool2 = SchoolNum
        frmTMSInsertStudent!lblThemeDisplay = !lblThemeDisplay
        frmTMSInsertStudent!lblThemeDisplay2 = !lblThemeDisplay
        frmTMSInsertStudent!lblSetting = !lblSetting
        frmTMSInsertStudent!lblSetting2 = !lblSetting
        frmTMSInsertStudent!lblSourceDisplay = !lblSourceDisplay
        frmTMSInsertStudent!lblSourceDisplay2 = !lblSourceDisplay
        frmTMSInsertStudent!lblBroOnly = !lblBroOnly
        
        With frmTMSInsertStudent
        '
        'Display TalkNo description on form
        '
        Select Case CurrentTalkNo
        Case "P":
            !lblAssignment = "Opening Prayer"
            !lblAssignment2 = "Opening Prayer"
        Case "S":
            !lblAssignment = "Speech Quality Talk"
            !lblAssignment2 = "Speech Quality Talk"
        Case "1":
            !lblAssignment = "Talk No 1"
            !lblAssignment2 = "Talk No 1"
        Case "B":
            !lblAssignment = "Bible Highlights"
            !lblAssignment2 = "Bible Highlights"
        Case "2":
            !lblAssignment = "Reading Assignment"
            !lblAssignment2 = "Reading Assignment"
        Case "3":
            !lblAssignment = "Talk No 3"
            !lblAssignment2 = "Talk No 3"
        Case "4":
            !lblAssignment = "Talk No 4"
            !lblAssignment2 = "Talk No 4"
        Case Else:
            !lblAssignment = ""
            !lblAssignment2 = ""
        
        End Select
        End With
        
        With frmTMSInsertStudent
        .FormAssignmentDate = CurrentAssignmentDate
        .FormSchoolNo = SchoolNum
        .FormTalkNo = CurrentTalkNo
        .FormStudent = TheTMS.GetTMSStudent(CurrentAssignmentDate, CurrentTalkNo, SchoolNum, 0)
        .FormAssistant = TheTMS.GetTMSAssistant(CurrentAssignmentDate, CurrentTalkNo, SchoolNum, 0)
        End With
    End If
    
    End With
       
    
    ActualTalkDate = Format(GetDateOfGivenDay(CurrentAssignmentDate, _
                        GlobalParms.GetValue("MidWeekMeetingDay", "NumVal")), "mm/dd/yyyy")
    
    
    Select Case frmTMSInsertStudent.chkBrothersOnly.value
    Case vbChecked
        BroOnlySQL = " AND c.GenderMF = 'M' "
    Case vbUnchecked
        BroOnlySQL = ""
    End Select
    
    Select Case frmTMSInsertStudent.chkShowEandMSOnly.value
    Case vbChecked
        If CurrentTalkNo = "4" Then
            TaskSQL = " Task IN (33,34,35) "
        Else
            TaskSQL = " Task = " & rstTMSSchedule!TaskNo & " "
        End If
    Case vbUnchecked
        TaskSQL = " Task = " & rstTMSSchedule!TaskNo & " "
    End Select
    
    SuspendDateSQL = " AND NOT EXISTS (SELECT 1 " & _
                    "FROM tblTaskPersonSuspendDates f " & _
                    "WHERE Task = " & rstTMSSchedule!TaskNo & _
                    " AND a.Person = f.Person " & _
                    " AND (SuspendStartDate <= #" & ActualTalkDate & "# AND SuspendEndDate >= #" & ActualTalkDate & "#)) "
                        
    Select Case frmTMSInsertStudent.optIntelligent.value
    Case True
    
        '
        'Prayer weighting is separate to other talk weighting. Otherwise
        ' a brother that does prayers and talks would end up doing less talks
        ' than brothers that don't do prayers....
        '
        If CurrentTalkNo <> "P" Then
            OrderingSQL = "ORDER BY b.TMSWeighting"
        Else
            OrderingSQL = "ORDER BY b.TMSPrayerWeighting"
        End If
        
        
        If TheGridName = "flxInsertStudent" Then
            '
            'ie The InsertStudent grid is currently being populated
            '
            'Take account of suspend dates
            '
            Set rstStudentList = CMSDB.OpenRecordset("SELECT DISTINCT a.Person, " & _
                                                "b.TMSWeighting, " & _
                                                "b.TMSPrayerWeighting " & _
                                                "FROM ((tblTaskPersonSuspendDates a " & _
                                                "INNER JOIN tblTMSWeightings b ON " & _
                                                "a.Person = b.PersonID) " & _
                                                "INNER JOIN tblNameAddress c ON " & _
                                                " (c.ID = b.PersonID)) " & _
                                                "WHERE " & TaskSQL & _
                                                SuspendDateSQL & " AND Active = TRUE " & _
                                                " AND EXISTS  (SELECT 1 " & _
                                                                "FROM tblTaskPersonSuspendDates d " & _
                                                                " WHERE d.Person = a.Person " & _
                                                                " AND Task = " & SchoolNumForSQL & ") " & _
                                                BroOnlySQL & OrderingSQL, dbOpenForwardOnly)
                                                                                                
        Else
'            '
'            'ie The InsertAssistant grid is currently being populated
'            '


            '
            'Take account of suspend dates
            '
            Set rstStudentList = CMSDB.OpenRecordset("SELECT DISTINCT a.Person, " & _
                                                "b.TMSWeighting " & _
                                                "FROM ((tblTaskPersonSuspendDates a " & _
                                                "INNER JOIN tblTMSWeightings b ON " & _
                                                " a.Person = b.PersonID) " & _
                                                ") INNER JOIN tblNameAddress c ON c.ID = a.Person " & _
                                                "WHERE Task IN (42,43)" & _
                                                SuspendDateSQL & " AND Active = TRUE " & _
                                                " AND EXISTS  (SELECT 1 " & _
                                                                "FROM tblTaskPersonSuspendDates d " & _
                                                                " WHERE d.Person = a.Person " & _
                                                                " AND Task = " & SchoolNumForSQL & ") " & _
                                                "ORDER BY b.TMSWeighting", dbOpenForwardOnly)
                                                                
        End If
        
    Case False
    
        If TheGridName = "flxInsertStudent" Then
            '
            'ie The InsertStudent grid is currently being populated
            '
            'Take account of suspend dates
            '
            Set rstStudentList = CMSDB.OpenRecordset("SELECT DISTINCT a.Person, c.FirstName, " & _
                                                    "c.MiddleName, c.LastName " & _
                                                "FROM tblTaskPersonSuspendDates a " & _
                                                "INNER JOIN tblNameAddress c ON " & _
                                                " (a.Person = c.ID) " & _
                                                "WHERE " & TaskSQL & _
                                                SuspendDateSQL & " AND Active = TRUE " & _
                                                " AND EXISTS  (SELECT 1 " & _
                                                                "FROM tblTaskPersonSuspendDates d " & _
                                                                " WHERE d.Person = a.Person " & _
                                                                " AND Task = " & SchoolNumForSQL & ") " & _
                                                BroOnlySQL & _
                                                " ORDER BY LastName, FirstName, MiddleName", dbOpenForwardOnly)
                                                                                                
        Else
            '
            'ie The InsertAssistant grid is currently being populated
            '
            'Take account of suspend dates
            '
            Set rstStudentList = CMSDB.OpenRecordset("SELECT DISTINCT a.Person, c.FirstName, " & _
                                                    "c.MiddleName, c.LastName " & _
                                                "FROM (tblTaskPersonSuspendDates a " & _
                                                "INNER JOIN tblNameAddress c ON " & _
                                                " a.Person = c.ID) " & _
                                                "WHERE Task IN (42,43) " & _
                                                SuspendDateSQL & " AND Active = TRUE " & _
                                                " AND EXISTS  (SELECT 1 " & _
                                                                "FROM tblTaskPersonSuspendDates d " & _
                                                                " WHERE d.Person = a.Person " & _
                                                                " AND Task = " & SchoolNumForSQL & ") " & _
                                                " ORDER BY LastName, FirstName, MiddleName", dbOpenForwardOnly)
                                                                
        End If
    End Select
    
    
    i = 2 'row of grid
    j = 0 'position on array storing each student's prev/next dates
    
    '
    'Fill grid with last/prev dates. Actual columns used is determined by the PrevXxxx
    ' variables which are set in frmTMSInsertStudent.SetUpGrids according to the
    ' talkno selected on frmTMSScheduling.
    '
    
    TheGrid.Rows = 2
    With rstStudentList
    If Not .BOF Then
        Do While Not .EOF
            TheGrid.Rows = i + 1
            
            If CongregationMember.TMSStudentOnThisWeekFAST(!Person, CurrentAssignmentDate) Then
                frmTMSInsertStudent.GridRowBackColourToGrey TheGrid, i, True, False
            Else
                frmTMSInsertStudent.GridRowBackColourToGrey TheGrid, i, False, False
            End If
                        
            TheGrid.TextMatrix(i, 0) = CongregationMember.LastFirstNameMiddleName(!Person)
            
            TheGrid.TextMatrix(i, TheGrid.Cols - 1) = !Person 'NameID used for programmatic purposes.
            
            If frmTMSInsertStudent.chkShowNextPrevDates.value = vbChecked Then
                If TheGridName = "flxInsertStudent" Then
                
                    ReDim Preserve NextPrevDates(j)
                    NextPrevDates(j).ThePersonID = !Person
                    
                    '
                    'Put next/prev dates into an array
                    '
                    NextPrevDates(j).NextPrevInfo(0) = CongregationMember.TMSPrevDateFAST(!Person, "P", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(1) = CongregationMember.TMSPrevDateFAST(!Person, "S", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(2) = CongregationMember.TMSPrevDateFAST(!Person, "1", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(3) = CongregationMember.TMSPrevDateFAST(!Person, "B", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(4) = CongregationMember.TMSPrevDateFAST(!Person, "2", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(5) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates(j).NextPrevInfo(6) = CongregationMember.TMSPrevDateFAST(!Person, "3", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(7) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates(j).NextPrevInfo(8) = CongregationMember.TMSPrevDateFAST(!Person, "4", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(9) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates(j).NextPrevInfo(10) = CongregationMember.TMSPrevDateFAST(!Person, "Asst", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(11) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates(j).NextPrevInfo(12) = CongregationMember.TMSNextDateFAST(!Person, "P", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(13) = CongregationMember.TMSNextDateFAST(!Person, "S", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(14) = CongregationMember.TMSNextDateFAST(!Person, "1", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(15) = CongregationMember.TMSNextDateFAST(!Person, "B", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(16) = CongregationMember.TMSNextDateFAST(!Person, "2", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(17) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates(j).NextPrevInfo(18) = CongregationMember.TMSNextDateFAST(!Person, "3", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(19) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates(j).NextPrevInfo(20) = CongregationMember.TMSNextDateFAST(!Person, "4", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(21) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates(j).NextPrevInfo(22) = CongregationMember.TMSNextDateFAST(!Person, "Asst", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(23) = CongregationMember.GetTMSSchoolNoForInsertForm
                    
                Else
                    '
                    'This is Assistant-grid. Get prev/next dates from array.
                    '
                    'Locate current student in array to get their prev/next dates
                    '
                    For j = 0 To UBound(NextPrevDates)
                        If NextPrevDates(j).ThePersonID = !Person Then
                            Exit For
                        End If
                    Next j
                    
                End If
                 
                '
                'Put array contents into student grid. If Student is not in array
                ' (ie because they're only a householder, so wouldn't have been included
                ' when processing Student Grid), then get dates from DB
                '
                If j <= UBound(NextPrevDates) Then
                    With TheGrid
                    .TextMatrix(i, PrevPrayer) = NextPrevDates(j).NextPrevInfo(0)
                    .TextMatrix(i, PrevSQ) = NextPrevDates(j).NextPrevInfo(1)
                    .TextMatrix(i, PrevNo1) = NextPrevDates(j).NextPrevInfo(2)
                    .TextMatrix(i, PrevBH) = NextPrevDates(j).NextPrevInfo(3)
                    .TextMatrix(i, PrevNo2) = NextPrevDates(j).NextPrevInfo(4)
                    .TextMatrix(i, PrevNo2School) = NextPrevDates(j).NextPrevInfo(5)
                    .TextMatrix(i, PrevNo3) = NextPrevDates(j).NextPrevInfo(6)
                    .TextMatrix(i, PrevNo3School) = NextPrevDates(j).NextPrevInfo(7)
                    .TextMatrix(i, PrevNo4) = NextPrevDates(j).NextPrevInfo(8)
                    .TextMatrix(i, PrevNo4School) = NextPrevDates(j).NextPrevInfo(9)
                    .TextMatrix(i, PrevAsst) = NextPrevDates(j).NextPrevInfo(10)
                    .TextMatrix(i, PrevAsstSchool) = NextPrevDates(j).NextPrevInfo(11)
                    .TextMatrix(i, NextPrayer) = NextPrevDates(j).NextPrevInfo(12)
                    .TextMatrix(i, NextSQ) = NextPrevDates(j).NextPrevInfo(13)
                    .TextMatrix(i, NextNo1) = NextPrevDates(j).NextPrevInfo(14)
                    .TextMatrix(i, NextBH) = NextPrevDates(j).NextPrevInfo(15)
                    .TextMatrix(i, NextNo2) = NextPrevDates(j).NextPrevInfo(16)
                    .TextMatrix(i, NextNo2School) = NextPrevDates(j).NextPrevInfo(17)
                    .TextMatrix(i, NextNo3) = NextPrevDates(j).NextPrevInfo(18)
                    .TextMatrix(i, NextNo3School) = NextPrevDates(j).NextPrevInfo(19)
                    .TextMatrix(i, NextNo4) = NextPrevDates(j).NextPrevInfo(20)
                    .TextMatrix(i, NextNo4School) = NextPrevDates(j).NextPrevInfo(21)
                    .TextMatrix(i, NextAsst) = NextPrevDates(j).NextPrevInfo(22)
                    .TextMatrix(i, NextAsstSchool) = NextPrevDates(j).NextPrevInfo(23)
                    End With
                Else
                    '
                    'Go to   s l o w   DB. Add new student to array, and get the dates....
                    '
                    ReDim Preserve NextPrevDates(j)
                    NextPrevDates(j).ThePersonID = !Person
                    
                    NextPrevDates(j).NextPrevInfo(0) = CongregationMember.TMSPrevDateFAST(!Person, "P", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(1) = CongregationMember.TMSPrevDateFAST(!Person, "S", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(2) = CongregationMember.TMSPrevDateFAST(!Person, "1", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(3) = CongregationMember.TMSPrevDateFAST(!Person, "B", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(4) = CongregationMember.TMSPrevDateFAST(!Person, "2", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(5) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates(j).NextPrevInfo(6) = CongregationMember.TMSPrevDateFAST(!Person, "3", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(7) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates(j).NextPrevInfo(8) = CongregationMember.TMSPrevDateFAST(!Person, "4", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(9) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates(j).NextPrevInfo(10) = CongregationMember.TMSPrevDateFAST(!Person, "Asst", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(11) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates(j).NextPrevInfo(12) = CongregationMember.TMSNextDateFAST(!Person, "P", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(13) = CongregationMember.TMSNextDateFAST(!Person, "S", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(14) = CongregationMember.TMSNextDateFAST(!Person, "1", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(15) = CongregationMember.TMSNextDateFAST(!Person, "B", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(16) = CongregationMember.TMSNextDateFAST(!Person, "2", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(17) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates(j).NextPrevInfo(18) = CongregationMember.TMSNextDateFAST(!Person, "3", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(19) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates(j).NextPrevInfo(20) = CongregationMember.TMSNextDateFAST(!Person, "4", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(21) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates(j).NextPrevInfo(22) = CongregationMember.TMSNextDateFAST(!Person, "Asst", CurrentAssignmentDate)
                    NextPrevDates(j).NextPrevInfo(23) = CongregationMember.GetTMSSchoolNoForInsertForm
                                        
                    With TheGrid
                    .TextMatrix(i, PrevPrayer) = NextPrevDates(j).NextPrevInfo(0)
                    .TextMatrix(i, PrevSQ) = NextPrevDates(j).NextPrevInfo(1)
                    .TextMatrix(i, PrevNo1) = NextPrevDates(j).NextPrevInfo(2)
                    .TextMatrix(i, PrevBH) = NextPrevDates(j).NextPrevInfo(3)
                    .TextMatrix(i, PrevNo2) = NextPrevDates(j).NextPrevInfo(4)
                    .TextMatrix(i, PrevNo2School) = NextPrevDates(j).NextPrevInfo(5)
                    .TextMatrix(i, PrevNo3) = NextPrevDates(j).NextPrevInfo(6)
                    .TextMatrix(i, PrevNo3School) = NextPrevDates(j).NextPrevInfo(7)
                    .TextMatrix(i, PrevNo4) = NextPrevDates(j).NextPrevInfo(8)
                    .TextMatrix(i, PrevNo4School) = NextPrevDates(j).NextPrevInfo(9)
                    .TextMatrix(i, PrevAsst) = NextPrevDates(j).NextPrevInfo(10)
                    .TextMatrix(i, PrevAsstSchool) = NextPrevDates(j).NextPrevInfo(11)
                    .TextMatrix(i, NextPrayer) = NextPrevDates(j).NextPrevInfo(12)
                    .TextMatrix(i, NextSQ) = NextPrevDates(j).NextPrevInfo(13)
                    .TextMatrix(i, NextNo1) = NextPrevDates(j).NextPrevInfo(14)
                    .TextMatrix(i, NextBH) = NextPrevDates(j).NextPrevInfo(15)
                    .TextMatrix(i, NextNo2) = NextPrevDates(j).NextPrevInfo(16)
                    .TextMatrix(i, NextNo2School) = NextPrevDates(j).NextPrevInfo(17)
                    .TextMatrix(i, NextNo3) = NextPrevDates(j).NextPrevInfo(18)
                    .TextMatrix(i, NextNo3School) = NextPrevDates(j).NextPrevInfo(19)
                    .TextMatrix(i, NextNo4) = NextPrevDates(j).NextPrevInfo(20)
                    .TextMatrix(i, NextNo4School) = NextPrevDates(j).NextPrevInfo(21)
                    .TextMatrix(i, NextAsst) = NextPrevDates(j).NextPrevInfo(22)
                    .TextMatrix(i, NextAsstSchool) = NextPrevDates(j).NextPrevInfo(23)
                    End With
                                        
                End If
                 
                 '
                 'Now highlight in green the most recent
                 ' next & previous talks given by each student. This aids user in
                 ' spreading talks more evenly.
                 '
                 With TheGrid
                 '
                 'Put pertinent PrevDates into an array
                 ' (Must change format since can only use variants in Array function)
                 '
                 PrevDates = Array(Format(NextPrevDates(j).NextPrevInfo(1), "yyyy/mm/dd"), _
                                Format(NextPrevDates(j).NextPrevInfo(2), "yyyy/mm/dd"), _
                                Format(NextPrevDates(j).NextPrevInfo(3), "yyyy/mm/dd"), _
                                Format(NextPrevDates(j).NextPrevInfo(4), "yyyy/mm/dd"), _
                                Format(NextPrevDates(j).NextPrevInfo(6), "yyyy/mm/dd"), _
                                Format(NextPrevDates(j).NextPrevInfo(8), "yyyy/mm/dd"))
                 
                 '
                 'Sort desc
                 '
                 BubbleSort PrevDates, , True
                 
                 .Row = i
                 
                If PrevDates(0) <> "" Then
                    Select Case Format(PrevDates(0), "dd/mm/yy")
                    Case NextPrevDates(j).NextPrevInfo(1)
                       .col = PrevSQ
                       .CellBackColor = PrevTalkColour
                    Case NextPrevDates(j).NextPrevInfo(2)
                       .col = PrevNo1
                       .CellBackColor = PrevTalkColour
                    Case NextPrevDates(j).NextPrevInfo(3)
                       .col = PrevBH
                       .CellBackColor = PrevTalkColour
                    Case NextPrevDates(j).NextPrevInfo(4)
                       .col = PrevNo2
                       .CellBackColor = PrevTalkColour
                       .col = PrevNo2School
                       .CellBackColor = PrevTalkColour
                    Case NextPrevDates(j).NextPrevInfo(6)
                       .col = PrevNo3
                       .CellBackColor = PrevTalkColour
                       .col = PrevNo3School
                       .CellBackColor = PrevTalkColour
                    Case NextPrevDates(j).NextPrevInfo(8)
                       .col = PrevNo4
                       .CellBackColor = PrevTalkColour
                       .col = PrevNo4School
                       .CellBackColor = PrevTalkColour
                    End Select
                End If
                
                 '
                 'Put pertinent NextDates into an array
                 ' (Must change format since can only use variants in Array function)
                 '
                 NextDates = Array(Format(NextPrevDates(j).NextPrevInfo(13), "yyyy/mm/dd"), _
                                Format(NextPrevDates(j).NextPrevInfo(14), "yyyy/mm/dd"), _
                                Format(NextPrevDates(j).NextPrevInfo(15), "yyyy/mm/dd"), _
                                Format(NextPrevDates(j).NextPrevInfo(16), "yyyy/mm/dd"), _
                                Format(NextPrevDates(j).NextPrevInfo(18), "yyyy/mm/dd"), _
                                Format(NextPrevDates(j).NextPrevInfo(20), "yyyy/mm/dd"))
                 
                 '
                 'Sort asc
                 '
                 BubbleSort NextDates, , False
                 
                 '
                 'Find the minimum non-blank date
                 '
                For n = 0 To 5
                    If NextDates(n) <> "" Then
                        Exit For
                    End If
                Next n
                 
                If n > 5 Then
                    n = 5
                End If
                 
                 .Row = i
                 
                If NextDates(n) <> "" Then
                    Select Case Format(NextDates(n), "dd/mm/yy")
                    Case NextPrevDates(j).NextPrevInfo(13)
                       .col = NextSQ
                       .CellBackColor = NextTalkColour
                    Case NextPrevDates(j).NextPrevInfo(14)
                       .col = NextNo1
                       .CellBackColor = NextTalkColour
                    Case NextPrevDates(j).NextPrevInfo(15)
                       .col = NextBH
                       .CellBackColor = NextTalkColour
                    Case NextPrevDates(j).NextPrevInfo(16)
                       .col = NextNo2
                       .CellBackColor = NextTalkColour
                       .col = NextNo2School
                       .CellBackColor = NextTalkColour
                    Case NextPrevDates(j).NextPrevInfo(18)
                       .col = NextNo3
                       .CellBackColor = NextTalkColour
                       .col = NextNo3School
                       .CellBackColor = NextTalkColour
                    Case NextPrevDates(j).NextPrevInfo(20)
                       .col = NextNo4
                       .CellBackColor = NextTalkColour
                       .col = NextNo4School
                       .CellBackColor = NextTalkColour
                    End Select
                End If
                End With
                
            End If
                                        
            i = i + 1
            
            If TheGridName = "flxInsertStudent" Then
                j = j + 1
            End If
            
            .MoveNext
        Loop
    End If
    
    End With
    
    Screen.MousePointer = vbNormal
    
    
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Public Sub FillTMSStudentGrid_2009(TheGrid As MSFlexGrid)
'
'This grid allows replacement student to be picked.
'
'This proc is executed for both flxInsertStudent and flxInsertAssistant, depending on
' the parm. When executed for flsInsertStudent, Prev/Next dates are acquired from the DB]
' via the TMSPrevDateFAST and TMSNextDateFAST methods. These are stored in array NextPrevDates_2009().
' This array has elements of Type TMSPersonAndNextPrevDates_2009, one for each Student to
' be dislayed on the grid. Each student has a set of 23 elements in the UDT, each element storing
' a Next/Prev date/school. This stored data is then used to populate flxInsertAssistant.
'
Dim CurrentAssignmentDate As Date, CurrentTalkNo As String
Dim rstStudentList As Recordset, SQLStr As String, BroOnlyFlag As Boolean, SchoolNum As Long
Dim SchoolNumForSQL As Long, CurrentTalkNoForSQL As Long, i As Long, PersonStore As String
Dim ActualTalkDate As Date, OrderingSQL As String, PrevDates() As Variant, n As Integer
Dim a As String, BroOnlySQL As String, SisOnlySQL As String, SuspendDateSQL As String, NextDates() As Variant
Dim AverageWeighting As Long, rstAvWtg As Recordset, NoPersons As Long, j As Long
Dim TheGridName As String, TaskSQL  As String
Dim lSuspTask As Long, TopSQL As String, lTOPVal As Long, sName As String
Dim bIsSourceless As Boolean, SourcelessSQL As String, CompleteSQL As String
Dim SchoolTaskSQL As String, sPrevAsstDate As String, sNextAsstDate As String
Dim sPrevTalkDate As String, sNextTalkDate As String, lCurrentPersonInGrid As Long
Dim ExclApptMenSQL As String, BroOnlyAsstSQL As String, SisOnlyAsstSQL As String
Dim lPrevStudent As Long

On Error GoTo ErrorTrap

    Screen.MousePointer = vbHourglass
    
    TheGridName = TheGrid.Name

    With frmTMSScheduling
    '
    'Need to use the actual TaskNo for the school-no a student is able to have talks in
    ' Required for the SQL later...
    '
    Select Case True
    Case !opt1stSchool:
        SchoolNum = 1
        SchoolNumForSQL = 44
    Case !opt2ndSchool:
        SchoolNum = 2
        SchoolNumForSQL = 45
    Case !opt3rdSchool:
        SchoolNum = 3
        SchoolNumForSQL = 46
    End Select
    
    Select Case !lblBroOnly.Caption
    Case "Yes":
        BroOnlyFlag = True
    Case "":
        BroOnlyFlag = False
    End Select
    
    
    CurrentTalkNo = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 1)
    CurrentAssignmentDate = CDate(!flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 0))
    
    If frmTMSInsertStudent.SourcelessFilter Then
        If (TheGridName = "flxInsertStudent") And (CurrentTalkNo = "2" Or CurrentTalkNo = "3") Then
            bIsSourceless = TheTMS.IsItSourceless(CurrentAssignmentDate, CurrentTalkNo)
        Else
            bIsSourceless = False
        End If
    Else
        bIsSourceless = False
    End If
    
    If TheGridName = "flxInsertStudent" Then 'only need to do this bit once
        frmTMSInsertStudent.FillAssignmentListBox CurrentAssignmentDate
    End If
    
    If TheGridName = "flxInsertStudent" Then
    
        With frmTMSInsertStudent
        
        .FormAssignmentDate = CurrentAssignmentDate
        .FormSchoolNo = SchoolNum
        .FormTalkNo = CurrentTalkNo
        
        lPrevStudent = .FormStudent
        .FormStudent = TheTMS.GetTMSStudent(CurrentAssignmentDate, CurrentTalkNo, SchoolNum, 0)
        If lPrevStudent <> .FormStudent Then
            'we have a new student selected. Determine whether we need to show only sisters on the
            ' Assistants tab...
            Select Case CurrentTalkNo
            Case "2", "3"
                If CongregationMember.GetGender(.FormStudent) = Female Then
                    frmTMSInsertStudent.chkSistersAsst.value = vbChecked
                Else
                    frmTMSInsertStudent.chkSistersAsst.value = vbUnchecked
                End If
            Case Else
                frmTMSInsertStudent.chkSistersAsst.value = vbUnchecked
            End Select
        End If
        
        .FormAssistant = TheTMS.GetTMSAssistant(CurrentAssignmentDate, CurrentTalkNo, SchoolNum, 0)
        End With
        
        If CurrentTalkNo <> "R" And CurrentTalkNo <> "MR" Then
            frmTMSInsertStudent!lblCurrentStudent = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 2)
        Else
            If frmTMSInsertStudent.FormStudent > 0 Then
                frmTMSInsertStudent!lblCurrentStudent = CongregationMember.FirstAndLastName(frmTMSInsertStudent.FormStudent)
            Else
                frmTMSInsertStudent!lblCurrentStudent = ""
            End If
        End If
        
        frmTMSInsertStudent!lblCurrentAssistant = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 3)
        frmTMSInsertStudent!lblCurrentStudent2 = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 2)
        frmTMSInsertStudent!lblCurrentAssistant2 = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 3)
        frmTMSInsertStudent!lblAssignmentDate = Format$(CurrentAssignmentDate, "dd/mm/yyyy")
        frmTMSInsertStudent!lblAssignmentDate2 = Format$(CurrentAssignmentDate, "dd/mm/yyyy")
        frmTMSInsertStudent!lblSchool1 = SchoolNum
        frmTMSInsertStudent!lblSchool2 = SchoolNum
        frmTMSInsertStudent!lblThemeDisplay = !lblThemeDisplay
        frmTMSInsertStudent!lblThemeDisplay2 = !lblThemeDisplay
        frmTMSInsertStudent!lblSetting = !lblSetting
        frmTMSInsertStudent!lblSetting2 = !lblSetting
        frmTMSInsertStudent!lblSourceDisplay = !lblSourceDisplay
        frmTMSInsertStudent!lblSourceDisplay2 = !lblSourceDisplay
        frmTMSInsertStudent!lblBroOnly = !lblBroOnly
        frmTMSInsertStudent!lblSQ = .lblCounsel
        frmTMSInsertStudent!lblSQ2 = .lblCounsel
        frmTMSInsertStudent!lblSQ2 = .lblCounsel
        If AwkwardCounselPoint(frmTMSInsertStudent!lblSQ) Then
            frmTMSInsertStudent!lblSQ.ForeColor = vbRed
            frmTMSInsertStudent!lblSQ2.ForeColor = vbRed
        Else
            frmTMSInsertStudent!lblSQ.ForeColor = vbBlack
            frmTMSInsertStudent!lblSQ2.ForeColor = vbBlack
        End If
        frmTMSInsertStudent!lblSlipPrinted = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 5)
        frmTMSInsertStudent!lblSlipPrinted2 = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 5)
        If frmTMSInsertStudent!lblSlipPrinted = "Y" Then
            frmTMSInsertStudent!lblSlipPrinted.ForeColor = vbRed
            frmTMSInsertStudent!lblSlipPrinted2.ForeColor = vbRed
            frmTMSInsertStudent!lblSlipPrinted.FontBold = True
            frmTMSInsertStudent!lblSlipPrinted2.FontBold = True
        Else
            frmTMSInsertStudent!lblSlipPrinted.ForeColor = vbBlack
            frmTMSInsertStudent!lblSlipPrinted2.ForeColor = vbBlack
            frmTMSInsertStudent!lblSlipPrinted.FontBold = False
            frmTMSInsertStudent!lblSlipPrinted2.FontBold = False
        End If
        
        With frmTMSInsertStudent
        '
        'Display TalkNo description on form
        '
        Select Case CurrentTalkNo
        Case "P":
            !lblAssignment = "Opening Prayer"
            !lblAssignment2 = "Opening Prayer"
        Case "B":
            !lblAssignment = "Bible Highlights"
            !lblAssignment2 = "Bible Highlights"
        Case "1":
            !lblAssignment = "Reading Assignment"
            !lblAssignment2 = "Reading Assignment"
        Case "2":
            !lblAssignment = "Talk No 2"
            !lblAssignment2 = "Talk No 2"
        Case "3":
            !lblAssignment = "Talk No 3"
            !lblAssignment2 = "Talk No 3"
        Case "R", "MR":
            !lblAssignment = "Oral Review Reading"
            !lblAssignment2 = ""
        Case Else:
            !lblAssignment = ""
            !lblAssignment2 = ""
        
        End Select
        End With
        
    End If
    
    End With
       
    
    ActualTalkDate = Format(GetDateOfGivenDay(CurrentAssignmentDate, _
                        GlobalParms.GetValue("MidWeekMeetingDay", "NumVal")), "mm/dd/yyyy")
    
    
    Select Case frmTMSInsertStudent.chkBrothersOnly.value
    Case vbChecked
        BroOnlySQL = " AND c.GenderMF = 'M' "
    Case vbUnchecked
        BroOnlySQL = ""
    End Select
    
    Select Case frmTMSInsertStudent.chkSistersOnly.value
    Case vbChecked
        SisOnlySQL = " AND c.GenderMF = 'F' "
    Case vbUnchecked
        SisOnlySQL = ""
    End Select
    
    'asst check box filters
    Select Case frmTMSInsertStudent.chkexclEMSAsst.value
    Case vbChecked
        Select Case CurrentTalkNo
        Case "2", "3"
            ExclApptMenSQL = " AND NOT EXISTS (SELECT 1 FROM tblTaskAndPerson t " & _
                      "                 WHERE t.Task IN (51,52) " & _
                                      " AND t.Person = a.Person) "
                                      
        Case Else
            ExclApptMenSQL = " "
        End Select
    Case vbUnchecked
        ExclApptMenSQL = " "
    End Select

    'asst check box filters
    Select Case frmTMSInsertStudent.chkBrosAsst.value
    Case vbChecked
        BroOnlyAsstSQL = " AND c.GenderMF = 'M' "
    Case vbUnchecked
        BroOnlyAsstSQL = ""
    End Select
    
    'asst check box filters
    Select Case frmTMSInsertStudent.chkSistersAsst.value
    Case vbChecked
        SisOnlyAsstSQL = " AND c.GenderMF = 'F' "
    Case vbUnchecked
        SisOnlyAsstSQL = ""
    End Select
    
    
    
    Select Case frmTMSInsertStudent.chkShowEandMSOnly.value
    Case vbChecked
        If CurrentTalkNo = "3" Then
            TaskSQL = " a.Task = " & rstTMSSchedule!TaskNo & " " & _
                      " AND EXISTS (SELECT 1 FROM tblTaskAndPerson t " & _
                      "                 WHERE t.Task IN (51,52) " & _
                                      " AND t.Person = a.Person) "
        Else
            TaskSQL = " a.Task = " & rstTMSSchedule!TaskNo & " "
        End If
    Case vbUnchecked
        Select Case frmTMSInsertStudent.chkExclAppointedMen.value
        Case vbChecked
            Select Case CurrentTalkNo
            Case "1", "3"
                TaskSQL = " a.Task = " & rstTMSSchedule!TaskNo & " " & _
                          " AND NOT EXISTS (SELECT 1 FROM tblTaskAndPerson t " & _
                          "                 WHERE t.Task IN (51,52) " & _
                                          " AND t.Person = a.Person) "
                                          
            Case "R", "MR" 'should never be the case when chkExclAppointedMen is checked
                TaskSQL = " a.Task = 103 " 'oral review reader
            Case Else
                TaskSQL = " a.Task = " & rstTMSSchedule!TaskNo & " "
            End Select
        Case vbUnchecked
            Select Case CurrentTalkNo
            Case "R", "MR"
                TaskSQL = " a.Task = 103 "
            Case Else
                TaskSQL = " a.Task = " & rstTMSSchedule!TaskNo & " "
            End Select
        End Select
    End Select
    
    If bIsSourceless Then
        Select Case CurrentTalkNo
        Case "2", "3"
            SourcelessSQL = " AND EXISTS (SELECT 1 FROM tblTaskAndPerson t2 " & _
                             "                 WHERE t2.Task = 39 " & _
                                      " AND t2.Person = a.Person) "
                                      
        Case Else
            SourcelessSQL = " "
        End Select
    Else
        SourcelessSQL = " "
    End If
    
    'if listing brothers for talk #2 then fill list from bros that do talk #3
    If CurrentTalkNo = "2" Then
        If BroOnlySQL <> "" Then
            TaskSQL = " a.Task = 101 "
        End If
    End If
    
    Select Case CurrentTalkNo
    Case "R", "MR"
        lSuspTask = 103
    Case Else
        If CurrentTalkNo = "2" Then
            If BroOnlySQL = "" Then
                lSuspTask = IIf(TheGridName = "flxInsertStudent", rstTMSSchedule!TaskNo, 43)
            Else
                lSuspTask = 101
            End If
        Else
            lSuspTask = IIf(TheGridName = "flxInsertStudent", rstTMSSchedule!TaskNo, 43)
        End If
    End Select
    
    SuspendDateSQL = " AND NOT EXISTS (SELECT 1 " & _
                    "FROM tblTaskPersonSuspendDates f " & _
                    "WHERE Task = " & lSuspTask & _
                    " AND a.Person = f.Person " & _
                    " AND (SuspendStartDate <= #" & ActualTalkDate & "# AND SuspendEndDate >= #" & ActualTalkDate & "#)) "
                        
    Select Case frmTMSInsertStudent.optIntelligent.value
    Case True 'intelligent
    
        lTOPVal = GlobalParms.GetValue("TMS_NumberOfRowsToShowOnInsertForm", "NumVal", 0)
        If lTOPVal = 0 Then
            TopSQL = ""
         Else
            TopSQL = " TOP " & lTOPVal & " "
        End If
    
        '
        'Prayer weighting is separate to other talk weighting. Otherwise
        ' a brother that does prayers and talks would end up doing less talks
        ' than brothers that don't do prayers....
        '
        If CurrentTalkNo <> "P" Then
            OrderingSQL = "ORDER BY b.TMSWeighting"
        Else
            OrderingSQL = "ORDER BY b.TMSPrayerWeighting"
        End If
        
        
        If TheGridName = "flxInsertStudent" Then
            '
            'ie The InsertStudent grid is currently being populated
            '
            'Take account of suspend dates
            '
            Select Case CurrentTalkNo
            Case "R", "MR"
                'not taking account of school for oral review reader
                CompleteSQL = "SELECT DISTINCT " & TopSQL & " a.Person, " & _
                                "b.TMSWeighting, " & _
                                "b.TMSPrayerWeighting " & _
                                "FROM ((tblTaskPersonSuspendDates a " & _
                                "INNER JOIN tblTMSWeightings b ON " & _
                                "a.Person = b.PersonID) " & _
                                "INNER JOIN tblNameAddress c ON " & _
                                " (c.ID = b.PersonID)) " & _
                                "WHERE " & TaskSQL & _
                                SuspendDateSQL & " AND Active = TRUE " & _
                                BroOnlySQL & SisOnlySQL & SourcelessSQL & OrderingSQL
            Case Else
                CompleteSQL = "SELECT DISTINCT " & TopSQL & " a.Person, " & _
                                "b.TMSWeighting, " & _
                                "b.TMSPrayerWeighting " & _
                                "FROM (((tblTaskPersonSuspendDates a " & _
                                "INNER JOIN tblTMSWeightings b ON " & _
                                "a.Person = b.PersonID) " & _
                                "INNER JOIN tblNameAddress c ON " & _
                                " (c.ID = b.PersonID)) " & _
                                "INNER JOIN tblTMSSchoolAndPerson d ON " & _
                                " d.PersonID = a.Person AND d.Task = a.Task) " & _
                                "WHERE " & TaskSQL & _
                                SuspendDateSQL & " AND Active = TRUE " & _
                                " AND d.SchoolNo = " & SchoolNum & " " & _
                                BroOnlySQL & SisOnlySQL & SourcelessSQL & OrderingSQL
            End Select
            
                            
            Set rstStudentList = CMSDB.OpenRecordset(CompleteSQL, dbOpenForwardOnly)
                                                                                                
        Else
'            '
'            'ie The InsertAssistant grid is currently being populated
'            '


            '
            'Take account of suspend dates
            '
            CompleteSQL = "SELECT DISTINCT " & TopSQL & " a.Person, " & _
                            "b.TMSAsstWeighting " & _
                            "FROM (((tblTaskPersonSuspendDates a " & _
                            "INNER JOIN tblTMSWeightings b ON " & _
                            "a.Person = b.PersonID) " & _
                            "INNER JOIN tblNameAddress c ON " & _
                            " (c.ID = b.PersonID)) " & _
                            "INNER JOIN tblTMSSchoolAndPerson d ON " & _
                            " d.PersonID = a.Person AND d.Task = a.Task) " & _
                            "WHERE a.Task IN (42,43)" & _
                            SuspendDateSQL & " AND Active = TRUE " & _
                            " AND d.SchoolNo = " & SchoolNum & " " & _
                            ExclApptMenSQL & BroOnlyAsstSQL & SisOnlyAsstSQL & _
                            "ORDER BY b.TMSAsstWeighting"
            
            Set rstStudentList = CMSDB.OpenRecordset(CompleteSQL, dbOpenForwardOnly)
                                                                
        End If
        
    Case False 'non-intelligent
    
        If TheGridName = "flxInsertStudent" Then
            '
            'ie The InsertStudent grid is currently being populated
            '
            'Take account of suspend dates
            '
            Set rstStudentList = CMSDB.OpenRecordset("SELECT DISTINCT a.Person, c.FirstName, " & _
                                                    "c.MiddleName, c.LastName " & _
                                                "FROM ((tblTaskPersonSuspendDates a " & _
                                                "INNER JOIN tblNameAddress c ON " & _
                                                " (a.Person = c.ID)) " & _
                                                "INNER JOIN tblTMSSchoolAndPerson d ON " & _
                                                " d.PersonID = a.Person AND d.Task = a.Task) " & _
                                                "WHERE " & TaskSQL & _
                                                SuspendDateSQL & " AND Active = TRUE " & _
                                                " AND d.SchoolNo = " & SchoolNum & " " & _
                                                BroOnlySQL & SisOnlySQL & SourcelessSQL & _
                                                " ORDER BY LastName, FirstName, MiddleName", dbOpenForwardOnly)
                                                                                                
        Else
            '
            'ie The InsertAssistant grid is currently being populated
            '
            'Take account of suspend dates
            '
            Set rstStudentList = CMSDB.OpenRecordset("SELECT DISTINCT a.Person, c.FirstName, " & _
                                                    "c.MiddleName, c.LastName " & _
                                                "FROM ((tblTaskPersonSuspendDates a " & _
                                                "INNER JOIN tblNameAddress c ON " & _
                                                " (a.Person = c.ID)) " & _
                                                "INNER JOIN tblTMSSchoolAndPerson d ON " & _
                                                " d.PersonID = a.Person AND d.Task = a.Task) " & _
                                                "WHERE a.Task IN (42,43) " & _
                                                SuspendDateSQL & " AND Active = TRUE " & _
                                                " AND d.SchoolNo = " & SchoolNum & " " & _
                                                " ORDER BY LastName, FirstName, MiddleName", dbOpenForwardOnly)
                                                                
        End If
    End Select
    
    
    i = 2 'row of grid
    j = 0 'position on array storing each student's prev/next dates
    
    '
    'Fill grid with last/prev dates. Actual columns used is determined by the PrevXxxx
    ' variables which are set in frmTMSInsertStudent.SetUpGrids according to the
    ' talkno selected on frmTMSScheduling.
    '
    
    TheGrid.Rows = 2
    With rstStudentList
    If Not .BOF Then
        Do While Not .EOF
            TheGrid.Rows = i + 1
            
            If CongregationMember.TMSStudentOnThisWeekFAST(!Person, CurrentAssignmentDate) Then
                frmTMSInsertStudent.GridRowBackColourToGrey TheGrid, i, True, False
            Else
                frmTMSInsertStudent.GridRowBackColourToGrey TheGrid, i, False, False
            End If
                        
            TheGrid.TextMatrix(i, 0) = CongregationMember.LastFirstNameMiddleName(!Person)
            
            lCurrentPersonInGrid = !Person
            TheGrid.TextMatrix(i, TheGrid.Cols - 1) = lCurrentPersonInGrid 'NameID used for programmatic purposes.
            
            If frmTMSInsertStudent.chkShowNextPrevDates.value = vbChecked Then
                If TheGridName = "flxInsertStudent" Then
                
                    ReDim Preserve NextPrevDates_2009(j)
                    NextPrevDates_2009(j).ThePersonID = !Person
                    
                    '
                    'Put next/prev dates into an array
                    '
                    NextPrevDates_2009(j).NextPrevInfo(0) = CongregationMember.TMSPrevDateFAST(!Person, "B", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(1) = CongregationMember.TMSPrevDateFAST(!Person, "1", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(2) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2009(j).NextPrevInfo(3) = CongregationMember.TMSPrevDateFAST(!Person, "2", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(4) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2009(j).NextPrevInfo(5) = CongregationMember.TMSPrevDateFAST(!Person, "3", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(6) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2009(j).NextPrevInfo(7) = CongregationMember.TMSPrevDateFAST(!Person, "Asst", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(8) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2009(j).NextPrevInfo(9) = CongregationMember.TMSNextDateFAST(!Person, "B", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(10) = CongregationMember.TMSNextDateFAST(!Person, "1", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(11) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2009(j).NextPrevInfo(12) = CongregationMember.TMSNextDateFAST(!Person, "2", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(13) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2009(j).NextPrevInfo(14) = CongregationMember.TMSNextDateFAST(!Person, "3", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(15) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2009(j).NextPrevInfo(16) = CongregationMember.TMSNextDateFAST(!Person, "Asst", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(17) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2009(j).NextPrevInfo(18) = CongregationMember.TMSPrevDateFAST(!Person, "R", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(19) = CongregationMember.TMSNextDateFAST(!Person, "R", CurrentAssignmentDate)
                    
                Else
                    '
                    'This is Assistant-grid. Get prev/next dates from array.
                    '
                    'Locate current student in array to get their prev/next dates
                    '
                    For j = 0 To UBound(NextPrevDates_2009)
                        If NextPrevDates_2009(j).ThePersonID = !Person Then
                            Exit For
                        End If
                    Next j
                    
                End If
                 
                '
                'Put array contents into student grid. If Student is not in array
                ' (ie because they're only a householder, so wouldn't have been included
                ' when processing Student Grid), then get dates from DB
                '
                If j <= UBound(NextPrevDates_2009) Then
                    With TheGrid
                    .TextMatrix(i, PrevBH) = NextPrevDates_2009(j).NextPrevInfo(0)
                    .TextMatrix(i, PrevNo1) = NextPrevDates_2009(j).NextPrevInfo(1)
                    .TextMatrix(i, PrevNo1School) = NextPrevDates_2009(j).NextPrevInfo(2)
                    .TextMatrix(i, PrevNo3) = NextPrevDates_2009(j).NextPrevInfo(5)
                    .TextMatrix(i, NextNo3) = NextPrevDates_2009(j).NextPrevInfo(14)
                    .TextMatrix(i, NextNo3School) = NextPrevDates_2009(j).NextPrevInfo(15)
                    .TextMatrix(i, PrevNo3School) = NextPrevDates_2009(j).NextPrevInfo(6)
                    .TextMatrix(i, NextBH) = NextPrevDates_2009(j).NextPrevInfo(9)
                    .TextMatrix(i, NextNo1) = NextPrevDates_2009(j).NextPrevInfo(10)
                    .TextMatrix(i, NextNo1School) = NextPrevDates_2009(j).NextPrevInfo(11)
                    
                    If CurrentTalkNo <> "R" And CurrentTalkNo <> "MR" Then
                        .TextMatrix(i, PrevNo2) = NextPrevDates_2009(j).NextPrevInfo(3)
                        .TextMatrix(i, PrevNo2School) = NextPrevDates_2009(j).NextPrevInfo(4)
                        .TextMatrix(i, PrevAsst) = NextPrevDates_2009(j).NextPrevInfo(7)
                        .TextMatrix(i, PrevAsstSchool) = NextPrevDates_2009(j).NextPrevInfo(8)
                        .TextMatrix(i, NextNo2) = NextPrevDates_2009(j).NextPrevInfo(12)
                        .TextMatrix(i, NextNo2School) = NextPrevDates_2009(j).NextPrevInfo(13)
                        .TextMatrix(i, NextAsst) = NextPrevDates_2009(j).NextPrevInfo(16)
                        .TextMatrix(i, NextAsstSchool) = NextPrevDates_2009(j).NextPrevInfo(17)
                    Else
                        .TextMatrix(i, PrevReview) = NextPrevDates_2009(j).NextPrevInfo(18)
                        .TextMatrix(i, NextReview) = NextPrevDates_2009(j).NextPrevInfo(19)
                    End If
                    
                    End With
                Else
                    '
                    'Go to   s l o w   DB. Add new student to array, and get the dates....
                    '
                    ReDim Preserve NextPrevDates_2009(j)
                    NextPrevDates_2009(j).ThePersonID = !Person
                    
                    NextPrevDates_2009(j).NextPrevInfo(0) = CongregationMember.TMSPrevDateFAST(!Person, "B", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(1) = CongregationMember.TMSPrevDateFAST(!Person, "1", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(2) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2009(j).NextPrevInfo(5) = CongregationMember.TMSPrevDateFAST(!Person, "3", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(6) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2009(j).NextPrevInfo(9) = CongregationMember.TMSNextDateFAST(!Person, "B", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(10) = CongregationMember.TMSNextDateFAST(!Person, "1", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(11) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2009(j).NextPrevInfo(14) = CongregationMember.TMSNextDateFAST(!Person, "3", CurrentAssignmentDate)
                    NextPrevDates_2009(j).NextPrevInfo(15) = CongregationMember.GetTMSSchoolNoForInsertForm
                                        
                    If CurrentTalkNo <> "R" And CurrentTalkNo <> "MR" Then
                        NextPrevDates_2009(j).NextPrevInfo(3) = CongregationMember.TMSPrevDateFAST(!Person, "2", CurrentAssignmentDate)
                        NextPrevDates_2009(j).NextPrevInfo(4) = CongregationMember.GetTMSSchoolNoForInsertForm
                        NextPrevDates_2009(j).NextPrevInfo(7) = CongregationMember.TMSPrevDateFAST(!Person, "Asst", CurrentAssignmentDate)
                        NextPrevDates_2009(j).NextPrevInfo(8) = CongregationMember.GetTMSSchoolNoForInsertForm
                        NextPrevDates_2009(j).NextPrevInfo(12) = CongregationMember.TMSNextDateFAST(!Person, "2", CurrentAssignmentDate)
                        NextPrevDates_2009(j).NextPrevInfo(13) = CongregationMember.GetTMSSchoolNoForInsertForm
                        NextPrevDates_2009(j).NextPrevInfo(16) = CongregationMember.TMSNextDateFAST(!Person, "Asst", CurrentAssignmentDate)
                        NextPrevDates_2009(j).NextPrevInfo(17) = CongregationMember.GetTMSSchoolNoForInsertForm
                    Else
                        NextPrevDates_2009(j).NextPrevInfo(18) = CongregationMember.TMSPrevDateFAST(!Person, "R", CurrentAssignmentDate)
                        NextPrevDates_2009(j).NextPrevInfo(19) = CongregationMember.TMSNextDateFAST(!Person, "R", CurrentAssignmentDate)
                    End If
                                        
                                        
                    With TheGrid
                    .TextMatrix(i, PrevBH) = NextPrevDates_2009(j).NextPrevInfo(0)
                    .TextMatrix(i, PrevNo1) = NextPrevDates_2009(j).NextPrevInfo(1)
                    .TextMatrix(i, PrevNo1School) = NextPrevDates_2009(j).NextPrevInfo(2)
                    .TextMatrix(i, PrevNo3) = NextPrevDates_2009(j).NextPrevInfo(5)
                    .TextMatrix(i, PrevNo3School) = NextPrevDates_2009(j).NextPrevInfo(6)
                    .TextMatrix(i, NextBH) = NextPrevDates_2009(j).NextPrevInfo(9)
                    .TextMatrix(i, NextNo1) = NextPrevDates_2009(j).NextPrevInfo(10)
                    .TextMatrix(i, NextNo1School) = NextPrevDates_2009(j).NextPrevInfo(11)
                    .TextMatrix(i, NextNo3) = NextPrevDates_2009(j).NextPrevInfo(14)
                    .TextMatrix(i, NextNo3School) = NextPrevDates_2009(j).NextPrevInfo(15)
                    
                    If CurrentTalkNo <> "R" And CurrentTalkNo <> "MR" Then
                        .TextMatrix(i, PrevNo2) = NextPrevDates_2009(j).NextPrevInfo(3)
                        .TextMatrix(i, PrevNo2School) = NextPrevDates_2009(j).NextPrevInfo(4)
                        .TextMatrix(i, PrevAsst) = NextPrevDates_2009(j).NextPrevInfo(7)
                        .TextMatrix(i, PrevAsstSchool) = NextPrevDates_2009(j).NextPrevInfo(8)
                        .TextMatrix(i, NextNo2) = NextPrevDates_2009(j).NextPrevInfo(12)
                        .TextMatrix(i, NextNo2School) = NextPrevDates_2009(j).NextPrevInfo(13)
                        .TextMatrix(i, NextAsst) = NextPrevDates_2009(j).NextPrevInfo(16)
                        .TextMatrix(i, NextAsstSchool) = NextPrevDates_2009(j).NextPrevInfo(17)
                    Else
                        .TextMatrix(i, PrevReview) = NextPrevDates_2009(j).NextPrevInfo(18)
                        .TextMatrix(i, NextReview) = NextPrevDates_2009(j).NextPrevInfo(19)
                    End If
                    
                    
                    End With
                                        
                End If
                 
                 '
                 'Now highlight in green the most recent
                 ' next & previous talks given by each student. This aids user in
                 ' spreading talks more evenly.
                 '
                 With TheGrid
                 '
                 'Put pertinent PrevDates into an array
                 ' (Must change format since can only use variants in Array function)
                 '
                 
                 'col | talk type
                 ' 0  |   Prev BH
                 ' 1  |   Prev #1
                 ' 3  |   Prev #2
                 ' 5  |   Prev #3
                 ' 7  |   Prev Asst
                 
                 
                 
                PrevDates = Array(Format(NextPrevDates_2009(j).NextPrevInfo(0), "yyyy/mm/dd"), _
                               Format(NextPrevDates_2009(j).NextPrevInfo(1), "yyyy/mm/dd"), _
                               Format(NextPrevDates_2009(j).NextPrevInfo(3), "yyyy/mm/dd"), _
                               Format(NextPrevDates_2009(j).NextPrevInfo(5), "yyyy/mm/dd"))
                               
                sPrevAsstDate = Format(NextPrevDates_2009(j).NextPrevInfo(7), "yyyy/mm/dd")
                                 
                 
                 '
                 'Sort desc
                 '
                 BubbleSort PrevDates, , True
                 
                 .Row = i
                 
                If PrevDates(0) <> "" Then
                    Select Case Format(PrevDates(0), "dd/mm/yy")
                    Case NextPrevDates_2009(j).NextPrevInfo(0)
                       .col = PrevBH
                       .CellBackColor = PrevTalkColour
                    Case NextPrevDates_2009(j).NextPrevInfo(1)
                       .col = PrevNo1
                       .CellBackColor = PrevTalkColour
                       .col = PrevNo1School
                       .CellBackColor = PrevTalkColour
                    Case NextPrevDates_2009(j).NextPrevInfo(3)
                       .col = PrevNo2
                       .CellBackColor = PrevTalkColour
                       .col = PrevNo2School
                       .CellBackColor = PrevTalkColour
                    Case NextPrevDates_2009(j).NextPrevInfo(5)
                       .col = PrevNo3
                       .CellBackColor = PrevTalkColour
                       .col = PrevNo3School
                       .CellBackColor = PrevTalkColour
                    End Select
                End If
                
                sPrevTalkDate = PrevDates(0)
                
                 '
                 'Put pertinent NextDates into an array
                 ' (Must change format since can only use variants in Array function)
                 '
                 
                 'col | talk type
                 ' 9  |   Next BH
                 ' 10 |   Next #1
                 ' 12 |   Next #2
                 ' 14 |   Next #3
                 ' 16 |   Next Asst

                NextDates = Array(Format(NextPrevDates_2009(j).NextPrevInfo(9), "yyyy/mm/dd"), _
                               Format(NextPrevDates_2009(j).NextPrevInfo(10), "yyyy/mm/dd"), _
                               Format(NextPrevDates_2009(j).NextPrevInfo(12), "yyyy/mm/dd"), _
                               Format(NextPrevDates_2009(j).NextPrevInfo(14), "yyyy/mm/dd"))
                               
                sNextAsstDate = Format(NextPrevDates_2009(j).NextPrevInfo(16), "yyyy/mm/dd")
                 
                 '
                 'Sort asc
                 '
                 BubbleSort NextDates, , False
                 
                 '
                 'Find the minimum non-blank date
                 '
                For n = 0 To UBound(NextDates)
                    If NextDates(n) <> "" Then
                        Exit For
                    End If
                Next n
                 
                If n > UBound(NextDates) Then
                    n = UBound(NextDates)
                End If
                 
                 .Row = i
                 
                If NextDates(n) <> "" Then
                    Select Case Format(NextDates(n), "dd/mm/yy")
                    Case NextPrevDates_2009(j).NextPrevInfo(9)
                       .col = NextBH
                       .CellBackColor = NextTalkColour
                    Case NextPrevDates_2009(j).NextPrevInfo(10)
                       .col = NextNo1
                       .CellBackColor = NextTalkColour
                       .col = NextNo1School
                       .CellBackColor = NextTalkColour
                    Case NextPrevDates_2009(j).NextPrevInfo(12)
                       .col = NextNo2
                       .CellBackColor = NextTalkColour
                       .col = NextNo2School
                       .CellBackColor = NextTalkColour
                    Case NextPrevDates_2009(j).NextPrevInfo(14)
                       .col = NextNo3
                       .CellBackColor = NextTalkColour
                       .col = NextNo3School
                       .CellBackColor = NextTalkColour
                    End Select
                End If
                
                sNextTalkDate = NextDates(n)
                
                End With
                
            End If
            
            
            'so we now have the most recent and next talk dates,
            ' and the most recent and next asst dates.
            'If the asst dates are closer than the talk dates, highlight the asst dates too.
            
            
             With TheGrid
             
             If sPrevAsstDate <> "" Then
                If (sPrevAsstDate > sPrevTalkDate) Or (sPrevTalkDate = "") Then
                    .col = PrevAsst
                    .CellBackColor = PrevAsstColour
                    .col = PrevAsstSchool
                    .CellBackColor = PrevAsstColour
                End If
             End If
             
             If sNextAsstDate <> "" Then
                If (sNextAsstDate < sNextTalkDate) Or (sNextTalkDate = "") Then
                    .col = NextAsst
                    .CellBackColor = NextAsstColour
                    .col = NextAsstSchool
                    .CellBackColor = NextAsstColour
                End If
             End If
             
            
             End With
            
                                        
            i = i + 1
            
            If TheGridName = "flxInsertStudent" Then
                j = j + 1
            End If
            
            .MoveNext
        Loop
    End If
    
    End With
    
    Screen.MousePointer = vbNormal
    
    
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Public Function FillTMSStudentGrid_2016(TheGrid As MSFlexGrid) As Boolean
'
'This grid allows replacement student to be picked.
'
'This proc is executed for both flxInsertStudent and flxInsertAssistant, depending on
' the parm. When executed for flsInsertStudent, Prev/Next dates are acquired from the DB]
' via the TMSPrevDateFAST and TMSNextDateFAST methods. These are stored in array NextPrevDates_2016().
' This array has elements of Type TMSPersonAndNextPrevDates_2016, one for each Student to
' be dislayed on the grid. Each student has a set of 23 elements in the UDT, each element storing
' a Next/Prev date/school. This stored data is then used to populate flxInsertAssistant.
'
Dim CurrentAssignmentDate As Date, CurrentTalkNo As String, currentItemsSeqNum As Long
Dim rstStudentList As Recordset, SQLStr As String, BroOnlyFlag As Boolean, SchoolNum As Long
Dim SchoolNumForSQL As Long, CurrentTalkNoForSQL As Long, i As Long, PersonStore As String
Dim ActualTalkDate As Date, OrderingSQL As String, PrevDates() As Variant, n As Integer
Dim a As String, BroOnlySQL As String, SisOnlySQL As String, SuspendDateSQL As String, NextDates() As Variant
Dim AverageWeighting As Long, rstAvWtg As Recordset, NoPersons As Long, j As Long
Dim TheGridName As String, TaskSQL  As String
Dim lSuspTask As Long, TopSQL As String, lTOPVal As Long, sName As String
Dim CompleteSQL As String
Dim SchoolTaskSQL As String, sPrevAsstDate As String, sNextAsstDate As String
Dim sPrevTalkDate As String, sNextTalkDate As String, lCurrentPersonInGrid As Long
Dim ExclApptMenSQL As String, BroOnlyAsstSQL As String, SisOnlyAsstSQL As String
Dim lPrevStudent As Long

On Error GoTo ErrorTrap

    Screen.MousePointer = vbHourglass
    
    TheGridName = TheGrid.Name

    With frmTMSScheduling
    '
    'Need to use the actual TaskNo for the school-no a student is able to have talks in
    ' Required for the SQL later...
    '
    Select Case True
    Case !opt1stSchool:
        SchoolNum = 1
        SchoolNumForSQL = 44
    Case !opt2ndSchool:
        SchoolNum = 2
        SchoolNumForSQL = 45
    Case !opt3rdSchool:
        SchoolNum = 3
        SchoolNumForSQL = 46
    End Select
    
    Select Case !lblBroOnly.Caption
    Case "Yes":
        BroOnlyFlag = True
    Case "":
        BroOnlyFlag = False
    End Select
    
    
    CurrentTalkNo = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 1)
    CurrentAssignmentDate = CDate(!flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 0))
    currentItemsSeqNum = CLng(!flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 4))
        
    If TheGridName = "flxInsertStudent" Then 'only need to do this bit once
        frmTMSInsertStudent.FillAssignmentListBox CurrentAssignmentDate
    End If
    
    If TheGridName = "flxInsertStudent" Then
    
        With frmTMSInsertStudent
        
        .FormAssignmentDate = CurrentAssignmentDate
        .FormSchoolNo = SchoolNum
        .FormTalkNo = CurrentTalkNo
        
        lPrevStudent = .FormStudent
        .FormStudent = TheTMS.GetTMSStudent(CurrentAssignmentDate, CurrentTalkNo, SchoolNum, currentItemsSeqNum)
        
        .FormAssistant = TheTMS.GetTMSAssistant(CurrentAssignmentDate, CurrentTalkNo, SchoolNum, currentItemsSeqNum)
        End With
        
        frmTMSInsertStudent!lblCurrentStudent = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 2)
        
        If CurrentTalkNo <> "A" Then
            frmTMSInsertStudent!lblCurrentAssistant = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 3)
            frmTMSInsertStudent!lblCurrentStudent2 = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 2)
            frmTMSInsertStudent!lblCurrentAssistant2 = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 3)
            frmTMSInsertStudent!lblAssignmentDate = Format$(CurrentAssignmentDate, "dd/mm/yyyy")
            frmTMSInsertStudent!lblAssignmentDate2 = Format$(CurrentAssignmentDate, "dd/mm/yyyy")
            frmTMSInsertStudent!lblSchool1 = SchoolNum
            frmTMSInsertStudent!lblSchool2 = SchoolNum
            frmTMSInsertStudent!lblThemeDisplay = !lblThemeDisplay
            frmTMSInsertStudent!lblThemeDisplay2 = !lblThemeDisplay
            frmTMSInsertStudent!lblSetting = !lblSetting
            frmTMSInsertStudent!lblSetting2 = !lblSetting
            frmTMSInsertStudent!lblSourceDisplay = !lblSourceDisplay
            frmTMSInsertStudent!lblSourceDisplay2 = !lblSourceDisplay
            frmTMSInsertStudent!lblBroOnly = !lblBroOnly
            frmTMSInsertStudent!lblSQ = .lblCounsel
            frmTMSInsertStudent!lblSQ2 = .lblCounsel
            frmTMSInsertStudent!lblSQ2 = .lblCounsel
            If AwkwardCounselPoint(frmTMSInsertStudent!lblSQ) Then
                frmTMSInsertStudent!lblSQ.ForeColor = vbRed
                frmTMSInsertStudent!lblSQ2.ForeColor = vbRed
            Else
                frmTMSInsertStudent!lblSQ.ForeColor = vbBlack
                frmTMSInsertStudent!lblSQ2.ForeColor = vbBlack
            End If
            frmTMSInsertStudent!lblSlipPrinted = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 5)
            frmTMSInsertStudent!lblSlipPrinted2 = !flxTMSSchedule.TextMatrix(!flxTMSSchedule.Row, 5)
            If frmTMSInsertStudent!lblSlipPrinted = "Y" Then
                frmTMSInsertStudent!lblSlipPrinted.ForeColor = vbRed
                frmTMSInsertStudent!lblSlipPrinted2.ForeColor = vbRed
                frmTMSInsertStudent!lblSlipPrinted.FontBold = True
                frmTMSInsertStudent!lblSlipPrinted2.FontBold = True
            Else
                frmTMSInsertStudent!lblSlipPrinted.ForeColor = vbBlack
                frmTMSInsertStudent!lblSlipPrinted2.ForeColor = vbBlack
                frmTMSInsertStudent!lblSlipPrinted.FontBold = False
                frmTMSInsertStudent!lblSlipPrinted2.FontBold = False
            End If
            
            With frmTMSInsertStudent
            '
            'Display TalkNo description on form
            '
            !lblAssignment = TheTMS.GetTMSTalkDescription(CurrentTalkNo, Format$(CurrentAssignmentDate, "dd/mm/yyyy"))
            !lblAssignment2 = !lblAssignment
            
            End With
        Else
            frmTMSInsertStudent!lblCurrentAssistant = ""
            frmTMSInsertStudent!lblCurrentStudent2 = ""
            frmTMSInsertStudent!lblCurrentAssistant2 = ""
            frmTMSInsertStudent!lblAssignmentDate = ""
            frmTMSInsertStudent!lblAssignmentDate2 = ""
            frmTMSInsertStudent!lblSchool1 = ""
            frmTMSInsertStudent!lblSchool2 = ""
            frmTMSInsertStudent!lblThemeDisplay = ""
            frmTMSInsertStudent!lblThemeDisplay2 = ""
            frmTMSInsertStudent!lblSetting = ""
            frmTMSInsertStudent!lblSetting2 = ""
            frmTMSInsertStudent!lblSourceDisplay = ""
            frmTMSInsertStudent!lblSourceDisplay2 = ""
            frmTMSInsertStudent!lblBroOnly = ""
            frmTMSInsertStudent!lblSQ = ""
            frmTMSInsertStudent!lblSQ2 = ""
            frmTMSInsertStudent!lblSQ2 = ""
            frmTMSInsertStudent!lblSlipPrinted = ""
            frmTMSInsertStudent!lblSlipPrinted2 = ""
            
            With frmTMSInsertStudent
            '
            'Display TalkNo description on form
            '
            !lblAssignment = ""
            !lblAssignment2 = ""
            
            End With
            
            FillTMSStudentGrid_2016 = True
            Exit Function
            
        End If
        
    Else
    
        If CurrentTalkNo = "A" Then
            FillTMSStudentGrid_2016 = True
            Exit Function
        End If
        
    End If
    
    End With
       
    
    ActualTalkDate = Format(GetDateOfGivenDay(CurrentAssignmentDate, _
                        GlobalParms.GetValue("MidWeekMeetingDay", "NumVal")), "mm/dd/yyyy")
    
    
    Select Case frmTMSInsertStudent.chkBrothersOnly.value
    Case vbChecked
        BroOnlySQL = " AND c.GenderMF = 'M' "
    Case vbUnchecked
        BroOnlySQL = ""
    End Select
    
    Select Case frmTMSInsertStudent.chkSistersOnly.value
    Case vbChecked
        SisOnlySQL = " AND c.GenderMF = 'F' "
    Case vbUnchecked
        SisOnlySQL = ""
    End Select
    
    'asst check box filters
    Select Case frmTMSInsertStudent.chkexclEMSAsst.value
    Case vbChecked
        Select Case CurrentTalkNo
        Case "IC", "RV", "BS", "O"
            ExclApptMenSQL = " AND NOT EXISTS (SELECT 1 FROM tblTaskAndPerson t " & _
                      "                 WHERE t.Task IN (51,52) " & _
                                      " AND t.Person = a.Person) "
                                      
        Case Else
            ExclApptMenSQL = " "
        End Select
    Case vbUnchecked
        ExclApptMenSQL = " "
    End Select

    'asst check box filters
    Select Case frmTMSInsertStudent.chkBrosAsst.value
    Case vbChecked
        BroOnlyAsstSQL = " AND c.GenderMF = 'M' "
    Case vbUnchecked
        BroOnlyAsstSQL = ""
    End Select
    
    'asst check box filters
    Select Case frmTMSInsertStudent.chkSistersAsst.value
    Case vbChecked
        SisOnlyAsstSQL = " AND c.GenderMF = 'F' "
    Case vbUnchecked
        SisOnlyAsstSQL = ""
    End Select
    
    
    
    Select Case frmTMSInsertStudent.chkShowEandMSOnly.value
    Case vbChecked
        TaskSQL = " a.Task = " & rstTMSSchedule!TaskNo & " "
    Case vbUnchecked
        Select Case frmTMSInsertStudent.chkExclAppointedMen.value
        Case vbChecked
            TaskSQL = " a.Task = " & rstTMSSchedule!TaskNo & " " & _
                      " AND NOT EXISTS (SELECT 1 FROM tblTaskAndPerson t " & _
                      "                 WHERE t.Task IN (51,52) " & _
                                      " AND t.Person = a.Person) "
                                          
        Case vbUnchecked
            TaskSQL = " a.Task = " & rstTMSSchedule!TaskNo & " "
        End Select
    End Select
    
      
    lSuspTask = IIf(TheGridName = "flxInsertStudent", rstTMSSchedule!TaskNo, 43)
    
    SuspendDateSQL = " AND NOT EXISTS (SELECT 1 " & _
                    "FROM tblTaskPersonSuspendDates f " & _
                    "WHERE Task = " & lSuspTask & _
                    " AND a.Person = f.Person " & _
                    " AND (SuspendStartDate <= #" & ActualTalkDate & "# AND SuspendEndDate >= #" & ActualTalkDate & "#)) "
                        
    Select Case frmTMSInsertStudent.optIntelligent.value
    Case True 'intelligent
    
        lTOPVal = GlobalParms.GetValue("TMS_NumberOfRowsToShowOnInsertForm", "NumVal", 0)
        If lTOPVal = 0 Then
            TopSQL = ""
         Else
            TopSQL = " TOP " & lTOPVal & " "
        End If
    
        OrderingSQL = "ORDER BY b.TMSWeighting"
        
        
        If TheGridName = "flxInsertStudent" Then
            '
            'ie The InsertStudent grid is currently being populated
            '
            'Take account of suspend dates
            '
            Select Case CurrentTalkNo
            Case "R", "MR"
                'not taking account of school for oral review reader
                CompleteSQL = "SELECT DISTINCT " & TopSQL & " a.Person, " & _
                                "b.TMSWeighting, " & _
                                "b.TMSPrayerWeighting " & _
                                "FROM ((tblTaskPersonSuspendDates a " & _
                                "INNER JOIN tblTMSWeightings b ON " & _
                                "a.Person = b.PersonID) " & _
                                "INNER JOIN tblNameAddress c ON " & _
                                " (c.ID = b.PersonID)) " & _
                                "WHERE " & TaskSQL & _
                                SuspendDateSQL & " AND Active = TRUE " & _
                                BroOnlySQL & SisOnlySQL & OrderingSQL
            Case Else
                CompleteSQL = "SELECT DISTINCT " & TopSQL & " a.Person, " & _
                                "b.TMSWeighting, " & _
                                "b.TMSPrayerWeighting " & _
                                "FROM (((tblTaskPersonSuspendDates a " & _
                                "INNER JOIN tblTMSWeightings b ON " & _
                                "a.Person = b.PersonID) " & _
                                "INNER JOIN tblNameAddress c ON " & _
                                " (c.ID = b.PersonID)) " & _
                                "INNER JOIN tblTMSSchoolAndPerson d ON " & _
                                " d.PersonID = a.Person AND d.Task = a.Task) " & _
                                "WHERE " & TaskSQL & _
                                SuspendDateSQL & " AND Active = TRUE " & _
                                " AND d.SchoolNo = " & SchoolNum & " " & _
                                BroOnlySQL & SisOnlySQL & OrderingSQL
            End Select
            
                            
            Set rstStudentList = CMSDB.OpenRecordset(CompleteSQL, dbOpenForwardOnly)
                                                                                                
        Else
'            '
'            'ie The InsertAssistant grid is currently being populated
'            '


            '
            'Take account of suspend dates
            '
            CompleteSQL = "SELECT DISTINCT " & TopSQL & " a.Person, " & _
                            "b.TMSAsstWeighting " & _
                            "FROM (((tblTaskPersonSuspendDates a " & _
                            "INNER JOIN tblTMSWeightings b ON " & _
                            "a.Person = b.PersonID) " & _
                            "INNER JOIN tblNameAddress c ON " & _
                            " (c.ID = b.PersonID)) " & _
                            "INNER JOIN tblTMSSchoolAndPerson d ON " & _
                            " d.PersonID = a.Person AND d.Task = a.Task) " & _
                            "WHERE a.Task IN (42,43)" & _
                            SuspendDateSQL & " AND Active = TRUE " & _
                            " AND d.SchoolNo = " & SchoolNum & " " & _
                            ExclApptMenSQL & BroOnlyAsstSQL & SisOnlyAsstSQL & _
                            "ORDER BY b.TMSAsstWeighting"
            
            Set rstStudentList = CMSDB.OpenRecordset(CompleteSQL, dbOpenForwardOnly)
                                                                
        End If
        
    Case False 'non-intelligent
    
        If TheGridName = "flxInsertStudent" Then
            '
            'ie The InsertStudent grid is currently being populated
            '
            'Take account of suspend dates
            '
            Set rstStudentList = CMSDB.OpenRecordset("SELECT DISTINCT a.Person, c.FirstName, " & _
                                                    "c.MiddleName, c.LastName " & _
                                                "FROM ((tblTaskPersonSuspendDates a " & _
                                                "INNER JOIN tblNameAddress c ON " & _
                                                " (a.Person = c.ID)) " & _
                                                "INNER JOIN tblTMSSchoolAndPerson d ON " & _
                                                " d.PersonID = a.Person AND d.Task = a.Task) " & _
                                                "WHERE " & TaskSQL & _
                                                SuspendDateSQL & " AND Active = TRUE " & _
                                                " AND d.SchoolNo = " & SchoolNum & " " & _
                                                BroOnlySQL & SisOnlySQL & _
                                                " ORDER BY LastName, FirstName, MiddleName", dbOpenForwardOnly)
                                                                                                
        Else
            '
            'ie The InsertAssistant grid is currently being populated
            '
            'Take account of suspend dates
            '
            Set rstStudentList = CMSDB.OpenRecordset("SELECT DISTINCT a.Person, c.FirstName, " & _
                                                    "c.MiddleName, c.LastName " & _
                                                "FROM ((tblTaskPersonSuspendDates a " & _
                                                "INNER JOIN tblNameAddress c ON " & _
                                                " (a.Person = c.ID)) " & _
                                                "INNER JOIN tblTMSSchoolAndPerson d ON " & _
                                                " d.PersonID = a.Person AND d.Task = a.Task) " & _
                                                "WHERE a.Task IN (42,43) " & _
                                                SuspendDateSQL & " AND Active = TRUE " & _
                                                " AND d.SchoolNo = " & SchoolNum & " " & _
                                                " ORDER BY LastName, FirstName, MiddleName", dbOpenForwardOnly)
                                                                
        End If
    End Select
    
    
    FillTMSStudentGrid_2016 = True 'init
    
    i = 2 'row of grid
    j = 0 'position on array storing each student's prev/next dates
    
    '
    'Fill grid with last/prev dates. Actual columns used is determined by the PrevXxxx
    ' variables which are set in frmTMSInsertStudent.SetUpGrids according to the
    ' talkno selected on frmTMSScheduling.
    '
    
    TheGrid.Rows = 2
    With rstStudentList
    If Not .BOF Then
        Do While Not .EOF
            TheGrid.Rows = i + 1
            
            If CongregationMember.TMSStudentOnThisWeekFAST(!Person, CurrentAssignmentDate) Then
                frmTMSInsertStudent.GridRowBackColourToGrey TheGrid, i, True, False
            Else
                frmTMSInsertStudent.GridRowBackColourToGrey TheGrid, i, False, False
            End If
                        
            TheGrid.TextMatrix(i, 0) = CongregationMember.LastFirstNameMiddleName(!Person)
            
            lCurrentPersonInGrid = !Person
            TheGrid.TextMatrix(i, TheGrid.Cols - 1) = lCurrentPersonInGrid 'NameID used for programmatic purposes.
            
            If frmTMSInsertStudent.chkShowNextPrevDates.value = vbChecked Then
                If TheGridName = "flxInsertStudent" Then
                
                    ReDim Preserve NextPrevDates_2016(j)
                    NextPrevDates_2016(j).ThePersonID = !Person
                    
                    '
                    'Put next/prev dates into an array
                    '
                    NextPrevDates_2016(j).NextPrevInfo(0) = CongregationMember.TMSPrevDateFAST(!Person, "'BR','1'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(1) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(2) = CongregationMember.TMSNextDateFAST(!Person, "'BR','1'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(3) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(4) = CongregationMember.TMSPrevDateFAST(!Person, "'IC','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(5) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(6) = CongregationMember.TMSNextDateFAST(!Person, "'IC','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(7) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(8) = CongregationMember.TMSPrevDateFAST(!Person, "'RV','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(9) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(10) = CongregationMember.TMSNextDateFAST(!Person, "'RV','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(11) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(12) = CongregationMember.TMSPrevDateFAST(!Person, "'BS','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(13) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(14) = CongregationMember.TMSNextDateFAST(!Person, "'BS','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(15) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(16) = CongregationMember.TMSPrevDateFAST(!Person, "Asst", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(17) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(18) = CongregationMember.TMSNextDateFAST(!Person, "Asst", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(19) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(20) = CongregationMember.TMSPrevDateFAST(!Person, "'O','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(21) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(22) = CongregationMember.TMSNextDateFAST(!Person, "'O','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(23) = CongregationMember.GetTMSSchoolNoForInsertForm
                    
                Else
                    '
                    'This is Assistant-grid. Get prev/next dates from array.
                    '
                    'Locate current student in array to get their prev/next dates
                    '
                    For j = 0 To UBound(NextPrevDates_2016)
                        If NextPrevDates_2016(j).ThePersonID = !Person Then
                            Exit For
                        End If
                    Next j
                    
                End If
                 
                '
                'Put array contents into student grid. If Student is not in array
                ' (ie because they're only a householder, so wouldn't have been included
                ' when processing Student Grid), then get dates from DB
                '
                If j <= UBound(NextPrevDates_2016) Then
                    With TheGrid
                    .TextMatrix(i, PrevBR) = NextPrevDates_2016(j).NextPrevInfo(0)
                    .TextMatrix(i, PrevBRSchool) = NextPrevDates_2016(j).NextPrevInfo(1)
                    .TextMatrix(i, NextBR) = NextPrevDates_2016(j).NextPrevInfo(2)
                    .TextMatrix(i, NextBRSchool) = NextPrevDates_2016(j).NextPrevInfo(3)
                    .TextMatrix(i, PrevIC) = NextPrevDates_2016(j).NextPrevInfo(4)
                    .TextMatrix(i, PrevICSchool) = NextPrevDates_2016(j).NextPrevInfo(5)
                    .TextMatrix(i, NextIC) = NextPrevDates_2016(j).NextPrevInfo(6)
                    .TextMatrix(i, NextICSchool) = NextPrevDates_2016(j).NextPrevInfo(7)
                    .TextMatrix(i, PrevRV) = NextPrevDates_2016(j).NextPrevInfo(8)
                    .TextMatrix(i, PrevRVSchool) = NextPrevDates_2016(j).NextPrevInfo(9)
                    .TextMatrix(i, NextRV) = NextPrevDates_2016(j).NextPrevInfo(10)
                    .TextMatrix(i, NextRVSchool) = NextPrevDates_2016(j).NextPrevInfo(11)
                    .TextMatrix(i, PrevBS) = NextPrevDates_2016(j).NextPrevInfo(12)
                    .TextMatrix(i, PrevBSSchool) = NextPrevDates_2016(j).NextPrevInfo(13)
                    .TextMatrix(i, NextBS) = NextPrevDates_2016(j).NextPrevInfo(14)
                    .TextMatrix(i, NextBSSchool) = NextPrevDates_2016(j).NextPrevInfo(15)
                    .TextMatrix(i, PrevAsst) = NextPrevDates_2016(j).NextPrevInfo(16)
                    .TextMatrix(i, PrevAsstSchool) = NextPrevDates_2016(j).NextPrevInfo(17)
                    .TextMatrix(i, NextAsst) = NextPrevDates_2016(j).NextPrevInfo(18)
                    .TextMatrix(i, NextAsstSchool) = NextPrevDates_2016(j).NextPrevInfo(19)
                    .TextMatrix(i, PrevO) = NextPrevDates_2016(j).NextPrevInfo(20)
                    .TextMatrix(i, PrevOSchool) = NextPrevDates_2016(j).NextPrevInfo(21)
                    .TextMatrix(i, NextO) = NextPrevDates_2016(j).NextPrevInfo(22)
                    .TextMatrix(i, NextOSchool) = NextPrevDates_2016(j).NextPrevInfo(23)
                    
                   
                    End With
                Else
                    '
                    'Go to   s l o w   DB. Add new student to array, and get the dates....
                    '
                    ReDim Preserve NextPrevDates_2016(j)
                    NextPrevDates_2016(j).ThePersonID = !Person
                    
                    NextPrevDates_2016(j).NextPrevInfo(0) = CongregationMember.TMSPrevDateFAST(!Person, "'BR','1'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(1) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(2) = CongregationMember.TMSNextDateFAST(!Person, "'BR','1'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(3) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(4) = CongregationMember.TMSPrevDateFAST(!Person, "'IC','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(5) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(6) = CongregationMember.TMSNextDateFAST(!Person, "'IC','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(7) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(8) = CongregationMember.TMSPrevDateFAST(!Person, "'RV','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(9) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(10) = CongregationMember.TMSNextDateFAST(!Person, "'RV','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(11) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(12) = CongregationMember.TMSPrevDateFAST(!Person, "'BS','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(13) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(14) = CongregationMember.TMSNextDateFAST(!Person, "'BS','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(15) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(16) = CongregationMember.TMSPrevDateFAST(!Person, "Asst", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(17) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(18) = CongregationMember.TMSNextDateFAST(!Person, "Asst", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(19) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(20) = CongregationMember.TMSPrevDateFAST(!Person, "'O','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(21) = CongregationMember.GetTMSSchoolNoForInsertForm
                    NextPrevDates_2016(j).NextPrevInfo(22) = CongregationMember.TMSNextDateFAST(!Person, "'O','1','2','3'", CurrentAssignmentDate)
                    NextPrevDates_2016(j).NextPrevInfo(23) = CongregationMember.GetTMSSchoolNoForInsertForm
                                        
                                        
                    With TheGrid
                    .TextMatrix(i, PrevBR) = NextPrevDates_2016(j).NextPrevInfo(0)
                    .TextMatrix(i, PrevBRSchool) = NextPrevDates_2016(j).NextPrevInfo(1)
                    .TextMatrix(i, NextBR) = NextPrevDates_2016(j).NextPrevInfo(2)
                    .TextMatrix(i, NextBRSchool) = NextPrevDates_2016(j).NextPrevInfo(3)
                    .TextMatrix(i, PrevIC) = NextPrevDates_2016(j).NextPrevInfo(4)
                    .TextMatrix(i, PrevICSchool) = NextPrevDates_2016(j).NextPrevInfo(5)
                    .TextMatrix(i, NextIC) = NextPrevDates_2016(j).NextPrevInfo(6)
                    .TextMatrix(i, NextICSchool) = NextPrevDates_2016(j).NextPrevInfo(7)
                    .TextMatrix(i, PrevRV) = NextPrevDates_2016(j).NextPrevInfo(8)
                    .TextMatrix(i, PrevRVSchool) = NextPrevDates_2016(j).NextPrevInfo(9)
                    .TextMatrix(i, NextRV) = NextPrevDates_2016(j).NextPrevInfo(10)
                    .TextMatrix(i, NextRVSchool) = NextPrevDates_2016(j).NextPrevInfo(11)
                    .TextMatrix(i, PrevBS) = NextPrevDates_2016(j).NextPrevInfo(12)
                    .TextMatrix(i, PrevBSSchool) = NextPrevDates_2016(j).NextPrevInfo(13)
                    .TextMatrix(i, NextBS) = NextPrevDates_2016(j).NextPrevInfo(14)
                    .TextMatrix(i, NextBSSchool) = NextPrevDates_2016(j).NextPrevInfo(15)
                    .TextMatrix(i, PrevAsst) = NextPrevDates_2016(j).NextPrevInfo(16)
                    .TextMatrix(i, PrevAsstSchool) = NextPrevDates_2016(j).NextPrevInfo(17)
                    .TextMatrix(i, NextAsst) = NextPrevDates_2016(j).NextPrevInfo(18)
                    .TextMatrix(i, NextAsstSchool) = NextPrevDates_2016(j).NextPrevInfo(19)
                    .TextMatrix(i, PrevO) = NextPrevDates_2016(j).NextPrevInfo(20)
                    .TextMatrix(i, PrevOSchool) = NextPrevDates_2016(j).NextPrevInfo(21)
                    .TextMatrix(i, NextO) = NextPrevDates_2016(j).NextPrevInfo(22)
                    .TextMatrix(i, NextOSchool) = NextPrevDates_2016(j).NextPrevInfo(23)
                    
                    
                    End With
                                        
                End If
                 
                 '
                 'Now highlight in green the most recent
                 ' next & previous talks given by each student. This aids user in
                 ' spreading talks more evenly.
                 '
                 With TheGrid
                 '
                 'Put pertinent PrevDates into an array
                 ' (Must change format since can only use variants in Array function)
                 '
                 
                 'col  | talk type
                 ' 0   |   Prev BR
                 ' 4   |   Prev IC
                 ' 8   |   Prev RV
                 ' 12  |   Prev BS
                 ' 16  |   Prev Asst
                 ' 20  |   Prev O
                 
                 
                 
                PrevDates = Array(Format(NextPrevDates_2016(j).NextPrevInfo(0), "yyyy/mm/dd"), _
                               Format(NextPrevDates_2016(j).NextPrevInfo(4), "yyyy/mm/dd"), _
                               Format(NextPrevDates_2016(j).NextPrevInfo(8), "yyyy/mm/dd"), _
                               Format(NextPrevDates_2016(j).NextPrevInfo(12), "yyyy/mm/dd"), _
                               Format(NextPrevDates_2016(j).NextPrevInfo(20), "yyyy/mm/dd"))
                               
                sPrevAsstDate = Format(NextPrevDates_2016(j).NextPrevInfo(16), "yyyy/mm/dd")
                                 
                 
                 '
                 'Sort desc
                 '
                 BubbleSort PrevDates, , True
                 
                 .Row = i
                 
                If PrevDates(0) <> "" Then
                    Select Case Format(PrevDates(0), "dd/mm/yy")
                    Case NextPrevDates_2016(j).NextPrevInfo(0)
                       .col = PrevBR
                       .CellBackColor = PrevTalkColour
                       .col = PrevBRSchool
                       .CellBackColor = PrevTalkColour
                    Case NextPrevDates_2016(j).NextPrevInfo(4)
                       .col = PrevIC
                       .CellBackColor = PrevTalkColour
                       .col = PrevICSchool
                       .CellBackColor = PrevTalkColour
                    Case NextPrevDates_2016(j).NextPrevInfo(8)
                       .col = PrevRV
                       .CellBackColor = PrevTalkColour
                       .col = PrevRVSchool
                       .CellBackColor = PrevTalkColour
                    Case NextPrevDates_2016(j).NextPrevInfo(12)
                       .col = PrevBS
                       .CellBackColor = PrevTalkColour
                       .col = PrevBSSchool
                       .CellBackColor = PrevTalkColour
                    Case NextPrevDates_2016(j).NextPrevInfo(20)
                       .col = PrevO
                       .CellBackColor = PrevTalkColour
                       .col = PrevOSchool
                       .CellBackColor = PrevTalkColour
                    End Select
                End If
                
                sPrevTalkDate = PrevDates(0)
                
                 '
                 'Put pertinent NextDates into an array
                 ' (Must change format since can only use variants in Array function)
                 '
                 
                 'col | talk type
                 ' 2  |   Next BR
                 ' 6  |   Next IC
                 ' 10 |   Next RV
                 ' 14 |   Next BS
                 ' 18 |   Next Asst
                 ' 22 |   Next O

                NextDates = Array(Format(NextPrevDates_2016(j).NextPrevInfo(2), "yyyy/mm/dd"), _
                               Format(NextPrevDates_2016(j).NextPrevInfo(6), "yyyy/mm/dd"), _
                               Format(NextPrevDates_2016(j).NextPrevInfo(10), "yyyy/mm/dd"), _
                               Format(NextPrevDates_2016(j).NextPrevInfo(14), "yyyy/mm/dd"), _
                               Format(NextPrevDates_2016(j).NextPrevInfo(22), "yyyy/mm/dd"))
                               
                sNextAsstDate = Format(NextPrevDates_2016(j).NextPrevInfo(18), "yyyy/mm/dd")
                 
                 '
                 'Sort asc
                 '
                 BubbleSort NextDates, , False
                 
                 '
                 'Find the minimum non-blank date
                 '
                For n = 0 To UBound(NextDates)
                    If NextDates(n) <> "" Then
                        Exit For
                    End If
                Next n
                 
                If n > UBound(NextDates) Then
                    n = UBound(NextDates)
                End If
                 
                 .Row = i
                 
                If NextDates(n) <> "" Then
                    Select Case Format(NextDates(n), "dd/mm/yy")
                    Case NextPrevDates_2016(j).NextPrevInfo(2)
                       .col = NextBR
                       .CellBackColor = NextTalkColour
                       .col = NextBRSchool
                       .CellBackColor = NextTalkColour
                    Case NextPrevDates_2016(j).NextPrevInfo(6)
                       .col = NextIC
                       .CellBackColor = NextTalkColour
                       .col = NextICSchool
                       .CellBackColor = NextTalkColour
                    Case NextPrevDates_2016(j).NextPrevInfo(10)
                       .col = NextRV
                       .CellBackColor = NextTalkColour
                       .col = NextRVSchool
                       .CellBackColor = NextTalkColour
                    Case NextPrevDates_2016(j).NextPrevInfo(14)
                       .col = NextBS
                       .CellBackColor = NextTalkColour
                       .col = NextBSSchool
                       .CellBackColor = NextTalkColour
                    Case NextPrevDates_2016(j).NextPrevInfo(22)
                       .col = NextO
                       .CellBackColor = NextTalkColour
                       .col = NextOSchool
                       .CellBackColor = NextTalkColour
                    End Select
                End If
                
                sNextTalkDate = NextDates(n)
                
                End With
                
            End If
            
            
            'so we now have the most recent and next talk dates,
            ' and the most recent and next asst dates.
            'If the asst dates are closer than the talk dates, highlight the asst dates too.
            
            
             With TheGrid
             
             If sPrevAsstDate <> "" Then
                If (sPrevAsstDate > sPrevTalkDate) Or (sPrevTalkDate = "") Then
                    .col = PrevAsst
                    .CellBackColor = PrevAsstColour
                    .col = PrevAsstSchool
                    .CellBackColor = PrevAsstColour
                End If
             End If
             
             If sNextAsstDate <> "" Then
                If (sNextAsstDate < sNextTalkDate) Or (sNextTalkDate = "") Then
                    .col = NextAsst
                    .CellBackColor = NextAsstColour
                    .col = NextAsstSchool
                    .CellBackColor = NextAsstColour
                End If
             End If
             
            
             End With
            
                                        
            i = i + 1
            
            If TheGridName = "flxInsertStudent" Then
                j = j + 1
            End If
            
            .MoveNext
        Loop
    Else
        'no students to display
        FillTMSStudentGrid_2016 = False
    End If
    
    End With
    
    Screen.MousePointer = vbNormal
    
    
    Exit Function
ErrorTrap:
    EndProgram
    
End Function



Public Function SeekTheFile(TheFileName As String, _
                            TheDialogTitle As String, _
                            FileOpenDialog As CommonDialog) As Boolean
Dim SaveCurDir As String

    On Error GoTo ExitNow

    SeekTheFile = False
    
    '
    'Set up dialogue parms
    '
    FileOpenDialog.Filter = "All files (*.*)|*.*|Text files|*.txt"
    FileOpenDialog.FilterIndex = 2
    FileOpenDialog.DefaultExt = "txt"
    FileOpenDialog.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist Or _
        cdlOFNNoReadOnlyReturn Or cdlOFNNoChangeDir
    FileOpenDialog.DialogTitle = TheDialogTitle
    'FileOpenDialog.InitDir = CurDir$
'    FileOpenDialog.FileName = FileName
    
    FileOpenDialog.CancelError = True ' Exit if user presses Cancel.
    
    FileOpenDialog.ShowOpen
    TheFileName = FileOpenDialog.Filename
    SeekTheFile = True

ExitNow:

End Function

Public Sub RefreshGrid(Optional ReselectRow As Boolean = False)
Dim SchoolNum As Long, StoreRow As Long
    On Error GoTo ErrorTrap
    
    '
    'Refresh the grid on frmTMSScheduling
    '
    
    Screen.MousePointer = vbHourglass
    
    With frmTMSScheduling
    
    !flxTMSSchedule.Redraw = False
    '
    'StoreRow used to maintain row selection following refresh
    '
    If !flxTMSSchedule.Row > 0 And ReselectRow Then
        StoreRow = !flxTMSSchedule.Row
    Else
        StoreRow = 0
    End If
    
    If !cmbYear.ListIndex > -1 And !cmbMonth.ListIndex > -1 Then

        Select Case True
        Case !opt1stSchool:
            SchoolNum = 1
        Case !opt2ndSchool:
            SchoolNum = 2
        Case !opt3rdSchool:
            SchoolNum = 3
        End Select

        '
        'Acquire 2 recsets - one for the items, the other for the students assigned to items.
        '
        GetTMSScheduleRecordset !cmbYear.ItemData(!cmbYear.ListIndex), !cmbMonth.ItemData(!cmbMonth.ListIndex), SchoolNum, rstTMSSchedule, rstTMSStudents
        FillTMSScheduleGrid
        
        '
        'StoreRow used to maintain row selection following refresh
        '
        If StoreRow > 0 And StoreRow < .flxTMSSchedule.Rows - 1 Then
            !flxTMSSchedule.Row = StoreRow
            '
            'Make selected row red again since FillTMSScheduleGrid rebuilds grid as black
            '
            .SortRowColour
            .flxTMSSchedule_Click
        Else
'            .GridToBlack
            !lblSourceDisplay = ""
            !lblThemeDisplay = ""
            !lblBroOnly = ""
            !lblSetting = ""
            !lblCounsel = ""
        End If
        
    Else
        !flxTMSSchedule.Rows = 1
        .lblSch1.ForeColor = vbBlack
        .lblSch2.ForeColor = vbBlack
        .lblSch3.ForeColor = vbBlack
    End If
        
    
    '09/10/2006 - this line was stopping searched for names from being blue. Don't think
    ' we need it anyway
    '.GridToBlack
    
    !flxTMSSchedule.Redraw = True
    
    End With
    

    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub


Public Function BuildTMSWeightingsTable() As Boolean
Dim TheError As Integer, rstAllTMSMembers As Recordset

On Error GoTo ErrorTrap
    
    DelAllRows "tblTMSWeightings"
    
    Set rstAllTMSMembers = CMSDB.OpenRecordset("SELECT DISTINCT Person " & _
                                        "FROM tblTaskPersonSuspendDates " & _
                                        "WHERE TaskCategory = " & 4 & _
                                        " AND TaskSubCategory = " & 6, dbOpenDynaset)
    
    With rstAllTMSMembers
    
    If Not .BOF Then
        BuildTMSWeightingsTable = True
        Do Until .EOF
            CMSDB.Execute "INSERT INTO tblTMSWeightings " & _
                              "(PersonID, TMSWeighting, TMSPrayerWeighting, TMSAsstWeighting) " & _
                              "VALUES (" & !Person & ", " & 0 & ", " & 0 & ", " & 0 & ")"
                              
            .MoveNext
        Loop
    Else
        BuildTMSWeightingsTable = False
        Exit Function
    End If
    
    End With
    
    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Sub CalculateTMSWeightings()
On Error GoTo ErrorTrap
Dim weekno As Long, i As Long, rstTMSWeightings As Recordset
Dim TempDate As Date, SQLStr As String, rstTMSStudentsOnSchedule As Recordset, PrevDate As Date
Dim TheWeighting As Double, TheAsstWeighting As Double, ProcessDate As Date, sTalkNo As String
Dim bInsertingAsst As Boolean, TheAsstListWtg As Double, lStudentID As Long, lAsstID As Long
Dim sPrevStuOrAsst As String, LastAssignmentDate As Date

    Screen.MousePointer = vbHourglass
    
    CMSDB.Execute "UPDATE tblTMSWeightings " & _
                    "SET TMSWeighting = 0, " & _
                        "TMSPrayerWeighting = 0, " & _
                        "TMSAsstWeighting = 0 "

    
    SQLStr = "SELECT tblTMSSchedule.AssignmentDate, " & _
            "tblTMSSchedule.TalkNo, " & _
            "tblTMSSchedule.PersonID, " & _
            "tblTMSSchedule.Assistant1ID, " & _
            "tblTMSSchedule.SchoolNo, " & _
            "tblTMSSchedule.ScheduleSeqNum " & _
            "FROM tblTMSSchedule " & _
            "WHERE TalkDefaulted = FALSE " & _
            "ORDER BY 1"

            

    Set rstTMSStudentsOnSchedule = CMSDB.OpenRecordset(SQLStr, dbOpenDynaset)
    
    
    weekno = WeeksToCheck
    
    Set rstTMSWeightings = CMSDB.OpenRecordset("SELECT PersonID, TMSWeighting, TMSAsstWeighting, TMSPrayerWeighting " & _
                                        "FROM tblTMSWeightings ", dbOpenDynaset)
        
    TempDate = CDate(Format(frmTMSScheduling!flxTMSSchedule.TextMatrix(frmTMSScheduling!flxTMSSchedule.Row, 0), "dd/mm/yyyy"))
    
    PrevDate = TempDate
    
    With rstTMSStudentsOnSchedule
    '
    'Work back from current date, updating each student's weighting
    '
    .FindLast "AssignmentDate <= #" & Format(TempDate, "mm/dd/yyyy") & "#"
    
    
    If Not .NoMatch Then
        ProcessDate = CDate(Format(!AssignmentDate, "dd/mm/yyyy"))
        
        'these two weightings are combined to give weightings for frmInsertStudent's *Insert Student* list
        TheWeighting = GetTMSWeighting(Abs(DateDiff("ww", ProcessDate, TempDate)), !TalkNo, ProcessDate)
        TheAsstWeighting = GetTMSWeighting(Abs(DateDiff("ww", ProcessDate, TempDate)), "Asst", ProcessDate)
        
        'this weighting is for frmInsertStudent's *Insert Assistant* list
        TheAsstListWtg = GetTMSAsstListWeighting(Abs(DateDiff("ww", ProcessDate, TempDate)), !TalkNo, ProcessDate)
        
        Do While weekno >= 0
            lStudentID = !PersonID
            rstTMSWeightings.FindFirst "PersonID = " & lStudentID
            If Not rstTMSWeightings.NoMatch Then
                rstTMSWeightings.Edit
                If !TalkNo <> "P" Then
                    rstTMSWeightings!TMSWeighting = rstTMSWeightings!TMSWeighting + TheWeighting
                    rstTMSWeightings.Update
                    If !TalkNo = "IC" Or !TalkNo = "RV" Or !TalkNo = "BS" Or !TalkNo = "O" Then
                        rstTMSWeightings.Edit
                        'rstTMSWeightings!TMSWeighting = rstTMSWeightings!TMSWeighting + TheAsstWeighting
                        rstTMSWeightings!TMSAsstWeighting = rstTMSWeightings!TMSAsstWeighting + TheAsstListWtg
                        rstTMSWeightings.Update
                    End If
                Else
                    rstTMSWeightings!TMSPrayerWeighting = rstTMSWeightings!TMSPrayerWeighting + TheWeighting
                    rstTMSWeightings.Update
                End If
                
            End If
            
            lAsstID = !Assistant1ID
            If lAsstID <> 0 Then
                rstTMSWeightings.FindFirst "PersonID = " & lAsstID
                If Not rstTMSWeightings.NoMatch Then
                    rstTMSWeightings.Edit
                    
                    If CongregationMember.Does_Only_TMS_Asst(!Assistant1ID) Then  'asst only
                        rstTMSWeightings!TMSAsstWeighting = rstTMSWeightings!TMSAsstWeighting + _
                                                            TheAsstListWtg * TMSWeightingIfAssistantOnly
                        rstTMSWeightings.Update
                    Else
                        rstTMSWeightings!TMSWeighting = rstTMSWeightings!TMSWeighting + TheAsstWeighting
                        rstTMSWeightings!TMSAsstWeighting = rstTMSWeightings!TMSAsstWeighting + TheAsstListWtg
                        rstTMSWeightings.Update
                    End If
                        
                End If
            End If
            
            
            .MovePrevious
            If .BOF Then
                Exit Do
            End If
            
            If PrevDate <> !AssignmentDate Then
                weekno = weekno - 1
                ProcessDate = CDate(Format(!AssignmentDate, "dd/mm/yyyy"))
            End If
            
            'these two weightings are combined to give weightings for frmInsertStudent's *Insert Student* list
            TheWeighting = GetTMSWeighting(Abs(DateDiff("ww", ProcessDate, TempDate)), !TalkNo, ProcessDate)
            TheAsstWeighting = GetTMSWeighting(Abs(DateDiff("ww", ProcessDate, TempDate)), "Asst", ProcessDate)
            
            'this weighting is for frmInsertStudent's *Insert Assistant* list
            TheAsstListWtg = GetTMSAsstListWeighting(Abs(DateDiff("ww", ProcessDate, TempDate)), !TalkNo, ProcessDate)
            
            PrevDate = !AssignmentDate
        Loop
    End If
    
    '
    'Work forward from current date, updating each student's weighting
    '
    .FindFirst "AssignmentDate >= #" & Format(TempDate, "mm/dd/yyyy") & "#"
    PrevDate = TempDate

    weekno = 0
    
    If Not .NoMatch Then
        ProcessDate = CDate(Format(!AssignmentDate, "dd/mm/yyyy"))
        
        'these two weightings are combined to give weightings for frmInsertStudent's *Insert Student* list
        TheWeighting = GetTMSWeighting(Abs(DateDiff("ww", ProcessDate, TempDate)), !TalkNo, ProcessDate)
        TheAsstWeighting = GetTMSWeighting(Abs(DateDiff("ww", ProcessDate, TempDate)), "Asst", ProcessDate)
        
        'this weighting is for frmInsertStudent's *Insert Assistant* list
        TheAsstListWtg = GetTMSAsstListWeighting(Abs(DateDiff("ww", ProcessDate, TempDate)), !TalkNo, ProcessDate)
        
        Do While weekno <= WeeksToCheck
            lStudentID = !PersonID
            rstTMSWeightings.FindFirst "PersonID = " & lStudentID
            If Not rstTMSWeightings.NoMatch Then
                rstTMSWeightings.Edit
                If !TalkNo <> "P" Then
                    rstTMSWeightings!TMSWeighting = rstTMSWeightings!TMSWeighting + TheWeighting
                    rstTMSWeightings.Update
                    If !TalkNo = "IC" Or !TalkNo = "RV" Or !TalkNo = "BS" Or !TalkNo = "O" Then
                        rstTMSWeightings.Edit
                        'rstTMSWeightings!TMSWeighting = rstTMSWeightings!TMSWeighting + TheAsstWeighting
                        rstTMSWeightings!TMSAsstWeighting = rstTMSWeightings!TMSAsstWeighting + TheAsstListWtg
                        rstTMSWeightings.Update
                    End If
                Else
                    rstTMSWeightings!TMSPrayerWeighting = rstTMSWeightings!TMSPrayerWeighting + TheWeighting
                    rstTMSWeightings.Update
                End If
                
            End If
            
            lAsstID = !Assistant1ID
            If lAsstID <> 0 Then
                rstTMSWeightings.FindFirst "PersonID = " & lAsstID
                If Not rstTMSWeightings.NoMatch Then
                    rstTMSWeightings.Edit
                    
                    If CongregationMember.Does_Only_TMS_Asst(!Assistant1ID) Then  'asst only
                        rstTMSWeightings!TMSAsstWeighting = rstTMSWeightings!TMSAsstWeighting + _
                                                            TheAsstListWtg * TMSWeightingIfAssistantOnly
                        rstTMSWeightings.Update
                    Else
                        rstTMSWeightings!TMSWeighting = rstTMSWeightings!TMSWeighting + TheAsstWeighting
                        rstTMSWeightings!TMSAsstWeighting = rstTMSWeightings!TMSAsstWeighting + TheAsstListWtg
                        rstTMSWeightings.Update
                    End If
                        
                End If
            End If
            
            
            .MoveNext
            If .EOF Then
                Exit Do
            End If
    
            
            If PrevDate <> !AssignmentDate Then
                weekno = weekno + 1
                ProcessDate = CDate(Format(!AssignmentDate, "dd/mm/yyyy"))
            End If
            
            'these two weightings are combined to give weightings for frmInsertStudent's *Insert Student* list
            TheWeighting = GetTMSWeighting(Abs(DateDiff("ww", ProcessDate, TempDate)), !TalkNo, ProcessDate)
            TheAsstWeighting = GetTMSWeighting(Abs(DateDiff("ww", ProcessDate, TempDate)), "Asst", ProcessDate)
            
            'this weighting is for frmInsertStudent's *Insert Assistant* list
            TheAsstListWtg = GetTMSAsstListWeighting(Abs(DateDiff("ww", ProcessDate, TempDate)), !TalkNo, ProcessDate)
            
            PrevDate = !AssignmentDate
        Loop
    End If
    
    End With
    
    '
    'Now scale down the weightings of all brothers that do No3 talks so that they
    ' compare more favourably with sister's weightings. Otherwise, brothers
    ' won't get a look in!
    '
    With frmTMSScheduling!flxTMSSchedule
     If NewMtgArrangementStarted(CStr(TempDate)) = CLM2016 Then
        Select Case .TextMatrix(.Row, 1)
        Case "IC", "RV", "BS", "O"
            CMSDB.Execute "UPDATE tblTMSWeightings " & _
                          "SET TMSWeighting = (TMSWeighting * " & TMSScaleNo4BroWeightings + 0.01 & ")" & _
                          "WHERE PersonID IN " & _
                          " (SELECT Person " & _
                          "  FROM tblTaskAndPerson " & _
                          "  INNER JOIN tblNameAddress " & _
                          "  ON tblTaskAndPerson.Person = tblNameAddress.ID " & _
                          "  WHERE GenderMF = 'M' " & _
                          "  AND Task IN (105, 106, 107, 108))"
        End Select
   ElseIf NewMtgArrangementStarted(CStr(TempDate)) = TMS2009 Then
        If .TextMatrix(.Row, 1) = "3" Then
            CMSDB.Execute "UPDATE tblTMSWeightings " & _
                          "SET TMSWeighting = (TMSWeighting * " & TMSScaleNo4BroWeightings + 0.01 & ")" & _
                          "WHERE PersonID IN " & _
                          " (SELECT Person " & _
                          "  FROM tblTaskAndPerson " & _
                          "  INNER JOIN tblNameAddress " & _
                          "  ON tblTaskAndPerson.Person = tblNameAddress.ID " & _
                          "  WHERE GenderMF = 'M' " & _
                          "  AND Task = 101)"
        End If
    ElseIf NewMtgArrangementStarted(CStr(TempDate)) = Pre2009 Then
        If .TextMatrix(.Row, 1) = "4" Then
            CMSDB.Execute "UPDATE tblTMSWeightings " & _
                          "SET TMSWeighting = (TMSWeighting * " & TMSScaleNo4BroWeightings & ")" & _
                          "WHERE PersonID IN " & _
                          " (SELECT Person " & _
                          "  FROM tblTaskAndPerson " & _
                          "  INNER JOIN tblNameAddress " & _
                          "  ON tblTaskAndPerson.Person = tblNameAddress.ID " & _
                          "  WHERE GenderMF = 'M' " & _
                          "  AND TaskCategory = 4 " & _
                          "  AND TaskSubCategory = 6 " & _
                          "  AND Task IN (40, 41))"
        End If
    End If
    End With
    
    'now work through the wtg table and scale wtg to try and alternate talk/asst better
    If gbTMSAltAsstStu Then
        'sTalkNo = frmTMSInsertStudent.FormTalkNo
        sTalkNo = frmTMSScheduling.CurrentTalkNum
    '    bInsertingAsst = (frmTMSInsertStudent.tabInsertStudent = 1)
        If sTalkNo = "IC" Or sTalkNo = "RV" Or sTalkNo = "BS" Or sTalkNo = "O" Then

            With rstTMSWeightings

            .MoveFirst
            Do Until .BOF Or .EOF
                If Not CongregationMember.Does_Only_TMS_Asst(!PersonID) And _
                    Not CongregationMember.Does_TMS_2_3_4_NOT_Asst_2016(!PersonID) Then

                        sPrevStuOrAsst = CongregationMember.TMSPreviousAssignment_Student_or_Asst(!PersonID, _
                                                                             frmTMSScheduling.CurrentAssignmentDate, _
                                                                             LastAssignmentDate)


                        If frmTMSScheduling.CurrentAssignmentDate <= DateAdd("m", glTMSMaxNoMonthsForSistersTalks, LastAssignmentDate) Then

                            If sPrevStuOrAsst = "STUDENT" Then
                                rstTMSWeightings.Edit
                                rstTMSWeightings!TMSWeighting = rstTMSWeightings!TMSWeighting * glTMSAltAsstStuWtg
                                rstTMSWeightings.Update
                            Else
                                If sPrevStuOrAsst = "ASST" Then
                                    rstTMSWeightings.Edit
                                    rstTMSWeightings!TMSAsstWeighting = rstTMSWeightings!TMSAsstWeighting * glTMSAltAsstStuWtg
                                    rstTMSWeightings.Update
                                End If
                            End If
                            
                        End If

                End If

                .MoveNext
            Loop


            End With


        End If
    End If
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Function GetTMSWeighting(WeekBeingProcessed As Long, TalkNo As String, AssignmentDate As Date) As Double
On Error GoTo ErrorTrap
Dim bNewArr As MidweekMtgVersion

    Select Case Abs(WeekBeingProcessed)
    Case Is = 0
        GetTMSWeighting = 3.3E+100
    Case Is <= 3
        GetTMSWeighting = 3.3E+50
    Case Is = 4
        GetTMSWeighting = 260000000000#
    Case Is = 5
        GetTMSWeighting = 20000000000#
    Case Is = 6
        GetTMSWeighting = 150000000
    Case Is = 7
        GetTMSWeighting = 11000000
    Case Is = 8
        GetTMSWeighting = 800000
    Case Is = 9
        GetTMSWeighting = 600000
    Case Is = 10
        GetTMSWeighting = 400000
    Case Is = 11
        GetTMSWeighting = 150000
    Case Is = 12
        GetTMSWeighting = 90000
    Case Is = 13
        GetTMSWeighting = 50000
    Case Is = 14
        GetTMSWeighting = 20000
    Case Is = 15
        GetTMSWeighting = 9000
    Case Is < 20
        GetTMSWeighting = 5000
    Case Is < 30
        GetTMSWeighting = 1000
    Case Is < 40
        GetTMSWeighting = 100
    Case Else
        GetTMSWeighting = 1
    End Select
    
     
     
     
    '
    'Weighting coeffs are acquired from GlobalParms object when frmTMSScheduling loads
    '
    
    If WeekBeingProcessed = 0 Then
        GetTMSWeighting = 1000 * GetTMSWeighting
    Else
        Select Case TalkNo
        Case "BS", "IC", "RV", "O", "Asst", "BR"
            Select Case TalkNo
            Case "BR"
                GetTMSWeighting = TMSBibleReadingWeighting_2016 * GetTMSWeighting
            Case "IC"
                GetTMSWeighting = TMSInitialCallWeighting_2016 * GetTMSWeighting
            Case "RV"
                GetTMSWeighting = TMSReturnVisitWeighting_2016 * GetTMSWeighting
            Case "BS"
                GetTMSWeighting = TMSBibleStudyWeighting_2016 * GetTMSWeighting
            Case "O"
                GetTMSWeighting = TMSOtherWeighting_2016 * GetTMSWeighting
            Case "Asst"
                GetTMSWeighting = TMSAsstWeighting * GetTMSWeighting
            End Select
        Case "P", "S", "B", "4", "R", "MR"
            Select Case TalkNo
            Case "P"
                GetTMSWeighting = TMSPrayerWeighting * GetTMSWeighting
            Case "S"
                GetTMSWeighting = TMSSQWeighting * GetTMSWeighting
            Case "B"
                GetTMSWeighting = TMSBHWeighting * GetTMSWeighting
            Case "4"
                GetTMSWeighting = TMSNo4Weighting * GetTMSWeighting
            Case "R", "MR"
                GetTMSWeighting = TMSReviewReaderWeighting * GetTMSWeighting
            End Select
        Case "1", "2", "3"
            bNewArr = NewMtgArrangementStarted(CStr(AssignmentDate))
            Select Case TalkNo
            Case "1"
                If bNewArr = TMS2009 Then
                    GetTMSWeighting = TMSNo1Weighting_2009 * GetTMSWeighting
                Else
                    GetTMSWeighting = TMSNo1Weighting * GetTMSWeighting
                End If
            Case "2"
                If bNewArr = TMS2009 Then
                    GetTMSWeighting = TMSNo2Weighting_2009 * GetTMSWeighting
                Else
                    GetTMSWeighting = TMSNo2Weighting * GetTMSWeighting
                End If
            Case "3"
                If bNewArr = TMS2009 Then
                    GetTMSWeighting = TMSNo3Weighting_2009 * GetTMSWeighting
                Else
                    GetTMSWeighting = TMSNo3Weighting * GetTMSWeighting
                End If
            End Select
        End Select
    End If
    


    Exit Function
ErrorTrap:
    EndProgram
End Function

Public Function GetTMSAsstListWeighting(WeekBeingProcessed As Long, TalkNo As String, AssignmentDate As Date) As Double
On Error GoTo ErrorTrap
Dim bNewArr As MidweekMtgVersion

    Select Case Abs(WeekBeingProcessed)
    Case Is = 0
        GetTMSAsstListWeighting = 5E+100
    Case Is <= 3
        GetTMSAsstListWeighting = 5E+50
    Case Is < 6
        GetTMSAsstListWeighting = 100000000000#
    Case Is = 6
        GetTMSAsstListWeighting = 1000000000
    Case Is = 7
        GetTMSAsstListWeighting = 100000000
    Case Is = 8
        GetTMSAsstListWeighting = 10000000
    Case Is = 9
        GetTMSAsstListWeighting = 1000000
    Case Is = 10
        GetTMSAsstListWeighting = 100000
    Case Is = 11
        GetTMSAsstListWeighting = 100000
    Case Is = 12
        GetTMSAsstListWeighting = 10000
    Case Is = 13
        GetTMSAsstListWeighting = 1000
    Case Is = 14
        GetTMSAsstListWeighting = 1000
    Case Is = 15
        GetTMSAsstListWeighting = 1000
    Case Is = 16
        GetTMSAsstListWeighting = 10000
    Case Is = 17
        GetTMSAsstListWeighting = 100000
    Case Is = 18
        GetTMSAsstListWeighting = 1000000
    Case Is = 19
        GetTMSAsstListWeighting = 10000000
    Case Is = 20
        GetTMSAsstListWeighting = 100000000
    Case Else
        GetTMSAsstListWeighting = 1000000000000#
    End Select
'
'    Select Case Abs(WeekBeingProcessed)
'    Case Is = 0
'        GetTMSAsstListWeighting = 100000000000#
'    Case Is < 6
'        GetTMSAsstListWeighting = 1000000
'    Case Is = 6
'        GetTMSAsstListWeighting = 100000
'    Case Is = 7
'        GetTMSAsstListWeighting = 1000
'    Case Is = 8
'        GetTMSAsstListWeighting = 100
'    Case Is = 9
'        GetTMSAsstListWeighting = 1000
'    Case Is = 10
'        GetTMSAsstListWeighting = 100000
'    Case Is = 11
'        GetTMSAsstListWeighting = 1000000
'    Case Is > 11
'        GetTMSAsstListWeighting = 1000000
'    Case Else
'        GetTMSAsstListWeighting = 1000000
'    End Select
    
'   *** 14/12/2010 ***
'    Select Case Abs(WeekBeingProcessed)
'    Case Is < 6
'        GetTMSAsstListWeighting = 1000000
'    Case Is < 7
'        GetTMSAsstListWeighting = 750000
'    Case Is < 8
'        GetTMSAsstListWeighting = 400000
'    Case Is < 9
'        GetTMSAsstListWeighting = 100000
'    Case Is < 10
'        GetTMSAsstListWeighting = 10000
'    Case Is < 12
'        GetTMSAsstListWeighting = 100
'    Case Is < 15
'        GetTMSAsstListWeighting = 100
'    Case Is < 18
'        GetTMSAsstListWeighting = 10000
'    Case Is < 20
'        GetTMSAsstListWeighting = 100000
'    Case Is < 30
'        GetTMSAsstListWeighting = 1000000
'    Case Else
'        GetTMSAsstListWeighting = 1000000
'    End Select
     
    '
    'Weighting coeffs are acquired from GlobalParms object when frmTMSScheduling loads
    '
    
    If WeekBeingProcessed = 0 Then
        GetTMSAsstListWeighting = 1000 * GetTMSAsstListWeighting
    Else
        Select Case TalkNo
        Case "BS", "IC", "RV", "O", "Asst", "BR"
            Case "BR"
                GetTMSAsstListWeighting = TMSBibleReadingWeighting_2016 * GetTMSAsstListWeighting
            Case "IC"
                GetTMSAsstListWeighting = TMSInitialCallWeighting_2016 * GetTMSAsstListWeighting
            Case "RV"
                GetTMSAsstListWeighting = TMSReturnVisitWeighting_2016 * GetTMSAsstListWeighting
            Case "BS"
                GetTMSAsstListWeighting = TMSBibleStudyWeighting_2016 * GetTMSAsstListWeighting
            Case "O"
                GetTMSAsstListWeighting = TMSOtherWeighting_2016 * GetTMSAsstListWeighting
            Case "Asst"
                GetTMSAsstListWeighting = TMSAsstWeighting * GetTMSAsstListWeighting
        Case "P", "S", "B", "4"
            Select Case TalkNo
            Case "P"
                GetTMSAsstListWeighting = TMSPrayerWeighting * GetTMSAsstListWeighting
            Case "S"
                GetTMSAsstListWeighting = TMSSQWeighting * GetTMSAsstListWeighting
            Case "B"
                GetTMSAsstListWeighting = TMSBHWeighting * GetTMSAsstListWeighting
            Case "4"
                GetTMSAsstListWeighting = TMSNo4Weighting * GetTMSAsstListWeighting
            End Select
        Case "1", "2", "3"
            bNewArr = NewMtgArrangementStarted(CStr(AssignmentDate))
            Select Case TalkNo
            Case "1"
                If bNewArr = TMS2009 Then
                    GetTMSAsstListWeighting = TMSNo1Weighting_2009 * GetTMSAsstListWeighting
                Else
                    GetTMSAsstListWeighting = TMSNo1Weighting * GetTMSAsstListWeighting
                End If
            Case "2"
                If bNewArr = TMS2009 Then
                    GetTMSAsstListWeighting = TMSNo2Weighting_2009 * GetTMSAsstListWeighting * 3
                Else
                    GetTMSAsstListWeighting = TMSNo2Weighting * GetTMSAsstListWeighting * 3
                End If
            Case "3"
                If bNewArr = TMS2009 Then
                    GetTMSAsstListWeighting = TMSNo3Weighting_2009 * GetTMSAsstListWeighting * 3
                Else
                    GetTMSAsstListWeighting = TMSNo3Weighting * GetTMSAsstListWeighting * 3
                End If
            End Select
        End Select
    End If
    


    Exit Function
ErrorTrap:
    EndProgram
End Function



Public Function GetTMSCounselHistoryRecordset(TheStudentID As Long) As Recordset
Dim SQLStr As String
'
'

On Error GoTo ErrorTrap

    SQLStr = "SELECT AssignmentDate, " & _
            "TalkNo, " & _
            "CounselPoint, " & _
            "CounselPointAssignedDate, " & _
            "CounselPointCompletedDate, " & _
            "Comment, " & _
            "ScheduleSeqNum, " & _
            "SchoolNo, " & _
            "Setting, " & _
            "PersonID, " & _
            "Assistant1ID, " & _
            "TalkCompleted, " & _
            "TalkDefaulted, " & _
            "IsVolunteer, " & _
            "ExerciseComplete, " & _
            "DiscussedWithStudent, " & _
            "ItemsSeqNum " & _
            "FROM tblTMSSchedule " & _
            "WHERE PersonId = " & TheStudentID & _
            " AND  TalkNo NOT IN ('P','R','MR') " & _
            " ORDER BY AssignmentDate DESC"
            
            '" AND SchoolNo <= " & GlobalParms.GetValue("TMSNoSchoolsForCounsel", "NumVal", 1) & _

    Set GetTMSCounselHistoryRecordset = CMSDB.OpenRecordset(SQLStr, dbOpenDynaset)

    Exit Function
ErrorTrap:
    EndProgram
    
End Function


Public Sub FillTheTMSCounselHistoryGrid(CounselHistory As Recordset, TopRow As Long)
Dim i As Byte, j As Integer

On Error GoTo ErrorTrap

'
'Populate grid from recordset
'
     'clear grid's non-fixed rows
    frmTMSCounselPoints!flxCounselHistory.Rows = 1
    
    With CounselHistory
        
    If Not .BOF Then
        .MoveFirst
        j = 1
        Do Until .EOF
            frmTMSCounselPoints!flxCounselHistory.Rows = j + 1
            frmTMSCounselPoints!flxCounselHistory.TextMatrix(j, 0) = !AssignmentDate
            frmTMSCounselPoints!flxCounselHistory.TextMatrix(j, 1) = !TalkNo
            
            If !CounselPoint > 0 Then
                Select Case NewMtgArrangementStarted(!AssignmentDate)
                Case CLM2016
                    frmTMSCounselPoints!flxCounselHistory.TextMatrix(j, 2) = !CounselPoint
                Case TMS2009
                    Select Case !TalkNo
                    Case "B"
                        If CongregationMember.TMS_AllowCounselOnNo1AndBH(!PersonID) Then
                            frmTMSCounselPoints!flxCounselHistory.TextMatrix(j, 2) = !CounselPoint
                        Else
                            frmTMSCounselPoints!flxCounselHistory.TextMatrix(j, 2) = ""
                        End If
                    Case "3" And (CongregationMember.ElderDate(!PersonID) > 0 Or _
                                  CongregationMember.ServantDate(!PersonID) > 0)
                        If CongregationMember.TMS_AllowCounselOnNo1AndBH(!PersonID) Then
                            frmTMSCounselPoints!flxCounselHistory.TextMatrix(j, 2) = !CounselPoint
                        Else
                            frmTMSCounselPoints!flxCounselHistory.TextMatrix(j, 2) = ""
                        End If
                    Case Else
                        frmTMSCounselPoints!flxCounselHistory.TextMatrix(j, 2) = !CounselPoint
                    End Select
                Case Else
                    Select Case !TalkNo
                    Case "1", "B"
                        If CongregationMember.TMS_AllowCounselOnNo1AndBH(!PersonID) Then
                            frmTMSCounselPoints!flxCounselHistory.TextMatrix(j, 2) = !CounselPoint
                        Else
                            frmTMSCounselPoints!flxCounselHistory.TextMatrix(j, 2) = ""
                        End If
                    Case Else
                        frmTMSCounselPoints!flxCounselHistory.TextMatrix(j, 2) = !CounselPoint
                    End Select
                End Select
            Else
                frmTMSCounselPoints!flxCounselHistory.TextMatrix(j, 2) = ""
            End If
            
            frmTMSCounselPoints!flxCounselHistory.TextMatrix(j, 3) = j - 1
            j = j + 1
            .MoveNext
        Loop
    End If
        
    On Error Resume Next
    If TopRow > 0 Then
        If TopRow < frmTMSCounselPoints!flxCounselHistory.Rows Then
            frmTMSCounselPoints!flxCounselHistory.TopRow = TopRow
        End If
    End If
    On Error GoTo ErrorTrap
           
    
    End With
    
    
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Public Function LoadCMSFileToDB(ScheduleFilePath As String, FileType As Long) As Boolean
Dim FreeFileNum As Integer, TheScheduleFileRecord As String, TheRecordCount As Long
Dim InputRecCollection As New Collection, i As Long, AString As String, TheDate As Date
Dim FileIsOpen As Boolean

On Error GoTo ErrorTrap

    LoadCMSFileToDB = False

    FreeFileNum = FreeFile() 'Get next available file number
    
    TheTMSItemFilePath = ScheduleFilePath
    
    On Error Resume Next
    
    Open ScheduleFilePath For Input As #FreeFileNum
    
    If Err.number <> 0 Then
        MsgBox "Could not open " & ScheduleFilePath & ".", vbOKOnly + vbCritical, AppName
        End
    End If
    
    FileIsOpen = True
    On Error GoTo ErrorTrap
    
'    Line Input #FreeFileNum, TheScheduleFileRecord
    TheRecordCount = 1
    
    For i = 1 To InputRecCollection.Count
        InputRecCollection.Remove i
    Next i
    
    TheDate = CDate(frmTMSLoadItems!txtFirstDate)
    
    mbBroOnlyTalkFound = False 'init
    
    '
    'Read through CMS Schedule file
    '
    Do While (Not EOF(FreeFileNum))
        Line Input #FreeFileNum, TheScheduleFileRecord
        '
        'Add the new record to collection for later display if necessary
        '
        InputRecCollection.Add TheScheduleFileRecord
        
        Select Case FileType
        Case 1 'tab delimited
            If ProcessTabDelimitedRow(TheScheduleFileRecord, TheDate) Then
                TheRecordCount = TheRecordCount + 1
            Else
                Screen.MousePointer = vbNormal
                MsgBox "Error on line " & TheRecordCount & " of " & ScheduleFilePath & _
                        vbCrLf & "Please correct and try again. The following dialogue will show the erroneous data.", vbOKOnly + vbExclamation, AppName
                For i = 1 To InputRecCollection.Count
                    AString = AString & InputRecCollection.Item(i) & vbNewLine
                Next i
                TMSWeeksItems = AString
                frmTMSShowErroneousItems.Show vbModal
                Exit Function
            End If
        Case 2 'normal format
            If CheckRecordTypeAndActAccordingly(TheScheduleFileRecord, TheDate) Then
                TheRecordCount = TheRecordCount + 1
            Else
                Screen.MousePointer = vbNormal
                MsgBox "Error on line " & TheRecordCount & " of " & ScheduleFilePath & _
                        vbCrLf & "Please correct and try again. The following dialogue will show the erroneous data.", vbOKOnly + vbExclamation, AppName
                For i = 1 To InputRecCollection.Count
                    AString = AString & InputRecCollection.Item(i) & vbNewLine
                Next i
                TMSWeeksItems = AString
                frmTMSShowErroneousItems.Show vbModal
                Exit Function
            End If
        End Select
    Loop
        
    Close FreeFileNum
    
    Dim s As String
    If GlobalParms.GetValue("TMSLoadNo3BroOnlyWarning", "TrueFalse", False) Then
        s = vbCrLf & vbCrLf & _
            IIf(mbBroOnlyTalkFound, "", "No #3 assignments marked as 'Brother Only' " & _
                "were found. You may need to mark these as such manually.")
    End If
    
    MsgBox "Items successfully loaded. " & s, vbOKOnly + vbInformation, AppName
    
            
    LoadCMSFileToDB = True

    Exit Function
    
ErrorTrap:
    
    If FileIsOpen Then
        Close #FreeFileNum
    End If
    
    EndProgram
    

End Function
Private Function CheckRecordTypeAndActAccordingly(InputRecord As String, CurrentAssignmentDate As Date) As Boolean

    If Not NewMtgArrangementStarted(CStr(CurrentAssignmentDate)) Then
        CheckRecordTypeAndActAccordingly = CheckRecordTypeAndActAccordingly_2002(InputRecord, CurrentAssignmentDate)
    Else
        CheckRecordTypeAndActAccordingly = CheckRecordTypeAndActAccordingly_2009(InputRecord, CurrentAssignmentDate)
    End If

End Function

Private Function CheckRecordTypeAndActAccordingly_2002(InputRecord As String, CurrentAssignmentDate As Date) As Boolean
On Error GoTo ErrorTrap

Static OneWeeksTMSSchedule As TMSScheduleRecord
Dim TMSScheduleDay As Long, DateRecDetails() As String, TMSScheduleMonth As Long
Dim i As Long
Static DateRecFound As Boolean
Static SQRecFound As Boolean
Static No1RecFound As Boolean
Static No2RecFound As Boolean
Static No3RecFound As Boolean
Static No4RecFound As Boolean
Static OralReviewRecFound As Boolean

    DoEvents
    '
    'Get rid of tabs and extraneous spaces. Double-up the quotes.
    '
    InputRecord = Trim(Replace(InputRecord, vbTab, " "))
    InputRecord = RemoveExtraSpacesToLeaveSingleSpacedWords(InputRecord)
    InputRecord = DoubleUpSingleQuotes(InputRecord)
    
    CheckRecordTypeAndActAccordingly_2002 = True
    
    With OneWeeksTMSSchedule
    
    Select Case True
    Case IsDateRecord(InputRecord)
        '
        'Split out each word to an element of the array
        '
        DateRecDetails() = Split(InputRecord)
        If UBound(DateRecDetails) >= 7 Then 'at least 7 words - maybe more for Bible books
                                            ' such as "Song of Solomon 1-3"
            '
            'Construct the date
            '
            TMSScheduleMonth = GetMonthNumber(Left(DateRecDetails(0), 3))
            TMSScheduleDay = CLng(DateRecDetails(1))
            .TMSUDT_AssignmentDate = CDate(DateSerial(CInt(Right(frmTMSLoadItems.txtFirstDate, 4)), _
                                        CInt(TMSScheduleMonth), CInt(TMSScheduleDay)))
            
            '
            'Check if the song is final part of the line
            '
            If DateRecDetails(UBound(DateRecDetails) - 1) = "Song" Then
                .TMSUDT_SongNo = DateRecDetails(UBound(DateRecDetails))
                '
                'Build the BH Source - depends how long it is. Relies on fact that last two
                ' words on the record are "Song nnn"
                '
                For i = 4 To UBound(DateRecDetails) - 2
                    .TMSUDT_BHSource = .TMSUDT_BHSource & DateRecDetails(i) & " "
                Next i
            ElseIf DateRecDetails(2) = "Song" Then
                .TMSUDT_SongNo = DateRecDetails(3)
                '
                'Build the BH Source - depends how long it is. Relies on fact that record
                ' ends with "Bible Reading: <Scripture ref>"
                '
                For i = 6 To UBound(DateRecDetails)
                    .TMSUDT_BHSource = .TMSUDT_BHSource & DateRecDetails(i) & " "
                Next i
            Else
                CheckRecordTypeAndActAccordingly_2002 = False
                Exit Function
            End If
            
            'BH source should now be successfully constructed...
            .TMSUDT_BHSource = Trim(.TMSUDT_BHSource)
            
                                        
            '
            'Compare file's date with that maintained internally
            '
            If .TMSUDT_AssignmentDate <> CurrentAssignmentDate Then
                CheckRecordTypeAndActAccordingly_2002 = False
                Exit Function
            End If
            
            DateRecFound = True
        Else
            CheckRecordTypeAndActAccordingly_2002 = False
            Exit Function
        End If
    Case Left(InputRecord, 14) = "Speech Quality"
        .TMSUDT_SQSource = AcquireTheSourceMaterial(InputRecord, "(", ")", False)
        .TMSUDT_SQTheme = AcquireTheTheme(InputRecord, "Speech Quality", "(")
        SQRecFound = True
    Case Left(InputRecord, 11) = "Oral Review" Or Left(InputRecord, 33) = "Theocratic Ministry School Review"
        .TMSUDT_OralReview = "Oral Review"
        OralReviewRecFound = True
        If AllRecsComplete(DateRecFound, _
                            SQRecFound, _
                            No1RecFound, _
                            No2RecFound, _
                            No3RecFound, _
                            No4RecFound, _
                            OralReviewRecFound) Then
            CheckRecordTypeAndActAccordingly_2002 = True
            UpdateDB OneWeeksTMSSchedule, OralReviewRecFound
            '
            'Initialise structure for next logical group...
            '
            OneWeeksTMSSchedule = InitSchedRec(0, "", 0, "", "", "", "", "", "", "", "", "", "", False, False, "")
            
            DateRecFound = False
            SQRecFound = False
            No1RecFound = False
            No2RecFound = False
            No3RecFound = False
            No4RecFound = False
            OralReviewRecFound = False
            
            CurrentAssignmentDate = CurrentAssignmentDate + 7
        Else
            CheckRecordTypeAndActAccordingly_2002 = False
        End If
    Case Left(InputRecord, 5) = "No. 1"
        .TMSUDT_No1Source = AcquireTheSourceMaterial(InputRecord, "(", ")", False)
        .TMSUDT_No1Theme = AcquireTheTheme(InputRecord, "No. 1", "(")
        No1RecFound = True
    Case Left(InputRecord, 5) = "No. 2"
        .TMSUDT_No2Source = AcquireTheSourceMaterial(InputRecord, "No. 2", "", False)
        .TMSUDT_No2Theme = "Reading Assignment"
        No2RecFound = True
    Case Left(InputRecord, 5) = "No. 3"
        .TMSUDT_No3Source = AcquireTheSourceMaterial(InputRecord, "(", ")", False)
        .TMSUDT_No3Theme = AcquireTheTheme(InputRecord, "No. 3", "(")
         No3RecFound = True
    Case Left(InputRecord, 5) = "No. 4"
        .TMSUDT_No4Source = AcquireTheSourceMaterial(InputRecord, "(", ")", False)
        .TMSUDT_No4Theme = AcquireTheTheme(InputRecord, "No. 4", "(")
        
        If Left(.TMSUDT_No4Theme, 1) = "*" Then
            .TMSUDT_BroOnlyForNo4 = True
            mbBroOnlyTalkFound = True
        Else
            .TMSUDT_BroOnlyForNo4 = False
        End If
        
        No4RecFound = True
        If AllRecsComplete(DateRecFound, _
                            SQRecFound, _
                            No1RecFound, _
                            No2RecFound, _
                            No3RecFound, _
                            No4RecFound, _
                            OralReviewRecFound) Then
            CheckRecordTypeAndActAccordingly_2002 = True
            UpdateDB OneWeeksTMSSchedule, OralReviewRecFound
            '
            'Initialise structure for next logical group...
            '
            OneWeeksTMSSchedule = InitSchedRec(0, "", 0, "", "", "", "", "", "", "", "", "", "", False, False, "")
            
            DateRecFound = False
            SQRecFound = False
            No1RecFound = False
            No2RecFound = False
            No3RecFound = False
            No4RecFound = False
            OralReviewRecFound = False
            
            CurrentAssignmentDate = CurrentAssignmentDate + 7
            
        Else
            CheckRecordTypeAndActAccordingly_2002 = False
        End If
    Case InputRecord = vbCr Or InputRecord = "" Or InputRecord = vbLf
         'Do nothing
    Case Left(InputRecord, 10) = "theocratic" Or Left(InputRecord, 10) = "Theocratic" Or _
         Left(InputRecord, 10) = "THEOCRATIC"
         'Do nothing
    Case Else
        CheckRecordTypeAndActAccordingly_2002 = False
        Exit Function
    End Select
    
    End With

    Exit Function
    
ErrorTrap:
    EndProgram
    
End Function

Private Function CheckRecordTypeAndActAccordingly_2009(InputRecord As String, CurrentAssignmentDate As Date) As Boolean
On Error GoTo ErrorTrap

Static OneWeeksTMSSchedule As TMSScheduleRecord
Dim TMSScheduleDay As Long, DateRecDetails() As String, TMSScheduleMonth As Long
Dim i As Long
Static DateRecFound As Boolean
Static SQRecFound As Boolean
Static No1RecFound As Boolean
Static No2RecFound As Boolean
Static No3RecFound As Boolean
Static No4RecFound As Boolean
Static OralReviewRecFound As Boolean

    DoEvents
    '
    'Get rid of tabs and extraneous spaces. Double-up the quotes.
    '
    InputRecord = Trim(Replace(InputRecord, vbTab, " "))
    InputRecord = RemoveExtraSpacesToLeaveSingleSpacedWords(InputRecord)
    InputRecord = DoubleUpSingleQuotes(InputRecord)
    
    CheckRecordTypeAndActAccordingly_2009 = True
    
    With OneWeeksTMSSchedule
    
    Select Case True
    Case IsDateRecord(InputRecord)
        '
        'Split out each word to an element of the array
        '
        DateRecDetails() = Split(InputRecord)
        If UBound(DateRecDetails) >= 5 Then 'at least 6 words - maybe more for Bible books
                                            ' such as "Song of Solomon 1-3"
            '
            'Construct the date
            '
            TMSScheduleMonth = GetMonthNumber(Left(DateRecDetails(0), 3))
            TMSScheduleDay = CLng(DateRecDetails(1))
            .TMSUDT_AssignmentDate = CDate(DateSerial(CInt(Right(frmTMSLoadItems.txtFirstDate, 4)), _
                                        CInt(TMSScheduleMonth), CInt(TMSScheduleDay)))
            
            '
            'Get BH Source
            '
            For i = 4 To UBound(DateRecDetails)
                .TMSUDT_BHSource = .TMSUDT_BHSource & DateRecDetails(i) & " "
            Next i
            
            
            'BH source should now be successfully constructed...
            .TMSUDT_BHSource = Trim(.TMSUDT_BHSource)
            
            'no song supplied in 2009+ schedules
            .TMSUDT_SongNo = 0
                                        
            '
            'Compare file's date with that maintained internally
            '
            If .TMSUDT_AssignmentDate <> CurrentAssignmentDate Then
                CheckRecordTypeAndActAccordingly_2009 = False
                Exit Function
            End If
            
            DateRecFound = True
        Else
            CheckRecordTypeAndActAccordingly_2009 = False
            Exit Function
        End If
    Case Left(InputRecord, 11) = "Oral Review" Or Left(InputRecord, 33) = "Theocratic Ministry School Review"
        .TMSUDT_OralReview = "Oral Review"
        OralReviewRecFound = True
        If AllRecsComplete_2009(DateRecFound, _
                                No1RecFound, _
                                No2RecFound, _
                                No3RecFound, _
                                OralReviewRecFound) Then
            CheckRecordTypeAndActAccordingly_2009 = True
            UpdateDB_2009 OneWeeksTMSSchedule, OralReviewRecFound
            '
            'Initialise structure for next logical group...
            '
            OneWeeksTMSSchedule = InitSchedRec(0, "", 0, "", "", "", "", "", "", "", "", "", "", False, False, "")
            
            DateRecFound = False
            No1RecFound = False
            No2RecFound = False
            No3RecFound = False
            OralReviewRecFound = False
            
            CurrentAssignmentDate = CurrentAssignmentDate + 7
        Else
            CheckRecordTypeAndActAccordingly_2009 = False
        End If
    Case Left(InputRecord, 5) = "No. 1"
        .TMSUDT_No1Source = AcquireTheSourceMaterial(InputRecord, "No. 1", "", True)
        .TMSUDT_No1Theme = "Reading Assignment"
        No1RecFound = True
    Case Left(InputRecord, 5) = "No. 2"
        .TMSUDT_No2Source = AcquireTheSourceMaterial(InputRecord, "(", ")", True)
        .TMSUDT_No2Theme = AcquireTheTheme(InputRecord, "No. 2", "(")
        No2RecFound = True
    Case Left(InputRecord, 5) = "No. 3"
        .TMSUDT_No3Source = AcquireTheSourceMaterial(InputRecord, "(", ")", True)
        .TMSUDT_No3Theme = AcquireTheTheme(InputRecord, "No. 3", "(")
        No3RecFound = True
        
        If Left(.TMSUDT_No3Theme, 1) = "*" And Left(.TMSUDT_No3Theme, 2) <> "**" Then
            .TMSUDT_BroOnlyForNo3 = True
            mbBroOnlyTalkFound = True
        Else
            .TMSUDT_BroOnlyForNo3 = False
        End If
        
        If AllRecsComplete_2009(DateRecFound, _
                                No1RecFound, _
                                No2RecFound, _
                                No3RecFound, _
                                OralReviewRecFound) Then
            CheckRecordTypeAndActAccordingly_2009 = True
            UpdateDB_2009 OneWeeksTMSSchedule, OralReviewRecFound
            '
            'Initialise structure for next logical group...
            '
            OneWeeksTMSSchedule = InitSchedRec(0, "", 0, "", "", "", "", "", "", "", "", "", "", False, False, "")
            
            DateRecFound = False
            No1RecFound = False
            No2RecFound = False
            No3RecFound = False
            OralReviewRecFound = False
            
            CurrentAssignmentDate = CurrentAssignmentDate + 7
            
        Else
            CheckRecordTypeAndActAccordingly_2009 = False
        End If
        
    Case InputRecord = vbCr Or InputRecord = "" Or InputRecord = vbLf
         'Do nothing
    Case Left(InputRecord, 10) = "theocratic" Or Left(InputRecord, 10) = "Theocratic" Or _
         Left(InputRecord, 10) = "THEOCRATIC"
         'Do nothing
    Case Else
        CheckRecordTypeAndActAccordingly_2009 = False
        Exit Function
    End Select
    
    End With

    Exit Function
    
ErrorTrap:
    EndProgram
    
End Function


Private Function ProcessTabDelimitedRow(InputRecord As String, _
                                        CurrentAssignmentDate As Date) As Boolean
                                        
    If Not NewMtgArrangementStarted(CStr(CurrentAssignmentDate)) Then
        ProcessTabDelimitedRow = ProcessTabDelimitedRow_2002(InputRecord, CurrentAssignmentDate)
    Else
        ProcessTabDelimitedRow = ProcessTabDelimitedRow_2009(InputRecord, CurrentAssignmentDate)
    End If
                                        
End Function
Private Function ProcessTabDelimitedRow_2002(InputRecord As String, _
                                                CurrentAssignmentDate As Date) As Boolean
On Error GoTo ErrorTrap

Static OneWeeksTMSSchedule As TMSScheduleRecord
Dim RowDetails() As String, TempStrArray() As String
Dim i As Long
Static OralReviewRecFound As Boolean

    DoEvents
    '
    'Double-up the quotes.
    '
    InputRecord = DoubleUpSingleQuotes(InputRecord)
    
    ProcessTabDelimitedRow_2002 = True
    
    '
    'Split out each chunk to an element of the array
    '
    RowDetails() = Split(InputRecord, vbTab)
    
    If UBound(RowDetails) <> 11 Then 'wrong no of columns
        ProcessTabDelimitedRow_2002 = False
        Exit Function
    End If
    
    'remove leading/trailing spaces from each field and check...
    For i = 0 To 11
        RowDetails(i) = Trim$(RowDetails(i))
        If RowDetails(i) = "" Then
            ProcessTabDelimitedRow_2002 = False
            Exit Function
        End If
    Next i
    
    If UCase$(RowDetails(0)) = "DATE" Then 'file header
        Exit Function
    End If
    
    With OneWeeksTMSSchedule
    
    'Assignment date
    RowDetails(0) = Replace(RowDetails(0), ".", "/")
    RowDetails(0) = Replace(RowDetails(0), ":", "/")
    TempStrArray() = Split(RowDetails(0), "/")
    If UBound(TempStrArray) <> 2 Then 'dodgy date
        ProcessTabDelimitedRow_2002 = False
        Exit Function
    End If
    For i = 0 To 2
        If Not IsNumber(TempStrArray(i), False, False, False) Then
            ProcessTabDelimitedRow_2002 = False
            Exit Function
        End If
    Next i
    If Not ValidDate(RowDetails(0)) Then
        ProcessTabDelimitedRow_2002 = False
        Exit Function
    End If
    
    .TMSUDT_AssignmentDate = CDate(Format$(RowDetails(0), "dd/mm/yyyy"))
    
    If .TMSUDT_AssignmentDate <> CurrentAssignmentDate Then
        ProcessTabDelimitedRow_2002 = False
        Exit Function
    End If
    
    'Bible Highlights
    .TMSUDT_BHSource = RowDetails(1)
    
    'Song
    .TMSUDT_SongNo = RowDetails(2)
    If Not IsNumber(RowDetails(2), False, False, False) Then
        ProcessTabDelimitedRow_2002 = False
        Exit Function
    End If
    
    
    'SQ
    .TMSUDT_SQTheme = RowDetails(3)
    .TMSUDT_SQSource = Replace$(RowDetails(4), Chr$(167), "par")
    .TMSUDT_SQSource = Replace$(RowDetails(4), Chr$(182), "par")
    
    '#1
    .TMSUDT_No1Theme = RowDetails(5)
    .TMSUDT_No1Source = Replace$(RowDetails(6), Chr$(167), "par")
    .TMSUDT_No1Source = Replace$(RowDetails(6), Chr$(182), "par")
    
    If UCase$(RowDetails(5)) = "ORAL REVIEW" Then
        OralReviewRecFound = True
        .TMSUDT_OralReview = "Oral Review"
    Else
        OralReviewRecFound = False
    End If
    
    '#2
    .TMSUDT_No2Theme = "Reading Assignment"
    .TMSUDT_No2Source = RowDetails(7)
    
    '#3
    .TMSUDT_No3Theme = RowDetails(8)
    .TMSUDT_No3Source = Replace$(RowDetails(9), Chr$(167), "par")
    .TMSUDT_No3Source = Replace$(RowDetails(9), Chr$(182), "par")
    
    '#4
    .TMSUDT_No4Theme = RowDetails(10)
    .TMSUDT_No4Source = Replace$(RowDetails(11), Chr$(167), "par")
    .TMSUDT_No4Source = Replace$(RowDetails(11), Chr$(182), "par")
        
    If Left(.TMSUDT_No4Theme, 1) = "*" Then
        .TMSUDT_BroOnlyForNo4 = True
        mbBroOnlyTalkFound = True
    Else
        .TMSUDT_BroOnlyForNo4 = False
    End If
        
    'Save record
    UpdateDB OneWeeksTMSSchedule, OralReviewRecFound
    
    '
    'Initialise structure for next logical group...
    '
    OneWeeksTMSSchedule = InitSchedRec(0, "", 0, "", "", "", "", "", "", "", "", "", "", False, False, "")
        
    OralReviewRecFound = False
    
    CurrentAssignmentDate = CurrentAssignmentDate + 7
    
        
    End With

    Exit Function
    
ErrorTrap:
    EndProgram
    
End Function
Private Function ProcessTabDelimitedRow_2009(InputRecord As String, _
                                            CurrentAssignmentDate As Date) As Boolean
On Error GoTo ErrorTrap

Static OneWeeksTMSSchedule As TMSScheduleRecord
Dim RowDetails() As String, TempStrArray() As String
Dim i As Long
Static OralReviewRecFound As Boolean

    DoEvents
    '
    'Double-up the quotes.
    '
    InputRecord = DoubleUpSingleQuotes(InputRecord)
    
    ProcessTabDelimitedRow_2009 = True
    
    '
    'Split out each chunk to an element of the array
    '
    RowDetails() = Split(InputRecord, vbTab)
    
    If UBound(RowDetails) <> 11 Then 'wrong no of columns
        ProcessTabDelimitedRow_2009 = False
        Exit Function
    End If
    
    'remove leading/trailing spaces from each field and check...
    For i = 0 To 11
        RowDetails(i) = Trim$(RowDetails(i))
        If RowDetails(i) = "" Then
            ProcessTabDelimitedRow_2009 = False
            Exit Function
        End If
    Next i
    
    If UCase$(RowDetails(0)) = "DATE" Then 'file header
        Exit Function
    End If
    
    With OneWeeksTMSSchedule
    
    'Assignment date
    RowDetails(0) = Replace(RowDetails(0), ".", "/")
    RowDetails(0) = Replace(RowDetails(0), ":", "/")
    TempStrArray() = Split(RowDetails(0), "/")
    If UBound(TempStrArray) <> 2 Then 'dodgy date
        ProcessTabDelimitedRow_2009 = False
        Exit Function
    End If
    For i = 0 To 2
        If Not IsNumber(TempStrArray(i), False, False, False) Then
            ProcessTabDelimitedRow_2009 = False
            Exit Function
        End If
    Next i
    If Not ValidDate(RowDetails(0)) Then
        ProcessTabDelimitedRow_2009 = False
        Exit Function
    End If
    
    .TMSUDT_AssignmentDate = CDate(Format$(RowDetails(0), "dd/mm/yyyy"))
    
    If .TMSUDT_AssignmentDate <> CurrentAssignmentDate Then
        ProcessTabDelimitedRow_2009 = False
        Exit Function
    End If
    
    'Bible Highlights
    .TMSUDT_BHSource = RowDetails(1)
        
    If UCase$(RowDetails(2)) = "ORAL REVIEW" Then
        OralReviewRecFound = True
        .TMSUDT_OralReview = "Oral Review"
    Else
        OralReviewRecFound = False
    End If
    
    '#1
    .TMSUDT_No1Theme = "Reading Assignment"
    .TMSUDT_No1Source = RowDetails(2)
    
    '#2
    .TMSUDT_No2Theme = RowDetails(3)
    .TMSUDT_No2Source = Replace$(RowDetails(4), Chr$(167), "par")
    .TMSUDT_No2Source = Replace$(RowDetails(4), Chr$(182), "par")
    
    '#3
    .TMSUDT_No3Theme = RowDetails(5)
    .TMSUDT_No3Source = Replace$(RowDetails(6), Chr$(167), "par")
    .TMSUDT_No3Source = Replace$(RowDetails(6), Chr$(182), "par")
    
    If Left(.TMSUDT_No3Theme, 1) = "*" And Left(.TMSUDT_No3Theme, 2) <> "**" Then
        .TMSUDT_BroOnlyForNo3 = True
        mbBroOnlyTalkFound = True
    Else
        .TMSUDT_BroOnlyForNo3 = False
    End If
        
    'Save record
    UpdateDB_2009 OneWeeksTMSSchedule, OralReviewRecFound
    
    '
    'Initialise structure for next logical group...
    '
    OneWeeksTMSSchedule = InitSchedRec(0, "", 0, "", "", "", "", "", "", "", "", "", "", False, False, "")
        
    OralReviewRecFound = False
    
    CurrentAssignmentDate = CurrentAssignmentDate + 7
    
        
    End With

    Exit Function
    
ErrorTrap:
    EndProgram
    
End Function

Function InitSchedRec(TMSAssignmentDate As Date, _
                        TMSBHSource As String, _
                        TMSSongNo As Long, _
                        TMSSQTheme As String, _
                        TMSSQSource As String, _
                        TMSNo1Theme As String, _
                        TMSNo1Source As String, _
                        TMSNo2Theme As String, _
                        TMSNo2Source As String, _
                        TMSNo3Theme As String, _
                        TMSNo3Source As String, _
                        TMSNo4Theme As String, _
                        TMSNo4Source As String, _
                        TMSBroOnlyForNo3 As Boolean, _
                        TMSBroOnlyForNo4 As Boolean, _
                        TMSOralReview As String) As TMSScheduleRecord
On Error GoTo ErrorTrap

    With InitSchedRec
    .TMSUDT_AssignmentDate = TMSAssignmentDate
    .TMSUDT_BHSource = TMSBHSource
    .TMSUDT_No1Source = TMSNo1Source
    .TMSUDT_No1Theme = TMSNo1Theme
    .TMSUDT_No2Source = TMSNo2Source
    .TMSUDT_No2Theme = TMSNo2Theme
    .TMSUDT_No3Source = TMSNo3Source
    .TMSUDT_No3Theme = TMSNo3Theme
    .TMSUDT_No4Source = TMSNo4Source
    .TMSUDT_No4Theme = TMSNo4Theme
    .TMSUDT_SongNo = TMSSongNo
    .TMSUDT_SQTheme = TMSSQTheme
    .TMSUDT_SQSource = TMSSQSource
    .TMSUDT_OralReview = TMSOralReview
    .TMSUDT_BroOnlyForNo3 = TMSBroOnlyForNo3
    .TMSUDT_BroOnlyForNo4 = TMSBroOnlyForNo4
    End With

    Exit Function
    
ErrorTrap:
    EndProgram
    
End Function

Private Function IsDateRecord(TheString) As Boolean
On Error GoTo ErrorTrap

    Select Case LCase(Left(TheString, 3))
    Case "jan", _
        "feb", _
        "mar", _
        "apr", _
        "may", _
        "jun", _
        "jul", _
        "aug", _
        "sep", _
        "oct", _
        "nov", _
        "dec"
            IsDateRecord = True
    Case Else
        IsDateRecord = False
    End Select

    Exit Function
    
ErrorTrap:
    EndProgram
End Function

Private Sub AddRecordToDB(AssignmentDate As Date, TalkNo As String, TaskNo As Long, TalkSeqNum As Long, _
                            TalkTheme As String, SourceMaterial As String, Difficulty As Byte, BroOnly As Boolean)
On Error GoTo ErrorTrap

    CMSDB.Execute "INSERT INTO tblTMSItems " & _
                      "(AssignmentDate, TalkNo, " & _
                      "TaskNo, TalkSeqNum, " & _
                        "TalkTheme, " & _
                        "SourceMaterial, DifficultyRating0to5, " & _
                        "BrotherOnly) " & _
                      "VALUES (#" & Format(AssignmentDate, "mm/dd/yyyy") & "#, '" & TalkNo & "', " & TaskNo & ", " & TalkSeqNum & ", '" & _
                      TalkTheme & "', '" & SourceMaterial & "', " & Difficulty & ", " & _
                      BroOnly & ")"
    
    Exit Sub
    
ErrorTrap:
    EndProgram
End Sub


Private Function AcquireTheTheme(TheRecord As String, AfterString, BeforeString) As String
On Error GoTo ErrorTrap
Dim RightChunk As String, LeftChunk As String, PosOfThemeEnd As Long
    
    RightChunk = Right(TheRecord, Len(TheRecord) - Len(AfterString))
    
    If RightChunk = "" Then
        AcquireTheTheme = ""
        Exit Function
    End If
    
    PosOfThemeEnd = InStr(1, RightChunk, BeforeString)
    
    If PosOfThemeEnd = 0 Then
        'ie there's no source material
        AcquireTheTheme = RightChunk
    Else
        AcquireTheTheme = Left(RightChunk, PosOfThemeEnd - 1)
    End If
        
    
    If Left(AcquireTheTheme, 1) = ":" Then
        AcquireTheTheme = Right(AcquireTheTheme, Len(AcquireTheTheme) - 1)
    End If
    
    AcquireTheTheme = Trim(AcquireTheTheme)
    
    AcquireTheTheme = Replace(AcquireTheTheme, Chr(182), "par ")
    AcquireTheTheme = Replace(AcquireTheTheme, Chr(151), " - ") 'em dash
'    AcquireTheTheme = Replace(AcquireTheTheme, "-", " - ")
    
    
    Exit Function
    
ErrorTrap:
    EndProgram
End Function

Private Function AcquireTheSourceMaterial(TheRecord As String, AfterString, BeforeString, NewArrangement As Boolean) As String
On Error GoTo ErrorTrap
Dim RightChunk As String, LeftChunk As String, PosOfSourceStart As Long, sReadingTypeStr As String
    
    sReadingTypeStr = IIf(NewArrangement, "No. 1", "No. 2")
    
    If AfterString <> sReadingTypeStr Then
        PosOfSourceStart = InStr(1, TheRecord, AfterString)
    Else
        PosOfSourceStart = Len(AfterString)
    End If
        
    RightChunk = Right(TheRecord, Len(TheRecord) - PosOfSourceStart)
    
    If Len(RightChunk) > 2 Then
        If Right(RightChunk, 1) <> ")" And AfterString <> sReadingTypeStr Then
            AcquireTheSourceMaterial = ""
            Exit Function
        End If
    End If
    
    If AfterString <> sReadingTypeStr Then
        AcquireTheSourceMaterial = Left(RightChunk, Len(RightChunk) - 1)
    Else
        AcquireTheSourceMaterial = Trim(RightChunk)
    End If
    
    '
    'Get rid of the funny paragraph markers and replace with "par"
    '
    AcquireTheSourceMaterial = Replace(AcquireTheSourceMaterial, Chr(182), "par")
    AcquireTheSourceMaterial = Replace(AcquireTheSourceMaterial, Chr(151), " - ") 'em dash
    AcquireTheSourceMaterial = Replace(AcquireTheSourceMaterial, Chr(150), " - ") 'en dash
'    AcquireTheSourceMaterial = Replace(AcquireTheSourceMaterial, "-", " - ")
    If Left(AcquireTheSourceMaterial, 1) = ":" Then
        AcquireTheSourceMaterial = Trim(Right(AcquireTheSourceMaterial, Len(AcquireTheSourceMaterial) - 1))
    End If
    
    
    Exit Function
    
ErrorTrap:
    EndProgram
End Function


Public Property Get TMSWeeksItems() As String
    TMSWeeksItems = theitems
End Property

Public Property Let TMSWeeksItems(ByVal vNewValue As String)
    theitems = vNewValue
End Property

Public Property Get TheTMSItemFilePath() As String
    TheTMSItemFilePath = FilePath
End Property

Public Property Let TheTMSItemFilePath(ByVal vNewValue As String)
    FilePath = vNewValue
End Property

Private Function AllRecsComplete(DateRecFound As Boolean, _
                            SQRecFound As Boolean, _
                            No1RecFound As Boolean, _
                            No2RecFound As Boolean, _
                            No3RecFound As Boolean, _
                            No4RecFound As Boolean, _
                            OralReviewRecFound As Boolean) As Boolean
    On Error GoTo ErrorTrap

    '
    'Reached end of logical grouping. Are all values filled?
    '
        If OralReviewRecFound Then
            If DateRecFound And SQRecFound Then
                AllRecsComplete = True
            Else
                AllRecsComplete = False
                Exit Function
            End If
        Else
            If DateRecFound And SQRecFound And No1RecFound And No2RecFound And _
                No3RecFound And No4RecFound Then
                AllRecsComplete = True
            Else
                AllRecsComplete = False
                Exit Function
            End If
        End If
    
    Exit Function
    
ErrorTrap:
    EndProgram
End Function
Private Function AllRecsComplete_2009(DateRecFound As Boolean, _
                            No1RecFound As Boolean, _
                            No2RecFound As Boolean, _
                            No3RecFound As Boolean, _
                            OralReviewRecFound As Boolean) As Boolean
    On Error GoTo ErrorTrap

    '
    'Reached end of logical grouping. Are all values filled?
    '
        If OralReviewRecFound Then
            If DateRecFound Then
                AllRecsComplete_2009 = True
            Else
                AllRecsComplete_2009 = False
                Exit Function
            End If
        Else
            If DateRecFound And No1RecFound And No2RecFound And _
                No3RecFound Then
                AllRecsComplete_2009 = True
            Else
                AllRecsComplete_2009 = False
                Exit Function
            End If
        End If
    
    Exit Function
    
ErrorTrap:
    EndProgram
End Function

Private Sub UpdateDB(TheWeeksItems As TMSScheduleRecord, OralReviewRecFound As Boolean)
    On Error GoTo ErrorTrap

    With TheWeeksItems

    '
    'Now add records to tblTMSItems for each TalkNo in turn...
    '
    AddRecordToDB .TMSUDT_AssignmentDate, "P", 47, 0, "Opening Prayer", "Song " & CStr(.TMSUDT_SongNo), 0, False
    AddRecordToDB .TMSUDT_AssignmentDate, "S", 33, 1, .TMSUDT_SQTheme, .TMSUDT_SQSource, 0, False
    AddRecordToDB .TMSUDT_AssignmentDate, "B", 34, 3, "Bible Highlights", .TMSUDT_BHSource, 0, False
    
    If Not OralReviewRecFound Then
        AddRecordToDB .TMSUDT_AssignmentDate, "1", 35, 2, .TMSUDT_No1Theme, .TMSUDT_No1Source, 0, False
        
        AddRecordToDB .TMSUDT_AssignmentDate, "2", 36, 4, .TMSUDT_No2Theme, .TMSUDT_No2Source, 0, False
        
        If .TMSUDT_No3Source = "" Then
            AddRecordToDB .TMSUDT_AssignmentDate, "3", 39, 5, .TMSUDT_No3Theme, .TMSUDT_No3Source, 0, False
        Else
            AddRecordToDB .TMSUDT_AssignmentDate, "3", 38, 5, .TMSUDT_No3Theme, .TMSUDT_No3Source, 0, False
        End If
    
        If .TMSUDT_No4Source = "" Then
            AddRecordToDB .TMSUDT_AssignmentDate, "4", 40, 6, .TMSUDT_No4Theme, .TMSUDT_No4Source, 0, .TMSUDT_BroOnlyForNo4
        Else
            AddRecordToDB .TMSUDT_AssignmentDate, "4", 41, 6, .TMSUDT_No4Theme, .TMSUDT_No4Source, 0, .TMSUDT_BroOnlyForNo4
        End If
    Else
        AddRecordToDB .TMSUDT_AssignmentDate, "R", 86, 4, "Oral Review", "Oral Review", 0, False
    End If
        
    End With
    Exit Sub
    
ErrorTrap:
    EndProgram
End Sub
Private Sub UpdateDB_2009(TheWeeksItems As TMSScheduleRecord, OralReviewRecFound As Boolean)
    On Error GoTo ErrorTrap

    With TheWeeksItems

    '
    'Now add records to tblTMSItems for each TalkNo in turn...
    '
    AddRecordToDB .TMSUDT_AssignmentDate, "P", 47, 0, "Opening Prayer", "", 0, False
    AddRecordToDB .TMSUDT_AssignmentDate, "B", 34, 1, "Bible Highlights", .TMSUDT_BHSource, 0, False
    
    If Not OralReviewRecFound Then
    
        AddRecordToDB .TMSUDT_AssignmentDate, "1", 99, 2, .TMSUDT_No1Theme, .TMSUDT_No1Source, 0, False
        
        AddRecordToDB .TMSUDT_AssignmentDate, "2", 100, 3, .TMSUDT_No2Theme, .TMSUDT_No2Source, 0, False
        
        AddRecordToDB .TMSUDT_AssignmentDate, "3", 101, 4, .TMSUDT_No3Theme, .TMSUDT_No3Source, 0, .TMSUDT_BroOnlyForNo3
    
    Else
        AddRecordToDB .TMSUDT_AssignmentDate, "R", 86, 2, "Oral Review", "Oral Review", 0, False
    End If
        
    End With
    Exit Sub
    
ErrorTrap:
    EndProgram
End Sub

Public Sub AnyUnprintedTMSSlips()
On Error GoTo ErrorTrap
Dim rstCheckTMS As Recordset, rstCheckTMS2 As Recordset, SQLStr As String, NoDays As Long

    If NewMtgArrangementStarted(CStr(Now)) <> CLM2016 Then Exit Sub

    Select Case True
    Case PersonHasAccess(gCurrentUserCode, CompleteAccess), _
         PersonHasAccess(gCurrentUserCode, TMSOverseer), _
         PersonHasAccess(gCurrentUserCode, TMSPrinting)
         
         NoDays = GlobalParms.GetValue("TMSWarnOfUnprintedSlipsDays", "NumVal")
         
         If NoDays > 0 Then
         
            SQLStr = "SELECT * " & _
                   "FROM tblTMSSchedule " & _
                   "WHERE AssignmentDate BETWEEN #" & Format(Now, "mm/dd/yyyy") & "# " & _
                   "AND #" & Format(DateAdd("d", NoDays, date), "mm/dd/yyyy") & "# " & _
                   "AND SlipPrinted = FALSE " & _
                   "AND PersonID <> 0 " & _
                   "AND TalkNo IN ('IC', 'BR', 'RV', 'BS', 'O') " & _
                   "AND SchoolNo <= " & GlobalParms.GetValue("TMSNoSchoolsForCounsel", "NumVal", 1)
            
            Set rstCheckTMS = CMSDB.OpenRecordset(SQLStr, dbOpenSnapshot)
            
            If Not rstCheckTMS.BOF Then
               MsgBox "Theocratic Ministry School Assignment Slips are due to be printed", _
               vbOKOnly + vbInformation, AppName
            Else
               SQLStr = "SELECT * " & _
                       "FROM tblTMSItems " & _
                       "WHERE AssignmentDate BETWEEN #" & Format(Now, "mm/dd/yyyy") & "# " & _
                       "AND #" & Format(DateAdd("d", NoDays, date), "mm/dd/yyyy") & "# " & _
                       "AND TalkNo IN ('IC', 'BR', 'RV', 'BS', 'O')"
               
               Set rstCheckTMS = CMSDB.OpenRecordset(SQLStr, dbOpenSnapshot)
               
               SQLStr = "SELECT * " & _
                       "FROM tblTMSSchedule " & _
                       "WHERE AssignmentDate BETWEEN #" & Format(Now, "mm/dd/yyyy") & "# " & _
                       "AND #" & Format(DateAdd("d", NoDays, date), "mm/dd/yyyy") & "# " & _
                       "AND TalkNo IN ('IC', 'BR', 'RV', 'BS', 'O')"
               
               Set rstCheckTMS2 = CMSDB.OpenRecordset(SQLStr, dbOpenDynaset)
               
               With rstCheckTMS
               If Not .BOF Then
                   .MoveFirst
                   Do Until .EOF
                       rstCheckTMS2.FindFirst "AssignmentDate = #" & Format(!AssignmentDate, "mm/dd/yyyy") & "# " & _
                                               "AND TalkNo = '" & !TalkNo & "' AND PersonID <> 0"
                                               
                       If rstCheckTMS2.NoMatch Then
                            If Not IsCircuitOrDistrictAssemblyWeek(!AssignmentDate) And _
                                Not (IsOralReviewWeek(!AssignmentDate) And !TalkNo <> "B") Then
                                
                                MsgBox "There are unassigned Theocratic Ministry School items due for printing", _
                                vbOKOnly + vbInformation, AppName
                                Exit Do
                                
                            End If
                       End If
                       .MoveNext
                   Loop
                   rstCheckTMS2.Close
               Else
                   MsgBox "You need to load new Theocratic Ministry School items", _
                   vbOKOnly + vbInformation, AppName
               End If
               End With
            End If
            rstCheckTMS.Close
         End If
    End Select
    
    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub

Public Sub TMSScheduleDueToPrint()
On Error GoTo ErrorTrap
Dim NoDays As Long, sNextDate As String, lCurrDays As Long

    If Not NewMtgArrangementStarted(CStr(Now)) Then Exit Sub

    Select Case True
    Case PersonHasAccess(gCurrentUserCode, CompleteAccess), _
         PersonHasAccess(gCurrentUserCode, TMSOverseer), _
         PersonHasAccess(gCurrentUserCode, TMSPrinting)
         
         NoDays = GlobalParms.GetValue("TMSWarnOfUnprintedScheduleDays", "NumVal")
         
         If NoDays > 0 Then
         
            sNextDate = HandleNull(GlobalParms.GetValue("NextTMSSchedulePrintStartDate", "DateVal"), "")
            
            If sNextDate <> "" Then
                lCurrDays = DateDiff("d", Now, CDate(sNextDate))
                If lCurrDays <= NoDays Then
                    MsgBox "Theocratic Ministry School Schedule is due to be published", _
                    vbOKOnly + vbInformation, AppName
                End If
            End If
                                    
         End If
         
    End Select
    
    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub



Public Sub ShowIndiviualNextPrevData(TheGrid As MSFlexGrid, PersonID As Long, CurrentAssignmentDate As Date)

    If NewMtgArrangementStarted(CStr(CurrentAssignmentDate)) Then
        ShowIndiviualNextPrevData_2009 TheGrid, PersonID, CurrentAssignmentDate
    Else
        ShowIndiviualNextPrevData_2002 TheGrid, PersonID, CurrentAssignmentDate
    End If

End Sub

Public Sub ShowIndiviualNextPrevData_2002(TheGrid As MSFlexGrid, PersonID As Long, CurrentAssignmentDate As Date)
On Error GoTo ErrorTrap
Dim OneStudentsDates As TMSPersonAndNextPrevDates
Dim PrevDates() As Variant, NextDates() As Variant, n As Integer

    '
    'Display single row's next/prev dates on frmTMSInsertStudent if
    ' chkShowNextPrevDates is off
    '
    If frmTMSInsertStudent.chkShowNextPrevDates = vbChecked Then Exit Sub

    '
    'Load all next/prev data to array
    '
    OneStudentsDates.NextPrevInfo(0) = CongregationMember.TMSPrevDateFAST(PersonID, "P", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(1) = CongregationMember.TMSPrevDateFAST(PersonID, "S", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(2) = CongregationMember.TMSPrevDateFAST(PersonID, "1", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(3) = CongregationMember.TMSPrevDateFAST(PersonID, "B", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(4) = CongregationMember.TMSPrevDateFAST(PersonID, "2", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(5) = CongregationMember.GetTMSSchoolNoForInsertForm
    OneStudentsDates.NextPrevInfo(6) = CongregationMember.TMSPrevDateFAST(PersonID, "3", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(7) = CongregationMember.GetTMSSchoolNoForInsertForm
    OneStudentsDates.NextPrevInfo(8) = CongregationMember.TMSPrevDateFAST(PersonID, "4", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(9) = CongregationMember.GetTMSSchoolNoForInsertForm
    OneStudentsDates.NextPrevInfo(10) = CongregationMember.TMSPrevDateFAST(PersonID, "Asst", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(11) = CongregationMember.GetTMSSchoolNoForInsertForm
    OneStudentsDates.NextPrevInfo(12) = CongregationMember.TMSNextDateFAST(PersonID, "P", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(13) = CongregationMember.TMSNextDateFAST(PersonID, "S", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(14) = CongregationMember.TMSNextDateFAST(PersonID, "1", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(15) = CongregationMember.TMSNextDateFAST(PersonID, "B", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(16) = CongregationMember.TMSNextDateFAST(PersonID, "2", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(17) = CongregationMember.GetTMSSchoolNoForInsertForm
    OneStudentsDates.NextPrevInfo(18) = CongregationMember.TMSNextDateFAST(PersonID, "3", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(19) = CongregationMember.GetTMSSchoolNoForInsertForm
    OneStudentsDates.NextPrevInfo(20) = CongregationMember.TMSNextDateFAST(PersonID, "4", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(21) = CongregationMember.GetTMSSchoolNoForInsertForm
    OneStudentsDates.NextPrevInfo(22) = CongregationMember.TMSNextDateFAST(PersonID, "Asst", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(23) = CongregationMember.GetTMSSchoolNoForInsertForm

    '
    'Put array contents on the grid
    '
    With TheGrid
    .TextMatrix(.Row, PrevPrayer) = OneStudentsDates.NextPrevInfo(0)
    .TextMatrix(.Row, PrevSQ) = OneStudentsDates.NextPrevInfo(1)
    .TextMatrix(.Row, PrevNo1) = OneStudentsDates.NextPrevInfo(2)
    .TextMatrix(.Row, PrevBH) = OneStudentsDates.NextPrevInfo(3)
    .TextMatrix(.Row, PrevNo2) = OneStudentsDates.NextPrevInfo(4)
    .TextMatrix(.Row, PrevNo2School) = OneStudentsDates.NextPrevInfo(5)
    .TextMatrix(.Row, PrevNo3) = OneStudentsDates.NextPrevInfo(6)
    .TextMatrix(.Row, PrevNo3School) = OneStudentsDates.NextPrevInfo(7)
    .TextMatrix(.Row, PrevNo4) = OneStudentsDates.NextPrevInfo(8)
    .TextMatrix(.Row, PrevNo4School) = OneStudentsDates.NextPrevInfo(9)
    .TextMatrix(.Row, PrevAsst) = OneStudentsDates.NextPrevInfo(10)
    .TextMatrix(.Row, PrevAsstSchool) = OneStudentsDates.NextPrevInfo(11)
    .TextMatrix(.Row, NextPrayer) = OneStudentsDates.NextPrevInfo(12)
    .TextMatrix(.Row, NextSQ) = OneStudentsDates.NextPrevInfo(13)
    .TextMatrix(.Row, NextNo1) = OneStudentsDates.NextPrevInfo(14)
    .TextMatrix(.Row, NextBH) = OneStudentsDates.NextPrevInfo(15)
    .TextMatrix(.Row, NextNo2) = OneStudentsDates.NextPrevInfo(16)
    .TextMatrix(.Row, NextNo2School) = OneStudentsDates.NextPrevInfo(17)
    .TextMatrix(.Row, NextNo3) = OneStudentsDates.NextPrevInfo(18)
    .TextMatrix(.Row, NextNo3School) = OneStudentsDates.NextPrevInfo(19)
    .TextMatrix(.Row, NextNo4) = OneStudentsDates.NextPrevInfo(20)
    .TextMatrix(.Row, NextNo4School) = OneStudentsDates.NextPrevInfo(21)
    .TextMatrix(.Row, NextAsst) = OneStudentsDates.NextPrevInfo(22)
    .TextMatrix(.Row, NextAsstSchool) = OneStudentsDates.NextPrevInfo(23)
    End With

     '
     'Now highlight in green the most recent
     ' next & previous talks given by each student. This aids user in
     ' spreading talks more evenly.
     '
     With TheGrid
     '
     'Put pertinent PrevDates into an array
     ' (Must change format since can only use variants in Array function)
     '
     PrevDates = Array(Format(OneStudentsDates.NextPrevInfo(1), "yyyy/mm/dd"), _
                    Format(OneStudentsDates.NextPrevInfo(2), "yyyy/mm/dd"), _
                    Format(OneStudentsDates.NextPrevInfo(3), "yyyy/mm/dd"), _
                    Format(OneStudentsDates.NextPrevInfo(4), "yyyy/mm/dd"), _
                    Format(OneStudentsDates.NextPrevInfo(6), "yyyy/mm/dd"), _
                    Format(OneStudentsDates.NextPrevInfo(8), "yyyy/mm/dd"))
     
     '
     'Sort desc
     '
     BubbleSort PrevDates, , True
     
    If PrevDates(0) <> "" Then
        Select Case Format(PrevDates(0), "dd/mm/yy")
        Case OneStudentsDates.NextPrevInfo(1)
           .col = PrevSQ
           .CellBackColor = PrevTalkColour
        Case OneStudentsDates.NextPrevInfo(2)
           .col = PrevNo1
           .CellBackColor = PrevTalkColour
        Case OneStudentsDates.NextPrevInfo(3)
           .col = PrevBH
           .CellBackColor = PrevTalkColour
        Case OneStudentsDates.NextPrevInfo(4)
           .col = PrevNo2
           .CellBackColor = PrevTalkColour
           .col = PrevNo2School
           .CellBackColor = PrevTalkColour
        Case OneStudentsDates.NextPrevInfo(6)
           .col = PrevNo3
           .CellBackColor = PrevTalkColour
           .col = PrevNo3School
           .CellBackColor = PrevTalkColour
        Case OneStudentsDates.NextPrevInfo(8)
           .col = PrevNo4
           .CellBackColor = PrevTalkColour
           .col = PrevNo4School
           .CellBackColor = PrevTalkColour
        End Select
    End If
    
     '
     'Put pertinent NextDates into an array
     ' (Must change format since can only use variants in Array function)
     '
     NextDates = Array(Format(OneStudentsDates.NextPrevInfo(13), "yyyy/mm/dd"), _
                    Format(OneStudentsDates.NextPrevInfo(14), "yyyy/mm/dd"), _
                    Format(OneStudentsDates.NextPrevInfo(15), "yyyy/mm/dd"), _
                    Format(OneStudentsDates.NextPrevInfo(16), "yyyy/mm/dd"), _
                    Format(OneStudentsDates.NextPrevInfo(18), "yyyy/mm/dd"), _
                    Format(OneStudentsDates.NextPrevInfo(20), "yyyy/mm/dd"))
     
     '
     'Sort asc
     '
     BubbleSort NextDates, , False
     
     '
     'Find the minimum non-blank date
     '
    For n = 0 To 5
        If NextDates(n) <> "" Then
            Exit For
        End If
    Next n
     
    If n > 5 Then
        n = 5
    End If
     
    If NextDates(n) <> "" Then
        Select Case Format(NextDates(n), "dd/mm/yy")
        Case OneStudentsDates.NextPrevInfo(13)
           .col = NextSQ
           .CellBackColor = NextTalkColour
        Case OneStudentsDates.NextPrevInfo(14)
           .col = NextNo1
           .CellBackColor = NextTalkColour
        Case OneStudentsDates.NextPrevInfo(15)
           .col = NextBH
           .CellBackColor = NextTalkColour
        Case OneStudentsDates.NextPrevInfo(16)
           .col = NextNo2
           .CellBackColor = NextTalkColour
           .col = NextNo2School
           .CellBackColor = NextTalkColour
        Case OneStudentsDates.NextPrevInfo(18)
           .col = NextNo3
           .CellBackColor = NextTalkColour
           .col = NextNo3School
           .CellBackColor = NextTalkColour
        Case OneStudentsDates.NextPrevInfo(20)
           .col = NextNo4
           .CellBackColor = NextTalkColour
           .col = NextNo4School
           .CellBackColor = NextTalkColour
        End Select
    End If
    End With

    Exit Sub
    
ErrorTrap:
    EndProgram
End Sub

Public Sub ShowIndiviualNextPrevData_2009(TheGrid As MSFlexGrid, PersonID As Long, CurrentAssignmentDate As Date)
On Error GoTo ErrorTrap
Dim OneStudentsDates As TMSPersonAndNextPrevDates_2009
Dim PrevDates() As Variant, NextDates() As Variant, n As Integer, bHighlightAsstDates As Boolean
    '
    'Display single row's next/prev dates on frmTMSInsertStudent if
    ' chkShowNextPrevDates is off
    '
    If frmTMSInsertStudent.chkShowNextPrevDates = vbChecked Then Exit Sub
    

    '
    'Load all next/prev data to array
    '
    OneStudentsDates.NextPrevInfo(0) = CongregationMember.TMSPrevDateFAST(PersonID, "B", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(1) = CongregationMember.TMSPrevDateFAST(PersonID, "1", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(2) = CongregationMember.GetTMSSchoolNoForInsertForm
    OneStudentsDates.NextPrevInfo(3) = CongregationMember.TMSPrevDateFAST(PersonID, "2", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(4) = CongregationMember.GetTMSSchoolNoForInsertForm
    OneStudentsDates.NextPrevInfo(5) = CongregationMember.TMSPrevDateFAST(PersonID, "3", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(6) = CongregationMember.GetTMSSchoolNoForInsertForm
    OneStudentsDates.NextPrevInfo(7) = CongregationMember.TMSPrevDateFAST(PersonID, "Asst", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(8) = CongregationMember.GetTMSSchoolNoForInsertForm
    OneStudentsDates.NextPrevInfo(9) = CongregationMember.TMSNextDateFAST(PersonID, "B", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(10) = CongregationMember.TMSNextDateFAST(PersonID, "1", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(11) = CongregationMember.GetTMSSchoolNoForInsertForm
    OneStudentsDates.NextPrevInfo(12) = CongregationMember.TMSNextDateFAST(PersonID, "2", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(13) = CongregationMember.GetTMSSchoolNoForInsertForm
    OneStudentsDates.NextPrevInfo(14) = CongregationMember.TMSNextDateFAST(PersonID, "3", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(15) = CongregationMember.GetTMSSchoolNoForInsertForm
    OneStudentsDates.NextPrevInfo(16) = CongregationMember.TMSNextDateFAST(PersonID, "Asst", CurrentAssignmentDate)
    OneStudentsDates.NextPrevInfo(17) = CongregationMember.GetTMSSchoolNoForInsertForm

    '
    'Put array contents on the grid
    '
    With TheGrid
    .TextMatrix(.Row, PrevBH) = OneStudentsDates.NextPrevInfo(0)
    .TextMatrix(.Row, PrevNo1) = OneStudentsDates.NextPrevInfo(1)
    .TextMatrix(.Row, PrevNo1School) = OneStudentsDates.NextPrevInfo(2)
    .TextMatrix(.Row, PrevNo2) = OneStudentsDates.NextPrevInfo(3)
    .TextMatrix(.Row, PrevNo2School) = OneStudentsDates.NextPrevInfo(4)
    .TextMatrix(.Row, PrevNo3) = OneStudentsDates.NextPrevInfo(5)
    .TextMatrix(.Row, PrevNo3School) = OneStudentsDates.NextPrevInfo(6)
    .TextMatrix(.Row, PrevAsst) = OneStudentsDates.NextPrevInfo(7)
    .TextMatrix(.Row, PrevAsstSchool) = OneStudentsDates.NextPrevInfo(8)
    .TextMatrix(.Row, NextBH) = OneStudentsDates.NextPrevInfo(9)
    .TextMatrix(.Row, NextNo1) = OneStudentsDates.NextPrevInfo(10)
    .TextMatrix(.Row, NextNo1School) = OneStudentsDates.NextPrevInfo(11)
    .TextMatrix(.Row, NextNo2) = OneStudentsDates.NextPrevInfo(12)
    .TextMatrix(.Row, NextNo2School) = OneStudentsDates.NextPrevInfo(13)
    .TextMatrix(.Row, NextNo3) = OneStudentsDates.NextPrevInfo(14)
    .TextMatrix(.Row, NextNo3School) = OneStudentsDates.NextPrevInfo(15)
    .TextMatrix(.Row, NextAsst) = OneStudentsDates.NextPrevInfo(16)
    .TextMatrix(.Row, NextAsstSchool) = OneStudentsDates.NextPrevInfo(17)
    End With

     '
     'Now highlight in green the most recent
     ' next & previous talks given by each student. This aids user in
     ' spreading talks more evenly.
     '
     With TheGrid
     '
     'Put pertinent PrevDates into an array
     ' (Must change format since can only use variants in Array function)
     '
     If Not bHighlightAsstDates Then
        PrevDates = Array(Format(OneStudentsDates.NextPrevInfo(0), "yyyy/mm/dd"), _
                       Format(OneStudentsDates.NextPrevInfo(1), "yyyy/mm/dd"), _
                       Format(OneStudentsDates.NextPrevInfo(3), "yyyy/mm/dd"), _
                       Format(OneStudentsDates.NextPrevInfo(5), "yyyy/mm/dd"))
    Else
        PrevDates = Array(Format(OneStudentsDates.NextPrevInfo(0), "yyyy/mm/dd"), _
                       Format(OneStudentsDates.NextPrevInfo(1), "yyyy/mm/dd"), _
                       Format(OneStudentsDates.NextPrevInfo(3), "yyyy/mm/dd"), _
                       Format(OneStudentsDates.NextPrevInfo(5), "yyyy/mm/dd"), _
                       Format(OneStudentsDates.NextPrevInfo(7), "yyyy/mm/dd"))
    End If
     
     '
     'Sort desc
     '
     BubbleSort PrevDates, , True
     
    If PrevDates(0) <> "" Then
        Select Case Format(PrevDates(0), "dd/mm/yy")
        Case OneStudentsDates.NextPrevInfo(0)
           .col = PrevBH
           .CellBackColor = PrevTalkColour
        Case OneStudentsDates.NextPrevInfo(1)
           .col = PrevNo1
           .CellBackColor = PrevTalkColour
           .col = PrevNo1School
           .CellBackColor = PrevTalkColour
        Case OneStudentsDates.NextPrevInfo(3)
           .col = PrevNo2
           .CellBackColor = PrevTalkColour
           .col = PrevNo2School
           .CellBackColor = PrevTalkColour
        Case OneStudentsDates.NextPrevInfo(5)
           .col = PrevNo3
           .CellBackColor = PrevTalkColour
           .col = PrevNo3School
           .CellBackColor = PrevTalkColour
        Case OneStudentsDates.NextPrevInfo(7)
           .col = PrevAsst
           .CellBackColor = PrevTalkColour
           .col = PrevAsstSchool
           .CellBackColor = PrevTalkColour
        End Select
    End If
    
     '
     'Put pertinent NextDates into an array
     ' (Must change format since can only use variants in Array function)
     '
     If Not bHighlightAsstDates Then
        NextDates = Array(Format(OneStudentsDates.NextPrevInfo(9), "yyyy/mm/dd"), _
                       Format(OneStudentsDates.NextPrevInfo(10), "yyyy/mm/dd"), _
                       Format(OneStudentsDates.NextPrevInfo(12), "yyyy/mm/dd"), _
                       Format(OneStudentsDates.NextPrevInfo(14), "yyyy/mm/dd"))
    Else
        NextDates = Array(Format(OneStudentsDates.NextPrevInfo(9), "yyyy/mm/dd"), _
                       Format(OneStudentsDates.NextPrevInfo(10), "yyyy/mm/dd"), _
                       Format(OneStudentsDates.NextPrevInfo(12), "yyyy/mm/dd"), _
                       Format(OneStudentsDates.NextPrevInfo(14), "yyyy/mm/dd"), _
                       Format(OneStudentsDates.NextPrevInfo(16), "yyyy/mm/dd"))
    End If
     
     '
     'Sort asc
     '
     BubbleSort NextDates, , False
     
     '
     'Find the minimum non-blank date
     '
    For n = 0 To UBound(NextDates)
        If NextDates(n) <> "" Then
            Exit For
        End If
    Next n
     
    If n > UBound(NextDates) Then
        n = UBound(NextDates)
    End If
     
    If NextDates(n) <> "" Then
        Select Case Format(NextDates(n), "dd/mm/yy")
        Case OneStudentsDates.NextPrevInfo(9)
           .col = NextBH
           .CellBackColor = NextTalkColour
        Case OneStudentsDates.NextPrevInfo(10)
           .col = NextNo1
           .CellBackColor = NextTalkColour
           .col = NextNo1School
           .CellBackColor = NextTalkColour
        Case OneStudentsDates.NextPrevInfo(12)
           .col = NextNo2
           .CellBackColor = NextTalkColour
           .col = NextNo2School
           .CellBackColor = NextTalkColour
        Case OneStudentsDates.NextPrevInfo(14)
           .col = NextNo3
           .CellBackColor = NextTalkColour
           .col = NextNo3School
           .CellBackColor = NextTalkColour
        Case OneStudentsDates.NextPrevInfo(16)
           .col = NextAsst
           .CellBackColor = NextTalkColour
           .col = NextAsstSchool
           .CellBackColor = NextTalkColour
        End Select
    End If
    End With

    Exit Sub
    
ErrorTrap:
    EndProgram
End Sub

Public Sub AddWeekNumbersToItems(TheYear As Integer)
Dim rstItems As Recordset, i As Long, OralReviewFound As Boolean
On Error GoTo ErrorTrap

    Set rstItems = CMSDB.OpenRecordset("SELECT * " & _
                                       "FROM tblTMSItems " & _
                                       "WHERE Year(AssignmentDate) = " & TheYear, _
                                       dbOpenDynaset)
    
    If rstItems.BOF Then
        Exit Sub
    End If
    
    With rstItems
    
    Do Until .EOF
        If !TalkSeqNum = 0 Then
            If OralReviewFound Then
                i = 1
                OralReviewFound = False
            Else
                i = i + 1
            End If
        End If
        If !TalkNo = "R" Then
            OralReviewFound = True
        End If
        .Edit
        !weekno = i
        .Update
        .MoveNext
    Loop
    
    End With
    
    Exit Sub
    
ErrorTrap:
    EndProgram
End Sub

Public Function AwkwardCounselPoint(ByVal CounselPoint As String)

On Error GoTo ErrorTrap

Dim str As String

    If Trim(CounselPoint) = "" Then
        AwkwardCounselPoint = False
        Exit Function
    End If

    If GlobalParms.GetValue("TMS_AwkwardCounselPoints", "TrueFalse") = True Then

        str = GlobalParms.GetValue("TMS_AwkwardCounselPoints", "AlphaVal")
        
        AwkwardCounselPoint = (InStr(1, str, "~" & CounselPoint & "~") > 0)
    
    Else
    
        AwkwardCounselPoint = False
    
    End If
    
    Exit Function
    
ErrorTrap:
    EndProgram

End Function



