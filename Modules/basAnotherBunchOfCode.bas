Attribute VB_Name = "basAnotherBunchOfCode"
Option Explicit

Public Sub PrintMissingReports()
On Error GoTo ErrorTrap
Dim MinSQL As String, rstPrintReport As Recordset, rstPrintTable As Recordset
Dim DateString_US As String, DateString2_US As String, DateString2_UK As String
Dim lsNormalYear As String, lsNormalYear2 As String
Dim PrevGroup As Long, PrevDate As Date, PrevPerson As Long, TheSpacer As String
Dim str As String, sLastReportingPeriod As String, MinSQL2 As String, mlStorePersonID As Long
Dim sMonths As String, bUseWord As cmsPrintUsingWord, MinSQL3 As String, ErrorCode As Integer
Dim bAddComma As Boolean

    
    bUseWord = PrintUsingWord
    
    If bUseWord = cmsDontPrint Then Exit Sub
    
    Screen.MousePointer = vbHourglass

    'first build the missing report print table
    
    DelAllRows "tblPrintCongMinByGroup"
    
    'find the date range for reporting...
    With frmCongStats
    Select Case True
    Case .optCalendar
        
        If .Month_PubNo = 0 Or .Month_PubNo2 = 0 Or _
           .Year_PubNo = 0 Or .Year_PubNo2 = 0 Then
            MsgBox "Invalid date range", vbOKOnly + vbExclamation, AppName
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
        
        'Date from which we're interested - as selected in combos
        DateString_US = CStr(.Month_PubNo) & "/01/" & .Year_PubNo
        'Date up to which we're interested, as selected in combos
        DateString2_US = CStr(.Month_PubNo2) & "/01/" & .Year_PubNo2
        
    Case .optServiceYear
        
        If .Year_PubNo = 0 Then
            MsgBox "Invalid date range", vbOKOnly + vbExclamation, AppName
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
    
        lsNormalYear = CStr(ConvertServiceYearToNormalYear(CDate("01/09" & "/" & .Year_PubNo)))
        
        'Date from which we're interested - ie start of service year
        DateString_US = "09/01/" & lsNormalYear
        
        'Date up to which we're interested, as selected in combos
        If .Month_PubNo = 0 Then
            DateString2_US = "08/01/" & CStr(.Year_PubNo)
        Else
            lsNormalYear2 = CStr(ConvertServiceYearToNormalYear(CDate("01/" & .Month_PubNo & "/" & .Year_PubNo)))
            DateString2_US = CStr(.Month_PubNo) & "/01/" & lsNormalYear2
        End If
    
    End Select
    End With
    
    'get last complete Society reporting period
    sLastReportingPeriod = "01/" & GetLastSocietyReportingPeriodMMYYYY(Now)
            
    'now, if the selected end-date is one month after the last complete reporting period
    ' (ie, we want to include THIS month's missing reports not yet logged on
    ' tblMissingReports) we need to UNION this SQL with the existing report SQL...
    If DateDiff("m", CDate(sLastReportingPeriod), CDate(frmCongStats.ActualEndDate)) = 1 Then

        str = " UNION ALL " & _
            " SELECT DISTINCTROW ID,  " & _
                        "LastName, FirstName, MiddleName, " & _
                        "tblBookGroups.GroupNo, tblBookGroups.GroupName, '" & _
                        frmCongStats.ActualEndDate & "' AS ActualMinDate " & _
                "FROM ((tblNameAddress INNER JOIN tblPublisherDates ON " & _
                "      tblNameAddress.ID = tblPublisherDates.PersonID) INNER JOIN " & _
                "      tblBookGroupMembers ON " & _
                "      tblPublisherDates.PersonID = tblBookGroupMembers.PersonID) " & _
                "      INNER JOIN tblBookGroups ON tblBookGroups.GroupNo = " & _
                "                                      tblBookGroupMembers.GroupNo " & _
                "WHERE StartDate <= #" & Format(frmCongStats.ActualEndDate, "mm/dd/yyyy") & "# " & _
                "AND EndDate >= #" & Format(frmCongStats.ActualEndDate, "mm/dd/yyyy") & "# " & _
                " AND StartReason <> 2 " & _
                "AND tblPublisherDates.PersonID NOT IN " & _
                "            (SELECT PersonID " & _
                            " FROM tblMinReports " & _
                            " WHERE ActualMinPeriod = #" & Format(frmCongStats.ActualEndDate, "mm/dd/yyyy") & "#) "

    Else

        str = ""

    End If
    
    '
    'Driving Recset to get all missing reports logged on tblMissingReports
    '
    MinSQL = "SELECT DISTINCTROW ID, " & _
                    "LastName, FirstName, MiddleName, " & _
                    "tblBookGroups.GroupNo, tblBookGroups.GroupName, " & _
                    "tblMissingReports.ActualMinDate " & _
             "FROM tblMissingReports INNER JOIN (tblPublisherDates INNER JOIN " & _
             "(tblNameAddress INNER JOIN (tblBookGroupMembers INNER JOIN " & _
             "tblBookGroups ON tblBookGroupMembers.GroupNo = " & _
             "tblBookGroups.GroupNo) ON tblNameAddress.ID = " & _
             "tblBookGroupMembers.PersonID) ON tblPublisherDates.PersonID = " & _
             "tblNameAddress.ID) ON tblMissingReports.PersonID = " & _
             "tblPublisherDates.PersonID " & _
             "WHERE StartDate <= #" & _
              DateString2_US & _
             "# AND EndDate >= #" & DateString_US & "# " & _
             "AND ZeroReport = FALSE " & _
            " AND StartReason <> 2 " & _
             "AND ActualMinDate BETWEEN #" & DateString_US & _
             "# AND #" & DateString2_US & "# " & _
            str & " ORDER BY 6, 7, 2, 3, 4, 1 "
             
    Set rstPrintReport = CMSDB.OpenRecordset(MinSQL, dbOpenDynaset)
    Set rstPrintTable = CMSDB.OpenRecordset("tblPrintCongMinByGroup", dbOpenDynaset)
    
    If rstPrintReport.BOF Then
        MsgBox "Nothing to print.", vbOKOnly + vbExclamation, AppName
        Screen.MousePointer = vbNormal
        Exit Sub
    Else
        
        If bUseWord = cmsUseWord Then
                    
            If str <> "" Then
                'we've determined already that we want to report on missing reports
                ' up to current month not yet logged on tblMissingReports (that's why
                ' str<>"") SQL needed for this is....
                str = " UNION ALL " & _
                    " SELECT DISTINCTROW tblBookGroups.GroupName AS TheGrp,  " & _
                            "LastName & ', ' &  FirstName & ' ' & MiddleName AS PersonName, '" & _
                                Format(frmCongStats.ActualEndDate, "mmmm yyyy") & "' AS MinDate, '" & _
                                Format(frmCongStats.ActualEndDate, "yyyymmdd") & "' AS MinDate2, '" & _
                                " frmCongStats.ActualEndDate ' AS MinDate3, ID " & _
                        "FROM ((tblNameAddress INNER JOIN tblPublisherDates ON " & _
                        "      tblNameAddress.ID = tblPublisherDates.PersonID) INNER JOIN " & _
                        "      tblBookGroupMembers ON " & _
                        "      tblPublisherDates.PersonID = tblBookGroupMembers.PersonID) " & _
                        "      INNER JOIN tblBookGroups ON tblBookGroups.GroupNo = " & _
                        "                                      tblBookGroupMembers.GroupNo " & _
                        "WHERE StartDate <= #" & Format(frmCongStats.ActualEndDate, "mm/dd/yyyy") & "# " & _
                        "AND EndDate >= #" & Format(frmCongStats.ActualEndDate, "mm/dd/yyyy") & "# " & _
                        " AND StartReason <> 2 " & _
                        "AND tblPublisherDates.PersonID NOT IN " & _
                        "            (SELECT PersonID " & _
                                    " FROM tblMinReports " & _
                                    " WHERE ActualMinPeriod = #" & Format(frmCongStats.ActualEndDate, "mm/dd/yyyy") & "#) "
            
            End If
            
            
            MinSQL = "SELECT DISTINCTROW tblBookGroups.GroupName AS TheGrp, " & _
                            "LastName & ', ' &  FirstName & ' ' & MiddleName AS PersonName, " & _
                            "format(tblMissingReports.ActualMinDate, 'mmmm yyyy') AS MinDate, " & _
                            "format(tblMissingReports.ActualMinDate, 'yyyymmdd') AS MinDate2, " & _
                            "tblMissingReports.ActualMinDate AS MinDate3, ID " & _
                     "FROM tblMissingReports INNER JOIN (tblPublisherDates INNER JOIN " & _
                     "(tblNameAddress INNER JOIN (tblBookGroupMembers INNER JOIN " & _
                     "tblBookGroups ON tblBookGroupMembers.GroupNo = " & _
                     "tblBookGroups.GroupNo) ON tblNameAddress.ID = " & _
                     "tblBookGroupMembers.PersonID) ON tblPublisherDates.PersonID = " & _
                     "tblNameAddress.ID) ON tblMissingReports.PersonID = " & _
                     "tblPublisherDates.PersonID " & _
                     "WHERE StartDate <= #" & _
                      DateString2_US & _
                     "# AND EndDate >= #" & DateString_US & "# " & _
                     "AND ZeroReport = FALSE " & _
                     " AND StartReason <> 2 " & _
                     "AND ActualMinDate BETWEEN #" & DateString_US & _
                     "# AND #" & DateString2_US & "# " & _
                    str & _
                     "ORDER BY 1, 2, 4"
                                          
            '
            'Now do SELECT from SELECT.... This is needed because we need to order by a
            ' column which is not included in the MSWord SQL, and if we do try to include
            ' it the dll will complain that the no of recset fields <> number of report
            ' fields! So do the order-by within the inner select....
            '
            MinSQL2 = "SELECT TheGrp, ID, PersonName, MinDate3 FROM (" & MinSQL & ")"
           
           'build the print table
            
            DeleteTable "tblPrintMissingReports"
            CreateTable ErrorCode, "tblPrintMissingReports", "PersonID", "LONG", , , False
            CreateField ErrorCode, "tblPrintMissingReports", "GroupName", "MEMO"
            CreateField ErrorCode, "tblPrintMissingReports", "PersonName", "MEMO"
            CreateField ErrorCode, "tblPrintMissingReports", "Months", "MEMO"
            
            Set rstPrintReport = CMSDB.OpenRecordset(MinSQL2, dbOpenDynaset)
            Set rstPrintTable = CMSDB.OpenRecordset("tblPrintMissingReports", dbOpenDynaset)
            
            With rstPrintReport

            Do Until .EOF 'add new print rec to tblPrintMissingReports for each row found

                rstPrintTable.AddNew

                rstPrintTable!PersonID = !ID
                rstPrintTable!GroupName = !TheGrp
                rstPrintTable!PersonName = !PersonName

                mlStorePersonID = !ID
                sMonths = ""
                bAddComma = False
                
                
                Do Until .EOF Or .BOF
                    If mlStorePersonID <> !ID Then Exit Do
                    mlStorePersonID = !ID
                    sMonths = sMonths & IIf(bAddComma, ", ", "") & GetMonthName(Month(!MinDate3)) & _
                                                     Chr(160) & CStr(year(!MinDate3)) 'chr(160) is non-breaking space
                    bAddComma = True
                    .MoveNext
                Loop

                rstPrintTable!Months = sMonths

                rstPrintTable.Update

            Loop

            End With

            MinSQL3 = "SELECT GroupName, PersonName, Months FROM tblPrintMissingReports"
        
            GenerateMissingReportPrintInWord MinSQL3, Format(DateString_US, "mm/dd/yyyy"), _
                                                     IIf(str = "", Format(DateString2_US, "mm/dd/yyyy"), frmCongStats.ActualEndDate)
        Else
        
            With rstPrintReport
    
            Do Until .EOF 'add new print rec to tblPrintCongMinByGroup for each row found
    
                rstPrintTable.AddNew
    
                rstPrintTable!CalendarMonthAndYear = GetMonthName(Month(!ActualMinDate)) & _
                                                    " " & CStr(year(!ActualMinDate))
                rstPrintTable!PersonID = !ID
    
                rstPrintTable.Update
    
                .MoveNext
            Loop
    
            End With
        
            GenerateMissingReportPrint Format(DateString_US, "mm/dd/yyyy"), _
                                       IIf(str = "", Format(DateString2_US, "mm/dd/yyyy"), frmCongStats.ActualEndDate)
        End If
        
    
    End If
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
ErrorTrap:
    EndProgram
    
    
End Sub
Private Sub GenerateMissingReportPrintInWord(TheSQL As String, StartDateUK As String, EndDateUK As String)
On Error GoTo ErrorTrap
    
'print to word

Dim RptTitle As String
Dim reporter As MSWordReportingTool2.RptTool

On Error GoTo ErrorTrap

    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    RptTitle = MonthName(Month(CDate(StartDateUK)), True) & " " & _
                CStr(year(CDate(StartDateUK)))
    
    If StartDateUK <> EndDateUK Then
        RptTitle = RptTitle & " to " & _
                                 MonthName(Month(CDate(EndDateUK)), True) & " " & _
                                 CStr(year((EndDateUK)))
    End If
    
    .ReportSQL = TheSQL

    .ReportTitle = "Missing Field Service Reports"
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 10
    .RightMargin = 10
    .ShowPageNumber = True
    .ReportFooterFontName = "Arial"
    .ReportFooterFontSize = 8
    .ReportTitleFontName = "Times New Roman"
    .ReportTitleFontSize = 13
    .ApplyTableFormatting = True
    .ClientName = AppName
    .AdditionalReportHeading = RptTitle

    .AdditionalReportHeadingFontName = "Arial"
    .AdditionalReportHeadingFontSize = 10
    .AdditionalReportHeadingBold = False
    .AdditionalReportHeadingItalic = False
    .GroupingColumn = 1
    .HideWordWhileBuilding = True
    .RaiseReportCompleteEvent = False
    
    .AddTableColumnAttribute "Group", 50, , , , , 10, 10, True, , , , True
    .AddTableColumnAttribute "Publishers", 50, , , , , 10, 10, True
    .AddTableColumnAttribute "Month", 50, , , , , 10, 10, True
    .PageFormat = cmsPortrait
    .SaveDoc = True
    .DocPath = gsDocsDirectory & "\" & "Missing Field Service Reports " & RptTitle & " " & _
                                Replace(Replace(Now, ":", "-"), "/", "-")
    
    .GenerateReport
       
    End With
    
    SwitchOnDAO
    
    Screen.MousePointer = vbNormal

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Function GenerateMissingReportPrint(FromDateUK As String, _
                                           ToDateUK As String) As Boolean
Dim WorkSheetTopMargin As Single
Dim WorkSheetBottomMargin As Single
Dim WorkSheetLeftMargin As Single
Dim WorkSheetRightMargin As Single
Dim RptTitle As String

On Error GoTo ErrorTrap

    Screen.MousePointer = vbHourglass
    
    '
    'Arrange page margins before we close the db connection
    '
    WorkSheetTopMargin = 566.929 * (GlobalParms.GetValue("A4TopMargin", "NumFloat"))
    WorkSheetBottomMargin = 566.929 * (GlobalParms.GetValue("A4BottomMargin", "NumFloat"))
    WorkSheetLeftMargin = 566.929 * (GlobalParms.GetValue("A4LeftMargin", "NumFloat"))
    WorkSheetRightMargin = 566.929 * (GlobalParms.GetValue("A4RightMargin", "NumFloat"))

    DestroyGlobalObjects
    CMSDB.Close
    
    '
    'GENERAL ADO WARNING...
    ' If we refer to DataReport prior to 'Showing' it, thus opening new ADODB connection
    ' while DAO connection still open, we get funny results.. eg missing fields on report.
    '
        
    MissingReportByGroup.TopMargin = WorkSheetTopMargin '<----- At this point, TMSStudentDetails.Initialize runs.
    MissingReportByGroup.BottomMargin = WorkSheetBottomMargin
    MissingReportByGroup.LeftMargin = WorkSheetLeftMargin
    MissingReportByGroup.RightMargin = WorkSheetRightMargin
    
    Screen.MousePointer = vbNormal
    
    RptTitle = "Missing Field Service Reports - " & _
                MonthName(Month(CDate(FromDateUK)), True) & " " & _
                CStr(year(CDate(FromDateUK)))
    
    If FromDateUK <> ToDateUK Then
        RptTitle = RptTitle & " to " & _
                                 MonthName(Month(CDate(ToDateUK)), True) & " " & _
                                 CStr(year((ToDateUK)))
    End If
    
    MissingReportByGroup.Sections("PageHeader").Controls("lblTitle").Caption = RptTitle

    MissingReportByGroup.Show vbModal
    
    InstantiateGlobalObjects

    GenerateMissingReportPrint = True

    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Function PrintAdvancedReport(RepSQL As String) As Boolean
Dim WorkSheetTopMargin As Single
Dim WorkSheetBottomMargin As Single
Dim WorkSheetLeftMargin As Single
Dim WorkSheetRightMargin As Single
Dim RptTitle As String, rstRepData As Recordset, TheSQL As String, UseWord As cmsPrintUsingWord

On Error GoTo ErrorTrap

    UseWord = PrintUsingWord
    
    If UseWord = cmsDontPrint Then
        PrintAdvancedReport = True
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    '
    'Build the reporting table
    '
    DelAllRows "tblAdvancedMinReportingPrint"
    
    Set rstRepData = CMSDB.OpenRecordset(RepSQL, dbOpenForwardOnly)
    
    With rstRepData
    
    Do Until .EOF Or .BOF
        TheSQL = "INSERT INTO tblAdvancedMinReportingPrint " & _
                          "(PersonID, " & _
                            "PersonName, " & _
                            "AvgBooks, " & _
                            "AvgBooklets, " & _
                            "AvgHours, " & _
                            "AvgMagazines, " & _
                            "AvgReturnVisits, " & _
                            "AvgStudies, " & _
                            "AvgTracts, " & _
                            "BookGroupName) " & _
                          "VALUES (" & !PersonID & ", '" & _
                                       DoubleUpSingleQuotes(!TheFullName) & "', " & _
                                       Round(!AvgBooks, 2) & ", " & _
                                       Round(!AvgBooklets, 2) & ", " & _
                                       Round(!AvgHours, 2) & ", " & _
                                       Round(!AvgMagazines, 2) & ", " & _
                                       Round(!AvgReturnVisits, 2) & ", " & _
                                       Round(!AvgStudies, 2) & ", " & _
                                       Round(!AvgTracts, 2) & ", '" & _
                                       GetGroupName(!BookGroupID) & "')"
        
        CMSDB.Execute TheSQL
        
        .MoveNext
    Loop
    
    .Close
    End With
    
    If UseWord = cmsUseWord Then
        PrintAdvancedReportToWord
        Exit Function
    End If
    
    '
    'Arrange page margins before we close the db connection
    '
    WorkSheetTopMargin = 566.929 * (GlobalParms.GetValue("A4TopMargin", "NumFloat"))
    WorkSheetBottomMargin = 566.929 * (GlobalParms.GetValue("A4BottomMargin", "NumFloat"))
    WorkSheetLeftMargin = 566.929 * (GlobalParms.GetValue("A4LeftMargin", "NumFloat"))
    WorkSheetRightMargin = 566.929 * (GlobalParms.GetValue("A4RightMargin", "NumFloat"))

    SwitchOffDAO
        
    '
    'GENERAL ADO WARNING...
    ' If we refer to DataReport prior to 'Showing' it, thus opening new ADODB connection
    ' while DAO connection still open, we get funny results.. eg missing fields on report.
    '
        
    AdvancedMinReport.TopMargin = WorkSheetTopMargin '<----- At this point, TMSStudentDetails.Initialize runs.
    AdvancedMinReport.BottomMargin = WorkSheetBottomMargin
    AdvancedMinReport.LeftMargin = WorkSheetLeftMargin
    AdvancedMinReport.RightMargin = WorkSheetRightMargin
    
    Screen.MousePointer = vbNormal
        
    AdvancedMinReport.Show vbModal
    
    SwitchOnDAO

    PrintAdvancedReport = True

    Exit Function
ErrorTrap:
    EndProgram
    
End Function
Public Sub PrintAdvancedReportToWord()

Dim reporter As MSWordReportingTool2.RptTool

On Error GoTo ErrorTrap

    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .ReportSQL = "SELECT PersonName, " & _
                    "       AvgBooks, " & _
                    "       AvgBooklets, " & _
                    "       AvgHours, " & _
                    "       AvgMagazines, " & _
                    "       AvgReturnVisits, " & _
                    "       AvgStudies, " & _
                    "       AvgTracts " & _
                    "FROM tblAdvancedMinReportingPrint "

    .ReportTitle = "Field Service Averages"
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 10
    .RightMargin = 10
    .ReportFooterFontName = "Arial"
    .ReportFooterFontSize = 8
    .ReportTitleFontName = "Times New Roman"
    .ReportTitleFontSize = 16
    .ApplyTableFormatting = True
    .ClientName = AppName
    .AdditionalReportHeading = vbCr & frmAdvancedCongStats.Criteria
    .AdditionalReportHeadingFontName = "Arial"
    .AdditionalReportHeadingFontSize = 8
    .AdditionalReportHeadingBold = False
    .AdditionalReportHeadingItalic = True
    .ShowPageNumber = True
    .GroupingColumn = 0
    .HideWordWhileBuilding = True
    .SaveDoc = True
    .NumberFormat = "0.00"
    
    .DocPath = gsDocsDirectory & "\" & "Congregation Stats " & _
                                Replace(Replace(Now, ":", "-"), "/", "-")
    
    
    .AddTableColumnAttribute "Name", 50, , , , , 10, 10, True, True, , , True
    .AddTableColumnAttribute "Av Bks", 20, , cmsRightTop, , , 10, 10, True, True
    .AddTableColumnAttribute "Av Bklts", 20, , cmsRightTop, , , 10, 10, True, True
    .AddTableColumnAttribute "Av Hrs", 20, , cmsRightTop, , , 10, 10, True, True
    .AddTableColumnAttribute "Av Mags", 20, , cmsRightTop, , , 10, 10, True, True
    .AddTableColumnAttribute "Av RVs", 20, , cmsRightTop, , , 10, 10, True, True
    .AddTableColumnAttribute "Av Std", 20, , cmsRightTop, , , 10, 10, True, True
    .AddTableColumnAttribute "Av Tra", 20, , cmsRightTop, , , 10, 10, True, True
    .PageFormat = cmsPortrait
    
    .GenerateReport
       
    End With
    
    SwitchOnDAO
    
    Screen.MousePointer = vbNormal

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Public Sub PrintBookGroups()
On Error GoTo ErrorTrap
Dim MinSQL As String, rstPrintReport As Recordset, rstPrintTable As Recordset

    Screen.MousePointer = vbHourglass

    'first build the print table
    
    DelAllRows "tblPrintBookGroups"
    
    '
    'Driving Recset to get all data
    '
    MinSQL = "SELECT ID, " & _
                    "LastName, FirstName, MiddleName, " & _
                    "tblBookGroups.GroupNo, tblBookGroups.GroupName " & _
             "FROM tblNameAddress INNER JOIN " & _
                    "(tblBookGroupMembers INNER JOIN tblBookGroups ON " & _
                        "tblBookGroupMembers.GroupNo = tblBookGroups.GroupNo) " & _
                        "ON tblNameAddress.ID = tblBookGroupMembers.PersonID " & _
             "WHERE Active = TRUE " & _
             "ORDER BY tblBookGroups.GroupName, " & _
             "         LastName, FirstName, MiddleName "
             
    Set rstPrintReport = CMSDB.OpenRecordset(MinSQL, dbOpenDynaset)
    Set rstPrintTable = CMSDB.OpenRecordset("tblPrintBookGroups", dbOpenDynaset)
    
    If rstPrintReport.BOF Then
        MsgBox "Nothing to print.", vbOKOnly + vbExclamation, AppName
        Screen.MousePointer = vbNormal
        Exit Sub
    Else
        With rstPrintReport
        
        Do Until .EOF 'add new print rec to tblPrintCongMinByGroup for each row found
        
            rstPrintTable.AddNew
                                    
            rstPrintTable!GroupName = !GroupName
            rstPrintTable!PersonName = !LastName & ", " & !FirstName & " " & !MiddleName
            rstPrintTable!GroupID = !GroupNo
            rstPrintTable!PersonID = !ID
            
            If CongregationMember.DoesRole(!ID, CInt(GlobalDefaultCong), 4, 4, 14) Then
                rstPrintTable!IsOverseer = "X"
            Else
                rstPrintTable!IsOverseer = ""
            End If
                
            If CongregationMember.DoesRole(!ID, CInt(GlobalDefaultCong), 4, 4, 15) Then
                rstPrintTable!IsAssistant = "X"
            Else
                rstPrintTable!IsAssistant = ""
            End If
                
            If CongregationMember.DoesRole(!ID, CInt(GlobalDefaultCong), 4, 4, 16) Then
                rstPrintTable!IsReader = "X"
            Else
                rstPrintTable!IsReader = ""
            End If
                
            If CongregationMember.DoesRole(!ID, CInt(GlobalDefaultCong), 4, 4, 17) Then
                rstPrintTable!IsPrayer = "X"
            Else
                rstPrintTable!IsPrayer = ""
            End If
                
            If CongregationMember.DoesRole(!ID, CInt(GlobalDefaultCong), 4, 4, 90) Then
                rstPrintTable!HasKM = "X"
                rstPrintTable!HasKM_Num = "1"
            Else
                rstPrintTable!HasKM = ""
                rstPrintTable!HasKM_Num = "0"
            End If
                                                   
            rstPrintTable.Update
            
            .MoveNext
        Loop
        
        End With
        
        Select Case PrintUsingWord
        Case cmsUseWord
            GenerateBookGroupPrintInWord
        Case cmsUseMSDatareport
            GenerateBookGroupPrint
        End Select
    
    End If
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
ErrorTrap:
    EndProgram
    
    
End Sub
Public Sub PrintNamesAndAddresses()
On Error GoTo ErrorTrap
Dim MinSQL As String, rstPrintReport As Recordset, rstPrintTable As Recordset

    Screen.MousePointer = vbHourglass

    'first build the print table
    
    DelAllRows "tblPrintBookGroups"
    
    '
    'Driving Recset to get all data
    '
    MinSQL = "SELECT ID, " & _
                    "LastName, FirstName, MiddleName, " & _
                    "tblBookGroups.GroupNo, tblBookGroups.GroupName " & _
             "FROM tblNameAddress INNER JOIN " & _
                    "(tblBookGroupMembers INNER JOIN tblBookGroups ON " & _
                        "tblBookGroupMembers.GroupNo = tblBookGroups.GroupNo) " & _
                        "ON tblNameAddress.ID = tblBookGroupMembers.PersonID " & _
             "WHERE Active = TRUE " & _
             "ORDER BY tblBookGroups.GroupName, " & _
             "         LastName, FirstName, MiddleName "
             
    Set rstPrintReport = CMSDB.OpenRecordset(MinSQL, dbOpenDynaset)
    Set rstPrintTable = CMSDB.OpenRecordset("tblPrintBookGroups", dbOpenDynaset)
    
    If rstPrintReport.BOF Then
        MsgBox "Nothing to print.", vbOKOnly + vbExclamation, AppName
        Screen.MousePointer = vbNormal
        Exit Sub
    Else
        With rstPrintReport
        
        Do Until .EOF 'add new print rec to tblPrintCongMinByGroup for each row found
        
            rstPrintTable.AddNew
                                    
            rstPrintTable!GroupName = !GroupName
            rstPrintTable!PersonName = !LastName & ", " & !FirstName & " " & !MiddleName
            rstPrintTable!GroupID = !GroupNo
            rstPrintTable!PersonID = !ID
            
            If CongregationMember.DoesRole(!ID, CInt(GlobalDefaultCong), 4, 4, 14) Then
                rstPrintTable!IsOverseer = "X"
            Else
                rstPrintTable!IsOverseer = ""
            End If
                
            If CongregationMember.DoesRole(!ID, CInt(GlobalDefaultCong), 4, 4, 15) Then
                rstPrintTable!IsAssistant = "X"
            Else
                rstPrintTable!IsAssistant = ""
            End If
                
            If CongregationMember.DoesRole(!ID, CInt(GlobalDefaultCong), 4, 4, 16) Then
                rstPrintTable!IsReader = "X"
            Else
                rstPrintTable!IsReader = ""
            End If
                
            If CongregationMember.DoesRole(!ID, CInt(GlobalDefaultCong), 4, 4, 17) Then
                rstPrintTable!IsPrayer = "X"
            Else
                rstPrintTable!IsPrayer = ""
            End If
                
            If CongregationMember.DoesRole(!ID, CInt(GlobalDefaultCong), 4, 4, 90) Then
                rstPrintTable!HasKM = "X"
                rstPrintTable!HasKM_Num = "1"
            Else
                rstPrintTable!HasKM = ""
                rstPrintTable!HasKM_Num = "0"
            End If
                                                   
            rstPrintTable.Update
            
            .MoveNext
        Loop
        
        End With
        
        Select Case PrintUsingWord
        Case cmsUseWord
            GenerateBookGroupPrintInWord
        Case cmsUseMSDatareport
            GenerateBookGroupPrint
        End Select
    
    End If
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
ErrorTrap:
    EndProgram
    
    
End Sub


Public Sub PrintBookGroupReportList(GroupList As String, PrintType As Long)
On Error GoTo ErrorTrap
Dim MinSQL As String, rstPrintReport As Recordset, rstPrintTable As Recordset
Dim PubDate As String, rs As Recordset, col As New Collection, lnum As Long, i As Long
Dim bSomethingPrinted As Boolean

    Screen.MousePointer = vbHourglass

    'first build the print table
    
    DelAllRows "tblPrintBookGroups"
    
    '
    'Driving Recset to get all data
    '
    Select Case PrintType
    Case 1 'publishers for report list
        PubDate = "#" & Month(Now) & "/01/" & year(Now) & "#"
        
        MinSQL = "SELECT ID, " & _
                        "LastName, FirstName, MiddleName, " & _
                        "tblBookGroups.GroupNo, tblBookGroups.GroupName " & _
                 "FROM ((tblNameAddress INNER JOIN tblBookGroupMembers ON tblNameAddress.ID = tblBookGroupMembers.PersonID) " & _
                        "              INNER JOIN tblBookGroups ON tblBookGroupMembers.GroupNo = tblBookGroups.GroupNo) " & _
                            "          INNER JOIN tblPublisherDates ON tblNameAddress.ID = tblPublisherDates.PersonID " & _
                 "WHERE Active = TRUE " & _
                 " AND StartDate <= " & PubDate & _
                 " AND EndDate >= " & PubDate & _
                 " AND StartReason <> 2 " & _
                 "ORDER BY tblBookGroups.GroupName, " & _
                 "         LastName, FirstName, MiddleName "
    Case 2 'Has KM list
        MinSQL = "SELECT ID, " & _
                        "LastName, FirstName, MiddleName, " & _
                        "tblBookGroups.GroupNo, tblBookGroups.GroupName " & _
                 "FROM ((tblNameAddress INNER JOIN tblBookGroupMembers ON tblNameAddress.ID = tblBookGroupMembers.PersonID) " & _
                        "              INNER JOIN tblBookGroups ON tblBookGroupMembers.GroupNo = tblBookGroups.GroupNo) " & _
                            "          INNER JOIN tblTaskAndPerson ON tblTaskAndPerson.Person = tblNameAddress.ID " & _
                 "WHERE Active = TRUE " & _
                 " AND Task = 90 " & _
                 "ORDER BY tblBookGroups.GroupName, " & _
                 "         LastName, FirstName, MiddleName "
    
    End Select
             
    Set rstPrintReport = CMSDB.OpenRecordset(MinSQL, dbOpenDynaset)
    Set rstPrintTable = CMSDB.OpenRecordset("tblPrintBookGroups", dbOpenDynaset)
    
    If rstPrintReport.BOF Then
        MsgBox "Nothing to print.", vbOKOnly + vbExclamation, AppName
        Screen.MousePointer = vbNormal
        Exit Sub
    Else
        With rstPrintReport
        
        Do Until .EOF 'add new print rec to tblPrintBookGroups for each row found
        
            rstPrintTable.AddNew
                                    
            rstPrintTable!GroupName = !GroupName
            rstPrintTable!PersonName = !LastName & ", " & !FirstName & " " & !MiddleName
            rstPrintTable!GroupID = !GroupNo
            rstPrintTable!PersonID = !ID
            
            If CongregationMember.DoesRole(!ID, CInt(GlobalDefaultCong), 4, 4, 14) Then
                rstPrintTable!IsOverseer = "X"
            Else
                rstPrintTable!IsOverseer = ""
            End If
                
            If CongregationMember.DoesRole(!ID, CInt(GlobalDefaultCong), 4, 4, 15) Then
                rstPrintTable!IsAssistant = "X"
            Else
                rstPrintTable!IsAssistant = ""
            End If
                
            If CongregationMember.DoesRole(!ID, CInt(GlobalDefaultCong), 4, 4, 16) Then
                rstPrintTable!IsReader = "X"
            Else
                rstPrintTable!IsReader = ""
            End If
                
            If CongregationMember.DoesRole(!ID, CInt(GlobalDefaultCong), 4, 4, 17) Then
                rstPrintTable!IsPrayer = "X"
            Else
                rstPrintTable!IsPrayer = ""
            End If
                
            If CongregationMember.DoesRole(!ID, CInt(GlobalDefaultCong), 4, 4, 90) Then
                rstPrintTable!HasKM = "X"
                rstPrintTable!HasKM_Num = "1"
            Else
                rstPrintTable!HasKM = ""
                rstPrintTable!HasKM_Num = "0"
            End If
                                                   
            rstPrintTable.Update
            
            .MoveNext
        Loop
        
        End With
        
        If PrintUsingWord(False) = cmsUseWord Then
            Set rs = CMSDB.OpenRecordset("SELECT GroupNo FROM tblBookGroups " & _
                                         "WHERE GroupNo IN (" & GroupList & ") ", _
                                            dbOpenDynaset)
            With rs
            Do Until .BOF Or .EOF
                lnum = !GroupNo
                col.Add lnum
                .MoveNext
            Loop
            End With
            rs.Close
            Set rs = Nothing
            
            '
            'Now look at tblPrintBookGroups. For each required group,
            ' check that something exists on the table, then print to Word.
            ' Remember that the RptTool disconnects DAO each time, so must
            ' rebuild recordset each time round the loop!
            '
            For i = 1 To col.Count
                
                Set rs = CMSDB.OpenRecordset("SELECT GroupID " & _
                                             "FROM tblPrintBookGroups " & _
                                             "WHERE GroupID = " & col.Item(i), _
                                            dbOpenSnapshot)
                                                    
                If Not rs.BOF Then
                    GenerateBookGroupReportListPrintInWord col.Item(i), _
                                                           (i < col.Count), _
                                                           PrintType
                    bSomethingPrinted = True
                End If
                
            Next i
            
            If Not bSomethingPrinted Then
                MsgBox "Nothing to print", vbOKOnly + vbExclamation, AppName
            End If
            
            On Error Resume Next
            rs.Close
            Set rs = Nothing
            On Error GoTo ErrorTrap
                            
            Set col = Nothing
        Else
            MsgBox "Microsoft Word not installed", vbOKOnly + vbExclamation, AppName
        End If
    
    End If
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
ErrorTrap:
    EndProgram
    
    
End Sub


Public Function GenerateBookGroupPrint() As Boolean
Dim WorkSheetTopMargin As Single
Dim WorkSheetBottomMargin As Single
Dim WorkSheetLeftMargin As Single
Dim WorkSheetRightMargin As Single
Dim RptTitle As String

On Error GoTo ErrorTrap

    Screen.MousePointer = vbHourglass
    
    '
    'Arrange page margins before we close the db connection
    '
    WorkSheetTopMargin = 566.929 * (GlobalParms.GetValue("A4TopMargin", "NumFloat"))
    WorkSheetBottomMargin = 566.929 * (GlobalParms.GetValue("A4BottomMargin", "NumFloat"))
    WorkSheetLeftMargin = 566.929 * (GlobalParms.GetValue("A4LeftMargin", "NumFloat"))
    WorkSheetRightMargin = 566.929 * (GlobalParms.GetValue("A4RightMargin", "NumFloat"))

    SwitchOffDAO
    
    '
    'GENERAL ADO WARNING...
    ' If we refer to DataReport prior to 'Showing' it, thus opening new ADODB connection
    ' while DAO connection still open, we get funny results.. eg missing fields on report.
    '
        
    BookGroupDetails.TopMargin = WorkSheetTopMargin '<----- At this point, Initialize runs.
    BookGroupDetails.BottomMargin = WorkSheetBottomMargin
    BookGroupDetails.LeftMargin = WorkSheetLeftMargin
    BookGroupDetails.RightMargin = WorkSheetRightMargin
    
    Screen.MousePointer = vbNormal
    
    BookGroupDetails.Show vbModal
    
    SwitchOnDAO

    GenerateBookGroupPrint = True

    Exit Function
ErrorTrap:
    EndProgram
    
End Function
Public Function GenerateBookGroupPrintInWord() As Boolean

Dim reporter As MSWordReportingTool2.RptTool

On Error GoTo ErrorTrap

    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .ReportSQL = "SELECT GroupName, " & _
                 "       PersonName, " & _
                 "       IsOverseer, " & _
                 "       IsAssistant, " & _
                 "       IsReader, " & _
                 "       IsPrayer, " & _
                 "       HasKM " & _
                 "FROM   tblPrintBookGroups " & _
                 "   GROUP BY GroupName, " & _
                 "            PersonName, " & _
                 "            IsOverseer, " & _
                 "            IsAssistant, " & _
                 "            IsReader, " & _
                 "            IsPrayer, " & _
                 "            HasKM " & _
                 "   ORDER BY 1,2 "
                    

    .ReportTitle = "Book Group Details"
    
    .SaveDoc = True
    .DocPath = gsDocsDirectory & "\" & "Field Service Group Details " & _
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
    .ShowPageNumber = True
    .GroupingColumn = 1
    .HideWordWhileBuilding = True
    
'    If FormIsOpen("frmCongregationSetUp") Then
        .ShowProgress = True
'    Else
'        .ShowProgress = False
'    End If
        
    .ShowProgress = True
    .AdditionalReportHeading = Format$(Now, "mmmm d") & _
                                GetLettersForOrdinalNumber(CLng(Day(Now))) & _
                                " " & CStr(year(Now))
    .AdditionalReportHeadingBold = False
    .AdditionalReportHeadingFontSize = 12
    
    .AddTableColumnAttribute "Group", 50, , , "Times New Roman", , 10, 10, True, True, , , True
    .AddTableColumnAttribute "Person", 50, , , "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Overseer", 18, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Asst", 18, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Reader", 18, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Prayer", 18, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Has KM", 18, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    
    .PageFormat = cmsPortrait
    
    .GenerateReport
       
    End With
    
    SwitchOnDAO
    
    Screen.MousePointer = vbNormal

    Exit Function
ErrorTrap:
    EndProgram
    
    
End Function
Public Function GenerateBookGroupReportListPrintInWord(lGrpID As Long, _
                                                       KeepWordHidden As Boolean, _
                                                       ListType As Long) As Boolean

Dim reporter As MSWordReportingTool2.RptTool

On Error GoTo ErrorTrap

    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .ReportSQL = "SELECT GroupName, " & _
                 "       PersonName, ' ' , ' ' , ' ' , ' ' , ' ' , ' ' , " & _
                 "                   ' ' , ' ' , ' ' , ' ' , ' ' , ' '  " & _
                 "FROM   tblPrintBookGroups " & _
                 "WHERE GroupID = " & lGrpID & _
                 "   ORDER BY 2 "

    Select Case ListType
    Case 1
        .ReportTitle = "Field Service Group Report Checklist"
    Case 2
        .ReportTitle = "Field Service Group Kingdom Ministry Checklist"
    Case Else
        .ReportTitle = "WHAT'S THIS REPORT CALLED?!"
    End Select
    
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
    .ShowPageNumber = True
    .GroupingColumn = 1
    .PageFormat = cmsLandscape
    .HideWordWhileBuilding = True
    .HideWordWhenDone = KeepWordHidden
            
    .ShowProgress = True
    
    .AddTableColumnAttribute "Group", 50, , , "Times New Roman", , 10, 10, True, True, , , True
    .AddTableColumnAttribute "Person", 50, , , "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Sep", 13, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Oct", 13, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Nov", 13, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Dec", 13, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Jan", 13, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Feb", 13, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Mar", 13, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Apr", 13, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "May", 13, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Jun", 13, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Jul", 13, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
    .AddTableColumnAttribute "Aug", 13, cmsCentreTop, cmsCentreTop, "Times New Roman", , 10, 10, True, True
        
            
    .GenerateReport
       
    End With
    
    SwitchOnDAO
    
    Screen.MousePointer = vbNormal

    Exit Function
ErrorTrap:
    EndProgram
    
    
End Function

Public Function ValidateEmailAddress(EmailAddress As String, _
                                     Optional MandatoryValue As Boolean = False) As Boolean

On Error GoTo ErrorTrap

    If EmailAddress <> "" Then
        If InStr(1, EmailAddress, "@") = 0 Or _
           InStr(1, EmailAddress, "@") = 1 Or _
           InStr(1, EmailAddress, "@") = Len(EmailAddress) Then
                ValidateEmailAddress = False
                Exit Function
        End If
        
        If InStr(1, EmailAddress, ".") = 0 Or _
           InStr(1, EmailAddress, ".") = 1 Or _
           InStr(1, EmailAddress, ".") = Len(EmailAddress) Then
                ValidateEmailAddress = False
                Exit Function
        End If

        If InStr(1, EmailAddress, " ") > 0 Then
            ValidateEmailAddress = False
            Exit Function
        End If
    Else
        If MandatoryValue Then
            ValidateEmailAddress = False
            Exit Function
        End If
    End If
           
    ValidateEmailAddress = True

    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Function PrintUsingWord(Optional AskUser As Boolean = True) As cmsPrintUsingWord

On Error GoTo ErrorTrap

    If AskUser Then
    
        If GlobalParms.GetValue("UseWordForReports", "TrueFalse") Then
            If IsOfficeAppPresent(cmsWord) Then
                Select Case MsgBox("Do you want to use Microsoft Word?", vbYesNoCancel + vbQuestion, AppName)
                Case vbYes
                    PrintUsingWord = cmsUseWord
                Case vbNo
                    PrintUsingWord = cmsUseMSDatareport
                Case Else
                    PrintUsingWord = cmsDontPrint
                End Select
            Else
                Select Case MsgBox("Microsoft Word is not installed correctly. " & _
                           "Microsoft Data-Report will be used.", vbQuestion + vbYesNo, _
                           AppName)
                Case vbYes
                    PrintUsingWord = cmsUseMSDatareport
                Case vbNo
                    PrintUsingWord = cmsDontPrint
                End Select
            End If
        Else
            Select Case MsgBox("Use Microsoft Data-Report?", vbQuestion + vbYesNo, _
                       AppName)
            Case vbYes
                PrintUsingWord = cmsUseMSDatareport
            Case vbNo
                PrintUsingWord = cmsDontPrint
            End Select
        End If
        
    Else
        
        If GlobalParms.GetValue("UseWordForReports", "TrueFalse") Then
            If IsOfficeAppPresent(cmsWord) Then
                PrintUsingWord = cmsUseWord
            Else
                PrintUsingWord = cmsUseMSDatareport
            End If
        Else
            PrintUsingWord = cmsUseMSDatareport
        End If
        
    End If

    Exit Function
ErrorTrap:
    EndProgram
    
End Function



Public Function DealWithUnknownPersonName(TheName As String, IfUnknown As String) As String
    If Left(TheName, 1) <> "?" Then
        DealWithUnknownPersonName = TheName
    Else
        DealWithUnknownPersonName = IfUnknown
    End If
End Function
