Attribute VB_Name = "basNotMoreCode"
Option Explicit

Public Function TotalForReportField(ByVal TheYear As Long, _
                                    ByVal TheMonth As Long, _
                                    ByVal ReportField As String, _
                                    ByVal PublisherType As Long) As Double
On Error GoTo ErrorTrap
Dim TheString As String, rstRecSet As Recordset, DateString As String
Dim EndOfYear As String, StartOfYear As String, TheTable As String
Dim AnAmount As Double, sPieceOfSQL As String

    '
    'NB: tblMinReports is based on SERVICE YEARS (except for ActualMinPeriod and
    '      SocietyReportingPeriod which are calendar dates),
    '    whereas tblPublisherDates is based on NORMAL YEARS.
    '
    
    Select Case PublisherType
    Case 1 'pub
        TheTable = "tblPublisherDates"
        'exclude min done in previous cong for those that have moved in,
        ' but only for figures sent to branch
        sPieceOfSQL = " AND StartReason <> 2 "
    Case 2 'aux pio
        TheTable = "tblAuxPioDates"
        sPieceOfSQL = ""
    Case 3 'reg pio
        TheTable = "tblRegPioDates"
        sPieceOfSQL = ""
    Case 4 'spec pio
        TheTable = "tblSpecPioDates"
        sPieceOfSQL = ""
    End Select
    
    If TheMonth > 0 Then
        DateString = CStr(TheMonth) & "/01/" & _
                        CStr(ConvertServiceYearToNormalYear(CDate("01/" & CStr(TheMonth) & _
                                                                                "/" & CStr(TheYear))))
    
        If frmCongStats.optMinInMonth = True Then
            TheString = "SELECT SUM(" & ReportField & ") AS TheSum " & _
                        "FROM tblMinReports " & _
                        "     INNER JOIN " & TheTable & _
                        "     ON  tblMinReports.PersonID = " & TheTable & ".PersonID " & _
                        "WHERE tblMinReports.MinistryDoneInMonth = " & TheMonth & _
                        "  AND tblMinReports.MinistryDoneInYear = " & TheYear & _
                        "  AND " & TheTable & ".StartDate <= #" & DateString & _
                        "# AND " & TheTable & ".EndDate >= #" & DateString & "#"
        Else
            TheString = "SELECT SUM(" & ReportField & ") AS TheSum " & _
                        "FROM tblMinReports " & _
                        "     INNER JOIN " & TheTable & _
                        "     ON  tblMinReports.PersonID = " & TheTable & ".PersonID " & _
                        "WHERE tblMinReports.SocietyReportingMonth = " & TheMonth & _
                        "  AND tblMinReports.SocietyReportingYear = " & TheYear & _
                        "  AND " & TheTable & ".StartDate <= #" & DateString & _
                        "# AND " & TheTable & ".EndDate >= #" & DateString & "#"
                        
            'exclude min done in previous cong for those that have moved in,
            ' but only for figures sent to branch
            TheString = TheString & sPieceOfSQL
        End If
    Else
        StartOfYear = "09/01/" & _
                        CStr(ConvertServiceYearToNormalYear(CDate("01/09/" & CStr(TheYear))))
        EndOfYear = "08/01/" & _
                        CStr(ConvertServiceYearToNormalYear(CDate("01/08/" & CStr(TheYear))))
        
        TheString = "SELECT SUM(" & ReportField & ") AS TheSum " & _
                    "FROM tblMinReports " & _
                    "     INNER JOIN " & TheTable & _
                    "     ON  tblMinReports.PersonID = " & TheTable & ".PersonID " & _
                    "WHERE tblMinReports.MinistryDoneInYear = " & TheYear & _
                    "  AND " & TheTable & ".StartDate <= #" & EndOfYear & _
                    "# AND " & TheTable & ".EndDate >= #" & StartOfYear & "#" & _
                    "  AND " & TheTable & ".StartDate <= tblMinReports.ActualMinPeriod " & _
                    "  AND " & TheTable & ".EndDate >= tblMinReports.ActualMinPeriod "
    End If
                      
    Set rstRecSet = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)
    
    With rstRecSet
    
    If Not .BOF Then
        If Not IsNull(!TheSum) Then
            TotalForReportField = !TheSum
        Else
            TotalForReportField = 0
        End If
    Else
        TotalForReportField = 0
    End If
    
    End With
    
    '
    'Now add in decimal part from any previous months....
    '
    If ReportField = "NoHours" And PublisherType = 1 Then
'        If TheMonth > 0 And frmCongStats.optMinRepToSociety Then
        If TheMonth > 0 Then
            If frmCongStats.optMinRepToSociety = True Then
                TheString = "SELECT SUM(NoHours) AS TheSum " & _
                            "FROM tblMinReports " & _
                            "WHERE tblMinReports.SocietyReportingPeriod < #" & DateString & _
                              "#"
            Else
                TheString = "SELECT SUM(NoHours) AS TheSum " & _
                            "FROM tblMinReports " & _
                            "WHERE tblMinReports.ActualMinPeriod < #" & DateString & _
                              "#"
            End If
                        
            Set rstRecSet = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)
            
            With rstRecSet
            
            If Not .BOF Then
                If Not IsNull(!TheSum) Then
                    AnAmount = GetFractionPart(!TheSum)
                Else
                    AnAmount = 0
                End If
            Else
                AnAmount = 0
            End If
            
            TotalForReportField = Fix(TotalForReportField + AnAmount)
                       
            End With
        Else
            TheString = "SELECT SUM(NoHours) AS TheSum " & _
                        "FROM tblMinReports " & _
                        "WHERE tblMinReports.MinistryDoneInYear < " & TheYear

            Set rstRecSet = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)

            With rstRecSet

            If Not .BOF Then
                If Not IsNull(!TheSum) Then
                    AnAmount = GetFractionPart(!TheSum)
                Else
                    AnAmount = 0
                End If
            Else
                AnAmount = 0
            End If

            TotalForReportField = Fix(TotalForReportField + AnAmount)

            End With
                
        End If
    Else
        TotalForReportField = Fix(TotalForReportField)
    End If
            

    Exit Function
ErrorTrap:
    EndProgram

    
End Function

Public Function FormatMMYYYY(ByVal TheString As String) As String
On Error GoTo ErrorTrap

Dim TempString As String, i As Long, j As Long, k As String
Dim TempString2 As String
Dim TempString3 As String
    '
    'Pass in a user-entered date, function returns date in MM/YYYY format
    '
    
    '
    'First find number of "/" in string. If > 1, remove everything upto and
    ' including second to last "/"
    '
    i = StringCount(TheString, "/")
    
    If i > 1 Then
        j = InStrRev(TheString, "/")
        j = InStrRev(TheString, "/", j - 1)
        TempString = Right(TheString, Len(TheString) - j)
    Else
        TempString = TheString
    End If
    
    '
    'If there is NO "/", then....
    '
    If i = 0 Then
        
        TempString = RemoveNonNumerics(TempString)
               
        Select Case True
        Case Len(TempString) = 0
        'Nothing in string. Use current date.
            FormatMMYYYY = Format(Month(Now), "00") & "/" & Format(year(Now), "0000")
        Case CLng(TempString) >= 1 And CLng(TempString) <= 12
        'We'll say this is a month value. Let's add current year.
            FormatMMYYYY = Format(TempString, "00") & "/" & Format(year(Now), "0000")
        Case Len(TempString) = 3
        'Assume this is a year value
            FormatMMYYYY = Format(Month(Now), "00") & "/" & Left(year(Now), 1) & TempString
        Case Len(TempString) = 2
        'Assume this is a year value
            FormatMMYYYY = Format(Month(Now), "00") & "/" & Left(year(Now), 2) & TempString
        Case Len(TempString) = 1
        'Assume this is a year value
            FormatMMYYYY = Format(Month(Now), "00") & "/" & Left(year(Now), 3) & TempString
        Case CLng(Left$(TempString, 2)) >= 1 And CLng(Left$(TempString, 2)) <= 12
            FormatMMYYYY = Format(CLng(Left$(TempString, 2)), "00") & "/" & Format(year(Now), "0000")
        Case Else
            FormatMMYYYY = Format(Month(Now), "00") & "/" & Format(Left(TempString, 4), "0000")
        End Select
    Else
        '
        'At this point, we have a string with ONE "/" in it
        'First, find position of the "/"..
        '
        i = InStr(1, TempString, "/")
        
        'Now check string to left of "/"
        TempString2 = Left(TempString, i - 1)
        TempString2 = RemoveNonNumerics(TempString2)
        If Len(TempString2) > 2 Then
            If CLng(Left(TempString2, 2)) >= 1 And CLng(Left(TempString2, 2)) <= 12 Then
                TempString2 = Format(Left(TempString2, 2), "00")
            Else
                TempString2 = "12"
            End If
        ElseIf Len(TempString2) = 0 Then
            TempString2 = Format(Month(Now), "00")
        Else
            If CLng(TempString2) >= 1 And CLng(TempString2) <= 12 Then
                TempString2 = Format(TempString2, "00")
            Else
                TempString2 = Format(Month(Now), "00")
            End If
        End If
            
        'Now check string to right of "/"
        TempString3 = Right(TempString, Len(TempString) - i)
        TempString3 = RemoveNonNumerics(TempString3)
        If Len(TempString3) = 4 Then
        Else
            If Len(TempString3) = 3 Then
                TempString3 = Left(year(Now), 1) & TempString3
            ElseIf Len(TempString3) = 2 Then
                TempString3 = Left(year(Now), 2) & TempString3
            ElseIf Len(TempString3) = 1 Then
                TempString3 = Left(year(Now), 3) & TempString3
            ElseIf Len(TempString3) = 0 Then
                TempString3 = year(Now)
            Else
                TempString3 = Left(TempString3, 4)
            End If
        End If
        
        FormatMMYYYY = TempString2 & "/" & TempString3
        
    End If
        
    Exit Function
ErrorTrap:
    EndProgram

    
End Function

Public Function RemoveNonNumerics(ByVal TheString As String) As String
On Error GoTo ErrorTrap

Dim TempString As String, i As Long, j As Long, k As String
Dim TempString2 As String

    TempString = TheString
        
    If Not IsNumber(TempString, False, False, False) Then
        'There are non-numerics in the string. Lose 'em.
        TempString2 = TempString
        For j = 1 To Len(TempString2)
            k = Mid$(TempString2, j, 1)
            If Not IsNumber(k, False, False, False) Then
                TempString = DelSubstr(TempString, k, False)
            End If
        Next j
    End If
    
    RemoveNonNumerics = TempString

    Exit Function
ErrorTrap:
    EndProgram

    
End Function

Public Function FormatDateAsUKDateString(TheDate As Date)

    FormatDateAsUKDateString = Format(TheDate, "dd/mm/yyyy")

End Function

Public Function GetSocietyReportingPeriodMMYYYY(ByVal CurrentDate As Date, _
                                                 Optional TheMM, _
                                                 Optional TheYYYY) As String
On Error GoTo ErrorTrap
Dim TheMMBit As String, TheYYYYBit As String

    '
    'If TheMM and TheYYYY are provided, format these as MM/YYYY
    ' Otherwise....
    'If current day of month is < 6 then we are sending report for prev month.
    ' Otherwise, report is for this month.
    '
    
    If IsMissing(TheMM) Then
        If Day(CurrentDate) > GlobalParms.GetValue("DayOfMonthForReportToSociety", "NumVal") Then
            'This month
            TheMMBit = Format(Month(CurrentDate), "00")
            TheYYYYBit = Format(ServiceYear(CurrentDate), "0000")
            GetSocietyReportingPeriodMMYYYY = TheMMBit & "/" & TheYYYYBit
        Else
            'Last month
            TheMMBit = Format(Month(DateAdd("m", -1, CurrentDate)), "00")
            TheYYYYBit = Format(ServiceYear(DateAdd("m", -1, CurrentDate)), "00")
            GetSocietyReportingPeriodMMYYYY = TheMMBit & "/" & TheYYYYBit
        End If
    Else
        If IsMissing(TheYYYY) Then
            TheYYYY = CLng(year(Now))
            TheMMBit = Format(TheMM, "00")
            TheYYYYBit = Format(TheYYYY, "00")
            GetSocietyReportingPeriodMMYYYY = TheMMBit & "/" & TheYYYYBit
        End If
    End If
    
    Exit Function
ErrorTrap:
    EndProgram
    

End Function

Public Function GetLastSocietyReportingPeriodMMYYYY(ByVal CurrentDate As Date) As String
On Error GoTo ErrorTrap
Dim TheMMBit As String, TheYYYYBit As String

    '
    'If TheMM and TheYYYY are provided, format these as MM/YYYY
    ' Otherwise....
    'If current day of month is < 6 then we are sending report for prev month.
    ' Otherwise, report is for this month.
    '
    
    If Day(CurrentDate) > GlobalParms.GetValue("DayOfMonthForReportToSociety", "NumVal") Then
        TheMMBit = Format(Month(DateAdd("m", -1, CurrentDate)), "00")
        TheYYYYBit = Format(ServiceYear(DateAdd("m", -1, CurrentDate)), "00")
        GetLastSocietyReportingPeriodMMYYYY = TheMMBit & "/" & TheYYYYBit
    Else
        TheMMBit = Format(Month(DateAdd("m", -2, CurrentDate)), "00")
        TheYYYYBit = Format(ServiceYear(DateAdd("m", -2, CurrentDate)), "00")
        GetLastSocietyReportingPeriodMMYYYY = TheMMBit & "/" & TheYYYYBit
    End If
    
    Exit Function
ErrorTrap:
    EndProgram
    

End Function


Public Sub RefreshMinistryStatusForPerson(ByVal PersonID As Long)
On Error GoTo ErrorTrap

    'start from scratch for this person...
    DeleteSomeRows "tblIrregularPubs", "PersonID = ", PersonID
    
    WriteToLogFile "frmMinReportyEntry.SaveReport ------------- 0010"
    
    DeleteMissingReportsForPerson PersonID
    
    WriteToLogFile "frmMinReportyEntry.SaveReport ------------- 0014"
    
    PutAllMissingReportsIntoTable "01/09/2003", "01/12/9999", PersonID
    
    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Public Sub RemoveReportsOutsidePubPeriod(ByVal PersonID As Long)
On Error GoTo ErrorTrap
Dim rstTemp As Recordset, str As String, str2 As String

    Set rstTemp = CMSDB.OpenRecordset("SELECT StartDate, " & _
                                      "       EndDate, " & _
                                      "       SeqNum " & _
                                      "FROM tblPublisherDates " & _
                                      "WHERE PersonID = " & PersonID _
                                      , dbOpenDynaset)
                                      
    With rstTemp
    
    str = ""
    
    Do Until .BOF Or .EOF
    
        str = str & " AND ActualMinPeriod NOT BETWEEN #" & Format(!StartDate, "mm/dd/yyyy") & "# " & _
                        "AND #" & Format(!EndDate, "mm/dd/yyyy") & "# "
                        
        .MoveNext
    
    Loop
    
    End With
    
    If str <> "" Then
        
        str2 = "SELECT 1 FROM tblMinReports WHERE PersonID = " & PersonID & str
        
        Set rstTemp = CMSDB.OpenRecordset(str2, dbOpenDynaset)
        
        If Not rstTemp.BOF Then
            If MsgBox("Field service has been recorded for " & CongregationMember.NameWithMiddleInitial(PersonID) & _
                    " in months outside of the periods during which " & IIf(CongregationMember.GetGender(PersonID) = Male, "he ", "she ") & _
                    "is a publisher. Do you want to delete this ministry?", vbYesNo + vbQuestion + vbDefaultButton2, AppName) = vbYes Then
                    
                CMSDB.Execute "DELETE FROM tblMinReports WHERE PersonID = " & PersonID & str
            
            End If
        End If
    End If
   
                                      
GetOut:

    rstTemp.Close
    Set rstTemp = Nothing
    
    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Public Sub DeleteAllMinistryForPerson(ByVal PersonID As Long, _
                                      ByVal StartDate As Date, _
                                      ByVal EndDate As Date)
On Error GoTo ErrorTrap

    'start from scratch for this person...
    DeleteSomeRows "tblIrregularPubs", "PersonID = ", PersonID
    DeleteMissingReportsForPerson PersonID
    
    'now delete reports....
    CMSDB.Execute ("DELETE FROM tblMinReports " & _
                  "WHERE PersonID = " & PersonID & _
                  " AND ActualMinPeriod BETWEEN #" & Format(StartDate, "mm/dd/yyyy") & "# " & _
                  "                       AND #" & Format(EndDate, "mm/dd/yyyy") & "# ")

    PutAllMissingReportsIntoTable "01/09/2003", "01/12/9999", PersonID
    
    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Public Sub CleanUpDocs()
On Error Resume Next
Dim fso As FileSystemObject, fil As File, lDays As Long

    lDays = HandleNull(GlobalParms.GetValue("DocsDeleteDays", "NumVal"))
    
    If lDays = 0 Then Exit Sub

    Set fso = New FileSystemObject
    
    For Each fil In fso.GetFolder(gsDocsDirectory).Files
        If DateDiff("d", fil.DateCreated, Now) > lDays Then
            fil.Delete True
        End If
    Next
    
    Set fso = Nothing
    
End Sub


