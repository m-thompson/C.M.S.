Attribute VB_Name = "basGeneral4"
Option Explicit

Public Sub PopulateComboWithCurrentPublishers(TheControl As Control, _
                                              MinType As MinistryType, _
                                              Optional AsOfDate_UK, _
                                              Optional BookGroup, _
                                              Optional MaintainSelection As Boolean = True)
On Error GoTo ErrorTrap
Dim TheTable As String, psDateString_US As String, SQLStr As String
Dim BookGroupFilter As String, BookGroupFilter2 As String

    If Not (TypeOf TheControl Is ListBox) And _
       Not (TypeOf TheControl Is ComboBox) Then
       
        EndProgram "Wrong type of control (basGeneral4.PopulateComboWithCurrentPublishers)"
    End If
    
    Select Case MinType
    Case IsPublisher
        TheTable = "tblPublisherDates"
    Case IsAuxPio
        TheTable = "tblAuxPioDates"
    Case IsRegPio
        TheTable = "tblRegPioDates"
    Case IsSpecPio
        TheTable = "tblSpecPioDates"
    End Select
    
    If Not IsMissing(AsOfDate_UK) Then
        psDateString_US = Format$(AsOfDate_UK, "mm/dd/yyyy")
        SQLStr = "AND StartDate <= #" & psDateString_US & "# " & _
                 "AND EndDate >= #" & psDateString_US & "# "
    Else
        SQLStr = ""
    End If
    
    If IsMissing(BookGroup) Then
        BookGroup = 0
    End If
    
    If BookGroup > 0 Then
        BookGroupFilter = " AND GroupNo = " & BookGroup
        
        BookGroupFilter2 = "INNER JOIN tblBookGroupMembers ON " & _
                           " tblBookGroupMembers.PersonID = tblNameAddress.ID "
    Else
        BookGroupFilter = ""
        BookGroupFilter2 = ""
    End If
    
    HandleListBox.PopulateListBox TheControl, _
        "SELECT DISTINCTROW tblNameAddress.ID, " & _
        "tblNameAddress.FirstName & ' ' & tblNameAddress.MiddleName, " & _
        "tblNameAddress.LastName " & _
        "FROM (tblNameAddress INNER JOIN " & TheTable & " ON " & _
        "tblNameAddress.ID =  " & TheTable & ".PersonID) " & _
        BookGroupFilter2 & _
        " WHERE Active = TRUE " & SQLStr & BookGroupFilter & _
        " ORDER BY tblNameAddress.LastName, tblNameAddress.FirstName" _
        , CMSDB, 0, ", ", MaintainSelection, 2, 1
        
    Exit Sub
ErrorTrap:
    EndProgram

    
End Sub

Public Sub PutAllMissingReportsIntoTable(StartDate_UK As String, _
                                 EndDate_UK As String, _
                                 Optional ThePerson _
                                 )
                                 
On Error GoTo ErrorTrap
Dim psDateString_US As String, psTheString As String, rstRecSet As Recordset
Dim psDateString_UK As String, TheEndDate_UK As String, SQLPiece As String
Dim NextGrpID As Long, rstRecSet2 As Recordset, lbZeroReport As Boolean
Dim rstRecSet3 As Recordset, psTheString3 As String

    
    
    If IsMissing(ThePerson) Then
        SQLPiece = ""
    Else
        SQLPiece = "AND PersonID = " & ThePerson
    End If
        
    '
    'Gets all missing reports for date range. Dates supplied are NOT Service Year
    ' format.
    'Dates should be supplied as "01/nn/nnnn" - including End-Date, for which
    ' that whole month is included.
    '
    psDateString_UK = StartDate_UK
    
    'If EndDate_UK is future, only check for missing reports up to last
    ' complete reporting month...
    If DateDiff("m", Now, CDate(EndDate_UK)) >= -1 Then 'EndDate is future
        If Day(Now) >= GlobalParms.GetValue("DayOfMonthForCongStats", "NumVal") Then
            TheEndDate_UK = Format$((DateAdd("m", -1, Now)), "dd/mm/yyyy")
        Else
            TheEndDate_UK = Format$((DateAdd("m", -2, Now)), "dd/mm/yyyy")
        End If
    Else
        TheEndDate_UK = EndDate_UK
    End If
    
    'make dates begin with 1st of month to match format on tblMinReports and tblPublisherDates
    
    TheEndDate_UK = CStr("01/" & _
                    Format$(Month(CDate(TheEndDate_UK)), "00") & "/" & _
                    Format$(year(CDate(TheEndDate_UK)), "0000"))
                    
    psDateString_UK = CStr("01/" & _
                    Format$(Month(CDate(psDateString_UK)), "00") & "/" & _
                    Format$(year(CDate(psDateString_UK)), "0000"))
    
    psDateString_US = Format$(psDateString_UK, "mm/dd/yyyy")
    
    Set rstRecSet2 = CMSDB.OpenRecordset("tblMinReports", dbOpenDynaset)
    
    Do Until Format$(psDateString_UK, "yyyymmdd") > Format$(TheEndDate_UK, "yyyymmdd")
        
        'put all pubs for whom there's no report into the recset
        psTheString = "SELECT a.PersonID " & _
                      "FROM tblPublisherDates a " & _
                      "WHERE StartDate <= #" & psDateString_US & "# " & _
                      "AND EndDate >= #" & psDateString_US & "# " & _
                      SQLPiece & _
                      " AND (NOT EXISTS " & _
                      "            (SELECT 1 " & _
                                  " FROM tblMinReports b " & _
                                  " WHERE b.ActualMinPeriod = #" & psDateString_US & "# " & _
                                  " AND b.PersonID = a.PersonID )" & _
                      " OR EXISTS " & _
                      "            (SELECT 1 " & _
                                  " FROM tblMinReports c " & _
                                  " WHERE c.ActualMinPeriod = #" & psDateString_US & "# " & _
                                  " AND c.PersonID = a.PersonID " & _
                                  " AND NoHours = 0))"
                                  
        Set rstRecSet = CMSDB.OpenRecordset(psTheString, dbOpenDynaset)
        
        'Now put all pubs found with no report into tblMissingReports
        'If report has zero hours, it's still inserted to tblMissingReports but
        ' with ZeroReport=TRUE
        '
        
        With rstRecSet
        
        
        If Not .BOF Then
        
            Do Until .EOF
                
                NextGrpID = DetermineNextMissingRptGrpID(!PersonID, psDateString_UK)
                
                With rstRecSet2
                
                .FindFirst "PersonID = " & rstRecSet!PersonID & _
                           " AND ActualMinPeriod = #" & psDateString_US & "#" & _
                           " AND NoHours = 0"
                           
                If .NoMatch Then 'no zero report
                    lbZeroReport = False
                Else
                    lbZeroReport = True
                End If
                    
                End With

                On Error Resume Next 'In case duplicate insert
                               
                CMSDB.Execute "INSERT INTO tblMissingReports " & _
                                  "(PersonID, ServiceYear, ServiceMonth, ActualMinDate, MissingReportGroupID, ZeroReport) " & _
                                  "VALUES (" & !PersonID & ", " & _
                                         ServiceYear(CDate(psDateString_UK)) & ", " & _
                                         CLng(Mid$(psDateString_UK, 4, 2)) & ", #" & _
                                         psDateString_US & "#, " & _
                                         NextGrpID & ", " & lbZeroReport & ")"
                                         
                On Error GoTo ErrorTrap
                
                'now set the GrpID of any subsequent contiguous missing reports
                ' to that of above
                psTheString3 = "SELECT MissingReportGroupID " & _
                              "FROM tblMissingReports " & _
                              "WHERE PersonID = " & !PersonID & _
                              " AND ActualMinDate = #" & _
                                Format$(DateAdd("m", 1, psDateString_UK), "mm/dd/yyyy") & "#"
            
                Set rstRecSet3 = CMSDB.OpenRecordset(psTheString3, dbOpenSnapshot)
                  
                With rstRecSet3
                If Not .BOF Then
                    CMSDB.Execute "UPDATE tblMissingReports " & _
                                  "SET MissingReportGroupID = " & NextGrpID & _
                                  " WHERE MissingReportGroupID = " & !MissingReportGroupID
                                
                End If
                End With
                
                DetermineInactivePublishers NextGrpID
                
                DetermineIrregularPublishers !PersonID, _
                                              CDate(psDateString_UK), _
                                              True
                
                .MoveNext
            Loop
            
            rstRecSet3.Close
            
        End If
        
        End With
    
        'increment date by one month
        psDateString_UK = CStr(DateAdd("m", 1, CDate(psDateString_UK)))
        
        psDateString_US = Format$(psDateString_UK, "mm/dd/yyyy")
   
    Loop
    
    
    
    Exit Sub
ErrorTrap:
    EndProgram

    
End Sub

Public Function DetermineNextMissingRptGrpID(ThePerson As Long, CurrentMinDate_UK As String) As Long
                                 
On Error GoTo ErrorTrap
Dim psTheString As String, rstRecSet As Recordset, TheDate_US As String

    '
    'Determine the Missing Report Group ID from previous month for supplied publisher.
    ' If no missing report last month, then try next month
    '

    psTheString = "SELECT MissingReportGroupID " & _
                  "FROM tblMissingReports " & _
                  "WHERE PersonID = " & ThePerson & _
                  " AND ActualMinDate = #" & _
                    Format$(DateAdd("m", -1, CurrentMinDate_UK), "mm/dd/yyyy") & "#"

    Set rstRecSet = CMSDB.OpenRecordset(psTheString, dbOpenSnapshot)
      
    With rstRecSet
    If Not .BOF Then
        'there was a missing report last month, so use same grp id
        DetermineNextMissingRptGrpID = !MissingReportGroupID
        Exit Function
    Else
        'no missing report last month, so try next month....
    End If
    End With
    
    '
    'Determine the Missing Report Group ID from next month for supplied publisher.
    ' If no missing report next month, then create a new Group ID as MAX + 1
    '

    psTheString = "SELECT MissingReportGroupID " & _
                  "FROM tblMissingReports " & _
                  "WHERE PersonID = " & ThePerson & _
                  " AND ActualMinDate = #" & _
                    Format$(DateAdd("m", 1, CurrentMinDate_UK), "mm/dd/yyyy") & "#"

    Set rstRecSet = CMSDB.OpenRecordset(psTheString, dbOpenSnapshot)
      
    With rstRecSet
    If Not .BOF Then
        'there is a missing report next month, so use same grp id
        DetermineNextMissingRptGrpID = !MissingReportGroupID
    Else
        'no missing report next month, so get a new grp id
        DetermineNextMissingRptGrpID = GetNextMissingReportGroupID
    End If
    End With
    
    rstRecSet.Close
 
    Exit Function
ErrorTrap:
    EndProgram

    
End Function
Public Sub DealWithSpecialPioneerHours()
                                 
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
Dim rs As Recordset, dtMaxDate As Date, dtTempDate As Date
Dim str As String, sSpecPioHrs As String
                          
    Set rs = CMSDB.OpenRecordset("SELECT MAX(ActualMinPeriod) AS MaxDate " & _
                                 "FROM tblMinreports ", dbOpenForwardOnly)
    
    If Not IsNull(rs!MaxDate) Then
        dtMaxDate = rs!MaxDate
        
        Set rs = CMSDB.OpenRecordset("SELECT PersonID, StartDate " & _
                                 "FROM tblSpecPioDates ", dbOpenForwardOnly)
        
        If Not rs.BOF Then
            On Error Resume Next
            sSpecPioHrs = GlobalParms.GetValue("SpecialPioHours", "NumVal")
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
                                                "0, 0, " & sSpecPioHrs & ", 0, 0, 0, '', #" & _
                                                Format(dtTempDate, "mm/dd/yyyy") & "#, #" & _
                                                Format(dtTempDate, "mm/dd/yyyy") & "#,'',0)")
                                                
                    
                    dtTempDate = DateAdd("m", 1, dtTempDate)
                Loop
                rs.MoveNext
            Loop
            On Error GoTo ErrorTrap
        End If
        
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Public Sub DealWithLongTermInactive()
                                 
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
Dim rs As Recordset, dtTempDate As Date
Dim str As String
                          
    Set rs = CMSDB.OpenRecordset("SELECT Person " & _
                             "FROM tbltaskAndPerson " & _
                             "WHERE Task = 95", dbOpenForwardOnly) 'all longterm inactive
    
    If Not rs.BOF Then
    
        dtTempDate = CurrentReportMonth
        dtTempDate = Format(DateAdd("m", -1, Now), "dd/mm/yyyy")
        dtTempDate = CDate("01/" & Month(dtTempDate) & "/" & year(dtTempDate))
        
        On Error Resume Next 'ignore duplicate inserts (lazy guy....)
        Do Until rs.EOF 'for each inactive pub
            str = "INSERT INTO tblPubRecCardRowPrinted " & _
                    "VALUES (" & rs!Person & ", #" & _
                                Format(dtTempDate, "mm/dd/yyyy") & "#, FALSE)"
            
            CMSDB.Execute (str)
                                        
            CMSDB.Execute ("INSERT INTO tblMinReports " & _
                            "VALUES (" & rs!Person & ", " & _
                                        Month(dtTempDate) & ", " & _
                                        year(ConvertNormalDateToServiceDate(dtTempDate)) & ", " & _
                                        Month(dtTempDate) & ", " & _
                                        year(ConvertNormalDateToServiceDate(dtTempDate)) & ", " & _
                                        "0, 0, 0, 0, 0, 0, 'No Report', #" & _
                                        Format(dtTempDate, "mm/dd/yyyy") & "#, #" & _
                                        Format(dtTempDate, "mm/dd/yyyy") & "#,'',0)")
                                        
            
            rs.MoveNext 'next pub
        Loop
        
        On Error GoTo ErrorTrap
        
    End If
        
    rs.Close
    Set rs = Nothing
    
    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Public Function GetNextMissingReportGroupID() As Long
                                 
On Error GoTo ErrorTrap
Dim psTheString As String, rstRecSet As Recordset
'
'This function just finds next sequence number without checking whether person
' already has a missing report on the table.
'
    psTheString = "SELECT MAX(MissingReportGroupID) AS MaxGrpID " & _
                  "FROM tblMissingReports "

    Set rstRecSet = CMSDB.OpenRecordset(psTheString, dbOpenSnapshot)
      
    With rstRecSet
    If Not IsNull(!MaxGrpID) Then
        GetNextMissingReportGroupID = !MaxGrpID + 1
    Else
        GetNextMissingReportGroupID = 0
    End If
    End With
    
    rstRecSet.Close
 
    Exit Function
ErrorTrap:
    EndProgram

    
End Function
Public Function IsPersonIrregularThisMonth(PersonID As Long, MinDate_UK As Date) As Boolean
                                 
On Error GoTo ErrorTrap
Dim psTheString As String, rstRecSet As Recordset
        
    psTheString = "SELECT * " & _
                  "FROM tblIrregularPubs " & _
                  "WHERE PersonID = " & PersonID & _
                 " AND MinistryDate = #" & _
                   Format$(MinDate_UK, "mm/dd/yyyy") & "#"
                  
    Set rstRecSet = CMSDB.OpenRecordset(psTheString, dbOpenDynaset)
    
    If rstRecSet.BOF Then 'no irregular pub entry
        IsPersonIrregularThisMonth = False
    Else
        IsPersonIrregularThisMonth = True
    End If
    
    rstRecSet.Close
 
    Exit Function
ErrorTrap:
    EndProgram

    
End Function

Public Sub DeleteMissingReportRec(ThePerson As Long, _
                                  ActualMinDate_UK As String, _
                                  ZeroReport As Boolean)
                                 
On Error GoTo ErrorTrap
Dim psTheString As String, rstRecSet As Recordset, SaveMissingRptID As Long
Dim NextMissingRptID As Long

    '
    'Find the GrpID of missing-report being deleted
    '
    psTheString = "SELECT MissingReportGroupID " & _
                  "FROM tblMissingReports " & _
                  "WHERE PersonID = " & ThePerson & _
                  " AND ActualMinDate = #" & _
                    Format$(ActualMinDate_UK, "mm/dd/yyyy") & "#"

    Set rstRecSet = CMSDB.OpenRecordset(psTheString, dbOpenDynaset)
      
    With rstRecSet
    If Not .BOF Then
        SaveMissingRptID = !MissingReportGroupID
    Else
        Exit Sub 'no missing report for this month
    End If
    End With
    
    '
    'Now delete missing report if it's not a zero report. Otherwise just
    ' set the zero report flag
    '
    If Not ZeroReport Then
        CMSDB.Execute "DELETE FROM tblMissingReports " & _
                      "WHERE PersonID = " & ThePerson & _
                      " AND ActualMinDate = #" & _
                      Format$(ActualMinDate_UK, "mm/dd/yyyy") & "#"
    Else
        CMSDB.Execute "UPDATE tblMissingReports " & _
                      "SET ZeroReport = TRUE" & _
                      " WHERE PersonID = " & ThePerson & _
                      " AND ActualMinDate = #" & _
                      Format$(ActualMinDate_UK, "mm/dd/yyyy") & "#"
    End If
                    
    
    If Not ZeroReport Then
        '
        'Now change Grp ID of new group of missing reports created by deletion
        ' (If any)
        '
        
        NextMissingRptID = GetNextMissingReportGroupID
        
        CMSDB.Execute "UPDATE tblMissingReports " & _
                      "SET MissingReportGroupID = " & NextMissingRptID & _
                      " WHERE PersonID = " & ThePerson & _
                      " AND ActualMinDate > #" & _
                        Format$(ActualMinDate_UK, "mm/dd/yyyy") & "#" & _
                       " AND MissingReportGroupID = " & SaveMissingRptID
        
        'Now determine whether pub is inactive for existing and newly created groups
        DetermineInactivePublishers SaveMissingRptID
        DetermineInactivePublishers NextMissingRptID
        
        DetermineIrregularPublishers ThePerson, _
                                    CDate(ActualMinDate_UK), _
                                     IIf(ZeroReport, True, False)
   End If
   
    rstRecSet.Close
    Set rstRecSet = Nothing
    
    Exit Sub
ErrorTrap:
    EndProgram

    
End Sub
Public Function DeleteMissingReportsForPerson(ByVal ThePerson As Long) As Boolean
                                 
On Error GoTo ErrorTrap
Dim psTheString As String, rstRecSet As Recordset

    '
    'Find the GrpIDs of missing-reports being deleted
    '
    psTheString = "SELECT MissingReportGroupID " & _
                  "FROM tblMissingReports " & _
                  "WHERE PersonID = " & ThePerson

    Set rstRecSet = CMSDB.OpenRecordset(psTheString, dbOpenDynaset)
      
    WriteToLogFile "DeleteMissingReportsForPerson ------------- 0011"
      
    '
    'For each missing report group, delete recs from tblInactivePubs
    '
    With rstRecSet
    If Not .BOF Then
        Do Until .EOF
            CMSDB.Execute "DELETE FROM tblInactivePubs " & _
                          "WHERE MissingReportGroupID = " & !MissingReportGroupID
            
            .MoveNext
        Loop
    Else
        DeleteMissingReportsForPerson = True
        Exit Function 'no missing reports for this person!
    End If
    End With
    
    WriteToLogFile "DeleteMissingReportsForPerson ------------- 0012"
    
    '
    'Now delete from tblMissingReports
    '
    CMSDB.Execute "DELETE FROM tblMissingReports " & _
                  "WHERE PersonID = " & ThePerson
   
    rstRecSet.Close
    
    WriteToLogFile "DeleteMissingReportsForPerson ------------- 0013"
    
    DeleteMissingReportsForPerson = True
    
    Exit Function
ErrorTrap:
    EndProgram

    
End Function


Public Sub DetermineInactivePublishers(ByRef MissingReportGroupID As Long)
                                 
On Error GoTo ErrorTrap
Dim psTheString As String, rstRecSet As Recordset, SaveMissingRptID As Long
Dim psTheString2 As String, rstRecSet2 As Recordset
'Dim psTheString3 As String, rstRecSet3 As Recordset

    '
    'Count missing reports in the group
    '
    psTheString = "SELECT COUNT(*) AS MissingRptCount, " & _
                  "MAX(ActualMinDate) AS MaxDate, " & _
                  "MIN(ActualMinDate) AS MinDate " & _
                  "FROM tblMissingReports " & _
                  "WHERE MissingReportGroupID = " & MissingReportGroupID

    Set rstRecSet = CMSDB.OpenRecordset(psTheString, dbOpenDynaset)
      
    With rstRecSet
    If !MissingRptCount >= 6 Then
        
'        'Find the Start Date of inactivity - the 6th full month of no report
        
        'Anything on tblInactivePubs?
        psTheString2 = "SELECT MissingReportGroupID " & _
                      "FROM tblInactivePubs " & _
                      "WHERE MissingReportGroupID = " & MissingReportGroupID
                      
        Set rstRecSet2 = CMSDB.OpenRecordset(psTheString2, dbOpenDynaset)
        
        If rstRecSet2.BOF Then 'no inactive pub entry
            CMSDB.Execute "INSERT INTO tblInactivePubs " & _
                              "(MissingReportGroupID, StartDate, EndDate) " & _
                              "VALUES (" & MissingReportGroupID & ", #" & _
                                     Format$(DateAdd("m", 5, !MinDate), "mm/dd/yyyy") & "#, #" & _
                                     Format$(!MaxDate, "mm/dd/yyyy") & "#) "
            
        Else
            CMSDB.Execute "UPDATE tblInactivePubs " & _
                              "SET StartDate = #" & Format$(DateAdd("m", 5, !MinDate), "mm/dd/yyyy") & "#, " & _
                              "  EndDate = #" & Format$(!MaxDate, "mm/dd/yyyy") & "# " & _
                              "WHERE MissingReportGroupID = " & MissingReportGroupID
        End If
        rstRecSet2.Close
'        rstRecSet3.Close
    Else
        CMSDB.Execute "DELETE FROM tblInactivePubs " & _
                      "WHERE MissingReportGroupID = " & MissingReportGroupID
    End If
    End With
        
    rstRecSet.Close
    
    
    Exit Sub
ErrorTrap:
    EndProgram

    
End Sub

Public Sub DetermineIrregularPublishers(ByRef PersonID As Long, _
                                        ByRef MinDate_UK As Date, _
                                        ByRef NoReport As Boolean)

On Error GoTo ErrorTrap
Dim psTheString As String, rstRecSet As Recordset, SaveMissingRptID As Long
Dim psTheString2 As String, rstRecSet2 As Recordset, NoOfMonths As Long
Dim MinDateMod_UK As Date, ReportingDay As Long

'NB: This will INCLUDE inactive pubs. frmCongStats filters out from the
'    irregular pubs any that are inactive for the whole of a specified
'    period - such pubs are no long irregular, but rather inactive.

    MinDateMod_UK = DateAdd("m", 5, MinDate_UK)

    Select Case NoReport
    Case True 'There is no report for the supplied month
        'insert Irregular Pub entry on tblIrregularPubs for this and next 5 months
        On Error Resume Next 'in case of duplicate insert
        CMSDB.Execute "INSERT INTO tblIrregularPubs " & _
                          "(PersonID, IrregStartDate, IrregEndDate) " & _
                          "VALUES (" & PersonID & ", #" & _
                                 Format$(MinDate_UK, "mm/dd/yyyy") & "#, #" & _
                                 Format$(MinDateMod_UK, "mm/dd/yyyy") & "#) "
        On Error GoTo ErrorTrap
        
    Case False
        CMSDB.Execute "DELETE FROM tblIrregularPubs " & _
                      "WHERE PersonID = " & PersonID & _
                     " AND IrregStartDate = #" & _
                       Format$(MinDate_UK, "mm/dd/yyyy") & "#"
    
    End Select
    
    Exit Sub
ErrorTrap:
    EndProgram

    
End Sub

Public Sub DeleteRangeOfMissingReportRecs(ByVal PersonID As Long, _
                                          ByVal StartDate_UK As Date, _
                                          ByVal EndDate_UK As Date)
                                 
Dim TempDate As Date, NoMonths As Long, i As Long
On Error GoTo ErrorTrap

    NoMonths = DateDiff("m", StartDate_UK, EndDate_UK)
    
    TempDate = StartDate_UK
    
    For i = 1 To NoMonths
        DeleteMissingReportRec PersonID, CStr(TempDate), False
        TempDate = DateAdd("m", 1, TempDate)
    Next
        
    
    Exit Sub
ErrorTrap:
    EndProgram

    
End Sub

Public Sub DeleteMinReport(ActualReportMonth As Long, _
                           ActualReportYear As Long, _
                           ThePerson As Long)
On Error GoTo ErrorTrap

Dim TheString As String

    TheString = Format("01/" & CStr(ActualReportMonth) & "/" & _
                ConvertServiceYearToNormalYear(CDate("01/" & _
                           CStr(ActualReportMonth) & _
                           "/" & CStr(ActualReportYear))), "dd/mm/yyyy")
              
    CMSDB.Execute "DELETE FROM tblMinReports " & _
                  "WHERE PersonID = " & ThePerson & _
                  " AND MinistryDoneInMonth = " & ActualReportMonth & _
                  " AND MinistryDoneInYear = " & ActualReportYear
          
    CMSDB.Execute "DELETE FROM tblPubRecCardRowPrinted " & _
                  "WHERE PersonID = " & ThePerson & _
                  " AND ActualMinPeriod = #" & Format$(TheString, "mm/dd/yyyy") & "#"
          
    '
    'If report is deleted, then it's now missing!
    '
    RefreshMinistryStatusForPerson ThePerson
'    PutAllMissingReportsIntoTable TheString, TheString, ThePerson
                        

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Public Sub SetPersonInactive(ThePerson As Long)
On Error GoTo ErrorTrap

Dim TheString As String
Dim rs As Recordset

    '
    'set the person's publisher/pioneer end dates
    '
    MsgBox "Now update the publisher's status...", vbOKOnly + vbInformation, AppName
    
    frmFieldServiceRoles.FormPersonID = ThePerson
    frmFieldServiceRoles.PersonSetInactive = True
    frmFieldServiceRoles.Show vbModal
          
    '
    'Remove person from all cong roles (except baptism/associated)
    '
    DeleteSomeRows "tblEldersAndServants", "PersonID = ", ThePerson
'    DeleteSomeRows "tblTaskPersonSuspendDates", "Person = ", ThePerson
'    DeleteSomeRows "tblTaskAndPerson", "Person = ", ThePerson
    CMSDB.Execute ("DELETE FROM tblTaskAndPerson " & _
                        "WHERE Task NOT IN (55, 56)" & _
                        " AND Person = " & ThePerson)
    CMSDB.Execute ("DELETE FROM tblTaskPersonSuspendDates " & _
                        "WHERE Task NOT IN (55, 56)" & _
                        " AND Person = " & ThePerson)
                        
    
    'now check if the person has any linked addresses anywhere...
    Set rs = CMSDB.OpenRecordset("SELECT ID FROM tblNameAddress " & _
                                 "WHERE LinkedAddressPerson = " & ThePerson & _
                                 " AND Active = TRUE", dbOpenDynaset)
    
    If Not rs.BOF Then
        MsgBox "At least one other congregation member's address is linked to that of " & _
                    CongregationMember.NameWithMiddleInitial(ThePerson) & _
                ". These may need to be updated.", vbOKOnly + vbInformation, AppName
                    
        CMSDB.Execute "UPDATE tblNameAddress SET LinkedAddressPerson = 0 WHERE LinkedAddressPerson = " & ThePerson
    End If
       
    On Error Resume Next
    rs.Close
    Set rs = Nothing

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Public Function LastReportMonth(Optional ForCongStats As Boolean = True) As Date
On Error GoTo ErrorTrap

Dim TheString As String
                        
    If Day(Now) >= GlobalParms.GetValue(IIf(ForCongStats, "DayOfMonthForCongStats", "DayOfMonthForReportToSociety"), "NumVal") Then
        TheString = Format$((DateAdd("m", -1, Now)), "dd/mm/yyyy")
    Else
        TheString = Format$((DateAdd("m", -2, Now)), "dd/mm/yyyy")
    End If
        
    LastReportMonth = CDate("01/" & _
                     Month(TheString) & "/" & _
                     year(TheString))
                        

    Exit Function
ErrorTrap:
    EndProgram
    

End Function
Public Function CurrentReportMonth() As Date
On Error GoTo ErrorTrap

Dim TheString As String
                        
    If Day(Now) > GlobalParms.GetValue("DayOfMonthForReportToSociety", "NumVal") Then
        'past 6th of month
        TheString = Format$(Now, "dd/mm/yyyy")
    Else
        'not up to 6th of month, so check up to current month -1
        TheString = Format$((DateAdd("m", -1, Now)), "dd/mm/yyyy")
    End If
        
    CurrentReportMonth = CDate("01/" & _
                     Month(TheString) & "/" & _
                     year(TheString))
                        

    Exit Function
ErrorTrap:
    EndProgram
    

End Function
Public Function TruncateTextToFit(TheString As String, _
                                  MaxLengthCM As Single, _
                                  FontName As String, _
                                  FontSize As Single) As String
On Error GoTo ErrorTrap

Dim TempString As String
                        
    Printer.ScaleMode = vbCentimeters
    Printer.Font.Name = FontName
    Printer.Font.Size = FontSize
    
    TempString = TheString
    
    Do Until Printer.TextWidth(TempString) <= MaxLengthCM
        TempString = Left(TempString, Len(TempString) - 1)
        If Len(TempString) = 0 Then
            Exit Do
        End If
    Loop
    
    TruncateTextToFit = TempString

    Exit Function
ErrorTrap:
    EndProgram
    

End Function
Public Sub SyncroniseOutlookContact(PersonID As Long, Optional NewEntryID As String = "")
On Error GoTo ErrorTrap
Dim bOutLookWasOpen As Boolean
'Dim oOutlook As Outlook.Application
'Dim oContactItem As Outlook.ContactItem
'Dim oItem As Outlook.ContactItem
'Dim oNameSpace As Outlook.NameSpace
'Dim oFldrs As Outlook.Folders
'Dim oFldr As Outlook.MAPIFolder
'Dim oFind As Outlook.ContactItem
'Dim oItems As Outlook.Items
Dim oOutlook As Object
Dim oContactItem As Object
Dim oItem As Object
Dim oNameSpace As Object
Dim oFldrs As Object
Dim oFldr As Object
Dim oFind As Object
Dim oItems As Object
Dim bOutlookOpen As Boolean
Dim lPersonID As Long
Dim sEntryID As String
Dim bAllowEmailAccess As Boolean
Dim lPerson As Long
Dim rsCMS As Recordset
Dim bSetFldsBlank As Boolean

Dim rs As Recordset, arr() As String
    
    bSetFldsBlank = GlobalParms.GetValue("OutlookSynch_RemoveCMSFldWhenBlankOutlookFld", "TrueFalse", True)
    
    If PersonID > 0 Then
    
        If MsgBox("Synchronise " & CongregationMember.NameWithMiddleInitial(PersonID) & _
                  " to selected Outlook entry? ", vbYesNo + vbQuestion + vbDefaultButton2, AppName) = vbNo Then
                
            ShowMessage "Operation cancelled", 1250, IIf(FormIsOpen("frmPersonalDetails"), frmPersonalDetails, frmMainMenu)
            Exit Sub
        End If
        Set rsCMS = GetGeneralRecordset("SELECT ID, OfficialFirstName, FirstName, MiddleName, LastName " & _
                                            "FROM tblNameAddress " & _
                                        "WHERE ID = " & PersonID)
    Else
        If MsgBox("Re-synchronise to Outlook contacts? Names in CMS will not be updated.", vbYesNo + vbQuestion + vbDefaultButton2, AppName) = vbNo Then
                
            ShowMessage "Operation cancelled", 1250, IIf(FormIsOpen("frmPersonalDetails"), frmPersonalDetails, frmMainMenu)
            Exit Sub
        End If
        
        Set rsCMS = GetGeneralRecordset("SELECT ID, OutlookEntryID " & _
                                            "FROM (tblNameAddress a LEFT JOIN tblVisitingSpeakers b " & _
                                        "     ON a.ID = b.PersonID) " & _
                                        "  INNER JOIN tblOutlookSync c ON c.PersonID = a.ID " & _
                                        "WHERE Active = TRUE OR b.PersonID IS NOT NULL ")
    End If
        
    Screen.MousePointer = vbHourglass
    
    On Error Resume Next
    Set oOutlook = GetObject(, "Outlook.Application")
    bOutLookWasOpen = True
    If Err.number <> 0 Then
      bOutLookWasOpen = False
      Err.Clear
      Set oOutlook = CreateObject("Outlook.Application")
    End If
    
    If Err.number <> 0 Then
        ShowMessage "Could not open Outlook", 1500, IIf(FormIsOpen("frmPersonalDetails"), frmPersonalDetails, frmMainMenu)
        Set oOutlook = Nothing
        bOutlookOpen = False
        GoTo GetOut
    Else
        Set oNameSpace = oOutlook.GetNamespace("MAPI")
        Set oFldrs = oNameSpace.Folders
        bOutlookOpen = True
    End If
    On Error GoTo ErrorTrap
    
    Do While Not rsCMS.BOF And Not rsCMS.EOF
    
        lPerson = rsCMS!ID
        
        If NewEntryID = "" Then
            sEntryID = rsCMS!OutlookEntryID
        Else
            sEntryID = NewEntryID
        End If
        
        On Error Resume Next
        Set oContactItem = oNameSpace.GetItemFromID(sEntryID)
        If Err.number <> 0 Then ' something's wrong with this EntryID
            CMSDB.Execute "DELETE FROM tblOutlookSync WHERE OutlookEntryID = '" & sEntryID & "'"
            GoTo LoopNext
        End If
        
        Set rs = GetGeneralRecordset("SELECT Address1, Address2, Address3, Address4, PostCode, HomePhone, MobilePhone, Email, " & _
                                    " OfficialFirstName, FirstName, MiddleName, LastName " & _
                                    "FROM tblNameAddress " & _
                                    "WHERE ID = " & lPerson)
        
        With rs
        
        If .BOF Then
            MsgBox "CMS entry not located - logic error", vbOKOnly + vbExclamation, AppName
            GoTo GetOut
        End If
        
        .Edit
        
        'only synch address if not linked to someone else's
        If CongregationMember.LinkedAddressPerson(lPerson) = 0 Then
            If oContactItem.HomeAddressStreet <> "" Then
                'try to split Outlook's HomeAddressStreet into multiple lines, delimited by CR or ','
                arr() = Split(oContactItem.HomeAddressStreet, vbCrLf)
                If UBound(arr) = 0 Then
                    arr() = Split(oContactItem.HomeAddressStreet, ",")
                End If
                If UBound(arr) >= 0 Then
                    !Address1 = arr(0)
                Else
                    If bSetFldsBlank Then !Address1 = ""
                End If
                If UBound(arr) >= 1 Then
                    !Address2 = arr(1)
                Else
                    If bSetFldsBlank Then !Address2 = ""
                End If
                If UBound(arr) >= 2 Then
                    !Address3 = arr(2)
                Else
                    If bSetFldsBlank Then !Address3 = ""
                End If
            End If
            If oContactItem.HomeAddressCity <> "" Then
                !Address4 = oContactItem.HomeAddressCity
            Else
                If bSetFldsBlank Then !Address4 = ""
            End If
            If oContactItem.HomeAddressPostalCode <> "" Then
                !PostCode = oContactItem.HomeAddressPostalCode
            Else
                If bSetFldsBlank Then !PostCode = ""
            End If
            If oContactItem.HomeTelephoneNumber <> "" Then
                !HomePhone = oContactItem.HomeTelephoneNumber
            Else
                If bSetFldsBlank Then !HomePhone = ""
            End If
        End If
        
        If GlobalParms.GetValue("OutlookSynch_AllowEmailAddrAccess", "TrueFalse") Then
            If oContactItem.Email1Address <> "" Then
                !Email = oContactItem.Email1Address
            Else
                If bSetFldsBlank Then !Email = ""
            End If
        End If
        
        If oContactItem.MobileTelephoneNumber <> "" Then
            !MobilePhone = oContactItem.MobileTelephoneNumber
        Else
            If bSetFldsBlank Then !MobilePhone = ""
        End If
        
        If PersonID > 0 Then
            If oContactItem.FirstName <> "" And oContactItem.LastName <> "" Then
                If !FirstName <> oContactItem.FirstName Or _
                    !OfficialFirstName <> oContactItem.FirstName Or _
                    !MiddleName <> oContactItem.MiddleName Or _
                    !LastName <> oContactItem.LastName Then
                    If MsgBox("Do you want the name to be updated? The First Name and Offical First Name " & _
                                "will be updated.", vbYesNo + vbQuestion, AppName) = vbYes Then
                        !FirstName = oContactItem.FirstName
                        !OfficialFirstName = oContactItem.FirstName
                        !LastName = oContactItem.LastName
                    End If
                End If
            End If
        End If

        .Update
    
        End With
        
        'now update tblOutlookSync
        If NewEntryID <> "" And PersonID > 0 Then
            Set rs = GetGeneralRecordset("SELECT PersonID, OutlookEntryID " & _
                                         "FROM tblOutlookSync " & _
                                         "WHERE PersonID = " & PersonID)
            
            With rs
            
            
            If .BOF Then
                .AddNew
                !PersonID = PersonID
            Else
                .Edit
            End If
            
            !OutlookEntryID = NewEntryID
            
            .Update
            
            End With
        End If
        
LoopNext:
        
        rsCMS.MoveNext
    
    Loop
    
    If FormIsOpen("frmPersonalDetails") Then
        frmPersonalDetails.RefreshNamesList
        frmPersonalDetails.GoToPerson PersonID
    End If
    
    If FormIsOpen("frmSyncToOutlook") Then
        frmSyncToOutlook.ShowCurrentMatches
    End If
    
    Screen.MousePointer = vbNormal
    
    If PersonID > 0 Then
        MsgBox CongregationMember.NameWithMiddleInitial(PersonID) & _
                " synchronised to Outlook entry", vbOKOnly + vbInformation, AppName
    Else
        MsgBox "Outlook synchronisation complete", vbOKOnly + vbInformation, AppName
    End If

    
GetOut:

    On Error Resume Next

    rs.Close
    Set rs = Nothing
    rsCMS.Close
    Set rsCMS = Nothing
    
    Screen.MousePointer = vbNormal

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Public Function UnzipFile(sPathToZip As String, _
                          sPathToExtractedFile As String, _
                          Optional bOverwriteExisting As Boolean = True) As Boolean
                         
'
'Must add reference to Microsoft Scripting Runtime to use the File System Object
'
Dim FileSysObj As New Scripting.FileSystemObject, TheDrive As Scripting.Drive
Dim sExtractedFile As String, sDirOfExtractedFile As String, InitialDirectory As String
Dim i As Long, bOK As Boolean
Dim m_cUnzip As cUnzip

On Error Resume Next

    UnzipFile = False
    
    Set m_cUnzip = New cUnzip
            
    With FileSysObj
    
    If .FileExists(sPathToZip) Then
    
        sExtractedFile = .GetFileName(sPathToExtractedFile)
        sDirOfExtractedFile = .GetParentFolderName(sPathToExtractedFile)
        
        '
        'Prepare for Unzip...
        '
        With m_cUnzip
        .OverwriteExisting = bOverwriteExisting
        .ZipFile = sPathToZip
        .Directory 'This adds the zip's contents to the filecount
        
        'select file in zip
        bOK = False
        For i = 1 To .FileCount
            If .Filename(i) = sExtractedFile Then
                .FileSelected(i) = True
                bOK = True
                Exit For
            End If
        Next i
        
        If Not bOK Then
            UnzipFile = False
            Exit Function
        End If
        
        .UnzipFolder = sDirOfExtractedFile
        .Unzip
        If Err.number > 0 Then
            UnzipFile = False
            Exit Function
        End If
        
        End With
    Else
        UnzipFile = False
        Exit Function
    End If
    
    End With
    
    UnzipFile = True

End Function


