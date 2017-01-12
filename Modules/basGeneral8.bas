Attribute VB_Name = "basGeneral8"
Option Explicit

Public Function GetInfirmityLevelDescription(lLevel As Long) As String

On Error GoTo ErrorTrap

    Dim s As String
    
    'changes here shuld be reflected in FillInfirmityCombo()
    
    Select Case lLevel
    Case 0
        s = "No health problems"
    Case 1
        s = "Minor health problems"
    Case 2
        s = "Major health problems"
    Case 3
        s = "Severe disability"
    End Select

    GetInfirmityLevelDescription = s
    

    Exit Function
ErrorTrap:
    EndProgram

End Function


Public Sub GetTableUpdateDates()

On Error GoTo ErrorTrap
    
Dim tdf As TableDef

    DelAllRows "tblTableUpdateDateTimes"
    
    For Each tdf In CMSDB.TableDefs
        If Left(tdf.Name, 3) = "tbl" And tdf.Name <> "tblTableUpdateDateTimes" Then
            CMSDB.Execute "INSERT INTO tblTableUpdateDateTimes (TableName, LastUpdateDate) " & _
                          "VALUES ('" & tdf.Name & "', " & GetDateTimeStringForSQLWhere(tdf.LastUpdated) & ")"
        
        End If
    Next

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Function HaveAnyTablesChanged() As Boolean

On Error GoTo ErrorTrap
    
Dim tdf As TableDef, rs As Recordset

    Set rs = CMSDB.OpenRecordset("tblTableUpdateDateTimes", dbOpenDynaset)

    With rs
    For Each tdf In CMSDB.TableDefs
        If Left(tdf.Name, 3) = "tbl" And tdf.Name <> "tblTableUpdateDateTimes" Then
            .FindFirst "TableName = '" & tdf.Name & "'"
            If .NoMatch Then 'table has since been deleted
                HaveAnyTablesChanged = True
                GoTo GetOut
            End If
            
            If CStr(tdf.LastUpdated) <> CStr(!LastUpdateDate) Then
                HaveAnyTablesChanged = True
                GoTo GetOut
            End If
                
        End If
    Next
    End With
    
    HaveAnyTablesChanged = False
    
GetOut:
    On Error Resume Next
    rs.Close
    Set rs = Nothing

    Exit Function
ErrorTrap:
    EndProgram

End Function
Public Sub ShowContactDetails(PersonID As Long, _
                              Optional MessageBody As String = "", _
                              Optional MessageTitle As String = "", _
                              Optional bModal As Boolean = False, _
                              Optional ByVal ParentForm As Form, _
                              Optional EmailType As cmsEmailTypes = cmsGeneral)
                              
On Error Resume Next

Dim TheForm As Form

    If ParentForm Is Nothing Then
        Set TheForm = frmMainMenu
    End If

    With frmContactDetails
    .PersonID = PersonID
    .MessageBody = MessageBody
    .MessageTitle = MessageTitle
    .EmailType = EmailType
    Set .ParentForm = ParentForm
    
    If bModal Then
        .Show vbModal, ParentForm
    Else
        .Show vbModeless, ParentForm
        If Err.number <> 0 Then
            .Show vbModal, ParentForm
        End If
    End If
            
    End With

End Sub

Public Function GetFlexGridColFromXPos(TheGrid As MSFlexGrid, XPos As Single) As Long

On Error GoTo ErrorTrap
Dim i As Long, lAccWidth As Long
                
    With TheGrid
            
    For i = 0 To .Cols - 1
        
        lAccWidth = lAccWidth + .ColWidth(i)
        
        If XPos <= lAccWidth Then
            GetFlexGridColFromXPos = i
            Exit Function
        End If
    
    Next i
    
    End With
                
    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Function GetPersonGroupingName(PersonGroupingID) As String

Dim rs As Recordset
On Error GoTo ErrorTrap

    Set rs = CMSDB.OpenRecordset("tblPersonGroupingNames", dbOpenSnapshot)
    With rs
    .FindFirst "PersonGroupingID = " & PersonGroupingID
    If Not .NoMatch Then
        GetPersonGroupingName = !PersonGroupingName
    Else
        GetPersonGroupingName = ""
    End If
    End With
    rs.Close
    Set rs = Nothing

    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Function NewMtgArrangementStarted(WeekBeginning As String) As MidweekMtgVersion

On Error GoTo ErrorTrap

    If Not ValidDate(WeekBeginning) Then
        Err.Raise vbObjectError + 1, "NewMtgArrangementStarted", _
                "Invalid date supplied: " & WeekBeginning
    End If

    Select Case True
    Case CDate(WeekBeginning) >= CDate(WEEK_OF_NEW_CLAM_MTG_ARRANGEMENT)
        NewMtgArrangementStarted = CLM2016
    Case CDate(WeekBeginning) >= CDate(WEEK_OF_NEW_MTG_ARRANGEMENT)
        NewMtgArrangementStarted = TMS2009
    Case Else
        NewMtgArrangementStarted = Pre2009
    End Select


    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Function ServiceMtgItemsBetweenDates(TheStartDate As Date, TheEndDate As Date) As Boolean
On Error GoTo ErrorTrap
Dim rsttmsquery As Recordset
'
    Set rsttmsquery = CMSDB.OpenRecordset("SELECT DISTINCT 1 " & _
                                       "FROM tblServiceMtgs " & _
                                       "WHERE MeetingDate BETWEEN #" & Format(TheStartDate, "mm/dd/yyyy") & "#" & _
                                       " AND #" & Format(TheEndDate, "mm/dd/yyyy") & "#" & _
                                       " AND PersonID > 0 ", dbOpenDynaset)
    
    
    
    ServiceMtgItemsBetweenDates = Not rsttmsquery.BOF
    
    rsttmsquery.Close
    Set rsttmsquery = Nothing
    
    Exit Function
ErrorTrap:
    EndProgram

End Function
Public Function CongBibleStudyBetweenDates(TheStartDate As Date, TheEndDate As Date) As Boolean
On Error GoTo ErrorTrap
Dim rsttmsquery As Recordset
'
    Set rsttmsquery = CMSDB.OpenRecordset("SELECT DISTINCT 1 " & _
                                       "FROM tblCongBibleStudyRota " & _
                                       "WHERE MeetingDate BETWEEN #" & Format(TheStartDate, "mm/dd/yyyy") & "#" & _
                                       " AND #" & Format(TheEndDate, "mm/dd/yyyy") & "#" & _
                                       " AND (ConductorID > 0 or ReaderID > 0) ", dbOpenDynaset)
    
    
    
    CongBibleStudyBetweenDates = Not rsttmsquery.BOF
    
    rsttmsquery.Close
    Set rsttmsquery = Nothing
    
    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Function CongBibleStudySongForWeek(ByVal MondayDate_UK As Date) As Long
On Error GoTo ErrorTrap
Dim rst As Recordset, str As String
Dim var As Variant

    str = "SELECT OpeningSong " & _
          "FROM tblCongBibleStudyRota " & _
          "WHERE MeetingDate = #" & Format(MondayDate_UK, "mm/dd/yyyy") & "# "
          
    Set rst = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    If Not rst.BOF Then
        CongBibleStudySongForWeek = HandleNull(rst!OpeningSong)
    Else
        CongBibleStudySongForWeek = 0
    End If
    
    rst.Close
    Set rst = Nothing
        
    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Sub DeleteServiceMtg(StartDate As Date, EndDate As Date)
On Error GoTo ErrorTrap
Dim rsttmsquery As Recordset
        
    CMSDB.Execute "DELETE FROM tblServiceMtgs WHERE MeetingDate BETWEEN " & _
                    GetDateStringForSQLWhere(CStr(StartDate)) & " AND " & GetDateStringForSQLWhere(CStr(EndDate))
    
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Public Function GetFirstServMtgBro(ByVal MeetingDate As Date) As Long
On Error GoTo ErrorTrap
Dim rst As Recordset, str As String
Dim var As Variant

    str = "SELECT TOP 1 PersonID " & _
          "FROM tblServiceMtgs " & _
          "WHERE MeetingDate = #" & Format(MeetingDate, "mm/dd/yyyy") & "# " & _
          "AND PersonID <> 0 " & _
          "ORDER BY SeqNum "
          
    Set rst = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    If Not rst.BOF Then
        GetFirstServMtgBro = rst!PersonID
    Else
        GetFirstServMtgBro = 0
    End If
    
    rst.Close
    Set rst = Nothing
        
    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Function GetFirstServMtgItemName(ByVal MeetingDate As Date) As String
On Error GoTo ErrorTrap
Dim rst As Recordset, str As String
Dim var As Variant

    str = "SELECT TOP 1 ItemName " & _
          "FROM tblServiceMtgs " & _
          "WHERE MeetingDate = #" & Format(MeetingDate, "mm/dd/yyyy") & "# " & _
          "AND PersonID <> 0 " & _
          "ORDER BY SeqNum "
          
    Set rst = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    If Not rst.BOF Then
        GetFirstServMtgItemName = rst!ItemName
    Else
        GetFirstServMtgItemName = ""
    End If
    
    rst.Close
    Set rst = Nothing
        
    Exit Function
ErrorTrap:
    EndProgram

End Function


Public Function CongBibleStudyConductor(ByVal MondayDate_UK As Date) As Long
On Error GoTo ErrorTrap
Dim rst As Recordset, str As String
Dim var As Variant

    str = "SELECT MeetingDate, " & _
          "       ConductorID, " & _
          "       ReaderID " & _
          "FROM tblCongBibleStudyRota " & _
          "WHERE MeetingDate = #" & Format(MondayDate_UK, "mm/dd/yyyy") & "# "
          
    Set rst = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    If Not rst.BOF Then
        CongBibleStudyConductor = rst!ConductorID
    Else
         CongBibleStudyConductor = 0
    End If
    
    rst.Close
    Set rst = Nothing
        
    Exit Function
ErrorTrap:
    EndProgram

End Function
Public Function CongBibleStudyReader(ByVal MondayDate_UK As Date) As Long
On Error GoTo ErrorTrap
Dim rst As Recordset, str As String
Dim var As Variant

    str = "SELECT MeetingDate, " & _
          "       ConductorID, " & _
          "       ReaderID " & _
          "FROM tblCongBibleStudyRota " & _
          "WHERE MeetingDate = #" & Format(MondayDate_UK, "mm/dd/yyyy") & "# "
          
    Set rst = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    If Not rst.BOF Then
        CongBibleStudyReader = rst!ReaderID
    Else
         CongBibleStudyReader = 0
    End If
    
    rst.Close
    Set rst = Nothing
        
    Exit Function
ErrorTrap:
    EndProgram

End Function
Public Function CongBibleStudyPrayer(ByVal MondayDate_UK As Date) As Long
On Error GoTo ErrorTrap
Dim rst As Recordset, str As String
Dim var As Variant

    str = "SELECT MeetingDate, " & _
          "       PrayerID " & _
          "FROM tblCongBibleStudyRota " & _
          "WHERE MeetingDate = #" & Format(MondayDate_UK, "mm/dd/yyyy") & "# "
          
    Set rst = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    If Not rst.BOF Then
        CongBibleStudyPrayer = rst!PrayerID
    Else
         CongBibleStudyPrayer = 0
    End If
    
    rst.Close
    Set rst = Nothing
        
    Exit Function
ErrorTrap:
    EndProgram

End Function


Public Function SetPrinter(ByVal sPrinterName As String) As Boolean
On Error GoTo ErrorTrap
Dim bOK As Boolean
            
    Dim p As VB.Printer
    
    For Each p In VB.Printers
        If p.DeviceName = sPrinterName Then
            Set Printer = p
            bOK = True
            Exit For
        End If
    Next
    
    SetPrinter = bOK
    
    Exit Function
ErrorTrap:
    EndProgram

End Function
