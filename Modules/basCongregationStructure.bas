Attribute VB_Name = "basCongregationStructure"
Option Explicit



Public Function GetGroupCleaning(MeetingDate_wcMonday As Date) As Long
On Error GoTo ErrorTrap
Dim rstCleaning As Recordset, str As String
   
    If TableExists("tblPrevCleaningRota") Then
        str = "SELECT * " & _
                "FROM tblPrevCleaningRota " & _
                "UNION ALL " & _
                "SELECT * " & _
                "FROM tblCleaningRota " & _
                "ORDER BY RotaDate "
    Else
        str = "SELECT * " & _
                "FROM tblCleaningRota " & _
                "ORDER BY RotaDate "
    End If
   
    Set rstCleaning = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    With rstCleaning
    .FindFirst "RotaDate = #" & Format$(MeetingDate_wcMonday, "mm/dd/yyyy") & "#"
    If Not .NoMatch Then
        GetGroupCleaning = !GroupNo
    Else
        GetGroupCleaning = 0
    End If
    End With
    
    rstCleaning.Close
    Set rstCleaning = Nothing
            
    Exit Function
ErrorTrap:
    EndProgram

End Function
Public Function GetMidWkMtgStartTime(MeetingDate As Date, bReturnDefault As Boolean, Optional sDefaultVal As String = "") As String
On Error GoTo ErrorTrap
Dim rs As Recordset, str As String
   
    str = "SELECT * " & _
            "FROM tblMidWkMtgTempStartTime " & _
            "WHERE MeetingDate = " & GetDateStringForSQLWhere(CStr(MeetingDate))
   
    Set rs = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    With rs
    If .BOF Then
        If bReturnDefault Then
            str = GlobalParms.GetValue("MidWeekMeetingStartTime", "TimeVal")
        Else
            str = sDefaultVal
        End If
    Else
        str = CStr(!NewTime)
    End If
    End With
    
    GetMidWkMtgStartTime = Format(str, "HH:MM")
    
    rs.Close
    Set rs = Nothing
            
    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Sub SaveMidWkMtgTempStartTime(MeetingDate As String, NewTime As String)
On Error GoTo ErrorTrap
Dim rs As Recordset, str As String
   
    str = "SELECT * " & _
            "FROM tblMidWkMtgTempStartTime "
   
    Set rs = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    With rs
    .FindFirst "MeetingDate = " & GetDateStringForSQLWhere(MeetingDate)
    If .NoMatch Then
        .AddNew
        !MeetingDate = MeetingDate
    Else
        .Edit
    End If
    !NewTime = NewTime
    .Update
    End With
    
    rs.Close
    Set rs = Nothing
            
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Sub DeleteMidWkMtgTempStartTime(MeetingDate As String)
On Error GoTo ErrorTrap
Dim rs As Recordset, str As String
   
    CMSDB.Execute "DELETE " & _
                    "FROM tblMidWkMtgTempStartTime " & _
                    "WHERE MeetingDate = " & GetDateStringForSQLWhere(MeetingDate)
               
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Function GetPubCardVersionDesc(CardTypeID) As String
On Error GoTo ErrorTrap
Dim rst As Recordset, str As String
   
    str = "SELECT CardTypeDesc FROM tblPubCardTypes WHERE CardTypeID = " & CardTypeID
   
    Set rst = CMSDB.OpenRecordset(str, dbOpenForwardOnly)
    
    With rst
    If Not .BOF Then
        GetPubCardVersionDesc = !CardTypeDesc
    Else
        GetPubCardVersionDesc = ""
    End If
    End With
    
    rst.Close
    Set rst = Nothing
            
    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Function GetTravellingOverseers() As Recordset
On Error GoTo ErrorTrap
Dim sSQL As String
   
    sSQL = "SELECT ID, " & _
                  "LastName & ', ' & FirstName & ' ' & MiddleName AS TheName " & _
           "FROM tblNameAddress INNER JOIN tblVisitingSpeakers ON " & _
           "tblNameAddress.ID = tblVisitingSpeakers.PersonID " & _
           "WHERE CongNo = 32767 " & _
            " ORDER BY 2 "
    
    Set GetTravellingOverseers = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
                
    Exit Function
ErrorTrap:
    EndProgram

End Function
Public Function GetMapName(MapNo As Long) As String
On Error GoTo ErrorTrap
Dim sSQL As String, rs As Recordset
   
    sSQL = "SELECT MapName " & _
           "FROM tblTerritoryMaps " & _
           "WHERE MapNo = " & MapNo
    
    Set rs = CMSDB.OpenRecordset(sSQL, dbOpenForwardOnly)
    
    If rs.BOF Then
        GetMapName = ""
    Else
        GetMapName = HandleNull(rs!MapNAme, "")
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    EndProgram

End Function


Public Function GetCongregationName(CongNo As Long) As String
On Error GoTo ErrorTrap
Dim rs As Recordset
   
    Set rs = CMSDB.OpenRecordset("tblCong", dbOpenDynaset)
    
    With rs
    .FindFirst "CongNo = " & CongNo
    If Not .NoMatch Then
        GetCongregationName = !CongName
    Else
        GetCongregationName = ""
    End If
    End With
    
    Set rs = Nothing
            
    Exit Function
ErrorTrap:
    EndProgram

End Function
Public Function GetCongSundayMtgTime(CongNo As Long) As String
On Error GoTo ErrorTrap
Dim rs As Recordset
   
    Set rs = CMSDB.OpenRecordset("tblCong", dbOpenDynaset)
    
    With rs
    .FindFirst "CongNo = " & CongNo
    If Not .NoMatch Then
        GetCongSundayMtgTime = Format(CStr(!SundayMtgTime), "HH:MM")
    Else
        GetCongSundayMtgTime = ""
    End If
    End With
    
    Set rs = Nothing
            
    Exit Function
ErrorTrap:
    EndProgram

End Function
Public Function GetCongSundayMtgDay(CongNo As Long) As Long
On Error GoTo ErrorTrap
Dim rs As Recordset
   
    Set rs = CMSDB.OpenRecordset("tblCong", dbOpenDynaset)
    
    With rs
    .FindFirst "CongNo = " & CongNo
    If Not .NoMatch Then
        GetCongSundayMtgDay = !SundayMtgDay
    Else
        GetCongSundayMtgDay = 0
    End If
    End With
    
    Set rs = Nothing
            
    Exit Function
ErrorTrap:
    EndProgram

End Function


Public Function GetGroupName(GroupID As Long, Optional ZeroValue As String = "", Optional RemoveWordGroupFromEnd As Boolean = False) As String
On Error GoTo ErrorTrap
Dim rstGroups As Recordset, str As String
   
    If ZeroValue <> "" And GroupID = 0 Then
        GetGroupName = ZeroValue
        Exit Function
    End If
    
    Set rstGroups = CMSDB.OpenRecordset("tblBookGroups", dbOpenDynaset)
    With rstGroups
    .FindFirst "GroupNo = " & GroupID
    If Not .NoMatch Then
        str = !GroupName
    Else
        str = ""
    End If
    End With
    
    If LCase(Right(str, 6)) = " group" Then
        str = Left(str, Len(str) - 6)
    End If
    
    GetGroupName = str
    
    rstGroups.Close
    Set rstGroups = Nothing
            
    Exit Function
ErrorTrap:
    EndProgram

End Function
Public Function GetBookStudyOverseer(GroupID As Long) As Long
On Error GoTo ErrorTrap
Dim rs As Recordset, str As String

    str = "SELECT B.PersonID " & _
          "FROM tblTaskAndPerson A INNER JOIN tblBookGroupMembers B " & _
          "          ON A.Person = B.PersonID " & _
          "WHERE B.GroupNo = " & GroupID & _
          " AND A.Task = 97 "
   
    Set rs = CMSDB.OpenRecordset(str, dbOpenSnapshot)
    With rs
    If Not .BOF Then
        GetBookStudyOverseer = !PersonID
    Else
        GetBookStudyOverseer = 0
    End If
    End With
    
    rs.Close
    Set rs = Nothing
            
    Exit Function
ErrorTrap:
    EndProgram

End Function
Public Function GetBookStudyAssistant(GroupID As Long) As Long
On Error GoTo ErrorTrap
Dim rs As Recordset, str As String

    str = "SELECT B.PersonID " & _
          "FROM tblTaskAndPerson A INNER JOIN tblBookGroupMembers B " & _
          "          ON A.Person = B.PersonID " & _
          "WHERE B.GroupNo = " & GroupID & _
          " AND A.Task = 98 "
   
    Set rs = CMSDB.OpenRecordset(str, dbOpenSnapshot)
    With rs
    If Not .BOF Then
        GetBookStudyAssistant = !PersonID
    Else
        GetBookStudyAssistant = 0
    End If
    End With
    
    rs.Close
    Set rs = Nothing
            
    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Function GetSongTitle(SongNo As Long) As String
On Error GoTo ErrorTrap
Dim rstSongs As Recordset
   
    Set rstSongs = CMSDB.OpenRecordset("tblSongs", dbOpenDynaset)
    With rstSongs
    .FindFirst "SongNo = " & SongNo
    If Not .NoMatch Then
        GetSongTitle = !SongTitle
    Else
        GetSongTitle = ""
    End If
    End With
    Set rstSongs = Nothing
            
    Exit Function
ErrorTrap:
    EndProgram

End Function
Public Function GetTalkTitle(TalkNo As Long) As String
On Error GoTo ErrorTrap
Dim rstTalks As Recordset
   
    Set rstTalks = CMSDB.OpenRecordset("tblPublicTalkOutlines", dbOpenDynaset)
    With rstTalks
    .FindFirst "TalkNo = " & TalkNo
    If Not .NoMatch Then
        GetTalkTitle = !TalkTitle
    Else
        GetTalkTitle = ""
    End If
    End With
    Set rstTalks = Nothing
            
    Exit Function
ErrorTrap:
    EndProgram

End Function
Public Function GetServiceMtgItemType(ItemNo As Long) As String
On Error GoTo ErrorTrap
   
    Select Case ItemNo
    Case 0, 1
        GetServiceMtgItemType = "Item"
    Case 2
        GetServiceMtgItemType = "Prayer"
    End Select
            
    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Function SuspendedSQL(PersonFieldName As String, _
                                  TheDate_UK As Date, _
                                  UseWHERE As Boolean, _
                                  Optional vTaskCat, _
                                  Optional vTaskSubCat, _
                                  Optional vTask) As String

On Error GoTo ErrorTrap

Dim TaskCat As Long, TaskSubCat As Long, Task As Long
Dim strSQL As String

    If Not IsMissing(vTaskCat) Then
        TaskCat = CLng(vTaskCat)
        strSQL = strSQL & " AND TaskCategory = " & TaskCat & " "
    Else
        TaskCat = 0
    End If
    
    If Not IsMissing(vTaskSubCat) Then
        TaskSubCat = CLng(vTaskSubCat)
        strSQL = strSQL & " AND TaskSubCategory = " & TaskSubCat & " "
    Else
        TaskSubCat = 0
    End If
    
    If Not IsMissing(vTask) Then
        Task = CLng(vTask)
        strSQL = strSQL & " AND Task = " & Task & " "
    Else
        Task = 0
    End If
    
    If Task = 0 And TaskSubCat = 0 And TaskCat = 0 Then
        EndProgram "SuspendedSQL - must supply at least one Task tree element"
    End If
        
        
        
    SuspendedSQL = IIf(UseWHERE, " WHERE ", " AND ") & PersonFieldName & _
                        " NOT IN " & _
                        " (SELECT Person " & _
                        "FROM tblTaskPersonSuspendDates " & _
                        "WHERE (SuspendStartDate <= #" & Format(TheDate_UK, "mm/dd/yyyy") & _
                        "# AND SuspendEndDate >= #" & Format(TheDate_UK, "mm/dd/yyyy") & "#) " & _
                        strSQL & ") "
                                            
    
    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Function AttendantsToday(DateToCheck As Date, _
                                DateIsMondayDate As Boolean, _
                                MeetingType As cmsMeetingTypes, _
                                Optional AttendantType As cmsAttendantTypes = 0) _
                                As Collection

On Error GoTo ErrorTrap
Dim rsRota As Recordset, dteUS As String, lSunMtgDay As Long, lMidWkMtgDay As Long
Dim lNoAtts As Integer, lNoPlat As Integer, lNoSound As Integer, lNoMics As Integer
Dim l1stPos As Integer, i As Long, str As String
Dim lCount As Long, AttendantsFound As New Collection
Dim lPerson As Long

    RemoveAllItemsFromCollection AttendantsFound

    If GlobalParms.GetValue("RotaTimesPerWeek", "NumVal") = 1 Then
        'SPAM bros listed once per week (under monday date) so translate
        ' supplied date to monday date
        dteUS = Format$(GetDateOfGivenDay(DateToCheck, vbMonday, False), "mm/dd/yyyy")
    Else
        If Not DateIsMondayDate Then
            'actual date
            dteUS = Format$(DateToCheck, "mm/dd/yyyy")
        Else
            'monday date, so need to derive actual SPAM rota date depending
            ' on meeting type
            If IsCOVisitWeek(DateToCheck) Then
                Select Case MeetingType
                Case cmsSundayMtg
                    dteUS = Format$(GetDateOfGivenDay(DateToCheck, vbSunday, True), "mm/dd/yyyy")
                Case cmsMidWkMtg
                    dteUS = Format$(GetDateOfGivenDay(DateToCheck, vbTuesday, True), "mm/dd/yyyy")
                Case cmsBookstudy
                    dteUS = Format$(GetDateOfGivenDay(DateToCheck, vbThursday, True), "mm/dd/yyyy")
                End Select
            Else
                lSunMtgDay = GlobalParms.GetValue("SundayMeetingDay", "NumVal")
                lMidWkMtgDay = GlobalParms.GetValue("MidWeekMeetingDay", "NumVal")
            
                Select Case MeetingType
                Case cmsSundayMtg
                    dteUS = Format$(GetDateOfGivenDay(DateToCheck, lSunMtgDay, True), "mm/dd/yyyy")
                Case cmsMidWkMtg
                    dteUS = Format$(GetDateOfGivenDay(DateToCheck, lMidWkMtgDay, True), "mm/dd/yyyy")
                End Select
            End If
        End If
    End If
                
'    If TableExists("tblPrevRota") Then
'
'        CMSDB.Execute ("DELETE * FROM tblPrevRota " & _
'                       "WHERE RotaDate IN (" & _
'                       "SELECT RotaDate FROM tblRota)")
'
'        Set rsRota = CMSDB.OpenRecordset("SELECT * " & _
'                                        "FROM tblPrevRota " & _
'                                        "WHERE RotaDate = #" & dteUS & "# " & _
'                                        "UNION ALL " & _
'                                        "SELECT * " & _
'                                        "FROM tblRota " & _
'                                        "WHERE RotaDate = #" & dteUS & "# " _
'                                        , dbOpenSnapshot)
'
'    Else
'        Set rsRota = CMSDB.OpenRecordset("SELECT * " & _
'                                        "FROM tblRota " & _
'                                        "WHERE RotaDate = #" & dteUS & "# " _
'                                        , dbOpenSnapshot)
'    End If
    
    Set rsRota = GetSPAMRota(CDate(Format(dteUS, "mm/dd/yyyy")), _
                             CDate(Format(dteUS, "mm/dd/yyyy")))
                                        
    With rsRota
    If .BOF Then
        Set AttendantsToday = AttendantsFound
        rsRota.Close
        Set rsRota = Nothing
        Exit Function
    End If
    
    '
    'Now use rota's structure to determine Bros on selected week
    '
    AcquireRotaStructure2 lNoAtts, lNoMics, lNoSound, lNoPlat, l1stPos, rsRota
    
    If AttendantType = 0 Then 'all att types searched
        For i = l1stPos To .Fields.Count - 1
            lPerson = .Fields(i)
            AttendantsFound.Add lPerson
        Next i
    Else
        'search only fields for specified Att type....
        'Need to construct field name and use DAO trickery...
        Select Case AttendantType
        Case cmsAttendantAtt
            str = "Attendant_"
            lCount = lNoAtts
        Case cmsMicrophonesAtt
            str = "RovingMic_"
            lCount = lNoMics
        Case cmsPlatformAtt
            str = "Platform_"
            lCount = lNoPlat
        Case cmsSoundAtt
            str = "Sound_"
            lCount = lNoSound
        End Select
        
        For i = 1 To lCount
            lPerson = .Fields(str & Format$(i, "00"))
            AttendantsFound.Add lPerson
        Next i
        
    End If
    
    End With
    
    rsRota.Close
    Set rsRota = Nothing
                                        
    Set AttendantsToday = AttendantsFound

    Exit Function
    
ErrorTrap:
    EndProgram
End Function

Public Function AttendantsAfterDate(DateToCheck As Date, _
                                    NumberOfMtgsAfterDate As Long, _
                                    Optional AttendantType As cmsAttendantTypes = 0) _
                                    As Collection

On Error GoTo ErrorTrap
Dim rsRota As Recordset, dteUS As String, lSunMtgDay As Long, lMidWkMtgDay As Long
Dim lNoAtts As Integer, lNoPlat As Integer, lNoSound As Integer, lNoMics As Integer
Dim l1stPos As Integer, i As Long, str As String
Dim lCount As Long, AttendantsFound As New Collection
Dim lPerson As Long, dteRotaDate As Date

    RemoveAllItemsFromCollection AttendantsFound
    
    dteUS = Format$(DateToCheck, "mm/dd/yyyy")
                            
'    If TableExists("tblPrevRota") Then
'
'        CMSDB.Execute ("DELETE * FROM tblPrevRota " & _
'                       "WHERE RotaDate IN (" & _
'                       "SELECT RotaDate FROM tblRota)")
'
'        Set rsRota = CMSDB.OpenRecordset("SELECT * " & _
'                                        "FROM tblPrevRota " & _
'                                        "WHERE RotaDate > #" & dteUS & "# " & _
'                                        "UNION ALL " & _
'                                        "SELECT * " & _
'                                        "FROM tblRota " & _
'                                        "WHERE RotaDate > #" & dteUS & "# " & _
'                                        "ORDER BY RotaDate " _
'                                        , dbOpenSnapshot)
'    Else
'        Set rsRota = CMSDB.OpenRecordset("SELECT * " & _
'                                        "FROM tblRota " & _
'                                        "WHERE RotaDate > #" & dteUS & "# " & _
'                                        "ORDER BY RotaDate " _
'                                        , dbOpenSnapshot)
'    End If

    'find att dates AFTER the supplied meeting date
    Set rsRota = GetSPAMRota(CDate(DateAdd("d", 1, Format(dteUS, "mm/dd/yyyy"))), _
                            CDate("31/12/9999"))

                                        
    With rsRota
    If .BOF Then
        Set AttendantsAfterDate = AttendantsFound
        rsRota.Close
        Set rsRota = Nothing
        Exit Function
    End If
    
    .Move NumberOfMtgsAfterDate
    
    If .EOF Then
        Set AttendantsAfterDate = AttendantsFound
        rsRota.Close
        Set rsRota = Nothing
        Exit Function
    End If
    
    dteRotaDate = !RotaDate
    AttendantsFound.Add dteRotaDate
    
    '
    'Now use rota's structure to determine Bros on selected week
    '
    AcquireRotaStructure2 lNoAtts, lNoMics, lNoSound, lNoPlat, l1stPos, rsRota
    
    If AttendantType = 0 Then 'all att types searched
        For i = l1stPos To .Fields.Count - 1
            lPerson = .Fields(i)
            AttendantsFound.Add lPerson
        Next i
    Else
        'search only fields for specified Att type....
        'Need to construct field name and use DAO trickery...
        Select Case AttendantType
        Case cmsAttendantAtt
            str = "Attendant_"
            lCount = lNoAtts
        Case cmsMicrophonesAtt
            str = "RovingMic_"
            lCount = lNoMics
        Case cmsPlatformAtt
            str = "Platform_"
            lCount = lNoPlat
        Case cmsSoundAtt
            str = "Sound_"
            lCount = lNoSound
        End Select
        
        For i = 1 To lCount
            lPerson = .Fields(str & Format$(i, "00"))
            AttendantsFound.Add lPerson
        Next i
        
    End If
    
    End With
    
    rsRota.Close
    Set rsRota = Nothing
                                        
    Set AttendantsAfterDate = AttendantsFound

    Exit Function
    
ErrorTrap:
    EndProgram
End Function

Public Function GetSPAMRota(ByVal StartDate As Date, _
                            ByVal EndDate As Date, _
                            Optional SQL_String_OUT As String) As Recordset
On Error GoTo ErrorTrap
   
Dim str As String, rs As Recordset
Dim str2 As String

    'get list of all SPAM rotas. Use MAX function on the rota version number to remove dupes
    str = "SELECT RotaTbl & ' ' & VerNo AS RotaTable FROM (" & _
          "SELECT LEFT(RotaTableName,33) as RotaTbl, " & _
          "MAX(RIGHT(RotaTableName, 6)) as VerNo " & _
          "FROM tblStoredSPAMRotas " & _
          "GROUP BY LEFT(RotaTableName,33))"
    
    Set rs = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    'for all rotas that match the current structure, add to a big UNION statement
    With rs
    Do Until .BOF Or .EOF
        If CompatibleSPAMRotaStructure(!RotaTable) Then
            If str2 <> "" Then
                str2 = str2 & vbCrLf & " UNION ALL "
            End If
            str2 = str2 & vbCrLf & " SELECT * FROM [" & !RotaTable & "] "
        End If
    
        .MoveNext
    Loop
    End With
    
    'create a SELECT stmt for the resulting recset
    If str2 <> "" Then
    
        str2 = "SELECT * FROM (" & str2 & ") " & _
               "WHERE RotaDate BETWEEN " & GetDateStringForSQLWhere(CStr(StartDate)) & _
                " AND " & GetDateStringForSQLWhere(CStr(EndDate)) & _
                " ORDER BY RotaDate "
        
        rs.Close
        Set rs = Nothing
                
        Set GetSPAMRota = CMSDB.OpenRecordset(str2, dbOpenDynaset)
        SQL_String_OUT = str2
        
    Else 'no rotas available - return an empty dummy recset
    
        Set GetSPAMRota = CMSDB.OpenRecordset("SELECT * FROM tblConstants WHERE FldName = 'DUMMY'", dbOpenDynaset)
        SQL_String_OUT = ""
    End If
        
            
    Exit Function
ErrorTrap:
    EndProgram

End Function
Public Function CompatibleSPAMRotaStructure(ByVal TableName As String) As Boolean
On Error GoTo ErrorTrap
   
Dim str As String, rs As Recordset
Dim TempNoAttending As Integer
Dim TempNoOnMics As Integer
Dim TempNoOnSound As Integer
Dim TempNoOnPlatform As Integer
Dim FirstJobPos As Integer

    Set rs = CMSDB.OpenRecordset(TableName, dbOpenDynaset)
    
    AcquireRotaStructure2 TempNoAttending, TempNoOnMics, TempNoOnSound, _
                          TempNoOnPlatform, FirstJobPos, rs
                          
    With GlobalParms
    
    If TempNoAttending = .GetValue("NoAttending", "NumVal") And _
       TempNoOnMics = .GetValue("NoOnMics", "NumVal") And _
       TempNoOnSound = .GetValue("NoOnSound", "NumVal") And _
       TempNoOnPlatform = .GetValue("NoOnPlatform", "NumVal") And _
       FirstJobPos = .GetValue("Pos1stJobOnRota", "NumVal") Then
       
        CompatibleSPAMRotaStructure = True
    Else
        CompatibleSPAMRotaStructure = False
    End If
    
    End With
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Function GetNoPublishersThisMonth(TheMonth As Long, TheYear As Long) As Long
On Error GoTo ErrorTrap
        
Dim TheString As String, DateString_US As String, rstRecSet As Recordset

    'Construct SQL to give number of publishers in selected month
    
    DateString_US = CStr(TheMonth) & "/01/" & CStr(TheYear)

    TheString = "SELECT COUNT(tblPublisherDates.PersonID) as CountPubs " & _
                "FROM  (SELECT DISTINCT tblPublisherDates.PersonID " & _
                        "FROM tblPublisherDates " & _
                        "WHERE StartDate <= #" & _
                         DateString_US & _
                        "# AND EndDate >= #" & DateString_US & "# " & _
                        " AND StartReason <> 2)"
                
    Set rstRecSet = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)
    
    If Not IsNull(rstRecSet!CountPubs) Then
        GetNoPublishersThisMonth = rstRecSet!CountPubs
    Else
        GetNoPublishersThisMonth = 0
    End If
    
    rstRecSet.Close
    Set rstRecSet = Nothing
        
    
    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Function GetTotalPubsThisYear(TheServiceYear As Long) As Long
On Error GoTo ErrorTrap
        
Dim TheString As String, rstRecSet As Recordset

    TheString = "SELECT COUNT(tblPublisherDates.PersonID) as CountPubs " & _
                "FROM  (SELECT DISTINCT tblPublisherDates.PersonID " & _
                        "FROM tblPublisherDates " & _
                        "WHERE StartDate <= #08/01/" & TheServiceYear & _
                        "# AND EndDate >= #09/01/" & TheServiceYear - 1 & "# " & _
                        " AND StartReason <> 2)"
                
    Set rstRecSet = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)
    
    If Not IsNull(rstRecSet!CountPubs) Then
        GetTotalPubsThisYear = rstRecSet!CountPubs
    Else
        GetTotalPubsThisYear = 0
    End If
    
    rstRecSet.Close
    Set rstRecSet = Nothing
        
    
    Exit Function
ErrorTrap:
    EndProgram
    
End Function
Public Function GetTotalPubsInPeriod(StartDateUK As Date, EndDateUK As Date) As Long
On Error GoTo ErrorTrap
        
Dim TheString As String, rstRecSet As Recordset

    TheString = "SELECT COUNT(tblPublisherDates.PersonID) as CountPubs " & _
                "FROM  (SELECT DISTINCT tblPublisherDates.PersonID " & _
                        "FROM tblPublisherDates " & _
                        "WHERE StartDate <= " & GetDateStringForSQLWhere(CStr(EndDateUK)) & _
                        " AND EndDate >= " & GetDateStringForSQLWhere(CStr(StartDateUK)) & _
                        " AND StartReason <> 2)"
                
    Set rstRecSet = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)
    
    If Not IsNull(rstRecSet!CountPubs) Then
        GetTotalPubsInPeriod = rstRecSet!CountPubs
    Else
        GetTotalPubsInPeriod = 0
    End If
    
    rstRecSet.Close
    Set rstRecSet = Nothing
        
    
    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Function GetAveragePubsInPeriod(StartDateUK As Date, EndDateUK As Date) As Double
On Error GoTo ErrorTrap
        
Dim TheString As String, rstRecSet As Recordset
Dim dteStart As Date, dteEnd As Date, dteTemp As Date
Dim lTotPubs As Long

    dteStart = CDate("01/" & Month(StartDateUK) & "/" & year(StartDateUK))
    dteEnd = CDate("01/" & Month(EndDateUK) & "/" & year(EndDateUK))
    dteTemp = dteStart
    lTotPubs = 0
    
    Do Until DateDiff("d", dteEnd, dteTemp) > 0
        
        TheString = "SELECT COUNT(tblPublisherDates.PersonID) as CountPubs " & _
                    "FROM  (SELECT DISTINCT tblPublisherDates.PersonID " & _
                            "FROM tblPublisherDates " & _
                            "WHERE StartDate <= " & GetDateStringForSQLWhere(CStr(dteTemp)) & _
                            " AND EndDate >= " & GetDateStringForSQLWhere(CStr(dteTemp)) & _
                            " AND StartReason <> 2) "
                    
        Set rstRecSet = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)
    
        lTotPubs = lTotPubs + CLng(HandleNull(rstRecSet!CountPubs))
        
        dteTemp = DateAdd("m", 1, dteTemp)
        
    Loop
    
    rstRecSet.Close
    Set rstRecSet = Nothing
    
    GetAveragePubsInPeriod = lTotPubs / (DateDiff("m", dteStart, dteEnd) + 1)
    
    Exit Function
ErrorTrap:
    EndProgram
    
End Function


Public Sub TidyUpBookGroupTables()
On Error GoTo ErrorTrap
        
Dim TheString As String, rst As Recordset, rst2 As Recordset

    'remove pubs from non-existant groups
    CMSDB.Execute "DELETE FROM tblBookGroupMembers " & _
                  "WHERE GroupNo NOT IN " & _
                  "(SELECT GroupNo FROM tblBookGroups)"
    
    
    'remove attendance figures for non-existant groups
    CMSDB.Execute "DELETE FROM tblGroupAttendance " & _
                  "WHERE GroupNo NOT IN " & _
                  "(SELECT GroupNo FROM tblBookGroups)"
    
    
    'update total Group atts on tblMeetingAttendance
    TheString = "SELECT SUM(Attendance) AS SumAtt, WeekBeginning " & _
                "FROM tblGroupAttendance " & _
                "GROUP BY WeekBeginning"
                
    Set rst = CMSDB.OpenRecordset(TheString, dbOpenDynaset)
    
    With rst
    If .BOF Or .EOF Then
        CMSDB.Execute "DELETE FROM tblMeetingAttendance " & _
                      "WHERE MeetingTypeID = 4"
    Else
        
        Set rst2 = CMSDB.OpenRecordset("SELECT MeetingTypeID, WeekBeginning, Attendance " & _
                                     "FROM tblMeetingAttendance " & _
                                     "WHERE MeetingTypeID = 4", dbOpenDynaset)
        
        Do Until .EOF
            rst2.FindFirst "WeekBeginning = " & GetDateStringForSQLWhere(CStr(!WeekBeginning))
            If Not rst2.NoMatch Then
                If HandleNull(!SumAtt) = 0 Then
                    If Not IsCOVisitWeek(!WeekBeginning) Then 'just in case
                        rst2.Delete
                    End If
                Else
                    rst2.Edit
                    rst2!Attendance = !SumAtt
                    rst2.Update
                End If
            Else
                rst2.AddNew
                rst2!MeetingTypeID = 4
                rst2!WeekBeginning = !WeekBeginning
                rst2!Attendance = !SumAtt
                rst2.Update
            End If
            
            .MoveNext
        Loop
    
    End If
    End With
    
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    rst2.Close
    Set rst2 = Nothing
    
    
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub




