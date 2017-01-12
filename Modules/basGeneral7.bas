Attribute VB_Name = "basGeneral7"
Option Explicit

Public Sub GenerateAddressList(sGroupList As String)
Dim arr() As String, i As Long, bDataFound As Boolean, sNum As String, Prnt As cmsPrintUsingWord
On Error GoTo ErrorTrap

    Screen.MousePointer = vbHourglass
    
    Prnt = PrintUsingWord
    
    If Prnt = cmsUseMSDatareport Then
        If Not PrintAddressList(sGroupList) Then
            MsgBox "Nothing to print", vbOKOnly + vbInformation, AppName
        End If
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    'Print using Word...
    If Prnt = cmsUseWord Then
        If sGroupList <> "" Then
            arr() = Split(sGroupList, ",")
            
            For i = 0 To UBound(arr)
                sNum = Trim$(arr(i))
                If IsNumber(sNum, False, False, False) Then
                    If GenerateAddressListForGroup(CLng(sNum)) Then
                        bDataFound = True
                    End If
                End If
            Next i
        Else
            bDataFound = GenerateAddressListForGroup(0)
        End If
        
        If Not bDataFound Then
            MsgBox "Nothing to print", vbOKOnly + vbInformation, AppName
        Else
            MsgBox "Address list generated in Word", vbOKOnly + vbInformation, AppName
            
        End If
    End If

    Screen.MousePointer = vbNormal

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub
Public Function CheckDataExists(TheSQL As String, NoDataMessage As String) As Boolean
On Error GoTo ErrorTrap
Dim rs As Recordset

    Set rs = CMSDB.OpenRecordset(TheSQL, dbOpenForwardOnly)
    
    If rs.BOF Or rs.EOF Then
        MsgBox NoDataMessage, vbOKOnly + vbInformation, AppName
        CheckDataExists = False
    Else
        CheckDataExists = True
    End If
    
    rs.Close
    Set rs = Nothing

    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function SetUpTranCode(Update As Boolean, _
                              TranCode As String, _
                              Description As String, _
                              InOutTypeID As Long, _
                              AutoDayOfMonth As Long, _
                              Amount As Double, _
                              OnReceipt As Boolean, _
                              Ref As Long, _
                              AccountID As Long, _
                              TfrAccountID As Long) As Boolean
On Error GoTo ErrorTrap
Dim rs As Recordset

    'THIS IS CALLED BY UPGRADE CODE ONLY. DON'T RE-USE!!!!!

    Set rs = CMSDB.OpenRecordset("tblTransactionTypes", dbOpenDynaset)
    
    With rs
    
    .FindFirst "TranCode = '" & TranCode & "' AND InOutTypeID = " & InOutTypeID
    
    If Update Then
        If .NoMatch Then
            SetUpTranCode = False
            GoTo GetOut
        End If
        .Edit
    Else
        If Not .NoMatch Then
            SetUpTranCode = False
            GoTo GetOut
        End If
        .AddNew
        !TranCode = TranCode
        !InOutTypeID = InOutTypeID
    End If
            
    !Description = DoubleUpSingleQuotes(Description)
    !AutoDayOfMonth = AutoDayOfMonth
    !Amount = Amount
    !OnReceipt = OnReceipt
    !Ref = Ref
    !AccountID = AccountID
    !TfrAccountID = TfrAccountID
    .Update
    
    SetUpTranCode = True
    
    End With
    
GetOut:
    rs.Close
    Set rs = Nothing

    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function GetPublicTalkCoordinator(CongNo As Long) As Long
On Error GoTo ErrorTrap
Dim rs As Recordset, TheSQL As String

    TheSQL = "SELECT ID " & _
           "FROM tblNameAddress INNER JOIN tblTaskAndPerson " & _
           "         ON tblNameAddress.ID = tblTaskAndPerson.Person " & _
           "WHERE GenderMF = 'M' " & _
           " AND Task = 73 AND CongNo = " & CongNo

    Set rs = CMSDB.OpenRecordset(TheSQL, dbOpenForwardOnly)
    
    If rs.BOF Or rs.EOF Then
        GetPublicTalkCoordinator = 0
    Else
        GetPublicTalkCoordinator = rs!ID
    End If
    
    rs.Close
    Set rs = Nothing

    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function GetPublicMeetingDetails(MeetingDate As Date, _
                                        CongWhereMeetingIs As Long, _
                                        Outbound As Boolean) As PublicMeetingDetails
On Error GoTo ErrorTrap
Dim rs As Recordset, mtg As PublicMeetingDetails

    Set rs = CMSDB.OpenRecordset("SELECT MeetingDate, " & _
                                       "SpeakerID, " & _
                                       "SpeakerID2, " & _
                                       "TalkNo, " & _
                                       "ChairmanID, " & _
                                       "WTReaderID, " & _
                                       "Info, " & _
                                       "Provisional " & _
                                    "FROM tblPublicMtgSchedule " & _
                                    "WHERE MeetingDate = #" & _
                                      Format(MeetingDate, "mm/dd/yyyy") & "# " & _
                                      " AND CongNoWhereMtgIs = " & CongWhereMeetingIs _
                                                              , dbOpenDynaset)
                                                              
    With rs
    
    If Not .BOF Then
        mtg.MeetingDate = HandleNull(!MeetingDate)
        mtg.Chairman = HandleNull(!ChairmanID)
        mtg.Reader = HandleNull(!WTReaderID)
        mtg.SpeakerA = HandleNull(!SpeakerID)
        mtg.SpeakerB = HandleNull(!SpeakerID2)
        mtg.TalkNo = HandleNull(!TalkNo)
        mtg.Info = HandleNull(!Info, "")
        mtg.TalkCoordinator = GetPublicTalkCoordinator(IIf(Outbound, CongWhereMeetingIs, CongregationMember.CongForPerson(!SpeakerID)))
        mtg.CongWhereMtgIs = CongWhereMeetingIs
        mtg.Provisional = HandleNull(!Provisional, False)
    Else
        mtg.MeetingDate = 0
        mtg.Chairman = 0
        mtg.Reader = 0
        mtg.SpeakerA = 0
        mtg.SpeakerB = 0
        mtg.TalkNo = 0
        mtg.TalkCoordinator = 0
        mtg.Info = ""
        mtg.CongWhereMtgIs = 0
        mtg.Provisional = False
    End If
    
    End With
    
    GetPublicMeetingDetails = mtg
    
    rs.Close
    Set rs = Nothing

    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function GetGoneInactiveInPeriod(StartDate As Date, EndDate As Date) As Recordset
On Error GoTo ErrorTrap

Dim TheString As String, DateString_US As String, DateString2_US As String
    
    
    'Date from which we're interested - as selected in combos
    DateString_US = Format(StartDate, "mm/dd/yyyy")
    'Date up to which we're interested, as selected in combos
    DateString2_US = Format(EndDate, "mm/dd/yyyy")
    
    
    'Construct SQL to give us names of inactive publishers
    'This is equivalent of "SELECT COUNT(DISTINCT PersonID)" - doesn't work in Access.
       
    TheString = "SELECT DISTINCT a.PersonID " & _
                "FROM (tblMissingReports AS a INNER JOIN " & _
                "       tblInactivePubs AS b ON a.MissingReportGroupID = b.MissingReportGroupID) " & _
                "     INNER JOIN tblPublisherDates ON tblPublisherDates.PersonID = a.PersonID " & _
                "WHERE EXISTS " & _
                          " (SELECT 1 FROM tblMinReports " & _
                          "  WHERE tblMinReports.ActualMinPeriod = DateAdd('m', -6, b.StartDate) " & _
                          "  AND a.PersonID = tblMinReports.PersonID " & _
                          "  AND tblMinReports.NoHours > 0) " & _
                  "AND b.StartDate >= #" & _
                       DateString_US & _
                "# AND b.StartDate <= #" & DateString2_US & "# " & _
                "AND tblPublisherDates.StartDate <= #" & _
                 DateString2_US & _
                "# AND tblPublisherDates.EndDate >= #" & DateString_US & "# "

    Set GetGoneInactiveInPeriod = CMSDB.OpenRecordset(TheString, dbOpenDynaset)

    Exit Function
ErrorTrap:
    EndProgram
    
End Function
Public Function GetReactivatedInPeriod(StartDate As Date, EndDate As Date) As Recordset
On Error GoTo ErrorTrap

Dim TheString As String, DateString_US As String, DateString2_US As String
    
    
    'Date from which we're interested - as selected in combos
    DateString_US = Format(StartDate, "mm/dd/yyyy")
    'Date up to which we're interested, as selected in combos
    DateString2_US = Format(EndDate, "mm/dd/yyyy")
    
    
    'Construct SQL to give us names of inactive publishers
    'This is equivalent of "SELECT COUNT(DISTINCT PersonID)" - doesn't work in Access.
       
    TheString = "SELECT DISTINCT a.PersonID " & _
                "FROM tblMinReports AS a INNER JOIN  tblPublisherDates ON tblPublisherDates.PersonID = a.PersonID " & _
                "WHERE (((Exists (SELECT 1  FROM " & _
                                  "tblMissingReports INNER JOIN tblInactivePubs ON " & _
                                  "tblMissingReports.MissingReportGroupID = " & _
                                  "tblInactivePubs.MissingReportGroupID " & _
                                   " WHERE tblInactivePubs.EndDate = DateAdd('m', -1, a.ActualMinPeriod) " & _
                                   " AND tblMissingReports.PersonID = a.PersonID)))) " & _
                  " AND a.NoHours > 0 " & _
                  "AND ActualMinPeriod >= #" & _
                       DateString_US & _
                "# AND ActualMinPeriod <= #" & DateString2_US & "# " & _
                "AND tblPublisherDates.StartDate <= #" & _
                 DateString2_US & _
                "# AND tblPublisherDates.EndDate >= #" & DateString_US & "# "

    Set GetReactivatedInPeriod = CMSDB.OpenRecordset(TheString, dbOpenDynaset)

    Exit Function
ErrorTrap:
    EndProgram
    
End Function
Public Function GetNumberActivePubsInPeriod(StartDate As Date, EndDate As Date) As Long
On Error GoTo ErrorTrap

Dim TheString As String, DateString_US As String, DateString2_US As String
Dim rstRecSet As Recordset
    
    
    'Date from which we're interested - as selected in combos
    DateString_US = Format(StartDate, "mm/dd/yyyy")
    'Date up to which we're interested, as selected in combos
    DateString2_US = Format(EndDate, "mm/dd/yyyy")
    
    
    'Construct SQL to give us number of active publishers
    'This is equivalent of "SELECT COUNT(DISTINCT PersonID)" - doesn't work in Access.
    TheString = "SELECT COUNT(tblPublisherDates.PersonID) as CountActive " & _
                "FROM (SELECT DISTINCT tblPublisherDates.PersonID " & _
                "FROM tblPublisherDates " & _
                "       INNER JOIN tblNameAddress AS c " & _
                        "        ON c.ID = tblPublisherDates.PersonID " & _
                "WHERE StartDate <= #" & _
                 DateString2_US & _
                "# AND EndDate >= #" & DateString_US & "# " & _
                "AND Active = TRUE " & _
                "AND tblPublisherDates.PersonID NOT IN " & _
                        "(SELECT DISTINCT PersonID " & _
                        "FROM tblMissingReports INNER JOIN tblInactivePubs ON " & _
                        "    (tblMissingReports.MissingReportGroupID = " & _
                            " tblInactivePubs.MissingReportGroupID) " & _
                        "WHERE StartDate <= #" & _
                         DateString2_US & _
                        "# AND EndDate >= #" & DateString_US & "#))"

    Set rstRecSet = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)
                                                              
    GetNumberActivePubsInPeriod = rstRecSet!CountActive
    
    rstRecSet.Close
    Set rstRecSet = Nothing

    Exit Function
ErrorTrap:
    EndProgram
    
End Function


Public Sub DeleteGiftAid(PersonID As Long)
Dim lGiftAidNo As Long, sDate As String
On Error GoTo ErrorTrap

    sDate = "#" & Format(DateAdd("d", -1, Now), "mm/dd/yyyy") & "# "

    lGiftAidNo = CongregationMember.GiftAidNoFromPerson(PersonID)
    
    CMSDB.Execute "UPDATE tblGiftAidPayerActiveDates " & _
                  "SET EndDate = " & sDate & _
                  "WHERE GiftAidNo = " & lGiftAidNo & _
                  " AND EndDate > " & sDate
                  
    CMSDB.Execute "UPDATE tblGiftAidPayers " & _
                  "SET PersonID = 0 " & _
                  "WHERE GiftAidNo = " & lGiftAidNo
                  
'    DeleteSomeRows "tblGiftAidPayers", "PersonID = ", PersonID

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Public Sub HideAllOpenForms(Optional UnHide As Boolean = False)
Dim frm As Form

On Error Resume Next

    If UnHide Then
        For Each frm In Forms
            If frm.Name <> "frmMainMenu" And _
               frm.Name <> "frmSetUpMenu" Then
                  frm.Visible = True          ' show  the form
            End If
        Next
    Else
        For Each frm In Forms
            If frm.Name <> "frmMainMenu" And _
               frm.Name <> "frmSetUpMenu" Then
                  frm.Visible = False        ' hide the form
            End If
        Next
    End If

End Sub

Public Function GetOrgName(OrgID As Long) As String
Dim rs As Recordset
On Error GoTo ErrorTrap

    Set rs = CMSDB.OpenRecordset("SELECT OrgName FROM tblExtOrgs " & _
                                "WHERE OrgID = " & OrgID, dbOpenSnapshot)
                                
    With rs
    
    If .BOF Then
        GetOrgName = ""
    Else
        GetOrgName = HandleNull(!OrgName, "")
    End If
    
    End With
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function GiftAidNoExists(GiftAidNo As Long) As Boolean
Dim rs As Recordset
On Error GoTo ErrorTrap

    Set rs = CMSDB.OpenRecordset("SELECT 1 " & _
                               "FROM tblGiftAidPayers " & _
                               "WHERE GiftAidNo = " & GiftAidNo, dbOpenForwardOnly)
                                       
    If rs.BOF Then
        GiftAidNoExists = False
    Else
        GiftAidNoExists = True
    End If
            
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function GiftAidNoActiveForDate(GiftAidNo As Long, TheDate As Date) As Boolean
Dim rs As Recordset
On Error GoTo ErrorTrap

    Set rs = CMSDB.OpenRecordset("SELECT 1 " & _
                               "FROM tblGiftAidPayerActiveDates " & _
                               "WHERE GiftAidNo = " & GiftAidNo & _
                               " AND #" & Format(TheDate, "mm/dd/yyyy") & "# " & _
                               "      >= StartDate AND #" & _
                               Format(TheDate, "mm/dd/yyyy") & "# " & _
                               "      <= EndDate ", dbOpenForwardOnly)
    
    If rs.BOF Then
        GiftAidNoActiveForDate = False
    Else
        GiftAidNoActiveForDate = True
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function IsTranTypeFromCongReceipt(TranCodeID As Long) As Boolean
Dim rs As Recordset
On Error GoTo ErrorTrap

    Set rs = CMSDB.OpenRecordset("SELECT OnReceipt FROM tblTransactionTypes " & _
                                "WHERE TranCodeID = " & TranCodeID, dbOpenSnapshot)
                                
    With rs
    
    If .BOF Then
        IsTranTypeFromCongReceipt = False
    Else
        If HandleNull(!OnReceipt, False) = True Then
            IsTranTypeFromCongReceipt = True
        Else
            IsTranTypeFromCongReceipt = False
        End If
    End If
    
    End With
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function TranTypeContributedAtGroups(TranCode As String) As Boolean
On Error GoTo ErrorTrap

    TranTypeContributedAtGroups = (InStr(1, GlobalParms.GetValue("TranTypesContributedAtGroups", "AlphaVal", ""), TranCode & ",") > 0)
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function

Public Function GetAccountAmountBetweenDates(StartDate_UK As String, _
                                             EndDate_UK As String, _
                                             OnReceipt As Boolean, _
                                             BookGroupID As Long, _
                                             Optional AccountID As Long = 0) As Double
                                             
Dim rs As Recordset, str As String, lGrpID As Long

On Error GoTo ErrorTrap

    If ValidDate(StartDate_UK) And ValidDate(EndDate_UK) Then
    Else
        GetAccountAmountBetweenDates = 0
        Exit Function
    End If
    
     

    str = "SELECT SUM(a.Amount) AS TotAmount " & _
         "FROM tblTransactionDates a " & _
         " INNER JOIN tblTransactionTypes b ON a.TranCodeID = b.TranCodeID " & _
        "WHERE TranDate BETWEEN " & GetDateStringForSQLWhere(StartDate_UK) & _
            " AND " & GetDateStringForSQLWhere(EndDate_UK) & _
            IIf(OnReceipt, " AND OnReceipt = TRUE ", "") & _
            " AND BookGroupNo IN (" & IIf(BookGroupID <= 0, "-1,0", BookGroupID) & ")" & _
            " AND a.AccountID = " & AccountID

    Set rs = CMSDB.OpenRecordset(str, dbOpenSnapshot)
    
    GetAccountAmountBetweenDates = HandleNull(rs!TotAmount, 0)
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function GetGeneralRecordset(SQL As String) As Recordset
                                             
On Error GoTo ErrorTrap

    Set GetGeneralRecordset = CMSDB.OpenRecordset(SQL, dbOpenDynaset)
        
    Exit Function
ErrorTrap:
    Call EndProgram
End Function

Public Sub ShowMessage(MessageText As String, _
                        DisplayTime_ms As Long, _
                        ParentForm As Form, _
                        Optional SwitchOff As Boolean = False, _
                        Optional Colour As ColorConstants = vbBlue, _
                        Optional ReFocusOnCallingForm As Boolean = False, _
                        Optional PauseSeconds As Integer = 0)
                                             
    On Error Resume Next
    
    If SwitchOff Or gbSuppressMsg Then Exit Sub
    
    If gbShowMsgBox Then
        MsgBox MessageText, vbOKOnly + vbInformation, AppName
        If ReFocusOnCallingForm Then
            ParentForm.SetFocus
        End If
        Exit Sub
    End If
    
    With frmMsg
    .DisplayText = MessageText
    .DisplayTime_ms = DisplayTime_ms
    .TextColour = Colour
    .ReFocusOnCallingForm = ReFocusOnCallingForm
    End With
    
    frmMsg.Show vbModeless, ParentForm
    If Err.number <> 0 Then
        Err.Clear
        frmMsg.Show vbModal, ParentForm
        If Err.number <> 0 Then
            MsgBox MessageText, vbOKOnly + vbInformation, AppName
        Else
            If PauseSeconds > 0 Then
                DoEvents
                Pause PauseSeconds
            End If
        End If
    Else
        If PauseSeconds > 0 Then
            DoEvents
            Pause PauseSeconds
        End If
    End If
    
    
    If ReFocusOnCallingForm Then
        ParentForm.SetFocus
    End If
    

End Sub


Public Function GetTransactionDetails(TransactionID As Long) As TransactionDetails
Dim rs As Recordset, str As String, TheTran As TransactionDetails
On Error GoTo ErrorTrap

    str = "SELECT a.TranCodeID, " & _
         "a.TranDate, " & _
         "a.Amount, " & _
         "a.TranDescription, " & _
         "a.RefNo, " & _
         "b.TranCode, " & _
         "b.Description AS TranTypeDesc, " & _
         "c.InOutTypeID, " & _
         "c.Description AS InOutTypeDesc, " & _
         "d.InOutID, " & _
         "d.Description AS InOutDesc, " & _
         "b.AutoDayOfMonth, " & _
         "a.BookGroupNo, " & _
         "a.TranSubTypeID, " & _
         "a.AccountID, " & _
         "a.TfrAccountID " & _
         "FROM ((tblTransactionDates a " & _
         " INNER JOIN tblTransactionTypes b ON a.TranCodeID = b.TranCodeID) " & _
         " INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
         " INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID " & _
        "WHERE TranID = " & TransactionID

    Set rs = CMSDB.OpenRecordset(str, dbOpenSnapshot)
                                
    With rs
    
    If Not .BOF Then
        TheTran.TransactionID = HandleNull(TransactionID)
        TheTran.TransactionCodeID = HandleNull(!TranCodeID)
        TheTran.TransactionCode = HandleNull(!TranCode, "")
        TheTran.InOutTypeID = HandleNull(!InOutTypeID)
        TheTran.InOutID = HandleNull(!InOutID)
        TheTran.InOutTypeDescription = HandleNull(!InOutTypeDesc, "")
        TheTran.InOutDescription = HandleNull(!InOutDesc, "")
        TheTran.TransactionDate = HandleNull(!TranDate)
        TheTran.Amount = HandleNull(!Amount)
        TheTran.TransactionDescription = HandleNull(!TranDescription, "")
        TheTran.TransactionTypeDescription = HandleNull(!TranTypeDesc, "")
        TheTran.RefNo = HandleNull(!RefNo)
        TheTran.AutoDayOfMonth = HandleNull(!AutoDayOfMonth)
        TheTran.BookGroupNo = HandleNull(!BookGroupNo)
        TheTran.TransactionSubTypeID = HandleNull(!TranSubTypeID)
        TheTran.AccountID = HandleNull(!AccountID)
        TheTran.TfrAccountID = HandleNull(!TfrAccountID)
    Else
        TheTran.TransactionID = 0
        TheTran.TransactionCodeID = 0
        TheTran.TransactionCode = ""
        TheTran.InOutTypeID = 0
        TheTran.InOutID = 0
        TheTran.InOutTypeDescription = ""
        TheTran.InOutDescription = ""
        TheTran.TransactionDate = 0
        TheTran.Amount = 0
        TheTran.TransactionDescription = ""
        TheTran.TransactionTypeDescription = ""
        TheTran.RefNo = 0
        TheTran.AutoDayOfMonth = 0
        TheTran.BookGroupNo = -1
        TheTran.TransactionSubTypeID = 0
        TheTran.AccountID = -1
        TheTran.TfrAccountID = -1
    End If
    
    End With
    
    GetTransactionDetails = TheTran
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function

Public Function GetTranTypeDesc_TOP1(TransCode As String, Optional bIncludeSuppressed As Boolean = False) As String
Dim rs As Recordset, str As String
On Error GoTo ErrorTrap

    str = "SELECT TOP 1 Description AS TranTypeDesc " & _
            "FROM tblTransactionTypes " & _
            "WHERE TranCode = '" & TransCode & "' " & _
            IIf(bIncludeSuppressed, "", " AND Suppressed = FALSE ") & _
            "ORDER BY TranCodeID "

    Set rs = CMSDB.OpenRecordset(str, dbOpenSnapshot)
                                
    With rs
    
    If Not .BOF Then
        GetTranTypeDesc_TOP1 = !TranTypeDesc
    Else
        GetTranTypeDesc_TOP1 = ""
    End If
    
    End With
       
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function


Public Function GetTransactionCodeStuff(TransCodeID As Long) As TransactionDetails
Dim rs As Recordset, str As String, TheTran As TransactionDetails
On Error GoTo ErrorTrap

    str = "SELECT b.TranCode, " & _
          "b.TranCodeID, " & _
         "b.Description AS TranTypeDesc, " & _
         "b.Ref, " & _
         "c.InOutTypeID, " & _
         "c.Description AS InOutTypeDesc, " & _
         "d.InOutID, " & _
         "d.Description AS InOutDesc, " & _
         "b.AutoDayOfMonth, " & _
         "b.Amount, " & _
         "b.OnReceipt, " & _
         "b.AccountID, " & _
         "b.TfrAccountID, " & _
         "b.Suppressed " & _
         "FROM (tblTransactionTypes b " & _
         " INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
         " INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID " & _
        "WHERE TranCodeID = " & TransCodeID

    Set rs = CMSDB.OpenRecordset(str, dbOpenSnapshot)
                                
    With rs
    
    If Not .BOF Then
        TheTran.TransactionID = 0
        TheTran.TransactionCodeID = HandleNull(!TranCodeID)
        TheTran.TransactionCode = HandleNull(!TranCode, "")
        TheTran.InOutTypeID = HandleNull(!InOutTypeID)
        TheTran.InOutID = HandleNull(!InOutID)
        TheTran.InOutTypeDescription = HandleNull(!InOutTypeDesc, "")
        TheTran.InOutDescription = HandleNull(!InOutDesc, "")
        TheTran.TransactionDate = 0
        TheTran.Amount = HandleNull(!Amount, 0)
        TheTran.TransactionDescription = ""
        TheTran.TransactionTypeDescription = HandleNull(!TranTypeDesc, "")
        TheTran.RefNo = HandleNull(!Ref, 0)
        TheTran.AutoDayOfMonth = HandleNull(!AutoDayOfMonth)
        TheTran.OnReceipt = HandleNull(!OnReceipt, False)
        TheTran.Suppressed = HandleNull(!Suppressed, False)
        TheTran.AccountID = HandleNull(!AccountID, -1)
        TheTran.TfrAccountID = HandleNull(!TfrAccountID, -1)
    Else
        TheTran.TransactionID = 0
        TheTran.TransactionCodeID = 0
        TheTran.TransactionCode = ""
        TheTran.InOutTypeID = 0
        TheTran.InOutID = 0
        TheTran.InOutTypeDescription = ""
        TheTran.InOutDescription = ""
        TheTran.TransactionDate = 0
        TheTran.Amount = 0
        TheTran.TransactionDescription = ""
        TheTran.TransactionTypeDescription = ""
        TheTran.RefNo = 0
        TheTran.AutoDayOfMonth = 0
        TheTran.OnReceipt = False
        TheTran.Suppressed = False
        TheTran.AccountID = -1
        TheTran.TfrAccountID = -1
    End If
    
    End With
    
    GetTransactionCodeStuff = TheTran
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function GetTranCodeIDFromTranCode(TransCode As String, InOutTypeID As Long) As Long
Dim rs As Recordset, str As String
On Error GoTo ErrorTrap

    str = "SELECT TranCodeID " & _
         "FROM tblTransactionTypes " & _
        "WHERE TranCode = '" & TransCode & "' " & _
        " AND InOutTypeID = " & InOutTypeID

    Set rs = CMSDB.OpenRecordset(str, dbOpenSnapshot)
                                
    With rs
    
    If Not .BOF Then
        GetTranCodeIDFromTranCode = rs!TranCodeID
    Else
        GetTranCodeIDFromTranCode = -1
    End If
    
    End With
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Sub CheckIfPersonInBookGroup(PersonID As Long)
On Error GoTo ErrorTrap

    If CongregationMember.BookGroup(PersonID) = -1 Then
        If MsgBox(CongregationMember.FirstAndLastName(PersonID) & _
                    " has not been assigned to a group. Do you want to do this now?", vbYesNo + vbQuestion, AppName) = vbYes Then
            If FormIsOpen("frmBookGroupMembers") Then
                Unload frmBookGroupMembers
            End If
            lnkPersonID = PersonID
            frmBookGroupMembers.Show vbModal
        End If
    End If
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Public Function GenerateAddressListForGroup(GroupNo As Long) As Boolean

On Error GoTo ErrorTrap

Dim sSQL As String, rstPrintTable As Recordset, rs As Recordset, sAddr As String
Dim sMbl1 As String, sMbl2 As String, sMbl As String

    'first build the print table
    
    DelAllRows "tblPrintAddresses"
    
    '
    'Driving Recset to get all data
    '
    If GroupNo > 0 Then
        sSQL = "SELECT ID " & _
                 "FROM tblNameAddress INNER JOIN " & _
                        "tblBookGroupMembers ON tblNameAddress.ID = tblBookGroupMembers.PersonID " & _
                 "WHERE Active = TRUE " & _
                 "AND tblBookGroupMembers.GroupNo = " & GroupNo
    Else
        sSQL = "SELECT ID " & _
                 "FROM tblNameAddress " & _
                 "WHERE Active = TRUE "
    End If
             
    Set rs = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    Set rstPrintTable = CMSDB.OpenRecordset("tblPrintAddresses", dbOpenDynaset)
    
    If rs.BOF Then
        GenerateAddressListForGroup = False
        Exit Function
    Else
        GenerateAddressListForGroup = True
        With rs
        
        Do Until .EOF 'add new print rec for each row found
        
            rstPrintTable.AddNew
                                    
            rstPrintTable!PersonName = CongregationMember.LastFirstNameWithMiddleInitial(!ID)
            
            sAddr = CongregationMember.GetPersonsAddress(!ID)
            sAddr = Replace(sAddr, vbCrLf, " ")
            sAddr = Replace(sAddr, vbLf, " ")
            sAddr = Replace(sAddr, vbCr, " ")
            rstPrintTable!Address = sAddr
            
            rstPrintTable!HomePhone = CongregationMember.HomePhone(!ID)
            
            sMbl1 = CongregationMember.MobilePhone(!ID)
            sMbl2 = CongregationMember.MobilePhone2(!ID)
            If sMbl1 <> "" Then
                If sMbl2 <> "" Then
                    sMbl = sMbl1 & " / " & sMbl2
                Else
                    sMbl = sMbl1
                End If
            Else
                If sMbl2 <> "" Then
                    sMbl = sMbl2
                Else
                    sMbl = ""
                End If
            End If
            rstPrintTable!MobilePhone = sMbl
            rstPrintTable!MobilePhone2 = ""
                                                   
            rstPrintTable.Update
            
            .MoveNext
        Loop
        
        End With
        
        If PrintUsingWord(False) = cmsUseWord Then
            GenerateAddressPrintInWord GetGroupName(GroupNo)
        End If
    
    End If

    Exit Function
ErrorTrap:
    Call EndProgram
End Function

Public Function PrintAddressList(GroupNos As String) As Boolean

On Error GoTo ErrorTrap

Dim sSQL As String, rstPrintTable As Recordset, rs As Recordset
Dim lGrp  As Long
Dim RotaTopMargin As Single, RotaBottomMargin As Single, RotaLeftMargin As Single
Dim RotaRightMargin As Single


    'first build the print table
    
    DelAllRows "tblPrintAddresses"
    
    '
    'Driving Recset to get all data
    '
    sSQL = "SELECT ID " & _
             "FROM tblNameAddress LEFT JOIN " & _
                    "tblBookGroupMembers ON tblNameAddress.ID = tblBookGroupMembers.PersonID " & _
             "WHERE Active = TRUE " & _
             IIf(GroupNos <> "", "AND tblBookGroupMembers.GroupNo IN (" & GroupNos & ") ", " ")
             
    Set rs = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    Set rstPrintTable = CMSDB.OpenRecordset("tblPrintAddresses", dbOpenDynaset)
    
    If rs.BOF Then
        PrintAddressList = False
        Exit Function
    Else
        PrintAddressList = True
        With rs
        
        Do Until .EOF 'add new print rec for each row found
        
            rstPrintTable.AddNew
                                    
            rstPrintTable!PersonName = CongregationMember.LastFirstNameWithMiddleInitial(!ID)
            rstPrintTable!Address = CongregationMember.GetPersonsAddress(!ID)
            rstPrintTable!HomePhone = CongregationMember.HomePhone(!ID)
            
            If CongregationMember.MobilePhone2(!ID) <> "" Then
                rstPrintTable!MobilePhone = CongregationMember.MobilePhone(!ID) & " / " & _
                                            CongregationMember.MobilePhone2(!ID)
            Else
                rstPrintTable!MobilePhone = CongregationMember.MobilePhone(!ID)
            End If
            
            rstPrintTable!MobilePhone2 = CongregationMember.MobilePhone2(!ID)
            
            lGrp = CongregationMember.BookGroup(!ID)
            
            If GroupNos <> "" Then
                rstPrintTable!GroupName = IIf(lGrp > 0, GetGroupName(lGrp), "-- No Group Set --")
            Else
                rstPrintTable!GroupName = GlobalParms.GetValue("DefaultCong", "AlphaVal") & " Congregation"
            End If
                                                   
            rstPrintTable.Update
            
            .MoveNext
        Loop
        
        End With
        
        '
        'Arrange page margins before we close the db connection
        '
        RotaTopMargin = 566.929 * (GlobalParms.GetValue("A4TopMargin", "NumFloat"))
        RotaBottomMargin = 566.929 * (GlobalParms.GetValue("A4BottomMargin", "NumFloat"))
        RotaLeftMargin = 566.929 * (GlobalParms.GetValue("A4LeftMargin", "NumFloat"))
        RotaRightMargin = 566.929 * (GlobalParms.GetValue("A4RightMargin", "NumFloat"))
            
        DestroyGlobalObjects
        CMSDB.Close
        
        '
        'Set up the datareport
        '
        '
        'GENERAL ADO WARNING...
        ' If we refer to DataReport prior to 'Showing' it, thus opening new ADODB connection
        ' while DAO connection still open, we get funny results.. eg missing fields on report.
        '
        PrintAllCongMembers.TopMargin = RotaTopMargin '<----- At this point, 'PrintCleaningRota.Initialize' runs.
        PrintAllCongMembers.BottomMargin = RotaBottomMargin
        PrintAllCongMembers.LeftMargin = RotaLeftMargin
        PrintAllCongMembers.RightMargin = RotaRightMargin
        
        Screen.MousePointer = vbNormal
        
        PrintAllCongMembers.Show vbModal
        
        '
        'Global Objects are destroyed when report is generated
        ' due to DB Disconnect. So, instantiate them once more....
        '
        SwitchOnDAO
    
    End If

    Exit Function
ErrorTrap:
    Call EndProgram
End Function

Public Function GenerateAddressPrintInWord(sGroup As String) As Boolean

Dim reporter As MSWordReportingTool2.RptTool, str As String

On Error GoTo ErrorTrap

    str = GlobalParms.GetValue("DefaultCong", "AlphaVal") & " Congregation"
    
    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.
    
    .ReportSQL = "SELECT PersonName, " & _
                 "       Address, " & _
                 "       HomePhone, " & _
                 "       MobilePhone " & _
                 "FROM   tblPrintAddresses " & _
                 "   ORDER BY 1 "
                 
    .SaveDoc = True
    
    If sGroup <> "" Then
        .ReportTitle = sGroup & " Field Service Group" & vbCrLf & GetMonthName(Month(Now)) & _
                                            " " & Day(Now) & GetLettersForOrdinalNumber(Day(Now)) & _
                                            " " & year(Now)
                                            
        .DocPath = gsDocsDirectory & "\" & sGroup & " Field Service Addresses " & _
                                    Replace(Replace(Now, ":", "-"), "/", "-")
    Else
        .ReportTitle = str & vbCrLf & GetMonthName(Month(Now)) & _
                        " " & Day(Now) & GetLettersForOrdinalNumber(Day(Now)) & _
                        " " & year(Now)
                        
        .DocPath = gsDocsDirectory & "\" & str & " Addresses " & _
                                    Replace(Replace(Now, ":", "-"), "/", "-")
    End If
    
        
    .AdditionalReportHeadingItalic = True
    .AdditionalReportHeadingBold = True
    .AdditionalReportHeadingFontSize = 12
    .AdditionalReportHeading = ""
    
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 15
    .RightMargin = 15
    .ReportFooterFontName = "Arial"
    .ReportFooterFontSize = 10
    .ReportTitleFontName = "Times New Roman"
    .ReportTitleFontSize = 16
    .ApplyTableFormatting = True
    .ClientName = AppName
    .ShowPageNumber = True
    .GroupingColumn = 0
    .PageFormat = cmsPortrait
    .HideWordWhileBuilding = True
    .HideWordWhenDone = False
            
    .ShowProgress = True
    
    .AddTableColumnAttribute "Name", 40, , , "Times New Roman", , 9, 9, True, True, , , True
    .AddTableColumnAttribute "Address", 80, , , "Times New Roman", , 9, 9, True, True
    .AddTableColumnAttribute "Home Phone", 30, , , "Times New Roman", , 9, 9, True, True
    .AddTableColumnAttribute "Mobile Phone", 30, , , "Times New Roman", , 9, 9, True, True
            
    .GenerateReport
       
    End With
    
    SwitchOnDAO
    
    Screen.MousePointer = vbNormal

    Exit Function
ErrorTrap:
    EndProgram
    
End Function
Public Sub RemoveDatabaseOrphans()
On Error GoTo ErrorTrap
Dim sSQL As String, rs As Recordset

    sSQL = " NOT IN (SELECT ID FROM tblNameAddress) "

    DeleteSomeRows "tblIDWeightings", "ID" & sSQL
    DeleteSomeRows "tblMarriage", "ID" & sSQL
    DeleteSomeRows "tblMarriage", "Spouse" & sSQL
    DeleteSomeRows "tblChildren", "Parent" & sSQL
    DeleteSomeRows "tblChildren", "Child" & sSQL
    DeleteSomeRows "tblTaskAndPerson", "Person" & sSQL
    DeleteSomeRows "tblTaskPersonSuspendDates", "Person" & sSQL
    DeleteSomeRows "tblTMSCounselPoints", "StudentID" & sSQL
    DeleteSomeRows "tblBaptismDates", "PersonID" & sSQL
    DeleteSomeRows "tblBookGroupMembers", "PersonID" & sSQL
    DeleteSomeRows "tblEldersAndServants", "PersonID" & sSQL
    DeleteSomeRows "tblIrregularPubs", "PersonID" & sSQL
    DeleteSomeRows "tblMinReports", "PersonID" & sSQL
    DeleteSomeRows "tblPublisherDates", "PersonID" & sSQL
    DeleteSomeRows "tblPubCardTypeForPerson", "PersonID" & sSQL
    DeleteSomeRows "tblRegPioDates", "PersonID" & sSQL
    DeleteSomeRows "tblAuxPioDates", "PersonID" & sSQL
    DeleteSomeRows "tblSpecPioDates", "PersonID" & sSQL
    DeleteSomeRows "tblPubRecCardRowPrinted", "PersonID" & sSQL
    DeleteSomeRows "tblVisitingSpeakers", "PersonID" & sSQL
    DeleteSomeRows "tblIndividualSPAMWeightings", "PersonID" & sSQL
    DeleteSomeRows "tblPioHourCredit", "PersonID" & sSQL
    DeleteSomeRows "tblSpeakersTalks", "PersonID" & sSQL
    DeleteSomeRows "tblIndividualPioTarget", "PersonID" & sSQL
    DeleteSomeRows "tblGeneralNotesForPerson", "PersonID" & sSQL
    
    'a trickier delete arrangement is required for missing reports....
    
    Set rs = CMSDB.OpenRecordset("SELECT DISTINCT PersonID " & _
                                 "FROM tblMissingReports " & _
                                 "WHERE PersonID " & sSQL, dbOpenForwardOnly)

    With rs
    
    Do Until .EOF Or .BOF
        DeleteMissingReportsForPerson CLng(!PersonID)
        .MoveNext
    Loop
      
    End With
    
    '....and Gift Aiders
    
    Set rs = CMSDB.OpenRecordset("SELECT DISTINCT PersonID " & _
                                 "FROM tblGiftAidPayers " & _
                                 "WHERE PersonID " & sSQL, dbOpenForwardOnly)

    With rs
    
    Do Until .EOF Or .BOF
        DeleteGiftAid CLng(!PersonID)
        .MoveNext
    Loop
      
    End With
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub


Public Function PioneerTargetInfo(PersonID As Long, _
                                   MinType As MinistryType, _
                                   CurrentDate As Date) As String
On Error GoTo ErrorTrap
Dim lServiceYear As Long, DateString1 As String, DateString2 As String
Dim lNormalYear As Long, lNumMonths As Long, lTargetAnnualHours As Long
Dim lTargetMonthlyHours As Long, fActualMonthlyHours As Single
Dim lActualHoursToDate As Long, lTargetHoursToDate As Long
Dim rs As Recordset, str As String, fRequiredAverage As Single
Dim PioStartDate As Date, lTotalMonths As Long
Dim StartServiceYear As String, lHourDiff As Long

    If MinType <> IsRegPio And MinType <> IsSpecPio Then
        Exit Function
    End If
    
    lServiceYear = year(ConvertNormalDateToServiceDate(CDate(Format$(CurrentDate, "dd/mm/yyyy"))))
    
    lNormalYear = ConvertServiceYearToNormalYear(CDate("01/09" & "/" & lServiceYear))
    
    'Date from which we're interested - ie start of service year
    DateString1 = "01/09/" & lNormalYear
    StartServiceYear = DateString1
    
    'Date up to which we're interested
    DateString2 = "01/" & Month(CurrentDate) & "/" & year(CurrentDate)
        
    lTargetAnnualHours = CongregationMember.GetIndividualPioTargetHours(PersonID, lServiceYear)
        
    If MinType = IsRegPio Then
        If lTargetAnnualHours = 0 Then
            lTargetAnnualHours = GlobalParms.GetValue("AnnualRegPioHours", "NumVal")
        End If
        Set rs = CMSDB.OpenRecordset("SELECT StartDate, EndDate " & _
                                     "FROM tblRegPioDates " & _
                                     "WHERE PersonID = " & PersonID & _
                                     " AND StartDate <= #" & _
                                        Format(DateString2, "mm/dd/yyyy") & _
                                        "# AND EndDate >= #" & Format(DateString1, "mm/dd/yyyy") & _
                                        "# ORDER BY StartDate DESC", dbOpenForwardOnly)
    Else
        If lTargetAnnualHours = 0 Then
            lTargetAnnualHours = GlobalParms.GetValue("AnnualSpecPioHours", "NumVal")
        End If
        Set rs = CMSDB.OpenRecordset("SELECT StartDate, EndDate " & _
                                     "FROM tblSpecPioDates " & _
                                     "WHERE PersonID = " & PersonID & _
                                     " AND StartDate <= #" & _
                                        Format(DateString2, "mm/dd/yyyy") & _
                                        "# AND EndDate >= #" & Format(DateString1, "mm/dd/yyyy") & _
                                        "# ORDER BY StartDate DESC", dbOpenForwardOnly)
    End If
        
    
    With rs
    If Not .BOF Then
        PioStartDate = !StartDate
    Else
        PioneerTargetInfo = " - Unable to calculate details"
        Exit Function
    End If
    
    lNumMonths = DateDiff("m", StartServiceYear, PioStartDate)
        
    lTargetMonthlyHours = lTargetAnnualHours / 12
    
    If lNumMonths > 0 Then 'pio period starts after the start of this service year
        'look for any previous pio period...
        .MoveNext
        Do Until .EOF
        
            If DateDiff("m", !EndDate, PioStartDate) = 1 And _
               DateDiff("m", StartServiceYear, !EndDate) >= 0 Then
                'two pio periods contiguous, and prev period ends within this service year
                If DateDiff("m", !StartDate, StartServiceYear) > 0 Then 'prev period start prior to this service year
                    DateString1 = StartServiceYear
                    Exit Do
                Else
                    DateString1 = CStr(!StartDate)
                    Exit Do
                End If
            Else
                DateString1 = StartServiceYear
                Exit Do
            End If
            
            .MoveNext
            
        Loop
        
        lNumMonths = DateDiff("m", StartServiceYear, DateString1)
        lTotalMonths = 12 - lNumMonths
        
        lTargetAnnualHours = lTotalMonths * lTargetMonthlyHours
                
    Else 'pio period start on or before start of service year
        lTotalMonths = 12
    End If
    
    End With
    rs.Close
    Set rs = Nothing
        
    Set rs = CMSDB.OpenRecordset("SELECT SUM(NoHours) AS TotHrs " & _
                                 "FROM tblMinReports " & _
                                 "WHERE PersonID = " & PersonID & _
                                 " AND ActualMinPeriod BETWEEN #" & _
                                    Format(DateString1, "mm/dd/yyyy") & _
                                    "# AND #" & Format(DateString2, "mm/dd/yyyy") & _
                                    "#", dbOpenForwardOnly)
    
    With rs
    If IsNull(!TotHrs) Then
        lActualHoursToDate = 0
    Else
        lActualHoursToDate = !TotHrs
    End If
    End With
    
    Set rs = CMSDB.OpenRecordset("SELECT SUM(NoHours) AS TotHrs " & _
                                 "FROM tblPioHourCredit " & _
                                 "WHERE PersonID = " & PersonID & _
                                 " AND MinDate BETWEEN #" & _
                                    Format(DateString1, "mm/dd/yyyy") & _
                                    "# AND #" & Format(DateString2, "mm/dd/yyyy") & _
                                    "#", dbOpenForwardOnly)
    
    With rs
    If Not IsNull(!TotHrs) Then
        lActualHoursToDate = lActualHoursToDate + !TotHrs
    End If
    End With
    
    rs.Close
    Set rs = Nothing
    
    
    lNumMonths = DateDiff("m", DateString1, DateString2) + 1
    
    lTargetHoursToDate = lTargetMonthlyHours * lNumMonths

    fActualMonthlyHours = lActualHoursToDate / lNumMonths
    

    str = vbCrLf & "As at end of " & GetMonthName(CLng(Month(DateString2))) & " " & year(DateString2) & ": "
    
    If lNumMonths < lTotalMonths Then
        If lActualHoursToDate < lTargetHoursToDate Then
            str = str & lTargetHoursToDate - lActualHoursToDate & " hours behind target. "
            fRequiredAverage = ((lTargetAnnualHours - lActualHoursToDate) / (lTotalMonths - lNumMonths))
            If lNumMonths < 11 Then
                str = str & Round(fRequiredAverage, 2) & " hours required for each month."
            Else
                str = str & Round(fRequiredAverage, 2) & " hours required next month."
            End If
        Else
            If lActualHoursToDate >= lTargetAnnualHours Then
                str = str & "Annual requirement reached"
            Else
                If lActualHoursToDate > lTargetHoursToDate Then
                    str = str & lActualHoursToDate - lTargetHoursToDate & " hours ahead of target. "
                Else
                    str = str & "On target"
                End If
            End If
        End If
    Else
        lHourDiff = lActualHoursToDate - lTargetAnnualHours
        If lHourDiff >= 0 Then
            str = str & "Annual requirement reached (" & lActualHoursToDate & "/" & lTargetAnnualHours & " hours" & _
                     IIf(lHourDiff > 0, ", " & Abs(lHourDiff) & " ahead of target)", ")")
        Else
            str = str & "Failed to reach annual requirement (" & lActualHoursToDate & "/" & lTargetAnnualHours & " hours" & _
                      ", " & Abs(lHourDiff) & " behind target)"
        End If
    End If
    
    PioneerTargetInfo = str

    Exit Function
ErrorTrap:
    PioneerTargetInfo = " - Unable to calculate details"

End Function

Public Function TransactionsForAccountID(AccountID As Long) As Long
Dim rs As Recordset
On Error GoTo ErrorTrap

    Set rs = CMSDB.OpenRecordset("SELECT COUNT(*) AS TheNo FROM tblTransactionDates " & _
                                "WHERE AccountID = " & AccountID & _
                                " OR TfrAccountID = " & AccountID, dbOpenForwardOnly)
                                
    If Not rs.BOF Then
        If Not IsNull(rs!TheNo) Then
            TransactionsForAccountID = rs!TheNo
        Else
            TransactionsForAccountID = 0
        End If
    Else
        TransactionsForAccountID = 0
    End If
        
    rs.Close
    Set rs = Nothing

    Exit Function
ErrorTrap:
    Call EndProgram
End Function

Public Function GetDatabaseTableScalar(SQL As String, Optional ValueIfNull = 0) As Variant
Dim rs As Recordset
On Error GoTo ErrorTrap

    Set rs = CMSDB.OpenRecordset(SQL, dbOpenForwardOnly)
                                
    If Not rs.BOF Then
        GetDatabaseTableScalar = HandleNull(rs.Fields(0), ValueIfNull)
    Else
        GetDatabaseTableScalar = ValueIfNull
    End If
        
    rs.Close
    Set rs = Nothing

    Exit Function
ErrorTrap:
    Call EndProgram
End Function


