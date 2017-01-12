VERSION 5.00
Begin VB.Form frmAlarms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " C.M.S. Calendar Alarms for Today"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   Icon            =   "frmAlarms.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit..."
      Height          =   255
      Left            =   4209
      TabIndex        =   5
      Top             =   2430
      Width           =   930
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print..."
      Height          =   525
      Left            =   2157
      TabIndex        =   2
      Top             =   2160
      Width           =   930
   End
   Begin VB.CommandButton cmdSnooze 
      Caption         =   "Snooze"
      Height          =   525
      Left            =   5235
      TabIndex        =   6
      Top             =   2160
      Width           =   930
   End
   Begin VB.TextBox txtSnoozeDays 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6195
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2280
      Width           =   555
   End
   Begin VB.ListBox lstAlarms 
      Height          =   1860
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   210
      Width           =   8130
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   525
      Left            =   7305
      TabIndex        =   8
      Top             =   2160
      Width           =   930
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View..."
      Height          =   255
      Left            =   4209
      TabIndex        =   4
      Top             =   2160
      Width           =   930
   End
   Begin VB.CommandButton cmdDeSelectAll 
      Caption         =   "De-Select All"
      Height          =   525
      Left            =   1131
      TabIndex        =   1
      Top             =   2160
      Width           =   930
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All"
      Height          =   525
      Left            =   105
      TabIndex        =   0
      Top             =   2160
      Width           =   930
   End
   Begin VB.CommandButton cmdAcknowledge 
      Caption         =   "Dismiss"
      Height          =   525
      Left            =   3183
      TabIndex        =   3
      Top             =   2160
      Width           =   930
   End
   Begin VB.Label Label2 
      Caption         =   "days"
      Height          =   240
      Left            =   6810
      TabIndex        =   10
      Top             =   2310
      Width           =   390
   End
End
Attribute VB_Name = "frmAlarms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim msinfo As String
Dim mdteCurrentDate As Date

Private Sub cmdAcknowledge_Click()
Dim i As Integer, dteTemp As Date, lStoreDay As Long
Dim bFound As Boolean
On Error GoTo ErrorTrap

    If lstAlarms.SelCount > 0 Then
        With GlobalCalendar
        For i = 0 To lstAlarms.ListCount - 1
            If lstAlarms.Selected(i) Then
                .GetAnEvent lstAlarms.ItemData(i)
                If .PeriodID = 0 Then 'not recurring
                    .SetAlarmAcknowledged lstAlarms.ItemData(i), True
                Else
                    dteTemp = .EventStartDate
                    lStoreDay = Weekday(dteTemp)
                    bFound = False 'init
                    Do Until dteTemp > .EventEndDate And dteTemp > mdteCurrentDate
                        Select Case .PeriodID
                        Case 1 'daily
                            dteTemp = DateAdd("d", .PeriodCycle, dteTemp)
                        Case 2 'weekly
                            dteTemp = DateAdd("ww", .PeriodCycle, dteTemp)
                        Case 3 'monthly by date
                            dteTemp = DateAdd("m", .PeriodCycle, dteTemp)
                        Case 4 'monthly by day
                            dteTemp = DateAdd("m", .PeriodCycle, dteTemp)
                            Select Case .WhichWeekOfMonth
                            Case 5
                                dteTemp = DateOfLastDayOfMonth(lStoreDay, year(dteTemp), Month(dteTemp))
                            Case Else
                                dteTemp = DateOfNthDay(lStoreDay, year(dteTemp), Month(dteTemp), .WhichWeekOfMonth)
                            End Select
                        End Select
                        If dteTemp > mdteCurrentDate And dteTemp <= .EventEndDate Then
                            bFound = True
                            Exit Do
                        End If
                    Loop
                    
                    If bFound Then
                        .UpdateAlarmDateOnCurrentEvent dteTemp
                        .UpdateStartDateOnCurrentEvent dteTemp
                    Else
                        .SetAlarmAcknowledged lstAlarms.ItemData(i), True
                    End If
                    
                End If
            End If
        Next i
        End With
        AnyAlarms
        
        'no point having form open with no events on it!
        If lstAlarms.ListCount = 0 Then
            Unload Me
        End If
        
    Else
        MsgBox "You must select checkbox for items to dismiss.", vbExclamation + vbOKOnly, AppName
    End If

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDeSelectAll_Click()

On Error GoTo ErrorTrap

    SelectAllInListBox lstAlarms, cmsSelectNone

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub cmdEdit_Click()
Dim lSeq As Long
On Error GoTo ErrorTrap
    
    If lstAlarms.ListIndex > -1 Then
    
        lSeq = lstAlarms.ItemData(lstAlarms.ListIndex)
        
        GlobalCalendar.GetAnEvent lSeq
        
        If Not GlobalCalendar.NoEventsFound Then
        
            With frmAddNewEvent
            '
            'Fill frmAddNewEvent before it's displayed
            '
            .DoNotTrigger = True
            
            .UpdateMode = cmsEdit
            .Caption = "C.M.S. Change Calendar Event"
            .txtCong = GetCongregationName(GlobalDefaultCong)
            .txtStartDate = Format(GlobalCalendar.EventStartDate, "DD/MM/YYYY")
            .txtEndDate = Format(GlobalCalendar.EventEndDate, "DD/MM/YYYY")
            .txtEndTime = Format(GlobalCalendar.EventEndTime, "HH:MM")
            .txtNote = GlobalCalendar.NoteMemo
            .txtStartTime = Format(GlobalCalendar.EventStartTime, "HH:MM")
            
            Select Case GlobalCalendar.AlarmDateTime
            Case 0, 1, 2 'if alarmdatetime is 2, there is no alarm
                .txtAlarmDate = ""
                .txtAlarmTime = ""
                .chkAlarm.value = vbUnchecked
            Case 3 'if alarmdatetime is 3, there is an alarm for recurring event
                .txtAlarmDate = ""
                .txtAlarmTime = ""
                .chkAlarm.value = vbChecked
            Case Else
                .txtAlarmDate = Format(GlobalCalendar.AlarmDateTime, "DD/MM/YYYY")
                .txtAlarmTime = Format(GlobalCalendar.AlarmDateTime, "HH:MM")
                .chkAlarm.value = vbChecked
            End Select
            
            .chkPrivate = IIf(GlobalCalendar.PrivateEntry, vbChecked, vbUnchecked)
            
            HandleListBox.PopulateListBox .cmbEvent, "SELECT EventID, EventName FROM tblEventLookup " & _
                                                        "WHERE ShowInCalendar = TRUE ", _
                            CMSDB, 0, "", False, 1
                            
            .EventTypeID = GlobalCalendar.EventID
            .EventID = GlobalCalendar.SeqNum
            
            '
            'Set for single/recurring event according to content of LinkSeqNum field...
            '
            If GlobalCalendar.PeriodID = 0 Then
                .optSingle = True
            Else
                .optRecurring = True
            End If
        
            .DoNotTrigger = False
            
            '
            'Show frmAddNewEvent
            '
            .Show vbModal, Me
            
            AnyAlarms
            
            End With
    
        Else
            MsgBox "Event not found - data error.", vbOKOnly + vbCritical, AppName
            Exit Sub
        End If
        
    Else
        MsgBox "First select an event to edit", vbOKOnly + vbExclamation, AppName
        Exit Sub
    End If
    
    
    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub cmdPrint_Click()

On Error GoTo ErrorTrap

Dim ErrorCode As Integer, i As Long, reporter As MSWordReportingTool2.RptTool
Dim rs As Recordset
Dim RotaTopMargin As Single, RotaBottomMargin As Single, RotaLeftMargin As Single
Dim RotaRightMargin As Single


    If lstAlarms.SelCount = 0 Then
        ShowMessage "No entry selected", 1000, Me
        GoTo GetOut
    End If

    DeleteTable "tblTempAlarms"
    CreateTable ErrorCode, "tblTempAlarms", "TheText", "MEMO", , , True
    
    For i = 0 To lstAlarms.ListCount - 1
        lstAlarms.ListIndex = i
        If lstAlarms.Selected(i) Then
            CMSDB.Execute "INSERT INTO tblTempAlarms (TheText) VALUES ('" & DoubleUpSingleQuotes(lstAlarms.text) & "')"
        End If
    Next i
    
    lstAlarms.ListIndex = -1
    
    Set rs = CMSDB.OpenRecordset("SELECT * FROM tblTempAlarms", dbOpenForwardOnly)
    
    If rs.BOF Or rs.EOF Then GoTo GetOut
    
    Select Case PrintUsingWord
    Case cmsUseWord
        Screen.MousePointer = vbHourglass
        
        SwitchOffDAO
    
        Set reporter = New RptTool
        
        With reporter
        
        .DB_PathAndName = CompletePathToTheMDBFileAndExt
        
        .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.
    
        .SaveDoc = True
        .DocPath = gsDocsDirectory & "\" & "Reminders on " & _
                                    Replace(Replace(CStr(mdteCurrentDate), ":", "-"), "/", "-")
    
        
        .ReportSQL = "SELECT TheText " & _
                     "FROM tblTempAlarms "
    
        .ReportTitle = "Reminders"
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
        .AdditionalReportHeading = Format(mdteCurrentDate, "dd/mm/yyyy")
        .GroupingColumn = 0
        .HideWordWhileBuilding = True
        
        .AddTableColumnAttribute "Reminders", 120, , , , , 14, 13, True, True, , , True
        
        .PageFormat = cmsPortrait
        
        .GenerateReport
    
        End With
        
        Set reporter = Nothing
        SwitchOnDAO
        
        Screen.MousePointer = vbNormal
        
    Case cmsUseMSDatareport
    
        Screen.MousePointer = vbHourglass
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
        PrintReminders.TopMargin = RotaTopMargin '<----- At this point, 'PrintCleaningRota.Initialize' runs.
        PrintReminders.BottomMargin = RotaBottomMargin
        PrintReminders.LeftMargin = RotaLeftMargin
        PrintReminders.RightMargin = RotaRightMargin
        
        Screen.MousePointer = vbNormal
        
        msinfo = Format(mdteCurrentDate, "dd/mm/yyyy")
        PrintReminders.Show vbModeless, Me
        
        '
        'Global Objects are destroyed when report is generated
        ' due to DB Disconnect. So, instantiate them once more....
        '
        SwitchOnDAO
        
        Screen.MousePointer = vbNormal
                
    End Select
    
    
GetOut:
    On Error Resume Next
    rs.Close
    Set rs = Nothing

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdSelectAll_Click()

On Error GoTo ErrorTrap

    SelectAllInListBox lstAlarms, cmsSelectAll

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdSnooze_Click()
On Error GoTo ErrorTrap
Dim str As String, i As Long

    If lstAlarms.SelCount = 0 Then
        MsgBox "Please select an alarm entry", vbOKOnly + vbExclamation, AppName
        Exit Sub
    End If

    If IsNumber(txtSnoozeDays, False, False, False) Then
    
        If CLng(txtSnoozeDays) > 0 Then
        
            With GlobalCalendar
            
            For i = 0 To lstAlarms.ListCount - 1
            
                If lstAlarms.Selected(i) Then
                
                    .GetAnEvent lstAlarms.ItemData(i)
                    
                    If .AlarmDateTime <> "" Then
                        .UpdateAlarmDateOnCurrentEvent CDate(DateAdd("d", CDbl(txtSnoozeDays), CDate(Format(mdteCurrentDate, "dd/mm/yyyy"))))
                    End If
                    
                End If
                
            Next i
            
            End With
            
        Else
            TextFieldGotFocus txtSnoozeDays, True
            MsgBox "Invalid snooze value", vbOKOnly + vbExclamation, AppName
        End If
        
    Else
    
        TextFieldGotFocus txtSnoozeDays, True
        MsgBox "Invalid snooze value", vbOKOnly + vbExclamation, AppName
        
    End If

    AnyAlarms
    If lstAlarms.ListCount = 0 Then
        Unload Me
    End If

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdView_Click()
Dim lSeq As Long
On Error GoTo ErrorTrap
    
    If lstAlarms.ListIndex > -1 Then
    
        lSeq = lstAlarms.ItemData(lstAlarms.ListIndex)
        
        With GlobalCalendar
        
        .GetAnEvent lSeq
        
        If Not .NoEventsFound Then
            frmViewEvent.Show vbModal, Me
        Else
            MsgBox "Event not found - data error.", vbOKOnly + vbCritical, AppName
            Exit Sub
        End If
        
        End With
    Else
        MsgBox "First select an event to view", vbOKOnly + vbExclamation, AppName
        Exit Sub
    End If
    
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Function AnyAlarms() As Boolean
On Error GoTo ErrorTrap
Dim bAlarmFound As Boolean
Dim mlStore As Long, mlStoreIX As Long
    
    'get all events for today. Loop through them and, if alarm falls today,
    ' add to listbox on frmAlarms
    
    If lstAlarms.ListIndex > -1 Then
        mlStore = lstAlarms.ItemData(lstAlarms.ListIndex)
    Else
        mlStore = -1
    End If
    
    mlStoreIX = -1
        
    lstAlarms.Clear
    
    With GlobalCalendar
    .GetDaysEvents CDate(Format(mdteCurrentDate, "dd/mm/yyyy")), True, True
    
    lstAlarms.Clear
    
    If Not .NoEventsFound Then
        
        Do Until .NoMoreEvents
            If ShowAlert(.AutoAlarmID) Then
                lstAlarms.AddItem .EventStartDate & ": " & .EventName & " - " & .NoteMemo
                lstAlarms.ItemData(lstAlarms.NewIndex) = .SeqNum
                If mlStore = .SeqNum Then
                    mlStoreIX = lstAlarms.NewIndex
                End If
            End If
            .GetNextEvent
        Loop
    
        AnyAlarms = (lstAlarms.ListCount > 0)
    Else
        AnyAlarms = False
    End If
    
    End With
    
    If mlStoreIX > -1 Then
        lstAlarms.ListIndex = mlStoreIX
    Else
        lstAlarms.ListIndex = -1
    End If
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function


Private Sub Form_Load()

On Error GoTo ErrorTrap

'    cmdPrint.Enabled = PrintUsingWord(False)

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Function ShowAlert(AutoAlarmID As Long) As Boolean

On Error GoTo ErrorTrap

Dim rs As Recordset, str As String

    str = "SELECT AlarmDays, AlarmMemo, TriggerSQL, AccessLevels " & _
          "FROM tblEventAutoAlarms " & _
          "WHERE AutoAlarmID = " & AutoAlarmID
          
    Set rs = CMSDB.OpenRecordset(str, dbOpenForwardOnly)

    If rs.BOF Then
        ShowAlert = True
        GoTo GetOut
    End If

    If rs!AccessLevels <> "" Then
        If Not UserHasSecurityLevel(gCurrentUserCode, rs!AccessLevels) Then
            ShowAlert = False
            GoTo GetOut
        End If
    End If
    
    If rs!TriggerSQL <> "" Then
        Set rs = GetGeneralRecordset(rs!TriggerSQL)
        If rs.BOF Then
            ShowAlert = False
            GoTo GetOut
        End If
    End If
    
    ShowAlert = True

GetOut:

    rs.Close
    Set rs = Nothing

    Exit Function
ErrorTrap:
    EndProgram

End Function


Private Sub txtSnoozeDays_KeyPress(KeyAscii As Integer)
    KeyPressValid KeyAscii, cmsunSignedIntegers, True
End Sub

Public Property Get Info() As String
    Info = msinfo
End Property

Public Property Let CurrentDate(ByVal vNewValue As Date)
    mdteCurrentDate = vNewValue
End Property
