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
      TabIndex        =   5
      Top             =   2160
      Width           =   930
   End
   Begin VB.TextBox txtSnoozeDays 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   6195
      MaxLength       =   3
      TabIndex        =   6
      Top             =   2280
      Width           =   555
   End
   Begin VB.ListBox lstAlarms 
      Height          =   1860
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   210
      Width           =   8130
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   525
      Left            =   7305
      TabIndex        =   7
      Top             =   2160
      Width           =   930
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View..."
      Height          =   525
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
      TabIndex        =   9
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

Private Sub cmdAcknowledge_Click()
Dim i As Integer, dteTemp As Date, lStoreDay As Long
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
                    lStoreDay = Day(dteTemp)
                    Do Until dteTemp > .EventEndDate And dteTemp > Now
                        Select Case .PeriodID
                        Case 1 'daily
                            dteTemp = DateAdd("d", .PeriodCycle, dteTemp)
                        Case 2 'weekly
                            dteTemp = DateAdd("ww", .PeriodCycle, dteTemp)
                        Case 3 'monthly by date
                            dteTemp = DateAdd("m", .PeriodCycle, dteTemp)
                        Case 4 'monthly by day
                            dteTemp = DateAdd("m", .PeriodCycle, dteTemp)
                            
                        End Select
                    Loop
                    
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

Private Sub cmdPrint_Click()

On Error GoTo ErrorTrap

Dim ErrorCode As Integer, i As Long, reporter As MSWordReportingTool1.rpttool
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

    Screen.MousePointer = vbHourglass
    
    If PrintUsingWord Then
        
        SwitchOffDAO
    
        Set reporter = New rpttool
        
        With reporter
        
        .DB_PathAndName = CompletePathToTheMDBFileAndExt
        
        .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.
    
        .SaveDoc = True
        .DocPath = gsDocsDirectory & "\" & "Reminders on " & _
                                    Replace(Replace(Now, ":", "-"), "/", "-")
    
        
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
        .AdditionalReportHeading = Format(Now, "dd/mm/yyyy")
        .GroupingColumn = 0
        .HideWordWhileBuilding = True
        
        .AddTableColumnAttribute "Reminders", 120, , , , , 14, 13, True, True, , , True
        
        .PageFormat = cmsPortrait
        
        .GenerateReport
    
        End With
        
        Set reporter = Nothing
        SwitchOnDAO
        
        Screen.MousePointer = vbNormal
        
    Else
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
        
        msinfo = Format(Now, "dd/mm/yyyy")
        PrintReminders.Show vbModeless, Me
        
        '
        'Global Objects are destroyed when report is generated
        ' due to DB Disconnect. So, instantiate them once more....
        '
        SwitchOnDAO
        
        Screen.MousePointer = vbNormal
                
    End If
    
    
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
                        .UpdateAlarmDateOnCurrentEvent CDate(DateAdd("d", CDbl(txtSnoozeDays), CDate(Format(Now, "dd/mm/yyyy"))))
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
    
    'get all events for today. Loop through them and, if alarm falls today,
    ' add to listbox on frmAlarms
    
    lstAlarms.Clear
    
    With GlobalCalendar
    .GetDaysEvents CDate(Format(Now, "dd/mm/yyyy")), True, True
    
    lstAlarms.Clear
    
    If Not .NoEventsFound Then
        
        Do Until .NoMoreEvents
            lstAlarms.AddItem .EventStartDate & ": " & .EventName & " - " & .NoteMemo
            lstAlarms.ItemData(lstAlarms.NewIndex) = .SeqNum
            .GetNextEvent
        Loop
    
        AnyAlarms = True
    Else
        AnyAlarms = False
    End If
    
    End With
    
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

Private Sub txtSnoozeDays_KeyPress(KeyAscii As Integer)
    KeyPressValid KeyAscii, cmsUnsignedIntegers, True
End Sub

Public Property Get Info() As String
    Info = msinfo
End Property

