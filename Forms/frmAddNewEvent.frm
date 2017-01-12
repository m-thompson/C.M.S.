VERSION 5.00
Begin VB.Form frmAddNewEvent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.M.S. Update Calendar"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "frmAddNewEvent.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdShowCalendar1 
      DownPicture     =   "frmAddNewEvent.frx":0442
      Height          =   315
      Left            =   2205
      Picture         =   "frmAddNewEvent.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   120
      Width           =   420
   End
   Begin VB.CommandButton cmdShowCalendar2 
      DownPicture     =   "frmAddNewEvent.frx":0CC6
      Height          =   315
      Left            =   2205
      Picture         =   "frmAddNewEvent.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   840
      Width           =   420
   End
   Begin VB.Frame fraAlarm 
      Height          =   510
      Left            =   1650
      TabIndex        =   31
      Top             =   2565
      Width           =   2445
      Begin VB.CommandButton cmdSetAlarmToStartDate 
         Caption         =   "="
         Height          =   315
         Left            =   2055
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Set alarm date to event start date"
         Top             =   135
         Width           =   240
      End
      Begin VB.CommandButton cmdShowCalendar3 
         DownPicture     =   "frmAddNewEvent.frx":154A
         Height          =   315
         Left            =   1635
         Picture         =   "frmAddNewEvent.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   135
         Width           =   420
      End
      Begin VB.TextBox txtAlarmDate 
         Height          =   315
         Left            =   660
         MaxLength       =   10
         TabIndex        =   7
         Top             =   135
         Width           =   978
      End
      Begin VB.TextBox txtAlarmTime 
         Height          =   315
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   8
         Top             =   150
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         Caption         =   "Time"
         Height          =   255
         Left            =   840
         TabIndex        =   33
         Top             =   180
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Date"
         Height          =   255
         Left            =   180
         TabIndex        =   32
         Top             =   180
         Width           =   450
      End
   End
   Begin VB.CheckBox chkAlarm 
      Alignment       =   1  'Right Justify
      Caption         =   "Alarm"
      Height          =   210
      Left            =   180
      TabIndex        =   6
      Top             =   2760
      Width           =   1230
   End
   Begin VB.CheckBox chkPrivate 
      Alignment       =   1  'Right Justify
      Caption         =   "Don't export"
      Height          =   315
      Left            =   2985
      TabIndex        =   9
      ToolTipText     =   "Do not include ths entry when exporting the database"
      Top             =   105
      Width           =   1185
   End
   Begin VB.Frame fraRecur 
      Caption         =   "Recurring Event Details"
      Height          =   1155
      Left            =   195
      TabIndex        =   27
      Top             =   3375
      Width           =   6375
      Begin VB.TextBox txtFrequency 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3000
         TabIndex        =   18
         Top             =   660
         Width           =   570
      End
      Begin VB.ComboBox cmbWeekOfMonth 
         Height          =   315
         Left            =   4485
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   660
         Width           =   1605
      End
      Begin VB.ComboBox cmbPeriod 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   660
         Width           =   1605
      End
      Begin VB.Label lblWeekOfMonth 
         Caption         =   "Week of Month"
         Height          =   225
         Left            =   4485
         TabIndex        =   30
         Top             =   405
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Cycle"
         Height          =   225
         Left            =   3030
         TabIndex        =   29
         Top             =   420
         Width           =   660
      End
      Begin VB.Label Label8 
         Caption         =   "Period"
         Height          =   225
         Left            =   480
         TabIndex        =   28
         Top             =   435
         Width           =   1575
      End
   End
   Begin VB.Frame optgrpRecurring 
      Height          =   990
      Left            =   5385
      TabIndex        =   16
      Top             =   2115
      Width           =   1170
      Begin VB.OptionButton optRecurring 
         Caption         =   "Recurring"
         Height          =   240
         Left            =   75
         TabIndex        =   11
         Top             =   600
         Width           =   1050
      End
      Begin VB.OptionButton optSingle 
         Caption         =   "Single"
         Height          =   240
         Left            =   75
         TabIndex        =   10
         Top             =   255
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdEventMaint 
      Caption         =   "Event Type Manager..."
      Height          =   555
      Left            =   5385
      TabIndex        =   14
      Top             =   1485
      Width           =   1155
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   555
      Left            =   5385
      TabIndex        =   13
      Top             =   735
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   555
      Left            =   5385
      TabIndex        =   12
      Top             =   120
      Width           =   1155
   End
   Begin VB.ComboBox cmbEvent 
      Height          =   315
      ItemData        =   "frmAddNewEvent.frx":1DCE
      Left            =   1215
      List            =   "frmAddNewEvent.frx":1DD0
      TabIndex        =   4
      Text            =   "cmbEvent"
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox txtNote 
      Height          =   600
      Left            =   1215
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox txtEndTime 
      Height          =   315
      Left            =   1215
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1200
      Width           =   525
   End
   Begin VB.TextBox txtEndDate 
      Height          =   315
      Left            =   1215
      MaxLength       =   10
      TabIndex        =   2
      Top             =   840
      Width           =   978
   End
   Begin VB.TextBox txtStartTime 
      Height          =   315
      Left            =   1215
      MaxLength       =   5
      TabIndex        =   1
      Top             =   480
      Width           =   525
   End
   Begin VB.TextBox txtStartDate 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1215
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   978
   End
   Begin VB.TextBox txtCong 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2370
      TabIndex        =   15
      Top             =   4995
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Label lblEvent 
      Caption         =   "Event"
      Height          =   255
      Left            =   195
      TabIndex        =   26
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label lblNote 
      Caption         =   "Note"
      Height          =   255
      Left            =   195
      TabIndex        =   25
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label lblEndTime 
      Caption         =   "End Time"
      Height          =   255
      Left            =   195
      TabIndex        =   24
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label lblEndDate 
      Caption         =   "End Date"
      Height          =   255
      Left            =   195
      TabIndex        =   23
      Top             =   900
      Width           =   975
   End
   Begin VB.Label lblStartTime 
      Caption         =   "Start Time"
      Height          =   255
      Left            =   195
      TabIndex        =   22
      Top             =   540
      Width           =   975
   End
   Begin VB.Label lblStart 
      Caption         =   "Start Date"
      Enabled         =   0   'False
      Height          =   255
      Left            =   195
      TabIndex        =   21
      Top             =   180
      Width           =   975
   End
   Begin VB.Label lblCong 
      Caption         =   "Congregation"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1350
      TabIndex        =   20
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmAddNewEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AmIInitialising As Boolean
Dim OralReviewAction As Long, StoreEventID As Integer
Dim mUpdateMode As cmsUpdateModes, miEventTypeID As Integer
Dim mbDoNotTrigger As Boolean
Dim WithEvents frmCal As frmMiniCalendar
Attribute frmCal.VB_VarHelpID = -1
Dim bClickedStartDate As Boolean
Dim mbClickedAlarmDate As Boolean
Dim TheCal As New clsCalendar
Dim mlEventID As Long
Dim msPrevStartDate As String

Private Sub cmdSetAlarmToStartDate_Click()
    txtAlarmDate = txtStartDate
End Sub

Private Sub cmdShowCalendar3_Click()
On Error GoTo ErrorTrap

    mbClickedAlarmDate = True
    
    With frmCal
    .FormDate = txtAlarmDate
    .SetPos = True
    .XPos = Me.Left + fraAlarm.Left + cmdShowCalendar3.Left + cmdShowCalendar3.Width
    .YPos = Me.Top + fraAlarm.Top + cmdShowCalendar3.Top + cmdShowCalendar3.Height
    .Show vbModal, Me
    End With

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub frmCal_InsertDate(TheDate As String)
On Error GoTo ErrorTrap

    If mbClickedAlarmDate Then
        txtAlarmDate = TheDate
    Else
        If bClickedStartDate Then
            txtStartDate = TheDate
        Else
            txtEndDate = TheDate
        End If
    End If

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub


Private Sub cmdShowCalendar1_Click()

On Error GoTo ErrorTrap

    bClickedStartDate = True
    mbClickedAlarmDate = False
    
    With frmCal
    .FormDate = txtStartDate
    .SetPos = True
    .XPos = Me.Left + cmdShowCalendar1.Left + cmdShowCalendar1.Width
    .YPos = Me.Top + cmdShowCalendar1.Top + cmdShowCalendar1.Height
    .Show vbModal, Me
    End With

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdShowCalendar2_Click()
On Error GoTo ErrorTrap

    bClickedStartDate = False
    mbClickedAlarmDate = False
    
    With frmCal
    .FormDate = txtEndDate
    .SetPos = True
    .XPos = Me.Left + cmdShowCalendar2.Left + cmdShowCalendar2.Width
    .YPos = Me.Top + cmdShowCalendar2.Top + cmdShowCalendar2.Height
    .Show vbModal, Me
    End With

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub chkAlarm_Click()
On Error GoTo ErrorTrap

    SetStatusOfAlarmEntry

    If mbDoNotTrigger Then Exit Sub
'    fraAlarm.Enabled = IIf(optRecurring, False, chkAlarm.value)
    If chkAlarm.value = vbUnchecked Then
        txtAlarmDate = ""
        txtAlarmTime = ""
    Else
        txtAlarmDate = txtStartDate
        txtAlarmTime = txtStartTime
    End If

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmbEvent_Click()
On Error GoTo ErrorTrap

     If cmbEvent.ListIndex > -1 Then
        If Not RecurringEventType(cmbEvent.ItemData(Me.cmbEvent.ListIndex)) Then
            optSingle.value = True
            optgrpRecurring.Enabled = False
            optRecurring.Enabled = False
        Else
            optRecurring.ToolTipText = ""
            If mUpdateMode = cmsView Then
                optgrpRecurring.Enabled = False
            Else
                optgrpRecurring.Enabled = True
                optRecurring.Enabled = True
            End If
        End If
        
        'autofill end date and description as appropriate
        With cmbEvent
        If txtStartDate.text <> "" And mUpdateMode = cmsAdd Then
            Select Case .ItemData(.ListIndex)
            Case 1 'circuit assembly
                txtEndDate.text = txtStartDate.text
            Case 3 'district convention
                txtEndDate.text = DateAdd("d", 2, txtStartDate.text)
            Case 4 'co visit
                txtEndDate.text = DateAdd("d", 5, txtStartDate.text)
            Case 5  ' host visit
                txtEndDate.text = DateAdd("d", 5, txtStartDate.text)
            Case 6 'memorial
            Case 15 'reminder
            End Select
            
            chkAlarm = IIf(.ItemData(.ListIndex) = 15, vbChecked, vbUnchecked)
            
            If Trim(txtNote) = "" Then
                txtNote = GlobalCalendar.GetEventTypeDescription(.ItemData(.ListIndex))
            End If

        End If
    
    End With
        
    End If
        
    Exit Sub
ErrorTrap:
    EndProgram
        

End Sub

Private Sub cmbEvent_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap

    If KeyCode = 46 Then
        cmbEvent.ListIndex = -1
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmbEvent_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorTrap
    
    AutoCompleteCombo Me!cmbEvent, KeyAscii
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmbPeriod_Click()
On Error GoTo ErrorTrap
   
    'If Initialising Then Exit Sub
    
    If cmbPeriod.ItemData(cmbPeriod.ListIndex) = 4 Then 'ie Monthly by day
        cmbWeekOfMonth.Visible = True
        lblWeekOfMonth.Visible = True
    Else
        cmbWeekOfMonth.Visible = False
        lblWeekOfMonth.Visible = False
    End If
    
    AdjustWeekCombo
        
    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub cmdClose_Click()
On Error GoTo ErrorTrap
   
    Unload Me
    
    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub cmdEventMaint_Click()
On Error GoTo ErrorTrap
    frmEventTypeManager.Show vbModal, Me
    HandleListBox.Requery cmbEvent, True
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdOK_Click()
Dim CheckIfWeSaveAndExit As Boolean

On Error GoTo ErrorTrap
   
    If mUpdateMode = cmsAdd Or mUpdateMode = cmsEdit Then
        Call ApplyChanges(CheckIfWeSaveAndExit)
        If CheckIfWeSaveAndExit Then
            Unload Me
        End If
    Else
        Unload Me
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub ApplyChanges(SaveAndExit As Boolean)
Dim PeriodID As Integer, PeriodCycle As Integer, WhichWeekOfMonth As Integer
Dim sAlarmDateTime As String, rs As Recordset, str As String, bFound As Boolean
Dim lEvent As Long

On Error GoTo ErrorTrap

    If cmbEvent.ListIndex = -1 Then
        HandleListBox.SelectItem cmbEvent, 15 'default to 'reminder'
    End If
    
    If EntryValidatedOK Then
    
        Select Case optRecurring
        Case True:
            PeriodID = CInt(cmbPeriod.ItemData(cmbPeriod.ListIndex))
            PeriodCycle = CInt(txtFrequency)
            WhichWeekOfMonth = CInt(cmbWeekOfMonth.ItemData(cmbWeekOfMonth.ListIndex))
        Case False:
            PeriodID = 0
            PeriodCycle = 0
            WhichWeekOfMonth = 0
        End Select
        
        lEvent = cmbEvent.ItemData(cmbEvent.ListIndex)
            
        If Not EventConflictsWithTMS(CInt(cmbEvent.ItemData(cmbEvent.ListIndex)), _
                                                            CDate(txtStartDate)) Then
            
            If Not EventConflictsWithPublicMtg(CInt(cmbEvent.ItemData(cmbEvent.ListIndex)), _
                                                            CDate(txtStartDate)) Then
                                                                            
                If (ServiceMtgItemsBetweenDates(GetDateOfGivenDay(CDate(txtStartDate), vbMonday, False), _
                                                    GetDateOfGivenDay(CDate(txtStartDate), vbMonday, False)) _
                    Or CongBibleStudyBetweenDates(GetDateOfGivenDay(CDate(txtStartDate), vbMonday, False), _
                                                    GetDateOfGivenDay(CDate(txtStartDate), vbMonday, False))) _
                    And (lEvent = 1 Or lEvent = 3 Or _
                            (lEvent = 2 And NewMtgArrangementStarted(CStr(GetDateOfGivenDay(CDate(txtStartDate), vbMonday, False))))) Then
                                                                        
                     MsgBox "The Service Meeting and/or Congregation Bible Study conflicts with this event. " & _
                            "First delete entries in the schedules for this week.", vbOKOnly + vbExclamation, AppName
                Else
                                                                        
                    If txtAlarmDate = "" Then
                        If optSingle Then
                            sAlarmDateTime = "01/01/1900" '00:00:00"
                        Else
                            If chkAlarm.value = vbChecked Then
                                sAlarmDateTime = txtStartDate
                            Else
                                sAlarmDateTime = "02/01/1900" ' 00:00:00"
                            End If
                        End If
                    Else
                        sAlarmDateTime = txtAlarmDate  '& " " & txtAlarmTime
                    End If
                    
                    Select Case mUpdateMode
                    Case cmsAdd
                        GlobalCalendar.AddEvent CInt(GlobalDefaultCong), (txtStartDate), _
                            CInt(cmbEvent.ItemData(cmbEvent.ListIndex)), (txtEndDate), txtStartTime, _
                            txtEndTime, txtNote, PeriodID, PeriodCycle, 0, _
                             WhichWeekOfMonth, 0, 0, #1/1/1900#, 0, _
                             chkPrivate, sAlarmDateTime, False
                    Case cmsEdit
                    
                        GlobalCalendar.GetAnEvent CLng(mlEventID)
                        msPrevStartDate = CStr(GlobalCalendar.EventStartDate)
                        Select Case StoreEventID
                        Case 1, 2, 3 'Circuit/District Assembly
                            CMSDB.Execute "DELETE FROM tblTMSSchedule " & _
                                              "WHERE AssignmentDate = #" & Format(GetDateOfGivenDay(CDate(msPrevStartDate), vbMonday, False), "mm/dd/yyyy") & "# " & _
                                              " AND TalkNo = 'A'"
                                              
                            CMSDB.Execute "DELETE FROM tblTMSSchedule " & _
                                              "WHERE AssignmentDate = #" & Format(GetDateOfGivenDay(CDate(msPrevStartDate) + 7, vbMonday, False), "mm/dd/yyyy") & "# " & _
                                              " AND TalkNo = 'MR'"
                                              
                        Case 4 'CO Visit
                            CMSDB.Execute "DELETE FROM tblTMSSchedule " & _
                                              "WHERE AssignmentDate = #" & Format(GetDateOfGivenDay(CDate(msPrevStartDate), vbMonday, False), "mm/dd/yyyy") & "# " & _
                                              " AND TalkNo = 'CO'"
                            CMSDB.Execute "DELETE FROM tblTMSSchedule " & _
                                              "WHERE AssignmentDate = #" & Format(GetDateOfGivenDay(CDate(msPrevStartDate) + 7, vbMonday, False), "mm/dd/yyyy") & "# " & _
                                              " AND TalkNo = 'MR'"
                        End Select
                        
                        GlobalCalendar.GetAnEvent TheCal.SeqNum
                    
                        GlobalCalendar.UpdateEvent TheCal.SeqNum, CInt(GlobalDefaultCong), (txtStartDate), _
                             CInt(cmbEvent.ItemData(cmbEvent.ListIndex)), (txtEndDate), txtStartTime, _
                            txtEndTime, txtNote, PeriodID, PeriodCycle, 0, _
                            WhichWeekOfMonth, GlobalCalendar.LinkSeqNum, 0, #1/1/1900#, GlobalCalendar.LinkEventID, _
                             chkPrivate, sAlarmDateTime, False
                            
                        If FormIsOpen("frmCalendar") Then
                            frmCalendar!MonthView1.value = txtStartDate
                        End If
                        
                    End Select
                    
                    '
                    'Now insert event on tblTMSSchedule as appropriate
                    '
                    Select Case CInt(cmbEvent.ItemData(cmbEvent.ListIndex))
                    Case 1, 3 'Circuit/District Assembly
                        CMSDB.Execute "INSERT INTO tblTMSSchedule " & _
                                          "(AssignmentDate, TalkNo, TalkSeqNum, SchoolNo, PersonID, Assistant1ID,Assistant2ID, Setting) " & _
                                          "VALUES (#" & Format(GetDateOfGivenDay(CDate(txtStartDate), vbMonday, False), "mm/dd/yyyy") & "#, 'A', " & _
                                                  "99, 1, 0, 0, 0, 0)"
                    Case 4 'CO Visit
                        CMSDB.Execute "INSERT INTO tblTMSSchedule " & _
                                          "(AssignmentDate, TalkNo, TalkSeqNum,SchoolNo, PersonID, Assistant1ID,Assistant2ID, Setting) " & _
                                          "VALUES (#" & Format(GetDateOfGivenDay(CDate(txtStartDate), vbMonday, False), "mm/dd/yyyy") & "#, 'CO', " & _
                                                 "99, 1, 0, 0, 0, 0)"
                    Case 2 'spec assbly
                        If NewMtgArrangementStarted(txtStartDate) Then
                            CMSDB.Execute "INSERT INTO tblTMSSchedule " & _
                                              "(AssignmentDate, TalkNo, TalkSeqNum, SchoolNo, PersonID, Assistant1ID,Assistant2ID, Setting) " & _
                                              "VALUES (#" & Format(GetDateOfGivenDay(CDate(txtStartDate), vbMonday, False), "mm/dd/yyyy") & "#, 'A', " & _
                                                     "99, 1, 0, 0, 0, 0)"
                        End If
                    End Select
                
                            
                    If AccessAllowed("frmMainMenu", "cmdServiceMtg") Then
                        
                        Select Case CInt(cmbEvent.ItemData(cmbEvent.ListIndex))
                        Case 4, 5 'Circuit/Host visit
                            
                            str = "SELECT tblServiceMtgs.SeqNum, " & _
                                  "       tblServiceMtgs.MeetingDate, " & _
                                  "       tblServiceMtgs.ItemTypeID, " & _
                                  "       tblServiceMtgs.ItemName, " & _
                                  "       tblServiceMtgs.ItemLength, " & _
                                  "       tblServiceMtgs.PersonID, " & _
                                  "       tblServiceMtgs.Announcements " & _
                                  "FROM tblServiceMtgs " & _
                                  "WHERE MeetingDate = #" & Format(GetDateOfGivenDay(txtStartDate, vbMonday, False), "mm/dd/yyyy") & "# " & _
                                  "ORDER BY SeqNum"
                                  
                            Set rs = CMSDB.OpenRecordset(str, dbOpenDynaset)
                                                                                      
                            With rs
                                                                                                       
                            If Not .BOF Then
                            
                                Do Until .EOF
                                    If CongregationMember.IsTravellingOverseer(!PersonID) Then
                                        bFound = True
                                        Exit Do
                                    End If
                                    .MoveNext
                                Loop
                                
                                If Not bFound Then
                                    MsgBox "A Service Meeting has already been scheduled for this week. It appears " & _
                                            "that no CO's talk has been entered however.", vbOKOnly + vbExclamation, AppName
                                End If
                                    
                            End If
                            
                            End With
                            
                            rs.Close
                            Set rs = Nothing
                                                                                      
                        End Select
                    End If
                    
                    Select Case mUpdateMode
                    Case cmsAdd
                        GlobalCalendar.AddEventAlert CInt(cmbEvent.ItemData(cmbEvent.ListIndex)), CDate(txtStartDate)
                    Case cmsEdit
                        GlobalCalendar.AddEventAlert CInt(cmbEvent.ItemData(cmbEvent.ListIndex)), CDate(txtStartDate), TheCal.SeqNum
                    End Select
                        
                    If FormIsOpen("frmCalendar") Then
                        frmCalendar.RefreshOtherForms
                        frmCalendar.SetUpGrid
                    End If
                    
                    SaveAndExit = True
                End If
            End If
        End If
    Else
        SaveAndExit = False
    End If
    
        
    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Function EntryValidatedOK() As Boolean

On Error GoTo ErrorTrap
    
    txtEndDate = Trim$(txtEndDate)
    txtStartDate = Trim$(txtStartDate)
    txtEndTime = Trim$(txtEndTime)
    txtStartTime = Trim$(txtStartTime)
    txtNote = Trim$(txtNote)
    txtFrequency = Trim$(txtFrequency)
    txtAlarmDate = Trim$(txtAlarmDate)
    txtAlarmTime = Trim$(txtAlarmTime)

    If Not IsDate(Me!txtStartDate) Then
        EntryValidatedOK = False
        MsgBox "Start Date is not valid. ", vbOKOnly + vbExclamation, AppName
        TextFieldGotFocus txtStartDate, True
        Exit Function
    Else
        EntryValidatedOK = True
    End If
        
    If Me!txtEndDate = "" Then
        Me!txtEndDate = Me!txtStartDate
    End If
    
    If Not IsDate(Me!txtEndDate) Then
        EntryValidatedOK = False
        MsgBox "End Date is not valid. ", vbOKOnly + vbExclamation, AppName
        TextFieldGotFocus txtEndDate, True
        Exit Function
    Else
        EntryValidatedOK = True
    End If
    
    If txtStartTime <> "" Then
        If Not IsTime(Me!txtStartTime) Then
            EntryValidatedOK = False
            MsgBox "Start time is not valid. ", vbOKOnly + vbExclamation, AppName
            TextFieldGotFocus txtStartTime, True
            Exit Function
        Else
            EntryValidatedOK = True
        End If
    Else
        EntryValidatedOK = True
    End If
    
    If txtEndTime <> "" Then
        If Not IsTime(Me!txtEndTime) Then
            EntryValidatedOK = False
            MsgBox "End time is not valid. ", vbOKOnly + vbExclamation, AppName
            TextFieldGotFocus txtEndTime, True
            Exit Function
        Else
            EntryValidatedOK = True
        End If
    Else
        EntryValidatedOK = True
    End If
    
    If CDate(txtStartDate) > CDate(txtEndDate) Then
        EntryValidatedOK = False
        MsgBox "Start date must be prior to End date.", vbOKOnly + vbExclamation, AppName
        TextFieldGotFocus txtEndDate, True
        Exit Function
    Else
        EntryValidatedOK = True
    End If
    
    If CDate(txtStartDate) = CDate(txtEndDate) Then
        If txtStartTime <> "" And txtEndTime <> "" Then
            If CDate(txtStartTime) > CDate(txtEndTime) Then
                EntryValidatedOK = False
                MsgBox "Start time must be prior to End time.", vbOKOnly + vbExclamation, AppName
                TextFieldGotFocus txtEndDate, True
                Exit Function
            Else
                EntryValidatedOK = True
            End If
        End If
    Else
        EntryValidatedOK = True
    End If
    
    If chkAlarm Then
        If optSingle Then
            If txtAlarmDate = "" Then
                txtAlarmDate = txtStartDate
            End If
            If Not IsDate(Me!txtAlarmDate) Then
                EntryValidatedOK = False
                MsgBox "Alarm Date is not valid. ", vbOKOnly + vbExclamation, AppName
                TextFieldGotFocus txtAlarmDate, True
                Exit Function
            Else
                EntryValidatedOK = True
            End If
            If txtAlarmTime = "" Then
                txtAlarmTime = "09:00"
            End If
            If Not IsTime(Me!txtAlarmTime) Then
                EntryValidatedOK = False
                MsgBox "Alarm time is not valid. ", vbOKOnly + vbExclamation, AppName
                TextFieldGotFocus txtAlarmTime, True
                Exit Function
            Else
                EntryValidatedOK = True
            End If
        Else
'            chkAlarm = vbUnchecked
            EntryValidatedOK = True
        End If
    End If
    
    If cmbEvent.ListIndex = -1 Then
        EntryValidatedOK = False
        MsgBox "Please select an Event from the list.", vbOKOnly + vbExclamation, AppName
        Me!cmbEvent.SetFocus
        Exit Function
    Else
        EntryValidatedOK = True
    End If

    If Len(txtNote) > 150 Then
        EntryValidatedOK = False
        MsgBox "The 'Note' field should contain no more than 150 characters.", vbOKOnly + vbExclamation, AppName
        TextFieldGotFocus txtNote, True
        Exit Function
    Else
        EntryValidatedOK = True
    End If
    
    If optRecurring Then
        If cmbPeriod.ListIndex = -1 Then
            EntryValidatedOK = False
            MsgBox "Please select a recurring period", vbOKOnly + vbExclamation, AppName
            Me!cmbPeriod.SetFocus
            Exit Function
        Else
            EntryValidatedOK = True
        End If
        If cmbWeekOfMonth.ListIndex = -1 Then
            EntryValidatedOK = False
            MsgBox "Please select a week", vbOKOnly + vbExclamation, AppName
            Me!cmbWeekOfMonth.SetFocus
            Exit Function
        Else
            EntryValidatedOK = True
        End If
        If IsNumeric(txtFrequency) Then
            If txtFrequency < 1 Or txtFrequency > 1000 Then
                EntryValidatedOK = False
                MsgBox "Frequency should be a number between 1 and 1000", vbOKOnly + vbExclamation, AppName
                Me!txtFrequency.SetFocus
                Exit Function
            Else
                EntryValidatedOK = True
            End If
        Else
            EntryValidatedOK = False
            MsgBox "Frequency should be a number between 1 and 1000", vbOKOnly + vbExclamation, AppName
            Me!txtFrequency.SetFocus
            Exit Function
        End If
    Else
        EntryValidatedOK = True
    End If
    

    '
    'THE FOLLOWING 4 IF STMTS SHOULD COME LAST IN THIS PROC!!!
    '
    If txtStartTime = "" And txtEndTime <> "" Then
        txtStartTime = txtEndTime
    End If
    
    If txtEndTime = "" And txtStartTime <> "" Then
        txtEndTime = txtStartTime
    End If
        
    
    If txtStartTime = "" Then
        txtStartTime = "00:00"
    End If
        
    If txtEndTime = "" Then
        txtEndTime = "00:00"
    End If
                
    Exit Function
ErrorTrap:
    EndProgram
    
End Function


Private Sub Form_Activate()
    On Error GoTo ErrorTrap

    If mUpdateMode = cmsEdit Then
        StoreEventID = miEventTypeID
        HandleListBox.SelectItem cmbEvent, CLng(miEventTypeID)
    Else
        If AmIInitialising Then
            cmbEvent.ListIndex = -1
            optSingle = True
        End If
    End If
    
    TheCal.GetAnEvent mlEventID
        
    If Not TheCal.NoEventsFound Then
        If TheCal.PeriodID = 0 Then
        'NOT a recurring event
            optSingle = True
            fraRecur.Visible = False
        Else
            AmIInitialising = False
            optRecurring = True
            AmIInitialising = True
            fraRecur.Visible = True
            HandleListBox.PopulateListBox cmbPeriod, "SELECT ID, PeriodDescription FROM " & _
                "tblEventRecurringLookup", CMSDB, 0, "", False, 1
            HandleListBox.PopulateListBox cmbWeekOfMonth, "SELECT ID, WeekNum FROM " & _
                "tblWeekOfMonth", CMSDB, 0, "", False, 1
                
            HandleListBox.SelectItem cmbPeriod, TheCal.PeriodID
            HandleListBox.SelectItem cmbWeekOfMonth, TheCal.WhichWeekOfMonth
            txtFrequency = TheCal.PeriodCycle
        End If
    End If
    
    Select Case mUpdateMode
    Case cmsView
        lblCong.Enabled = False
        lblstart.Enabled = False
        lblStartTime.Enabled = False
        lblEndDate.Enabled = False
        lblEndTime.Enabled = False
        lblEvent.Enabled = False
        lblNote.Enabled = False
        txtStartDate.Enabled = False
        cmdShowCalendar1.Enabled = False
        cmdShowCalendar2.Enabled = False
        txtEndDate.Enabled = False
        txtEndTime.Enabled = False
        txtNote.Enabled = False
        txtStartTime.Enabled = False
        cmbEvent.Enabled = False
        cmbPeriod.Enabled = False
        txtFrequency.Enabled = False
        cmbWeekOfMonth.Enabled = False
        cmdOK.Visible = False
        optgrpRecurring.Enabled = False
        fraAlarm.Enabled = False
        chkAlarm.Enabled = False
        chkPrivate.Enabled = False
        txtAlarmDate.Enabled = False
        cmdShowCalendar3.Enabled = False
        cmdSetAlarmToStartDate.Enabled = False
        txtAlarmTime.Enabled = False
        lblDate.Enabled = False
        lblTime.Enabled = False
    Case cmsAdd
        lblCong.Enabled = False
        lblstart.Enabled = True
        lblStartTime.Enabled = True
        lblEndDate.Enabled = True
        lblEndTime.Enabled = True
        lblEvent.Enabled = True
        lblNote.Enabled = True
        txtStartDate.Enabled = True
        cmdShowCalendar1.Enabled = True
        cmdShowCalendar2.Enabled = True
        txtEndDate.Enabled = True
        txtEndTime.Enabled = True
        txtNote.Enabled = True
        txtStartTime.Enabled = True
        cmbEvent.Enabled = True
        cmbPeriod.Enabled = True
        txtFrequency.Enabled = True
        cmbWeekOfMonth.Enabled = True
        cmdOK.Visible = True
        optgrpRecurring.Enabled = True
        chkAlarm.Enabled = True
        SetStatusOfAlarmEntry
        chkPrivate.Enabled = True
    Case cmsEdit
        lblCong.Enabled = False
        lblstart.Enabled = True
        lblStartTime.Enabled = True
        lblEndDate.Enabled = True
        lblEndTime.Enabled = True
        lblEvent.Enabled = True
        lblNote.Enabled = True
        txtStartDate.Enabled = True
        cmdShowCalendar1.Enabled = True
        cmdShowCalendar2.Enabled = True
        txtEndDate.Enabled = True
        txtEndTime.Enabled = True
        txtNote.Enabled = True
        txtStartTime.Enabled = True
        cmbEvent.Enabled = True
        cmbPeriod.Enabled = True
        txtFrequency.Enabled = True
        cmbWeekOfMonth.Enabled = True
        cmdOK.Visible = True
        optgrpRecurring.Enabled = True
        chkAlarm.Enabled = True
        SetStatusOfAlarmEntry
        chkPrivate.Enabled = True
    End Select
        
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrorTrap
    
    AmIInitialising = True
            
'    ControlForTestOnly chkAlarm, True, False
'    ControlForTestOnly chkPrivate, True, False
'
''    ControlForTestOnly fraAlarm,   True, False
'

    Set frmCal = New frmMiniCalendar

    AmIInitialising = False
    
    Exit Sub
ErrorTrap:
    EndProgram

    
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Unload frmCal
    Set frmCal = Nothing
    
    Set TheCal = Nothing

End Sub

Private Sub optRecurring_Click()
    On Error GoTo ErrorTrap
    
    If AmIInitialising Then Exit Sub
    
    fraRecur.Visible = True
    HandleListBox.PopulateListBox cmbPeriod, "SELECT ID, PeriodDescription FROM " & _
        "tblEventRecurringLookup", CMSDB, 0, "", False, 1
    HandleListBox.PopulateListBox cmbWeekOfMonth, "SELECT ID, WeekNum FROM " & _
        "tblWeekOfMonth", CMSDB, 0, "", False, 1
        
    AdjustWeekCombo
        
    cmbPeriod.ListIndex = -1
    'cmbWeekOfMonth.ListIndex = -1
    txtFrequency = ""
    
    Me.Height = 5040
    
'    fraAlarm.Enabled = IIf(optRecurring, False, chkAlarm.value)
    
    If chkAlarm.value = vbUnchecked Then
        txtAlarmDate = ""
        txtAlarmTime = ""
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub optSingle_Click()
    On Error GoTo ErrorTrap
    
    If AmIInitialising Then Exit Sub
    
    fraRecur.Visible = False
    
    Me.Height = 3660
    
    SetStatusOfAlarmEntry
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub SetStatusOfAlarmEntry()
On Error GoTo ErrorTrap
    
    If optRecurring Then
'        fraAlarm.Enabled = False
'        lblDate.Enabled = False
'        lblTime.Enabled = False
        chkAlarm.ToolTipText = "Alarm date will be that of recurring event"
    Else
        If chkAlarm.value = vbChecked Then
            fraAlarm.Enabled = True
            lblDate.Enabled = True
'            lblTime.Enabled = True
            chkAlarm.ToolTipText = "Enter date you want to be alerted"
        Else
'            fraAlarm.Enabled = False
            lblDate.Enabled = False
 '           lblTime.Enabled = False
            chkAlarm.ToolTipText = ""
        End If
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub



Private Sub txtstartdate_Change()

On Error GoTo ErrorTrap

    If ValidDate(txtStartDate) Then
        If ValidDate(txtEndDate) Then
            If CDate(txtEndDate) < CDate(txtStartDate) Then
                txtEndDate = txtStartDate
            End If
        End If
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub txtAlarmDate_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorTrap

    KeyPressValid KeyAscii, cmsDates

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub txtAlarmDate_LostFocus()
    txtAlarmDate = Format(Trim(txtAlarmDate), "DD/MM/YYYY")
End Sub
Private Sub txtAlarmTime_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorTrap

    KeyPressValid KeyAscii, cmsTimes

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub txtAlarmTime_LostFocus()
    txtAlarmTime = Trim(txtAlarmTime)
End Sub

Private Sub txtEndDate_LostFocus()
    txtEndDate = Format(Trim(txtEndDate), "DD/MM/YYYY")
End Sub

Private Sub txtEndTime_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorTrap

'Must be numeric. Allow Backspace (8) and colon (58). Delete and arrow keys seem to be allowed by default.

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 58 Then
        KeyAscii = 0
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub txtEndTime_LostFocus()
    txtEndTime = Trim(txtEndTime)
End Sub



Private Sub txtFrequency_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorTrap

'Must be numeric. Allow Backspace (8) and colon (58). Delete and arrow keys seem to be allowed by default.

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram


End Sub



Private Sub txtNote_GotFocus()
    TextFieldGotFocus txtNote
End Sub

Private Sub txtStartDate_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorTrap

'Must be numeric. Allow Backspace (8) and forward-slash (47). Delete and arrow keys seem to be allowed by default.

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 47 Then
        KeyAscii = 0
    End If
    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub txtEndDate_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorTrap

'Must be numeric. Allow Backspace (8) and forward-slash (47). Delete and arrow keys seem to be allowed by default.

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 47 Then
        KeyAscii = 0
    End If
    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub txtStartDate_LostFocus()
On Error GoTo ErrorTrap
    
    txtStartDate = Format$(Trim$(txtStartDate), "DD/MM/YYYY", vbFriday, vbFirstJan1)
    AdjustWeekCombo
    
    Exit Sub
ErrorTrap:
    EndProgram

    
End Sub

Private Sub txtStartTime_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorTrap

'Must be numeric. Allow Backspace (8) and colon (58). Delete and arrow keys seem to be allowed by default.

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 58 Then
        KeyAscii = 0
    End If
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub txtStartTime_LostFocus()
    txtStartTime = Trim(txtStartTime)
End Sub

Public Property Get Initialising() As Boolean
    Initialising = AmIInitialising
End Property

Public Property Let Initialising(ByVal vNewValue As Boolean)
    AmIInitialising = vNewValue
End Property
Public Property Get DoNotTrigger() As Boolean
    DoNotTrigger = mbDoNotTrigger
End Property

Public Property Let DoNotTrigger(ByVal vNewValue As Boolean)
    mbDoNotTrigger = vNewValue
End Property

Public Property Let UpdateMode(ByVal vNewValue As cmsUpdateModes)
    mUpdateMode = vNewValue
End Property
Public Property Let EventTypeID(ByVal vNewValue As Integer)
    miEventTypeID = vNewValue
End Property
Public Property Let EventID(ByVal vNewValue As Long)
    mlEventID = vNewValue
End Property

Private Sub AdjustWeekCombo()
    On Error GoTo ErrorTrap
        
    If IsDate(txtStartDate) Then
        If OrdinalPosOfDay(CDate(txtStartDate)) = "L" Then
            HandleListBox.SelectItem cmbWeekOfMonth, 5
        Else
            HandleListBox.SelectItem cmbWeekOfMonth, CLng(OrdinalPosOfDay(CDate(txtStartDate)))
        End If
    Else
        cmbWeekOfMonth.ListIndex = -1
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Function EventConflictsWithTMS(EventID As Integer, EventDate As Date) As Boolean
Dim rsttmsquery As Recordset, OriginalOralReviewDate As Date

    On Error GoTo ErrorTrap
        
    Select Case EventID
    Case 1, 3 'Circuit/District Assembly week - all TMS items should be cancelled
    
        '
        'Select all scheduled students which fall in same week/year of assembly date
        '
        Set rsttmsquery = CMSDB.OpenRecordset("SELECT * " & _
                                           "FROM tblTMSSchedule " & _
                                           "WHERE AssignmentDate = " & GetDateStringForSQLWhere(GetDateOfGivenDay(EventDate, vbMonday)) & _
                                            " AND (PersonID > 0 OR Assistant1ID > 0)" _
                                           , dbOpenDynaset)
        
        With rsttmsquery
        
        If .BOF Then 'No conflicts found with student talks
            '
            'Is there an Oral review this week?
            '
            If OralReviewConflict(EventDate, EventID) Then
                EventConflictsWithTMS = True
                Exit Function
            Else
                EventConflictsWithTMS = False
                Exit Function
            End If
        Else
            If Not OralReviewConflict(EventDate, EventID) Then
                If MsgBox("This event conflicts with scheduled assignments on the School. " & _
                            "Do you want to delete these assignments?", vbYesNo + vbQuestion, AppName) = vbYes Then
                    
                    rsttmsquery.FindFirst "SlipPrinted = TRUE"
                    If Not rsttmsquery.NoMatch Then
                        If MsgBox("A slip has been printed for at least one of these " & _
                                  "assignments. Do you want to proceed with this operation?", vbYesNo + vbQuestion, AppName) = vbYes Then
                            
                            TheTMS.DeleteTMSAssignmentsOnRecSet rsttmsquery
                            EventConflictsWithTMS = False
                        Else
                            EventConflictsWithTMS = True
                        End If
                    Else
                        TheTMS.DeleteTMSAssignmentsOnRecSet rsttmsquery
                        EventConflictsWithTMS = False
                    End If
                Else
                    EventConflictsWithTMS = True
                End If
            Else
                EventConflictsWithTMS = True
                Exit Function
            End If
        End If
        
        End With
    Case 4 'CO Visit. Items 2,3,4 should be cancelled. Oral Review should be moved to next week.
        '
        'Select all scheduled students which fall in same week/year of CO Visit
        '
        If NewMtgArrangementStarted(CStr(GetDateOfGivenDay(EventDate, vbMonday, False))) Then
            Set rsttmsquery = CMSDB.OpenRecordset("SELECT * " & _
                                               "FROM tblTMSSchedule " & _
                                               "WHERE AssignmentDate = " & GetDateStringForSQLWhere(GetDateOfGivenDay(EventDate, vbMonday)) & _
                                                " AND (PersonID > 0 OR Assistant1ID > 0)" & _
                                                " AND TalkNo IN ('1', '2', '3')" _
                                               , dbOpenDynaset)
        Else
            Set rsttmsquery = CMSDB.OpenRecordset("SELECT * " & _
                                               "FROM tblTMSSchedule " & _
                                               "WHERE AssignmentDate = " & GetDateStringForSQLWhere(GetDateOfGivenDay(EventDate, vbMonday)) & _
                                                " AND (PersonID > 0 OR Assistant1ID > 0)" & _
                                                " AND TalkNo IN ('2', '3', '4')" _
                                               , dbOpenDynaset)
        End If
        
        With rsttmsquery
        
        If .BOF Then 'No conflicts found with student talks
            '
            'Is there an Oral review this week?
            '
            If OralReviewConflict(EventDate, EventID) Then
                EventConflictsWithTMS = True
                Exit Function
            Else
                EventConflictsWithTMS = False
                Exit Function
            End If
        Else
            If Not OralReviewConflict(EventDate, EventID) Then
                If Not NewMtgArrangementStarted(CStr(EventDate)) Then
                    If MsgBox("This event conflicts with scheduled assignments on the School. " & _
                                "Do you want to delete these assignments?", vbYesNo + vbQuestion, AppName) = vbYes Then
                        
                        rsttmsquery.FindFirst "SlipPrinted = TRUE"
                        If Not rsttmsquery.NoMatch Then
                            If MsgBox("A slip has been printed for at least one of these " & _
                                      "assignments. Do you want to proceed with this operation?", vbYesNo + vbQuestion, AppName) = vbYes Then
                                
                                TheTMS.DeleteTMSAssignmentsOnRecSet rsttmsquery
                                EventConflictsWithTMS = False
                            Else
                                EventConflictsWithTMS = True
                            End If
                        Else
                            TheTMS.DeleteTMSAssignmentsOnRecSet rsttmsquery
                            EventConflictsWithTMS = False
                        End If
                    Else
                        EventConflictsWithTMS = True
                    End If
                Else
                    EventConflictsWithTMS = False
                End If
            Else
                EventConflictsWithTMS = True
                Exit Function
            End If
        End If
        
        End With

    
    Case 6 'Memorial night

        '
        'Select all scheduled students which fall on same day as memorial date
        '
        If Weekday(EventDate) = glMidWkMtgDay Then
            Set rsttmsquery = CMSDB.OpenRecordset("SELECT * " & _
                                               "FROM tblTMSSchedule " & _
                                               "WHERE AssignmentDate = " & GetDateStringForSQLWhere(GetDateOfGivenDay(EventDate, vbMonday)) & _
                                             " AND (PersonID > 0 OR Assistant1ID > 0)" _
                                              , dbOpenDynaset)
        Else
            EventConflictsWithTMS = False
            Exit Function
        End If
        

        With rsttmsquery

        If .BOF Then 'No apparent conflicts found with student talks
            '
            'Is there an Oral review this week, and on day of memorial?
            '
            If Weekday(EventDate) = GlobalParms.GetValue("MidWeekMeetingDay", "NumVal") Then
                If OralReviewConflict(EventDate, EventID) Then
                    EventConflictsWithTMS = True
                    Exit Function
                Else
                    EventConflictsWithTMS = False
                    Exit Function
                End If
            Else
                EventConflictsWithTMS = False
                Exit Function
            End If
        Else 'There is a conflict between the memorial and assigned talks
            If Weekday(EventDate) = GlobalParms.GetValue("MidWeekMeetingDay", "NumVal") Then
                If OralReviewConflict(EventDate, EventID) Then
                    EventConflictsWithTMS = True
                    Exit Function
                Else
                    If MsgBox("This event conflicts with scheduled assignments on the School. " & _
                                "Do you want to delete these assignments?", vbYesNo + vbQuestion, AppName) = vbYes Then

                        rsttmsquery.FindFirst "SlipPrinted = TRUE"
                        If Not rsttmsquery.NoMatch Then
                            If MsgBox("A slip has been printed for at least one of these " & _
                                      "assignments. Do you want to proceed with this operation?", vbYesNo + vbQuestion, AppName) = vbYes Then

                                TheTMS.DeleteTMSAssignmentsOnRecSet rsttmsquery
                                EventConflictsWithTMS = False
                            Else
                                EventConflictsWithTMS = True
                            End If
                        Else
                            TheTMS.DeleteTMSAssignmentsOnRecSet rsttmsquery
                            EventConflictsWithTMS = False
                        End If
                    Else
                        EventConflictsWithTMS = True
                    End If
                End If
            Else
                EventConflictsWithTMS = False
                Exit Function
            End If
        End If

        End With

    Case Else
        EventConflictsWithTMS = False
    End Select
    
    Exit Function
ErrorTrap:
    EndProgram

End Function
Private Function EventConflictsWithPublicMtg(EventID As Integer, EventDate As Date) As Boolean
Dim rst As Recordset, dte As Date

    On Error GoTo ErrorTrap
    
    dte = GetDateOfGivenDay(EventDate, vbMonday, False)
        
    Select Case EventID
    Case 1, 3  'Circuit/District Assembly week - meeting should be cancelled
    
    
        '
        'Select meeting which falls in same week of assembly date
        '
        Set rst = CMSDB.OpenRecordset("SELECT * " & _
                                        "FROM tblPublicMtgSchedule " & _
                                        "WHERE MeetingDate = #" & Format(dte, "mm/dd/yyyy") & "# " & _
                                        "AND (SpeakerID > 0 OR TalkNo > 0 OR ChairmanID > 0 " & _
                                        "     OR WTReaderID > 0 OR (Info <> '' AND Info IS NOT NULL)) " & _
                                        "AND CongNoWhereMtgIs = " & GlobalDefaultCong _
                                        , dbOpenDynaset)
        
        With rst
        
        If .BOF Then 'No conflicts found
            EventConflictsWithPublicMtg = False
            
        Else
            If MsgBox("This assembly conflicts with a scheduled Public Meeting. " & _
                        "Do you want to delete the meeting?", vbYesNo + vbQuestion, AppName) = vbYes Then
                
               .Delete
                
                EventConflictsWithPublicMtg = False
            Else
                EventConflictsWithPublicMtg = True
            End If
        End If
        
        End With
        
    Case 4 'CO Visit.
        '
        'Select meeting which falls in same week of CO Visit
        '
        Set rst = CMSDB.OpenRecordset("SELECT * " & _
                                        "FROM tblPublicMtgSchedule  " & _
                                        "WHERE MeetingDate = #" & Format(dte, "mm/dd/yyyy") & "# " & _
                                        "AND CongNoWhereMtgIs = " & GlobalDefaultCong _
                                        , dbOpenDynaset)
        
        With rst
        
        If .BOF Then 'No meeting found
            EventConflictsWithPublicMtg = False
        Else
            If CongregationMember.CongForPerson(!SpeakerID) <> 32767 Then 'pub mtg found, but speaker not CO
                If MsgBox("This Circuit Visit conflicts with a scheduled Public Meeting. " & _
                            "Do you want to delete the meeting?", vbYesNo + vbQuestion, AppName) = vbYes Then
                    
                   .Delete
                   
                   EventConflictsWithPublicMtg = False
                    
                Else
                    EventConflictsWithPublicMtg = True
                End If
            Else
                EventConflictsWithPublicMtg = False
                .Edit 'no reader required
                !WTReaderID = 0
                .Update
            End If
        End If
              
        End With
        
    Case Else
        EventConflictsWithPublicMtg = False
    End Select
    
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    On Error GoTo ErrorTrap
    
    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Property Get DoWhatWithOralReview() As Long
    DoWhatWithOralReview = OralReviewAction
End Property

Public Property Let DoWhatWithOralReview(ByVal vNewValue As Long)
    OralReviewAction = vNewValue
End Property


Private Function OralReviewConflict(EventDate As Date, EventID As Integer) As Boolean
Dim rsttmsquery As Recordset, OriginalOralReviewDate As Date, COVisitMsg As String, tmpDate As Date
Dim TMSDate As Date
Dim bOK As Boolean
On Error GoTo ErrorTrap


    TMSDate = GetDateOfGivenDay(EventDate, vbMonday, False)
    
    If IsMovedOralReviewWeek(TMSDate) Then
        OralReviewConflict = True
        Exit Function
    End If
    
    
                                       
    If Not IsOralReviewWeek(TMSDate) Then
        'No Oral review on event week
        OralReviewConflict = False
        Exit Function
    Else
        OriginalOralReviewDate = TMSDate
        frmDoWhatWithOralReview.Show vbModal
        Select Case DoWhatWithOralReview
        Case 0 'Do nothing at all (ie don't insert assembly)
            OralReviewConflict = True
            Exit Function
        Case 1 'Move Oral Review to next available week
        
            If NewMtgArrangementStarted(CStr(OriginalOralReviewDate)) Then
            
                'find next meeting where no assemblies or CO visit...
                tmpDate = OriginalOralReviewDate + 7
        
                '
                'Any assignments no 1,2,3 next week?
                '
                
                Set rsttmsquery = CMSDB.OpenRecordset("SELECT * " & _
                                                   "FROM tblTMSSchedule " & _
                                                    "WHERE AssignmentDate = #" & Format(tmpDate, "mm/dd/yyyy") & "# " & _
                                                     " AND TalkNo IN ('1', '2', '3')" & _
                                                     " AND (PersonID > 0 OR Assistant1ID > 0)" _
                                                    , dbOpenDynaset)
                                                    
    
                Select Case EventID
                Case 1 'circuit assembly
                    COVisitMsg = "These assignments will be handled at the assembly."
                Case 3 'district convention
                    COVisitMsg = ""
                Case 4 'CO visit
                    COVisitMsg = "Assignments 1, 2 and 3 will be " & _
                                "brought forward from week commencing " & tmpDate & _
                                ", but the students will be DELETED."
                Case Else
                    COVisitMsg = ""
                End Select
                
                If MsgBox("Moving the Oral Review to the week commencing " & tmpDate & " will mean the DELETION of " & _
                            "assignments 1, 2, 3 on that week. " & COVisitMsg & " Is this OK?", vbYesNo + vbQuestion, AppName) = vbYes Then
                    
                    rsttmsquery.FindFirst "SlipPrinted = TRUE"
                    If Not rsttmsquery.NoMatch Then
                        If MsgBox("A slip has been printed for at least one of these " & _
                                  "assignments. Do you want to proceed with this operation?", vbYesNo + vbQuestion, AppName) = vbYes Then
                            
                            TheTMS.DeleteTMSAssignmentsOnRecSet rsttmsquery
                        Else
                            OralReviewConflict = True
                            Exit Function
                        End If
                    Else
                        TheTMS.DeleteTMSAssignmentsOnRecSet rsttmsquery
                    End If
                Else
                    OralReviewConflict = True
                    Exit Function
                End If
        
            Else ' pre-2009
        
                'find next meeting where no assemblies or CO visit...
                tmpDate = OriginalOralReviewDate + 7
'                bOK = False
'                Do Until bOK
'                    If IsCOVisitWeek(TmpDate) Or _
'                        IsCircuitOrDistrictAssemblyWeek(TmpDate) Then
'
'                        TmpDate = TmpDate + 7
'                    Else
'                        bOK = True
'                    End If
'                Loop
        
                '
                'Any assignments no 1,2,3,4 next week?
                '
                
                Set rsttmsquery = CMSDB.OpenRecordset("SELECT * " & _
                                                   "FROM tblTMSSchedule " & _
                                                    "WHERE AssignmentDate = #" & Format(tmpDate, "mm/dd/yyyy") & "# " & _
                                                     " AND TalkNo IN ('1', '2', '3', '4')" & _
                                                     " AND (PersonID > 0 OR Assistant1ID > 0)" _
                                                    , dbOpenDynaset)
                                                    
    
                Select Case EventID
                Case 1, 3 'assembly/convention
                    COVisitMsg = "The Speech Quality and Bible Highlights will be " & _
                                "moved to week commencing " & tmpDate & _
                                " but the speakers will be DELETED."
                Case 4 'CO visit
                    COVisitMsg = "Talk No 1 will be " & _
                                "brought forward from week commencing " & tmpDate & _
                                " but the speaker will be DELETED."
                Case Else
                    COVisitMsg = ""
                End Select
                
                If MsgBox("Moving the Oral Review to the week commencing " & tmpDate & " will mean the DELETION of " & _
                            "assignments 1, 2, 3 and 4 on that week. " & COVisitMsg & " Is this OK?", vbYesNo + vbQuestion, AppName) = vbYes Then
                    
                    rsttmsquery.FindFirst "SlipPrinted = TRUE"
                    If Not rsttmsquery.NoMatch Then
                        If MsgBox("A slip has been printed for at least one of these " & _
                                  "assignments. Do you want to proceed with this operation?", vbYesNo + vbQuestion, AppName) = vbYes Then
                            
                            TheTMS.DeleteTMSAssignmentsOnRecSet rsttmsquery
                        Else
                            OralReviewConflict = True
                            Exit Function
                        End If
                    Else
                        TheTMS.DeleteTMSAssignmentsOnRecSet rsttmsquery
                    End If
                Else
                    OralReviewConflict = True
                    Exit Function
                End If
                
            End If
            
            CMSDB.Execute "INSERT INTO tblTMSSchedule " & _
                              "(AssignmentDate, TalkNo, TalkSeqNum, SchoolNo, PersonID, Assistant1ID,Assistant2ID, Setting) " & _
                              "VALUES (#" & Format(tmpDate, "mm/dd/yyyy") & "#, 'MR', " & _
                                      "4, 1, 0, 0, 0, 0)"
                                     
                 
            OralReviewConflict = False
            Exit Function
                 
        Case 2 'Insert event and do nothing with Oral Review. Effectively: delete review
            OralReviewConflict = False
            Exit Function
        End Select
    End If

    Exit Function
ErrorTrap:
    EndProgram
End Function
