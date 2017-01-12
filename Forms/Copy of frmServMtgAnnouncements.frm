VERSION 5.00
Begin VB.Form frmServMtgAnnouncements 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " C.M.S. Service Meeting Announcements"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   Icon            =   "frmServMtgAnnouncements.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAutoText 
      Caption         =   "Auto-Text >"
      Height          =   315
      Left            =   135
      TabIndex        =   5
      ToolTipText     =   "Print Using Microsoft Word"
      Top             =   3825
      Width           =   1005
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   315
      Left            =   4830
      TabIndex        =   1
      ToolTipText     =   "Print Using Microsoft Word"
      Top             =   3825
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Caption         =   "Announcements"
      Height          =   3165
      Left            =   135
      TabIndex        =   3
      Top             =   555
      Width           =   6750
      Begin VB.TextBox txtAnnouncements 
         Height          =   2745
         Left            =   105
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   255
         Width           =   6510
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   315
      Left            =   5865
      TabIndex        =   2
      Top             =   3825
      Width           =   1005
   End
   Begin VB.Label lblDesc 
      Height          =   300
      Left            =   150
      TabIndex        =   4
      Top             =   180
      Width           =   4605
   End
   Begin VB.Menu mnuAction 
      Caption         =   "Action"
      Visible         =   0   'False
      Begin VB.Menu mnuAttendants 
         Caption         =   "Attendants"
      End
      Begin VB.Menu mnuReports 
         Caption         =   "Field Service Reports"
      End
      Begin VB.Menu mnuCleaningThisWeek 
         Caption         =   "Group Cleaning"
      End
      Begin VB.Menu mnuDays 
         Caption         =   "List Days of Week"
      End
   End
End
Attribute VB_Name = "frmServMtgAnnouncements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mdteMeetingDate As Date, msAnnouncements As String
Dim mbChangeMade As Boolean, mbFormLoading As Boolean
Dim mbSaveChanged As Boolean

Private Function SaveData() As Boolean
Dim rs As Recordset
On Error GoTo ErrorTrap

    msAnnouncements = txtAnnouncements
    frmServiceMtg.Announcements = msAnnouncements
    
    Set rs = CMSDB.OpenRecordset("tblServiceMtgs", dbOpenDynaset)
    
    With rs
    
    .FindFirst "MeetingDate = #" & Format(mdteMeetingDate, "mm/dd/yyyy") & "# AND ItemTypeID = 0"
     
     If Not .NoMatch Then
        .Edit
        !Announcements = msAnnouncements
        .Update
        MsgBox "Announcements saved.", vbOKOnly + vbInformation, AppName
     Else
        MsgBox "Announcements will be saved when meeting scheduling is saved on the Service " & _
                "Meeting dialogue.", vbOKOnly + vbInformation, AppName
     End If
     
    End With
    
    mbChangeMade = False
    cmdApply.Enabled = False

    Exit Function
ErrorTrap:
    Call EndProgram
End Function

Private Sub FillTheForm()
Dim rstTemp As Recordset, lResponse As Long
On Error GoTo ErrorTrap

    txtAnnouncements.text = msAnnouncements
    
    mbChangeMade = False
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub



Private Sub cmdApply_Click()

On Error GoTo ErrorTrap

    SaveData

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdAutoText_Click()

On Error GoTo ErrorTrap

    Me.PopupMenu mnuAction

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdClose_Click()

On Error GoTo ErrorTrap

    Unload Me

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub



Private Sub Form_Load()

On Error GoTo ErrorTrap

    mbFormLoading = True
    
    FillTheForm
    
    mbChangeMade = False

    mbFormLoading = False
    
    mbSaveChanged = frmServiceMtg.ChangeMade

    lblDesc.Caption = "Announcements for Service Meeting: " & CStr(mdteMeetingDate)
    
    cmdApply.Enabled = False
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo ErrorTrap

    If mbChangeMade Then
        If MsgBox("Abandon changes?", vbYesNo + vbQuestion, AppName) = vbYes Then
            Cancel = False
            If mbSaveChanged Then
             'frmServiceMtgs had already been changed prior to opening announcements form
            Else
                frmServiceMtg.ChangeMade = False
                frmServiceMtg.cmdApply.Enabled = False
            End If
        Else
            Cancel = True
        End If
    Else
        Cancel = False
    End If
    
    Exit Sub
ErrorTrap:
    Call EndProgram
    
End Sub


Private Sub mnuAttendants_Click()
On Error GoTo ErrorTrap
Dim NewText As String, TempDate As Date, bDone As Boolean
Dim col As New Collection, lSundayMtgDay As Long, lMidweekMtgDay As Long

    RemoveAllItemsFromCollection col
    bDone = False 'init

    If GlobalParms.GetValue("RotaTimesPerWeek", "NumVal") = 1 Then
        NewText = vbCrLf & _
                  "> Sound and Attendants: " & vbCrLf
                  
        'find the next Monday date...
        
        TempDate = GetDateOfGivenDay(mdteMeetingDate, vbMonday, False)
        TempDate = CDate(DateAdd("ww", 1, TempDate))
        Do Until bDone
            If IsCircuitOrDistrictAssemblyWeek(TempDate) Then
                NewText = NewText & vbTab & "Next week: " & vbTab & "ASSEMBLY" & vbCrLf
                TempDate = CDate(DateAdd("ww", 1, TempDate))
            Else
                NewText = NewText & vbTab & "Next week: " & vbCrLf
                bDone = True
            End If
        Loop
        
        'use this date to determine attendants:
        
        'Attendants
        Set col = AttendantsToday(TempDate, _
                                  True, _
                                  0, cmsAttendantAtt)
        
        If col.Count = 0 Then
            Set col = Nothing
            Exit Sub
        End If
        
        NewText = NewText & vbTab & "Attendants: " & BuildBroList(col) & vbCrLf
    
        'Roving
        Set col = AttendantsToday(TempDate, _
                                  True, _
                                  0, cmsMicrophonesAtt)
        
        NewText = NewText & vbTab & "Microphones: " & BuildBroList(col) & vbCrLf
    
        'Platform
        Set col = AttendantsToday(TempDate, _
                                  True, _
                                  0, cmsPlatformAtt)
        
        NewText = NewText & vbTab & "Platform: " & BuildBroList(col) & vbCrLf
    
        'Sound
        Set col = AttendantsToday(TempDate, _
                                  True, _
                                  0, cmsSoundAtt)
        
        NewText = NewText & vbTab & "Sound: " & BuildBroList(col) & vbCrLf
    Else
        'RotaTimesPerWeek > 1
        
        TempDate = GetDateOfGivenDay(mdteMeetingDate, _
                                     CLng(GlobalParms.GetValue("MidWeekMeetingDay", "NumVal")), _
                                     True)
            
        '>>>> Next mtg:
        
        Set col = AttendantsAfterDate(TempDate, _
                                  0, cmsAttendantAtt)
                                  
        If col.Count > 0 Then
        
            NewText = vbCrLf & _
                      "> Sound and Attendants for "
                      
            
            'use this date to determine attendants:
            
            'Attendants
            
            NewText = NewText & Format$(col.Item(1), "ddd mmm d") & vbCrLf & vbTab & "Attendants: " & BuildBroListWithDate(col) & vbCrLf
        
            'Roving
            Set col = AttendantsAfterDate(TempDate, _
                                       0, cmsMicrophonesAtt)
            
            NewText = NewText & vbTab & "Microphones: " & BuildBroListWithDate(col) & vbCrLf
        
            'Platform
            Set col = AttendantsAfterDate(TempDate, _
                                      0, cmsPlatformAtt)
            
            NewText = NewText & vbTab & "Platform: " & BuildBroListWithDate(col) & vbCrLf
        
            'Sound
            Set col = AttendantsAfterDate(TempDate, _
                                      0, cmsSoundAtt)
            
            NewText = NewText & vbTab & "Sound: " & BuildBroListWithDate(col) & vbCrLf
        End If
              
        '>>>> Next mtg:
        
        'Attendants
        Set col = AttendantsAfterDate(TempDate, _
                                  1, cmsAttendantAtt)
                                  
        If col.Count > 0 Then
        
            NewText = NewText & vbCrLf & _
                      "> Sound and Attendants for "
                  
        
            'use this date to determine attendants:
        

        
            NewText = NewText & Format$(col.Item(1), "ddd mmm d") & vbCrLf & vbTab & "Attendants: " & BuildBroListWithDate(col) & vbCrLf
        
            'Roving
            Set col = AttendantsAfterDate(TempDate, _
                                       1, cmsMicrophonesAtt)
            
            NewText = NewText & vbTab & "Microphones: " & BuildBroListWithDate(col) & vbCrLf
        
            'Platform
            Set col = AttendantsAfterDate(TempDate, _
                                      1, cmsPlatformAtt)
            
            NewText = NewText & vbTab & "Platform: " & BuildBroListWithDate(col) & vbCrLf
        
            'Sound
            Set col = AttendantsAfterDate(TempDate, _
                                      1, cmsSoundAtt)
            
            NewText = NewText & vbTab & "Sound: " & BuildBroListWithDate(col) & vbCrLf
        
        End If
              
        '>>>> Next mtg:
        
        'Attendants
        Set col = AttendantsAfterDate(TempDate, _
                                  2, cmsAttendantAtt)
                                  
        If col.Count > 0 Then
        
            NewText = NewText & vbCrLf & _
                      "> Sound and Attendants for "
                  
        
            'use this date to determine attendants:
        

        
            NewText = NewText & Format$(col.Item(1), "ddd mmm d") & vbCrLf & vbTab & "Attendants: " & BuildBroListWithDate(col) & vbCrLf
        
            'Roving
            Set col = AttendantsAfterDate(TempDate, _
                                       2, cmsMicrophonesAtt)
            
            NewText = NewText & vbTab & "Microphones: " & BuildBroListWithDate(col) & vbCrLf
        
            'Platform
            Set col = AttendantsAfterDate(TempDate, _
                                      2, cmsPlatformAtt)
            
            NewText = NewText & vbTab & "Platform: " & BuildBroListWithDate(col) & vbCrLf
        
            'Sound
            Set col = AttendantsAfterDate(TempDate, _
                                      2, cmsSoundAtt)
            
            NewText = NewText & vbTab & "Sound: " & BuildBroListWithDate(col) & vbCrLf
        
        End If
              
    End If
              
    If NewText <> "" Then
        txtAnnouncements = InsertSubstr(txtAnnouncements, _
                                       NewText & vbCrLf & vbCrLf, _
                                       txtAnnouncements.SelStart)
                                       
        SetTextBoxInsertPoint txtAnnouncements, -1, True
                                       
    Else
        MsgBox "No attendants on rota beyond this date", vbOKOnly + vbInformation, AppName
    End If
                                   
    Set col = Nothing
    

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub mnuCleaningThisWeek_Click()
Dim lSave As Long
On Error GoTo ErrorTrap
Dim NewText As String, TempDate As Date, bDone As Boolean

    NewText = vbCrLf & _
              "> Cleaning: " & vbTab & "This week: " & vbTab & GetGroupName(GetGroupCleaning(mdteMeetingDate)) & _
              vbCrLf
              
    TempDate = CDate(DateAdd("ww", 1, mdteMeetingDate))
    Do Until bDone
        If IsCircuitOrDistrictAssemblyWeek(TempDate) Then
            NewText = NewText & vbTab & vbTab & "Next week: " & vbTab & "Assembly" & vbCrLf
        Else
            NewText = NewText & vbTab & vbTab & "Next week: " & vbTab & GetGroupName(GetGroupCleaning(TempDate)) & _
                                 vbCrLf
            bDone = True
        End If
        TempDate = CDate(DateAdd("ww", 1, TempDate))
    Loop
              
    txtAnnouncements = InsertSubstr(txtAnnouncements, _
                                   NewText & vbCrLf & vbCrLf, _
                                   txtAnnouncements.SelStart)
                                   
    SetTextBoxInsertPoint txtAnnouncements, -1, True
                                   
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub
Private Function BuildBroList(col As Collection) As String

On Error GoTo ErrorTrap
Dim NewText As String, TempDate As Date, bDone As Boolean, i As Long

    
    For i = 1 To col.Count
        NewText = NewText & CongregationMember.NameWithOneFirstInitial(col.Item(i))
        If i < col.Count Then
            NewText = NewText & ", "
        End If
    Next i
              
    BuildBroList = NewText

    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Private Function BuildBroListWithDate(col As Collection) As String

On Error GoTo ErrorTrap
Dim NewText As String, TempDate As Date, bDone As Boolean, i As Long

    
    For i = 2 To col.Count
        NewText = NewText & CongregationMember.NameWithOneFirstInitial(col.Item(i))
        If i < col.Count Then
            NewText = NewText & ", "
        End If
    Next i
              
    BuildBroListWithDate = NewText

    Exit Function
ErrorTrap:
    Call EndProgram
End Function

Private Sub mnuDays_Click()
Dim lStartDay As Long, sDays As String, i As Long
On Error GoTo ErrorTrap

    lStartDay = GlobalParms.GetValue("MidWeekMeetingDay", "NumVal")
    lStartDay = GetNextInSequence(True, lStartDay, 7, 1, 1) 'get mtg day
    sDays = GetDayName(GetNextInSequence(False, lStartDay, 7, 1, 1)) & vbCrLf 'get day after mtg day
    
    For i = 2 To 7
        sDays = sDays & GetDayName(GetNextInSequence(False, lStartDay, 7, 1, 1)) & " " & vbCrLf
    Next i

    txtAnnouncements = InsertSubstr(txtAnnouncements, _
                                   sDays & vbCrLf, _
                                   txtAnnouncements.SelStart)

    SetTextBoxInsertPoint txtAnnouncements, -1, True

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub mnuReports_Click()

On Error GoTo ErrorTrap
Dim NewText As String, TheMonth As String

    If Day(Now) >= 1 And Day(Now) < GlobalParms.GetValue("DayOfMonthForReportToSociety", "NumVal") Then
        TheMonth = GetMonthName(Month(DateAdd("m", -1, date)))
    Else
        TheMonth = GetMonthName(Month(date))
    End If
    
    NewText = vbCrLf & _
              "> Please hand your " & TheMonth & _
              " field service reports in to your Bookstudy overseer."
              
    txtAnnouncements = InsertSubstr(txtAnnouncements, _
                                   NewText & vbCrLf & vbCrLf, _
                                   txtAnnouncements.SelStart)
                                   
    SetTextBoxInsertPoint txtAnnouncements, -1, True

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub txtAnnouncements_Change()

On Error GoTo ErrorTrap

    mbChangeMade = True
    cmdApply.Enabled = True

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

'Private Sub txtAnnouncements_GotFocus()
'
'On Error GoTo ErrorTrap
'
'    TextFieldGotFocus txtAnnouncements
'
'    Exit Sub
'ErrorTrap:
'    Call EndProgram
'End Sub


Public Property Get MeetingDate() As Date
    MeetingDate = mdteMeetingDate
End Property

Public Property Let MeetingDate(ByVal vNewValue As Date)
    mdteMeetingDate = vNewValue
End Property
Public Property Get Announcements() As String
    Announcements = msAnnouncements
End Property

Public Property Let Announcements(ByVal vNewValue As String)
    msAnnouncements = vNewValue
End Property
