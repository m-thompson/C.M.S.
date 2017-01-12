VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmAssignmentsForBro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.M.S. Assignments List"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "frmAssignmentsForBro.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGoTo 
      Caption         =   "Go to"
      Height          =   315
      Left            =   3510
      TabIndex        =   10
      Top             =   1110
      Width           =   2370
   End
   Begin VB.ComboBox cmbNames 
      Height          =   315
      Left            =   90
      TabIndex        =   8
      Top             =   375
      Width           =   3255
   End
   Begin VB.CommandButton cmdCLose 
      Caption         =   "Close"
      Height          =   315
      Left            =   4995
      TabIndex        =   7
      Top             =   7335
      Width           =   870
   End
   Begin ComctlLib.TreeView tv 
      Height          =   5565
      Left            =   105
      TabIndex        =   0
      Top             =   1725
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   9816
      _Version        =   327682
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.TextBox txtLastDate 
      Height          =   315
      Left            =   1605
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1110
      Width           =   978
   End
   Begin VB.TextBox txtFirstDate 
      Height          =   315
      Left            =   90
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1110
      Width           =   978
   End
   Begin VB.CommandButton cmdShowCalendar1 
      DownPicture     =   "frmAssignmentsForBro.frx":0442
      Height          =   315
      Left            =   1080
      Picture         =   "frmAssignmentsForBro.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1110
      Width           =   420
   End
   Begin VB.CommandButton cmdShowCalendar2 
      DownPicture     =   "frmAssignmentsForBro.frx":0CC6
      Height          =   315
      Left            =   2610
      Picture         =   "frmAssignmentsForBro.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1110
      Width           =   420
   End
   Begin VB.Label Label3 
      Caption         =   "Name"
      Height          =   255
      Left            =   105
      TabIndex        =   9
      Top             =   150
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Start Date"
      Height          =   255
      Left            =   90
      TabIndex        =   6
      Top             =   885
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "End Date"
      Height          =   255
      Left            =   1605
      TabIndex        =   5
      Top             =   885
      Width           =   1275
   End
   Begin VB.Menu mnuGoTo 
      Caption         =   "Go to..."
      Visible         =   0   'False
      Begin VB.Menu mnuGo 
         Caption         =   "Go to"
      End
   End
End
Attribute VB_Name = "frmAssignmentsForBro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlPersonID As Long
Dim msStartDate As String, msEndDate As String
Dim WithEvents frmCal As frmMiniCalendar
Attribute frmCal.VB_VarHelpID = -1
Dim bClickedStartDate As Boolean, mbIgnore As Boolean
Dim mvNodeLoc As Variant 'has currently selected node#
Dim mlModality As FormShowConstants


Private Sub cmbNames_Click()
Dim SkillLevel As Long, TheCheckBox As Control

On Error GoTo ErrorTrap
    
    If Me!cmbNames.ListIndex > -1 Then
        mlPersonID = Me!cmbNames.ItemData(Me!cmbNames.ListIndex)
        FillTreeview
    Else
        tv.Nodes.Clear
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub cmbNames_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap

    If KeyCode = 46 Then
        cmbNames.ListIndex = -1
    End If
    
Exit Sub
    
ErrorTrap:
    EndProgram
    
End Sub

Private Sub cmbNames_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorTrap
    
    AutoCompleteCombo Me!cmbNames, KeyAscii

Exit Sub
    
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdGoTo_Click()

On Error GoTo ErrorTrap

    mnuGo_Click

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub Form_Unload(Cancel As Integer)
    msStartDate = ""
    msEndDate = ""
    BringForwardMainMenuWhenItsTheLastFormOpen
End Sub

Private Sub frmCal_InsertDate(TheDate As String)
On Error GoTo ErrorTrap

    If bClickedStartDate Then
        txtFirstDate = TheDate
    Else
        txtLastDate = TheDate
    End If

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub cmdShowCalendar1_Click()

On Error GoTo ErrorTrap

    bClickedStartDate = True
    
    With frmCal
    .FormDate = txtFirstDate
    .SetPos = True
    .XPos = Me.Left + cmdShowCalendar1.Left + cmdShowCalendar1.Width
    .YPos = Me.Top + cmdShowCalendar1.Top + cmdShowCalendar1.Height
    .Show mlModality, Me
    End With

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdShowCalendar2_Click()
On Error GoTo ErrorTrap

    bClickedStartDate = False
    
    With frmCal
    .FormDate = txtLastDate
    .SetPos = True
    .XPos = Me.Left + cmdShowCalendar2.Left + cmdShowCalendar2.Width
    .YPos = Me.Top + cmdShowCalendar2.Top + cmdShowCalendar2.Height
    .Show mlModality, Me
    End With

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub


Private Sub mnuGo_Click()

On Error GoTo ErrorTrap

Dim arr() As String, rs As Recordset, bCannotOpenRota As Boolean

    On Error Resume Next
    Set mvNodeLoc = tv.SelectedItem
    If Err.number <> 0 Then
        Set mvNodeLoc = Nothing
        Exit Sub
    End If
'    On Error GoTo ErrorTrap

    If TypeOf mvNodeLoc Is Node Then
    
        If mvNodeLoc.Children = 0 Then
        
            Select Case mvNodeLoc.Parent.text
            Case "Attendant"
                If Not IsNumber(mvNodeLoc.Tag) Then
                    MsgBox "Cannot perform operation - invalid reference", vbOKOnly + vbExclamation, AppName
                    Exit Sub
                End If
            Case "Service Meeting", "Public Meeting", "Congregation Bible Study"
                If Not ValidDate(mvNodeLoc.Tag) Then
                    MsgBox "Cannot perform operation - invalid date", vbOKOnly + vbExclamation, AppName
                    Exit Sub
                End If
            Case "Student Talks"
                arr() = Split(mvNodeLoc.Tag, "||")
                If Not ValidDate(arr(0)) Then
                    MsgBox "Cannot perform operation - invalid date", vbOKOnly + vbExclamation, AppName
                    Exit Sub
                End If
                If Not IsNumber(arr(1), False, False, False) Then
                    MsgBox "Cannot perform operation - invalid school", vbOKOnly + vbExclamation, AppName
                    Exit Sub
                End If
            End Select
        
            Select Case mvNodeLoc.Parent.text
            Case "Attendant"
                If ValidDate(Right(mvNodeLoc.text, 10)) Then
                    Set rs = CMSDB.OpenRecordset("SELECT 1 FROM tblRota WHERE RotaDate = " & GetDateStringForSQLWhere(Right(mvNodeLoc.text, 10)), _
                                                dbOpenForwardOnly) 'check exists on current rota
                    If Not rs.BOF Then
                        If FormIsOpen("frmEditRota") Then
                            Unload frmEditRota
                        End If
'                        On Error Resume Next
                        rs.Close
                        Set rs = Nothing
'                        On Error GoTo ErrorTrap
                        frmEditRota.SeqNum = mvNodeLoc.Tag
                        frmEditRota.Show vbModal, Me
                    Else
                        ShowMessage "Cannot edit/display rota for selected date." & vbCrLf & _
                                    "Try selecting the appropriate rota date range...", 3000, Me, , , , 5
                        bCannotOpenRota = True
                    End If
'                    On Error Resume Next
                    rs.Close
                    Set rs = Nothing
'                    On Error GoTo ErrorTrap
                Else
                    ShowMessage "Cannot edit/display rota for selected date." & vbCrLf & _
                                "Try selecting the appropriate rota date range...", 3000, Me, , , , 5
                    bCannotOpenRota = True
                End If
                If bCannotOpenRota Then
                    bCannotOpenRota = False
                    If FormIsOpen("frmSoundAndPlatformRota") Then
                        Unload frmSoundAndPlatformRota
                    End If
                    frmSoundAndPlatformRota.Show vbModeless, frmMainMenu
                End If
            Case "Service Meeting"
                If FormIsOpen("frmServiceMtg") Then
                    Unload frmServiceMtg
                End If
                frmServiceMtg.MeetingDate = CDate(mvNodeLoc.Tag)
                frmServiceMtg.Show mlModality, Me
            Case "Student Talks"
                If FormIsOpen("frmTMSScheduling") Then
                    Unload frmTMSScheduling
                End If
                frmTMSScheduling.FormDate = CDate(arr(0))
                frmTMSScheduling.FormYear = year(arr(0))
                frmTMSScheduling.FormMonth = Month(arr(0))
                frmTMSScheduling.CurrentSchool = CLng(arr(1))
                frmTMSScheduling.Show mlModality, Me
            Case "Public Meeting"
                If FormIsOpen("frmPublicMtgSchedule") Then
                    Unload frmPublicMtgSchedule
                End If
                frmPublicMtgSchedule.FormDate = CDate(mvNodeLoc.Tag)
                frmPublicMtgSchedule.Show mlModality, Me
                
            Case "Congregation Bible Study"
                If FormIsOpen("frmCongBibleStudyRota") Then
                    Unload frmCongBibleStudyRota
                End If
                frmCongBibleStudyRota.FormDate = CDate(mvNodeLoc.Tag)
                frmCongBibleStudyRota.Show mlModality, Me
           
            End Select
            
        End If
        
    End If


    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub tv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim vDummy
On Error GoTo ErrorTrap
        
    On Error Resume Next
    Set mvNodeLoc = tv.HitTest(X, Y)
    Set vDummy = mvNodeLoc
    If Err.number <> 0 Then
        cmdGoTo.Enabled = False
        Exit Sub
    End If
    vDummy = mvNodeLoc.Children
    If Err.number <> 0 Then
        cmdGoTo.Enabled = False
        Exit Sub
    End If
    
    On Error GoTo ErrorTrap
    
    If Not IsEmpty(mvNodeLoc) Then
    
        If mvNodeLoc.Children = 0 Then
            mvNodeLoc.Selected = True
            
            Select Case mvNodeLoc.Parent.text
            Case "Attendant"
                mnuGo.Caption = "Go to Attendants"
                mnuGo.Enabled = AccessAllowed("frmMainMenu", "cmdOpenSPAMRota")
            Case "Service Meeting"
                mnuGo.Caption = "Go to Service Meeting"
                mnuGo.Enabled = AccessAllowed("frmMainMenu", "cmdServiceMtg")
            Case "Student Talks"
                mnuGo.Caption = "Go to Student Talks"
                mnuGo.Enabled = AccessAllowed("frmMainMenu", "cmdOpenTMS")
            Case "Public Meeting"
                mnuGo.Caption = "Go to Public Meeting"
                mnuGo.Enabled = AccessAllowed("frmMainMenu", "cmdPublicMeeting")
            Case "Congregation Bible Study"
                mnuGo.Caption = "Go to Congregation Bible Study"
                mnuGo.Enabled = AccessAllowed("frmMainMenu", "cmdPublicMeeting")
            End Select
            
            cmdGoTo.Enabled = mnuGo.Enabled
            cmdGoTo.Caption = mnuGo.Caption
            
            If Button = vbRightButton Then
                Me.PopupMenu mnuGoTo
            End If
            
        Else
            cmdGoTo.Enabled = False
            cmdGoTo.Caption = "Go to"
            
        End If
            
    End If
    
    Set mvNodeLoc = Nothing
        
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub tv_NodeClick(ByVal Node As ComctlLib.Node)
On Error GoTo ErrorTrap
    Set mvNodeLoc = Node
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub txtFirstDate_Change()
On Error GoTo ErrorTrap

    If mbIgnore Then Exit Sub
    
    FillTreeview

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub txtFirstDate_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorTrap

    KeyPressValid KeyAscii, cmsDates, True

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub


Private Sub txtLastDate_Change()
On Error GoTo ErrorTrap

    If mbIgnore Then Exit Sub

    FillTreeview

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub txtLastDate_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorTrap

    KeyPressValid KeyAscii, cmsDates, True

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Public Property Let PersonID(ByVal vNewValue As Long)
    mlPersonID = vNewValue
End Property

Public Property Let StartDate(ByVal vNewValue As String)
    msStartDate = vNewValue
End Property
Public Property Let EndDate(ByVal vNewValue As String)
    msEndDate = vNewValue
End Property

Private Sub Form_Load()

On Error GoTo ErrorTrap

    Set frmCal = New frmMiniCalendar
    
    HandleListBox.PopulateListBox Me!cmbNames, _
        "SELECT DISTINCTROW tblNameAddress.ID, " & _
        "tblNameAddress.FirstName & ' ' & tblNameAddress.MiddleName, " & _
        "tblNameAddress.LastName " & _
        "FROM tblNameAddress " & _
        "WHERE Active = TRUE " & _
        " ORDER BY tblNameAddress.LastName, tblNameAddress.FirstName" _
        , CMSDB, 0, ", ", True, 2, 1
        
    If mlPersonID > 0 Then
        HandleListBox.SelectItem cmbNames, mlPersonID
    End If
    
    If Not ValidDate(msStartDate) Or Not ValidDate(msEndDate) Then
        msStartDate = DateAdd("m", -6, Format(Now, "mm/dd/yyyy"))
        msEndDate = DateAdd("m", 6, Format(Now, "mm/dd/yyyy"))
    End If
    
    txtFirstDate = msStartDate
    txtLastDate = msEndDate
    
    cmdGoTo.Enabled = False
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub FillTreeview()

On Error GoTo ErrorTrap

Dim AttCol As New Collection, nd As Node, i As Long
Dim TMSCol As New Collection
Dim SMtgCol As New Collection
Dim PubMtgCol As New Collection
Dim CongBibStdyCol As New Collection
Dim arr() As String, str1 As String, str2 As String

    tv.Nodes.Clear
    
    cmdGoTo.Enabled = False
    cmdGoTo.Caption = "Go to"
    
    If ValidDate(txtFirstDate) And ValidDate(txtLastDate) And cmbNames.ListIndex > -1 Then
    Else
        Exit Sub
    End If
    
    If CDate(txtFirstDate) > CDate(txtLastDate) Then
        Exit Sub
    End If
    
    '
    'SPAM
    '
    Set AttCol = GetDatesOnSPAM(mlPersonID, txtFirstDate, txtLastDate)
    
    If AttCol.Count > 0 Then
        
        Set nd = tv.Nodes.Add(, , "Root1", "Attendant")
        
        For i = 1 To AttCol.Count
            arr() = Split(AttCol.Item(i), "||")
            Set nd = tv.Nodes.Add("Root1", tvwChild, , arr(0))
            nd.Tag = arr(1)
        Next i
        
    End If

    '
    'TMS
    '
    Set TMSCol = GetTMSDates(mlPersonID, txtFirstDate, txtLastDate)
    
    If TMSCol.Count > 0 Then
        
        Set nd = tv.Nodes.Add(, , "Root2", "Student Talks")
        
        For i = 1 To TMSCol.Count
            arr() = Split(TMSCol.Item(i), "||")
            Set nd = tv.Nodes.Add("Root2", tvwChild, , arr(0))
            nd.Tag = arr(1) & "||" & arr(2)
        Next i
        
    End If

    '
    'Service Mtg
    '
    Set SMtgCol = GetSrvMtgDates(mlPersonID, txtFirstDate, txtLastDate)
    
    If SMtgCol.Count > 0 Then
        
        Set nd = tv.Nodes.Add(, , "Root3", "Service Meeting")
        
        For i = 1 To SMtgCol.Count
            arr() = Split(SMtgCol.Item(i), "||")
            Set nd = tv.Nodes.Add("Root3", tvwChild, , arr(0))
            nd.Tag = arr(1)
        Next i
        
    End If
    
    '
    'Public Mtg
    '
    Set PubMtgCol = GetPubMtgDates(mlPersonID, txtFirstDate, txtLastDate)
    
    If PubMtgCol.Count > 0 Then
                
        Set nd = tv.Nodes.Add(, , "Root4", "Public Meeting")
        
        For i = 1 To PubMtgCol.Count
            arr() = Split(PubMtgCol.Item(i), "||")
            Set nd = tv.Nodes.Add("Root4", tvwChild, , arr(0))
            nd.Tag = arr(1)
        Next i
        
    End If
    
    '
    'Cong Bible Study
    '
    Set CongBibStdyCol = GetCongBibleStudyDates(mlPersonID, txtFirstDate, txtLastDate)
    
    If CongBibStdyCol.Count > 0 Then
                
        Set nd = tv.Nodes.Add(, , "Root5", "Congregation Bible Study")
        
        For i = 1 To CongBibStdyCol.Count
            arr() = Split(CongBibStdyCol.Item(i), "||")
            Set nd = tv.Nodes.Add("Root5", tvwChild, , arr(0))
            nd.Tag = arr(1)
        Next i
        
    End If
    
    Set AttCol = Nothing
    Set TMSCol = Nothing
    Set SMtgCol = Nothing
    Set PubMtgCol = Nothing
    Set CongBibStdyCol = Nothing

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Function GetSPAMDesc(TheString As String) As String

On Error GoTo ErrorTrap

    Select Case LCase$(Left(TheString, 3))
    Case "att"
        GetSPAMDesc = "Attendant"
    Case "rov"
        GetSPAMDesc = "Roving Microphone"
    Case "pla"
        GetSPAMDesc = "Platform"
    Case "sou"
        GetSPAMDesc = "Sound"
    End Select
    
    Exit Function
ErrorTrap:
    EndProgram

End Function

Private Function GetDatesOnSPAM(ByVal ThePerson As Long, _
                                ByVal StartDate_UK As String, _
                                ByVal EndDate_UK As String) As Collection
                                 
On Error GoTo ErrorTrap
Dim rst As Recordset, TempCol As New Collection, str As String
Dim var As Variant, i As Long, fld As DAO.Field
Dim str2 As String

    RemoveAllItemsFromCollection TempCol
 
 'return collection with all details
    
    Set rst = GetSPAMRota(StartDate_UK, EndDate_UK)
 
    With rst
    
    Do Until .EOF Or .BOF 'for each date
    
        For Each fld In rst.Fields
            
            Select Case LCase$(Left(fld.Name, 3))
            Case "sou", "pla", "att", "rov"
                If .Fields(fld.Name) = ThePerson Then
                    TempCol.Add GetSPAMDesc(fld.Name) & ": " & !RotaDate & "||" & !SeqNum
                End If
            End Select
        
        Next
        
        .MoveNext
        
    Loop 'rotadates
        
    End With
     
    Set GetDatesOnSPAM = TempCol
    
    rst.Close
    Set rst = Nothing
        
    Exit Function
ErrorTrap:
    EndProgram

End Function

Private Function GetTMSDates(ByVal ThePerson As Long, _
                                ByVal StartDate_UK As String, _
                                ByVal EndDate_UK As String) As Collection
                                 
On Error GoTo ErrorTrap
Dim rst As Recordset, TempCol As New Collection, str As String
Dim var As Variant, i As Long, fld As DAO.Field
Dim str2 As String

    RemoveAllItemsFromCollection TempCol
 
 'return collection with all details
    
    str = "SELECT AssignmentDate, TalkNo, SchoolNo, PersonID, Assistant1ID " & _
          "FROM tblTMSSchedule " & _
          "WHERE AssignmentDate BETWEEN " & GetDateStringForSQLWhere(StartDate_UK) & _
          "                         AND " & GetDateStringForSQLWhere(EndDate_UK) & _
          " AND (PersonID = " & ThePerson & _
          " OR Assistant1ID = " & ThePerson & _
          ") ORDER BY 1 "

    Set rst = CMSDB.OpenRecordset(str, dbOpenDynaset)
 
    With rst
    
    Do Until .EOF Or .BOF 'for each date
    
        If !PersonID = ThePerson Then
            TempCol.Add TheTMS.GetTMSTalkDescription(!TalkNo, FormatDateAsUKDateString(!AssignmentDate)) & ": " & !AssignmentDate & "||" & !AssignmentDate & "||" & !SchoolNo
        Else
            TempCol.Add TheTMS.GetTMSTalkDescription(!TalkNo, FormatDateAsUKDateString(!AssignmentDate)) & " Assistant: " & !AssignmentDate & "||" & !AssignmentDate & "||" & !SchoolNo
        End If
        
        .MoveNext
        
    Loop
        
    End With
     
    Set GetTMSDates = TempCol
    
    rst.Close
    Set rst = Nothing
        
    Exit Function
ErrorTrap:
    EndProgram

End Function
Private Function GetSrvMtgDates(ByVal ThePerson As Long, _
                                ByVal StartDate_UK As String, _
                                ByVal EndDate_UK As String) As Collection
                                 
On Error GoTo ErrorTrap
Dim rst As Recordset, TempCol As New Collection, str As String
Dim var As Variant, i As Long, fld As DAO.Field
Dim str2 As String

    RemoveAllItemsFromCollection TempCol
 
 'return collection with all details
    
    str = "SELECT SeqNum, MeetingDate, ItemTypeID, ItemLength, PersonID " & _
          "FROM tblServiceMtgs " & _
          "WHERE MeetingDate BETWEEN " & GetDateStringForSQLWhere(StartDate_UK) & _
          "                         AND " & GetDateStringForSQLWhere(EndDate_UK) & _
          " AND PersonID = " & ThePerson & _
          " ORDER BY 2 "

    Set rst = CMSDB.OpenRecordset(str, dbOpenDynaset)
 
    With rst
    
    Do Until .EOF Or .BOF 'for each date
    
        Select Case !ItemTypeID
        Case 0, 1
            TempCol.Add "Item (" & !ItemLength & " mins): " & !MeetingDate & "||" & !MeetingDate
        Case 2
            TempCol.Add "Concluding Prayer: " & !MeetingDate & "||" & !MeetingDate
        End Select
        
        .MoveNext
        
    Loop
        
    End With
     
    Set GetSrvMtgDates = TempCol
    
    rst.Close
    Set rst = Nothing
        
    Exit Function
ErrorTrap:
    EndProgram

End Function

Private Function GetPubMtgDates(ByVal ThePerson As Long, _
                                ByVal StartDate_UK As String, _
                                ByVal EndDate_UK As String) As Collection
                                 
On Error GoTo ErrorTrap
Dim rst As Recordset, TempCol As New Collection, str As String
Dim var As Variant, i As Long, fld As DAO.Field
Dim str2 As String

    RemoveAllItemsFromCollection TempCol
 
 'return collection with all details
    
    str = "SELECT MeetingDate, SpeakerID, SpeakerID2, ChairmanID, WTReaderID, CongNoWhereMtgIs " & _
          "FROM tblPublicMtgSchedule " & _
          "WHERE MeetingDate BETWEEN " & GetDateStringForSQLWhere(StartDate_UK) & _
          "                         AND " & GetDateStringForSQLWhere(EndDate_UK) & _
          " AND (SpeakerID = " & ThePerson & _
          " OR SpeakerID2 = " & ThePerson & _
          " OR ChairmanID = " & ThePerson & _
          " OR WTReaderID = " & ThePerson & _
          ") ORDER BY 1 "

    Set rst = CMSDB.OpenRecordset(str, dbOpenDynaset)
 
    With rst
    
    Do Until .EOF Or .BOF 'for each date
    
        Select Case True
        Case !SpeakerID = ThePerson
            TempCol.Add "Public Talk at " & GetCongregationName(!CongNoWhereMtgIs) & ": " & !MeetingDate & "||" & !MeetingDate
        Case !SpeakerID2 = ThePerson
            TempCol.Add "Public Talk at " & GetCongregationName(!CongNoWhereMtgIs) & ": " & !MeetingDate & "||" & !MeetingDate
        Case !ChairmanID = ThePerson
            TempCol.Add "Chairman" & ": " & !MeetingDate & "||" & !MeetingDate
        Case !WTReaderID = ThePerson
            TempCol.Add "Watchtower Reader" & ": " & !MeetingDate & "||" & !MeetingDate
        End Select
        
        .MoveNext
        
    Loop
        
    End With
     
    Set GetPubMtgDates = TempCol
    
    rst.Close
    Set rst = Nothing
        
    Exit Function
ErrorTrap:
    EndProgram

End Function
Private Function GetCongBibleStudyDates(ByVal ThePerson As Long, _
                                        ByVal StartDate_UK As String, _
                                        ByVal EndDate_UK As String) As Collection
                                 
On Error GoTo ErrorTrap
Dim rst As Recordset, TempCol As New Collection, str As String
Dim var As Variant, i As Long, fld As DAO.Field
Dim str2 As String

    RemoveAllItemsFromCollection TempCol
 
 'return collection with all details
    
    str = "SELECT MeetingDate, ConductorID, ReaderID, PrayerID " & _
          "FROM tblCongBibleStudyRota " & _
          "WHERE MeetingDate BETWEEN " & GetDateStringForSQLWhere(StartDate_UK) & _
          "                         AND " & GetDateStringForSQLWhere(EndDate_UK) & _
          " AND (ConductorID = " & ThePerson & _
          " OR ReaderID = " & ThePerson & _
          " OR PrayerID = " & ThePerson & _
          ") ORDER BY 1 "

    Set rst = CMSDB.OpenRecordset(str, dbOpenDynaset)
 
    With rst
    
    Do Until .EOF Or .BOF 'for each date
    
        Select Case True
        Case !ConductorID = ThePerson
            TempCol.Add "Conductor: " & !MeetingDate & "||" & !MeetingDate
        Case !ReaderID = ThePerson
            TempCol.Add "Reader: " & !MeetingDate & "||" & !MeetingDate
        Case !PrayerID = ThePerson
            TempCol.Add "Prayer: " & !MeetingDate & "||" & !MeetingDate
        End Select
        
        .MoveNext
        
    Loop
        
    End With
     
    Set GetCongBibleStudyDates = TempCol
    
    rst.Close
    Set rst = Nothing
        
    Exit Function
ErrorTrap:
    EndProgram

End Function
Public Property Get Modality() As FormShowConstants
    Modality = mlModality
End Property

Public Property Let Modality(ByVal vNewValue As FormShowConstants)
    mlModality = vNewValue
End Property
