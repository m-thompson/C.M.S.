VERSION 5.00
Begin VB.Form frmAttendantsDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.M.S. Attendants Details"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmAttendantsDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   300
      Left            =   3930
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   375
      Width           =   675
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   300
      Left            =   3255
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   375
      Width           =   675
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   3810
      TabIndex        =   8
      Top             =   2850
      Width           =   795
   End
   Begin VB.Frame frmAssignments 
      Height          =   1785
      Left            =   195
      TabIndex        =   4
      Top             =   930
      Width           =   4410
      Begin VB.CheckBox chkPlatSundayOnly 
         Caption         =   "Sunday Only"
         Height          =   255
         Left            =   1470
         TabIndex        =   17
         Top             =   1350
         Width           =   1365
      End
      Begin VB.CheckBox chkPlatMidWeekOnly 
         Caption         =   "Midweek Only"
         Height          =   255
         Left            =   2925
         TabIndex        =   16
         Top             =   1350
         Width           =   1305
      End
      Begin VB.CheckBox chkRovSundayOnly 
         Caption         =   "Sunday Only"
         Height          =   255
         Left            =   1470
         TabIndex        =   15
         Top             =   960
         Width           =   1365
      End
      Begin VB.CheckBox chkRovMidWeekOnly 
         Caption         =   "Midweek Only"
         Height          =   255
         Left            =   2925
         TabIndex        =   14
         Top             =   960
         Width           =   1305
      End
      Begin VB.CheckBox chkSoundSundayOnly 
         Caption         =   "Sunday Only"
         Height          =   255
         Left            =   1470
         TabIndex        =   13
         Top             =   585
         Width           =   1365
      End
      Begin VB.CheckBox chkSoundMidWeekOnly 
         Caption         =   "Midweek Only"
         Height          =   255
         Left            =   2925
         TabIndex        =   12
         Top             =   585
         Width           =   1305
      End
      Begin VB.CheckBox chkAttSundayOnly 
         Caption         =   "Sunday Only"
         Height          =   255
         Left            =   1470
         TabIndex        =   11
         Top             =   195
         Width           =   1365
      End
      Begin VB.CheckBox chkAttMidWeekOnly 
         Caption         =   "Midweek Only"
         Height          =   255
         Left            =   2925
         TabIndex        =   10
         Top             =   195
         Width           =   1305
      End
      Begin VB.CheckBox chkPlatform 
         Caption         =   "Platform"
         Height          =   255
         Left            =   195
         TabIndex        =   9
         Top             =   1350
         Width           =   1260
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "Sound"
         Height          =   255
         Left            =   195
         TabIndex        =   7
         Top             =   585
         Width           =   1260
      End
      Begin VB.CheckBox chkRoving 
         Caption         =   "Roving Mics"
         Height          =   255
         Left            =   195
         TabIndex        =   6
         Top             =   960
         Width           =   1260
      End
      Begin VB.CheckBox chkAttendant 
         Caption         =   "Attendant"
         Height          =   255
         Left            =   195
         TabIndex        =   5
         Top             =   195
         Width           =   1260
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000005&
         X1              =   210
         X2              =   4215
         Y1              =   1290
         Y2              =   1290
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000005&
         X1              =   210
         X2              =   4215
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000003&
         X1              =   210
         X2              =   4215
         Y1              =   1275
         Y2              =   1275
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         X1              =   210
         X2              =   4215
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   210
         X2              =   4215
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   210
         X2              =   4215
         Y1              =   495
         Y2              =   495
      End
   End
   Begin VB.ComboBox cmbBrothers 
      Height          =   315
      Left            =   195
      TabIndex        =   2
      Text            =   "cmbBrothers"
      Top             =   375
      Width           =   3060
   End
   Begin VB.ComboBox cmbCongregation 
      Enabled         =   0   'False
      Height          =   315
      Left            =   270
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4155
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Brothers"
      Height          =   255
      Left            =   195
      TabIndex        =   3
      Top             =   135
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Congregation"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frmAttendantsDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TheSelectedBrother As Long, TheSelectedCong As Integer, InitControl As Boolean
Dim mbSwitchOffMsg As Boolean
Dim WithEvents frmNewBro As frmAddPerson
Attribute frmNewBro.VB_VarHelpID = -1

Private Sub chkAttendant_Click()
On Error GoTo ErrorTrap

    If InitControl Then Exit Sub

    If chkAttendant.value = vbChecked Then
        chkAttSundayOnly.Enabled = True
        chkAttMidWeekOnly.Enabled = True
        CongregationMember.AddPersonToRole CLng(TheSelectedCong), 6, 11, 60, CLng(TheSelectedBrother)
        ShowMessage "Change saved", 300, Me, mbSwitchOffMsg
    Else
        CongregationMember.DeletePersonFromRole CLng(TheSelectedCong), 6, 11, 60, CLng(TheSelectedBrother)
        ShowMessage "Change saved", 300, Me, mbSwitchOffMsg
        InitControl = True
        chkAttSundayOnly.value = vbUnchecked
        chkAttMidWeekOnly.value = vbUnchecked
        InitControl = False
        chkAttSundayOnly.Enabled = False
        chkAttMidWeekOnly.Enabled = False
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram
    
    
End Sub

Private Sub chkAttMidWeekOnly_Click()
On Error GoTo ErrorTrap

    If InitControl Then Exit Sub

    CongregationMember.SetMidweekOnly CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 11, 60, _
                                    IIf(chkAttMidWeekOnly = vbChecked, True, False)
                                    
    ShowMessage "Change saved", 300, Me, mbSwitchOffMsg

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub chkAttSundayOnly_Click()
On Error GoTo ErrorTrap

    If InitControl Then Exit Sub

    CongregationMember.SetSundayOnly CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 11, 60, _
                                    IIf(chkAttSundayOnly = vbChecked, True, False)

    ShowMessage "Change saved", 300, Me, mbSwitchOffMsg
    
    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub chkPlatform_Click()
On Error GoTo ErrorTrap

    If InitControl Then Exit Sub
    
    If chkPlatform.value = vbChecked Then
        chkPlatSundayOnly.Enabled = True
        chkPlatMidWeekOnly.Enabled = True
        CongregationMember.AddPersonToRole CLng(TheSelectedCong), 6, 10, 58, CLng(TheSelectedBrother)
        ShowMessage "Change saved", 300, Me, mbSwitchOffMsg
    Else
        CongregationMember.DeletePersonFromRole CLng(TheSelectedCong), 6, 10, 58, CLng(TheSelectedBrother)
        ShowMessage "Change saved", 300, Me, mbSwitchOffMsg
        InitControl = True
        chkPlatSundayOnly.value = vbUnchecked
        chkPlatMidWeekOnly.value = vbUnchecked
        InitControl = False
        chkPlatSundayOnly.Enabled = False
        chkPlatMidWeekOnly.Enabled = False
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub chkPlatMidWeekOnly_Click()
On Error GoTo ErrorTrap

    If InitControl Then Exit Sub

    CongregationMember.SetMidweekOnly CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 10, 58, _
                                    IIf(chkPlatMidWeekOnly = vbChecked, True, False)
                                    
    ShowMessage "Change saved", 300, Me, mbSwitchOffMsg

    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub chkPlatSundayOnly_Click()
On Error GoTo ErrorTrap

    If InitControl Then Exit Sub

    CongregationMember.SetSundayOnly CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 10, 58, _
                                    IIf(chkPlatSundayOnly = vbChecked, True, False)
    ShowMessage "Change saved", 300, Me, mbSwitchOffMsg
    
    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub chkRoving_Click()
On Error GoTo ErrorTrap

    If InitControl Then Exit Sub

    If chkRoving.value = vbChecked Then
        chkRovSundayOnly.Enabled = True
        chkRovMidWeekOnly.Enabled = True
        CongregationMember.AddPersonToRole CLng(TheSelectedCong), 6, 10, 59, CLng(TheSelectedBrother)
        ShowMessage "Change saved", 300, Me, mbSwitchOffMsg
    Else
        CongregationMember.DeletePersonFromRole CLng(TheSelectedCong), 6, 10, 59, CLng(TheSelectedBrother)
        ShowMessage "Change saved", 300, Me, mbSwitchOffMsg
        InitControl = True
        chkRovSundayOnly.value = vbUnchecked
        chkRovMidWeekOnly.value = vbUnchecked
        InitControl = False
        chkRovSundayOnly.Enabled = False
        chkRovMidWeekOnly.Enabled = False
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub chkRovMidWeekOnly_Click()
On Error GoTo ErrorTrap

    If InitControl Then Exit Sub

    CongregationMember.SetMidweekOnly CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 10, 59, _
                                    IIf(chkRovMidWeekOnly = vbChecked, True, False)

    ShowMessage "Change saved", 300, Me, mbSwitchOffMsg
    
    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub chkRovSundayOnly_Click()
On Error GoTo ErrorTrap

    If InitControl Then Exit Sub

    CongregationMember.SetSundayOnly CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 10, 59, _
                                    IIf(chkRovSundayOnly = vbChecked, True, False)

    ShowMessage "Change saved", 300, Me, mbSwitchOffMsg

    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub chkSound_Click()
On Error GoTo ErrorTrap

    If InitControl Then Exit Sub

    If chkSound.value = vbChecked Then
        chkSoundSundayOnly.Enabled = True
        chkSoundMidWeekOnly.Enabled = True
        CongregationMember.AddPersonToRole CLng(TheSelectedCong), 6, 10, 57, CLng(TheSelectedBrother)
        ShowMessage "Change saved", 300, Me, mbSwitchOffMsg
    Else
        CongregationMember.DeletePersonFromRole CLng(TheSelectedCong), 6, 10, 57, CLng(TheSelectedBrother)
        ShowMessage "Change saved", 300, Me, mbSwitchOffMsg
        InitControl = True
        chkSoundSundayOnly.value = vbUnchecked
        chkSoundMidWeekOnly.value = vbUnchecked
        InitControl = False
        chkSoundSundayOnly.Enabled = False
        chkSoundMidWeekOnly.Enabled = False
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram
    


End Sub

Private Sub chkSoundMidWeekOnly_Click()
On Error GoTo ErrorTrap

    If InitControl Then Exit Sub

    CongregationMember.SetMidweekOnly CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 10, 57, _
                                    IIf(chkSoundMidWeekOnly = vbChecked, True, False)

    ShowMessage "Change saved", 300, Me, mbSwitchOffMsg
    
    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub chkSoundSundayOnly_Click()
On Error GoTo ErrorTrap

    If InitControl Then Exit Sub

    CongregationMember.SetSundayOnly CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 10, 57, _
                                    IIf(chkSoundSundayOnly = vbChecked, True, False)

    ShowMessage "Change saved", 300, Me, mbSwitchOffMsg
    
    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub cmbBrothers_Click()
Dim SkillLevel As Long, TheChkBox As Control

On Error GoTo ErrorTrap
    
    InitControl = True
    
    If Me!cmbBrothers.ListIndex > -1 And Me!cmbCongregation.ListIndex > -1 Then
        TheSelectedBrother = CInt(Me!cmbBrothers.ItemData(Me!cmbBrothers.ListIndex))
        TheSelectedCong = CInt(Me!cmbCongregation.ItemData(Me!cmbCongregation.ListIndex))
        
        Me!frmAssignments.Enabled = True
        
        EnableChks
        
        '
        'Now set/unset checkboxes depending on which assignments bro does
        '
        If CongregationMember.IsAttendant(CInt(TheSelectedBrother), TheSelectedCong) Then
            Me!chkAttendant.value = vbChecked
            If CongregationMember.IsSundayOnly(CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 11, 60) Then
                chkAttSundayOnly.value = vbChecked
            Else
                chkAttSundayOnly.value = vbUnchecked
            End If
            If CongregationMember.IsMidweekOnly(CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 11, 60) Then
                chkAttMidWeekOnly.value = vbChecked
            Else
                chkAttMidWeekOnly.value = vbUnchecked
            End If
        Else
            Me!chkAttendant.value = vbUnchecked
        End If
        
        If CongregationMember.IsRovingMic(CLng(TheSelectedBrother), CLng(TheSelectedCong)) Then
            If CongregationMember.IsSundayOnly(CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 10, 59) Then
                chkRovSundayOnly.value = vbChecked
            Else
                chkRovSundayOnly.value = vbUnchecked
            End If
            If CongregationMember.IsMidweekOnly(CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 10, 59) Then
                chkRovMidWeekOnly.value = vbChecked
            Else
                chkRovMidWeekOnly.value = vbUnchecked
            End If
            Me!chkRoving.value = vbChecked
        Else
            Me!chkRoving.value = vbUnchecked
        End If
        
        If CongregationMember.IsSound(CLng(TheSelectedBrother), CLng(TheSelectedCong)) Then
            Me!chkSound.value = vbChecked
            If CongregationMember.IsSundayOnly(CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 10, 57) Then
                chkSoundSundayOnly.value = vbChecked
            Else
                chkSoundSundayOnly.value = vbUnchecked
            End If
            If CongregationMember.IsMidweekOnly(CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 10, 57) Then
                chkSoundMidWeekOnly.value = vbChecked
            Else
                chkSoundMidWeekOnly.value = vbUnchecked
            End If
        Else
            Me!chkSound.value = vbUnchecked
        End If
        
        If CongregationMember.IsPlatform(CLng(TheSelectedBrother), CLng(TheSelectedCong)) Then
            If CongregationMember.IsSundayOnly(CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 10, 58) Then
                chkPlatSundayOnly.value = vbChecked
            Else
                chkPlatSundayOnly.value = vbUnchecked
            End If
            If CongregationMember.IsMidweekOnly(CLng(TheSelectedBrother), CLng(TheSelectedCong), 6, 10, 58) Then
                chkPlatMidWeekOnly.value = vbChecked
            Else
                chkPlatMidWeekOnly.value = vbUnchecked
            End If
            Me!chkPlatform.value = vbChecked
        Else
            Me!chkPlatform.value = vbUnchecked
        End If
        
        InitControl = False
        mbSwitchOffMsg = True
        chkAttendant_Click
        chkPlatform_Click
        chkSound_Click
        chkRoving_Click
        mbSwitchOffMsg = False
    
    Else
        chkAttendant.value = vbUnchecked
        chkPlatform.value = vbUnchecked
        chkSound.value = vbUnchecked
        chkRoving.value = vbUnchecked
        DisableChks
    End If
    
    InitControl = False
    
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub cmbBrothers_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap

    If KeyCode = 46 Then
        cmbBrothers.ListIndex = -1
    End If
    
Exit Sub
    
ErrorTrap:
    EndProgram

End Sub

Private Sub cmbBrothers_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorTrap
    
    AutoCompleteCombo Me!cmbBrothers, KeyAscii

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

Private Sub cmdEdit_Click()
On Error GoTo ErrorTrap
    
    If cmbBrothers.ListIndex > -1 Then
    
        Set frmNewBro = New frmAddPerson
        
        With frmNewBro
        .PersonID = cmbBrothers.ItemData(cmbBrothers.ListIndex)
        .Show vbModal, Me
        End With
                
    Else
        ShowMessage "Please select a brother to edit", 1200, Me
        Exit Sub
    End If

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdNew_Click()
On Error GoTo ErrorTrap

    Set frmNewBro = New frmAddPerson
    
    With frmNewBro
    .PersonID = 0
    .Show vbModal, Me
    End With
        

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub Form_Load()
On Error GoTo ErrorTrap
    
    'Populate cmbCongrgation
    
    If CongregationIsSetUp Then
        HandleListBox.PopulateListBox Me!cmbCongregation, _
                                      "SELECT CongName, CongNo FROM tblCong", _
                                      CMSDB, 1, "", False, 0
                                      
        HandleListBox.SelectItem Me!cmbCongregation, _
                                 CLng(GlobalParms.GetValue("DefaultCong", "NumVal"))
    End If
                             
    Me!cmbCongregation.Enabled = False
    
    'Populate cmbBrothers
    HandleListBox.PopulateListBox Me!cmbBrothers, _
        "SELECT tblNameAddress.ID, " & _
        "tblNameAddress.FirstName & ' ' & tblNameAddress.MiddleName, " & _
        "tblNameAddress.LastName " & _
        "FROM tblNameAddress " & _
        "WHERE Active = TRUE " & _
        "AND GenderMF = 'M'" & _
        " ORDER BY tblNameAddress.LastName, tblNameAddress.FirstName" _
        , CMSDB, 0, ", ", True, 2, 1
        
    If TheSelectedBrother > 0 Then
        HandleListBox.SelectItem cmbBrothers, TheSelectedBrother
        If HandleListBox.ErrorCode = CMSItemNotInList Then
            DisableChks
        End If
    Else
        DisableChks
    End If
        
    
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DisableChks()
On Error GoTo ErrorTrap
Dim TMSFormControl As Control
    
    For Each TMSFormControl In Me.Controls
        If (TypeOf TMSFormControl Is CheckBox) Then
            TMSFormControl.Enabled = False
        End If
    Next
    

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub EnableChks()
On Error GoTo ErrorTrap
Dim TMSFormControl As Control
    
    For Each TMSFormControl In Me.Controls
        If TypeOf TMSFormControl Is CheckBox Then
            TMSFormControl.Enabled = True
        End If
    Next

    Exit Sub
ErrorTrap:
    EndProgram

End Sub



Public Property Get FormPersonID() As Long
    FormPersonID = TheSelectedBrother
End Property

Public Property Let FormPersonID(ByVal vNewValue As Long)
    TheSelectedBrother = CLng(vNewValue)
End Property

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap

    On Error Resume Next
    Set frmNewBro = Nothing

    BringForwardMainMenuWhenItsTheLastFormOpen
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub frmNewBro_PersonUpdated(ThePerson As Long)
On Error GoTo ErrorTrap

    HandleListBox.Requery cmbBrothers, False
    HandleListBox.SelectItem cmbBrothers, ThePerson
        
    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub
