VERSION 5.00
Begin VB.Form frmAddNewCMSUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  C.M.S.  User Edit"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmAddNewCMSUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDefaultDates 
      Height          =   316
      Left            =   1170
      Picture         =   "frmAddNewCMSUser.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Set to today"
      Top             =   1753
      Width           =   345
   End
   Begin VB.TextBox txtActiveToDate 
      ForeColor       =   &H00000000&
      Height          =   316
      Left            =   150
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2428
      Width           =   978
   End
   Begin VB.TextBox txtActiveFromDate 
      ForeColor       =   &H00000000&
      Height          =   316
      Left            =   150
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1753
      Width           =   978
   End
   Begin VB.TextBox txtEmailAddress 
      Height          =   300
      Left            =   150
      MaxLength       =   150
      TabIndex        =   4
      Top             =   3105
      Width           =   4830
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   480
      Left            =   3990
      TabIndex        =   6
      Top             =   960
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   480
      Left            =   3990
      TabIndex        =   5
      Top             =   390
      Width           =   990
   End
   Begin VB.TextBox txtPassword 
      Height          =   300
      Left            =   150
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1094
      Width           =   1920
   End
   Begin VB.TextBox txtNewUserName 
      Height          =   300
      Left            =   150
      MaxLength       =   20
      TabIndex        =   0
      Top             =   435
      Width           =   1920
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Active To Date"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   150
      TabIndex        =   11
      Top             =   2193
      Width           =   1380
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Active From Date"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   150
      TabIndex        =   10
      Top             =   1507
      Width           =   1380
   End
   Begin VB.Label Label3 
      Caption         =   "Email Address"
      Height          =   210
      Left            =   150
      TabIndex        =   9
      Top             =   2880
      Width           =   1875
   End
   Begin VB.Label Label2 
      Caption         =   "Initial Password"
      Height          =   210
      Left            =   150
      TabIndex        =   8
      Top             =   851
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "New Username"
      Height          =   210
      Left            =   150
      TabIndex        =   7
      Top             =   195
      Width           =   1875
   End
End
Attribute VB_Name = "frmAddNewCMSUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbEditMode As Boolean, mlUserCode As Long, mbNoEmailAddress As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDefaultDates_Click()
    txtActiveFromDate = Format(Now, "dd/mm/yyyy")
    txtActiveToDate = MAX_DATE
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrorTrap
Dim NewUserName As String, NewPassword As String, rstCheckDupe As Recordset
Dim NewEmail As String, rstTemp As Recordset
   
    NewUserName = DoubleUpSingleQuotes(Trim(txtNewUserName.text))
    NewPassword = DoubleUpSingleQuotes(Trim(txtPassword.text))
    NewEmail = DoubleUpSingleQuotes(Trim(txtEmailAddress.text))
    
    txtPassword.text = Trim(txtPassword.text)
    txtNewUserName.text = Trim(txtNewUserName.text)
    txtEmailAddress.text = Trim(txtEmailAddress.text)
    
    If Len(txtNewUserName.text) < 3 Then
        MsgBox "Username should be between 3 and 20 characters in length", vbOKOnly + vbExclamation, AppName
        TextFieldGotFocus txtNewUserName, True
        Exit Sub
    End If
    
    
    If Not EditMode Then
        If Len(txtPassword.text) < 3 Then
            MsgBox "Password should be between 3 and 20 characters in length", vbOKOnly + vbExclamation, AppName
            TextFieldGotFocus txtPassword, True
            Exit Sub
        End If
        
        If txtPassword.text = "CMSPASSWORD" Then
            MsgBox "CMSPASSWORD is only to be used when reseting an existing password.", vbOKOnly + vbExclamation, AppName
            TextFieldGotFocus txtPassword, True
            Exit Sub
        End If
    End If
    
    If Not ValidDate(txtActiveFromDate.text) Then
        MsgBox "Invalid Active From Date", vbOKOnly + vbExclamation, AppName
        TextFieldGotFocus txtActiveFromDate, True
        Exit Sub
    End If
    
    If Not ValidDate(txtActiveToDate.text) Then
        MsgBox "Invalid Active To Date", vbOKOnly + vbExclamation, AppName
        TextFieldGotFocus txtActiveToDate, True
        Exit Sub
    End If
    
    If CDate(txtActiveFromDate.text) > CDate(txtActiveToDate.text) Then
        MsgBox "From Date should be earlier than To Date!", vbOKOnly + vbExclamation, AppName
        TextFieldGotFocus txtActiveFromDate, True
        Exit Sub
    End If
    
    If txtEmailAddress.text <> "" Then
        If InStr(1, txtEmailAddress.text, "@") = 0 Or _
           InStr(1, txtEmailAddress.text, "@") = 1 Or _
           InStr(1, txtEmailAddress.text, "@") = Len(txtEmailAddress.text) Then
                MsgBox "Email Address invalid.", vbOKOnly + vbExclamation, AppName
                TextFieldGotFocus txtEmailAddress, True
                Exit Sub
        End If
        
        If InStr(1, txtEmailAddress.text, ".") = 0 Or _
           InStr(1, txtEmailAddress.text, ".") = 1 Or _
           InStr(1, txtEmailAddress.text, ".") = Len(txtEmailAddress.text) Then
                MsgBox "Email Address invalid.", vbOKOnly + vbExclamation, AppName
                TextFieldGotFocus txtEmailAddress, True
                Exit Sub
        End If

        If InStr(1, txtEmailAddress.text, " ") > 0 Then
            MsgBox "Email Address invalid.", vbOKOnly + vbExclamation, AppName
            TextFieldGotFocus txtEmailAddress, True
            Exit Sub
        End If
    End If
           
    '
    '
    'Does this username already exist?
    '
    If Not EditMode Then
        Set rstCheckDupe = CMSDB.OpenRecordset("SELECT TheUserID " & _
                                               "FROM tblSecurity " & _
                                               "WHERE TheUserID = '" & NewUserName & "'" _
                                               , dbOpenSnapshot)
                                               
        If rstCheckDupe.BOF Then
        Else
            MsgBox "Username already exists", vbOKOnly + vbExclamation, AppName
            TextFieldGotFocus txtNewUserName, True
            Exit Sub
        End If
        rstCheckDupe.Close
    End If
                                           
    '
    'Everything ok so save new user or edit existing user as appropriate...
    '
    If EditMode Then
        CMSDB.Execute "UPDATE tblSecurity " & _
                      "SET TheUserID = '" & txtNewUserName & "', " & _
                      "    ActiveFromDate = #" & Format(txtActiveFromDate, "mm/dd/yyyy") & "#, " & _
                      "    ActiveToDate = #" & Format(txtActiveToDate, "mm/dd/yyyy") & "# " & _
                      "WHERE UserCode = " & mlUserCode
                      
        If mbNoEmailAddress Then
            CMSDB.Execute "INSERT INTO tblCMSUsersEmailAddresses " & _
                      "(UserCode, EmailAddress) " & _
                      "VALUES " & _
                      "(" & mlUserCode & ", '" & _
                      NewEmail & "')"
        Else
            CMSDB.Execute "UPDATE tblCMSUsersEmailAddresses " & _
                          "SET EmailAddress = '" & txtEmailAddress & "' " & _
                          "WHERE UserCode = " & mlUserCode
        End If
    Else
        CMSDB.Execute "INSERT INTO tblSecurity " & _
                  "(TheUserID, ThePassword, ActiveFromDate, ActiveToDate) " & _
                  "VALUES " & _
                  "('" & NewUserName & "', '" & _
                  NewPassword & "', " & _
                  "#" & Format(txtActiveFromDate, "mm/dd/yyyy") & "#, " & _
                  "#" & Format(txtActiveToDate, "mm/dd/yyyy") & "#)"
                  
        Set rstTemp = CMSDB.OpenRecordset("SELECT UserCode " & _
                                          "FROM tblSecurity " & _
                                          "WHERE TheUserID ='" & txtNewUserName & "' ", _
                                          dbOpenForwardOnly)
                                          
        CMSDB.Execute "INSERT INTO tblCMSUsersEmailAddresses " & _
                  "(UserCode, EmailAddress) " & _
                  "VALUES " & _
                  "(" & rstTemp!UserCode & ", '" & _
                  NewEmail & "')"
    End If

    '
    'Refresh parent form
    '
    HandleListBox.Requery frmSecurityAdmin.cmbUsers, False, CMSDB
    frmSecurityAdmin.cmbUsers.ListIndex = -1
    frmSecurityAdmin.cmbUsers_Click
    
    If EditMode Then
        MsgBox "User edited", vbOKOnly + vbInformation, AppName
    Else
        MsgBox "New user added", vbOKOnly + vbInformation, AppName
    End If
    
    
    Unload Me
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Property Get EditMode() As Boolean
    EditMode = mbEditMode
End Property

Public Property Let EditMode(ByVal vNewValue As Boolean)
    mbEditMode = vNewValue
End Property

Private Sub Form_Load()
On Error GoTo ErrorTrap

Dim rstEmail As Recordset, rstTemp As Recordset

    txtPassword.text = ""
    
    If EditMode Then
        txtPassword.Enabled = False
        
        txtNewUserName.text = frmSecurityAdmin.cmbUsers.text
        Set rstEmail = CMSDB.OpenRecordset("SELECT EmailAddress " & _
                                           "FROM tblCMSUsersEmailAddresses " & _
                                           "WHERE UserCode = " & mlUserCode, _
                                           dbOpenForwardOnly)
        
        If Not rstEmail.BOF Then
            mbNoEmailAddress = False
            If Not IsNull(rstEmail!EmailAddress) Then
                txtEmailAddress.text = rstEmail!EmailAddress
            Else
                txtEmailAddress.text = ""
            End If
        Else
            mbNoEmailAddress = True
            txtEmailAddress.text = ""
        End If
        
        Set rstTemp = CMSDB.OpenRecordset("SELECT ActiveFromDate, ActiveToDate " & _
                                          "FROM tblSecurity " & _
                                          "WHERE UserCode = " & mlUserCode, _
                                          dbOpenForwardOnly)
        
        If Not rstTemp.BOF Then
            txtActiveFromDate = IIf(IsNull(rstTemp!ActiveFromDate), "", rstTemp!ActiveFromDate)
            txtActiveToDate = IIf(IsNull(rstTemp!ActiveToDate), "", rstTemp!ActiveToDate)
        Else
            txtActiveFromDate = ""
            txtActiveToDate = ""
        End If
        
        If mlUserCode <= 2 Then
            txtNewUserName.Enabled = False
            txtActiveFromDate.Enabled = False
            txtActiveToDate.Enabled = False
        Else
            txtNewUserName.Enabled = True
            
            If FormUserCode = gCurrentUserCode Then
                txtActiveFromDate.Enabled = False
                txtActiveToDate.Enabled = False
            Else
                txtActiveFromDate.Enabled = True
                txtActiveToDate.Enabled = True
            End If
        End If
        
        rstEmail.Close
        rstTemp.Close
    Else
        txtPassword.Enabled = True
    End If

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Public Property Get FormUserCode() As Long
    FormUserCode = mlUserCode
End Property

Public Property Let FormUserCode(ByVal vNewValue As Long)
    mlUserCode = vNewValue
End Property


Private Sub txtActiveFromDate_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorTrap

'Must be numeric. Allow Backspace (8) and forward-slash (47). Delete and arrow keys seem to be allowed by default.

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 47 Then
        KeyAscii = 0
    End If
    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Private Sub txtActiveFromDate_LostFocus()
On Error GoTo ErrorTrap

    If IsDate(txtActiveFromDate) Then
        txtActiveFromDate = Format(txtActiveFromDate, "dd/mm/yyyy")
    ElseIf Trim(txtActiveFromDate.text) = "" Then
        txtActiveFromDate = Format(Now, "dd/mm/yyyy")
    End If

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub txtActiveToDate_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorTrap

'Must be numeric. Allow Backspace (8) and forward-slash (47). Delete and arrow keys seem to be allowed by default.

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 47 Then
        KeyAscii = 0
    End If
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub txtActiveToDate_LostFocus()
On Error GoTo ErrorTrap

    If IsDate(txtActiveToDate) Then
        txtActiveToDate = Format(txtActiveToDate, "dd/mm/yyyy")
    ElseIf Trim(txtActiveToDate.text) = "" Then
        txtActiveToDate = "31/12/9999"
    End If

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
