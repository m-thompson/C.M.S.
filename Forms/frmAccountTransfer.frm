VERSION 5.00
Begin VB.Form frmAccountTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.M.S. Transfer to/from Another Account"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   Icon            =   "frmAccountTransfer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtRefNo 
      Height          =   285
      Left            =   1230
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1620
      Width           =   1050
   End
   Begin VB.TextBox txtTranDesc 
      Height          =   285
      Left            =   1230
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1260
      Width           =   4380
   End
   Begin VB.ComboBox cmbTfrFromAccount 
      Height          =   315
      Left            =   1230
      TabIndex        =   1
      Top             =   480
      Width           =   4395
   End
   Begin VB.CommandButton cmdShowCalendar1 
      DownPicture     =   "frmAccountTransfer.frx":0442
      Height          =   285
      Left            =   2220
      Picture         =   "frmAccountTransfer.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   420
   End
   Begin VB.TextBox txtFirstDate 
      Height          =   285
      Left            =   1230
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   978
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   3750
      TabIndex        =   6
      Top             =   2160
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4710
      TabIndex        =   7
      Top             =   2160
      Width           =   885
   End
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   1230
      MaxLength       =   11
      TabIndex        =   5
      Top             =   1980
      Width           =   1200
   End
   Begin VB.ComboBox cmbTfrToAccount 
      Height          =   315
      Left            =   1230
      TabIndex        =   2
      Top             =   870
      Width           =   4395
   End
   Begin VB.Label lblRefNo 
      Caption         =   "Cheque No"
      Height          =   255
      Left            =   180
      TabIndex        =   14
      Top             =   1650
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Description"
      Height          =   225
      Left            =   180
      TabIndex        =   13
      Top             =   1290
      Width           =   840
   End
   Begin VB.Label Label4 
      Caption         =   "Date"
      Height          =   255
      Left            =   180
      TabIndex        =   12
      Top             =   180
      Width           =   810
   End
   Begin VB.Label Label2 
      Caption         =   "Transfer From: "
      Height          =   255
      Left            =   165
      TabIndex        =   10
      Top             =   540
      Width           =   1125
   End
   Begin VB.Label Label6 
      Caption         =   "Amount"
      Height          =   255
      Left            =   180
      TabIndex        =   9
      Top             =   2010
      Width           =   810
   End
   Begin VB.Label Label5 
      Caption         =   "Transfer to"
      Height          =   240
      Left            =   165
      TabIndex        =   8
      Top             =   915
      Width           =   855
   End
End
Attribute VB_Name = "frmAccountTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event TransactionEntered()
Dim mlTransactionID As Long
Dim mlAccountID As Long
Dim mlAccountIDParm As Long
Dim mlTfrAccountID As Long
Dim mlInOutTypeID As Long
Dim mlInOutID As Long
Dim mbEditMode As Boolean
Dim mbRegularTran As Boolean
Dim WithEvents frmCal As frmMiniCalendar
Attribute frmCal.VB_VarHelpID = -1

Private Sub cmbTfrFromAccount_Click()

On Error GoTo ErrorTrap
            
    SetAccountNos
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmbTfrFromAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        cmbTfrFromAccount.ListIndex = -1
    End If
End Sub

Private Sub cmbTfrFromAccount_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorTrap
    
    AutoCompleteCombo Me!cmbTfrFromAccount, KeyAscii
    
    If KeyAscii = 13 Then
        cmbTfrToAccount.SetFocus
    Else
        KeyAscii = 0 'This stops the BEEP every time Enter is pressed
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub cmbTfrToAccount_Click()

On Error GoTo ErrorTrap
        
    SetAccountNos

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub SetAccountNos()
Dim trn As TransactionDetails
On Error GoTo ErrorTrap
        
    If cmbTfrFromAccount.ListIndex > 0 Then
        mlTfrAccountID = cmbTfrFromAccount.ItemData(cmbTfrFromAccount.ListIndex)
    Else
        If cmbTfrToAccount.ListIndex > 0 Then
            mlTfrAccountID = cmbTfrToAccount.ItemData(cmbTfrToAccount.ListIndex)
        Else
            mlTfrAccountID = -1
        End If
    End If
    
    If cmbTfrFromAccount.ListIndex = 0 Then
        mlInOutID = 2
        mlInOutTypeID = 2
    Else
        mlInOutID = 1
        mlInOutTypeID = 1
    End If
    
    If (Not mbEditMode) And (mlTfrAccountID > 0) Then
        If mlInOutID = 1 Then
            txtTranDesc = "Account Transfer - From " & GetAccountName(mlTfrAccountID)
        Else
            txtTranDesc = "Account Transfer - To " & GetAccountName(mlTfrAccountID)
        End If
    End If
        
    mlAccountID = 0

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmbTfrToAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        cmbTfrFromAccount.ListIndex = -1
    End If
End Sub

Private Sub cmbTfrToAccount_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorTrap
    
    AutoCompleteCombo Me!cmbTfrToAccount, KeyAscii
    
    If KeyAscii = 13 Then
        txtTranDesc.SetFocus
    Else
        KeyAscii = 0 'This stops the BEEP every time Enter is pressed
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOK.SetFocus
        KeyAscii = 0 'This stops the BEEP every time Enter is pressed
    Else
        KeyPressValid KeyAscii, cmsUnsignedDecimals
    End If
End Sub

Private Sub txtFirstDate_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorTrap

    '
    'Move to next field when ENTER pressed
    '
    If KeyAscii = 13 Then
        cmbTfrFromAccount.SetFocus
        KeyAscii = 0 'This stops the BEEP every time Enter is pressed
    Else
        KeyPressValid KeyAscii, cmsDates, True
    End If

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub txtRefNo_GotFocus()
    TextFieldGotFocus txtRefNo
End Sub

Private Sub txtRefNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAmount.SetFocus
        KeyAscii = 0 'This stops the BEEP every time Enter is pressed
    Else
        KeyPressValid KeyAscii, cmsUnsignedIntegers
    End If
End Sub
Private Sub txtTranDesc_GotFocus()
    TextFieldGotFocus txtTranDesc
End Sub


Private Sub cmdShowCalendar1_Click()

On Error GoTo ErrorTrap
   
    With frmCal
    .SetPos = True
    .FormDate = txtFirstDate
    .XPos = Me.Left + cmdShowCalendar1.Left + cmdShowCalendar1.Width
    .YPos = Me.Top + cmdShowCalendar1.Top + cmdShowCalendar1.Height
    .Show vbModal, Me
    End With

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Unload frmCal
    Set frmCal = Nothing
    
End Sub


Private Sub frmCal_InsertDate(TheDate As String)

On Error GoTo ErrorTrap

    txtFirstDate = TheDate

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub


Private Sub txtAmount_GotFocus()
    TextFieldGotFocus txtAmount
End Sub
Private Sub txtAmount_LostFocus()
    txtAmount = Format(txtAmount, "0.00")
End Sub


Private Sub txtFirstDate_GotFocus()
    TextFieldGotFocus txtFirstDate
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim rs As Recordset, bReceipted As Boolean, str As String, lgrp As Long
On Error GoTo ErrorTrap

    If Not ValidateEntry Then Exit Sub
    
    Set rs = CMSDB.OpenRecordset("tblTransactionDates", dbOpenDynaset)
    
    With rs
    
    If mlTransactionID = 0 Then
        .AddNew
    Else
        .FindFirst "TranID = " & mlTransactionID
        .Edit
    End If
    
    !TranCodeID = GetTranCodeIDFromTranCode(gsAccountTransferTranCode, mlInOutTypeID)
    str = txtTranDesc
    
    If txtRefNo = "" Then
        !RefNo = 0
    Else
        !RefNo = CLng(txtRefNo)
    End If
    !TranDate = CDate(txtFirstDate)
    !Amount = CDbl(txtAmount) * IIf(mlInOutID = 1, 1, -1)
    !FinancialYear = GetFinancialYear(CDate(txtFirstDate))
    !FinancialMonth = GetFinancialMonth(CDate(txtFirstDate))
    !FinancialQuarter = GetFinancialQuarter(CDate(txtFirstDate))
    !TranSubTypeID = 0
    !AccountID = 0
    !TfrAccountID = mlTfrAccountID
    !BookGroupNo = -1
    !TranDescription = txtTranDesc
    
    .Update
    
    RaiseEvent TransactionEntered

    End With
    
    rs.Close
    Set rs = Nothing
    
    Unload Me

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub Form_Load()
Dim rs As Recordset, TranStuff As TransactionDetails
On Error GoTo ErrorTrap

    Set frmCal = New frmMiniCalendar
        
    
    HandleListBox.PopulateListBox cmbTfrFromAccount, _
                              "SELECT AccountID, AccountName " & _
                              "FROM tblBankAccounts ORDER BY 1 ", CMSDB, 0, "", False, 1
                              
    
                              
    HandleListBox.PopulateListBox cmbTfrToAccount, _
                              "SELECT AccountID, AccountName " & _
                              "FROM tblBankAccounts ORDER BY 1 ", CMSDB, 0, "", False, 1
                              
    
    If mlTransactionID = 0 Then
        ClearForm
        mbEditMode = False
        
        If mlAccountIDParm >= 0 Then
            HandleListBox.SelectItem cmbTfrFromAccount, mlAccountIDParm
        Else
            cmbTfrFromAccount.ListIndex = -1
        End If
        
        If cmbTfrToAccount.ListCount = 2 Then
            If mlAccountIDParm > 0 Then
                cmbTfrToAccount.ListIndex = 0
            Else
                cmbTfrToAccount.ListIndex = 1
            End If
        Else
            cmbTfrToAccount.ListIndex = -1
        End If
    Else
        mbEditMode = True
        TranStuff = GetTransactionDetails(mlTransactionID)
        With TranStuff
        txtAmount = Format(.Amount * IIf(.Amount < 0, -1, 1), "0.00")
        txtFirstDate = .TransactionDate
        txtTranDesc = .TransactionDescription
        txtRefNo = IIf(.RefNo = 0, "", .RefNo)
        Select Case .InOutID
        Case 1
            HandleListBox.SelectItem cmbTfrFromAccount, .TfrAccountID
            HandleListBox.SelectItem cmbTfrToAccount, 0
        Case 2
            HandleListBox.SelectItem cmbTfrFromAccount, 0
            HandleListBox.SelectItem cmbTfrToAccount, .TfrAccountID
        End Select
        End With
    End If
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub
Private Sub ClearForm()
On Error GoTo ErrorTrap
Dim txt As Control
        
    txtAmount = ""
    txtFirstDate = ""
    txtTranDesc = ""
    cmbTfrFromAccount.ListIndex = -1
    cmbTfrToAccount.ListIndex = -1
    
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Function ValidateEntry() As Boolean
Dim rs As Recordset, str As String, bMatch As Boolean, lnum As Long, lNum2 As Long, lNum3 As Long
On Error GoTo ErrorTrap

    str = Format(txtAmount, "0.00")
    
    If Len(str) > 11 Then
        MsgBox "Maximum tranaction amount is 99999999.99", vbOKOnly + vbExclamation, AppName
        TextFieldGotFocus txtAmount, True
        ValidateEntry = False
        Exit Function
    Else
        txtAmount = str
    End If
        
    If txtTranDesc = "" Then
        MsgBox "Please enter a Transaction Description", vbOKOnly + vbExclamation, AppName
        TextFieldGotFocus txtTranDesc, True
        ValidateEntry = False
        Exit Function
    End If
      
    If Not ValidDate(txtFirstDate) Then
        MsgBox "Please enter a valid Transaction Date", vbOKOnly + vbExclamation, AppName
        TextFieldGotFocus txtFirstDate, True
        ValidateEntry = False
        Exit Function
    End If
    
    If GetAccountStartDate(mlAccountID) > _
        CDate(txtFirstDate) Or _
       GetAccountStartDate(mlTfrAccountID) > _
        CDate(txtFirstDate) Then
        MsgBox "Transaction Date should not be prior to " & _
                  "account start dates.", vbOKOnly + vbExclamation, AppName
        TextFieldGotFocus txtFirstDate, True
        ValidateEntry = False
        Exit Function
    End If
      
    If txtRefNo <> "" Then
        If Not IsNumber(txtRefNo, False, False, False) Then
            MsgBox "Please enter a valid Cheque Number", vbOKOnly + vbExclamation, AppName
            TextFieldGotFocus txtRefNo, True
            ValidateEntry = False
            Exit Function
        End If
    End If

    If Not IsNumber(txtAmount, True, False, True) Then
        MsgBox "Please enter a valid Transaction Amount", vbOKOnly + vbExclamation, AppName
        TextFieldGotFocus txtAmount, True
        ValidateEntry = False
        Exit Function
    End If
        
    If CDbl(txtAmount) <= 0 Then
        MsgBox "Amount should be greater than 0.00", vbOKOnly + vbExclamation, AppName
        TextFieldGotFocus txtAmount, True
        ValidateEntry = False
        Exit Function
    End If
    
    If txtRefNo = "" Then
        lnum = 0
    Else
        lnum = CLng(txtRefNo)
    End If
    
    If mlTfrAccountID <> 0 And mlAccountID <> 0 Then
        MsgBox "One of the accounts must be the Current Account", vbOKOnly + vbExclamation, AppName
        cmbTfrFromAccount.SetFocus
        ValidateEntry = False
        Exit Function
    End If
    
    If mlTfrAccountID = -1 Or mlAccountID = -1 Then
        MsgBox "'From' and 'To' accounts must be specified. One should be the Current Account.", vbOKOnly + vbExclamation, AppName
        cmbTfrFromAccount.SetFocus
        ValidateEntry = False
        Exit Function
    End If
    
    If mlTfrAccountID = mlAccountID Then
        MsgBox "'From Account' cannot be the same as the 'To Account'", vbOKOnly + vbExclamation, AppName
        cmbTfrFromAccount.SetFocus
        ValidateEntry = False
        Exit Function
    End If
    
    If cmbTfrFromAccount.ListIndex = -1 Or cmbTfrToAccount.ListIndex = -1 Then
        MsgBox "'From' and 'To' accounts must be specified. One should be the Current Account.", vbOKOnly + vbExclamation, AppName
        cmbTfrFromAccount.SetFocus
        ValidateEntry = False
        Exit Function
    End If
    
    If cmbTfrFromAccount.ListIndex <> 0 And cmbTfrToAccount.ListIndex <> 0 Then
        MsgBox "One of the accounts must be the Current Account", vbOKOnly + vbExclamation, AppName
        cmbTfrFromAccount.SetFocus
        ValidateEntry = False
        Exit Function
    End If
        
    If GlobalParms.GetValue("AccountsCheckDupeEntries", "TrueFalse") Then
        Set rs = CMSDB.OpenRecordset("SELECT 1 " & _
                                     "FROM tblTransactionDates " & _
                                     "WHERE TranDate = " & GetDateStringForSQLWhere(txtFirstDate) & _
                                     " AND TranCodeID = " & GetTranCodeIDFromTranCode(gsAccountTransferTranCode, mlInOutTypeID) & _
                                     " AND Amount = " & CDbl(txtAmount) & _
                                     " AND RefNo = " & lnum & _
                                     " AND AccountID = " & mlAccountID & _
                                     " AND TfrAccountID = " & mlTfrAccountID, dbOpenForwardOnly)
                                     
        If Not rs.BOF Then
            If MsgBox("An identical transaction already exists. Do you want to continue?", vbYesNo + vbQuestion, AppName) = vbNo Then
                TextFieldGotFocus txtFirstDate, True
                ValidateEntry = False
                Exit Function
            End If
        End If
    End If
    
    ValidateEntry = True
    
    On Error Resume Next
    rs.Close
    Set rs = Nothing

    Exit Function
ErrorTrap:
    Call EndProgram
End Function

Public Property Get TransactionID() As Long
    TransactionID = mlTransactionID
End Property

Public Property Let TransactionID(ByVal vNewValue As Long)
    mlTransactionID = vNewValue
End Property
Public Property Let AccountID(ByVal vNewValue As Long)
    mlAccountIDParm = vNewValue
End Property
Public Property Get RegularTran() As Boolean
    RegularTran = mbRegularTran
End Property

Public Property Let RegularTran(ByVal vNewValue As Boolean)
    mbRegularTran = vNewValue
End Property


Private Sub txtFirstDate_LostFocus()
Dim str As String

On Error GoTo ErrorTrap

    If IsNumber(txtFirstDate, False, False, False) Then
        With frmAccountSummary
        str = txtFirstDate & "/" & .cmbMonth.ItemData(.cmbMonth.ListIndex) & "/" & _
              GetNormalYearFromFinancialYear(CDate("01/" & .cmbMonth.ItemData(.cmbMonth.ListIndex) & "/" & _
                                                .cmbYear.text))
        End With
    Else
        str = txtFirstDate
    End If


    If Not IsDate(str) Then Exit Sub
    
    txtFirstDate = Format(str, "dd/mm/yyyy")

    Exit Sub
ErrorTrap:
    EndProgram

    
End Sub

Private Sub txtTranDesc_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorTrap

    If KeyAscii = 13 Then
        txtRefNo.SetFocus
        KeyAscii = 0 'This stops the BEEP every time Enter is pressed
    End If

    Exit Sub
ErrorTrap:
    EndProgram


End Sub
