VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAccountSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.M.S. Account Summary"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10245
   Icon            =   "frmAccountSummary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBigFrame 
      Height          =   7005
      Left            =   90
      TabIndex        =   14
      Top             =   60
      Width           =   9990
      Begin VB.ComboBox cmbAccount 
         Height          =   315
         Left            =   885
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   195
         Width           =   4395
      End
      Begin VB.Frame fraMode 
         Caption         =   "Display Mode"
         Height          =   1425
         Left            =   150
         TabIndex        =   15
         Top             =   600
         Width           =   2010
         Begin VB.OptionButton optMonth 
            Caption         =   "Tax Year && Month"
            Height          =   195
            Left            =   135
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   360
            Width           =   1770
         End
         Begin VB.OptionButton optDateRange 
            Caption         =   "Date Range"
            Height          =   195
            Left            =   135
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   705
            Width           =   1770
         End
         Begin VB.OptionButton optSearch 
            Caption         =   "Transaction Search"
            Height          =   195
            Left            =   135
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1035
            Width           =   1770
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxAccounts 
         Height          =   4770
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2070
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   8414
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
      End
      Begin VB.Frame fraMonth 
         Height          =   1425
         Left            =   2250
         TabIndex        =   51
         Top             =   600
         Width           =   7590
         Begin VB.CommandButton cmdGoToCurrentMonth 
            Height          =   315
            Left            =   2760
            Picture         =   "frmAccountSummary.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   61
            TabStop         =   0   'False
            ToolTipText     =   "Go to current month"
            Top             =   675
            Width           =   345
         End
         Begin VB.CheckBox chkSuppressWarnings 
            Caption         =   "Suppress Warnings"
            Height          =   390
            Left            =   6360
            TabIndex        =   54
            ToolTipText     =   "Decide whether to be warned about missing transactions"
            Top             =   630
            Width           =   1125
         End
         Begin VB.ComboBox cmbYear 
            Height          =   315
            Left            =   210
            Style           =   2  'Dropdown List
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   675
            Width           =   765
         End
         Begin VB.ComboBox cmbMonth 
            Height          =   315
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   675
            Width           =   1455
         End
         Begin MSComCtl2.UpDown ctlDateUpDown 
            Height          =   315
            Left            =   2475
            TabIndex        =   55
            ToolTipText     =   "Use +/- keys to scroll through months"
            Top             =   675
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   556
            _Version        =   393216
            Alignment       =   0
            OrigLeft        =   2055
            OrigTop         =   455
            OrigRight       =   2295
            OrigBottom      =   770
            Enabled         =   -1  'True
         End
         Begin VB.Label Label22 
            Caption         =   "Tax Year"
            Height          =   285
            Left            =   210
            TabIndex        =   58
            Top             =   465
            Width           =   825
         End
         Begin VB.Label Label2 
            Caption         =   "Month"
            Height          =   255
            Left            =   1035
            TabIndex        =   57
            Top             =   465
            Width           =   975
         End
         Begin VB.Label lblInfo 
            Height          =   255
            Left            =   3180
            TabIndex        =   56
            Top             =   735
            Width           =   2985
         End
      End
      Begin VB.Frame fraDateRange 
         Height          =   1425
         Left            =   2250
         TabIndex        =   44
         Top             =   600
         Width           =   7590
         Begin VB.CommandButton cmdShowCalendar2 
            DownPicture     =   "frmAccountSummary.frx":048B
            Height          =   315
            Left            =   2700
            Picture         =   "frmAccountSummary.frx":08CD
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   690
            Width           =   420
         End
         Begin VB.CommandButton cmdShowCalendar1 
            DownPicture     =   "frmAccountSummary.frx":0D0F
            Height          =   315
            Left            =   1125
            Picture         =   "frmAccountSummary.frx":1151
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   690
            Width           =   420
         End
         Begin VB.TextBox txtFirstDate 
            Height          =   315
            Left            =   135
            MaxLength       =   10
            TabIndex        =   46
            Top             =   690
            Width           =   978
         End
         Begin VB.TextBox txtLastDate 
            Height          =   315
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   45
            Top             =   690
            Width           =   978
         End
         Begin VB.Label Label7 
            Caption         =   "End Date"
            Height          =   255
            Left            =   1710
            TabIndex        =   50
            Top             =   465
            Width           =   1275
         End
         Begin VB.Label Label8 
            Caption         =   "Start Date"
            Height          =   255
            Left            =   135
            TabIndex        =   49
            Top             =   465
            Width           =   1275
         End
      End
      Begin VB.Frame fraSearch 
         Height          =   1425
         Left            =   2250
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   7590
         Begin VB.TextBox txtSearchDate2 
            Height          =   315
            Left            =   1050
            MaxLength       =   10
            TabIndex        =   35
            Top             =   615
            Width           =   978
         End
         Begin VB.TextBox txtSearchDate1 
            Height          =   315
            Left            =   1050
            MaxLength       =   10
            TabIndex        =   34
            Top             =   255
            Width           =   978
         End
         Begin VB.CommandButton cmdSearchDate1 
            DownPicture     =   "frmAccountSummary.frx":1593
            Height          =   315
            Left            =   2040
            Picture         =   "frmAccountSummary.frx":19D5
            Style           =   1  'Graphical
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   255
            Width           =   420
         End
         Begin VB.CommandButton cmdSearchDate2 
            DownPicture     =   "frmAccountSummary.frx":1E17
            Height          =   315
            Left            =   2040
            Picture         =   "frmAccountSummary.frx":2259
            Style           =   1  'Graphical
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   615
            Width           =   420
         End
         Begin VB.TextBox txtDesc 
            Height          =   315
            Left            =   2985
            MaxLength       =   100
            TabIndex        =   31
            ToolTipText     =   "Use '|' as logical OR"
            Top             =   615
            Width           =   2940
         End
         Begin VB.TextBox txtMaxAmount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2985
            MaxLength       =   9
            TabIndex        =   30
            Top             =   975
            Width           =   978
         End
         Begin VB.TextBox txtMinAmount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1050
            MaxLength       =   9
            TabIndex        =   29
            Top             =   975
            Width           =   978
         End
         Begin VB.TextBox txtMinRef 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4830
            MaxLength       =   9
            TabIndex        =   28
            ToolTipText     =   "Gift Aid or Cheque Number"
            Top             =   975
            Width           =   978
         End
         Begin VB.TextBox txtMaxRef 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6522
            MaxLength       =   9
            TabIndex        =   27
            ToolTipText     =   "Gift Aid or Cheque Number"
            Top             =   975
            Width           =   978
         End
         Begin VB.CommandButton cmdTranCodes 
            Caption         =   "Codes"
            Height          =   315
            Left            =   4185
            TabIndex        =   26
            ToolTipText     =   "No Transaction Code criteria selected"
            Top             =   255
            Width           =   855
         End
         Begin VB.ComboBox cmbTranType 
            Height          =   315
            Left            =   2970
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   255
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   315
            Left            =   6810
            TabIndex        =   24
            Top             =   255
            Width           =   690
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
            Height          =   315
            Left            =   6090
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "Shortcut: F5"
            Top             =   255
            Width           =   690
         End
         Begin VB.CommandButton cmdTranSubCodes 
            Caption         =   "Sub-Tp"
            Height          =   315
            Left            =   5070
            TabIndex        =   22
            ToolTipText     =   "No Transaction Sub-Type criteria selected"
            Top             =   255
            Width           =   855
         End
         Begin VB.CommandButton cmdBkGrp 
            Caption         =   "Bk Grp"
            Height          =   315
            Left            =   6090
            TabIndex        =   21
            ToolTipText     =   "No Book-Group criteria selected"
            Top             =   615
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label Label9 
            Caption         =   "Start Date"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   300
            Width           =   1275
         End
         Begin VB.Label Label10 
            Caption         =   "End Date"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   667
            Width           =   810
         End
         Begin VB.Label Label11 
            Caption         =   "Text"
            Height          =   255
            Left            =   2595
            TabIndex        =   41
            Top             =   667
            Width           =   825
         End
         Begin VB.Label Label12 
            Caption         =   "Min Amount"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1035
            Width           =   885
         End
         Begin VB.Label Label13 
            Caption         =   "Max Amount"
            Height          =   255
            Left            =   2070
            TabIndex        =   39
            Top             =   1035
            Width           =   945
         End
         Begin VB.Label Label15 
            Caption         =   "Max Ref"
            Height          =   255
            Left            =   5865
            TabIndex        =   38
            Top             =   1035
            Width           =   660
         End
         Begin VB.Label Label16 
            Caption         =   "Min Ref"
            Height          =   255
            Left            =   4230
            TabIndex        =   37
            Top             =   1035
            Width           =   555
         End
         Begin VB.Label Label14 
            Caption         =   "Type"
            Height          =   255
            Left            =   2595
            TabIndex        =   36
            Top             =   300
            Width           =   420
         End
      End
      Begin VB.Label Label17 
         Caption         =   "Account"
         Height          =   285
         Left            =   210
         TabIndex        =   60
         Top             =   255
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   390
      Left            =   9210
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   8370
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Current Account"
      Height          =   1575
      Left            =   90
      TabIndex        =   2
      Top             =   7185
      Width           =   3000
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Opening Balance"
         Height          =   210
         Left            =   105
         TabIndex        =   12
         Top             =   375
         Width           =   1305
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Income"
         Height          =   210
         Left            =   105
         TabIndex        =   11
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Expense"
         Height          =   210
         Left            =   105
         TabIndex        =   10
         Top             =   810
         Width           =   1305
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Closing Balance"
         Height          =   210
         Left            =   105
         TabIndex        =   9
         Top             =   1260
         Width           =   1305
      End
      Begin VB.Label lblOpeningBal 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1515
         TabIndex        =   8
         Top             =   375
         Width           =   1170
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Surplus"
         Height          =   210
         Left            =   105
         TabIndex        =   7
         Top             =   1035
         Width           =   1305
      End
      Begin VB.Label lblIncome 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1515
         TabIndex        =   6
         Top             =   600
         Width           =   1170
      End
      Begin VB.Label lblExpense 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1515
         TabIndex        =   5
         Top             =   810
         Width           =   1170
      End
      Begin VB.Label lblSurplus 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1515
         TabIndex        =   4
         Top             =   1035
         Width           =   1170
      End
      Begin VB.Label lblClosingBal 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1515
         TabIndex        =   3
         Top             =   1260
         Width           =   1170
      End
   End
   Begin VB.Frame fraOtherAccounts 
      Caption         =   "Other Accounts"
      Height          =   1575
      Left            =   3270
      TabIndex        =   0
      Top             =   7185
      Width           =   5745
      Begin VB.TextBox txtOtherAccounts 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1245
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   5595
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9315
      Top             =   7290
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Actions"
      Begin VB.Menu mnuNewTransaction 
         Caption         =   "New Transaction (n)"
      End
      Begin VB.Menu mnuNewTransfer 
         Caption         =   "New Account Transfer (t)"
      End
      Begin VB.Menu mnuAddRegular 
         Caption         =   "Add Month's Regular Transactions (r)"
      End
      Begin VB.Menu mnuDivider 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTranTypes 
         Caption         =   "Transaction Types"
      End
      Begin VB.Menu mnuGiftAid 
         Caption         =   "Gift Aid"
      End
      Begin VB.Menu mnuDiv3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMissingTrans 
         Caption         =   "Show Missing Transactions"
      End
      Begin VB.Menu mnuInterestAlert 
         Caption         =   "Alert for Missing Bank Interest"
      End
      Begin VB.Menu mnuDiv2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpeningBal 
         Caption         =   "Manage Bank Accounts"
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReporting 
         Caption         =   "Reporting..."
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export to Spreadsheet..."
      End
      Begin VB.Menu mnuDiv4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options..."
      End
      Begin VB.Menu mnuDiv20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "Calculator..."
      End
   End
   Begin VB.Menu mnuActions2 
      Caption         =   "Actions2"
      Visible         =   0   'False
      Begin VB.Menu mnuNewTran 
         Caption         =   "New Transaction (n)"
      End
      Begin VB.Menu mnuAccTfr 
         Caption         =   "New Account Transfer (t)"
      End
      Begin VB.Menu mnuDivB 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditTran 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuMove 
         Caption         =   "Move Date"
      End
      Begin VB.Menu mnuDivA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddRegular2 
         Caption         =   "Add Month's Regular Transactions (r)"
      End
   End
End
Attribute VB_Name = "frmAccountSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlFormTaxYear As Long, mlFormNormalYear As Long, mlFormMonth As Long
Dim mlTranID As Long, mlPrevRow As Long
Dim mlAccountID As Long, mfCurrentAccClosingBal As Double
Dim WithEvents frmTranEntry As frmTransactionEntry
Attribute frmTranEntry.VB_VarHelpID = -1
Dim WithEvents frmAccTransfer As frmAccountTransfer
Attribute frmAccTransfer.VB_VarHelpID = -1
Dim WithEvents frmTransactionTypes As frmTranTypes
Attribute frmTransactionTypes.VB_VarHelpID = -1
Dim WithEvents frmSelectInOutType As frmAccountInOutTypeSelect
Attribute frmSelectInOutType.VB_VarHelpID = -1
Dim WithEvents frmReg As frmRegularTransactions
Attribute frmReg.VB_VarHelpID = -1
Dim WithEvents frmCal1 As frmMiniCalendar
Attribute frmCal1.VB_VarHelpID = -1
Dim WithEvents frmCal2 As frmMiniCalendar
Attribute frmCal2.VB_VarHelpID = -1
Dim WithEvents frmCal3 As frmMiniCalendar
Attribute frmCal3.VB_VarHelpID = -1
Dim WithEvents frmTrnCdSel As frmTranCodeSelect
Attribute frmTrnCdSel.VB_VarHelpID = -1
Dim WithEvents frmMissingReceiptsList As frmShowList
Attribute frmMissingReceiptsList.VB_VarHelpID = -1
Dim UpperLimit As Byte, bSuppress As Boolean, mbFindMode As Boolean
Dim msStartDate As String, msEndDate As String, mbOrderingCols As Boolean
Dim msTranCodeSearch As String
Dim msTranSubCodeSearch As String
Dim msBkGrpSearch As String
Dim mbExcludeTranCodes As Boolean
Dim mbExcludeTranSubCodes As Boolean
Dim mbExcludeBookGroups As Boolean
Dim mlMouseFlxX As Long
Dim mlMouseFlxY As Long
Dim mlCodeType As Long, mbFormLoading As Boolean


Private Sub chkSuppressWarnings_Click()
On Error GoTo ErrorTrap

    If chkSuppressWarnings.value = vbChecked Then
        GlobalParms.Save "AccountsWarningsDefault", "TrueFalse", False
    Else
        GlobalParms.Save "AccountsWarningsDefault", "TrueFalse", True
    End If

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmbAccount_Click()

On Error GoTo ErrorTrap

    If mbFormLoading Then Exit Sub

    If cmbAccount.ListIndex = -1 Then
        cmbAccount.ListIndex = 0
    End If
    
    mlAccountID = cmbAccount.ItemData(cmbAccount.ListIndex)
    
    SetUpFormControls
    
    cmbMonth_Click

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmbMonth_Click()

On Error GoTo ErrorTrap

    If bSuppress Then Exit Sub
    
    If cmbYear.ListIndex > -1 And cmbMonth.ListIndex > -1 Then
        GetDates
        GetTotals
        FillAccountsGrid
    Else
        ClearForm
    End If

    Exit Sub
ErrorTrap:
    EndProgram

End Sub




Private Sub cmbYear_Click()

On Error GoTo ErrorTrap

    If cmbYear.ListIndex > -1 And cmbMonth.ListIndex > -1 Then
        GetDates
        GetTotals
        FillAccountsGrid
    Else
        ClearForm
    End If

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub ClearForm()

On Error GoTo ErrorTrap

'    mlFormMonth = 0
'    mlFormNormalYear = 0
'    mlFormTaxYear = 0
    
    flxAccounts.Rows = 1
    
    lblOpeningBal = ""
    lblIncome = ""
    lblExpense = ""
    lblSurplus = ""
    lblClosingBal = ""

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdClear_Click()
On Error Resume Next
    
    txtSearchDate1 = ""
    txtSearchDate2 = ""
    txtMinAmount = ""
    txtMaxAmount = ""
    txtDesc = ""
    cmbTranType.ListIndex = 0
    txtMinRef = ""
    txtMaxRef = ""
    
    msTranCodeSearch = ""
    msTranSubCodeSearch = ""
    msBkGrpSearch = ""
    mbExcludeBookGroups = False
    mbExcludeTranCodes = False
    mbExcludeTranSubCodes = False
    
    mlCodeType = 1
    TranCodeButtonText SuppressMsg:=True
    mlCodeType = 2
    TranCodeButtonText SuppressMsg:=True
    mlCodeType = 3
    TranCodeButtonText SuppressMsg:=True
   
    ClearForm
        
End Sub

Private Sub cmdClose_Click()

On Error GoTo ErrorTrap

    Unload Me

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub GetDates()
Dim str As String, str1 As String, str2 As String, str3 As String
Dim arr() As String, str4 As String, str5 As String, i As Long
On Error GoTo ErrorTrap

    If mlAccountID > 0 Then Exit Sub

    mlFormMonth = cmbMonth.ItemData(cmbMonth.ListIndex)
    mlFormTaxYear = cmbYear.text
    mlFormNormalYear = GetNormalYearFromFinancialYear(CDate("01/" & mlFormMonth & "/" & mlFormTaxYear))
    
    msStartDate = txtFirstDate
    msEndDate = txtLastDate
    
    lblInfo = "Calendar month: " & GetMonthName(mlFormMonth) & " " & mlFormNormalYear
    
    If chkSuppressWarnings.value = vbUnchecked And optMonth Then
        str1 = CheckForMissingInterest
        str2 = CheckForMissingReceipts
        str3 = CheckForMissingRegularPayments
        
        If str2 <> "" Then
            arr() = Split(Split(str2, "|")(1), ",")
            str = IIf(str1 = "", "", str1 & vbCrLf & vbCrLf) & _
                  IIf(str2 = "", "", Split(str2, "|")(0) & vbCrLf & vbCrLf) & _
                  str3
        Else
            str = IIf(str1 = "", "", str1 & vbCrLf & vbCrLf) & _
                  IIf(str2 = "", "", str2 & vbCrLf & vbCrLf) & _
                  str3
        End If
        
        If str2 <> "" Then 'are there missing receipts?
            str = str & "Do you want to remove 'missing receipts' alerts?"
        End If
        
        If str <> "" Then
            If str2 = "" Then
                MsgBox str, vbOKOnly + vbInformation, AppName & " - Accounts"
            Else
                If MsgBox(str, vbYesNo + vbQuestion, AppName & " - Accounts") = vbYes Then
                    '2nd element of this array will contain a date list for no transactions
                    
                    For i = 0 To UBound(arr)
                        
                        If str5 = "" Then
                            str5 = CStr(CLng(CDate(arr(i)))) & "," & Format(arr(i), "dddd mmmm dd yyyy")
                        Else
                            str5 = str5 & "|" & CStr(CLng(CDate(arr(i)))) & "," & Format(arr(i), "dddd mmmm dd yyyy")
                        End If
                        
                    
                    Next i
                    
                    Set frmMissingReceiptsList = New frmShowList
                    
                    With frmMissingReceiptsList
                    
                    .FormCaption = "C.M.S. Missing Receipts"
                    .FormMessage = "Select dates for which there were no receipts"
                    .FormSQL = ""
                    .ListToShow = str5
                    .CheckBoxStyle = True
                    .Show vbModal, Me
                    
                    End With
                    
                    On Error Resume Next
                    Set frmMissingReceiptsList = Nothing
                    
                    On Error GoTo ErrorTrap
                End If
            End If
        End If
    End If
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdGoToCurrentMonth_Click()
    GoToCurrentMonth
End Sub

Private Sub cmdSearch_Click()

On Error Resume Next

    If Trim(txtDesc) <> "" Or _
        ValidDate(txtSearchDate1) Or _
        ValidDate(txtSearchDate2) Or _
        Trim(txtMinAmount) <> "" Or _
        Trim(txtMaxAmount) <> "" Or _
        Trim(txtMinRef) <> "" Or _
        Trim(txtMaxRef) <> "" Or _
        Trim(msTranCodeSearch) <> "" Or _
        Trim(msTranSubCodeSearch) <> "" Or _
        Trim(msBkGrpSearch) <> "" Or _
        cmbTranType.ListIndex > 0 Then
    
            cmdSearch.SetFocus
            FillAccountsGridSearch
            
    Else
    
        ShowMessage "Please enter search criteria", 1500, Me, , vbRed
            
    End If

End Sub

Private Sub cmdSearchDate1_Click()

On Error GoTo ErrorTrap
    
    If FormIsOpen("frmMiniCalendar") Then Exit Sub
    
    With frmCal1
    
    .SetPos = True
    .XPos = Me.Left + fraBigFrame.Left + fraSearch.Left + cmdSearchDate1.Left + cmdSearchDate1.Width
    .YPos = Me.Top + fraBigFrame.Top + fraSearch.Top + cmdSearchDate1.Top + cmdSearchDate1.Height
    .FormDate = txtSearchDate1
    .Show vbModeless, Me

    End With

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub
Private Sub cmdSearchDate2_Click()

On Error GoTo ErrorTrap

    If FormIsOpen("frmMiniCalendar") Then Exit Sub
    
    With frmCal2
    
    .SetPos = True
    .XPos = Me.Left + fraBigFrame.Left + fraSearch.Left + cmdSearchDate2.Left + cmdSearchDate2.Width
    .YPos = Me.Top + fraBigFrame.Top + fraSearch.Top + cmdSearchDate2.Top + cmdSearchDate2.Height
    .FormDate = txtSearchDate2
    .Show vbModeless, Me

    End With

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdShowCalendar1_Click()

On Error GoTo ErrorTrap

    If FormIsOpen("frmMiniCalendar") Then Exit Sub
    
    Set frmCal1 = New frmMiniCalendar
    
    With frmCal1
    
    .SetPos = True
    .XPos = Me.Left + fraBigFrame.Left + fraDateRange.Left + cmdShowCalendar1.Left + cmdShowCalendar1.Width
    .YPos = Me.Top + fraBigFrame.Top + fraDateRange.Top + cmdShowCalendar1.Top + cmdShowCalendar1.Height
    .FormDate = txtFirstDate
    .Show vbModeless, Me

    End With
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub
Private Sub cmdShowCalendar2_Click()

On Error GoTo ErrorTrap

    If FormIsOpen("frmMiniCalendar") Then Exit Sub
    
    Set frmCal2 = New frmMiniCalendar
    
    With frmCal2
    
    .SetPos = True
    .XPos = Me.Left + fraBigFrame.Left + fraDateRange.Left + cmdShowCalendar2.Left + cmdShowCalendar2.Width
    .YPos = Me.Top + fraBigFrame.Top + fraDateRange.Top + cmdShowCalendar2.Top + cmdShowCalendar2.Height
    .FormDate = txtLastDate
    .Show vbModeless, Me

    End With
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdTranCodes_Click()

On Error GoTo ErrorTrap
    
    mlCodeType = 1
    frmTrnCdSel.TranCodes = msTranCodeSearch
    frmTrnCdSel.ExcludeCodes = mbExcludeTranCodes
    frmTrnCdSel.Show vbModal, Me
        
    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub cmdTranSubCodes_Click()

On Error GoTo ErrorTrap
    
    mlCodeType = 2
    frmTrnCdSel.TranCodes = msTranSubCodeSearch
    frmTrnCdSel.ExcludeCodes = mbExcludeTranSubCodes
    frmTrnCdSel.Show vbModal, Me
        
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdBkGrp_Click()

On Error GoTo ErrorTrap
    
    mlCodeType = 3
    frmTrnCdSel.TranCodes = msBkGrpSearch
    frmTrnCdSel.ExcludeCodes = mbExcludeBookGroups
    frmTrnCdSel.Show vbModal, Me
        
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub ctlDateUpDown_DownClick()
On Error GoTo ErrorTrap
    
    If cmbMonth.ListIndex = 0 Then
        If cmbYear.ListIndex > 0 Then
            bSuppress = True
            cmbMonth.ListIndex = UpperLimit
            bSuppress = False
            cmbYear.ListIndex = cmbYear.ListIndex - 1
        End If
    Else
        cmbMonth.ListIndex = cmbMonth.ListIndex - 1
    End If

    Exit Sub
    
ErrorTrap:
    EndProgram
    

End Sub

Private Sub ctlDateUpDown_UpClick()
On Error GoTo ErrorTrap
    
    If cmbMonth.ListIndex = UpperLimit Then
        If cmbYear.ListIndex < cmbYear.ListCount - 1 Then
            bSuppress = True
            cmbMonth.ListIndex = 0
            bSuppress = False
            cmbYear.ListIndex = cmbYear.ListIndex + 1
        End If
    Else
        cmbMonth.ListIndex = cmbMonth.ListIndex + 1
    End If

    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub


Private Sub flxAccounts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorTrap
Dim TheRow As Long, sDate As String, Tran As TransactionDetails
'
'X and Y are relative to top left of the flxGrid
'

    mlMouseFlxX = CLng(X)
    mlMouseFlxY = CLng(Y)

    If optSearch Or optDateRange Then
        flxAccounts.ToolTipText = ""
        Exit Sub
    End If

    If cmbYear.ListIndex > -1 And cmbMonth.ListIndex > -1 Then
        With flxAccounts
                    
        '
        'Use current Y position to work out which row of the grid has been hovered
        '
        TheRow = (Ceiling(CDbl(Y) / .RowHeight(0))) + .TopRow - 2
        If TheRow <= .Rows - 1 And TheRow > 0 Then
                        
            sDate = .TextMatrix(TheRow, 0)
            
            If Not IsNumber(.TextMatrix(TheRow, 4), False, False, False) Then
                .ToolTipText = ""
                Exit Sub
            End If
            
            Tran = GetTransactionDetails(CLng(.TextMatrix(TheRow, 4)))
            
            If ValidDate(sDate) Then
                
'                .ToolTipText = "Contributions for " & sDate & ": " & _
'                                Format(GetFiguresForDate(sDate), "£0.00")
                .ToolTipText = "Contributions for " & sDate & ": " & _
                                Format(GetAccountAmountBetweenDates(sDate, sDate, True, Tran.BookGroupNo), "£0.00")
                
            Else
                
                .ToolTipText = ""
                
            End If
            
        Else
            .ToolTipText = ""
        End If
        
        End With
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Function GetFiguresForDate(TheDate As String, Optional AccountID As Long = 0) As Double
On Error GoTo ErrorTrap
Dim sSQL As String, rs As Recordset

    'get total of all receipts
    sSQL = "SELECT ABS(SUM(e.Amount)) AS ReceiptSum " & _
           "FROM ((tblTransactionTypes b " & _
           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
           "INNER JOIN tblTransactionDates e ON e.TranCodeID = b.TranCodeID " & _
           "Where d.InOutID = 1 " & _
           "AND b.OnReceipt = TRUE " & _
           "AND e.TranDate = #" & Format(TheDate, "mm/dd/yyyy") & "# " & _
           "AND e.AccountID = " & AccountID
         
    Set rs = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    With rs
    
    If .BOF Or .EOF Then
        GetFiguresForDate = 0
    Else
        If IsNull(!ReceiptSum) Then
            GetFiguresForDate = 0
        Else
            GetFiguresForDate = !ReceiptSum
        End If
    End If
    
    End With

    rs.Close
    Set rs = Nothing

    Exit Function
ErrorTrap:
    Call EndProgram
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If KeyCode = 116 Then 'F5
        If optSearch Then
            cmdSearch_Click
        End If
    End If
    
    If optMonth Then
        If KeyCode = vbKeyN Then 'letter N
            mnuNewTran_Click
        End If
        If KeyCode = vbKeyR Then 'letter r
            mnuAddRegular_Click
        End If
        If KeyCode = vbKeyT Then 'letter T
            mnuAccTfr_Click
        End If
    End If
    

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next

    Select Case KeyAscii
    Case 43 '   +
        ctlDateUpDown_UpClick
    Case 61 '   =
        ctlDateUpDown_UpClick
    Case 45 '   -
        ctlDateUpDown_DownClick
    Case 95 ' underscore
        ctlDateUpDown_DownClick
    End Select


End Sub

Private Sub Form_Load()
Dim i As Long
On Error GoTo ErrorTrap

    mbFormLoading = True

    HandleListBox.PopulateListBox cmbAccount, _
                              "SELECT AccountID, AccountName " & _
                              "FROM tblBankAccounts ORDER BY 1 ", CMSDB, 0, "", False, 1
                              
    cmbAccount.ListIndex = 0
    
    UpperLimit = 11
    
    For i = year(Now) - GlobalParms.GetValue("YearsHistoryToInclude_Accounts", "NumVal") To year(Now) + 2
        cmbYear.AddItem i
    Next i
    
    If mlFormTaxYear = 0 Then
        cmbYear.text = CStr(GetFinancialYear(Now))
    Else
        cmbYear.text = CStr(mlFormTaxYear)
    End If

    ConstructGrid
    
    chkSuppressWarnings.value = IIf(GlobalParms.GetValue("AccountsWarningsDefault", "TrueFalse"), vbUnchecked, vbChecked)
    
    '
    'Populate cmbMonth
    '
    HandleListBox.PopulateListBox Me!cmbMonth, "SELECT MonthNum, " & _
                                               "       MonthName " & _
                                               "FROM tblMonthName " & _
                                               "ORDER BY OrderForFiscalYear ASC", _
                                               CMSDB, 0, "", False, 1
                                               
    optMonth = True
                                               
    HandleListBox.SelectItem cmbMonth, Month(Now)
    
    
    With cmbTranType
    .AddItem "All"
    .AddItem "Income"
    .AddItem "Expenditure"
    .ListIndex = 0
    End With
    
    mnuInterestAlert.Checked = GlobalParms.GetValue("AlertForMissingBankInterest", "TrueFalse")
    gbNewTransaction_SkipDesc = GlobalParms.GetValue("NewTransaction_SkipDesc", "TrueFalse")
    gsAccountTransferTranCode = GlobalParms.GetValue("AccountTransferTransactionCode", "AlphaVal")
        
    Set frmTranEntry = New frmTransactionEntry
    Set frmAccTransfer = New frmAccountTransfer
    Set frmTransactionTypes = New frmTranTypes
    Set frmSelectInOutType = New frmAccountInOutTypeSelect
    Set frmReg = New frmRegularTransactions
    Set frmTrnCdSel = New frmTranCodeSelect
    Set frmCal1 = New frmMiniCalendar
    Set frmCal2 = New frmMiniCalendar
    Set frmCal3 = New frmMiniCalendar
    
    mbFormLoading = False
    
    'cater for new field service group arrangement
    cmdBkGrp.Visible = (year(Now) = 2009)

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub


Public Sub ConstructGrid()
On Error GoTo ErrorTrap
Dim TheYear As Long, rs As Recordset, str As String, j As Long

    'Set and bold the column headings
    With flxAccounts
    
    .Cols = 5
    .Rows = 1
    
    .TextMatrix(0, 0) = "Date"
    .TextMatrix(0, 1) = "Description"
    .TextMatrix(0, 2) = "Amount In"
    .TextMatrix(0, 3) = "Amount Out"
    .TextMatrix(0, 4) = "TranID"
        
    .Row = 0
    For j = 0 To .Cols - 1
        .col = j
        .CellFontBold = True
    Next j
    
    .ColWidth(4) = 0
    .ColWidth(0) = 1050
    .ColWidth(1) = 6000
    .ColWidth(2) = 1100
    .ColWidth(3) = 1100
    
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter
           
    End With

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub
Public Sub GoToCurrentMonth()
On Error GoTo ErrorTrap

    If mlFormTaxYear = 0 Then
        cmbYear.text = CStr(GetFinancialYear(Now))
    Else
        cmbYear.text = CStr(mlFormTaxYear)
    End If
    
    HandleListBox.SelectItem cmbMonth, Month(Now)

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Public Sub SortRowColour()
Dim i As Integer, NewRow As Integer, SaveCol As Long

On Error GoTo ErrorTrap

'
'Set clicked row to red, previous row back to black
'
    With flxAccounts
    
    
    If .Row > 0 Then
    
        SaveCol = .col
        
        If mlPrevRow < .Rows Then  'could be that prevrow has been deleted
            NewRow = .Row
            .Row = mlPrevRow
            
            For i = 0 To .Cols - 1
                .col = i
                .CellForeColor = QBColor(0) 'Change previous row to black
            Next i
            
            .Row = NewRow
        End If
    
        For i = 0 To .Cols - 1
            .col = i
            .CellForeColor = QBColor(12) 'Bright Red text
        Next i
        
        mlPrevRow = .Row
        
        .col = SaveCol
                    
    End If
    
    End With
    

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Public Sub GridToBlack()
Dim i As Integer, j As Integer
    On Error GoTo ErrorTrap

    'set each cell text to black
    With flxAccounts
    .Row = 0
    .RowSel = .Rows - 1
    .col = 0
    .ColSel = .Cols - 1
    .CellForeColor = QBColor(0)
    End With
    
    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub
Public Sub FillAccountsGrid(Optional SavePos As Boolean = False)
On Error GoTo ErrorTrap
Dim TheYear As Long, rs As Recordset, str As String, str2 As String
Dim i As Long, j As Long, store As Long, lCong As Long, lRef As Long
Dim sDateSQL As String, sZeroAmountSQL As String, lAccountID As Long

    With flxAccounts
        
    If SavePos Then
        If .Rows > 1 Then
            store = .TopRow
        Else
            store = 1
        End If
    End If
    
    lAccountID = mlAccountID
    
    .Rows = 1
        
    Select Case True
    Case optMonth
    
        If cmbYear.ListIndex = -1 Or cmbMonth.ListIndex = -1 Then
            Exit Sub
        End If
        
        sDateSQL = "WHERE FinancialYear = " & CLng(cmbYear.text) & _
                   " AND Month(TranDate) = " & cmbMonth.ItemData(cmbMonth.ListIndex)
                   
    Case optDateRange
    
        sDateSQL = "WHERE TranDate BETWEEN " & GetDateStringForSQLWhere(msStartDate) & _
                   " AND " & GetDateStringForSQLWhere(msEndDate)

    End Select
    
    If GlobalParms.GetValue("ShowZeroAmountTransactions", "TrueFalse") Then
        sZeroAmountSQL = " "
    Else
        sZeroAmountSQL = " AND a.Amount <> 0 "
    End If
    
    str = "SELECT a.TranCodeID, " & _
         "a.TranID, " & _
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
         "a.AccountID " & _
         "FROM ((tblTransactionDates a " & _
         " INNER JOIN tblTransactionTypes b ON a.TranCodeID = b.TranCodeID) " & _
         " INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
         " INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID " & _
          sDateSQL & sZeroAmountSQL & _
         " AND " & IIf(lAccountID = 0, " a.AccountID = ", "a.TfrAccountID = ") & lAccountID & _
         IIf(lAccountID = 0, " AND a.AccountID = 0 ", " ") & _
         " ORDER BY 3, 2"

    Set rs = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    i = 1
    Do Until rs.BOF Or rs.EOF
        .Rows = i + 1
        .TextMatrix(i, 0) = rs!TranDate
        
        lRef = HandleNull(rs!RefNo, "")
        .TextMatrix(i, 1) = HandleNull(rs!TranDescription) & _
                         " " & IIf(lRef = 0, "", "(" & lRef & ")")
        If rs!BookGroupNo > 0 Then
            .TextMatrix(i, 1) = .TextMatrix(i, 1) & " (" & GetGroupName(rs!BookGroupNo, "Congregation") & ")"
        End If
        
        If lAccountID = 0 Then
            .TextMatrix(i, 2) = IIf(rs!Amount >= 0, Format(rs!Amount, "0.00"), "")
            .TextMatrix(i, 3) = IIf(rs!Amount < 0, Format(-1 * rs!Amount, "0.00"), "")
        Else
            If rs!AccountID >= 0 Then
                .TextMatrix(i, 3) = IIf(rs!Amount >= 0, Format(rs!Amount, "0.00"), "")
                .TextMatrix(i, 2) = IIf(rs!Amount < 0, Format(-1 * rs!Amount, "0.00"), "")
            Else
                .TextMatrix(i, 2) = IIf(rs!Amount >= 0, Format(rs!Amount, "0.00"), "")
                .TextMatrix(i, 3) = IIf(rs!Amount < 0, Format(-1 * rs!Amount, "0.00"), "")
            End If
        End If
            
        .TextMatrix(i, 4) = rs!TranID
        
        rs.MoveNext
        i = i + 1
    Loop
    
    If SavePos Then
    
        If .Rows > 1 Then
            .TopRow = store
        End If
        
    End If
    
    End With
    
    RowShadingGroups flxAccounts, 0, vbWhite, RGB(240, 240, 240)

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Public Sub FillAccountsGridSearch(Optional SavePos As Boolean = False, Optional OrderByCol As Long = -1)
On Error GoTo ErrorTrap
Dim TheYear As Long, rs As Recordset, str As String, str2 As String
Dim i As Long, j As Long, store As Long, lCong As Long, lRef As Long
Dim sDateSQL1 As String, sDateSQL2 As String, sDescSQL As String
Dim sMinAmount As String, sMaxAmount As String, sTranTypeSQL As String
Dim dIncome As Double, dExpense As Double, dTotal As Double, arr() As String
Dim sOrderBySQL As String, sTranCodeSQL As String, sRefSQL1 As String, sRefSQL2 As String
Static bOrderSwitch As Boolean, sOrderAscDesc As String
Dim sTranSubCodeSQL As String, sAccountIDSQL As String
Dim sBkGrpSQL As String, sZeroAmountSQL As String

    With flxAccounts
        
    If SavePos Then
        If .Rows > 1 Then
            store = .TopRow
        Else
            store = 1
        End If
    End If
    
    .Rows = 1
                
    If ValidDate(txtSearchDate1) Then
        sDateSQL1 = " AND TranDate >= " & GetDateStringForSQLWhere(txtSearchDate1)
    Else
        sDateSQL1 = ""
        txtSearchDate1 = ""
    End If
    If ValidDate(txtSearchDate2) Then
        sDateSQL2 = " AND TranDate <= " & GetDateStringForSQLWhere(txtSearchDate2)
    Else
        sDateSQL2 = ""
        txtSearchDate2 = ""
    End If
    If IsNumber(txtMinAmount, True, False, True) Then
        sMinAmount = " AND ABS(a.Amount) >= " & txtMinAmount
    Else
        sMinAmount = ""
        txtMinAmount = ""
    End If
    If IsNumber(txtMaxAmount, True, False, True) Then
        sMaxAmount = " AND ABS(a.Amount) <= " & txtMaxAmount
    Else
        sMaxAmount = ""
        txtMaxAmount = ""
    End If
    If IsNumber(txtMinRef, False, False, False) Then
        If CLng(txtMinRef) > 0 Then
            sRefSQL1 = " AND ABS(a.RefNo) >= " & CLng(txtMinRef)
        Else
            sRefSQL1 = ""
        End If
    Else
        sRefSQL1 = ""
    End If
    If IsNumber(txtMaxRef, False, False, False) Then
        If CLng(txtMaxRef) > 0 Then
            sRefSQL2 = " AND ABS(a.RefNo) <= " & CLng(txtMaxRef) & " AND ABS(a.RefNo) > 0 "
        Else
            sRefSQL2 = ""
        End If
    Else
        sRefSQL2 = ""
    End If
    If Trim(txtDesc) <> "" Then
        
        sDescSQL = " AND (LCASE(TranDescription) LIKE '*"
        arr() = Split(LCase(txtDesc), "|")
        For i = 0 To UBound(arr)
            sDescSQL = sDescSQL & DoubleUpSingleQuotes(arr(i)) & "*' "
            If i < UBound(arr) Then
                sDescSQL = sDescSQL & " OR TranDescription LIKE '*"
            End If
        Next i
        
        sDescSQL = sDescSQL & " OR LCASE(e.GroupName) LIKE '*"
        arr() = Split(LCase(txtDesc), "|")
        For i = 0 To UBound(arr)
            sDescSQL = sDescSQL & DoubleUpSingleQuotes(arr(i)) & "*' "
            If i < UBound(arr) Then
                sDescSQL = sDescSQL & " OR e.GroupName LIKE '*"
            End If
        Next i
        
        sDescSQL = sDescSQL & ") "
    
    Else
        sDescSQL = ""
        txtDesc = ""
    End If
    Select Case cmbTranType.ListIndex
    Case 0
        sTranTypeSQL = " AND d.InOutID IN (1,2) "
    Case 1
        sTranTypeSQL = " AND d.InOutID = 1 "
    Case 2
        sTranTypeSQL = " AND d.InOutID = 2 "
    End Select
    
    If msTranCodeSearch <> "" Then
        If Not mbExcludeTranCodes Then
            sTranCodeSQL = " AND b.TranCodeID IN (" & msTranCodeSearch & ") "
        Else
            sTranCodeSQL = " AND b.TranCodeID NOT IN (" & msTranCodeSearch & ") "
        End If
    Else
        sTranCodeSQL = ""
    End If
    
    If msTranSubCodeSearch <> "" Then
        If Not mbExcludeTranSubCodes Then
            sTranSubCodeSQL = " AND a.TranSubTypeID IN (" & msTranSubCodeSearch & ") "
        Else
            sTranSubCodeSQL = " AND a.TranSubTypeID NOT IN (" & msTranSubCodeSearch & ") "
        End If
    Else
        sTranSubCodeSQL = ""
    End If
    
    If msBkGrpSearch <> "" Then
        If Not mbExcludeBookGroups Then
            sBkGrpSQL = " AND a.BookGroupNo IN (" & msBkGrpSearch & ") "
        Else
            sBkGrpSQL = " AND a.BookGroupNo NOT IN (" & msBkGrpSearch & ") "
        End If
    Else
        sBkGrpSQL = ""
    End If
    
    If OrderByCol > -1 Then
        bOrderSwitch = Not bOrderSwitch
    End If
    sOrderAscDesc = IIf(bOrderSwitch, " ASC ", " DESC ")
    Select Case OrderByCol
    Case -1
         sOrderBySQL = " ORDER BY 3 " & sOrderAscDesc & ", 2 "
    Case 0
         sOrderBySQL = " ORDER BY 3 " & sOrderAscDesc & ", 2 "
    Case 1
         sOrderBySQL = " ORDER BY 5 " & sOrderAscDesc & ", 2 "
    Case 2, 3
         sOrderBySQL = " ORDER BY 15 " & sOrderAscDesc & ", 2 "
    Case Else
         sOrderBySQL = " ORDER BY 3 " & sOrderAscDesc & ", 2 "
    End Select
    
    If GlobalParms.GetValue("ShowZeroAmountTransactions", "TrueFalse") Then
        sZeroAmountSQL = " "
    Else
        sZeroAmountSQL = " AND a.Amount <> 0 "
    End If
    
    sAccountIDSQL = " AND a.AccountID IN (0) "
    
    str = "SELECT TOP " & glMaxResultRows & " a.TranCodeID, " & _
         "a.TranID, " & _
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
         "BookgroupNo, " & _
         "ABS(a.Amount) AS AbsAmount " & _
         "FROM (((tblTransactionDates a " & _
         " INNER JOIN tblTransactionTypes b ON a.TranCodeID = b.TranCodeID) " & _
         " INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
         " INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
         " LEFT JOIN tblBookGroups e ON e.GroupNo = a.BookGroupNo " & _
         " WHERE 1=1 " & _
          sDateSQL1 & sDateSQL2 & sMinAmount & sMaxAmount & sDescSQL & sTranTypeSQL & _
           sRefSQL1 & sRefSQL2 & sTranCodeSQL & sTranSubCodeSQL & sBkGrpSQL & sZeroAmountSQL & _
           sAccountIDSQL & _
          sOrderBySQL

        ' DON'T FORGET TO MODIFY TOTALS CALCULATIONS BELOW!!!
        
    Set rs = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    If rs.BOF Then
        ShowMessage "No matching transactions found", 1250, Me
    End If
    
    i = 1
    Do Until rs.BOF Or rs.EOF
        .Rows = i + 1
        .TextMatrix(i, 0) = rs!TranDate
        
        lRef = HandleNull(rs!RefNo, "")
        .TextMatrix(i, 1) = HandleNull(rs!TranDescription) & _
                         " " & IIf(lRef = 0, "", "(" & lRef & ")")
        If rs!BookGroupNo > 0 Then
            .TextMatrix(i, 1) = .TextMatrix(i, 1) & " (" & GetGroupName(rs!BookGroupNo, "Congregation") & ")"
        End If
        .TextMatrix(i, 2) = IIf(rs!Amount >= 0, Format(rs!Amount, "0.00"), "")
        .TextMatrix(i, 3) = IIf(rs!Amount < 0, Format(-1 * rs!Amount, "0.00"), "")
        .TextMatrix(i, 4) = rs!TranID
        
        rs.MoveNext
        i = i + 1
    Loop
    
    If SavePos Then
    
        If .Rows > 1 Then
            .TopRow = store
        End If
        
    End If
    
    End With
    
    RowShadingGroups flxAccounts, 0, vbWhite, RGB(240, 240, 240)
    
    'show totals...
    
    Select Case cmbTranType.ListIndex
    Case 0, 1
        str = "SELECT SUM(a.Amount) AS TotalIncome " & _
             "FROM (((tblTransactionDates a " & _
             " INNER JOIN tblTransactionTypes b ON a.TranCodeID = b.TranCodeID) " & _
             " INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
             " INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
             " LEFT JOIN tblBookGroups e ON e.GroupNo = a.BookGroupNo " & _
             " WHERE 1=1 " & _
              sDateSQL1 & sDateSQL2 & sMinAmount & sMaxAmount & sDescSQL & sTranCodeSQL & _
              sTranSubCodeSQL & sRefSQL1 & sRefSQL2 & sBkGrpSQL & sAccountIDSQL & _
              " AND c.InOutTypeID IN (1, 3) "
    
        Set rs = CMSDB.OpenRecordset(str, dbOpenDynaset)
        
        dIncome = CDbl(HandleNull(rs!TotalIncome))
    Case Else
        dIncome = 0
    End Select

    Select Case cmbTranType.ListIndex
    Case 0, 2
        str = "SELECT SUM(a.Amount) AS TotalExpense " & _
             "FROM (((tblTransactionDates a " & _
             " INNER JOIN tblTransactionTypes b ON a.TranCodeID = b.TranCodeID) " & _
             " INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
             " INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
             " LEFT JOIN tblBookGroups e ON e.GroupNo = a.BookGroupNo " & _
             " WHERE 1=1 " & _
              sDateSQL1 & sDateSQL2 & sMinAmount & sMaxAmount & sDescSQL & sTranCodeSQL & _
              sRefSQL1 & sRefSQL2 & sTranSubCodeSQL & sBkGrpSQL & sAccountIDSQL & _
              " AND c.InOutTypeID IN (2, 4) "
    
        Set rs = CMSDB.OpenRecordset(str, dbOpenDynaset)
        
        dExpense = CDbl(HandleNull(rs!TotalExpense))
    Case Else
        dExpense = 0
    End Select
    
    dTotal = dExpense + dIncome
    
    lblOpeningBal = "-"
    lblIncome = Format(dIncome, "£0.00")
    lblExpense = Format(Abs(dExpense), "£0.00")
    lblSurplus = "-"
    lblClosingBal = "-"
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub


Private Sub flxAccounts_Click()

On Error GoTo ErrorTrap
    
    With flxAccounts
    If .Row > 0 Then
        mlTranID = CLng(.TextMatrix(.Row, 4))
    Else
        mlTranID = 0
    End If
    End With
        
    If Not mbOrderingCols Then
        SortRowColour
    Else
        mbOrderingCols = False
    End If

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub
Private Sub flxAccounts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorTrap
Dim TheRow As Long
'
'X and Y are relative to top left of the flxGrid.
'

'    If optSearch Then Exit Sub
    
    If cmbYear.ListIndex > -1 And cmbMonth.ListIndex > -1 Then
        '
        'Use current Y position to work out which row of the grid has beeen clicked
        '
        With flxAccounts
        If Button = vbRightButton Then
                        
            TheRow = (Ceiling(CDbl(Y) / .RowHeight(0))) + .TopRow - 2
            If TheRow <= .Rows - 1 Then
                .Row = TheRow
                flxAccounts_Click
                mnuEditTran.Enabled = (mlTranID <> 0)
                mnuDelete.Enabled = (mlTranID <> 0)
                mnuMove.Enabled = (mlTranID <> 0)
                Me.PopupMenu mnuActions2
            End If
            
        Else
            If Button = vbLeftButton Then
                TheRow = (Ceiling(CDbl(Y) / .RowHeight(0))) - 1
                If TheRow = 0 Then
                    If optSearch And .Rows > 1 Then
                        mbOrderingCols = True
                        FillAccountsGridSearch , .col
                    End If
                End If
            End If
        End If

        End With
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Public Property Get PreviousRow() As Long
    PreviousRow = mlPrevRow
End Property

Public Property Let PreviousRow(ByVal NewValue As Long)
    mlPrevRow = NewValue
End Property


Public Property Get FormTaxYear() As Long
    FormTaxYear = mlFormTaxYear
End Property

Public Property Let FormTaxYear(ByVal vNewValue As Long)
    mlFormTaxYear = vNewValue
End Property
Public Property Get FormNormalYear() As Long
    FormNormalYear = mlFormNormalYear
End Property

Public Property Let FormNormalYear(ByVal vNewValue As Long)
    mlFormNormalYear = vNewValue
End Property
Public Property Get FormMonth() As Long
    FormMonth = mlFormMonth
End Property

Public Property Let FormMonth(ByVal vNewValue As Long)
    mlFormMonth = vNewValue
End Property
Public Property Get TranID() As Long
    TranID = mlTranID
End Property

Public Property Let TranID(ByVal vNewValue As Long)
    mlTranID = vNewValue
End Property
Public Property Get AccountID() As Long
    AccountID = mlAccountID
End Property

Public Property Let AccountID(ByVal vNewValue As Long)
    mlAccountID = vNewValue
End Property

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    mlFormTaxYear = 0
    
    Set frmTranEntry = Nothing
    Set frmAccTransfer = Nothing
    Set frmTransactionTypes = Nothing
    Set frmSelectInOutType = Nothing
    Set frmReg = Nothing
    Set frmCal1 = Nothing
    Set frmCal2 = Nothing
    Set frmCal3 = Nothing
    Set frmTrnCdSel = Nothing
    
    BringForwardMainMenuWhenItsTheLastFormOpen
    
End Sub

Private Sub frmAccTransfer_TransactionEntered()

On Error GoTo ErrorTrap

    Select Case True
    Case optDateRange Or optMonth
        FillAccountsGrid True
        GetTotals
    Case optSearch
        FillAccountsGridSearch True
    End Select

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub frmCal1_InsertDate(TheDate As String)
    
    Select Case True
    Case optMonth
    Case optDateRange
        If ValidDate(TheDate) Then
            txtFirstDate = TheDate
        Else
            txtFirstDate = ""
        End If
    Case optSearch
        If ValidDate(TheDate) Then
            txtSearchDate1 = TheDate
        Else
            txtSearchDate1 = ""
        End If
    End Select
        
End Sub
Private Sub frmCal2_InsertDate(TheDate As String)
    Select Case True
    Case optMonth
    Case optDateRange
        If ValidDate(TheDate) Then
            txtLastDate = TheDate
        Else
            txtLastDate = ""
        End If
    Case optSearch
        If ValidDate(TheDate) Then
            txtSearchDate2 = TheDate
        Else
            txtSearchDate2 = ""
        End If
    End Select
End Sub

Private Sub frmCal3_InsertDate(TheDate As String)
On Error GoTo ErrorTrap
Dim rs As Recordset

    With flxAccounts
    If .Row < 1 Or .Row > .Rows - 1 Then
        ShowMessage "Cannot move transaction", 1250, Me, , vbRed, True
        Exit Sub
    End If
    If Not IsNumber(.TextMatrix(.Row, 4), False, False, False) Then
        ShowMessage "Cannot move transaction", 1250, Me, , vbRed, True
        Exit Sub
    End If
    If CLng(.TextMatrix(.Row, 4)) <= 0 Then
        ShowMessage "Cannot move transaction", 1250, Me, , vbRed, True
        Exit Sub
    End If
    If mlTranID <= 0 Then
        ShowMessage "Cannot move transaction", 1250, Me, , vbRed, True
        Exit Sub
    End If
    End With
    
    If GetAccountStartDate(0) > _
        CDate(TheDate) Then
        ShowMessage "Transaction Date should not be prior to " & _
                  GetAccountStartDate(0), 2000, Me, , vbRed, True
        Exit Sub
    End If
        
    Set rs = CMSDB.OpenRecordset("tblTransactionDates", dbOpenDynaset)
    
    With rs
    
    .FindFirst "TranID = " & mlTranID
    .Edit
    !TranDate = CDate(TheDate)
    .Update
    
    End With
    
    rs.Close
    Set rs = Nothing
    
    Select Case True
    Case optDateRange Or optMonth
        FillAccountsGrid True
        GetTotals
    Case optSearch
        FillAccountsGridSearch True
    End Select
    
    ShowMessage "Transaction moved to " & TheDate, 1750, Me, , , True
    
'    Me.SetFocus
        
    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub frmMissingReceiptsList_ListGenerated(StuffList As String)
Dim rs As Recordset, i As Long, arr() As String, sDate As String
On Error GoTo ErrorTrap
    
    arr() = Split(StuffList, ",")
    
    For i = 0 To UBound(arr)
    
        Set rs = CMSDB.OpenRecordset("tblTransactionDates", dbOpenDynaset)
        
        sDate = arr(i)
        With rs
        
        .AddNew
        
        !TranCodeID = GetTranCodeIDFromTranCode(GlobalParms.GetValue("CongContributionTransactionCode", "AlphaVal"), 1)
        !TranDescription = "ZERO ENTRY"
        !RefNo = 0
        !Amount = 0
        !TranDate = CDate(sDate)
        !FinancialYear = GetFinancialYear(CDate(sDate))
        !FinancialMonth = GetFinancialMonth(CDate(sDate))
        !FinancialQuarter = GetFinancialQuarter(CDate(sDate))
        !BookGroupNo = -1
        !TranSubTypeID = 0
        !AccountID = 0
        !TfrAccountID = -1
        
        .Update
        
        End With
            
    Next i
    
    If i > 0 Then
        ShowMessage "Missing receipts processed", 2000, frmMissingReceiptsList
    Else
        ShowMessage "Nothing processed", 2000, frmMissingReceiptsList
    End If
        

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub frmReg_TransSelected(colTrans As Collection)
Dim rs As Recordset, TranStuff As TransactionDetails, i As Long
Dim sDate As String, sMsg As String, bCannotContinue As Boolean
On Error GoTo ErrorTrap

    bCannotContinue = False
      
    For i = 1 To colTrans.Count
        TranStuff = GetTransactionCodeStuff(CLng(colTrans(i)))
        
        Set rs = CMSDB.OpenRecordset("tblTransactionDates", dbOpenDynaset)
        
        With rs
              
        sDate = TranStuff.AutoDayOfMonth & "/" & mlFormMonth & "/" & mlFormNormalYear
        sDate = Format(sDate, "dd/mm/yyyy")
        If Not ValidDate(sDate) Then
            sDate = "28/" & mlFormMonth & "/" & mlFormNormalYear
        End If
        
        If TranStuff.TransactionCode = gsAccountTransferTranCode Then
            If GetAccountStartDate(TranStuff.AccountID) > CDate(sDate) Or _
                GetAccountStartDate(TranStuff.TfrAccountID) > CDate(sDate) Then
                
                sMsg = sMsg & TranStuff.TransactionTypeDescription & " (£" & _
                        Abs(TranStuff.Amount) & ") has date prior to account start date." & vbCrLf & vbCrLf
                        
                bCannotContinue = True
            End If
        Else
            If GetAccountStartDate(TranStuff.AccountID) > CDate(sDate) Then
                sMsg = sMsg & TranStuff.TransactionTypeDescription & " (£" & _
                        Abs(TranStuff.Amount) & ") has date prior to account start date." & vbCrLf & vbCrLf
                
                bCannotContinue = True
            End If
        End If
        
        sDate = Format(sDate, "mm/dd/yyyy")
        
        .FindFirst "TranDate = #" & sDate & "# AND TranCodeID = " & TranStuff.TransactionCodeID
        
        If Not .NoMatch Then
            sMsg = sMsg & TranStuff.TransactionTypeDescription & " (£" & _
                    Abs(TranStuff.Amount) & ") already entered." & vbCrLf & vbCrLf
        End If
        
        End With
        
    Next i
      
    If sMsg <> "" Then
        If Not bCannotContinue Then
            If MsgBox(sMsg & "Continue?", vbYesNo + vbQuestion + vbDefaultButton2, AppName) = vbYes Then
            Else
                GoTo TidyUp
            End If
        Else
            MsgBox sMsg, vbOKOnly + vbExclamation, AppName
            GoTo TidyUp
        End If
    End If
    
    For i = 1 To colTrans.Count
    
        TranStuff = GetTransactionCodeStuff(CLng(colTrans(i)))
        
        Set rs = CMSDB.OpenRecordset("tblTransactionDates", dbOpenDynaset)
        
        sDate = TranStuff.AutoDayOfMonth & "/" & mlFormMonth & "/" & mlFormNormalYear
        sDate = Format(sDate, "dd/mm/yyyy")
        If Not ValidDate(sDate) Then
            sDate = "28/" & mlFormMonth & "/" & mlFormNormalYear
            sDate = Format(sDate, "dd/mm/yyyy")
        End If
        
        If (TranStuff.TransactionCode <> GlobalParms.GetValue("GiftAidTransactionCode", "AlphaVal")) Or _
            (TranStuff.TransactionCode = GlobalParms.GetValue("GiftAidTransactionCode", "AlphaVal") And _
                                             GiftAidNoActiveForDate(TranStuff.RefNo, CDate(sDate))) Then
        
            With rs
            
            .AddNew
            
            !TranCodeID = TranStuff.TransactionCodeID
            !TranDescription = TranStuff.TransactionTypeDescription
            !RefNo = TranStuff.RefNo
            !Amount = TranStuff.Amount
            !TranDate = CDate(sDate)
            !FinancialYear = GetFinancialYear(CDate(sDate))
            !FinancialMonth = GetFinancialMonth(CDate(sDate))
            !FinancialQuarter = GetFinancialQuarter(CDate(sDate))
            !BookGroupNo = -1
            !TranSubTypeID = 0
            !AccountID = TranStuff.AccountID
            !TfrAccountID = TranStuff.TfrAccountID
            
            .Update
            
            End With
            
        End If
        
    Next i
    
TidyUp:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    On Error GoTo ErrorTrap
    
    FillAccountsGrid True
    GetTotals
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Private Sub frmTranEntry_TransactionEntered()

On Error GoTo ErrorTrap

    Select Case True
    Case optDateRange Or optMonth
        FillAccountsGrid True
        GetTotals
    Case optSearch
        FillAccountsGridSearch True
    End Select
    

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub GetTotals()
On Error GoTo ErrorTrap

    fraOtherAccounts.Visible = False
    
    Select Case True
    Case optDateRange
        GetPeriodTotals
    Case optMonth
        GetMonthTotals
    End Select
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Private Sub frmTrnCdSel_ReturnTranCodes(TranCodes As String, ExcludeTheCodes As Boolean)

On Error GoTo ErrorTrap

    Select Case mlCodeType
    Case 1
        msTranCodeSearch = TranCodes
        mbExcludeTranCodes = ExcludeTheCodes
    Case 2
        msTranSubCodeSearch = TranCodes
        mbExcludeTranSubCodes = ExcludeTheCodes
    Case 3
        msBkGrpSearch = TranCodes
        mbExcludeBookGroups = ExcludeTheCodes
    End Select
        
    TranCodeButtonText
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub TranCodeButtonText(Optional SuppressMsg As Boolean = False)

On Error GoTo ErrorTrap

    Select Case mlCodeType
    Case 1
        If msTranCodeSearch = "" Then
            cmdTranCodes.Caption = "Codes"
            cmdTranCodes.ToolTipText = "No Transaction Code criteria selected"
            ShowMessage "No Transaction Code criteria selected", 750, Me, SuppressMsg
        Else
            cmdTranCodes.Caption = "* Codes"
            cmdTranCodes.ToolTipText = "Transaction Code criteria selected"
            ShowMessage "Transaction Code criteria selected", 750, Me, SuppressMsg
        End If
    Case 2
        If msTranSubCodeSearch = "" Then
            cmdTranSubCodes.Caption = "Sub Cd"
            cmdTranSubCodes.ToolTipText = "No Transaction Sub-Type criteria selected"
            ShowMessage "No Transaction Sub-Type criteria selected", 750, Me, SuppressMsg
        Else
            cmdTranSubCodes.Caption = "* Sub Cd"
            cmdTranSubCodes.ToolTipText = "Transaction Sub-Type criteria selected"
            ShowMessage "Transaction Sub-Type criteria selected", 750, Me, SuppressMsg
        End If
    Case 3
        If msBkGrpSearch = "" Then
            cmdBkGrp.Caption = "Bk Grp"
            cmdBkGrp.ToolTipText = "No Book-Group criteria selected"
            ShowMessage "No Book-Group criteria selected", 750, Me, SuppressMsg
        Else
            cmdBkGrp.Caption = "* Bk Grp"
            cmdBkGrp.ToolTipText = "Book-Group criteria selected"
            ShowMessage "Book-Group criteria selected", 750, Me, SuppressMsg
        End If
    End Select
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub mnuAccTfr_Click()
Dim dte As Date
On Error GoTo ErrorTrap

    If Not optMonth Then
        ShowMessage "Account transfers allowed only in display mode 'Tax Year & Month'", 2000, Me
        Exit Sub
    End If
    
    If GetDatabaseTableScalar("SELECT COUNT(*) FROM tblBankAccounts") = 1 Then
        ShowMessage "No other accounts set up for transfers", 1500, Me
        Exit Sub
    End If
    
    dte = GetAccountStartDate(0)
    
    dte = CDate("01/" & Month(dte) & "/" & year(dte))

    If DateDiff("m", dte, CDate("01/" & mlFormMonth & "/" & mlFormNormalYear)) <= 0 Then
        dte = DateAdd("m", 1, dte)
        MsgBox "Cannot enter transactions earlier than " & GetMonthName(Month(dte)) & " " & year(dte), vbOKOnly + vbExclamation, AppName
    Else
        frmAccTransfer.TransactionID = 0
        frmAccTransfer.AccountID = mlAccountID
        frmAccTransfer.RegularTran = False
        frmAccTransfer.Show vbModal, Me
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub mnuAddRegular_Click()
Dim dte As Date
On Error GoTo ErrorTrap
    
    If Not optMonth Then
        ShowMessage "Cannot add regular transactions in this view", 1500, Me
        Exit Sub
    End If

    dte = GetAccountStartDate(0)
    
    dte = CDate("01/" & Month(dte) & "/" & year(dte))

    If DateDiff("m", dte, CDate("01/" & mlFormMonth & "/" & mlFormNormalYear)) <= 0 Then
        dte = DateAdd("m", 1, dte)
        MsgBox "Cannot enter transactions earlier than " & GetMonthName(Month(dte)) & " " & year(dte), vbOKOnly + vbExclamation, AppName
    Else
        frmReg.Show vbModal, Me
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub mnuAddRegular2_Click()

On Error GoTo ErrorTrap

    mnuAddRegular_Click

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub mnuCalculator_Click()
On Error GoTo ErrorTrap

    Shell "calc", vbNormalFocus

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub mnuDelete_Click()

On Error GoTo ErrorTrap

    If MsgBox("Delete selected transaction?", vbYesNo + vbDefaultButton2 + vbQuestion, AppName) = vbNo Then
        Exit Sub
    End If
    
    CMSDB.Execute "DELETE FROM tblTransactionDates WHERE TranID = " & mlTranID
    
    Select Case True
    Case optDateRange Or optMonth
        FillAccountsGrid True
        GetTotals
    Case optSearch
        FillAccountsGridSearch True
    End Select
    

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub mnuEditTran_Click()
Dim trn As TransactionDetails
On Error GoTo ErrorTrap

    trn = GetTransactionDetails(mlTranID)

    If trn.TfrAccountID > 0 And trn.AccountID = 0 Then
        frmAccTransfer.TransactionID = mlTranID
        frmAccTransfer.RegularTran = False
        frmAccTransfer.Show vbModal, Me
    Else
        frmTranEntry.NonCurrentAccount = (trn.TfrAccountID > 0)
        frmTranEntry.AccountID = mlAccountID
        frmTranEntry.TransactionID = mlTranID
        frmTranEntry.RegularTran = False
        frmTranEntry.Show vbModal, Me
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub mnuExport_Click()

On Error GoTo ErrorTrap

    If flxAccounts.Rows < 2 Then
        MsgBox "Nothing to export", vbOKCancel + vbExclamation, AppName
        Exit Sub
    End If
    
    ExportFlexGridToCSV gsDocsDirectory, _
                        "Congregation Accounts", _
                        "csv", _
                        "C.M.S. Export Accounts to CSV File", _
                        flxAccounts, _
                        CommonDialog1

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub mnuGiftAid_Click()

On Error GoTo ErrorTrap

    frmGiftAid.Show vbModal, Me

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub mnuInterestAlert_Click()

On Error GoTo ErrorTrap

    If GlobalParms.GetValue("AlertForMissingBankInterest", "TrueFalse") Then
        mnuInterestAlert.Checked = False
        GlobalParms.Save "AlertForMissingBankInterest", "TrueFalse", False
    Else
        mnuInterestAlert.Checked = True
        GlobalParms.Save "AlertForMissingBankInterest", "TrueFalse", True
    End If
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub mnuMenu_Click()

On Error GoTo ErrorTrap
    
    mnuExport.Enabled = (flxAccounts.Rows > 1)

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub mnuMissingTrans_Click()
Dim str As String, str1 As String, str2 As String, str3 As String
On Error GoTo ErrorTrap

    str1 = CheckForMissingInterest(AlwaysCheckForInterest:=True)
    str2 = CheckForMissingReceipts
    If str2 <> "" Then
        str2 = Split(str2, "|")(0)
    Else
        str2 = ""
    End If
    str3 = CheckForMissingRegularPayments
    
    str = IIf(str1 = "", "", str1 & vbCrLf & vbCrLf) & _
          IIf(str2 = "", "", str2 & vbCrLf & vbCrLf) & _
          str3
    
    If str <> "" Then
        MsgBox str, vbOKOnly + vbInformation, AppName & " - Accounts"
    Else
        MsgBox "No missing transactions for " & GetMonthName(mlFormMonth) & " " & mlFormNormalYear, _
         vbOKOnly + vbInformation, AppName & " - Accounts"
        
    End If
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub mnuMove_Click()
On Error GoTo ErrorTrap


    If mlTranID <= 0 Then
        ShowMessage "Cannot move transaction", 1250, Me, , vbRed
        Exit Sub
    End If
    If Not ValidDate(flxAccounts.TextMatrix(flxAccounts.Row, 0)) Then
        ShowMessage "Cannot move transaction", 1250, Me, , vbRed
        Exit Sub
    End If
    
    On Error Resume Next
    Unload frmCal3
    Set frmCal3 = Nothing
    If FormIsOpen("frmMiniCalendar") Then Exit Sub
    Set frmCal3 = New frmMiniCalendar
    On Error GoTo ErrorTrap
    
    With frmCal3
    
    .SetPos = True
    .XPos = Me.Left + fraBigFrame.Left + flxAccounts.Left + 1050
    .YPos = Me.Top + fraBigFrame.Top + flxAccounts.Top + flxAccounts.CellTop
    .FormDate = flxAccounts.TextMatrix(flxAccounts.Row, 0)
    .Show vbModeless, Me

    End With

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub mnuNewTran_Click()
Dim dte As Date
On Error GoTo ErrorTrap

    dte = GetAccountStartDate(0)
    
    dte = CDate("01/" & Month(dte) & "/" & year(dte))

    If DateDiff("m", dte, CDate("01/" & mlFormMonth & "/" & mlFormNormalYear)) <= 0 Then
        dte = DateAdd("m", 1, dte)
        MsgBox "Cannot enter transactions earlier than " & GetMonthName(Month(dte)) & " " & year(dte), vbOKOnly + vbExclamation, AppName
    Else
        frmTranEntry.TransactionID = 0
        frmTranEntry.RegularTran = False
        frmTranEntry.NonCurrentAccount = (mlAccountID > 0)
        frmTranEntry.AccountID = mlAccountID
        frmTranEntry.Show vbModal, Me
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub mnuNewTransaction_Click()

On Error GoTo ErrorTrap

    mnuNewTran_Click

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub mnuNewTransfer_Click()
    mnuAccTfr_Click
End Sub

Private Sub mnuOpeningBal_Click()

On Error GoTo ErrorTrap

    If MsgBox("Changes made here could affect ALL accounts totals. Are you sure you want to continue?", _
                vbYesNo + vbQuestion + vbDefaultButton2, AppName) = vbYes Then
        frmOpeningBalance.Show vbModal, Me
        GetTotals
        HandleListBox.Requery cmbAccount, True
    Else
        ShowMessage "Operation cancelled", 1000, Me
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub mnuOptions_Click()
On Error GoTo ErrorTrap

    frmAccountOptions.Show vbModal, Me

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub mnuReporting_Click()

On Error GoTo ErrorTrap

    frmAccountsReporting.Show vbModeless, Me

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub mnuTranTypes_Click()

On Error GoTo ErrorTrap

    frmTranTypeList.ViewMode = False
    frmTranTypeList.Show vbModal, Me

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Function CheckForMissingInterest(Optional AlwaysCheckForInterest As Boolean) As String
Dim rs As Recordset, str As String, sDateToCheck As String
Dim mlMonthToCheck As Long, mlYearToCheck As Long

On Error GoTo ErrorTrap

    If Not GlobalParms.GetValue("AlertForMissingBankInterest", "TrueFalse") Then
        CheckForMissingInterest = ""
        Exit Function
    End If
    
    If IsMissing(AlwaysCheckForInterest) Then
        AlwaysCheckForInterest = False
    End If

    'if form date later than current date, no need to check for missing interest
    If DateDiff("m", "01/" & Month(Now) & "/" & year(Now), "01/" & mlFormMonth & "/" & mlFormNormalYear) > 0 Then
        CheckForMissingInterest = ""
        Exit Function
    End If

    If Not AlwaysCheckForInterest Then
        'only check for missing interest transaction if we're in the last 7 days
        ' of the month.
        If mlFormMonth = Month(Now) And mlFormNormalYear = year(Now) And _
         Day(Now) < NoDaysInMonth(mlFormMonth, mlFormNormalYear) - 7 Then
            CheckForMissingInterest = ""
            Exit Function
        End If
    End If
        
    If AlwaysCheckForInterest Then
        mlMonthToCheck = mlFormMonth
        mlYearToCheck = mlFormNormalYear
    Else
        If CDate(DateSerial(CInt(mlFormNormalYear), mlFormMonth, 1)) < _
                        CDate(DateSerial(year(Now), Month(Now), 1)) Then
            
            mlMonthToCheck = mlFormMonth
            mlYearToCheck = mlFormNormalYear
        Else
            'check for last month's interest
            sDateToCheck = Format$(DateAdd("m", -1, Now), "dd/mm/yyyy")
            mlMonthToCheck = Month(sDateToCheck)
            mlYearToCheck = year(sDateToCheck)
        End If
    End If
    
    'are there any interest transactions in the specified month?
    str = "SELECT 1 " & _
          "FROM (tblTransactionDates a " & _
          "INNER JOIN tblTransactionTypes b ON a.TranCodeID = b.TranCodeID) INNER JOIN " & _
          "tblAccInOutTypes c ON c.InOutTypeID = b.InOutTypeID " & _
          "WHERE b.TranCode = '" & GlobalParms.GetValue("BankInterestTransactionCode", "AlphaVal") & "' " & _
          "AND c.InOutID = 1 " & _
          "AND Year(TranDate) = " & mlYearToCheck & _
          " AND Month(TranDate) = " & mlMonthToCheck & _
          " AND a.AccountID = 0 "
          
    Set rs = CMSDB.OpenRecordset(str, dbOpenSnapshot)
    
    If rs.BOF Then
        CheckForMissingInterest = "No bank interest entered for " & GetMonthName(mlMonthToCheck) & " " & mlYearToCheck
    Else
        CheckForMissingInterest = ""
    End If
    
    rs.Close
    Set rs = Nothing

    Exit Function
ErrorTrap:
    EndProgram

End Function
Private Function CheckForMissingRegularPayments() As String
Dim rs As Recordset, str As String, sDateToCheck As String
Dim mlMonthToCheck As Long, mlYearToCheck As Long

On Error GoTo ErrorTrap

    'if form date later than current date, no need to check for missing interest
    If DateDiff("m", "01/" & Month(Now) & "/" & year(Now), "01/" & mlFormMonth & "/" & mlFormNormalYear) > 0 Then
        CheckForMissingRegularPayments = ""
        Exit Function
    End If
    
    'any regular payments set up?
    str = "SELECT 1 " & _
          "FROM tblTransactionTypes " & _
          "WHERE InOutTypeID IN (3,4)"
          
    Set rs = CMSDB.OpenRecordset(str, dbOpenSnapshot)
    
    If rs.BOF Then
        CheckForMissingRegularPayments = ""
        Exit Function
    End If
    
    'are there any regular transactions in the specified month?
    str = "SELECT 1 " & _
          "FROM tblTransactionDates a " & _
          "INNER JOIN tblTransactionTypes b ON a.TranCodeID = b.TranCodeID " & _
          "WHERE b.InOutTypeID IN (3,4) " & _
          "AND Year(TranDate) = " & mlFormNormalYear & _
          " AND Month(TranDate) = " & mlFormMonth & _
          " AND a.AccountID = 0 " & _
          " AND b.Suppressed = FALSE "
    
    Set rs = CMSDB.OpenRecordset(str, dbOpenSnapshot)
    
    If rs.BOF Then
        CheckForMissingRegularPayments = "No regular payments entered for " & _
            GetMonthName(mlFormMonth) & " " & mlFormNormalYear & ". "
    Else
        CheckForMissingRegularPayments = ""
    End If
        
    rs.Close
    Set rs = Nothing

    Exit Function
ErrorTrap:
    EndProgram

End Function
Private Function CheckForMissingReceipts() As String
Dim rs As Recordset, str As String, i As Long, lSunMtgDay As Long, lMidMtgDay As Long
Dim sMsg As String, sMoreSQL As String, sTempDate As String, sDateList As String
On Error GoTo ErrorTrap

    If DateDiff("m", "01/" & Month(Now) & "/" & year(Now), "01/" & mlFormMonth & "/" & mlFormNormalYear) > 0 Then
        CheckForMissingReceipts = ""
        Exit Function
    End If
    
    lSunMtgDay = GlobalParms.GetValue("SundayMeetingDay", "NumVal")
    lMidMtgDay = GlobalParms.GetValue("MidWeekMeetingDay", "NumVal")
    sMoreSQL = GlobalParms.GetValue("TransactionCodesForCongContribs", "AlphaVal")
    sMoreSQL = "'" & Replace(sMoreSQL, ",", "','") & "'" 'put single quotes around each tran-code for the SQL
    
    str = "SELECT a.TranDate " & _
          "FROM (tblTransactionDates a " & _
          "INNER JOIN tblTransactionTypes b ON a.TranCodeID = b.TranCodeID) INNER JOIN " & _
          "tblAccInOutTypes c ON c.InOutTypeID = b.InOutTypeID " & _
          "WHERE c.InOutID = 1 " & _
          "AND Year(TranDate) = " & mlFormNormalYear & _
          " AND Month(TranDate) = " & mlFormMonth & _
          " AND b.TranCode IN (" & sMoreSQL & ") " & _
          " AND a.BookgroupNo <= 0 " & _
          " AND a.AccountID = 0 " & _
          " AND b.Suppressed = FALSE "
          
    Set rs = CMSDB.OpenRecordset(str, dbOpenSnapshot)
    
    For i = 1 To NoDaysInMonth(mlFormMonth, mlFormNormalYear)
        If mlFormMonth = Month(Now) And mlFormNormalYear = year(Now) And Day(Now) <= i Then
            Exit For
        End If
        
        sTempDate = DateSerial(CInt(mlFormNormalYear), CInt(mlFormMonth), CInt(i))
                
        Select Case True
        Case Weekday(sTempDate) = lSunMtgDay Or _
            Weekday(sTempDate) = lMidMtgDay Or _
            IsMemorialDay(CDate(sTempDate)) Or _
            (IsCOVisitWeek(GetDateOfGivenDay(CDate(sTempDate), vbMonday, False)) And _
                (Weekday(sTempDate) = vbTuesday Or Weekday(sTempDate) = vbThursday))
            
            
            If (IsCircuitOrDistrictAssemblyWeek(GetDateOfGivenDay(CDate(sTempDate), vbMonday, False)) And _
                 (Weekday(sTempDate) = lMidMtgDay Or Weekday(sTempDate) = lSunMtgDay)) Or _
                 ((IsCOVisitWeek(GetDateOfGivenDay(CDate(sTempDate), vbMonday, False)) And _
                    (Weekday(sTempDate) = lMidMtgDay And Weekday(sTempDate) <> vbTuesday And Weekday(sTempDate) <> vbThursday))) Then
            Else
                rs.FindFirst "TranDate = #" & Format(sTempDate, "mm/dd/yyyy") & "#"
                
                If rs.NoMatch Then
                    sMsg = sMsg & "No congregation contributions found for " & _
                        IIf(IsMemorialDay(CDate(sTempDate)), "memorial", "meeting") & " on " & _
                            Format(sTempDate, "dddd") & " " & _
                            Day(sTempDate) & GetLettersForOrdinalNumber(Day(sTempDate)) & " " & _
                            Format(sTempDate, "mmmm") & vbCrLf
                    
                    If sDateList = "" Then
                        sDateList = sTempDate
                    Else
                        sDateList = sDateList & "," & sTempDate
                    End If
                    
                End If
                
            End If
                    
        End Select
    Next i
    
    CheckForMissingReceipts = sMsg & IIf(sDateList <> "", "|", "") & sDateList 'sDateList used to construct list later
    
    rs.Close
    Set rs = Nothing

    Exit Function
ErrorTrap:
    EndProgram

End Function


Private Sub GetMonthTotals(Optional AccountID As Long = 0)
Dim fOpeningBalance As Double, fExpense As Double, fIncome As Double, fNum As Double
Dim fClosingBalance As Double, fSurplus As Double
Dim rs As Recordset, sSQL As String

On Error GoTo ErrorTrap

    'get balance at start of this month
    sSQL = "SELECT SUM(a.Amount) AS Tot " & _
          "FROM (tblTransactionDates a " & _
         " INNER JOIN tblTransactionTypes b ON a.TranCodeID = b.TranCodeID) " & _
         " INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID " & _
          "WHERE TranDate < #" & mlFormMonth & "/01/" & mlFormNormalYear & "# " & _
          "AND c.InOutTypeID IN (1,2,3,4) " & _
          " AND a.AccountID = " & AccountID
    
    Set rs = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    fNum = CDbl(HandleNull(rs!Tot))
    
    fOpeningBalance = GetAccountStartAmount(AccountID) + fNum
    
    'get expenditure this month
    sSQL = "SELECT SUM(a.Amount) AS Tot " & _
          "FROM (tblTransactionDates a " & _
         " INNER JOIN tblTransactionTypes b ON a.TranCodeID = b.TranCodeID) " & _
         " INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID " & _
          "WHERE month(TranDate) = " & mlFormMonth & " AND Year(TranDate) = " & mlFormNormalYear & _
          " AND c.InOutTypeID IN (2, 4) " & _
          " AND a.AccountID = " & AccountID
    
    Set rs = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    fExpense = CDbl(HandleNull(rs!Tot)) * -1
        
    'get income this month
    sSQL = "SELECT SUM(a.Amount) AS Tot " & _
          "FROM (tblTransactionDates a " & _
         " INNER JOIN tblTransactionTypes b ON a.TranCodeID = b.TranCodeID) " & _
         " INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID " & _
          "WHERE month(TranDate) = " & mlFormMonth & " AND Year(TranDate) = " & mlFormNormalYear & _
          " AND c.InOutTypeID IN (1, 3) " & _
          " AND a.AccountID = " & AccountID
    
    Set rs = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    fIncome = CDbl(HandleNull(rs!Tot))
    
    'Surplus
    fSurplus = fIncome - fExpense
    
    'month closing balance
    fClosingBalance = fOpeningBalance + fSurplus
    
    mfCurrentAccClosingBal = fClosingBalance
    
    lblOpeningBal = Format(fOpeningBalance, "£0.00")
    lblIncome = Format(fIncome, "£0.00")
    lblExpense = Format(Abs(fExpense), "£0.00")
    lblSurplus = Format(fSurplus, "£0.00")
    lblClosingBal = Format(fClosingBalance, "£0.00")

    rs.Close
    Set rs = Nothing
    
    GetMonthTotalsForOtherAccounts
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub GetMonthTotalsForOtherAccounts(Optional AccountID As Long = 0)
Dim dStartBal As Double, dAmount As Double, dStartBalTot As Double, fNum As Double
Dim dEndBal As Double, dEndBalTot As Double
Dim rs As Recordset, sSQL As String, rsStartBalOtherAccs As Recordset, rsEndBalOtherAccs As Recordset
Dim sBalanceOtherAccsText As String, bOtherAccsFound  As Boolean
Dim lAccountID As Long

On Error GoTo ErrorTrap
   
    'get balances for other accounts at end of period
    sSQL = "SELECT c.AccountID, SUM(b.Amount * iif(b.AccountID = -1, 1, -1)) AS  EndBalance " & _
           "FROM tblTransactionDates b " & _
           "RIGHT JOIN tblBankAccounts c ON b.TfrAccountID = c.AccountID " & _
           "WHERE b.TranDate <= #" & mlFormMonth & "/" & NoDaysInMonth(mlFormMonth, mlFormNormalYear) & "/" & mlFormNormalYear & "# " & _
           " AND  c.AccountID <> 0 " & _
           "GROUP BY c.AccountID " & _
           "UNION ALL " & _
           "SELECT AccountID, 0 " & _
           "FROM tblBankAccounts " & _
           "WHERE AccountID NOT IN " & _
           "(SELECT TfrAccountID " & _
           " FROM tblTransactionDates " & _
           " WHERE TfrAccountID > 0 " & _
           " AND TranDate <= #" & mlFormMonth & "/" & NoDaysInMonth(mlFormMonth, mlFormNormalYear) & "/" & mlFormNormalYear & "#) " & _
           "AND AccountID <> 0 "
         
    Set rsEndBalOtherAccs = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)

    'put in start/end balances and surplus for other accounts
    
    sBalanceOtherAccsText = vbCrLf
    Do Until (rsEndBalOtherAccs.BOF Or rsEndBalOtherAccs.EOF)
             
        bOtherAccsFound = True
        lAccountID = rsEndBalOtherAccs!AccountID
        dAmount = GetAccountStartAmount(lAccountID)
        dEndBal = HandleNull(rsEndBalOtherAccs!EndBalance)
        
        sBalanceOtherAccsText = sBalanceOtherAccsText & GetAccountName(lAccountID) & " closing balance: " & _
                                 RightAlignString(Format(dEndBal + dAmount, "£0.00"), 11) & vbCrLf
                
        dEndBalTot = dEndBalTot + dEndBal + dAmount
        
        rsEndBalOtherAccs.MoveNext
        
    Loop
    
    If bOtherAccsFound Then
        sBalanceOtherAccsText = sBalanceOtherAccsText & vbCrLf
            
        sBalanceOtherAccsText = sBalanceOtherAccsText & "TOTAL BALANCE AT END OF PERIOD:" & RightAlignString(Format(dEndBalTot + mfCurrentAccClosingBal, "£0.00"), 11)
        
        txtOtherAccounts = sBalanceOtherAccsText
        fraOtherAccounts.Visible = True
    Else
        fraOtherAccounts.Visible = False
    End If
    
    On Error Resume Next

    rsEndBalOtherAccs.Close
    Set rsEndBalOtherAccs = Nothing
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Private Sub GetPeriodTotals(Optional AccountID As Long = 0)
Dim fOpeningBalance As Double, fExpense As Double, fIncome As Double, fNum As Double
Dim fClosingBalance As Double, fSurplus As Double
Dim rs As Recordset, sSQL As String

On Error GoTo ErrorTrap

    'get balance at start of this period
    sSQL = "SELECT SUM(a.Amount) AS Tot " & _
          "FROM (tblTransactionDates a " & _
         " INNER JOIN tblTransactionTypes b ON a.TranCodeID = b.TranCodeID) " & _
         " INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID " & _
          "WHERE TranDate < " & GetDateStringForSQLWhere(msStartDate) & _
          " AND c.InOutTypeID IN (1,2,3,4) " & _
          " AND a.AccountID = " & AccountID
    
    Set rs = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    fNum = CDbl(HandleNull(rs!Tot))
    
    fOpeningBalance = GetAccountStartAmount(0) + fNum
    
    'get expenditure this period
    sSQL = "SELECT SUM(a.Amount) AS Tot " & _
          "FROM (tblTransactionDates a " & _
         " INNER JOIN tblTransactionTypes b ON a.TranCodeID = b.TranCodeID) " & _
         " INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID " & _
          "WHERE TranDate BETWEEN " & GetDateStringForSQLWhere(msStartDate) & _
                          " AND " & GetDateStringForSQLWhere(msEndDate) & _
          " AND c.InOutTypeID IN (2, 4) " & _
          " AND a.AccountID = " & AccountID
    
    Set rs = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    fExpense = CDbl(HandleNull(rs!Tot)) * -1
        
    'get income this period
    sSQL = "SELECT SUM(a.Amount) AS Tot " & _
          "FROM (tblTransactionDates a " & _
         " INNER JOIN tblTransactionTypes b ON a.TranCodeID = b.TranCodeID) " & _
         " INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID " & _
          "WHERE TranDate BETWEEN " & GetDateStringForSQLWhere(msStartDate) & _
                          " AND " & GetDateStringForSQLWhere(msEndDate) & _
          " AND c.InOutTypeID IN (1, 3) " & _
          " AND a.AccountID = " & AccountID
    
    Set rs = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    fIncome = CDbl(HandleNull(rs!Tot))
    
    'Surplus
    fSurplus = fIncome - fExpense
    
    'period closing balance
    fClosingBalance = fOpeningBalance + fSurplus
    
    lblOpeningBal = Format(fOpeningBalance, "£0.00")
    lblIncome = Format(fIncome, "£0.00")
    lblExpense = Format(Abs(fExpense), "£0.00")
    lblSurplus = Format(fSurplus, "£0.00")
    lblClosingBal = Format(fClosingBalance, "£0.00")

    rs.Close
    Set rs = Nothing
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub



Private Sub SetKeyPreview()

On Error GoTo ErrorTrap

    Select Case True
    Case optMonth
        Me.KeyPreview = True
    Case optDateRange
        Me.KeyPreview = False
    Case optSearch
        Me.KeyPreview = True
    End Select
        
    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub optMonth_Click()

On Error GoTo ErrorTrap

    SetUpFormControls
    
    cmbMonth_Click
    
    SetKeyPreview
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub optDateRange_Click()

On Error GoTo ErrorTrap

    SetUpFormControls
    
    RespondToDateRangeChange
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub optSearch_Click()

On Error GoTo ErrorTrap
   
    ClearForm
    
    SetUpFormControls
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub SetUpFormControls()

On Error GoTo ErrorTrap

    If cmbAccount.ListIndex > 0 Then
        optMonth = True
        optDateRange.Enabled = False
        optSearch.Enabled = False
    Else
        optDateRange.Enabled = True
        optSearch.Enabled = True
    End If

    Select Case True
    Case optMonth
        fraMonth.Visible = True
        fraDateRange.Visible = False
        fraSearch.Visible = False
        cmbAccount.Enabled = True
    Case optDateRange
        fraMonth.Visible = False
        fraDateRange.Visible = True
        fraSearch.Visible = False
        cmbAccount.ListIndex = 0
        cmbAccount.Enabled = False
    Case optSearch
        fraMonth.Visible = False
        fraDateRange.Visible = False
        fraSearch.Visible = True
        cmbAccount.ListIndex = 0
        cmbAccount.Enabled = False
    End Select
    
    fraOtherAccounts.Visible = optMonth

    mnuAddRegular.Enabled = optMonth And cmbAccount.ListIndex = 0
    mnuAddRegular2.Enabled = mnuAddRegular.Enabled
    mnuMissingTrans.Enabled = optMonth And cmbAccount.ListIndex = 0
    mnuAccTfr.Enabled = optMonth
    mnuNewTransfer.Enabled = mnuAccTfr.Enabled
    
    mnuOpeningBal.Enabled = optMonth Or optDateRange
    mnuNewTran.Enabled = optMonth Or optDateRange
    mnuNewTransaction.Enabled = mnuNewTran.Enabled
        
    mnuActions2.Enabled = True
    Frame2.Visible = True
    
    ClearForm
    
    cmdSearch.Default = optSearch
    
    SetKeyPreview
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub



Private Sub txtDesc_GotFocus()
    TextFieldGotFocus txtDesc
End Sub

Private Sub txtFirstDate_Change()

On Error GoTo ErrorTrap

    If bSuppress Then Exit Sub
    
    RespondToDateRangeChange

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub txtFirstDate_KeyPress(KeyAscii As Integer)
    KeyPressValid KeyAscii, cmsDates, True
End Sub

Private Sub txtFirstDate_LostFocus()

On Error GoTo ErrorTrap

    If ValidDate(txtFirstDate) Then
        txtFirstDate = Format(txtFirstDate)
    End If

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub txtLastDate_Change()

On Error GoTo ErrorTrap

    If bSuppress Then Exit Sub
    
    RespondToDateRangeChange
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub RespondToDateRangeChange()

On Error GoTo ErrorTrap

    If ValidDate(txtFirstDate) And ValidDate(txtLastDate) Then
        If DateDiff("d", txtFirstDate, txtLastDate) >= 0 Then
            GetDates
            GetTotals
            FillAccountsGrid
        Else
            ClearForm
        End If
    Else
        ClearForm
    End If

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub txtLastDate_KeyPress(KeyAscii As Integer)
    KeyPressValid KeyAscii, cmsDates, True
End Sub

Private Sub txtLastDate_LostFocus()

On Error GoTo ErrorTrap

    If ValidDate(txtLastDate) Then
        txtLastDate = Format(txtLastDate)
    End If

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub txtMaxAmount_GotFocus()
    TextFieldGotFocus txtMaxAmount
End Sub

Private Sub txtMaxRef_GotFocus()
    If IsNumber(txtMinRef, False, False, False) And txtMaxRef = "" Then
        txtMaxRef = txtMinRef
    End If
    TextFieldGotFocus txtMaxRef
End Sub

Private Sub txtMaxRef_KeyPress(KeyAscii As Integer)
    KeyPressValid KeyAscii, cmsUnsignedIntegers, True
End Sub

Private Sub txtMaxRef_LostFocus()
    If IsNumber(txtMaxRef, False, False, False) Then
        If CLng(txtMaxRef) = 0 Then
            txtMaxRef = ""
        End If
    End If
        
End Sub

Private Sub txtMinAmount_GotFocus()
    TextFieldGotFocus txtMinAmount
End Sub

Private Sub txtMinAmount_KeyPress(KeyAscii As Integer)
    KeyPressValid KeyAscii, cmsUnsignedDecimals, True
End Sub
Private Sub txtMaxAmount_KeyPress(KeyAscii As Integer)
    KeyPressValid KeyAscii, cmsUnsignedDecimals, True
End Sub

Private Sub txtMinAmount_LostFocus()
    txtMinAmount = Format(txtMinAmount, "0.00")
End Sub
Private Sub txtMaxAmount_LostFocus()
    txtMaxAmount = Format(txtMaxAmount, "0.00")
End Sub

Private Sub txtMinRef_GotFocus()
    TextFieldGotFocus txtMinRef
End Sub

Private Sub txtMinRef_KeyPress(KeyAscii As Integer)
    KeyPressValid KeyAscii, cmsUnsignedIntegers, True
End Sub

Private Sub txtMinRef_LostFocus()
    If IsNumber(txtMinRef, False, False, False) Then
        If CLng(txtMinRef) = 0 Then
            txtMinRef = ""
            Exit Sub
        End If
    End If
End Sub

Private Sub txtSearchDate1_GotFocus()
    TextFieldGotFocus txtSearchDate1
End Sub
Private Sub txtSearchDate2_GotFocus()
    TextFieldGotFocus txtSearchDate2
End Sub

Private Sub txtSearchDate1_KeyPress(KeyAscii As Integer)
    KeyPressValid KeyAscii, cmsDates, True
End Sub

Private Sub txtSearchDate1_LostFocus()
    If ValidDate(txtSearchDate1) Then
        txtSearchDate1 = Format(txtSearchDate1, "dd/mm/yyyy")
    End If
End Sub

Private Sub txtSearchDate2_KeyPress(KeyAscii As Integer)
    KeyPressValid KeyAscii, cmsDates, True
End Sub
Private Sub txtSearchDate2_LostFocus()
    If ValidDate(txtSearchDate2) Then
        txtSearchDate2 = Format(txtSearchDate2, "dd/mm/yyyy")
    End If
End Sub


Public Property Get CodeType() As Long
    CodeType = mlCodeType
End Property

