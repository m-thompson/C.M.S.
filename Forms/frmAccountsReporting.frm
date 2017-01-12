VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmAccountsReporting 
   Caption         =   " C.M.S. Accounts Reporting"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12600
   Icon            =   "frmAccountsReporting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9945
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   330
      Left            =   135
      TabIndex        =   9
      ToolTipText     =   "Print text to default printer"
      Top             =   9570
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   330
      Left            =   1260
      TabIndex        =   10
      ToolTipText     =   "Copy text to clipboard"
      Top             =   9570
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   330
      Left            =   11355
      TabIndex        =   11
      Top             =   9570
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   8280
      Left            =   135
      TabIndex        =   13
      Top             =   1185
      Width           =   12330
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   7965
         Left            =   60
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   195
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   14049
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmAccountsReporting.frx":0442
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1020
      Left            =   135
      TabIndex        =   12
      Top             =   150
      Width           =   12330
      Begin VB.CommandButton cmdLastYear 
         Caption         =   "Last Year"
         Height          =   315
         Left            =   5130
         TabIndex        =   6
         ToolTipText     =   "Set dates to last financial year"
         Top             =   480
         Width           =   915
      End
      Begin VB.CheckBox chkIncludeTempIncome 
         Caption         =   "Include Temporary Income"
         Height          =   210
         Left            =   6420
         TabIndex        =   7
         Top             =   540
         Width           =   2295
      End
      Begin VB.CommandButton cmdYTD 
         Caption         =   "YTD"
         Height          =   315
         Left            =   4185
         TabIndex        =   5
         ToolTipText     =   "Set dates to 'Year to Date'"
         Top             =   480
         Width           =   915
      End
      Begin VB.CommandButton cmdMTD 
         Caption         =   "MTD"
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         ToolTipText     =   "Set dates to 'Month to Date'"
         Top             =   480
         Width           =   915
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "Run"
         Height          =   315
         Left            =   11130
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtLastDate 
         Height          =   315
         Left            =   1665
         MaxLength       =   10
         TabIndex        =   2
         Top             =   480
         Width           =   978
      End
      Begin VB.TextBox txtFirstDate 
         Height          =   315
         Left            =   165
         MaxLength       =   10
         TabIndex        =   0
         Top             =   480
         Width           =   978
      End
      Begin VB.CommandButton cmdShowCalendar1 
         DownPicture     =   "frmAccountsReporting.frx":04C4
         Height          =   315
         Left            =   1155
         Picture         =   "frmAccountsReporting.frx":0906
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   420
      End
      Begin VB.CommandButton cmdShowCalendar2 
         DownPicture     =   "frmAccountsReporting.frx":0D48
         Height          =   315
         Left            =   2655
         Picture         =   "frmAccountsReporting.frx":118A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   420
      End
      Begin VB.Label Label1 
         Caption         =   "Start Date"
         Height          =   255
         Left            =   165
         TabIndex        =   15
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "End Date"
         Height          =   255
         Left            =   1665
         TabIndex        =   14
         Top             =   255
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmAccountsReporting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents frmCal As frmMiniCalendar
Attribute frmCal.VB_VarHelpID = -1
Dim bClickedStartDate As Boolean, mbIgnore As Boolean
Dim mbIncTmp As Boolean

Private Sub chkIncludeTempIncome_Click()
On Error GoTo ErrorTrap

    mbIncTmp = (chkIncludeTempIncome = vbChecked)

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub cmdCopy_Click()

On Error GoTo ErrorTrap

    With RichTextBox1

    .SelStart = 0
    .SelLength = Len(.text)
    
    CopyTextToClipBoard .SelRTF, vbCFRTF
    
    .SelLength = 0
    
    ShowMessage "Accounts copied to clipboard", 1500, Me
    
    End With

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdLastYear_Click()
On Error GoTo ErrorTrap

    txtFirstDate = "01/04/" & (year(Now) - 1)
    txtLastDate = "31/03/" & year(Now)
    
    GenerateReport

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub cmdMTD_Click()

On Error GoTo ErrorTrap

    txtFirstDate = "01/" & Format(Month(Now), "00") & "/" & year(Now)
    txtLastDate = Format(Now, "dd/mm/yyyy")
    
    GenerateReport

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdPrint_Click()

On Error GoTo ErrorTrap

Dim lSaveOrientation As Long, lSavePaperHeight As Long, lSavePaperWidth As Long


    With RichTextBox1
        
    If MsgBox("Print text to the default printer?", vbQuestion + vbYesNo, AppName) = vbYes Then
        
        lSaveOrientation = Printer.Orientation
        lSavePaperHeight = Printer.Height
        lSavePaperWidth = Printer.Width
        
        .SelLength = 0
        Printer.Orientation = vbPRORLandscape
        Printer.Width = 567 * GlobalParms.GetValue("AccountsReportingPrintWidthCM", "NumFloat", 20.5)
        Printer.Height = 567 * GlobalParms.GetValue("AccountsReportingPrintHeightCM", "NumFloat", 29)
        
        .SelPrint Printer.hdc
        
        Printer.Width = lSavePaperWidth
        Printer.Height = lSavePaperHeight
        Printer.Orientation = lSaveOrientation
    
    End If
    
    End With

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdYTD_Click()

On Error GoTo ErrorTrap

    txtFirstDate = "01/04/" & GetFinancialYear(Now)
    txtLastDate = Format(Now, "dd/mm/yyyy")
    
    GenerateReport

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdRun_Click()

On Error GoTo ErrorTrap

    If ValidDate(txtFirstDate) And ValidDate(txtLastDate) Then
        GenerateReport
    Else
        MsgBox "Invalid date range", vbOKOnly + vbExclamation, AppName
    End If

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdShowCalendar1_Click()

On Error GoTo ErrorTrap

    bClickedStartDate = True
    
    frmCal.SetPos = True
    frmCal.FormDate = txtFirstDate
    frmCal.XPos = Me.Left + Frame1.Left + cmdShowCalendar1.Left + cmdShowCalendar1.Width
    frmCal.YPos = Me.Top + Frame1.Top + cmdShowCalendar1.Top + cmdShowCalendar1.Height
    frmCal.Show vbModeless, Me

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdShowCalendar2_Click()
On Error GoTo ErrorTrap

    bClickedStartDate = False
    
    frmCal.SetPos = True
    frmCal.FormDate = txtLastDate
    frmCal.XPos = Me.Left + Frame1.Left + cmdShowCalendar2.Left + cmdShowCalendar2.Width
    frmCal.YPos = Me.Top + Frame1.Top + cmdShowCalendar2.Top + cmdShowCalendar2.Height
    frmCal.Show vbModeless, Me

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub Form_Load()

On Error GoTo ErrorTrap

    Set frmCal = New frmMiniCalendar

    cmdYTD_Click 'populate dates
    
    GenerateReport
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub Form_Resize()
On Error Resume Next

    With Me
    
    .Width = 12720
    
    cmdCopy.Top = .Height - 780
    cmdPrint.Top = .Height - 780
    cmdClose.Top = .Height - 780
    
    If .Height < 3600 Then
        .Height = 3600
    End If
    If .Height > 10350 Then
        .Height = 10350
    End If
    
    Frame2.Height = .Height - 2070
    RichTextBox1.Height = Frame2.Height - 315
    
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Unload frmCal
    Set frmCal = Nothing
    
    frmAccountSummary.SetFocus
    
End Sub



Private Sub frmCal_InsertDate(TheDate As String)
On Error GoTo ErrorTrap

    If bClickedStartDate Then
        txtFirstDate = TheDate
    Else
        txtLastDate = TheDate
    End If
    
    CheckDatesAndRun
    
    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub
Private Sub CheckDatesAndRun()
On Error GoTo ErrorTrap

    If ValidDate(txtFirstDate) And ValidDate(txtLastDate) Then
        If CDate(txtLastDate) >= CDate(txtFirstDate) Then
            GenerateReport
        Else
            RichTextBox1.text = ""
        End If
    Else
        RichTextBox1.text = ""
    End If

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub txtFirstDate_Change()
    CheckDatesAndRun
End Sub

Private Sub txtFirstDate_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorTrap

    KeyPressValid KeyAscii, cmsDates, True

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub txtFirstDate_LostFocus()

    txtFirstDate = Format(txtFirstDate, "dd/mm/yyyy")
    
    CheckDatesAndRun

End Sub

Private Sub txtLastDate_Change()
    CheckDatesAndRun
End Sub

Private Sub txtLastDate_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorTrap

    KeyPressValid KeyAscii, cmsDates, True

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub GenerateReport()
On Error GoTo ErrorTrap

Dim rsInTotals As Recordset, sSQL As String
Dim rsInOtherDetail As Recordset, rsGiftAidReceipts As Recordset
Dim rsOutTotals As Recordset, rsOutOtherDetail As Recordset
Dim rsOutGrandTotal As Recordset, rsInGrandTotal As Recordset
Dim rsStartBal As Recordset, rsEndBal As Recordset
Dim rsStartBalOtherAccs As Recordset, rsEndBalOtherAccs As Recordset
Dim rsCircuitExpenses As Recordset
Dim rsSubTranReceipts As Recordset
Dim rsInTranTypes As Recordset
Dim rsOutTranTypes As Recordset
Dim rsContsToSociety As Recordset
Dim rsBkGrpConts As Recordset

Dim sHeadingText As String
Dim sDetailText As String
Dim sDetailText2 As String
Dim sReceiptsHeadingText As String
Dim sExpensesHeadingText As String
Dim sBalanceText As String
Dim sBalanceOtherAccsText As String
Dim sGiftAidText As String
Dim sOtherReceiptsText As String
Dim sTranSubTypeText As String
Dim sOtherExpensesText As String
Dim sCircuitContributionsText As String
Dim sGiftAidHeadingText As String
Dim sOtherReceiptsHeadingText As String
Dim sTranSubTypeHeadingText As String
Dim sOtherExpensesHeadingText As String
Dim sCircuitContHeadingText As String
Dim sContsToSocietyText As String
Dim sContsToSocietyHeadingText As String
Dim sBkGrpBreakdownText As String
Dim sBkGrpBreakdownHeadingText As String

Dim dAmount As Double
Dim dStartBal As Double
Dim dEndBal As Double
Dim dStartBalTot As Double
Dim dEndBalTot As Double
Dim lAccountID As Long
Dim bOtherAccsFound As Boolean
Dim sLine As String

Dim lPosSoFar As Long

Dim sTmpInStr As String

    sLine = "-----------------------------------------------------------"
    sTmpInStr = IIf(mbIncTmp, "", "AND c.InOutTypeID NOT IN (5,7) ")

'    'get summary of all receipts grouped by Tran Code
'    sSQL = "SELECT b.TranCodeID, b.TranCode, b.Description, SUM(e.Amount) AS ReceiptSum " & _
'           "FROM ((tblTransactionTypes b " & _
'           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
'           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
'           "LEFT JOIN tblTransactionDates e ON e.TranCodeID = b.TranCodeID " & _
'           "Where d.InOutID = 1 " & _
'           sTmpInStr & _
'           "AND (e.TranDate BETWEEN #" & Format(txtFirstDate, "mm/dd/yyyy") & "# AND " & _
'           " #" & Format(txtLastDate, "mm/dd/yyyy") & "# " & _
'            "   OR e.TranDate IS NULL)" & _
'           "GROUP BY b.TranCodeID, b.TranCode, b.Description"
'
'    Set rsInTotals = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    'get receipt tran-types
    sSQL = "SELECT b.TranCodeID, b.TranCode, b.Description " & _
           "FROM (tblTransactionTypes b " & _
           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID " & _
           "WHERE d.InOutID = 1 " & _
           " AND b.Suppressed = FALSE " & _
           sTmpInStr & _
           " ORDER BY b.TranCode "
           
    Set rsInTranTypes = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)

    'get breakdown of all 'other' receipts
    sSQL = "SELECT b.TranCodeID, b.TranCode, e.TranDescription, ABS(e.Amount) AS Amount " & _
           "FROM ((tblTransactionTypes b " & _
           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
           "INNER JOIN tblTransactionDates e ON e.TranCodeID = b.TranCodeID " & _
           "WHERE d.InOutID = 1 " & _
           sTmpInStr & _
           "AND  b.TranCode = 'O' " & _
           "AND  e.AccountID = 0 " & _
           "AND (e.TranDate BETWEEN #" & Format(txtFirstDate, "mm/dd/yyyy") & "# AND " & _
           " #" & Format(txtLastDate, "mm/dd/yyyy") & "# " & _
            "   OR e.amount IS NULL)"

    Set rsInOtherDetail = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)


    'get Book-Group contributions
    sSQL = "SELECT iif(e.BookGroupNo = 0,'> Congregation',f.GroupName) AS TheGroupName , SUM(e.Amount) AS Amount " & _
           "FROM (((tblTransactionTypes b " & _
           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
           "INNER JOIN tblTransactionDates e ON e.TranCodeID = b.TranCodeID) " & _
           "LEFT JOIN tblBookGroups f ON f.GroupNo = e.BookGroupNo " & _
           "WHERE d.InOutID = 1 " & _
           sTmpInStr & _
           " AND BookGroupNo > -1 " & _
           "AND  b.TranCode = '" & GlobalParms.GetValue("CongContributionTransactionCode", "AlphaVal") & "' " & _
           "AND  e.AccountID = 0 " & _
           "AND (e.TranDate BETWEEN #" & Format(txtFirstDate, "mm/dd/yyyy") & "# AND " & _
           " #" & Format(txtLastDate, "mm/dd/yyyy") & "# " & _
            "   OR e.amount IS NULL) " & _
           "GROUP BY iif(e.BookGroupNo = 0,'> Congregation',f.GroupName)"

    Set rsBkGrpConts = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    
    'get all 'Transaction Sub-Type' receipts
    sSQL = "SELECT b.TranCode, f.Description, SUM(e.Amount) AS Amount " & _
           "FROM (((tblTransactionTypes b " & _
           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
           "INNER JOIN tblTransactionDates e ON e.TranCodeID = b.TranCodeID) " & _
           "INNER JOIN tblTransactionSubtypes f ON f.TranSubCodeID = e.TranSubTypeID " & _
           "WHERE d.InOutID = 1 " & _
           sTmpInStr & _
           "AND  e.TranSubTypeID > 0 " & _
           "AND  e.AccountID = 0 " & _
           "AND (e.TranDate BETWEEN #" & Format(txtFirstDate, "mm/dd/yyyy") & "# AND " & _
           " #" & Format(txtLastDate, "mm/dd/yyyy") & "# " & _
            "   OR e.amount IS NULL) " & _
           "AND b.Suppressed = FALSE " & _
           " AND f.Suppressed = FALSE " & _
           "GROUP BY b.TranCode, f.Description"

    Set rsSubTranReceipts = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    'get breakdown of all Gift Aid receipts, grouped by Gift Aid No
    sSQL = "SELECT b.TranCodeID, b.TranCode, 'Gift Aid No ' & e.RefNo AS GiftAidDesc, SUM(e.Amount) AS GiftAidSum " & _
           "FROM ((tblTransactionTypes b " & _
           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
           "INNER JOIN tblTransactionDates e ON e.TranCodeID = b.TranCodeID " & _
           "WHERE d.InOutID = 1 " & _
           sTmpInStr & _
           "AND  b.TranCode = 'G' " & _
           "AND  e.AccountID = 0 " & _
           "AND (e.TranDate BETWEEN #" & Format(txtFirstDate, "mm/dd/yyyy") & "# AND " & _
           " #" & Format(txtLastDate, "mm/dd/yyyy") & "# " & _
            "   OR e.amount IS NULL)" & _
           "GROUP BY b.TranCodeID, b.TranCode, e.RefNo " & _
           "ORDER BY e.RefNo"
         
    Set rsGiftAidReceipts = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
'    'get summary of all expenses grouped by Tran Code
'    sSQL = "SELECT b.TranCodeID, b.TranCode, b.Description, ABS(SUM(e.Amount)) AS ExpenseSum " & _
'           "FROM ((tblTransactionTypes b " & _
'           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
'           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
'           "LEFT JOIN tblTransactionDates e ON e.TranCodeID = b.TranCodeID " & _
'           "Where d.InOutID = 2 " & _
'           "AND c.InOutTypeID <> 6 " & _
'           "AND (e.TranDate BETWEEN #" & Format(txtFirstDate, "mm/dd/yyyy") & "# AND " & _
'           " #" & Format(txtLastDate, "mm/dd/yyyy") & "# " & _
'            "   OR e.amount IS NULL)" & _
'           "GROUP BY b.TranCodeID, b.TranCode, b.Description"
'
'    Set rsOutTotals = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)

    'get expense tran-types
    sSQL = "SELECT b.TranCodeID, b.TranCode, b.Description " & _
           "FROM (tblTransactionTypes b " & _
           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID " & _
           "Where d.InOutID = 2 " & _
           "AND b.Suppressed = FALSE " & _
           " ORDER BY b.TranCode "
           
    Set rsOutTranTypes = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)


    'get breakdown of all 'other' expenses
    sSQL = "SELECT b.TranCodeID, b.TranCode, e.TranDescription, ABS(e.Amount) AS Amount " & _
           "FROM ((tblTransactionTypes b " & _
           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
           "INNER JOIN tblTransactionDates e ON e.TranCodeID = b.TranCodeID " & _
           "WHERE d.InOutID = 2 " & _
           "AND (e.TranDate BETWEEN #" & Format(txtFirstDate, "mm/dd/yyyy") & "# AND " & _
           " #" & Format(txtLastDate, "mm/dd/yyyy") & "# " & _
            "   OR e.amount IS NULL)" & _
           "AND c.InOutTypeID <> 6 " & _
           "AND  e.AccountID = 0 " & _
           "AND  b.TranCode = 'O'"
         
    Set rsOutOtherDetail = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    'get breakdown of all 'circuit' expenses
    sSQL = "SELECT b.TranCodeID, b.TranCode, e.TranDescription, ABS(e.Amount)  AS Amount  " & _
           "FROM ((tblTransactionTypes b " & _
           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
           "INNER JOIN tblTransactionDates e ON e.TranCodeID = b.TranCodeID " & _
           "WHERE d.InOutID = 2 " & _
           "AND (e.TranDate BETWEEN #" & Format(txtFirstDate, "mm/dd/yyyy") & "# AND " & _
           " #" & Format(txtLastDate, "mm/dd/yyyy") & "# " & _
            "   OR e.amount IS NULL) " & _
           "AND  e.AccountID = 0 " & _
           "AND c.InOutTypeID <> 6 " & _
           "AND  b.TranCode = 'A'"
           
    Set rsCircuitExpenses = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
         
    'get breakdown of all donations in boxes for IBSA and WBTS
    sSQL = "SELECT b.TranCodeID, b.TranCode, e.TranDescription, ABS(SUM(e.Amount)) AS SocietySum  " & _
           "FROM ((tblTransactionTypes b " & _
           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
           "INNER JOIN tblTransactionDates e ON e.TranCodeID = b.TranCodeID " & _
           "WHERE  e.TranDate BETWEEN #" & Format(txtFirstDate, "mm/dd/yyyy") & "# AND " & _
           " #" & Format(txtLastDate, "mm/dd/yyyy") & "# " & _
           "AND c.InOutTypeID = 5 " & _
           "AND  e.AccountID = 0 " & _
           " GROUP BY b.TranCodeID, b.TranCode, e.TranDescription" & _
           " ORDER BY b.TranCode "
           
    Set rsContsToSociety = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    'get total of all expenses
    sSQL = "SELECT ABS(SUM(e.Amount)) AS ExpenseSum " & _
           "FROM ((tblTransactionTypes b " & _
           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
           "LEFT JOIN tblTransactionDates e ON e.TranCodeID = b.TranCodeID " & _
           "Where d.InOutID = 2 " & _
           "AND c.InOutTypeID <> 6 " & _
           "AND  e.AccountID = 0 " & _
           "AND (e.TranDate BETWEEN #" & Format(txtFirstDate, "mm/dd/yyyy") & "# AND " & _
           " #" & Format(txtLastDate, "mm/dd/yyyy") & "# " & _
            "   OR e.amount IS NULL)"
         
    Set rsOutGrandTotal = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    'get total of all receipts
    sSQL = "SELECT ABS(SUM(e.Amount)) AS ReceiptSum " & _
           "FROM ((tblTransactionTypes b " & _
           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
           "LEFT JOIN tblTransactionDates e ON e.TranCodeID = b.TranCodeID " & _
           "Where d.InOutID = 1 " & _
           sTmpInStr & _
           " AND  e.AccountID = 0 " & _
           "AND e.TranDate BETWEEN #" & Format(txtFirstDate, "mm/dd/yyyy") & "# AND " & _
           " #" & Format(txtLastDate, "mm/dd/yyyy") & "# "
         
    Set rsInGrandTotal = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
     
    'get balance at start of period
    sSQL = "SELECT iif(isnull(SUM(e.Amount)),0,SUM(e.Amount)) AS StartBalance " & _
           "FROM ((tblTransactionTypes b " & _
           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
           "INNER JOIN tblTransactionDates e ON e.TranCodeID = b.TranCodeID " & _
           "WHERE e.TranDate < #" & Format(txtFirstDate, "mm/dd/yyyy") & "# " & _
           " AND  e.AccountID = 0 " & _
            "AND c.InOutTypeID NOT IN (5,6,7) "

         
    Set rsStartBal = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    'get balance at end of period
    sSQL = "SELECT iif(isnull(SUM(e.Amount)),0,SUM(e.Amount)) AS EndBalance " & _
           "FROM ((tblTransactionTypes b " & _
           "INNER JOIN tblAccInOutTypes c ON b.InOutTypeID = c.InOutTypeID) " & _
           "INNER JOIN tblAccInOut d ON c.InOutID = d.InOutID) " & _
           "INNER JOIN tblTransactionDates e ON e.TranCodeID = b.TranCodeID " & _
           "WHERE e.TranDate <= #" & Format(txtLastDate, "mm/dd/yyyy") & "# " & _
           " AND  e.AccountID = 0 " & _
            "AND c.InOutTypeID NOT IN (5,6,7) "
         
    Set rsEndBal = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    'get balances for other accounts at start of period
    sSQL = "SELECT c.AccountID, iif(isnull(SUM(b.Amount)),0,SUM(iif(b.AccountID = 0,-1,1) * b.Amount)) AS StartBalance " & _
           "FROM tblTransactionDates b " & _
           "INNER JOIN tblBankAccounts c ON b.TfrAccountID = c.AccountID " & _
           "WHERE b.TranDate < #" & Format(txtFirstDate, "mm/dd/yyyy") & "# " & _
           " AND  c.AccountID > 0 " & _
           " AND  b.AccountID <= 0 " & _
           "GROUP BY c.AccountID " & _
           "UNION ALL " & _
           "SELECT AccountID, 0 " & _
           "FROM tblBankAccounts " & _
           "WHERE AccountID NOT IN " & _
           "(SELECT TfrAccountID " & _
           " FROM tblTransactionDates " & _
           " WHERE TfrAccountID > 0 " & _
           " AND TranDate < #" & Format(txtFirstDate, "mm/dd/yyyy") & "#) " & _
           "AND AccountID > 0 "
           
    Set rsStartBalOtherAccs = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    'get balances for other accounts at end of period
    sSQL = "SELECT c.AccountID, iif(isnull(SUM(b.Amount)),0,SUM(iif(b.AccountID = 0,-1,1) * b.Amount)) AS EndBalance " & _
           "FROM tblTransactionDates b " & _
           "INNER JOIN tblBankAccounts c ON b.TfrAccountID = c.AccountID " & _
           "WHERE b.TranDate <= #" & Format(txtLastDate, "mm/dd/yyyy") & "# " & _
           " AND  c.AccountID > 0 " & _
           " AND  b.AccountID <= 0 " & _
           "GROUP BY c.AccountID " & _
           "UNION ALL " & _
           "SELECT AccountID, 0 " & _
           "FROM tblBankAccounts " & _
           "WHERE AccountID NOT IN " & _
           "(SELECT TfrAccountID " & _
           " FROM tblTransactionDates " & _
           " WHERE TfrAccountID > 0 " & _
           " AND TranDate <= #" & Format(txtLastDate, "mm/dd/yyyy") & "#) " & _
           "AND AccountID > 0 "
         
    Set rsEndBalOtherAccs = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    '---------
    
    'set up RichTextBox's fixed tabs...
    SetTBTabStops RichTextBox1, 50, 557
    
    With RichTextBox1
    
    '----- Put in the raw text - format it all later
    
    'clear existing text
    .text = ""
    
    'put in heading
    If txtFirstDate = txtLastDate Then
        sHeadingText = "Accounts Breakdown for " & txtFirstDate & vbCr & vbCr
    Else
        sHeadingText = "Accounts Breakdown for period " & txtFirstDate & " to " & txtLastDate & vbCr & vbCr
    End If
        
    .text = sHeadingText
    
    'put in receipts heading
    sReceiptsHeadingText = "Current Account Receipts" & vbCr & vbCr
    .text = .text & sReceiptsHeadingText
    
    'put in the receipts
    sDetailText = ""
    With rsInTranTypes
    Do Until .EOF Or .BOF
        
        sSQL = "SELECT ABS(iif(isnull(SUM(e.Amount)),0,SUM(e.Amount))) AS ReceiptSum " & _
               "FROM tblTransactionDates e " & _
               "WHERE e.TranDate BETWEEN #" & Format(txtFirstDate, "mm/dd/yyyy") & "# AND " & _
               " #" & Format(txtLastDate, "mm/dd/yyyy") & "# AND " & _
               "TranCodeID = " & !TranCodeID & _
                " AND AccountID = 0 "
               
        Set rsInTranTypes = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
        
        sDetailText = sDetailText & _
                      !TranCode & vbTab & !Description & ":" & _
                            vbTab & RightAlignString(Format(rsInTranTypes!ReceiptSum, "£0.00"), 11) & vbCr
                            
        .MoveNext
    
    Loop
    End With
    
    sDetailText = sDetailText & _
                     ">>>>>" & vbTab & "Total Receipts:" & _
                          vbTab & RightAlignString(Format(rsInGrandTotal!ReceiptSum, "£0.00"), 11) & vbCr
                          
    sDetailText = sDetailText & vbCr
    .text = .text & sDetailText
    
    'put in the 'Expenses' subheading
    sExpensesHeadingText = "Current Account Expenses" & vbCr & vbCr
    .text = .text & sExpensesHeadingText
    
    'put in the expenses
    sDetailText2 = ""
    With rsOutTranTypes
    Do Until .EOF Or .BOF
        
        sSQL = "SELECT ABS(iif(isnull(SUM(e.Amount)),0,SUM(e.Amount))) AS ExpenseSum " & _
               "FROM tblTransactionDates e " & _
               "WHERE e.TranDate BETWEEN #" & Format(txtFirstDate, "mm/dd/yyyy") & "# AND " & _
               " #" & Format(txtLastDate, "mm/dd/yyyy") & "# AND " & _
               "TranCodeID = " & !TranCodeID & _
                " AND AccountID = 0 "
               
        Set rsOutTranTypes = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
        
        sDetailText2 = sDetailText2 & _
                      !TranCode & vbTab & !Description & ":" & _
                            vbTab & RightAlignString(Format(rsOutTranTypes!ExpenseSum, "£0.00"), 11) & vbCr
                            
        .MoveNext
    
    Loop
    End With
    
    sDetailText2 = sDetailText2 & _
                     ">>>>>" & vbTab & "Total Expenses:" & _
                          vbTab & RightAlignString(Format(rsOutGrandTotal!ExpenseSum, "£0.00"), 11) & vbCr
    sDetailText2 = sDetailText2 & vbCr
    .text = .text & sDetailText2
    
    'put in start/end balances and surplus
    
    dAmount = GetAccountStartAmount(0)
    
    sBalanceText = GetAccountName(0) & " Balance at Start of Period:" & vbTab & RightAlignString(Format(rsStartBal!StartBalance + dAmount, "£0.00"), 11) & vbCr
    
    sBalanceText = sBalanceText & _
                   "Surplus / (Deficit):" & vbTab & _
                   RightAlignString(Format(rsEndBal!EndBalance - rsStartBal!StartBalance, "£0.00"), 11) & vbCr
    
    sBalanceText = sBalanceText & GetAccountName(0) & " Balance at End of Period:" & vbTab & RightAlignString(Format(rsEndBal!EndBalance + dAmount, "£0.00"), 11) & _
                    vbCr & vbCr & vbCr
                    
    dStartBalTot = rsStartBal!StartBalance + dAmount
    dEndBalTot = rsEndBal!EndBalance + dAmount
    
    .text = .text & sBalanceText
    
    'put in start/end balances and surplus for other accounts
    
    sBalanceOtherAccsText = ""
    Do Until (rsStartBalOtherAccs.BOF Or rsStartBalOtherAccs.EOF) And _
             (rsEndBalOtherAccs.BOF Or rsEndBalOtherAccs.EOF)
             
        bOtherAccsFound = True
        lAccountID = rsStartBalOtherAccs!AccountID
        dAmount = GetAccountStartAmount(lAccountID)
        dStartBal = rsStartBalOtherAccs!StartBalance
        dEndBal = rsEndBalOtherAccs!EndBalance
        
        sBalanceOtherAccsText = sBalanceOtherAccsText & GetAccountName(lAccountID) & " Balance at Start of Period:" & vbTab & RightAlignString(Format(dStartBal + dAmount, "£0.00"), 11) & vbCr
        
        sBalanceOtherAccsText = sBalanceOtherAccsText & _
                                "Surplus / (Deficit):" & vbTab & _
                                RightAlignString(Format(dEndBal - dStartBal, "£0.00"), 11) & vbCr
        
        sBalanceOtherAccsText = sBalanceOtherAccsText & GetAccountName(lAccountID) & " Balance at End of Period:" & vbTab & RightAlignString(Format(dEndBal + dAmount, "£0.00"), 11) & _
                                vbCr & vbCr
        
        dStartBalTot = dStartBalTot + dStartBal + dAmount
        dEndBalTot = dEndBalTot + dEndBal + dAmount
        
        rsStartBalOtherAccs.MoveNext
        rsEndBalOtherAccs.MoveNext
        
    Loop
    
    If bOtherAccsFound Then
        sBalanceOtherAccsText = sBalanceOtherAccsText & vbCr
    
        
        sBalanceOtherAccsText = sBalanceOtherAccsText & "TOTAL BALANCE AT START OF PERIOD:" & vbTab & RightAlignString(Format(dStartBalTot, "£0.00"), 11) & vbCr
        
        sBalanceOtherAccsText = sBalanceOtherAccsText & _
                                "SURPLUS / (DEFICIT):" & vbTab & _
                                RightAlignString(Format(dEndBalTot - dStartBalTot, "£0.00"), 11) & vbCr
        
        sBalanceOtherAccsText = sBalanceOtherAccsText & "TOTAL BALANCE AT END OF PERIOD:" & vbTab & RightAlignString(Format(dEndBalTot, "£0.00"), 11) & _
                                vbCr & vbCr & vbCr
        
    End If
    
    .text = .text & sBalanceOtherAccsText
    
    'put in the 'Gift Aid' subheading
    sGiftAidHeadingText = "Gift Aid Contributions Breakdown" & vbCr & vbCr
    .text = .text & sGiftAidHeadingText
    
    'put in Gift Aid Details
    sGiftAidText = ""
    With rsGiftAidReceipts
    Do Until .EOF Or .BOF
        
        sGiftAidText = sGiftAidText & _
                      !TranCode & vbTab & !GiftAidDesc & ": " & _
                            vbTab & RightAlignString(Format(!GiftAidSum, "£0.00"), 11) & vbCr
                            
        .MoveNext
    
    Loop
    End With
    
    sGiftAidText = sGiftAidText & vbCr & vbCr
    .text = .text & sGiftAidText
    
'    'put in the 'Book-Group Breakdown' subheading
'    sBkGrpBreakdownHeadingText = "Book-Group Contributions Breakdown" & vbCr & vbCr
'    .text = .text & sBkGrpBreakdownHeadingText
'
'    'put in Book-Group Breakdown Details
'    sBkGrpBreakdownText = ""
'    With rsBkGrpConts
'    Do Until .EOF Or .BOF
'
'        sBkGrpBreakdownText = sBkGrpBreakdownText & _
'                       !TheGroupName & ": " & _
'                            vbTab & RightAlignString(Format(!Amount, "£0.00"), 11) & vbCr
'
'        .MoveNext
'
'    Loop
'    End With
'
'    sBkGrpBreakdownText = sBkGrpBreakdownText & vbCr & vbCr
'    .text = .text & sBkGrpBreakdownText
    

    'put in 'Other receipts' Heading
    sOtherReceiptsHeadingText = "Other Receipts Breakdown" & vbCr & vbCr
    .text = .text & sOtherReceiptsHeadingText
    
    'put in 'Other receipts' Details
    sOtherReceiptsText = ""
    With rsInOtherDetail
    Do Until .EOF Or .BOF
        
        sOtherReceiptsText = sOtherReceiptsText & _
                      !TranCode & vbTab & !TranDescription & ": " & _
                            vbTab & RightAlignString(Format(!Amount, "£0.00"), 11) & vbCr
                            
        .MoveNext
    
    Loop
    End With
    
    sOtherReceiptsText = sOtherReceiptsText & vbCr & vbCr
    .text = .text & sOtherReceiptsText

    'put in 'Transaction Sub-Type Breakdown' Heading
    sTranSubTypeHeadingText = "Transaction Sub-Type Breakdown" & vbCr & vbCr
    .text = .text & sTranSubTypeHeadingText
    
    'put in 'Transaction Sub-Type Breakdown' Details
    sTranSubTypeText = ""
    With rsSubTranReceipts
    Do Until .EOF Or .BOF
        
        sTranSubTypeText = sTranSubTypeText & _
                      !TranCode & vbTab & !Description & ": " & _
                            vbTab & RightAlignString(Format(!Amount, "£0.00"), 11) & vbCr
                            
        .MoveNext
    
    Loop
    End With
    
    sTranSubTypeText = sTranSubTypeText & vbCr & vbCr
    .text = .text & sTranSubTypeText
    
    'put in 'Other Expenses' Heading
    sOtherExpensesHeadingText = "Other Expenses Breakdown" & vbCr & vbCr
    .text = .text & sOtherExpensesHeadingText
    
    'put in 'Other Expenses' Details
    sOtherExpensesText = ""
    With rsOutOtherDetail
    Do Until .EOF Or .BOF

        sOtherExpensesText = sOtherExpensesText & _
                      !TranCode & vbTab & !TranDescription & ": " & _
                            vbTab & RightAlignString(Format(!Amount, "£0.00"), 11) & vbCr

        .MoveNext

    Loop
    End With

    sOtherExpensesText = sOtherExpensesText & vbCr & vbCr
    .text = .text & sOtherExpensesText
    
    'put in 'Circuit Expenses' Heading
    sCircuitContHeadingText = "Circuit Contributions Breakdown" & vbCr & vbCr
    .text = .text & sCircuitContHeadingText
    
    'put in 'Circuit Expenses' Details
    sCircuitContributionsText = ""
    With rsCircuitExpenses
    Do Until .EOF Or .BOF

        sCircuitContributionsText = sCircuitContributionsText & _
                      !TranCode & vbTab & !TranDescription & ": " & _
                            vbTab & RightAlignString(Format(!Amount, "£0.00"), 11) & vbCr

        .MoveNext

    Loop
    End With

    sCircuitContributionsText = sCircuitContributionsText & vbCr & vbCr
    .text = .text & sCircuitContributionsText
    
    'put in 'Contributions to IBSA/WBTS in boxes' Heading
    sContsToSocietyHeadingText = "Contributions to IBSA/WBTS in Boxes" & vbCr & vbCr
    .text = .text & sContsToSocietyHeadingText
    
    'put in 'Contributions to IBSA/WBTS in boxes' Details
    sContsToSocietyText = ""
    With rsContsToSociety
    Do Until .EOF Or .BOF

        sContsToSocietyText = sContsToSocietyText & _
                      !TranCode & vbTab & !TranDescription & ": " & _
                            vbTab & RightAlignString(Format(!SocietySum, "£0.00"), 11) & vbCr

        .MoveNext

    Loop
    End With

    sContsToSocietyText = sContsToSocietyText & vbCr & vbCr
    .text = .text & sContsToSocietyText
    
    
  '--- Now apply formatting to added text
    
    'format the heading
    .SelStart = 0
    .SelLength = Len(sHeadingText)
    .SelFontName = "Arial"
    .SelFontSize = 14
    .SelUnderline = True
    .SelBold = True
    .SelAlignment = rtfLeft
    lPosSoFar = Len(sHeadingText)
    
    'format 'Receipts' heading
    .SelStart = lPosSoFar
    .SelLength = Len(sReceiptsHeadingText)
    .SelFontName = "Arial"
    .SelFontSize = 10
    .SelUnderline = False
    .SelBold = True
    lPosSoFar = lPosSoFar + Len(sReceiptsHeadingText)
    
    'format receipts details just added
    .SelStart = lPosSoFar
    .SelLength = Len(sDetailText)
    .SelFontName = "Courier New"
    .SelFontSize = 10
    .SelUnderline = True
    .SelBold = False
    lPosSoFar = lPosSoFar + Len(sDetailText)
    
    'format 'Expenses' heading
    .SelStart = lPosSoFar
    .SelLength = Len(sExpensesHeadingText)
    .SelFontName = "Arial"
    .SelFontSize = 10
    .SelUnderline = False
    .SelBold = True
    lPosSoFar = lPosSoFar + Len(sExpensesHeadingText)

    'format expenses details just added
    .SelStart = lPosSoFar
    .SelLength = Len(sDetailText2)
    .SelFontName = "Courier New"
    .SelFontSize = 10
    .SelUnderline = True
    .SelBold = False
    lPosSoFar = lPosSoFar + Len(sDetailText2)
    
    'format balance details
    .SelStart = lPosSoFar
    .SelLength = Len(sBalanceText)
    .SelFontName = "Courier New"
    .SelFontSize = 10
    .SelUnderline = True
    .SelBold = True
    lPosSoFar = lPosSoFar + Len(sBalanceText)
    
    If bOtherAccsFound Then
        'format other account balance details
        .SelStart = lPosSoFar
        .SelLength = Len(sBalanceOtherAccsText)
        .SelFontName = "Courier New"
        .SelFontSize = 10
        .SelUnderline = True
        .SelBold = True
        lPosSoFar = lPosSoFar + Len(sBalanceOtherAccsText)
    End If
    
    'format 'Gift Aid Breakdown' heading
    .SelStart = lPosSoFar
    .SelLength = Len(sGiftAidHeadingText)
    .SelFontName = "Arial"
    .SelFontSize = 10
    .SelUnderline = False
    .SelBold = True
    lPosSoFar = lPosSoFar + Len(sGiftAidHeadingText)

    'format Gift Aid Breakdown
    .SelStart = lPosSoFar
    .SelLength = Len(sGiftAidText)
    .SelFontName = "Courier New"
    .SelFontSize = 10
    .SelUnderline = True
    .SelBold = False
    lPosSoFar = lPosSoFar + Len(sGiftAidText)
        
    'format 'Book-Group Breakdown' heading
    .SelStart = lPosSoFar
    .SelLength = Len(sBkGrpBreakdownHeadingText)
    .SelFontName = "Arial"
    .SelFontSize = 10
    .SelUnderline = False
    .SelBold = True
    lPosSoFar = lPosSoFar + Len(sBkGrpBreakdownHeadingText)

    'format Book-Group Breakdown
    .SelStart = lPosSoFar
    .SelLength = Len(sBkGrpBreakdownText)
    .SelFontName = "Courier New"
    .SelFontSize = 10
    .SelUnderline = True
    .SelBold = False
    lPosSoFar = lPosSoFar + Len(sBkGrpBreakdownText)
        
    'format 'other receipts' Breakdown heading
    .SelStart = lPosSoFar
    .SelLength = Len(sOtherReceiptsHeadingText)
    .SelFontName = "Arial"
    .SelFontSize = 10
    .SelUnderline = False
    .SelBold = True
    lPosSoFar = lPosSoFar + Len(sOtherReceiptsHeadingText)

    'format 'other receipts' Breakdown
    .SelStart = lPosSoFar
    .SelLength = Len(sOtherReceiptsText)
    .SelFontName = "Courier New"
    .SelFontSize = 10
    .SelUnderline = True
    .SelBold = False
    lPosSoFar = lPosSoFar + Len(sOtherReceiptsText)
    
    'format 'Transaction Sub-Type' Breakdown heading
    .SelStart = lPosSoFar
    .SelLength = Len(sTranSubTypeHeadingText)
    .SelFontName = "Arial"
    .SelFontSize = 10
    .SelUnderline = False
    .SelBold = True
    lPosSoFar = lPosSoFar + Len(sTranSubTypeHeadingText)

    'format 'Transaction Sub-Type' Breakdown
    .SelStart = lPosSoFar
    .SelLength = Len(sTranSubTypeText)
    .SelFontName = "Courier New"
    .SelFontSize = 10
    .SelUnderline = True
    .SelBold = False
    lPosSoFar = lPosSoFar + Len(sTranSubTypeText)
    
    'format 'other expenses' Breakdown heading
    .SelStart = lPosSoFar
    .SelLength = Len(sOtherExpensesHeadingText)
    .SelFontName = "Arial"
    .SelFontSize = 10
    .SelUnderline = False
    .SelBold = True
    lPosSoFar = lPosSoFar + Len(sOtherExpensesHeadingText)

    'format 'other expenses' Breakdown
    .SelStart = lPosSoFar
    .SelLength = Len(sOtherExpensesText)
    .SelFontName = "Courier New"
    .SelFontSize = 10
    .SelUnderline = True
    .SelBold = False
    lPosSoFar = lPosSoFar + Len(sOtherExpensesText)
    
    'format 'Circuit expenses' Breakdown heading
    .SelStart = lPosSoFar
    .SelLength = Len(sCircuitContHeadingText)
    .SelFontName = "Arial"
    .SelFontSize = 10
    .SelUnderline = False
    .SelBold = True
    lPosSoFar = lPosSoFar + Len(sCircuitContHeadingText)

    'format 'Circuit expenses' Breakdown
    .SelStart = lPosSoFar
    .SelLength = Len(sCircuitContributionsText)
    .SelFontName = "Courier New"
    .SelFontSize = 10
    .SelUnderline = True
    .SelBold = False
    lPosSoFar = lPosSoFar + Len(sCircuitContributionsText)
    
    'format 'IBSA/WBTS' Breakdown heading
    .SelStart = lPosSoFar
    .SelLength = Len(sContsToSocietyHeadingText)
    .SelFontName = "Arial"
    .SelFontSize = 10
    .SelUnderline = False
    .SelBold = True
    lPosSoFar = lPosSoFar + Len(sContsToSocietyHeadingText)

    'format 'IBSA/WBTS' Breakdown
    .SelStart = lPosSoFar
    .SelLength = Len(sContsToSocietyText)
    .SelFontName = "Courier New"
    .SelFontSize = 10
    .SelUnderline = True
    .SelBold = False
    lPosSoFar = lPosSoFar + Len(sContsToSocietyText)
    
    
    'move to start of text
    .SelStart = 0
    .SelLength = 0
    
    End With
    
    
    On Error Resume Next
    
    rsInTotals.Close
    Set rsInTotals = Nothing
    rsInOtherDetail.Close
    Set rsInOtherDetail = Nothing
    rsGiftAidReceipts.Close
    Set rsGiftAidReceipts = Nothing
    rsOutTotals.Close
    Set rsOutTotals = Nothing
    rsOutOtherDetail.Close
    Set rsOutOtherDetail = Nothing
    rsInGrandTotal.Close
    Set rsInGrandTotal = Nothing
    rsOutGrandTotal.Close
    Set rsOutGrandTotal = Nothing
    rsStartBal.Close
    Set rsStartBal = Nothing
    rsEndBal.Close
    Set rsEndBal = Nothing
    rsInTranTypes.Close
    Set rsInTranTypes = Nothing
    rsOutTranTypes.Close
    Set rsOutTranTypes = Nothing
    rsCircuitExpenses.Close
    Set rsCircuitExpenses = Nothing
    rsContsToSociety.Close
    Set rsContsToSociety = Nothing
    rsBkGrpConts.Close
    Set rsBkGrpConts = Nothing
    rsStartBalOtherAccs.Close
    Set rsStartBalOtherAccs = Nothing
    rsEndBalOtherAccs.Close
    Set rsEndBalOtherAccs = Nothing
    
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub txtLastDate_LostFocus()

    txtLastDate = Format(txtLastDate, "dd/mm/yyyy")
    
    CheckDatesAndRun

End Sub
