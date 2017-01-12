VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAdvancedCongStats2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.M.S. Advanced Congregation Stats - Screen 2"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   Icon            =   "frmAdvancedCongStats2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   75
      Top             =   5055
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   345
      Left            =   3330
      TabIndex        =   54
      Top             =   8895
      Width           =   930
   End
   Begin VB.CommandButton cmdBaptised 
      Caption         =   "&Baptised"
      Height          =   345
      Left            =   2340
      TabIndex        =   48
      Top             =   8895
      Width           =   930
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   345
      Left            =   4335
      TabIndex        =   12
      Top             =   8895
      Width           =   930
   End
   Begin VB.Frame Frame2 
      Caption         =   "Individual Averages"
      ForeColor       =   &H00FF0000&
      Height          =   3660
      Left            =   180
      TabIndex        =   44
      Top             =   5145
      Width           =   8230
      Begin MSFlexGridLib.MSFlexGrid flxIndividualAverages 
         Height          =   3360
         Left            =   60
         TabIndex        =   45
         Top             =   240
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   5927
         _Version        =   393216
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         ScrollBars      =   2
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   345
      Left            =   5325
      TabIndex        =   13
      Top             =   8895
      Width           =   930
   End
   Begin VB.Frame fraStats 
      Caption         =   "Summary Stats"
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   180
      TabIndex        =   20
      Top             =   3180
      Width           =   8230
      Begin VB.TextBox txtAvgTra 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   6420
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   645
         Width           =   885
      End
      Begin VB.TextBox txtTotTra 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   6420
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   990
         Width           =   885
      End
      Begin VB.CommandButton cmdCopyToClipBoard 
         Caption         =   "Copy to Clipboard"
         Height          =   330
         Left            =   1080
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "Paste into MS Word, then convert text to table."
         Top             =   1410
         Width           =   1455
      End
      Begin VB.TextBox txtCount 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   6420
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1410
         Width           =   885
      End
      Begin VB.TextBox txtTotStu 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   990
         Width           =   885
      End
      Begin VB.TextBox txtTotRVs 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   4635
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   990
         Width           =   885
      End
      Begin VB.TextBox txtTotMags 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   3750
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   990
         Width           =   885
      End
      Begin VB.TextBox txtTotHrs 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   2865
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   990
         Width           =   885
      End
      Begin VB.TextBox txtTotBro 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   1980
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   990
         Width           =   885
      End
      Begin VB.TextBox txtTotBooks 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   990
         Width           =   885
      End
      Begin VB.TextBox txtAvgStu 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   645
         Width           =   885
      End
      Begin VB.TextBox txtAvgRVs 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   4635
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   645
         Width           =   885
      End
      Begin VB.TextBox txtAvgMags 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   3750
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   645
         Width           =   885
      End
      Begin VB.TextBox txtAvgHrs 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   2865
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   645
         Width           =   885
      End
      Begin VB.TextBox txtAvgBro 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   1980
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   645
         Width           =   885
      End
      Begin VB.TextBox txtAvgBooks 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   645
         Width           =   885
      End
      Begin VB.Label Label8 
         Caption         =   "Tracts"
         Height          =   255
         Left            =   6450
         TabIndex        =   57
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Count"
         Height          =   255
         Left            =   5655
         TabIndex        =   43
         Top             =   1455
         Width           =   705
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Totals"
         Height          =   255
         Left            =   600
         TabIndex        =   40
         Top             =   1035
         Width           =   420
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Av"
         Height          =   255
         Left            =   705
         TabIndex        =   39
         Top             =   690
         Width           =   315
      End
      Begin VB.Label Label16 
         Caption         =   "Bible Studies"
         Height          =   435
         Left            =   5685
         TabIndex        =   38
         Top             =   270
         Width           =   525
      End
      Begin VB.Label Label15 
         Caption         =   "Return Visits"
         Height          =   450
         Left            =   4800
         TabIndex        =   37
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label14 
         Caption         =   "Magazines"
         Height          =   255
         Left            =   3780
         TabIndex        =   36
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label13 
         Caption         =   "Hours"
         Height          =   255
         Left            =   3105
         TabIndex        =   35
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label12 
         Caption         =   "Brochures"
         Height          =   255
         Left            =   2040
         TabIndex        =   34
         Top             =   270
         Width           =   840
      End
      Begin VB.Label Label11 
         Caption         =   "Books"
         Height          =   255
         Left            =   1275
         TabIndex        =   33
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Criteria"
      ForeColor       =   &H00FF0000&
      Height          =   2790
      Left            =   150
      TabIndex        =   14
      Top             =   255
      Width           =   8230
      Begin VB.CheckBox chkExclHourCredits 
         Caption         =   "Exclude Hour Credits"
         Height          =   210
         Left            =   4005
         TabIndex        =   52
         Top             =   2430
         Width           =   1830
      End
      Begin VB.ComboBox cmbBookgroups 
         Height          =   315
         Left            =   3645
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   855
         Width           =   2685
      End
      Begin VB.CheckBox chkIncludeZeroHours 
         Caption         =   "Include Zero Hours"
         Height          =   210
         Left            =   1455
         TabIndex        =   50
         Top             =   2430
         Width           =   1830
      End
      Begin VB.CheckBox chkIncSpecPios 
         Caption         =   "Include Special Pioneers"
         Height          =   210
         Left            =   4005
         TabIndex        =   49
         Top             =   2145
         Width           =   2115
      End
      Begin VB.CheckBox chkCurrentMembersOnly 
         Caption         =   "Current Cong Members Only"
         Height          =   210
         Left            =   1455
         TabIndex        =   47
         Top             =   2145
         Width           =   2400
      End
      Begin VB.ComboBox cmbAttribute2 
         Height          =   315
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   2130
      End
      Begin VB.ComboBox cmbRelOp2 
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1680
         Width           =   675
      End
      Begin VB.TextBox txtValue2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4350
         MaxLength       =   8
         TabIndex        =   10
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox cmbNames 
         Height          =   315
         Left            =   3645
         TabIndex        =   4
         Text            =   "cmbNames"
         Top             =   855
         Width           =   2685
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "&Run"
         Default         =   -1  'True
         Height          =   360
         Left            =   6225
         TabIndex        =   11
         Top             =   2295
         Width           =   825
      End
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4350
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1290
         Width           =   855
      End
      Begin VB.ComboBox cmbRelOp 
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1290
         Width           =   675
      End
      Begin VB.ComboBox cmbAttribute 
         Height          =   315
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1290
         Width           =   2130
      End
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   5340
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   435
         Width           =   765
      End
      Begin VB.ComboBox cmbWhichPeople 
         Height          =   315
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   855
         Width           =   2130
      End
      Begin VB.ComboBox cmbPeriod 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   435
         Width           =   1725
      End
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Left            =   3825
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   435
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "and"
         Height          =   285
         Left            =   1125
         TabIndex        =   46
         Top             =   1740
         Width           =   300
      End
      Begin VB.Label Label4 
         Caption         =   "with"
         Height          =   285
         Left            =   1125
         TabIndex        =   41
         Top             =   1350
         Width           =   300
      End
      Begin VB.Label Label3 
         Caption         =   "for"
         Height          =   285
         Left            =   1140
         TabIndex        =   19
         Top             =   915
         Width           =   285
      End
      Begin VB.Label Label2 
         Caption         =   "to"
         Height          =   285
         Left            =   3630
         TabIndex        =   18
         Top             =   480
         Width           =   210
      End
      Begin VB.Label Label23 
         Caption         =   "Month"
         Height          =   255
         Left            =   3825
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "Srv Year"
         Height          =   315
         Left            =   5340
         TabIndex        =   16
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Field Ministry in"
         Height          =   285
         Left            =   675
         TabIndex        =   15
         Top             =   495
         Width           =   1140
      End
   End
   Begin VB.Menu Actions 
      Caption         =   "mnuActions"
      Visible         =   0   'False
      Begin VB.Menu mnuPubDetails 
         Caption         =   "Show Details"
      End
      Begin VB.Menu mnuReports 
         Caption         =   "Show Ministry for"
      End
   End
End
Attribute VB_Name = "frmAdvancedCongStats2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlEndYear As Long, mlEndMonth As Long, mlThePerson As Long
Dim mbDatesChanged As Boolean, mbCreditOptionChanged As Boolean, rstGridSQL As Recordset, mbGridFilled As Boolean
Dim msGridSQL As String, mbRecSetActive As Boolean, msReportSQL As String, msOrderBy As String
Dim msCriteria As String, mbHitRunButton As Boolean, mlBookGroup As Long
Dim mbReportsUpdated As Boolean, mbChangedCurrentPubsFlag As Boolean
Dim mlNoRows As Long, mbForm2RebuiltTable As Boolean


Private Sub cmdExport_Click()

On Error GoTo ErrorTrap

    ExportFlexGridToCSV gsDocsDirectory, _
                        "Congregation Stats ", _
                        "csv", _
                        "C.M.S. Export Stats to CSV File", _
                        flxIndividualAverages, _
                        CommonDialog1

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdCopyToClipBoard_Click()

On Error GoTo ErrorTrap

Dim str  As String

    str = ",Books,Brochures,Hours,Magazines,Return Visits, Studies, Tracts" & vbCrLf
    str = str & _
          "Averages," & txtAvgBooks & "," & txtAvgBro & "," & txtAvgHrs & "," & txtAvgMags & _
          "," & txtAvgRVs & "," & txtAvgStu & "," & txtAvgTra & vbCrLf & _
          "Totals," & txtTotBooks & "," & txtTotBro & "," & txtTotHrs & "," & txtTotMags & _
          "," & txtTotRVs & "," & txtTotStu & "," & txtTotTra & vbCrLf & "Publishers," & txtCount
          
    Clipboard.Clear
    Clipboard.SetText str
    
    ShowMessage "Ministry copied to clipboard", 1500, Me

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap

    BringForwardMainMenuWhenItsTheLastFormOpen

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub chkCurrentMembersOnly_Click()
    
On Error GoTo ErrorTrap

    mbChangedCurrentPubsFlag = True
    
    ClearResults
    
    If mbHitRunButton Then
        RunQuery
    End If

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub chkExclHourCredits_Click()
On Error GoTo ErrorTrap

    mbCreditOptionChanged = True
    
    ClearResults
    
    If mbHitRunButton Then
        RunQuery
    End If

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub



Private Sub chkIncludeZeroHours_Click()
On Error GoTo ErrorTrap

    ClearResults
    
    If mbHitRunButton Then
        RunQuery
    End If

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub chkIncSpecPios_Click()
On Error GoTo ErrorTrap

    ClearResults
    
    If mbHitRunButton Then
        RunQuery
    End If
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmbAttribute_Click()
On Error GoTo ErrorTrap

    If cmbAttribute.ItemData(cmbAttribute.ListIndex) = 0 Then
        cmbRelOp.Visible = False
        txtValue.Visible = False
    Else
        cmbRelOp.Visible = True
        txtValue.Visible = True
    End If
    
    ClearResults

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub cmbAttribute2_Click()
On Error GoTo ErrorTrap

    If cmbAttribute2.ItemData(cmbAttribute2.ListIndex) = 0 Then
        cmbRelOp2.Visible = False
        txtValue2.Visible = False
    Else
        cmbRelOp2.Visible = True
        txtValue2.Visible = True
    End If
    
    ClearResults

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub cmbBookgroups_Click()
On Error GoTo ErrorTrap

    If cmbBookgroups.ListIndex > -1 Then
        mlBookGroup = cmbBookgroups.ItemData(cmbBookgroups.ListIndex)
    Else
        mlBookGroup = 0
    End If
    
    ClearResults

    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub cmbMonth_Click()
On Error GoTo ErrorTrap

    If cmbMonth.ListIndex > -1 Then
        mlEndMonth = cmbMonth.ItemData(cmbMonth.ListIndex)
    Else
        mlEndMonth = 0
    End If
    
    mbDatesChanged = True

    ClearResults

    Exit Sub
ErrorTrap:
    EndProgram


End Sub



Private Sub cmbNames_Click()
On Error GoTo ErrorTrap

    If cmbNames.ListIndex > -1 Then
        mlThePerson = cmbNames.ItemData(cmbNames.ListIndex)
    Else
        mlThePerson = 0
    End If
    
    ClearResults

    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub cmbPeriod_Click()
    ClearResults
    mbDatesChanged = True
End Sub

Private Sub cmbRelOp_Click()
    ClearResults
End Sub
Private Sub cmbRelOp2_Click()
    ClearResults
End Sub

Private Sub cmbWhichPeople_Click()
On Error GoTo ErrorTrap

    Select Case cmbWhichPeople.ItemData(cmbWhichPeople.ListIndex)
    Case 8 'person
        cmbNames.Visible = True
        cmbBookgroups.Visible = False
        cmbAttribute.ListIndex = 0 'for a specific person don't need attribute set
        cmbAttribute.Enabled = False
        cmbAttribute2.ListIndex = 0 'for a specific person don't need attribute set
        cmbAttribute2.Enabled = False
    Case 11 'Group
        cmbNames.Visible = False
        cmbBookgroups.Visible = True
        cmbAttribute.Enabled = True
        cmbAttribute2.Enabled = True
    Case Else
        cmbNames.Visible = False
        cmbBookgroups.Visible = False
        cmbAttribute.Enabled = True
        cmbAttribute2.Enabled = True
    End Select
    
    ClearResults

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmbYear_Click()
On Error GoTo ErrorTrap

    If cmbYear.ListIndex > -1 Then
        mlEndYear = CLng(cmbYear.text)
    Else
        mlEndYear = 0
    End If
    
    ClearResults
    
    mbDatesChanged = True

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdBaptised_Click()

On Error GoTo ErrorTrap

    frmBaptised.Show vbModal, Me

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdPrint_Click()
On Error GoTo ErrorTrap

    mbRecSetActive = False

    If mlNoRows > 0 Then
        PrintAdvancedReport msReportSQL
    Else
        ShowMessage "Nothing to print", 800, Me
    End If
        
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdRun_Click()
On Error GoTo ErrorTrap

    RunQuery
    
    mbHitRunButton = True

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub RunQuery()
On Error GoTo ErrorTrap

    Screen.MousePointer = vbHourglass
    cmdRun.Enabled = False
    cmdRun.Caption = "Wait..."

    
    If ValidQuery Then
        ConstructAndRunQuery
        cmdPrint.Enabled = True
        cmdExport.Enabled = True
    Else
        cmdPrint.Enabled = False
        cmdExport.Enabled = False
    End If
    
    cmdRun.Enabled = True
    cmdRun.Caption = "&Run"
    Screen.MousePointer = vbNormal


    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub BuildReportingDB(StartDate_US As String, _
                             EndDate_US As String)
On Error GoTo ErrorTrap
Dim TheSQL As String, rstPubs As Recordset, MinDate_US As String
Dim MinDate_UK As String, EndDate_UK As String, rstStats As Recordset
Dim TheSQL2 As String, rstReportingTbl As Recordset, StartDate_UK As String
Dim TempPubStartDate_UK As String, TempPubStartDate_US As String
Dim TempPubEndDate_UK As String, TempPubEndDate_US As String
Dim bUseAltPubDates As Boolean, TempNoMonths As Long
Dim dTempAvgHours As Double, dTempAvgRVs As Double, dTempAvgBooks As Double
Dim dTempAvgBooklets As Double, dTempAvgStudies As Double, dTempAvgMags As Double
Dim dTempAvgTracts As Double
Dim rstTempPubdata As Recordset, TempPubSQL As String
Dim bArtificialRegPio As Boolean, bArtificialSpecPio As Boolean, lHourCredit As Long

    bUseAltPubDates = False
    
    mbForm2RebuiltTable = True
    frmAdvancedCongStats.Form1RebuiltTable = False
    
    DelAllRows "tblAdvancedMinreporting"
    
    Set rstReportingTbl = CMSDB.OpenRecordset("tblAdvancedMinreporting", dbOpenDynaset)
    
    '
    'put all dates in range for each publisher, and add any ministry done
    '
'    TheSQL = "SELECT DISTINCT PersonID " & _
'                "FROM tblPublisherDates " & _
'                " WHERE StartDate <= #" & _
'                 EndDate_US & _
'                "# AND EndDate >= #" & StartDate_US & "# "

    TheSQL = "SELECT MAX(StartDate), PersonID " & _
                "FROM tblPublisherDates " & _
                " WHERE StartDate <= #" & _
                 EndDate_US & _
                "# AND EndDate >= #" & StartDate_US & "# " & _
                IIf(chkCurrentMembersOnly.value = vbChecked, " AND StartReason <> 2 ", " ") & _
                "GROUP BY PersonID"

    Set rstPubs = CMSDB.OpenRecordset(TheSQL, dbOpenForwardOnly)
    
    TheSQL2 = "SELECT * FROM tblMinReports " & _
                    "INNER JOIN tblPublisherDates ON tblPublisherDates.PersonID = " & _
                                 " tblMinReports.PersonID" & _
                " WHERE ActualMinPeriod " & _
                " BETWEEN #" & StartDate_US & "# AND #" & EndDate_US & "#"
    
    Set rstStats = CMSDB.OpenRecordset(TheSQL2, dbOpenDynaset)
    
    EndDate_UK = Format$(EndDate_US, "mm/dd/yyyy")
    StartDate_UK = Format$(StartDate_US, "mm/dd/yyyy")
    
    With rstPubs
    
    Do Until .BOF Or .EOF 'for each person
    
        MinDate_US = StartDate_US
        MinDate_UK = Format$(StartDate_US, "mm/dd/yyyy")
        
        Do Until CDate(MinDate_UK) > CDate(EndDate_UK) ' for each date
        
            'find min done for person (outer loop) for date (inner loop)
            rstStats.FindFirst "tblMinReports.PersonID = " & rstPubs!PersonID & _
                        " AND ActualMinPeriod = #" & MinDate_US & "#"
            
            'add person and their ministry to reporting table
            'if no ministry found for month, insert zeroes
            With rstReportingTbl
            .AddNew
            !PersonID = rstPubs!PersonID
            !ActualMinDate = MinDate_UK
            !BookGroupID = CongregationMember.BookGroup(rstPubs!PersonID)
            If Not rstStats.NoMatch Then
                If chkExclHourCredits = vbUnchecked Then
                    lHourCredit = CongregationMember.GetPioHourCredit(rstPubs!PersonID, CDate(MinDate_UK), CDate(MinDate_UK))
                Else
                    lHourCredit = 0
                End If
                !NoHours = rstStats!NoHours + lHourCredit
                !NoBooks = rstStats!NoBooks
                !NoBooklets = rstStats!NoBooklets
                !NoMagazines = rstStats!NoMagazines
                !NoReturnVisits = rstStats!NoReturnVisits
                !NoStudies = rstStats!NoStudies
                !NoTracts = rstStats!NoTracts
                !ArtificialValue = False
            Else
                'no min reported for this date
                ' Is person a publisher in this month?
                If CongregationMember.IsPublisher(rstPubs!PersonID, CDate(MinDate_UK)) Then
                    'person is publisher, just didn't report this month!
                    !NoHours = 0
                    !NoBooks = 0
                    !NoBooklets = 0
                    !NoMagazines = 0
                    !NoReturnVisits = 0
                    !NoStudies = 0
                    !NoTracts = 0
                    !ArtificialValue = False
                    bUseAltPubDates = False
                    bArtificialRegPio = False
                    bArtificialSpecPio = False
                Else
                    'person is not a publisher at this point, so must
                    ' populate fields with their average during time as a
                    ' publisher in order to display meaningful averages for
                    ' them later...
                    !ArtificialValue = True
                    If Not bUseAltPubDates Then
                        'not yet found pub dates and averages for this pub, so
                        ' find them now...
                    
                        bUseAltPubDates = True
                        
                        'find the publisher's start/end dates. Ensure they
                        ' are within the start/end reporting dates
                        TempPubEndDate_UK = CStr(CongregationMember.PublisherEndDate( _
                                                rstPubs!PersonID, CDate(StartDate_UK)))
                        If CDate(TempPubEndDate_UK) > CDate(EndDate_UK) Then
                            TempPubEndDate_UK = EndDate_UK
                        End If
                        
                        TempPubStartDate_UK = CStr(CongregationMember.PublisherStartDate( _
                                                rstPubs!PersonID, CDate(EndDate_UK)))
                        If CDate(TempPubStartDate_UK) < CDate(StartDate_UK) Then
                            TempPubStartDate_UK = StartDate_UK
                        End If
                        
                        'determine if person is reg pio or spec pio on first month
                        ' of their being a pub in cong. If so, this is the MinType that'll
                        ' be used for the 'artificial' figures.
                        Select Case True
                        Case CongregationMember.IsRegPio(rstPubs!PersonID, CDate(TempPubStartDate_UK))
                            bArtificialRegPio = True
                            bArtificialSpecPio = False
                        Case CongregationMember.IsSpecPio(rstPubs!PersonID, CDate(TempPubStartDate_UK))
                            bArtificialRegPio = False
                            bArtificialSpecPio = True
                        Case Else
                            bArtificialRegPio = False
                            bArtificialSpecPio = False
                        End Select
                        
                        TempPubStartDate_US = Format$(TempPubStartDate_UK, "mm/dd/yyyy")
                        TempPubEndDate_US = Format$(TempPubEndDate_UK, "mm/dd/yyyy")
                        
                        'now find averages between the calculated dates
                        TempNoMonths = Abs(DateDiff("m", TempPubStartDate_UK, TempPubEndDate_UK)) + 1
                        
                        TempPubSQL = "SELECT SUM(NoHours) AS TempHours, " & _
                                            "SUM(NoBooks) AS TempBooks, " & _
                                            "SUM(NoBooklets) AS TempBooklets, " & _
                                            "SUM(NoMagazines) AS TempMagazines, " & _
                                            "SUM(NoReturnVisits) AS TempReturnVisits, " & _
                                            "SUM(NoStudies) AS TempStudies, " & _
                                            "SUM(NoTracts) AS TempTracts " & _
                                            "FROM tblMinReports " & _
                                            "WHERE ActualMinPeriod BETWEEN #" & _
                                                TempPubStartDate_US & "# AND #" & _
                                                TempPubEndDate_US & "# " & _
                                            "AND PersonID = " & rstPubs!PersonID
                                            
                        Set rstTempPubdata = CMSDB.OpenRecordset(TempPubSQL, dbOpenForwardOnly)
                        
                        With rstTempPubdata
                        
                        If Not .BOF Then
                            If chkExclHourCredits = vbUnchecked Then
                                lHourCredit = CongregationMember.GetPioHourCredit(rstPubs!PersonID, CDate(TempPubStartDate_UK), CDate(TempPubEndDate_UK))
                            Else
                                lHourCredit = 0
                            End If
                            dTempAvgHours = (IIf(IsNull(!TempHours), 0, !TempHours) + lHourCredit) / TempNoMonths
                            dTempAvgBooklets = IIf(IsNull(!TempBooklets), 0, !TempBooklets) / TempNoMonths
                            dTempAvgBooks = IIf(IsNull(!TempBooks), 0, !TempBooks) / TempNoMonths
                            dTempAvgMags = IIf(IsNull(!TempMagazines), 0, !TempMagazines) / TempNoMonths
                            dTempAvgRVs = IIf(IsNull(!TempReturnVisits), 0, !TempReturnVisits) / TempNoMonths
                            dTempAvgStudies = IIf(IsNull(!TempStudies), 0, !TempStudies) / TempNoMonths
                            dTempAvgTracts = IIf(IsNull(!TempTracts), 0, !TempTracts) / TempNoMonths
                        Else
                            dTempAvgHours = 0
                            dTempAvgBooklets = 0
                            dTempAvgBooks = 0
                            dTempAvgMags = 0
                            dTempAvgRVs = 0
                            dTempAvgStudies = 0
                            dTempAvgTracts = 0
                        End If
                        
                        End With
                        
                        'save to db
                        !NoHours = dTempAvgHours
                        !NoBooks = dTempAvgBooks
                        !NoBooklets = dTempAvgBooklets
                        !NoMagazines = dTempAvgMags
                        !NoReturnVisits = dTempAvgRVs
                        !NoStudies = dTempAvgStudies
                        !NoTracts = dTempAvgTracts
                        
                    Else
                        'already found dates and averages for this pub, so use
                        ' saved values...
                        !NoHours = dTempAvgHours
                        !NoBooks = dTempAvgBooks
                        !NoBooklets = dTempAvgBooklets
                        !NoMagazines = dTempAvgMags
                        !NoReturnVisits = dTempAvgRVs
                        !NoStudies = dTempAvgStudies
                        !NoTracts = dTempAvgTracts
                    End If
                End If
            End If
            
            If bUseAltPubDates Then
                'for artificially added figures, what mintype is it?
                If bArtificialRegPio Then
                    !MinType = IsRegPio
                ElseIf bArtificialSpecPio Then
                    !MinType = IsSpecPio
                Else
                    !MinType = IsPublisher
                End If
            Else
                Select Case True
                Case CongregationMember.IsAuxPio(rstPubs!PersonID, CDate(MinDate_UK))
                    !MinType = IsAuxPio
                Case CongregationMember.IsRegPio(rstPubs!PersonID, CDate(MinDate_UK))
                    !MinType = IsRegPio
                Case CongregationMember.IsSpecPio(rstPubs!PersonID, CDate(MinDate_UK))
                    !MinType = IsSpecPio
                Case Else
                    !MinType = IsPublisher
                End Select
            End If
            
            Select Case CongregationMember.BaptismDate(rstPubs!PersonID)
            Case 0
                 !IsBaptised = False
            Case Is <= CDate(MinDate_UK)
                !IsBaptised = True
            Case Else
                !IsBaptised = False
            End Select
            
            Select Case CongregationMember.ElderDate(rstPubs!PersonID)
            Case 0
                 !ElderMS = ""
            Case Is <= CDate(MinDate_UK)
                !ElderMS = "E"
            Case Else
                !ElderMS = ""
            End Select
            
            Select Case CongregationMember.ServantDate(rstPubs!PersonID)
            Case 0
                 !ElderMS = ""
            Case Is <= CDate(MinDate_UK)
                !ElderMS = "MS"
            Case Else
                !ElderMS = ""
            End Select
            
            .Update
            End With
                                                           
            'add one month then repeat inner loop
            MinDate_UK = DateAdd("m", 1, MinDate_UK)
            MinDate_US = Format$(MinDate_UK, "mm/dd/yyyy")
            
        Loop
        .MoveNext 'go to next person and repeat outer loop
        bUseAltPubDates = False
        bArtificialRegPio = False
        bArtificialSpecPio = False
    Loop
    
    End With
    
    mbDatesChanged = False
    mbCreditOptionChanged = False
    
    rstReportingTbl.Close
    rstStats.Close
    rstPubs.Close
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub ConstructAndRunQuery()
On Error GoTo ErrorTrap
Dim DateSQL As String, StartDate_US As String, EndDate_US As String
Dim lsNormalYear As String, lsNormalYear2 As String, StartDate_UK As String
Dim EndDate_UK As String, rstGetPubsCount As Recordset, CountPubsSQL As String
Dim mlNoMonths As Long, SQLStr As String, rstStats As Recordset, rstNamesStats As Recordset
Dim AgeStr As String, AgeStr2 As String
Dim TheRelOp As String, TheRelOp2 As String
Dim sElderStr As String, strPerson As String, AvgSQL As String, MinTypeSQL As String
Dim GenderSQL As String, BaptismSQL As String, YearsBaptisedSQL As String, AvgRVsSQL As String
Dim YearsBaptisedSQL2 As String, FullWHERESQL As String, ActiveSQL As String
Dim SpecPioSQL As String, ZeroHourSQL As String, lMths As Long, sBookGroup As String

    'Date up to which we're interested, as selected in combos
    lsNormalYear2 = CStr(ConvertServiceYearToNormalYear(CDate("01/" & mlEndMonth & "/" & mlEndYear)))
    EndDate_UK = "01/" & CStr(mlEndMonth) & "/" & lsNormalYear2
    EndDate_US = CStr(mlEndMonth) & "/01/" & lsNormalYear2

'    'Start date
'    Select Case cmbPeriod.ItemData(cmbPeriod.ListIndex)
'    Case 0
'        mlNoMonths = -2
'    Case 1
'        mlNoMonths = -5
'    Case 2
'        mlNoMonths = -11
'    End Select

    mlNoMonths = cmbPeriod.ItemData(cmbPeriod.ListIndex) * -1
    
    StartDate_UK = DateAdd("m", mlNoMonths, EndDate_UK)
    StartDate_US = Format$(StartDate_UK, "mm/dd/yyyy")
    
    'build reporting db if date range changes - prevents unnecessary table
    ' accesses
    If mbDatesChanged Or mbCreditOptionChanged Or mbReportsUpdated _
        Or frmAdvancedCongStats.Form1RebuiltTable Or mbChangedCurrentPubsFlag Then
        BuildReportingDB StartDate_US, EndDate_US
        mbReportsUpdated = False
    End If
    
    Select Case cmbRelOp.ListIndex
    Case 0
        TheRelOp = "="
    Case 1
        TheRelOp = "<"
    Case 2
        TheRelOp = ">"
    Case 3
        TheRelOp = "<="
    Case 4
        TheRelOp = ">="
    Case 5
        TheRelOp = "<>"
    End Select
    
    Select Case cmbRelOp2.ListIndex
    Case 0
        TheRelOp2 = "="
    Case 1
        TheRelOp2 = "<"
    Case 2
        TheRelOp2 = ">"
    Case 3
        TheRelOp2 = "<="
    Case 4
        TheRelOp2 = ">="
    Case 5
        TheRelOp2 = "<>"
    End Select
    
    'Age criteria
    If cmbAttribute.ItemData(cmbAttribute.ListIndex) = 1 Then
        AgeStr = " AND Year(DOB) " & TheRelOp & _
                 year(DateAdd("yyyy", -1 * CDbl(txtValue), Now))
    Else
        AgeStr = " "
    End If
    
    If cmbAttribute2.ItemData(cmbAttribute2.ListIndex) = 1 Then
        AgeStr2 = " AND Year(DOB) " & TheRelOp2 & _
                 year(DateAdd("yyyy", -1 * CDbl(txtValue2), Now))
    Else
        AgeStr2 = " "
    End If
        
    'Years baptised criteria
    If cmbAttribute.ItemData(cmbAttribute.ListIndex) = 2 Then
        YearsBaptisedSQL = " AND PersonID IN " & _
                            "(SELECT PersonID " & _
                            " FROM tblBaptismDates " & _
                            " WHERE Year(BaptismDate) " & TheRelOp & _
                 year(DateAdd("yyyy", -1 * CDbl(txtValue), Now)) & ") "
    Else
        YearsBaptisedSQL2 = " "
    End If
    
    If cmbAttribute2.ItemData(cmbAttribute2.ListIndex) = 2 Then
        YearsBaptisedSQL2 = " AND PersonID IN " & _
                            "(SELECT PersonID " & _
                            " FROM tblBaptismDates " & _
                            " WHERE Year(BaptismDate) " & TheRelOp2 & _
                 year(DateAdd("yyyy", -1 * CDbl(txtValue2), Now)) & ") "
    Else
        YearsBaptisedSQL2 = " "
    End If
    
    'elder / servant criteria
    Select Case cmbWhichPeople.ItemData(cmbWhichPeople.ListIndex)
    Case 4 'elders. Only include men appointed prior to startdate
        sElderStr = " AND PersonID IN " & _
                    "(SELECT PersonID FROM tblEldersAndServants " & _
                    " WHERE AppointmentDate <= # " & EndDate_US & "# " & _
                    " AND ElderOrServant = 'E') "
    Case 5 'servants
        sElderStr = " AND PersonID IN " & _
                    "(SELECT PersonID FROM tblEldersAndServants " & _
                    " WHERE AppointmentDate <= # " & EndDate_US & "# " & _
                    " AND ElderOrServant = 'MS') "
    Case Else
        sElderStr = ""
    End Select
        
    'specific person criteria
    Select Case cmbWhichPeople.ItemData(cmbWhichPeople.ListIndex)
    Case 8 'person
        strPerson = " AND PersonID = " & mlThePerson & " "
    Case Else
        strPerson = ""
    End Select
        
    'specific bookgroup criteria
    Select Case cmbWhichPeople.ItemData(cmbWhichPeople.ListIndex)
    Case 11 'bookgroups
        sBookGroup = " AND BookGroupID = " & mlBookGroup & " "
    Case Else
        sBookGroup = ""
    End Select
    
    'Ministry Type
    Select Case cmbWhichPeople.ItemData(cmbWhichPeople.ListIndex)
    Case 1
        MinTypeSQL = " AND MinType = " & IsPublisher & " "
    Case 2
        MinTypeSQL = " AND MinType = " & IsRegPio & " "
    Case 3
        MinTypeSQL = " AND MinType = " & IsAuxPio & " "
    Case Else
        MinTypeSQL = " "
    End Select
        
    'Gender
    Select Case cmbWhichPeople.ItemData(cmbWhichPeople.ListIndex)
    Case 9
        GenderSQL = " AND GenderMF = 'M' "
    Case 10
        GenderSQL = " AND GenderMF = 'F' "
    Case Else
        GenderSQL = " "
    End Select
    
    'Baptism SQL
    Select Case cmbWhichPeople.ItemData(cmbWhichPeople.ListIndex)
    Case 6
        BaptismSQL = " AND IsBaptised = True "
    Case 7
        BaptismSQL = " AND IsBaptised = False "
    Case Else
        BaptismSQL = " "
    End Select
    
    'Include current cong members only?
    If chkCurrentMembersOnly.value Then
        ActiveSQL = " AND Active = TRUE "
    Else
        ActiveSQL = " "
    End If
    
    'Include reports where Hours are zero?
    If chkIncludeZeroHours.value Then
        ZeroHourSQL = " "
    Else
        ZeroHourSQL = " AND NoHours > 0 "
    End If
    
    'Include Special Pioneers?
    If chkIncSpecPios.value Then
        SpecPioSQL = " "
    Else
        If MinTypeSQL = " " Then
            SpecPioSQL = " AND MinType <> " & IsSpecPio & " "
        Else
            SpecPioSQL = " "
        End If
    End If
    
    '
    'now find totals for all selected publishers
    '
    FullWHERESQL = AgeStr & AgeStr2 & sElderStr & strPerson & MinTypeSQL & _
                             BaptismSQL & GenderSQL & YearsBaptisedSQL & YearsBaptisedSQL2 & _
                             ActiveSQL & SpecPioSQL & ZeroHourSQL & sBookGroup
    
    Select Case cmbAttribute.ItemData(cmbAttribute.ListIndex)
    'criteria includes average hrs, so different SQL is required
    Case 3
        Select Case cmbAttribute2.ItemData(cmbAttribute2.ListIndex)
        Case 3
            SQLStr = MinSQLStatsTotals(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 1, _
                                 cmbRelOp2.text, txtValue2.text, 1)
        Case 4
            SQLStr = MinSQLStatsTotals(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 1, _
                                 cmbRelOp2.text, txtValue2.text, 2)
        Case Else
            SQLStr = MinSQLStatsTotals(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 1)
        End Select
    Case 4
        'criteria includes average RVs, so different SQL is required
        Select Case cmbAttribute2.ItemData(cmbAttribute2.ListIndex)
        Case 3
            SQLStr = MinSQLStatsTotals(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 2, _
                                 cmbRelOp2.text, txtValue2.text, 1)
        Case 4
            SQLStr = MinSQLStatsTotals(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 2, _
                                 cmbRelOp2.text, txtValue2.text, 2)
        Case Else
            SQLStr = MinSQLStatsTotals(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 2)
        End Select
    Case Else
        Select Case cmbAttribute2.ItemData(cmbAttribute2.ListIndex)
        Case 3
            SQLStr = MinSQLStatsTotals(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 , , , _
                                 cmbRelOp2.text, txtValue2.text, 1)
        Case 4
            SQLStr = MinSQLStatsTotals(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 , , , _
                                 cmbRelOp2.text, txtValue2.text, 2)
        Case Else
            SQLStr = MinSQLStatsTotals(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL)
        End Select
    
    End Select
    
    Set rstStats = CMSDB.OpenRecordset(SQLStr, dbOpenDynaset)
    
    With rstStats
    
    If Not .BOF Then
        txtTotBro = IIf(Not IsNull(!TotSumBooklets), Round(!TotSumBooklets, 2), 0)
        txtTotBooks = IIf(Not IsNull(!TotSumBooks), Round(!TotSumBooks, 2), 0)
        txtTotHrs = IIf(Not IsNull(!TotSumHours), Round(!TotSumHours, 2), 0)
        txtTotMags = IIf(Not IsNull(!TotSumMagazines), Round(!TotSumMagazines, 2), 0)
        txtTotRVs = IIf(Not IsNull(!TotSumReturnVisits), Round(!TotSumReturnVisits, 2), 0)
        txtTotStu = IIf(Not IsNull(!TotSumStudies), Round(!TotSumStudies, 2), 0)
        txtTotTra = IIf(Not IsNull(!TotSumTracts), Round(!TotSumTracts, 2), 0)
        txtCount = IIf(Not IsNull(!CountPersons), Round(!CountPersons, 2), 0)
    Else
        txtTotBro = 0
        txtTotBooks = 0
        txtTotHrs = 0
        txtTotMags = 0
        txtTotRVs = 0
        txtTotStu = 0
        txtTotTra = 0
        txtCount = 0
    End If
    
    End With
    
    
    '
    'now find  averages for all selected publishers
    '
    Select Case cmbAttribute.ItemData(cmbAttribute.ListIndex)
    'criteria includes average hrs, so different SQL is required
    Case 3
        Select Case cmbAttribute2.ItemData(cmbAttribute2.ListIndex)
        Case 3
            SQLStr = MinSQLStatsAvgs(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 1, _
                                 cmbRelOp2.text, txtValue2.text, 1)
        Case 4
            SQLStr = MinSQLStatsAvgs(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 1, _
                                 cmbRelOp2.text, txtValue2.text, 2)
        Case Else
            SQLStr = MinSQLStatsAvgs(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 1)
        End Select
    Case 4
        'criteria includes average RVs, so different SQL is required
        Select Case cmbAttribute2.ItemData(cmbAttribute2.ListIndex)
        Case 3
            SQLStr = MinSQLStatsAvgs(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 2, _
                                 cmbRelOp2.text, txtValue2.text, 1)
        Case 4
            SQLStr = MinSQLStatsAvgs(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 2, _
                                 cmbRelOp2.text, txtValue2.text, 2)
        Case Else
            SQLStr = MinSQLStatsAvgs(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 2)
        End Select
    Case Else
        Select Case cmbAttribute2.ItemData(cmbAttribute2.ListIndex)
        Case 3
            SQLStr = MinSQLStatsAvgs(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 , , , _
                                 cmbRelOp2.text, txtValue2.text, 1)
        Case 4
            SQLStr = MinSQLStatsAvgs(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 , , , _
                                 cmbRelOp2.text, txtValue2.text, 2)
        Case Else
            SQLStr = MinSQLStatsAvgs(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL)
        End Select
    
    End Select
    
    Set rstStats = CMSDB.OpenRecordset(SQLStr, dbOpenDynaset)
    
    With rstStats
    
    If Not .BOF Then
        lMths = HandleNull(!TotalMonths)
        txtAvgBooks = IIf(Not IsNull(!TotSumBooks), Round(!TotSumBooks / lMths, 2), 0)
        txtAvgBro = IIf(Not IsNull(!TotSumBooklets), Round(!TotSumBooklets / lMths, 2), 0)
        txtAvgHrs = IIf(Not IsNull(!TotSumHours), Round(!TotSumHours / lMths, 2), 0)
        txtAvgMags = IIf(Not IsNull(!TotSumMagazines), Round(!TotSumMagazines / lMths, 2), 0)
        txtAvgRVs = IIf(Not IsNull(!TotSumReturnVisits), Round(!TotSumReturnVisits / lMths, 2), 0)
        txtAvgStu = IIf(Not IsNull(!TotSumStudies), Round(!TotSumStudies / lMths, 2), 0)
        txtAvgTra = IIf(Not IsNull(!TotSumTracts), Round(!TotSumTracts / lMths, 2), 0)
    Else
        txtAvgBooks = 0
        txtAvgBro = 0
        txtAvgHrs = 0
        txtAvgMags = 0
        txtAvgRVs = 0
        txtAvgStu = 0
        txtAvgTra = 0
    End If
    
    End With
    
    
    
    Select Case cmbAttribute.ItemData(cmbAttribute.ListIndex)
    'criteria includes average hrs, so different SQL is required
    Case 3
        Select Case cmbAttribute2.ItemData(cmbAttribute2.ListIndex)
        Case 3
            msGridSQL = MinSQLNames(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 1, _
                                 cmbRelOp2.text, txtValue2.text, 1)
        Case 4
            msGridSQL = MinSQLNames(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 1, _
                                 cmbRelOp2.text, txtValue2.text, 2)
        Case Else
            msGridSQL = MinSQLNames(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 1)
        End Select
    Case 4
        'criteria includes average RVs, so different SQL is required
        Select Case cmbAttribute2.ItemData(cmbAttribute2.ListIndex)
        Case 3
            msGridSQL = MinSQLNames(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 2, _
                                 cmbRelOp2.text, txtValue2.text, 1)
        Case 4
            msGridSQL = MinSQLNames(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 2, _
                                 cmbRelOp2.text, txtValue2.text, 2)
        Case Else
            msGridSQL = MinSQLNames(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 cmbRelOp.text, txtValue.text, 2)
        End Select
    Case Else
        Select Case cmbAttribute2.ItemData(cmbAttribute2.ListIndex)
        Case 3
            msGridSQL = MinSQLNames(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 , , , _
                                 cmbRelOp2.text, txtValue2.text, 1)
        Case 4
            msGridSQL = MinSQLNames(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL, _
                                 , , , _
                                 cmbRelOp2.text, txtValue2.text, 2)
        Case Else
            msGridSQL = MinSQLNames(StartDate_US, _
                                 EndDate_US, _
                                 FullWHERESQL)
        End Select
    
    End Select
    
    FillGrid
    
    '
    'construct descriptive text for print report...
    '
    If StartDate_UK <> EndDate_UK Then
        msCriteria = "Averages between " & Format$(StartDate_UK, "mmmm yyyy") & _
                                         " and " & Format$(EndDate_UK, "mmmm yyyy") & _
                                         " inclusive"
    Else
        msCriteria = "Averages on " & Format$(StartDate_UK, "dd/mm/yyyy")
    End If

    Select Case cmbWhichPeople.ItemData(cmbWhichPeople.ListIndex)
    Case 8 'person
        msCriteria = msCriteria & " for " & _
            CongregationMember.FullNameFromDB(mlThePerson) & " " & vbCrLf
'        msCriteria = msCriteria & " for " & _
'            CongregationMember.FullNameFromDB(cmbNames.ItemData(cmbNames.ListIndex)) & _
'                                                        " " & vbCrLf
    Case 11 'bookgroup
        msCriteria = msCriteria & " for " & _
            GetGroupName(mlBookGroup) & " " & vbCrLf
    Case Else
        msCriteria = msCriteria & " for " & LCase$(cmbWhichPeople.text) & " " & vbCrLf
    End Select
    
    If cmbAttribute.ItemData(cmbAttribute.ListIndex) > 0 Then
        msCriteria = msCriteria & "where " & LCase$(cmbAttribute.text) & " " & _
                                    cmbRelOp.text & " " & txtValue.text
        If cmbAttribute2.ItemData(cmbAttribute2.ListIndex) > 0 Then
            msCriteria = msCriteria & " and " & LCase$(cmbAttribute2.text) & " " & _
                                        cmbRelOp2.text & " " & txtValue2.text & vbCrLf
        End If
    Else
        If cmbAttribute2.ItemData(cmbAttribute2.ListIndex) > 0 Then
            msCriteria = msCriteria & "where " & LCase$(cmbAttribute2.text) & " " & _
                                        cmbRelOp2.text & " " & txtValue2.text & vbCrLf
        End If
    End If
    
    rstStats.Close

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Private Sub Form_Load()
On Error GoTo ErrorTrap
Dim TheYear As Long, TheString As String, i As Long
   
    Me.Left = frmAdvancedCongStats.Left + 567
    Me.Top = frmAdvancedCongStats.Top - 567
   
   '
    'Populate cmbYear_PubNos and cmbMonth_PubNos with past 5 years, current year and next year!
    ' Averages/Totals calculated up to this month/year
    '
    For TheYear = year(Now) - GlobalParms.GetValue("YearsHistoryToInclude_Min", "NumVal") To year(Now) + 2
        cmbYear.AddItem TheYear
    Next TheYear
    
    TheString = CStr(LastReportMonth)
    cmbYear.text = CStr(ServiceYear(CDate(TheString)))
    
'    TheString = GetSocietyReportingPeriodMMYYYY(Now)
'    cmbYear.text = CStr(ServiceYear(CDate("01/" & TheString)))
    cmbYear_Click

    HandleListBox.PopulateListBox Me!cmbMonth, "SELECT MonthNum, " & _
                                               "       MonthName " & _
                                               "FROM tblMonthName " & _
                                               "ORDER BY OrderForServiceYear ASC", _
                                               CMSDB, 0, "", False, 1
    
    HandleListBox.SelectItem Me!cmbMonth, CLng(Month(DateAdd("m", -1, Now))) 'set to last month
    
    '
    'Now populate cmbPeriod. Want to be able to calc averages & totals over
    ' period of 3/6/12 months...
    '
    cmbPeriod.AddItem "1 Month"
    cmbPeriod.ItemData(cmbPeriod.NewIndex) = 0
    cmbPeriod.AddItem "2 Months"
    cmbPeriod.ItemData(cmbPeriod.NewIndex) = 1
    cmbPeriod.AddItem "3 Months"
    cmbPeriod.ItemData(cmbPeriod.NewIndex) = 2
    cmbPeriod.AddItem "4 Months"
    cmbPeriod.ItemData(cmbPeriod.NewIndex) = 3
    cmbPeriod.AddItem "5 Months"
    cmbPeriod.ItemData(cmbPeriod.NewIndex) = 4
    cmbPeriod.AddItem "6 Months"
    cmbPeriod.ItemData(cmbPeriod.NewIndex) = 5
    cmbPeriod.AddItem "7 Months"
    cmbPeriod.ItemData(cmbPeriod.NewIndex) = 6
    cmbPeriod.AddItem "8 Months"
    cmbPeriod.ItemData(cmbPeriod.NewIndex) = 7
    cmbPeriod.AddItem "9 Months"
    cmbPeriod.ItemData(cmbPeriod.NewIndex) = 8
    cmbPeriod.AddItem "10 Months"
    cmbPeriod.ItemData(cmbPeriod.NewIndex) = 9
    cmbPeriod.AddItem "11 Months"
    cmbPeriod.ItemData(cmbPeriod.NewIndex) = 10
    cmbPeriod.AddItem "12 Months"
    cmbPeriod.ItemData(cmbPeriod.NewIndex) = 11
    
    cmbPeriod.ListIndex = 5 'default to 6 mths
    
    '
    'Now populate cmbWhichPeople. Used to select category of person for which to
    ' calc totals/averages
    '
    cmbWhichPeople.AddItem "All Persons"
    cmbWhichPeople.ItemData(cmbWhichPeople.NewIndex) = 0
    cmbWhichPeople.AddItem "Publishers only"
    cmbWhichPeople.ItemData(cmbWhichPeople.NewIndex) = 1
    cmbWhichPeople.AddItem "Regular Pioneers"
    cmbWhichPeople.ItemData(cmbWhichPeople.NewIndex) = 2
    cmbWhichPeople.AddItem "Auxiliary Pioneers"
    cmbWhichPeople.ItemData(cmbWhichPeople.NewIndex) = 3
    cmbWhichPeople.AddItem "Elders"
    cmbWhichPeople.ItemData(cmbWhichPeople.NewIndex) = 4
    cmbWhichPeople.AddItem "Ministerial Servants"
    cmbWhichPeople.ItemData(cmbWhichPeople.NewIndex) = 5
    cmbWhichPeople.AddItem "Baptised Publishers"
    cmbWhichPeople.ItemData(cmbWhichPeople.NewIndex) = 6
    cmbWhichPeople.AddItem "Unbaptised Publishers"
    cmbWhichPeople.ItemData(cmbWhichPeople.NewIndex) = 7
    cmbWhichPeople.AddItem "Males"
    cmbWhichPeople.ItemData(cmbWhichPeople.NewIndex) = 9
    cmbWhichPeople.AddItem "Females"
    cmbWhichPeople.ItemData(cmbWhichPeople.NewIndex) = 10
    
    cmbWhichPeople.AddItem "This person >>>"
    cmbWhichPeople.ItemData(cmbWhichPeople.NewIndex) = 8
    
    cmbWhichPeople.AddItem "This group >>>"
    cmbWhichPeople.ItemData(cmbWhichPeople.NewIndex) = 11
    
    cmbWhichPeople.ListIndex = 0
    
    '
    'Now populate cmbAttribute.
    '
    With cmbAttribute
    .AddItem "No Criteria"
    .ItemData(.NewIndex) = 0
    .AddItem "Age"
    .ItemData(.NewIndex) = 1
    .AddItem "Years Baptised"
    .ItemData(.NewIndex) = 2
    .AddItem "Average hours"
    .ItemData(.NewIndex) = 3
    .AddItem "Average RVs"
    .ItemData(.NewIndex) = 4
    .ListIndex = 0
    End With
    
    With cmbAttribute2
    .AddItem "No Criteria"
    .ItemData(.NewIndex) = 0
    .AddItem "Age"
    .ItemData(.NewIndex) = 1
    .AddItem "Years Baptised"
    .ItemData(.NewIndex) = 2
    .AddItem "Average hours"
    .ItemData(.NewIndex) = 3
    .AddItem "Average RVs"
    .ItemData(.NewIndex) = 4
    .ListIndex = 0
    End With
    
    '
    'Now populate cmbRelOp.
    '
    With cmbRelOp
    .AddItem "="
    .AddItem ">"
    .AddItem "<"
    .AddItem ">="
    .AddItem "<="
    .AddItem "<>"
    End With
    
    With cmbRelOp2
    .AddItem "="
    .AddItem ">"
    .AddItem "<"
    .AddItem ">="
    .AddItem "<="
    .AddItem "<>"
    End With
    
    'Populate cmbNames
    PopulateComboWithCurrentPublishers Me!cmbNames, _
                                       IsPublisher
                                       
    'populate bookgroups combo
    HandleListBox.PopulateListBox Me.cmbBookgroups, _
                              "SELECT GroupNo, GroupName " & _
                              "FROM tblBookGroups " & _
                              "ORDER BY 2 ASC", CMSDB, 0, "", False, 1
    
    If cmbBookgroups.ListCount > 0 Then
        cmbBookgroups.ListIndex = 0
    End If
    
    'rebuild reporting db on first run of query
    mbDatesChanged = True
    mbCreditOptionChanged = True
    
    '
    'Set up flx
    '
    With flxIndividualAverages
    .Rows = 2
    .FixedRows = 1
    .FixedCols = 0
    .Cols = 9
        
    .Row = 0
    For i = 0 To 7
        .col = i
        .CellFontBold = True
    Next i
    For i = 1 To 7
        .col = i
        .ColAlignment(i) = flexAlignRightCenter
    Next i
    For i = 1 To 7
        .col = i
        .CellAlignment = flexAlignCenterCenter
    Next i
        
    .ColAlignment(0) = flexAlignLeftCenter 'name column
    
    .TextMatrix(0, 0) = "Name"
    .TextMatrix(0, 1) = "Bks"
    .TextMatrix(0, 2) = "Bro"
    .TextMatrix(0, 3) = "Hrs"
    .TextMatrix(0, 4) = "Mgs"
    .TextMatrix(0, 5) = "RVs"
    .TextMatrix(0, 6) = "Stu"
    .TextMatrix(0, 7) = "Tra"
    .TextMatrix(0, 8) = "ID"
    
    .ColWidth(0) = 2613
    .ColWidth(1) = 740
    .ColWidth(2) = 740
    .ColWidth(3) = 740
    .ColWidth(4) = 740
    .ColWidth(5) = 740
    .ColWidth(6) = 740
    .ColWidth(7) = 740
    .ColWidth(8) = 0
    
    End With
    
    mbGridFilled = False
    mbHitRunButton = False
    
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Private Sub flxIndividualAverages_Click()
Static SortAsc As Boolean

    If mbGridFilled Then
        With flxIndividualAverages
        
        If .MouseRow = 0 Then
            If .col < 8 Then
                If SortAsc = True Then
                    SortAsc = False
                Else
                    SortAsc = True
                End If
                FillGrid .col + 1, SortAsc
            End If
        End If
        
        End With
    End If

End Sub

Private Function ValidQuery() As Boolean
On Error GoTo ErrorTrap
Dim TheControl As Control

    ValidQuery = True
    
    For Each TheControl In Me.Controls
        If TypeOf TheControl Is ComboBox Then
            If TheControl.ListIndex = -1 And _
                TheControl.Visible And _
                TheControl.Name <> "cmbNames" Then
                MsgBox "Ensure that all options are selected.", vbOKOnly + vbExclamation, AppName
                ValidQuery = False
                TheControl.SetFocus
                Exit Function
            End If
        End If
    Next
    
    txtValue = Trim$(txtValue)
    txtValue2 = Trim$(txtValue2)
    
    If txtValue = "" And txtValue.Visible Then
        MsgBox "Please enter a value.", vbOKOnly + vbExclamation, AppName
        ValidQuery = False
        txtValue.SetFocus
        Exit Function
    End If
    
    If txtValue.Visible Then
        If Not IsNumber(txtValue, True, False, False) Then
            MsgBox "Invalid entry", vbOKOnly + vbExclamation, AppName
            ValidQuery = False
            txtValue.SetFocus
            Exit Function
        End If
    End If
    
    If txtValue2 = "" And txtValue2.Visible Then
        MsgBox "Please enter a value.", vbOKOnly + vbExclamation, AppName
        ValidQuery = False
        txtValue2.SetFocus
        Exit Function
    End If
    
    If txtValue2.Visible Then
        If Not IsNumber(txtValue2, True, False, False) Then
            MsgBox "Invalid entry", vbOKOnly + vbExclamation, AppName
            ValidQuery = False
            txtValue2.SetFocus
            Exit Function
        End If
    End If
    
    If cmbWhichPeople.ItemData(cmbWhichPeople.ListIndex) = 7 And _
        (cmbAttribute.ItemData(cmbAttribute.ListIndex) = 2 Or _
         cmbAttribute2.ItemData(cmbAttribute2.ListIndex) = 2) Then
        MsgBox "Cannot process 'Years Baptised' for 'Unbaptised Publishers'!.", _
                vbOKOnly + vbExclamation, AppName
        ValidQuery = False
        cmbAttribute.SetFocus
        Exit Function
    End If
        
    ValidQuery = True

    Exit Function
ErrorTrap:
    EndProgram

End Function

Private Function MinSQLStatsTotals(StartDate_US As String, _
                             EndDate_US As String, _
                             ExtraWhereClause As String, _
                             Optional RelOp, _
                             Optional AvgValue, _
                             Optional AvgType As Long = 1, _
                             Optional RelOp2, _
                             Optional AvgValue2, _
                             Optional AvgType2 As Long = 1) As String
On Error GoTo ErrorTrap

Dim AvgValueSQL As String

    If IsMissing(RelOp) Then
        AvgValueSQL = " "
    Else
        Select Case AvgType
        Case 1
            AvgValueSQL = " HAVING Avg(NoHours) " & RelOp & " " & AvgValue & " "
        Case 2
            AvgValueSQL = " HAVING Avg(NoReturnVisits) " & RelOp & " " & AvgValue & " "
        End Select
    End If
    
    If IsMissing(RelOp2) Then
        'IF stmt above will have taken care of it
    Else
        If IsMissing(RelOp) Then
            Select Case AvgType2
            Case 1
                AvgValueSQL = " HAVING Avg(NoHours) " & RelOp2 & " " & AvgValue2 & " "
            Case 2
                AvgValueSQL = " HAVING Avg(NoReturnVisits) " & RelOp2 & " " & AvgValue2 & " "
            End Select
        Else
            Select Case AvgType2
            Case 1
                AvgValueSQL = AvgValueSQL & " AND Avg(NoHours) " & RelOp2 & " " & AvgValue2 & " "
            Case 2
                AvgValueSQL = AvgValueSQL & " AND Avg(NoReturnVisits) " & RelOp2 & " " & AvgValue2 & " "
            End Select
        End If
    End If
    
    MinSQLStatsTotals = "SELECT SUM(SumBooks) AS TotSumBooks, " & _
                       "  SUM(SumBooklets) AS TotSumBooklets, " & _
                       "  SUM(SumHours) AS TotSumHours,  " & _
                       "  SUM(SumMagazines) AS TotSumMagazines, " & _
                       "  SUM(SumReturnVisits) AS TotSumReturnVisits, " & _
                       "  SUM(SumStudies) AS TotSumStudies, " & _
                       "  SUM(SumTracts) AS TotSumTracts, " & _
                       "  COUNT(PersonID) AS CountPersons "

    MinSQLStatsTotals = MinSQLStatsTotals & _
                        "FROM " & _
                        "(SELECT SUM(NoBooks) AS SumBooks, " & _
                            "SUM(NoBooklets) AS SumBooklets, " & _
                            "SUM(NoHours) AS SumHours, " & _
                            "SUM(NoMagazines) AS SumMagazines, " & _
                            "SUM(NoReturnVisits) AS SumReturnVisits, " & _
                            "SUM(NoStudies) AS SumStudies, " & _
                            "SUM(NoTracts) AS SumTracts, " & _
                            "Count(PersonID) AS CountMonths , " & _
                            "PersonID " & _
                            "FROM tblAdvancedMinReporting " & _
                            "  INNER JOIN tblNameAddress ON " & _
                                " tblAdvancedMinReporting.PersonID = tblNameAddress.ID " & _
                            "WHERE ActualMinDate BETWEEN #" & StartDate_US & _
                            "# AND #" & EndDate_US & "# " & _
                            "AND ArtificialValue = FALSE " & _
                            ExtraWhereClause & _
                            "GROUP BY PersonID " & AvgValueSQL & ")"
    
    Exit Function
ErrorTrap:
    EndProgram
End Function

Private Function MinSQLStatsAvgs(StartDate_US As String, _
                             EndDate_US As String, _
                             ExtraWhereClause As String, _
                             Optional RelOp, _
                             Optional AvgValue, _
                             Optional AvgType As Long = 1, _
                             Optional RelOp2, _
                             Optional AvgValue2, _
                             Optional AvgType2 As Long = 1) As String
On Error GoTo ErrorTrap

Dim AvgValueSQL As String

    If IsMissing(RelOp) Then
        AvgValueSQL = " "
    Else
        Select Case AvgType
        Case 1
            AvgValueSQL = " HAVING Avg(NoHours) " & RelOp & " " & AvgValue & " "
        Case 2
            AvgValueSQL = " HAVING Avg(NoReturnVisits) " & RelOp & " " & AvgValue & " "
        End Select
    End If
    
    If IsMissing(RelOp2) Then
        'IF stmt above will have taken care of it
    Else
        If IsMissing(RelOp) Then
            Select Case AvgType2
            Case 1
                AvgValueSQL = " HAVING Avg(NoHours) " & RelOp2 & " " & AvgValue2 & " "
            Case 2
                AvgValueSQL = " HAVING Avg(NoReturnVisits) " & RelOp2 & " " & AvgValue2 & " "
            End Select
        Else
            Select Case AvgType2
            Case 1
                AvgValueSQL = AvgValueSQL & " AND Avg(NoHours) " & RelOp2 & " " & AvgValue2 & " "
            Case 2
                AvgValueSQL = AvgValueSQL & " AND Avg(NoReturnVisits) " & RelOp2 & " " & AvgValue2 & " "
            End Select
        End If
    End If
    
    MinSQLStatsAvgs = "SELECT AVG(AvgBooks) AS TotAvgBooks, " & _
                       " AVG(AvgBooklets) AS TotAvgBooklets, " & _
                       "  AVG(AvgHours)  AS TotAvgHours, " & _
                       "  AVG(AvgMagazines) AS TotAvgMagazines,  " & _
                       "  AVG(AvgReturnVisits) AS TotAvgReturnVisits, " & _
                       "  AVG(AvgStudies) AS TotAvgStudies, " & _
                       "  AVG(AvgTracts) AS TotAvgTracts, " & _
                       "Sum(SumBooks) AS TotSumBooks, " & _
                       " Sum(SumBooklets) AS TotSumBooklets, " & _
                       "  Sum(SumHours)  AS TotSumHours, " & _
                       "  Sum(SumMagazines) AS TotSumMagazines,  " & _
                       "  Sum(SumReturnVisits) AS TotSumReturnVisits, " & _
                       "  Sum(SumStudies) AS TotSumStudies, " & _
                       "  Sum(SumTracts) AS TotSumTracts, " & _
                       "  COUNT(PersonID) AS CountPersons, " & _
                       "  SUM(CountMonths) AS TotalMonths "

    MinSQLStatsAvgs = MinSQLStatsAvgs & _
                    "FROM " & _
                    "(SELECT AVG(NoBooks) AS AvgBooks, " & _
                        "AVG(NoBooklets) AS AvgBooklets, " & _
                        "AVG(NoHours) AS AvgHours, " & _
                        "AVG(NoMagazines) AS AvgMagazines, " & _
                        "AVG(NoReturnVisits) AS AvgReturnVisits, " & _
                        "AVG(NoStudies) AS AvgStudies, " & _
                        "AVG(NoTracts) AS AvgTracts, " & _
                        "SUM(NoBooks) AS SumBooks, " & _
                        "SUM(NoBooklets) AS SumBooklets, " & _
                        "SUM(NoHours) AS SumHours, " & _
                        "SUM(NoMagazines) AS SumMagazines, " & _
                        "SUM(NoReturnVisits) AS SumReturnVisits, " & _
                        "SUM(NoStudies) AS SumStudies, " & "SUM(NoTracts) AS SumTracts, " & _
                        "Count(PersonID) AS CountMonths , " & _
                        "PersonID " & _
                        "FROM tblAdvancedMinReporting " & _
                        "  INNER JOIN tblNameAddress ON " & _
                            " tblAdvancedMinReporting.PersonID = tblNameAddress.ID " & _
                        "WHERE ActualMinDate BETWEEN #" & StartDate_US & _
                        "# AND #" & EndDate_US & "# " & _
                        "AND ArtificialValue = FALSE " & _
                        ExtraWhereClause & _
                        "GROUP BY PersonID " & AvgValueSQL & ")"
        
'    MinSQLStatsAvgs = "SELECT AVG(AvgBooks) AS TotAvgBooks, " & _
'                       " AVG(AvgBooklets) AS TotAvgBooklets, " & _
'                       "  AVG(AvgHours)  AS TotAvgHours, " & _
'                       "  AVG(AvgMagazines) AS TotAvgMagazines,  " & _
'                       "  AVG(AvgReturnVisits) AS TotAvgReturnVisits, " & _
'                       "  AVG(AvgStudies) AS TotAvgStudies, " & _
'                       "  COUNT(PersonID) AS CountPersons "
'
'    MinSQLStatsAvgs = MinSQLStatsAvgs & _
'                    "FROM " & _
'                    "(SELECT AVG(NoBooks) AS AvgBooks, " & _
'                        "AVG(NoBooklets) AS AvgBooklets, " & _
'                        "AVG(NoHours) AS AvgHours, " & _
'                        "AVG(NoMagazines) AS AvgMagazines, " & _
'                        "AVG(NoReturnVisits) AS AvgReturnVisits, " & _
'                        "AVG(NoStudies) AS AvgStudies, " & _
'                        "Count(PersonID) AS CountMonths , " & _
'                        "PersonID " & _
'                        "FROM tblAdvancedMinReporting " & _
'                        "  INNER JOIN tblNameAddress ON " & _
'                            " tblAdvancedMinReporting.PersonID = tblNameAddress.ID " & _
'                        "WHERE ActualMinDate BETWEEN #" & StartDate_US & _
'                        "# AND #" & EndDate_US & "# " & _
'                        ExtraWhereClause & _
'                        "GROUP BY PersonID " & AvgValueSQL & ")"
        
    Exit Function
ErrorTrap:
    EndProgram
End Function



Private Function MinSQLNames(StartDate_US As String, _
                             EndDate_US As String, _
                             ExtraWhereClause As String, _
                             Optional RelOp, _
                             Optional AvgValue, _
                             Optional AvgType As Long = 1, _
                             Optional RelOp2, _
                             Optional AvgValue2, _
                             Optional AvgType2 As Long = 1) As String

On Error GoTo ErrorTrap

Dim AvgValueSQL As String

    If IsMissing(RelOp) Then
        AvgValueSQL = " "
    Else
        Select Case AvgType
        Case 1
            AvgValueSQL = " HAVING Avg(NoHours) " & RelOp & " " & AvgValue & " "
        Case 2
            AvgValueSQL = " HAVING Avg(NoReturnVisits) " & RelOp & " " & AvgValue & " "
        End Select
    End If
    
    If IsMissing(RelOp2) Then
        'IF stmt above will have taken care of it
    Else
        If IsMissing(RelOp) Then
            Select Case AvgType2
            Case 1
                AvgValueSQL = " HAVING Avg(NoHours) " & RelOp2 & " " & AvgValue2 & " "
            Case 2
                AvgValueSQL = " HAVING Avg(NoReturnVisits) " & RelOp2 & " " & AvgValue2 & " "
            End Select
        Else
            Select Case AvgType2
            Case 1
                AvgValueSQL = AvgValueSQL & " AND Avg(NoHours) " & RelOp2 & " " & AvgValue2 & " "
            Case 2
                AvgValueSQL = AvgValueSQL & " AND Avg(NoReturnVisits) " & RelOp2 & " " & AvgValue2 & " "
            End Select
        End If
    End If

    MinSQLNames = "SELECT " & _
                    "LastName & ', ' & FirstName & ' ' & MiddleName AS TheFullName, " & _
                    "AVG(NoBooks) AS AvgBooks, " & _
                    "AVG(NoBooklets) AS AvgBooklets, " & _
                    "AVG(NoHours) AS AvgHours, " & _
                    "AVG(NoMagazines) AS AvgMagazines, " & _
                    "AVG(NoReturnVisits) AS AvgReturnVisits, " & _
                    "AVG(NoStudies) AS AvgStudies, " & _
                    "AVG(NoTracts) AS AvgTracts, " & _
                    "PersonID, " & _
                    "BookGroupID " & _
                    "FROM tblAdvancedMinReporting " & _
                    "  INNER JOIN tblNameAddress ON " & _
                        " tblAdvancedMinReporting.PersonID = tblNameAddress.ID " & _
                    "WHERE ActualMinDate BETWEEN #" & StartDate_US & _
                    "# AND #" & EndDate_US & "# " & _
                    ExtraWhereClause & _
                    "GROUP BY PersonID, LastName & ', ' & FirstName & ' ' & MiddleName, " & _
                    "BookGroupID " & _
                    AvgValueSQL

    Exit Function
ErrorTrap:
    EndProgram
End Function
Private Sub cmbNames_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap

    If KeyCode = 46 Then
        cmbNames.ListIndex = -1
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Public Sub ClearResults()
On Error GoTo ErrorTrap
Dim i As Long
        txtTotBro = ""
        txtTotBooks = ""
        txtTotHrs = ""
        txtTotMags = ""
        txtTotRVs = ""
        txtTotStu = ""
        txtTotTra = ""
        txtAvgBooks = ""
        txtAvgBro = ""
        txtAvgHrs = ""
        txtAvgMags = ""
        txtAvgRVs = ""
        txtAvgStu = ""
        txtAvgTra = ""
        flxIndividualAverages.Rows = 2
        
        With flxIndividualAverages
        If .Cols > 2 Then
            For i = 0 To 7
                flxIndividualAverages.TextMatrix(1, i) = ""
            Next i
        End If
        End With
        txtCount = ""
        cmdPrint.Enabled = False
        cmdExport.Enabled = False
        
        mbGridFilled = False
        
        mbHitRunButton = False
    
    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub FillGrid(Optional OrderByCol = 1, Optional SortAscending = True)
On Error GoTo ErrorTrap
Dim i As Long, j As Long, SortStr As String

    SortStr = " ORDER BY " & OrderByCol & IIf(SortAscending = True, " ASC ", " DESC ")
    msReportSQL = msGridSQL & SortStr
    
    Select Case OrderByCol
    Case 1
        msOrderBy = " PersonName " & IIf(SortAscending = True, " ASC ", " DESC ")
    Case 2
        msOrderBy = " AvgBooks " & IIf(SortAscending = True, " ASC ", " DESC ")
    Case 3
        msOrderBy = " AvgBooklets " & IIf(SortAscending = True, " ASC ", " DESC ")
    Case 4
        msOrderBy = " AvgHours " & IIf(SortAscending = True, " ASC ", " DESC ")
    Case 5
        msOrderBy = " AvgMagazines " & IIf(SortAscending = True, " ASC ", " DESC ")
    Case 6
        msOrderBy = " AvgReturnVisits " & IIf(SortAscending = True, " ASC ", " DESC ")
    Case 7
        msOrderBy = " AvgStudies " & IIf(SortAscending = True, " ASC ", " DESC ")
    Case 8
        msOrderBy = " AvgTracts " & IIf(SortAscending = True, " ASC ", " DESC ")
    End Select
    
    Set rstGridSQL = CMSDB.OpenRecordset(msReportSQL, dbOpenSnapshot)
    
    mbRecSetActive = True
    
    If Not rstGridSQL.BOF Then
        rstGridSQL.MoveLast
        rstGridSQL.MoveFirst
        mlNoRows = rstGridSQL.RecordCount
    Else
        mlNoRows = 0
        Exit Sub
    End If
    
    With flxIndividualAverages
    
    .Rows = rstGridSQL.RecordCount + 1
    
    j = 1
    
    Do Until rstGridSQL.BOF Or rstGridSQL.EOF
        For i = 0 To 7
            If i > 0 Then
                .TextMatrix(j, i) = Format$(Round(rstGridSQL.Fields(i), 2), "0.00")
            Else
                .TextMatrix(j, i) = rstGridSQL.Fields(i)
            End If
        Next i
        j = j + 1
        rstGridSQL.MoveNext
    Loop
    
    ShadeOddGridRowBands flxIndividualAverages, RGB(235, 235, 235)
    
    'indicate which column is sorted
    'indicate which column is sorted
    Select Case OrderByCol
    Case 1
        .TextMatrix(0, 0) = "Name" & " *"
        .TextMatrix(0, 1) = "Bks"
        .TextMatrix(0, 2) = "Bro"
        .TextMatrix(0, 3) = "Hrs"
        .TextMatrix(0, 4) = "Mgs"
        .TextMatrix(0, 5) = "RVs"
        .TextMatrix(0, 6) = "Stu"
        .TextMatrix(0, 7) = "Tra"
    Case 2
        .TextMatrix(0, 0) = "Name"
        .TextMatrix(0, 1) = "Bks" & " *"
        .TextMatrix(0, 2) = "Bro"
        .TextMatrix(0, 3) = "Hrs"
        .TextMatrix(0, 4) = "Mgs"
        .TextMatrix(0, 5) = "RVs"
        .TextMatrix(0, 6) = "Stu"
        .TextMatrix(0, 7) = "Tra"
    Case 3
        .TextMatrix(0, 0) = "Name"
        .TextMatrix(0, 1) = "Bks"
        .TextMatrix(0, 2) = "Bro" & " *"
        .TextMatrix(0, 3) = "Hrs"
        .TextMatrix(0, 4) = "Mgs"
        .TextMatrix(0, 5) = "RVs"
        .TextMatrix(0, 6) = "Stu"
        .TextMatrix(0, 7) = "Tra"
    Case 4
        .TextMatrix(0, 0) = "Name"
        .TextMatrix(0, 1) = "Bks"
        .TextMatrix(0, 2) = "Bro"
        .TextMatrix(0, 3) = "Hrs" & " *"
        .TextMatrix(0, 4) = "Mgs"
        .TextMatrix(0, 5) = "RVs"
        .TextMatrix(0, 6) = "Stu"
        .TextMatrix(0, 7) = "Tra"
    Case 5
        .TextMatrix(0, 0) = "Name"
        .TextMatrix(0, 1) = "Bks"
        .TextMatrix(0, 2) = "Bro"
        .TextMatrix(0, 3) = "Hrs"
        .TextMatrix(0, 4) = "Mgs" & " *"
        .TextMatrix(0, 5) = "RVs"
        .TextMatrix(0, 6) = "Stu"
        .TextMatrix(0, 7) = "Tra"
    Case 6
        .TextMatrix(0, 0) = "Name"
        .TextMatrix(0, 1) = "Bks"
        .TextMatrix(0, 2) = "Bro"
        .TextMatrix(0, 3) = "Hrs"
        .TextMatrix(0, 4) = "Mgs"
        .TextMatrix(0, 5) = "RVs" & " *"
        .TextMatrix(0, 6) = "Stu"
        .TextMatrix(0, 7) = "Tra"
    Case 7
        .TextMatrix(0, 0) = "Name"
        .TextMatrix(0, 1) = "Bks"
        .TextMatrix(0, 2) = "Bro"
        .TextMatrix(0, 3) = "Hrs"
        .TextMatrix(0, 4) = "Mgs"
        .TextMatrix(0, 5) = "RVs"
        .TextMatrix(0, 6) = "Stu" & " *"
        .TextMatrix(0, 7) = "Tra"
    Case 8
        .TextMatrix(0, 0) = "Name"
        .TextMatrix(0, 1) = "Bks"
        .TextMatrix(0, 2) = "Bro"
        .TextMatrix(0, 3) = "Hrs"
        .TextMatrix(0, 4) = "Mgs"
        .TextMatrix(0, 5) = "RVs"
        .TextMatrix(0, 6) = "Stu"
        .TextMatrix(0, 7) = "Tra" & " *"
    End Select
    
    End With
    mbGridFilled = True
    rstGridSQL.Close
    
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

Private Sub txtValue_Change()
    ClearResults
End Sub
Private Sub txtValue2_Change()
    ClearResults
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)

    'Must be numeric. Allow Backspace (8) and full stop (46)
    'Delete and arrow keys seem to be allowed by default.
    
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
            KeyAscii = 0
        End If

End Sub
Private Sub txtValue2_KeyPress(KeyAscii As Integer)

    'Must be numeric. Allow Backspace (8) and full stop (46)
    'Delete and arrow keys seem to be allowed by default.
    
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
            KeyAscii = 0
        End If

End Sub

Public Property Get OrderBy() As String
    OrderBy = msOrderBy
End Property

Public Property Get Criteria() As String
    Criteria = msCriteria
End Property

Public Property Let ReportUpdated(ByVal vNewValue As Boolean)
    mbReportsUpdated = vNewValue
End Property
Public Property Get Form2RebuiltTable() As Boolean
    Form2RebuiltTable = mbForm2RebuiltTable
End Property

Public Property Let Form2RebuiltTable(ByVal vNewValue As Boolean)
    mbForm2RebuiltTable = vNewValue
End Property

