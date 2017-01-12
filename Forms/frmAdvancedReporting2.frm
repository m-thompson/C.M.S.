VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAdvancedReporting2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.M.S. Advanced Reporting - Screen 2"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15180
   Icon            =   "frmAdvancedReporting2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   15180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1140
      Top             =   8850
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Results"
      ForeColor       =   &H00FF0000&
      Height          =   3690
      Left            =   90
      TabIndex        =   12
      Top             =   5295
      Width           =   14985
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export..."
         Height          =   345
         Left            =   12990
         TabIndex        =   16
         Top             =   3180
         Width           =   930
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   345
         Left            =   13965
         TabIndex        =   15
         Top             =   3180
         Width           =   930
      End
      Begin MSFlexGridLib.MSFlexGrid flxQueryResults 
         Height          =   2865
         Left            =   135
         TabIndex        =   13
         Top             =   285
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   5054
         _Version        =   393216
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   1
      End
      Begin VB.Label lblRecords 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   150
         TabIndex        =   14
         Top             =   3225
         Width           =   5715
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Query Builder"
      ForeColor       =   &H00FF0000&
      Height          =   5010
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   14985
      Begin VB.CommandButton cmdClear 
         Cancel          =   -1  'True
         Caption         =   "&Clear..."
         Height          =   450
         Left            =   13980
         TabIndex        =   7
         ToolTipText     =   "[ESC]"
         Top             =   4410
         Width           =   885
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "&Run"
         Height          =   450
         Left            =   11085
         TabIndex        =   4
         ToolTipText     =   "F5"
         Top             =   4410
         Width           =   885
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save..."
         Height          =   450
         Left            =   13020
         TabIndex        =   6
         Top             =   4410
         Width           =   885
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&Open..."
         Height          =   450
         Left            =   12045
         TabIndex        =   5
         Top             =   4410
         Width           =   885
      End
      Begin VB.ListBox lstColumns 
         Height          =   4350
         Left            =   3330
         TabIndex        =   3
         Top             =   540
         Width           =   3000
      End
      Begin VB.TextBox txtTableDesc 
         BackColor       =   &H80000000&
         Height          =   1290
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3600
         Width           =   3180
      End
      Begin VB.TextBox txtQuery 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3795
         Left            =   6360
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   540
         Width           =   8490
      End
      Begin VB.ListBox lstTables 
         Height          =   2985
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   540
         Width           =   3180
      End
      Begin VB.Label lblQueryName 
         Caption         =   "Query"
         Height          =   240
         Left            =   6375
         TabIndex        =   11
         Top             =   330
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Columns"
         Height          =   225
         Left            =   3585
         TabIndex        =   10
         Top             =   330
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Tables"
         Height          =   225
         Left            =   150
         TabIndex        =   8
         Top             =   330
         Width           =   1695
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "Actions"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Table to query"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select all"
      End
   End
End
Attribute VB_Name = "frmAdvancedReporting2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbQueryChanged As Boolean, msQueryName As String, msQueryDesc As String
Dim WithEvents frmOpenQry As frmOpenUserQuery
Attribute frmOpenQry.VB_VarHelpID = -1

Private Sub frmOpenQry_PassQuery(TheQueryName As String, TheQueryDesc As String, TheQueryString As String)

On Error GoTo ErrorTrap
    
    msQueryDesc = TheQueryDesc
    msQueryName = TheQueryName
    txtQuery = TheQueryString

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Unload frmOpenQry
    Set frmOpenQry = Nothing

    BringForwardMainMenuWhenItsTheLastFormOpen

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdClear_Click()
On Error GoTo ErrorTrap

    If Len(txtQuery) > 0 Then
        If MsgBox("Clear query?", vbYesNo + vbQuestion, AppName) = vbYes Then
            InitQuery
        End If
    End If

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub
Private Sub InitQuery()

On Error Resume Next
    txtQuery.SetFocus
    
On Error GoTo ErrorTrap

    txtQuery = "SELECT "
    txtQuery.SelStart = Len(txtQuery)
    
    msQueryName = ""
    msQueryDesc = ""
    lblRecords = ""
    
    cmdExport.Enabled = False

    mbQueryChanged = False
    Frame2.Caption = "Results"

    Exit Sub
ErrorTrap:
    Call EndProgram


End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()

On Error GoTo ErrorTrap


    ExportFlexGridToCSV gsDocsDirectory, _
                        "Query", _
                        "csv", _
                        "C.M.S. Query Results", _
                        flxQueryResults, _
                        CommonDialog1


    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub cmdOpen_Click()

On Error GoTo ErrorTrap

    frmOpenQry.Screen2 = True
    frmOpenQry.Show vbModal, Me
    mbQueryChanged = False
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdRun_Click()
On Error GoTo ErrorTrap

    txtQuery = Trim(txtQuery)
    
    If Len(txtQuery) < 10 Then
        MsgBox "Enter a valid query", vbOKOnly + vbExclamation, AppName
        TextFieldGotFocus txtQuery, True
        Exit Sub
    End If
    
    ClearGrid
    
'    If LCase(Left(txtQuery, 6)) <> "select" And _
'        LCase(Left(txtQuery, 6)) <> "update" And _
'        LCase(Left(txtQuery, 6)) <> "delete" And _
'        LCase(Left(txtQuery, 6)) <> "insert" Then
'        MsgBox "SQL not valid - must begin with SELECT/UPDATE/INSERT/DELETE", vbOKOnly + vbExclamation, AppName
'        TextFieldGotFocus txtQuery, True
'        Exit Sub
'    End If
    
    If LCase(Left(txtQuery, 6)) = "update" Or _
        LCase(Left(txtQuery, 6)) = "delete" Or _
        LCase(Left(txtQuery, 6)) = "insert" Then
        
        If MsgBox("This SQL will modify the database contents. Continue?", vbYesNo + vbQuestion + vbDefaultButton2, AppName) = vbNo Then
            MsgBox "Operation cancelled", vbOKOnly + vbInformation, AppName
            TextFieldGotFocus txtQuery, True
            Exit Sub
        Else
            DoUpdateSQL
        End If
    Else
        FillGrid
    End If
    
    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub cmdSave_Click()

On Error GoTo ErrorTrap

    If txtQuery = "" Then
        MsgBox "No query entered", vbExclamation + vbOKOnly, AppName
        txtQuery.SetFocus
        Exit Sub
    End If
    
    With frmSaveQuery
    .QueryDesc = msQueryDesc
    .QueryName = msQueryName
    .QueryString = txtQuery
    .Screen2 = True
    .Show vbModal, Me
    End With
    
    If flxQueryResults.Rows > 0 Then
        Frame2 = "Results - " & Replace(msQueryName, "&", "&&")
    Else
        Frame2 = "Results"
    End If

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorTrap

    If KeyCode = 116 Then
        cmdRun_Click
    End If

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub


Private Sub Form_Load()
On Error GoTo ErrorTrap

    Me.Left = frmAdvancedReporting.Left + 567
    Me.Top = frmAdvancedReporting.Top - 567
   
    FillTableList
    
    InitQuery
    
    flxQueryResults.Rows = 0
    
    Set frmOpenQry = New frmOpenUserQuery

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub
Private Sub FillTableList()
Dim tdf As TableDef
On Error GoTo ErrorTrap

    lstTables.Clear
    
    For Each tdf In CMSDB.TableDefs
        Select Case tdf.Name
        Case "tblAdvancedMinReportingPrint", "tblCMSBackups", _
             "tblExportDetails", "tblIDWeightings", "tblLastExportDate", _
             "tblObjectSecurity_OLD", "tblPrintBookGroups", "tblPrintCongMinByGroup", _
             "tblPrintPublicMtgSchedule", "tblPrintSPAMRota", "tblPrintVisitingTalkSchedule", _
             "tblPublicMtgWtg", "tblServiceMtgSchedulePrint", "tblStoredSPAMRotas", _
             "tblStoredTMSSchedules", "tblTMSAssignmentSlips", "tblTMSPrintSchedule", _
             "tblTMSPrintStudentDetails", "tblTMSPrintTypes", "tblTMSPrintWorkSheet", _
             "tblTMSSkillLookup", "tblTMSSkillRatings", "tblTMSWeightings", _
             "tblUserQueries", "tblAdvancedMinReporting", "tblSecurity"
             
        Case Else
            If InStr(1, tdf.Name, ":") = 0 Then
                If Left(tdf.Name, 3) = "tbl" And Left(tdf.Name, 7) <> "tblTemp" Then
                    lstTables.AddItem tdf.Name
                End If
            End If
        End Select
    Next
    
    lstTables.ListIndex = 0

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub
Private Sub FillColumnList()
Dim fld As Field
On Error GoTo ErrorTrap

    lstColumns.Clear
    
    If lstTables.ListIndex = -1 Then
        Exit Sub
    End If
    
    For Each fld In CMSDB.TableDefs(lstTables.text).Fields
        lstColumns.AddItem fld.Name
    Next
    
    lstColumns.ListIndex = 0

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub


Private Sub lstColumns_Click()
On Error GoTo ErrorTrap
Dim fld As Field

    If lstColumns.ListIndex = -1 Then
        Exit Sub
    End If
    
    Set fld = CMSDB.TableDefs(lstTables.text).Fields(lstColumns.text)

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub

Private Sub lstColumns_DblClick()
On Error GoTo ErrorTrap
Dim txt As String, Pos As Long

    If lstColumns.ListIndex = -1 Then
        Exit Sub
    End If
    
    If InStr(1, lstColumns.text, " ") = 0 Then
        txt = lstColumns.text
    Else
        txt = "[" & lstColumns.text & "]"
    End If

    Pos = txtQuery.SelStart

    txtQuery = InsertSubstr(txtQuery, _
                            txt, _
                            txtQuery.SelStart)
                            
    txtQuery.SelStart = Pos + Len(txt)
    txtQuery.SetFocus
                            

    Exit Sub
    
ErrorTrap:
    Call EndProgram

End Sub

Private Sub lstTables_Click()

On Error GoTo ErrorTrap

    If lstTables.ListIndex = -1 Then
        txtTableDesc.text = ""
        lstColumns.Clear
        Exit Sub
    End If

    On Error Resume Next
    txtTableDesc = CMSDB.TableDefs(lstTables.text).Properties("Description")
    If Err.number <> 0 Then
        txtTableDesc = ""
    End If
    
    On Error GoTo ErrorTrap
    FillColumnList

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub
Private Sub ClearGrid()

On Error GoTo ErrorTrap

    flxQueryResults.Rows = 0
    lblRecords = ""
    cmdExport.Enabled = False
    Frame2 = "Results"
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub lstTables_DblClick()
On Error GoTo ErrorTrap
Dim txt As String, Pos As Long

    If lstTables.ListIndex = -1 Then
        Exit Sub
    End If
    
    
    txt = AddSquareBracketsIfNecessary(lstTables.text)

    Pos = txtQuery.SelStart
    
    txtQuery = InsertSubstr(txtQuery, _
                            txt, _
                            txtQuery.SelStart)
                            
    txtQuery.SelStart = Pos + Len(txt)
    txtQuery.SetFocus

    Exit Sub
    
ErrorTrap:
    Call EndProgram

End Sub
Public Function AddSquareBracketsIfNecessary(TableName As String) As String
On Error GoTo ErrorTrap
Dim txt As String, Pos As Long

    If InStr(1, TableName, " ") = 0 Then
        AddSquareBracketsIfNecessary = TableName
    Else
        AddSquareBracketsIfNecessary = "[" & TableName & "]"
    End If

    Exit Function
    
ErrorTrap:
    Call EndProgram

End Function

Private Sub FillGrid()
On Error GoTo ErrorTrap
Dim i As Long, j As Long, rstGridSQL As Recordset, fld As Field
Dim lMaxRows As Long, lRows As Long
        
    On Error Resume Next
    
    Set rstGridSQL = CMSDB.OpenRecordset(txtQuery, dbOpenSnapshot)
    
    If Err.number <> 0 Then
        MsgBox "The following error occurred - check, amend and re-run the query." & _
                    vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly
                    
        TextFieldGotFocus txtQuery, True
        Exit Sub
    End If
    
    On Error GoTo ErrorTrap
    
    ClearGrid
    
    If Not rstGridSQL.BOF Then
        rstGridSQL.MoveLast
        rstGridSQL.MoveFirst
    Else
        MsgBox "Query returned no data", vbOKOnly + vbInformation, AppName
        Exit Sub
    End If
        
    With flxQueryResults
    
    lRows = rstGridSQL.RecordCount
    
    lblRecords = lRows & " record" & IIf(lRows = 1, "", "s")
    
    lMaxRows = GlobalParms.GetValue("AdvancedReportMaxRows", "NumVal")
    
    If lRows > lMaxRows Then
        lRows = lMaxRows
        lblRecords = lblRecords & " (capped at " & lMaxRows & ")"
    End If
    
    .Rows = lRows + 1
    .Cols = rstGridSQL.Fields.Count
    
    .FixedRows = 1
    .FixedCols = 0
    
    .Row = 0
    For i = 0 To rstGridSQL.Fields.Count - 1
        .col = i
        .CellFontBold = True
    Next i
    
    For i = 0 To rstGridSQL.Fields.Count - 1
        .TextMatrix(0, i) = rstGridSQL.Fields(i).Name
    Next i
    
    For i = 0 To rstGridSQL.Fields.Count - 1
        Select Case rstGridSQL.Fields(i).Type
        Case dbBigInt, dbByte, dbCurrency, dbDecimal, dbDouble, _
                dbFloat, dbInteger, dbLong, dbNumeric, dbSingle
            
            .ColAlignment(i) = flexAlignRightCenter
        Case Else
            .ColAlignment(i) = flexAlignLeftCenter
        End Select
    Next i
    
    .Row = 0
    
    For i = 0 To rstGridSQL.Fields.Count - 1
        .col = i
        .CellAlignment = flexAlignLeftCenter
    Next i
                
    j = 1
    
    Do Until rstGridSQL.BOF Or rstGridSQL.EOF Or j > lMaxRows
        For i = 0 To rstGridSQL.Fields.Count - 1
            .TextMatrix(j, i) = HandleNull(rstGridSQL.Fields(i), "")
        Next i
        j = j + 1
        rstGridSQL.MoveNext
    Loop

    'shade all odd rows
    For j = 1 To .Rows - 1 Step 2
        .Row = j
        For i = 0 To .Cols - 1
            .col = i
            .CellBackColor = RGB(235, 235, 235) ' light grey
        Next i
    Next j
    
    End With
    
    SetFlexGridColumnWidths Me, flxQueryResults, 1.6
    
    rstGridSQL.Close
    Set rstGridSQL = Nothing
    
    cmdExport.Enabled = True
    
    If msQueryName <> "" Then
        Frame2.Caption = "Results - " & Replace(msQueryName, "&", "&&") & IIf(mbQueryChanged, " (Modified)", "")
    Else
        Frame2.Caption = "Results"
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DoUpdateSQL()
On Error GoTo ErrorTrap
    
    ClearGrid
    
    On Error Resume Next
    
    CMSDB.Execute (txtQuery)
    
    If Err.number <> 0 Then
        MsgBox "The following error occurred - check, amend and re-run the SQL." & _
                    vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly
                    
        TextFieldGotFocus txtQuery, True
        Exit Sub
    Else
        MsgBox "SQL executed", vbInformation + vbOKOnly, AppName
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Private Sub lsttables_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorTrap

Dim TheRow As Long

    If Button = vbRightButton Then
        
        'an item must be selected before the right-click will work... How odd....
        If lstTables.ListCount > 0 Then
            lstTables.ListIndex = lstTables.TopIndex
        Else
            Exit Sub
        End If
    
        If lstTables.SelCount > 0 Then
            
            mnuAdd.Enabled = True
            
            'select appropriate row
            '
            'Use current Y position to work out which row of the listbox has been clicked
            '
            TheRow = (Ceiling(CDbl(Y) / _
                        (GetListItemHeightInTwips(lstTables)))) + _
                                                     lstTables.TopIndex - 1
            
            If TheRow > lstTables.ListCount - 1 Then
                TheRow = lstTables.ListCount - 1
            End If
            
            lstTables.ListIndex = TheRow
            
            If lstTables.SelCount > 1 Then
                mnuAdd.Caption = "Add tables to query"
                mnuSelectAll.Enabled = False
                mnuSelectAll.Caption = "Select all"
            Else
                mnuAdd.Caption = "Add " & CreateJoinSQL & " to query"
                mnuSelectAll.Caption = "Select all from " & CreateJoinSQL
                mnuSelectAll.Enabled = True
            End If
            
            Me.PopupMenu mnuActions
                                                                  
        Else
            mnuAdd.Enabled = False
        End If
        
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub




Private Sub mnuAdd_Click()
On Error GoTo ErrorTrap
Dim txt As String, Pos As Long

    Pos = txtQuery.SelStart
    
    txt = CreateJoinSQL
    
    txtQuery = InsertSubstr(txtQuery, _
                            txt, _
                            txtQuery.SelStart)
                            
    txtQuery.SelStart = Pos + Len(txt)
    txtQuery.SetFocus

    Exit Sub
    
ErrorTrap:
    Call EndProgram
End Sub

Private Sub mnuSelectAll_Click()

On Error GoTo ErrorTrap

    'use CreateJoinSQL to get table name of select item
    txtQuery.text = "SELECT * " & vbCrLf & "FROM " & CreateJoinSQL

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub txtQuery_Change()

On Error GoTo ErrorTrap

'    ClearGrid
    mbQueryChanged = True

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub
Private Function CreateJoinSQL() As String
Dim str As String, lNoTables As Long, arrTables() As String
Dim sOpeningBrackets As String
Dim i As Long, j As Long
On Error GoTo ErrorTrap

    lNoTables = lstTables.SelCount
    
    If lNoTables = 0 Then Exit Function
    
    If lNoTables > 26 Then
        MsgBox "Maximum of 26 tables allowed", vbOKOnly + vbExclamation, AppName
        Exit Function
    End If
    
    
    With lstTables
    
    'create array to store selected table names
    
    ReDim arrTables(lNoTables - 1)
    
    'get tablenames into array
    For i = 0 To .ListCount - 1
        If .Selected(i) Then
            arrTables(j) = AddSquareBracketsIfNecessary(.List(i))
            j = j + 1
        End If
    Next i
    
    'if only one table selected, return it
    If lNoTables = 1 Then
        CreateJoinSQL = arrTables(0)
        Exit Function
    End If
    
    'if only 2 tables selected, return the simple join
    If lNoTables = 2 Then
        CreateJoinSQL = arrTables(0) & " a " & vbCrLf & "INNER JOIN " & arrTables(1) & " b ON "
        Exit Function
    End If
        
    'if > 3 tables selected, create the necessary number of opening brackets
    For i = 3 To lNoTables
        sOpeningBrackets = sOpeningBrackets & "("
    Next i
        
    'first table
    str = arrTables(0) & " a " & vbCrLf
        
    'now join each additional table in array, adding a letter (a-z) for table synonym
    For i = 1 To lNoTables - 1
    
        str = str & " INNER JOIN " & arrTables(i) & " " & LCase$(Alphabet(i + 1)) & " ON "
              
        If i < (lNoTables - 1) Then
            str = str & ") "
        End If
        
        str = str & vbCrLf
        
    Next i
        
    str = sOpeningBrackets & str
    
    End With

    CreateJoinSQL = str

    Exit Function
ErrorTrap:
    Call EndProgram
End Function

Public Property Get QueryName() As String
    QueryName = msQueryName
End Property

Public Property Let QueryName(ByVal vNewValue As String)
    msQueryName = vNewValue
End Property
Public Property Get QueryDesc() As String
    QueryName = msQueryDesc
End Property

Public Property Let QueryDesc(ByVal vNewValue As String)
    msQueryDesc = vNewValue
End Property

Public Property Get QueryChanged() As Boolean
    QueryChanged = mbQueryChanged
End Property

Public Property Let QueryChanged(ByVal vNewValue As Boolean)
    mbQueryChanged = vNewValue
End Property
