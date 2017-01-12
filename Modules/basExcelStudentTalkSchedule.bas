Attribute VB_Name = "basExcelStudentTalkSchedule"
Option Explicit

'Dim oExcelApp As Excel.Application
'Dim oExcelDoc As Excel.Workbook
'Dim oExcelResultSheet As Excel.Worksheet
'Dim oExcelNamesSheet As Excel.Worksheet
'Dim oExcelWkSht As Excel.Worksheet

Dim oExcelApp As Object
Dim oExcelDoc As Object
Dim oExcelResultSheet As Object
Dim oExcelNamesSheet As Object
Dim oExcelWkSht As Object

Dim mbExcelWasOpen As Boolean



Public Function OpenExcel() As Boolean

On Error Resume Next

    Set oExcelApp = GetObject(, "Excel.Application")
    mbExcelWasOpen = True
    If Err.number <> 0 Then
      mbExcelWasOpen = False
      Err.Clear
      Set oExcelApp = CreateObject("Excel.Application")
    End If
    
    If Err.number <> 0 Then
        ShowMessage "Could not open Excel", 1500, frmTMSPrinting
        Set oExcelApp = Nothing
        OpenExcel = False
        Exit Function
    Else
        OpenExcel = True
    End If
    
End Function

Public Sub GenerateExcelStudentTalkSchedule(StartDate As Date, EndDate As Date)
On Error GoTo ErrorTrap
Dim i As Long, j As Long, rs As Recordset, sSQL As String, sDate1 As String, sDate2 As String
Dim bForgetAboutExcel As Boolean, TempDate As Date
Dim sDateDesc As String, str As String, bOK As Boolean, oWkSht As Object, dte As Date
Dim str1 As String, str2 As String, CurrDate As Date, PrevDate As Date

    If Not gFSO.FolderExists(gsDocsDirectory) Then
        ShowMessage "No valid folder for documents", 1500, frmTMSPrinting
        Exit Sub
    End If
    
    If Not BuildPrintTable(StartDate, EndDate) Then
        ShowMessage "Could not generate data for extract", 1500, frmTMSPrinting
        Exit Sub
    End If
    
    
    Set rs = CMSDB.OpenRecordset("tblTMSPrintSchedule", dbOpenDynaset)
        
    If rs.BOF Or rs.EOF Then
        ShowMessage "No schedule exists!", 1100, frmTMSPrinting
        Exit Sub
    End If
    
    If Not OpenExcel Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    oExcelApp.Visible = False
    
    bForgetAboutExcel = True 'init
    
    'set up the document
    
    Set oExcelDoc = oExcelApp.Workbooks.Add
    Set oExcelResultSheet = oExcelDoc.Worksheets.Add
    oExcelResultSheet.Name = "Student Talks Schedule"
    
    'remove the 3 default sheets
    On Error Resume Next
    oExcelApp.DisplayAlerts = False
    oExcelDoc.Worksheets("Sheet1").Delete
    oExcelDoc.Worksheets("Sheet2").Delete
    oExcelDoc.Worksheets("Sheet3").Delete
    oExcelApp.DisplayAlerts = True
    On Error GoTo ErrorTrap
         
     'put in headings on new worksheet
     oExcelResultSheet.Range("A1").value = "Date"
     oExcelResultSheet.Range("B1").value = "Assignment"
     oExcelResultSheet.Range("C1").value = "Theme"
     oExcelResultSheet.Range("D1").value = "Student 1"
     oExcelResultSheet.Range("E1").value = "Assistant 1"
     oExcelResultSheet.Range("F1").value = "Counsel 1"
     oExcelResultSheet.Range("G1").value = "Student 2"
     oExcelResultSheet.Range("H1").value = "Assistant 2"
     oExcelResultSheet.Range("I1").value = "Counsel 2"
     oExcelResultSheet.Range("J1").value = "Student 3"
     oExcelResultSheet.Range("K1").value = "Assistant 3"
     oExcelResultSheet.Range("L1").value = "Counsel 3"
    
    
     'freeze the top row
     oExcelResultSheet.Range("2:2").Select
     oExcelApp.ActiveWindow.FreezePanes = True
     
    i = 2
    
       
    With rs
    
    CurrDate = !AssignmentDate
    PrevDate = !AssignmentDate
    
    Do Until .BOF Or .EOF
        
        
        bForgetAboutExcel = False
        
             
        oExcelResultSheet.Range("A" & i).value = !AssignmentDateStr
        oExcelResultSheet.Range("B" & i).value = !TalkType
        oExcelResultSheet.Range("C" & i).value = !Theme
        oExcelResultSheet.Range("D" & i).value = !Student1
        oExcelResultSheet.Range("E" & i).value = !Assistant1
        oExcelResultSheet.Range("F" & i).value = !SQ1
        oExcelResultSheet.Range("G" & i).value = !Student2
        oExcelResultSheet.Range("H" & i).value = !Assistant2
        oExcelResultSheet.Range("I" & i).value = !SQ2
        oExcelResultSheet.Range("J" & i).value = !Student3
        oExcelResultSheet.Range("K" & i).value = !Assistant3
        oExcelResultSheet.Range("L" & i).value = !SQ3
        
        PrevDate = !AssignmentDate
        
        .MoveNext
        
        If Not .EOF Then
            CurrDate = !AssignmentDate
            
            If PrevDate <> CurrDate Then
                'put border between dates
                With oExcelResultSheet.Range("A" & i & ":L" & i).Borders(9)              'xlEdgeBottom
                  .LineStyle = 1 'xlContinuous
                  .Weight = 2  'xlthin
                End With
            End If
        End If
        
        i = i + 1
             
    Loop
             
    End With
    
      'bold headings
      oExcelResultSheet.Range("A1:L1").Font.Bold = True
      
      'auto-size column widths
      oExcelResultSheet.Columns("A:L").AutoFit
      
      'put border around
      oExcelResultSheet.Range("A1", "L" & i).BorderAround , -4138   'xlMedium
      
      'put border under top row
      With oExcelResultSheet.Range("A1:L1").Borders(9)    'xlEdgeBottom
        .LineStyle = 1 'xlContinuous
        .Weight = 2  'xlthin
      End With
        
      'put border between Schools
      With oExcelResultSheet.Range("C1", "C" & i).Borders(10)    'xlEdgeRight
        .LineStyle = 1 'xlContinuous
        .Weight = 2  'xlthin
      End With
      'put border between Schools
      With oExcelResultSheet.Range("F1", "F" & i).Borders(10)    'xlEdgeRight
        .LineStyle = 1 'xlContinuous
        .Weight = 2  'xlthin
      End With
      'put border between Schools
      With oExcelResultSheet.Range("I1", "I" & i).Borders(10)    'xlEdgeRight
        .LineStyle = 1 'xlContinuous
        .Weight = 2  'xlthin
      End With
        
      'Shade top row
      oExcelResultSheet.Range("A1:L1").Interior.ColorIndex = 15 'grey
    
                                    
    If Not bForgetAboutExcel Then

        
        'save the doc - first get unique file name
        bOK = False
        Do Until bOK
            str = gsDocsDirectory & "\Student Talks Schedule - " & _
                        MakeStringValidForFileName(CStr(StartDate), "-") & _
                        " to " & _
                        MakeStringValidForFileName(CStr(EndDate), "-") & _
                        " (" & _
                            Replace(Replace(Now, ":", "-"), "/", "-") & ")"
            bOK = Not gFSO.FileExists(str)
        Loop
        
        oExcelApp.Visible = True
        
        oExcelDoc.SaveAs str

        Screen.MousePointer = vbNormal
        
        ShowMessage "Excel spreadsheet generated", 1200, frmTMSPrinting
        
                
    Else
        ShowMessage "Nothing scheduled!", 1000, frmTMSPrinting
    End If
    
    
    'clean up and leave
    
    On Error Resume Next
    
    Screen.MousePointer = vbNormal
    
    rs.Close
    Set rs = Nothing

    
    If Not bForgetAboutExcel Then
        oExcelApp.Visible = True
    Else
        If Not mbExcelWasOpen Then
            'if there's nothing to show and Excel wasn't open before, shut it now.
            oExcelApp.DisplayAlerts = False
            oExcelApp.Quit
        End If
    End If
    
    Set oExcelApp = Nothing
    Set oExcelDoc = Nothing
    Set oExcelResultSheet = Nothing
    
    Screen.MousePointer = vbNormal
    
'    Application.Windows(oExcelDoc).Activate

    Exit Sub
    
ErrorTrap:

    On Error Resume Next
     
    str = Err.Description
     
    Screen.MousePointer = vbNormal
    
    If Not mbExcelWasOpen Then
        oExcelApp.DisplayAlerts = False
        oExcelApp.Quit
    Else
        oExcelApp.Visible = True
    End If
    
    Set oExcelApp = Nothing
    Set oExcelDoc = Nothing
    Set oExcelResultSheet = Nothing

'    Call EndProgram

    MsgBox "A problem occurred while processing the spreadsheet: " & str, vbOKOnly + vbExclamation, AppName

End Sub

Private Function BuildPrintTable(Optional pStartDate As Date, Optional pEndDate As Date)
Dim rstNewPrintedSchedule As Recordset
Dim rstSchedule As Recordset, ScheduleSQL As String
Dim TheStartDate As String, TheEndDate As String
Dim TheStartDateUS As String, TheEndDateUS As String
Dim PrevAssignmentDate As Date
Dim PersonName As String
Dim AssistantName As String
Dim bNewArr_Start As MidweekMtgVersion, bNewArr_End As MidweekMtgVersion
Dim i As Long, sThemeAndSource As String
Dim dte As Date, ErrorCode As Integer
Dim PrevItemsSeqNum As Long
Dim lSQ As Long, sSQ As String
Dim sTheme As String



On Error GoTo ErrorTrap

    DelAllRows "tblTMSPrintSchedule"
    
    
    'build the print table
     DeleteTable "tblTMSPrintSchedule"
     CreateTable ErrorCode, "tblTMSPrintSchedule", "ItemsSeqNum", "LONG", , , True, "SeqNum"
     CreateField ErrorCode, "tblTMSPrintSchedule", "TalkSeqNum", "LONG"
     CreateField ErrorCode, "tblTMSPrintSchedule", "AssignmentDate", "DATE"
     CreateField ErrorCode, "tblTMSPrintSchedule", "AssignmentDateStr", "TEXT"
     CreateField ErrorCode, "tblTMSPrintSchedule", "TalkType", "TEXT"
     CreateField ErrorCode, "tblTMSPrintSchedule", "Theme", "TEXT", "255"
     CreateField ErrorCode, "tblTMSPrintSchedule", "Student1", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "Assistant1", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "SQ1", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "Student2", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "Assistant2", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "SQ2", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "Student3", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "Assistant3", "TEXT", "100"
     CreateField ErrorCode, "tblTMSPrintSchedule", "SQ3", "TEXT", "100"
    
    
    With frmTMSPrinting
    
    If Not ValidDate(CStr(pStartDate)) Then
        If DateDiff("d", CDate(.cmbStartDate.text), CDate(.cmbEndDate.text)) > 366 Then
            MsgBox "Schedule should be no longer than one year", vbExclamation + vbOKOnly, AppName
            BuildPrintTable = False
            Exit Function
        End If
    
        bNewArr_Start = NewMtgArrangementStarted(.cmbStartDate.text)
        bNewArr_End = NewMtgArrangementStarted(.cmbEndDate.text)
        
        If bNewArr_Start <> bNewArr_End Then
            MsgBox "You cannot print a schedule spanning the old and new meeting arrangement " _
            , vbInformation + vbOKOnly, AppName
            BuildPrintTable = False
            Exit Function
        End If
        
        TheStartDateUS = Format(.cmbStartDate.text, "mm/dd/yyyy")
        TheEndDateUS = Format(.cmbEndDate.text, "mm/dd/yyyy")
        
    Else
        TheStartDateUS = Format(CStr(pStartDate), "mm/dd/yyyy")
        TheEndDateUS = Format(CStr(pEndDate), "mm/dd/yyyy")
    End If
    
    
    End With
    
    '
    'Driving Recset to get all assignments in date range
    '
    ScheduleSQL = "SELECT * FROM tblTMSSchedule " & _
                " WHERE AssignmentDate BETWEEN #" & TheStartDateUS & "# AND #" & _
                TheEndDateUS & "# " & _
                " AND (PersonID > 0 OR TalkNo IN ('A')) " & _
                " AND SchoolNo <= " & GlobalParms.GetValue("TMSNoSchoolsForSchedulePrint", "NumVal", 1) & " " & _
                " ORDER BY AssignmentDate, TalkSeqNum, ItemsSeqNum,  SchoolNo "

    Set rstSchedule = CMSDB.OpenRecordset(ScheduleSQL, dbOpenDynaset)
    
    
    If rstSchedule.BOF Then
        BuildPrintTable = False
        Exit Function
    Else
        PrevItemsSeqNum = rstSchedule!ItemsSeqNum
    End If
    
        
    Set rstNewPrintedSchedule = CMSDB.OpenRecordset("Select * FROM tblTMSPrintSchedule " & _
                                               " ORDER By AssignmentDate, TalkSeqNum, ItemsSeqNum", dbOpenDynaset)
        
    With rstNewPrintedSchedule
    Do Until rstSchedule.EOF
            
        .AddNew
        !AssignmentDateStr = CStr(Format(rstSchedule!AssignmentDate, "dd/mm/yyyy"))
        !AssignmentDate = rstSchedule!AssignmentDate
        !TalkSeqNum = rstSchedule!TalkSeqNum
        !ItemsSeqNum = rstSchedule!ItemsSeqNum
        !TalkType = TheTMS.GetTMSTalkDescription(rstSchedule!TalkNo, CStr(Format(rstSchedule!AssignmentDate, "dd/mm/yyyy")))
         
        If TheTMS.GetTMSItemThemeAndSource(rstSchedule!AssignmentDate, _
                                           rstSchedule!TalkNo, rstSchedule!ItemsSeqNum) = TMSOK Then
            sThemeAndSource = TheTMS.TMSThemeAndSource
        Else
            sThemeAndSource = ""
        End If
        
        !Theme = sThemeAndSource
        
        .Update
            
        Do While Not rstSchedule.EOF

            AssistantName = CongregationMember.NameWithMiddleInitial(rstSchedule!Assistant1ID)
            If Left(AssistantName, 1) = "?" Then
                AssistantName = ""
            End If
            
            PersonName = CongregationMember.NameWithMiddleInitial(rstSchedule!PersonID)
            If Left(PersonName, 1) = "?" Then
                PersonName = ""
            End If
            
            If rstSchedule!TalkNo <> "A" Then
                lSQ = rstSchedule!CounselPoint
                If lSQ > 0 Then
                    sSQ = CStr(lSQ) & "-" & TheTMS.GetTMSCounselDescription(lSQ)
                Else
                    sSQ = ""
                End If
            Else
                sSQ = ""
            End If
            
            .Requery
            .MoveLast
            .Edit
            .Fields("Student" & rstSchedule!SchoolNo) = PersonName
            .Fields("Assistant" & rstSchedule!SchoolNo) = AssistantName
            .Fields("SQ" & rstSchedule!SchoolNo) = sSQ
            
            .Update
            
            PrevItemsSeqNum = rstSchedule!ItemsSeqNum
            
            rstSchedule.MoveNext
            
            If Not rstSchedule.EOF Then
                If PrevItemsSeqNum <> rstSchedule!ItemsSeqNum Then
                    Exit Do
                End If
            End If
            
        Loop
            
        
        If rstSchedule.EOF Then
            Exit Do
        End If
            
    Loop
    End With
    
    If CopyPrintSchedule Then
        BuildPrintTable = True
    Else
        BuildPrintTable = False
        Exit Function
    End If
    
    Exit Function
ErrorTrap:
    EndProgram
    
End Function









