Attribute VB_Name = "basTMSWorkSheets"
Option Explicit
Public Sub PrintWorkTMSSheetsUsingWord()
Dim reporter As MSWordReportingTool2.RptTool
On Error GoTo ErrorTrap

    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .ReportSQL = "SELECT AssignmentDate, SchoolNo, OpeningSong, PrayerBroName, " & _
                "SQBroName, SQTheme, SQSource, No1BroName, No1Theme, " & _
                "No1Source, No1CounselPoint, No1CounselSubPoint1, No1CounselSubPoint2, " & _
                "No1CounselSubPoint3, No1CounselSubPoint4, No1CounselSubPoint5, " & _
                "No1Comment, BHBroName, BHSource, BHCounselPoint, BHCounselSubPoint1, " & _
                "BHCounselSubPoint2, BHCounselSubPoint3, BHCounselSubPoint4, " & _
                "BHCounselSubPoint5, BHComment, No2BroName, No2Theme, No2Source, " & _
                "No2CounselPoint, No2CounselSubPoint1, No2CounselSubPoint2, " & _
                "No2CounselSubPoint3, No2CounselSubPoint4, No2CounselSubPoint5, " & _
                "No2Comment, No3StudentName, No3AssistantName, No3Setting, No3Theme, " & _
                "No3Source, No3CounselPoint, No3CounselSubPoint1, No3CounselSubPoint2, " & _
                "No3CounselSubPoint3, No3CounselSubPoint4, No3CounselSubPoint5, " & _
                "No3Comment, No4StudentName, No4AssistantName, No4Setting, No4Theme, " & _
                "No4Source, No4CounselPoint, No4CounselSubPoint1, No4CounselSubPoint2, " & _
                "No4CounselSubPoint3, " & _
                "No4CounselSubPoint4, No4CounselSubPoint5, No4Comment, ConcludingSong " & _
                "FROM tblTMSPrintWorkSheet"


    .PageFormat = cmsPortrait
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 10
    .RightMargin = 10
    .ClientName = AppName
    .ShowPageNumber = False
    .HideWordWhileBuilding = True
    .NumberFormat = ""
    .ShowProgress = True
    .ReportTitle = "Theocratic Ministry School Worksheet"
    .SaveDoc = True
    .DocPath = gsDocsDirectory & "\" & "Theocratic Ministry School" & _
                                 " Worksheet " & _
                                Replace(Replace(Now, ":", "-"), "/", "-")
    
    
    
    .DocumentType = cmsPagePerDBRow
    
    ''
    'Because DocumentType = cmsPagePerDBRow, must supply cell-specific details....
    ''
    
    .TableCols_Adv = 10
    .TableRows_Adv = 41
    
    'put in fixed text, ie field labels. Not required if using a Word template
    .AddFixedCellText_Adv 1, 1, "School: ", True
    .AddFixedCellText_Adv 1, 3, "Theocratic Ministry School Worksheet"
    .AddFixedCellText_Adv 2, 1, "Song: "
    .AddFixedCellText_Adv 2, 7, "Prayer: "
    .AddFixedCellText_Adv 3, 1, "Speech Quality: "
    .AddFixedCellText_Adv 4, 1, "Talk No 1: "
    .AddFixedCellText_Adv 5, 1, "Counsel: "
    .AddFixedCellText_Adv 5, 4, "Last Comment: "
    .AddFixedCellText_Adv 5, 9, "Timing: " & vbCr & "Next Pt:"
    .AddFixedCellText_Adv 11, 5, "Bible Highlights: "
    .AddFixedCellText_Adv 11, 1, "B/H: "
    .AddFixedCellText_Adv 12, 1, "Counsel: "
    .AddFixedCellText_Adv 12, 4, "Last Comment: "
    .AddFixedCellText_Adv 12, 9, "Timing: " & vbCr & "Next Pt:"
    .AddFixedCellText_Adv 18, 1, "Talk No 2: "
    .AddFixedCellText_Adv 19, 1, "Counsel: "
    .AddFixedCellText_Adv 19, 4, "Last Comment: "
    .AddFixedCellText_Adv 19, 9, "Timing: " & vbCr & "Next Pt:"
    .AddFixedCellText_Adv 25, 1, "Talk No 3: "
    .AddFixedCellText_Adv 26, 1, "Asst: "
    .AddFixedCellText_Adv 26, 4, "Setting: "
    .AddFixedCellText_Adv 27, 1, "Counsel: "
    .AddFixedCellText_Adv 27, 4, "Last Comment: "
    .AddFixedCellText_Adv 27, 9, "Timing: " & vbCr & "Next Pt:"
    .AddFixedCellText_Adv 33, 1, "Talk No 4: "
    .AddFixedCellText_Adv 34, 1, "Asst: "
    .AddFixedCellText_Adv 34, 4, "Setting: "
    .AddFixedCellText_Adv 35, 1, "Counsel: "
    .AddFixedCellText_Adv 35, 4, "Last Comment: "
    .AddFixedCellText_Adv 35, 9, "Timing: " & vbCr & "Next Pt:"
'    .AddFixedCellText_Adv 41, 1, "Invite Service Meeting chairman to the platform"
    
    'Mapping of DB fields to table cells.
    ' listed in order of fields in ReportSQL.
    
    .AddFieldMapping_Adv 1, 10, True    'assignment date
    .AddFieldMapping_Adv 1, 1           'school no
    .AddFieldMapping_Adv 2, 2           'Song
    .AddFieldMapping_Adv 2, 8           'Prayer Bro
    .AddFieldMapping_Adv 3, 3           'SQ bro
    .AddFieldMapping_Adv 3, 5           'SQ theme
    .AddFieldMapping_Adv 3, 8           'SQ source

    .AddFieldMapping_Adv 4, 3           'No1 Bro
    .AddFieldMapping_Adv 4, 5           'No1 theme
    .AddFieldMapping_Adv 4, 8           'No1 source
    .AddFieldMapping_Adv 5, 3           'No1 counsel point
    .AddFieldMapping_Adv 6, 1           'No1 counsel sub-point1
    .AddFieldMapping_Adv 7, 1           'No1 counsel sub-point2
    .AddFieldMapping_Adv 8, 1           'No1 counsel sub-point3
    .AddFieldMapping_Adv 9, 1           'No1 counsel sub-point4
    .AddFieldMapping_Adv 10, 1          'No1 counsel sub-point5
    .AddFieldMapping_Adv 5, 5          'No1 comment

    .AddFieldMapping_Adv 11, 3           'BH Bro
    .AddFieldMapping_Adv 11, 8           'BH source
    .AddFieldMapping_Adv 12, 3           'BH counsel point
    .AddFieldMapping_Adv 13, 1           'BH counsel sub-point1
    .AddFieldMapping_Adv 14, 1           'BH counsel sub-point2
    .AddFieldMapping_Adv 15, 1           'BH counsel sub-point3
    .AddFieldMapping_Adv 16, 1           'BH counsel sub-point4
    .AddFieldMapping_Adv 17, 1          'BH counsel sub-point5
    .AddFieldMapping_Adv 12, 5          'BH comment

    .AddFieldMapping_Adv 18, 3           'No2  Bro
    .AddFieldMapping_Adv 18, 5           'No2  theme
    .AddFieldMapping_Adv 18, 8           'No2  source
    .AddFieldMapping_Adv 19, 3           'No2  counsel point
    .AddFieldMapping_Adv 20, 1           'No2  counsel sub-point1
    .AddFieldMapping_Adv 21, 1           'No2  counsel sub-point2
    .AddFieldMapping_Adv 22, 1           'No2  counsel sub-point3
    .AddFieldMapping_Adv 23, 1           'No2  counsel sub-point4
    .AddFieldMapping_Adv 24, 1          'No2  counsel sub-point5
    .AddFieldMapping_Adv 19, 5          'No2  comment

    .AddFieldMapping_Adv 25, 3           'No3  Bro
    .AddFieldMapping_Adv 26, 3           'No3  asst
    .AddFieldMapping_Adv 26, 5           'No3  setting
    .AddFieldMapping_Adv 25, 5           'No3  theme
    .AddFieldMapping_Adv 25, 8           'No3  source
    .AddFieldMapping_Adv 27, 3           'No3  counsel point
    .AddFieldMapping_Adv 28, 1           'No3  counsel sub-point1
    .AddFieldMapping_Adv 29, 1           'No3  counsel sub-point2
    .AddFieldMapping_Adv 30, 1           'No3  counsel sub-point3
    .AddFieldMapping_Adv 31, 1           'No3  counsel sub-point4
    .AddFieldMapping_Adv 32, 1          'No3  counsel sub-point5
    .AddFieldMapping_Adv 27, 5          'No3  comment

    .AddFieldMapping_Adv 33, 3           'No4  Bro
    .AddFieldMapping_Adv 34, 3           'No4  asst
    .AddFieldMapping_Adv 26, 5           'No3  setting
    .AddFieldMapping_Adv 33, 5           'No4  theme
    .AddFieldMapping_Adv 33, 8           'No4  source
    .AddFieldMapping_Adv 35, 3           'No4  counsel point
    .AddFieldMapping_Adv 36, 1           'No4  counsel sub-point1
    .AddFieldMapping_Adv 37, 1           'No4  counsel sub-point2
    .AddFieldMapping_Adv 38, 1           'No4  counsel sub-point3
    .AddFieldMapping_Adv 39, 1           'No4  counsel sub-point4
    .AddFieldMapping_Adv 40, 1          'No4  counsel sub-point5
    .AddFieldMapping_Adv 35, 5          'No4  comment
    .AddFieldMapping_Adv 41, 1          'Concluding Song
'
    '
    'these settings would be ignorred if using word template with pre-built table
    '
    'apply font size to whole table
    .AddCellAttributes_Adv 1, 1, 41, 10, cmsleftTop, "Times New Roman", 10, cmsOptionfalse, _
                                                        cmsOptionfalse, , , , True
                                                        
    'apply bold, centred formatting to the "TMS Worksheet" cells
    .AddCellAttributes_Adv 1, 3, 1, 8, cmsCentreTop, , 12, cmsOptionTrue, _
                                                        cmsOptionTrue
                           
    'bold for School No
    .AddCellAttributes_Adv 1, 1, , , cmsleftTop, , 12, cmsOptionTrue, cmsOptionTrue
    
    'bold for Talk No1 label
    .AddCellAttributes_Adv 4, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No1 counsel label
    .AddCellAttributes_Adv 5, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No1 comment label
    .AddCellAttributes_Adv 5, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No1 timing, next pt labels
    .AddCellAttributes_Adv 5, 9, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
'
    'bold for BH label
    .AddCellAttributes_Adv 11, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for BH counsel label
    .AddCellAttributes_Adv 12, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for BH comment label
    .AddCellAttributes_Adv 12, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for BH timing, next pt labels
    .AddCellAttributes_Adv 12, 9, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for School No
    .AddCellAttributes_Adv 1, 1, , , cmsleftTop, , 12, cmsOptionTrue, cmsOptionTrue
    
'
    'bold for Talk No2 label
    .AddCellAttributes_Adv 18, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No2 counsel label
    .AddCellAttributes_Adv 19, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No2 comment label
    .AddCellAttributes_Adv 19, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No2 timing, next pt labels
    .AddCellAttributes_Adv 19, 9, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse

'
    'bold for Talk No3 label
    .AddCellAttributes_Adv 25, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No3 ASSISTANT
    .AddCellAttributes_Adv 26, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No3 setting
    .AddCellAttributes_Adv 26, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No3 counsel label
    .AddCellAttributes_Adv 27, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No3 comment label
    .AddCellAttributes_Adv 27, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No3 timing, next pt labels
    .AddCellAttributes_Adv 27, 9, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse

'
    'bold for Talk No4 label
    .AddCellAttributes_Adv 33, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No4 ASSISTANT
    .AddCellAttributes_Adv 34, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No4 setting
    .AddCellAttributes_Adv 34, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No4 counsel label
    .AddCellAttributes_Adv 35, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No4 comment label
    .AddCellAttributes_Adv 35, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No4 timing, next pt labels
    .AddCellAttributes_Adv 35, 9, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for concluding song label
    .AddCellAttributes_Adv 41, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    
    
    'bold, right-aligned assignment date
    .AddCellAttributes_Adv 1, 9, 1, 10, cmsRightTop, , 12, cmsOptionTrue, cmsOptionTrue
    
    'song no label
    .AddCellAttributes_Adv 2, 1, , , cmsleftTop, , , cmsOptionTrue, cmsOptionfalse
    
    'prayer label
    .AddCellAttributes_Adv 2, 7, , , cmsleftTop, , , cmsOptionTrue, cmsOptionfalse
    
    'SQ label
    .AddCellAttributes_Adv 3, 1, , , cmsleftTop, , , cmsOptionTrue, cmsOptionfalse
    
    'apply thicker border to divide by assignment
    .AddCellAttributes_Adv 3, 1, 3, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    .AddCellAttributes_Adv 10, 1, 10, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    .AddCellAttributes_Adv 17, 1, 17, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    .AddCellAttributes_Adv 24, 1, 24, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    .AddCellAttributes_Adv 32, 1, 32, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    .AddCellAttributes_Adv 40, 1, 40, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    
    
    '
    'Join cells. ie, make border between them invisible
    '
    .AddJoinedCellRange_Adv 2, 1, 2, 6, True
    .AddJoinedCellRange_Adv 2, 7, 2, 10
    .AddJoinedCellRange_Adv 5, 1, 5, 2
    .AddJoinedCellRange_Adv 5, 9, 5, 10
'    .AddJoinedCellRange_Adv 41, 2, 41, 3
'    .AddJoinedCellRange_Adv 41, 3, 41, 4
    
    '
    'Merge cells. This should be done from right to left and bottom to top
    '
    .AddMergeCellRange_Adv 41, 1, 41, 10, True
    .AddMergeCellRange_Adv 40, 4, 40, 10
    .AddMergeCellRange_Adv 40, 1, 40, 3
    .AddMergeCellRange_Adv 39, 4, 39, 10
    .AddMergeCellRange_Adv 39, 1, 39, 3
    .AddMergeCellRange_Adv 38, 4, 38, 10
    .AddMergeCellRange_Adv 38, 1, 38, 3
    .AddMergeCellRange_Adv 37, 4, 37, 10
    .AddMergeCellRange_Adv 37, 1, 37, 3
    .AddMergeCellRange_Adv 36, 4, 36, 10
    .AddMergeCellRange_Adv 36, 1, 36, 3
    .AddMergeCellRange_Adv 35, 5, 35, 8
    .AddMergeCellRange_Adv 35, 2, 35, 3
    .AddMergeCellRange_Adv 34, 4, 34, 10
    .AddMergeCellRange_Adv 34, 2, 34, 3
    .AddMergeCellRange_Adv 33, 8, 33, 10
    .AddMergeCellRange_Adv 33, 5, 33, 7
    .AddMergeCellRange_Adv 33, 2, 33, 4
        
    .AddMergeCellRange_Adv 32, 4, 32, 10
    .AddMergeCellRange_Adv 32, 1, 32, 3
    .AddMergeCellRange_Adv 31, 4, 31, 10
    .AddMergeCellRange_Adv 31, 1, 31, 3
    .AddMergeCellRange_Adv 30, 4, 30, 10
    .AddMergeCellRange_Adv 30, 1, 30, 3
    .AddMergeCellRange_Adv 29, 4, 29, 10
    .AddMergeCellRange_Adv 29, 1, 29, 3
    .AddMergeCellRange_Adv 28, 4, 28, 10
    .AddMergeCellRange_Adv 28, 1, 28, 3
    .AddMergeCellRange_Adv 27, 5, 27, 8
    .AddMergeCellRange_Adv 27, 2, 27, 3
    .AddMergeCellRange_Adv 26, 4, 26, 10
    .AddMergeCellRange_Adv 26, 2, 26, 3
    .AddMergeCellRange_Adv 25, 8, 25, 10
    .AddMergeCellRange_Adv 25, 5, 25, 7
    .AddMergeCellRange_Adv 25, 2, 25, 4
    
    .AddMergeCellRange_Adv 24, 4, 24, 10
    .AddMergeCellRange_Adv 24, 1, 24, 3
    .AddMergeCellRange_Adv 23, 4, 23, 10
    .AddMergeCellRange_Adv 23, 1, 23, 3
    .AddMergeCellRange_Adv 22, 4, 22, 10
    .AddMergeCellRange_Adv 22, 1, 22, 3
    .AddMergeCellRange_Adv 21, 4, 21, 10
    .AddMergeCellRange_Adv 21, 1, 21, 3
    .AddMergeCellRange_Adv 20, 4, 20, 10
    .AddMergeCellRange_Adv 20, 1, 20, 3
    .AddMergeCellRange_Adv 19, 5, 19, 8
    .AddMergeCellRange_Adv 19, 2, 19, 3
    .AddMergeCellRange_Adv 18, 8, 18, 10
    .AddMergeCellRange_Adv 18, 5, 18, 7
    .AddMergeCellRange_Adv 18, 2, 18, 4
    
    .AddMergeCellRange_Adv 17, 4, 17, 10
    .AddMergeCellRange_Adv 17, 1, 17, 3
    .AddMergeCellRange_Adv 16, 4, 16, 10
    .AddMergeCellRange_Adv 16, 1, 16, 3
    .AddMergeCellRange_Adv 15, 4, 15, 10
    .AddMergeCellRange_Adv 15, 1, 15, 3
    .AddMergeCellRange_Adv 14, 4, 14, 10
    .AddMergeCellRange_Adv 14, 1, 14, 3
    .AddMergeCellRange_Adv 13, 4, 13, 10
    .AddMergeCellRange_Adv 13, 1, 13, 3
    .AddMergeCellRange_Adv 12, 5, 12, 8
    .AddMergeCellRange_Adv 12, 2, 12, 3
    .AddMergeCellRange_Adv 11, 8, 11, 10
    .AddMergeCellRange_Adv 11, 5, 11, 7
    .AddMergeCellRange_Adv 11, 2, 11, 4
    
    .AddMergeCellRange_Adv 10, 4, 10, 10
    .AddMergeCellRange_Adv 10, 1, 10, 3
    .AddMergeCellRange_Adv 9, 4, 9, 10
    .AddMergeCellRange_Adv 9, 1, 9, 3
    .AddMergeCellRange_Adv 8, 4, 8, 10
    .AddMergeCellRange_Adv 8, 1, 8, 3
    .AddMergeCellRange_Adv 7, 4, 7, 10
    .AddMergeCellRange_Adv 7, 1, 7, 3
    .AddMergeCellRange_Adv 6, 4, 6, 10
    .AddMergeCellRange_Adv 6, 1, 6, 3
    .AddMergeCellRange_Adv 5, 5, 5, 8
    .AddMergeCellRange_Adv 5, 2, 5, 3
    .AddMergeCellRange_Adv 4, 8, 4, 10
    .AddMergeCellRange_Adv 4, 5, 4, 7
    .AddMergeCellRange_Adv 4, 2, 4, 4
    
    .AddMergeCellRange_Adv 3, 8, 3, 10
    .AddMergeCellRange_Adv 3, 5, 3, 7
    .AddMergeCellRange_Adv 3, 3, 3, 4
    .AddMergeCellRange_Adv 3, 1, 3, 2
    .AddMergeCellRange_Adv 2, 8, 2, 10
    .AddMergeCellRange_Adv 2, 2, 2, 6
    .AddMergeCellRange_Adv 1, 9, 1, 10
    .AddMergeCellRange_Adv 1, 3, 1, 8
    .AddMergeCellRange_Adv 1, 1, 1, 2
    
    .GenerateReport
       
    End With
    
    SwitchOnDAO
    
    Set reporter = Nothing
    
    Screen.MousePointer = vbNormal

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Public Sub PrintWorkTMSSheetsUsingWord_2009()
Dim reporter As MSWordReportingTool2.RptTool
On Error GoTo ErrorTrap

    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .ReportSQL = "SELECT AssignmentDate2, SchoolNo, " & _
                "BHBroName, BHSource, BHCounselPoint, BHCounselSubPoint1, " & _
                "BHCounselSubPoint2, BHCounselSubPoint3, BHCounselSubPoint4, " & _
                "BHCounselSubPoint5, BHComment, No1BroName, No1Theme, No1Source, " & _
                "No1CounselPoint, No1CounselSubPoint1, No1CounselSubPoint2, " & _
                "No1CounselSubPoint3, No1CounselSubPoint4, No1CounselSubPoint5, " & _
                "No1Comment, No2BroName, No2AssistantName, No2Setting, No2Theme, " & _
                "No2Source, No2CounselPoint, No2CounselSubPoint1, No2CounselSubPoint2, " & _
                "No2CounselSubPoint3, No2CounselSubPoint4, No2CounselSubPoint5, " & _
                "No2Comment, No3StudentName, No3AssistantName, No3Setting, No3Theme, " & _
                "No3Source, No3CounselPoint, No3CounselSubPoint1, No3CounselSubPoint2, " & _
                "No3CounselSubPoint3, " & _
                "No3CounselSubPoint4, No3CounselSubPoint5, No3Comment, ConcludingSong " & _
                "FROM tblTMSPrintWorkSheet"


    .PageFormat = cmsPortrait
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 10
    .RightMargin = 10
    .ClientName = AppName
    .ShowPageNumber = False
    .HideWordWhileBuilding = True
    .NumberFormat = ""
    .ShowProgress = True
    .ReportTitle = ""
    .SaveDoc = True
    .DocPath = gsDocsDirectory & "\" & "Theocratic Ministry School" & _
                                 " Worksheet for Week " & Replace(frmTMSPrinting.cmbStartDate.text, "/", "-") & " (" & _
                                Replace(Replace(Now, ":", "-"), "/", "-") & ")"
    
    
    'use this if using a word template with pre-built table in it
'    .WordTemplatePath_Adv = App.Path & "\School.dot"
    
    .DocumentType = cmsPagePerDBRow
    
    ''
    'Because DocumentType = cmsPagePerDBRow, must supply cell-specific details....
    ''
    
    .TableCols_Adv = 10
    .TableRows_Adv = 33
    
    'put in fixed text, ie field labels. Not required if using a Word template
    .AddFixedCellText_Adv 1, 1, "School: ", True
    .AddFixedCellText_Adv 1, 3, "Theocratic Ministry School Worksheet"
'    .AddFixedCellText_Adv 2, 1, "Song: "
'    .AddFixedCellText_Adv 2, 7, "Prayer: "
    .AddFixedCellText_Adv 3, 5, "Bible Highlights: "
    .AddFixedCellText_Adv 3, 1, "B/H: "
    .AddFixedCellText_Adv 4, 1, "Counsel: "
    .AddFixedCellText_Adv 4, 4, "Last Comment: "
    .AddFixedCellText_Adv 4, 9, "Timing: " & vbCr & "Ex: " & vbCr & "Next Pt: "
    .AddFixedCellText_Adv 10, 1, "Talk No 1: "
    .AddFixedCellText_Adv 11, 1, "Counsel: "
    .AddFixedCellText_Adv 11, 4, "Last Comment: "
    .AddFixedCellText_Adv 11, 9, "Timing: " & vbCr & "Ex: " & vbCr & "Next Pt: "
    .AddFixedCellText_Adv 17, 1, "Talk No 2: "
    .AddFixedCellText_Adv 18, 1, "Asst: "
    .AddFixedCellText_Adv 18, 4, "Setting: "
    .AddFixedCellText_Adv 19, 1, "Counsel: "
    .AddFixedCellText_Adv 19, 4, "Last Comment: "
    .AddFixedCellText_Adv 19, 9, "Timing: " & vbCr & "Ex: " & vbCr & "Next Pt: "
    .AddFixedCellText_Adv 25, 1, "Talk No 3: "
    .AddFixedCellText_Adv 26, 1, "Asst: "
    .AddFixedCellText_Adv 26, 4, "Setting: "
    .AddFixedCellText_Adv 27, 1, "Counsel: "
    .AddFixedCellText_Adv 27, 4, "Last Comment: "
    .AddFixedCellText_Adv 27, 9, "Timing: " & vbCr & "Ex: " & vbCr & "Next Pt: "
    
    'Mapping of DB fields to table cells.
    ' listed in order of fields in ReportSQL.
    
    .AddFieldMapping_Adv 1, 10, True    'assignment date
    .AddFieldMapping_Adv 1, 1           'school no
'    .AddFieldMapping_Adv 2, 2           'Song
'    .AddFieldMapping_Adv 2, 8           'Prayer Bro

    .AddFieldMapping_Adv 3, 3           'BH Bro
    .AddFieldMapping_Adv 3, 8           'BH source
    .AddFieldMapping_Adv 4, 3           'BH counsel point
    .AddFieldMapping_Adv 5, 1           'BH counsel sub-point1
    .AddFieldMapping_Adv 6, 1           'BH counsel sub-point2
    .AddFieldMapping_Adv 7, 1           'BH counsel sub-point3
    .AddFieldMapping_Adv 8, 1           'BH counsel sub-point4
    .AddFieldMapping_Adv 9, 1          'BH counsel sub-point5
    .AddFieldMapping_Adv 4, 5          'BH comment

    .AddFieldMapping_Adv 10, 3           'No1  Bro
    .AddFieldMapping_Adv 10, 5           'No1  theme
    .AddFieldMapping_Adv 10, 8           'No1  source
    .AddFieldMapping_Adv 11, 3           'No1  counsel point
    .AddFieldMapping_Adv 12, 1           'No1  counsel sub-point1
    .AddFieldMapping_Adv 13, 1           'No1  counsel sub-point2
    .AddFieldMapping_Adv 14, 1           'No1  counsel sub-point3
    .AddFieldMapping_Adv 15, 1           'No1  counsel sub-point4
    .AddFieldMapping_Adv 16, 1          'No1  counsel sub-point5
    .AddFieldMapping_Adv 11, 5          'No1  comment

    .AddFieldMapping_Adv 17, 3           'No2  Bro
    .AddFieldMapping_Adv 18, 3           'No2  asst
    .AddFieldMapping_Adv 18, 5           'No2  setting
    .AddFieldMapping_Adv 17, 5            'No2  theme
    .AddFieldMapping_Adv 17, 8           'No2  source
    .AddFieldMapping_Adv 19, 3           'No2  counsel point
    .AddFieldMapping_Adv 20, 1           'No2  counsel sub-point1
    .AddFieldMapping_Adv 21, 1           'No2  counsel sub-point2
    .AddFieldMapping_Adv 22, 1           'No2  counsel sub-point3
    .AddFieldMapping_Adv 23, 1           'No2  counsel sub-point4
    .AddFieldMapping_Adv 24, 1          'No2  counsel sub-point5
    .AddFieldMapping_Adv 19, 5          'No2  comment

    .AddFieldMapping_Adv 25, 3           'No3  Bro
    .AddFieldMapping_Adv 26, 3           'No3  asst
    .AddFieldMapping_Adv 26, 5           'No3  setting
    .AddFieldMapping_Adv 25, 5           'No3  theme
    .AddFieldMapping_Adv 25, 8           'No3  source
    .AddFieldMapping_Adv 27, 3           'No3  counsel point
    .AddFieldMapping_Adv 28, 1           'No3  counsel sub-point1
    .AddFieldMapping_Adv 29, 1           'No3  counsel sub-point2
    .AddFieldMapping_Adv 30, 1           'No3  counsel sub-point3
    .AddFieldMapping_Adv 31, 1           'No3  counsel sub-point4
    .AddFieldMapping_Adv 32, 1          'No3  counsel sub-point5
    .AddFieldMapping_Adv 27, 5          'No3  comment
    .AddFieldMapping_Adv 33, 1          'Concluding Song
'
    '
    'these settings would be ignorred if using word template with pre-built table
    '
    'apply font size to whole table
    .AddCellAttributes_Adv 1, 1, 33, 10, cmsleftTop, "Times New Roman", 10, cmsOptionfalse, _
                                                        cmsOptionfalse, , , , True
                                                        
    'apply bold, centred formatting to the "TMS Worksheet" cells
    .AddCellAttributes_Adv 1, 3, 1, 8, cmsCentreTop, , 12, cmsOptionTrue, _
                                                        cmsOptionTrue
                           
    'bold for School No
    .AddCellAttributes_Adv 1, 1, , , cmsleftTop, , 12, cmsOptionTrue, cmsOptionTrue
        
'
    'bold for BH label
    .AddCellAttributes_Adv 3, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for BH counsel label
    .AddCellAttributes_Adv 4, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for BH comment label
    .AddCellAttributes_Adv 4, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for BH timing, next pt labels
    .AddCellAttributes_Adv 4, 9, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
       
'
    'bold for Talk No1 label
    .AddCellAttributes_Adv 10, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No1 counsel label
    .AddCellAttributes_Adv 11, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No1 comment label
    .AddCellAttributes_Adv 11, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No1 timing, next pt labels
    .AddCellAttributes_Adv 11, 9, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse

'
    'bold for Talk No2 label
    .AddCellAttributes_Adv 17, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No2 ASSISTANT
    .AddCellAttributes_Adv 18, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No2 setting
    .AddCellAttributes_Adv 18, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No2 counsel label
    .AddCellAttributes_Adv 19, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No2 comment label
    .AddCellAttributes_Adv 19, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No2 timing, next pt labels
    .AddCellAttributes_Adv 19, 9, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse

'
    'bold for Talk No3 label
    .AddCellAttributes_Adv 25, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No3 ASSISTANT
    .AddCellAttributes_Adv 26, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No3 setting
    .AddCellAttributes_Adv 26, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No3 counsel label
    .AddCellAttributes_Adv 27, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No3 comment label
    .AddCellAttributes_Adv 27, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for Talk No3 timing, next pt labels
    .AddCellAttributes_Adv 27, 9, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    'bold for concluding song label
    .AddCellAttributes_Adv 33, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    
    
    'bold, right-aligned assignment date
    .AddCellAttributes_Adv 1, 9, 1, 10, cmsRightTop, , 12, cmsOptionTrue, cmsOptionTrue
    
'    'song no label
'    .AddCellAttributes_Adv 2, 1, , , cmsleftTop, , , cmsOptionTrue, cmsOptionfalse
'
'    'prayer label
'    .AddCellAttributes_Adv 2, 7, , , cmsleftTop, , , cmsOptionTrue, cmsOptionfalse
    
    
    'apply thicker border to divide by assignment
    .AddCellAttributes_Adv 2, 1, 2, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    .AddCellAttributes_Adv 9, 1, 9, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    .AddCellAttributes_Adv 16, 1, 16, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    .AddCellAttributes_Adv 24, 1, 24, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    .AddCellAttributes_Adv 32, 1, 32, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    
    
    '
    'Join cells. ie, make border between them invisible
    '
    .AddJoinedCellRange_Adv 2, 1, 2, 6, True
    .AddJoinedCellRange_Adv 2, 7, 2, 10
    
    '
    'Merge cells. This should be done from right to left and bottom to top
    '
    .AddMergeCellRange_Adv 33, 1, 33, 10, True
    .AddMergeCellRange_Adv 32, 4, 32, 10
    .AddMergeCellRange_Adv 32, 1, 32, 3
    .AddMergeCellRange_Adv 31, 4, 31, 10
    .AddMergeCellRange_Adv 31, 1, 31, 3
    .AddMergeCellRange_Adv 30, 4, 30, 10
    .AddMergeCellRange_Adv 30, 1, 30, 3
    .AddMergeCellRange_Adv 29, 4, 29, 10
    .AddMergeCellRange_Adv 29, 1, 29, 3
    .AddMergeCellRange_Adv 28, 4, 28, 10
    .AddMergeCellRange_Adv 28, 1, 28, 3
    .AddMergeCellRange_Adv 27, 5, 27, 8
    .AddMergeCellRange_Adv 27, 2, 27, 3
    .AddMergeCellRange_Adv 26, 4, 26, 10
    .AddMergeCellRange_Adv 26, 2, 26, 3
    .AddMergeCellRange_Adv 25, 8, 25, 10
    .AddMergeCellRange_Adv 25, 5, 25, 7
    .AddMergeCellRange_Adv 25, 2, 25, 4
        
    .AddMergeCellRange_Adv 24, 4, 24, 10
    .AddMergeCellRange_Adv 24, 1, 24, 3
    .AddMergeCellRange_Adv 23, 4, 23, 10
    .AddMergeCellRange_Adv 23, 1, 23, 3
    .AddMergeCellRange_Adv 22, 4, 22, 10
    .AddMergeCellRange_Adv 22, 1, 22, 3
    .AddMergeCellRange_Adv 21, 4, 21, 10
    .AddMergeCellRange_Adv 21, 1, 21, 3
    .AddMergeCellRange_Adv 20, 4, 20, 10
    .AddMergeCellRange_Adv 20, 1, 20, 3
    .AddMergeCellRange_Adv 19, 5, 19, 8
    .AddMergeCellRange_Adv 19, 2, 19, 3
    .AddMergeCellRange_Adv 18, 4, 18, 10
    .AddMergeCellRange_Adv 18, 2, 18, 3
    .AddMergeCellRange_Adv 17, 8, 17, 10
    .AddMergeCellRange_Adv 17, 5, 17, 7
    .AddMergeCellRange_Adv 17, 2, 17, 4
    
    .AddMergeCellRange_Adv 16, 4, 16, 10
    .AddMergeCellRange_Adv 16, 1, 16, 3
    .AddMergeCellRange_Adv 15, 4, 15, 10
    .AddMergeCellRange_Adv 15, 1, 15, 3
    .AddMergeCellRange_Adv 14, 4, 14, 10
    .AddMergeCellRange_Adv 14, 1, 14, 3
    .AddMergeCellRange_Adv 13, 4, 13, 10
    .AddMergeCellRange_Adv 13, 1, 13, 3
    .AddMergeCellRange_Adv 12, 4, 12, 10
    .AddMergeCellRange_Adv 12, 1, 12, 3
    .AddMergeCellRange_Adv 11, 5, 11, 8
    .AddMergeCellRange_Adv 11, 2, 11, 3
    .AddMergeCellRange_Adv 10, 8, 10, 10
    .AddMergeCellRange_Adv 10, 5, 10, 7
    .AddMergeCellRange_Adv 10, 2, 10, 4
    
    .AddMergeCellRange_Adv 9, 4, 9, 10
    .AddMergeCellRange_Adv 9, 1, 9, 3
    .AddMergeCellRange_Adv 8, 4, 8, 10
    .AddMergeCellRange_Adv 8, 1, 8, 3
    .AddMergeCellRange_Adv 7, 4, 7, 10
    .AddMergeCellRange_Adv 7, 1, 7, 3
    .AddMergeCellRange_Adv 6, 4, 6, 10
    .AddMergeCellRange_Adv 6, 1, 6, 3
    .AddMergeCellRange_Adv 5, 4, 5, 10
    .AddMergeCellRange_Adv 5, 1, 5, 3
    .AddMergeCellRange_Adv 4, 5, 4, 8
    .AddMergeCellRange_Adv 4, 2, 4, 3
    .AddMergeCellRange_Adv 3, 8, 3, 10
    .AddMergeCellRange_Adv 3, 5, 3, 7
    .AddMergeCellRange_Adv 3, 2, 3, 4
    
    .AddMergeCellRange_Adv 2, 1, 2, 10
    .AddMergeCellRange_Adv 1, 8, 1, 10
    .AddMergeCellRange_Adv 1, 3, 1, 7
    .AddMergeCellRange_Adv 1, 1, 1, 2
    
    .GenerateReport
       
    End With
    
    SwitchOnDAO
    
    Set reporter = Nothing
    
    Screen.MousePointer = vbNormal
    
    If GlobalParms.GetValue("TMS_PrintSQListWithWorksheet", "TrueFalse", True) = True Then
        If MsgBox("Print Speech Quality update list?", vbYesNo + vbQuestion, AppName) = vbYes Then
            PrintBEBookUpdateList frmTMSPrinting.cmbStartDate.text
        End If
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Public Sub PrintWorkTMSSheetsUsingWord_2016()
Dim reporter As MSWordReportingTool2.RptTool
Dim rs As Recordset, sSQL As String, i As Integer, iNoAssignments As Integer, iTableRowsPerTalkNo As Integer
Dim bMergeInit As Boolean

On Error GoTo ErrorTrap

    sSQL = "SELECT AssignmentDate2, SchoolNo, " & _
            "Item1TalkNo, Item1StudentName, Item1AssistantName, Item1Theme, Item1Source, Item1CounselPoint, Item1CounselSubPoint1, Item1CounselSubPoint2, Item1CounselSubPoint3, Item1CounselSubPoint4, Item1CounselSubPoint5, Item1Comment, Item1SettingName, Item1SettingTitle, " & _
            "Item2TalkNo, Item2StudentName, Item2AssistantName, Item2Theme, Item2Source, Item2CounselPoint, Item2CounselSubPoint1, Item2CounselSubPoint2, Item2CounselSubPoint3, Item2CounselSubPoint4, Item2CounselSubPoint5, Item2Comment, Item2SettingName, Item2SettingTitle, " & _
            "Item3TalkNo, Item3StudentName, Item3AssistantName, Item3Theme, Item3Source, Item3CounselPoint, Item3CounselSubPoint1, Item3CounselSubPoint2, Item3CounselSubPoint3, Item3CounselSubPoint4, Item3CounselSubPoint5, Item3Comment, Item3SettingName, Item3SettingTitle, " & _
            "Item4TalkNo, Item4StudentName, Item4AssistantName, Item4Theme, Item4Source, Item4CounselPoint, Item4CounselSubPoint1, Item4CounselSubPoint2, Item4CounselSubPoint3, Item4CounselSubPoint4, Item4CounselSubPoint5, Item4Comment, Item4SettingName, Item4SettingTitle, " & _
            "Item5TalkNo, Item5StudentName, Item5AssistantName, Item5Theme, Item5Source, Item5CounselPoint, Item5CounselSubPoint1, Item5CounselSubPoint2, Item5CounselSubPoint3, Item5CounselSubPoint4, Item5CounselSubPoint5, Item5Comment, Item5SettingName, Item5SettingTitle " & _
            "FROM tblTMSPrintWorkSheet"
            
    Set rs = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)
    
    iNoAssignments = 0
    iTableRowsPerTalkNo = 8
    
    For i = 1 To 5
        If rs.Fields("Item" & i & "TalkNo") <> "" Then
            iNoAssignments = iNoAssignments + 1
        End If
    Next i
    
    rs.Close
    Set rs = Nothing
       
    If iNoAssignments = 0 Then
        ShowMessage "Nothing to print!", 1500, frmTMSPrinting
        Exit Sub
    End If
    
    SwitchOffDAO
    
    sSQL = "SELECT AssignmentDate2, SchoolNo, " & _
            "Item1TalkNo, Item1StudentName, Item1AssistantName, Item1Theme, Item1Source, Item1CounselPoint, Item1CounselSubPoint1, Item1CounselSubPoint2, Item1CounselSubPoint3, Item1CounselSubPoint4, Item1CounselSubPoint5, Item1Comment, Item1SettingName, Item1SettingTitle "

    If iNoAssignments >= 2 Then
        sSQL = sSQL & " ,Item2TalkNo, Item2StudentName, Item2AssistantName, Item2Theme, Item2Source, Item2CounselPoint, Item2CounselSubPoint1, Item2CounselSubPoint2, Item2CounselSubPoint3, Item2CounselSubPoint4, Item2CounselSubPoint5, Item2Comment, Item2SettingName, Item2SettingTitle "
    End If
    
    If iNoAssignments >= 3 Then
        sSQL = sSQL & " ,Item3TalkNo, Item3StudentName, Item3AssistantName, Item3Theme, Item3Source, Item3CounselPoint, Item3CounselSubPoint1, Item3CounselSubPoint2, Item3CounselSubPoint3, Item3CounselSubPoint4, Item3CounselSubPoint5, Item3Comment, Item3SettingName, Item3SettingTitle "
    End If
    
    If iNoAssignments >= 4 Then
        sSQL = sSQL & " ,Item4TalkNo, Item4StudentName, Item4AssistantName, Item4Theme, Item4Source, Item4CounselPoint, Item4CounselSubPoint1, Item4CounselSubPoint2, Item4CounselSubPoint3, Item4CounselSubPoint4, Item4CounselSubPoint5, Item4Comment, Item4SettingName, Item4SettingTitle "
    End If

    If iNoAssignments = 5 Then
        sSQL = sSQL & " ,Item5TalkNo, Item5StudentName, Item5AssistantName, Item5Theme, Item5Source, Item5CounselPoint, Item5CounselSubPoint1, Item5CounselSubPoint2, Item5CounselSubPoint3, Item5CounselSubPoint4, Item5CounselSubPoint5, Item5Comment, Item5SettingName, Item5SettingTitle "
    End If
    
    sSQL = sSQL & " FROM tblTMSPrintWorkSheet"
    
    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .ReportSQL = sSQL
        
    .PageFormat = cmsPortrait
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 10
    .RightMargin = 10
    .ClientName = AppName
    .ShowPageNumber = False
    .HideWordWhileBuilding = True
    .NumberFormat = ""
    .ShowProgress = True
    .ReportTitle = ""
    .SaveDoc = True
    .DocPath = gsDocsDirectory & "\" & "Student Talks" & _
                                 " Worksheet for Week " & Replace(frmTMSPrinting.cmbStartDate.text, "/", "-") & " (" & _
                                Replace(Replace(Now, ":", "-"), "/", "-") & ")"
    
    
    'use this if using a word template with pre-built table in it
'    .WordTemplatePath_Adv = App.Path & "\School.dot"
    
    .DocumentType = cmsPagePerDBRow
    
    ''
    'Because DocumentType = cmsPagePerDBRow, must supply cell-specific details....
    ''
    
    .TableCols_Adv = 10
    
    '1st row of table contains School No, Heading and date
    ' then each group of [iTableRowsPerTalkNo] rows contains info for each student talk
    ' calculate total number of Word table rows
    .TableRows_Adv = 1 + (iTableRowsPerTalkNo * iNoAssignments)
    
    
    'put in fixed text, ie field labels. Not required if using a Word template
    .AddFixedCellText_Adv 1, 1, "School: ", True
    .AddFixedCellText_Adv 1, 3, "Student Talks Worksheet"
    
    .AddFixedCellText_Adv 3, 1, "Asst: "
    .AddFixedCellText_Adv 4, 1, "Counsel: "
    .AddFixedCellText_Adv 4, 4, "Last Comment: "
    .AddFixedCellText_Adv 4, 9, "Timing: " & vbCr & "Ex: " & vbCr & "Next Pt: "
    
    If iNoAssignments >= 2 Then
        .AddFixedCellText_Adv 11, 1, "Asst: "
        .AddFixedCellText_Adv 12, 1, "Counsel: "
        .AddFixedCellText_Adv 12, 4, "Last Comment: "
        .AddFixedCellText_Adv 12, 9, "Timing: " & vbCr & "Ex: " & vbCr & "Next Pt: "
    End If
    
    
    If iNoAssignments >= 3 Then
        .AddFixedCellText_Adv 19, 1, "Asst: "
        .AddFixedCellText_Adv 20, 1, "Counsel: "
        .AddFixedCellText_Adv 20, 4, "Last Comment: "
        .AddFixedCellText_Adv 20, 9, "Timing: " & vbCr & "Ex: " & vbCr & "Next Pt: "
    End If
    
    If iNoAssignments >= 4 Then
        .AddFixedCellText_Adv 27, 1, "Asst: "
        .AddFixedCellText_Adv 28, 1, "Counsel: "
        .AddFixedCellText_Adv 28, 4, "Last Comment: "
        .AddFixedCellText_Adv 28, 9, "Timing: " & vbCr & "Ex: " & vbCr & "Next Pt: "
    End If
    
    If iNoAssignments = 5 Then
        .AddFixedCellText_Adv 35, 1, "Asst: "
        .AddFixedCellText_Adv 36, 1, "Counsel: "
        .AddFixedCellText_Adv 36, 4, "Last Comment: "
        .AddFixedCellText_Adv 36, 9, "Timing: " & vbCr & "Ex: " & vbCr & "Next Pt: "
    End If
    
    'Mapping of DB fields to table cells.
    ' listed in order of fields in ReportSQL.
    
    .AddFieldMapping_Adv 1, 10, True    'assignment date
    .AddFieldMapping_Adv 1, 1           'school no

    .AddFieldMapping_Adv 2, 1           'Talk No
    .AddFieldMapping_Adv 2, 3           'student name
    .AddFieldMapping_Adv 3, 3           'asst name
    .AddFieldMapping_Adv 2, 5           'theme
    .AddFieldMapping_Adv 2, 8           'source
    .AddFieldMapping_Adv 4, 3           'counsel pt
    .AddFieldMapping_Adv 5, 1          'counsel sub-point1
    .AddFieldMapping_Adv 6, 1          'counsel sub-point2
    .AddFieldMapping_Adv 7, 1          'counsel sub-point3
    .AddFieldMapping_Adv 8, 1          'counsel sub-point4
    .AddFieldMapping_Adv 9, 1          'counsel sub-point5
    .AddFieldMapping_Adv 4, 5          'comment
    .AddFieldMapping_Adv 3, 5           'setting
    .AddFieldMapping_Adv 3, 4           'settingTitle

    
    If iNoAssignments >= 2 Then
        .AddFieldMapping_Adv 10, 1           'Talk No
        .AddFieldMapping_Adv 10, 3           'student name
        .AddFieldMapping_Adv 11, 3           'asst name
        .AddFieldMapping_Adv 10, 5           'theme
        .AddFieldMapping_Adv 10, 8           'source
        .AddFieldMapping_Adv 12, 3           'counsel pt
        .AddFieldMapping_Adv 13, 1          'counsel sub-point1
        .AddFieldMapping_Adv 14, 1          'counsel sub-point2
        .AddFieldMapping_Adv 15, 1          'counsel sub-point3
        .AddFieldMapping_Adv 16, 1          'counsel sub-point4
        .AddFieldMapping_Adv 17, 1          'counsel sub-point5
        .AddFieldMapping_Adv 12, 5          'comment
        .AddFieldMapping_Adv 11, 5           'setting
        .AddFieldMapping_Adv 11, 4           'settingTitle
    End If

    If iNoAssignments >= 3 Then
        .AddFieldMapping_Adv 18, 1           'Talk No
        .AddFieldMapping_Adv 18, 3           'student name
        .AddFieldMapping_Adv 19, 3           'asst name
        .AddFieldMapping_Adv 18, 5           'theme
        .AddFieldMapping_Adv 18, 8           'source
        .AddFieldMapping_Adv 20, 3           'counsel pt
        .AddFieldMapping_Adv 21, 1          'counsel sub-point1
        .AddFieldMapping_Adv 22, 1          'counsel sub-point2
        .AddFieldMapping_Adv 23, 1          'counsel sub-point3
        .AddFieldMapping_Adv 24, 1          'counsel sub-point4
        .AddFieldMapping_Adv 25, 1          'counsel sub-point5
        .AddFieldMapping_Adv 20, 5          'comment
        .AddFieldMapping_Adv 19, 5           'setting
        .AddFieldMapping_Adv 19, 4           'settingTitle
    End If
    
    If iNoAssignments >= 4 Then
        .AddFieldMapping_Adv 26, 1           'Talk No
        .AddFieldMapping_Adv 26, 3           'student name
        .AddFieldMapping_Adv 27, 3           'asst name
        .AddFieldMapping_Adv 26, 5           'theme
        .AddFieldMapping_Adv 26, 8           'source
        .AddFieldMapping_Adv 28, 3           'counsel pt
        .AddFieldMapping_Adv 29, 1          'counsel sub-point1
        .AddFieldMapping_Adv 30, 1          'counsel sub-point2
        .AddFieldMapping_Adv 31, 1          'counsel sub-point3
        .AddFieldMapping_Adv 32, 1          'counsel sub-point4
        .AddFieldMapping_Adv 33, 1          'counsel sub-point5
        .AddFieldMapping_Adv 28, 5          'comment
        .AddFieldMapping_Adv 27, 5           'setting
        .AddFieldMapping_Adv 27, 4           'settingTitle
    End If
    
    If iNoAssignments = 5 Then
        .AddFieldMapping_Adv 34, 1           'Talk No
        .AddFieldMapping_Adv 34, 3           'student name
        .AddFieldMapping_Adv 35, 3           'asst name
        .AddFieldMapping_Adv 34, 5           'theme
        .AddFieldMapping_Adv 34, 8           'source
        .AddFieldMapping_Adv 36, 3           'counsel pt
        .AddFieldMapping_Adv 37, 1          'counsel sub-point1
        .AddFieldMapping_Adv 38, 1          'counsel sub-point2
        .AddFieldMapping_Adv 39, 1          'counsel sub-point3
        .AddFieldMapping_Adv 40, 1          'counsel sub-point4
        .AddFieldMapping_Adv 41, 1          'counsel sub-point5
        .AddFieldMapping_Adv 36, 5          'comment
        .AddFieldMapping_Adv 35, 5           'setting
        .AddFieldMapping_Adv 35, 4           'settingTitle
    End If

    '
    '
    'these settings would be ignorred if using word template with pre-built table
    '
    'apply font size to whole table
    .AddCellAttributes_Adv 1, 1, .TableRows_Adv, 10, cmsleftTop, "Times New Roman", 10, cmsOptionfalse, _
                                                         cmsOptionfalse, , , , True
                                                        
    'apply bold, centred formatting to the "TMS Worksheet" cells
    .AddCellAttributes_Adv 1, 3, 1, 8, cmsCentreTop, , 12, cmsOptionTrue, _
                                                        cmsOptionTrue
                           
    'bold for School No
    .AddCellAttributes_Adv 1, 1, , , cmsleftTop, , 12, cmsOptionTrue, cmsOptionTrue
        
    
    'bold, right-aligned assignment date
    .AddCellAttributes_Adv 1, 9, 1, 10, cmsRightTop, , 12, cmsOptionTrue, cmsOptionTrue
           
    'Make various cells bold
    .AddCellAttributes_Adv 2, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    .AddCellAttributes_Adv 3, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    .AddCellAttributes_Adv 3, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    .AddCellAttributes_Adv 4, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    .AddCellAttributes_Adv 4, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    .AddCellAttributes_Adv 4, 9, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    
    If iNoAssignments >= 2 Then
        .AddCellAttributes_Adv 10, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 11, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 11, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 12, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 12, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 12, 9, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    End If

    If iNoAssignments >= 3 Then
        .AddCellAttributes_Adv 18, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 19, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 19, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 20, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 20, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 20, 9, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    End If
    
    If iNoAssignments >= 4 Then
        .AddCellAttributes_Adv 26, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 27, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 27, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 28, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 28, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 28, 9, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    End If

    If iNoAssignments = 5 Then
        .AddCellAttributes_Adv 34, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 35, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 35, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 36, 1, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 36, 4, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
        .AddCellAttributes_Adv 36, 9, , , cmsleftTop, , 9, cmsOptionTrue, cmsOptionfalse
    End If
    
    'Darker lines between assignments
    .AddCellAttributes_Adv 1, 1, 1, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    .AddCellAttributes_Adv 9, 1, 9, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    
    If iNoAssignments >= 2 Then
        .AddCellAttributes_Adv 17, 1, 17, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    End If

    If iNoAssignments >= 3 Then
        .AddCellAttributes_Adv 25, 1, 25, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    End If
    
    If iNoAssignments >= 4 Then
        .AddCellAttributes_Adv 33, 1, 33, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    End If

    If iNoAssignments = 5 Then
        .AddCellAttributes_Adv 41, 1, 41, 10, , , , , , , cmsBottomBorder, cmsLineWidth150pt
    End If
    


    '
    'Merge cells. Should be the last formatting done.
    ' This should be done from right to left and bottom to top
    '
    
    bMergeInit = True 'init
    
    If iNoAssignments = 5 Then
        .AddMergeCellRange_Adv 41, 4, 41, 10, bMergeInit
        bMergeInit = False
        .AddMergeCellRange_Adv 41, 1, 41, 3
        .AddMergeCellRange_Adv 40, 4, 40, 10
        .AddMergeCellRange_Adv 40, 1, 40, 3
        .AddMergeCellRange_Adv 39, 4, 39, 10
        .AddMergeCellRange_Adv 39, 1, 39, 3
        .AddMergeCellRange_Adv 38, 4, 38, 10
        .AddMergeCellRange_Adv 38, 1, 38, 3
        .AddMergeCellRange_Adv 37, 4, 37, 10
        .AddMergeCellRange_Adv 37, 1, 37, 3
        .AddMergeCellRange_Adv 36, 5, 36, 8
        .AddMergeCellRange_Adv 36, 2, 36, 3
        .AddMergeCellRange_Adv 35, 4, 35, 10
        .AddMergeCellRange_Adv 35, 2, 35, 3
        .AddMergeCellRange_Adv 34, 8, 34, 10
        .AddMergeCellRange_Adv 34, 5, 34, 7
        .AddMergeCellRange_Adv 34, 2, 34, 4
    End If

    If iNoAssignments >= 4 Then
        .AddMergeCellRange_Adv 33, 4, 33, 10, bMergeInit
        bMergeInit = False
        .AddMergeCellRange_Adv 33, 1, 33, 3
        .AddMergeCellRange_Adv 32, 4, 32, 10
        .AddMergeCellRange_Adv 32, 1, 32, 3
        .AddMergeCellRange_Adv 31, 4, 31, 10
        .AddMergeCellRange_Adv 31, 1, 31, 3
        .AddMergeCellRange_Adv 30, 4, 30, 10
        .AddMergeCellRange_Adv 30, 1, 30, 3
        .AddMergeCellRange_Adv 29, 4, 29, 10
        .AddMergeCellRange_Adv 29, 1, 29, 3
        .AddMergeCellRange_Adv 28, 5, 28, 8
        .AddMergeCellRange_Adv 28, 2, 28, 3
        .AddMergeCellRange_Adv 27, 4, 27, 10
        .AddMergeCellRange_Adv 27, 2, 27, 3
        .AddMergeCellRange_Adv 26, 8, 26, 10
        .AddMergeCellRange_Adv 26, 5, 26, 7
        .AddMergeCellRange_Adv 26, 2, 26, 4
    End If

    If iNoAssignments >= 3 Then
        .AddMergeCellRange_Adv 25, 4, 25, 10, bMergeInit
        bMergeInit = False
        .AddMergeCellRange_Adv 25, 1, 25, 3
        .AddMergeCellRange_Adv 24, 4, 24, 10
        .AddMergeCellRange_Adv 24, 1, 24, 3
        .AddMergeCellRange_Adv 23, 4, 23, 10
        .AddMergeCellRange_Adv 23, 1, 23, 3
        .AddMergeCellRange_Adv 22, 4, 22, 10
        .AddMergeCellRange_Adv 22, 1, 22, 3
        .AddMergeCellRange_Adv 21, 4, 21, 10
        .AddMergeCellRange_Adv 21, 1, 21, 3
        .AddMergeCellRange_Adv 20, 5, 20, 8
        .AddMergeCellRange_Adv 20, 2, 20, 3
        .AddMergeCellRange_Adv 19, 4, 19, 10
        .AddMergeCellRange_Adv 19, 2, 19, 3
        .AddMergeCellRange_Adv 18, 8, 18, 10
        .AddMergeCellRange_Adv 18, 5, 18, 7
        .AddMergeCellRange_Adv 18, 2, 18, 4
    End If
    
    If iNoAssignments >= 2 Then
        .AddMergeCellRange_Adv 17, 4, 17, 10, bMergeInit
        bMergeInit = False
        .AddMergeCellRange_Adv 17, 1, 17, 3
        .AddMergeCellRange_Adv 16, 4, 16, 10
        .AddMergeCellRange_Adv 16, 1, 16, 3
        .AddMergeCellRange_Adv 15, 4, 15, 10
        .AddMergeCellRange_Adv 15, 1, 15, 3
        .AddMergeCellRange_Adv 14, 4, 14, 10
        .AddMergeCellRange_Adv 14, 1, 14, 3
        .AddMergeCellRange_Adv 13, 4, 13, 10
        .AddMergeCellRange_Adv 13, 1, 13, 3
        .AddMergeCellRange_Adv 12, 5, 12, 8
        .AddMergeCellRange_Adv 12, 2, 12, 3
        .AddMergeCellRange_Adv 11, 4, 11, 10
        .AddMergeCellRange_Adv 11, 2, 11, 3
        .AddMergeCellRange_Adv 10, 8, 10, 10
        .AddMergeCellRange_Adv 10, 5, 10, 7
        .AddMergeCellRange_Adv 10, 2, 10, 4
    End If
    
    .AddMergeCellRange_Adv 9, 4, 9, 10, bMergeInit
    bMergeInit = False
    .AddMergeCellRange_Adv 9, 1, 9, 3
    .AddMergeCellRange_Adv 8, 4, 8, 10
    .AddMergeCellRange_Adv 8, 1, 8, 3
    .AddMergeCellRange_Adv 7, 4, 7, 10
    .AddMergeCellRange_Adv 7, 1, 7, 3
    .AddMergeCellRange_Adv 6, 4, 6, 10
    .AddMergeCellRange_Adv 6, 1, 6, 3
    .AddMergeCellRange_Adv 5, 4, 5, 10
    .AddMergeCellRange_Adv 5, 1, 5, 3
    .AddMergeCellRange_Adv 4, 5, 4, 8
    .AddMergeCellRange_Adv 4, 2, 4, 3
    .AddMergeCellRange_Adv 3, 4, 3, 10
    .AddMergeCellRange_Adv 3, 2, 3, 3
    .AddMergeCellRange_Adv 2, 8, 2, 10
    .AddMergeCellRange_Adv 2, 5, 2, 7
    .AddMergeCellRange_Adv 2, 2, 2, 4
    .AddMergeCellRange_Adv 1, 8, 1, 10
    .AddMergeCellRange_Adv 1, 3, 1, 7
    .AddMergeCellRange_Adv 1, 1, 1, 2
    
    
    .GenerateReport
       
    End With
    
    SwitchOnDAO
    
    Set reporter = Nothing
    
    Screen.MousePointer = vbNormal
    
    If GlobalParms.GetValue("TMS_PrintSQListWithWorksheet", "TrueFalse", True) = True Then
        If MsgBox("Print Speech Quality update list?", vbYesNo + vbQuestion, AppName) = vbYes Then
            PrintBEBookUpdateList frmTMSPrinting.cmbStartDate.text
        End If
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Public Sub PrintBEBookUpdateList(AssignmentDate As String)
On Error GoTo ErrorTrap

    Select Case GlobalParms.GetValue("TMS_PrintSQListWithWorksheet_Type", "AlphaVal", "PerStudent")
    Case "PerStudent"
        PrintBEBookUpdateList_DoIt Get_PrintBEBookUpdateList_SQL(AssignmentDate, "'1'"), AssignmentDate, False, True
        PrintBEBookUpdateList_DoIt Get_PrintBEBookUpdateList_SQL(AssignmentDate, "'2'"), AssignmentDate, False, True
        PrintBEBookUpdateList_DoIt Get_PrintBEBookUpdateList_SQL(AssignmentDate, "'3'"), AssignmentDate, False, True
    Case "OneSheet"
        PrintBEBookUpdateList_DoIt Get_PrintBEBookUpdateList_SQL(AssignmentDate, "'1','2','3'"), AssignmentDate, False, False
    End Select


    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Public Function Get_PrintBEBookUpdateList_SQL(AssignmentDate As String, TalkNo As String) As String
On Error GoTo ErrorTrap

    Dim sSQL As String

    sSQL = "SELECT b.LastName & ', ' & b.FirstName & ' ' & b.MiddleName AS Name,  " & _
            "AssignmentDate, CounselPoint, CounselPointAssignedDate, CounselPointCompletedDate " & _
            "FROM tblTMSSchedule a INNER JOIN tblNameAddress b ON a.PersonID = b.ID " & _
            "Where a.CounselPoint > 0 " & _
            "AND a.DiscussedWithStudent = FALSE " & _
            "AND a.TalkDefaulted = FALSE " & _
            "AND a.AssignmentDate > #01/01/2009# " & _
            "AND a.AssignmentDate < " & GetDateStringForSQLWhere(AssignmentDate) & " " & _
            "AND a.TalkNo IN('1','2','3') " & _
            "And b.Active = TRUE " & _
            "AND b.ID IN " & _
            "(SELECT c.PersonID " & _
            " FROM tblTMSSchedule c " & _
            " Where c.AssignmentDate = " & GetDateStringForSQLWhere(AssignmentDate) & _
            " AND c.TalkNo IN (" & TalkNo & ")) " & _
            "ORDER BY 1, 2 "



    Get_PrintBEBookUpdateList_SQL = sSQL

    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Sub PrintBEBookUpdateList_DoIt(sSQL As String, AssignmentDate As String, bAsk As Boolean, bPerStudent As Boolean)

On Error GoTo ErrorTrap

Dim rs As Recordset
Dim bNowt As Boolean
Dim reporter As MSWordReportingTool2.RptTool
Dim sName As String
    
           
    Set rs = CMSDB.OpenRecordset(sSQL, dbOpenForwardOnly)
    
    bNowt = rs.BOF
    
    If Not bNowt Then
        sName = rs!Name
    End If
    
    rs.Close
    Set rs = Nothing
    
    If bNowt Then Exit Sub
    
    If bAsk Then
        If MsgBox("Print Speech Quality update list?", vbYesNo + vbQuestion, AppName) = vbNo Then
            Exit Sub
        End If
    End If
    
    '
    'Start the print using Word.....
    '
    
    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .SaveDoc = True
    
    If bPerStudent Then
        .DocPath = gsDocsDirectory & "\" & "TMS Speech Quality Update List for " & _
                    MakeStringValidForFileName(sName) & " - Week " & _
                      Format(CDate(AssignmentDate), "dd-mm-yyyy") & " (" & _
                                    Replace(Replace(Now, ":", "-"), "/", "-") & ")"
    Else
        .DocPath = gsDocsDirectory & "\" & "TMS Speech Quality Update List for Week " & _
                      Format(CDate(AssignmentDate), "dd-mm-yyyy") & " (" & _
                                    Replace(Replace(Now, ":", "-"), "/", "-") & ")"
    End If

    
    .ReportSQL = sSQL

    .ReportTitle = "TMS Speech Quality Update List"
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 10
    .RightMargin = 10
    .ReportFooterFontName = "Arial"
    .ReportFooterFontSize = 8
    .ReportTitleFontName = "Times New Roman"
    .ReportTitleFontSize = 15
    .ApplyTableFormatting = True
    .ClientName = AppName
    .AdditionalReportHeading = "Assignment Week: " & AssignmentDate
    .GroupingColumn = 1
    .HideWordWhileBuilding = True
    
    .AddTableColumnAttribute "Name", 50, , , , , 12, 10, True, True, , , True
    .AddTableColumnAttribute "Talk Date", 30, , , , , 12, 10, True, True
    .AddTableColumnAttribute "SQ", 10, , , , , 12, 10, True, True
    .AddTableColumnAttribute "Assigned", 30, , , , , 12, 10, True, True
    .AddTableColumnAttribute "Completed", 30, , , , , 12, 10, True, True
    
    .PageFormat = cmsPortrait
    
    .GenerateReport

    End With
    
    SwitchOnDAO
    
    Set reporter = Nothing
    Screen.MousePointer = vbNormal
    
        
        



    Exit Sub
ErrorTrap:
    EndProgram

End Sub



Public Function PrintWorkTMSSheets(Optional bDraft As Boolean = False) As Boolean
Dim WorkSheetTopMargin As Single
Dim WorkSheetBottomMargin As Single
Dim WorkSheetLeftMargin As Single
Dim WorkSheetRightMargin As Single
Dim WorkSheetType As Integer

On Error GoTo ErrorTrap

    Screen.MousePointer = vbHourglass
    '
    'Build the table used for the report....
    '
    If Not BuildPrintTable Then
        PrintWorkTMSSheets = False
        Screen.MousePointer = vbNormal
        Exit Function
    End If
    
    Select Case PrintUsingWord(False)
    Case cmsUseWord
        Select Case NewMtgArrangementStarted(frmTMSPrinting.cmbStartDate.text)
        Case CLM2016
            PrintWorkTMSSheetsUsingWord_2016
        Case TMS2009
            PrintWorkTMSSheetsUsingWord_2009
        Case Else
            PrintWorkTMSSheetsUsingWord
        End Select
        PrintWorkTMSSheets = True
        Exit Function
    Case cmsDontPrint
        Screen.MousePointer = vbNormal
        PrintWorkTMSSheets = True
        Exit Function
    End Select
    
    
    '
    'Arrange page margins before we close the db connection
    '
    WorkSheetTopMargin = 566.929 * (GlobalParms.GetValue("A4TopMargin", "NumFloat"))
    WorkSheetBottomMargin = 566.929 * (GlobalParms.GetValue("A4BottomMargin", "NumFloat"))
    WorkSheetLeftMargin = 566.929 * (GlobalParms.GetValue("A4LeftMargin", "NumFloat"))
    WorkSheetRightMargin = 566.929 * (GlobalParms.GetValue("A4RightMargin", "NumFloat"))

    '
    'Which type of worksheet shall we print?
    '
    WorkSheetType = GlobalParms.GetValue("TMSWorksheetToUse", "NumVal")
    
    DestroyGlobalObjects
    CMSDB.Close
    
    '
    'GENERAL ADO WARNING...
    ' If we refer to DataReport prior to 'Showing' it, thus opening new ADODB connection
    ' while DAO connection still open, we get funny results.. eg missing fields on report.
    '
        
    If NewMtgArrangementStarted(frmTMSPrinting.cmbStartDate.text) Then
        Select Case WorkSheetType
        Case 1:
            'Display room for counsel on #1 and BH
            TMSPrintWorkSheet_2009.TopMargin = WorkSheetTopMargin '<----- At this point, TMSPrintWorkSheet.Initialize runs.
            TMSPrintWorkSheet_2009.BottomMargin = WorkSheetBottomMargin
            TMSPrintWorkSheet_2009.LeftMargin = WorkSheetLeftMargin
            TMSPrintWorkSheet_2009.RightMargin = WorkSheetRightMargin
            Screen.MousePointer = vbNormal
            
            TMSPrintWorkSheet_2009.Show vbModal
        Case 2:
            'No room for counsel on #1 and BH
            TMSPrintWorkSheet2_2009.TopMargin = WorkSheetTopMargin '<----- At this point, TMSPrintWorkSheet.Initialize runs.
            TMSPrintWorkSheet2_2009.BottomMargin = WorkSheetBottomMargin
            TMSPrintWorkSheet2_2009.LeftMargin = WorkSheetLeftMargin
            TMSPrintWorkSheet2_2009.RightMargin = WorkSheetRightMargin
            Screen.MousePointer = vbNormal
            
            TMSPrintWorkSheet2_2009.Show vbModal
        End Select
    Else
        Select Case WorkSheetType
        Case 1:
            'Display room for counsel on #1 and BH
            TMSPrintWorkSheet.TopMargin = WorkSheetTopMargin '<----- At this point, TMSPrintWorkSheet.Initialize runs.
            TMSPrintWorkSheet.BottomMargin = WorkSheetBottomMargin
            TMSPrintWorkSheet.LeftMargin = WorkSheetLeftMargin
            TMSPrintWorkSheet.RightMargin = WorkSheetRightMargin
            Screen.MousePointer = vbNormal
            
            TMSPrintWorkSheet.Show vbModal
        Case 2:
            'No room for counsel on #1 and BH
            TMSPrintWorkSheet2.TopMargin = WorkSheetTopMargin '<----- At this point, TMSPrintWorkSheet.Initialize runs.
            TMSPrintWorkSheet2.BottomMargin = WorkSheetBottomMargin
            TMSPrintWorkSheet2.LeftMargin = WorkSheetLeftMargin
            TMSPrintWorkSheet2.RightMargin = WorkSheetRightMargin
            Screen.MousePointer = vbNormal
            
            TMSPrintWorkSheet2.Show vbModal
        End Select
    End If
    
    '
    'rstTMSQuery  and the Global Objects are destroyed when report is generated
    ' due to DB Disconnect. So, instantiate them once more....
    '
'    frmTMSPrinting.SetUpTMSSlipPrintMainRecSet
    InstantiateGlobalObjects

    PrintWorkTMSSheets = True

    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Function PrintStudentDetails() As Boolean
Dim WorkSheetTopMargin As Single
Dim WorkSheetBottomMargin As Single
Dim WorkSheetLeftMargin As Single
Dim WorkSheetRightMargin As Single
Dim b2009 As Boolean

On Error GoTo ErrorTrap

    Screen.MousePointer = vbHourglass
    
    b2009 = GlobalParms.GetValue("TMSPrintStudentDtlsFor2009Format", "TrueFalse")
    
    '
    'Build the table used for the report....
    '
    If Not BuildStudentDetailsPrintTable Then
        PrintStudentDetails = False
        Screen.MousePointer = vbNormal
        Exit Function
    End If
    
    Select Case PrintUsingWord(True)
    Case cmsUseWord
        If b2009 Then
            TMSPrintStudentDetailsWithWord_2009
        Else
            TMSPrintStudentDetailsWithWord
        End If
    Case cmsUseMSDatareport
    
        '
        'Arrange page margins before we close the db connection
        '
        WorkSheetTopMargin = 566.929 * (GlobalParms.GetValue("A4TopMargin", "NumFloat"))
        WorkSheetBottomMargin = 566.929 * (GlobalParms.GetValue("A4BottomMargin", "NumFloat"))
        WorkSheetLeftMargin = 566.929 * (GlobalParms.GetValue("A4LeftMargin", "NumFloat"))
        WorkSheetRightMargin = 566.929 * (GlobalParms.GetValue("A4RightMargin", "NumFloat"))
    
    
        DestroyGlobalObjects
        CMSDB.Close
        
        '
        'GENERAL ADO WARNING...
        ' If we refer to DataReport prior to 'Showing' it, thus opening new ADODB connection
        ' while DAO connection still open, we get funny results.. eg missing fields on report.
        '
        If b2009 Then
            TMSStudentDetails_2009.TopMargin = WorkSheetTopMargin '<----- At this point, TMSStudentDetails.Initialize runs.
            TMSStudentDetails_2009.BottomMargin = WorkSheetBottomMargin
            TMSStudentDetails_2009.LeftMargin = WorkSheetLeftMargin
            TMSStudentDetails_2009.RightMargin = WorkSheetRightMargin
            Screen.MousePointer = vbNormal
            TMSStudentDetails_2009.Show vbModal
        Else
            TMSStudentDetails.TopMargin = WorkSheetTopMargin '<----- At this point, TMSStudentDetails.Initialize runs.
            TMSStudentDetails.BottomMargin = WorkSheetBottomMargin
            TMSStudentDetails.LeftMargin = WorkSheetLeftMargin
            TMSStudentDetails.RightMargin = WorkSheetRightMargin
            Screen.MousePointer = vbNormal
            TMSStudentDetails.Show vbModal
        End If
        
        InstantiateGlobalObjects
    End Select

    PrintStudentDetails = True
    Screen.MousePointer = vbNormal

    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Function PrintCounselForm() As Boolean

On Error GoTo ErrorTrap

Dim i As Long, PersonID As Long

    Screen.MousePointer = vbHourglass
    
    
    With frmTMSPrinting.lstStudents
    
    For i = 0 To .ListCount - 1
    
    '
    'Build the table used for the report....
    '
    
        If .Selected(i) Then
        
            PersonID = .ItemData(i)
    
            If Not BuildCounselFormPrintTable(PersonID) Then
                PrintCounselForm = False
                Screen.MousePointer = vbNormal
                Exit Function
            End If
        
            Select Case PrintUsingWord(False)
            Case cmsUseWord
                TMSPrintCounselPointsWithWord (PersonID)
            Case Else
                ShowMessage "Could not open Word", 1000, frmTMSPrinting
                PrintCounselForm = False
                Exit Function
            End Select
            
        End If
        
    Next i
    
    End With

    PrintCounselForm = True
    Screen.MousePointer = vbNormal

    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Private Function BuildPrintTable() As Boolean
Dim rstPrtCounsel As Recordset
Dim rstPrtSchedule As Recordset
Dim rstPrtItems As Recordset
Dim rstPrtWorkSheet As Recordset
Dim rstPrtComment As Recordset
Dim rstSettings As Recordset
Dim TheAssignmentDate As Date
Dim ActiveSchoolsSQL As String
Dim ScheduleSQL As String
Dim ItemsSQL As String
Dim CounselSQL As String
Dim CommentSQL As String
Dim PrevSchool As Long
Dim TheComment As String
Dim TheCounsel As String
Dim AssistantName As String
Dim SubPt1 As String, SubPt2 As String, SubPt3 As String, SubPt4 As String, SubPt5 As String
Dim IsMovedOralReview As Boolean, NextWeekIsMovedOralReview As Boolean
Dim LastWeekWasAssemblyWeek As Boolean, IsCOVisitWeek As Boolean, LastWeekWasCOVisit As Boolean
Dim IsOralReview As Boolean
Dim TheSettingAndDesc As String, sSettingTitle As String
Dim SongNo As Long, sSongNo As String, sStudentName As String
Dim lSong2 As Long, sSong2 As String, sSrvMtgSQL As String
Dim sSrvMtgPerson As String, bNewArr As MidweekMtgVersion
Dim lPers As Long, sItemName As String
Dim i As Long
Dim lMidweekMtgDay As Long
Dim iItemIX As Integer


On Error GoTo ErrorTrap

    DelAllRows "tblTMSPrintWorkSheet"
    
    With frmTMSPrinting
    '
    'Build chunk of SQL to select appropriate schools from tblSchedule
    '
    ActiveSchoolsSQL = " AND SchoolNo IN ("
    For i = 0 To 2
        If .chkSchool(CInt(i)).value = vbChecked Then
            ActiveSchoolsSQL = ActiveSchoolsSQL & CStr(i + 1)
            ActiveSchoolsSQL = ActiveSchoolsSQL & ", "
        End If
    Next i
    
    ActiveSchoolsSQL = Left(ActiveSchoolsSQL, Len(ActiveSchoolsSQL) - 2)
    ActiveSchoolsSQL = ActiveSchoolsSQL & ") "
    
    
    TheAssignmentDate = CDate(.cmbStartDate.text)
    
    bNewArr = NewMtgArrangementStarted(CStr(TheAssignmentDate))
    
    End With
    
    '
    'Get 1st bro from Service Meeting
    '
    
    lPers = GetFirstServMtgBro(TheAssignmentDate)
    
    If lPers > 0 Then
        
        If GlobalParms.GetValue("TMS_InclSrvMtgItemNameInWorkshtPrt", "TrueFalse", False) = True Then
            sItemName = GetFirstServMtgItemName(TheAssignmentDate)
        Else
            sItemName = ""
        End If
            
        
        
        If sItemName <> "" Then
            sItemName = ": " & sItemName
        End If
        
        sSrvMtgPerson = "Invite " & CongregationMember.NameWithMiddleInitial(lPers) & _
                            " to the platform to start the Service Meeting" & _
                            IIf(sItemName = "", "", sItemName)
    Else
        sSrvMtgPerson = "Invite                              to the platform."
    End If
    
    '
    'Driving Recset to get all assignments for given date and schools
    '
    ScheduleSQL = "SELECT * FROM tblTMSSchedule " & _
                " WHERE AssignmentDate = #" & Format(TheAssignmentDate, "mm/dd/yyyy") & "# " & _
                ActiveSchoolsSQL & _
                " AND TalkNo IN ('BR','IC','RV','BS','O') " & _
                " AND PersonID > 0 " & _
                " AND ItemsSeqNum > 0 " & _
                 " ORDER BY SchoolNo, TalkSeqNum, ItemsSeqNum "

    Set rstPrtSchedule = CMSDB.OpenRecordset(ScheduleSQL, dbOpenDynaset)
    
    
    If rstPrtSchedule.BOF Then
        BuildPrintTable = False
        Exit Function
    Else
        PrevSchool = rstPrtSchedule!SchoolNo
    End If
    
    
    '
    'All items for this and next week's Assignment Date. (Next/Previous week's is provided
    ' in case of moved Oral Review scenario
    '
    ItemsSQL = "SELECT * FROM tblTMSItems " & _
                "WHERE AssignmentDate BETWEEN #" & Format(DateAdd("ww", -1, TheAssignmentDate), "mm/dd/yyyy") & "# " & _
                "AND #" & Format(DateAdd("ww", 1, TheAssignmentDate), "mm/dd/yyyy") & "# "
    
    Set rstPrtItems = CMSDB.OpenRecordset(ItemsSQL, dbOpenDynaset)
    
    If rstPrtItems.BOF Then
        BuildPrintTable = False
        Exit Function
    End If
    
    
    Set rstSettings = CMSDB.OpenRecordset("tblTMSSettings", dbOpenDynaset)
    
    Set rstPrtWorkSheet = CMSDB.OpenRecordset("Select * FROM tblTMSPrintWorkSheet " & _
                                               " ORDER By SchoolNo", dbOpenDynaset)
                                               
    lMidweekMtgDay = GlobalParms.GetValue("MidWeekMeetingDay", "NumVal")
    
    iItemIX = 0
        
    With rstPrtWorkSheet
'    rstPrtSchedule.MoveLast
    Do Until rstPrtSchedule.EOF
'    Do Until rstPrtSchedule.BOF
        .AddNew
        !SchoolNo = rstPrtSchedule!SchoolNo
        !AssignmentDate2 = CStr(TheAssignmentDate) & " (" & _
                            GetDateOfGivenDay(TheAssignmentDate, lMidweekMtgDay, True) & ")"
        !AssignmentDate = CStr(TheAssignmentDate)
        .Update
'        Do While Not rstPrtSchedule.BOF
        Do While Not rstPrtSchedule.EOF
        
            iItemIX = iItemIX + 1
            
            SubPt1 = ""
            SubPt2 = ""
            SubPt3 = ""
            SubPt4 = ""
            SubPt5 = ""
            TheCounsel = ""
            
            'Find Item data for scheduled assignment...
            CheckForCOOrAssembly IsMovedOralReview, _
                                 NextWeekIsMovedOralReview, _
                                 LastWeekWasAssemblyWeek, _
                                 IsCOVisitWeek, _
                                 LastWeekWasCOVisit, _
                                 rstPrtSchedule!AssignmentDate
                                 
            CheckForOralReview IsOralReview, rstPrtSchedule!AssignmentDate
            
            Select Case bNewArr
            Case CLM2016
                rstPrtItems.FindFirst "AssignmentDate = #" & Format(TheAssignmentDate, "mm/dd/yyyy") & "# " & _
                                        "AND TalkNo = '" & rstPrtSchedule!TalkNo & "' " & _
                                        " AND ItemsSeqNum = " & rstPrtSchedule!ItemsSeqNum
            Case TMS2009
                If NextWeekIsMovedOralReview And IsCOVisitWeek And _
                    (rstPrtSchedule!TalkNo = "1" Or rstPrtSchedule!TalkNo = "2" Or rstPrtSchedule!TalkNo = "3") Then
                    rstPrtItems.FindFirst "AssignmentDate = #" & Format(TheAssignmentDate + 7, "mm/dd/yyyy") & "# " & _
                                            "AND TalkNo = '" & rstPrtSchedule!TalkNo & "' "
                Else
                    rstPrtItems.FindFirst "AssignmentDate = #" & Format(TheAssignmentDate, "mm/dd/yyyy") & "# " & _
                                            "AND TalkNo = '" & rstPrtSchedule!TalkNo & "'"
                End If
                
            Case Else
                If LastWeekWasAssemblyWeek And IsMovedOralReview Then
                    rstPrtItems.FindFirst "AssignmentDate = #" & Format(TheAssignmentDate - 7, "mm/dd/yyyy") & "# " & _
                                            "AND TalkNo = '" & rstPrtSchedule!TalkNo & "'"
                ElseIf NextWeekIsMovedOralReview And IsCOVisitWeek And rstPrtSchedule!TalkNo = "1" Then
                    rstPrtItems.FindFirst "AssignmentDate = #" & Format(TheAssignmentDate + 7, "mm/dd/yyyy") & "# " & _
                                            "AND TalkNo = '1'"
                Else
                    rstPrtItems.FindFirst "AssignmentDate = #" & Format(TheAssignmentDate, "mm/dd/yyyy") & "# " & _
                                            "AND TalkNo = '" & rstPrtSchedule!TalkNo & "'"
                End If
            End Select
            
            If Not (IsNull(rstPrtSchedule!CounselPoint) And rstPrtSchedule!CounselPoint = 0) Then
                'Find Counsel Point info for scheduled assignment
                CounselSQL = "SELECT tblTMSCounselPointList.CounselPoint, " & _
                            "        tblTMSCounselPointList.CounselDescription, " & _
                            "        tblTMSCounselPointComponents.SubPointDescription, " & _
                            "        tblTMSCounselPointComponents.CounselSubPoint, " & _
                            "        tblTMSCounselPointList.PageOfBeBook " & _
                            " FROM tblTMSCounselPointList INNER JOIN tblTMSCounselPointComponents " & _
                            " ON (tblTMSCounselPointComponents.CounselPoint = tblTMSCounselPointList.CounselPoint) " & _
                            " WHERE tblTMSCounselPointList.CounselPoint = " & rstPrtSchedule!CounselPoint
                
                Set rstPrtCounsel = CMSDB.OpenRecordset(CounselSQL, dbOpenForwardOnly)
                
                With rstPrtCounsel
                If .BOF Then
                    TheCounsel = ""
                Else
                    TheCounsel = !CounselPoint & " - " & !CounselDescription & " (" & !PageOfBeBook & ")"
                    Do While Not .EOF
                        Select Case !CounselSubPoint
                        Case 1
                            SubPt1 = !SubPointDescription
                        Case 2
                            SubPt2 = !SubPointDescription
                        Case 3
                            SubPt3 = !SubPointDescription
                        Case 4
                            SubPt4 = !SubPointDescription
                        Case 5
                            SubPt5 = !SubPointDescription
                        End Select
                        rstPrtCounsel.MoveNext
                    Loop
                End If
                End With
            End If
            
            '
            'Get last comment
            '
            If frmTMSPrinting.chkIncludeCommentOnWorksheets.value = vbChecked Then
                CommentSQL = "SELECT Comment FROM tblTMSSchedule " & _
                            " WHERE AssignmentDate < #" & Format(TheAssignmentDate, "mm/dd/yyyy") & "# " & _
                            "AND PersonID = " & rstPrtSchedule!PersonID & _
                            " AND TalkNo NOT IN ('P', 'S')" & _
                            " ORDER BY AssignmentDate"
                            
                Set rstPrtComment = CMSDB.OpenRecordset(CommentSQL, dbOpenDynaset)
                
                If Not rstPrtComment.BOF Then
                    rstPrtComment.MoveLast
                    If IsNull(rstPrtComment!Comment) Then
                        TheComment = ""
                    Else
                        TheComment = Left(rstPrtComment!Comment, 250)
                    End If
                Else
                    TheComment = ""
                End If
            Else
                TheComment = ""
            End If
            
            AssistantName = CongregationMember.NameWithMiddleInitial(rstPrtSchedule!Assistant1ID)
            If Left(AssistantName, 1) = "?" Then
                AssistantName = ""
            End If
            
            '
            'Get the setting
            '
            If IsNull(rstPrtSchedule!Setting) Or rstPrtSchedule!Setting = 0 Then
                TheSettingAndDesc = ""
            Else
                rstSettings.FindFirst "SettingNo = " & rstPrtSchedule!Setting
                TheSettingAndDesc = CStr(rstSettings!SettingNo) & " - " & _
                                    rstSettings!SettingDesc
            End If
            
            Select Case rstPrtSchedule!TalkNo
            Case "O"
                sSettingTitle = "Setting:"
            Case Else
                sSettingTitle = ""
            End Select
            
            sStudentName = CongregationMember.NameWithMiddleInitial(rstPrtSchedule!PersonID)
            sStudentName = IIf(Left(sStudentName, 1) = "?", "", sStudentName)
            
           
            '
            'Now Populate the Print table....  at last....
            '
            .Requery
            .MoveLast
            .Edit
            
            Select Case bNewArr
            Case CLM2016
            
                .Fields("Item" & iItemIX & "TalkNo") = TheTMS.GetTMSTalkDescription(rstPrtSchedule!TalkNo, CStr(TheAssignmentDate))
                .Fields("Item" & iItemIX & "StudentName") = sStudentName
                .Fields("Item" & iItemIX & "AssistantName") = AssistantName
                .Fields("Item" & iItemIX & "Theme") = rstPrtItems!TalkTheme
                .Fields("Item" & iItemIX & "Source") = rstPrtItems!SourceMaterial
                .Fields("Item" & iItemIX & "CounselPoint") = TheCounsel
                .Fields("Item" & iItemIX & "CounselSubPoint1") = SubPt1
                .Fields("Item" & iItemIX & "CounselSubPoint2") = SubPt2
                .Fields("Item" & iItemIX & "CounselSubPoint3") = SubPt3
                .Fields("Item" & iItemIX & "CounselSubPoint4") = SubPt4
                .Fields("Item" & iItemIX & "CounselSubPoint5") = SubPt5
                .Fields("Item" & iItemIX & "Comment") = TheComment
                .Fields("Item" & iItemIX & "SettingName") = TheSettingAndDesc
                .Fields("Item" & iItemIX & "SettingTitle") = sSettingTitle 'this is either "Setting:" or ""
            
            Case Else
                Select Case rstPrtSchedule!TalkNo
                Case "P"
                    sSongNo = RemoveNonNumerics(rstPrtItems!SourceMaterial)
                    If IsNumber(sSongNo) Then
                        !OpeningSong = sSongNo & " - " & GetSongTitle(CLng(sSongNo))
                    Else
                        !OpeningSong = ""
                    End If
                    !PrayerBroName = sStudentName
                Case "S"
                    !SQTheme = rstPrtItems!TalkTheme
                    !SQBroName = sStudentName
                    !SQSource = rstPrtItems!SourceMaterial
                Case "1"
                    !No1BroName = sStudentName
                    !No1Theme = rstPrtItems!TalkTheme
                    !No1Source = rstPrtItems!SourceMaterial
                    !No1CounselPoint = TheCounsel
                    !No1CounselSubPoint1 = SubPt1
                    !No1CounselSubPoint2 = SubPt2
                    !No1CounselSubPoint3 = SubPt3
                    !No1CounselSubPoint4 = SubPt4
                    !No1CounselSubPoint5 = SubPt5
                    !No1Comment = TheComment
                'MJT 8/8/11
                Case "R", "MR"
                    !No1Theme = "ORAL REVIEW"
                    !No1BroName = "Reader:" & sStudentName
                Case "B"
                    !bhBroName = sStudentName
                    !BHSource = rstPrtItems!SourceMaterial
                    !bhCounselPoint = TheCounsel
                    !bhCounselSubPoint1 = SubPt1
                    !bhCounselSubPoint2 = SubPt2
                    !bhCounselSubPoint3 = SubPt3
                    !bhCounselSubPoint4 = SubPt4
                    !bhCounselSubPoint5 = SubPt5
                    !BHComment = TheComment
                    'MJT 8/8/11
    '                If IsOralReview Or IsMovedOralReview Then
    '                    !No1Theme = "ORAL REVIEW"
    '                End If
                Case "2"
                    !No2BroName = sStudentName
                    !No2AssistantName = AssistantName
                    !No2Theme = rstPrtItems!TalkTheme
                    !No2Source = rstPrtItems!SourceMaterial
                    !No2CounselPoint = TheCounsel
                    !No2CounselSubPoint1 = SubPt1
                    !No2CounselSubPoint2 = SubPt2
                    !No2CounselSubPoint3 = SubPt3
                    !No2CounselSubPoint4 = SubPt4
                    !No2CounselSubPoint5 = SubPt5
                    !No2Comment = TheComment
                    !No2Setting = TheSettingAndDesc
                Case "3"
                    !No3StudentName = sStudentName
                    !No3AssistantName = AssistantName
                    !No3Theme = rstPrtItems!TalkTheme
                    !No3Source = rstPrtItems!SourceMaterial
                    !No3CounselPoint = TheCounsel
                    !No3CounselSubPoint1 = SubPt1
                    !No3CounselSubPoint2 = SubPt2
                    !No3CounselSubPoint3 = SubPt3
                    !No3CounselSubPoint4 = SubPt4
                    !No3CounselSubPoint5 = SubPt5
                    !No3Comment = TheComment
                    !No3Setting = TheSettingAndDesc
                Case "4"
                    !No4StudentName = sStudentName
                    !No4AssistantName = AssistantName
                    !No4Theme = rstPrtItems!TalkTheme
                    !No4Source = rstPrtItems!SourceMaterial
                    !No4CounselPoint = TheCounsel
                    !No4CounselSubPoint1 = SubPt1
                    !No4CounselSubPoint2 = SubPt2
                    !No4CounselSubPoint3 = SubPt3
                    !No4CounselSubPoint4 = SubPt4
                    !No4CounselSubPoint5 = SubPt5
                    !No4Comment = TheComment
                    !No4Setting = TheSettingAndDesc
                Case "CO"
                    If Not bNewArr Then
                        !No2Theme = "CO VISIT"
                    End If
                End Select
                
    '            !ConcludingSong = sSong2
                !ConcludingSong = IIf(rstPrtSchedule!SchoolNo = 1, sSrvMtgPerson, "")
            End Select
            
            .Update
            PrevSchool = rstPrtSchedule!SchoolNo
            rstPrtSchedule.MoveNext
            If Not rstPrtSchedule.EOF Then
                If PrevSchool <> rstPrtSchedule!SchoolNo Then
                    Exit Do
                End If
            End If
        Loop
        If rstPrtSchedule.EOF Then
            Exit Do
        End If
    Loop
    End With
    
    BuildPrintTable = True
    
    Exit Function
ErrorTrap:
    EndProgram
    
End Function
Private Function BuildCounselFormPrintTable(PersonID As Long) As Boolean
Dim rstPoints As Recordset, NamesSQL As String, PrevStudent As Long
Dim rstPrintTable As Recordset, rstStudent As Recordset
Dim TheError As Integer, TalkSQL As String

Dim i As Long

On Error GoTo ErrorTrap

    On Error Resume Next
    DeleteTable "tblTMSCounselPointPrint"
    On Error GoTo ErrorTrap
    CreateTable TheError, "tblTMSCounselPointPrint", "Point", "TEXT", "10"
    CreateField TheError, "tblTMSCounselPointPrint", "CounselName", "TEXT", "200"
    CreateField TheError, "tblTMSCounselPointPrint", "DateAssigned", "TEXT", "10"
    CreateField TheError, "tblTMSCounselPointPrint", "DateComplete", "TEXT", "10"
    CreateField TheError, "tblTMSCounselPointPrint", "ExerciseComplete", "TEXT", "10"
        
    '
    'Driving Recset
    '
    NamesSQL = "SELECT DISTINCT " & _
               "CounselPoint " & _
               ",CounselDescription " & _
               "FROM tblTMSCounselPointList " & _
               "WHERE CounselPoint > 0 " & _
               "ORDER BY 1 "

    Set rstPoints = CMSDB.OpenRecordset(NamesSQL, dbOpenDynaset)
    
    Set rstPrintTable = CMSDB.OpenRecordset("tblTMSCounselPointPrint", dbOpenDynaset)
    
    If rstPoints.BOF Then
        BuildCounselFormPrintTable = False
        Exit Function
    End If
    
    With rstPrintTable
    
    Do Until rstPoints.EOF
    
        NamesSQL = "SELECT TOP 1 MAX(CounselPointAssignedDate) AS AssignedDate, " & _
                                 " CounselPointCompletedDate AS CompletedDate, " & _
                                 " ExerciseComplete " & _
                   " FROM tblTMSSchedule " & _
                   " WHERE PersonID = " & PersonID & _
                   " AND CounselPoint = " & rstPoints!CounselPoint & _
                   " AND CounselPointAssignedDate IS NOT NULL " & _
                   " AND CounselPointAssignedDate <> 0 " & _
                   " GROUP BY CounselPointCompletedDate, assignmentdate, ExerciseComplete " & _
                   " ORDER BY AssignmentDate DESC "

        Set rstStudent = CMSDB.OpenRecordset(NamesSQL, dbOpenDynaset)
        
        .AddNew
        !Point = rstPoints!CounselPoint
        !CounselName = rstPoints!CounselDescription
        If rstStudent.BOF Or rstStudent.EOF Then
            !DateAssigned = ""
            !DateComplete = ""
            !ExerciseComplete = ""
        Else
            !DateAssigned = Format(rstStudent!AssignedDate, "dd/mm/yyyy")
            !DateComplete = HandleNull(Format(rstStudent!CompletedDate, "dd/mm/yyyy"), "")
            !ExerciseComplete = IIf(rstStudent!ExerciseComplete, "Y", "")
        End If
        .Update
        
        rstPoints.MoveNext
        
    Loop
    
    End With
    
    On Error Resume Next
    rstPoints.Close
    rstStudent.Close
    rstPrintTable.Close
    Set rstPoints = Nothing
    Set rstStudent = Nothing
    Set rstPrintTable = Nothing
    
    BuildCounselFormPrintTable = True
    
    Exit Function
ErrorTrap:
    EndProgram
    
End Function
Private Function BuildStudentDetailsPrintTable() As Boolean
Dim rstStudents As Recordset, NamesSQL As String, PrevStudent As Long
Dim rstPrintTable As Recordset, b2009 As Boolean

Dim i As Long

On Error GoTo ErrorTrap

    b2009 = GlobalParms.GetValue("TMSPrintStudentDtlsFor2009Format", "TrueFalse")
    
    DelAllRows "tblTMSPrintStudentDetails"
    
    '
    'Driving Recset
    '
    NamesSQL = "SELECT DISTINCT " & _
               "tblNameAddress.ID " & _
               ",tblNameAddress.FirstName " & _
               ",tblNameAddress.MiddleName " & _
               ",tblNameAddress.LastName " & _
               ",tblTaskAndPerson.Task " & _
               "FROM tblTaskAndPerson " & _
               "INNER JOIN tblNameAddress " & _
               "ON tblNameAddress.ID = tblTaskAndPerson.Person " & _
               "WHERE TaskCategory = 4 " & _
               "AND TaskSubcategory = 6 " & _
               "AND Active = TRUE " & _
               "ORDER BY 4, 2, 3"

    Set rstStudents = CMSDB.OpenRecordset(NamesSQL, dbOpenDynaset)
    
    Set rstPrintTable = CMSDB.OpenRecordset("tblTMSPrintStudentDetails", dbOpenDynaset)
    
    If rstStudents.BOF Then
        BuildStudentDetailsPrintTable = False
        Exit Function
    Else
        PrevStudent = 0
    End If
    
    With rstPrintTable
    
    If b2009 Then
        Do Until rstStudents.EOF
            .AddNew
            !StudentName = rstStudents!LastName & ", " & rstStudents!FirstName & " " & rstStudents!MiddleName
            !School1 = IIf(CongregationMember.TMSPersonAssignedToSchool(rstStudents!ID, 1), "Y", "")
            !School2 = IIf(CongregationMember.TMSPersonAssignedToSchool(rstStudents!ID, 2), "Y", "")
            !School3 = IIf(CongregationMember.TMSPersonAssignedToSchool(rstStudents!ID, 3), "Y", "")
            .Update
            PrevStudent = rstStudents!ID
            Do Until rstStudents.EOF
                .MoveLast
                .Edit
                Select Case rstStudents!Task
                Case 34 'B/H
                    !DoesBH = IIf(CongregationMember.PersonIsSuspended2(rstStudents!ID, date, rstStudents!Task, True), "S", "Y")
                Case 99 '#1
                    !DoesNo1 = IIf(CongregationMember.PersonIsSuspended2(rstStudents!ID, date, rstStudents!Task, True), "S", "Y")
                Case 100 '#2
                    !DoesNo2 = IIf(CongregationMember.PersonIsSuspended2(rstStudents!ID, date, rstStudents!Task, True), "S", "Y")
                Case 101 '#3
                    !DoesNo3 = IIf(CongregationMember.PersonIsSuspended2(rstStudents!ID, date, rstStudents!Task, True), "S", "Y")
                Case 42, 43 'Assistant
                    !DoesAsst = IIf(CongregationMember.PersonIsSuspended2(rstStudents!ID, date, rstStudents!Task, True), "S", "Y")
                Case 47 'Prayer
                    !DoesPrayer = IIf(CongregationMember.PersonIsSuspended2(rstStudents!ID, date, rstStudents!Task, True), "S", "Y")
                End Select
                            
                .Update
                PrevStudent = rstStudents!ID
                rstStudents.MoveNext
                If Not rstStudents.EOF Then
                    If PrevStudent <> rstStudents!ID Then
                        Exit Do
                    End If
                End If
            Loop
        Loop
    Else
        Do Until rstStudents.EOF
            .AddNew
            !StudentName = rstStudents!LastName & ", " & rstStudents!FirstName & " " & rstStudents!MiddleName
            .Update
            PrevStudent = rstStudents!ID
            Do Until rstStudents.EOF
                .MoveLast
                .Edit
                Select Case rstStudents!Task
                Case 33 'S/Q
                    !DoesSQ = "Y"
                Case 34 'B/H
                    !DoesBH = "Y"
                Case 35 '#1
                    !DoesNo1 = "Y"
                Case 36, 37 '#2
                    !DoesNo2 = "Y"
                Case 38, 39 '#3
                    !DoesNo3 = "Y"
                Case 40, 41 '#4
                    !DoesNo4 = "Y"
                Case 42, 43 'Assistant
                    !DoesAsst = "Y"
                Case 44 '1st School
                    !School1 = "Y"
                Case 45 '2nd school
                    !School2 = "Y"
                Case 46 '3rd school
                    !School3 = "Y"
                Case 47 'Prayer
                    !DoesPrayer = "Y"
                End Select
                            
                .Update
                PrevStudent = rstStudents!ID
                rstStudents.MoveNext
                If Not rstStudents.EOF Then
                    If PrevStudent <> rstStudents!ID Then
                        Exit Do
                    End If
                End If
            Loop
        Loop
    End If
    
    End With
    
    BuildStudentDetailsPrintTable = True
    
    Exit Function
ErrorTrap:
    EndProgram
    End Function

Private Sub TMSPrintStudentDetailsWithWord()

On Error GoTo ErrorTrap

Dim reporter As MSWordReportingTool2.RptTool

    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .SaveDoc = True
    .DocPath = gsDocsDirectory & "\" & "Theocratic Ministry School" & _
                                 " Student Details " & _
                                Replace(Replace(Now, ":", "-"), "/", "-")

    
    .ReportSQL = "SELECT StudentName, DoesPrayer, DoesSQ, DoesNo1, DoesBH, " & _
                    "DoesNo2, DoesNo3, DoesNo4, DoesAsst, School1, School2, School3 " & _
                 "FROM tblTMSPrintStudentDetails " & _
                 "ORDER BY 1"

    .ReportTitle = "Theocratic Ministry School" & vbCrLf & "Student Details"
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 10
    .RightMargin = 10
    .ReportFooterFontName = "Arial"
    .ReportFooterFontSize = 8
    .ReportTitleFontName = "Times New Roman"
    .ReportTitleFontSize = 16
    .ApplyTableFormatting = True
    .ClientName = AppName
    .AdditionalReportHeading = ""
    .GroupingColumn = 0
    .HideWordWhileBuilding = True
    .ShowProgress = True
    
    .AddTableColumnAttribute "Student Name", 48, , , , , 9, 10, True, , , , True
    .AddTableColumnAttribute "Pray", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "S/Q", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "No1", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "B/H", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "No2", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "No3", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "No4", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "Asst", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "Sch1", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "Sch2", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "Sch3", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    
    .PageFormat = cmsPortrait
    
    .GenerateReport

    End With
    
    SwitchOnDAO
    
    Set reporter = Nothing
    Screen.MousePointer = vbNormal

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub TMSPrintStudentDetailsWithWord_2009()

On Error GoTo ErrorTrap

Dim reporter As MSWordReportingTool2.RptTool

    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .SaveDoc = True
    .DocPath = gsDocsDirectory & "\" & "Theocratic Ministry School" & _
                                 " Student Details " & _
                                Replace(Replace(Now, ":", "-"), "/", "-")

    
    .ReportSQL = "SELECT StudentName, DoesBH, " & _
                    "DoesNo1, DoesNo2, DoesNo3, DoesAsst, School1, School2, School3 " & _
                 "FROM tblTMSPrintStudentDetails " & _
                 "ORDER BY 1"

    .ReportTitle = "Theocratic Ministry School" & vbCrLf & "Student Details"
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 10
    .RightMargin = 10
    .ReportFooterFontName = "Arial"
    .ReportFooterFontSize = 8
    .ReportTitleFontName = "Times New Roman"
    .ReportTitleFontSize = 16
    .ApplyTableFormatting = True
    .ClientName = AppName
    .AdditionalReportHeading = ""
    .GroupingColumn = 0
    .HideWordWhileBuilding = True
    .ShowProgress = True
    
    .AddTableColumnAttribute "Student Name", 48, , , , , 9, 10, True, , , , True
    .AddTableColumnAttribute "B/H", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "No1", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "No2", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "No3", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "Asst", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "Sch1", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "Sch2", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    .AddTableColumnAttribute "Sch3", 13, cmsCentreTop, cmsCentreTop, , , 9, 10, True
    
    .PageFormat = cmsPortrait
    
    .GenerateReport

    End With
    
    SwitchOnDAO
    
    Set reporter = Nothing
    Screen.MousePointer = vbNormal

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub TMSPrintCounselPointsWithWord(PersonID As Long)

On Error GoTo ErrorTrap

Dim reporter As MSWordReportingTool2.RptTool, str As String

    str = CongregationMember.FullName(PersonID)

    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .SaveDoc = True
    .DocPath = gsDocsDirectory & "\" & "Theocratic Ministry School" & _
                                 " Counsel Form for " & str & " (" & _
                                Replace(Replace(Now, ":", "-"), "/", "-") & ")"

    
    .ReportSQL = "SELECT CLng(Point), CounselName, " & _
                    "DateAssigned, DateComplete, ExerciseComplete " & _
                 "FROM tblTMSCounselPointPrint " & _
                 "ORDER BY 1"

    .ReportTitle = "Theocratic Ministry School Counsel Form" & vbCrLf & str
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 10
    .RightMargin = 10
    .ReportFooterFontName = "Arial"
    .ReportFooterFontSize = 8
    .ReportTitleFontName = "Times New Roman"
    .ReportTitleFontSize = 16
    .ApplyTableFormatting = True
    .ClientName = AppName
    .AdditionalReportHeading = Format(Now, "dd/mm/yyyy")
    .GroupingColumn = 0
    .HideWordWhileBuilding = True
    .ShowProgress = True
    
    .AddTableColumnAttribute "No", 10, , , , , 9, 8.5, True, , , , True
    .AddTableColumnAttribute "Study Name", 80, , , , , 9, 8.5, True
    .AddTableColumnAttribute "Assigned", 25, , , , , 9, 8.5, True
    .AddTableColumnAttribute "Completed", 25, , , , , 9, 8.5, True
    .AddTableColumnAttribute "Ex", 10, cmsCentreTop, cmsCentreTop, , , 9, 8.5, True
    
    .PageFormat = cmsPortrait
    
    .GenerateReport

    End With
    
    SwitchOnDAO
    
    Set reporter = Nothing
    Screen.MousePointer = vbNormal

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
