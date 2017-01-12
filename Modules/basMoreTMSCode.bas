Attribute VB_Name = "basMoreTMSCode"
Option Explicit
Dim rstStuffForSlips As Recordset
Dim MinTopMargin As Single
Dim MinLeftMargin As Single
Dim StudentXPos As Single
Dim StudentYPos As Single
Dim StudentMaxHeight As Single
Dim StudentMaxWidth As Single
Dim AssistantXPos As Single
Dim AssistantYPos As Single
Dim AssistantMaxHeight As Single
Dim AssistantMaxWidth As Single
Dim ThemeXPos As Single
Dim ThemeYPos As Single
Dim ThemeMaxHeight As Single
Dim ThemeMaxWidth As Single
Dim SourceXPos As Single
Dim SourceYPos As Single
Dim SourceMaxHeight As Single
Dim SourceMaxWidth As Single
Dim CounselXPos As Single
Dim CounselYPos As Single
Dim CounselMaxHeight As Single
Dim CounselMaxWidth As Single
Dim TalkNoXPos As Single
Dim TalkNoYPos As Single
Dim TalkNoMaxHeight As Single
Dim TalkNoMaxWidth As Single
Dim AssignmentDateXPos As Single
Dim AssignmentDateYPos As Single
Dim AssignmentDateMaxHeight As Single
Dim AssignmentDateMaxWidth As Single
Dim SchoolNo1XPos As Single
Dim SchoolNo1YPos As Single
Dim SchoolNo1MaxHeight As Single
Dim SchoolNo1MaxWidth As Single
Dim SchoolNo2XPos As Single
Dim SchoolNo2YPos As Single
Dim SchoolNo2MaxHeight As Single
Dim SchoolNo2MaxWidth As Single
Dim SchoolNo3XPos As Single
Dim SchoolNo3YPos As Single
Dim SchoolNo3MaxHeight As Single
Dim SchoolNo3MaxWidth As Single
Dim UserTweakX As Single
Dim UserTweakY As Single
Dim PaperHeight As Single
Dim PaperWidth As Single
Dim MaxSourceWidth As Single
Dim MaxThemeWidth As Single
Dim StudentNoteX As Single
Dim StudentNoteY As Single
Dim StudentNote2X As Single
Dim StudentNote2Y As Single
Dim SettingX As Single
Dim SettingY As Single
Dim EndLineY As Single
Dim HeadingX As Single
Dim HeadingY As Single
Dim SubSlipCount As Integer
Dim SchoolNoXPos As Single
Dim SchoolNoYPos As Single

Dim SaveX As Single
Dim SaveY As Single

Dim SlipFontSize As Long
Dim SlipFontName As String
Dim CurrentFontSize As Long

Dim DefaultPageHeight As Single
Dim DefaultPageWidth As Single
Dim DefaultOrientation As Long

Dim SourceMaterialA As String, SourceMaterialB As String
Dim ThemeA As String, ThemeB As String
Dim ThemeHasTwoLines As Boolean
Dim SourceHasTwoLines As Boolean


Public Function PrintTheAssignmentSlips() As Boolean

On Error GoTo ErrorTrap

Dim i As Long, lMaxSlipIX As Long
Dim frm As Form
Dim bSingleSlip As Boolean
Dim fDefaultPaperHeight As Single
Dim fDefaultPaperWidth  As Single
Dim NewArr As MidweekMtgVersion

    fDefaultPaperHeight = Printer.Height
    fDefaultPaperWidth = Printer.Width

    SetUpPrintParameters
    
    
    Set rstStuffForSlips = CMSDB.OpenRecordset("tblTMSAssignmentSlips", dbOpenDynaset)
    
    
    With rstStuffForSlips
    
    .MoveLast
    .MoveFirst
    
    NewArr = NewMtgArrangementStarted(CStr(rstStuffForSlips!AssignmentDate))
    
    bSingleSlip = (.RecordCount = 1)
    
    'remove this line to reinstate single slip print mode
    bSingleSlip = False
    
    If NewArr = CLM2016 Then
        lMaxSlipIX = 3
    Else
        If bSingleSlip Then
            lMaxSlipIX = 0
        Else
            lMaxSlipIX = 2
        End If
    End If
    
    
    If Not .BOF Then
    '
    'Loop round each slip
    '

        SubSlipCount = 0
        
        Select Case NewArr
        Case CLM2016
            Set frm = New frmPrintTMSSlip_2016
        Case Else
            If Not bSingleSlip Then
                Set frm = New frmPrintTMSSlip
            Else
                Set frm = New frmPrintTMSSlipSingle
            End If
        End Select
        
        frm.Visible = False
        
        Load frm
    
        Do While Not .EOF
            
            With frm
            
            If NewArr = CLM2016 Then
            
                For i = 0 To lMaxSlipIX
                    .lblAsst(i) = ""
                    .lblDate(i) = ""
                    .lblNote1(i) = ""
                    .lblNote2(i) = ""
                    .lblSchool1(i) = ""
                    .lblSchool2(i) = ""
                    .lblSchool3(i) = ""
                    .lblSQ(i) = ""
                    .lblStudent(i) = ""
                    .lblThemeB(i) = ""
                    .lblBR(i) = ""
                    .lblIC(i) = ""
                    .lblRV(i) = ""
                    .lblBS(i) = ""
                    .lblO(i) = ""
                Next i
    
                For SubSlipCount = 0 To lMaxSlipIX
                    CheckThemeAndSourceLength
                    .lblAsst(SubSlipCount) = rstStuffForSlips!AssistantName
                    .lblDate(SubSlipCount) = rstStuffForSlips!AssignmentDate
                    .lblNote1(SubSlipCount) = rstStuffForSlips!StudentNote
                    .lblNote2(SubSlipCount) = rstStuffForSlips!StudentNote2
                    
                    Select Case rstStuffForSlips!SchoolNo
                    Case 1
                        .lblSchool1(SubSlipCount) = "X"
                        .lblSchool2(SubSlipCount) = ""
                        .lblSchool3(SubSlipCount) = ""
                    Case 2
                        .lblSchool1(SubSlipCount) = ""
                        .lblSchool2(SubSlipCount) = "X"
                        .lblSchool3(SubSlipCount) = ""
                    Case 3
                        .lblSchool1(SubSlipCount) = ""
                        .lblSchool2(SubSlipCount) = ""
                        .lblSchool3(SubSlipCount) = "X"
                    End Select
                    
                    Select Case rstStuffForSlips!TalkNo
                    Case "BR"
                        .lblBR(SubSlipCount) = "X"
                        .lblIC(SubSlipCount) = ""
                        .lblRV(SubSlipCount) = ""
                        .lblBS(SubSlipCount) = ""
                        .lblO(SubSlipCount) = ""
                    Case "IC"
                        .lblBR(SubSlipCount) = ""
                        .lblIC(SubSlipCount) = "X"
                        .lblRV(SubSlipCount) = ""
                        .lblBS(SubSlipCount) = ""
                        .lblO(SubSlipCount) = ""
                    Case "RV"
                        .lblBR(SubSlipCount) = ""
                        .lblIC(SubSlipCount) = ""
                        .lblRV(SubSlipCount) = "X"
                        .lblBS(SubSlipCount) = ""
                        .lblO(SubSlipCount) = ""
                    Case "BS"
                        .lblBR(SubSlipCount) = ""
                        .lblIC(SubSlipCount) = ""
                        .lblRV(SubSlipCount) = ""
                        .lblBS(SubSlipCount) = "X"
                        .lblO(SubSlipCount) = ""
                    Case "O"
                        .lblBR(SubSlipCount) = ""
                        .lblIC(SubSlipCount) = ""
                        .lblRV(SubSlipCount) = ""
                        .lblBS(SubSlipCount) = ""
                        .lblO(SubSlipCount) = "X"
                    End Select
                    
                    .lblSQ(SubSlipCount) = rstStuffForSlips!CounselPointNo
                    .lblStudent(SubSlipCount) = rstStuffForSlips!StudentName
                    .lblThemeB(SubSlipCount) = ThemeB
            
                    rstStuffForSlips.MoveNext
                    If rstStuffForSlips.EOF Then
                        Exit For
                    End If
                                
                Next SubSlipCount
            
            Else
            
                For i = 0 To lMaxSlipIX
                    .lblAsst(i) = ""
                    .lblDate(i) = ""
                    .lblNote1(i) = ""
                    .lblNote2(i) = ""
                    .lblSchool1(i) = ""
                    .lblSchool2(i) = ""
                    .lblSchool3(i) = ""
                    .lblSetting(i) = ""
                    .lblSourceA(i) = ""
                    .lblSourceB(i) = ""
                    .lblSQ(i) = ""
                    .lblStudent(i) = ""
                    .lblTalkNo(i) = ""
                    .lblThemeA(i) = ""
                    .lblThemeB(i) = ""
                Next i
    
                For SubSlipCount = 0 To lMaxSlipIX
                    CheckThemeAndSourceLength
                    .lblAsst(SubSlipCount) = rstStuffForSlips!AssistantName
                    .lblDate(SubSlipCount) = rstStuffForSlips!AssignmentDate
                    .lblNote1(SubSlipCount) = rstStuffForSlips!StudentNote
                    .lblNote2(SubSlipCount) = rstStuffForSlips!StudentNote2
                    
                    Select Case rstStuffForSlips!SchoolNo
                    Case 1
                        .lblSchool1(SubSlipCount) = "X"
                        .lblSchool2(SubSlipCount) = ""
                        .lblSchool3(SubSlipCount) = ""
                    Case 2
                        .lblSchool1(SubSlipCount) = ""
                        .lblSchool2(SubSlipCount) = "X"
                        .lblSchool3(SubSlipCount) = ""
                    Case 3
                        .lblSchool1(SubSlipCount) = ""
                        .lblSchool2(SubSlipCount) = ""
                        .lblSchool3(SubSlipCount) = "X"
                    End Select
                    
                    .lblSetting(SubSlipCount) = rstStuffForSlips!Setting
                    .lblSourceA(SubSlipCount) = SourceMaterialA
                    .lblSourceB(SubSlipCount) = SourceMaterialB
                    .lblSQ(SubSlipCount) = rstStuffForSlips!CounselPointNo
                    .lblStudent(SubSlipCount) = rstStuffForSlips!StudentName
                    .lblTalkNo(SubSlipCount) = rstStuffForSlips!TalkNo
                    .lblThemeA(SubSlipCount) = ThemeA
                    .lblThemeB(SubSlipCount) = ThemeB
                    
                    rstStuffForSlips.MoveNext
                    If rstStuffForSlips.EOF Then
                        Exit For
                    End If
                                
                Next SubSlipCount
               
            End If
                
            
            If bSingleSlip Then
                Printer.Height = 566.929 * PaperHeight 'set in setupprintparameters earlier in this proc
                Printer.Width = 566.929 * PaperWidth
            End If
            
            .PrintForm
            Printer.EndDoc
            
            If bSingleSlip Then
                Printer.Height = fDefaultPaperHeight
                Printer.Width = fDefaultPaperWidth
            End If
            
            End With
                        
                
        Loop
        
        Unload frm
        Set frm = Nothing
'        Unload frmPrintTMSSlip
'        Set frmPrintTMSSlip = Nothing
    
    End If
    
    End With
                
    rstStuffForSlips.Close
    
    If MsgBox("The selected Assignment Slips have been sent to the printer. " & _
           vbCrLf & "Have ALL slips been successfully printed?", vbQuestion + vbYesNo, AppName) = vbNo Then
        
        PrintTheAssignmentSlips = False
        MsgBox "Please check which slips need to be reprinted, then try again.", _
                vbOKOnly + vbInformation, AppName
                
                
    Else
        PrintTheAssignmentSlips = True
        
        If PrintUsingWord(False) Then
            If MsgBox("Do you want to print a checklist? ", vbQuestion + vbYesNo, AppName) = vbYes Then
                PrintAssignmentSlipCheckList
                frmTMSPrinting.SetUpTMSSlipPrintMainRecSet 're-connect recsets
            End If
        End If
    End If

    Exit Function
ErrorTrap:
    On Error Resume Next
    Unload frm
    Set frm = Nothing
    MsgBox "There was a printing error. Please check printer and try again. (" & Err.Description & ")", vbOKOnly + vbCritical, AppName
    PrintTheAssignmentSlips = False
    rstStuffForSlips.Close
    Exit Function
    
End Function

Public Sub PrintAssignmentSlipCheckList()

Dim reporter As MSWordReportingTool2.RptTool

On Error GoTo ErrorTrap

    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .SaveDoc = True
    .DocPath = gsDocsDirectory & "\" & "TMS Assignment Slip Checklist " & _
                                Replace(Replace(Now, ":", "-"), "/", "-")

    
    .ReportSQL = "SELECT StudentName, " & _
                 "       AssignmentDate, " & _
                 "       TalkNo, " & _
                 "       SchoolNo, " & _
                 "       ' ' " & _
                 "FROM tblTMSAssignmentSlips "

    .ReportTitle = "TMS Assignment Slip Checklist"
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 10
    .RightMargin = 10
    .ReportFooterFontName = "Arial"
    .ReportFooterFontSize = 8
    .ReportTitleFontName = "Times New Roman"
    .ReportTitleFontSize = 18
    .ApplyTableFormatting = True
    .ClientName = AppName
    .AdditionalReportHeading = ""
    .GroupingColumn = 0
    .HideWordWhileBuilding = True
    
    .AddTableColumnAttribute "Student Name", 60, , , , , 10, 10, True, True, , , True
    .AddTableColumnAttribute "Assignment Date", 30, , , , , 10, 10, True, True
    .AddTableColumnAttribute "Talk No", 20, , , , , 10, 10, True, True
    .AddTableColumnAttribute "School No", 20, , , , , 10, 10, True, True
    .AddTableColumnAttribute "Slip Given", 35, , , , , 10, 10, True, True
    
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



Private Sub SetUpPrintParameters()

On Error GoTo ErrorTrap

     Printer.ScaleMode = vbCentimeters
     MinTopMargin = GlobalParms.GetValue("TMSSlipMinTopMargin", "NumFloat")
     MinLeftMargin = GlobalParms.GetValue("TMSSlipMinLeftMargin", "NumFloat")
     StudentXPos = GlobalParms.GetValue("TMSSlipStudentXPos", "NumFloat")
     StudentYPos = GlobalParms.GetValue("TMSSlipStudentYPos", "NumFloat")
     StudentMaxHeight = GlobalParms.GetValue("TMSSlipStudentMaxHeight", "NumFloat")
     StudentMaxWidth = GlobalParms.GetValue("TMSSlipStudentMaxWidth", "NumFloat")
     AssistantXPos = GlobalParms.GetValue("TMSSlipAssistantXPos", "NumFloat")
     AssistantYPos = GlobalParms.GetValue("TMSSlipAssistantYPos", "NumFloat")
     AssistantMaxHeight = GlobalParms.GetValue("TMSSlipAssistantMaxHeight", "NumFloat")
     AssistantMaxWidth = GlobalParms.GetValue("TMSSlipAssistantMaxWidth", "NumFloat")
     ThemeXPos = GlobalParms.GetValue("TMSSlipThemeXPos", "NumFloat")
     ThemeYPos = GlobalParms.GetValue("TMSSlipThemeYPos", "NumFloat")
     ThemeMaxHeight = GlobalParms.GetValue("TMSSlipThemeMaxHeight", "NumFloat")
     ThemeMaxWidth = GlobalParms.GetValue("TMSSlipThemeMaxWidth", "NumFloat")
     SourceXPos = GlobalParms.GetValue("TMSSlipSourceXPos", "NumFloat")
     SourceYPos = GlobalParms.GetValue("TMSSlipSourceYPos", "NumFloat")
     SourceMaxHeight = GlobalParms.GetValue("TMSSlipSourceMaxHeight", "NumFloat")
     SourceMaxWidth = GlobalParms.GetValue("TMSSlipSourceMaxWidth", "NumFloat")
     CounselXPos = GlobalParms.GetValue("TMSSlipCounselXPos", "NumFloat")
     CounselYPos = GlobalParms.GetValue("TMSSlipCounselYPos", "NumFloat")
     CounselMaxHeight = GlobalParms.GetValue("TMSSlipCounselMaxHeight", "NumFloat")
     CounselMaxWidth = GlobalParms.GetValue("TMSSlipCounselMaxWidth", "NumFloat")
     TalkNoXPos = GlobalParms.GetValue("TMSSlipTalkNoXPos", "NumFloat")
     TalkNoYPos = GlobalParms.GetValue("TMSSlipTalkNoYPos", "NumFloat")
     TalkNoMaxHeight = GlobalParms.GetValue("TMSSlipTalkNoMaxHeight", "NumFloat")
     TalkNoMaxWidth = GlobalParms.GetValue("TMSSlipTalkNoMaxWidth", "NumFloat")
     AssignmentDateXPos = GlobalParms.GetValue("TMSSlipAssignmentDateXPos", "NumFloat")
     AssignmentDateYPos = GlobalParms.GetValue("TMSSlipAssignmentDateYPos", "NumFloat")
     AssignmentDateMaxHeight = GlobalParms.GetValue("TMSSlipAssignmentDateMaxHeight", "NumFloat")
     AssignmentDateMaxWidth = GlobalParms.GetValue("TMSSlipAssignmentDateMaxWidth", "NumFloat")
     UserTweakX = GlobalParms.GetValue("TMSSlipUserTweakX", "NumFloat")
     UserTweakY = GlobalParms.GetValue("TMSSlipUserTweakY", "NumFloat")
     SchoolNo1XPos = GlobalParms.GetValue("TMSSlipSchoolNo1XPos", "NumFloat")
     SchoolNo1YPos = GlobalParms.GetValue("TMSSlipSchoolNo1YPos", "NumFloat")
     SchoolNo1MaxHeight = GlobalParms.GetValue("TMSSlipSchoolNo1MaxHeight", "NumFloat")
     SchoolNo1MaxWidth = GlobalParms.GetValue("TMSSlipSchoolNo1MaxWidth", "NumFloat")
     SchoolNo2XPos = GlobalParms.GetValue("TMSSlipSchoolNo2XPos", "NumFloat")
     SchoolNo2YPos = GlobalParms.GetValue("TMSSlipSchoolNo2YPos", "NumFloat")
     SchoolNo2MaxHeight = GlobalParms.GetValue("TMSSlipSchoolNo2MaxHeight", "NumFloat")
     SchoolNo2MaxWidth = GlobalParms.GetValue("TMSSlipSchoolNo2MaxWidth", "NumFloat")
     SchoolNo3XPos = GlobalParms.GetValue("TMSSlipSchoolNo3XPos", "NumFloat")
     SchoolNo3YPos = GlobalParms.GetValue("TMSSlipSchoolNo3YPos", "NumFloat")
     SchoolNo3MaxHeight = GlobalParms.GetValue("TMSSlipSchoolNo3MaxHeight", "NumFloat")
     SchoolNo3MaxWidth = GlobalParms.GetValue("TMSSlipSchoolNo3MaxWidth", "NumFloat")
     PaperHeight = GlobalParms.GetValue("TMSSlipPaperHeight", "NumFloat")
     PaperWidth = GlobalParms.GetValue("TMSSlipPaperWidth", "NumFloat")
     MaxSourceWidth = GlobalParms.GetValue("TMSSlipMaxSourceWidth", "NumFloat")
     MaxThemeWidth = GlobalParms.GetValue("TMSSlipMaxThemeWidth", "NumFloat")
     StudentNoteX = GlobalParms.GetValue("TMSSlipStudentNoteXPos", "NumFloat")
     StudentNoteY = GlobalParms.GetValue("TMSSlipStudentNoteYPos", "NumFloat")
     StudentNote2X = GlobalParms.GetValue("TMSSlipStudentNote2XPos", "NumFloat")
     StudentNote2Y = GlobalParms.GetValue("TMSSlipStudentNote2YPos", "NumFloat")
     SettingX = GlobalParms.GetValue("TMSSlipSettingXPos", "NumFloat")
     SettingY = GlobalParms.GetValue("TMSSlipSettingYPos", "NumFloat")
     
     SlipFontSize = GlobalParms.GetValue("TMSSlipFontSize", "NumFloat")
     SlipFontName = GlobalParms.GetValue("TMSSlipFontName", "AlphaVal")

     Printer.Font.Name = SlipFontName
     Printer.Font.Size = SlipFontSize


    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub SetUpPrintParameters_Sub()

On Error GoTo ErrorTrap

     Printer.ScaleMode = vbCentimeters
     
     StudentXPos = GlobalParms.GetValue("TMSSlipStudentXPos_Sub", "NumFloat")
     StudentYPos = GlobalParms.GetValue("TMSSlipStudentYPos_Sub", "NumFloat")
     AssistantXPos = GlobalParms.GetValue("TMSSlipAssistantXPos_Sub", "NumFloat")
     AssistantYPos = GlobalParms.GetValue("TMSSlipAssistantYPos_Sub", "NumFloat")
     ThemeXPos = GlobalParms.GetValue("TMSSlipThemeXPos_Sub", "NumFloat")
     ThemeYPos = GlobalParms.GetValue("TMSSlipThemeYPos_Sub", "NumFloat")
     SourceXPos = GlobalParms.GetValue("TMSSlipSourceXPos_Sub", "NumFloat")
     SourceYPos = GlobalParms.GetValue("TMSSlipSourceYPos_Sub", "NumFloat")
     CounselXPos = GlobalParms.GetValue("TMSSlipCounselXPos_Sub", "NumFloat")
     CounselYPos = GlobalParms.GetValue("TMSSlipCounselYPos_Sub", "NumFloat")
     TalkNoXPos = GlobalParms.GetValue("TMSSlipTalkNoXPos_Sub", "NumFloat")
     TalkNoYPos = GlobalParms.GetValue("TMSSlipTalkNoYPos_Sub", "NumFloat")
     AssignmentDateXPos = GlobalParms.GetValue("TMSSlipAssignmentDateXPos_Sub", "NumFloat")
     AssignmentDateYPos = GlobalParms.GetValue("TMSSlipAssignmentDateYPos_Sub", "NumFloat")
     SchoolNoXPos = GlobalParms.GetValue("TMSSlipSchoolNoXPos_Sub", "NumFloat")
     SchoolNoYPos = GlobalParms.GetValue("TMSSlipSchoolNoYPos_Sub", "NumFloat")
     MaxSourceWidth = GlobalParms.GetValue("TMSSlipMaxSourceWidth_Sub", "NumFloat")
     MaxThemeWidth = GlobalParms.GetValue("TMSSlipMaxThemeWidth_Sub", "NumFloat")
     StudentNoteX = GlobalParms.GetValue("TMSSlipStudentNoteXPos_Sub", "NumFloat")
     StudentNoteY = GlobalParms.GetValue("TMSSlipStudentNoteYPos_Sub", "NumFloat")
     StudentNote2X = GlobalParms.GetValue("TMSSlipStudentNote2XPos_Sub", "NumFloat")
     StudentNote2Y = GlobalParms.GetValue("TMSSlipStudentNote2YPos_Sub", "NumFloat")
     SettingX = GlobalParms.GetValue("TMSSlipSettingXPos_Sub", "NumFloat")
     SettingY = GlobalParms.GetValue("TMSSlipSettingYPos_Sub", "NumFloat")
     PaperHeight = GlobalParms.GetValue("TMSSlipPaperHeight_Sub", "NumFloat")
     EndLineY = GlobalParms.GetValue("TMSSlipEndLineY_Sub", "NumFloat")
     HeadingX = GlobalParms.GetValue("TMSSlipHeadingX_Sub", "NumFloat")
     HeadingY = GlobalParms.GetValue("TMSSlipHeadingY_Sub", "NumFloat")
     UserTweakX = GlobalParms.GetValue("A4LeftMargin", "NumFloat")
     UserTweakY = GlobalParms.GetValue("A4TopMargin", "NumFloat")
     
     SlipFontSize = GlobalParms.GetValue("TMSSlipFontSize_Sub", "NumFloat")
     SlipFontName = GlobalParms.GetValue("TMSSlipFontName_Sub", "AlphaVal")

     Printer.Font.Name = SlipFontName
     Printer.Font.Size = SlipFontSize


    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintStudentName()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = StudentXPos
    .CurrentY = StudentYPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + UserTweakX + MinLeftMargin
    .CurrentY = .CurrentY + UserTweakY + MinTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print rstStuffForSlips!StudentName
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintThemeA()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = ThemeXPos
    .CurrentY = ThemeYPos - 2 * .ScaleY(SlipFontSize, vbPoints, vbCentimeters) - 0.05
    .CurrentX = .CurrentX + UserTweakX + MinLeftMargin
    .CurrentY = .CurrentY + UserTweakY + MinTopMargin
    
    Printer.Print ThemeA
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintThemeB()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = ThemeXPos
    .CurrentY = ThemeYPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters) '+ 0.4
    .CurrentX = .CurrentX + UserTweakX + MinLeftMargin
    .CurrentY = .CurrentY + UserTweakY + MinTopMargin
    
    Printer.Print ThemeB
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintAssistantName()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = AssistantXPos
    .CurrentY = AssistantYPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + UserTweakX + MinLeftMargin
    .CurrentY = .CurrentY + UserTweakY + MinTopMargin
    
    Printer.Print rstStuffForSlips!AssistantName
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintSourceA()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = SourceXPos
    .CurrentY = SourceYPos - 2 * .ScaleY(SlipFontSize, vbPoints, vbCentimeters) - 0.05
    .CurrentX = .CurrentX + UserTweakX + MinLeftMargin
    .CurrentY = .CurrentY + UserTweakY + MinTopMargin
    
    Printer.Print SourceMaterialA
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintSourceB()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = SourceXPos
    .CurrentY = SourceYPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters) '+ 0.4
    .CurrentX = .CurrentX + UserTweakX + MinLeftMargin
    .CurrentY = .CurrentY + UserTweakY + MinTopMargin
    
    Printer.Print SourceMaterialB
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintCounselPoint()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CounselXPos
    .CurrentY = CounselYPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + UserTweakX + MinLeftMargin
    .CurrentY = .CurrentY + UserTweakY + MinTopMargin
    
    Printer.Print rstStuffForSlips!CounselPointNo
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintTalkNo()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = TalkNoXPos
    .CurrentY = TalkNoYPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + UserTweakX + MinLeftMargin
    .CurrentY = .CurrentY + UserTweakY + MinTopMargin
    
    If rstStuffForSlips!TalkNo = "B" Then
        Printer.Print "BH"
    ElseIf rstStuffForSlips!TalkNo = "S" Then
        Printer.Print "SQ"
    Else
        Printer.Print rstStuffForSlips!TalkNo
    End If
    
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintStudentNote()

On Error GoTo ErrorTrap

    With Printer
    
    .CurrentX = StudentNoteX
    .CurrentY = StudentNoteY - .ScaleY(SlipFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + UserTweakX + MinLeftMargin
    .CurrentY = .CurrentY + UserTweakY + MinTopMargin
    
    .FontBold = True
    .FontItalic = True
    
    Printer.Print rstStuffForSlips!StudentNote
    
    .FontBold = False
    .FontItalic = False
    
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintStudentNote2()

On Error GoTo ErrorTrap

    With Printer
    
    .CurrentX = StudentNote2X
    .CurrentY = StudentNote2Y - .ScaleY(SlipFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + UserTweakX + MinLeftMargin
    .CurrentY = .CurrentY + UserTweakY + MinTopMargin
    
    .FontBold = False
    .FontItalic = True
    
    Printer.Print rstStuffForSlips!StudentNote2
    
    .FontBold = False
    .FontItalic = False
    
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintSchoolNo()

On Error GoTo ErrorTrap

    With Printer
    .CurrentY = SchoolNo1YPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters)
    .CurrentY = .CurrentY + UserTweakY + MinTopMargin
    
    Select Case rstStuffForSlips!SchoolNo
    Case 1
        .CurrentX = SchoolNo1XPos
    Case 2
        .CurrentX = SchoolNo2XPos
    Case 3
        .CurrentX = SchoolNo3XPos
    End Select
        
    .CurrentX = .CurrentX + UserTweakX + MinLeftMargin
    Printer.Print "x"
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintSetting()

On Error GoTo ErrorTrap

    With Printer
    .CurrentY = SettingY - .ScaleY(SlipFontSize, vbPoints, vbCentimeters)
    .CurrentY = .CurrentY + UserTweakY + MinTopMargin
    .CurrentX = SettingX
    .CurrentX = .CurrentX + UserTweakX + MinLeftMargin
    Printer.Print rstStuffForSlips!Setting
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintAssignmentDate()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = AssignmentDateXPos
    .CurrentY = AssignmentDateYPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + UserTweakX + MinLeftMargin
    .CurrentY = .CurrentY + UserTweakY + MinTopMargin
    
    Printer.Print rstStuffForSlips!AssignmentDate
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub CheckThemeAndSourceLength()

On Error GoTo ErrorTrap

Dim str As String

    If Printer.TextWidth(rstStuffForSlips!SourceMaterial) > MaxSourceWidth Then
        SourceMaterialA = Left(rstStuffForSlips!SourceMaterial, 25)
        If Len(rstStuffForSlips!SourceMaterial) > 25 Then
            SourceMaterialB = Right(rstStuffForSlips!SourceMaterial, Len(rstStuffForSlips!SourceMaterial) - 25)
        Else
            SourceMaterialB = SourceMaterialA
            SourceMaterialA = ""
        End If
        
    Else
        SourceMaterialA = ""
        SourceMaterialB = rstStuffForSlips!SourceMaterial
    End If
    
    If Printer.TextWidth(rstStuffForSlips!TalkTheme) > MaxThemeWidth Then
        ThemeA = Left(rstStuffForSlips!TalkTheme, 85)
        If Len(rstStuffForSlips!TalkTheme) > 85 Then
            ThemeB = Right(rstStuffForSlips!TalkTheme, Len(rstStuffForSlips!TalkTheme) - 85)
        Else
            ThemeB = ThemeA
            ThemeA = ""
        End If
    Else
        ThemeA = ""
        ThemeB = rstStuffForSlips!TalkTheme
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub PrintStudentName_Sub()

On Error GoTo ErrorTrap

    With Printer
    .CurrentY = StudentYPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters) + _
                UserTweakY + SubSlipCount * PaperHeight
    .CurrentX = StudentXPos + UserTweakX
    
    '
    'Ridiculously, CurrentX and CurrentY are wiped after a Print, so save 'em....
    '
    SaveX = .CurrentX
    SaveY = .CurrentY
    
    .FontBold = True
    Printer.Print "NAME: "
    .FontBold = False
    
    .CurrentX = SaveX + .TextWidth("NAME: ")
    .CurrentY = SaveY
    
    Printer.Print rstStuffForSlips!StudentName
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintHeading_Sub()


On Error GoTo ErrorTrap

    With Printer
    .CurrentY = HeadingY - .ScaleY(SlipFontSize, vbPoints, vbCentimeters) + _
                UserTweakY + SubSlipCount * PaperHeight
    .CurrentX = HeadingX + UserTweakX
        
    .FontBold = True
    Printer.Print "THEOCRATIC MINISTRY SCHOOL ASSIGNMENT"
    .FontBold = False
    
    End With

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Private Sub PrintThemeA_Sub()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = ThemeXPos + UserTweakX
    .CurrentY = ThemeYPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters) + _
                UserTweakY + SubSlipCount * PaperHeight
    
    SaveX = .CurrentX
    SaveY = .CurrentY
    
    .FontBold = True
    Printer.Print "THEME: "
    .FontBold = False
    
    .CurrentX = SaveX + .TextWidth("THEME: ")
    .CurrentY = SaveY
    
    Printer.Print ThemeA
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintThemeB_Sub()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = ThemeXPos + UserTweakX + .TextWidth("THEME: ")
    .CurrentY = ThemeYPos + 0.5 * .ScaleY(SlipFontSize, vbPoints, vbCentimeters) + _
                UserTweakY + SubSlipCount * PaperHeight
    
    Printer.Print ThemeB
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintAssistantName_Sub()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = AssistantXPos + UserTweakX
    .CurrentY = AssistantYPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters) + _
                UserTweakY + SubSlipCount * PaperHeight
    
    SaveX = .CurrentX
    SaveY = .CurrentY
    
    .FontBold = True
    Printer.Print "ASSISTANT: "
    .FontBold = False
    
    .CurrentX = SaveX + .TextWidth("ASSISTANT: ")
    .CurrentY = SaveY
    Printer.Print rstStuffForSlips!AssistantName
    
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintSourceA_Sub()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = SourceXPos + UserTweakX
    .CurrentY = SourceYPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters) + _
                UserTweakY + SubSlipCount * PaperHeight
    
    SaveX = .CurrentX
    SaveY = .CurrentY
    
    .FontBold = True
    Printer.Print "SOURCE: "
    .FontBold = False
    
    .CurrentX = SaveX + .TextWidth("SOURCE: ")
    .CurrentY = SaveY
    
    Printer.Print SourceMaterialA
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintSourceB_Sub()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = SourceXPos + UserTweakX + .TextWidth("SOURCE: ")
    .CurrentY = SourceYPos + 0.5 * .ScaleY(SlipFontSize, vbPoints, vbCentimeters) + _
                UserTweakY + SubSlipCount * PaperHeight
    
    Printer.Print SourceMaterialB
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintCounselPoint_Sub()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CounselXPos + UserTweakX
    .CurrentY = CounselYPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters) + _
                UserTweakY + SubSlipCount * PaperHeight
    
    SaveX = .CurrentX
    SaveY = .CurrentY
    
    .FontBold = True
    Printer.Print "COUNSEL POINT: "
    .FontBold = False
    
    .CurrentX = SaveX + .TextWidth("COUNSEL POINT: ")
    .CurrentY = SaveY
    
    Printer.Print rstStuffForSlips!CounselPointNo
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintTalkNo_Sub()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = TalkNoXPos + UserTweakX
    .CurrentY = TalkNoYPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters) + _
                UserTweakY + SubSlipCount * PaperHeight
    
    SaveX = .CurrentX
    SaveY = .CurrentY
    
    .FontBold = True
    Printer.Print "TALK NO: "
    .FontBold = False
    
    .CurrentX = SaveX + .TextWidth("TALK NO: ")
    .CurrentY = SaveY
    
    If rstStuffForSlips!TalkNo = "B" Then
        Printer.Print "BH"
    ElseIf rstStuffForSlips!TalkNo = "S" Then
        Printer.Print "SQ"
    Else
        Printer.Print rstStuffForSlips!TalkNo
    End If
    
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintStudentNote_Sub()

On Error GoTo ErrorTrap

    If rstStuffForSlips!StudentNote <> "" Then
        With Printer
        
        .CurrentX = StudentNoteX + UserTweakX
        .CurrentY = StudentNoteY - .ScaleY(SlipFontSize, vbPoints, vbCentimeters) + _
                    UserTweakY + SubSlipCount * PaperHeight
        
        SaveX = .CurrentX
        SaveY = .CurrentY
        
        .FontBold = True
        
        Printer.Print
        
        .FontBold = False
        
        .CurrentX = SaveX + .TextWidth("NOTE: ")
        .CurrentY = SaveY
        Printer.Print rstStuffForSlips!StudentNote
        
        End With
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintStudentNote2_Sub()

On Error GoTo ErrorTrap

    If rstStuffForSlips!StudentNote2 <> "" Then
        With Printer
        
        .CurrentX = StudentNote2X + UserTweakX
        .CurrentY = StudentNote2Y - .ScaleY(SlipFontSize, vbPoints, vbCentimeters) + _
                    UserTweakY + SubSlipCount * PaperHeight
                        
        .FontBold = False
        
        Printer.Print rstStuffForSlips!StudentNote2
        
        End With
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintSchoolNo_Sub()

On Error GoTo ErrorTrap

    With Printer
    .CurrentY = SchoolNoYPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters) + _
                UserTweakY + SubSlipCount * PaperHeight
    .CurrentX = SchoolNoXPos + UserTweakX
    
    SaveX = .CurrentX
    SaveY = .CurrentY
    
    .FontBold = True
    Printer.Print "SCHOOL: "
    .FontBold = False
    
    .CurrentX = SaveX + .TextWidth("SCHOOL: ")
    .CurrentY = SaveY
    
    Printer.Print rstStuffForSlips!SchoolNo
    
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintSetting_Sub()

On Error GoTo ErrorTrap

    With Printer
    .CurrentY = SettingY - .ScaleY(SlipFontSize, vbPoints, vbCentimeters) + _
                UserTweakY + SubSlipCount * PaperHeight
    .CurrentX = SettingX + UserTweakX
    
    SaveX = .CurrentX
    SaveY = .CurrentY
    
    .FontBold = True
    Printer.Print "SETTING: "
    .FontBold = False
    
    .CurrentX = SaveX + .TextWidth("SETTING: ")
    .CurrentY = SaveY
    
    Printer.Print rstStuffForSlips!Setting
    
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintAssignmentDate_Sub()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = AssignmentDateXPos + UserTweakX
    .CurrentY = AssignmentDateYPos - .ScaleY(SlipFontSize, vbPoints, vbCentimeters) + _
                UserTweakY + SubSlipCount * PaperHeight
    
    SaveX = .CurrentX
    SaveY = .CurrentY
    
    .FontBold = True
    Printer.Print "DATE: "
    .FontBold = False
    
    .CurrentX = SaveX + .TextWidth("DATE: ")
    .CurrentY = SaveY
    
    Printer.Print rstStuffForSlips!AssignmentDate
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintLine_Sub()


On Error GoTo ErrorTrap

    With Printer
    .CurrentX = UserTweakX
    .CurrentY = EndLineY + UserTweakY + SubSlipCount * PaperHeight
    
    Printer.Print "----------------------------------------------------------------" & _
                  "----------------------------------------------------------------" & _
                  "--------------"
    End With
    

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub CheckThemeAndSourceLength_Sub()

On Error GoTo ErrorTrap

    If Printer.TextWidth(rstStuffForSlips!SourceMaterial) > MaxSourceWidth Then
        SourceMaterialA = Left(rstStuffForSlips!SourceMaterial, 35)
        If Len(rstStuffForSlips!SourceMaterial) > 35 Then
            SourceMaterialB = Left(Right(rstStuffForSlips!SourceMaterial, Len(rstStuffForSlips!SourceMaterial) - 35), 35)
        End If
    Else
        SourceMaterialA = rstStuffForSlips!SourceMaterial
        SourceMaterialB = ""
    End If
    
    If Printer.TextWidth(rstStuffForSlips!TalkTheme) > MaxThemeWidth Then
        ThemeA = Left(rstStuffForSlips!TalkTheme, 90)
        If Len(rstStuffForSlips!TalkTheme) > 90 Then
            ThemeB = Left(Right(rstStuffForSlips!TalkTheme, Len(rstStuffForSlips!TalkTheme) - 90), 90)
        End If
    Else
        ThemeA = rstStuffForSlips!TalkTheme
        ThemeB = ""
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


