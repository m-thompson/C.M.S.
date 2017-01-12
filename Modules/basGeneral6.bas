Attribute VB_Name = "basGeneral6"
Option Explicit
Public Function GetFile(TheFileName As String, _
                           TheDialogTitle As String, _
                           FileOpenDialog As CommonDialog, _
                           InitialDir As String, _
                           InitialFileName As String, _
                           FilterString As String, _
                           IncludeALLFilesFilter As Boolean, _
                           DefaultExt As String) As Boolean

Dim SaveCurDir As String

    On Error GoTo ExitNow

    GetFile = False
    
    '
    'Set up dialogue parms
    '
    If FilterString <> "" Then
        FileOpenDialog.Filter = FilterString & _
                                IIf(IncludeALLFilesFilter, "|All Files|*.*", "")
    Else
        FileOpenDialog.Filter = "All Files|*.*"
    End If
        
    FileOpenDialog.FilterIndex = 1
    FileOpenDialog.DefaultExt = DefaultExt
    FileOpenDialog.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist Or _
        cdlOFNNoReadOnlyReturn Or cdlOFNNoChangeDir
    FileOpenDialog.DialogTitle = TheDialogTitle
    FileOpenDialog.InitDir = InitialDir
    
    If InitialFileName <> "" Then
        FileOpenDialog.Filename = InitialFileName
    End If
    
    FileOpenDialog.CancelError = True ' Exit (via raised error) if user presses Cancel.
    
    FileOpenDialog.ShowOpen
    TheFileName = FileOpenDialog.Filename
    GetFile = True

ExitNow:
    Exit Function
    
End Function
Public Function GetPersonNameForSMS(PersonID As Long) As String
Dim str As String, rs As Recordset

    'if listing names for an SMS, we want just first initial with lastname.
    ' But if there's more than one (eg 'M Thompson'), use full first name.

    On Error GoTo ErrorTrap

    str = "SELECT COUNT(*) AS TheNum " & _
          "FROM tblNameAddress " & _
          "WHERE LEFT(FirstName,1) = '" & Left(CongregationMember.GetFirstName(PersonID), 1) & "' " & _
          "AND LastName = '" & CongregationMember.GetLastName(PersonID) & "'"
          
    Set rs = CMSDB.OpenRecordset(str, dbOpenForwardOnly)
    
    If HandleNull(rs!TheNum) <= 1 Then
        GetPersonNameForSMS = CongregationMember.NameWithOneFirstInitial(PersonID)
    Else
        GetPersonNameForSMS = CongregationMember.FirstAndLastName(PersonID)
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
    
End Function
Public Function NumberOfActivePeopleInDB() As Long

Dim str As String, rs As Recordset

    On Error GoTo ErrorTrap

    str = "SELECT COUNT(1) AS num FROM tblNameAddress " & _
            "WHERE Active = TRUE "
          
    Set rs = CMSDB.OpenRecordset(str, dbOpenForwardOnly)
    
    NumberOfActivePeopleInDB = HandleNull(rs!num)
        
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
    
End Function
Public Function NumberOfFieldServiceGroups() As Long

Dim str As String, rs As Recordset

    On Error GoTo ErrorTrap

    str = "SELECT COUNT(1) AS num FROM tblBookGroups "
          
    Set rs = CMSDB.OpenRecordset(str, dbOpenForwardOnly)
    
    NumberOfFieldServiceGroups = HandleNull(rs!num)
        
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
    
End Function
Public Sub CheckSetUp()

On Error GoTo ErrorTrap

    If NumberOfActivePeopleInDB > 0 Then
        GlobalParms.Save "PeopleAreInTheSystem", "TrueFalse", True
        SystemContainsPeople = True
        DealWithCongSetup
        frmMainMenu.EnforceSecurity
    Else
        GlobalParms.Save "PeopleAreInTheSystem", "TrueFalse", False
        CheckIfCongSetUp
        CheckIfBrothersInDB
    End If
    
    Exit Sub
ErrorTrap:
    Call EndProgram
    
End Sub


Public Function GetDNCsForMap(MapNo As Long) As String
Dim rs As Recordset, str As String
Dim rs2 As Recordset, str2 As String
Dim sDNCs As String
On Error GoTo ErrorTrap

    str = "SELECT StreetName, StreetSeqNum " & _
          "FROM tblTerritoryStreets " & _
          "WHERE MapNo = " & MapNo
          
    Set rs = CMSDB.OpenRecordset(str, dbOpenForwardOnly)
    
    Do Until rs.EOF Or rs.BOF
        
        
        str2 = "SELECT HouseNo " & _
                "FROM tblTerritoryDNCs " & _
                "WHERE StreetSeqNum = " & rs!StreetSeqNum & _
                " ORDER BY 1 "
        
        Set rs2 = CMSDB.OpenRecordset(str2, dbOpenForwardOnly)
        
        If Not rs2.BOF Then
        
            sDNCs = sDNCs & rs!StreetName & ": "
        
            With rs2
            
            Do Until .BOF Or .EOF
            
                sDNCs = sDNCs & !HouseNo
            
                .MoveNext
                
                If Not .EOF Then
                    sDNCs = sDNCs & ","
                End If
                
            Loop
            
            End With
            
        End If
        
        rs.MoveNext
        
        If (Not rs.EOF) And (Not rs2.BOF) Then
            sDNCs = sDNCs & vbCrLf
        End If
        
    Loop
    
    GetDNCsForMap = sDNCs
    
    On Error Resume Next
    
    rs.Close
    rs2.Close
    Set rs = Nothing
    Set rs2 = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function

Public Function LoadPicToImageBox(TheImage As Image, _
                             TheFilePath As String, _
                             Optional DesiredHeight_cm As Single, _
                             Optional DesiredWidth_cm As Single, _
                             Optional Auto As Boolean = False, _
                             Optional CentreInContainer As Boolean = True) As Boolean
On Error Resume Next
    
Dim lStoreHeight As Long, lStoreWidth As Long, bFixHeight As Boolean, bFixWidth As Boolean
    
    TheImage.Stretch = False 'image will resize to pic's actual dimensions
    
    TheImage.Picture = LoadPicture(TheFilePath)
    
    TheImage.Stretch = True 'now we can resize the pic, squeezing/stretching it as appropriate
    
    If Err.number <> 0 Then
        LoadPicToImageBox = False
        Exit Function
    End If
    
    If Auto Then
        If TheImage.Height > TheImage.Width Then
            bFixHeight = True
            bFixWidth = False
        ElseIf TheImage.Width > TheImage.Height Then
            bFixHeight = False
            bFixWidth = True
        Else
            bFixHeight = True
            bFixWidth = True
        End If
    End If
        
    lStoreHeight = TheImage.Height
    lStoreWidth = TheImage.Width
    
    If bFixHeight And bFixWidth Then
        TheImage.Width = 567 * DesiredWidth_cm
        TheImage.Height = 567 * DesiredHeight_cm
    Else
        If IsMissing(DesiredHeight_cm) Or bFixWidth Then
            'no height supplied, so it can vary. Fix the width
            TheImage.Width = 567 * DesiredWidth_cm
            TheImage.Height = (567 * TheImage.Height * DesiredWidth_cm / lStoreWidth)
        ElseIf IsMissing(DesiredWidth_cm) Or bFixHeight Then
            'no Width supplied, so it can vary. Fix the Height
            TheImage.Height = 567 * DesiredHeight_cm
            TheImage.Width = (567 * TheImage.Width * DesiredWidth_cm / lStoreHeight)
        Else
            TheImage.Width = 567 * DesiredWidth_cm
            TheImage.Height = 567 * DesiredHeight_cm
        End If
    End If
    
    If CentreInContainer Then
        If TheImage.Container.Height > TheImage.Height Then
            TheImage.Top = (TheImage.Container.Height - TheImage.Height) / 2
        End If
        If TheImage.Container.Width > TheImage.Width Then
            TheImage.Left = (TheImage.Container.Width - TheImage.Width) / 2
        End If
    End If
    
    If Err.number <> 0 Then
        LoadPicToImageBox = False
        Exit Function
    End If
    
    LoadPicToImageBox = True

End Function

Public Sub WriteToLogFile(StringToWrite As String)
On Error Resume Next

    tsLogFileTextStream.WriteLine Now & vbTab & StringToWrite & vbCrLf

End Sub
Public Function MinReportSentToBranch(SocietyReportingPeriod As Date) As Boolean
On Error GoTo ErrorTrap
Dim rs As Recordset, bFlag As Boolean
    
    Set rs = GetGeneralRecordset("SELECT 1 FROM tblReportSentToBranch " & _
                                " WHERE SocietyReportingPeriod = " & GetDateStringForSQLWhere(CStr(SocietyReportingPeriod)))
    
    bFlag = Not rs.BOF
    
    rs.Close
    Set rs = Nothing
    
    MinReportSentToBranch = bFlag

    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function SaveMinReportSentToBranchFlag(SocietyReportingPeriod As Date, Sent As Boolean)
On Error GoTo ErrorTrap
    
Dim rs As Recordset, bFlag As Boolean
    
    Set rs = GetGeneralRecordset("tblReportSentToBranch")
    
    rs.FindFirst "SocietyReportingPeriod = " & GetDateStringForSQLWhere(CDate(SocietyReportingPeriod))
    
    If Sent Then
        If rs.NoMatch Then
            rs.AddNew
        Else
            rs.Edit
        End If
        rs!SocietyReportingPeriod = SocietyReportingPeriod
        rs.Update
    Else
        rs.Delete
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function AddStringListDelimiter(ExistingString As String, StringToAdd As String, Optional TheDelimiter As String = ", ") As String
On Error GoTo ErrorTrap
    
    If ExistingString = "" Then
        AddStringListDelimiter = StringToAdd
    Else
        AddStringListDelimiter = ExistingString & TheDelimiter & StringToAdd
    End If
        
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function PublicTalkIsProvisional(MeetingDate As Date, _
                                        SpeakerID As Long, _
                                        CongWhereMtgIs As Long) As Boolean
                                        
On Error GoTo ErrorTrap
    
Dim rs As Recordset, str As String

    str = "SELECT Provisional " & _
          "FROM tblPublicMtgSchedule " & _
          "WHERE MeetingDate = #" & _
                Format(MeetingDate, "mm/dd/yyyy") & "# " & _
          " AND CongNoWhereMtgIs = " & CongWhereMtgIs & _
          " AND SpeakerID = " & SpeakerID
                 
    Set rs = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    If Not rs.BOF Then
        PublicTalkIsProvisional = rs!Provisional
    Else
        PublicTalkIsProvisional = False
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Sub SetPublicTalkToProvisional(MeetingDate As Date, _
                                    SpeakerID As Long, _
                                    CongWhereMtgIs As Long, _
                                    ProvisionalFlag As Boolean)
                                        
On Error GoTo ErrorTrap
    
Dim rs As Recordset, str As String

    str = "SELECT Provisional " & _
          "FROM tblPublicMtgSchedule " & _
          "WHERE MeetingDate = #" & _
                Format(MeetingDate, "mm/dd/yyyy") & "# " & _
          " AND CongNoWhereMtgIs = " & CongWhereMtgIs & _
          " AND SpeakerID = " & SpeakerID
                 
    Set rs = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    If Not rs.BOF Then
        rs.Edit
        rs!Provisional = ProvisionalFlag
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Public Function BuildCongRolesPrintTable() As Boolean
Dim rstStudents As Recordset, NamesSQL As String, PersonID As Integer
Dim rstPrintTable As Recordset, ErrCode As Integer

Dim i As Long

On Error GoTo ErrorTrap

    DeleteTable "tblPrintCongRoles"
    CreateTable ErrCode, "tblPrintCongRoles", "PersonID", "LONG", , , False
    CreateField ErrCode, "tblPrintCongRoles", "PersonName", "TEXT", "150"
    CreateField ErrCode, "tblPrintCongRoles", "Elder", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "Servant", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "PresidingOverseer", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "Secretary", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "TMSOverseer", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "TMSOverseerAsst", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "ServiceOverseer", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "WTConductor", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "WTConductorAsst", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "CongBibStudyConductor", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "PublicSpeaker", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "OutboundSpeaker", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "WTReader", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "PubMtgChairman", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "CongPrayers", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "SrvMtgAnnouncements", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "SrvMtgItems", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "Accounts", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "AccountsAsst", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "Literature", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "LiteratureAsst", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "LiteratureCoord", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "Magazines", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "MagsAsst", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "Territory", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "AttendantOverseer", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "Attendant", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "RovingMics", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "Sound", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "Platform", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "FireCoordinator", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "FireCoordinatorAsst", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "AuxCllr", "YESNO"
    CreateField ErrCode, "tblPrintCongRoles", "AuxCllrAsst", "YESNO"
    
    '
    'Driving Recset to get all baptised bros
    '
    NamesSQL = "SELECT DISTINCT " & _
               "tblNameAddress.ID " & _
               ",tblNameAddress.FirstName " & _
               ",tblNameAddress.MiddleName " & _
               ",tblNameAddress.LastName " & _
               "FROM tblBaptismDates " & _
               "INNER JOIN tblNameAddress " & _
               "ON tblNameAddress.ID = tblBaptismDates.PersonID " & _
               "WHERE Active = TRUE " & _
               "AND GenderMF = 'M' " & _
               "ORDER BY 4, 2, 3"

    Set rstStudents = CMSDB.OpenRecordset(NamesSQL, dbOpenDynaset)
    Set rstPrintTable = CMSDB.OpenRecordset("tblPrintCongRoles", dbOpenDynaset)
    
    If rstStudents.BOF Then
        BuildCongRolesPrintTable = False
        Exit Function
    End If
    
    With rstPrintTable
    
    Do Until rstStudents.EOF
        PersonID = rstStudents!ID
        .AddNew
        
        !PersonID = PersonID
        !PersonName = rstStudents!LastName & ", " & rstStudents!FirstName & " " & rstStudents!MiddleName
        !PublicSpeaker = CongregationMember.DoesRole2(PersonID, 49)
        !CongPrayers = CongregationMember.DoesCongPrayers(PersonID)
        !WTConductor = CongregationMember.IsWatchtower(PersonID, giGlobalDefaultCong)
        !WTConductorAsst = CongregationMember.DoesRole2(PersonID, 11)
        !PubMtgChairman = CongregationMember.DoesRole2(PersonID, 48)
        !SrvMtgAnnouncements = CongregationMember.DoesRole2(PersonID, 19)
        !SrvMtgItems = CongregationMember.DoesRole2(PersonID, 21)
        !TMSOverseer = CongregationMember.IsSchool(PersonID, giGlobalDefaultCong)
        !TMSOverseerAsst = CongregationMember.DoesRole2(PersonID, 31)
        !CongBibStudyConductor = CongregationMember.DoesRole2(PersonID, 14)
        !WTReader = CongregationMember.DoesRole2(PersonID, 10)
        !Sound = CongregationMember.IsSound(PersonID, giGlobalDefaultCong)
        !Platform = CongregationMember.IsPlatform(PersonID, giGlobalDefaultCong)
        !RovingMics = CongregationMember.IsRovingMic(PersonID, giGlobalDefaultCong)
        !Attendant = CongregationMember.IsAttendant(PersonID, giGlobalDefaultCong)
        !AttendantOverseer = CongregationMember.DoesRole2(PersonID, 96)
        !PresidingOverseer = CongregationMember.DoesRole2(PersonID, 71)
        !Secretary = CongregationMember.DoesRole2(PersonID, 74)
        !ServiceOverseer = CongregationMember.DoesRole2(PersonID, 67)
        !Accounts = CongregationMember.DoesRole2(PersonID, 81)
        !AccountsAsst = CongregationMember.DoesRole2(PersonID, 82)
        !Magazines = CongregationMember.DoesRole2(PersonID, 66)
        !MagsAsst = CongregationMember.DoesRole2(PersonID, 65)
        !Literature = CongregationMember.DoesRole2(PersonID, 63)
        !LiteratureAsst = CongregationMember.DoesRole2(PersonID, 64)
        !LiteratureCoord = CongregationMember.DoesRole2(PersonID, 62)
        !Territory = CongregationMember.DoesRole2(PersonID, 69)
        !OutboundSpeaker = CongregationMember.DoesRole2(PersonID, 93)
        !FireCoordinator = CongregationMember.DoesRole2(PersonID, 77)
        !FireCoordinatorAsst = CongregationMember.DoesRole2(PersonID, 78)
        !Elder = CongregationMember.DoesRole2(PersonID, 51)
        !Servant = CongregationMember.DoesRole2(PersonID, 52)
        !AuxCllr = CongregationMember.DoesRole2(PersonID, 27)
        !AuxCllrAsst = CongregationMember.DoesRole2(PersonID, 28)
        
        .Update
        rstStudents.MoveNext
    Loop
    
    End With
    
    BuildCongRolesPrintTable = True
    
    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Sub PrintCongRolesWithWord()

On Error GoTo ErrorTrap

Dim reporter As MSWordReportingTool2.RptTool

    SwitchOffDAO

    Screen.MousePointer = vbHourglass
    
    Set reporter = New RptTool
    
    With reporter
    
    .DB_PathAndName = CompletePathToTheMDBFileAndExt
    
    .DBPassword = TheDBPassword 'ignorred if the DB doesn't need a password.

    .SaveDoc = True
    .DocPath = gsDocsDirectory & "\" & "Congregation Roles " & _
                                Replace(Replace(Now, ":", "-"), "/", "-")

    
    .ReportSQL = "SELECT PersonName, Elder, Servant, PresidingOverseer, Secretary, " & _
                 "  TMSOverseer, TMSOverseerAsst,AuxCllr,AuxCllrAsst,ServiceOverseer, WTConductor, " & _
                 "  WTConductorAsst, CongBibStudyConductor, PublicSpeaker, OutboundSpeaker, " & _
                 "  WTReader, PubMtgChairman, CongPrayers, SrvMtgAnnouncements, SrvMtgItems, " & _
                 "  Accounts, AccountsAsst, Literature, LiteratureAsst, LiteratureCoord, " & _
                 "  Magazines, MagsAsst, Territory, AttendantOverseer, Attendant, RovingMics, " & _
                 "  Sound, Platform, FireCoordinator, FireCoordinatorAsst " & _
                 "FROM tblPrintCongRoles " & _
                 "ORDER BY 1"

    .ReportTitle = "Congregation Roles - " & Format(Now, "mmmm dd yyyy")
    .TopMargin = 15
    .BottomMargin = 15
    .LeftMargin = 10
    .RightMargin = 10
    .ReportFooterFontName = "Arial"
    .ReportFooterFontSize = 8
    .ReportTitleFontName = "Times New Roman"
    .ReportTitleFontSize = 12
    .ApplyTableFormatting = True
    .ClientName = AppName
    .AdditionalReportHeading = ""
    .GroupingColumn = 0
    .HideWordWhileBuilding = True
    .ShowProgress = True
    
    
    .AddTableColumnAttribute "", 48, , , , , 9, 9, True, , , , True
    .AddTableColumnAttribute "Elder", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Servant", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Coordinator", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Secretary", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "TMSO", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "TMSO Asst", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Aux Clr", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Aux Clr Ast", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Serv O'seer", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "WT O'seer", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "WT Asst", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "B.St Cond", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Publ Spkr", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Vis Spkr", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Reader", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Chairman", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Prayers", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Announce", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Sv Mtg", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Accounts", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Acc Asst", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Lit Serv", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Lit Asst", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Lit Coord", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Magazines", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Mags Asst", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Territory", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Att O'seer", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Attendant", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Rov Mics", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Sound", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Platform", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Fire Coord", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    .AddTableColumnAttribute "Fire Asst", 6, cmsLeftCentre, cmsCentreTop, , , 8, 9, , , , , , "X", cmsUp
    
    .PageFormat = cmsLandscape
    
    .GenerateReport

    End With
    
    SwitchOnDAO
    
    Set reporter = Nothing
    Screen.MousePointer = vbNormal

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

