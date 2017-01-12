Attribute VB_Name = "basGlobalVariables"
Option Explicit

'
'CONSTANTS
'
Public Const WORD_APP = "Word.Application"
Public Const WEEK_OF_NEW_MTG_ARRANGEMENT = "29/12/2008"
Public Const WEEK_OF_NEW_CLAM_MTG_ARRANGEMENT = "04/01/2016"
Public Const MAX_DATE = "31/12/9999"
Public Const MAX_DATE_US = "12/31/9999"

Public TheMDBFile As String, TheDBPassword As String
Public LoginSucceeded As Boolean
Public gCurrentUserID As String, gCurrentPassword As String, gCurrentUserCode As Long
Public gWindowsUsername As String
Public gdActiveFromDate As String, gdActiveToDate As String
Public gCurrentAccessLevel As Long
Public gbResetPassword As Boolean, gbImportedResetPassword As Boolean
Public gAppRunDate As Date
Public gbWindowsLogon As Boolean

Public gFSO As FileSystemObject

'log file handled by this object
Public tsLogFileTextStream As TextStream


Public gbShowMsgBox As Boolean, gbSuppressMsg As Boolean
Public gbAutoSelectPersonIfOnlyMatch As Boolean

Public gsAccountTransferTranCode As String

Public gstrAppVersion As String
Public gstrDBVersion As String
Public GlobalParms As clsApplicationConstants
Public GlobalCalendar As clsCalendar
Public CongregationMember As clsCongregationMember
Public TheTMS As clsTMS
Public rstOneDayContents As Recordset 'Applies to calendar functions (frmCalendar)
Public rstTMSItems As Recordset 'For frmTMSItems
Public rstTMSSchedule As Recordset, rstTMSStudents As Recordset 'for frmTMSScheduling
Public rstTMSInsertStudent As Recordset
Public TheMDBFileAndExt As String, CompletePathToTheMDBFileAndExt As String, JustTheDirectory As String
Public gsLogPath As String
Public gsDocsDirectory As String
Public CompletePathToCopyDB As String
Public SystemContainsPeople As Boolean, CongregationIsSetUp As Boolean
Public NextCMSDBSeqNo As Integer
Public gbHandleForeignChars As Boolean

Public glNoMonthsForMinReporting As Long
Public gbNewTransaction_SkipDesc As Boolean

Public glMaxResultRows As Long

Public gbAllowHighlightsInSch2 As Boolean
Public gbAllowHighlightsInSch3 As Boolean

Public glTMSAssistantRepeatMonths As Long
Public glTMSMaxNoMonthsForSistersTalks As Long

Public gbPrintSlipsForOralReviewReader As Boolean

Public glMidWkMtgDay As Long
Public glSundayMtgDay As Long

'
'TMSGridDoubleClicked is used to correct what I believe to be a bug in the flexgrid:
' If double-clicking a flexgrid opens another form with a flexgrid, the 2nd grid
' has its click event triggered if it's in the same position on the screen as the
' location of the mouse's double-click.
'
Public TMSGridDoubleClicked As Boolean

Public lnkTMSDraftSchedule As Boolean

'
'Following refer to columns on frmTMSInsertStudent.flxInsertStudent
'
Public PrevPrayer As Long, NextPrayer As Long, PrevNo1 As Long, NextNo1 As Long
Public PrevSQ As Long, NextSQ As Long, PrevBH As Long, NextBH As Long
Public PrevReview As Long, NextReview As Long
Public PrevNo2 As Long, NextNo2 As Long, PrevNo3 As Long, NextNo3 As Long
Public PrevNo4 As Long, NextNo4 As Long, PrevNo2School As Long, NextNo2School As Long
Public PrevNo1School As Long, NextNo1School As Long
Public PrevNo3School As Long, NextNo3School As Long, PrevNo4School As Long
Public NextNo4School As Long, PrevAsst As Long, NextAsst As Long
Public PrevAsstSchool As Long, NextAsstSchool As Long, WeeksToCheck As Long
Public PrevBR As Long, NextBR As Long, PrevBRSchool As Long, NextBRSchool As Long
Public PrevIC As Long, NextIC As Long, PrevICSchool As Long, NextICSchool As Long
Public PrevRV As Long, NextRV As Long, PrevRVSchool As Long, NextRVSchool As Long
Public PrevBS As Long, NextBS As Long, PrevBSSchool As Long, NextBSSchool As Long
Public PrevO As Long, NextO As Long, PrevOSchool As Long, NextOSchool As Long
Public TMS_BR_Weighting As Double
Public TMS_IC_Weighting As Double
Public TMS_RV_Weighting As Double
Public TMS_BS_Weighting As Double
Public TMS_O_Weighting As Double
Public TMSAsstWeighting As Double
Public GlobalDefaultCong As Long, giGlobalDefaultCong As Integer
Public PrevTalkColour As Long, CanNotShow As Boolean
Public TMSPrayerWeighting As Double, NextTalkColour As Long
Public PrevAsstColour As Long
Public NextAsstColour As Long
Public TMSSQWeighting As Double
Public TMSNo1Weighting As Double
Public TMSBHWeighting As Double
Public TMSNo2Weighting As Double
Public TMSNo3Weighting As Double
Public TMSNo4Weighting As Double
Public TMSReviewReaderWeighting As Double
Public TMSNo1Weighting_2009 As Double
Public TMSNo2Weighting_2009 As Double
Public TMSNo3Weighting_2009 As Double
Public TMSBibleReadingWeighting_2016 As Double
Public TMSInitialCallWeighting_2016 As Double
Public TMSReturnVisitWeighting_2016 As Double
Public TMSBibleStudyWeighting_2016 As Double
Public TMSOtherWeighting_2016 As Double
Public TMSWeightingIfAssistantOnly As Double
Public TMSScaleNo4BroWeightings As Double
Public gbMultiUserMode As Boolean

'
'Used in calculating TMS weightings
'
Public glTMSAltAsstStuWtg As Double
Public gbTMSAltAsstStu As Boolean


Enum CMSSubcategories
    TMS_Subcat = 6
End Enum

Public Type TMSScheduleRecord
    TMSUDT_AssignmentDate As Date
    TMSUDT_BHSource As String
    TMSUDT_SongNo As Long
    TMSUDT_SQTheme As String
    TMSUDT_SQSource As String
    TMSUDT_No1Theme As String
    TMSUDT_No1Source As String
    TMSUDT_No2Theme As String
    TMSUDT_No2Source As String
    TMSUDT_No3Theme As String
    TMSUDT_No3Source As String
    TMSUDT_No4Theme As String
    TMSUDT_No4Source As String
    TMSUDT_BroOnlyForNo2 As Boolean
    TMSUDT_BroOnlyForNo3 As Boolean
    TMSUDT_BroOnlyForNo4 As Boolean
    TMSUDT_OralReview As String
    TMSUDT_COVisit As Boolean
End Type

Public Type TheTMSPersonAndDate
    TheAssignmentDate As Date
    ThePersonID As Long
    TheSchool As Long
End Type

Public Type StartAndEndDate
    StartDate As Date
    EndDate As Date
End Type

Public Type TransactionDetails
    TransactionID As Long
    TransactionCodeID As Long
    TransactionCode As String
    InOutTypeID As Long
    InOutID As Long
    InOutTypeDescription As String
    InOutDescription As String
    TransactionDate As Date
    Amount As Double
    TransactionDescription As String
    TransactionTypeDescription As String
    RefNo As Long
    AutoDayOfMonth As Long
    OnReceipt As Boolean
    BookGroupNo As Long
    TransactionSubTypeID As Long
    AccountID As Long
    TfrAccountID As Long
    Suppressed As Boolean
End Type

Public Type TMSPrevNextTalkDates
    PersonID As Long
    BaseAssignmentDate As Date
    LastPrayerDate As Date
    NextPrayerDate As Date
    LastSQDate As Date
    NextSQDate As Date
    LastNo1Date As Date
    NextNo1Date As Date
    LastBHDate As Date
    NextBHDate As Date
    LastReviewDate As Date
    NextReviewDate As Date
    LastNo2Date As Date
    NextNo2Date As Date
    LastNo3Date As Date
    NextNo3Date As Date
    LastNo4Date As Date
    NextNo4Date As Date
    LastAsstDate As Date
    NextAsstDate As Date
End Type

Public Type PublicMeetingDetails
    MeetingDate As Date
    SpeakerA As Long
    SpeakerB As Long
    CongWhereMtgIs As Long
    TalkCoordinator As Long
    TalkNo As Long
    Chairman As Long
    Reader As Long
    Info As String
    Provisional As Boolean
End Type

Public Type TMS_ThemeAndSource
    MeetingDate As Date
    Theme As String
    source As String
    ThemeAndSource As String
End Type

Public Enum SecurityAccessLevels
    CompleteAccess = 1
    SPAMPrinting = 2
    TMSPrinting = 3
    TMSOverseer = 4
    GeneralAdmin = 5
End Enum

Public Enum MinistryType
    IsPublisher = 1
    IsAuxPio = 2
    IsRegPio = 3
    IsSpecPio = 4
End Enum

Public Enum cmsOfficeAppConstants
    cmsWord = 0
    cmsAccess = 1
    cmsExcel = 2
    cmsPowerpoint = 3
    cmsOutlook = 4
End Enum

Public Enum cmsListSelection
    cmsSelectNone = 0
    cmsSelectAll = 1
End Enum

Public Enum cmsTextEntryTypes
    cmsUnsignedIntegers = 0
    cmsUnsignedDecimals
    cmsSignedIntegers
    cmsSignedDecimals
    cmsDates
    cmsTimes
    cmsAlphabetic
    cmsAlphaNumeric
    cmsAlphaNumericPunctuation
End Enum

Public Enum cmsUpdateModes
    cmsView = 0
    cmsEdit
    cmsAdd
    cmsDelete
End Enum

Public Enum cmsEmailTypes
    cmsGeneral = 0
    cmsMissingReport
End Enum

Public Enum cmsPrinterOrientation
    cmsPortrait = 1
    cmsLandscape = 2
End Enum

Public Enum cmsJobType
    cmsElder = 0
    cmsMinisterialServant
    cmsAnnouncements
    cmsServiceMtgPrayer
    cmsServiceMtgItems
End Enum

Public Enum cmsPublicMtgPersonnelTypes
    cmsChairman = 0
    cmsLocalPublicSpeaker
    cmsWatchtowerReader
    cmsOutboundPublicSpeaker
End Enum

Public Enum cmsAttendantTypes
    cmsSoundAtt = 1
    cmsPlatformAtt
    cmsAttendantAtt
    cmsMicrophonesAtt
End Enum

Public Enum cmsMeetingTypes
    cmsSundayMtg = 1
    cmsMidWkMtg
    cmsPublicTalk
    cmsWatchtower
    cmsBookstudy
    cmsServiceMtg
    cmsTMS

End Enum

Public Enum cmsPubFiguresType
    NoHours = 1
    NoBooks
    NoBrochures
    NoRVs
    NoStudies
    NoMags
    NoTracts
    SocietyReportingPeriod
    Remarks
    OtherComments
End Enum

Public Enum cmsPrintUsingWord
    cmsUseWord = 1
    cmsUseMSDatareport
    cmsDontPrint
End Enum

Public Enum eSpecialFolders
  SpecialFolder_AppData = &H1A        'for the current Windows user, on any computer on the network [Windows 98 or later]
  SpecialFolder_CommonAppData = &H23  'for all Windows users on this computer [Windows 2000 or later]
  SpecialFolder_LocalAppData = &H1C   'for the current Windows user, on this computer only [Windows 2000 or later]
  SpecialFolder_Documents = &H5       'the Documents folder for the current Windows user
End Enum

Public Enum MidweekMtgVersion
    Pre2009 = 1
    TMS2009
    CLM2016
End Enum
    
    


