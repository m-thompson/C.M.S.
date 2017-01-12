Attribute VB_Name = "basLinkage"
Option Explicit

Public lnkCongNo As Integer 'Pass CongNo to screen
Public lnkPersonID As Integer 'Pass Person ID to screen
Public lnkLimitTaskSubCatsSQL As String 'Piece of SQL to limit Task Sub Categories (eg "AND TaskSubCat IN .....")
Public lnkLimitTaskCatsSQL As String 'Piece of SQL to limit Task Categories
Public lnkLimitTaskSubCatsSQL2 As String 'Piece of SQL to limit Task Sub Categories (eg "AND TaskSubCat IN .....")
Public lnkLimitTaskCatsSQL2 As String 'Piece of SQL to limit Task Categories
Public lnkPersDtlsFormIsOpen As Boolean
Public lnkNameFormatForPrint As Byte
Public lnkDateFormatForPrint As Byte
Public TMSUpdateItem As Boolean, TMSAddItem As Boolean
Public lnkTMSAssignmentDate As Date, lnkTMSTalkNo As String, lnkTMSTheme As String
Public lnkTMSSource As String, lnkTMSDifficulty As Byte, lnkTMSBroOnly As Boolean
Public lnkTMSSeqNo As Long
Public lnkTMSInsertStudentFormActive As Boolean
Public lnkTMSAddItemsForm_IsActive As Boolean, lnkTMSAddItemsAssignmentDate As Date
Public lnkTMSPrintingFormOpen As Boolean
Public lnkTMSScheduleSearch() As TheTMSPersonAndDate
Public lnkTMSScheduleSearchStarted As Boolean
Public lnkTMSCounselPointsFormIsOpen_CurrentPoint As Boolean
Public lnkTMSCounselPointsFormIsOpen_NextPoint As Boolean
Public lnkNoOfSchools As Integer
Public lnkTMSSchedulingFormOpen As Boolean
Public lnkTMSCounselPointsFormIsOpen As Boolean
Public lnkTMSScheduleName As String
Public lnkSetUpMenuFormOpen As Boolean
Public lnkTMSStudentDetailsOpen As Boolean
Public lnkSuspendFormOpen1 As Boolean
Public lnkSuspendFormOpen2 As Boolean
Public lnkTMSOptionsOpen As Boolean
Public lnkTMSMenuOpen As Boolean
Public lnkSPAMFormOpen As Boolean
Public lnkTMSAdvancedCounselUpdateOpen As Boolean
Public lnkPubDatesFormOpen1 As Boolean
Public lnkPubDatesFormOpen2 As Boolean
