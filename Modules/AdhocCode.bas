Attribute VB_Name = "AdhocCode"
Option Explicit


Public Sub FixMeetingTypes()

    CMSDB.Execute "UPDATE tblMeetingAttendance " & _
                  "SET MeetingTypeID = 5 WHERE MeetingTypeID = 0 AND WeekBeginning > #04/30/2007#"
    CMSDB.Execute "UPDATE tblMeetingAttendance " & _
                  "SET MeetingTypeID = 6 WHERE MeetingTypeID = 1 AND WeekBeginning > #04/30/2007#"
    CMSDB.Execute "UPDATE tblMeetingAttendance " & _
                  "SET MeetingTypeID = 7 WHERE MeetingTypeID = 2 AND WeekBeginning > #04/30/2007#"
    CMSDB.Execute "UPDATE tblMeetingAttendance " & _
                  "SET MeetingTypeID = 8 WHERE MeetingTypeID = 3 AND WeekBeginning > #04/30/2007#"

    CMSDB.Execute "UPDATE tblMeetingAttendance " & _
                  "SET MeetingTypeID = 2 WHERE MeetingTypeID = 5 AND WeekBeginning > #04/30/2007#"
    CMSDB.Execute "UPDATE tblMeetingAttendance " & _
                  "SET MeetingTypeID = 3 WHERE MeetingTypeID = 6 AND WeekBeginning > #04/30/2007#"
    CMSDB.Execute "UPDATE tblMeetingAttendance " & _
                  "SET MeetingTypeID = 0 WHERE MeetingTypeID = 7 AND WeekBeginning > #04/30/2007#"
    CMSDB.Execute "UPDATE tblMeetingAttendance " & _
                  "SET MeetingTypeID = 1 WHERE MeetingTypeID = 8 AND WeekBeginning > #04/30/2007#"

End Sub
