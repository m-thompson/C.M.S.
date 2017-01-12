VERSION 5.00
Begin VB.Form frmAddEventType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " C.M.S. Add/Edit Event Type"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frmAddEventType.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2415
      TabIndex        =   2
      Top             =   765
      Width           =   1020
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1245
      TabIndex        =   1
      Top             =   765
      Width           =   1020
   End
   Begin VB.TextBox txtEventType 
      Height          =   315
      Left            =   180
      MaxLength       =   50
      TabIndex        =   0
      Top             =   240
      Width           =   4260
   End
End
Attribute VB_Name = "frmAddEventType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbUpdateMode As Boolean, FormInit As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrorTrap
Dim rstTemp As Recordset

    If mbUpdateMode Then
        If frmEventTypeManager.lstEventTypes.ListIndex > -1 Then
            CMSDB.Execute "UPDATE tblEventLookup " & _
                          "SET EventName = '" & Trim(DoubleUpSingleQuotes(txtEventType.text)) & _
                          "' WHERE EventID = " & frmEventTypeManager.lstEventTypes.ItemData( _
                                                frmEventTypeManager.lstEventTypes.ListIndex)
        End If
    Else
        Set rstTemp = CMSDB.OpenRecordset("SELECT MAX(EventID) as MaxID " & _
                                          "FROM tblEventLookup", dbOpenSnapshot)
                                          
        CMSDB.Execute "INSERT INTO tblEventLookup " & _
                      "(EventID, EventName, ShowInCalendar) " & _
                      "VALUES (" & rstTemp!MaxID + 1 & ", '" & _
                      Trim(DoubleUpSingleQuotes(txtEventType.text)) & _
                      "', True" & ")"
    End If
            
    Unload Me
    
    Exit Sub
    
ErrorTrap:
    EndProgram
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrorTrap

    FormInit = True
    
    If mbUpdateMode Then
        txtEventType.text = frmEventTypeManager.lstEventTypes.text
    End If
    
    cmdOK.Enabled = False
    
    FormInit = False
            
    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub

Public Property Get UpdateMode() As Boolean
    UpdateMode = mbUpdateMode
End Property

Public Property Let UpdateMode(ByVal vNewValue As Boolean)
    mbUpdateMode = vNewValue
End Property

Private Sub txtEventType_Change()
On Error GoTo ErrorTrap

    If FormInit Then Exit Sub
    
    cmdOK.Enabled = True
            
    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub
