VERSION 5.00
Begin VB.Form frmAccountInOutTypeSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.M.S. Account In/Out Type Select"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "frmAccountInOutTypeSelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   4185
      TabIndex        =   1
      Top             =   1410
      Width           =   630
   End
   Begin VB.ListBox lstTypes 
      Height          =   1230
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   4740
   End
End
Attribute VB_Name = "frmAccountInOutTypeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event InOutTypeSelected(InOutID As Long, InOutTypeID As Long, TranCodeID As Long)
Dim msTranCode As String, mlInOutID As Long, mlInOutTypeID As Long

Private Sub cmdOK_Click()
Dim str As String, trn As TransactionDetails, i As Long
On Error GoTo ErrorTrap

    If lstTypes.ListIndex = -1 Then Exit Sub

    trn = GetTransactionCodeStuff(lstTypes.ItemData(lstTypes.ListIndex))
    RaiseEvent InOutTypeSelected(trn.InOutID, trn.InOutTypeID, trn.TransactionCodeID)
    
    Unload Me

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub Form_Load()
Dim str As String, trn As TransactionDetails, i As Long
On Error GoTo ErrorTrap
    
    str = "SELECT TranCodeID, Description " & _
            "FROM tblTransactionTypes " & _
            "WHERE TranCode = '" & msTranCode & "'" & _
            " AND Suppressed = FALSE " & _
           " ORDER BY Description "
            
    HandleListBox.PopulateListBox lstTypes, str, CMSDB, 0, "", False, 1
    
    If lstTypes.ListCount = 0 Then
        RaiseEvent InOutTypeSelected(0, 0, 0)
        Unload Me
    ElseIf lstTypes.ListCount = 1 Then
        trn = GetTransactionCodeStuff(lstTypes.ItemData(0))
        RaiseEvent InOutTypeSelected(trn.InOutID, trn.InOutTypeID, trn.TransactionCodeID)
        Unload Me
    Else
        For i = 0 To lstTypes.ListCount - 1
            trn = GetTransactionCodeStuff(lstTypes.ItemData(i))
            Select Case trn.InOutTypeID
            Case 1, 5
                lstTypes.ListIndex = i
                Exit For
            End Select
        Next i
    End If
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Public Property Get TranCode() As String
    TranCode = msTranCode
End Property

Public Property Let TranCode(ByVal vNewValue As String)
    msTranCode = vNewValue
End Property

Private Sub lstTypes_DblClick()
On Error GoTo ErrorTrap

    cmdOK_Click

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
