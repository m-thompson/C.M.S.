VERSION 5.00
Begin VB.Form frmAccountOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.M.S. Accounts Options"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmAccountOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkMoveFocusToTranDesc 
      Caption         =   "Move focus to Description Field on Transaction Entry form."
      Height          =   435
      Left            =   210
      TabIndex        =   1
      ToolTipText     =   "After selecting Transaction Code, determines whether focus shifts to Description or skips past it"
      Top             =   150
      Width           =   2805
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   300
      Left            =   3960
      TabIndex        =   0
      Top             =   225
      Width           =   660
   End
End
Attribute VB_Name = "frmAccountOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mbDoNotTrigger As Boolean

Private Sub chkMoveFocusToTranDesc_Click()
On Error GoTo ErrorTrap

    If mbDoNotTrigger Then Exit Sub

    If chkMoveFocusToTranDesc.value = vbChecked Then
        GlobalParms.Save "NewTransaction_SkipDesc", "TrueFalse", False
        gbNewTransaction_SkipDesc = False
    Else
        GlobalParms.Save "NewTransaction_SkipDesc", "TrueFalse", True
        gbNewTransaction_SkipDesc = True
    End If
    
    ShowMessage "Saved", 1000, Me


    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo ErrorTrap

    mbDoNotTrigger = True

    chkMoveFocusToTranDesc.value = IIf(GlobalParms.GetValue("NewTransaction_SkipDesc", "TrueFalse", False), vbUnchecked, vbChecked)
    
    
    mbDoNotTrigger = False

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub
