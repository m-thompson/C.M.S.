VERSION 5.00
Begin VB.Form frmAdjustPubMtgWtgs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " C.M.S. Adjust Weightings for Reader/Chairman Rota"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmAdjustPubMtgWtgs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3960
      TabIndex        =   5
      Top             =   1035
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   2625
      TabIndex        =   4
      Top             =   1035
      Width           =   915
   End
   Begin VB.TextBox txtChairman 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1380
      MaxLength       =   2
      TabIndex        =   3
      ToolTipText     =   "Reduce this figure to increase the rate at which brothers are chairman"
      Top             =   1035
      Width           =   540
   End
   Begin VB.TextBox txtReader 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   345
      MaxLength       =   2
      TabIndex        =   2
      ToolTipText     =   "Reduce this figure to increase the rate at which brothers read."
      Top             =   1035
      Width           =   540
   End
   Begin VB.Label Label3 
      Caption         =   $"frmAdjustPubMtgWtgs.frx":0442
      Height          =   660
      Left            =   300
      TabIndex        =   6
      Top             =   60
      Width           =   4560
   End
   Begin VB.Label Label2 
      Caption         =   "Chairman"
      Height          =   210
      Left            =   1335
      TabIndex        =   1
      Top             =   780
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Reader"
      Height          =   210
      Left            =   345
      TabIndex        =   0
      Top             =   780
      Width           =   900
   End
End
Attribute VB_Name = "frmAdjustPubMtgWtgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

On Error GoTo ErrorTrap

    Unload Me

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdOK_Click()

On Error GoTo ErrorTrap

    If IsNumber(txtReader, False, False, False) Then
        If CLng(txtReader) > 0 Then
            GlobalParms.Save "ReaderWeighting", "NumVal", CLng(txtReader)
        Else
            GoTo Bad
        End If
    Else
        GoTo Bad
    End If
    
    If IsNumber(txtChairman, False, False, False) Then
        If CLng(txtChairman) > 0 Then
            GlobalParms.Save "ChairmanWeighting", "NumVal", CLng(txtChairman)
        Else
            GoTo Bad
        End If
    Else
        GoTo Bad
    End If

    Unload Me

    Exit Sub
    
Bad:
    MsgBox "Invalid Entry", vbOKOnly + vbExclamation, AppName
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub Form_Load()

On Error GoTo ErrorTrap

    txtChairman = GlobalParms.GetValue("ChairmanWeighting", "NumVal")
    txtReader = GlobalParms.GetValue("ReaderWeighting", "NumVal")

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub


Private Sub txtChairman_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorTrap

    KeyPressValid KeyAscii, cmsUnsignedIntegers

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub


Private Sub txtReader_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorTrap

    KeyPressValid KeyAscii, cmsUnsignedIntegers

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub
