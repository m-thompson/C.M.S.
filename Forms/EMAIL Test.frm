VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1260
      Left            =   1605
      TabIndex        =   0
      Top             =   900
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Emailer As New SendSMS.EmailTool

Private Sub Command1_Click()

    With Emailer
    
    .AddEmailAddress "mjt@orange.net"
    .AddEmailAddress "michael.j.thompson@serco.com"

    .SendEmail "Test", "This is a test"
    
    End With
    
    Set Emailer = Nothing

End Sub
