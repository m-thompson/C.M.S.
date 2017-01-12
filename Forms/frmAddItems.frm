VERSION 5.00
Begin VB.Form frmAddItems 
   Caption         =   "C.M.S. Add School Items"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAssignmentDate 
      Height          =   315
      Left            =   570
      TabIndex        =   5
      Top             =   1125
      Width           =   1380
   End
   Begin VB.TextBox txtTalkNo 
      Height          =   315
      Left            =   2610
      TabIndex        =   4
      Top             =   1125
      Width           =   525
   End
   Begin VB.TextBox txtTheme 
      Height          =   585
      Left            =   555
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1890
      Width           =   3450
   End
   Begin VB.TextBox txtSource 
      Height          =   315
      Left            =   4815
      TabIndex        =   2
      Top             =   1110
      Width           =   1380
   End
   Begin VB.TextBox txtDifficulty 
      Height          =   315
      Left            =   4800
      TabIndex        =   1
      Top             =   1860
      Width           =   525
   End
   Begin VB.CheckBox chkBroOnly 
      Caption         =   "Brother Only?"
      Height          =   315
      Left            =   4800
      TabIndex        =   0
      Top             =   2430
      Width           =   1800
   End
   Begin VB.Label Label3 
      Caption         =   "Assignment Date"
      Height          =   255
      Left            =   570
      TabIndex        =   10
      Top             =   900
      Width           =   1845
   End
   Begin VB.Label Label4 
      Caption         =   "TalkNo"
      Height          =   255
      Left            =   2610
      TabIndex        =   9
      Top             =   900
      Width           =   1845
   End
   Begin VB.Label Label5 
      Caption         =   "Theme"
      Height          =   255
      Left            =   555
      TabIndex        =   8
      Top             =   1665
      Width           =   1845
   End
   Begin VB.Label Label6 
      Caption         =   "Source Material"
      Height          =   255
      Left            =   4815
      TabIndex        =   7
      Top             =   885
      Width           =   1845
   End
   Begin VB.Label Label7 
      Caption         =   "Difficulty Rating"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   1650
      Width           =   1845
   End
End
Attribute VB_Name = "frmAddItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
