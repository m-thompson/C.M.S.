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
   Begin VB.TextBox Text1 
      Height          =   450
      Left            =   1005
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1665
      Width           =   3285
   End
   Begin VB.ComboBox cmbNames2 
      Height          =   315
      Left            =   495
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   795
      Width           =   3630
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AutoCompleteCombo(Combo1 As ComboBox, KeyAscii As Integer)
    Dim lCnt As Long 'Generic long counter
    Dim lMax As Long
    Dim sComboItem As String
    Dim sComboText As String 'Text currently in combobox
    Dim sText As String 'Text after keypressed
    Dim miSelStart As Integer
    'MsgBox KeyAscii
     
    miSelStart = Combo1.SelStart 'Set the entry point of the selected text
    If KeyAscii = 13 Or KeyAscii = 27 Then Exit Sub 'Charage Return or escape
     
    With Combo1

        lMax = .ListCount - 1 'The number of values in the list
        sComboText = .text 'The current text in the list
         
        '**********************************************
        'This block either removes or adds a character
        '***********************************************
        If KeyAscii = 8 Then 'Backspace
            'If there is one character left, remove the text
            If Len(Combo1.text) = 1 Or Len(Combo1.text) = 0 Then
                Combo1.text = ""
                miSelStart = 0
                Exit Sub
            'There is one character and seltext (our matched text)
            ElseIf Len(Combo1.text) - (Len(Combo1.text) - miSelStart) = 1 Then
                Combo1.text = ""
                miSelStart = 0
                Exit Sub
            End If
             
            'An ordinary backspace, decrement 1 character
            Combo1.text = Left(sComboText, miSelStart - 1)
            sText = Left(sComboText, miSelStart - 1) 'reset our sText var sText = combo1.text??
        Else
            'A char other than backspace, return or escape was pressed
            sText = Left(sComboText, miSelStart) & Chr(KeyAscii) 'Increment our string with the new char
        End If

        KeyAscii = 0 'Reset key pressed
         
        '**********************************************
        'This block performs our lookup
        '**********************************************
        For lCnt = 0 To lMax
            sComboItem = .List(lCnt) 'Current item
             
            If UCase(sText) = UCase(Left(sComboItem, _
                                         Len(sText))) Then 'A match was found
                '.ListIndex = lCnt 'Not sure why this is needed
                .text = sComboItem
                .SelStart = Len(sText) 'Start the highlighting after manually entered text
                .SelLength = Len(sComboItem) - (Len(sText)) 'Highlight to the end of the text
                 
                Exit For
            Else
                .text = sText 'Set the text value = the new text
                .SelStart = Len(.text) 'Set selstart to the end of the string
            End If
             
        Next 'lCnt
    End With


End Sub

Private Sub cmbNames2_Change()

End Sub

Private Sub cmbNames2_KeyPress(KeyAscii As Integer)
    AutoCompleteCombo Me!cmbNames2, KeyAscii
End Sub

Private Sub Form_Load()
    HandleListBox.PopulateListBox Me!cmbNames2, _
        "SELECT DISTINCTROW tblNameAddress.ID, " & _
        "tblNameAddress.FirstName & ' ' & tblNameAddress.MiddleName, " & _
        "tblNameAddress.LastName " & _
        "FROM tblNameAddress INNER JOIN tblTaskAndPerson " & _
        "ON tblNameAddress.ID = tblTaskAndPerson.Person " & _
        "WHERE Active = TRUE " & _
        "AND CongNo = " & GlobalDefaultCong & _
        " ORDER BY tblNameAddress.LastName, tblNameAddress.FirstName" _
        , CMSDB, 0, ", ", True, 2, 1

End Sub

