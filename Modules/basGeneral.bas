Attribute VB_Name = "basGeneral"
Option Explicit
Public Function SpecialFolder(pFolder As eSpecialFolders) As String
'Returns the path to the specified special folder (AppData etc)

Dim objShell  As Object
Dim objFolder As Object

  Set objShell = CreateObject("Shell.Application")
  Set objFolder = objShell.Namespace(CLng(pFolder))

  If (Not objFolder Is Nothing) Then SpecialFolder = objFolder.Self.Path

  Set objFolder = Nothing
  Set objShell = Nothing

  If SpecialFolder = "" Then Err.Raise 513, "SpecialFolder", "The folder path could not be detected"

End Function

Public Sub FlexGridToColour(TheGrid As MSFlexGrid, _
                           Optional lRGB As Long = 0)
                           
Dim i As Integer, j As Integer
    On Error GoTo ErrorTrap

    With TheGrid
        
    For i = 0 To .Rows - 1
        .Row = i
        For j = 0 To .Cols - 1
            .col = j
            .CellForeColor = QBColor(0)
        Next j
    Next i
    
    End With
    
    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub

Public Sub SelectFlexGridRow(TheGrid As MSFlexGrid, Optional StartAtCol As Long = 1)

On Error GoTo ErrorTrap
    
    With TheGrid
    
    If .RowSel <> .Row Then
        .RowSel = .Row
    End If
    
    Select Case .Row
    Case 0
        .HighLight = flexHighlightNever
    Case Else
        .HighLight = flexHighlightAlways
        .col = StartAtCol
        .ColSel = .Cols - 1
    End Select
        
    End With
    

    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub


Public Function KeyPressValid(KeyAscii As Integer, _
                              TextEntryType As cmsTextEntryTypes, _
                              Optional ModifyKeyAscii As Boolean = True) As Boolean

    Select Case TextEntryType
    Case cmsUnsignedIntegers
    'Must be numeric. Allow Backspace (8)
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
            KeyPressValid = False
            If ModifyKeyAscii Then KeyAscii = 0
        Else
            KeyPressValid = True
        End If
    Case cmsUnsignedDecimals
    'Must be numeric. Allow Backspace (8) and full stop (46)
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 Then
            KeyPressValid = False
            If ModifyKeyAscii Then KeyAscii = 0
        Else
            KeyPressValid = True
        End If
    Case cmsDates
    'Must be numeric. Allow Backspace (8) and forward-slash (47).
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 47 Then
            KeyPressValid = False
            If ModifyKeyAscii Then KeyAscii = 0
        Else
            KeyPressValid = True
        End If
    Case cmsTimes
    'Must be numeric. Allow Backspace (8) and colon (58).
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 58 Then
            KeyPressValid = False
            If ModifyKeyAscii Then KeyAscii = 0
        Else
            KeyPressValid = True
        End If
    Case cmsAlphabetic
    'Must be alpha. Allow Backspace (8).
        If ((KeyAscii >= 65 And KeyAscii <= 90) Or _
            (KeyAscii >= 97 And KeyAscii <= 122)) Or KeyAscii = 8 Then
            KeyPressValid = True
        Else
            If ModifyKeyAscii Then KeyAscii = 0
            KeyPressValid = False
        End If
    Case cmsAlphaNumeric
    'Must be alpha/number. Allow Backspace (8).
        If ((KeyAscii >= 65 And KeyAscii <= 90) Or _
            (KeyAscii >= 97 And KeyAscii <= 122) Or _
            (KeyAscii >= 48 And KeyAscii <= 57)) Or KeyAscii = 8 Then
            KeyPressValid = True
        Else
            If ModifyKeyAscii Then KeyAscii = 0
            KeyPressValid = False
        End If
    Case cmsAlphaNumericPunctuation
    'Must be alpha/number. Allow Backspace (8).
        If ((KeyAscii >= 65 And KeyAscii <= 90) Or _
            (KeyAscii >= 97 And KeyAscii <= 122) Or _
            (KeyAscii >= 48 And KeyAscii <= 57)) Or KeyAscii = 8 _
            Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 45 Or KeyAscii = 44 Or KeyAscii = 95 Then
            KeyPressValid = True
        Else
            If ModifyKeyAscii Then KeyAscii = 0
            KeyPressValid = False
        End If
    Case Else
        KeyPressValid = True
    End Select

End Function

Public Sub RemoveAllItemsFromCollection(ByRef TheCollection)

Dim i As Long

    For i = 1 To TheCollection.Count
        TheCollection.Remove 1
    Next i
    
End Sub
Public Sub LimitTextFieldNumber(TheTextBox As TextBox, _
                                KeyAscii As Integer, _
                                Min As Long, Max As Long)

    If KeyAscii > 0 Then
        If IsNumber(Chr(KeyAscii), False, False, False) Then
            Select Case CLng(TheTextBox.text & Chr(KeyAscii))
            Case Min To Max
                'OK
            Case Else
                KeyAscii = 0
            End Select
        End If
    End If
    
End Sub
Public Sub SelectAllInListBox(ByRef TheListBox As ListBox, SelectWhat As cmsListSelection)
    Dim i As Long, saveIndex As Long, saveTop As Long
    
    If TheListBox.MultiSelect = True Or TheListBox.Style = 1 Then
        ' Save current state.
        saveIndex = TheListBox.ListIndex
        saveTop = TheListBox.TopIndex
        ' Make the list box invisible to avoid flickering.
        TheListBox.Visible = False
        ' Change the select state for all items.
        For i = 0 To TheListBox.ListCount - 1
            TheListBox.Selected(i) = SelectWhat
        Next
        ' Restore original state, and make the list box visible again.
        TheListBox.TopIndex = saveTop
        TheListBox.ListIndex = saveIndex
        TheListBox.Visible = True
    End If
    
End Sub

Function GetListBoxRowFromTwips(TheListBox As ListBox, Y_Twips_Pos As Single) As Long

Dim TheRow As Long

    If TheListBox.ListCount > 0 Then
        If Y_Twips_Pos > 0 Then
            TheRow = Ceiling((Y_Twips_Pos / _
                            (GetListItemHeightInTwips(TheListBox)))) + _
                                                              TheListBox.TopIndex - 1
                                                                      
            If TheRow > TheListBox.ListCount - 1 Then
                TheRow = TheListBox.ListCount - 1
            End If
            
            GetListBoxRowFromTwips = TheRow
        Else
            GetListBoxRowFromTwips = 0
        End If
    Else
        GetListBoxRowFromTwips = -1
    End If
                                                              
End Function
Function GetDateStringForSQLWhere(Date_UK As String) As String

    GetDateStringForSQLWhere = " #" & Format$(Date_UK, "mm/dd/yyyy") & "# "
                                                              
End Function
Function GetDateStringForSQLWhere_US(Date_US As String) As String

    GetDateStringForSQLWhere_US = " #" & Date_US & "# "
                                                              
End Function
Function GetDateTimeStringForSQLWhere(DateTime_UK As String) As String

    GetDateTimeStringForSQLWhere = " #" & Format$(DateTime_UK, "mm/dd/yyyy hh:mm:ss") & "# "
                                                              
End Function


Function GetFractionPart(num As Double) As Double
'
' Gets the fractional (decimal) part of a number
'
Dim TheNumber As String, Pos As Long

    TheNumber = CStr(num)
    Pos = InStr(1, TheNumber, ".")
    
    If Pos = 0 Then
        GetFractionPart = 0
    Else
        GetFractionPart = CDbl("0" & Right(TheNumber, Len(TheNumber) - (Pos - 1)))
    End If
    
End Function
Function ConvertTimeToHHMM_12Hr_No_AM_PM(TheTime As Variant) As String

Dim TheNumber As String, Pos As Long, str As String

    If TheTime = "" Then
        ConvertTimeToHHMM_12Hr_No_AM_PM = ""
        Exit Function
    End If

    str = Format(TheTime, "H:MM AMPM")
    
    ConvertTimeToHHMM_12Hr_No_AM_PM = Trim(Left(str, Len(str) - 2))
    
End Function


Function Alphabet(num As Integer) As String
'
' Converts a number to the equivalent letter for looking up spreadsheet
'  columns
'
' Chris Boutal ITNET 20 Jul 1999
'
' Return blank if an error occurs
'
Alphabet = " "
On Error Resume Next
Select Case num
    Case Is < 27: Alphabet = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", num, 1)
    Case Is < 53: Alphabet = Mid$("AAABACADAEAFAGAHAIAJAKALAMANAOAPAQARASATAUAVAWAXAYAZ", 2 * num - 1, 2)
    Case Else: 'returns blank if number out of range
End Select
Exit Function
    
End Function
Public Sub SetFlexGridColumnWidths(TheForm As Form, _
                                   TheGrid As MSFlexGrid, _
                                   Optional IncreaseWidthByFactor As Double = 1.5)
Dim i As Long, j As Long, lMaxTextWidth As Long, lWidth As Long
On Error GoTo ErrorTrap

    With TheGrid
    
    For i = 0 To .Cols - 1
    
        lMaxTextWidth = 0
        
        For j = 0 To .Rows - 1
            
            lWidth = GetTextWidthInTwips(.TextMatrix(j, i), TheForm)
            
            If lWidth > lMaxTextWidth Then
                lMaxTextWidth = lWidth
            End If
    
        Next j
        
        .ColWidth(i) = lMaxTextWidth * IncreaseWidthByFactor
        
    Next i
    
    End With

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Function UnSpace(text As String) As String
'
' Function to remove spaces, brackets and commas and
'  change all letters to uppercase for comparing strings
'  in freeish format
'
' Chris Boutal ITNET 01 Sep 1999
'
Dim i As Integer
Dim Letter As String
Dim Record As String
'
UnSpace = ""
On Error Resume Next
Record = ""
For i = 1 To Len(text)
    Letter = Mid$(text, i, 1)
    Select Case Letter
        Case " ": 'ignore
        Case "(": 'ignore
        Case ")": 'ignore
        Case ",": 'ignore
        Case Else: Record = Record & Letter
    End Select
Next i
UnSpace = UCase$(Record)
Exit Function

End Function
Public Function RemoveNonPrintingChars(TheText As String) As String
'
' Function to remove non-printing characters and convert any
'  funnies to something more useful....
'
Dim i As Integer
Dim Letter As String
Dim PurgedString As String

    RemoveNonPrintingChars = ""
    PurgedString = ""
    
    For i = 1 To Len(TheText)
        Letter = Mid$(TheText, i, 1)
        Select Case Asc(Letter)
        Case 32 To 126
            PurgedString = PurgedString & Letter
        Case 146 'this is a funny type of apostrophe. Convert to normal type...
            PurgedString = PurgedString & "'"
        Case 151 'this is the long hyphen, called "em-dash". Convert to normal hyphen.
            PurgedString = PurgedString & " - "
        Case Else
            
        End Select
    Next i
    
    RemoveNonPrintingChars = PurgedString
    
    Exit Function

End Function
Public Function MakeStringValidForFileName(TheText As String, Optional ReplaceDodgyCharWith As String = "_") As String
'
'
Dim i As Integer
Dim Letter As String
Dim PurgedString As String

    
    For i = 1 To Len(TheText)
        Letter = Mid$(TheText, i, 1)
        Select Case Asc(Letter)
        Case 92, 47, 58, 34, 42, 63, 60, 62, 124 '\ / : " * ? < > |
            PurgedString = PurgedString & ReplaceDodgyCharWith
        Case 32 To 126
            PurgedString = PurgedString & Letter
        Case 146 'this is a funny type of apostrophe. Convert to normal type...
            PurgedString = PurgedString & "'"
        Case 151 'this is the long hyphen, called "em-dash". Convert to normal hyphen.
            PurgedString = PurgedString & " - "
        Case Else
            
        End Select
    Next i
    
    MakeStringValidForFileName = PurgedString
    
    Exit Function

End Function

Public Function RemovePuncFromPersName(TheText As String) As String
'
' Function to remove non-printing characters and convert any
'  funnies to something more useful....
'
Dim i As Integer
Dim Letter As String
Dim PurgedString As String

    RemovePuncFromPersName = ""
    PurgedString = ""
    
    For i = 1 To Len(TheText)
        Letter = Mid$(TheText, i, 1)
        Select Case Letter
        Case Chr(146) 'this is a funny type of apostrophe. Convert to normal type...
        Case Chr(151) 'this is the long hyphen, called "em-dash". Convert to normal hyphen.
        Case "'", "-", "!", "£", "$", "%", "^", "&", "*", "(", ")", "_", "+", "=", "{", _
             "[", "}", "]", ":", ";", "@", "'", "~", "#", "<", ",", ">", ".", "?", "/", _
             "|", "\", "¬", "`", "¦", """"
        Case Else
            PurgedString = PurgedString & Letter
        End Select
    Next i
    
    RemovePuncFromPersName = PurgedString
    
    Exit Function

End Function

Public Function RemoveNonNumericChars(TheText As String) As String
'
' Function to remove non-numeric characters
'
Dim i As Integer
Dim Letter As String
Dim PurgedString As String

    RemoveNonNumericChars = ""
    PurgedString = ""
    
    For i = 1 To Len(TheText)
        Letter = Mid$(TheText, i, 1)
        Select Case Asc(Letter)
        Case 48 To 57
            PurgedString = PurgedString & Letter
        Case Else
            
        End Select
    Next i
    
    RemoveNonNumericChars = PurgedString
    
    Exit Function

End Function
Public Function RemoveNumericChars(TheText As String, _
                                   Optional RemoveNonPrintables As Boolean = False) _
                                                                          As String
'
' Function to remove numeric characters
'
Dim i As Integer
Dim Letter As String
Dim PurgedString As String

    RemoveNumericChars = ""
    PurgedString = ""
    
    For i = 1 To Len(TheText)
        Letter = Mid$(TheText, i, 1)
        Select Case Asc(Letter)
        Case 48 To 57
        Case Else
            PurgedString = PurgedString & Letter
        End Select
    Next i
    
    RemoveNumericChars = IIf(RemoveNonPrintables, _
                             RemoveNonPrintingChars(PurgedString), _
                             PurgedString)
    
    Exit Function

End Function
Public Function RemoveNonTextChars(TheText As String) As String
'
Dim i As Integer
Dim Letter As String
Dim PurgedString As String

    RemoveNonTextChars = ""
    PurgedString = ""
    
    For i = 1 To Len(TheText)
        Letter = Mid$(TheText, i, 1)
        Select Case Asc(Letter)
        Case 48 To 57
            PurgedString = PurgedString & Letter
        Case 65 To 90
            PurgedString = PurgedString & Letter
        Case 97 To 122
            PurgedString = PurgedString & Letter
        Case Else
        End Select
    Next i
    
    RemoveNonTextChars = PurgedString
    
    Exit Function

End Function


Function DeleteTable(TheTable As String) As Boolean
Dim MsgBoxResult As Integer
'NOTE tabledefs methods seem rather unreliable when using via VB6
    
    'M J Thompson ITNET August 1999
    '
    'Deletes a table. (Data AND Structure!!!)
    'Parm is name of table to delete (eg "tblTempMapping")
    'If successful delete, DeleteTable returned to calling function as TRUE,
    ' else FALSE
    '
    
    On Error Resume Next
    'CMSDB.TableDefs.Delete TheTable
    CMSDB.Execute "DROP TABLE [" & TheTable & "]"
    ' 0 = No error; 2008/3211/3008 = Table in use; 2580/3265/3295/3376 = Table doesn't exist

    Select Case Err.number
    Case 0, 2580, 3265, 3295, 3371, 3376
        DeleteTable = True
    Case 2008, 3211, 3008, 3260, 3262
        MsgBox TheTable & " table is in use - close it down", vbExclamation, "ERROR"
        DeleteTable = False
        EndProgram
        Exit Function
    Case Else
        MsgBoxResult = MsgBox("Error " & Err.number & " occured while deleting table " & TheTable _
           , vbExclamation + vbOKOnly, AppName)
        DeleteTable = False
        EndProgram
        Exit Function
    End Select
    

End Function



Public Function DelAllRows(TableName As String) As Boolean
'
'M J Thompson ITNET August 1999
'
'Delete data in table. Parm is name of table whose data is to be
'deleted (eg "tblChannelMapping")
'

On Error GoTo TableInUse

    CMSDB.Execute ("DELETE * FROM " & TableName)
    
    DelAllRows = True
    
    Exit Function
    
TableInUse:
    Select Case Err.number
    Case 0, 2580, 3265
        DelAllRows = True
    Case 2008, 3211, 3008, 3260, 3262
        MsgBox TableName & " table is in use - close it down", vbExclamation, "MsgboxTitleMap"
        DelAllRows = False
        EndProgram
        Exit Function
    Case Else
        MsgBox "Error " & Err.number & " occured while deleting table " & TableName _
           , vbExclamation, AppName
        DelAllRows = False
        EndProgram
        Exit Function
    End Select

End Function



Public Function DelSubstr(StrToSearch As String, _
                          StrToRemove As String, _
                          CaseSens As Boolean, _
                          Optional StartPos As Long = 1 _
                          ) As String

'M J Thompson - August 1999 - ITNET
'
'This function deletes a specified substring (StrToRemove) from a string (StrToSearch)
' and closes the gap. If CaseSens is TRUE, StrToRemove must have identical case to
' the substring to be deleted.

Dim StrLeft As String, StrRight As String, PosStrToRemove As Integer, CaseSensDig As Byte


On Error GoTo ErrorTrap

    CaseSensDig = IIf(CaseSens, 0, 1)  'Check if case sensitivity required
    
    PosStrToRemove = InStr(StartPos, StrToSearch, StrToRemove, CaseSensDig) 'find position of substring to remove
    If PosStrToRemove <> 0 Then
        StrLeft = Left$(StrToSearch, PosStrToRemove - 1)    'if substring present, separate left part of string before start of substring
        StrRight = Right$(StrToSearch, (Len(StrToSearch) - PosStrToRemove + 1))
        Mid$(StrRight, 1, Len(StrToRemove)) = Space(Len(StrToRemove)) 'replace substring with spaces
        StrRight = Trim(StrRight)   'Remove leading/trailing spaces
        DelSubstr = StrLeft & StrRight 'New string
    Else
        DelSubstr = StrToSearch
    End If
    
    Exit Function
    

ErrorTrap:
    EndProgram
        
End Function

Public Function InsertSubstr(OriginalString As String, _
                             StrToInsert As String, _
                             InsertPos As Long _
                                 ) As String


Dim StrLeft As String, StrRight As String, PosStrToRemove As Integer

On Error GoTo ErrorTrap

'insert a string into another string. InsertPos is zero-based.

    If InsertPos <= 0 Then
        InsertSubstr = StrToInsert & OriginalString
    Else
        If InsertPos >= Len(OriginalString) Then
            InsertSubstr = OriginalString & StrToInsert
        Else
            InsertSubstr = Left(OriginalString, InsertPos) & _
                            StrToInsert & _
                            Right(OriginalString, Len(OriginalString) - InsertPos)
        End If
    End If
    
    Exit Function
    
ErrorTrap:
    EndProgram
        
End Function



Public Function FieldInTable(StringToFind As String, KeyField As String, KeyNo As Long, _
                             TableToSearch As String, FieldToSearch As String) As Boolean
'
'M J Thompson - August 1999 - ITNET
'
'Used to check whether a string ("StringToFind") exists in a table in field "FieldToSearch".
'Can be used for validation.
'If field does exist, FieldInTable is returned as TRUE and the corresponding
'KeyField (eg 'SeqNum') is returned as "KeyNo".
'
'Pass in the String to validate as "StringToFind"
'Pass in the tablename used to do validation as "TableToSearch"
'Pass in the fieldname on the validation table as "FieldToSearch"


On Error GoTo ErrorTrap

Dim SearchSQL As String, rstRecSet As Recordset

    SearchSQL = "SELECT " & KeyField & " FROM " & TableToSearch & " WHERE " & FieldToSearch & " = """ & StringToFind & """"
    Set rstRecSet = CMSDB.OpenRecordset(SearchSQL, dbOpenForwardOnly)
     
    With rstRecSet
    If Not .BOF Then 'Field found in table
        FieldInTable = True
        KeyNo = .Fields(KeyField)
    Else
        FieldInTable = False
        KeyNo = 0
    End If
    Exit Function
    End With
    

    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Function CreateTable(ErrorCode As Integer, _
                            TableName As String, _
                            FieldName As String, _
                            DataType As String, _
                            Optional FieldSize As String = "50", _
                            Optional Not_Null As String = "NOT NULL", _
                            Optional CreateKey As Boolean = True, _
                            Optional AutoKeyName As String = "SeqNum", _
                            Optional JetRequiredField As Boolean = False) As Boolean
'
'M J Thompson - ITNET - August 1999
'
'Creates a table with two fields, one (SeqNum) generated internally to this function as the
'key. Tablename and fieldname passed in as parms. Errorcode is returned to calling proc.
'If ErrorCode = 3010, Table already exists.
'If CREATE fails, CreateTable returned as FALSE.
'To see full list of DataTypes, see Help under "Creating Fields in Tables" - "Data Types" -
' "Comparison of Data Types".
'
On Error GoTo ErrorTrap

    If CreateKey Then
        If DataType = "TEXT" Then
            CMSDB.Execute "CREATE TABLE " & TableName & _
                " (" & AutoKeyName & " COUNTER CONSTRAINT Constr " _
                & "PRIMARY KEY, " & FieldName & " " & DataType & " (" & FieldSize & ") " & Not_Null & ");"
        Else
            CMSDB.Execute "CREATE TABLE " & TableName & _
                " (" & AutoKeyName & " COUNTER CONSTRAINT Constr " _
                & "PRIMARY KEY, " & FieldName & " " & DataType & " " & Not_Null & ");"
           
        End If
    Else
        If DataType = "TEXT" Then
            CMSDB.Execute "CREATE TABLE " & TableName & _
                " (" & FieldName & " " & DataType & " (" & FieldSize & ") " & Not_Null & ");"
        Else
            CMSDB.Execute "CREATE TABLE " & TableName & _
                " (" & FieldName & " " & DataType & " " & Not_Null & ");"
           
        End If
    End If
    
    CMSDB.TableDefs.Refresh
    CMSDB.TableDefs(TableName).Fields(FieldName).Required = JetRequiredField
    CMSDB.TableDefs.Refresh
    
    CreateTable = True
    
    Exit Function
    

ErrorTrap:
    EndProgram
        
End Function

Public Function CreateField(ErrorCode As Integer, TableName As String, _
                FieldName As String, DataType As String, Optional FieldSize As String = "50", _
                Optional Not_Null As String = "NOT NULL", Optional JetRequiredField As Boolean = False) As Boolean
'
'M J Thompson - ITNET - August 1999
'
'Adds a field to a table.
'Tablename and fieldname passed in as parms. Errorcode is returned to calling proc. If
'ErrorCode = 3380, Field already exists. If ALTER fails, CreateTable returned as FALSE.
'To see full list of DataTypes, see Help under "Creating Fields in Tables" - "Data Types" -
' "Comparison of Data Types".
'
'Set Not_Null to "" if not required
'
On Error GoTo ErrorTrap
    
    If DataType = "TEXT" Then
        CMSDB.Execute "ALTER TABLE " & TableName & " ADD COLUMN " & _
        FieldName & " " & DataType & " (" & FieldSize & ") " & Not_Null & ";"
    Else
        CMSDB.Execute "ALTER TABLE " & TableName & " ADD COLUMN " & _
        FieldName & " " & DataType & " " & Not_Null & ";"
    End If
    
    CMSDB.TableDefs.Refresh
    CMSDB.TableDefs(TableName).Fields(FieldName).Required = JetRequiredField
    CMSDB.TableDefs.Refresh
        
    CreateField = True

    Exit Function
    
    
ErrorTrap:
    EndProgram
        
    
End Function
Public Function CreateIndex(ErrorCode As Integer, TableName As String, FieldName As String, _
                            IndexName As String, UniqueIX As Boolean, AllowNull As Boolean, Optional CreatePrimary As Boolean = False) _
                            As Boolean
'
'M J Thompson - ITNET - August 1999
'
'Adds index to a field in a table.
'Tablename, fieldname and Index Name passed in as parms. Errorcode is returned to calling proc. If
'ErrorCode = 3409, field doesn't exist. If ErrorCode = 3375, Index Name already exists.
'IndexName can be anything, but must be unique for each Indexed field.
'If operation fails, CreateIndex returned as FALSE.
'
                            
                            
                            
Dim strUnique As String, strNull As String

    
    On Error Resume Next
    
    CMSDB.Execute "DROP INDEX " & IndexName & " ON " & TableName & ";"


    
    On Error GoTo ErrorTrap

    strUnique = IIf(UniqueIX, "UNIQUE", "")
                            
'    strNull = IIf(AllowNull, "", "WITH DISALLOW NULL")
'
    If CreatePrimary Then
        strNull = "WITH PRIMARY"
    Else
        If Not AllowNull Then
            strNull = "WITH DISALLOW NULL"
        End If
    End If
    
    
    CMSDB.Execute "CREATE " & strUnique & " INDEX " & IndexName _
                    & " ON " & TableName & " (" & FieldName & " ) " _
                    & strNull & ";"
                    
    CreateIndex = True
    
    Exit Function
    
ErrorTrap:
    EndProgram
        
                
End Function

Public Function DropIndex(ErrorCode As Integer, TableName As String, _
                            IndexName As String) _
                            As Boolean
'
'M J Thompson - ITNET - August 1999
'
'
                            
On Error GoTo ErrorTrap
                            
Dim strUnique As String, strNull As String

                        
    CMSDB.Execute "DROP INDEX " & IndexName _
                    & " ON " & TableName & ";"
                    
    DropIndex = True
    
    Exit Function
    
ErrorTrap:
    EndProgram
        
                
End Function
Public Function ValidDate(DateStr As String) As Boolean
Dim TheMonth As Integer, TheYear As Integer, TheDay As Integer
Dim IsLeapYear As Boolean, TestDate As String


On Error GoTo ErrorTrap

    DateStr = Trim(DateStr)
    
    If IsDate(DateStr) Then
        TestDate = Format$(DateStr, "DD/MM/YYYY")
        
        TheMonth = Mid$(TestDate, 4, 2)
        TheYear = Right$(TestDate, 4)
        TheDay = Left$(TestDate, 2)
        
        If Left(TestDate, 5) <> Left(DateStr, 5) Then
            ValidDate = False
            Exit Function
        End If
        
        If (TheYear Mod 100) = 0 Then
            If (TheYear Mod 4) = 0 Then
                IsLeapYear = True
            Else
                IsLeapYear = False
            End If
        Else
            If (TheYear Mod 4) = 0 Then
                IsLeapYear = True
            Else
                IsLeapYear = False
            End If
        End If
        
        If ((Not IsLeapYear) And TheMonth = 2 And TheDay = 29) Or _
            ((TheDay = 31) And (TheMonth = 4 Or TheMonth = 6 Or TheMonth = 9 Or TheMonth = 11)) _
            Then
            
            ValidDate = False
        Else
            ValidDate = True
        End If
    Else
        ValidDate = False
    End If
    

    Exit Function
ErrorTrap:
    EndProgram
End Function

Public Function DropField(ErrorCode As Integer, TableName As String, FieldName As String) As Boolean
'
'M J Thompson - ITNET - August 1999
'
'Remove a field from a table.
'Tablename and fieldname passed in as parms. Errorcode is returned to calling proc.
'
On Error Resume Next
    
    CMSDB.Execute "ALTER TABLE " & TableName & " DROP COLUMN " & FieldName & ";"

    DropField = (Err.number = 0)

    
End Function

Public Function DeleteSomeRows(TableName As String, Criteria As String, Optional TheValue)
'
'Delete data in table.
'

On Error GoTo TableInUse

    If IsMissing(TheValue) Then
        TheValue = ""
    End If

    CMSDB.Execute ("DELETE * FROM " & TableName & " WHERE " & Criteria & TheValue)
    
    DeleteSomeRows = True
    
    Exit Function
    
TableInUse:
    Select Case Err.number
    Case 0, 2580, 3265
        DeleteSomeRows = True
    Case 2008, 3211, 3008, 3260, 3262
        DeleteSomeRows = False
        EndProgram "Table locked: " & TableName
    Case Else
        DeleteSomeRows = False
        EndProgram "Error deleting from " & TableName
    End Select

End Function

Public Function FindMaxVal(arr As Variant)
    Dim Index As Long, index2 As Long
    Dim firstItem As Long, LastItem As Long, i As Long, MaxVal As Variant

On Error GoTo ErrorTrap

    ' exit if it is not an array
    If VarType(arr) < vbArray Then Exit Function
    
    firstItem = LBound(arr)
    LastItem = UBound(arr)
    For i = firstItem To LastItem
        If arr(i) > MaxVal Then
            MaxVal = arr(i)
        End If
    Next i
    
    FindMaxVal = MaxVal

    Exit Function
ErrorTrap:
    EndProgram
        
End Function

Public Function TableExists(TableName As String) As Boolean
Dim temp As Recordset
'NOTE tabledefs methods seem rather unreliable when using via VB6
On Error GoTo NoTable
    
    Set temp = CMSDB.OpenRecordset(TableName, dbOpenSnapshot)
    
    TableExists = True
    
    On Error Resume Next
    temp.Close
    Set temp = Nothing

    Exit Function
    
NoTable:
    TableExists = False
    Set temp = Nothing
    
    Err.Clear
    
End Function
Public Function TableExists_ExtDB(TheDB As Database, TableName As String) As Boolean
Dim temp As Recordset
On Error GoTo NoTable
    
    Set temp = TheDB.OpenRecordset(TableName, dbOpenSnapshot)
    
    TableExists_ExtDB = True
    
    On Error Resume Next
    temp.Close
    Set temp = Nothing

    Exit Function
    
NoTable:
    TableExists_ExtDB = False
    
    Set temp = Nothing
    
    Err.Clear
    
End Function

Public Sub CopyTable(NewTableName As String, TableToCopy As String, DBName As DAO.Database)

On Error GoTo ErrorTrap

    DeleteTable NewTableName
    
    DBName.Execute "select * Into [" & NewTableName & "] From [" & TableToCopy & "]"


    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub



Public Sub EndProgram(Optional CustomMessage As String = "", _
                      Optional RestorePrevBackup As Boolean = True)
                      
Dim ErrNum As Long, ErrDesc As String, ErrSrc As String
    
    Screen.MousePointer = vbNormal
    
    ErrNum = Err.number
    ErrDesc = Err.Description
    ErrSrc = Err.source
    
    WriteToLogFile "Error " & ErrNum & " (" & ErrDesc & ") has occurred in " & _
    ErrSrc & vbCr & vbCr & "Further info: " & CustomMessage & vbCr & vbCr & _
    "The application will now terminate."
    
    
    MsgBox "Error " & ErrNum & " (" & ErrDesc & ") has occurred in " & _
    ErrSrc & vbCr & vbCr & "Further info: " & CustomMessage & vbCr & vbCr & _
    "The application will now terminate.", vbCritical + vbOKOnly, AppName
        
    End
    
End Sub

Public Function IsTime(TheTimeToCheck As String) As Boolean
Dim i As Long
    
    TheTimeToCheck = Trim$(TheTimeToCheck)
    
    If Len(TheTimeToCheck) > 5 Then
        IsTime = False
        Exit Function
    End If
    If Len(TheTimeToCheck) < 4 Then
        IsTime = False
        Exit Function
    End If
    
    i = InStr(TheTimeToCheck, ":")
    
    If i < 2 Or i > 3 Then
        IsTime = False
        Exit Function
    End If
    
    If i = 2 Then
        If Len(TheTimeToCheck) > 4 Then
            IsTime = False
            Exit Function
        Else
            TheTimeToCheck = "0" & TheTimeToCheck
        End If
    End If
    
    If Not IsNumeric(Left$(TheTimeToCheck, 2)) Then
        IsTime = False
        Exit Function
    End If
    
    If Not IsNumeric(Right$(TheTimeToCheck, 2)) Then
        IsTime = False
        Exit Function
    End If
    
    If Left$(TheTimeToCheck, 2) > 23 Then
        IsTime = False
        Exit Function
    End If
    
    If Right$(TheTimeToCheck, 2) > 59 Then
        IsTime = False
        Exit Function
    End If
    
    IsTime = True
    
End Function

Public Function WeekOfMonth(TheDate As Date) As Integer
Dim WeekOfYear_1 As Integer, WeekOfYear_2 As Integer, TheMonth As Integer, TheYear As Integer
    
'
'Week of month for TheDate is week of year minus week of year in which the 1st of month falls
'
    
    TheMonth = CInt(DatePart("m", TheDate, vbMonday, vbFirstJan1))
    TheYear = CInt(DatePart("yyyy", TheDate, vbMonday, vbFirstJan1))
    
    WeekOfYear_1 = DatePart("ww", DateSerial(TheYear, TheMonth, 1), vbMonday, vbFirstJan1)
    WeekOfYear_2 = DatePart("ww", TheDate, vbMonday, vbFirstJan1)
    
    WeekOfMonth = WeekOfYear_2 - WeekOfYear_1 + 1

End Function

Public Function OrdinalPosOfDay(TheDate As Date) As Variant
 '
'Let's say TheDate is the 5th Tuesday of the month. Function returns "L" - last Tuesday
' of month.
'
Dim TempDate As Date, i As Byte, DateOfFirstDay As Date, PrevMonth As Byte

    'first go to first day of the month
    TempDate = DateSerial(year(TheDate), Month(TheDate), 1)
    
    i = 1
    
    'Now find date of first (eg) Tuesday in month
    Do Until Weekday(TheDate) = Weekday(TempDate)
        TempDate = DateAdd("d", 1, TempDate)
    Loop
        
    'TempDate now contains date of first day
    DateOfFirstDay = TempDate
    
    Do Until Month(TheDate) <> Month(TempDate) Or TempDate = TheDate
        TempDate = DateAdd("ww", 1, TempDate)
        i = i + 1
    Loop
    
    'i now contains ordinal position of day
    
    Select Case i
    Case 5:
        'last (eg) tUESDAY OF MONTH
        OrdinalPosOfDay = "L"
    Case 4:
        '4th (eg) tUESDAY OF MONTH.... but is it also the last....?
        TempDate = DateAdd("ww", 1, TempDate)
        If Month(TheDate) = Month(TempDate) Then
            OrdinalPosOfDay = 4
        Else
            OrdinalPosOfDay = "L"
        End If
    Case Else
        OrdinalPosOfDay = i
    End Select

End Function

Public Function NumberOfWeeksInMonth(TheDay As Long, TheYear As Long, TheMonth As Long) As Long
Dim i As Long
'
'Let's say we want to find number of Wednesdays in a month. Supply following parms:
' TheDay = vbWednesday, TheYear =2003, TheMonth =2
'
Dim TempDate As Date, DayPos As Variant

    'first go to first (eg) Wednesday of the month
    TempDate = DateOfNthDay(TheDay, TheYear, TheMonth, 1)
    
    i = 0
    
    'count number of (eg) Wednesdays in month
    Do Until DayPos = "L"
        DayPos = OrdinalPosOfDay(TempDate)
        i = i + 1
        TempDate = DateAdd("ww", 1, TempDate)
    Loop
        
    'TempDate now contains date of Nth day
    NumberOfWeeksInMonth = i
    
End Function


Public Function DateOfNthDay(TheDay As Long, TheYear As Long, TheMonth As Long, WhichDayPosition As Long) As Date
Dim i As Long
'
'Let's say we want to find the 3rd Wednesday of February 2003. Supply following parms:
' TheDay = vbWednesday, TheYear =2003, TheMonth =2, WhichDayPosition =3
'
Dim TempDate As Date

    'first go to first day of the month
    TempDate = DateSerial(TheYear, TheMonth, 1)
    
    i = 0
    
    'Now find date of Nth (eg) Wednesday in month
    Do
        If Weekday(TempDate) = TheDay Then
            i = i + 1
            If i = WhichDayPosition Then
                Exit Do
            End If
        End If
        TempDate = DateAdd("d", 1, TempDate)
    Loop
        
    'TempDate now contains date of Nth day
    DateOfNthDay = TempDate
    
End Function
Public Function DateOfLastDayOfMonth(TheDay As Long, TheYear As Long, TheMonth As Long) As Date
Dim i As Long
'
'Let's say we want to find the last Wednesday of February 2003. Supply following parms:
' TheDay = vbWednesday, TheYear =2003, TheMonth =2
'
Dim TempDate As Date

    'first go to 4th occurence of TheDay of the month
    TempDate = DateOfNthDay(TheDay, TheYear, TheMonth, 4)
    
    If OrdinalPosOfDay(TempDate) <> "L" Then
        TempDate = DateAdd("ww", 1, TempDate)
    End If
    
    DateOfLastDayOfMonth = TempDate
    
End Function

Public Function GetDateOfGivenDay(TheDate As Date, _
                                  TheWeekDay, Optional bAfterTheDate) As Date
'
'Let's say we want to find the Monday of the week containing date 15/10/03. Supply parm
' TheDate=15/10/03, TheWeekDay = vbMonday. Function will return 13/10/03.
'
'bAfterTheDate:
' if TRUE, forces function to return date AFTER TheDate
' if FALSE, forces function to return date BEFORE TheDate
' if not set, whether TheDate is before/after the returned date is undefined.

Dim TempDate As Date, WeekDayOfParmDate As Long, DayDiff As Long

    '
    'First find weekday of given date:
    '
    WeekDayOfParmDate = Weekday(TheDate, vbMonday)   'Monday is day 1, Sunday is day 7
    
    DayDiff = WeekDayOfParmDate - TheWeekDay + 1
    
    TempDate = TheDate - DayDiff
    
    If IsMissing(bAfterTheDate) Then
        GetDateOfGivenDay = TempDate
        Exit Function
    End If
    
    If bAfterTheDate = True Then
        If TempDate < TheDate Then
            GetDateOfGivenDay = TempDate + 7
        Else
            GetDateOfGivenDay = TempDate
        End If
    Else
        If TempDate > TheDate Then
            GetDateOfGivenDay = TempDate - 7
        Else
            GetDateOfGivenDay = TempDate
        End If
    End If
    
End Function

'Public Function ConvertSunDayOneToMonDayOne(TheWeekDay As Long) As Long
''
'    If TheWeekDay = 1 Then
'        ConvertSunDayOneToMonDayOne = 7
'    Else
'        ConvertSunDayOneToMonDayOne = TheWeekDay - 1
'    End If
'
'End Function


Public Function GetDateOfFirstWeekDayOfMonth(TheDate As Date, TheWeekDay) As Date
Dim HoldDate As Date
'
'Let's say we want to find the FIRST Monday of the week containing date 01/06/05.
' Supply parm TheDate=01/06/05, TheWeekDay = vbMonday.
' Function will return 06/06/05
'
Dim TempDate As Date, WeekDayOfParmDate As Long, DayDiff As Long

    HoldDate = GetDateOfGivenDay(TheDate, TheWeekDay)
    
    If Month(HoldDate) <> Month(TheDate) Then
        GetDateOfFirstWeekDayOfMonth = DateAdd("ww", 1, HoldDate)
    Else
        GetDateOfFirstWeekDayOfMonth = HoldDate
    End If
    
End Function
Public Function RemoveLeadingDotsFromString(TheString) As String
Dim str As String, i As Long

    str = TheString
    
    For i = 1 To Len(TheString)
        If Left(str, 1) = "." Then
            If Len(str) > 1 Then
                str = Right(str, Len(str) - 1)
            Else
                str = ""
                Exit For
            End If
        Else
            Exit For
        End If
    Next i
    
    RemoveLeadingDotsFromString = str
    
End Function


Function DoubleUpSingleQuotes(strInput) As String
    DoubleUpSingleQuotes = Replace(strInput, "'", "''")
End Function

Public Function GetMonthNumber(TheString) As Long
On Error GoTo ErrorTrap

    Select Case TheString
    Case "Jan"
        GetMonthNumber = 1
    Case "Feb"
        GetMonthNumber = 2
    Case "Mar"
        GetMonthNumber = 3
    Case "Apr"
        GetMonthNumber = 4
    Case "May"
        GetMonthNumber = 5
    Case "Jun"
        GetMonthNumber = 6
    Case "Jul"
        GetMonthNumber = 7
    Case "Aug"
        GetMonthNumber = 8
    Case "Sep"
        GetMonthNumber = 9
    Case "Oct"
        GetMonthNumber = 10
    Case "Nov"
        GetMonthNumber = 11
    Case "Dec"
        GetMonthNumber = 12
    End Select

    Exit Function
    
ErrorTrap:
    EndProgram
End Function

Public Function GetMonthName(ByVal TheMonth As Long, _
                             Optional ByVal bAbbrev As Boolean = False) As String
On Error GoTo ErrorTrap
Dim str As String

    If TheMonth > 12 Or TheMonth < 1 Then
        EndProgram "GetMonthName(" & TheMonth & ") - Invalid month argument"
    End If
    
    str = Format("01/" & CStr(TheMonth) & "/2004", "mmmm")
    
    If Not bAbbrev Then
        GetMonthName = str
    Else
        GetMonthName = Left(str, 3)
    End If

    Exit Function
    
ErrorTrap:
    EndProgram
End Function
Public Function RemoveExtraSpacesToLeaveSingleSpacedWords(InputString As String) As String
On Error GoTo ErrorTrap
Dim TheString As String

    TheString = Replace(InputString, "  ", " ")
    
    Do While InStr(1, TheString, "  ") > 0
        TheString = Replace(TheString, "  ", " ")
    Loop

    RemoveExtraSpacesToLeaveSingleSpacedWords = TheString

    Exit Function
    
ErrorTrap:
    EndProgram
    
End Function



Public Sub DestroyGlobalObjects(Optional bExcludeFSO As Boolean = False)
Dim TheString As String

    Set GlobalParms = Nothing
    Set GlobalCalendar = Nothing
    Set TheTMS = Nothing
    Set CongregationMember = Nothing
    
    If Not bExcludeFSO Then
        Set gFSO = Nothing
    End If

End Sub
Public Sub InstantiateGlobalObjects(Optional bExcludeFSO As Boolean = False)
On Error GoTo ErrorTrap
Dim TheString As String

    Set GlobalParms = New clsApplicationConstants
    Set GlobalCalendar = New clsCalendar
    Set TheTMS = New clsTMS
    Set CongregationMember = New clsCongregationMember
    
    If Not bExcludeFSO Then
        Set gFSO = New FileSystemObject
    End If
    
    If FormIsOpen("frmPersonalDetails") Then
        frmPersonalDetails.SetUpNameAddressRecSets
    End If
    
    If FormIsOpen("frmTMSScheduling") Then
        RefreshGrid True
    End If
    
    

    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub



Sub BubbleSort(arr As Variant, Optional numEls As Variant, _
    Optional descending As Boolean)
' Bubble Sort an array of any type
' BubbleSort is especially convenient with small arrays (1,000
' items or fewer) or with arrays that are already almost sorted
'
' NUMELS is the index of the last item to be sorted, and is
' useful if the array is only partially filled.
'
' Works with any kind of array, except UDTs and fixed-length
' strings, and including objects if your are sorting on their
' default property. String are sorted in case-sensitive mode.
'
' You can write faster procedures if you modify the first two lines
' to account for a specific data type, eg.
' Sub BubbleSortS(arr() As Single, Optional numEls As Variant,
'  '     Optional descending As Boolean)
'   Dim value As Single

    Dim value As Variant
    Dim Index As Long
    Dim firstItem As Long
    Dim indexLimit As Long, lastSwap As Long

    ' account for optional arguments
    If IsMissing(numEls) Then numEls = UBound(arr)
    firstItem = LBound(arr)
    lastSwap = numEls

    Do
        indexLimit = lastSwap - 1
        lastSwap = 0
        For Index = firstItem To indexLimit
            value = arr(Index)
            If (value > arr(Index + 1)) Xor descending Then
                ' if the items are not in order, swap them
                arr(Index) = arr(Index + 1)
                arr(Index + 1) = value
                lastSwap = Index
            End If
        Next
    Loop While lastSwap
End Sub

Function Ceiling(number As Double) As Long
'
'Returns next integer for a given number. eg if number=1.00001, Ceiling returned
' as 2
'
    Ceiling = -Int(-number)
End Function



Public Sub CheckForCOOrAssembly(IsMovedOralReview As Boolean, _
                                 NextWeekIsMovedOralReview As Boolean, _
                                 LastWeekWasAssemblyWeek As Boolean, _
                                 IsCOVisitThisWeek As Boolean, _
                                 LastWeekWasCOVisit As Boolean, _
                                 AssignmentDate As Date)
                                 
Dim TheCong As Long, rsttmsquery As Recordset
                                 
On Error GoTo ErrorTrap
                                 
    TheCong = GlobalParms.GetValue("DefaultCong", "NumVal")
            
    '
    'First check for moved Oral Review. An Oral Review will ONLY appear in
    ' tblTMSSchedule if it has been moved from it's normal week due to eg
    ' a CO visit....
    '
    Set rsttmsquery = CMSDB.OpenRecordset("SELECT AssignmentDate " & _
                                       "FROM tblTMSSchedule " & _
                                       "WHERE AssignmentDate = #" & Format(AssignmentDate, "mm/dd/yyyy") & _
                                        "# AND TalkNo = 'MR'" _
                                       , dbOpenSnapshot)
                                       
    If rsttmsquery.BOF Then
        IsMovedOralReview = False
    Else
        IsMovedOralReview = True
    End If
    
    '
    'Now check if NEXT week has Moved oral Review...
    '
    Set rsttmsquery = CMSDB.OpenRecordset("SELECT AssignmentDate " & _
                                       "FROM tblTMSSchedule " & _
                                        "WHERE AssignmentDate = #" & Format(AssignmentDate + 7, "mm/dd/yyyy") & _
                                        "# AND TalkNo = 'MR'" _
                                       , dbOpenSnapshot)
                                       
    If rsttmsquery.BOF Then
        NextWeekIsMovedOralReview = False
    Else
        NextWeekIsMovedOralReview = True
    End If
    
                                        
                                        
    '
    'Now check Calendar for assembly last week...
    '
    LastWeekWasAssemblyWeek = IsCircuitOrDistrictAssemblyWeek(AssignmentDate - 7)
    
    '
    'Now check Calendar for CO Visit this week...
    '
    IsCOVisitThisWeek = IsCOVisitWeek(AssignmentDate)
    
    '
    'Now check Calendar for CO Visit this week...
    '
    LastWeekWasCOVisit = IsCOVisitWeek(AssignmentDate - 7)
    

    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Public Sub CheckForOralReview(IsOralReview As Boolean, _
                              AssignmentDate As Date)
                                 
Dim TheCong As Long, rsttmsquery As Recordset
                                 
On Error GoTo ErrorTrap
                                 
    TheCong = GlobalParms.GetValue("DefaultCong", "NumVal")
            
    '
    'Now check for normal Oral Review.
    '
    Set rsttmsquery = CMSDB.OpenRecordset("SELECT AssignmentDate " & _
                                       "FROM tblTMSItems " & _
                                       "WHERE AssignmentDate = #" & Format(AssignmentDate, "mm/dd/yyyy") & _
                                        "# AND TalkNo = 'R'" _
                                       , dbOpenSnapshot)
                                       
    If rsttmsquery.BOF Then
        IsOralReview = False
    Else
        IsOralReview = True
    End If
    
    rsttmsquery.Close
    
    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Public Function IsOralReviewWeek(AssignmentDate As Date) As Boolean
                                 
Dim rs1 As Recordset, bIsOralReview As Boolean, bIsMovedReview As Boolean
Dim bMovedReviewSoon As Boolean
                                 
On Error GoTo ErrorTrap
                                 
    Set rs1 = CMSDB.OpenRecordset("SELECT 1 " & _
                                    "FROM tblTMSItems " & _
                                    "WHERE AssignmentDate = #" & Format(AssignmentDate, "mm/dd/yyyy") & _
                                     "# AND TalkNo = 'R'" _
                                    , dbOpenSnapshot)
                                       
    bIsOralReview = Not rs1.BOF
      
    Set rs1 = CMSDB.OpenRecordset("SELECT 1 " & _
                                    "FROM tblTMSSchedule " & _
                                    "WHERE AssignmentDate = #" & Format(AssignmentDate, "mm/dd/yyyy") & _
                                     "# AND TalkNo = 'MR'" _
                                    , dbOpenSnapshot)
                                       
    bIsMovedReview = Not rs1.BOF
     
    Set rs1 = CMSDB.OpenRecordset("SELECT 1 " & _
                                    "FROM tblTMSSchedule " & _
                                    "WHERE AssignmentDate BETWEEN " & GetDateStringForSQLWhere(CStr(AssignmentDate)) & _
                                                        " AND " & GetDateStringForSQLWhere(CStr(AssignmentDate + 21)) & _
                                    " AND TalkNo = 'MR'" _
                                    , dbOpenSnapshot)
                                       
    bMovedReviewSoon = Not rs1.BOF
    
    If bIsMovedReview Then
        IsOralReviewWeek = True
        Exit Function
    End If
    
    If bIsOralReview And Not bMovedReviewSoon Then
        IsOralReviewWeek = True
        Exit Function
    End If
    
    IsOralReviewWeek = False
     
    On Error Resume Next
    rs1.Close
    Set rs1 = Nothing
    
    Exit Function
ErrorTrap:
    EndProgram


End Function

Public Function IsMovedOralReviewWeek(AssignmentDate As Date) As Boolean
                                 
Dim rs1 As Recordset
                                 
On Error GoTo ErrorTrap
                                 
     
    Set rs1 = CMSDB.OpenRecordset("SELECT 1 " & _
                                    "FROM tblTMSSchedule " & _
                                    "WHERE AssignmentDate = #" & Format(AssignmentDate, "mm/dd/yyyy") & _
                                     "# AND TalkNo = 'MR'" _
                                    , dbOpenSnapshot)
                                       
    IsMovedOralReviewWeek = Not rs1.BOF
         
     
    On Error Resume Next
    rs1.Close
    Set rs1 = Nothing
    
    Exit Function
ErrorTrap:
    EndProgram


End Function



Public Function AccessAllowed(FormNameProperty, ControlNameProperty) As Boolean
'
'SECURITY PHILOSOPHY....
'~~~~~~~~~~~~~~~~~~~~~~~
'
'tblSecurity holds all users and their passwords. They also have a numeric 'UserCode'
'
'UserCode references entries on tblAccessLevels. The Access-Level in a numeric code
' which identifies controls that the user has access to
'
'Any 'secured' controls (eg command buttons etc) are stored by name on tblObjectSecurity,
' qualified by container form name. Each control will be listed along with ALL Access-Levels
' that are allowed to use that control
'
'So, in this function, each of the supplied control's Access-levels is checked against
' the Access-Levels assigned to the current user. If there's a match, the user is allowed
' access (ie, this function returns TRUE, which the calling procedure can use to
' then determine whether to enable/disable/show/hide/lock/unlock the control in the
' form-load sub.
'
'

                                 
Dim rstCheckAccess As Recordset, rstSecurityLevels As Recordset
                                 
On Error GoTo ErrorTrap

    AccessAllowed = False
                                 
    If gbResetPassword Then Exit Function
    
    Set rstCheckAccess = CMSDB.OpenRecordset("SELECT SecurityLevel " & _
                                       "FROM tblObjectSecurity " & _
                                       "WHERE FormNameProperty = '" & FormNameProperty & _
                                       "' AND ControlNameProperty = '" & ControlNameProperty & _
                                        "'", dbOpenSnapshot)
                                                                           
    Set rstSecurityLevels = CMSDB.OpenRecordset("SELECT AccessLevel " & _
                                       "FROM tblAccessLevels " & _
                                       "WHERE UserCode = " & gCurrentUserCode, _
                                       dbOpenSnapshot)
                                           
    With rstCheckAccess
    
    If Not rstSecurityLevels.BOF Then
        If Not .BOF Then
            Do While Not .EOF 'loop through all allowable access-levels for this control
                'Check if user has matching access-level
                rstSecurityLevels.MoveFirst
                rstSecurityLevels.FindFirst "AccessLevel = " & !SecurityLevel
                If Not rstSecurityLevels.NoMatch Then
                    AccessAllowed = True
                    Exit Do
                End If
                .MoveNext
            Loop
        End If
    End If
    
    End With
    
    rstCheckAccess.Close
    rstSecurityLevels.Close
    
    Exit Function
ErrorTrap:
    EndProgram


End Function
Public Function UserHasSecurityLevel(UserCode As Long, AccessLevels_CommaDelimString As String) As Boolean
                                 
Dim rstSecurityLevels As Recordset
                                 
On Error GoTo ErrorTrap

                                                                           
    Set rstSecurityLevels = CMSDB.OpenRecordset("SELECT 1 " & _
                                       "FROM tblAccessLevels " & _
                                       "WHERE UserCode = " & UserCode & _
                                       " AND AccessLevel IN (" & AccessLevels_CommaDelimString & ") ", _
                                       dbOpenSnapshot)
                                           
    UserHasSecurityLevel = Not rstSecurityLevels.BOF
    
    rstSecurityLevels.Close
    Set rstSecurityLevels = Nothing
    
    Exit Function
ErrorTrap:
    EndProgram


End Function


Public Function SaveCSVFile(TheFileName As String, TheDialogTitle As String, FileSaveDialog As CommonDialog) As Boolean
Dim SaveCurDir As String

    On Error GoTo ExitNow

    SaveCSVFile = False
    
    '
    'Note - No file is created until data is written to the path contained in
    '       TheFileName after this routine has run.
    '
    
    '
    'Set up dialogue parms
    '
    FileSaveDialog.Filter = "All files (*.*)|*.*|CSV files|*.csv"
'    FileSaveDialog.Filter = "All files (*.*)|*.*|Text files|*.txt"
    FileSaveDialog.FilterIndex = 2
    FileSaveDialog.DefaultExt = "csv"
    FileSaveDialog.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or _
                           cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
    FileSaveDialog.DialogTitle = TheDialogTitle
    FileSaveDialog.Filename = TheFileName
    
    FileSaveDialog.CancelError = True ' Allow Exit if user presses Cancel.
    
    FileSaveDialog.ShowSave
    TheFileName = FileSaveDialog.Filename
    SaveCSVFile = True

ExitNow:

End Function
Public Function SaveTextFile(TheFileName As String, _
                             TheDialogTitle As String, _
                             FileSaveDialog As CommonDialog, _
                             FileExt_NoDot As String) As Boolean
Dim SaveCurDir As String

    On Error GoTo ExitNow

    SaveTextFile = False
    
    '
    'Note - No file is created until data is written to the path contained in
    '       TheFileName after this routine has run.
    '
    
    '
    'Set up dialogue parms
    '
    FileSaveDialog.Filter = "All files (*.*)|*.*|" & FileExt_NoDot & " files|*." & FileExt_NoDot
'    FileSaveDialog.Filter = "All files (*.*)|*.*|CSV files|*.csv"
'    FileSaveDialog.Filter = "All files (*.*)|*.*|Text files|*.txt"
    FileSaveDialog.FilterIndex = 2
    FileSaveDialog.DefaultExt = FileExt_NoDot
    FileSaveDialog.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or _
                           cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
    FileSaveDialog.DialogTitle = TheDialogTitle
    FileSaveDialog.Filename = TheFileName
    
    FileSaveDialog.CancelError = True ' Allow Exit if user presses Cancel.
    
    FileSaveDialog.ShowSave
    TheFileName = FileSaveDialog.Filename
    SaveTextFile = True

ExitNow:

End Function

Public Sub CloseAllOpenForms(ReLogOn As Boolean, Optional GeneralSettings As Boolean)
Dim frm As Form

On Error Resume Next

    If IsMissing(GeneralSettings) Then
        GeneralSettings = False
    End If
    
    If ReLogOn Then
        For Each frm In Forms
            If frm.Name <> "frmMainMenu" And _
               frm.Name <> "frmSetUpMenu" Then
    '              frm.Hide            ' hide the form
                  WriteToLogFile "Closing " & frm.Name
                  Unload frm          ' deactivate the form
                  Set frm = Nothing   ' remove from memory
            End If
        Next
    Else
        If GeneralSettings Then
            For Each frm In Forms
                If frm.Name <> "frmMainMenu" And _
                   frm.Name <> "frmSetUpMenu" And _
                   frm.Name <> "frmOptionsMenu" And _
                   frm.Name <> "frmFieldMinistryMenu" And _
                   frm.Name <> "frmTMSMenu" And _
                   frm.Name <> "frmGeneralSettings" Then
        '              frm.Hide            ' hide the form
                      WriteToLogFile "Closing " & frm.Name
                      Unload frm          ' deactivate the form
                      Set frm = Nothing   ' remove from memory
                End If
            Next
        Else
            For Each frm In Forms
                If frm.Name <> "frmMainMenu" Then
                    WriteToLogFile "Closing " & frm.Name
                    Unload frm          ' deactivate the form
                    Set frm = Nothing   ' remove from memory
                End If
            Next
        End If
    End If


End Sub

Public Function FormIsOpen(ByVal FormName As String) As Boolean
Dim frm As Form

On Error GoTo ErrorTrap

    For Each frm In Forms
        If frm.Name = FormName Then
            FormIsOpen = True
            Set frm = Nothing
            Exit Function
        End If
    Next

    FormIsOpen = False
    
    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Sub HideAllOpenFormsExceptMenus()
Dim frm As Form

On Error GoTo ModalFormDetected

    For Each frm In Forms
        If frm.Name <> "frmMainMenu" And _
           frm.Name <> "frmSetUpMenu" And _
           frm.Name <> "frmTMSMenu" And _
           frm.Name <> "frmLogin" And _
           frm.Name <> "frmFieldMinistryMenu" And _
           frm.Name <> "frmSplash" Then
              frm.Hide            ' hide the form
        End If
    Next
    
    CanNotShow = False
    
    frmMainMenu.SetFocus

    Exit Sub
    
ModalFormDetected:
    '
    'can't hide a modal form
    '
    CanNotShow = True
    Beep
End Sub

Public Sub ShowAllOpenFormsExceptMenus()
Dim frm As Form

On Error GoTo ErrorTrap

    If Not CanNotShow Then
        For Each frm In Forms
            If frm.Name <> "frmMainMenu" And _
               frm.Name <> "frmSetUpMenu" And _
               frm.Name <> "frmTMSMenu" And _
               frm.Name <> "frmPublicMeetingMenu" And _
               frm.Name <> "frmLogin" And _
               frm.Name <> "frmFieldMinistryMenu" And _
               frm.Name <> "frmSplash" Then
                  frm.Show
            End If
        Next
    Else
        Beep
    End If
    

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Public Sub BringForwardMainMenuWhenItsTheLastFormOpen()

    On Error Resume Next 'in case of error 5 on the setfocus
        
    Select Case True
    Case Forms.Count = 2 'ie current form and frmMainMenu
        frmMainMenu.SetFocus
    Case Forms.Count = 3 And (FormIsOpen("frmSetUpMenu") Or _
                              FormIsOpen("frmTMSMenu") Or _
                              FormIsOpen("frmPublicMeetingMenu") Or _
                              FormIsOpen("frmCustomRotas") Or _
                              FormIsOpen("frmMeetingAttendance") Or _
                              FormIsOpen("frmFieldMinistryMenu") Or _
                              FormIsOpen("frmServiceMtg") Or _
                              FormIsOpen("frmPersonalDetails"))
        frmMainMenu.SetFocus
    End Select
    
'    If Err.number = 5 Then 'can't setfocus as current form is modal
'        frmMainMenu.BringMeForward CurrentForm
'    End If

End Sub

Public Function PersonHasAccess(UserCode As Long, AccessLevel As SecurityAccessLevels) As Boolean
'
                             
Dim rstSecurityLevels As Recordset
                                 
On Error GoTo ErrorTrap

    PersonHasAccess = False
                                                                                                            
    Set rstSecurityLevels = CMSDB.OpenRecordset("SELECT * " & _
                                       "FROM tblAccessLevels " & _
                                       "WHERE UserCode = " & UserCode & _
                                       " AND AccessLevel = " & AccessLevel, _
                                       dbOpenSnapshot)
    
    If rstSecurityLevels.BOF Then
        PersonHasAccess = False
    Else
        PersonHasAccess = True
    End If
    
    rstSecurityLevels.Close
    
    Exit Function
ErrorTrap:
    EndProgram


End Function

Function IsNumber(TheNumber As String, _
                  Optional DecValue As Boolean = False, _
                  Optional AllowSign As Boolean = False, _
                  Optional AllowInitialAndTrailingDP As Boolean = False) As Boolean
    
    Dim i As Integer
    
    If Not IsNumeric(TheNumber) Then
        IsNumber = False
        Exit Function
    End If
    
    If IsMissing(DecValue) Then
        DecValue = False
    End If
    If IsMissing(AllowSign) Then
        AllowSign = False
    End If
    If IsMissing(AllowInitialAndTrailingDP) Then
        AllowInitialAndTrailingDP = False
    End If
    
    For i = 1 To Len(TheNumber)
        Select Case Mid$(TheNumber, i, 1)
        Case "0" To "9"
        Case "-", "+"
            If AllowSign Then
                ' Minus/plus signs are only allowed as leading chars.
                If i > 1 Then Exit Function
            Else
                Exit Function
            End If
        Case "."
            ' Exit if decimal values not allowed.
            If Not DecValue Then Exit Function
            ' Only one decimal separator is allowed.
            If InStr(TheNumber, ".") < i Then Exit Function
            ' Decimal point can't be at start or end
            If InStr(TheNumber, ".") = 1 Or _
               InStr(TheNumber, ".") = Len(TheNumber) Then
                If Not AllowInitialAndTrailingDP Then
                    Exit Function
                End If
            End If
        Case Else
            ' Reject all other characters.
            Exit Function
        End Select
    Next i
    
    IsNumber = True
    
End Function

Function GetLettersForOrdinalNumber(TheNumber As Long) As String
'
'Put the 'th' or 'st' or 'rd' after ordinal number as appropriate
'
    
    Dim NoAsString As String
    
    NoAsString = CStr(TheNumber)
    
    Select Case Right(NoAsString, 1)
    Case 0
        GetLettersForOrdinalNumber = "th"
    Case 1
        GetLettersForOrdinalNumber = "st"
        If Len(NoAsString) > 1 Then
            '
            'eg compare '11th' with '21st'
            '              ^^          ^^
            
            If Left((Right(NoAsString, 2)), 1) = 1 Then
                GetLettersForOrdinalNumber = "th"
            Else
                GetLettersForOrdinalNumber = "st"
            End If
        Else
            GetLettersForOrdinalNumber = "st"
        End If
    Case 2
        If Len(NoAsString) > 1 Then
            If Left((Right(NoAsString, 2)), 1) = 1 Then
                GetLettersForOrdinalNumber = "th"
            Else
                GetLettersForOrdinalNumber = "nd"
            End If
        Else
            GetLettersForOrdinalNumber = "nd"
        End If
    Case 3
        If Len(NoAsString) > 1 Then
            If Left((Right(NoAsString, 2)), 1) = 1 Then
                GetLettersForOrdinalNumber = "th"
            Else
                GetLettersForOrdinalNumber = "rd"
            End If
        Else
            GetLettersForOrdinalNumber = "rd"
        End If
    Case 4
        GetLettersForOrdinalNumber = "th"
    Case 5
        GetLettersForOrdinalNumber = "th"
    Case 6
        GetLettersForOrdinalNumber = "th"
    Case 7
        GetLettersForOrdinalNumber = "th"
    Case 8
        GetLettersForOrdinalNumber = "th"
    Case 9
        GetLettersForOrdinalNumber = "th"
    Case Else
        GetLettersForOrdinalNumber = ""
    End Select
    
    
End Function


Sub OpenWindowsExplorer(ByVal ThePath As String, _
                        ByVal SelectFile As Boolean, _
                        ByVal SinglePane As Boolean, _
                        Optional ByVal WindowSize As VbAppWinStyle = vbMaximizedFocus)
    
' Open Explorer window. If SelectFile is TRUE, then ThePath must lead to the file
'  to be selected.

    Shell "explorer " & IIf(SinglePane, " /n", " /e") & _
          IIf(SelectFile, ", /select,", ", /root, ") & """" & ThePath & """", WindowSize
        
End Sub

Function StringCount(source As String, Search As String) As Long
    ' You get the number of substrings by subtracting the length of the
    ' original string from the length of the string that you obtain by
    ' replacing the substring with another string that is one char longer.
    StringCount = Len(Replace(source, Search, Search & "*")) - Len(source)
End Function
Function ConvertPixelsToTwipsX(ByVal lPixels_X As Long) As Long

    ConvertPixelsToTwipsX = Screen.TwipsPerPixelX * lPixels_X

End Function
Function ConvertPixelsToTwipsY(ByVal lPixels_Y As Long) As Long

    ConvertPixelsToTwipsY = Screen.TwipsPerPixelY * lPixels_Y

End Function
Public Function ConvertForeignChars(TheChar As String) As String

    Select Case True
    Case InStr(1, "àáâãäå", TheChar) > 0
        ConvertForeignChars = "a"
    Case InStr(1, "èéêë", TheChar) > 0
        ConvertForeignChars = "e"
    Case InStr(1, "ìíîîï", TheChar) > 0
        ConvertForeignChars = "i"
    Case InStr(1, "ðòôóõöø", TheChar) > 0
        ConvertForeignChars = "o"
    Case InStr(1, "ùúûü", TheChar) > 0
        ConvertForeignChars = "u"
    Case InStr(1, "ÁÂÃÄÅ", TheChar) > 0
        ConvertForeignChars = "A"
    Case InStr(1, "ÈÉÊË", TheChar) > 0
        ConvertForeignChars = "E"
    Case InStr(1, "ÌÍÎÏ", TheChar) > 0
        ConvertForeignChars = "I"
    Case InStr(1, "ÒÓÔÕÖØ", TheChar) > 0
        ConvertForeignChars = "O"
    Case InStr(1, "ÙÚÛÜ", TheChar) > 0
        ConvertForeignChars = "U"
    Case Else
        ConvertForeignChars = TheChar
    End Select
    
End Function
Public Function ConvertStringWithForeignChars(TheString As String) As String
Dim i As Long, str As String

    str = ""
    
    For i = 1 To Len(TheString)
        str = str & ConvertForeignChars(Mid(TheString, i, 1))
    Next i
    
    ConvertStringWithForeignChars = str
    
End Function

Sub SetGridRowForeColour(ByVal TheGrid As MSFlexGrid, _
                         ByVal TheRow As Long, _
                         ByVal VBColour As Integer)
                         
Dim SaveCol As Long, SaveRow As Long, i As Long
    
    With TheGrid
    
    SaveCol = .col
    SaveRow = .Row
    
    .Row = TheRow
    
    For i = 1 To .Cols - 1
        .col = i
        .CellForeColor = QBColor(VBColour)
    Next i
    
    .col = SaveCol
    .Row = SaveRow
    
    End With
    
End Sub
Public Sub TextFieldGotFocus(txtField As TextBox, _
                             Optional SetTheFocus As Boolean = False)
                             
On Error Resume Next 'in case of error 5 if set focus on loading form

    If SetTheFocus Then txtField.SetFocus
    txtField.SelStart = 0
    txtField.SelLength = Len(txtField.text)

End Sub
Public Sub AutoCompleteComboKeyDown(TheCombo As ComboBox, KeyCode As Integer)
                             
    If KeyCode = 46 Then
        TheCombo.ListIndex = -1
    End If

End Sub
Public Sub AutoCompleteComboLostFocus(TheCombo As ComboBox, _
                                      Optional DefaultListIndex As Integer = 0)
                    
    'use this in combo's lost focus event only if there is a [None] or [All] option
    ' with ListIndex=0
    
    If TheCombo.text = "" Then
        TheCombo.ListIndex = DefaultListIndex
    End If

End Sub
Public Function AutoCompleteCombo(TheCombo As ComboBox, KeyAscii As Integer) As Long
Dim sOldEnteredText As String, sComboText As String, lMaxCount As Long
Dim lEntryPoint As Long, sEnteredChar As String, i As Long
Dim sNewEnteredText As String, bMatched As Boolean, bBackSpace As Boolean

    With TheCombo
        
    If KeyAscii = 13 Or KeyAscii = 27 Then Exit Function 'Carriage Return or escape
    
    lEntryPoint = .SelStart 'Current text entry point
    lMaxCount = .ListCount - 1
    
    sEnteredChar = Chr$(KeyAscii) 'New text entry
        
    sOldEnteredText = Left$(.text, .SelStart) 'text keyed by user prior to this new entry
    
    If KeyAscii <> 8 Then 'not backspace
        sNewEnteredText = sOldEnteredText & sEnteredChar 'append newly entered char
        bBackSpace = False
    Else
        'backspace
        If Len(sOldEnteredText) = 0 Then
            .ListIndex = -1
            Exit Function
        End If
        'remove a char from what user previously entered
        sNewEnteredText = Left$(sOldEnteredText, Len(sOldEnteredText) - 1)
        bBackSpace = True
    End If
            
    
    
    KeyAscii = 0 'don't want VB to put text in combo
    
    If sNewEnteredText = "" Then
        .ListIndex = -1
        Exit Function
    End If
    
    bMatched = False 'initialise
    'search combo for 1st entry which matches user's entry
    For i = 0 To lMaxCount
'        If UCase$(ConvertStringWithForeignChars(Left(.List(i), Len(sNewEnteredText)))) = UCase$(sNewEnteredText) Then
        If UCase$(Left(.List(i), Len(sNewEnteredText))) = UCase$(sNewEnteredText) Then
            bMatched = True
            Exit For
        End If
    Next i
    
    If bMatched Then
        'Match found, so select entry in combo and set new text entrypoint
        .ListIndex = i
        .SelStart = Len(sNewEnteredText)
'        AutoCompleteCombo = i
    Else
        'No match found so set text entry point to what it was before
        .SelStart = Len(sOldEnteredText)
'        AutoCompleteCombo = .ListIndex
    End If
    
    .SelLength = Len(.text) - .SelStart 'show highlight on selected text
    
    End With

End Function

Public Function HandleNull(ByVal TheValue, Optional ByVal ValueIfNull = 0) As Variant

    If IsNull(TheValue) Then
        HandleNull = ValueIfNull
    Else
        On Error Resume Next
        HandleNull = TheValue
        If Err.number <> 0 Then
            HandleNull = ValueIfNull
        End If
    End If

End Function
Public Function GetYearOfBirthFromAge(lAge As Long) As Long

    GetYearOfBirthFromAge = year(Now) - lAge

End Function
Public Function GetAgeInYears(DateOfBirth As Date) As Long

    GetAgeInYears = Fix(DateDiff("m", DateOfBirth, Now) / 12)

End Function
Public Function GetAgeAsYearsAndMonths(DateOfBirth As Date) As String
Dim lyears As Long, lMonths As Long

    lyears = Fix(DateDiff("m", DateOfBirth, Now) / 12)
    lMonths = DateDiff("m", DateOfBirth, Now) - 12 * lyears
    
    GetAgeAsYearsAndMonths = lyears & IIf(lyears <> 1, " Years ", " Year ") & _
                             lMonths & IIf(lMonths <> 1, " Months", " Month")

End Function

Public Function ListItemID(ListBox As Control) As Long

    With ListBox
    
    If .ListIndex > -1 Then
        ListItemID = .ItemData(.ListIndex)
    Else
        ListItemID = -1
    End If
    
    End With

End Function

' returns True if an item in a collection actually exists

Function ItemExistsInCollection(col As Collection, Key As String) As Boolean
    Dim dummy As Variant
    On Error Resume Next
    dummy = col.Item(Key)
    ItemExistsInCollection = (Err = 0)
End Function

Public Sub SelectFlexGridCells(ByVal FlexGrid As MSFlexGrid, _
                               ByVal TheRow As Long, _
                               ByVal TheCol As Long)

    FlexGrid.RowSel = TheRow
    FlexGrid.ColSel = TheCol

End Sub
Public Sub ShadeOddGridRowBands(ByVal FlexGrid As MSFlexGrid, _
                                ByVal TheColour As Long)
Dim i As Long, j As Long
    'shade all odd rows
    With FlexGrid
    For j = 1 To .Rows - 1 Step 2
        .Row = j
        For i = 0 To .Cols - 1
            .col = i
            .CellBackColor = TheColour ' light grey
        Next i
    Next j
    End With
End Sub

Public Sub SetTopRowOfGridToWeek(flxGrid As MSFlexGrid, _
                                  lRowsAboveSelected As Long, _
                                  lDateColumn, _
                                  Optional ByVal dteSetToDate)
                                  
On Error GoTo ErrorTrap
Dim i As Long, TheDate As Date
    
    'set top row of grid so that current week is displayed
    If IsMissing(dteSetToDate) Then
        dteSetToDate = Now
    End If
    
    TheDate = CDate(Format(GetDateOfGivenDay(CDate(dteSetToDate), vbMonday), "dd/mm/yyyy"))
    
    With flxGrid
    
    If .Rows > 1 Then
        If TheDate > CDate(.TextMatrix(.Rows - 1, lDateColumn)) Or _
           TheDate < CDate(.TextMatrix(1, lDateColumn)) Then
            .TopRow = 1
            Exit Sub
        End If
    End If
    
    For i = 1 To .Rows - 1
        If CDate(.TextMatrix(i, lDateColumn)) >= TheDate Then
            If i > lRowsAboveSelected Then
                .TopRow = i - lRowsAboveSelected
            Else
                .TopRow = 1
            End If
            Exit For
        End If
    Next
    End With
    
    Exit Sub
ErrorTrap:
    Call EndProgram
    
End Sub

Public Function FileNameValid(sFileName) As Boolean
Dim arr() As String, i As Long

    If sFileName = "" Then
        FileNameValid = False
        Exit Function
    End If
    
    arr() = Split("\ / : "" * ? < > |")
    
    sFileName = DoubleUpSingleQuotes(sFileName)
    
    For i = 0 To UBound(arr)
        If InStr(1, sFileName, arr(i)) > 0 Then
            FileNameValid = False
            Exit Function
        End If
    Next i
    
    FileNameValid = True

End Function
Public Function GetFinancialYear(TheDate As Date) As Long

    Select Case Month(TheDate)
    Case 1 To 3
      GetFinancialYear = year(TheDate) - 1
    Case Else
      GetFinancialYear = year(TheDate)
    End Select

End Function
Public Function GetNormalYearFromFinancialYear(TheDate As Date) As Long

    Select Case Month(TheDate)
    Case 1 To 3
      GetNormalYearFromFinancialYear = year(TheDate) + 1
    Case Else
      GetNormalYearFromFinancialYear = year(TheDate)
    End Select

End Function
Public Function GetFinancialMonth(TheDate As Date) As Long

    Select Case Month(TheDate)
    Case 1 To 3
      GetFinancialMonth = Month(TheDate) + 9
    Case Else
      GetFinancialMonth = Month(TheDate) - 3
    End Select

End Function
Public Function GetFinancialQuarter(TheDate As Date) As Long

    Select Case Month(TheDate)
    Case 1 To 3
      GetFinancialQuarter = 4
    Case 4 To 6
      GetFinancialQuarter = 1
    Case 7 To 9
      GetFinancialQuarter = 2
    Case 10 To 12
      GetFinancialQuarter = 3
    End Select

End Function
Public Function GetNormalQuarter(TheDate As Date) As Long

    Select Case Month(TheDate)
    Case 1 To 3
      GetNormalQuarter = 1
    Case 4 To 6
      GetNormalQuarter = 2
    Case 7 To 9
      GetNormalQuarter = 3
    Case 10 To 12
      GetNormalQuarter = 4
    End Select

End Function

Public Function DeleteItemFromListBox(lst As Control, TheListIndex As Long) As Boolean

    If TypeOf lst Is ListBox Or TypeOf lst Is ComboBox Then
    Else
        Exit Function
    End If
    
    With lst
    If .ListIndex = -1 Then
        Exit Function
    End If
    
    .RemoveItem TheListIndex
    
    End With
    
    DeleteItemFromListBox = True

End Function
Public Sub SetTextBoxInsertPoint(TheTextBox As Control, _
                                 InsertPoint As Long, _
                                 SetTheFocus As Boolean)

'if InsertPoint is -1, caret moved to END of text
'if InsertPoint is -2, caret moved to START of text
'otherwise, caret is moved to whatever InsertPoint is

    If TypeOf TheTextBox Is TextBox Or TypeOf TheTextBox Is RichTextBox Then
    Else
        Exit Sub
    End If
    
    Select Case InsertPoint
    Case -2
        TheTextBox.SelStart = 0
    Case -1
        TheTextBox.SelStart = Len(TheTextBox)
    Case Else
        TheTextBox.SelStart = Abs(InsertPoint)
    End Select
    
    TheTextBox.SelLength = 0
    
    If SetTheFocus Then
        TheTextBox.SetFocus
    End If
    
End Sub

Public Sub RowShadingGroups(TheFlexGrid As MSFlexGrid, _
                            GroupingColumn As Long, _
                            Colour1 As Long, _
                            Colour2 As Long)
Dim sPrev As String, i As Long, j As Long
Dim bSwitch As Boolean, lColour As Long
On Error GoTo ErrorTrap
    With TheFlexGrid
    'shade rows differently on change of date
    For j = 1 To .Rows - 1
        .Row = j
        If sPrev <> .TextMatrix(j, GroupingColumn) Then
            bSwitch = Not bSwitch
            lColour = IIf(bSwitch, Colour1, Colour2)
        End If
        For i = 0 To .Cols - 1
            .col = i
            .CellBackColor = lColour
        Next i
        sPrev = .TextMatrix(j, GroupingColumn)
    Next j
    End With
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Public Sub ExportFlexGridToCSV(TheFileDir_NoFinalSlash As String, _
                                ExportTitle_NoFileExt As String, _
                                FileExt_NoDot As String, _
                                FileSaveDialogueTitle As String, _
                                TheFlexGrid As MSFlexGrid, _
                                TheCommonDialog As CommonDialog, _
                                Optional HasColumnHeaders As Boolean = True)

Dim FilePath As String, FileNum As Integer, rstPrintScheduleRecSet As Recordset
Dim StringToPrint As String, FileIsOpen As Boolean, i As Long, j As Long, str As String
Dim bEncloseInQuotes As Boolean

On Error GoTo ErrorTrap

    If TheFlexGrid.Rows < IIf(HasColumnHeaders, 2, 1) Then
        MsgBox "Nothing to export", vbExclamation + vbOKOnly, AppName
        Exit Sub
    End If

    bEncloseInQuotes = GlobalParms.GetValue("CSVExportHasDoubleQuotes", "TrueFalse")
    
    '
    'Create a text file in which to store the data from the grid.
    '
    
    str = Replace(ExportTitle_NoFileExt & " " & Now, ":", "-")
    FilePath = Replace(TheFileDir_NoFinalSlash & "\" & str & "." & FileExt_NoDot, "/", "-")
    
    '
    'Opens filesave dialogue
    '
    If Not SaveTextFile(FilePath, _
           FileSaveDialogueTitle, _
            TheCommonDialog, "csv") Then
        Exit Sub
    End If
    
    FileNum = FreeFile()
    Open FilePath For Output As #FileNum
    
    With TheFlexGrid
    
    'if first 2 letters of Excel file are "ID", file will not open (?!) Convert to "id"...
    'This isn't a problem when fields enclosed in quotes.
    If Not bEncloseInQuotes Then
        If Left$(.TextMatrix(0, 0), 2) = "ID" Then
            .TextMatrix(0, 0) = Replace$(.TextMatrix(0, 0), "ID", "id", 1)
        End If
    End If
    
    For i = 0 To .Rows - 1
    
        StringToPrint = ""
        
        For j = 0 To .Cols - 1
        
            If bEncloseInQuotes Then
                StringToPrint = StringToPrint & """" & .TextMatrix(i, j) & """"
            Else
                StringToPrint = StringToPrint & Replace(.TextMatrix(i, j), ",", " ")
            End If
            
            If j < .Cols - 1 Then
                StringToPrint = StringToPrint & ","
            End If
            
        Next j
        
        StringToPrint = StringToPrint & vbCrLf
        Print #FileNum, StringToPrint;
        
    Next i
    
    End With
    
    Close #FileNum
        
    If MsgBox("Data successfully exported as '" & FilePath & _
           "'. You may open this file in a spreadsheet. Do you want to " & _
           "work with the file now?", vbYesNo + vbQuestion, AppName) = vbYes Then
           
           On Error Resume Next 'Any prob opening Explorer - don't abend prog.
           
           OpenWindowsExplorer FilePath, True, True
           
           If Err.number > 0 Then
                MsgBox "Error opening explorer.", vbOKOnly + vbExclamation, AppName
           End If
           
           On Error GoTo ErrorTrap
    
    End If
    
    Exit Sub
ErrorTrap:
    If FileIsOpen Then
        Close #FileNum
    End If
    EndProgram
End Sub
Public Sub CopyFlexGridToClipboard(TheFlexGrid As MSFlexGrid, _
                                   Optional HasColumnHeaders As Boolean = True)

Dim rstPrintScheduleRecSet As Recordset
Dim StringToPrint As String, i As Long, j As Long, str As String

On Error GoTo ErrorTrap

    If TheFlexGrid.Rows < IIf(HasColumnHeaders, 2, 1) Then
        MsgBox "Nothing to copy", vbExclamation + vbOKOnly, AppName
        Exit Sub
    End If

    With TheFlexGrid
        
    StringToPrint = ""
    
    For i = 0 To .Rows - 1
    
        
        For j = 0 To .Cols - 1
            StringToPrint = StringToPrint & Replace(.TextMatrix(i, j), ",", " ")
            If j < .Cols - 1 Then
                StringToPrint = StringToPrint & ","
            End If
        Next j
        
        StringToPrint = StringToPrint & vbCrLf
        
    Next i
    
    End With
                 
    CopyTextToClipBoard StringToPrint
           
    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Public Sub SetTabStops(frm As Form, value As Boolean)
    Dim ctrl As Control
    On Error Resume Next
    For Each ctrl In frm.Controls
        ctrl.TabStop = value
    Next
End Sub
Public Sub FontDialogue(TheCommonDialog As CommonDialog, _
                        TheRichTextBox As RichTextBox)
On Error Resume Next

With TheCommonDialog
    .Flags = cdlCFBoth Or cdlCFForceFontExist Or cdlCFEffects
    If IsNull(TheRichTextBox.SelFontName) Then
        .Flags = .Flags Or cdlCFNoFaceSel
    Else
        .FontName = TheRichTextBox.SelFontName
    End If
    If IsNull(TheRichTextBox.SelFontSize) Then
        .Flags = .Flags Or cdlCFNoSizeSel
    Else
        .FontSize = TheRichTextBox.SelFontSize
    End If
    If IsNull(TheRichTextBox.SelBold) Or _
       IsNull(TheRichTextBox.SelItalic) Or _
       IsNull(TheRichTextBox.SelUnderline) Or _
       IsNull(TheRichTextBox.SelStrikeThru) Then
        .Flags = .Flags Or cdlCFNoStyleSel
    Else
        .FontBold = TheRichTextBox.SelBold
        .FontItalic = TheRichTextBox.SelItalic
        .FontUnderline = TheRichTextBox.SelUnderline
        .FontStrikethru = TheRichTextBox.SelStrikeThru
    End If
    If IsNull(TheRichTextBox.SelColor) Then
        .Flags = .Flags Or cdlCFNoStyleSel
    Else
        .Color = TheRichTextBox.SelColor
    End If
    .CancelError = True
    .ShowFont
    If Err = 0 Then
        TheRichTextBox.SelFontName = .FontName
        TheRichTextBox.SelBold = .FontBold
        TheRichTextBox.SelItalic = .FontItalic
        TheRichTextBox.SelColor = .Color
        If (.Flags And cdlCFNoSizeSel) = 0 Then
            TheRichTextBox.SelFontSize = .FontSize
        End If
        TheRichTextBox.SelUnderline = .FontUnderline
        TheRichTextBox.SelStrikeThru = .FontStrikethru
    End If
End With

End Sub

Public Function ItemDataInCombo(TheCombo As ComboBox, TheItemData As Long) As Boolean
Dim i As Long, bMatch As Boolean
On Error GoTo ErrorTrap

    With TheCombo
    For i = 0 To .ListCount - 1
        If .ItemData(i) = TheItemData Then
            bMatch = True
            Exit For
        End If
    Next i
    End With
    
    If bMatch Then
        ItemDataInCombo = True
    Else
        ItemDataInCombo = False
    End If
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function
Public Function AddApostropheToPersonName(TheName As String) As String
On Error GoTo ErrorTrap

    AddApostropheToPersonName = IIf(Right(TheName, 1) = "s", TheName & "'", TheName & "'s")
    
    Exit Function
ErrorTrap:
    Call EndProgram
End Function

Public Function CopyTextToClipBoard(TheText As String, Optional TextFormat As Variant = vbCFText) As Boolean

On Error Resume Next

Dim str  As String
          
    Clipboard.Clear
    
    Clipboard.SetText TheText, TextFormat

    If Err.number = 0 Then
        CopyTextToClipBoard = True
    Else
        CopyTextToClipBoard = False
    End If
    
    
End Function

Public Function NumberIsOdd(TheNumber As Long) As Boolean
    
    NumberIsOdd = (TheNumber Mod 2 = 1)
    
End Function
Public Function NumberIsEven(TheNumber As Long) As Boolean
    
    NumberIsEven = (TheNumber Mod 2 = 0)
    
End Function
Public Function RightAlignString(TheString As String, _
                                 StringLenForZeroSpaces As Long _
                                 ) As String

'works best if string is displayed with non-proportional font, eg courier

    RightAlignString = Space(StringLenForZeroSpaces - Len(TheString)) & TheString


End Function


Public Function GetRightOfString(StringToSearch As String, StringToFind As String) As String
Dim lPos As Long

    lPos = InStr(1, StringToSearch, StringToFind)
    
    If lPos = 0 Then
        GetRightOfString = StringToSearch
        Exit Function
    End If
    
    If Len(StringToSearch) - lPos - (Len(StringToFind) - 1) > 0 Then
        GetRightOfString = Right(StringToSearch, Len(StringToSearch) - lPos - (Len(StringToFind) - 1))
    Else
        GetRightOfString = ""
    End If
    
End Function
Public Function GetLeftOfString(StringToSearch As String, StringToFind As String) As String
Dim lPos As Long

    lPos = InStr(1, StringToSearch, StringToFind)
    
    If lPos = 0 Then
        GetLeftOfString = StringToSearch
        Exit Function
    End If
    
    If lPos - 1 > 0 Then
        GetLeftOfString = Left(StringToSearch, lPos - 1)
    Else
        GetLeftOfString = ""
    End If
    
End Function

