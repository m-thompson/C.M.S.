Attribute VB_Name = "basGeneral5"
Option Explicit

Private Declare Function DeleteMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32.dll" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwnewlong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Declare Function tapiRequestMakeCall Lib "TAPI32.DLL" (ByVal Dest As _
    String, ByVal AppName As String, ByVal CalledParty As String, _
    ByVal Comment As String) As Long
    
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const SC_ARRANGE = &HF110
Public Const SC_CLOSE = &HF060
Public Const SC_HOTKEY = &HF150
Public Const SC_HSCROLL = &HF080
Public Const SC_KEYMENU = &HF100
Public Const SC_MAXIMIZE = &HF030
Public Const SC_MINIMIZE = &HF020
Public Const SC_MOVE = &HF010
Public Const SC_NEXTWINDOW = &HF040
Public Const SC_PREVWINDOW = &HF050
Public Const SC_RESTORE = &HF120
Public Const SC_SIZE = &HF000
Public Const SC_VSCROLL = &HF070
Public Const SC_TASKLIST = &HF130
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const GWL_STYLE = (-16)

Public Enum T_WindowStyle
    WS_BORDER = &H800000
    WS_CAPTION = &HC00000
    WS_CHILD = &H40000000
    WS_CHILDWINDOW = (WS_CHILD)
    WS_CLIPCHILDREN = &H2000000
    WS_CLIPSIBLINGS = &H4000000
    WS_DISABLED = &H8000000
    WS_DLGFRAME = &H400000
    WS_EX_ACCEPTFILES = &H10&
    WS_EX_DLGMODALFRAME = &H1&
    WS_EX_NOPARENTNOTIFY = &H4&
    WS_EX_TOPMOST = &H8&
    WS_EX_TRANSPARENT = &H20&
    WS_GROUP = &H20000
    WS_HSCROLL = &H100000
    WS_MAXIMIZE = &H1000000
    WS_MAXIMIZEBOX = &H10000
    WS_MINIMIZE = &H20000000
    WS_MINIMIZEBOX = &H20000
    WS_OVERLAPPED = &H0&
    WS_ICONIC = WS_MINIMIZE
    WS_POPUP = &H80000000
    WS_VISIBLE = &H10000000
    WS_VSCROLL = &H200000
    WS_SYSMENU = &H80000
    WS_TABSTOP = &H10000
    WS_THICKFRAME = &H40000
    WS_TILED = WS_OVERLAPPED
    WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
    WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
    WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
    WS_SIZEBOX = WS_THICKFRAME
End Enum

Public Const UNKNOWN = _
"(Value Unknown Because System Call Failed)"

Public Declare Function GetUserName Lib "advapi32.dll" _
Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" _
Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As _
    Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As _
    Any, source As Any, ByVal numBytes As Long)
    
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Const LB_GETITEMHEIGHT = &H1A1
Const CB_GETITEMHEIGHT = &H154
Private Const EM_SETTABSTOPS = &HCB

Const KEY_READ = &H20019  ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
                          ' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
                          ' SYNCHRONIZE))

Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7
Const ERROR_MORE_DATA = 234

Public Const WINAPI_HandCursor = 32649&
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function GetCursor Lib "user32" () As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

' Remember the handle for the Hand cursor
Public HandCursor As Long

Public Function bOpen(strClass As String, strCaption As String) As Boolean
    bOpen = CBool(FindWindow(strClass$, strCaption$))
End Function


Public Sub SetHandCursor()
    On Error Resume Next

    ' Load the cursor only once
    If HandCursor = 0 Then
        HandCursor = LoadCursor(0, WINAPI_HandCursor)

    ' Only update the cursor if it isn't our current cursor
    ElseIf GetCursor() = HandCursor Then
        Exit Sub
    End If

    ' Finally, set the cursor to the Hand
    SetCursor HandCursor
End Sub

Function OpenEmailProgram(sDest As String, Optional sSubject As String = "", _
    Optional sBody As String = "", Optional sCC As String = "", Optional sBCC As String = "", _
    Optional ValidateEmailAddr As Boolean = True) As Boolean
    On Error Resume Next
            
    Dim str As String
            
    'to make OE Express the default, may need to run this on Run command line:
    ' <"C:\Program Files\Outlook Express\msimn.exe" /reg>
        
    If ValidateEmailAddr Then
        If Not ValidateEmailAddress(sDest) Then
            OpenEmailProgram = False
            Exit Function
        End If
    End If
    
    str = sDest
    
    If sSubject <> "" Then
        str = str & "?subject=" & sSubject
    End If
    
    'when using ShellExecute to open default email program,
    ' vbCrLf doesn't work. Need to use %0D%0A.
    If GlobalParms.GetValue("EmailShellExecuteLineBreak", "TrueFalse", "TRUE") = True Then
        sBody = Replace(sBody, vbCrLf, "%0D%0A")
    End If
    
    If sBody <> "" Then
        If str = sDest Then
            str = str & "?body=" & sBody
        Else
            str = str & "&body=" & sBody
        End If
    End If
        
    If sCC <> "" Then
        If str = sDest Then
            str = str & "?CC=" & sCC
        Else
            str = str & "&CC=" & sCC
        End If
    End If
        
    If sBCC <> "" Then
        If str = sDest Then
            str = str & "?BCC=" & sBCC
        Else
            str = str & "&BCC=" & sBCC
        End If
    End If
    
    ShellExecute 0, vbNullString, "mailto:" & str, 0&, 0&, 1
    
'    ShellExecute 0, vbNullString, "mailto:" & sDest & "?subject=" & sSubject & _
'        "&body=" & sBody & "&CC=" & sCC & "&BCC=" & sBCC, 0&, 0&, 1
        
    If Err.number > 0 Then
        OpenEmailProgram = False
    Else
        OpenEmailProgram = True
    End If
        
End Function


' make a phone call using TAPI
' return True if successfull
'
' Go to Control Panel|Phone and Modem Options to set up dialling settings

Function PhoneCall(ByVal PhoneNumber As String, ByVal DestName As String, _
    Optional ByVal Comment As String) As Boolean
    If tapiRequestMakeCall(Trim$(PhoneNumber), App.Title, Trim$(DestName), _
        Comment) = 0 Then
        PhoneCall = True
    End If
End Function


'Brings the specified form in the top most position, it will be over all other
'forms in the screen, even if they will receive the focus
Public Sub FormTopMost(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
End Sub

'Brings the form in his standard Z-Order
Public Sub FormNoTopMost(hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
End Sub

'Brings the form in the top position of the Z-Order, if another form takes the
'focus it will become the new top form
Public Sub FormTop(hWnd As Long)
    SetWindowPos hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
End Sub

'Remove the form's system menu, if RemoveClose is true the Close command inside the
'menu is removed too, in this case even the X key in the right upper cornet of the
'form will be removed
Public Sub RemoveSystemMenu(hWnd As Long, RemoveClose As Boolean)
    Dim hMenu As Long

    hMenu = GetSystemMenu(hWnd, False)
    DeleteMenu hMenu, SC_MAXIMIZE, MF_BYCOMMAND
    DeleteMenu hMenu, SC_MINIMIZE, MF_BYCOMMAND
    DeleteMenu hMenu, SC_SIZE, MF_BYCOMMAND
    DeleteMenu hMenu, SC_MOVE, MF_BYCOMMAND
    DeleteMenu hMenu, SC_RESTORE, MF_BYCOMMAND
    DeleteMenu hMenu, SC_NEXTWINDOW, MF_BYCOMMAND
    If RemoveClose Then
        DeleteMenu hMenu, SC_CLOSE, MF_BYCOMMAND
        DeleteMenu hMenu, 0, MF_BYPOSITION
    End If
End Sub

'Hides the upper right keys Maximize and minimize
Public Sub RemoveMaxMinButtons(hWnd As Long)
    Dim X As Long

    X = GetWindowLong(hWnd, GWL_STYLE)
    X = X And Not WS_MINIMIZEBOX
    X = X And Not WS_MAXIMIZEBOX
    SetWindowLong hWnd, GWL_STYLE, X
End Sub

'Shows the upper right keys Maximize and minimize
Public Sub AddMaxMinButtons(hWnd As Long)
    Dim X As Long

    X = GetWindowLong(hWnd, GWL_STYLE)
    X = X Or WS_MINIMIZEBOX
    X = X Or WS_MAXIMIZEBOX
    SetWindowLong hWnd, GWL_STYLE, X
End Sub

'Set the attribute of a window: the module has a public enum type that contains
'all the constants to define a window style (used by others Subs)
Public Sub SetWindowStyle(hWnd As Long, mAttribute As T_WindowStyle, Enable As Boolean)
    Dim X As Long

    X = GetWindowLong(hWnd, GWL_STYLE)
    If Enable Then
        X = X Or mAttribute
    Else
        X = X And Not mAttribute
    End If
    SetWindowLong hWnd, GWL_STYLE, X
End Sub

Public Sub ControlForTestOnly(ByVal TheControl As Control, _
                              ByVal MakeInvisible As Boolean, _
                              ByVal MakeDisabled As Boolean)

    If TheMDBFile = "CMS_LIVE" Then
        If MakeDisabled Then
            TheControl.Enabled = False
        Else
            TheControl.Enabled = True
        End If
        If MakeInvisible Then
            TheControl.Visible = False
        Else
            TheControl.Visible = True
        End If
    Else
        TheControl.Visible = True
        TheControl.Enabled = True
    End If
        
End Sub
Public Function ForTestOnly() As Boolean

    ForTestOnly = (TheMDBFile <> "CMS_LIVE")
        
End Function

Public Function GetComputerNameTool() As String

    Dim strString As String
    Dim lngBuffer As Long
    Dim lngRetVal As Long
    
    strString = String$(255, " ")
    lngBuffer = 255
    
    lngRetVal = GetComputerName(strString, lngBuffer)
    
    If lngRetVal <> 0 Then
        GetComputerNameTool = Left(strString, lngBuffer)
    Else
        GetComputerNameTool = "Not available"

'        Err.Raise Err.LastDllError, , _
'            "A system call returned an error code of " _
'            & Err.LastDllError
    End If
    
End Function

Public Function IsOfficeAppPresent(ByVal mbOfficeApp As cmsOfficeAppConstants) As _
                                                                            Boolean

' Check whether the specified Office application is present
' Note: require GetRegistryValue
'
' Example:
'    Dim sDescr As String
'    sDescr = "Word is installed: " & IsOfficeAppPresent(mbWord) & vbCrLf & '
'      "Access is installed: " & IsOfficeAppPresent(mbAccess) & vbCrLf & '
'   "Excel is installed: " & IsOfficeAppPresent(mbExcel) & vbCrLf & '
' "Powerpoint is installed: " & IsOfficeAppPresent(mbPowerpoint) & vbCrLf & '
'      "Outlook is installed: " & IsOfficeAppPresent(mbOutlook)
'    MsgBox sDescr
    
    
    Dim sApp As String
    Const HKEY_CLASSES_ROOT = &H80000000
    
    Select Case mbOfficeApp
        Case cmsWord
            sApp = "Word.Document\CurVer"
        Case cmsAccess
            sApp = "Access.Database\CurVer"
        Case cmsExcel
            sApp = "Excel.Sheet\CurVer"
        Case cmsPowerpoint
            sApp = "PowerPoint.Slide\CurVer"
        Case cmsOutlook
            sApp = "Outlook.Envelope\CurVer"
    End Select
    
    'if it reads a value, the key exists --> the application is installed
    IsOfficeAppPresent = (Len(GetRegistryValue(HKEY_CLASSES_ROOT, sApp, _
        "")) > 0)
        
End Function

' Read a Registry value
'
' Use KeyName = "" for the default value
' If the value isn't there, it returns the DefaultValue
' argument, or Empty if the argument has been omitted
'
' Supports DWORD, REG_SZ, REG_EXPAND_SZ, REG_BINARY and REG_MULTI_SZ
' REG_MULTI_SZ values are returned as a null-delimited stream of strings
' (VB6 users can use SPlit to convert to an array of string)

Function GetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, _
    ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
    Dim handle As Long
    Dim resLong As Long
    Dim resString As String
    Dim resBinary() As Byte
    Dim length As Long
    Dim retVal As Long
    Dim valueType As Long
    
    ' Prepare the default result
    GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)
    
    ' Open the key, exit if not found.
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then
        Exit Function
    End If
    
    ' prepare a 1K receiving resBinary
    length = 1024
    ReDim resBinary(0 To length - 1) As Byte
    
    ' read the registry key
    retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
        length)
    ' if resBinary was too small, try again
    If retVal = ERROR_MORE_DATA Then
        ' enlarge the resBinary, and read the value again
        ReDim resBinary(0 To length - 1) As Byte
        retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
            length)
    End If
    
    ' return a value corresponding to the value type
    Select Case valueType
        Case REG_DWORD
            CopyMemory resLong, resBinary(0), 4
            GetRegistryValue = resLong
        Case REG_SZ, REG_EXPAND_SZ
            ' copy everything but the trailing null char
            resString = Space$(length - 1)
            CopyMemory ByVal resString, resBinary(0), length - 1
            GetRegistryValue = resString
        Case REG_BINARY
            ' resize the result resBinary
            If length <> UBound(resBinary) + 1 Then
                ReDim Preserve resBinary(0 To length - 1) As Byte
            End If
            GetRegistryValue = resBinary()
        Case REG_MULTI_SZ
            ' copy everything but the 2 trailing null chars
            resString = Space$(length - 2)
            CopyMemory ByVal resString, resBinary(0), length - 2
            GetRegistryValue = resString
        Case Else
            RegCloseKey handle
            Err.Raise 1001, , "Unsupported value type"
    End Select
    
    ' close the registry key
    RegCloseKey handle
End Function


Function GetListItemHeightInPixels(ctrl As Control) As Long
' Return the height of each entry in a ListBox or ComboBox control (in pixels)
    Dim uMsg As Long
    If TypeOf ctrl Is ListBox Then
        uMsg = LB_GETITEMHEIGHT
    ElseIf TypeOf ctrl Is ComboBox Then
        uMsg = CB_GETITEMHEIGHT
    Else
        GetListItemHeightInPixels = -1
        Exit Function
    End If
    GetListItemHeightInPixels = SendMessage(ctrl.hWnd, uMsg, 0, ByVal 0&)
End Function
Function GetListItemHeightInTwips(ctrl As Control) As Long

    GetListItemHeightInTwips = ConvertPixelsToTwipsY(GetListItemHeightInPixels(ctrl))

End Function

Sub DisplayModemPanel()

    'Display Modem Settings
    Call Shell("rundll32.exe shell32.dll,Control_RunDLL modem.cpl", vbNormalFocus)
    
End Sub

''Display Printers
'Call Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @2", vbNormalFocus)
'
''Display Fonts
'Call Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @3", vbNormalFocus)
'
''Display Multimedia Settings
'Call Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl", vbNormalFocus)

Public Sub OpenWebPage(ByVal sURL As String)
On Error GoTo ErrorTrap

   Call ShellExecute(0&, vbNullString, sURL, vbNullString, _
                     vbNullString, vbNormalFocus)

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Public Function SetTBTabStops(TB As Object, _
ParamArray TabStops()) As Boolean

'PURPOSE: Set TabStops for a text box,
'a rich text box or any UserControl
'based on a text box
'that exposes the underlying text box's
'hwnd property.

'This creates columns whereby items
'in each column are separated by
'a tab character


'USAGE:
'Pass TextBox Object and a comma delimited
'list of tab stops.  Tab stops are expressed
'in dialog units which approximately equal
'1/4 the width of a character

'EXAMPLE:
'SetTBTabStops text1, 40, 80, 120
'text1.text = "Column1" & vbTab & "Column2" _
'& vbTab & "Column3" & vbTab & "Column4"

'This will create 3 columns separated by
'about 10 characters

Dim alTabStops() As Long
Dim lCtr As Long
Dim lColumns As Long
Dim lRet As Long

On Error GoTo ErrorHandler:

ReDim alTabStops(UBound(TabStops)) As Long

For lCtr = 0 To UBound(TabStops)
    alTabStops(lCtr) = TabStops(lCtr)
Next

lColumns = UBound(alTabStops) + 1


lRet = SendMessage(TB.hWnd, EM_SETTABSTOPS, _
lColumns, alTabStops(0))

SetTBTabStops = (lRet = 0)
Exit Function

ErrorHandler:
SetTBTabStops = False

End Function

