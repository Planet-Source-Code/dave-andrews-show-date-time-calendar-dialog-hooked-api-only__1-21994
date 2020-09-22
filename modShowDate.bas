Attribute VB_Name = "modShowDate"
'This Code was written by Dave Andrews
'Feel free to use or modify this module freely
'Special thanks to Joseph Huntley for the skeleton of API forms.

Option Explicit
Private Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Private Declare Function defWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type


Private Type POINTAPI
    x As Long
    y As Long
End Type


Private Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

'------Scrollbox Constants
Const SBS_VERT = &H1&
Const SBM_SETRANGE = &HE2
Const SBM_SETPOS = &HE0
Const SBM_GETPOS = &HE1
Const SBM_ENABLE_ARROWS = &HE4
Const ESB_ENABLE_BOTH = &H0
'-----------Edit Box Constants
Const ES_MULTILINE = &H4&
Const ES_CENTER = &H1&
Const ES_READONLY = &H800&
Const ES_NUMBER = &H2000
Const STM_SETIMAGE = 172
Const EM_GETLINE = &HC4
Const EN_SETFOCUS = &H100
'------Button Constants
Const BS_USERBUTTON = &H8&
Const BS_CENTER = 768
Const BS_PUSHBUTTON = &H0&
Const BS_AUTORADIOBUTTON = &H9&
Const BS_PUSHLIKE = &H1000&
Const BS_LEFTTEXT = &H20&
Const BM_SETSTATE = &HF3
Const BM_GETSTATE = &HF2
Const BM_SETCHECK = &HF1
Const BM_GETCHECK = &HF0
'----Static Constants---------
Const SS_WHITERECT = &H6&
Const SS_WHITEFRAME = &H9&
Const SS_CENTER = &H1&
'-----------Window Style Constants
Const WS_BORDER = &H800000
Const WS_CHILD = &H40000000
Const WS_OVERLAPPED = &H0&
Const WS_CAPTION = &HC00000 ' WS_BORDER Or WS_DLGFRAME
Const WS_SYSMENU = &H80000
Const WS_THICKFRAME = &H40000
Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000
Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Const WS_VISIBLE = &H10000000
Const WS_POPUP = &H80000000
Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Const WS_VSCROLL = &H200000
Const WS_EX_TOOLWINDOW = &H80
Const WS_EX_TOPMOST = &H8&
Const WS_EX_CLIENTEDGE = &H200&
Const WS_EX_WINDOWEDGE = &H100&
Public Const WS_EX_DLGMODALFRAME = &H1&
'-----ComboBox Constants-----
Const CBS_DROPDOWN = &H2&
Const CBS_DROPDOWNLIST = &H3&
Const CBS_SIMPLE = &H1&
Const CBS_AUTOHSCROLL = &H40&
Const CBS_DISABLENOSCROLL = &H800&
Const CBN_CLOSEUP = 8
Const CBN_SELCHANGE = 1
Const CBN_SELENDOK = 9
Const CB_ADDSTRING = &H143
Const CB_SHOWDROPDOWN = &H14F
Const CB_SETCURSEL = &H14E
Const CB_GETCURSEL = &H147
'-----------Window Messaging Constants
Const WM_DESTROY = &H2
Const WM_CLOSE = &H10
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_CTLCOLOREDIT = &H133
Const WM_COMMAND = &H111
Const WM_GETTEXT = &HD
Const WM_ENABLE = &HA
Const WM_KEYDOWN = &H100
Const WM_KEYUP = &H101
Const WM_SETTEXT = &HC
Const WM_VSCROLL = &H115
'--------Window Heiarchy Constants
Const GWL_WNDPROC = (-4)
Const GW_CHILD = 5
Const GW_OWNER = 4
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const SW_SHOWNORMAL = 1
'----------Misc Constants
Const CS_VREDRAW = &H1
Const CS_HREDRAW = &H2
Const CW_USEDEFAULT = &H80000000
Const COLOR_WINDOW = 5
Const SET_BACKGROUND_COLOR = 4103
Const IDC_ARROW = 32512&
Const IDI_APPLICATION = 32512&
Const MB_OK = &H0&
Const MB_ICONEXCLAMATION = &H30&

Dim MyMousePos As POINTAPI 'for getting the mouse positioning

Const gClassName = "Date Time API"

Dim gAppName As String
Dim CurMonth As Integer
Dim CurYear As Integer
Dim CurET As String
Dim CurPos As Integer
Dim TMax As Integer
Dim TMin As Integer
Dim CurDate As String
Dim CurTime As String
Dim Cancel As Boolean

Dim gHwnd As Long
Dim gMonthHwnd As Long
Dim gMonthOldProc As Long
Dim gYearHwnd As Long
Dim gYearOldProc As Long
Dim gDayLabel(6) As Long
Dim gDayHwnd(41) As Long

Dim gSpacerHwnd(2) As Long
Dim gHourHwnd As Long
Dim gHourOldProc As Long
Dim gMinHwnd As Long
Dim gMinOldProc As Long
Dim gSecHwnd As Long
Dim gSecOldProc As Long
Dim gAMPMHwnd As Long
Dim gAMPMOldProc As Long
Dim gUpHwnd As Long
Dim gUpOldProc As Long
Dim gDownHwnd As Long
Dim gDownOldProc As Long

Dim gOKHwnd As Long
Dim gOKOldProc As Long
Dim gCancelHwnd As Long
Dim gCancelOldProc As Long

Sub CreateCal()
Dim D As Integer
Dim sDay As Integer
Dim dCount As Integer
Dim tBuf As String
'Clear the text of all the buttons
SetDate
For D = 0 To 41
    Call SendMessage(gDayHwnd(D), WM_SETTEXT, 0&, ByVal CStr(""))
    Call SendMessage(gDayHwnd(D), BM_SETCHECK, False, 0&)
    Call EnableWindow(gDayHwnd(D), ByVal False) ' disable all the button-windows
Next D
CurMonth = SendMessage(gMonthHwnd&, CB_GETCURSEL, 0&, 0&) + 1
CurYear = Year(Now) + SendMessage(gYearHwnd&, CB_GETCURSEL, 0&, 0&) - 100
'the number of days in the current month = current month/year + 1 month - 1 day
dCount = Day(DateAdd("d", -1, DateAdd("m", 1, CDate(CurMonth & "/1/" & CurYear))))
sDay = Format(CDate(CurMonth & "/1/" & CurYear), "w") - 1
For D = sDay To sDay + dCount - 1
    Call SendMessage(gDayHwnd(D), WM_SETTEXT, 0&, ByVal CStr(D - sDay + 1))
    Call EnableWindow(gDayHwnd(D), ByVal True) ' enable the windows that have day-dates
    If (D - sDay + 1) = Day(CurDate) Then Call SendMessage(gDayHwnd(D), BM_SETCHECK, True, 0&)
Next D
End Sub

Sub Main()
    MsgBox ShowDate(Now(), "Select A Date")
    End
End Sub

Sub SetDate()
Dim D As Integer
Dim sDay As Integer
Dim dCount As Integer
Dim tBuf As String
Dim tVar(3) As String
dCount = Day(DateAdd("d", -1, DateAdd("m", 1, CDate(CurMonth & "/1/" & CurYear))))
sDay = Format(CDate(CurMonth & "/1/" & CurYear), "w") - 1
For D = sDay To sDay + dCount - 1
    If SendMessage(gDayHwnd(D), BM_GETCHECK, 0&, 0&) = 1 Then
        CurDate = CDate(CurMonth & "/" & (D - sDay + 1) & "/" & CurYear)
        Exit For
    End If
Next D
CurDate = Format(CurDate, "mm/dd/yyyy") & " " & CurTime
End Sub
Function EditClass() As WNDCLASS
EditClass.hbrBackground = vbRed
End Function


 Function ShowDate(Optional SelDate As String, Optional Title As String) As String
    If SelDate = "" Then SelDate = Now()
    If Title <> "" Then gAppName$ = Title Else gAppName$ = "Date/Time"
    Call GetCursorPos(MyMousePos)
    CurDate = SelDate
    CurMonth = Month(SelDate)
    CurYear = Year(SelDate)
    CurTime = Format(SelDate, "Long Time")
    Cancel = True
    Dim wMsg As Msg
    Dim tSec As String
    ''Call procedure to register window classname. If false, then exit.
    If RegisterWindowClass = False Then Exit Function
    
      ''Create window
      If CreateWindows() Then
         ''Loop will exit when WM_QUIT is sent to the window.
         Do While GetMessage(wMsg, 0&, 0&, 0&)
            ''TranslateMessage takes keyboard messages and converts
            ''them to WM_CHAR for easier processing.
            Call TranslateMessage(wMsg)
            ''Dispatchmessage calls the default window procedure
            ''to process the window message. (WndProc)
            Call DispatchMessage(wMsg)
            DoEvents
         Loop
      End If
    
    Call UnregisterClass(gClassName$, App.hInstance)
    If Not Cancel Then ShowDate = CurDate Else ShowDate = ""
End Function

 Function RegisterWindowClass() As Boolean

    Dim wc As WNDCLASS
    
    ''Registers our new window with windows so we can use our classname.
    
    wc.style = CS_HREDRAW Or CS_VREDRAW
    wc.lpfnwndproc = GetAddress(AddressOf WndProc) ''Address in memory of default window procedure.
    wc.hInstance = App.hInstance
    wc.hIcon = LoadIcon(0&, IDI_APPLICATION) ''Default application icon
    wc.hCursor = LoadCursor(0&, IDC_ARROW) ''Default arrow
    wc.hbrBackground = COLOR_WINDOW ''Default a color for window.
    wc.lpszClassName = gClassName$

    RegisterWindowClass = RegisterClass(wc) <> 0
    
End Function
 Function CreateWindows() As Boolean
    Dim ComboStyle As Long
    Dim LabelStyle As Long
    Dim TextStyle As Long
    Dim ButtonStyle As Long
    Dim ScrollStyle As Long
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim tDate As String
    Dim TempDC As Long
    ComboStyle = WS_CHILD Or CBS_DROPDOWN Or WS_VISIBLE Or WS_VSCROLL
    LabelStyle = WS_CHILD Or WS_VISIBLE Or SS_CENTER Or WS_BORDER
    TextStyle = WS_CHILD Or WS_VISIBLE Or ES_CENTER Or WS_BORDER Or ES_NUMBER
    ButtonStyle = WS_CHILD Or WS_VISIBLE Or BS_AUTORADIOBUTTON Or BS_PUSHLIKE Or WS_BORDER
    ScrollStyle = WS_CHILD Or WS_VISIBLE Or WS_BORDER
    'Create form window.
    gHwnd& = CreateWindowEx(WS_EX_TOOLWINDOW Or WS_EX_TOPMOST, gClassName$, gAppName$, WS_POPUPWINDOW Or WS_CAPTION Or WS_VISIBLE, MyMousePos.x, MyMousePos.y, 175, 262, 0&, 0&, App.hInstance, ByVal 0&)
    'Create Combos for month and year
    gMonthHwnd& = CreateWindowEx(0&, "COMBOBOX", "", ComboStyle, 0, 0, 96, 160, gHwnd&, 0&, App.hInstance, 0&)
    gYearHwnd& = CreateWindowEx(0&, "COMBOBOX", "", ComboStyle, 96, 0, 72, 160, gHwnd&, 0&, App.hInstance, 0&)
    'Create Labels for the days of the week
    gDayLabel(0) = CreateWindowEx(0&, "STATIC", "S", LabelStyle, 0, 25, 25, 19, gHwnd&, 0&, App.hInstance, 0&)
    gDayLabel(1) = CreateWindowEx(0&, "STATIC", "M", LabelStyle, 24, 25, 25, 19, gHwnd&, 0&, App.hInstance, 0&)
    gDayLabel(2) = CreateWindowEx(0&, "STATIC", "T", LabelStyle, 48, 25, 25, 19, gHwnd&, 0&, App.hInstance, 0&)
    gDayLabel(3) = CreateWindowEx(0&, "STATIC", "W", LabelStyle, 72, 25, 25, 19, gHwnd&, 0&, App.hInstance, 0&)
    gDayLabel(4) = CreateWindowEx(0&, "STATIC", "T", LabelStyle, 96, 25, 25, 19, gHwnd&, 0&, App.hInstance, 0&)
    gDayLabel(5) = CreateWindowEx(0&, "STATIC", "F", LabelStyle, 120, 25, 25, 19, gHwnd&, 0&, App.hInstance, 0&)
    gDayLabel(6) = CreateWindowEx(0&, "STATIC", "S", LabelStyle, 144, 25, 25, 19, gHwnd&, 0&, App.hInstance, 0&)
    'Create Buttons for the day-dates
    For i = 0 To 5
        For j = 0 To 6
            gDayHwnd(k) = CreateWindowEx(0&, "BUTTON", "30", ButtonStyle, (j * 24), 46 + (i * 24), 25, 25, gHwnd&, 0&, App.hInstance, 0&)
            k = k + 1
        Next j
    Next i
    'Create EDIT boxes for the time
    tDate = Hour(CurTime)
    If tDate = 0 Then tDate = 12
    gHourHwnd = CreateWindowEx(0&, "EDIT", IIf(tDate <= 12, tDate, tDate - 12), TextStyle, 0, 195, 30, 20, gHwnd&, 0&, App.hInstance, 0&)
    gSpacerHwnd(0) = CreateWindowEx(0&, "STATIC", ":", LabelStyle, 30, 195, 10, 20, gHwnd&, 0&, App.hInstance, 0&)
    gMinHwnd = CreateWindowEx(0&, "EDIT", Format(CurTime, "nn"), TextStyle, 40, 195, 30, 20, gHwnd&, 0&, App.hInstance, 0&)
    gSpacerHwnd(1) = CreateWindowEx(0&, "STATIC", ":", LabelStyle, 70, 195, 10, 20, gHwnd&, 0&, App.hInstance, 0&)
    gSecHwnd = CreateWindowEx(0&, "EDIT", Format(CurTime, "ss"), TextStyle, 80, 195, 30, 20, gHwnd&, 0&, App.hInstance, 0&)
    gSpacerHwnd(2) = CreateWindowEx(0&, "STATIC", " ", LabelStyle, 110, 195, 10, 20, gHwnd&, 0&, App.hInstance, 0&)
    gAMPMHwnd = CreateWindowEx(0&, "EDIT", Format(CurTime, "AMPM"), LabelStyle, 120, 195, 30, 20, gHwnd&, 0&, App.hInstance, 0&)
    gUpHwnd = CreateWindowEx(0&, "BUTTON", "", ScrollStyle, 152, 193, 16, 13, gHwnd&, 0&, App.hInstance, 0&)
    gDownHwnd = CreateWindowEx(0&, "BUTTON", "", ScrollStyle, 152, 204, 16, 13, gHwnd&, 0&, App.hInstance, 0&)
    'Create OK and Cancel Buttons
    gOKHwnd = CreateWindowEx(0&, "BUTTON", "OK", ScrollStyle, 0, 220, 85, 20, gHwnd&, 0&, App.hInstance, 0&)
    gCancelHwnd = CreateWindowEx(0&, "BUTTON", "CANCEL", ScrollStyle, 85, 220, 85, 20, gHwnd&, 0&, App.hInstance, 0&)
    
    
    'Fill The MONTH combobox with month names
    For i = 1 To 12
        tDate = Format(CDate(i & "/1/2000"), "mmmm")
        Call SendMessage(gMonthHwnd, CB_ADDSTRING, 0&, ByVal tDate$)
    Next i
    'Select the proper month for current Date
    Call SendMessage(gMonthHwnd, CB_SETCURSEL, Month(CurDate) - 1, 0&)
    'Fill the YEAR combobox
    For i = Year(CurDate) - 100 To Year(CurDate) + 100
        Call SendMessage(gYearHwnd, CB_ADDSTRING, 0&, ByVal CStr(i))
    Next i
    'Select the proper Year for current SelDate
    Call SendMessage(gYearHwnd, CB_SETCURSEL, 100, 0&)
    
    'Get the memory address of the default window
    'These staements are used to address the procedures associated with the
    'events of the controls on our form
    
    '-------Hook Month Combo-----------
    gMonthOldProc& = GetWindowLong(gMonthHwnd&, GWL_WNDPROC)
    Call SetWindowLong(gMonthHwnd&, GWL_WNDPROC, GetAddress(AddressOf MonthWndProc))
    '-------Hook Year Combo-----------
    gYearOldProc& = GetWindowLong(gYearHwnd&, GWL_WNDPROC)
    Call SetWindowLong(gYearHwnd&, GWL_WNDPROC, GetAddress(AddressOf YearWndProc))
    '-------Hook Hour EDIT-----------
    gHourOldProc& = GetWindowLong(gHourHwnd&, GWL_WNDPROC)
    Call SetWindowLong(gHourHwnd&, GWL_WNDPROC, GetAddress(AddressOf HourWndProc))
    '-------Hook Minute EDIT-----------
    gMinOldProc& = GetWindowLong(gMinHwnd&, GWL_WNDPROC)
    Call SetWindowLong(gMinHwnd&, GWL_WNDPROC, GetAddress(AddressOf MinWndProc))
    '-------Hook Seconds EDIT-----------
    gSecOldProc& = GetWindowLong(gSecHwnd&, GWL_WNDPROC)
    Call SetWindowLong(gSecHwnd&, GWL_WNDPROC, GetAddress(AddressOf SecWndProc))
    '-------Hook AMPM EDIT-----------
    gAMPMOldProc& = GetWindowLong(gAMPMHwnd&, GWL_WNDPROC)
    Call SetWindowLong(gAMPMHwnd&, GWL_WNDPROC, GetAddress(AddressOf AMPMWndProc))
    '-------Hook UPDOWN's-----------
    gUpOldProc& = GetWindowLong(gUpHwnd&, GWL_WNDPROC)
    Call SetWindowLong(gUpHwnd&, GWL_WNDPROC, GetAddress(AddressOf UpWndProc))
    gDownOldProc& = GetWindowLong(gDownHwnd&, GWL_WNDPROC)
    Call SetWindowLong(gDownHwnd&, GWL_WNDPROC, GetAddress(AddressOf DownWndProc))
    '-------Hook OK CANCEL-----------
    gOKOldProc& = GetWindowLong(gOKHwnd&, GWL_WNDPROC)
    Call SetWindowLong(gOKHwnd&, GWL_WNDPROC, GetAddress(AddressOf OKWndProc))
    gCancelOldProc& = GetWindowLong(gCancelHwnd&, GWL_WNDPROC)
    Call SetWindowLong(gCancelHwnd&, GWL_WNDPROC, GetAddress(AddressOf CancelWndProc))
    
    
    
    
    'Set the initial calendar
    CreateCal
    
    'Exit our function
    CreateWindows = (gHwnd& <> 0)
    
End Function
Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ''This our default window procedure for the window. It will handle all
    ''of our incoming window messages and we will write code based on the
    ''window message what the program should do.
    Dim i As Integer
      Select Case uMsg&
         Case WM_DESTROY:
            ''Since DefWindowProc doesn't automatically call
            ''PostQuitMessage (WM_QUIT). We need to do it ourselves.
            ''You can use DestroyWindow to get rid of the window manually.
            SetDate
            Call PostQuitMessage(0&)
      End Select
    ''Let windows call the default window procedure since we're done.
    WndProc = defWindowProc(hwnd&, uMsg&, wParam&, lParam&)

End Function
 Function MonthWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg&
        Case WM_COMMAND
            CreateCal
            
    End Select
    
  MonthWndProc = CallWindowProc(gMonthOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function
 Function MinWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg&
        Case WM_LBUTTONDOWN, WM_KEYUP
            Dim tLine As String
            tLine = Space(2)
            CurET = "MIN"
            Call SendMessage(gMinHwnd&, EM_GETLINE, 0&, ByVal tLine)
            CurPos = CLng(tLine)
            CurTime = Format(Hour(CurTime) & ":" & CurPos & ":" & Second(CurTime), "Long Time")
            TMax = 59
            TMin = 0
    End Select
    
  MinWndProc = CallWindowProc(gMinOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function
 Function HourWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg&
        Case WM_LBUTTONDOWN, WM_KEYUP
            Dim tLine As String
            tLine = Space(2)
            CurET = "HOUR"
            Call SendMessage(gHourHwnd&, EM_GETLINE, 0&, ByVal tLine)
            CurPos = CLng(tLine)
            CurTime = Format(CurPos & ":" & Minute(CurTime) & ":" & Second(CurTime) & " " & Format(CurTime, "AMPM"), "Long Time")
            TMax = 12
            TMin = 1
    End Select
    
  HourWndProc = CallWindowProc(gHourOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function
Function OKWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg&
        Case WM_LBUTTONDOWN
            Cancel = False
            Call SendMessage(gHwnd, WM_CLOSE, 0&, 0&)
    End Select
    
  OKWndProc = CallWindowProc(gOKOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function
Function CancelWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg&
        Case WM_LBUTTONDOWN
            Cancel = True
            Call SendMessage(gHwnd, WM_CLOSE, 0&, 0&)
    End Select
    
  CancelWndProc = CallWindowProc(gCancelOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function
 Function SecWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg&
        Case WM_LBUTTONDOWN, WM_KEYUP
            Dim tLine As String
            tLine = Space(2)
            CurET = "SEC"
            Call SendMessage(gSecHwnd&, EM_GETLINE, 0&, ByVal tLine)
            CurPos = CLng(tLine)
            CurTime = Format(Hour(CurTime) & ":" & Minute(CurTime) & ":" & CurPos, "Long Time")
            TMax = 59
            TMin = 0
    End Select
    
  SecWndProc = CallWindowProc(gSecOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function
 Function UpWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg&
        Case WM_LBUTTONUP
            CurPos = IIf(CurPos < TMax, CurPos + 1, TMin)
            Dim tWord As String
            'Form1.Caption = CurPos & " - " & wParam
            Select Case CurET
                Case "HOUR"
                    tWord = CStr(CurPos)
                    Call SendMessage(gHourHwnd&, WM_SETTEXT, 0&, ByVal tWord$)
                    CurTime = Format(CurPos & ":" & Minute(CurTime) & ":" & Second(CurTime) & " " & Format(CurTime, "AMPM"), "Long Time")
                Case "MIN"
                    tWord = Format(CurPos, "00")
                    Call SendMessage(gMinHwnd&, WM_SETTEXT, 0&, ByVal tWord$)
                    CurTime = Format(Hour(CurTime) & ":" & CurPos & ":" & Second(CurTime), "Long Time")
                Case "SEC"
                    tWord = Format(CurPos, "00")
                    Call SendMessage(gSecHwnd&, WM_SETTEXT, 0&, ByVal tWord$)
                    CurTime = Format(Hour(CurTime) & ":" & Minute(CurTime) & ":" & CurPos, "Long Time")
                Case "AMPM"
                    tWord = IIf(CurPos = 0, "AM", "PM")
                    Call SendMessage(gAMPMHwnd&, WM_SETTEXT, 0&, ByVal tWord$)
                    CurTime = Format(IIf(Hour(CurTime) <= 12, Hour(CurTime), Hour(CurTime) - 12) & ":" & Minute(CurTime) & ":" & Second(CurTime) & " " & tWord, "Long Time")
            End Select
    End Select
    
  UpWndProc = CallWindowProc(gUpOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function
Function DownWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg&
        Case WM_LBUTTONUP
            CurPos = IIf(CurPos > TMin, CurPos - 1, TMax)
            Dim tWord As String
            Select Case CurET
                Case "HOUR"
                    tWord = CStr(CurPos)
                    Call SendMessage(gHourHwnd&, WM_SETTEXT, 0&, ByVal tWord$)
                    CurTime = Format(CurPos & ":" & Minute(CurTime) & ":" & Second(CurTime) & " " & Format(CurTime, "AMPM"), "Long Time")
                Case "MIN"
                    tWord = Format(CurPos, "00")
                    Call SendMessage(gMinHwnd&, WM_SETTEXT, 0&, ByVal tWord$)
                    CurTime = Format(Hour(CurTime) & ":" & CurPos & ":" & Second(CurTime), "Long Time")
                Case "SEC"
                    tWord = Format(CurPos, "00")
                    Call SendMessage(gSecHwnd&, WM_SETTEXT, 0&, ByVal tWord$)
                    CurTime = Format(Hour(CurTime) & ":" & Minute(CurTime) & ":" & CurPos, "Long Time")
                Case "AMPM"
                    tWord = IIf(CurPos = 0, "AM", "PM")
                    Call SendMessage(gAMPMHwnd&, WM_SETTEXT, 0&, ByVal tWord$)
                    CurTime = Format(IIf(Hour(CurTime) <= 12, Hour(CurTime), Hour(CurTime) - 12) & ":" & Minute(CurTime) & ":" & Second(CurTime) & " " & tWord, "Long Time")
            End Select
    End Select
    
  DownWndProc = CallWindowProc(gDownOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function

 Function AMPMWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tLine As String
    Select Case uMsg&
        Case WM_LBUTTONDOWN
            tLine = Space(2)
            CurET = "AMPM"
            Call SendMessage(gAMPMHwnd&, EM_GETLINE, 0&, ByVal tLine) ' get the value of the edit box
            CurPos = IIf(tLine = "AM", 0, 1)
            CurTime = Format(IIf(Hour(CurTime) <= 12, Hour(CurTime), Hour(CurTime) - 12) & ":" & Minute(CurTime) & ":" & Second(CurTime) & " " & tLine, "Long Time")
            TMax = 1
            TMin = 0
        Case WM_KEYUP
            tLine = Space(1)
            CurET = "AMPM"
            Call SendMessage(gAMPMHwnd&, EM_GETLINE, 0&, ByVal tLine)
            tLine = Trim(UCase(tLine))
            If tLine = "A" Then
                tLine = "AM"
                Call SendMessage(gAMPMHwnd&, WM_SETTEXT, 0&, ByVal tLine$)
            Else
                tLine = "PM"
                Call SendMessage(gAMPMHwnd&, WM_SETTEXT, 0&, ByVal tLine$)
            End If
            CurPos = IIf(tLine = "AM", 0, 1)
            CurTime = Format(IIf(Hour(CurTime) <= 12, Hour(CurTime), Hour(CurTime) - 12) & ":" & Minute(CurTime) & ":" & Second(CurTime) & " " & tLine, "Long Time")
            TMax = 1
            TMin = 0
    End Select
        Debug.Print CurTime
    
  AMPMWndProc = CallWindowProc(gAMPMOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function
Function YearWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg&
        Case WM_COMMAND
            CreateCal
            
    End Select
    
  ''Since in MyCreateWindow we made the default window proc
  ''this procedure, we have to call the old one using CallWindowProc
  YearWndProc = CallWindowProc(gYearOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function

 Function GetAddress(ByVal lngAddr As Long) As Long
    ''Used with AddressOf to return the address in memory of a procedure.

    GetAddress = lngAddr&
    
End Function


