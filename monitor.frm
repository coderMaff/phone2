VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form_Monitor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chips Phone Monitor for Windows v1.0.1"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "monitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame_call 
      Caption         =   "Call Log"
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8775
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   240
         Width           =   8535
      End
   End
   Begin VB.Frame Frame_Stat 
      Caption         =   "Statistics"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   8775
      Begin VB.TextBox text_stat 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   8535
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5760
      Top             =   2400
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5160
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Menu mpopupSys 
      Caption         =   "System"
      Visible         =   0   'False
      Begin VB.Menu mpopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mpopexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form_Monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 28.06.2000 MRB Created
' 30.06.2000 MJW Iconify to tray
' 03.07.2000 MRB Saves stats after every phonecall, just in case

Dim sGlobalBuffer As String
Dim iTimeIn(10) As Integer
Dim iTimeOut(10) As Integer
Dim iCallsIn(10) As Integer
Dim iCallsOut(10) As Integer
Dim sLastNum(10) As String

' Functions

Function pad(i As Integer, iLen As Integer, sPad As String) As String
    Dim sSpace As String
    Dim X As Integer
    
    If Len(Format(i)) > iLen Then
        pad = Left(Format(i), iLen)
    Else
        Do While X < iLen
            X = X + 1
            sSpace = sSpace + sPad
        Loop
        pad = Left(sSpace, iLen - Len(Format(i))) + Format(i)
    End If
End Function

' Public procedure

Public Sub Update_Stats()
    Dim X As Integer
    X = 0
    text_stat = "E - #  In - # Out - # Total -  Time In - Time Out -    Total - Last Number" + Chr(13) + Chr(10)
    
    Do While X < 10
        text_stat = text_stat + Format(Val(X)) + " - "
        text_stat = text_stat + pad(iCallsIn(X), 5, " ") + " - "
        text_stat = text_stat + pad(iCallsOut(X), 5, " ") + " - "
        text_stat = text_stat + pad(iCallsIn(X) + iCallsOut(X), 7, " ") + " - "
        text_stat = text_stat + Str(TimeSerial(0, 0, iTimeIn(X))) + " - "
        text_stat = text_stat + Str(TimeSerial(0, 0, iTimeOut(X))) + " - "
        text_stat = text_stat + Str(TimeSerial(0, 0, iTimeIn(X) + iTimeOut(X))) + " - "
        text_stat = text_stat + sLastNum(X)
        text_stat = text_stat + Chr(13) + Chr(10)
        X = X + 1
    Loop
    
' Save stats every time, encase it crashes
    X = 0
    Open "c:\phone\phone.sta" For Output As #1
    Do While X < 10
        Write #1, iTimeIn(X), iTimeOut(X), iCallsIn(X), iCallsOut(X), sLastNum(X)
        X = X + 1
    Loop
    Close #1
End Sub

Public Sub update_log(sText As String)
    Text1.Text = Text1.Text + sText
    Open "c:\phone\Phone.Log" For Append As #1
        Write #1, Left(sText, 75)
    Close #1
End Sub

Private Sub Form_Load()
'the form must be fully visible before calling Shell_NotifyIcon       Me.Show
    Me.Refresh
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Phone Monitor" & vbNullChar
        End With
       Shell_NotifyIcon NIM_ADD, nid
    
    Dim X As Integer
    ' Open the serial port
    On Error GoTo CheckError
    
    If MSComm1.PortOpen = False Then
        MSComm1.CommPort = 2
        MSComm1.Settings = "9600,N,7,1"
        MSComm1.PortOpen = True
        MSComm1.InputLen = 0
        Text1.Text = "Port Open" + Chr(13) + Chr(10)
        Open "c:\phone\phone.sta" For Input As #1
            Do While X < 10
                Input #1, iTimeIn(X), iTimeOut(X), iCallsIn(X), iCallsOut(X), sLastNum(X)
                X = X + 1
            Loop
        Close #1
        Call Update_Stats
    End If
    
    Exit Sub
CheckError:
    If Err.Number <> 0 Then MsgBox (Error(Err.Number))
    ' Skip error when file dont exist, infact all errors but what the ell
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this procedure receives the callbacks from the System Tray icon.
      Dim Result As Long
      Dim msg As Long
       'the value of X will vary depending upon the scalemode setting
       If Me.ScaleMode = vbPixels Then
       msg = X
       Else
        msg = X / Screen.TwipsPerPixelX
        End If
        Select Case msg
        Case WM_LBUTTONUP        '514 restore form window
         If Me.WindowState = vbNormal Then
            Me.WindowState = vbMinimized
            Me.Hide
         Else
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
         End If
         Case WM_LBUTTONDBLCLK    '515 restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.Show
         Case WM_RBUTTONUP        '517 display popup menu
         Result = SetForegroundWindow(Me.hwnd)
         Me.PopupMenu Me.mpopupSys
       End Select
End Sub

Private Sub Form_Resize()
'this is necessary to assure that the minimized window is hidden
       If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'this removes the icon from the system tray
       Shell_NotifyIcon NIM_DELETE, nid
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
        Text1.Text = "Port Closed" + Chr(13) + Chr(10)
        Open "c:\phone\phone.sta" For Output As #1
            Do While X < 10
                Write #1, iTimeIn(X), iTimeOut(X), iCallsIn(X), iCallsOut(X), sLastNum(X)
                X = X + 1
            Loop
        Close #1
    End If
    
End Sub

Private Sub mpopexit_Click()
'called when user clicks the popup menu Exit command
Unload Me
End Sub

Private Sub mpopRestore_Click()
'called when the user clicks the popup menu Restore command
Dim Result As Long
Me.WindowState = vbNormal
Result = SetForegroundWindow(Me.hwnd)
       Me.Show
End Sub

Private Sub MSComm1_OnComm()
    Do While MSComm1.InBufferCount > 0
        sGlobalBuffer = sGlobalBuffer + MSComm1.Input
    Loop
End Sub

' String of data from phone thingy is 80 characters long including returns
'  1 Day of the week
'  4 SPACE
'  5 Day of the month
'  7 SPACE
'  8 Month
' 11 SPACE
' 12 Year
' 14 SPACE
' 15 Time
' 23 SPACE
' 24 Line
' 26 SPACE
' 27 Number dialed
' 28 SPACE
' 52 Duration hours
' 54 SPACE
' 55 Duration Minutes
' 57 SPACE
' 58 Duration Seconds
' 60 SPACE
' 72 Extension
' 75 SPACE

Private Sub Timer1_Timer()
    
    Dim sText As String
    Dim iLoop As Integer
    Dim fValid As Integer
    Dim iDur As Integer
    Dim iExt As Integer
        
    iLoop = 1
    fValid = 0
        
    If (sGlobalBuffer <> "") And Len(sGlobalBuffer) > 79 Then
                
        Do While iLoop < (Len(sGlobalBuffer) - 3)
        
            Select Case Mid(sGlobalBuffer, iLoop, 3)
                Case "MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"
                    If (Len(Mid(sGlobalBuffer, iLoop)) > 79) Then fValid = 1
            End Select
            
            If fValid = 1 Then
                sText = Mid(sGlobalBuffer, iLoop, 80)                           ' Grab my important text
                iDur = (Val(Mid(sText, 52, 2)) * 3600) + (Val(Mid(sText, 55, 2)) * 60) + Val(Mid(sText, 58, 2)) 'Nab the duration
                iExt = Val(Mid(sText, 74, 1))                                   ' Which extension dialed it
                sLastNum(iExt) = Mid(sText, 27, 11)                             ' Update last number dialed
                
                Select Case sLastNum(iExt)
                    Case "INCOMING   "
                        iTimeIn(iExt) = iTimeIn(iExt) + iDur                    ' Incrament calls in time
                        iCallsIn(iExt) = iCallsIn(iExt) + 1                     ' Incrament #calls in
                        Call update_log(sText)
                    Case "UNANSWERED "
                        Call update_log(Mid(sText, 1, 75) + Chr(13) + Chr(10))  ' Bung it on the screen
                    Case Else
                        iTimeOut(iExt) = iTimeOut(iExt) + iDur                  ' Incrament calls out time
                        iCallsOut(iExt) = iCallsOut(iExt) + 1                   ' Incrament #calls out
                        Call update_log(sText)                                  ' Bung it on the screen
                End Select
                sGlobalBuffer = Mid(sGlobalBuffer, iLoop + 80)                  ' Empty the delt-with bit from the global buffer
                iLoop = Len(sGlobalBuffer)                                      ' Jump loop
                Call Update_Stats
            Else
                iLoop = iLoop + 1
            End If
        Loop
    End If
                            
End Sub
