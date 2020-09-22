VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCheckWebSite 
   Caption         =   "Check Web Site"
   ClientHeight    =   3045
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "CheckWebSite.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "CheckWebSite.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblURL"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Image1(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Image1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Image1(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtMinutes"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtURL"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdViewSource"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtFailMinutes"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdCheckNow"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkMinimizeToTray"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Notify"
      TabPicture(1)   =   "CheckWebSite.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(3)=   "chkMsgBox"
      Tab(1).Control(4)=   "chkNetSend"
      Tab(1).Control(5)=   "chkEmail"
      Tab(1).Control(6)=   "txtNetSend"
      Tab(1).Control(7)=   "txtEmail"
      Tab(1).Control(8)=   "txtMailServer"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Message"
      TabPicture(2)   =   "CheckWebSite.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(1)=   "Label6"
      Tab(2).Control(2)=   "txtSubject"
      Tab(2).Control(3)=   "txtBody"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Help"
      TabPicture(3)   =   "CheckWebSite.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtHelp"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.TextBox txtMailServer 
         Height          =   285
         Left            =   -71940
         TabIndex        =   12
         Text            =   "mail.domain.com"
         ToolTipText     =   "SMTP mail server address"
         Top             =   1740
         Width           =   1455
      End
      Begin VB.TextBox txtHelp 
         BackColor       =   &H00C0C0C0&
         Height          =   2295
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "CheckWebSite.frx":04B2
         Top             =   480
         Width           =   4335
      End
      Begin VB.CheckBox chkMinimizeToTray 
         Caption         =   "Minimize to System Tray"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Minimize as icon in bottom right tray, not taskbar"
         Top             =   1980
         Width           =   2535
      End
      Begin VB.CommandButton cmdCheckNow 
         Caption         =   "Check Now"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Check if web page is up right now"
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtFailMinutes 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Text            =   "30"
         ToolTipText     =   "Number of minutes to wait"
         Top             =   1560
         Width           =   555
      End
      Begin VB.TextBox txtEmail 
         Height          =   525
         Left            =   -73680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "CheckWebSite.frx":0BFD
         ToolTipText     =   "seperate addresses with commas and no spaces"
         Top             =   2040
         Width           =   3195
      End
      Begin VB.TextBox txtNetSend 
         Height          =   525
         Left            =   -73680
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "apearso,jvannor"
         ToolTipText     =   "seperate ids with commas and no spaces"
         Top             =   1200
         Width           =   3195
      End
      Begin VB.TextBox txtBody 
         Height          =   1575
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "CheckWebSite.frx":0C24
         ToolTipText     =   "Message text to send when down"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.CommandButton cmdViewSource 
         Caption         =   "View Web Source HTML"
         Height          =   495
         Left            =   2280
         TabIndex        =   5
         ToolTipText     =   "View web page HTML"
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtSubject 
         Height          =   285
         Left            =   -73800
         TabIndex        =   14
         Text            =   "<Web Site> is down!"
         ToolTipText     =   "Subject of message to send when down"
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtURL 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Text            =   "http://fakesite.com/index.htm"
         ToolTipText     =   "Web page URL"
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtMinutes 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Text            =   "1"
         ToolTipText     =   "Number of minutes to wait"
         Top             =   1080
         Width           =   555
      End
      Begin VB.CheckBox chkEmail 
         Caption         =   "Send Email Message"
         Height          =   255
         Left            =   -74880
         TabIndex        =   11
         ToolTipText     =   "Send internet email message"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CheckBox chkNetSend 
         Caption         =   "Send Net Message (Net Send)"
         Height          =   255
         Left            =   -74880
         TabIndex        =   8
         ToolTipText     =   "Send message using DOS 'net send' command"
         Top             =   960
         Width           =   2655
      End
      Begin VB.CheckBox chkMsgBox 
         Caption         =   "Display Message Box on this machine"
         Height          =   255
         Left            =   -74880
         TabIndex        =   7
         ToolTipText     =   "Display local, modal message box"
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label9 
         Caption         =   "Mail Server"
         Height          =   255
         Left            =   -72840
         TabIndex        =   25
         Top             =   1800
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   600
         Picture         =   "CheckWebSite.frx":0C51
         Top             =   2280
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   120
         Top             =   2280
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   4080
         Picture         =   "CheckWebSite.frx":0F5B
         Top             =   2280
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label8 
         Caption         =   "On failure, wait"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1635
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "minute(s), then check again"
         Height          =   195
         Left            =   2160
         TabIndex        =   22
         Top             =   1635
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Message Body"
         Height          =   255
         Left            =   -74760
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Subject"
         Height          =   255
         Left            =   -74760
         TabIndex        =   20
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblURL 
         Caption         =   "Web Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "minute(s)"
         Height          =   195
         Left            =   2160
         TabIndex        =   18
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Check every"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1155
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "email addresses"
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "user logon ids"
         Height          =   255
         Left            =   -74880
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   120
      Top             =   2640
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4080
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   0
      Top             =   2160
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuCheckNow 
         Caption         =   "Check Now"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmCheckWebSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CheckTime As Date

'API to set order/positon of window
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

'API to minimize window to tray icon
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private SysIcon As NOTIFYICONDATA, RunningInTray As Boolean

'API to flash form icon
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Private Sub OnTopYes(WinHandle As Long)
    Call SetWindowPos(WinHandle, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Private Sub OnTopNo(WinHandle As Long)
    Call SetWindowPos(WinHandle, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Public Sub ShowIcon(ByRef TrayForm As Form)
    ' Show the systray icon. Use from another form : "SystrayIcon.ShowIcon"
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = TrayForm.hwnd
    SysIcon.uID = vbNull
    SysIcon.uFlags = 7
    SysIcon.uCallbackMessage = 512
    SysIcon.hIcon = TrayForm.Icon
    SysIcon.szTip = TrayForm.Caption + Chr(0)
    Shell_NotifyIcon 0, SysIcon
    RunningInTray = True
End Sub

Public Sub RemoveIcon(TrayForm As Form)
    ' Remove the systray icon. Use from another form : "SystrayIcon.RemoveIcon"
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = TrayForm.hwnd
    SysIcon.uID = vbNull
    SysIcon.uFlags = 7
    SysIcon.uCallbackMessage = vbNull
    SysIcon.hIcon = TrayForm.Icon
    SysIcon.szTip = Chr(0)
    Shell_NotifyIcon 2, SysIcon
    RunningInTray = False
End Sub

Private Sub Notify()
Dim subject As String
Dim body As String
Dim start As Integer
    subject = txtSubject
    subject = Replace(subject, "<Web Site>", Replace(txtURL, "http://", ""))
    subject = Replace(subject, "<date>", Now)
    
    body = txtBody
    body = Replace(body, "<Web Site>", Replace(txtURL, "http://", ""))
    body = Replace(body, "<date>", Now)
    
    If chkNetSend = 1 And txtNetSend <> "" Then
        Dim NetSend
        NetSend = txtNetSend.Text
        
        On Error Resume Next
        While InStr(NetSend, ",") > 0
            Shell "net send " & Left(NetSend, InStr(NetSend, ",") - 1) & " " & subject & vbCrLf & body, vbMinimizedNoFocus
            NetSend = Mid(NetSend, InStr(NetSend, ",") + 1)
        Wend
        Shell "net send " & NetSend & " " & subject & vbCrLf & body, vbMinimizedNoFocus
        On Error GoTo 0
    End If
    
    If chkEmail = 1 And txtEmail <> "" And txtMailServer <> "" Then
        Dim eml As New clsSendEmail
        Call eml.SendEmail(Me.Winsock1, txtMailServer.Text, "andy_pearson@uhc.com", txtEmail.Text, subject, body)
    End If

    If chkMsgBox = 1 Then
        On Error Resume Next
        Call OnTopYes(Me.hwnd)
        Beep
        Beep
        MsgBox subject & vbCrLf & vbCrLf & body
        Call OnTopNo(Me.hwnd)
        On Error GoTo 0
    End If
End Sub

Private Sub cmdCheckNow_Click()
    CheckTime = DateAdd("n", -1 * (txtMinutes + 1), Now)
    Call Timer1_Timer
End Sub

Private Sub cmdViewSource_Click()
Dim strPage As String
    On Error Resume Next
    strPage = Inet1.OpenURL(txtURL.Text)
    On Error GoTo 0
    MsgBox strPage
End Sub

Private Sub Form_Load()
    Left = GetSetting(App.EXEName, "Window", "X", 0)
    Top = GetSetting(App.EXEName, "Window", "Y", 0)
    If Left < 0 Then Left = 0
    If Top < 0 Then Top = 0
    txtURL = GetSetting(App.EXEName, "Settings", "WebSite", txtURL.Text)
    txtMinutes = GetSetting(App.EXEName, "Settings", "CheckMinutes", txtMinutes.Text)
    txtFailMinutes = GetSetting(App.EXEName, "Settings", "FailMinutes", txtFailMinutes.Text)
    txtNetSend = GetSetting(App.EXEName, "Notify", "NetSend", txtNetSend.Text)
    txtEmail = GetSetting(App.EXEName, "Notify", "Email", txtEmail.Text)
    chkMsgBox = GetSetting(App.EXEName, "Settings", "DisplayMessage", 0)
    chkNetSend = GetSetting(App.EXEName, "Settings", "NetSend", 0)
    chkEmail = GetSetting(App.EXEName, "Settings", "SendEmail", 0)
    txtSubject = GetSetting(App.EXEName, "Settings", "Subject", txtSubject.Text)
    txtBody = GetSetting(App.EXEName, "Settings", "Body", txtBody.Text)
    chkMinimizeToTray = GetSetting(App.EXEName, "Settings", "MinimizeToTray", 0)
    txtMailServer = GetSetting(App.EXEName, "Settings", "MailServer")
    If App.PrevInstance Then
        Unload Me
        Exit Sub
    End If
    
    DoEvents
    Me.Image1(0).Picture = Me.Icon
    CheckTime = DateAdd("n", -1 * (txtMinutes + 1), Now)
    If txtURL <> "http://fakesite.com/index.htm" Then 'avoid error for first use
        Call Timer1_Timer
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Left > 0 Then SaveSetting App.EXEName, "Window", "X", Left
    If Top > 0 Then SaveSetting App.EXEName, "Window", "Y", Top
    SaveSetting App.EXEName, "Settings", "WebSite", txtURL.Text
    SaveSetting App.EXEName, "Settings", "CheckMinutes", txtMinutes.Text
    SaveSetting App.EXEName, "Settings", "FailMinutes", txtFailMinutes.Text
    SaveSetting App.EXEName, "Notify", "NetSend", txtNetSend.Text
    SaveSetting App.EXEName, "Notify", "Email", txtEmail.Text
    SaveSetting App.EXEName, "Settings", "DisplayMessage", chkMsgBox.Value
    SaveSetting App.EXEName, "Settings", "NetSend", chkNetSend.Value
    SaveSetting App.EXEName, "Settings", "SendEmail", chkEmail.Value
    SaveSetting App.EXEName, "Settings", "Subject", txtSubject.Text
    SaveSetting App.EXEName, "Settings", "Body", txtBody.Text
    SaveSetting App.EXEName, "Settings", "MinimizeToTray", chkMinimizeToTray.Value
    If txtMailServer <> "" Then SaveSetting App.EXEName, "Settings", "MailServer", txtMailServer.Text
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        ' Put your code here to emulate systray icons events. Note: If there is any
        ' code for MouseMove, MouseDown, or MouseUp then a Double Click won't be
        ' caught.
        ' Only uncomment the events that your app will use, so as to avoid any
        ' strange errors.
      If RunningInTray Then
        Select Case x
            'Case 7680   ' MouseMove
            'Case 7695   ' Left MouseDown
            'Case 7710   ' Left MouseUp
            Case 7725   ' Left DoubleClick
                Me.WindowState = vbNormal   ' Or vbMaximized if you feel like it.
                Me.Show
                RemoveIcon Me
            Case 7740   ' Right MouseDown
                Me.PopupMenu Me.mnuFile
            'Case 7755   ' Right MouseUp
            'Case 7770   ' Right DoubleClick
        End Select
      End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Remove the icon when this form unload. Don't forget to unload this form!
    RemoveIcon Me 'Add your form's name here for the sub to work.
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized And chkMinimizeToTray = 1 Then
        ' This code hides the Form and puts the icon in the tray. Feel free to move
        ' it around if you like.
        Me.Hide
        ShowIcon Me
    End If
End Sub

Private Sub mnuCheckNow_Click()
    cmdCheckNow_Click
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuRestore_Click()
    Me.WindowState = vbNormal
    Me.Show
    RemoveIcon Me
End Sub

Private Sub Timer1_Timer()
Dim strPage As String
    If DateAdd("n", txtMinutes, CheckTime) < Now Then
        'check web page
        CheckTime = Now
        
        Inet1.RequestTimeout = 30
        On Error Resume Next
        strPage = Inet1.OpenURL(txtURL.Text)
        On Error GoTo 0
        If InStr(LCase(strPage), "</head>") <= 0 Or InStr(LCase(strPage), "404 not found") > 0 Or InStr(LCase(strPage), "dns") > 0 Then
            Notify
            CheckTime = DateAdd("n", txtFailMinutes, Now)
            Caption = Replace(txtURL, "http://", "") & " - Down - " & Format(Now, "hh:nn am/pm")
            Me.Icon = Me.Image1(1).Picture
            Me.Timer2.Enabled = True
        Else
            Caption = Replace(txtURL, "http://", "") & " - Up - " & Format(Now, "hh:nn am/pm")
            Me.Icon = Me.Image1(0).Picture
            If Me.Timer2.Enabled Then
                Me.Timer2.Enabled = False
                RemoveIcon Me
                ShowIcon Me
            End If
        End If
    End If
End Sub

Private Sub Timer2_Timer()
Static changeicon As Boolean
    FlashWindow hwnd, 1
    If changeicon Then
        Me.Icon = Me.Image1(1).Picture
    Else
        Me.Icon = Me.Image1(2).Picture
    End If
    changeicon = Not changeicon
    If Me.WindowState <> vbNormal Then
        RemoveIcon Me
        ShowIcon Me
    End If
End Sub

