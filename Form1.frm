VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3840
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      ParitySetting   =   2
      DataBits        =   7
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5520
      Width           =   6615
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   2880
      Top             =   6600
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2280
      Top             =   6600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1680
      Top             =   6600
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   3015
      Left            =   2880
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   5655
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   9975
      _cy             =   5318
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Y0_ON As String
Dim Y0_OFF As String
Dim Y1_ON As String
Dim Y1_OFF As String
Dim Y2_ON As String
Dim Y2_OFF As String
Dim Y3_ON As String
Dim Y3_OFF As String
Dim Y4_ON As String
Dim Y4_OFF As String
Dim Y5_ON As String
Dim Y5_OFF As String
Dim Y6_ON As String
Dim Y6_OFF As String
Dim Y7_ON As String
Dim Y7_OFF As String
Dim ASK_STATUS As String
Dim X0_ON As String
Dim X1_ON As String
Dim X2_ON As String
Dim X3_ON As String
Dim X4_ON As String
Dim X5_ON As String
Dim X6_ON As String
Dim X7_ON As String


Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
 
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" _
           (ByVal lpApplicationName As Long, _
            ByVal lpKeyName As Long, _
            ByVal lpDefault As Long, _
            ByVal lpReturnedString As Long, _
            ByVal nSize As Long, _
            ByVal lpFileName As Long) As Long
            
            
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Dim shell_tray As Long
          
  
Private Const SWP_NOSIZE = &H1

Private Const SWP_NOMOVE = &H2

Private Const HWND_TOPMOST = -1

Private Const HWND_NOTOPMOST = -2

Private Declare Function SetWindowPos Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal hWndInsertAfter As Long, _
             ByVal X As Long, ByVal Y As Long, _
             ByVal cx As Long, ByVal cy As Long, _
             ByVal wFlags As Long) As Long
  
  
  
  
  
  
' ################################
' [功能描述] 延时等待指定的毫秒数.
' [参数列表] Milliseconds  毫秒数.
' [返回类型] 无.
' ################################
Public Sub Delay(ByVal Milliseconds As Long)
    Dim lngTime As Long
     
    lngTime = timeGetTime
     
    While timeGetTime < lngTime + Milliseconds
        DoEvents
    Wend
End Sub
             
'------------
'读INI文件
'------------
Private Function GetValueFromINIFile(ByVal SectionName As String, _
        ByVal KeyName As String, _
        ByVal IniFileName As String) As String
     
    Dim strBuf As String
    '128个字符，初始化时用 0 填充
    strBuf = String(128, 0)
     
    GetPrivateProfileString StrPtr(SectionName), _
        StrPtr(KeyName), _
        StrPtr(""), _
        StrPtr(strBuf), _
        128, _
        StrPtr(IniFileName)
    '去除多余的 0
    strBuf = Replace(strBuf, Chr(0), "")
    GetValueFromINIFile = strBuf
End Function

'Private Sub Form_Unload(Cancel As Integer)
'Call SetWindowPos(shell_tray, 0, 0, 0, 0, 0, &O4 Or &H40)
'End Sub

Private Sub Form_Load()
'    shell_tray = FindWindow("Shell_TrayWnd", "")
'    Call SetWindowPos(shell_tray, 0, 0, 0, 0, 0, &O4 Or &H80)

    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

    Call Comm_initial
    Call Form_initial
    'Call WMP1_initial
    Call Protocol_initial
    Call YALL_OFF
    
    Timer1.Enabled = True
    Timer2.Enabled = True

    WindowsMediaPlayer1.URL = ""
    WMP1_play "\v\v0.mp4"
    WindowsMediaPlayer1.Controls.play
End Sub

Private Sub Comm_initial()
    Dim port As String
    port = GetValueFromINIFile("Port", "Value", App.Path & "\Port.INI")
    MSComm1.CommPort = port
    MSComm1.Settings = "9600,e,7,1"
    MSComm1.InputMode = comInputModeBinary
    MSComm1.InBufferSize = 1024                                                      ' 设置接收缓冲区为1024字节
    MSComm1.OutBufferSize = 4096                                                     ' 设置发送缓冲区为4096字节
    MSComm1.InBufferCount = 0                                                        ' 清空输入缓冲区
    MSComm1.OutBufferCount = 0                                                       ' 清空输出缓冲区
    MSComm1.SThreshold = 1                                                           ' 发送缓冲区空触发发送事件
    MSComm1.RThreshold = 1                                                           ' 每X个字符到接收缓冲区引起触发接收事件
    MSComm1.OutBufferCount = 0                                                       ' 清空发送缓冲区
    MSComm1.InBufferCount = 0
    MSComm1.PortOpen = True
End Sub

Private Sub Form_initial()
    Form1.Width = Screen.Width
    Form1.Height = Screen.Height
    Move 0, 0
    SetCursorPos Screen.Width, 0
End Sub

Private Sub WMP1_initial()
    WindowsMediaPlayer1.Move 0, 0, 0, 0
    WindowsMediaPlayer1.uiMode = "None"
    WindowsMediaPlayer1.enableContextMenu = False
    WindowsMediaPlayer1.Height = Screen.Height
    WindowsMediaPlayer1.Width = Screen.Width
End Sub

Private Sub Protocol_initial()
    
    Y0_ON = "023730303035034646"
    Y0_OFF = "023830303035033030"
    Y1_ON = "023730313035033030"
    Y1_OFF = "023830313035033031"
    Y2_ON = "023730323035033031"
    Y2_OFF = "023830333035033033"
    Y3_ON = "023730333035033032"
    Y3_OFF = "023830333035033033"
    Y4_ON = "023730343035033033"
    Y4_OFF = "023830343035033034"
    Y5_ON = "023730353035033034"
    Y5_OFF = "023830353035033035"
    Y6_ON = "023730363035033035"
    Y6_OFF = "023830363035033036"
    Y7_ON = "023730373035033036"
    Y7_OFF = "023830373035033037"
    ASK_STATUS = "0230303038303031033543"
    X0_ON = "023031033634"
    X1_ON = "023032033635"
    X2_ON = "023034033637"
    X3_ON = "023038033642"
    X4_ON = "023130033634"
    X5_ON = "023230033635"
    X6_ON = "023430033637"
    X7_ON = "023830033642"
        
End Sub

Private Sub MSComm1_OnComm()
On Error GoTo Err
    Select Case MSComm1.CommEvent
        Case comEvReceive
                Call hexrecv
        Case comEvSend
        Case Else
    End Select
Err:
End Sub

Private Sub WMP1_play(ByVal filename As String)
    
    WindowsMediaPlayer1.URL = App.Path & filename
    WindowsMediaPlayer1.Controls.play

End Sub

Private Sub WMP1_playstatus()
    If (WindowsMediaPlayer1.playState = wmppsPlaying) Then

    ElseIf WindowsMediaPlayer1.playState = wmppsStopped Then
        YALL_OFF
        'WindowsMediaPlayer1.Visible = False
        WMP1_play "\v\v0.mp4"
        
    Else
    End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        MSComm1.PortOpen = False
        End
    End If
End Sub


Private Sub hexsend(ByVal txtsend As String)
    Dim sd() As Byte
    Dim i As Integer
    If Len(txtsend) Mod 2 = 0 And Len(txtsend) <> 0 Then
        ReDim sd(Len(txtsend) / 2 - 1)
        For i = 0 To Len(txtsend) - 1 Step 2
           sd(i / 2) = Val("&H" & Mid(txtsend, i + 1, 2))
        Next
        If MSComm1.PortOpen = True Then
            MSComm1.Output = sd
        Else
            MSComm1.PortOpen = True
            MSComm1.Output = sd
        End If
    Else
        'MsgBox ("格式不对!")
    End If
End Sub

Private Sub hexrecv()
On Error GoTo Err
    Dim ReceiveArr() As Byte
    Dim receiveData As String
    Dim Counter As Integer
    Dim i As Integer
    If (MSComm1.InBufferCount > 0) Then
        Counter = MSComm1.InBufferCount
        receiveData = ""
        ReceiveArr = MSComm1.Input
        For i = 0 To (Counter - 1) Step 1
            If (ReceiveArr(i) < 16) Then
                receiveData = receiveData & "0" + Hex(ReceiveArr(i))
            Else
                receiveData = receiveData & Hex(ReceiveArr(i))
            End If
        Next i

        Text1.Text = Text1.Text + receiveData
                
    End If
Err:
End Sub

Private Sub YALL_OFF()
    
    Timer2.Enabled = False
c0:
    hexsend Y0_OFF
    If (Text1.Text <> "06") Then
    
    
    GoTo c0
    End If
c1:
    hexsend Y1_OFF
    If (Text1.Text <> "06") Then GoTo c1
    End If
c2:
    hexsend Y2_OFF
    If (Text1.Text <> "06") Then GoTo c2
    End If
c3:
    hexsend Y3_OFF
    If (Text1.Text <> "06") Then GoTo c3
    End If
c4:
    hexsend Y4_OFF
    If (Text1.Text <> "06") Then GoTo c4
    End If
c5:
    hexsend Y5_OFF
    If (Text1.Text <> "06") Then GoTo c5
    End If
c6:
    hexsend Y6_OFF
    If (Text1.Text <> "06") Then GoTo c6
    End If
c7:
    hexsend Y7_OFF
    If (Text1.Text <> "06") Then GoTo c7
    End If
    
    Timer2.Enabled = True

End Sub


Private Sub Timer1_Timer()

        'hexrecv
        
        If (Text1.Text = X0_ON) Then
           
           
           YALL_OFF
           
           Timer2.Enabled = False
c0:
           hexsend Y0_ON
           If (Text1.Text <> "06") Then GoTo c0
           End If
           WMP1_play "\v\v1.mp4"
           'Timer2.Enabled = True

        ElseIf (Text1.Text = X1_ON) Then
        
           
           hexsend Y1_ON
           WMP1_play "\v\v2.mp4"
           
        ElseIf (Text1.Text = X2_ON) Then
        
           
           hexsend Y2_ON
           WMP1_play "\v\v3.mp4"
           
        ElseIf (Text1.Text = X3_ON) Then
        
           
           hexsend Y3_ON
           WMP1_play "\v\v4.mp4"
           
        ElseIf (Text1.Text = X4_ON) Then
        
           
           hexsend Y4_ON
           WMP1_play "\v\v5.mp4"
           
        ElseIf (Text1.Text = X5_ON) Then
        
           
           hexsend Y5_ON
           WMP1_play "\v\v6.mp4"
           
        ElseIf (Text1.Text = X6_ON) Then
        
           
           hexsend Y6_ON
           WMP1_play "\v\v7.mp4"
           
        ElseIf (Text1.Text = X7_ON) Then
        
           
           hexsend Y7_ON
           WMP1_play "\v\v8.mp4"
           
        End If
        
        Text1.Text = ""

End Sub

Private Sub timer2_timer()

    hexsend ASK_STATUS
    
End Sub


Private Sub timer3_timer()

    WMP1_playstatus

End Sub
