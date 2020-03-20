VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置端口号"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleMode       =   0  'User
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "退出"
      Height          =   300
      Left            =   2400
      MaskColor       =   &H000000FF&
      MousePointer    =   4  'Icon
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   300
      Left            =   2400
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3840
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Text            =   "端口号"
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1732
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "请选择正确的端口号，否则程序将无法正常运行。"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        If Combo1.List(Combo1.ListIndex) <> "" Then
            Open App.Path & "\Port.ini" For Output As #1
            Print #1, "[Port]" & Chr(13) & Chr(10) & "Value=" & Right(Combo1.List(Combo1.ListIndex), Len(Combo1.List(Combo1.ListIndex)) - 3)
            Close #1
            Label2.Caption = "设置成功。"
        Else
            Label2.Caption = "设置失败。"
        End If
End Sub

Private Sub Command2_Click()
        End
End Sub

Private Sub Form_Load()
    j = 0
    For i = 1 To 16 Step 1
        If MSComm1.PortOpen = True Then                  '先关闭串口
            MSComm1.PortOpen = False
        End If
        MSComm1.CommPort = i
        On Error Resume Next                            '说明当一个运行时错误发生时，控件转到紧接着发生错误的语句之后的语句，并在此继续运行。访问对象时要使用这种形式而不使用 On Error GoTo。
        MSComm1.PortOpen = True
        If Err.Number <> 8002 Then                      '无效的串口号。这样可以检测到虚拟串口，如果用Err.Number = 0的话检测不到虚拟串口
            If j = 0 Then
                j = i
            End If
            Combo1.AddItem "COM" & i                   '生成串口选择列表
        End If
        MSComm1.PortOpen = False
    Next i
End Sub

