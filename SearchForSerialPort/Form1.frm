VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ö˿ں�"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "�˳�"
      Height          =   300
      Left            =   2400
      MaskColor       =   &H000000FF&
      MousePointer    =   4  'Icon
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
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
      Text            =   "�˿ں�"
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
      Caption         =   "��ѡ����ȷ�Ķ˿ںţ���������޷��������С�"
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
            Label2.Caption = "���óɹ���"
        Else
            Label2.Caption = "����ʧ�ܡ�"
        End If
End Sub

Private Sub Command2_Click()
        End
End Sub

Private Sub Form_Load()
    j = 0
    For i = 1 To 16 Step 1
        If MSComm1.PortOpen = True Then                  '�ȹرմ���
            MSComm1.PortOpen = False
        End If
        MSComm1.CommPort = i
        On Error Resume Next                            '˵����һ������ʱ������ʱ���ؼ�ת�������ŷ�����������֮�����䣬���ڴ˼������С����ʶ���ʱҪʹ��������ʽ����ʹ�� On Error GoTo��
        MSComm1.PortOpen = True
        If Err.Number <> 8002 Then                      '��Ч�Ĵ��ںš��������Լ�⵽���⴮�ڣ������Err.Number = 0�Ļ���ⲻ�����⴮��
            If j = 0 Then
                j = i
            End If
            Combo1.AddItem "COM" & i                   '���ɴ���ѡ���б�
        End If
        MSComm1.PortOpen = False
    Next i
End Sub

