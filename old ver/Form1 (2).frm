VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                   Moossage"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   6735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moossage 0.1 Alpha @Visual Basic ��***************************Powered By bonjour"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.CommandButton Command4 
         Caption         =   "��������������"
         Height          =   735
         Left            =   5520
         TabIndex        =   12
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   615
         Left            =   7320
         TabIndex        =   3
         Top             =   2160
         Width           =   735
      End
      Begin VB.Frame Frame4 
         Caption         =   "ʱ����"
         Height          =   1215
         Left            =   7200
         TabIndex        =   10
         Top             =   1800
         Width           =   1455
         Begin VB.Label Label2 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����-PUA"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   11
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "������Ҫ�������ݡ�Ȼ�󵥻� ��ʼ ��."
         Height          =   1215
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   6975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7320
         TabIndex        =   2
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "�����˳�����"
         Height          =   735
         Left            =   3720
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "��ʼ"
         Height          =   735
         Left            =   1920
         TabIndex        =   5
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��Ϣ�������"
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "�ظ�����"
         Height          =   1215
         Left            =   7200
         TabIndex        =   7
         Top             =   480
         Width           =   1455
         Begin VB.Label Label1 
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����-PUA"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   9
            Top             =   480
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() '��Ϣ������ϰ�ť
MsgBox "��ȷ��ô��"
'SendKeys ^C[, wait]
b = Len(Text1.Text)
Text1.SelStart = 0
Text1.SelLength = b
Text1.SetFocus
'Debug.Print Text1.SelText
'keybd_event VK_Ctrl, MapVirtualKey(VK_Ctrl, 0), 0, 0
SendKeys ("^C")
End Sub

Private Sub Command2_Click() '��ʼ��ť
rep = Text2
For m = 1 To Text2
SendKeys ("^V")
SendKeys ("{ENTER}")
Next
End Sub

Private Sub Command3_Click() '�����˳�����ť
Dim msg As Integer
Form1.Visible = False
msg = MsgBox("ȷ��Ҫ�Ƴ���", vbYesNo, "�˳������ȷ��")
If msg = vbYes Then
End
Else
Form1.Visible = True
End If
End Sub

Private Sub Form_Load()
Dim p As Integer, q As Integer
rep = Text2.Text
End Sub
