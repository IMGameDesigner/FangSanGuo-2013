VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "����ĩ��-������"
   ClientHeight    =   7935
   ClientLeft      =   855
   ClientTop       =   2925
   ClientWidth     =   11970
   Icon            =   "������.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "������.frx":08CA
   MousePointer    =   99  'Custom
   Picture         =   "������.frx":1598
   ScaleHeight     =   7935
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton ������Ϸ 
      Caption         =   "������Ϸ"
      Height          =   495
      Left            =   6000
      TabIndex        =   56
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton ��Ϸ�򵥻� 
      Caption         =   "��Ϸ�򵥻�"
      Height          =   495
      Left            =   4440
      TabIndex        =   55
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Timer ����6hei 
      Interval        =   500
      Left            =   10560
      Top             =   8040
   End
   Begin VB.Timer shijian6 
      Interval        =   500
      Left            =   10080
      Top             =   7440
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ȷ��ս��6"
      Height          =   375
      Left            =   1440
      TabIndex        =   34
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer ��ť 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   11040
      Top             =   9840
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Enabled         =   0   'False
      Height          =   1830
      Left            =   12000
      TabIndex        =   32
      Top             =   10560
      Width           =   3615
   End
   Begin VB.Timer ��ͼ 
      Interval        =   2000
      Left            =   12960
      Top             =   3000
   End
   Begin VB.Timer ��Ϸʤ�� 
      Interval        =   2000
      Left            =   12360
      Top             =   3480
   End
   Begin VB.Timer tishit 
      Interval        =   21000
      Left            =   12360
      Top             =   3000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   240
      TabIndex        =   30
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ȷ��ս��1"
      Height          =   375
      Left            =   2760
      TabIndex        =   29
      Top             =   7440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Text            =   "���"
      Top             =   6960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   12360
      TabIndex        =   26
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   12240
      TabIndex        =   25
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   13200
      TabIndex        =   24
      Text            =   "����Ҫ����ĵط�"
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      Caption         =   "�ر�"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   255
      Left            =   240
      TabIndex        =   57
      Top             =   10800
      Width           =   495
   End
   Begin VB.Label tishi 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   14160
      MouseIcon       =   "������.frx":1ABE
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   1200
      Width           =   8415
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      Height          =   735
      Index           =   6
      Left            =   9240
      TabIndex        =   41
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      Height          =   735
      Index           =   7
      Left            =   9480
      TabIndex        =   42
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      Height          =   735
      Index           =   8
      Left            =   9600
      MouseIcon       =   "������.frx":3108
      TabIndex        =   43
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      Height          =   735
      Index           =   9
      Left            =   9720
      TabIndex        =   44
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FF00&
      Height          =   735
      Index           =   10
      Left            =   9840
      TabIndex        =   45
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Height          =   735
      Index           =   1
      Left            =   10080
      TabIndex        =   36
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FFFF&
      Height          =   735
      Index           =   2
      Left            =   10320
      TabIndex        =   37
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FFFF&
      Height          =   735
      Index           =   3
      Left            =   10560
      TabIndex        =   38
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FFFF&
      Height          =   735
      Index           =   4
      Left            =   10680
      TabIndex        =   39
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H000080FF&
      Height          =   735
      Index           =   5
      Left            =   10800
      TabIndex        =   40
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ս��2"
      Height          =   615
      Left            =   18240
      TabIndex        =   53
      ToolTipText     =   "����13-300������>13-600"
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ս��1"
      Height          =   615
      Left            =   17160
      TabIndex        =   52
      ToolTipText     =   "ÿ�����ﶼ��ʹ���ҷ�Χ��ͬ"
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��ı2"
      Height          =   615
      Left            =   18240
      TabIndex        =   51
      ToolTipText     =   "12����"
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��ı1"
      Height          =   615
      Left            =   17160
      TabIndex        =   50
      ToolTipText     =   "14����"
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   16080
      TabIndex        =   49
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      TabIndex        =   48
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15000
      TabIndex        =   47
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15000
      TabIndex        =   46
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(0)"
      Height          =   615
      Index           =   0
      Left            =   8520
      TabIndex        =   35
      Top             =   7920
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   "�����3�������ƶ�����Ļ�ť"
      Height          =   375
      Left            =   10200
      TabIndex        =   33
      Top             =   9360
      Width           =   1215
   End
   Begin VB.Image �հ�ͼ 
      Height          =   735
      Left            =   11760
      Top             =   0
      Width           =   615
   End
   Begin VB.Image ս����ʽ 
      Height          =   5565
      Left            =   9240
      Picture         =   "������.frx":4752
      Top             =   10200
      Width           =   4425
   End
   Begin VB.Image Image4 
      Height          =   2445
      Left            =   11880
      Picture         =   "������.frx":CE7E
      Top             =   -240
      Visible         =   0   'False
      Width           =   12975
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "����������?f1������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12000
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "�޽������������ã�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   6480
      TabIndex        =   23
      Top             =   10200
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "��ʳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   6480
      TabIndex        =   22
      Top             =   9960
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "��Ǯ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6480
      TabIndex        =   21
      Top             =   9720
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "�˿�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   6480
      TabIndex        =   20
      Top             =   9480
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   19
      Top             =   10200
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "��ҵ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   18
      Top             =   9960
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "ũҵ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   17
      Top             =   9720
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   16
      Top             =   9480
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      ForeColor       =   &H00C00000&
      Height          =   1215
      Index           =   13
      Left            =   13320
      TabIndex        =   15
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   1215
      Index           =   12
      Left            =   13320
      TabIndex        =   14
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   1215
      Index           =   11
      Left            =   13200
      TabIndex        =   13
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   1215
      Index           =   10
      Left            =   13320
      TabIndex        =   12
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   1215
      Index           =   9
      Left            =   13200
      TabIndex        =   11
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   1215
      Index           =   8
      Left            =   13200
      TabIndex        =   10
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   1215
      Index           =   7
      Left            =   13200
      TabIndex        =   9
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   1215
      Index           =   6
      Left            =   13080
      TabIndex        =   8
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   1215
      Index           =   5
      Left            =   13080
      TabIndex        =   7
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   1215
      Index           =   4
      Left            =   12960
      TabIndex        =   6
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      ForeColor       =   &H00000000&
      Height          =   1215
      Index           =   3
      Left            =   13080
      TabIndex        =   5
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Index           =   2
      Left            =   14400
      TabIndex        =   4
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   1215
      Index           =   1
      Left            =   14400
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Label3��0��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   8520
      Width           =   8895
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   86
      Left            =   13440
      Top             =   2880
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   85
      Left            =   13560
      Top             =   2760
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   2550
      Index           =   84
      Left            =   13200
      Picture         =   "������.frx":149BE
      Top             =   2760
      Width           =   1785
   End
   Begin VB.Image ���� 
      Height          =   5325
      Index           =   83
      Left            =   13320
      Picture         =   "������.frx":15579
      Top             =   2640
      Width           =   8100
   End
   Begin VB.Image ���� 
      Height          =   2550
      Index           =   82
      Left            =   13440
      Picture         =   "������.frx":1DA11
      Top             =   2520
      Width           =   1785
   End
   Begin VB.Image ���� 
      Height          =   1425
      Index           =   81
      Left            =   13560
      Picture         =   "������.frx":1E8D8
      Top             =   2400
      Width           =   2100
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   80
      Left            =   13680
      Top             =   2280
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   2100
      Index           =   79
      Left            =   13800
      Picture         =   "������.frx":1F0C8
      Top             =   2160
      Width           =   2100
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   78
      Left            =   13920
      Top             =   2040
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   5250
      Index           =   77
      Left            =   14040
      Picture         =   "������.frx":1FB9F
      Top             =   1920
      Width           =   3690
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   76
      Left            =   14160
      Top             =   1800
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   75
      Left            =   13320
      Top             =   1560
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   3780
      Index           =   74
      Left            =   13440
      Picture         =   "������.frx":23418
      Top             =   1440
      Width           =   6885
   End
   Begin VB.Image ���� 
      Height          =   1305
      Index           =   73
      Left            =   13560
      Picture         =   "������.frx":27B5A
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Image ���� 
      Height          =   7125
      Index           =   72
      Left            =   13200
      Picture         =   "������.frx":28119
      Top             =   1320
      Width           =   10245
   End
   Begin VB.Image ���� 
      Height          =   3150
      Index           =   71
      Left            =   13320
      Picture         =   "������.frx":33D16
      Top             =   1200
      Width           =   2100
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   70
      Left            =   13440
      Top             =   1080
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   69
      Left            =   13560
      Top             =   960
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   68
      Left            =   13680
      Top             =   840
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   67
      Left            =   13800
      Top             =   720
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   66
      Left            =   13920
      Top             =   600
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   2550
      Index           =   65
      Left            =   14040
      Picture         =   "������.frx":3500A
      Top             =   480
      Width           =   1785
   End
   Begin VB.Image ���� 
      Height          =   2550
      Index           =   64
      Left            =   14160
      Picture         =   "������.frx":359B4
      Top             =   360
      Width           =   1785
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   63
      Left            =   13920
      Top             =   3600
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   2550
      Index           =   62
      Left            =   14040
      Picture         =   "������.frx":365A7
      Top             =   3480
      Width           =   1785
   End
   Begin VB.Image ���� 
      Height          =   5250
      Index           =   61
      Left            =   14160
      Picture         =   "������.frx":36FF6
      Top             =   3360
      Width           =   3705
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   60
      Left            =   14280
      Top             =   3240
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   59
      Left            =   14400
      Top             =   3120
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   14400
      Index           =   58
      Left            =   13800
      Picture         =   "������.frx":39E28
      Top             =   1680
      Width           =   9600
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   57
      Left            =   13200
      Top             =   3120
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   56
      Left            =   13320
      Top             =   3000
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   2550
      Index           =   55
      Left            =   13440
      Picture         =   "������.frx":4E6DF
      Top             =   2880
      Width           =   1785
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   54
      Left            =   13560
      Top             =   2760
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   53
      Left            =   13680
      Top             =   2640
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   2100
      Index           =   52
      Left            =   13800
      Picture         =   "������.frx":4F475
      Top             =   2520
      Width           =   2100
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   51
      Left            =   13920
      Top             =   2400
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   50
      Left            =   14040
      Top             =   2280
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   49
      Left            =   14160
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   2520
      Index           =   48
      Left            =   14280
      Picture         =   "������.frx":50571
      Top             =   2040
      Width           =   2100
   End
   Begin VB.Image ���� 
      Height          =   6000
      Index           =   47
      Left            =   13200
      Picture         =   "������.frx":5121C
      Top             =   1200
      Width           =   3600
   End
   Begin VB.Image ���� 
      Height          =   2550
      Index           =   46
      Left            =   13200
      Picture         =   "������.frx":55CE0
      Top             =   3000
      Width           =   1950
   End
   Begin VB.Image ���� 
      Height          =   2550
      Index           =   45
      Left            =   12120
      Picture         =   "������.frx":56838
      Top             =   2760
      Width           =   1785
   End
   Begin VB.Image ���� 
      Height          =   1905
      Index           =   44
      Left            =   12240
      Picture         =   "������.frx":574A2
      Top             =   2640
      Width           =   2100
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   43
      Left            =   12360
      Top             =   2520
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   5250
      Index           =   42
      Left            =   12480
      Picture         =   "������.frx":57B37
      Top             =   2400
      Width           =   3705
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   41
      Left            =   12600
      Top             =   2280
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   40
      Left            =   12720
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   5310
      Index           =   39
      Left            =   12840
      Picture         =   "������.frx":5B27D
      Top             =   2040
      Width           =   12030
   End
   Begin VB.Image ���� 
      Height          =   2100
      Index           =   38
      Left            =   12960
      Picture         =   "������.frx":67196
      Top             =   1920
      Width           =   2100
   End
   Begin VB.Image ���� 
      Height          =   3000
      Index           =   37
      Left            =   13080
      Picture         =   "������.frx":6804E
      Top             =   1800
      Width           =   3000
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   36
      Left            =   13200
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   35
      Left            =   13320
      Top             =   1560
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   2550
      Index           =   34
      Left            =   13440
      Picture         =   "������.frx":6B761
      Top             =   1440
      Width           =   1785
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   33
      Left            =   13560
      Top             =   1320
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   2700
      Index           =   32
      Left            =   13200
      Picture         =   "������.frx":6C1AC
      Top             =   1320
      Width           =   2700
   End
   Begin VB.Image ���� 
      Height          =   2550
      Index           =   31
      Left            =   13800
      Picture         =   "������.frx":709D4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   30
      Left            =   13920
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   29
      Left            =   14040
      Top             =   2040
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   2550
      Index           =   28
      Left            =   14160
      Picture         =   "������.frx":71666
      Top             =   1920
      Width           =   1785
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   27
      Left            =   14280
      Top             =   1800
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   26
      Left            =   14400
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   5250
      Index           =   25
      Left            =   13560
      Picture         =   "������.frx":723C0
      Top             =   1680
      Width           =   3705
   End
   Begin VB.Image ���� 
      Height          =   2550
      Index           =   24
      Left            =   13200
      Picture         =   "������.frx":76A33
      Top             =   1680
      Width           =   1785
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   23
      Left            =   13320
      Top             =   1560
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   5250
      Index           =   22
      Left            =   13440
      Picture         =   "������.frx":77444
      Top             =   1440
      Width           =   3705
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   21
      Left            =   13560
      Top             =   1320
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1395
      Index           =   20
      Left            =   13680
      Picture         =   "������.frx":796A3
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Image ���� 
      Height          =   2700
      Index           =   19
      Left            =   13800
      Picture         =   "������.frx":79D02
      Top             =   1080
      Width           =   2700
   End
   Begin VB.Image ���� 
      Height          =   2700
      Index           =   18
      Left            =   13920
      Picture         =   "������.frx":7C90C
      Top             =   960
      Width           =   2700
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   17
      Left            =   14040
      Top             =   840
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   16
      Left            =   14160
      Top             =   720
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   15
      Left            =   11760
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1500
      Index           =   14
      Left            =   11880
      Picture         =   "������.frx":80628
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   13
      Left            =   12000
      Top             =   1440
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   12
      Left            =   12120
      Top             =   1320
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   750
      Index           =   11
      Left            =   12240
      Picture         =   "������.frx":814F6
      Top             =   1200
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   10
      Left            =   12360
      Top             =   1080
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   9
      Left            =   12480
      Top             =   960
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   2550
      Index           =   8
      Left            =   12600
      Picture         =   "������.frx":81E63
      Top             =   840
      Width           =   1860
   End
   Begin VB.Image ���� 
      Height          =   5310
      Index           =   7
      Left            =   12720
      Picture         =   "������.frx":829C0
      Top             =   720
      Width           =   7530
   End
   Begin VB.Image ���� 
      Height          =   1860
      Index           =   6
      Left            =   12840
      Picture         =   "������.frx":89F1F
      Top             =   600
      Width           =   2100
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   5
      Left            =   12960
      Top             =   480
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   2550
      Index           =   4
      Left            =   13080
      Picture         =   "������.frx":8A829
      Top             =   360
      Width           =   1785
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   3
      Left            =   13200
      Top             =   240
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   2
      Left            =   13320
      Top             =   120
      Width           =   855
   End
   Begin VB.Image ���� 
      Height          =   1500
      Index           =   1
      Left            =   13440
      Picture         =   "������.frx":8B437
      Top             =   0
      Width           =   1500
   End
   Begin VB.Image ���� 
      Height          =   1455
      Index           =   0
      Left            =   13080
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   2730
      Index           =   3
      Left            =   -3240
      Picture         =   "������.frx":8C0E8
      Top             =   10200
      Width           =   10845
   End
   Begin VB.Image Image2 
      Height          =   4005
      Index           =   2
      Left            =   12840
      Picture         =   "������.frx":8D658
      Top             =   9600
      Width           =   4995
   End
   Begin VB.Label Label2 
      Caption         =   "diming"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   0
      Left            =   -120
      TabIndex        =   1
      Top             =   8880
      Width           =   3495
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   38
      Left            =   12840
      Picture         =   "������.frx":9B52C
      Top             =   6360
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   37
      Left            =   12840
      Picture         =   "������.frx":9B960
      Top             =   6000
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   36
      Left            =   12600
      Picture         =   "������.frx":9BD94
      Top             =   6840
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   35
      Left            =   12840
      Picture         =   "������.frx":9C1C8
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   34
      Left            =   12720
      Picture         =   "������.frx":9C5FC
      Top             =   6720
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   33
      Left            =   12960
      Picture         =   "������.frx":9CA30
      Top             =   6720
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   32
      Left            =   12840
      Picture         =   "������.frx":9CE64
      Top             =   6720
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   31
      Left            =   12840
      Picture         =   "������.frx":9D298
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   30
      Left            =   12720
      Picture         =   "������.frx":9D6CC
      Top             =   6720
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   29
      Left            =   12960
      Picture         =   "������.frx":9DB00
      Top             =   6360
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   28
      Left            =   12960
      Picture         =   "������.frx":9DF34
      Top             =   6360
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   27
      Left            =   13080
      Picture         =   "������.frx":9E368
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   26
      Left            =   12840
      Picture         =   "������.frx":9E79C
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   25
      Left            =   12840
      Picture         =   "������.frx":9EBD0
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   24
      Left            =   12840
      Picture         =   "������.frx":9F004
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   23
      Left            =   12960
      Picture         =   "������.frx":9F438
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   22
      Left            =   12720
      Picture         =   "������.frx":9F86C
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   21
      Left            =   12720
      Picture         =   "������.frx":9FCA0
      Top             =   6720
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   20
      Left            =   12960
      Picture         =   "������.frx":A00D4
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   19
      Left            =   12960
      Picture         =   "������.frx":A0508
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   18
      Left            =   13080
      Picture         =   "������.frx":A093C
      Top             =   6720
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   17
      Left            =   13320
      Picture         =   "������.frx":A0D70
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   16
      Left            =   13200
      Picture         =   "������.frx":A11A4
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   15
      Left            =   12720
      Picture         =   "������.frx":A15D8
      Top             =   7080
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   14
      Left            =   12720
      Picture         =   "������.frx":A1A0C
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   13
      Left            =   13080
      Picture         =   "������.frx":A1E40
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   12
      Left            =   13200
      Picture         =   "������.frx":A2274
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   11
      Left            =   12840
      Picture         =   "������.frx":A26A8
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   10
      Left            =   12480
      Picture         =   "������.frx":A2ADC
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   9
      Left            =   13440
      Picture         =   "������.frx":A2F10
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   8
      Left            =   13440
      Picture         =   "������.frx":A3344
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   7
      Left            =   12960
      Picture         =   "������.frx":A3778
      Top             =   6360
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   6
      Left            =   12720
      Picture         =   "������.frx":A3BAC
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   5
      Left            =   12960
      Picture         =   "������.frx":A3FE0
      Top             =   6360
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   4
      Left            =   12720
      Picture         =   "������.frx":A4414
      Top             =   6240
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   3
      Left            =   12840
      Picture         =   "������.frx":A4848
      Top             =   6720
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   2
      Left            =   12600
      Picture         =   "������.frx":A4C7C
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   1
      Left            =   12840
      Picture         =   "������.frx":A50B0
      Top             =   5880
      Width           =   1005
   End
   Begin VB.Image Image3 
      Height          =   1770
      Index           =   0
      Left            =   120
      Picture         =   "������.frx":A54E4
      Top             =   9120
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "����δѡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11280
      TabIndex        =   0
      Top             =   9480
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   0
      Left            =   9840
      Picture         =   "������.frx":A5918
      Top             =   8760
      Width           =   1050
   End
   Begin VB.Image Image2 
      Height          =   7725
      Index           =   1
      Left            =   7200
      Picture         =   "������.frx":A5DBC
      Top             =   9720
      Width           =   11415
   End
   Begin VB.Image Image1 
      Height          =   1965
      Index           =   25
      Left            =   7560
      Picture         =   "������.frx":B01FC
      Top             =   10080
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   1965
      Index           =   24
      Left            =   6480
      Picture         =   "������.frx":B14D9
      Top             =   10200
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   1965
      Index           =   23
      Left            =   5160
      Picture         =   "������.frx":B2723
      Top             =   10200
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   1965
      Index           =   22
      Left            =   3840
      Picture         =   "������.frx":B3BED
      Top             =   10320
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   1965
      Index           =   21
      Left            =   3000
      Picture         =   "������.frx":B4BC4
      Top             =   10320
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   1965
      Index           =   20
      Left            =   1560
      Picture         =   "������.frx":B5BAA
      Top             =   10200
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   1965
      Index           =   19
      Left            =   600
      Picture         =   "������.frx":B6DFB
      Top             =   10200
      Width           =   1965
   End
   Begin VB.Image Image1 
      Height          =   1770
      Index           =   18
      Left            =   8280
      Picture         =   "������.frx":B7B58
      Top             =   8880
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   1770
      Index           =   17
      Left            =   8640
      Picture         =   "������.frx":B7F8C
      Top             =   8760
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   1785
      Index           =   16
      Left            =   7800
      Picture         =   "������.frx":B83C0
      Top             =   8880
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   15
      Left            =   7320
      Picture         =   "������.frx":BBE18
      Top             =   8880
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   1785
      Index           =   14
      Left            =   6840
      Picture         =   "������.frx":BFBF0
      Top             =   8880
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   13
      Left            =   6600
      Picture         =   "������.frx":C1EEC
      Top             =   9000
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   1785
      Index           =   12
      Left            =   6120
      Picture         =   "������.frx":C5B70
      Top             =   8880
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   1785
      Index           =   11
      Left            =   5760
      Picture         =   "������.frx":C8190
      Top             =   9000
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   10
      Left            =   5280
      Picture         =   "������.frx":CBD9C
      Top             =   9000
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   1785
      Index           =   9
      Left            =   4680
      Picture         =   "������.frx":CFAFC
      Top             =   9000
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   8
      Left            =   4200
      Picture         =   "������.frx":D3CC4
      Top             =   9120
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   1785
      Index           =   7
      Left            =   3720
      Picture         =   "������.frx":D7A3C
      Top             =   9120
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   1785
      Index           =   6
      Left            =   3240
      Picture         =   "������.frx":DBA94
      Top             =   9120
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   1785
      Index           =   5
      Left            =   2880
      Picture         =   "������.frx":DDE1C
      Top             =   9000
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   1785
      Index           =   4
      Left            =   2400
      Picture         =   "������.frx":E1D7C
      Top             =   9000
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   1785
      Index           =   3
      Left            =   1920
      Picture         =   "������.frx":E5A3C
      Top             =   9000
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   1785
      Index           =   2
      Left            =   1440
      Picture         =   "������.frx":E94EC
      Top             =   9000
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   1785
      Index           =   1
      Left            =   960
      Picture         =   "������.frx":EBAF8
      Top             =   9000
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   0
      Left            =   9360
      Picture         =   "������.frx":EE178
      Top             =   8760
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   9015
      Index           =   26
      Left            =   12600
      Picture         =   "������.frx":EE61C
      Stretch         =   -1  'True
      Top             =   -360
      Width           =   15225
   End
   Begin VB.Image Image5 
      Height          =   6540
      Left            =   13320
      Picture         =   "������.frx":110CF9
      Top             =   9480
      Visible         =   0   'False
      Width           =   10950
   End
   Begin VB.Label Label15 
      Height          =   5175
      Left            =   12000
      TabIndex        =   54
      Top             =   1680
      Width           =   15375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim weiren As Long
     '&H00FF0000&����&H00FFFF00&ǳ��&H0000FF00&��&H000000FF&��&H0000FFFF&��&H000080FF&��&H0&��
     Dim �� As String
     Dim �� As String
     Dim �� As String
     Dim �� As String
     Dim ǳ�� As String
     Dim ���� As String
     Dim formse As String
     Dim youxijiandanhua As Boolean
    Dim yici1 As Long
    Dim yicizuobi As Long
Dim a1 As Long 'Dim fangjilian As Long
Dim an1 As Long '����16���---����������-++++δ   ��i
Dim huan1 As Long
Dim wanjiabeida As Boolean

Dim huihe As Long
Dim wojx As Long
Dim dijx As Long
Dim dij6 As Long
Dim jiliang6 As Long
Dim f61 As Long
Dim zhuang6 As Long
Dim time6 As Long
Dim wobing(4) As Long
Dim dibing(4) As Long
Dim woliang As Long
Dim diliang As Long
Dim j6 As Long
Dim dicheng6 As Long
Dim s As Long
Dim xiaodui16(17) As Long
Dim ����6ѡ�� As Long
Dim di����6ѡ�� As Long
Dim jiang80 As Long
Dim y As Long
Dim b1 As Long
Dim b2 As Long
Dim huan2 As Long
Dim huan3 As Long
Dim yyyy As Long
Dim wang As Long
Dim wang16(17) As Long '����
Dim jyuanwang(85) As Long '����
Dim min As Long
Dim way(39, 39) As Boolean
Dim f  As Long
Dim f1 As Long
Dim f2 As Long
Dim fx As Long
Dim fy As Long
Dim ff As Long
Dim diannaocheng(39) As Long
Dim diannaochengm As Long
Dim chuzhengmian As Boolean
Dim chuzhengmian2 As Boolean '�����Լ��ǳ�
Dim wanjiaxiaodui As Long '���С��
Dim xuanrentu As Boolean
Dim hongx As Long
Dim hongy As Long
Dim n As Long
Dim m As Long
Dim lvx As Long
Dim lvy As Long
Dim kongzhizhe(11, 8) As Long '0->diannao,1->p1,(2->p2) '�ǳر���
Dim kongzhizheij(39) As Long
Dim xianshichengchi As Long
Dim dizhiij(39) As String
Dim diming(11, 8) As String
Dim dizhi(11, 8) As Long '����ת��Ϊ����
Dim chengx(39) As Long '����ת��Ϊ����
Dim chengy(39) As Long
Dim suoshu(11, 8) As Long
Dim sbsuoshu(39) As Long
Dim sbkongzhizhe(39) As Long
Dim sbchengming(39) As String
'Dim sbtaishou(39) As String
Dim sbnongye(39) As Long
Dim sbshangye(39) As Long
Dim sbminzhong(39) As Long
Dim sbrenkou(39) As Long
Dim sbjinqian(39) As Long
Dim sbliangshi(39) As Long
Dim sbhoubeibingli(39) As Long '�ǳر���
Dim m2 As Long
Dim dichengjiang(100) As Long
Dim fuhuojming(10000) As String
Dim fuhuojzai(10000) As Long
Dim fuhuojwuli(10000) As Long
Dim fuhuojzhili(10000) As Long
Dim fuhuojbingzhong(10000) As Long
Dim fuhuojhao(10000) As Long
Dim fuhuojf As Long
Dim jm As Long
Dim jming(85) As String '�������1
Dim chengjianghao(100) As Long '0
Dim kongxian(85) As Boolean  '���н���2
Dim shiyongjiangling As Long
Dim jshenfen(85) As Long '1->wang,-1->fulu       3
Dim jwang(85) As Long '4
Dim jzai(85) As Long '5
Dim jji(85) As Long '6
Dim jwuli(85) As Long '7
Dim zuoyoujiangling As Long
Dim jzhili(85) As Long '8
Dim jzhong(85) As Long '9
Dim jjing(85) As Long '10
Dim jtili(85) As Long '11
Dim jbingzhong(85) As Long '12
Dim jbingli(85) As Long '�������'13
Dim kaitianfei As Long '���ñ���
Dim kaishangfei As Long
Private Sub ��ʳû��()
Dim cf1 As Long
Dim cf2 As Long
For cf1 = 1 To 38
If sbliangshi(cf1) < 0 Then
For cf2 = 1 To 84
If jbingli(cf2) > 0 And jzai(cf2) = cf1 Then
jbingli(cf2) = jbingli(cf2) \ 2
End If
Next
sbhoubeibingli(cf1) = 0
End If
Next
End Sub '�������½�
Private Sub ���켼��ʹ��() '�����
Dim cf1 As Long
Dim cf2 As Long
For cf1 = 1 To 84
If (jming(cf1) = "����") And jbingli(cf1) < 1500 And jshenfen(cf1) <> -1 Then
jbingli(cf1) = jbingli(cf1) + 400
 List1.AddItem "���������˼��ܡ�������--ʿ����Ϊ��" & jbingli(cf1) & "//"
End If
If (jming(cf1) = "������") And jbingli(cf1) < 9000 And jshenfen(cf1) <> -1 Then
jbingli(cf1) = jbingli(cf1) + 2400
 List1.AddItem "�����������˼��ܡ�������--ʿ����Ϊ��" & jbingli(cf1) & "//"
End If
If jming(cf1) = "����" And sbminzhong(jzai(cf1)) < 200 And jshenfen(cf1) <> -1 Then
sbminzhong(jzai(cf1)) = sbminzhong(jzai(cf1)) + 3
List1.AddItem "���ڷ������ܡ�������--���ұ�Ϊ��" & sbminzhong(jzai(cf1)) & "//"
End If


If jming(cf1) = "�ܲ�" And jbingli(cf1) > 2000 And jshenfen(cf1) <> -1 Then
 For cf2 = 1 To 38
 If way(cf2, jzai(cf1)) = True And sbliangshi(cf2) > 0 And sbsuoshu(cf2) <> jwang(39) Then
 sbliangshi(cf2) = 0
  If jming(sbsuoshu(cf2)) = "Ԭ��" And way(cf2, jzai(cf1)) = True Then
  Image3(cf2).Picture = Image3(jzai(39)).Picture
  Dim df4 As Long
  For df4 = 1 To 84
  If jzai(df4) = cf2 Then
  jwang(df4) = jwang(39)
  End If
  Next
  sbsuoshu(cf2) = jwang(39)
  If jwang(39) = wang Then
  kongzhizhe(chengx(cf2), chengy(cf2)) = 1
  End If
  End If
 End If
 Next
List1.AddItem "�ܲٷ������ܡ�ȼ����" & "//"
End If


If jming(cf1) = "���" And jshenfen(cf1) <> -1 And (jiang80 / 80) Mod 3 = 0 Then
For cf2 = 1 To 84
If jzai(cf2) = jzai(cf1) And jshenfen(cf2) <> -1 Then
inc jzhong(cf2)
End If
Next
List1.AddItem "��߷������ܡ�����" & "//"
End If
If jming(cf1) = "����" And jshenfen(cf1) <> -1 And jbingli(cf1) > 10000 Then
jbingli(cf1) = 5000
For cf2 = 1 To 84
If jzai(cf2) = jzai(cf1) And jshenfen(cf2) <> -1 Then
inc jwuli(cf2)
End If
Next
List1.AddItem "���������ܡ��Һ���" & "//"
End If
If jming(cf1) = "����" And jshenfen(cf1) <> -1 And (jiang80 / 80) Mod 5 = 0 Then
For cf2 = 1 To 84
If way(jzai(cf2), jzai(cf1)) = True And jshenfen(cf2) = 0 Then
jzhong(cf2) = jzhong(cf2) - 10
End If
Next
List1.AddItem "������������ʱ�뿪�����������ܡ���䡿" & "//"
End If
If jming(cf1) = "Ԭ��" And jshenfen(cf1) <> -1 And (jiang80 / 80) Mod 7 = 0 Then
sbnongye(jzai(cf1)) = sbnongye(jzai(cf1)) + 5000
List1.AddItem "Ԭ�ܷ������ܡ���ء�" & "//"
End If
''''''''''''����
Next
End Sub
Private Sub �й���(dij, wocheng) 'Ҫ�ģ�wang(����),�����ߣ�ͼ������ۣ�����(����)������Ʒ,��������{С�ӡ�xiaodui16��С������}�����Լ���²δ��
Dim sou As Long
Dim sou2 As Long '���⣺�������ڣ���ÿ�ؿɶ���ҳǣ�wang
Dim ji As Long
Dim yuanwang As Long
Dim yici2 As Long
yuanwang = sbsuoshu(wocheng)
ji = 0
Dim yici As Long
yici = 0
For sou2 = 1 To 84
If jzai(sou2) <> wocheng And jwang(sou2) = sbsuoshu(wocheng) And jshenfen(sou2) <> -1 And yici = 0 Then
ji = sou2 '��������һ��������ԭ��
yici = 1
End If
Next
If jzai(sbsuoshu(wocheng)) <> wocheng Then
ji = sbsuoshu(wocheng)
End If
If yici = 0 Then '���һ�ǳ�
If sbsuoshu(wocheng) = wang Then
wang = ji
ji = 0
End If
End If
jshenfen(ji) = 1
For sou = 1 To 84
If jwang(sou) = yuanwang Then
jwang(sou) = ji
End If
Next
For sou = 1 To 38
If sbsuoshu(sou) = yuanwang Then
sbsuoshu(sou) = ji
End If
Next
For sou = 1 To 84
If jzai(sou) = wocheng Then
jshenfen(sou) = -1
End If
Next
yici2 = 0
For sou = 1 To 84
If jzai(sou) = wocheng And jshenfen(sou) = -1 And yici2 = 0 Then '�У��н���²
jshenfen(sou) = 0
jwang(sou) = jwang(dij)
jzhong(sou) = 70
yici2 = 1
End If
Next
Image3(wocheng).Picture = Image3(jzai(jwang(dij))).Picture  'Image3(jwang(dij)).Picture '������
sbshangye(wocheng) = sbshangye(wocheng) - 800
sbminzhong(wocheng) = sbminzhong(wocheng) - 20
sbsuoshu(wocheng) = jwang(dij)
sbrenkou(wocheng) = sbrenkou(wocheng) - 20000
jjing(dij) = jjing(dij) + 10
If sbsuoshu(wocheng) = wang Then
wang = ji
End If
suoshu(chengx(wocheng), chengy(wocheng)) = jwang(dij)
kongzhizhe(chengx(wocheng), chengy(wocheng)) = 0
End Sub
Private Sub �н�����(dij) 'wanjiabeigong
Dim ff1 As Long
Dim ff2 As Long
Dim wocheng As Long
wocheng = 0
For ff1 = 1 To 38 '5�غ�һ�εд����ң��С��գ�
If way(ff1, jzai(dij)) = True And _
sbsuoshu(ff1) <> jwang(dij) And (sbsuoshu(ff1) = wang Or (youxijiandanhua = True Or jiang80 / 80 Mod 3 = 0)) _
Then
wocheng = ff1
End If
Next

If wocheng <> 0 Then '99####
''''''''''''''''''''''''''''''''''''''
Dim wobings As Long '
Dim dibings As Long
Dim czf As Long '

Dim j
Dim yici
'Ѱ�ҵн�
yici = 0 'dimj
For czf = 1 To 84 '���һ�Ž���:������ң���һ���б���--�Ѿ�
If jzai(czf) = wocheng And jwang(czf) = sbsuoshu(wocheng) And yici = 0 And jbingli(czf) > 0 And jshenfen(czf) <> -1 Then
yici = 1
j = czf
End If
Next
wobings = jbingli(j)
If yici = 0 Then
 tishi.Visible = True
  tishi.Caption = tishi.Caption & sbchengming(jzai(dij)) & jming(dij) & "�ѳǳ�" & sbchengming(wocheng) & "���ݣ��ճǣ�"
  wanjiabeida = True
  kongxian(dij) = False
   �й��� dij, wocheng
Else
jbingli(dij) = jbingli(dij) - jji(j) * 20 - (jzhili(j) - 10) * 200 - (jwuli(j) - 10) * 50 - jbingli(j) / 2
  jbingli(j) = jbingli(j) - jji(dij) * 20 - (jzhili(dij) - 10) * 200 - (jwuli(dij) - 10) * 50 - jbingli(dij) / 2
  If jbingli(j) < 0 Then
  tishi.Visible = True
  tishi.Caption = tishi.Caption & sbchengming(jzai(dij)) & jming(dij) & "�ѳǳ�" & sbchengming(wocheng) & "�������ˡ�"
  kongxian(dij) = False
  wanjiabeida = True
  �й��� dij, wocheng
  Else
  
  wanjiabeida = True
  kongxian(dij) = False
  If youxijiandanhua = False Or sbsuoshu(wocheng) = wang Then
  tishi.Visible = True
  tishi.Caption = tishi.Caption & _
  sbchengming(jzai(dij)) & jming(dij) & "����   ��" & sbchengming(wocheng) & "ԭ�б���" & wobings & "  ս��ʣ��" & jbingli(j) & "  ���е�һ���б�����--�سǽ��죺" & jming(j)
  End If
  End If
End If
 '''''''''''''''''''''''''''''''''''''''''
End If '99####


End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''��δ����
Private Sub diannaodong(jl As Long)
'�ҵĽ��Լ���***���н����ҳǳأ���������ѡ��ӭս,��²����ȷ,ͼ���ҽ��������ı�ط�
If sbminzhong(jzai(jl)) < 100 Then '��ʱ������
sbminzhong(jzai(jl)) = sbminzhong(jzai(jl)) + 1
If youxijiandanhua = True And sbminzhong(jzai(jl)) < 85 Then
kongxian(jl) = False
End If
End If
If sbjinqian(jzai(jl)) < 0 Then '�Ӷ�
sbjinqian(jzai(jl)) = sbjinqian(jzai(jl)) + 20000
sbminzhong(jzai(jl)) = sbminzhong(jzai(jl)) - 10
kongxian(jl) = False
End If
Dim jshu As Long
Dim fs As Long
'wanjiabeida = Falseת��-���غϽ���
a1 = 0
b1 = 0
If jiang80 > 1000000000 Then
jiang80 = 1
End If
'tishi.Visible = True
 jiang80 = jiang80 + 1 '��ʼ��
 
 'ũҵ����
 If sbliangshi(jzai(jl)) < 10000 And kongxian(jl) = True Then
 kongxian(jl) = False
 sbnongye(jzai(jl)) = sbnongye(jzai(jl)) + (jzhili(jl) + jji(jl) * 3) * 30
 sbjinqian(jzai(jl)) = sbjinqian(jzai(jl)) - 1000
 End If
 

If kongxian(jl) = True Then '�����³�
For f = 1 To 38
jshu = 0
For fs = 1 To 84
If jzai(fs) = f And jshenfen(fs) <> -1 Then '���˳Ǽ���
jshu = jshu + 1
End If
Next
If (way(jzai(jl), f) = True Or way(f, jzai(jl)) = True) And kongxian(jl) = True And sbsuoshu(f) = jwang(jl) And jshu = 0 Then
jzai(jl) = f
'kongxian(jl) = False�ҵĲ����ã�����Ҳ������
End If
Next
End If


             '�й�����ʼ
If (wanjiabeida = False Or youxijiandanhua = True) And kongxian(jl) = True And (jbingli(jl) > 1600 Or ((jiang80 / 80) Mod 4 = 0) And (jiang80 / 80) > 2) Then '4�غ�һ�Σ����ٿɹ���
�н����� jl
jjing(jl) = jjing(jl) + 2
End If 'digongji�ع�������

'abcd jl
Dim sb As Long
Dim sf As Long
sb = 0
For sf = 1 To 84
If jzai(sf) = jzai(jl) And jshenfen(sf) <> -1 Then
sb = sb + jbingli(sf)
End If
Next
If sbjinqian(jzai(jl)) > 0 And kongxian(jl) = True And sb < sbnongye(jzai(jl)) + 900 Then '�б�
    jjing(jl) = jjing(jl) + 1
    kongxian(jl) = False
sbjinqian(jzai(jl)) = sbjinqian(jzai(jl)) - 800
sbhoubeibingli(jzai(jl)) = sbhoubeibingli(jzai(jl)) + _
(jwuli(jl) * sbminzhong(jzai(jl)) * (sbrenkou(jzai(jl)) Mod 100000) / 80000) Mod 1000000 + (jwuli(jl) - 10) * 200
End If
Dim yuan As Long
yuan = sbhoubeibingli(jzai(jl))
If sbhoubeibingli(jzai(jl)) > 100 Then  '����
jbingli(jl) = jbingli(jl) + sbhoubeibingli(jzai(jl)) - 1
sbhoubeibingli(jzai(jl)) = 1
Label2(0).Caption = jming(jl) & yuan & "��"
End If
If sbjinqian(jzai(jl)) < 0 And kongxian(jl) = True Then
kongxian(jl) = False
sbshangye(jzai(jl)) = sbshangye(jzai(jl)) + (jzhili(jl) + jji(jl) * 3) * 30
sbjinqian(jzai(jl)) = sbjinqian(jzai(jl)) - kaishangfei
End If
'If KeyCode = vbKeyReturn And lvx = 4 And lvy = 1 And kongxian(shiyongjiangling) = True And jtili(shiyongjiangling) > 0 Then
If kongxian(jl) = True And (yici1 = 0 Or (yici1 < 4 And way(jzai(jl), jzai(wang)) = True)) Then 'Ȱ���н�
yici1 = yici1 + 1
tishi.Visible = True
tishi.Caption = tishi.Caption & "       ��" & sbchengming(jzai(jl)) & jming(jl) & "�Ѿ�ȥ������˵" & "   "
jtili(jl) = jtili(jl) - 16 + jzhili(jl)
Dim quan As Long
For quan = 1 To 84
If jwang(quan) <> jwang(jl) And jshenfen(jl) <> -1 And way(jzai(jl), jzai(quan)) = True Then
If jbingli(quan) < 1000 Then '(jzai(quan) <> jzai(jwang(quan)) Or (jiang80 / 80) Mod 2 = 0) And
jzhong(quan) = jzhong(quan) - jzhili(jl) + 7
'tishi.Caption = tishi.Caption & "z"
sbjinqian(jzai(jl)) = sbjinqian(jzai(jl)) - 2000
End If
If jzhong(quan) < 30 Then
If jwang(quan) = wang Then
List1.AddItem "����Ľ��챻" & jming(quan) & "˵����" & "//"
End If
jwang(quan) = jwang(jl)
jzai(quan) = jzai(jl)
jzhong(quan) = 51
tishi.Caption = tishi.Caption & jming(jl) & "���н��ˡ�" & jming(quan) & "//  "
End If
End If
Next
kongxian(jl) = False
End If
End Sub

Sub ����()
Dim xiaos
xiaos = xianshichengchi
Label2(1) = "����" & jming(sbsuoshu(xianshichengchi))
Label2(2) = "ũҵ" & sbnongye(xiaos)
Label2(3) = "��ҵ" & sbshangye(xiaos)
Label2(4) = "����" & sbminzhong(xiaos)
Label2(5) = "�˿�" & sbrenkou(xiaos)
Label2(6) = "��Ǯ" & sbjinqian(xiaos)
Label2(7) = "��ʳ" & sbliangshi(xiaos)
Label2(8) = "���ձ���" & sbhoubeibingli(xiaos)
If kongzhizhe(hongx, hongy) = 1 Then   'Or kongzhizheij(xiaos) = 1 Then
Label2(0).Caption = Label2(0).Caption & "p1"
End If
End Sub
Sub ����()
'Label3(1) = "����" & jming(chengjianghao(zuoyoujiangling))
 'If kongxian(chengjianghao(zuoyoujiangling)) = False Then
' Label3(2) = "��æ"
 'End If
Label3(1) = "����:" & jming(chengjianghao(zuoyoujiangling))
 If kongxian(chengjianghao(zuoyoujiangling)) = False Then
 Label3(2) = "��æ"
 Else
 Label3(2) = ""
 End If
If jshenfen(chengjianghao(zuoyoujiangling)) = 1 Then
Label3(3) = "�ϴ�"
 End If
 If jshenfen(chengjianghao(zuoyoujiangling)) = 0 Then
 Label3(3) = ""
 End If
 If jshenfen(chengjianghao(zuoyoujiangling)) = -1 Then
 Label3(3) = "��²"
 End If
 

Label3(4) = "����:" & jming(jwang(chengjianghao(zuoyoujiangling)))
For f = 1 To 38
For fx = 0 To 10
For fy = 0 To 7
If dizhi(fx, fy) = jzai(chengjianghao(zuoyoujiangling)) Then
Label3(5) = "����:" & diming(fx, fy)
End If
Next
Next
Next
Label3(6) = "�ȼ���" & jji(chengjianghao(zuoyoujiangling))
Label3(7) = "������" & jwuli(chengjianghao(zuoyoujiangling))
Label3(8) = "������" & jzhili(chengjianghao(zuoyoujiangling))
Label3(9) = "�ҳ϶ȣ�" & jzhong(chengjianghao(zuoyoujiangling))
Label3(10) = "���飺" & jjing(chengjianghao(zuoyoujiangling))
Label3(11) = "������" & jtili(chengjianghao(zuoyoujiangling))
If jbingzhong(chengjianghao(zuoyoujiangling)) = 1 Then
Label3(12) = "ǹ��"
End If
If jbingzhong(chengjianghao(zuoyoujiangling)) = 2 Then
Label3(12) = "����"
End If
If jbingzhong(chengjianghao(zuoyoujiangling)) = 3 Then
Label3(12) = "���"
End If
If jbingzhong(chengjianghao(zuoyoujiangling)) = 4 Then
Label3(12) = "ˮ��"
End If
Label3(13) = jbingli(chengjianghao(zuoyoujiangling))

End Sub
Private Sub inc(x As Long)
x = x + 1
End Sub
Private Sub jian1(x As Long)
x = x - 1
End Sub
Function xiaoduichengshuo(x As Long) As Long
If x = 3 Or x = 4 Or x = 6 Or x = 8 Or x = 9 Or x = 11 Or x = 12 Or x = 15 Or x = 16 Then
xiaoduichengshuo = 1
End If
If x = 5 Or x = 13 Or x = 14 Then
xiaoduichengshuo = 2
End If
If x = 1 Or x = 7 Then
xiaoduichengshuo = 3
End If
If x = 2 Or x = 10 Then
xiaoduichengshuo = 4
End If
End Function

Private Sub Command1_Click() '�ƶ�����ȷ����
Image5.Visible = False 'С��ͼ
For f = 1 To 38
If Text1.Text = dizhiij(f) Then
jzai(shiyongjiangling) = f
shiyongjiangling = 0
Label1.Caption = "" '
Text1.Visible = False '
Command1.Visible = False ''
Command2.Visible = False '
kongxian(shiyongjiangling) = False '������
Image1(26).Visible = True
Image1(0).Visible = True
For ff = 1 To 38
Image3(ff).Visible = True '��֪Ϊɶ��
Next
End If

Next

End Sub

Private Sub Command2_Click()
Image5.Visible = False 'С��ͼ
Text1.Visible = False '
Command1.Visible = False ''
Command2.Visible = False '
Image1(26).Visible = True
Image1(0).Visible = True
For f = 1 To 38
Image3(f).Visible = True '��֪Ϊɶ��
Next
End Sub
Private Sub �ҹ���(j, dicheng) 'Ҫ�ģ�wang(����),�����ߣ�ͼ������ۣ�����(����)������Ʒ,��������{С�ӡ�xiaodui16��С������}�����Լ���²δ��
Dim sou As Long
Dim sou2 As Long
Dim ji As Long
Dim yuanwang As Long
yuanwang = sbsuoshu(dicheng)
ji = 0
Dim yici As Long
yici = 0
For sou2 = 1 To 84 'xun wnag
If jzai(sou2) <> dicheng And jwang(sou2) = sbsuoshu(dicheng) And jshenfen(sou2) <> -1 And yici = 0 Then
ji = sou2 '��������������-��һ��һ������
yici = 1
End If
Next
If jzai(sbsuoshu(dicheng)) <> dicheng Then
ji = sbsuoshu(dicheng)
End If
If yici = 0 Then
ji = 0
End If
jshenfen(ji) = 1
For sou = 1 To 84
If jwang(sou) = yuanwang Then
jwang(sou) = ji
End If
Next
For sou = 1 To 38
If sbsuoshu(sou) = yuanwang Then
sbsuoshu(sou) = ji
End If
Next
For sou = 1 To 16
If xiaodui16(sou) = sbsuoshu(dicheng) Then 'δ�������һ��
'xiaodui16(sou) = jiȱ������
End If
Next
For sou = 1 To 84
If jzai(sou) = dicheng Then
jshenfen(sou) = -1
End If
Next
Image3(dicheng).Picture = Image1(an1).Picture
sbshangye(dicheng) = sbshangye(dicheng) - 800
sbminzhong(dicheng) = sbminzhong(dicheng) - 20
sbsuoshu(dicheng) = jwang(j)
sbrenkou(dicheng) = sbrenkou(dicheng) - 20000
jjing(j) = jjing(j) + 10
suoshu(chengx(dicheng), chengy(dicheng)) = jwang(j)
kongzhizhe(chengx(dicheng), chengy(dicheng)) = 1
End Sub

Private Sub ����(j, dicheng)
Dim wobing As Long
Dim dibing As Long
Dim dij As Long
Dim czf As Long
Dim yici As Long
wobing = jbingli(j)
'Ѱ�ҵн�
yici = 0
For czf = 1 To 84 '���һ�Ž���
If jzai(czf) = dicheng And jwang(czf) = sbsuoshu(dicheng) And yici = 0 And jbingli(czf) > 0 And jshenfen(czf) <> -1 Then
yici = 1
dij = czf
End If
Next
dibing = jbingli(dij)
If yici = 0 Then
kongxian(j) = False
 tishi.Visible = True
  tishi.Caption = "�ǳع��ݣ��ճǣ�"
   �ҹ��� j, dicheng
Else
jbingli(j) = jbingli(j) - jji(dij) * 20 - (jzhili(dij) - 10) * 200 - (jwuli(dij) - 10) * 50 - jbingli(dij) / 2
  jbingli(dij) = jbingli(dij) - jji(j) * 20 - (jzhili(j) - 10) * 200 - (jwuli(j) - 10) * 50 - jbingli(j) / 2
  If jbingli(dij) < 0 Then
  tishi.Visible = True
  tishi.Caption = "�ǳع���"
  kongxian(j) = False
  �ҹ��� j, dicheng
  Else
  kongxian(j) = False
  tishi.Visible = True
  tishi.Caption = "�г�" & sbchengming(dicheng) & "�˷���ԭ�б���" & dibing & "  ս��ʣ�ࣺ" & jbingli(dij) & "  �سǽ��죺" & jming(dij) & "    ��굥����ʧ"
  End If
End If
 
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''δ����
Private Sub Command3_Click() '����ѡ���˳ǳ�ȷ����
For f = 1 To 38 'δд�ɲ����Լ���;�ϳ���
If Text2.Text = dizhiij(f) And Text2.Text <> dizhiij(xianshichengchi) And way(f, xianshichengchi) = True And kongzhizhe(chengx(f), chengy(f)) = 0 Then ''And sbsuoshu(f) <> xiaodui
Image1(26).Visible = True
Image1(0).Visible = True
Text2.Visible = False '
Text2.Visible = False
Command3.Visible = False ''
Command4.Visible = False '
Command5.Visible = False
ս����ʽ.Visible = False
Image5.Visible = False
'����ѡ����������
yyyy = f
'chuzheng1
���� shiyongjiangling, yyyy
inc jjing(shiyongjiangling)
inc jjing(shiyongjiangling)
For ff = 1 To 38
Image3(ff).Visible = True '��֪Ϊɶ��
Next
End If
Next
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''δ����
Private Sub Command4_Click()
Text2.Visible = False '
Text2.Visible = False
Command3.Visible = False ''
Command4.Visible = False '
Command5.Visible = False
Image1(26).Visible = True
Image1(0).Visible = True
ս����ʽ.Visible = False
Image5.Visible = False
For f = 1 To 38
Image3(f).Visible = True
Next
End Sub



Private Sub Command5_Click() '��6����
For f = 1 To 38 'δд�ɲ����Լ���-û��;�ϳ���
If Text2.Text = dizhiij(f) And Text2.Text <> dizhiij(xianshichengchi) And way(f, xianshichengchi) = True And _
kongzhizhe(chengx(f), chengy(f)) = 0 Then 'And sbsuoshu(f) <> xiaodui
Text2.Visible = False '
Command3.Visible = False ''
Command4.Visible = False '
Command5.Value = False
ս����ʽ.Visible = False
Image5.Visible = False 'С��ͼ
Command5.Visible = False
kongxian(shiyongjiangling) = False
yyyy = f
����6 shiyongjiangling, yyyy
End If
Next
End Sub

Private Sub ����6(j As Long, dicheng As Long)


Dim yici As Long '''''''''''''''''''''''''''''
j6 = j
yici = 0: dicheng6 = dicheng
dij6 = 0
For f61 = 1 To 84 '���һ�Ž���
If jzai(f61) = dicheng And jwang(f61) = sbsuoshu(dicheng) And yici = 0 And jbingli(f61) > 0 And jshenfen(f61) <> -1 Then
yici = 1
dij6 = f61
End If
Next
If dij6 > 0 Then '����������������������
For f = 1 To 38
Image3(f).Visible = False
Next
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label13.Visible = True
Label14.Visible = True
For fx = 0 To 10
Label6(fx).Visible = True
Next
Image1(0).Visible = False
Image1(26).Visible = False
wojx = jwuli(j6) * 100 ''''''''''''
dijx = jwuli(dij6) * 100
wobing(1) = jbingli(j6) / 3
wobing(2) = jbingli(j6) / 3
wobing(3) = jbingli(j6) / 3
If jbingli(dij6) > 0 Then
dibing(1) = jbingli(dij6) / 3
dibing(2) = jbingli(dij6) / 3
dibing(3) = jbingli(dij6) / 3
Else
dibing(1) = 1
dibing(2) = 1
dibing(3) = 1
End If
If jbingli(j6) > 0 Then
woliang = 2200 * jbingli(j6) / 3 'һ��һ��
Else
woliang = 100
End If
If jbingli(dij6) > 0 Then
diliang = 8000 * (jbingli(dij6) + 1) / 3
Else
diliang = 10000
End If
'''''''''''''''''''''''''''''''Label6(1).Top = 1200
Label6(1).Left = 2520
Label6(2).Top = 360
Label6(2).Left = 1440
Label6(3).Top = 1200
Label6(3).Left = 1440
Label6(4).Top = 2040
Label6(4).Left = 1440
Label6(5).Top = 1200
Label6(5).Left = 360
Label6(6).Top = 1200
Label6(6).Left = 8400
Label6(7).Top = 360
Label6(7).Left = 9480
Label6(8).Top = 1200
Label6(8).Left = 9480
Label6(9).Top = 2040
Label6(9).Left = 9480
Label6(10).Top = 1200
Label6(10).Left = 10680

'Private Sub ����6��()
Label6(1).Caption = jming(j6) & wojx
Label6(6).Caption = jming(dij6) & dijx
Label6(2).Caption = jbingzhong(j6) & "��" & wobing(1)
Label6(3).Caption = jbingzhong(j6) & "��" & wobing(2)
Label6(4).Caption = jbingzhong(j6) & "��" & wobing(3)
Label6(7).Caption = jbingzhong(dij6) & "��" & dibing(1)
Label6(8).Caption = jbingzhong(dij6) & "��" & dibing(2)
Label6(9).Caption = jbingzhong(dij6) & "��" & dibing(3)
Label6(5).Caption = "����" & woliang
Label6(10).Caption = "����" & diliang
Label6(0).Caption = ����6ѡ��
time6 = 0
zhuang6 = 0
����6ѡ�� = 1
jiliang6 = 0
'End Sub

tishi.Visible = True
tishi.Caption = jming(dij6) & "��������ս��     ���1��2��3��4������깲ͬ���ƣ��б�Խ��Ӿ���Խ�ࡣ"

 '&H00FF0000&����&H00FFFF00&ǳ��&H0000FF00&��&H000000FF&��&H0000FFFF&��&H000080FF&��ɫ&H0&��
 Dim fjk
For fjk = 6 To 10
Label6(fjk).BackColor = &H0& '�߼�����Դ��f
Next

formse = Form1.BackColor
Form1.BackColor = &H0&

����6hei.Enabled = True
Else
If dij6 = 0 Then
tishi.Visible = True
tishi.Caption = "�˳��޽����ǡ�����ս������1�����ٹ��ǣ�����Ϊ�˼Ӿ��顿"
For fjk = 1 To 38
Image3(fjk).Visible = True
Next
Image1(0).Visible = True
Image1(26).Visible = True

End If
End If
End Sub







Private Sub Label11_Click() '����1
If ����6ѡ�� <> 0 Then
Dim f66 As Long
Dim yici As Long
yici = 0
For f66 = 6 To 10
If ����6ѡ�� = 1 And yici = 0 And Label6(f66).Left > Label6(1).Left - 1200 And Label6(f66).Left < Label6(1).Left + 1200 _
And Label6(f66).Top > Label6(1).Top - 1200 And Label6(f66).Top < Label6(1).Top + 1200 And Label6(1).Visible = True And jzhili(j6) >= 14 Then
If f66 = 6 And Label6(6).Visible = True Then
dijx = dijx - woliang / 2
woliang = woliang / 2
yici = 1
tishi.Visible = True
tishi.Caption = "�Ѿ��ö���֮һ������ȥ���𹥲���"
List1.AddItem jming(j6) & "�������ܡ��𹥡�" & "//"
End If
If f66 = 7 And yici = 0 And Label6(7).Visible = True Then
dibing(1) = dibing(1) - woliang / 2
woliang = woliang / 2
yici = 1
tishi.Visible = True
tishi.Caption = "�Ѿ��ö���֮һ������ȥ���𹥲���"
List1.AddItem jming(j6) & "�������ܡ��𹥡�" & "//"
End If
If f66 = 8 And yici = 0 And Label6(8).Visible = True Then
dibing(2) = dibing(2) - woliang / 2
woliang = woliang / 2
yici = 1
tishi.Visible = True
tishi.Caption = "�Ѿ��ö���֮һ������ȥ���𹥲���"
List1.AddItem jming(j6) & "�������ܡ��𹥡�" & "//"
End If
If f66 = 9 And yici = 0 And Label6(9).Visible = True Then
dibing(3) = dibing(3) - woliang / 2
woliang = woliang / 2
yici = 1
tishi.Visible = True
tishi.Caption = "�Ѿ��ö���֮һ������ȥ���𹥲���"
List1.AddItem jming(j6) & "�������ܡ��𹥡�" & "//"
End If
If f66 = 10 And yici = 0 And Label6(10).Visible = True Then
diliang = diliang - woliang * 100
woliang = woliang / 2
yici = 1
tishi.Visible = True
tishi.Caption = "�Ѿ��ö���֮һ������ȥ���𹥲���"
List1.AddItem jming(j6) & "�������ܡ��𹥡�" & "//"
End If
End If: Next: ����6��: End If
End Sub

Private Sub Label12_Click() '����2
If ����6ѡ�� = 1 And Label6(1).Visible = True And jzhili(j6) >= 12 Then
wojx = wojx - 10
 Dim df As Long
For df = 2 To 10
If df >= 2 And df <= 4 Then
Label6(df).BackColor = ��
End If
If df = 5 Then
Label6(df).BackColor = ��
End If
If df = 6 Then
Label6(df).BackColor = ����
End If
If df >= 7 And df <= 9 Then
Label6(df).BackColor = ǳ��
End If
If df = 10 Then
Label6(df).BackColor = ��
End If
Next
End If: ����6��
End Sub

Private Sub Label16_Click()
End
End Sub

Private Sub Timer1_Timer()
If tishi.Visible = False Then
tishi.Caption = ""
End If
End Sub

Private Sub ����6hei_Timer()

Dim df As Long
For df = 2 To 10
If Label6(df).Left > Label6(1).Left - 1200 And Label6(df).Left < Label6(1).Left + 1200 _
And Label6(df).Top > Label6(1).Top - 1200 And Label6(df).Top < Label6(1).Top + 1200 And Label6(1).Visible = True Then
If df >= 2 And df <= 4 Then
Label6(df).BackColor = ��
End If
If df = 5 Then
Label6(df).BackColor = ��
End If
If df = 6 Then
Label6(df).BackColor = ����
End If
If df >= 7 And df <= 9 Then
Label6(df).BackColor = ǳ��
End If
If df = 10 Then
Label6(df).BackColor = ��
End If
Else
Label6(df).BackColor = &H0&
End If
Next
End Sub
Private Sub Label13_Click() '����1
If ����6ѡ�� <> 0 Then
Dim f66 As Long
Dim yici As Long
yici = 0
For f66 = 6 To 10

If ����6ѡ�� = 1 And yici = 0 And Label6(f66).Left > Label6(1).Left - 1200 And Label6(f66).Left < Label6(1).Left + 1200 _
And Label6(f66).Top > Label6(1).Top - 1200 And Label6(f66).Top < Label6(1).Top + 1200 And Label6(1).Visible = True Then
If f66 = 6 And Label6(6).Visible = True Then
dijx = dijx - jwuli(j6)
yici = 1
End If
If f66 = 7 And yici = 0 And Label6(7).Visible = True Then
dibing(1) = dibing(1) - jwuli(j6)
yici = 1
End If
If f66 = 8 And yici = 0 And Label6(8).Visible = True Then
dibing(2) = dibing(2) - jwuli(j6)
yici = 1
End If
If f66 = 9 And yici = 0 And Label6(9).Visible = True Then
dibing(3) = dibing(3) - jwuli(j6)
yici = 1
End If
If f66 = 10 And yici = 0 And Label6(10).Visible = True Then
diliang = diliang - jwuli(j6) * 1000
yici = 1
End If
End If



If ����6ѡ�� = 2 And yici = 0 And Label6(f66).Left > Label6(2).Left - 1200 And Label6(f66).Left < Label6(2).Left + 1200 _
And Label6(f66).Top > Label6(2).Top - 1200 And Label6(f66).Top < Label6(2).Top + 1200 And Label6(2).Visible = True Then
If f66 = 6 And Label6(6).Visible = True Then
dijx = dijx - wobing(1) / 10
yici = 1
End If
If f66 = 7 And yici = 0 And Label6(7).Visible = True Then
dibing(1) = dibing(1) - wobing(1) / 10
yici = 1
End If
If f66 = 8 And yici = 0 And Label6(8).Visible = True Then
dibing(2) = dibing(2) - wobing(1) / 10
yici = 1
End If
If f66 = 9 And yici = 0 And Label6(9).Visible = True Then
dibing(3) = dibing(3) - wobing(1) / 10
yici = 1
End If
If f66 = 10 And yici = 0 And Label6(10).Visible = True Then
diliang = diliang - wobing(1) / 10 * 1000
yici = 1
End If
End If






If ����6ѡ�� = 3 And yici = 0 And Label6(f66).Left > Label6(3).Left - 1200 And Label6(f66).Left < Label6(3).Left + 1200 _
And Label6(f66).Top > Label6(3).Top - 1200 And Label6(f66).Top < Label6(3).Top + 1200 And Label6(3).Visible = True Then
If f66 = 6 And Label6(6).Visible = True Then
dijx = dijx - wobing(2) / 10
yici = 1
End If
If f66 = 7 And yici = 0 And Label6(7).Visible = True Then
dibing(1) = dibing(1) - wobing(2) / 10
yici = 1
End If
If f66 = 8 And yici = 0 And Label6(8).Visible = True Then
dibing(2) = dibing(2) - wobing(2) / 10
yici = 1
End If
If f66 = 9 And yici = 0 And Label6(9).Visible = True Then
dibing(3) = dibing(3) - wobing(2) / 10
yici = 1
End If
If f66 = 10 And yici = 0 And Label6(10).Visible = True Then
diliang = diliang - wobing(2) / 10 * 1000
yici = 1
End If
End If






If ����6ѡ�� = 4 And yici = 0 And Label6(f66).Left > Label6(4).Left - 1200 And Label6(f66).Left < Label6(4).Left + 1200 _
And Label6(f66).Top > Label6(4).Top - 1200 And Label6(f66).Top < Label6(4).Top + 1200 And Label6(4).Visible = True Then
If f66 = 6 And Label6(6).Visible = True Then
dijx = dijx - wobing(3) / 10
yici = 1
End If
If f66 = 7 And yici = 0 And Label6(7).Visible = True Then
dibing(1) = dibing(1) - wobing(3) / 10
yici = 1
End If
If f66 = 8 And yici = 0 And Label6(8).Visible = True Then
dibing(2) = dibing(2) - wobing(3) / 10
yici = 1
End If
If f66 = 9 And yici = 0 And Label6(9).Visible = True Then
dibing(3) = dibing(3) - wobing(3) / 10
yici = 1
End If
If f66 = 10 And yici = 0 And Label6(10).Visible = True Then
diliang = diliang - wobing(3) / 10 * 1000
yici = 1
End If
End If
Next
End If
����6��
End Sub
Private Sub ����6��()
Label6(1).Caption = jming(j6) & wojx
Label6(6).Caption = jming(dij6) & dijx
Label6(2).Caption = jbingzhong(j6) & "��" & wobing(1)
Label6(3).Caption = jbingzhong(j6) & "��" & wobing(2)
Label6(4).Caption = jbingzhong(j6) & "��" & wobing(3)
Label6(7).Caption = jbingzhong(dij6) & "��" & dibing(1)
Label6(8).Caption = jbingzhong(dij6) & "��" & dibing(2)
Label6(9).Caption = jbingzhong(dij6) & "��" & dibing(3)
Label6(5).Caption = "����" & woliang
Label6(10).Caption = "����" & diliang
Dim f66
For f66 = 2 To 4
If wobing(f66 - 1) < 0 Then
Label6(f66).Visible = False
End If
Next
For f66 = 7 To 9
If dibing(f66 - 6) < 0 Then
Label6(f66).Visible = False
End If
Next
If wojx < 0 Then
Label6(1).Visible = False
End If
If dijx < 0 Then
Label6(6).Visible = False
End If
If woliang < 0 Then
Label6(5).Visible = False
End If
If diliang < 0 Then
Label6(10).Visible = False
End If
If (Label6(6).Visible = False And Label6(7).Visible = False And Label6(8).Visible = False And Label6(9).Visible = False) Or Label6(10).Visible = False Then
 For fx = 0 To 10
Label6(fx).Visible = False
Next
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Image1(26).Visible = True
Image1(0).Visible = True
Dim dijqb As Long
dijqb = jbingli(dij6)
jbingli(dij6) = -1
jbingli(j6) = wobing(1) + wobing(2) + wobing(3)
inc jjing(dij6)
jjing(j6) = jjing(j6) + dijqb / 100
For f = 1 To 38
Image3(f).Visible = True
Next
tishi.Visible = True
tishi.Caption = "�ӵذ�ҵս��6��ʤ��"
Form1.BackColor = formse
Else
����6�ж�
End If
End Sub
Private Sub Label14_Click() '����2
If ����6ѡ�� <> 0 Then
Dim f66 As Long
Dim yici As Long
yici = 0
For f66 = 6 To 10
If ����6ѡ�� = 1 And yici = 0 And Label6(f66).Left > Label6(1).Left - 1200 And Label6(f66).Left < Label6(1).Left + 1200 _
And Label6(f66).Top > Label6(1).Top - 1200 And Label6(f66).Top < Label6(1).Top + 1200 And Label6(1).Visible = True And jwuli(j6) >= 14 Then
If f66 = 6 And Label6(6).Visible = True And jtili(j6) > 0 Then
dijx = dijx - 600
yici = 1
jtili(j6) = jtili(j6) - 5
End If
If f66 = 7 And yici = 0 And Label6(7).Visible = True And jtili(j6) > 0 Then
dibing(1) = dibing(1) - 600
yici = 1
jtili(j6) = jtili(j6) - 5
End If
If f66 = 8 And yici = 0 And Label6(8).Visible = True And jtili(j6) > 0 Then
dibing(2) = dibing(2) - 600
yici = 1
jtili(j6) = jtili(j6) - 5
End If
If f66 = 9 And yici = 0 And Label6(9).Visible = True And jtili(j6) > 0 Then
dibing(3) = dibing(3) - 600
yici = 1
jtili(j6) = jtili(j6) - 5
End If
If f66 = 10 And yici = 0 And Label6(10).Visible = True And jtili(j6) > 0 Then
tishi.Visible = True
tishi.Caption = "��������"
End If
If jtili(j6) < 0 Then
tishi.Visible = True
tishi.Caption = "��������"
End If
End If
If ����6ѡ�� = 1 And yici = 0 And Label6(f66).Left > Label6(1).Left - 1200 And Label6(f66).Left < Label6(1).Left + 1200 _
And Label6(f66).Top > Label6(1).Top - 1200 And Label6(f66).Top < Label6(1).Top + 1200 And Label6(1).Visible = True And jwuli(j6) = 13 Then
If f66 = 6 And Label6(6).Visible = True And jtili(j6) > 0 Then
dijx = dijx - 300
yici = 1
jtili(j6) = jtili(j6) - 5
End If
If f66 = 7 And yici = 0 And Label6(7).Visible = True And jtili(j6) > 0 Then
dibing(1) = dibing(1) - 300
yici = 1
jtili(j6) = jtili(j6) - 5
End If
If f66 = 8 And yici = 0 And Label6(8).Visible = True And jtili(j6) > 0 Then
dibing(2) = dibing(2) - 300
yici = 1
jtili(j6) = jtili(j6) - 5
End If
If f66 = 9 And yici = 0 And Label6(9).Visible = True And jtili(j6) > 0 Then
dibing(3) = dibing(3) - 300
yici = 1
jtili(j6) = jtili(j6) - 5
End If
If f66 = 10 And yici = 0 And Label6(10).Visible = True And jtili(j6) > 0 Then
tishi.Visible = True
tishi.Caption = "��������"
End If
If jtili(j6) < 0 Then
tishi.Visible = True
tishi.Caption = "��������"
End If
End If: Next: ����6��: End If
End Sub
Private Sub Label7_Click() '��
If ����6ѡ�� <> 0 And Label6(����6ѡ��).Top > 400 Then
Label6(����6ѡ��).Top = Label6(����6ѡ��).Top - 500
End If
����6�ж�
End Sub


Private Sub Label8_Click() '��
If ����6ѡ�� <> 0 And Label6(����6ѡ��).Top < 4300 Then
Label6(����6ѡ��).Top = Label6(����6ѡ��).Top + 500
End If
����6�ж�
End Sub

Private Sub Label9_Click() '��
If ����6ѡ�� <> 0 And Label6(����6ѡ��).Left > 400 Then
Label6(����6ѡ��).Left = Label6(����6ѡ��).Left - 500
End If
����6�ж�
End Sub
Private Sub Label10_Click() '��
If ����6ѡ�� <> 0 And Label6(����6ѡ��).Left < 11400 Then
Label6(����6ѡ��).Left = Label6(����6ѡ��).Left + 500
End If
����6�ж�
End Sub
Private Sub ����6�ж�()
'''''''ֻ������һ
''''1
If jiang80 / 80 > 5 Then
If jzhili(dij6) < 11 Then
If (diliang < jiliang6 - dibing(1) - dibing(2) - dibing(3) - 1000) Or jwuli(dij6) > 11 Then '�������һ��-����
zhuang6 = 1
End If
If zhuang6 = 1 Then
If Label6(6).Left > Label6(5).Left Then
Label6(6).Left = Label6(6).Left - 500
End If
If Label6(7).Left > Label6(5).Left Then
Label6(7).Left = Label6(7).Left - 500
End If
If Label6(7).Top > Label6(5).Top Then
Label6(7).Top = Label6(7).Top - 500
End If
If Label6(1).Visible = True Then
If Label6(9).Left > Label6(1).Left Then
Label6(9).Left = Label6(9).Left - 500
End If
If Label6(9).Top > Label6(1).Top Then
Label6(9).Top = Label6(9).Top - 500
End If
If Label6(9).Left < Label6(1).Left Then
Label6(9).Left = Label6(9).Left + 500
End If
If Label6(9).Top < Label6(1).Top Then
Label6(9).Top = Label6(9).Top + 500
End If
Else
If Label6(9).Left > Label6(5).Left Then
Label6(9).Left = Label6(9).Left - 500
End If
If Label6(9).Top > Label6(5).Top Then
Label6(9).Top = Label6(9).Top - 500
End If
If Label6(9).Left < Label6(5).Left Then
Label6(9).Left = Label6(9).Left + 500
End If
If Label6(9).Top < Label6(5).Top Then
Label6(9).Top = Label6(9).Top + 500
End If
End If
End If
di����6ѡ�� = 1
����6�н�����1
''''2
di����6ѡ�� = 2
����6�н�����1
''''4
di����6ѡ�� = 4
����6�н�����1
''''3
di����6ѡ�� = 3
Label6(8).Left = Label6(10).Left
����6�н�����1
jiliang6 = diliang
End If
''''''''''''''''''''''''''''''''''''''
If jzhili(dij6) > 11 Then
'If zhuang6 = 1 Then
Label6(6).Top = Label6(10).Top
Label6(6).Left = Label6(10).Left
Label6(7).Top = Label6(10).Top
Label6(7).Left = Label6(10).Left
Label6(9).Top = Label6(10).Top
Label6(9).Left = Label6(10).Left
'End If
di����6ѡ�� = 1
����6�н�����1
''''2
di����6ѡ�� = 2
����6�н�����1
''''4
di����6ѡ�� = 4
����6�н�����1
''''3
di����6ѡ�� = 3
Label6(8).Left = Label6(10).Left
����6�н�����1
End If
Else
di����6ѡ�� = 1
����6�н�����1
����6�н�����1
''''2
di����6ѡ�� = 2
����6�н�����1
''''4
di����6ѡ�� = 4
����6�н�����1
''''3
di����6ѡ�� = 3
Label6(8).Left = Label6(10).Left
����6�н�����1
End If
End Sub


Private Sub ����6�н�����1()

If di����6ѡ�� <> 0 Then
Dim f66 As Long
Dim yici As Long
yici = 0
For f66 = 1 To 5

If Label6(f66).Top > Label6(6).Top - 1200 And Label6(f66).Top < Label6(6).Top + 1200 _
And di����6ѡ�� = 1 And yici = 0 And Label6(f66).Left > Label6(6).Left - 1200 And Label6(f66).Left < Label6(6).Left + 1200 And Label6(6).Visible = True Then
If (f66 = 1 And Label6(1).Visible = True) Then
wojx = wojx - jwuli(dij6)
yici = 1
End If
If f66 = 2 And yici = 0 And Label6(2).Visible = True Then
wobing(1) = wobing(1) - jwuli(dij6)
yici = 1
End If
If f66 = 3 And yici = 0 And Label6(3).Visible = True Then
wobing(2) = wobing(2) - jwuli(dij6)
yici = 1
End If
If f66 = 4 And yici = 0 And Label6(4).Visible = True Then
wobing(3) = wobing(3) - jwuli(dij6)
yici = 1
End If
If f66 = 5 And yici = 0 And Label6(5).Visible = True Then
woliang = woliang - jwuli(dij6) * 1000
yici = 1
End If
End If




If Label6(f66).Top > Label6(7).Top - 1200 And Label6(f66).Top < Label6(7).Top + 1200 _
And di����6ѡ�� = 2 And yici = 0 And Label6(f66).Left > Label6(7).Left - 1200 And Label6(f66).Left < Label6(7).Left + 1200 And Label6(7).Visible = True Then
If f66 = 1 And Label6(1).Visible = True Then
wojx = wojx - dibing(1) / 10
yici = 1
End If
If f66 = 2 And yici = 0 And Label6(2).Visible = True Then
wobing(1) = wobing(1) - dibing(1) / 10
yici = 1
End If
If f66 = 3 And yici = 0 And Label6(3).Visible = True Then
wobing(2) = wobing(2) - dibing(1) / 10
yici = 1
End If
If f66 = 4 And yici = 0 And Label6(4).Visible = True Then
wobing(3) = wobing(3) - dibing(1) / 10
yici = 1
End If
If f66 = 5 And yici = 0 And Label6(5).Visible = True Then
woliang = woliang - dibing(1) / 10 * 1000
yici = 1

End If
End If



If Label6(f66).Top > Label6(8).Top - 1200 And Label6(f66).Top < Label6(8).Top + 1200 _
And di����6ѡ�� = 3 And yici = 0 And Label6(f66).Left > Label6(8).Left - 1200 And Label6(f66).Left < Label6(8).Left + 1200 And Label6(8).Visible = True Then
If f66 = 1 And Label6(1).Visible = True Then
wojx = wojx - dibing(2) / 10
yici = 1
End If
If f66 = 2 And yici = 0 And Label6(2).Visible = True Then
wobing(1) = wobing(1) - dibing(2) / 10
yici = 1
End If
If f66 = 3 And yici = 0 And Label6(3).Visible = True Then
wobing(2) = wobing(2) - dibing(2) / 10
yici = 1
End If
If f66 = 4 And yici = 0 And Label6(4).Visible = True Then
wobing(3) = wobing(3) - dibing(2) / 10
yici = 1
End If
If f66 = 5 And yici = 0 And Label6(5).Visible = True Then
woliang = woliang - dibing(2) / 10 * 1000
yici = 1
End If
End If






If Label6(f66).Top > Label6(9).Top - 1200 And Label6(f66).Top < Label6(9).Top + 1200 _
And di����6ѡ�� = 4 And yici = 0 And Label6(f66).Left > Label6(9).Left - 1200 And Label6(f66).Left < Label6(9).Left + 1200 And Label6(9).Visible = True Then
If f66 = 1 And Label6(1).Visible = True Then
wojx = wojx - dibing(3) / 10
yici = 1
End If
If f66 = 2 And yici = 0 And Label6(2).Visible = True Then
wobing(1) = wobing(1) - dibing(3) / 10
yici = 1
End If
If f66 = 3 And yici = 0 And Label6(3).Visible = True Then
wobing(2) = wobing(2) - dibing(3) / 10
yici = 1
End If
If f66 = 4 And yici = 0 And Label6(4).Visible = True Then
wobing(3) = wobing(3) - dibing(3) / 10
yici = 1
End If
If f66 = 5 And yici = 0 And Label6(5).Visible = True Then
woliang = woliang - dibing(3) / 10 * 1000
yici = 1
End If
End If







Next
End If
Label6(1).Caption = jming(j6) & wojx
Label6(6).Caption = jming(dij6) & dijx
Label6(2).Caption = jbingzhong(j6) & "��" & wobing(1)
Label6(3).Caption = jbingzhong(j6) & "��" & wobing(2)
Label6(4).Caption = jbingzhong(j6) & "��" & wobing(3)
Label6(7).Caption = jbingzhong(dij6) & "��" & dibing(1)
Label6(8).Caption = jbingzhong(dij6) & "��" & dibing(2)
Label6(9).Caption = jbingzhong(dij6) & "��" & dibing(3)
Label6(5).Caption = "����" & woliang
Label6(10).Caption = "����" & diliang
For f66 = 2 To 4
If wobing(f66 - 1) < 0 Then
Label6(f66).Visible = False
End If
Next
For f66 = 7 To 9
If dibing(f66 - 6) < 0 Then
Label6(f66).Visible = False
End If
Next
If wojx < 0 Then
Label6(1).Visible = False
End If
If dijx < 0 Then
Label6(6).Visible = False
End If
If woliang < 0 Then
Label6(5).Visible = False
End If
If diliang < 0 Then
Label6(10).Visible = False
End If
If (Label6(1).Visible = False And Label6(2).Visible = False And Label6(3).Visible = False And Label6(4).Visible = False) Or Label6(5).Visible = False Then
 For fx = 0 To 10
Label6(fx).Visible = False
Next
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Image1(26).Visible = True
Image1(0).Visible = True
Dim ji As Long
ji = jbingli(j6)
jbingli(j6) = -1
jbingli(dij6) = dibing(1) + dibing(2) + dibing(3)
inc jjing(j6)
jjing(dij6) = jjing(dij6) + 15
sbliangshi(jzai(j6)) = sbliangshi(jzai(j6)) - ji * 10
For f = 1 To 38
Image3(f).Visible = True
Next
tishi.Visible = True
tishi.Caption = "�ӵذ�ҵս��6��ʧ��"
Form1.BackColor = formse
End If
End Sub


Private Sub shijian6_Timer()
inc time6
If ����6ѡ�� = 1 Then
Label6(0).Caption = jming(j6)
End If
If ����6ѡ�� = 2 Then
Label6(0).Caption = "����1"
End If
If ����6ѡ�� = 3 Then
Label6(0).Caption = "����2"
End If
If ����6ѡ�� = 4 Then
Label6(0).Caption = "����3"
End If
Label6(0).Caption = Label6(0).Caption & "ʱ��" & time6
woliang = woliang - wobing(1) - wobing(2) - wobing(3)
diliang = diliang Mod 200000000 - dibing(1) - dibing(2) - dibing(3)
Label6(5).Caption = "����" & woliang
Label6(10).Caption = "����" & diliang
End Sub




'''''''''''''''''''''''''''''''''���ϳ���6
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''δ����
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyQ Then
'jbingli(1) = 54321
End If
If Image2(2).Visible = True And KeyCode = vbKeyF7 Then
Image2(2).Visible = False
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
Open "d:\����������.in" For Input As #1
Dim s As Long, fs As Long
Input #1, s
an1 = s
an1 = 1

Input #1, s
wang = s
For fs = 1 To 38
Input #1, s
sbsuoshu(fs) = s
If sbsuoshu(fs) = wang Then
kongzhizhe(chengx(fs), chengy(fs)) = 1
End If
Next
Dim j84h16(16) As Long, h16 As Long
For fs = 1 To 84
Input #1, s
jwang(fs) = s
If jwang(fs) = fs Then
h16 = h16 + 1
j84h16(h16) = fs
End If
Next
For fs = 1 To 38
If sbsuoshu(fs) = wang Then
Image3(fs).Picture = Image1(1).Picture
Else
Image3(fs).Picture = �հ�ͼ
End If
Next
tishi.Visible = True
tishi.Caption = "�����ľ�����һ�ַ������ǳ�����" '& an1 & wang

 For fs = 1 To 84
Input #1, s
jshenfen(fs) = s
Next
 For fs = 1 To 84
Input #1, s
kongxian(fs) = s
Next
 For fs = 1 To 84
Input #1, s
jzai(fs) = s
Next
 For fs = 1 To 84
Input #1, s
jji(fs) = s
Next
 For fs = 1 To 84
Input #1, s
jjing(fs) = s
Next
 For fs = 1 To 84
Input #1, s
jzhong(fs) = s
Next
 For fs = 1 To 84
Input #1, s
jtili(fs) = s
Next
 For fs = 1 To 84
Input #1, s
jbingli(fs) = s
Next
 For fs = 1 To 38
Input #1, s
sbnongye(fs) = s
Next
For fs = 1 To 38
Input #1, s
sbshangye(fs) = s
Next
For fs = 1 To 38
Input #1, s
sbminzhong(fs) = s
Next
For fs = 1 To 38
Input #1, s
sbrenkou(fs) = s
Next
For fs = 1 To 38
Input #1, s
sbjinqian(fs) = s
Next
For fs = 1 To 38
Input #1, s
sbliangshi(fs) = s
Next
For fs = 1 To 38
Input #1, s
sbhoubeibingli(fs) = s
Next
For fs = 1 To 84
Input #1, s
jwuli(fs) = s
Next
For fs = 1 To 84
Input #1, s
jzhili(fs) = s
Next
For fs = 1 To 100
Input #1, s
If s <> 0 Then
jming(s) = fuhuojming(fs)
End If
Next
Close 1
End If

If Image1(26).Visible = True And KeyCode = vbKeyF8 And Image2(2).Visible = False Then
Open "d:\����������.in" For Output As 2
Dim s2 As Long, fs2 As Long, ssss As String
s2 = an1
ssss = ssss & s2 & " "
s2 = wang
ssss = ssss & s2 & " "
For fs2 = 1 To 38
s2 = sbsuoshu(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 84
s2 = jwang(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 84
s2 = jshenfen(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 84
s2 = kongxian(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 84
s2 = jzai(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 84
s2 = jji(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 84
s2 = jjing(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 84
s2 = jzhong(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 84
s2 = jtili(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 84
s2 = jbingli(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 38
s2 = sbnongye(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 38
s2 = sbshangye(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 38
s2 = sbminzhong(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 38
s2 = sbrenkou(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 38
s2 = sbjinqian(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 38
s2 = sbliangshi(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 38
s2 = sbhoubeibingli(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 84
s2 = jwuli(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 84
s2 = jzhili(fs2)
ssss = ssss & s2 & " "
Next
For fs2 = 1 To 100
s2 = fuhuojhao(fs2)
ssss = ssss & s2 & " "
Next
Write #2, ssss
Close 2
tishi.Visible = True
tishi.Caption = "�Ѵ�"
End If

If Label6(0).Visible = True Then
If Label6(0).Visible = True And KeyCode = vbKey1 Then
����6ѡ�� = 1
End If
If Label6(0).Visible = True And KeyCode = vbKey2 Then
����6ѡ�� = 2
End If
If Label6(0).Visible = True And KeyCode = vbKey3 Then
����6ѡ�� = 3
End If
If Label6(0).Visible = True And KeyCode = vbKey4 Then
����6ѡ�� = 4
End If
End If

If Image4.Visible = True Then
If KeyCode = vbKeyLeft And jbingli(shiyongjiangling) - 100 > 0 Then '�������
sbhoubeibingli(xianshichengchi) = sbhoubeibingli(xianshichengchi) + 100
jbingli(shiyongjiangling) = jbingli(shiyongjiangling) - 100
����
End If
If KeyCode = vbKeyRight And sbhoubeibingli(xianshichengchi) - 100 > 0 Then
sbhoubeibingli(xianshichengchi) = sbhoubeibingli(xianshichengchi) - 100
jbingli(shiyongjiangling) = jbingli(shiyongjiangling) + 100
����
End If
End If
If Label3(1).Visible = True And KeyCode = vbKeyUp And kongxian(chengjianghao(zuoyoujiangling)) = True _
And jshenfen(chengjianghao(zuoyoujiangling)) <> -1 Then  'ѡ��
shiyongjiangling = chengjianghao(zuoyoujiangling) 'zuoyoujiangling
For f = 0 To 13
Label3(f).Visible = False
Next
Image2(0).Visible = True
Image2(1).Visible = True
Label1.Visible = True
Label1.Caption = jming(shiyongjiangling)
End If
If Image2(2).Visible = True Then  '�տ�ʼѡ�˿�ʼ
If KeyCode = vbKey1 Then
an1 = 1
wang = 1
kongzhizhe(0, 7) = 1
kongzhizhe(1, 6) = 1
kongzhizhe(2, 5) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode = vbKey2 Then
an1 = 2
wang = 19
kongzhizhe(6, 6) = 1
kongzhizhe(4, 6) = 1
kongzhizhe(7, 6) = 1
kongzhizhe(8, 7) = 1
kongzhizhe(6, 5) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode = vbKey3 Then
an1 = 3
wang = 59
kongzhizhe(10, 7) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode = vbKey4 Then
an1 = 4
wang = 37
kongzhizhe(9, 6) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode = vbKey5 Then
an1 = 5
wang = 62
kongzhizhe(3, 5) = 1
kongzhizhe(4, 5) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode = vbKey6 Then
an1 = 6
wang = 7
kongzhizhe(7, 5) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode = vbKey7 Then
an1 = 7
wang = 72
kongzhizhe(8, 4) = 1
kongzhizhe(8, 5) = 1
kongzhizhe(9, 4) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode = vbKey8 Then
an1 = 8
wang = 67
kongzhizhe(2, 4) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode = vbKey9 Then
an1 = 9
wang = 39
kongzhizhe(7, 4) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode = vbKey0 Then
an1 = 10
wang = 47
kongzhizhe(2, 1) = 1
kongzhizhe(3, 2) = 1
kongzhizhe(1, 2) = 1
kongzhizhe(1, 3) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode = vbKeyF1 Then
an1 = 11
wang = 16
kongzhizhe(5, 3) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode = vbKeyF2 Then
an1 = 12
wang = 55
kongzhizhe(6, 3) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode = vbKeyF3 Then
an1 = 13
wang = 77
kongzhizhe(8, 3) = 1
kongzhizhe(9, 3) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode = vbKeyF4 Then
an1 = 14
wang = 29
kongzhizhe(5, 2) = 1
kongzhizhe(6, 2) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode = vbKeyF5 Then
an1 = 15
wang = 12
kongzhizhe(6, 1) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode = vbKeyF6 Then
an1 = 16
wang = 64
kongzhizhe(8, 0) = 1
Image2(2).Visible = False
Image1(0).Visible = True
xuanrentu = False
End If
If KeyCode <> vbKeyF10 Then
jbingli(wang) = 150
End If
End If 'xuanrentu   �տ�ʼѡ�˽���
If Image1(0).Visible = True Then  'ԭʼ��ͼ��������
If KeyCode = vbKeyUp And hongy < 7 Then

hongy = hongy + 1
Image1(0).Top = Image1(0).Top - 1000
Label2(0).Caption = diming(hongx, hongy)
xianshichengchi = dizhi(hongx, hongy)
����
End If
If KeyCode = vbKeyDown And hongy > 0 Then

hongy = hongy - 1
Image1(0).Top = Image1(0).Top + 1000
Label2(0).Caption = diming(hongx, hongy)
xianshichengchi = dizhi(hongx, hongy)
����
End If
If KeyCode = vbKeyLeft And hongx > 0 Then

hongx = hongx - 1
Image1(0).Left = Image1(0).Left - 1125
Label2(0).Caption = diming(hongx, hongy)
xianshichengchi = dizhi(hongx, hongy)
����
End If
If KeyCode = vbKeyRight And hongx < 10 Then

hongx = hongx + 1
Image1(0).Left = Image1(0).Left + 1125
Label2(0).Caption = diming(hongx, hongy)
xianshichengchi = dizhi(hongx, hongy)
����
End If
End If 'chuzhengmian=f
If Image2(1).Visible = True Then     '��������������
If KeyCode = vbKeyUp And lvy < 4 Then
Image2(0).Top = Image2(0).Top - 1800
lvy = lvy + 1
End If
If KeyCode = vbKeyDown And lvy > 1 Then
Image2(0).Top = Image2(0).Top + 1800
lvy = lvy - 1
End If
If KeyCode = vbKeyLeft And lvx > 1 Then
Image2(0).Left = Image2(0).Left - 2700
lvx = lvx - 1
End If
If KeyCode = vbKeyRight And lvx < 4 Then
Image2(0).Left = Image2(0).Left + 2700
lvx = lvx + 1
End If
End If 'chuzhengmian=t
If Label3(0).Visible = True And KeyCode = vbKeyRight Then '�ҿ�����
If zuoyoujiangling + 1 <= m Then
zuoyoujiangling = zuoyoujiangling + 1
����
For f = 1 To 84
����(f).Visible = False '��ʧͼ
Next
With ����(chengjianghao(zuoyoujiangling))
.Visible = True
.Left = 3000
.Top = 1000
End With
'��ͼ.Enabled = True
End If


'''''
End If
If Label3(0).Visible = True And KeyCode = vbKeyLeft Then '�󿴽���
If zuoyoujiangling - 1 > 0 Then
zuoyoujiangling = zuoyoujiangling - 1
����
For f = 1 To 84
����(f).Visible = False '��ʧͼ
Next
With ����(chengjianghao(zuoyoujiangling))
.Visible = True
.Left = 3000
.Top = 1000
End With
'��ͼ.Enabled = True
End If


'''''
End If
If Image2(0).Visible = True And Image2(1).Visible = True Then '                            ���������
If KeyCode = vbKeyReturn And lvx = 1 And lvy = 4 Then  '��ѡ����
                                                        '����:��²���ã���²��æ�������
m = 0
Image2(0).Visible = False
Image2(1).Visible = False
Label1.Visible = False
For f = 1 To 84
If jzai(f) = xianshichengchi Then
m = m + 1
chengjianghao(m) = f
End If
Next
If m = 0 Then
For f = 1 To 84
chengjianghao(f) = 0
Next
End If
Label3(0).Caption = "(���Ҽ����������ϼ���ѡ��)���ǽ�������" & m
Label3(1) = "����:" & jming(chengjianghao(1))
 If kongxian(chengjianghao(1)) = False Then
 Label3(2) = "��æ"
 End If
If jshenfen(chengjianghao(1)) = 1 Then
Label3(3) = "�ϴ�"
Else
If jshenfen(chengjianghao(1)) = -1 Then
Label3(3) = "��²"
End If
End If
Label3(4) = "����:" & jming(jwang(chengjianghao(1)))
For f = 1 To 38
For fx = 0 To 10
For fy = 0 To 7
If dizhi(fx, fy) = jzai(chengjianghao(1)) Then
Label3(5) = "����:" & diming(fx, fy)
End If
Next
Next
Next
Label3(6) = "�ȼ���" & jji(chengjianghao(1))
Label3(7) = "������" & jwuli(chengjianghao(1))
Label3(8) = "������" & jzhili(chengjianghao(1))
Label3(9) = "�ҳ϶ȣ�" & jzhong(chengjianghao(1))
Label3(10) = "���飺" & jjing(chengjianghao(1))
Label3(11) = "������" & jtili(chengjianghao(1))
If jbingzhong(chengjianghao(1)) = 1 Then
Label3(12) = "ǹ��"
End If
If jbingzhong(chengjianghao(1)) = 2 Then
Label3(12) = "����"
End If
If jbingzhong(chengjianghao(1)) = 3 Then
Label3(12) = "���"
End If
If jbingzhong(chengjianghao(1)) = 4 Then
Label3(12) = "ˮ��"
End If
Label3(13) = jbingli(chengjianghao(1))
For f = 0 To 13
Label3(f).Visible = True
Next
zuoyoujiangling = 1
����
For f = 1 To 84
����(f).Visible = False '��ʧͼ
Next
With ����(chengjianghao(zuoyoujiangling))
.Visible = True
.Left = 3000
.Top = 1000
End With
'��ͼ.Enabled = True
End If

If KeyCode = vbKeyReturn And lvx = 2 And lvy = 4 And kongxian(shiyongjiangling) = True Then   '����ũ��
sbnongye(xianshichengchi) = sbnongye(xianshichengchi) + (jzhili(shiyongjiangling) + jji(shiyongjiangling) * 3) * 30
sbjinqian(xianshichengchi) = sbjinqian(xianshichengchi) - kaitianfei
Label2(2).Caption = "ũҵ��" & sbnongye(xianshichengchi)
Label2(6).Caption = "��Ǯ:" & sbjinqian(xianshichengchi)
Label1.Caption = ""
kongxian(shiyongjiangling) = False
shiyongjiangling = 0
End If '''
If KeyCode = vbKeyReturn And lvx = 3 And lvy = 4 And kongxian(shiyongjiangling) = True Then '������ҵ
Image2(0).Visible = False
Image2(1).Visible = False
Label1.Visible = False
sbshangye(xianshichengchi) = sbshangye(xianshichengchi) + (jzhili(shiyongjiangling) + jji(shiyongjiangling) * 3) * 30
sbjinqian(xianshichengchi) = sbjinqian(xianshichengchi) - kaishangfei
Label2(3).Caption = "��ҵ��" & sbshangye(xianshichengchi)
Label2(6).Caption = "��Ǯ:" & sbjinqian(xianshichengchi)
Label1.Caption = ""
kongxian(shiyongjiangling) = False
shiyongjiangling = 0
Image2(0).Visible = True
Image2(1).Visible = True
Label1.Visible = True
End If ''
If KeyCode = vbKeyReturn And lvx = 4 And lvy = 4 And kongxian(shiyongjiangling) = True Then   '��������
sbminzhong(xianshichengchi) = sbminzhong(xianshichengchi) + jji(shiyongjiangling)
If sbminzhong(xianshichengchi) > 100 Then
sbminzhong(xianshichengchi) = 100
End If
Label2(4).Caption = "���ң�" & sbminzhong(xianshichengchi)
Label1.Caption = ""
kongxian(shiyongjiangling) = False
shiyongjiangling = 0
End If '''
If KeyCode = vbKeyReturn And lvx = 3 And lvy = 3 And kongxian(shiyongjiangling) = True Then   '����
If sbjinqian(xianshichengchi) < sbliangshi(xianshichengchi) Then
sbjinqian(xianshichengchi) = sbjinqian(xianshichengchi) + (sbliangshi(xianshichengchi) / 2) / 2
sbliangshi(xianshichengchi) = sbliangshi(xianshichengchi) / 2
Label2(7).Caption = "��ʳ��" & sbliangshi(xianshichengchi)
Label2(6).Caption = "��Ǯ:" & sbjinqian(xianshichengchi)
Label1.Caption = ""
kongxian(shiyongjiangling) = False
shiyongjiangling = 0
Else
sbliangshi(xianshichengchi) = sbliangshi(xianshichengchi) + (sbjinqian(xianshichengchi) / 2) / 2
sbjinqian(xianshichengchi) = sbjinqian(xianshichengchi) / 2
Label2(7).Caption = "��ʳ��" & sbliangshi(xianshichengchi)
Label2(6).Caption = "��Ǯ:" & sbjinqian(xianshichengchi)
Label1.Caption = ""
kongxian(shiyongjiangling) = False
shiyongjiangling = 0
End If
End If '''
If KeyCode = vbKeyReturn And lvx = 4 And lvy = 3 And kongxian(shiyongjiangling) = True Then   '���뽫��
'inc (jzhong(shiyongjiangling))
jzhong(shiyongjiangling) = jzhong(shiyongjiangling) + (jzhili(wang) + jwuli(wang)) / 3 + 1
jtili(shiyongjiangling) = 100
sbjinqian(xianshichengchi) = sbjinqian(xianshichengchi) - 600
Label2(6).Caption = "��Ǯ:" & sbjinqian(xianshichengchi)
Label1.Caption = ""
kongxian(shiyongjiangling) = False
shiyongjiangling = 0
End If '''

If KeyCode = vbKeyReturn And lvx = 2 And lvy = 2 And kongxian(shiyongjiangling) = True Then   '�ƶ�����
Image5.Visible = True
Image5.Top = 0
Image5.Left = 4000
Text1.Visible = True
Command1.Visible = True
Command2.Visible = True
Image2(0).Visible = False
Image2(1).Visible = False
End If '''
If KeyCode = vbKeyReturn And lvx = 4 And lvy = 2 And kongxian(shiyongjiangling) = True And sbjinqian(xianshichengchi) > 0 Then '�б�
If jji(72) > 2 And jming(72) = "����" Then
tishi.Visible = True
tishi.Caption = "�����������������õ���������" '�����ߣ�ͼ������ۣ�����(����)
jming(72) = "������"
sbsuoshu(7) = 72
kongzhizhe(chengx(7), chengy(7)) = 1
Dim fd As Long
For fd = 1 To 84
If jzai(fd) = 7 Then
jshenfen(fd) = 0
jwang(fd) = 72
End If
Next
'ͼ
End If
sbhoubeibingli(xianshichengchi) = sbhoubeibingli(xianshichengchi) + jwuli(shiyongjiangling) * sbminzhong(xianshichengchi) _
* sbrenkou(xianshichengchi) / 80000 + (jwuli(shiyongjiangling) - 10) * 200
Label2(8).Caption = "δ��������" & sbhoubeibingli(xianshichengchi)
sbjinqian(xianshichengchi) = sbjinqian(xianshichengchi) - 800
jjing(shiyongjiangling) = jjing(shiyongjiangling) + 1
Label1.Caption = ""
kongxian(shiyongjiangling) = False
shiyongjiangling = 0
����
����
End If '''

If KeyCode = vbKeyReturn And lvx = 1 And lvy = 1 And kongxian(shiyongjiangling) = True Then   '�������
Image4.Visible = True '��������̳�
Image2(0).Visible = False
Image2(1).Visible = False
Label1.Visible = False
'�佫������ʾ
Image4.Visible = True

End If '''

If KeyCode = vbKeyReturn And lvx = 1 And lvy = 3 And kongxian(shiyongjiangling) = True Then   'Ȱ����²--�ɹ�
Dim fff As Long
Dim yici As Long
yici = 0
tishi.Visible = True
tishi.Caption = "�˴�û�з�²"
For fff = 1 To 84
If jzai(fff) = jzai(shiyongjiangling) And jshenfen(fff) = -1 And yici = 0 Then

yici = 1
tishi.Caption = "����Ȱ����������δ��50��"
jzhong(fff) = jzhong(fff) - jji(shiyongjiangling) * 2 - jzhili(shiyongjiangling) + 20 - jwuli(shiyongjiangling)

If jzhong(fff) < 50 Then
jshenfen(fff) = 0
inc jjing(shiyongjiangling)
jwang(fff) = jwang(shiyongjiangling)
jzhong(fff) = 60
tishi.Caption = "��Ȱ�������"
End If
Label1.Caption = ""
kongxian(shiyongjiangling) = False
shiyongjiangling = 0
End If
Next

End If '''
If KeyCode = vbKeyReturn And lvx = 2 And lvy = 1 And kongxian(shiyongjiangling) = True Then   '�Ӷ�
sbliangshi(xianshichengchi) = sbliangshi(xianshichengchi) + 10000
sbjinqian(xianshichengchi) = sbjinqian(xianshichengchi) + 10000
sbminzhong(xianshichengchi) = sbminzhong(xianshichengchi) - 10
����
Label1.Caption = ""
kongxian(shiyongjiangling) = False
shiyongjiangling = 0
End If '''
If KeyCode = vbKeyReturn And lvx = 3 And lvy = 2 And kongxian(shiyongjiangling) = True Then   'ί��
tishi.Visible = True
If kongzhizhe(chengx(xianshichengchi), chengy(xianshichengchi)) = 1 Then
kongzhizhe(chengx(xianshichengchi), chengy(xianshichengchi)) = 0
tishi.Caption = "�˳ǳ��Ѿ��������Թ���"
End If


End If '''
If KeyCode = vbKeyReturn And lvx = 4 And lvy = 1 And kongxian(shiyongjiangling) = True And jtili(shiyongjiangling) > 0 Then 'Ȱ���н�

tishi.Visible = True
tishi.Caption = "�Ѿ�ȥ������˵�����ؽ��춼ϣ��" & jming(shiyongjiangling) & "����   "
jtili(shiyongjiangling) = jtili(shiyongjiangling) - 16 + jzhili(shiyongjiangling)
Dim quan As Long
For quan = 1 To 84
If jwang(quan) <> jwang(shiyongjiangling) And jshenfen(quan) = 0 And jshenfen(shiyongjiangling) <> -1 And way(jzai(shiyongjiangling), jzai(quan)) = True Then
If jbingli(quan) < 1000 Then
jzhong(quan) = jzhong(quan) - jzhili(shiyongjiangling) + 7
tishi.Caption = tishi.Caption & jming(quan) & "�ҳ��½�- "
sbjinqian(jzai(shiyongjiangling)) = sbjinqian(jzai(shiyongjiangling)) - 2000
End If
If jzhong(quan) < 30 Then
jwang(quan) = jwang(shiyongjiangling)
jzai(quan) = jzai(shiyongjiangling)
jzhong(quan) = 51
tishi.Caption = tishi.Caption & "�н���" & jming(quan) & "//  "
End If
End If
Next
Label1.Caption = ""
kongxian(shiyongjiangling) = False
shiyongjiangling = 0
End If '''
If KeyCode = vbKeyReturn And lvx = 1 And lvy = 2 And kongxian(shiyongjiangling) = True Then   '�̵�
tishi.Visible = True
tishi.Caption = "�ȼ�����������δ����"
End If '''
If KeyCode = vbKeyReturn And lvx = 2 And lvy = 3 And kongxian(shiyongjiangling) = True Then   'ն��
Dim df As Long
Dim dyici As Long
dyici = 0
tishi.Visible = True
tishi.Caption = "�޷�²"
For df = 1 To 84
If jzai(df) = jzai(shiyongjiangling) And jshenfen(df) = -1 And dyici = 0 Then
dyici = 1
'si
jzai(df) = 0
tishi.Visible = True
tishi.Caption = jming(df) & "�Ѿ��� //    "
'huo
Dim gu As Long
gu = 0
If gu = 0 Then
'If sbsuoshu(fuhuojzai(fuhuojf + 1)) <> 0 And fuhuojzai(fuhuojf + 1) < 38 Then
'inc fuhuojzai(fuhuojf + 1)
'End If
fuhuojf = fuhuojf + 1
If fuhuojf > 9000 Then
fuhuojf = 1000
End If
jming(df) = fuhuojming(fuhuojf)
jshenfen(df) = 0
jzai(df) = fuhuojzai(fuhuojf) ''''''�Ѿ����
jwuli(df) = fuhuojwuli(fuhuojf)
jzhili(df) = fuhuojzhili(fuhuojf)
fuhuojhao(fuhuojf) = df
jbingzhong(df) = fuhuojbingzhong(fuhuojf)
jwang(df) = sbsuoshu(fuhuojzai(fuhuojf)) '''''�Լ���Ϊ���ҷ�²���˵���=�Ѿ����
����(df).Picture = �հ�ͼ.Picture
tishi.Visible = True
tishi.Caption = tishi.Caption & fuhuojf & sbchengming(fuhuojzai(fuhuojf)) & jming(sbsuoshu(fuhuojzai(fuhuojf))) & "�����½�" & jming(df)
jbingli(df) = 100
jjing(df) = 0
jtili(df) = 70
jji(df) = 1
jzhong(df) = 89
'tu
End If
jzhong(shiyongjiangling) = jzhong(shiyongjiangling) - 5
Label1.Caption = ""
kongxian(shiyongjiangling) = False
shiyongjiangling = 0
End If
Next
End If '''
If KeyCode = vbKeyReturn And lvx = 3 And lvy = 1 And kongxian(shiyongjiangling) = True And jbingli(shiyongjiangling) > -1 Then '����=
Text2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Image2(0).Visible = False
Image2(1).Visible = False
Label1.Visible = False
ս����ʽ.Visible = True
ս����ʽ.Left = 0
ս����ʽ.Top = -50
Image5.Visible = True
Image5.Top = 0
Image5.Left = 4000

End If '''
End If       '���������
If KeyCode = vbKeyReturn And Image1(0).Visible = True And sbsuoshu(dizhi(hongx, hongy)) = wang Then  ' '����Լ��ǳ�---Ӧ�÷��ڳ���ǰ������,���и�����
Label1.Caption = ""
Image2(0).Top = 800 '''
Image2(0).Left = 2600 '''
Image2(0).Visible = True
Image2(1).Top = 600 '''
Image2(1).Left = 1600 '''
Image2(1).Visible = True
lvx = 1
lvy = 4
Image1(26).Visible = False '''
Image1(0).Visible = False '''
For f = 1 To 38
Image3(f).Visible = False
Next

chuzhengmian = True '''



Label1.Visible = True
Label1.Left = Image2(1).Left + 200
Label1.Top = Image2(1).Top + 1200
End If '����Լ��ǳؽ���
If KeyCode = vbKeyReturn And kongzhizhe(hongx, hongy) <> 1 And Image1(0).Visible = True And sbsuoshu(dizhi(hongx, hongy)) <> wang Then '������˳ǳ�
Image2(3).Top = 600 '''
Image2(3).Left = 1600 '''
Image2(3).Visible = True
chuzhengmian2 = True '''
Image1(0).Visible = False '''



End If '������˳ǳؽ���
If KeyCode = vbKeyF1 And Label4.Visible = True Then '���ز��Խ���
Image1(0).Visible = True
Image1(26).Visible = True
Label4.Visible = False
End If
If Label4.Visible = True And KeyCode = vbKeyReturn Then '��������
Dim f89 As Long
f89 = 0

For f89 = 1 To 84
If jbingli(f89) > 50000 Then
jbingli(f89) = 50000
End If
Next

f89 = 0
For f89 = 1 To 38
If sbminzhong(f89) < 0 Then 'for f89=1to84--f89=40-end
sbminzhong(f89) = 0
End If
Next
yici1 = 0 '�з���
huihe = huihe + 1
List1.AddItem "�غ�" & huihe & "//"
'���黯Ϊ�ȼ�
For f = 1 To 84
If jjing(f) > 30 + (jji(f) - 1) * 25 Then
jjing(f) = jjing(f) - (30 + (jji(f) - 1) * 25)
jji(f) = jji(f) + 1
End If
Next

���켼��ʹ��
��ʳû��
wanjiabeida = False
'����й�����
Image1(0).Visible = True
Image1(26).Visible = True
Label4.Visible = False
For f = 1 To 84
kongxian(f) = True
Next
 For f = 1 To 38
 sbjinqian(f) = sbjinqian(f) + sbshangye(f)
 For fx = 1 To 84
 If jzai(fx) = f And jshenfen(fx) <> -1 Then 'If jwang(fx) = sbsuoshu(f) Then
 f2 = f2 + jbingli(fx)
 End If
 Next
 sbliangshi(f) = sbliangshi(f) + sbnongye(f) - sbhoubeibingli(f) - f2
 f2 = 0
 Next
 Dim shizhuang As Long '�˴�����f�����ºþ�
 For shizhuang = 1 To 84
 If kongxian(shizhuang) = True And jshenfen(shizhuang) <> -1 And kongzhizhe(chengx(jzai(shizhuang)), chengy(jzai(shizhuang))) = 0 Then 'jwang(shizhuang) <> wang Then '������������
 diannaodong shizhuang
 End If
 Next
For f = 1 To 38
sbrenkou(f) = sbrenkou(f) - 800 + sbminzhong(f) * 10 + sbrenkou(f) / 10000
Next

If youxijiandanhua = True Then '�ڶ���
For f89 = 1 To 84
If jbingli(f89) > 5000 Then
'jbingli(f89) = 5000
End If
Next
End If
����
����
End If

If KeyCode = vbKeyP Then '���ؼ�

If Image1(26).Visible = True And Image2(2).Visible = False And Image2(3).Visible = False And Image4.Visible = False Then '�򿪾��߽���
Image1(0).Visible = False
Image1(26).Visible = False
Label4.Visible = True
End If
If Image4.Visible = True Then '����
Image1(26).Visible = True
Image1(0).Visible = True
Image4.Visible = False
For f = 1 To 38
Image3(f).Visible = True '��֪Ϊɶ��
Next
End If

If Label3(0).Visible = True Then '����ѡ��ͼ
For f = 0 To 13
Label3(f).Visible = False
Next
Image2(0).Visible = True
Image2(1).Visible = True
Label1.Visible = True
End If

If chuzhengmian = True And Image2(0).Visible = True Then
'begin
Label1.Visible = False
chuzhengmian = False
Image1(26).Visible = True 'image1(26)shiditu
For f = 1 To 38
Image3(f).Visible = True
Next
Image2(1).Visible = False
Image2(0).Visible = False
Image1(0).Visible = True
'end
End If 'chuzhengmian
If chuzhengmian2 = True And Image2(3).Visible = True Then
Image2(3).Visible = False
chuzhengmian2 = False '''
Image1(0).Visible = True '''
End If 'chuzhengmian2
End If 'vbkeyP
'������������������
If KeyCode = vbKeyP Or KeyCode = vbKeyUp Then '��ʧ����
For f = 1 To 84
����(f).Visible = False '��ʧͼ
Next
tishi.Visible = False
tishi.Caption = ""
End If
End Sub
Private Sub Form_Load()
weiren = 0
youxijiandanhua = False
Label6(1).Top = 1200
Label6(1).Left = 2520
Label6(2).Top = 360
Label6(2).Left = 1440
Label6(3).Top = 1200
Label6(3).Left = 1440
Label6(4).Top = 2040
Label6(4).Left = 1440
Label6(5).Top = 1200
Label6(5).Left = 360
Label6(6).Top = 1200
Label6(6).Left = 8400
Label6(7).Top = 360
Label6(7).Left = 9480
Label6(8).Top = 1200
Label6(8).Left = 9480
Label6(9).Top = 2040
Label6(9).Left = 9480
Label6(10).Top = 1200
Label6(10).Left = 10680
Label6(0).Top = 6120
Label6(0).Left = 720
Label7.Top = 5400
Label7.Left = 6120
Label8.Top = 6120
Label8.Left = 6120
Label9.Top = 6120
Label9.Left = 5040
Label10.Top = 6120
Label10.Left = 7200
Label11.Top = 5400
Label11.Left = 8280
Label12.Top = 5400
Label12.Left = 9240
Label13.Top = 6120
Label13.Left = 8280
Label14.Top = 6120
Label14.Left = 9240
����6hei.Enabled = False
Label15.Top = 5520 '����6form����
Label15.Left = 0
     '����ǳ���̺�Ƴ�&H0&��
     �� = &HFF&
     �� = &HFFFF&
     �� = &H80FF&
     �� = &HFF00&
     ���� = &HFF0000
     ǳ�� = &HFFFF00
     
     For f = 1 To 9999
fuhuojming(f) = "��" & f
fuhuojzai(f) = 20 '�ٴ�
fuhuojwuli(f) = 11
fuhuojzhili(f) = 11
fuhuojbingzhong(f) = 1
Next
fuhuojf = 0
'�½�
fuhuojming(2) = "�����"
fuhuojzai(2) = 16 '���
fuhuojwuli(2) = 10
fuhuojzhili(2) = 16
fuhuojbingzhong(2) = 2
fuhuojming(3) = "��ͳ"
fuhuojzai(3) = 34 '���
fuhuojwuli(3) = 5
fuhuojzhili(3) = 17
fuhuojbingzhong(3) = 3
fuhuojming(4) = "С��"
fuhuojzai(4) = 34 '���
fuhuojwuli(4) = 5
fuhuojzhili(4) = 13
fuhuojbingzhong(4) = 1
fuhuojming(5) = "����"
fuhuojzai(5) = 3
fuhuojwuli(5) = 10
fuhuojzhili(5) = 14
fuhuojbingzhong(5) = 2
fuhuojming(6) = "��ά"
fuhuojzai(6) = 2
fuhuojwuli(6) = 13
fuhuojzhili(6) = 14
fuhuojbingzhong(6) = 3
'''''

Image4.Top = 1200
Image4.Left = 1500
tishi.Top = 2760
tishi.Left = 2520
Text1.Top = 6600
Text1.Left = 720
Command1.Top = 7080
Command1.Left = 720
Command2.Top = 7080
Command2.Left = 2280
List1.Top = 9120
List1.Left = 11520
Label4.Top = 1080
Label4.Left = 2760
For f = 1 To 38 'ʧ�ܴ��뿪ʼ
For fx = 0 To 10
For fy = 0 To 7
If dizhi(fx, fy) = f Then
kongzhizheij(f) = kongzhizhe(fx, fy)
End If
Next
Next
Next
''''''''''
'������2-1ת��


For fx = 0 To 10
For fy = 0 To 7
For f = 1 To 38
If chengx(f) = fx And chengy(f) = fy Then
kongzhizheij(f) = kongzhizhe(fx, fy)
End If
Next
Next
Next
'2-1ת������'ʧ�ܴ������
'kongzhizheij(2) = 1 '��֤2-1----Ϊʧ��
tishi.Visible = False
tishit.Enabled = False
Text2.Visible = False
Command3.Visible = False
Command4.Visible = False
ս����ʽ.Visible = False
way(1, 3) = True '''''''begin way
way(2, 5) = True
way(3, 4) = True ''
way(3, 10) = True
way(4, 7) = True
way(5, 8) = True
way(6, 12) = True
way(6, 9) = True ''
way(6, 10) = True
way(7, 16) = True
way(7, 8) = True
way(8, 11) = True
way(9, 12) = True ''
way(9, 15) = True
way(9, 10) = True
way(10, 15) = True
way(11, 13) = True
way(11, 14) = True ''
way(12, 17) = True
way(13, 16) = True
way(13, 21) = True
way(13, 14) = True
way(14, 22) = True ''
way(14, 23) = True
way(15, 18) = True
way(15, 19) = True
way(15, 16) = True
way(16, 20) = True ''
way(16, 21) = True
way(17, 18) = True
way(18, 29) = True
way(18, 19) = True
way(19, 24) = True ''
way(19, 20) = True
way(20, 24) = True
way(20, 30) = True
way(20, 25) = True
way(20, 21) = True ''
way(21, 25) = True
way(21, 22) = True
way(22, 26) = True
way(22, 23) = True
way(27, 23) = True ''
way(24, 29) = True
way(24, 30) = True
way(25, 31) = True
way(25, 26) = True
way(26, 32) = True ''
way(26, 27) = True
way(27, 33) = True
way(28, 29) = True
way(29, 34) = True
way(30, 35) = True ''
way(30, 31) = True
way(31, 36) = True
way(31, 32) = True
way(32, 37) = True
way(33, 38) = True ''
way(34, 35) = True
way(35, 36) = True
way(36, 37) = True
way(37, 38) = True
way(22, 25) = True '+
For fx = 1 To 38
For fy = 1 To 38
If way(fx, fy) = True Then
way(fy, fx) = True
End If
Next
Next 'end way

dizhiij(1) = "����"
dizhiij(2) = "����"
dizhiij(3) = "�ɶ�"
dizhiij(4) = "����"
dizhiij(5) = "����"
dizhiij(6) = "�Ϳ�"
dizhiij(7) = "����"
dizhiij(8) = "��ˮ"
dizhiij(9) = "����"
dizhiij(10) = "����"
dizhiij(11) = "����"
dizhiij(12) = "����"
dizhiij(13) = "����"
dizhiij(14) = "����"
dizhiij(15) = "����"
dizhiij(16) = "���"
dizhiij(17) = "����"
dizhiij(18) = "��ɳ"
dizhiij(19) = "����"
dizhiij(20) = "�ٴ�"
dizhiij(21) = "����"
dizhiij(22) = "��"
dizhiij(23) = "ƽԭ"
dizhiij(24) = "®��"
dizhiij(25) = "���"
dizhiij(26) = "���"
dizhiij(27) = "��Ƥ"
dizhiij(28) = "����"
dizhiij(29) = "��ɣ"
dizhiij(30) = "��ҵ"
dizhiij(31) = "С��"
dizhiij(32) = "����"
dizhiij(33) = "��ƽ"
dizhiij(34) = "���"
dizhiij(35) = "��"
dizhiij(36) = "����"
dizhiij(37) = "����"
dizhiij(38) = "����"

'�ƶ������õ������ؼ�
Text1.Visible = False
Command1.Visible = False
Command2.Visible = False
'����ѿ��̷�
kaitianfei = 1000
kaishangfei = 1000
'ԭʼ�ǳ�
For f = 1 To 38
sbnongye(f) = 1000
sbshangye(f) = 1000
sbminzhong(f) = 75
sbrenkou(f) = 60000
sbjinqian(f) = 10000
sbliangshi(f) = 10000
sbhoubeibingli(f) = 1000
Next
'�������
For f = 1 To 84
kongxian(f) = True '���н���2
jji(f) = 1 '6
jjing(f) = 0 '10
jtili(f) = 100 '11
jbingli(f) = 100 ' As Long'13
jzhong(f) = 65 '9
Next

 jming(1) = "����" '1
'chengjianghao(100) As Long '0
'shiyongjiangling As Long
jshenfen(1) = 1 '1->wang,-1->fulu       3
jwang(1) = 1 '4
jzai(1) = 8 '5
jwuli(1) = 11 '7
'zuoyoujiangling As Long
jzhili(1) = 10 '8
jzhong(1) = 100 '9
jbingzhong(1) = 3 ' As Long '12
 jming(2) = "��" '1
'jshenfen(2) = 0 '1->wang,-1->fulu       3
jwang(2) = 1 '4
jzai(2) = 5 '5
jwuli(2) = 12 '7
jzhili(2) = 10 '8
jbingzhong(2) = 3 ' As Long '12
'''''''
 jming(3) = "�ӵ�" '1
jwang(3) = 1 '4
jzai(3) = 8 '5
jwuli(3) = 12 '7
jzhili(3) = 10 '8
jbingzhong(3) = 2 ' As Long '12
 jming(4) = "����" '1
jwang(4) = 1 '4
jzai(4) = 2 '5
jwuli(4) = 10 '7
jzhili(4) = 10 '8
jbingzhong(4) = 3 ' As Long '12
 jming(5) = "���" '1
jwang(5) = 1 '4
jzai(5) = 8 '5
jwuli(5) = 10 '7
jzhili(5) = 10 '8
jbingzhong(5) = 3 ' As Long '12
 jming(6) = "�ź�" '1
jwang(6) = 1 '4
jzai(6) = 8 '5
jwuli(6) = 11 '7
jzhili(6) = 10 '8
jbingzhong(6) = 1 ' As Long '12
jming(7) = "����" '1
jwang(7) = 7 '4
jzai(7) = 26 '5
jwuli(7) = 14 '7
jzhili(7) = 10 '8
jbingzhong(7) = 3 ' As Long '12
 jming(8) = "����" '1
jwang(8) = 7 '4
jzai(8) = 26 '5
jwuli(8) = 12 '7
jzhili(8) = 11 '8
jbingzhong(8) = 3 ' As Long '12
jming(9) = "갰�" '1
jwang(9) = 7 '4
jzai(9) = 26 '5
jwuli(9) = 11 '7
jzhili(9) = 10 '8
jbingzhong(9) = 1 ' As Long '12
jming(10) = "κ��" '1
jwang(10) = 7 '4
jzai(10) = 26 '5
jwuli(10) = 11 '7
jzhili(10) = 10 '8
jbingzhong(10) = 2 ' As Long '12
jming(11) = "�¹�" '1
jwang(11) = 7 '4
jzai(11) = 26 '5
jwuli(11) = 10 '7
jzhili(11) = 12 '8
jbingzhong(11) = 2 ' As Long '12
jming(12) = "����" '1
jwang(12) = 12 '4
jzai(12) = 18 '5
jwuli(12) = 10 '7
jzhili(12) = 10 '8
jbingzhong(12) = 4 ' As Long '12
jming(13) = "����" '1
jwang(13) = 12 '4
jzai(13) = 18 '5
jwuli(13) = 10 '7
jzhili(13) = 10 '8
jbingzhong(13) = 2 ' As Long '12
jming(14) = "κ��" '1
jwang(14) = 12 '4
jzai(14) = 18 '5
jwuli(14) = 12 '7
jzhili(14) = 10 '8
jbingzhong(14) = 3 ' As Long '12
jming(15) = "����" '1
jwang(15) = 12 '4
jzai(15) = 18 '5
jwuli(15) = 10 '7
jzhili(15) = 10 '8
jbingzhong(15) = 1 ' As Long '12
jming(16) = "����" '1
jwang(16) = 16 '4
jzai(16) = 16 '5
jwuli(16) = 11 '7
jzhili(16) = 10 '8
jbingzhong(16) = 1 ' As Long '12
jm = 17
jming(jm) = "������" '1
jwang(jm) = 16 '4
jzai(jm) = 16 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 2 ' As Long '12
jm = 18
jming(jm) = "��ڼ" '1
jwang(jm) = 16 '4
jzai(jm) = 16 '5
jwuli(jm) = 10 '7
jzhili(jm) = 12 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 19
jming(jm) = "Ԭ��" '1
jwang(jm) = 19 '4
jzai(jm) = 22 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 20
jming(jm) = "����" '1
jwang(jm) = 19 '4
jzai(jm) = 22 '5
jwuli(jm) = 12 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 21
jming(jm) = "�ĳ�" '1
jwang(jm) = 19 '4
jzai(jm) = 22 '5
jwuli(jm) = 12 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 22
jming(jm) = "���" '1
jwang(jm) = 19 '4
jzai(jm) = 22 '5
jwuli(jm) = 10 '7
jzhili(jm) = 12 '8
jbingzhong(jm) = 2 ' As Long '12
jm = 23
jming(jm) = "��׼" '1
jwang(jm) = 19 '4
jzai(jm) = 22 '5
jwuli(jm) = 11 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 24
jming(jm) = "����" '1
jwang(jm) = 19 '4
jzai(jm) = 22 '5
jwuli(jm) = 10 '7
jzhili(jm) = 11 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 25
jming(jm) = "���" '1
jwang(jm) = 19 '4
jzai(jm) = 23 '5
jwuli(jm) = 12 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 26
jming(jm) = "���" '1
jwang(jm) = 19 '4
jzai(jm) = 27 '5
jwuli(jm) = 10 '7
jzhili(jm) = 11 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 27
jming(jm) = "����" '1
jwang(jm) = 19 '4
jzai(jm) = 33 '5
jwuli(jm) = 11 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 28
jming(jm) = "��ͼ" '1
jwang(jm) = 19 '4
jzai(jm) = 14 '5
jwuli(jm) = 10 '7
jzhili(jm) = 11 '8
jbingzhong(jm) = 2 ' As Long '12
jm = 29
jming(jm) = "����" '1
jwang(jm) = 29 '4
jzai(jm) = 15 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 4 ' As Long '12
jm = 30
jming(jm) = "��Խ" '1
jwang(jm) = 29 '4
jzai(jm) = 15 '5
jwuli(jm) = 10 '7
jzhili(jm) = 12 '8
jbingzhong(jm) = 2 ' As Long '12
jm = 31
jming(jm) = "����" '1
jwang(jm) = 29 '4
jzai(jm) = 15 '5
jwuli(jm) = 10 '7
jzhili(jm) = 11 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 32
jming(jm) = "����" '1
jwang(jm) = 29 '4
jzai(jm) = 15 '5
jwuli(jm) = 12 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 33
jming(jm) = "����" '1
jwang(jm) = 29 '4
jzai(jm) = 15 '5
jwuli(jm) = 10 '7
jzhili(jm) = 11 '8
jbingzhong(jm) = 4 ' As Long '12
jm = 34
jming(jm) = "��Ƹ" '1
jwang(jm) = 29 '4
jzai(jm) = 15 '5
jwuli(jm) = 11 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 4 ' As Long '12
jm = 35
jming(jm) = "����" '1
jwang(jm) = 29 '4
jzai(jm) = 15 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 2 ' As Long '12
jm = 36
jming(jm) = "����" '1
jwang(jm) = 29 '4
jzai(jm) = 19 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 37
jming(jm) = "����" '1
jwang(jm) = 37 '4
jzai(jm) = 37 '5
jwuli(jm) = 10 '7
jzhili(jm) = 11 '8
jbingzhong(jm) = 4 ' As Long '12
jm = 38
jming(jm) = "�䰲��" '1
jwang(jm) = 37 '4
jzai(jm) = 37 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 39
jming(jm) = "�ܲ�" '1
jwang(jm) = 39 '4
jzai(jm) = 25 '5
jwuli(jm) = 11 '7
jzhili(jm) = 12 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 40
jming(jm) = "�ĺ��" '1
jwang(jm) = 39 '4
jzai(jm) = 25 '5
jwuli(jm) = 12 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 41
jming(jm) = "�ĺ�Ԩ" '1
jwang(jm) = 39 '4
jzai(jm) = 25 '5
jwuli(jm) = 12 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 42
jming(jm) = "���" '1
jwang(jm) = 39 '4
jzai(jm) = 25 '5
jwuli(jm) = 10 '7
jzhili(jm) = 11 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 43
jming(jm) = "�ܺ�" '1
jwang(jm) = 39 '4
jzai(jm) = 25 '5
jwuli(jm) = 11 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 44
jming(jm) = "�ڽ�" '1
jwang(jm) = 39 '4
jzai(jm) = 25 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 45
jming(jm) = "����" '1
jwang(jm) = 39 '4
jzai(jm) = 25 '5
jwuli(jm) = 10 '7
jzhili(jm) = 12 '8
jbingzhong(jm) = 2 ' As Long '12
jm = 46
jming(jm) = "����" '1
jwang(jm) = 39 '4
jzai(jm) = 25 '5
jwuli(jm) = 10 '7
jzhili(jm) = 12 '8
jbingzhong(jm) = 2 ' As Long '12
jm = 47
jming(jm) = "���" '1
jwang(jm) = 47 '4
jzai(jm) = 3 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 2 ' As Long '12
jm = 48
jming(jm) = "����" '1
jwang(jm) = 47 '4
jzai(jm) = 3 '5
jwuli(jm) = 10 '7
jzhili(jm) = 11 '8
jbingzhong(jm) = 2 ' As Long '12
jm = 49
jming(jm) = "����" '1
jwang(jm) = 47 '4
jzai(jm) = 3 '5
jwuli(jm) = 11 '7
jzhili(jm) = 11 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 50
jming(jm) = "���" '1
jwang(jm) = 47 '4
jzai(jm) = 4 '5
jwuli(jm) = 11 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 4 ' As Long '12
jm = 51
jming(jm) = "����" '1
jwang(jm) = 47 '4
jzai(jm) = 3 '5
jwuli(jm) = 11 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 52
jming(jm) = "����" '1
jwang(jm) = 47 '4
jzai(jm) = 6 '5
jwuli(jm) = 10 '7
jzhili(jm) = 12 '8
jbingzhong(jm) = 4 ' As Long '12
jm = 53
jming(jm) = "��ܲ" '1
jwang(jm) = 47 '4
jzai(jm) = 10 '5
jwuli(jm) = 11 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 54
jming(jm) = "����" '1
jwang(jm) = 47 '4
jzai(jm) = 3 '5
jwuli(jm) = 11 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 2 ' As Long '12
jm = 55
jming(jm) = "Ԭ��" '1
jwang(jm) = 55 '4
jzai(jm) = 20 '5
jwuli(jm) = 10 '7
jzhili(jm) = 11 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 56
jming(jm) = "����" '1
jwang(jm) = 55 '4
jzai(jm) = 20 '5
jwuli(jm) = 10 '7
jzhili(jm) = 11 '8
jbingzhong(jm) = 4 ' As Long '12
jm = 57
jming(jm) = "����" '1
jwang(jm) = 55 '4
jzai(jm) = 20 '5
jwuli(jm) = 11 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 2 ' As Long '12
jm = 58
jming(jm) = "�ױ�" '1
jwang(jm) = 55 '4
jzai(jm) = 20 '5
jwuli(jm) = 11 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 59
jming(jm) = "�����" '1
jwang(jm) = 59 '4
jzai(jm) = 38 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 60
jming(jm) = "���￵" '1
jwang(jm) = 59 '4
jzai(jm) = 38 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 61
jming(jm) = "���﹧" '1
jwang(jm) = 59 '4
jzai(jm) = 38 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 62
jming(jm) = "���" '1
jwang(jm) = 62 '4
jzai(jm) = 11 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 2 ' As Long '12
jm = 63
jming(jm) = "����" '1
jwang(jm) = 62 '4
jzai(jm) = 13 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 4 ' As Long '12
jm = 64
jming(jm) = "����" '1
jwang(jm) = 64 '4
jzai(jm) = 28 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 2 ' As Long '12
jm = 65
jming(jm) = "���" '1
jwang(jm) = 64 '4
jzai(jm) = 28 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 66
jming(jm) = "�ų�" '1
jwang(jm) = 64 '4
jzai(jm) = 28 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 4 ' As Long '12
jm = 67
jming(jm) = "��³" '1
jwang(jm) = 67 '4
jzai(jm) = 7 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 2 ' As Long '12
jm = 68
jming(jm) = "����" '1
jwang(jm) = 67 '4
jzai(jm) = 7 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 69
jming(jm) = "����" '1
jwang(jm) = 67 '4
jzai(jm) = 7 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 70
jming(jm) = "����" '1
jwang(jm) = 67 '4
jzai(jm) = 7 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 4 ' As Long '12
jm = 71
jming(jm) = "����" '1
jwang(jm) = 67 '4
jzai(jm) = 7 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 72
jming(jm) = "����" '1
jwang(jm) = 72 '4
jzai(jm) = 32 '5
jwuli(jm) = 11 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 73
jming(jm) = "����" '1
jwang(jm) = 72 '4
jzai(jm) = 36 '5
jwuli(jm) = 13 '7
jzhili(jm) = 11 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 74
jming(jm) = "�ŷ�" '1
jwang(jm) = 72 '4
jzai(jm) = 32 '5
jwuli(jm) = 13 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 75
jming(jm) = "�¹�" '1
jwang(jm) = 72 '4
jzai(jm) = 32 '5
jwuli(jm) = 10 '7
jzhili(jm) = 11 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 76
jming(jm) = "��Ǭ" '1
jwang(jm) = 72 '4
jzai(jm) = 32 '5
jwuli(jm) = 10 '7
jzhili(jm) = 11 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 77
jming(jm) = "���" '1
jwang(jm) = 77 '4
jzai(jm) = 30 '5
jwuli(jm) = 12 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 1 ' As Long '12
jm = 78
jming(jm) = "̫ʷ��" '1
jwang(jm) = 77 '4
jzai(jm) = 35 '5
jwuli(jm) = 12 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 79
jming(jm) = "����" '1
jwang(jm) = 77 '4
jzai(jm) = 35 '5
jwuli(jm) = 12 '7
jzhili(jm) = 11 '8
jbingzhong(jm) = 4 ' As Long '12
jm = 80
jming(jm) = "�Ƹ�" '1
jwang(jm) = 77 '4
jzai(jm) = 30 '5
jwuli(jm) = 10 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 4 ' As Long '12
jm = 81
jming(jm) = "����" '1
jwang(jm) = 77 '4
jzai(jm) = 30 '5
jwuli(jm) = 12 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 2 ' As Long '12
jm = 82
jming(jm) = "��̩" '1
jwang(jm) = 77 '4
jzai(jm) = 30 '5
jwuli(jm) = 12 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 3 ' As Long '12
jm = 83
jming(jm) = "���" '1
jwang(jm) = 77 '4
jzai(jm) = 30 '5
jwuli(jm) = 10 '7
jzhili(jm) = 12 '8
jbingzhong(jm) = 4 ' As Long '12
jm = 84
jming(jm) = "����" '1
jwang(jm) = 77 '4
jzai(jm) = 30 '5
jwuli(jm) = 11 '7
jzhili(jm) = 10 '8
jbingzhong(jm) = 1 ' As Long '12


For f = 1 To 84
If jwang(f) = f Then '����
jzhong(f) = 100
jshenfen(f) = 1
End If
Next
 '�������


'tishi.Visible = True����
'tishi.Caption = wang16(16)
fx = 0
suoshu(0, 7) = 1 '�����������ǣ�С��
sbsuoshu(2) = 1
suoshu(1, 6) = 1
sbsuoshu(5) = 1
suoshu(2, 5) = 1
sbsuoshu(8) = 1
xiaodui16(1) = 1
suoshu(6, 5) = 19
sbsuoshu(22) = 19
suoshu(6, 6) = 19
sbsuoshu(23) = 19
suoshu(4, 6) = 19
sbsuoshu(14) = 19
suoshu(7, 6) = 19
sbsuoshu(27) = 19
suoshu(8, 7) = 19
sbsuoshu(33) = 19
xiaodui16(2) = 19
suoshu(10, 7) = 59
sbsuoshu(38) = 59
xiaodui16(3) = 59
suoshu(9, 6) = 37
sbsuoshu(37) = 37
xiaodui16(4) = 37
suoshu(3, 5) = 62
sbsuoshu(11) = 62
suoshu(4, 5) = 62
sbsuoshu(13) = 62
xiaodui16(5) = 62
suoshu(7, 5) = 7
sbsuoshu(26) = 7
xiaodui16(6) = 7
suoshu(8, 4) = 72
sbsuoshu(31) = 72
suoshu(8, 5) = 72
sbsuoshu(32) = 72
suoshu(9, 4) = 72
sbsuoshu(36) = 72
xiaodui16(7) = 72
suoshu(2, 4) = 67
sbsuoshu(7) = 67
xiaodui16(8) = 67
suoshu(7, 4) = 39
sbsuoshu(25) = 39
xiaodui16(9) = 39
suoshu(2, 1) = 47
sbsuoshu(6) = 47
suoshu(3, 2) = 47
sbsuoshu(10) = 47
suoshu(1, 2) = 47
sbsuoshu(3) = 47
suoshu(1, 3) = 47
sbsuoshu(4) = 47
xiaodui16(10) = 47
suoshu(5, 3) = 16
sbsuoshu(16) = 16
xiaodui16(11) = 16
suoshu(6, 3) = 55
sbsuoshu(20) = 55
xiaodui16(12) = 55
suoshu(8, 3) = 77
sbsuoshu(30) = 77
suoshu(9, 3) = 77
sbsuoshu(35) = 77
xiaodui16(13) = 77
suoshu(5, 2) = 29
sbsuoshu(15) = 29
suoshu(6, 2) = 29
sbsuoshu(19) = 29
xiaodui16(14) = 29
suoshu(6, 1) = 12
sbsuoshu(18) = 12
xiaodui16(15) = 12
suoshu(8, 0) = 64
sbsuoshu(28) = 64
xiaodui16(16) = 64

dizhi(0, 1) = 1 ' "����"''д���꿪ʼ
dizhi(0, 7) = 2 ' "����"
dizhi(1, 2) = 3 '"�ɶ�"
dizhi(1, 3) = 4 ' "����"
dizhi(1, 6) = 5 ' "����"
dizhi(2, 1) = 6 ' "�Ϳ�"
dizhi(2, 4) = 7 ' "����"
dizhi(2, 5) = 8 ' "��ˮ"
dizhi(3, 1) = 9 ' "����"
dizhi(3, 2) = 10 ' "����"
dizhi(3, 5) = 11 ' "����"
dizhi(4, 0) = 12 ' "����"
dizhi(4, 5) = 13 ' "����"
dizhi(4, 6) = 14 ' "����"
dizhi(5, 2) = 15 ' "����"
dizhi(5, 3) = 16 ' "���"
dizhi(6, 0) = 17 ' "����"
dizhi(6, 1) = 18 ' "��ɳ"
dizhi(6, 2) = 19 ' "����"
dizhi(6, 3) = 20 ' "�ٴ�"
dizhi(6, 4) = 21 ' "����"
dizhi(6, 5) = 22 ' "��"
dizhi(6, 6) = 23 ' "ƽԭ"
dizhi(7, 2) = 24 ' "®��"
dizhi(7, 4) = 25 ' "���"
dizhi(7, 5) = 26 ' "���"
dizhi(7, 6) = 27 ' "��Ƥ"
dizhi(8, 0) = 28 ' "����"
dizhi(8, 1) = 29 ' "��ɣ"
dizhi(8, 3) = 30 ' "��ҵ"
dizhi(8, 4) = 31 ' "С��"
dizhi(8, 5) = 32 ' "����"
dizhi(8, 7) = 33 ' "��ƽ"
dizhi(9, 2) = 34 ' "���"
dizhi(9, 3) = 35 ' "��"
dizhi(9, 4) = 36 ' "����"
dizhi(9, 6) = 37 ' "����"
dizhi(10, 7) = 38 ' "��ƽ" 'д�������

diming(0, 7) = "����" ''д������ʼ
diming(1, 6) = "����"
diming(1, 3) = "����"
diming(1, 2) = "�ɶ�"
diming(0, 1) = "����"
diming(2, 1) = "�Ϳ�"
diming(2, 4) = "����"
diming(2, 5) = "��ˮ"
diming(3, 1) = "����"
diming(3, 2) = "����"
diming(3, 5) = "����"
diming(4, 0) = "����"
diming(4, 5) = "����"
diming(4, 6) = "����"
diming(5, 2) = "����"
diming(5, 3) = "���"
diming(6, 0) = "����"
diming(6, 1) = "��ɳ"
diming(6, 2) = "����"
diming(6, 3) = "�ٴ�"
diming(6, 4) = "����"
diming(6, 5) = "��"
diming(6, 6) = "ƽԭ"
diming(7, 2) = "®��"
diming(7, 4) = "���"
diming(7, 5) = "���"
diming(7, 6) = "��Ƥ"
diming(8, 0) = "����"
diming(8, 1) = "��ɣ"
diming(8, 3) = "��ҵ"
diming(8, 4) = "С��"
diming(8, 5) = "����"
diming(8, 7) = "��ƽ"
diming(9, 2) = "���"
diming(9, 3) = "��"
diming(9, 4) = "����"
diming(9, 6) = "����"
diming(10, 7) = "��ƽ" 'д��������
For f = 1 To 38
For fx = 0 To 10
For fy = 0 To 7
If dizhi(fx, fy) = f Then '����ת��
chengx(f) = fx
chengy(f) = fy
sbchengming(f) = diming(fx, fy)
End If
Next
Next
Next

Image1(0).Top = 600
Image1(0).Visible = True
Image1(0).Left = 1100
hongx = 0
hongy = 6
For f = 1 To 25
Image1(f).Visible = False
Next
'begin
Label1.Visible = False
Image1(26).Visible = True 'image1(26)shiditu
Image1(26).Top = 0
Image1(26).Left = 0
Image2(2).Left = 5000
Image2(2).Top = 3000
Image2(2).Visible = True
Image2(1).Visible = False
Image2(0).Visible = False
Image1(0).Visible = True

Dim qw As Long
Dim qw2 As Long
For qw = 1 To 38
For qw2 = 1 To 16
If sbsuoshu(qw) = xiaodui16(qw2) Then
Image3(qw).Picture = Image1(qw2).Picture 'С��ͼ��ɫ
End If
Next
If sbsuoshu(qw) = 0 Then
Image3(qw).Picture = �հ�ͼ.Picture
End If
Next
For f = 1 To 38
For fx = 0 To 10
For fy = 0 To 7
If dizhi(fx, fy) = f Then
Image3(f).Left = zuobiaox(fx) 'С��ͼλ��
Image3(f).Top = zuobiaoy(fy)
End If
Next
Next
Next


For f = 1 To 13
Label3(f).Height = 300
Label3(f).Width = 2000
Label3(f).Top = f * 600
Label3(f).Left = 1000
Next
For f = 0 To 13
Label3(f).Visible = False
Next
Label3(0).Top = 8500
Label3(0).Left = 100
For f = 1 To 84
����(f).Top = 700
����(f).Left = 4000
����(f).Visible = False
Next
Label2(0).Top = 9500
Label2(0).Left = 0
'end

Image2(3).Visible = False
Image1(0).Visible = False
tishi.Visible = True
tishi.Caption = "f7��������һ���棻f8���棻�����=����Ϸ"

sbnongye(7) = 54321 '���кô�
sbjinqian(13) = 1000000 '�����ô�


Dim es As Long
For es = 1 To 84
sbhoubeibingli(jzai(es)) = 5
If jwang(es) = es And es Then
jbingli(es) = 2000
End If
Next
End Sub

Function zuobiaox(x As Long) As Long
zuobiaox = 1000 + (x) * 1150
End Function
Function zuobiaoy(y As Long) As Long
zuobiaoy = 6680 - (y) * 1000
End Function

Private Sub Label5_Click()
List1.Enabled = True
��ť.Enabled = True
End Sub





Private Sub tishi_Click()
tishi.Visible = False
tishit.Enabled = False
tishi.Caption = ""
End Sub

Private Sub tishit_Timer()
'tishi.Visible = False
'tishit.Enabled = False
End Sub




Private Sub ������Ϸ_Click()
������Ϸ.Visible = False
��Ϸ�򵥻�.Visible = False
End Sub

Private Sub ��ť_Timer()
List1.Enabled = False
��ť.Enabled = False
End Sub

Private Sub ��ͼ_Timer()
For f = 1 To 84
����(f).Visible = False
Next
��ͼ.Enabled = False

End Sub

Private Sub ��Ϸ�򵥻�_Click()
������Ϸ.Visible = False: ��Ϸ�򵥻�.Visible = False
youxijiandanhua = True
End Sub

Private Sub ��Ϸʤ��_Timer()
Dim dd As Long
Dim ds As Long
ds = 0
For f = 1 To 38
If sbsuoshu(f) <> wang Then
ds = 1
End If
Next ''''
If ds = 0 And Image2(2).Visible = False Then
'tishi.Visible = True
'tishi.Caption = "ȫ���ѱ���ͳһ�����ڿ���ȥͳһ�����ˣ����ɣ�"
End If
dd = 0
For f = 1 To 38
If sbsuoshu(f) = wang And wang <> 0 Then
dd = 1
End If
Next
If dd = 0 And Image2(2).Visible = False Then
'tishi.Visible = True
'tishi.Caption = "��Ϸ��������ʧ����"
End If
End Sub
