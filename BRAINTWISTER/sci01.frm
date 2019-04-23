VERSION 5.00
Begin VB.Form sci30 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8010
   ClientLeft      =   4530
   ClientTop       =   1935
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial Rounded MT Bold"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   540
      Left            =   9120
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Left            =   7560
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEXT"
      Height          =   615
      Left            =   9480
      TabIndex        =   10
      Top             =   6600
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Newton"
      Height          =   540
      Index           =   3
      Left            =   5880
      TabIndex        =   9
      Top             =   4515
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Pascal"
      Height          =   540
      Index           =   2
      Left            =   1320
      TabIndex        =   8
      Top             =   4560
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Joule"
      Height          =   540
      Index           =   1
      Left            =   5880
      TabIndex        =   7
      Top             =   2640
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Watt"
      Height          =   510
      Index           =   0
      Left            =   1320
      TabIndex        =   6
      Top             =   2580
      Width           =   2775
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   420
      Left            =   4200
      TabIndex        =   16
      Top             =   6840
      Width           =   210
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   420
      Left            =   1920
      TabIndex        =   15
      Top             =   6840
      Width           =   210
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "P="
      Height          =   420
      Left            =   1080
      TabIndex        =   14
      Top             =   6840
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "N="
      Height          =   420
      Left            =   3480
      TabIndex        =   13
      Top             =   6840
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "d"
      Height          =   495
      Index           =   4
      Left            =   5400
      TabIndex        =   5
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "c"
      Height          =   495
      Index           =   3
      Left            =   840
      TabIndex        =   4
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "b"
      Height          =   495
      Index           =   2
      Left            =   5400
      TabIndex        =   3
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "a"
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1.What is the SI of energy?"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   7935
   End
   Begin VB.Label Label1 
      Caption         =   "SCIENCE"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "sci30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Dim negcnt As Integer
Dim flag As Integer

Private Sub Command2_Click()

For i = 0 To 3
If Option1(i).Value = True And Option1(i).Caption = "Joule" Then
cnt = cnt + 1
flag = 1
Exit For
End If
Next
If flag = 0 Then
 negcnt = negcnt + 1
 End If

Text1.Text = cnt
Text2.Text = negcnt

sci31.Label4.Caption = sci30.Text1.Text
sci31.Label5.Caption = sci30.Text2.Text

sci31.Show
sci30.Hide
End Sub

