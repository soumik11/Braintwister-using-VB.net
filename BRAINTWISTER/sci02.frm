VERSION 5.00
Begin VB.Form sci31 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7590
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
   ScaleHeight     =   7590
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Height          =   540
      Left            =   10680
      TabIndex        =   12
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   540
      Left            =   9360
      TabIndex        =   11
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      Height          =   615
      Left            =   9480
      TabIndex        =   10
      Top             =   6600
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "500"
      Height          =   495
      Index           =   3
      Left            =   6000
      TabIndex        =   9
      Top             =   4440
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1000"
      Height          =   495
      Index           =   2
      Left            =   1200
      TabIndex        =   8
      Top             =   4440
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1500"
      Height          =   495
      Index           =   1
      Left            =   6000
      TabIndex        =   7
      Top             =   2880
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "2000"
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   6
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "N="
      Height          =   495
      Left            =   3840
      TabIndex        =   16
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "P="
      Height          =   615
      Left            =   1440
      TabIndex        =   15
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label5 
      Height          =   615
      Left            =   4440
      TabIndex        =   14
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label4 
      Height          =   615
      Left            =   2280
      TabIndex        =   13
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "d"
      Height          =   495
      Index           =   3
      Left            =   5520
      TabIndex        =   5
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "c"
      Height          =   495
      Index           =   2
      Left            =   720
      TabIndex        =   4
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "b"
      Height          =   495
      Index           =   1
      Left            =   5520
      TabIndex        =   3
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "a"
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2. How many calories are there in 1 kilo calories ?"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   8895
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
      Width           =   7455
   End
End
Attribute VB_Name = "sci31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Dim negcnt, flag As Integer
Private Sub Command1_Click()
cnt = sci30.Text1.Text
negcnt = sci30.Text2.Text

For i = 0 To 3
If Option1(i).Value = True And Option1(i).Caption = "1000" Then
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
sci32.Label4.Caption = sci31.Text1.Text
sci32.Label5.Caption = sci31.Text2.Text
sci32.Show
sci31.Hide
End Sub





