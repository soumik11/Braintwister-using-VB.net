VERSION 5.00
Begin VB.Form sci32 
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
      Left            =   9720
      TabIndex        =   12
      Text            =   "0"
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   540
      Left            =   8400
      TabIndex        =   11
      Text            =   "0"
      Top             =   360
      Visible         =   0   'False
      Width           =   975
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
      Caption         =   "0 degree  f"
      Height          =   495
      Index           =   3
      Left            =   6480
      TabIndex        =   9
      Top             =   4440
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "32 degree  f"
      Height          =   495
      Index           =   2
      Left            =   1560
      TabIndex        =   8
      Top             =   4440
      Width           =   3615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "100 degree  c"
      Height          =   495
      Index           =   1
      Left            =   6120
      TabIndex        =   7
      Top             =   2880
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      Caption         =   "0 degree  c"
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "N="
      Height          =   420
      Left            =   3600
      TabIndex        =   16
      Top             =   6120
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "P="
      Height          =   420
      Left            =   1080
      TabIndex        =   15
      Top             =   6120
      Width           =   450
   End
   Begin VB.Label Label5 
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   1680
      TabIndex        =   13
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "d"
      Height          =   495
      Index           =   3
      Left            =   5760
      TabIndex        =   5
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "c"
      Height          =   495
      Index           =   2
      Left            =   1080
      TabIndex        =   4
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "b"
      Height          =   495
      Index           =   1
      Left            =   5760
      TabIndex        =   3
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "a"
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "3. What is the freezing point of water on fahrenheit scale?"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   10095
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
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "sci32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Dim negcnt, flag As Integer
Private Sub Command1_Click()
cnt = sci31.Text1.Text
negcnt = sci31.Text2.Text

For i = 0 To 3
If Option1(i).Value = True And Option1(i).Caption = "32 degree  f" Then
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

sci33.Label4.Caption = sci32.Text1.Text
sci33.Label5.Caption = sci32.Text2.Text
sci33.Show
sci32.Hide

End Sub



