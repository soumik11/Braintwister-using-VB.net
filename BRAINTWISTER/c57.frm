VERSION 5.00
Begin VB.Form c57 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "c57"
   ClientHeight    =   7590
   ClientLeft      =   4530
   ClientTop       =   1935
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   5760
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   4320
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9480
      TabIndex        =   8
      Top             =   6600
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   2055
      Left            =   5880
      Picture         =   "c57.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   4320
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Height          =   2295
      Left            =   1800
      Picture         =   "c57.frx":1874
      ScaleHeight     =   2235
      ScaleWidth      =   2355
      TabIndex        =   6
      Top             =   4320
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      Height          =   1695
      Left            =   6120
      Picture         =   "c57.frx":3BE4
      ScaleHeight     =   1635
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   1800
      Picture         =   "c57.frx":4C76
      ScaleHeight     =   1515
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "a"
      Height          =   1695
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "b"
      Height          =   1695
      Index           =   1
      Left            =   5760
      TabIndex        =   13
      Top             =   2280
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "c"
      Height          =   2295
      Index           =   2
      Left            =   1560
      TabIndex        =   14
      Top             =   4320
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "d"
      Height          =   2055
      Index           =   3
      Left            =   5640
      TabIndex        =   15
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label11 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   21
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   20
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label Label9 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   19
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   18
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "COMPUTER"
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
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "8."
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "What is the logo of JAVA software?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   8415
   End
End
Attribute VB_Name = "c57"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Dim negcnt As Integer
Dim flag As Integer
Private Sub Command1_Click()
cnt = c56.Text1.Text
negcnt = c56.Text2.Text
For i = 0 To 3
If Option1(i).Value = True And Option1(i).Caption = "d" Then
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
c58.Label9.Caption = c57.Text1.Text
c58.Label11.Caption = c57.Text2.Text
c58.Show
c57.Hide
End Sub

