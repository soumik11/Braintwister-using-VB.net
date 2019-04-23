VERSION 5.00
Begin VB.Form c61 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "c61"
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
      Left            =   5520
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   4080
      TabIndex        =   12
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
      Picture         =   "c61.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   1635
      TabIndex        =   7
      Top             =   4440
      Width           =   1695
   End
   Begin VB.PictureBox Picture3 
      Height          =   2175
      Left            =   1800
      Picture         =   "c61.frx":2992
      ScaleHeight     =   2115
      ScaleWidth      =   1995
      TabIndex        =   6
      Top             =   4320
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Left            =   5880
      Picture         =   "c61.frx":3B10
      ScaleHeight     =   2115
      ScaleWidth      =   1995
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   1800
      Picture         =   "c61.frx":7874
      ScaleHeight     =   1995
      ScaleWidth      =   2355
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.OptionButton Option1 
      Caption         =   "a"
      Height          =   2055
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   2040
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "b"
      Height          =   2175
      Index           =   1
      Left            =   5640
      TabIndex        =   9
      Top             =   2040
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "d"
      Height          =   2055
      Index           =   3
      Left            =   5640
      TabIndex        =   11
      Top             =   4440
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "c"
      Height          =   2175
      Index           =   2
      Left            =   1560
      TabIndex        =   10
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label8 
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
      TabIndex        =   21
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label7 
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
      Left            =   5160
      TabIndex        =   20
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label5 
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
      TabIndex        =   19
      Top             =   5160
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
      TabIndex        =   18
      Top             =   2880
      Width           =   255
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
      Left            =   5040
      TabIndex        =   17
      Top             =   6840
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
      Left            =   4080
      TabIndex        =   16
      Top             =   6840
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
      Left            =   2880
      TabIndex        =   15
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label6 
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
      Left            =   2040
      TabIndex        =   14
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "12."
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
      TabIndex        =   2
      Top             =   1320
      Width           =   615
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
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Identify the founder of APPLE."
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
      Width           =   5895
   End
End
Attribute VB_Name = "c61"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Dim negcnt As Integer
Dim flag As Integer
Private Sub Command1_Click()
cnt = c60.Text1.Text
negcnt = c60.Text2.Text
For i = 0 To 3
If Option1(i).Value = True And Option1(i).Caption = "a" Then
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
c62.Label9.Caption = c61.Text1.Text
c62.Label11.Caption = c61.Text2.Text
c62.Show
c61.Hide
End Sub

