VERSION 5.00
Begin VB.Form c65 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "c65"
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
      Left            =   6120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   4680
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
      TabIndex        =   7
      Top             =   6600
      Width           =   1575
   End
   Begin VB.PictureBox Picture4 
      Height          =   2055
      Left            =   5880
      Picture         =   "c65.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   2715
      TabIndex        =   6
      Top             =   4560
      Width           =   2775
   End
   Begin VB.PictureBox Picture3 
      Height          =   2175
      Left            =   1560
      Picture         =   "c65.frx":33DE
      ScaleHeight     =   2115
      ScaleWidth      =   2235
      TabIndex        =   5
      Top             =   4440
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Height          =   2055
      Left            =   5880
      Picture         =   "c65.frx":8379
      ScaleHeight     =   1995
      ScaleWidth      =   2595
      TabIndex        =   4
      Top             =   2040
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   1560
      Picture         =   "c65.frx":958D
      ScaleHeight     =   1875
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "a"
      Height          =   2055
      Index           =   0
      Left            =   1320
      TabIndex        =   8
      Top             =   2040
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "b"
      Height          =   2055
      Index           =   1
      Left            =   5640
      TabIndex        =   9
      Top             =   2040
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      Caption         =   "c"
      Height          =   2175
      Index           =   2
      Left            =   1320
      TabIndex        =   10
      Top             =   4440
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "d"
      Height          =   2175
      Index           =   3
      Left            =   5640
      TabIndex        =   11
      Top             =   4440
      Width           =   3015
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
      Left            =   840
      TabIndex        =   21
      Top             =   5280
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
      Top             =   5280
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
      Left            =   840
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
      Left            =   4200
      TabIndex        =   17
      Top             =   6720
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
      Left            =   3240
      TabIndex        =   16
      Top             =   6720
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
      Left            =   2040
      TabIndex        =   15
      Top             =   6720
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
      Left            =   1200
      TabIndex        =   14
      Top             =   6720
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
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "16."
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
      Caption         =   "Which of the following is a social networking site?"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   10215
   End
End
Attribute VB_Name = "c65"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Dim negcnt As Integer
Dim flag As Integer
Private Sub Command1_Click()
cnt = c64.Text1.Text
negcnt = c64.Text2.Text
For i = 0 To 3
If Option1(i).Value = True And Option1(i).Caption = "c" Then
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
c66.Label9.Caption = c65.Text1.Text
c66.Label11.Caption = c65.Text2.Text
c66.Show
c65.Hide

End Sub

