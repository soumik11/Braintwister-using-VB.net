VERSION 5.00
Begin VB.Form Eng14 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eng14"
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
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5400
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "P.B.Shelley"
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
      Index           =   3
      Left            =   6360
      TabIndex        =   11
      Top             =   4080
      Width           =   4095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Robert Frost"
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
      Index           =   2
      Left            =   840
      TabIndex        =   10
      Top             =   4080
      Width           =   4095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Robert Lynd"
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
      Index           =   1
      Left            =   6360
      TabIndex        =   9
      Top             =   2880
      Width           =   4095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "William Wordsworth"
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
      Index           =   0
      Left            =   840
      TabIndex        =   8
      Top             =   2880
      Width           =   4095
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
   Begin VB.Label Label8 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   17
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   16
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "N="
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   15
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "P="
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   14
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "d"
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
      Index           =   3
      Left            =   5880
      TabIndex        =   6
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "c"
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
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "b"
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
      Index           =   1
      Left            =   5880
      TabIndex        =   4
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "a"
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
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Name the poet who wrote the poem'stopping by the woods on a snowwy evening'."
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   9975
   End
   Begin VB.Label Label2 
      Caption         =   "5."
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
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ENGLISH"
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
      Width           =   3495
   End
End
Attribute VB_Name = "Eng14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cnt As Integer
Dim negcnt As Integer
Dim flag As Integer
Private Sub Command1_Click()
cnt = Eng13.Text1.Text
negcnt = Eng13.Text2.Text

For i = 0 To 3
If Option1(i).Value = True And Option1(i).Caption = "William Wordsworth" Then
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
Eng15.Label7.Caption = Text1.Text
Eng15.Label8.Caption = Text2.Text
Eng15.Show
Eng14.Hide
End Sub
Private Sub Label5_Click()
label5.Caption = Eng13.Text1.Text
End Sub

Private Sub Label6_Click()
Label6.Caption = Eng13.Text2.Text
End Sub

Private Sub Label7_Click()
Label7.Caption = Eng13.Text1.Text
End Sub

Private Sub Label8_Click()
Label8.Caption = Eng13.Text2.Text
End Sub
