VERSION 5.00
Begin VB.Form a12 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ah12"
   ClientHeight    =   7590
   ClientLeft      =   4530
   ClientTop       =   1875
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2640
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   3960
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   975
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
      TabIndex        =   11
      Top             =   6600
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Horse"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   5280
      TabIndex        =   10
      Top             =   4440
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Dog"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   1680
      TabIndex        =   9
      Top             =   4440
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Sheep"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   5280
      TabIndex        =   8
      Top             =   3000
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Cow"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   1680
      TabIndex        =   7
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "N="
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   17
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "P="
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   16
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   15
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   14
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "D."
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
      Left            =   4680
      TabIndex        =   6
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "C."
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
      Left            =   1080
      TabIndex        =   5
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "B."
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
      Left            =   4680
      TabIndex        =   4
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "A."
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
      Left            =   1080
      TabIndex        =   3
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "The inhabitant of Hindus Valley Civilization were not familiar with ..........."
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1200
      TabIndex        =   2
      Top             =   1560
      Width           =   10095
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
      TabIndex        =   1
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "ARTS"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "a12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cnt As Integer, negcnt, flag As Integer
Private Sub Command1_Click()
cnt = a11.Text1.Text
negcnt = a11.Text2.Text
For i = 0 To 3
If Option1(i).Value = True And Option1(i).Caption = "Horse" Then
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
a13.Label5.Caption = a12.Text1.Text
a13.Label6.Caption = a12.Text2.Text
a13.Show
a12.Hide
End Sub

Private Sub Label5_Click()
Label5.Caption = a11.Text1.Text
End Sub

Private Sub Label6_Click()
Label6.Caption = a11.Text2.Text
End Sub
