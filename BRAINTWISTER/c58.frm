VERSION 5.00
Begin VB.Form c58 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "c58"
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
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   5880
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   4440
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Both a and c"
      Height          =   1335
      Index           =   3
      Left            =   6360
      TabIndex        =   11
      Top             =   4560
      Width           =   3855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Logical Operations"
      Height          =   1335
      Index           =   2
      Left            =   1440
      TabIndex        =   10
      Top             =   4560
      Width           =   3855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Transferring Data"
      Height          =   1335
      Index           =   1
      Left            =   6360
      TabIndex        =   9
      Top             =   2640
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      Height          =   615
      Left            =   9480
      TabIndex        =   8
      Top             =   6600
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Arithmetical Opertaions"
      Height          =   1335
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label Label11 
      Height          =   615
      Left            =   4440
      TabIndex        =   17
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "N"
      Height          =   615
      Left            =   3480
      TabIndex        =   16
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label Label9 
      Height          =   615
      Left            =   2280
      TabIndex        =   15
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "P"
      Height          =   615
      Left            =   1440
      TabIndex        =   14
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label Label7 
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
      Left            =   960
      TabIndex        =   7
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label6 
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
      Left            =   5880
      TabIndex        =   6
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label5 
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
      Left            =   5880
      TabIndex        =   5
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label4 
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
      Left            =   960
      TabIndex        =   4
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "CPU's main task is____________."
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   8415
   End
   Begin VB.Label Label2 
      Caption         =   "9."
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1680
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
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "c58"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Dim negcnt As Integer
Dim flag As Integer
Private Sub Command1_Click()
cnt = c57.Text1.Text
negcnt = c57.Text2.Text
For i = 0 To 3
If Option1(i).Value = True And Option1(i).Caption = "Both a and c" Then
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
c59.Label9.Caption = c58.Text1.Text
c59.Label11.Caption = c58.Text2.Text
c59.Show
c58.Hide
End Sub

