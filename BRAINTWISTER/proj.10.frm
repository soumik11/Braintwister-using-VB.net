VERSION 5.00
Begin VB.Form Eng10 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eng10"
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
      Height          =   540
      Left            =   5400
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   540
      Left            =   4080
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      Height          =   615
      Left            =   9480
      TabIndex        =   11
      Top             =   6600
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Belief"
      Height          =   495
      Index           =   3
      Left            =   5160
      TabIndex        =   10
      Top             =   3840
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Virtue"
      Height          =   495
      Index           =   2
      Left            =   840
      TabIndex        =   9
      Top             =   3840
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Incredible"
      Height          =   495
      Index           =   1
      Left            =   5160
      TabIndex        =   8
      Top             =   2520
      Width           =   3135
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Corrupt"
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   7
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   735
      Left            =   3000
      TabIndex        =   17
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      Height          =   735
      Left            =   1320
      TabIndex        =   16
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "N="
      Height          =   735
      Left            =   2160
      TabIndex        =   15
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "P="
      Height          =   735
      Left            =   600
      TabIndex        =   14
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "d"
      Height          =   615
      Index           =   3
      Left            =   4800
      TabIndex        =   6
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "c"
      Height          =   615
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label b 
      Caption         =   "b"
      Height          =   615
      Index           =   1
      Left            =   4800
      TabIndex        =   4
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "a"
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Synonyms for CREDENCE"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "1."
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   735
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
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "Eng10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cnt As Integer
Dim negcnt As Integer
Dim flag As Integer
Private Sub Command1_Click()

For i = 0 To 3
If Option1(i).Value = True And Option1(i).Caption = "Incredible" Then
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
Eng11.Label7.Caption = Text1.Text
Eng11.Label8.Caption = Text2.Text
Eng11.Show
Eng10.Hide
End Sub


