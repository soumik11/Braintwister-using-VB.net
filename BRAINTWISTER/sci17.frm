VERSION 5.00
Begin VB.Form sci46 
   BorderStyle     =   3  'Fixed Dialog
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
      Left            =   7680
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   540
      Left            =   6600
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   735
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
      Caption         =   "Oxygen"
      Height          =   495
      Index           =   3
      Left            =   5760
      TabIndex        =   9
      Top             =   4680
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Calcium"
      Height          =   495
      Index           =   2
      Left            =   1920
      TabIndex        =   8
      Top             =   4680
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Phosphorus"
      Height          =   495
      Index           =   1
      Left            =   5760
      TabIndex        =   7
      Top             =   3000
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Carbon"
      Height          =   495
      Index           =   0
      Left            =   1920
      TabIndex        =   6
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "N="
      Height          =   615
      Left            =   4560
      TabIndex        =   16
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "P="
      Height          =   615
      Left            =   2280
      TabIndex        =   15
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label5 
      Height          =   615
      Left            =   5280
      TabIndex        =   14
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label4 
      Height          =   615
      Left            =   3240
      TabIndex        =   13
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "d"
      Height          =   495
      Index           =   3
      Left            =   5280
      TabIndex        =   5
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "c"
      Height          =   495
      Index           =   2
      Left            =   1440
      TabIndex        =   4
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "b"
      Height          =   495
      Index           =   1
      Left            =   5280
      TabIndex        =   3
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "a"
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "17. Which of the following is not an element of the bone ?"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   9735
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
      Width           =   2295
   End
End
Attribute VB_Name = "sci46"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Dim negcnt, flag As Integer
Private Sub Command1_Click()
cnt = sci45.Text1.Text
negcnt = sci45.Text2.Text

For i = 0 To 3
If Option1(i).Value = True And Option1(i).Caption = "Carbon" Then
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
sci47.Show
sci46.Hide
sci47.Label4.Caption = sci46.Text1.Text
sci47.Label5.Caption = sci46.Text2.Text

End Sub

