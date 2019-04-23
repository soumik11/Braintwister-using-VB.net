VERSION 5.00
Begin VB.Form Eng11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eng11"
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
      Left            =   5640
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Height          =   540
      Left            =   4200
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   900
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
      Caption         =   "Legible"
      Height          =   615
      Index           =   3
      Left            =   5040
      TabIndex        =   10
      Top             =   3960
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Tradgedy"
      Height          =   615
      Index           =   2
      Left            =   840
      TabIndex        =   9
      Top             =   3960
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Queer"
      Height          =   615
      Index           =   1
      Left            =   5040
      TabIndex        =   8
      Top             =   2640
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Virtue"
      Height          =   615
      Index           =   0
      Left            =   840
      TabIndex        =   7
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label8 
      Height          =   615
      Left            =   3360
      TabIndex        =   17
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label7 
      Height          =   735
      Left            =   1440
      TabIndex        =   16
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "N="
      Height          =   615
      Left            =   2520
      TabIndex        =   15
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "P="
      Height          =   615
      Left            =   600
      TabIndex        =   14
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "d"
      Height          =   615
      Index           =   3
      Left            =   4560
      TabIndex        =   6
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "c"
      Height          =   615
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "b"
      Height          =   615
      Index           =   1
      Left            =   4560
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "a"
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Pick out the incorrectly spelt word:"
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   7455
   End
   Begin VB.Label Label2 
      Caption         =   "2."
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1560
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
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "Eng11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cnt As Integer
Dim negcnt As Integer
Dim flag As Integer
Private Sub Command1_Click()
Text1.Text = Eng10.Text1.Text
cnt = Eng10.Text1.Text
negcnt = Eng10.Text2.Text
For i = 0 To 3
If Option1(i).Value = True And Option1(i).Caption = "Tradgedy" Then
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
Eng12.Label7.Caption = Text1.Text
Eng12.Label8.Caption = Text2.Text
Eng12.Show
Eng11.Hide
End Sub

