VERSION 5.00
Begin VB.Form sci36 
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
   Begin VB.TextBox Text2 
      Height          =   540
      Left            =   8520
      TabIndex        =   12
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   540
      Left            =   7080
      TabIndex        =   11
      Text            =   "0"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
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
      Caption         =   "None of the above"
      Height          =   540
      Index           =   3
      Left            =   6360
      TabIndex        =   9
      Top             =   4560
      Width           =   4095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Mixture"
      Height          =   540
      Index           =   2
      Left            =   1920
      TabIndex        =   8
      Top             =   4560
      Width           =   3135
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Compound"
      Height          =   540
      Index           =   1
      Left            =   6360
      TabIndex        =   7
      Top             =   3000
      Width           =   4095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Elements"
      Height          =   540
      Index           =   0
      Left            =   1920
      TabIndex        =   6
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label Label6 
      Caption         =   "N="
      Height          =   735
      Left            =   3600
      TabIndex        =   16
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "P="
      Height          =   735
      Left            =   1560
      TabIndex        =   15
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label4 
      Height          =   735
      Left            =   4440
      TabIndex        =   14
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Height          =   735
      Left            =   2400
      TabIndex        =   13
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "d"
      Height          =   495
      Index           =   3
      Left            =   6000
      TabIndex        =   5
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "c"
      Height          =   495
      Index           =   2
      Left            =   1560
      TabIndex        =   4
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "b"
      Height          =   495
      Index           =   1
      Left            =   6000
      TabIndex        =   3
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "a"
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "7. The substance containig same type of atom is ?"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   9615
   End
   Begin VB.Label sci7 
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
      Width           =   5775
   End
End
Attribute VB_Name = "sci36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Dim negcnt, flag As Integer
Private Sub Command1_Click()
cnt = sci35.Text1.Text
negcnt = sci35.Text2.Text

For i = 0 To 3
If Option1(i).Value = True And Option1(i).Caption = "Element" Then
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

sci37.Label4.Caption = sci36.Text1.Text
sci37.Label5.Caption = sci36.Text2.Text
sci37.Show
sci36.Hide
End Sub

