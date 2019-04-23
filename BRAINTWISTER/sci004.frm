VERSION 5.00
Begin VB.Form sci33 
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
      Left            =   8160
      TabIndex        =   12
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   540
      Left            =   6840
      TabIndex        =   11
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   495
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
      Caption         =   "Newton"
      Height          =   495
      Index           =   3
      Left            =   5880
      TabIndex        =   9
      Top             =   4680
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Ampere"
      Height          =   495
      Index           =   2
      Left            =   1560
      TabIndex        =   8
      Top             =   4680
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Kelvin"
      Height          =   495
      Index           =   1
      Left            =   5880
      TabIndex        =   7
      Top             =   3000
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "J/kg/'c"
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "N="
      Height          =   735
      Left            =   4080
      TabIndex        =   16
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "P="
      Height          =   735
      Left            =   1080
      TabIndex        =   15
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label5 
      Height          =   735
      Left            =   5040
      TabIndex        =   14
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   " "
      Height          =   735
      Left            =   2160
      TabIndex        =   13
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "d"
      Height          =   495
      Index           =   3
      Left            =   5520
      TabIndex        =   5
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "c"
      Height          =   495
      Index           =   2
      Left            =   1080
      TabIndex        =   4
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "b"
      Height          =   495
      Index           =   1
      Left            =   5520
      TabIndex        =   3
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "a"
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "4. What is the SI unit of specific heat ?"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   8535
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
      Width           =   3495
   End
End
Attribute VB_Name = "sci33"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Dim negcnt, flag As Integer
Private Sub Command1_Click()
cnt = sci32.Text1.Text
negcnt = sci32.Text2.Text

For i = 0 To 3
If Option1(i).Value = True And Option1(i).Caption = "J/kg/'c" Then
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


sci34.Label4.Caption = sci33.Text1.Text
sci34.Label5.Caption = sci33.Text2.Text

sci34.Show
sci33.Hide
End Sub

