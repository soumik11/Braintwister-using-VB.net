VERSION 5.00
Begin VB.Form a05 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ag5"
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
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2760
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   4200
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
      Caption         =   "Tista"
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
      Left            =   5400
      TabIndex        =   10
      Top             =   4800
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Manas"
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
      Left            =   1800
      TabIndex        =   9
      Top             =   4800
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Bagmati"
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
      Left            =   5400
      TabIndex        =   8
      Top             =   3360
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Kaligandak"
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
      Left            =   1800
      TabIndex        =   7
      Top             =   3360
      Width           =   2535
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
      Left            =   2400
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
      Left            =   3120
      TabIndex        =   15
      Top             =   6600
      Width           =   975
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
      Height          =   615
      Left            =   1200
      TabIndex        =   14
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   8640
      X2              =   11160
      Y1              =   2160
      Y2              =   2160
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
      Top             =   5160
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
      Left            =   1200
      TabIndex        =   5
      Top             =   5160
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
      Top             =   3720
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
      Left            =   1200
      TabIndex        =   3
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Kathmandu is situated on the bank of river"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   10095
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
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   1680
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
Attribute VB_Name = "a05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt, negcnt, flag As Integer
Private Sub Command1_Click()
cnt = a04.Text1.Text
negcnt = a04.Text2.Text
For i = 0 To 3
If Option1(i).Value = True And Option1(i).Caption = "Bagmati" Then
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
a06.Label5.Caption = a05.Text1.Text
a06.Label6.Caption = a05.Text2.Text
a06.Show
a05.Hide
End Sub

