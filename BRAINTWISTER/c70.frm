VERSION 5.00
Begin VB.Form c70 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "c70"
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
   Begin VB.PictureBox Picture4 
      Height          =   2175
      Left            =   6600
      Picture         =   "c70.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   2355
      TabIndex        =   10
      Top             =   4920
      Width           =   2415
   End
   Begin VB.PictureBox Picture3 
      Height          =   2175
      Left            =   6480
      Picture         =   "c70.frx":4F9B
      ScaleHeight     =   2115
      ScaleWidth      =   2115
      TabIndex        =   9
      Top             =   2400
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   2040
      Picture         =   "c70.frx":7BE9
      ScaleHeight     =   2235
      ScaleWidth      =   2475
      TabIndex        =   8
      Top             =   5040
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   2040
      Picture         =   "c70.frx":9893
      ScaleHeight     =   2235
      ScaleWidth      =   2355
      TabIndex        =   7
      Top             =   2400
      Width           =   2415
   End
   Begin VB.OptionButton Option4 
      Height          =   2175
      Left            =   6240
      TabIndex        =   6
      Top             =   4920
      Width           =   2775
   End
   Begin VB.OptionButton Option3 
      Height          =   2295
      Left            =   1680
      TabIndex        =   5
      Top             =   5040
      Width           =   3015
   End
   Begin VB.OptionButton Option2 
      Height          =   2175
      Left            =   6240
      TabIndex        =   4
      Top             =   2400
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      Height          =   2415
      Left            =   1680
      TabIndex        =   3
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Which among the following is a Gaming Company?"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   1560
      Width           =   9735
   End
   Begin VB.Label Label2 
      Caption         =   "20."
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
Attribute VB_Name = "c70"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
