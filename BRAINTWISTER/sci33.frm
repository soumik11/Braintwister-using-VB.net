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
      Left            =   1080
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
      Left            =   1080
      TabIndex        =   6
      Top             =   3000
      Width           =   2655
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
      Left            =   720
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
      Left            =   720
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
