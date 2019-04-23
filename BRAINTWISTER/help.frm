VERSION 5.00
Begin VB.Form help 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   4530
   ClientTop       =   1935
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Arial Rounded MT Bold"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "help.frx":0000
   ScaleHeight     =   7440
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "BACK"
      Height          =   615
      Left            =   720
      TabIndex        =   8
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      Height          =   615
      Left            =   9840
      TabIndex        =   7
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "6. Finaly user will get to know about the score ."
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   6240
      Width           =   8415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "5. If the answer is wrong one marks will be deducted from the positive marks."
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   5400
      Width           =   10575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "4. Clicking the right answer will give one marks for each answer."
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   4440
      Width           =   8775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "3. Then user will got 20 questions per topic ."
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   3480
      Width           =   8775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "2. Next user have to choose the topic on which he wants to play."
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   2520
      Width           =   8775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1. First user needs to create a profile."
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Width           =   8415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INSTRUCTIONS"
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
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
End
Attribute VB_Name = "help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub help_Click()

End Sub

Private Sub Command1_Click()
loginpage.Show
help.Hide


End Sub

Private Sub Command2_Click()
introduction.Show
help.Hide

End Sub
