VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close me!"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtExp 
      Height          =   1455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.ListBox lstWords 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1425
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtWord 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Type first few letter of the word..."
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Full SQL view"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   2880
      Width           =   5535
   End
   Begin VB.Label Label5 
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   1920
      Width           =   5415
   End
   Begin VB.Label Label4 
      Height          =   615
      Left            =   3360
      TabIndex        =   7
      Top             =   3240
      Width           =   5415
   End
   Begin VB.Label Label3 
      Caption         =   "WHERE"
      Height          =   735
      Left            =   3360
      TabIndex        =   6
      Top             =   1080
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "FROM"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "SELECT"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Activate()
    txtWord.SelLength = Len(txtWord.Text)
End Sub

Private Sub Form_Load()
    'Setting up the main database
    'DBMain was dimensioned as global database in module declaration
    Set DBMain = OpenDatabase(App.Path & "\data.mdb")
End Sub


Private Sub lstWords_Click()
    'When you click the listbox then you will get the
    'words you specified
    'So we are calling our routine with all words of the word
    'Lenght of the word is not important
    'Because it's already calculated by routine
    Call QueryData(lstWords.List(lstWords.ListIndex))
End Sub

Private Sub txtWord_Change()
    'if there is a real text in the box then call QueryData
    'with text which you want to find as a parameter
    Call QueryData(txtWord.Text)
End Sub


