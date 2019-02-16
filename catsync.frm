VERSION 5.00
Begin VB.Form TediousLoadingPage 
   BackColor       =   &H8000000E&
   Caption         =   "CatSync LRMS"
   ClientHeight    =   5625
   ClientLeft      =   9600
   ClientTop       =   5295
   ClientWidth     =   9255
   Icon            =   "catsync.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   9255
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   120
   End
   Begin VB.ListBox MessageList 
      Height          =   450
      ItemData        =   "catsync.frx":1FBC8
      Left            =   840
      List            =   "catsync.frx":1FC05
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Linus Media Group Pvt. Ltd."
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   4320
      Width           =   5295
   End
   Begin VB.Label LoadingText 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "@Adobe Fan Heiti Std B"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   9015
   End
   Begin VB.Label CatSyncText 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "CatSync LRMS"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   3
      Top             =   3840
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   3555
      Left            =   3000
      Picture         =   "catsync.frx":1FDC3
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3540
   End
   Begin VB.Label VariableCount 
      Caption         =   "Label1"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "TediousLoadingPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer

Private Sub Form_Load()
    VariableCount.Caption = 0
End Sub

Private Sub Timer1_Timer()
    ' declaring the variables
    Dim message As String
    Dim count As Integer
    
    ' setting values to the values
    message = MessageList.List(1)
    count = VariableCount.Caption
    
    ' setting a random values and setting the
    ' seed value...
    Randomize (4)
    WaitCount = Int((Rnd * 3) + 1)
    Timer1.Interval = WaitCount * 1000
    ' Timer1.Interval = 50
    
    'setting the count to something... idk...
    count = count + 1
    VariableCount.Caption = count
    LoadingText.Caption = MessageList.List(count)
    
    If (count = MessageList.ListCount) Then
        FirstScreen.Show
        Unload Me
    End If
End Sub
