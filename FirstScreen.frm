VERSION 5.00
Begin VB.Form FirstScreen 
   BackColor       =   &H8000000E&
   ClientHeight    =   6675
   ClientLeft      =   9600
   ClientTop       =   4605
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FirstScreen.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6675
   ScaleWidth      =   9120
   Begin VB.CommandButton Users 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   0
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Image Image3 
      Height          =   660
      Left            =   8280
      Picture         =   "FirstScreen.frx":1FBC8
      Stretch         =   -1  'True
      Top             =   240
      Width           =   660
   End
   Begin VB.Label Label2 
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
      TabIndex        =   3
      Top             =   4200
      Width           =   5295
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
      TabIndex        =   2
      Top             =   3720
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   3555
      Left            =   3000
      Picture         =   "FirstScreen.frx":570B7
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3540
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Press start to begin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   5040
      Width           =   6015
   End
End
Attribute VB_Name = "FirstScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Books_Click()
    BooksForm.Show
    Unload Me
End Sub

Private Sub Command1_Click()
    UserForm.Show
    Unload Me
End Sub

Private Sub Image3_Click()
    End
End Sub

Private Sub Users_Click()
    Users.Caption = "Please wait..."
    RentBooks.Show
    Unload Me
End Sub
