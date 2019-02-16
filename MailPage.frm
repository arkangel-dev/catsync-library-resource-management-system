VERSION 5.00
Begin VB.Form MailPage 
   BackColor       =   &H8000000E&
   Caption         =   "Notifications"
   ClientHeight    =   5925
   ClientLeft      =   7920
   ClientTop       =   4650
   ClientWidth     =   13710
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MailPage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   13710
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   3
      Top             =   4680
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      FillColor       =   &H80000000&
      ForeColor       =   &H80000000&
      Height          =   4575
      Left            =   2400
      ScaleHeight     =   4700
      ScaleMode       =   0  'User
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   1080
      Width           =   135
   End
   Begin VB.ListBox NotificationList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   4590
      ItemData        =   "MailPage.frx":1FBC8
      Left            =   240
      List            =   "MailPage.frx":1FBCA
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1080
      Width           =   13215
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2880
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000004&
      Caption         =   "Notifications"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   615
      Left            =   -120
      Top             =   240
      Width           =   15135
   End
End
Attribute VB_Name = "MailPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Open App.path & "\notifications.txt" For Output As #1: Close #1
    NotificationList.Clear
    RentBooks.NotificationRedBar.Visible = False


    notification = DateTime.Now & " |  Notifications Cleared"
    Open App.path & "\notifications.txt" For Append As #1
    Print #1, notification
    Close #1
    
    
    RentBooks.NotificationRedBar.Visible = False
    Dim FileNo As Integer

    Dim TempStr As String
    FileNo = FreeFile
    Open App.path & "\notifications.txt" For Input As FileNo
       Do
          Line Input #FileNo, TempStr
          NotificationList.AddItem (TempStr)
          DoEvents
       Loop Until EOF(FileNo)
    Close #FileNo
End Sub

Private Sub Form_Load()
RentBooks.NotificationRedBar.Visible = False
Dim FileNo As Integer

   Dim TempStr As String
   FileNo = FreeFile
   Open App.path & "\notifications.txt" For Input As FileNo
      Do
         Line Input #FileNo, TempStr
         NotificationList.AddItem (TempStr)
         DoEvents
      Loop Until EOF(FileNo)
   Close #FileNo
End Sub
