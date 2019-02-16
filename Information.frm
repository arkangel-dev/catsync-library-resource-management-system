VERSION 5.00
Begin VB.Form Information 
   BackColor       =   &H8000000E&
   Caption         =   "About CatSync"
   ClientHeight    =   9300
   ClientLeft      =   9600
   ClientTop       =   3570
   ClientWidth     =   9630
   Icon            =   "Information.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   9630
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   8880
      Top             =   2040
   End
   Begin VB.Image Image2 
      Height          =   3585
      Left            =   3600
      Picture         =   "Information.frx":1FBC8
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3420
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   $"Information.frx":416D2
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   5
      Top             =   6000
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Copyright 2018 Isaam Rameez"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   5520
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Project Tergum / Isaam Rameez / MI College / 2018"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   8640
      Width           =   7455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   $"Information.frx":417A7
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      TabIndex        =   2
      Top             =   6840
      Width           =   7095
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "By Linus Media Group pvt ltd 2018 "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label Label1 
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
      Left            =   3720
      TabIndex        =   0
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   4065
      Left            =   3840
      Picture         =   "Information.frx":418DC
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   4380
   End
End
Attribute VB_Name = "Information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Image2.Visible = False
    Image1.Visible = True
End Sub
