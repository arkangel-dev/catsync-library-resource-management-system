VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form UserForm 
   Caption         =   "CatSync LRMS | Edit users"
   ClientHeight    =   6075
   ClientLeft      =   7875
   ClientTop       =   4950
   ClientWidth     =   13110
   Icon            =   "UserForm.frx":0000
   LinkTopic       =   "Users"
   ScaleHeight     =   6075
   ScaleWidth      =   13110
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9360
      TabIndex        =   20
      Top             =   360
      Width           =   3495
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "UserForm.frx":1FBC8
      Height          =   2295
      Left            =   8400
      TabIndex        =   18
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc userdatabaseadoc 
      Height          =   375
      Left            =   7680
      Top             =   4800
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DataConfig.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DataConfig.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from users"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      Caption         =   "Allowed to access IT resources"
      DataField       =   "accessToITResources"
      DataSource      =   "userdatabaseadoc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      DataField       =   "emailAddress"
      DataSource      =   "userdatabaseadoc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   13
      Top             =   2760
      Width           =   5535
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      DataField       =   "Institution"
      DataSource      =   "userdatabaseadoc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   12
      Top             =   2280
      Width           =   5535
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      DataField       =   "contactNumber"
      DataSource      =   "userdatabaseadoc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   11
      Top             =   1800
      Width           =   5535
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      DataField       =   "libraryID"
      DataSource      =   "userdatabaseadoc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   10
      Top             =   1320
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "lastName"
      DataSource      =   "userdatabaseadoc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   9
      Top             =   840
      Width           =   5535
   End
   Begin VB.TextBox BookNameText 
      Appearance      =   0  'Flat
      DataField       =   "firstName"
      DataSource      =   "userdatabaseadoc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      TabIndex        =   8
      Top             =   360
      Width           =   5535
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "Automatically send email notifications"
      DataField       =   "sendAutoMail"
      DataSource      =   "userdatabaseadoc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Staff Member"
      DataField       =   "staffMember"
      DataSource      =   "userdatabaseadoc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label9 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   19
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label10 
      Caption         =   "* This user is allowed to access IT resources "
      Height          =   495
      Left            =   2880
      TabIndex        =   17
      Top             =   5280
      Width           =   4335
   End
   Begin VB.Label Label8 
      Caption         =   "* This user is a staff member"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Label Label7 
      Caption         =   "* Use this function to send automated email to notify members to return book 2 days before the due date."
      Height          =   735
      Left            =   2880
      TabIndex        =   15
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Image DeleteRecordButton 
      Height          =   660
      Left            =   12120
      Picture         =   "UserForm.frx":1FBE7
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   660
   End
   Begin VB.Image SaveRecordButton 
      Height          =   660
      Left            =   11280
      Picture         =   "UserForm.frx":28B0B
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   660
   End
   Begin VB.Image AddRecordButton 
      Height          =   660
      Left            =   10440
      Picture         =   "UserForm.frx":3AF2C
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   660
   End
   Begin VB.Image LastRecordButton 
      Height          =   660
      Left            =   12120
      Picture         =   "UserForm.frx":41FF7
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   660
   End
   Begin VB.Image NextRecordButton 
      Height          =   660
      Left            =   11280
      Picture         =   "UserForm.frx":4BEB9
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   660
   End
   Begin VB.Image PreviousRecordButton 
      Height          =   660
      Left            =   10440
      Picture         =   "UserForm.frx":10BEFB
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   660
   End
   Begin VB.Image FirstRecordButton 
      Height          =   660
      Left            =   9600
      Picture         =   "UserForm.frx":114F3A
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Email Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Institution"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Contact Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Library ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' user selection database control...

Private Sub AddRecordButton_Click()
    userdatabaseadoc.Recordset.AddNew
End Sub

Private Sub DeleteRecordButton_Click()
    userdatabaseadoc.Recordset.Delete
End Sub

Private Sub FirstRecordButton_Click()
    userdatabaseadoc.Recordset.MoveFirst
End Sub

Private Sub LastRecordButton_Click()
    userdatabaseadoc.Recordset.MoveLast
End Sub

Private Sub NextRecordButton_Click()
    If userdatabaseadoc.Recordset.EOF = False Then
        userdatabaseadoc.Recordset.MoveNext
    End If
End Sub

Private Sub PreviousRecordButton_Click()
    If userdatabaseadoc.Recordset.BOF = False Then
        userdatabaseadoc.Recordset.MovePrevious
    End If
End Sub

Private Sub SaveRecordButton_Click()
    userdatabaseadoc.Recordset.Update
End Sub

' user interface controls...

Private Sub Image3_Click()
    Unload Me
End Sub

Private Sub Text6_Change()
    If Text6.Text = "" Then
        userdatabaseadoc.RecordSource = "select * from users"
        userdatabaseadoc.Refresh
    Else
        userdatabaseadoc.RecordSource = "select * from users where libraryID = '" & Text6.Text & "'"
        userdatabaseadoc.Refresh
        If userdatabaseadoc.Recordset.RecordCount = 0 Then
            Text6.BackColor = &H8080FF
        Else
            Text6.BackColor = &H80000005
        End If
    End If
        
End Sub
