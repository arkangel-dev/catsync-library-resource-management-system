VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form BooksForm 
   Caption         =   "CatSync LRMS | Edit Books"
   ClientHeight    =   7725
   ClientLeft      =   7920
   ClientTop       =   4125
   ClientWidth     =   13065
   Icon            =   "Books.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   13065
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Books.frx":1FBC8
      Height          =   4095
      Left            =   8520
      TabIndex        =   19
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7223
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
   Begin VB.CheckBox Check4 
      Alignment       =   1  'Right Justify
      Caption         =   "Book Borrowed"
      DataField       =   "Borrowed"
      DataSource      =   "DataControl"
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
      TabIndex        =   18
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox Text2 
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
      TabIndex        =   17
      Top             =   360
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc DataControl 
      Height          =   375
      Left            =   6840
      Top             =   5760
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
      RecordSource    =   "Select * from books"
      Caption         =   "DataCoruse"
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
      Caption         =   "Restricted to staff members"
      DataField       =   "StaffMembersOnly"
      DataSource      =   "DataControl"
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
      TabIndex        =   14
      Top             =   6840
      Width           =   2655
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "Cannot be borrowed"
      DataField       =   "CannotBeBorrowed"
      DataSource      =   "DataControl"
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
      TabIndex        =   13
      Top             =   6360
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Fictional"
      DataField       =   "Fictional"
      DataSource      =   "DataControl"
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
      TabIndex        =   12
      Top             =   5880
      Width           =   2655
   End
   Begin VB.ListBox List1 
      DataField       =   "BookGenre"
      DataSource      =   "DataControl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      ItemData        =   "Books.frx":1FBE2
      Left            =   2640
      List            =   "Books.frx":1FBF5
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   4320
      Width           =   5655
   End
   Begin VB.TextBox RemarksTextbox 
      DataField       =   "Remarks"
      DataSource      =   "DataControl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   2640
      TabIndex        =   9
      Top             =   2280
      Width           =   5655
   End
   Begin VB.TextBox AuthorTextbox 
      DataField       =   "Author"
      DataSource      =   "DataControl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   1800
      Width           =   5655
   End
   Begin VB.TextBox BookBarcodeTextbox 
      DataField       =   "BookBarcode"
      DataSource      =   "DataControl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   1320
      Width           =   5655
   End
   Begin VB.TextBox BookShortCodeTextbox 
      DataField       =   "BookShortCode"
      DataSource      =   "DataControl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   840
      Width           =   5655
   End
   Begin VB.TextBox BookNameText 
      DataField       =   "BookName"
      DataSource      =   "DataControl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Text            =   "(new)"
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label Label11 
      Caption         =   "* This book cannot be rented out"
      Height          =   375
      Left            =   3120
      TabIndex        =   21
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "* This book can be rented out by staff members only"
      Height          =   375
      Left            =   3120
      TabIndex        =   20
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label Label10 
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
      TabIndex        =   16
      Top             =   480
      Width           =   1095
   End
   Begin VB.Image FirstRecordButton 
      Height          =   660
      Left            =   9600
      Picture         =   "Books.frx":1FC25
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   660
   End
   Begin VB.Image PreviousRecordButton 
      Height          =   660
      Left            =   10440
      Picture         =   "Books.frx":29994
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   660
   End
   Begin VB.Image NextRecordButton 
      Height          =   660
      Left            =   11280
      Picture         =   "Books.frx":329D3
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   660
   End
   Begin VB.Image LastRecordButton 
      Height          =   660
      Left            =   12120
      Picture         =   "Books.frx":F2A15
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   660
   End
   Begin VB.Image AddRecordButton 
      Height          =   660
      Left            =   10440
      Picture         =   "Books.frx":FC8D7
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   660
   End
   Begin VB.Image SaveRecordButton 
      Height          =   660
      Left            =   11280
      Picture         =   "Books.frx":1039A2
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   660
   End
   Begin VB.Image DeleteRecordButton 
      Height          =   660
      Left            =   12120
      Picture         =   "Books.frx":115DC3
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   660
   End
   Begin VB.Label Label8 
      Caption         =   " * Include any physical or dinstinctive damages "
      Height          =   855
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Book Genre"
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
      TabIndex        =   10
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Remarks"
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
      Caption         =   "Author"
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
      Caption         =   "Book Barcode"
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
      Caption         =   "Book Short Code"
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
      Caption         =   "Book Name"
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
Attribute VB_Name = "BooksForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' database controls
' records and database controls...

Private Sub AddRecordButton_Click()
    DataControl.Recordset.AddNew
End Sub

Private Sub DeleteRecordButton_Click()
    DataControl.Recordset.Delete
End Sub

Private Sub FirstRecordButton_Click()
    If DataControl.Recordset.RecordCount > 1 Then
        DataControl.Recordset.MoveLast
    End If
End Sub

Private Sub LastRecordButton_Click()
    If DataControl.Recordset.RecordCount > 1 Then
        DataControl.Recordset.MoveFirst
    End If
End Sub

Private Sub NextRecordButton_Click()
    If DataControl.Recordset.EOF = False Then
        DataControl.Recordset.MoveNext
    End If
End Sub

Private Sub PreviousRecordButton_Click()
    If DataControl.Recordset.BOF = False Then
        DataControl.Recordset.MovePrevious
    End If
End Sub

Private Sub SaveRecordButton_Click()
    DataControl.Recordset.Update
End Sub

' user interface controls...


Private Sub Image6_Click()
    RentBooks.Show
    Unload Me
End Sub
Private Sub Image3_Click()
    Unload Me
End Sub
Private Sub ScrollButton_Click()
    Dim count As Integer
    If Me.Height = 9120 Then
        For count = 9120 To 1680 Step -1
            Me.Height = count
        Next count
    ElseIf Me.Height = 1680 Then
        For count = 1680 To 9120
            Me.Height = count
        Next count
    End If
End Sub
Private Sub Image2_Click()
    FirstScreen.Show
    Unload Me
End Sub

Private Sub Text2_Change()
    If Text2.Text = "" Then
        DataControl.RecordSource = "Select * from books"
        DataControl.Refresh
    Else
        DataControl.RecordSource = "Select * from books where BookName like '%" & Text2.Text & "%'"
        DataControl.Refresh
        If DataControl.Recordset.RecordCount = 0 Then
            Text2.BackColor = &H8080FF
        Else
            Text2.BackColor = &H80000005
        End If
    End If
End Sub
