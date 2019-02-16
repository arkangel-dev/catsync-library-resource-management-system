VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form RentBooks 
   BackColor       =   &H80000005&
   Caption         =   "CatSync LRMS | Rent Out Books"
   ClientHeight    =   15615
   ClientLeft      =   120
   ClientTop       =   300
   ClientWidth     =   28560
   Icon            =   "RentBooks.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   Begin VB.CommandButton Command4 
      Caption         =   "View Due Fees"
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
      Left            =   26280
      TabIndex        =   45
      Top             =   7920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox SearchField 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "RentBooks.frx":1FBC8
      Left            =   7680
      List            =   "RentBooks.frx":1FBDB
      TabIndex        =   44
      Text            =   "Book Name"
      Top             =   2280
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid BookGrid 
      Bindings        =   "RentBooks.frx":1FC10
      Height          =   9375
      Left            =   240
      TabIndex        =   3
      Top             =   5640
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   16536
      _Version        =   393216
      Appearance      =   0
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "RentBooks.frx":1FC2C
      Height          =   1095
      Left            =   16800
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1931
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
   Begin MSDataGridLib.DataGrid userFinderGrid 
      Bindings        =   "RentBooks.frx":1FC4C
      Height          =   1095
      Left            =   14160
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1931
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
   Begin MSAdodcLib.Adodc UserFinder 
      Height          =   375
      Left            =   11640
      Top             =   240
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
      RecordSource    =   "Select * from users "
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
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   16080
      Top             =   480
   End
   Begin MSAdodcLib.Adodc lateBooksControls 
      Height          =   330
      Left            =   9240
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
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
      RecordSource    =   "Select * from books where Borrowed = True"
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
   Begin VB.CheckBox borrowedBoolean 
      DataField       =   "Borrowed"
      DataSource      =   "bookscontrols"
      Height          =   375
      Left            =   13320
      TabIndex        =   41
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox BorrowedDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "BorrowedDate"
      DataSource      =   "bookscontrols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9600
      TabIndex        =   40
      Top             =   4920
      Width           =   5895
   End
   Begin VB.TextBox BorrowerID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "BorrowerID"
      DataSource      =   "bookscontrols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9600
      TabIndex        =   39
      Top             =   4440
      Width           =   5895
   End
   Begin VB.TextBox EmailAddress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "emailAddress"
      DataSource      =   "usercontrols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18960
      TabIndex        =   36
      Top             =   4440
      Width           =   6975
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Access to IT resources"
      DataField       =   "accessToITResources"
      DataSource      =   "usercontrols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   19560
      TabIndex        =   34
      Top             =   5040
      Width           =   3015
   End
   Begin VB.CheckBox staffMemberCheck 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Staff Member"
      DataField       =   "staffMember"
      DataSource      =   "usercontrols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   17400
      TabIndex        =   33
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "BookBarcode"
      DataSource      =   "bookscontrols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9600
      TabIndex        =   32
      Top             =   3960
      Width           =   5895
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "BookShortCode"
      DataSource      =   "bookscontrols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9600
      TabIndex        =   31
      Top             =   3480
      Width           =   5895
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "Author"
      DataSource      =   "bookscontrols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9600
      TabIndex        =   28
      Top             =   3000
      Width           =   5895
   End
   Begin VB.CommandButton ClearUser 
      Caption         =   "Refresh"
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
      Left            =   26280
      TabIndex        =   26
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "Remarks"
      DataSource      =   "bookscontrols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   1800
      TabIndex        =   25
      Top             =   3360
      Width           =   5895
   End
   Begin VB.TextBox SelectedBook 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "BookName"
      DataSource      =   "bookscontrols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   24
      Top             =   2880
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Return Book"
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
      Left            =   26280
      TabIndex        =   21
      Top             =   7200
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc RentedBookControls 
      Height          =   330
      Left            =   11640
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
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
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "Remarks"
      DataSource      =   "RentedBookControls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   19080
      TabIndex        =   20
      Top             =   7680
      Width           =   6975
   End
   Begin VB.TextBox SearchBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   17
      Top             =   2280
      Width           =   5895
   End
   Begin MSDataGridLib.DataGrid SelectedBookGrid 
      Bindings        =   "RentBooks.frx":1FC65
      Height          =   5175
      Left            =   17400
      TabIndex        =   16
      Top             =   9840
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   9128
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
   Begin VB.CommandButton Command3 
      Caption         =   "Add Book"
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
      Left            =   9840
      TabIndex        =   15
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get User"
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
      Left            =   26280
      TabIndex        =   14
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox LibraryID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18960
      TabIndex        =   13
      Top             =   2520
      Width           =   6975
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "Institution"
      DataSource      =   "usercontrols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18960
      TabIndex        =   11
      Top             =   3960
      Width           =   6975
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "lastName"
      DataSource      =   "usercontrols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18960
      TabIndex        =   10
      Top             =   3480
      Width           =   6975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "firstName"
      DataSource      =   "usercontrols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   18960
      TabIndex        =   9
      Top             =   3000
      Width           =   6975
   End
   Begin VB.TextBox SelectedBorrowedBook 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "BookName"
      DataSource      =   "RentedBookControls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   19080
      TabIndex        =   5
      Top             =   7200
      Width           =   6975
   End
   Begin MSAdodcLib.Adodc usercontrols 
      Height          =   330
      Left            =   10440
      Top             =   600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      RecordSource    =   "Select * from users"
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
   Begin MSAdodcLib.Adodc bookscontrols 
      Height          =   330
      Left            =   9240
      Top             =   600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin VB.CommandButton ClearButton 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Clear Search"
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
      Left            =   12000
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000004&
      Caption         =   "Find Books"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   48
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000004&
      Caption         =   "User's Rented Books"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   21720
      TabIndex        =   47
      Top             =   6240
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   615
      Left            =   17160
      Top             =   6120
      Width           =   11535
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000004&
      Caption         =   "Users"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   22440
      TabIndex        =   46
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Shape NotificationRedBar 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label BorrowedDateLabel 
      BackColor       =   &H8000000E&
      Caption         =   "Borrowed Date"
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
      Left            =   7920
      TabIndex        =   38
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label BorrowerIDLabel 
      BackColor       =   &H8000000E&
      Caption         =   "Borrower ID"
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
      Left            =   7920
      TabIndex        =   37
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Image Image7 
      Height          =   660
      Left            =   2760
      Picture         =   "RentBooks.frx":1FC86
      Stretch         =   -1  'True
      Top             =   240
      Width           =   660
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   1095
      Left            =   3600
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000E&
      Caption         =   "Email"
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
      Left            =   17400
      TabIndex        =   35
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   4800
      Picture         =   "RentBooks.frx":29304
      Stretch         =   -1  'True
      Top             =   240
      Width           =   660
   End
   Begin VB.Image Image5 
      Height          =   660
      Left            =   5640
      Picture         =   "RentBooks.frx":343EC
      Stretch         =   -1  'True
      Top             =   240
      Width           =   660
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000E&
      Caption         =   "Barcode"
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
      Left            =   7920
      TabIndex        =   30
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      Caption         =   "Book Code"
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
      Left            =   7920
      TabIndex        =   29
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
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
      Left            =   7920
      TabIndex        =   27
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
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
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
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
      TabIndex        =   23
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
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
      TabIndex        =   22
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
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
      Left            =   17280
      TabIndex        =   19
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "User ID"
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
      Left            =   17400
      TabIndex        =   12
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
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
      Left            =   17400
      TabIndex        =   8
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
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
      Left            =   17400
      TabIndex        =   7
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "User Name"
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
      Left            =   17400
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
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
      Left            =   17280
      TabIndex        =   4
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Image Image4 
      Height          =   660
      Left            =   1080
      Picture         =   "RentBooks.frx":4D7F0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   660
   End
   Begin VB.Image Image2 
      Height          =   660
      Left            =   240
      Picture         =   "RentBooks.frx":6559D
      Stretch         =   -1  'True
      Top             =   240
      Width           =   660
   End
   Begin VB.Image Image6 
      Height          =   660
      Left            =   1920
      Picture         =   "RentBooks.frx":6CC94
      Stretch         =   -1  'True
      Top             =   240
      Width           =   660
   End
   Begin VB.Image ScrollButton 
      Height          =   660
      Left            =   3960
      Picture         =   "RentBooks.frx":759B6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   660
   End
   Begin VB.Image Image3 
      Height          =   660
      Left            =   27600
      Picture         =   "RentBooks.frx":8E90D
      Stretch         =   -1  'True
      Top             =   240
      Width           =   660
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
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
      Left            =   22920
      TabIndex        =   0
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label CatSyncText 
      Alignment       =   2  'Center
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
      Left            =   22920
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   1215
      Left            =   -120
      Top             =   -120
      Width           =   29775
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   14775
      Left            =   17040
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   615
      Left            =   17160
      Top             =   1440
      Width           =   11535
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   615
      Left            =   0
      Top             =   1440
      Width           =   17175
   End
End
Attribute VB_Name = "RentBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub ClearButton_Click()
' this is the functioon to clear the search button...
    bookscontrols.RecordSource = "select * from books"
    bookscontrols.Refresh
End Sub


Private Sub ClearUser_Click()
' function to clear the current user
    LibraryID.Text = ""
    usercontrols.RecordSource = "select * from users where LibraryID = '0000'"
    usercontrols.Refresh
    RentedBookControls.RecordSource = "select * from books where BorrowerID = '0000'"
    RentedBookControls.Refresh
    Command4.Visible = False
End Sub

Private Sub Command1_Click()
' this is the function to return a book
    If Not LibraryID.Text = "" Then
        If RentedBookControls.Recordset.RecordCount > 0 Then
            RentedBookControls.Recordset.Fields("Borrowed") = False
            RentedBookControls.Recordset.Fields("BorrowerID") = ""
            'RentedBookControls.Recordset.Fields("BorrowedDate") = ""
            RentedBookControls.Recordset.Fields("DueTakenCareOf") = False
            RentedBookControls.Recordset.Update
            RentedBookControls.Refresh
            bookscontrols.Recordset.Update
            bookscontrols.Refresh
            RentedBookControls.Refresh
        Else
            MsgBox "This user has not rented out any books", , "CatSync LRMS"
        End If
    Else
        MsgBox "Enter A User Id", , "CatSync LRMS"
    End If

    
    
End Sub

Private Sub Command2_Click()
' the function to update the book the member rented out...
    usercontrols.RecordSource = "select * from users where LibraryID = '" & LibraryID.Text & "'"
    usercontrols.Refresh
    If usercontrols.Recordset.RecordCount = 0 Then
        MsgBox "User not found!"
    Else
        RentedBookControls.RecordSource = "select * from books where BorrowerID = '" & LibraryID.Text & "'"
        RentedBookControls.Refresh
        If Not RentedBookControls.Recordset.RecordCount = 0 Then
            maxLateDate = DateAdd("d", ReadXml("settings/functionsettings/maxrentduration"), RentedBookControls.Recordset.Fields("BorrowedDate"))
            If maxLateDate < DateTime.Now Then
                Command4.Visible = True
                MsgBox "This user has late fees!", , "CatSync LRMS"
            Else
                Command4.Visible = False
            End If
        Else
            Command4.Visible = False
        End If
    End If

End Sub

Private Sub Command3_Click()
' the function to rent out a book
    If Not bookscontrols.Recordset.RecordCount = 0 Then ' this condition check will see if any records are selected...
        If Not LibraryID.Text = "" Then ' this makes sure that a user is selected...
            If bookscontrols.Recordset.Fields("CannotBeBorrowed") = False Then ' this will check if the book is allowed to be borrowed...
                If (bookscontrols.Recordset.Fields("StaffMembersOnly") = True And usercontrols.Recordset.Fields("staffMember") = True) Or (bookscontrols.Recordset.Fields("StaffMembersOnly") = False) Then
                    ' the above condition will allow processing further if the (book is staff members only AND user is staff member) or (book is not staff members only
                    If bookscontrols.Recordset.Fields("Borrowed") = False Then ' this makes sure that the book is not already rented out..
                        bookscontrols.Recordset.Fields("Borrowed") = True
                        bookscontrols.Recordset.Fields("BorrowerID") = LibraryID.Text
                        bookscontrols.Recordset.Fields("BorrowedDate") = DateTime.Now
                        bookscontrols.Recordset.Fields("DueTakenCareOf") = False
                        bookscontrols.Recordset.Update
                        RentedBookControls.RecordSource = "select * from books where BorrowerID = '" & LibraryID.Text & "'"
                        RentedBookControls.Refresh
                    Else
                        MsgBox "This books is already rented.", , "CatSync LRMS"
                    End If
                Else
                    MsgBox "This member is not a staff member and is not allowed to take selected book", , "CatSync LRMS"
                End If
            Else
                MsgBox "This book cannot be borrowed.", , "CatSync LRMS"
            End If
        Else
            MsgBox "Enter A User ID", , "CatSync LRMS"
        End If
    Else
        MsgBox "Book not found.", , "CatSync LRMS"
    End If
        RentedBookControls.Refresh
End Sub


Private Sub Command4_Click()

    DueFees.usersLateBooks.Refresh
    DueFees.Show
End Sub

Private Sub Form_Load()
' this is the function that inits the whole rent out page
' this function will basically pre reset the whole page to its initial state
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    Me.WindowState = 2
    usercontrols.RecordSource = "select * from users where libraryID = '0000'"
    RentedBookControls.RecordSource = "select * from books where BookName = '0000'"
    usercontrols.Refresh
    RentedBookControls.Refresh
    CheckLateBooks
End Sub

Private Function ReadXml(path As String) As String
' this little function will return the read value from the xml
' file, which contains the settings...
    Dim doc As New MSXML2.DOMDocument60
    Dim success As Boolean
    Dim sResult As String

    success = doc.Load(App.path & "\config.xml")
    If success = False Then
        MsgBox doc.parseError.reason
    Else
        sResult = doc.selectSingleNode(path).Text
    End If
    ReadXml = sResult
End Function

Private Function WriteXml(value As String, path As String)
' this little function will write new values to the xml
' file...
    Dim doc As New MSXML2.DOMDocument60
    Dim success As Boolean

    success = doc.Load(App.path & "\config.xml")
    If success = False Then
        MsgBox doc.parseError.reason
    Else
        doc.selectSingleNode(path).Text = value
    End If
doc.Save (App.path & "\config.xml")
End Function



Private Sub Image1_Click()
' functions to show the settings page...
    Settings.Show
End Sub

Private Sub Image2_Click()
' this function will return the page to the main home page...
    FirstScreen.Show
    Unload Me
End Sub


Private Sub Image3_Click()
' this function will close the program
    Unload Me
    End
End Sub

Private Sub Image4_Click()
' this function will open the user editing page...
    UserForm.Show
End Sub

Private Sub Image5_Click()
' this function will display the funny help page...
    Information.Show
End Sub

Private Sub Image6_Click()
' this function will open the book editing page...
    BooksForm.Show
End Sub




Private Sub Image7_Click()
    MailPage.Show
End Sub

Private Sub ScrollButton_Click()
' this is the function to allow the user to control the visibility of the user interface
    Dim count As Integer
    Dim maxHeight As Integer
    maxHeight = Screen.Height
    If Me.WindowState = 2 Then
        Me.WindowState = 0
        For count = maxHeight To 1680 Step -2
            Me.Height = count
        Next count
    ElseIf Me.Height = 1680 Then
        For count = 1680 To maxHeight Step 10
            Me.Height = count
        Next count
        Me.WindowState = 2
    End If
End Sub

Private Sub SearchBox_Change()
' set the searchField variables
If SearchField.Text = "Book Name" Then
    Field = "BookName"
    ConditionStatement = "like"
ElseIf SearchField.Text = "Author" Then
    Field = "Author"
    ConditionStatement = "like"
ElseIf SearchField.Text = "Remarks" Then
    Field = "Remarks"
    ConditionStatement = "like"
Else
    MsgBox "Reference code error 448 | This search field is currently unsupported", , "CatSync LRMS"
End If

' this is the function to search for books
  searchQuery = SearchBox.Text
    If searchQuery = "" Then
        bookscontrols.RecordSource = "select * from books"
        bookscontrols.Refresh
    Else
        bookscontrols.RecordSource = "select * from books where " & Field & " " & ConditionStatement & " '%" & searchQuery & "%'"
        bookscontrols.Refresh
        ' this function will change the color based on the available results...
        If bookscontrols.Recordset.RecordCount = 0 Then
            SearchBox.BackColor = &H8080FF
        Else
            SearchBox.BackColor = &H80000004
        End If
    End If
End Sub

Private Function CheckLateBooks()
    Dim rentedbookcount As Integer
    Dim rentedate As Date
    Dim maxRentDate As Date
    Dim emailBody As String
    'MsgBox "CheckLateBook() Activated"
    rentedbookcount = 0
    lateBooksControls.Refresh
    
    If Not lateBooksControls.Recordset.RecordCount = 0 Then
        For X = 1 To lateBooksControls.Recordset.RecordCount
            If lateBooksControls.Recordset.Fields("Borrowed") = True Then
                rentedbookcount = rentedbookcount + 1
                rentdate = lateBooksControls.Recordset.Fields("BorrowedDate")
                maxRentDate = DateAdd("d", ReadXml("settings/functionsettings/maxrentduration"), rentdate)
                If DateTime.Now > maxRentDate Then
                    If lateBooksControls.Recordset.Fields("DueTakenCareOf") = False Then
                        BorrowerID = lateBooksControls.Recordset.Fields("BorrowerID")
                        lateBooksControls.Recordset.Fields("DueTakenCareOf") = True
                        BookName = lateBooksControls.Recordset.Fields("BookName")
                        UserFinder.RecordSource = "select * from users where libraryID = '" & BorrowerID & "'"
                        UserFinder.Refresh
                        BorrowerName = UserFinder.Recordset.Fields("lastName")
                        writeNotification (DateTime.Now & " |  Mr." & BorrowerName & " has a late book")
                        NotificationRedBar.Visible = True
                        
                        ' insert send mail function here...
                        If UserFinder.Recordset.Fields("sendAutoMail") Then
                            ' checking if the email address is empty...
                            If UserFinder.Recordset.Fields("emailAddress") = "" Then
                                MailPage.NotificationList.AddItem (DateTime.Now & " |  Mr." & BorrowerName & "'s email is missing from database")
                            Else
                                emailBody = "<h1>" & ReadXml("settings/emailsettings/emailtitle") & " | CatSync</h1><br><p> Recently you rented a book (" & BookName & ") on " & rentdate & ", and it expired today (" & DateTime.Now & "). Please return it with a due fee of " & ReadXml("settings/functionsettings/currency") & "." & ReadXml("settings/functionsettings/initiallatecost") & " and note that additional " & ReadXml("settings/functionsettings/currency") & "." & ReadXml("settings/functionsettings/dailylatecost") & " will be charged per late day.</p><p>Sincerely,</p><p>" & ReadXml("settings/functionsettings/institutionname") & "</p><br><tt>If you do not recognize or recall being enlisted on this mail list, please ignore this email.</tt><tt>Sent from CatSync Library Resource Management System 2018</tt> "
                                SendMail UserFinder.Recordset.Fields("emailAddress"), emailBody
                                writeNotification (DateTime.Now & " |  Email Sent To Mr." & BorrowerName & " regarding late book")
                                NotificationRedBar.Visible = True
                            End If
                        End If
                    End If
                End If
            End If
            lateBooksControls.Recordset.MoveNext
        Next X
    End If

End Function

Private Sub SearchButton_Click()
' this is the search button function
    searchQuery = SearchBox.Text
    If searchQuery = "" Then
        MsgBox "Fill the search box!"
    Else
        bookscontrols.RecordSource = "select * from books where BookName like '%" & searchQuery & "%'"
        bookscontrols.Refresh
    End If
    
End Sub

Private Sub SelectedBook_Change()
    If borrowedBoolean = False Then
        BorrowerIDLabel.Visible = False
        BorrowedDateLabel.Visible = False
        BorrowedDate.Visible = False
        BorrowerID.Visible = False
    Else
        BorrowerIDLabel.Visible = True
        BorrowedDateLabel.Visible = True
        BorrowedDate.Visible = True
        BorrowerID.Visible = True
    End If
End Sub

Private Function SendMail(reciever As String, message As String)
 
        Dim oSmtp As New EASendMailObjLib.Mail
        oSmtp.LicenseCode = "TryIt"
    
        ' Set your Gmail email address
        oSmtp.FromAddr = ReadXml("/settings/emailsettings/email")
    
        ' Add recipient email address
        oSmtp.AddRecipientEx reciever, 0
    
        ' Set email subject
        oSmtp.Subject = ReadXml("/settings/emailsettings/emailtitle")
        oSmtp.BodyFormat = 1
        ' Set email body
        oSmtp.BodyText = message
    
        ' Gmail SMTP server address
        oSmtp.ServerAddr = ReadXml("/settings/emailsettings/smtpserver")
    
        ' If you want to use direct SSL 465 port,
        ' Please add this line, otherwise TLS will be used.
        ' oSmtp.ServerPort = 465
    
        ' set 25 or 587 port
        oSmtp.ServerPort = ReadXml("/settings/emailsettings/smtpport")
    
        ' detect SSL/TLS automatically
        oSmtp.SSL_init
    
        oSmtp.UserName = ReadXml("/settings/emailsettings/email")
        oSmtp.Password = ReadXml("/settings/emailsettings/password")
        If Not oSmtp.SendMail() = 0 Then
            MsgBox "Failed to send email with the following error:" & oSmtp.GetLastErrDescription()
        End If
End Function

Private Sub Timer1_Timer()
    CheckLateBooks
End Sub

Private Function writeNotification(notification As String)
    Open App.path & "\notifications.txt" For Append As #1
    Print #1, notification
    Close #1
End Function
