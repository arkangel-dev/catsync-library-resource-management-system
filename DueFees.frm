VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form DueFees 
   BackColor       =   &H8000000E&
   Caption         =   "Due Fees"
   ClientHeight    =   3360
   ClientLeft      =   10350
   ClientTop       =   5865
   ClientWidth     =   8670
   Icon            =   "DueFees.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   8670
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "DueFees.frx":1FBC8
      Height          =   1815
      Left            =   4320
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3201
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
   Begin MSAdodcLib.Adodc usersLateBooks 
      Height          =   375
      Left            =   1920
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      RecordSource    =   "select * from books"
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
   Begin VB.Label Currency3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      Caption         =   "Currency3"
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
      Left            =   3480
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000004&
      Caption         =   "Late Fees"
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
      Left            =   3480
      TabIndex        =   9
      Top             =   480
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   615
      Left            =   0
      Top             =   360
      Width           =   9255
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Grand Total :"
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
      Left            =   1440
      TabIndex        =   8
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label GrandTotal 
      BackColor       =   &H8000000E&
      Caption         =   "GrandTotal"
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
      Left            =   4800
      TabIndex        =   7
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label InitialLateCost 
      BackColor       =   &H8000000E&
      Caption         =   "InitialLateCost"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label DailyLateCost 
      BackColor       =   &H8000000E&
      Caption         =   "DailyLateCost"
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
      Left            =   4800
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Currency2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      Caption         =   "Currency2"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Currency1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      Caption         =   "Currency1"
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
      Left            =   3600
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Daily Charges :"
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
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Initial Late Charge :"
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
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "DueFees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    usersLateBooks.RecordSource = "select * from books where BorrowerID = '" & RentBooks.LibraryID & "'"
    usersLateBooks.Refresh
    
    Currency1.Caption = ReadXml("settings/functionsettings/currency")
    Currency2.Caption = ReadXml("settings/functionsettings/currency")
    Currency3.Caption = ReadXml("settings/functionsettings/currency")
    
    Dim InitialCharges As Double
    Dim DailyCharges As Double
    InitialCharges = 0
    DailyCharges = 0
    If Not usersLateBooks.Recordset.RecordCount = 0 Then
        For X = 1 To usersLateBooks.Recordset.RecordCount
            maxRentDate = DateAdd("d", ReadXml("settings/functionsettings/maxrentduration"), usersLateBooks.Recordset.Fields("BorrowedDate"))
            If maxRentDate < DateTime.Now Then
                InitialCharges = InitialCharges + Val(ReadXml("settings/functionsettings/initiallatecost"))
                DailyCharges = DailyCharges + (ReadXml("settings/functionsettings/dailylatecost") * (usersLateBooks.Recordset.Fields("BorrowedDate") - DateTime.Now))
            End If
            usersLateBooks.Recordset.MoveNext
        Next X
    Else
        MsgBox "Missing records"
    End If
    InitialLateCost.Caption = InitialCharges
    DailyLateCost.Caption = DailyCharges
    GrandTotal.Caption = InitialCharges + DailyCharges
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

