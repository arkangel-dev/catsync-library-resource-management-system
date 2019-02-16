VERSION 5.00
Begin VB.Form Settings 
   BackColor       =   &H8000000E&
   Caption         =   "Settings"
   ClientHeight    =   7980
   ClientLeft      =   10350
   ClientTop       =   4470
   ClientWidth     =   8580
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   8580
   Begin VB.CommandButton Command6 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   26
      Top             =   7320
      Width           =   1575
   End
   Begin VB.ComboBox currencyField 
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
      Height          =   360
      ItemData        =   "Settings.frx":1FBC8
      Left            =   2280
      List            =   "Settings.frx":1FDDC
      TabIndex        =   24
      Top             =   7320
      Width           =   4215
   End
   Begin VB.TextBox dailyLateCharge 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "email"
      DataSource      =   "SettingsControls"
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
      Left            =   2280
      TabIndex        =   23
      Text            =   "10"
      Top             =   6840
      Width           =   4215
   End
   Begin VB.TextBox initLateCharge 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "email"
      DataSource      =   "SettingsControls"
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
      Left            =   2280
      TabIndex        =   21
      Text            =   "10"
      Top             =   6360
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   19
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox emailTitle 
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
      Height          =   360
      Left            =   2280
      TabIndex        =   18
      Text            =   "Return our book!"
      Top             =   3000
      Width           =   4215
   End
   Begin VB.TextBox maxRentField 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "email"
      DataSource      =   "SettingsControls"
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
      Left            =   2280
      TabIndex        =   16
      Text            =   "10"
      Top             =   5880
      Width           =   4215
   End
   Begin VB.TextBox institutionField 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "email"
      DataSource      =   "SettingsControls"
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
      Left            =   2280
      TabIndex        =   14
      Text            =   "Linus Media Group"
      Top             =   5400
      Width           =   4215
   End
   Begin VB.TextBox RecieverField 
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
      Height          =   360
      Left            =   2280
      TabIndex        =   11
      Top             =   3480
      Width           =   4215
   End
   Begin VB.CommandButton TestConnectionButton 
      Caption         =   "Test Connection"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox smtpportField 
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
      Height          =   360
      Left            =   2280
      TabIndex        =   7
      Text            =   "587"
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox smtpserverField 
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
      Height          =   360
      Left            =   2280
      TabIndex        =   6
      Text            =   "smtp.gmail.com"
      Top             =   2040
      Width           =   4215
   End
   Begin VB.TextBox passwordField 
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
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "48n4AV5^fr8u"
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox emailField 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "email"
      DataSource      =   "SettingsControls"
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
      Left            =   2280
      TabIndex        =   1
      Text            =   "notificationblender.samramirez@gmail.com"
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Label15 
      BackColor       =   &H8000000E&
      Caption         =   "Currency"
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
      Left            =   120
      TabIndex        =   25
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000E&
      Caption         =   "Initial Late Charge"
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
      Left            =   120
      TabIndex        =   22
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000E&
      Caption         =   "Charge per late day"
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
      Left            =   120
      TabIndex        =   20
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      Caption         =   "Test Email"
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
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      Caption         =   "Max Rent Days"
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
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Caption         =   "Institution Name"
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
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000004&
      Caption         =   "Institution and Monetary Policies"
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
      Left            =   2520
      TabIndex        =   12
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "Email Heading"
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
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000004&
      Caption         =   "Email Settings"
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
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "SMTP Port"
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
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "SMTP Server"
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
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Email Password"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Sender Email"
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
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   9255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   615
      Left            =   0
      Top             =   4320
      Width           =   9255
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub Form_Load()
' this function will read the xml file and init the fields...
    emailField.Text = ReadXml("settings/emailsettings/email")
    passwordField.Text = ReadXml("settings/emailsettings/password")
    smtpserverField.Text = ReadXml("settings/emailsettings/smtpserver")
    smtpportField.Text = ReadXml("settings/emailsettings/smtpport")
    emailTitle.Text = ReadXml("/settings/emailsettings/emailtitle")
    institutionField.Text = ReadXml("settings/functionsettings/institutionname")
    maxRentField.Text = ReadXml("settings/functionsettings/maxrentduration")
    initLateCharge.Text = ReadXml("settings/functionsettings/initiallatecost")
    dailyLateCharge.Text = ReadXml("settings/functionsettings/dailylatecost")
    currencyField.Text = ReadXml("settings/functionsettings/currency")
End Sub





' this is the button controls...
' ==============================


Private Sub Command6_Click()
    WriteXml currencyField, "settings/functionsettings/currency"
    WriteXml institutionField, "settings/functionsettings/institutionname"
    WriteXml initLateCharge.Text, "settings/functionsettings/initiallatecost"
    WriteXml dailyLateCharge, "settings/functionsettings/dailylatecost"
    WriteXml maxRentField, "settings/functionsettings/maxrentduration"
End Sub

Private Sub ShowSampleButton_Click()
    ColorSettingsPreview.Show
End Sub

Private Sub Command1_Click()
    WriteXml emailTitle, "/settings/emailsettings/emailtitle"
    WriteXml smtpserverField, "settings/emailsettings/smtpserver"
    WriteXml smtpportField, "settings/emailsettings/smtpport"
    WriteXml passwordField.Text, "/settings/emailsettings/password"
    WriteXml emailField.Text, "/settings/emailsettings/email"
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

Private Sub TestConnectionButton_Click()
    If Not RecieverField.Text = "" Then
        Dim oSmtp As New EASendMailObjLib.Mail
        oSmtp.LicenseCode = "TryIt"
        oSmtp.FromAddr = ReadXml("/settings/emailsettings/email") ' Set your Gmail email address
        oSmtp.AddRecipientEx RecieverField.Text, 0 ' Add recipient email address
        oSmtp.Subject = ReadXml("/settings/emailsettings/emailtitle") & " | Connection Test" ' Set email subject
        oSmtp.BodyFormat = 1
        ' Set email body
        oSmtp.BodyText = "<h1>Linus Media Group Library | CatSync</h1><br><p> Recently you rented a book on 1/1/2018, and it expired today. Please return it with a due fee of $100 </p><p>Sincerely,</p><p>Linus Media Group</p><br><tt>If you do not recognize or recall being enlisted on this mail list, please ignore this email.</tt><tt>Sent from CatSync Library Resource Management System 2018</tt> "
        oSmtp.ServerAddr = ReadXml("/settings/emailsettings/smtpserver") ' Gmail SMTP server address
        ' If you want to use direct SSL 465 port,
        ' Please add this line, otherwise TLS will be used.
        ' oSmtp.ServerPort = 465
        oSmtp.ServerPort = ReadXml("/settings/emailsettings/smtpport") ' set 25 or 587 port
        oSmtp.SSL_init ' detect SSL/TLS automatically
        oSmtp.UserName = ReadXml("/settings/emailsettings/email")
        oSmtp.Password = ReadXml("/settings/emailsettings/password")
        MsgBox "Testing connection. Please wait for confirmation or fail message."
        If oSmtp.SendMail() = 0 Then
            MsgBox "The test was successful."
        Else
            MsgBox "Failed to send email with the following error:" & oSmtp.GetLastErrDescription()
        End If
    Else
        MsgBox ("The recipient field cannot be empty to execute testing")
    End If
End Sub




