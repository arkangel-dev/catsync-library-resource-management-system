VERSION 5.00
Begin VB.Form DueForm 
   BackColor       =   &H8000000E&
   Caption         =   "Print Form"
   ClientHeight    =   0
   ClientLeft      =   8445
   ClientTop       =   1680
   ClientWidth     =   1800
   LinkTopic       =   "Form1"
   Picture         =   "DueForm.frx":0000
   ScaleHeight     =   0
   ScaleMode       =   0  'User
   ScaleWidth      =   89.955
End
Attribute VB_Name = "DueForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    printCoord "Jamba Juice", 48.68, 96.74
    printCoord "0001", 48.69, 102.92
    printCoord "27/10/2018", 48.69, 117
    printCoord "27/10/2018", 48.69, 127.58
    printCoord "1", 48.69, 123.61
    printCoord "MVR 50", 115.53, 148.96
    printCoord "MVR 22", 115.53, 153.98
    printCoord "MVR 72", 115.68, 165.89
    printCoord DateTime.Now, 73.81, 74.2
    
End Sub

Function printCoord(TextString As String, XLoc As Single, YLoc As Single)
    FormImage.CurrentX = XLoc
    FormImage.CurrentY = YLoc
    FormImage.Print TextString
End Function

Private Sub Command2_Click()
    'FormImage.Picture = FormImage.Image
    SavePicture FormImage.Picture, App.path & "\image.bmp"
End Sub

Private Sub FormImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Caption = X & "," & Y
End Sub

