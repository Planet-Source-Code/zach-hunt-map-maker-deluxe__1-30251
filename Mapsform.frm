VERSION 5.00
Begin VB.Form Mapsform 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tiles"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   631
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   300
      LargeChange     =   20
      Left            =   0
      Max             =   100
      SmallChange     =   5
      TabIndex        =   2
      Top             =   0
      Width           =   9180
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7620
      LargeChange     =   20
      Left            =   9180
      Max             =   100
      SmallChange     =   5
      TabIndex        =   1
      Top             =   0
      Width           =   300
   End
   Begin VB.PictureBox background 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   37500
      Left            =   0
      ScaleHeight     =   2500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1000
      TabIndex        =   0
      Top             =   300
      Width           =   15000
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   840
         Top             =   720
      End
   End
End
Attribute VB_Name = "Mapsform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub background_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim column As Integer
Dim row As Integer
Dim picture As PictureBox
Dim storepicture As PictureBox

'figures out the column and row that was clicked
column = x \ divide
row = y \ divide

'sets the leftclick and rightclick pictures
If Button = 1 Then Set picture = Mainform.leftclick
If Button = 2 Then Set picture = Mainform.rightclick
If Button = 1 Then Set storepicture = Mainform.leftclickstore
If Button = 2 Then Set storepicture = Mainform.rightclickstore

BitBlt picture.hDC, 0, 0, size, size, background.hDC, (column * divide) + addx, (row * divide) + addy, vbSrcCopy
BitBlt storepicture.hDC, 0, 0, size, size, background.hDC, (column * divide) + addx, (row * divide) + addy, vbSrcCopy
End Sub


Private Sub Form_Unload(Cancel As Integer)
'reloads the image into the map's picture
Mainform.Map.picture = Mainform.Mapstore.Image
End Sub

Private Sub HScroll1_Change()
background.Left = 0 - (HScroll1.Value * 5)
End Sub

Private Sub Timer1_Timer()
'make sure the leftclick and rightclick pictures
'stay when you switch to a different form
BitBlt Mainform.leftclick.hDC, 0, 0, 50, 50, Mainform.leftclickstore.hDC, 0, 0, vbSrcCopy
BitBlt Mainform.rightclick.hDC, 0, 0, 50, 50, Mainform.rightclickstore.hDC, 0, 0, vbSrcCopy
Timer1.Enabled = False
End Sub

Private Sub VScroll1_Change()
background.Top = 20 - (VScroll1.Value * 5)
End Sub
