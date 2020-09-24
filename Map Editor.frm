VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Mainform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Editor"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   506
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   659
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Cover 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   9600
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   16
      Top             =   7305
      Width           =   300
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   300
      LargeChange     =   2
      Left            =   1950
      Max             =   0
      TabIndex        =   14
      Top             =   7305
      Width           =   7650
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7230
      LargeChange     =   2
      Left            =   9600
      Max             =   0
      TabIndex        =   13
      Top             =   75
      Width           =   300
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   7650
      Left            =   0
      ScaleHeight     =   510
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   0
      Top             =   0
      Width           =   1875
      Begin VB.CommandButton zoom 
         Caption         =   "View Zoom Out"
         Height          =   450
         Left            =   75
         TabIndex        =   20
         Top             =   5550
         Width           =   1650
      End
      Begin VB.PictureBox rightclickstore 
         AutoRedraw      =   -1  'True
         Height          =   750
         Left            =   960
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   46
         TabIndex        =   18
         Top             =   4350
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.PictureBox leftclickstore 
         AutoRedraw      =   -1  'True
         Height          =   750
         Left            =   120
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   46
         TabIndex        =   17
         Top             =   4350
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.CommandButton Command8 
         Caption         =   "New Map"
         Height          =   450
         Left            =   75
         TabIndex        =   15
         Top             =   150
         Width           =   1650
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Save"
         Height          =   450
         Left            =   75
         TabIndex        =   11
         Top             =   1350
         Width           =   1650
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Open"
         Height          =   450
         Left            =   75
         TabIndex        =   10
         Top             =   750
         Width           =   1650
      End
      Begin VB.PictureBox rightclick 
         Height          =   750
         Left            =   975
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   46
         TabIndex        =   7
         Top             =   6600
         Width           =   750
      End
      Begin VB.PictureBox leftclick 
         Height          =   750
         Left            =   75
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   46
         TabIndex        =   6
         Top             =   6600
         Width           =   750
      End
      Begin VB.CommandButton Mapbutton 
         Caption         =   "Tile Set 5"
         Height          =   450
         Index           =   4
         Left            =   75
         TabIndex        =   5
         Top             =   4875
         Width           =   1650
      End
      Begin VB.CommandButton Mapbutton 
         Caption         =   "Tile Set 2"
         Height          =   450
         Index           =   1
         Left            =   75
         TabIndex        =   4
         Top             =   2850
         Width           =   1650
      End
      Begin VB.CommandButton Mapbutton 
         Caption         =   "Tile Set 3"
         Height          =   450
         Index           =   2
         Left            =   75
         TabIndex        =   3
         Top             =   3525
         Width           =   1650
      End
      Begin VB.CommandButton Mapbutton 
         Caption         =   "Tile Set 4"
         Height          =   450
         Index           =   3
         Left            =   75
         TabIndex        =   2
         Top             =   4200
         Width           =   1650
      End
      Begin VB.CommandButton Mapbutton 
         Caption         =   "Tile Set 1"
         Height          =   450
         Index           =   0
         Left            =   75
         TabIndex        =   1
         Top             =   2100
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Right Click"
         Height          =   195
         Left            =   960
         TabIndex        =   9
         Top             =   6300
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Left Click"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   6300
         Width           =   660
      End
   End
   Begin VB.PictureBox Map 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   1950
      ScaleHeight     =   481
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   510
      TabIndex        =   12
      Top             =   75
      Width           =   7650
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox Mapstore 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   7500
         Left            =   3480
         ScaleHeight     =   496
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   496
         TabIndex        =   19
         Top             =   3360
         Visible         =   0   'False
         Width           =   7500
      End
   End
   Begin VB.Line Line2 
      BorderWidth     =   7
      X1              =   125
      X2              =   700
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderWidth     =   7
      X1              =   125
      X2              =   125
      Y1              =   0
      Y2              =   510
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Map Maker Deluxe
'by Zach Hunt
''''''''''''''''''''
''''''''''''''''''''

Option Explicit

Private Sub Command6_Click()
'opens a map from a bmp file
CommonDialog1.CancelError = False
CommonDialog1.DialogTitle = "Open Map"
CommonDialog1.InitDir = App.Path
CommonDialog1.Filter = "Bitmaps (*.BMP)|*.BMP"
CommonDialog1.ShowOpen
On Error GoTo error
Mapstore.picture = LoadPicture(CommonDialog1.filename)
Map.picture = LoadPicture(CommonDialog1.filename)
error:
End Sub

Private Sub Command7_Click()
'saves a map to a bmp file
Dim filename As String
CommonDialog1.CancelError = False
CommonDialog1.DialogTitle = "Save Map"
CommonDialog1.DefaultExt = ".bmp"
CommonDialog1.InitDir = App.Path
CommonDialog1.Filter = "Bitmaps (*.BMP)|*.BMP"
CommonDialog1.ShowSave
On Error GoTo error
SavePicture Mapstore.Image, CommonDialog1.filename
error:
End Sub

Private Sub Command8_Click()
Dim columns As Integer
Dim rows As Integer
Dim x As Integer
Dim y As Integer

'starts a new map
On Error GoTo error
Do
    columns = InputBox("Enter map width" & vbCrLf & vbCrLf & "1-30", "Map Editor", 20)
Loop While columns > 30
Do
    rows = InputBox("Enter map height" & vbCrLf & vbCrLf & "1-30", "Map Editor", 20)
Loop While rows > 30
error:
'clears the map and draws the new lines if
'rows and columns are both valid
If columns <> 0 And rows <> 0 Then
    If columns > 11 Then HScroll1.Max = columns - 11
    If rows > 10 Then VScroll1.Max = rows - 10
    Map.Width = columns * 46
    Map.Height = rows * 46
    Mapstore.Width = columns * 46
    Mapstore.Height = rows * 46
    Map.picture = LoadPicture("")
    Mapstore.picture = LoadPicture("")
    newmap = True
    Mainform.Map.DrawWidth = 1.5
    Mainform.Mapstore.DrawWidth = 1.5
    For x = 0 To columns
    Map.Line (x * 46, 0)-(x * 46, rows * 46)
    Next x
    For y = 0 To rows
    Map.Line (0, y * 46)-(columns * 46, y * 46)
    Next y
    
    For x = 0 To columns
    Mapstore.Line (x * 46, 0)-(x * 46, rows * 46)
    Next x
    For y = 0 To rows
    Mapstore.Line (0, y * 46)-(columns * 46, y * 46)
    Next y
    Map.picture = Mapstore.Image
End If
End Sub

Private Sub zoom_Click()
'show the zoomed up version of the map
Load zoomform
zoomform.mapzoom.picture = LoadPicture("")
zoomform.mapzoom.Width = Mapstore.Width
zoomform.mapzoom.Height = Mapstore.Height
zoomform.mapzoom.PaintPicture Mapstore.Image, 0, 0, Mapstore.ScaleWidth / 2, Mapstore.ScaleHeight / 2
zoomform.Width = zoomform.mapzoom.Width * 8
zoomform.Height = zoomform.mapzoom.Height * 10
zoomform.Show
End Sub

Private Sub VScroll1_Change()
'changes the area viewed on the map
Map.Top = 5 - (VScroll1.Value * 46)
Map.picture = Mapstore.Image
End Sub

Private Sub HScroll1_Change()
'changes the area viewed on the map
Map.Left = 130 - (HScroll1.Value * 46)
Map.picture = Mapstore.Image
End Sub

Private Sub Map_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim row As Integer
Dim column As Integer
Dim picture As PictureBox
Dim zoom As Integer

'calculates the row and column clicked
column = x \ 46
row = y \ 46

If Button = 1 Then Set picture = leftclickstore
If Button = 2 Then Set picture = rightclickstore

'if a newmap has been started then it shows the picture
If newmap = True Then
    BitBlt Map.hDC, column * 46, row * 46, 46, 46, picture.hDC, 0, 0, vbSrcCopy
    BitBlt Mapstore.hDC, column * 46, row * 46, 46, 46, picture.hDC, 0, 0, vbSrcCopy
End If
End Sub

Private Sub Mapbutton_Click(Index As Integer)
Dim filename As String

Load Mapsform

'shows the different tile sets
filename = App.Path & "\RPG Tiles 0"
Select Case Index
Case 0
    filename = filename & "1.jpg"
Case 1
    filename = filename & "2.jpg"
Case 2
    filename = filename & "3.jpg"
Case 3
    filename = filename & "4.jpg"
Case 4
    filename = filename & "5.jpg"
End Select

Mapsform.background.picture = LoadPicture(filename)
Mapsform.Show
Mapsform.Left = 4650

'makes sure the sizes for the different tile sets are correct
Select Case Index
Case 0
    addx = 4
    addy = 4
    divide = 50
    size = 50
    Mapsform.VScroll1.Max = 25
    Mapsform.HScroll1.Max = 0
Case 1
    addx = 4
    addy = 3
    divide = 50
    size = 50
    Mapsform.VScroll1.Max = 0
    Mapsform.HScroll1.Max = 0
Case 2
    addx = 4
    addy = 3
    divide = 50
    size = 49
    Mapsform.VScroll1.Max = 0
    Mapsform.HScroll1.Max = 0
Case 3
    addx = 4
    addy = 3
    divide = 51
    size = 50
    Mapsform.VScroll1.Max = 0
    Mapsform.HScroll1.Max = 0
Case 4
    addx = 3
    addy = 3
    divide = 51
    size = 50
    Mapsform.VScroll1.Max = 375
    Mapsform.HScroll1.Max = 45
End Select

Mapsform.Timer1.Enabled = True
End Sub

