VERSION 5.00
Begin VB.Form zoomform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zoom"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   325
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox mapzoom 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3750
      Left            =   0
      ScaleHeight     =   250
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   0
      Top             =   0
      Width           =   3750
   End
End
Attribute VB_Name = "zoomform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
