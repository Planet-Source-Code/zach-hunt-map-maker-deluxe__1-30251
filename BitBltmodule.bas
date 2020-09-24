Attribute VB_Name = "BitBltmodule"
Public addy As Integer
Public addx As Integer
Public divide As Integer
Public size As Integer
Public newmap As Integer

'declaration for BitBlt
Declare Function BitBlt Lib "gdi32" _
                (ByVal hDestDC As Long, _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal nWidth As Long, _
                 ByVal nHeight As Long, _
                 ByVal hSrcDC As Long, _
                 ByVal xSrc As Long, _
                 ByVal ySrc As Long, _
                 ByVal dwRop As Long) As Long


