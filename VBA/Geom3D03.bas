Attribute VB_Name = "Geom3D03"
Option Explicit

Public Type TCoord
  x As Double
  y As Double
  Z As Double
End Type

Sub SetCoord(pt As TCoord, x1 As Double, y1 As Double, z1 As Double)
  pt.x = x1
  pt.y = y1
  pt.Z = z1
End Sub

Sub WriteCoord(fp As Integer, pt As TCoord)
  Print #fp, Format(pt.x, "0.000"), Format(pt.y, "0.000"), Format(pt.Z, "0.000")
End Sub


