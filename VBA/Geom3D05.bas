Attribute VB_Name = "Geom3D05"
Option Explicit

Public Type TNode
  x As Double
  y As Double
  Z As Double
  descr As String
End Type

Sub SetDescr(node As TNode, descr1 As String)
  node.descr = descr1
End Sub

Sub SetNode(node As TNode, descr1 As String, x1 As Double, y1 As Double, z1 As Double)
  node.descr = descr1
  node.x = x1
  node.y = y1
  node.Z = z1
End Sub

Sub WriteNODE(fp As Integer, node As TNode)
  Print #fp, node.descr, node.x, node.y, node.Z
End Sub




