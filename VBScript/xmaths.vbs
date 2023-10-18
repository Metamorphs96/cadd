Option Explicit

Function pi() 
  pi = CDbl(4*atn(1)) 'ArcCos(0) * 2
End Function

Function ToDegrees(x ) 
  ToDegrees = x * 180 / pi
End Function

Function ToRadians(x ) 
  ToRadians = x * pi / 180
End Function