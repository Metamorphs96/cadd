Option Explicit

Class clsCoordinate
	Public key
	Public x     
	Public y
	Public z
	  
	Sub initialise()
	  key = 0
	  x = 0
	  y = 0
	  z = 0
	End Sub
	
	Function ptStr2() 
	  ptStr2 = FormatNumber(x, 4) & "," & FormatNumber(y, 4)
	End Function	
	
	Function ptStr() 
	  ptStr = FormatNumber(x, 4) & "," & FormatNumber(y, 4) & "," & FormatNumber(z, 4)
	End Function
		
	Sub cprint()
	  Wscript.Echo  ptStr
	End Sub
	
	Sub fprint(fp)
	  fp.WriteLine(ptStr)
	End Sub
	
End Class

Function ptStr2D(Pt)
  ptStr2D = FormatNumber(Pt.x, 6) & "," & FormatNumber(Pt.y, 6)
End Function

Function pointStr2(x,y) 
  pointStr2 = FormatNumber(x, 4) & "," & FormatNumber(y, 4)
End Function
