Option Explicit

Public R14Acad
Public AcadR2000
Public Const dfLine = "CONTINUOUS"
'Public Const FmtStr = "####0.000000"
Public Const CADDFmtStr = "####0.000000"

Function pi() 
  pi = 4*atn(1) 'ArcCos(0) * 2
End Function

Function ToDegrees(x ) 
  ToDegrees = x * 180 / pi
End Function

Function ToRadians(x ) 
  ToRadians = x * pi / 180
End Function

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
	
	Sub setPoint(x1,y1,z1)
		x = x1
		y = y1
		z = z1
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


Sub MakeLayer(fp , layerName , Colour , LineType)
  fp.WriteLine("-LAYER m " & layerName)
  fp.WriteLine("color")
  fp.WriteLine(Colour)
  fp.WriteLine(layerName)
  fp.WriteLine("ltype")
  fp.WriteLine(LineType)
  fp.WriteLine(layerName)
  fp.WriteLine("")
End Sub


Function pointStr2(x,y) 
  pointStr2 = FormatNumber(x, 4) & "," & FormatNumber(y, 4)
End Function

Function displacementStr(deltaVector)
  displacementStr = "@" & FormatNumber(deltaVector.x, 4) & "," & FormatNumber(deltaVector.y, 4)
End Function

Function displacementStr2(x,y)
  displacementStr2 = "@" & FormatNumber(x, 4) & "," & FormatNumber(y, 4)
End Function

Function displacementStr3(deltaVector)
  displacementStr3 = "@" & FormatNumber(deltaVector.x, 4) & "," & FormatNumber(deltaVector.y, 4) & "," & FormatNumber(deltaVector.z, 4)
End Function

Sub startLine(fp)
	fp.WriteLine("LINE")
End Sub

Sub StartPline(fp)
  fp.WriteLine("PLINE")
End Sub

Sub closeLines(fp)
	fp.WriteLine("C")
End Sub

Sub drawLine(fp,pt1,pt2)
	fp.WriteLine("LINE " & pt1.ptStr & " " & pt2.ptStr)
	fp.WriteLine("")
End Sub

Sub drawLine2(fp,pt1,deltaVector)
	fp.WriteLine("LINE")
	fp.WriteLine(pt1.ptStr)
	fp.WriteLine(displacementStr3(deltaVector))
	fp.WriteLine("")
End Sub

'ptA: datum point taken to be bottom left hand corner
'Box drawn clockwise
'Lx=Length Horizontal Direction
'Ly=Length Vertical Direction
Sub drawBox(fp, ptA, Lx, Ly)
    Dim pts(3)
	Dim i
	
	Wscript.Echo "drawBox ..."
	For i=0 to ubound(pts)
		Set pts(i) = New clsCoordinate
		pts(i).initialise
		'NB: all z coordinates set to zero so can ignore
	Next
	
	'Bottom LH Corner
	pts(0).x = ptA.x
	pts(0).y = ptA.y
	
	'Top LH Corner
	pts(1).x = pts(0).x
	pts(1).y = pts(0).y + Ly
	
	'Top RH Corner
	pts(2).x = pts(1).x + Lx
	pts(2).y = pts(1).y
 
	'Bottom RH Corner
	pts(3).x = pts(2).x 
	pts(3).y = pts(2).y - Ly
	
	call startPline(fp)
	For i=0 to ubound(pts)
		pts(i).fprint(fp)
	Next
	call closeLines(fp)
		
	Wscript.Echo "...drawBox"
End Sub


Sub drawBox2(fp, ptA, Lx, Ly)
    Dim nextPt
	
	Wscript.Echo "drawBox2 ..."

	Set nextPt = New clsCoordinate
	nextPt.initialise
	'NB: z coordinates set to zero so can ignore

	call startPline(fp)
	
	'Bottom LH Corner
	nextPt.x = ptA.x
	nextPt.y = ptA.y
	nextPt.fprint(fp)
	
	'Top LH Corner
	nextPt.y = nextPt.y + Ly
	nextPt.fprint(fp)
	
	'Top RH Corner
	nextPt.x = nextPt.x + Lx
	nextPt.fprint(fp)
 
	'Bottom RH Corner
	nextPt.y = nextPt.y - Ly
	nextPt.fprint(fp)

	call closeLines(fp)
		
	Wscript.Echo "...drawBox2"
End Sub

Sub drawBox3(fp, ptA, Lx, Ly)
    Dim dist
	
	Wscript.Echo "drawBox3 ..."

	Set dist = New clsCoordinate
	dist.initialise

	call startPline(fp)
	'Bottom LH Corner
	ptA.fprint(fp)

	'Top LH Corner
	dist.x = 0
	dist.y = Ly
	fp.WriteLine("@" & dist.ptStr2 )

	'Top RH Corner
	dist.x = Lx
	dist.y = 0
	fp.WriteLine("@" & dist.ptStr2 )

	'Bottom RH Corner
	dist.x = 0
	dist.y = -Ly
	fp.WriteLine("@" & dist.ptStr2 )

	call closeLines(fp)
		
	Wscript.Echo "...drawBox3"
End Sub

Sub drawBox4(fp, ptA, Lx, Ly)
	
	Wscript.Echo "drawBox4 ..."
	call startPline(fp)
	ptA.fprint(fp)
	fp.WriteLine("@" & pointStr2(0,Ly) )
	fp.WriteLine("@" & pointStr2(Lx,0) )
	fp.WriteLine("@" & pointStr2(0,-Ly))
	call closeLines(fp)
		
	Wscript.Echo "...drawBox4"
End Sub

Sub drawBox4Z(fp, ptA, Lx, Ly, Zplane)
	Dim pt1
	 
	Set pt1 = New clsCoordinate
	pt1.initialise
	
	Wscript.Echo "drawBox4Z ..."
	call pt1.setPoint(ptA.x, ptA.y, Zplane)
	call startLine(fp)
	pt1.fprint(fp)
	fp.WriteLine("@" & pointStr2(0,Ly) )
	fp.WriteLine("@" & pointStr2(Lx,0) )
	fp.WriteLine("@" & pointStr2(0,-Ly))
	call closeLines(fp)
		
	Wscript.Echo "...drawBox4Z"
End Sub


Sub drawRectangularPrism(fp, ptA, Lx, Ly, Lz)
    Dim delta
	Dim pt1
	 
	Set delta = New clsCoordinate
	delta.initialise
	
	Set pt1 = New clsCoordinate
	pt1.initialise	
	call pt1.setPoint(ptA.x,ptA.y,ptA.z)
	 
	'Base
	call drawBox4Z(fp,pt1,Lx,Ly,pt1.z)
	
	'Top
	call drawBox4Z(fp,pt1,Lx,Ly,pt1.z+Lz)
	
	'Vertical Corners
	call delta.setPoint(0,0,Lz)
	call drawLine2(fp,pt1,delta)
	
	pt1.y = pt1.y + Ly 
	call drawLine2(fp,pt1,delta)
	
	pt1.x = pt1.x + Lx
	call drawLine2(fp,pt1,delta)
	
	pt1.y = pt1.y - Ly
	call drawLine2(fp,pt1,delta)	
	
End Sub

Sub drawRectangularBasedPyramid(fp, ptA, Lx, Ly, Lz, Zplane)
    Dim pt2
	Dim pt1
	 
	Set pt2 = New clsCoordinate
	pt2.initialise
	
	Set pt1 = New clsCoordinate
	pt1.initialise	
	call pt1.setPoint(ptA.x,ptA.y,Zplane)
	
	'Base
	call drawBox4Z(fp,pt1,Lx,Ly,Zplane)
	
	'Corner Lines
	call pt2.setPoint(Lx/2,Ly/2,Zplane+Lz)
	call drawLine(fp,pt1,pt2)
	
	pt1.y = pt1.y + Ly 
	call drawLine(fp,pt1,pt2)
	
	pt1.x = pt1.x + Lx
	call drawLine(fp,pt1,pt2)
	
	pt1.y = pt1.y - Ly
	call drawLine(fp,pt1,pt2)	
	
End Sub


'ptA=ridge point at eaves
Sub drawCentralDormer(fp, ptA, Lx, Ly, h1) 
	Dim pt1
	Dim pt2
	Dim pt3
	Dim pt4

	Set pt1 = New clsCoordinate
	pt1.initialise
	call pt1.setPoint(ptA.x,ptA.y,ptA.z)
	
	Set pt2 = New clsCoordinate
	pt2.initialise
	
	Set pt3 = New clsCoordinate
	pt3.initialise	
	
	Set pt4 = New clsCoordinate
	pt4.initialise		
	
	
	'Vertical
	call pt2.setPoint(pt1.x,pt1.y,pt1.z+h1)
	call drawLine(fp,pt1,pt2)	
	
	'Ridge
	call pt3.setPoint(pt2.x-Lx,pt2.y,pt2.z)
	call drawLine(fp,pt2,pt3)
	
	'Lhs
	call pt4.setPoint(pt1.x,pt1.y-Ly/2,pt1.z)
	call drawLine(fp,pt4,pt2)	
	call drawLine(fp,pt4,pt3)
	
	'Rhs
	call pt4.setPoint(pt1.x,pt1.y+Ly/2,pt1.z)
	call drawLine(fp,pt4,pt2)
	call drawLine(fp,pt4,pt3)
		
End Sub

Sub ScriptWriter1(scrFile)

	Dim ptA
	Dim L1, B1
	
	Set ptA = New clsCoordinate
	ptA.initialise
	
	B1=16	
	L1=3*B1
	call drawBox3(scrFile,ptA,L1,B1)
	scrFile.WriteLine("ZOOM E")
	
End Sub

Sub ScriptWriter2(scrFile)
	Dim pt0
	Dim ptA
	Dim BuildingLength, BuildingWidth
	Dim BuildingHeight
	
	Set pt0 = New clsCoordinate
	pt0.initialise	
	
	Set ptA = New clsCoordinate
	ptA.initialise
	
	BuildingHeight=2.4
	BuildingWidth=8	
	BuildingLength=2*BuildingWidth
	
	'Plan
	call drawBox4(scrFile,ptA,BuildingLength,BuildingWidth)
	
	'Elevation 1
	ptA.y = ptA.y-BuildingHeight
	call drawBox4(scrFile,ptA,BuildingLength,BuildingHeight)
	
	'Elevation 2
	ptA.x = ptA.x + BuildingLength
	call drawBox4(scrFile,ptA,BuildingWidth,BuildingHeight)
	
	'Elevation 3
	ptA.x = pt0.x - BuildingWidth
	call drawBox4(scrFile,ptA,BuildingWidth,BuildingHeight)
	
	'Elevation 4
	ptA.x = ptA.x - BuildingLength
	call drawBox4(scrFile,ptA,BuildingLength,BuildingHeight)

	'Section
	ptA.x = pt0.x + BuildingLength + BuildingWidth
	call drawBox4(scrFile,ptA,BuildingWidth,BuildingHeight)	
	
	scrFile.WriteLine("ZOOM E")
		
End Sub

Sub ScriptWriter3(scrFile)
    Dim dx,dy,dh
	Dim pt1, pt2, pt3
	Dim h, alpha, x 
	Dim h1,x1
	
	Dim Lx, Ly
	
	Set pt1 = New clsCoordinate
	pt1.initialise
	
	Set pt2 = New clsCoordinate
	pt2.initialise	
	
	Set pt3 = New clsCoordinate
	pt3.initialise	
	
	dx=1
	dy=1
	dh=1
	
	call pt1.setPoint(0,0,0)
	'call pt2.setPoint(1,1,1)
	'call drawLine(scrFile,pt1,pt2)
	'call drawBox4Z(scrFile,pt1,1,1,0)
	'call drawBox4Z(scrFile,pt1,1,1,1)
	
	call MakeLayer(scrFile,"member1", "red", dfLine)
	call drawRectangularPrism(scrFile,pt1,dx,dy,dh)
	
	
	x= dx/2
	alpha = ToRadians(30) 
	h = x * tan(alpha)
		
	call MakeLayer(scrFile,"roof", "magenta", dfLine)
	call drawRectangularBasedPyramid(scrFile,pt1,dx,dy,h,dh)
	
	call MakeLayer(scrFile,"dormer", "30", dfLine)
	Lx=dx/4
	alpha = ToRadians(30) 
	h1 = Lx * tan(alpha)

	call pt1.setPoint(dx,dy/2,dh)
	Ly = dy/2
	call drawCentralDormer(scrFile,pt1,Lx,Ly,h1)

    ' 'Vertical
	' call pt1.setPoint(dx,dy/2,dh)
	' call pt2.setPoint(pt1.x,pt1.y,pt1.z+h1)
	' call drawLine(scrFile,pt1,pt2)	
	
	' 'Ridge
	' call pt3.setPoint(dx/2+dx/4,dy/2,dh+h1)
	' call drawLine(scrFile,pt3,pt2)
	
	' 'Lhs
	' call pt1.setPoint(dx,dy/4,dh)
	' call drawLine(scrFile,pt1,pt2)	
	' call drawLine(scrFile,pt1,pt3)
	
	' 'Rhs
	' call pt1.setPoint(dx,dy/2+dy/4,dh)
	' call drawLine(scrFile,pt1,pt2)
	' call drawLine(scrFile,pt1,pt3)

	

	
	call MakeLayer(scrFile,"member5", "yellow", dfLine)
	call pt1.setPoint(0,0,0)
	call drawRectangularPrism(scrFile,pt1,dx/4,dy/4,dh/4)
	
	


End Sub

Sub CmdMain
	Dim fso
	Dim WshShell
	Dim objArgs
	
	Dim scrFile

	Dim startPath
	Dim MyDocPath
	Dim CADDscr

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set WshShell = CreateObject("WScript.Shell")
	Set objArgs = WScript.Arguments

	'Check Command Line Parameters if Needed
	MyDocPath = WshShell.SpecialFolders("MyDocuments")

	CADDscr = "mkviews.scr"
	startPath= MyDocPath & "\TestCAD\" & CADDscr
	Wscript.Echo startPath

    'Create Text File
	Set scrFile = fso.CreateTextFile(startPath, True)

	'Call Main Application
	Call ScriptWriter3(scrFile)

	scrFile.Close
	Wscript.Echo "All Done!"
	
End Sub

'-------------------------
'MAIN
'-------------------------

CmdMain

' END MAIN
'=========================