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

Function pointStr2(x,y) 
  pointStr2 = FormatNumber(x, 4) & "," & FormatNumber(y, 4)
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
	fp.WriteLine("LINE " & pt1.ptStr & " " & pt2.ptStr & " ")
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

	CADDscr = "DrawBox2.scr"
	startPath= MyDocPath & "\TestCAD\" & CADDscr
	Wscript.Echo startPath

    'Create Text File
	Set scrFile = fso.CreateTextFile(startPath, True)

	'Call Main Application
	Call ScriptWriter2(scrFile)

	scrFile.Close
	Wscript.Echo "All Done!"
	
End Sub

'-------------------------
'MAIN
'-------------------------

CmdMain

' END MAIN
'=========================