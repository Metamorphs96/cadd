Option Explicit

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
