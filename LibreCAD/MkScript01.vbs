Option Explicit

Sub ScriptWriter1(scrFile)

	Dim ptA
	Dim L1, B1
	
	Set ptA = New clsCoordinate
	ptA.initialise
	
	B1=16	
	L1=3*B1
	call drawBox3(scrFile,ptA,L1,B1)
	call zoomExtents(scrFile)
	
End Sub
