Option Explicit

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
	
	call zoomExtents(scrFile)
		
End Sub
