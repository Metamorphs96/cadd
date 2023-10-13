Option Explicit

'INTELLICAD OBJECTS
	Public CADserver 'As IntelliCAD.Application  ' Object

	Dim dwgDoc 'As IntelliCAD.Document
	Dim Block 'As IntelliCAD.Block
	Dim curBlock 'As IntelliCAD.Block
	Dim blockInstance 'As IntelliCAD.BlockInsert
	Dim attribs 'As IntelliCAD.Attributes
	Dim attr 'As IntelliCAD.Attribute
	Dim ssetObj 'As IntelliCAD.SelectionSet
	Dim pt 'As New IntelliCAD.Point

	Dim DocMspace
	Dim DocPspace
  

'GENERIC OBJECTS
	Dim ents 'entities collection
	Dim ent  'Single Entity
  
'SIMPLE VARIABLES
	Dim BlockName
	Dim count
	Dim fDrv, fPath, fName , fExt
	Dim i, j , r , C
	Dim isIncluded
	Dim countBlocks, countAttribs
	Dim attrTagName
	Dim attrTagValue

' This example demonstrates how to count the number of entities using
' the Count property.
Sub CountDrawingEntities(dwgDoc)
     Dim ents 'As Object
     Dim ct 'As Long

     Set ents = dwgDoc.ModelSpace
     ct = ents.count

     Wscript.Echo "No. of entities = " & ct

End Sub

  
'Launch the IntelliCAD application as a server controlled by this engine
Sub LaunchCADserver()

  'On Error GoTo ErrHandler_LaunchCADserver
  On Error Resume Next
  Set CADserver = Nothing
  
  Wscript.Echo "Attempting to Connect to Existing CAD server ..."
  'Launch the IntelliCAD application
  'Attempt to Connect to IntelliCAD if possible 
  'Set CADserver = CreateObject("Icad.Application.1")
  'Set CADserver = CreateObject("Icad.Application")
  Set CADserver = GetObject(, "Icad.Application")
  
  If Err <> 0 Then
	On Error GoTo 0
	call ErrorHandler("LaunchCADserver::1")
    Wscript.Echo  "... Error"
    Wscript.Echo  "Attempting to Connect to New Instance of CAD Server"
    
    On Error Resume Next
    Set CADserver = CreateObject("Icad.Application")
	'Set CADserver = GetObject(, "Icad.Application")
	If Err <> 0 Then
	    call ErrorHandler("LaunchCADserver::2")
		Close
		Wscript.Echo  "... LaunchCADserver: Exit Errors!"
		Wscript.Echo  Err.Number, Err.Description
		Stop	
	else
		CADserver.Visible = True
		Wscript.Echo  "CAD server launched ..."
		'Exit Sub
	End If
  Else
    CADserver.Visible = True
    Wscript.Echo "CAD server launched ..."
  End If

End Sub

Sub cMain
    Dim startPath
	Dim dwgPath
	
	LaunchCADserver
	
	dwgPath="C:\Users\workbench\Documents\TestCAD\ProgeCAD\test2023.dwg"
	If Not (CADserver is Nothing) then
		Set dwgDoc = CADserver.Documents.Open(dwgPath)
		Set DocMspace = dwgDoc.ModelSpace
		Set DocPspace = dwgDoc.PaperSpace		
		call CountDrawingEntities(dwgDoc)
		
		LaunchExcel
		If not(appExcel is Nothing) then
			If not(WrkBk is Nothing) then
				call MkCADWrkShts(WrkBk)
				call xtractLineTypes(dwgDoc)
				call ScanLayerTable(dwgDoc)
				call xtractTextStyles(dwgDoc)
				call xtractDimensionStyles(dwgDoc)
				call xtractNamedUCS(dwgDoc)
				call xtractNamedViews(dwgDoc)
				call xtractNamedViewPorts(dwgDoc)
				call xtractNamedBlocks(dwgDoc)
				call xtractRegisteredApplications(dwgDoc)
			End If
		End if
		
	End if
	
	
	Set dwgDoc = Nothing
	Set CADserver = Nothing
	
End Sub

Sub cmd2Main
    Dim startPath
	Dim dwgPath
	
	LaunchCADserver
	
	dwgPath="C:\Users\workbench\Documents\TestCAD\ProgeCAD\test2023.dwg"
	If Not (CADserver is Nothing) then
		Set dwgDoc = CADserver.Documents.Open(dwgPath)
		Set DocMspace = dwgDoc.ModelSpace
		Set DocPspace = dwgDoc.PaperSpace		
		call CountDrawingEntities(dwgDoc)
		
		LaunchExcel
		If not(appExcel is Nothing) then
			If not(WrkBk is Nothing) then
				call MkCADWrkShts(WrkBk)
				call xtractEntities(dwgDoc)
			End If
		End if
		
	End if
	
	
	Set dwgDoc = Nothing
	Set CADserver = Nothing
	
End Sub

Sub cmd3Main
    Dim startPath
	Dim dwgPath
	
	LaunchCADserver
	
	dwgPath="C:\Users\workbench\Documents\TestCAD\ProgeCAD\test2023.dwg"
	If Not (CADserver is Nothing) then
		Set dwgDoc = CADserver.Documents.Open(dwgPath)
		Set DocMspace = dwgDoc.ModelSpace
		Set DocPspace = dwgDoc.PaperSpace		
		call CountDrawingEntities(dwgDoc)
		
		LaunchExcel
		If not(appExcel is Nothing) then
			If not(WrkBk is Nothing) then
				call MkCADWrkShts(WrkBk)
				call xtractEntities2(dwgDoc)
			End If
		End if
		
	End if
	
	
	Set dwgDoc = Nothing
	Set CADserver = Nothing
	
End Sub