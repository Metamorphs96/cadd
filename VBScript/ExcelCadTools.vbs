Option Explicit

Public appExcel
Public WrkBk
Public WrkSht
Public dataBlock 'Range

Public WrkShtNames

'Excel Constants
Const xlCenter = -4108
Const xlSolid = 1
Const xlNone = -4142
Const xlContinuous = 1
Const xlAutomatic = -4105



Const xlDiagonalDown = 5
Const xlDiagonalUp = 6
Const xlEdgeLeft = 7
Const xlEdgeTop = 8
Const xlEdgeBottom = 9
Const xlEdgeRight = 10

Const xlInsideVertical = 11
Const xlInsideHorizontal = 12

Const xlThin = 2
Const xlThemeColorAccent2 = 6



Sub SetWrkShtNames()
  WrkShtNames = Array( _
						"LineStyles", _
						"Layers", _
						"TextStyles", _
						"DimStyles", _	
						"UCS", _						
						"Views", _	
						"ViewPorts", _
						"Blocks", _
						"XRefs", _
						"RegisteredApplications", _
						"Entities", _
						"Polylines" _
                      )
End Sub

Sub wbkWriteCellThemeHeader(ReportRange, r , c , V , colourCode , TintCode , formatStr1 )

  ReportRange.Offset(r, c).Value = V
  
  If formatStr1 <> "" Then
    ReportRange.Offset(r, c).NumberFormat = formatStr1
  End If
  
  ReportRange.Offset(r, c).Font.Bold = True
  ReportRange.Offset(r, c).HorizontalAlignment = xlCenter
  ReportRange.EntireColumn.AutoFit
  
  ReportRange.Offset(r, c).Interior.ThemeColor = colourCode
  ReportRange.Offset(r, c).Interior.TintAndShade = TintCode
  
  ReportRange.Offset(r, c).Interior.Pattern = xlSolid
  
  
  With ReportRange.Offset(r, c)
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone
    
    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
        
  End With
  
End Sub

Sub xtractLineTypes(dwgDoc)
    Dim LineStylesTable
    
    'IntelliCAD
    Dim LineStyle
  
    Dim i
       
    Set LineStylesTable = wrkBk.Worksheets("LineStyles").Range("A1:A1")
    
	call wbkWriteCellThemeHeader(LineStylesTable,0, 0, "Name", xlThemeColorAccent2, 0.6, "")
	
    i = 1
    With dwgDoc
      For Each LineStyle In .Linetypes
        LineStylesTable.Offset(i, 0).Value = LineStyle.name
        i = i + 1
      Next
    End With

End Sub





Sub ScanLayerTable(dwgDoc)
	Dim lyr, i
	Dim layerTable
	
	Set layerTable = WrkBk.Worksheets(WrkShtNames(1)).Range("A1:A1")
	
	call wbkWriteCellThemeHeader(layerTable,0, 0, "Name", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(layerTable,0, 1, "State1", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(layerTable,0, 2, "State2", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(layerTable,0, 3, "State3", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(layerTable,0, 4, "Colour", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(layerTable,0, 5, "Line Type", xlThemeColorAccent2, 0.6, "")
	
	
	call wbkWriteCellThemeHeader(layerTable,0, 11, "ColorMethod", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(layerTable,0, 12, "Red", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(layerTable,0, 13, "Blue", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(layerTable,0, 14, "Green", xlThemeColorAccent2, 0.6, "")

	i=1 'i=0 header row
    With dwgDoc
      For Each lyr In .Layers
        
        layerTable.Offset(i, 0).Value = lyr.name
        layerTable.Offset(i, 1).Value = lyr.LayerOn
        layerTable.Offset(i, 2).Value = lyr.Freeze
        layerTable.Offset(i, 3).Value = lyr.Lock
        layerTable.Offset(i, 4).Value = lyr.Color.ColorIndex '????
        layerTable.Offset(i, 5).Value = lyr.LineType
        
        'layerTable.Offset(i, 11).Value = lyr.Color.ColorMethod
        layerTable.Offset(i, 12).Value = lyr.Color.Red         'Appears to be colorindex
        'layerTable.Offset(i, 13).Value = lyr.Color.Blue        'zero
        'layerTable.Offset(i, 14).Value = lyr.Color.Green       'zero
        
        Set lyr = Nothing
        i = i + 1
      Next
	End With
End Sub


Sub xtractTextStyles(dwgDoc)
    Dim TextStylesTable
    
    'IntelliCAD
    Dim TextStyle
   
    Dim i
    
    Set TextStylesTable = wrkBk.Worksheets("TextStyles").Range("A1:A1")
	
    call wbkWriteCellThemeHeader(TextStylesTable,0, 0, "Name", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(TextStylesTable,0, 1, "Height", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(TextStylesTable,0, 2, "width", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(TextStylesTable,0, 3, "FontFile", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(TextStylesTable,0, 4, "ObliqueAngle", xlThemeColorAccent2, 0.6, "")
	
    i = 1
    With dwgDoc
      For Each TextStyle In .TextStyles
        TextStylesTable.Offset(i, 0).Value = TextStyle.name
        TextStylesTable.Offset(i, 1).Value = TextStyle.Height
        TextStylesTable.Offset(i, 2).Value = TextStyle.width
        TextStylesTable.Offset(i, 3).Value = TextStyle.FontFile
        TextStylesTable.Offset(i, 4).Value = TextStyle.ObliqueAngle
		
		Set TextStyle = Nothing
        i = i + 1
      Next
    End With

End Sub

Sub xtractDimensionStyles(dwgDoc)
    Dim DimStylesTable
    
    'IntelliCAD
    Dim dimStyle
   
    Dim i 
    
    Set DimStylesTable = wrkBk.Worksheets("DimStyles").Range("A1:A1")
    call wbkWriteCellThemeHeader(DimStylesTable,0, 0, "Name", xlThemeColorAccent2, 0.6, "")
	
    i = 1
    With dwgDoc
      For Each dimStyle In .DimensionStyles
        DimStylesTable.Offset(i, 0).Value = dimStyle.name
		
		Set dimStyle = Nothing
        i = i + 1
      Next
    End With
   

End Sub

Sub xtractNamedUCS(dwgDoc)
    Dim UCSTable
    
    'IntelliCAD
    Dim UCS
   
    Dim i
 
    Set UCSTable = wrkBk.Worksheets("UCS").Range("A1:A1")
    call wbkWriteCellThemeHeader(UCSTable,0, 0, "Name", xlThemeColorAccent2, 0.6, "")
	
    i = 1
    With dwgDoc
      For Each UCS In .UserCoordinateSystems
        UCSTable.Offset(i, 0).Value = UCS.name
		
		Set UCS = Nothing
        i = i + 1
      Next
    End With

End Sub

Sub xtractNamedViews(dwgDoc)
    Dim ViewsTable
    
    'IntelliCAD
    Dim View
   
    Dim i
 
    Set ViewsTable = wrkBk.Worksheets("Views").Range("A1:A1")
    call wbkWriteCellThemeHeader(ViewsTable,0, 0, "Name", xlThemeColorAccent2, 0.6, "")
	
    i = 1
    With dwgDoc
      For Each View In .Views
        ViewsTable.Offset(i, 0).Value = View.name
		
		Set View = Nothing
        i = i + 1
      Next
    End With
   
End Sub

Sub xtractNamedViewPorts(dwgDoc)
    Dim ViewPortsTable
    
    'IntelliCAD
    Dim ViewPort
   
    Dim i
    
    Set ViewPortsTable = wrkBk.Worksheets("ViewPorts").Range("A1:A1")
    call wbkWriteCellThemeHeader(ViewPortsTable,0, 0, "Name", xlThemeColorAccent2, 0.6, "")
	
    i = 1
    With dwgDoc
      For Each ViewPort In .Viewports
        ViewPortsTable.Offset(i, 0).Value = ViewPort.name
		
		Set ViewPort = Nothing
        i = i + 1
      Next
    End With
   
End Sub

Sub xtractNamedBlocks(dwgDoc)
    Dim BlocksTable
    Dim xrefTable
    
    Dim blockCount, xrefCount 
    
    'IntelliCAD
    Dim Block 
    Dim xref
   
    Dim i
    
    Set BlocksTable = wrkBk.Worksheets("Blocks").Range("A1:A1")
	call wbkWriteCellThemeHeader(BlocksTable,0, 0, "Name", xlThemeColorAccent2, 0.6, "")
	
    Set xrefTable = wrkBk.Worksheets("XRefs").Range("A1:A1")
    call wbkWriteCellThemeHeader(xrefTable,0, 0, "Name", xlThemeColorAccent2, 0.6, "")
	
    i = 1
    With dwgDoc
      blockCount = 0
      xrefCount = 0
      For Each Block In .Blocks
        If Block.IsXRef Then
          xrefCount = xrefCount + 1
          xrefTable.Offset(xrefCount, 0).Value = Block.name
        Else
          blockCount = blockCount + 1
          BlocksTable.Offset(blockCount, 0).Value = Block.name
        End If
      Next
    End With

End Sub


Sub xtractRegisteredApplications(dwgDoc)
    Dim RegisteredApplicationsTable
      
    'IntelliCAD
    Dim regApp
   
    Dim i

    Set RegisteredApplicationsTable = wrkBk.Worksheets("RegisteredApplications").Range("A1:A1")
    call wbkWriteCellThemeHeader(RegisteredApplicationsTable,0, 0, "Name", xlThemeColorAccent2, 0.6, "")
	
    i = 1
    With dwgDoc
      For Each regApp In .RegisteredApplications
        RegisteredApplicationsTable.Offset(i, 0).Value = regApp.name
		
		Set regApp = Nothing
        i = i + 1
      Next
    End With
   
End Sub

Sub xtractEntities(dwgDoc)
    Dim EntitiesTable
    
    'GENERIC OBJECTS
    Dim ents  'entities collection
    Dim ent   'Single Entity
    
    'SIMPLE VARIABLES
    Dim i

    Set EntitiesTable = wrkBk.Worksheets("Entities").Range("A1:A1")
    call wbkWriteCellThemeHeader(EntitiesTable,0, 0, "Item", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(EntitiesTable,0, 1, "EntityName", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(EntitiesTable,0, 2, "Handle", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(EntitiesTable,0, 3, "EntityType", xlThemeColorAccent2, 0.6, "")
	
    Set ents = dwgDoc.ModelSpace
    
    i = 1
    With dwgDoc
      For Each ent In ents
      
        EntitiesTable.Offset(i, 0).Value = i
        EntitiesTable.Offset(i, 1).Value = ent.EntityName
        EntitiesTable.Offset(i, 2).Value = ent.Handle
        EntitiesTable.Offset(i, 3).Value = ent.EntityType
        
        i = i + 1
      Next
    End With

End Sub

Sub ProcessLWPolyLine(dataTable, ent, ByRef k)
    Dim pt
	
    For Each pt In ent.Coordinates
        dataTable.Offset(k, 0).Value = ent.Handle
        dataTable.Offset(k, 1).Value = pt.x
        dataTable.Offset(k, 2).Value = pt.y
        dataTable.Offset(k, 3).Value = pt.z
        k = k + 1
    Next
	
End Sub


Sub xtractEntities2(dwgDoc)
    Dim EntitiesTable
    Dim PolyLineTable
	
    'GENERIC OBJECTS
    Dim ents  'entities collection
    Dim ent   'Single Entity
    
    'SIMPLE VARIABLES
    Dim i, k

    Set EntitiesTable = wrkBk.Worksheets("Entities").Range("A1:A1")
    call wbkWriteCellThemeHeader(EntitiesTable,0, 0, "Item", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(EntitiesTable,0, 1, "EntityName", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(EntitiesTable,0, 2, "Handle", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(EntitiesTable,0, 3, "EntityType", xlThemeColorAccent2, 0.6, "")
	
	Set PolyLineTable = wrkBk.Worksheets("Polylines").Range("A1:A1")
    call wbkWriteCellThemeHeader(PolyLineTable,0, 0, "Handle", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(PolyLineTable,0, 1, "X", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(PolyLineTable,0, 2, "Y", xlThemeColorAccent2, 0.6, "")
	call wbkWriteCellThemeHeader(PolyLineTable,0, 3, "Z", xlThemeColorAccent2, 0.6, "")	
	
    Set ents = dwgDoc.ModelSpace
    
    i = 1
	k = 1
    With dwgDoc
      For Each ent In ents
      
        EntitiesTable.Offset(i, 0).Value = i
        EntitiesTable.Offset(i, 1).Value = ent.EntityName
        EntitiesTable.Offset(i, 2).Value = ent.Handle
        EntitiesTable.Offset(i, 3).Value = ent.EntityType
		
		Select Case ent.EntityType
			Case 22 'LWPolyline
				Call ProcessLWPolyLine(PolyLineTable,ent,k)
			Case 25 'Polyline
				Wscript.Echo "Wrong Type of Polyline"
			Case else
				Wscript.Echo "Unknown Entity"
		End Select
        
        i = i + 1
      Next
    End With

End Sub


Sub MkCADWrkShts(WrkBk)
	Dim NumShts, i 
	
	SetWrkShtNames
	For i = UBound(WrkShtNames) To 0 Step -1
		Set wrksht = WrkBk.Worksheets.Add
		wrksht.Name = WrkShtNames(i)
	Next
	
End Sub


Sub LaunchExcel
	Set appExcel = CreateObject("Excel.Application")
	appExcel.visible=True

	If Err.Number <> 0 Then
		Wscript.Echo "Failed To Create New instance of Excel!"
	Else
		'Need to Add Workbook else Excel Opens and Exits
		Set WrkBk = appExcel.Workbooks.Add
		'Wscript.Echo "Dunno!"
	End If
End Sub

