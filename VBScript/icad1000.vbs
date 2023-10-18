Option Explicit

Sub processDrawing(dwgDoc)

  With CADserver
    Set dwgDoc = .Documents.Open(dwgFile, True)
    Set DocMspace = dwgDoc.ModelSpace
    Set DocPspace = dwgDoc.PaperSpace
    
  End With

End Sub


Sub scanDrawingDataTables()
'Excel Objects
  Dim wrkBk 'As Workbook
  
  Dim layerSht 'As Worksheet
  Dim wrkSht1 'As Worksheet
  Dim wrkSht2 'As Worksheet
  Dim wrkSht3 'As Worksheet
  Dim wrkSht4 'As Worksheet
  Dim wrkSht5 'As Worksheet
  Dim wrkSht6 'As Worksheet
  Dim wrkSht7 'As Worksheet
  Dim wrkSht8 'As Worksheet
  Dim wrkSht9 'As Worksheet
 
  Dim layerTable 'As Range
  Dim TextStylesTable 'As Range
  Dim ViewsTable 'As Range
  Dim ViewPortsTable 'As Range
  Dim UCSTable 'As Range
  Dim DimStylesTable 'As Range
  Dim LineStylesTable 'As Range
  Dim BlocksTable 'As Range
  Dim xrefTable 'As Range
  Dim RegisteredApplicationsTable 'As Range 
  
  
  
'IntelliCAD Objects
  Dim ssetObj 'As IntelliCAD.SelectionSet
  Dim ents 'As IntelliCAD.SelectionSet
  Dim en 'As IntelliCAD.Entity
  Dim en1 'As IntelliCAD.Circle
  Dim en2 'As IntelliCAD.Line
  Dim lyr 'As IntelliCAD.Layer
  
  Dim TextStyle 'As IntelliCAD.TextStyle
  Dim View 'As IntelliCAD.View
  Dim ViewPort 'As IntelliCAD.ViewPort
  Dim UCS 'As IntelliCAD.UserCoordSystem
  Dim dimStyle 'As IntelliCAD.DimensionStyle
  Dim LineStyle 'As IntelliCAD.LineType
  Dim Block 'As IntelliCAD.Block
  Dim regApp 'As IntelliCAD.RegisteredApplication
  Dim xref 'As IntelliCAD.ExternalReference

'Simple Variables  
  Dim x , y , z 
  Dim x1, y1 , z1 
  
  Dim i , j 
  Dim blockCount, xrefCount
  
  Dim ssetName As String
  Dim layerName As String
  
  On Error On Error Resume Next
  
  Wscript.Echo "scanDrawingDataTables ..."
  
  Set wrkBk = ActiveWorkbook
  
  ClearDrawingDataWorkSheets
  
  Set layerSht = wrkBk.Worksheets("Layers")
  Set layerTable = layerSht.Range("A1:A1")
  Set TextStylesTable = wrkBk.Worksheets("TextStyles").Range("A1:A1")
  Set ViewsTable = wrkBk.Worksheets("Views").Range("A1:A1")
  Set ViewPortsTable = wrkBk.Worksheets("ViewPorts").Range("A1:A1")
  Set UCSTable = wrkBk.Worksheets("UCS").Range("A1:A1")
  Set DimStylesTable = wrkBk.Worksheets("DimStyles").Range("A1:A1")
  Set LineStylesTable = wrkBk.Worksheets("LineStyles").Range("A1:A1")
  Set BlocksTable = wrkBk.Worksheets("Blocks").Range("A1:A1")
  Set xrefTable = wrkBk.Worksheets("XRefs").Range("A1:A1")
  Set RegisteredApplicationsTable = wrkBk.Worksheets("RegisteredApplications").Range("A1:A1")
    
  With CADserver
    
    j = 0
    i = 1
    With dwgDoc
      For Each lyr In .Layers
        'lyr.Color.ColorMethod = vicColorMethodByACI '195
'        Debug.Print lyr.Color.ColorMethod
        'Debug.Print lyr.Color.BookName
        'Debug.Print lyr.Color.EntityColor
'        Debug.Print lyr.Color.Red, lyr.Color.Blue, lyr.Color.Green
'        Debug.Print lyr.name, lyr.Color.ColorIndex
        
        'Debug.Print lyr.TrueColor 'error
        'Debug.Print lyr.Color.ColorName
        
        layerTable.Offset(i, 0).Value = lyr.name
        layerTable.Offset(i, 1).Value = lyr.LayerOn
        layerTable.Offset(i, 2).Value = lyr.Freeze
        layerTable.Offset(i, 3).Value = lyr.Lock
        layerTable.Offset(i, 4).Value = lyr.Color.ColorIndex '????
        layerTable.Offset(i, 5).Value = lyr.LineType
        
        layerTable.Offset(i, 11).Value = lyr.Color.ColorMethod
        layerTable.Offset(i, 12).Value = lyr.Color.Red         'Appears to be colorindex
        layerTable.Offset(i, 13).Value = lyr.Color.Blue        'zero
        layerTable.Offset(i, 14).Value = lyr.Color.Green       'zero
        
        'layerTable.Offset(i, 6).Value = lyr.TrueColor
        'layerTable.Offset(i, 7).Value = lyr.color.ColorName
        'layerTable.Offset(i, 7).Value = lyr.color.EntityColor
        Set lyr = Nothing
        i = i + 1
      Next lyr

      i = 1
      For Each TextStyle In .TextStyles
        TextStylesTable.Offset(i, 0).Value = TextStyle.name
        TextStylesTable.Offset(i, 1).Value = TextStyle.Height
        TextStylesTable.Offset(i, 2).Value = TextStyle.width
        TextStylesTable.Offset(i, 3).Value = TextStyle.FontFile
        TextStylesTable.Offset(i, 4).Value = TextStyle.ObliqueAngle
        
        i = i + 1
      Next TextStyle
      
      i = 1
      For Each View In .Views
        ViewsTable.Offset(i, j).Value = View.name
        i = i + 1
      Next View
      
      i = 1
      For Each ViewPort In .Viewports
        ViewPortsTable.Offset(i, j).Value = ViewPort.name
        i = i + 1
      Next ViewPort
      
      i = 1
      For Each UCS In .UserCoordinateSystems
        UCSTable.Offset(i, j).Value = UCS.name
        i = i + 1
      Next UCS

      i = 1
      For Each dimStyle In .DimensionStyles
        DimStylesTable.Offset(i, j).Value = dimStyle.name
        i = i + 1
      Next dimStyle
      
      i = 1
      For Each LineStyle In .Linetypes
        LineStylesTable.Offset(i, j).Value = LineStyle.name
        i = i + 1
      Next LineStyle
     
      blockCount = 0
      xrefCount = 0
      For Each Block In .Blocks
        If Block.IsXRef Then
          xrefCount = xrefCount + 1
          xrefTable.Offset(xrefCount, j).Value = Block.name
        Else
          blockCount = blockCount + 1
          BlocksTable.Offset(blockCount, j).Value = Block.name
        End If
      Next Block
      
      i = 1
      For Each regApp In .RegisteredApplications
        RegisteredApplicationsTable.Offset(i, j).Value = regApp.name
        i = i + 1
      Next regApp
      
    End With
    
  End With
  
  Wscript.Echo "... scanDrawingDataTables"
  
Exit_scanDrawingDataTables:
  Exit Sub
  
ErrHandler_scanDrawingDataTables:
  Close
  Call ErrorMessages("scanDrawingDataTables")
  Stop
End Sub

