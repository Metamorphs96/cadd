Attribute VB_Name = "ScriptSTR"
Option Explicit

Public Function sOpenDWG(fn As String) As String
  sOpenDWG = "OPEN " & fn
End Function

Public Function sCloseDWG() As String
  sCloseDWG = "CLOSE"
End Function

Public Function sNewDWG(fn As String) As String
  If R14Acad Then
    sNewDWG = "NEW " & "."
  Else
    sNewDWG = "NEW " & fn & "="
  End If
End Function

Public Function sSaveDWG() As String
  sSaveDWG = "QSAVE"
End Function

Public Function sSaveAS_DWG(fn As String) As String
  If R14Acad Then
    sSaveAS_DWG = "SAVEAS R14 " & fn
  Else
    sSaveAS_DWG = "SAVEAS " & fn
  End If
End Function

Public Function sDXFIN_DWG(fn As String) As String
  sDXFIN_DWG = "DXFIN " & fn
End Function

'Change property of AutoCAD entity
Public Function sChgProp(seltn As String, prpty As String, propVal As String) As String
  sChgProp = "CHPROP " & seltn & Space(2) & prpty & Space(1) & propVal & Space(1)
End Function

Public Function sMakeSlide(fn As String) As String
  sMakeSlide = "MSLIDE" & fn
End Function
Public Function sSetACADcolour(Acadcolour As String) As String
  sSetACADcolour = "Color " & Acadcolour
End Function

Public Function sViewSlide(fn As String) As String
  sViewSlide = "VSLIDE " & fn
End Function

Public Function sWriteVar(varName As String, varValue As String) As String
  sWriteVar = varName & " " & varValue
End Function

Public Function sChgTileMode(mode As Integer) As String
  sChgTileMode = "TILEMODE " & Format(mode, "0")
End Function

Public Function sDisplayUCSICON(state As String) As String
  sDisplayUCSICON = "UCSICON " & state
End Function

Public Function sCreateLayer(LayerName As String, Colour As String, LineType As String) As String
  sCreateLayer = R14Prefix(R14Acad) & "LAYER n " & LayerName & " color " & Colour & " " & LayerName & " ltype " & LineType & " " & LayerName & " "
End Function

Public Function sSetColourByLayer() As String
  sSetColourByLayer = "Color BYLAYER"
End Function

Public Function sSetLineTypeByLayer() As String
  sSetLineTypeByLayer = R14Prefix(R14Acad) & "Linetype s BYLAYER "
End Function

Public Function sSetLayer(LayerName As String) As String
  sSetLayer = R14Prefix(R14Acad) & "LAYER s " & LayerName & " "
End Function


'=====================================================


Public Function sMakeLayer(LayerName As String, Colour As String, LineType As String) As String
  sMakeLayer = R14Prefix(R14Acad) & "LAYER m " & LayerName & " color " & Colour & " " & LayerName & " ltype " & LineType & " " & LayerName & " "
End Function

Public Function sColourLayer(LayerName As String, ColourCode As String) As String
  sColourLayer = R14Prefix(R14Acad) & "LAYER c " & ColourCode & " " & LayerName & " "
End Function

Public Function swrite_mview(Cnr1 As TCoord, Cnr2 As TCoord) As String
  swrite_mview = "MVIEW " & ptStr2D(Cnr1) & " " & ptStr2D(Cnr2)
End Function

Public Function sAcadZoom(mode As String) As String
  sAcadZoom = "ZOOM " & mode
End Function

Public Function swrite_view(name As String) As String
  swrite_view = "VIEW S " & name
End Function

Public Function swrite_line(pt1 As TCoord, pt2 As TCoord) As String
  swrite_line = "LINE " & ptStr(pt1) & " " & ptStr(pt2) & " "
End Function

'Write two segment polyline
Public Function swrite_Pline(width As Double, pt1 As TCoord, pt2 As TCoord) As String
  swrite_Pline = "PLINE " & ptStr2D(pt1) & " w " & Format(width, "#0.0") & " " & Format(width, "#0.0") & " " & ptStr2D(pt2) & " "
End Function

Public Function sStart_Pline(width As Double, pt1 As TCoord, pt2 As TCoord) As String
  sStart_Pline = "PLINE " & ptStr2D(pt1) & " w " & Format(width, "#0.0") & " " & Format(width, "#0.0") & " " & ptStr2D(pt2)
End Function

Public Function sStart_Pline2() As String
  sStart_Pline2 = "PLINE"
End Function

Public Function sEnd_Pline() As String
  sEnd_Pline = "PLINEWID 0"
End Function

Public Function sClose_Pline() As String
  sClose_Pline = "C"
End Function

Public Function swrite_Vertex(pt1 As TCoord) As String
  swrite_Vertex = ptStr2D(pt1)
End Function

Public Function swrite_circle(pt1 As TCoord, diam As Double) As String
  swrite_circle = "CIRCLE " & ptStr2D(pt1) & " D " & Format(diam, "###0.00")
End Function
'Vertical assumed allowed
Public Function sDefTextStyle(sname As String, fontName As String) As String
  sDefTextStyle = R14Prefix(R14Acad) & "STYLE " & sname & " " & fontName & " 0 1 0 n n n"
End Function

'Vertical assumed NOT allowed
Public Function sDefTextStyle1(sname As String, fontName As String) As String
  sDefTextStyle1 = R14Prefix(R14Acad) & "STYLE " & sname & " " & fontName & " 0 1 0 n n"
End Function

Public Function sDefTextStyle2(sname As String, fontName As String, wdth As Double) As String
  sDefTextStyle2 = R14Prefix(R14Acad) & "STYLE " & sname & " " & fontName & " 0 " & Format(wdth, "#0.0") & " 0 n n n"
End Function

Public Function sDefTextStyleFull(sname As String, fontName As String, Hght As Double, wdth As Double, oblq As Double, txBack As String, TxUpside As String, txVert As String) As String
  sDefTextStyleFull = R14Prefix(R14Acad) & "STYLE " & sname & " " & fontName & " " & Format(Hght, "#0.0") & " " & Format(wdth, "#0.00") & " " & Format(oblq, "#0.00") & " " & txBack & " " & TxUpside & " " & txVert
End Function

Public Function swrite_JText(sname As String, TxtAlgn As String, pt1 As TCoord, Hght As Double, rot As Double, txt As String) As String
  swrite_JText = "TEXT STYLE " & sname & " J " & TxtAlgn & Space(1) & ptStr2D(pt1) & " " & Format(Hght, "#0.0") & " " & Format(rot, "#0.0") & " " & txt
End Function

Public Function swrite_Text(sname As String, pt1 As TCoord, Hght As Double, rot As Double, txt As String) As String
  swrite_Text = "TEXT STYLE " & sname & " " & ptStr2D(pt1) & " " & Format(Hght, "#0.0") & " " & Format(rot, "#0.0") & " " & txt
End Function

Public Function swrite_rectangle(fp As Integer, Cnr1 As TCoord, Cnr2 As TCoord) As String
  swrite_rectangle = "RECTANG " & ptStr2D(Cnr1) & " " & ptStr2D(Cnr2)
End Function






