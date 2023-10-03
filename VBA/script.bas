Attribute VB_Name = "script"
Option Explicit

Public Function R14Prefix(R14Acad As Boolean) As String
  If R14Acad Then
    R14Prefix = "-"
  Else
    R14Prefix = ""
  End If
End Function

Public Function R2000Prefix(R2000Acad As Boolean) As String
  If AcadR2000 Then
    R2000Prefix = "-"
  Else
    R2000Prefix = ""
  End If
End Function

Public Sub OpenDWG(fp As Integer, fn As String)
  Print #fp, "OPEN " & fn
End Sub

Public Sub CloseDWG(fp As Integer)
  Print #fp, "CLOSE"
End Sub

Public Sub NewDWG(fp As Integer, fn As String)
  If R14Acad Then
    Print #fp, "NEW " & "."
  ElseIf AcadR2000 Then
    Print #fp, "NEW " & fn
  Else
    Print #fp, "NEW " & fn & "="
  End If
End Sub

Public Sub SaveDWG(fp As Integer, fn As String)
  If fn <> "" Then
    If AcadR2000 Then
      Print #fp, "QSAVE " & fn
    Else
      Print #fp, "QSAVE"
    End If
  Else
    Print #fp, "QSAVE"
  End If
End Sub

Public Sub QuitDWG(fp As Integer)
  Print #fp, "QUIT"
End Sub


Public Sub SaveAS_DWG(fp As Integer, fn As String)
  If R14Acad Then
    Print #fp, "SAVEAS R14 " & fn
  ElseIf AcadR2000 Then
    Print #fp, "SAVEAS 2000 " & fn
  Else
    Print #fp, "SAVEAS " & fn
  End If
End Sub

Public Sub DXFIN_DWG(fp As Integer, fn As String)
  Print #fp, "DXFIN " & fn
End Sub

'Change property of AutoCAD entity
Public Sub ChgProp(fp As Integer, seltn As String, prpty As String, propVal As String)
  Print #fp, "CHPROP " & seltn & Space(2) & prpty & Space(1) & propVal & Space(1)
End Sub

Public Sub edDwgTxt(fp As Integer, pt As TCoord, tx1 As String)
  Print #fp, "CHANGE f " & ptStr2D(pt) & " @5,1 " & Space(7) & tx1
End Sub

Public Sub MakeSlide(fp As Integer, fn As String)
  Print #fp, "MSLIDE" & fn
End Sub

Public Sub SetACADcolour(fp As Integer, Acadcolour As String)
  Print #fp, "Color " & Acadcolour
End Sub

Public Sub ViewSlide(fp As Integer, fn As String)
  Print #fp, "VSLIDE " & fn
End Sub

Public Function ptStr2(x As Double, y As Double) As String
  ptStr2 = Format(x, CADDFmtStr) & "," & Format(y, CADDFmtStr)
End Function

Public Function ptStr3(x As Double, y As Double, z As Double) As String
  ptStr3 = Format(x, CADDFmtStr) & "," & Format(y, CADDFmtStr) & "," & Format(z, CADDFmtStr)
End Function

Public Function ptStr(pt As TCoord) As String
  ptStr = Format(pt.x, CADDFmtStr) & "," & Format(pt.y, CADDFmtStr) & "," & Format(pt.z, CADDFmtStr)
End Function

Public Function ptStr2D(pt As TCoord) As String
  ptStr2D = Format(pt.x, CADDFmtStr) & "," & Format(pt.y, CADDFmtStr)
End Function

Public Sub WriteVar(fp As Integer, varName As String, varValue As String)
  Print #fp, varName & " " & varValue
End Sub

Public Sub ChgTileMode(fp As Integer, mode As Integer)
  Print #fp, "TILEMODE " & Format(mode, "0")
End Sub

Public Sub DisplayUCSICON(fp As Integer, state As String)
  Print #fp, "UCSICON " & state
End Sub

Public Sub CreateNewLayer(fp As Integer, LayerName As String)
  Print #fp, R14Prefix(R14Acad) & "LAYER n " & LayerName & " "
End Sub

Public Sub CreateLayer(fp As Integer, LayerName As String, Colour As String, LineType As String)
  Print #fp, R14Prefix(R14Acad) & "LAYER n " & LayerName & " color " & Colour & " " & LayerName & " ltype " & LineType & " " & LayerName & " "
End Sub

Public Sub SetByLayer(fp As Integer)
  Print #fp, "Color BYLAYER"
  Print #fp, R14Prefix(R14Acad) & "Linetype s BYLAYER "
End Sub

Public Sub SetLayer(fp As Integer, LayerName As String)
  Dim cmd As String
  cmd = R14Prefix(R14Acad) & "LAYER s " & LayerName & " "
  Print #fp, cmd
End Sub

Public Sub MkLayer(fp As Integer, LayerName As String)
  Print #fp, R14Prefix(R14Acad) & "LAYER m " & LayerName & " "
End Sub

Public Sub MakeLayer(fp As Integer, LayerName As String, Colour As String, LineType As String)
  Print #fp, R14Prefix(R14Acad) & "LAYER m " & LayerName & " color " & Colour & " " & LayerName & " ltype " & LineType & " " & LayerName & " "
End Sub

Public Sub ColourLayer(fp As Integer, LayerName As String, ColourCode As String)
  Print #fp, R14Prefix(R14Acad) & "LAYER c " & ColourCode & " " & LayerName & " "
End Sub

Public Sub LayerPStyle(fp As Integer, LayerName As String, PStyleName As String)
  Print #fp, R14Prefix(R14Acad) & "LAYER PSTYLE " & PStyleName & " " & LayerName & " "
End Sub

Public Sub LayerLineType(fp As Integer, LayerName As String, LineType As String)
  Print #fp, R14Prefix(R14Acad) & "LAYER LTYPE " & LineType & " " & LayerName & " "
End Sub

Public Sub CreateLayout(fp As Integer, Lname As String)
  Print #fp, "-LAYOUT New " & Lname
End Sub

Public Sub SetVPLayers(fp As Integer, Lfilter As String)
  Print #fp, "VPLAYER " & Lfilter
End Sub


Public Sub SetLayout(fp As Integer, Lname As String)
  Print #fp, "-LAYOUT Set " & Lname
End Sub


Public Sub write_mview(fp As Integer, Cnr1 As TCoord, Cnr2 As TCoord)
  Print #fp, "MVIEW " & ptStr2D(Cnr1) & " " & ptStr2D(Cnr2)
End Sub

Public Sub AcadZoom(fp As Integer, mode As String)
  Print #fp, "ZOOM " & mode
End Sub

Public Sub write_view(fp As Integer, Cnr1 As TCoord, Cnr2 As TCoord, name As String)
  Print #fp, "ZOOM W " & ptStr(Cnr1) & " " & ptStr(Cnr2)
  Print #fp, "VIEW S " & name
End Sub

Public Sub write_line(fp As Integer, pt1 As TCoord, pt2 As TCoord)
  Print #fp, "LINE " & ptStr(pt1) & " " & ptStr(pt2) & " "
End Sub

'Write two segment polyline
Public Sub write_Pline(fp As Integer, width As Double, pt1 As TCoord, pt2 As TCoord)
  Print #fp, "PLINE " & ptStr2D(pt1) & " w " & Format(width, "#0.0") & " " & Format(width, "#0.0") & " " & ptStr2D(pt2) & " "
End Sub

Public Sub Start_Pline(fp As Integer, width As Double, pt1 As TCoord, pt2 As TCoord)
  Print #fp, "PLINE " & ptStr2D(pt1) & " w " & Format(width, "#0.0") & " " & Format(width, "#0.0") & " " & ptStr2D(pt2)
End Sub

Public Sub Start_Pline2(fp As Integer)
  Print #fp, "PLINE"
End Sub

Public Sub End_Pline(fp As Integer)
  Print #fp, ""
  Print #fp, "PLINEWID 0"
End Sub

Public Sub Close_Pline(fp As Integer)
  Print #fp, "C"
  Print #fp, "PLINEWID 0"
End Sub

Public Sub write_Vertex(fp As Integer, pt1 As TCoord)
  Print #fp, ptStr2D(pt1)
End Sub

Public Sub write_circle(fp As Integer, pt1 As TCoord, diam As Double)
  Print #fp, "CIRCLE " & ptStr(pt1) & " D " & Format(diam, "###0.00")
End Sub
'Vertical assumed allowed
Public Sub DefTextStyle(fp As Integer, sname As String, fontName As String)
  Print #fp, R14Prefix(R14Acad) & "STYLE " & sname & " " & fontName & " 0 1 0 n n n"
End Sub

'Vertical assumed NOT allowed
Public Sub DefTextStyle1(fp As Integer, sname As String, fontName As String)
  Print #fp, R14Prefix(R14Acad) & "STYLE " & sname & " " & fontName & " 0 1 0 n n"
End Sub

Public Sub DefTextStyle2(fp As Integer, sname As String, fontName As String, wdth As Double)
  Print #fp, R14Prefix(R14Acad) & "STYLE " & sname & " " & fontName & " 0 " & Format(wdth, "#0.0") & " 0 n n n"
End Sub

Public Sub DefTextStyleFull(fp As Integer, sname As String, fontName As String, Hght As Double, wdth As Double, oblq As Double, txBack As String, TxUpside As String, txVert As String)
  Print #fp, R14Prefix(R14Acad) & "STYLE " & sname & " " & fontName & " " & Format(Hght, "#0.0") & " " & Format(wdth, "#0.00") & " " & Format(oblq, "#0.00") & " " & txBack & " " & TxUpside & " " & txVert
End Sub

Public Sub write_JText(fp As Integer, sname As String, TxtAlgn As String, pt1 As TCoord, Hght As Double, rot As Double, txt As String)
  Print #fp, "TEXT STYLE " & sname & " J " & TxtAlgn & Space(1) & ptStr2D(pt1) & " " & Format(Hght, "#0.0") & " " & Format(rot, "#0.0") & " " & txt
End Sub

Public Sub write_Text(fp As Integer, sname As String, pt1 As TCoord, Hght As Double, rot As Double, txt As String)
  Print #fp, "TEXT STYLE " & sname & " " & ptStr2D(pt1) & " " & Format(Hght, "#0.0") & " " & Format(rot, "#0.0") & " " & txt
End Sub

Public Sub write_Box(fp As Integer, Cnr1 As TCoord, Cnr2 As TCoord)
  Dim Lx As Double, Ly As Double

  Lx = Cnr2.x - Cnr1.x
  Ly = Cnr2.y - Cnr1.y
  Print #fp, "LINE " & ptStr(Cnr1) & " " & ptStr2(Cnr1.x + Lx, Cnr1.y) _
    & " " & ptStr2(Cnr1.x + Lx, Cnr1.y - Ly) & " " & ptStr2(Cnr1.x, Cnr1.y - Ly) & " C"
End Sub

Public Sub write_rectangle(fp As Integer, Cnr1 As TCoord, Cnr2 As TCoord)
  Print #fp, "RECTANG " & ptStr2D(Cnr1) & " " & ptStr2D(Cnr2)
End Sub

'{Vertically Space Horizontal lines}
Public Sub VSpace(fp As Integer, Spt1 As TCoord, Spt2 As TCoord, yt As Double, Ly As Double)
  Dim pt1 As TCoord  '{ Start[base] point of Line }
  Dim pt2 As TCoord  '{ End[top] point of Line   }
     
  '{ Yt : Limit of lines along Y  axis}
  '{ Ly : Spacing along Y axis}

  '{Start at base and space upwards}
  '{printf("VSpace entered\n");}
  pt1.x = Spt1.x
  pt1.y = Spt1.y
  pt2.x = Spt2.x
  pt2.y = Spt2.y
  
  Do While (pt2.y <= yt)
    Call write_line(fp, pt1, pt2)
    pt1.y = pt1.y + Ly
    pt2.y = pt2.y + Ly
  Loop '{end while}
  
End Sub

'{Horizontally Space Vertical lines}
Public Sub HSpace(fp As Integer, Spt1 As TCoord, Spt2 As TCoord, Xt As Double, Lx As Double)
  Dim pt1 As TCoord  '{ Start[LHS] point of Line }
  Dim pt2 As TCoord  '{ End[RHS] point of Line   }
  
  '{ Xt : Limit of lines along X  axis}
  '{ Lx : Spacing along X axis}
  '{Start at LHS and space to RHS}
  pt1.x = Spt1.x
  pt1.y = Spt1.y
  pt2.x = Spt2.x
  pt2.y = Spt2.y
  
  Do While (pt2.x <= Xt)
    Call write_line(fp, pt1, pt2)
    pt1.x = pt1.x + Lx
    pt2.x = pt2.x + Lx
  Loop '{end while}
  
End Sub

'{Array Circle}
Public Sub ArrayCircle(fp As Integer, Spt1 As TCoord, diam As Double, Lx As Double, Ly As Double, nx As Integer, ny As Integer)
  Dim pt1 As TCoord  '{ Start[LHS] point }
  Dim i As Integer  '{ row counter}
  Dim j As Integer  '{ column counter}
     
  '{ Lx : Spacing along X axis}
  '{ Ly : Spacing along Y axis}

  '{Start at LHS and space to RHS}
  pt1.x = Spt1.x
  pt1.y = Spt1.y
  For i = 0 To ny - 1
    For j = 0 To nx - 1
      Call write_circle(fp, pt1, diam)
      pt1.x = pt1.x + Lx
    Next j
    pt1.x = Spt1.x
    pt1.y = pt1.y + Ly
  Next i
End Sub

'{Array Text}
Public Sub ArrayText(fp As Integer, Spt1 As TCoord, Hght As Double, rot As Double, txt As String, Lx As Double, Ly As Double, nx As Integer, ny As Integer)
  Dim pt1 As TCoord '{ Start[LHS] point }
  Dim i As Integer '{ row counter}
  Dim j As Integer '{ column counter}
 
  '{ Lx : Spacing along X axis}
  '{ Ly : Spacing along Y axis}

  '{Start at LHS and space to RHS}
  pt1.x = Spt1.x
  pt1.y = Spt1.y
  For i = 0 To ny - 1
    For j = 0 To nx - 1
      Call write_JText(fp, "NOTES", "M", pt1, Hght, rot, txt)
      pt1.x = pt1.x + Lx
    Next j
    pt1.x = Spt1.x
    pt1.y = pt1.y + Ly
  Next i
End Sub

'{Increment Text}
Public Sub IncrementText(fp As Integer, Spt1 As TCoord, Hght As Double, rot As Double, Lx As Double, Ly As Double, nx As Integer, ny As Integer, IncType As String, IncDir As Integer)
  Dim pt1 As TCoord '{ Start[LHS] point }
  Dim i As Integer '{ row counter}
  Dim j As Integer '{ column counter}
  Dim k As Integer '{ Text Counter}
  Dim s As String

  '{ Lx : Spacing along X axis}
  '{ Ly : Spacing along Y axis}

  If (IncDir = 1) Then
    k = 1
  ElseIf (IncDir = -1) Then
    If (nx = 1) Then
      k = ny
    ElseIf (ny = 1) Then
      k = nx
    End If
  End If
  
  '{Start at LHS and space to RHS}
  pt1.x = Spt1.x
  pt1.y = Spt1.y
  
  For i = 0 To ny - 1
    For j = 0 To nx - 1
      If (IncType = "N") Then
        s = Format(k, "###0") '.00")
      ElseIf (IncType = "A") Then
        s = Chr(k + 64)
      End If
      
      Call write_JText(fp, "NOTES", "M", pt1, Hght, rot, s)
      pt1.x = pt1.x + Lx
      If (IncDir = 1) Then
        k = k + 1
      ElseIf (IncDir = -1) Then
        k = k - 1
      End If
    Next j
    pt1.x = Spt1.x
    pt1.y = pt1.y + Ly
  Next i
End Sub

'{Array Lines}
Public Sub ArrayLine(fp As Integer, Spt1 As TCoord, Spt2 As TCoord, Lx As Double, Ly As Double, nx As Integer, ny As Integer)
  Dim pt1 As TCoord  '{ Start[LHS] point of Line }
  Dim pt2 As TCoord  '{ End[RHS] point of Line }
  Dim i As Integer  '{ row counter}"
  Dim j As Integer  '{ column counter}"

  '{ Lx : Spacing along X axis}
  '{ Ly : Spacing along Y axis}

  '{Start at LHS and space to RHS}
  pt1.x = Spt1.x
  pt1.y = Spt1.y
  pt2.x = Spt2.x
  pt2.y = Spt2.y
  For i = 0 To ny - 1
    For j = 0 To nx - 1
      Call write_line(fp, pt1, pt2)
      pt1.x = pt1.x + Lx
      pt2.x = pt2.x + Lx
    Next j
    pt1.x = Spt1.x
    pt2.x = Spt2.x
    pt1.y = pt1.y + Ly
    pt2.y = pt2.y + Ly
  Next i
End Sub

Public Sub CreateDimStyle(fp As Integer)
    '{Default Settings for A1 sheet}
    Print #fp, "DIMALT 0"
    Print #fp, "DIMALTD 2"
    Print #fp, "DIMALTF 0.039370"
    Print #fp, "DIMAPOST "
    Print #fp, "DIMASO 1"
    Print #fp, "DIMASZ 1.250000"
    Print #fp, "DIMBLK "
    Print #fp, "DIMBLK1 "
    Print #fp, "DIMBLK2 "
    Print #fp, "DIMCEN -1.750000"
    Print #fp, "DIMCLRD 256"
    Print #fp, "DIMCLRE 256"
    Print #fp, "DIMCLRT 256"
    Print #fp, "DIMGAP 1"
    Print #fp, "DIMDLE 0"
    Print #fp, "DIMDLI 5"
    Print #fp, "DIMEXE 1"
    Print #fp, "DIMEXO 1"
    Print #fp, "DIMLFAC 1"
    Print #fp, "DIMLIM 0"
    Print #fp, "DIMPOST "
    Print #fp, "DIMRND 0"
    Print #fp, "DIMSAH 0"
    Print #fp, "DIMSCALE 1"
    Print #fp, "DIMSE1 0"
    Print #fp, "DIMSE2 0"
    Print #fp, "DIMSHO 1"
    Print #fp, "DIMSOXD 0"
    Print #fp, "DIMTAD 0"
    Print #fp, "DIMTFAC 1"
    Print #fp, "DIMTIH 0"
    Print #fp, "DIMTIX 0"
    Print #fp, "DIMTM 0"
    Print #fp, "DIMTOFL 1"
    Print #fp, "DIMTOH 0"
    Print #fp, "DIMTOL 0"
    Print #fp, "DIMTP 0"
    Print #fp, "DIMTSZ 0"
    Print #fp, "DIMTVP 1.142900"
    Print #fp, "DIMTXT 1.750000"
    Print #fp, "DIMZIN 12"
    Print #fp, "DIMDSEP ."
    Print #fp, "DIM SAVE FULLSIZE EXIT"
End Sub

Public Sub GlobalSetUp(fp As Integer)
  '{Dialogue Settings Should be off(=0) for Scripts }
  'ie.
  'Print #fp, "FILEDIA 0"
  'Print #fp, "CMDDIA 0"
  'Print #fp, "ATTDIA 0"
  'achieve this by calling subroutine InitialiseForScript first

  '{SET Global Variables}
  
  '{System Settings}
  'Print #fp, "APERTURE 2"
  'Print #fp, "PICKBOX 2"

  '{Switch the UCSICON on in both model space and paperspace}
  Print #fp, "TILEMODE 1"
  Print #fp, "UCSICON ON"
  Print #fp, "TILEMODE 0"
  Print #fp, "UCSICON ON"
  Print #fp, "TILEMODE 1"

  '{General Settings}
  Print #fp, "units 2 2 2 0 0 n"
  Print #fp, "LIMMIN -118900.00,-84000.00"
  Print #fp, "LIMMAX 118900.00,84000.00"
  Print #fp, "LTSCALE 1"
  Print #fp, "PSLTSCALE 1"

  '{Drawing Modes}
  Print #fp, "BLIPMODE off"
  Print #fp, "FILLMODE 1"
  Print #fp, "SNAPUNIT 1000.00,1000.0"
  Print #fp, "snap off"
  Print #fp, "GRIDUNIT 5000.00,5000.0"
  Print #fp, "grid off"

  '{Current Entity Modes}
  Print #fp, "linetype s bylayer"
  Print #fp, ""
  Print #fp, "color bylayer"
  Print #fp, "thickness 0"
  Print #fp, "elevation 0"

  Print #fp, "pdmode 3"
  Print #fp, "CHAMFERA 0"
  Print #fp, "CHAMFERB 0"
  Print #fp, "FILLETRAD 0"

  '{Text Styles}
  'Print #fp, "style notes romans 0 1 0 n n n"
  Print #fp, "MIRRTEXT 0"
  Print #fp, "TEXTSIZE 1.75"

  '{Set Current Layer to 0}
  Print #fp, "layer s 0 on 0 "
End Sub

Public Sub InitialiseForScript(fp As Integer, fn As String)
  Print #fp, "FILEDIA 0"
  Print #fp, "CMDDIA 0"
  Print #fp, "NEW " & fn & ".dwg="
  Print #fp, "ATTDIA 0"
  '
  'ATTDIA causes a change to drawing file
  'and a need to save file if placed
  'before NEW
  '
End Sub

Public Sub SetDialogues(fp As Integer, mode As Integer)
  Print #fp, "FILEDIA " & Format(mode, "0")
  Print #fp, "CMDDIA " & Format(mode, "0")
  Print #fp, "ATTDIA " & Format(mode, "0")
End Sub

Public Sub WriteINSERT(fp As Integer, DwgBlockName As String, pt1 As TCoord)
  Dim s As String
  
  s = R14Prefix(R14Acad)
  s = s & "INSERT " & DwgBlockName & " " & ptStr(pt1) & " "
  s = s & "1 1 0"
  Print #fp, s
  
End Sub

Public Sub WriteINSERT2(fp As Integer, DwgBlockName As String, pt1 As TCoord, _
                        mag As Double, rot As Double, NumAtts As Integer)
  Dim s As String, i As Integer
  
  s = R14Prefix(R14Acad)
  s = s & "INSERT " & DwgBlockName & " " & ptStr(pt1) & " "
  s = s & Format(mag, "0.00") & " " & Format(mag, "0.00") & " " & Format(rot, "0.00") '"1 1 0"
  Print #fp, s
  For i = 1 To NumAtts
    Print #fp, "-"
  Next i
End Sub

Public Sub WriteDot(fp As Integer, pt1 As TCoord, diam As Double)
   Print #fp, "donut 0 " & Format(diam, "0.00") & " " & ptStr(pt1) & " "
End Sub

Public Sub NonPlotLayersOFF(fp As Integer)
  Print #fp, R14Prefix(R14Acad) & "layer off vports,pbdr-0,envelope-p1 "
End Sub

Public Sub DisplayTime(fp As Integer)
  Print #fp, "time "
End Sub

Public Sub SetModePaperSpace(fp As Integer)
  Print #fp, "TILEMODE 0"
  Print #fp, "PSPACE"
End Sub

Public Sub cadActivateModelVport(fp As Integer, vport As Integer)
  Print #fp, "MSPACE CVPORT " & Format(vport, "#")
End Sub

Public Sub cadActivatePaperSpace(fp As Integer)
  Print #fp, "PSPACE"
End Sub


Public Sub cadLoadLineTypes(fp As Integer, fn As String)

  Print #fp, R14Prefix(R14Acad) & "LINETYPE LOAD " & fn
  Print #fp,

End Sub

Sub WriteAttrib(fp As Integer, BlockName As String, _
  AttrName As String, AttrValue As String, FmtStr As String)
  
  Print #fp, "-attedit y " & BlockName & " " & AttrName
  Print #fp, "*"
  Print #fp, "C 0,0 841,594" 'A1 sheet
  Print #fp, "v r"
  Print #fp, Format(AttrValue, FmtStr)
  Print #fp, ""
End Sub
