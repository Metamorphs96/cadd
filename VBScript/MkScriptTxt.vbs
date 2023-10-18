Option Explicit

'Write Text Staggered up and down
Sub TextScriptWriter1(scrFile)
	Dim ptA
	Dim txtHght
	Dim TxtNote
	Dim i,n
	Dim ch
	Dim dx 
	
	R14Acad = True
	AcadR2000 = True
	
	Set ptA = New clsCoordinate	
	ptA.initialise

	txtHght = 2.5

	Call DefTextStyle(scrFile,"Notes", "Romans")
	
	TxtNote="This is Some Example Text"
	
	dx = txtHght*1.5
	n=len(TxtNote)
	For i = 1 to n
		ch = mid(TxtNote,i,1)
		
		ptA.x = ptA.x + dx
		If i mod 2 <> 0 then
			ptA.y = 5
		Else
			ptA.y = 0
		End If
		Call write_Text(scrFile, "Notes", ptA, txtHght, 0, ch)
	Next

	Call zoomExtents(scrFile)
	
End Sub

'Write Text Stepped Upwards
Sub TextScriptWriter2(scrFile)
	Dim ptA
	Dim txtHght
	Dim TxtNote
	Dim i,n
	Dim ch
	Dim dx, dy
	
	R14Acad = True
	AcadR2000 = True
	
	Set ptA = New clsCoordinate	
	ptA.initialise

	txtHght = 2.5

	Call DefTextStyle(scrFile,"Notes", "Romans")
	
	TxtNote="This is Some Example Text"
	
	dx = txtHght*1.5
	dy = 5
	n=len(TxtNote)
	For i = 1 to n
		ch = mid(TxtNote,i,1)
		
		ptA.x = ptA.x + dx
		ptA.y = ptA.y + dy
		Call write_Text(scrFile, "Notes", ptA, txtHght, 0, ch)
	Next

	Call zoomExtents(scrFile)
	
End Sub

'Write Text along Sine curve
'But keep text vertical
Sub TextScriptWriter3(scrFile)
	Dim ptA
	Dim txtHght
	Dim TxtNote
	Dim i,n,k
	Dim ch
	Dim x,y
	Dim dx, dy
	Dim dx2, dy2
	Dim Radius
	Dim arcLength
	Dim baseLength
	Dim scaleY
	
	R14Acad = True
	AcadR2000 = True
	
	Set ptA = New clsCoordinate	
	ptA.initialise

	txtHght = 2.5

	Call DefTextStyle(scrFile,"Notes", "Romans")
	
	TxtNote="This is Some Example Text"
	
	baseLength = 100
	n=len(TxtNote)
	dx=2*pi/n
	dx2 = baseLength/n
	x=0
	y=0
	scaleY = 10
	k=1
	For i=0 to n
		ch = mid(TxtNote,k,1)
		If i = 0 then
			x = 0
		else
			x = x + dx
		End If
		y = sin(x)
		ptA.x = ptA.x + dx2
		ptA.y = y*scaleY
		'call drawPoint(scrFile,ptA)
		Call write_Text(scrFile, "Notes", ptA, txtHght, 0, ch)
		k=k+1
	Next

	Call zoomExtents(scrFile)
	
End Sub

'Write Text along a sine curve
'Align text with tangent at point
Sub TextScriptWriter4(scrFile)
	Dim ptA
	Dim txtHght
	Dim TxtNote
	Dim i,n,k
	Dim ch
	Dim x,y
	Dim dx, dy
	Dim dx2, dy2
	Dim Radius
	Dim arcLength
	Dim baseLength
	Dim scaleY
	Dim theta
	
	R14Acad = True
	AcadR2000 = True
	
	Set ptA = New clsCoordinate	
	ptA.initialise

	txtHght = 2.5

	Call DefTextStyle(scrFile,"Notes", "Romans")
	
	TxtNote="This is Some Example Text"
	
	baseLength = 100
	n=len(TxtNote)
	dx=2*pi/n
	dx2 = baseLength/n
	x=0
	y=0
	scaleY = 10
	k=1
	For i=0 to n
		ch = mid(TxtNote,k,1)
		If i = 0 then
			x = 0
		else
			x = x + dx
		End If
		y = sin(x)
		theta = atn(cos(x)) 'derivative gives slope (m=dy/dx) convert to angle
		ptA.x = ptA.x + dx2
		ptA.y = y*scaleY
		'call drawPoint(scrFile,ptA)
		Call write_Text(scrFile, "Notes", ptA, txtHght, ToDegrees(theta), ch)
		k=k+1
	Next

	Call zoomExtents(scrFile)
	
End Sub

'Write Text around a circle
'f(x)=sqrt(r^2-x^2)
Sub TextScriptWriter5(scrFile)
	Dim ptA
	Dim txtHght
	Dim TxtNote
	Dim i,n,k
	Dim ch
	Dim x,y
	Dim dx, dy
	Dim dx2, dy2
	
	Dim Radius
	Dim arcLength
	Dim scaleY
	Dim theta
	Dim dTheta
	Dim thetaTxt
	
	R14Acad = True
	AcadR2000 = True
	
	Set ptA = New clsCoordinate	
	ptA.initialise

	

	Call DefTextStyle(scrFile,"Notes", "Romans")
	
	TxtNote="This is Some Example Text Around A Circle"
	
	TxtNote = TxtNote & " " 'allow extra space between start and end
	
	Radius = 200
	'call drawCircle(scrFile,ptA,2*Radius)
	
	txtHght = Radius/10
	n=len(TxtNote)
	theta=0
	dTheta=CDbl(2*pi/ n )
	
	Wscript.Echo n
	Wscript.Echo dTheta
	Wscript.Echo ToDegrees(dTheta)
	
	x=0
	y=0
	k=1
	For i=0 to n
		ch = mid(TxtNote,k,1)
		
		If i = 0 then
			theta = pi
		else
			theta = theta - dTheta
		End If
		x = Radius * cos(theta)
		y = Radius * sin(theta)
		
		ptA.x = x
		ptA.y = y
		thetaTxt =  ToDegrees(theta)
		
		If  (0 <= thetaTxt) and (thetaTxt <= 90) then
			thetaTxt = 270 + thetaTxt
		ElseIf  (90 < thetaTxt) and (thetaTxt <= 180) then
			thetaTxt = thetaTxt -90
	    ElseIf  (-90 <= thetaTxt) and (thetaTxt <=0) then
			thetaTxt = 270 + thetaTxt
		ElseIf  (-180 <= thetaTxt) and (thetaTxt < -90) then	
			thetaTxt = thetaTxt -90
		End If
		'call drawPoint(scrFile,ptA)
		If ch <> " " Then
			Call write_JText(scrFile, "Notes", "M", ptA, txtHght, thetaTxt, ch)
		else
			Call write_JText(scrFile, "Notes", "M", ptA, txtHght, thetaTxt, "*")
		End If
		'Call write_JText(scrFile, "Notes", "M", ptA, txtHght, thetaTxt, Cstr(ToDegrees(theta)))
		
		k=k+1
	Next

	Call zoomExtents(scrFile)
	
End Sub