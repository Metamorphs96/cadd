// execute in SciLAB then call CmdMain

function CmdMain
    scrFile = mopen('DrawBox2.scr', 'w');
	
	ScriptWriter2(scrFile)
	
    mclose(scrFile)
    
    disp('All Done!')
endfunction

function s=pointStr2(x,y)
    s = string(x) + ',' + string(y)
endfunction

function drawBox4(scrFile, PtA, Lx, Ly)
	mfprintf(scrFile,'PLINE\n')
	mfprintf(scrFile, pointStr2(PtA(1),PtA(2)) + '\n')
    mfprintf(scrFile,'@' + pointStr2(0,Ly) + '\n' )
    mfprintf(scrFile,'@' + pointStr2(Lx,0) + '\n' )   
    mfprintf(scrFile,'@' + pointStr2(0,-Ly) + '\n' )      
	mfprintf(scrFile,'C\n')
endfunction 

function  ScriptWriter1(scrFile)
    ptA = [0,0] 
    Ly=16
    Lx=2*Ly
    drawBox4(scrFile,ptA,Lx,Ly)
    mfprintf(scrFile,'ZOOM E')
endfunction

function ScriptWriter2(scrFile)
    pt0 = [0,0]
    ptA = [0,0]

    BuildingHeight=2.4
    BuildingWidth=8	
    BuildingLength=2*BuildingWidth
	
	//Plan
    drawBox4(scrFile,ptA,BuildingLength,BuildingWidth)

    //Elevation 1
    ptA(2) = ptA(2)-BuildingHeight
    drawBox4(scrFile,ptA,BuildingLength,BuildingHeight)

    //Elevation 2
    ptA(1) = ptA(1) + BuildingLength
    drawBox4(scrFile,ptA,BuildingWidth,BuildingHeight)

    //Elevation 3
    ptA(1) = pt0(1) - BuildingWidth
    drawBox4(scrFile,ptA,BuildingWidth,BuildingHeight)

    //Elevation 4
    ptA(1) = ptA(1) - BuildingLength
    drawBox4(scrFile,ptA,BuildingLength,BuildingHeight)

    //Section
    ptA(1) = pt0(1) + BuildingLength + BuildingWidth
    drawBox4(scrFile,ptA,BuildingWidth,BuildingHeight)	

    mfprintf(scrFile,'ZOOM E')
endfunction
