%% Open in FreeMat editor then Type CmdMain to command line

function CmdMain
    scrFile = fopen('DrawBox2.scr', 'w');
	
	ScriptWriter2(scrFile)
	
    fclose(scrFile)
    
    disp('All Done!')
end

function s=pointStr2(x,y)
    s = sprintf('%f,%f',x,y)
end

function drawBox4(scrFile, PtA, Lx, Ly)
	fprintf(scrFile,'PLINE\n')
	fprintf(scrFile,'%s\n',pointStr2(PtA(1),PtA(2)))
    fprintf(scrFile,'@%s\n',pointStr2(0,Ly) )
    fprintf(scrFile,'@%s\n',pointStr2(Lx,0) )   
    fprintf(scrFile,'@%s\n',pointStr2(0,-Ly) )      
	fprintf(scrFile,'C\n')
end 

function  ScriptWriter1(scrFile)
    ptA = [0,0] 
    Ly=16
    Lx=2*Ly
    drawBox4(scrFile,ptA,Lx,Ly)
    fprintf(scrFile,'ZOOM E')
end

function ScriptWriter2(scrFile)
    pt0 = [0,0]
    ptA = [0,0]

    BuildingHeight=2.4
    BuildingWidth=8	
    BuildingLength=2*BuildingWidth
	
	%%Plan
    drawBox4(scrFile,ptA,BuildingLength,BuildingWidth)

    %%Elevation 1
    ptA(2) = ptA(2)-BuildingHeight
    drawBox4(scrFile,ptA,BuildingLength,BuildingHeight)

    %%Elevation 2
    ptA(1) = ptA(1) + BuildingLength
    drawBox4(scrFile,ptA,BuildingWidth,BuildingHeight)

    %%Elevation 3
    ptA(1) = pt0(1) - BuildingWidth
    drawBox4(scrFile,ptA,BuildingWidth,BuildingHeight)

    %%Elevation 4
    ptA(1) = ptA(1) - BuildingLength
    drawBox4(scrFile,ptA,BuildingLength,BuildingHeight)

    %%Section
    ptA(1) = pt0(1) + BuildingLength + BuildingWidth
    drawBox4(scrFile,ptA,BuildingWidth,BuildingHeight)	

    fprintf(scrFile,'ZOOM E')
end
