#NB: python is case sensitive

class clsCoordinate:
    def __init__(self, xx,yy,zz):
        key = 0
        self.x = xx 
        self.y = yy
        self.z = zz

    def ptStr2(self):
        return( format(self.x,'.4f') + ',' + format(self.y,'.4f') )

    def ptStr(self):
        return( format(self.x,'.4f') + ',' + format(self.y,'.4f') + ',' + format(self.z,'.4f') ) 

    def cprint(self):
        print(self.ptStr()+'\n')

    def fprint(self,fp):
        fp.write(self.ptStr()+'\n')
 
def pointStr2(x,y):
    return(format(x, '.4f') + ',' + format(y, '.4f'))

def startLine(fp):
    fp.write('LINE\n')

def StartPline(fp):
    fp.write('PLINE\n')

def closeLines(fp):
    fp.write('C\n')

def drawLine(fp,pt1,pt2):
    fp.write('LINE ' & pt1.ptStr() & ' ' & pt2.ptStr() & ' ')

def drawBox3(fp, ptA, Lx, Ly):
    dist = clsCoordinate(0,0,0)

    StartPline(fp)
    #fp.write(ptA.ptStr2() + '\n')
    #Bottom LH Corner
    ptA.fprint(fp)

    #Top LH Corner
    dist.x = 0
    dist.y = Ly
    fp.write('@' + dist.ptStr2() + '\n' )

    #Top RH Corner
    dist.x = Lx
    dist.y = 0
    fp.write('@' + dist.ptStr2() + '\n' )   

    #Bottom RH Corner
    dist.x = 0
    dist.y = -Ly
    fp.write('@' + dist.ptStr2() + '\n' )      

    closeLines(fp)

def drawBox4(fp, ptA, Lx, Ly):
    StartPline(fp)
    ptA.fprint(fp)
    fp.write('@' + pointStr2(0,Ly) + '\n' )
    fp.write('@' + pointStr2(Lx,0) + '\n' )   
    fp.write('@' + pointStr2(0,-Ly) + '\n' )      

    closeLines(fp)   

def ScriptWriter1(scrFile):
    ptA = clsCoordinate(0,0,0)
    #ptA.x = 5
    #ptA.y=7
    #print(ptA.ptStr2())
    
    Ly=16
    Lx=2*Ly
    drawBox4(scrFile,ptA,Lx,Ly)
    scrFile.write('ZOOM E')
    
def ScriptWriter2(scrFile):
    pt0 = clsCoordinate(0,0,0)
    ptA = clsCoordinate(0,0,0)

    BuildingHeight=2.4
    BuildingWidth=8	
    BuildingLength=2*BuildingWidth
    
    #Plan
    drawBox4(scrFile,ptA,BuildingLength,BuildingWidth)

    #Elevation 1
    ptA.y = ptA.y-BuildingHeight
    drawBox4(scrFile,ptA,BuildingLength,BuildingHeight)

    #Elevation 2
    ptA.x = ptA.x + BuildingLength
    drawBox4(scrFile,ptA,BuildingWidth,BuildingHeight)

    #Elevation 3
    ptA.x = pt0.x - BuildingWidth
    drawBox4(scrFile,ptA,BuildingWidth,BuildingHeight)

    #Elevation 4
    ptA.x = ptA.x - BuildingLength
    drawBox4(scrFile,ptA,BuildingLength,BuildingHeight)

    #Section
    ptA.x = pt0.x + BuildingLength + BuildingWidth
    drawBox4(scrFile,ptA,BuildingWidth,BuildingHeight)	

    scrFile.write('ZOOM E')
 
    
def CmdMain():
    scrFile = open('DrawBox2.scr', 'w')
   
    
    ScriptWriter2(scrFile)
    
    scrFile.close()
    
    print('All Done!')
    
    
#-------------------------
#MAIN
#-------------------------    
    
CmdMain()

# END MAIN
#=========================