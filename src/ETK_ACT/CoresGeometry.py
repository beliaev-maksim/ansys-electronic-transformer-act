# already done: E, EI, EC = ETD = EQ = ER(with argument), EFD, EP, PQ, P=PH=PT(with  argument), RM, U, UI    

# super class Cores
# all sub classes for core creation will inherit from here basic parameters from GUI and 
# take all functions for primitive creation like createBox, createPolyhedron, move, rename and so on
class Cores():
    CS='Global'
    def __init__(self):
        global step1,step2, coreDimensionsSetup, filledDict
        
        self.oProject = oDesktop.GetActiveProject()
        self.oDesign = self.oProject.GetActiveDesign()     
        self.oEditor = self.oDesign.SetActiveEditor("3D Modeler") 
        
        self.coreDimensions = []
        for i in range(1,9):
            if step1.Properties["coreProperties/coreType/D_" + str(i)].Visible == True:
                self.coreDimensions.append(str(step1.Properties["coreProperties/coreType/D_" + str(i)].Value))
            else:
                self.coreDimensions.append('0')
        coreDimensionsSetup = self.coreDimensions # need in setup       
        # Draw in Base Core                                                                                
        self.DimD1 = float(self.coreDimensions[0])
        self.DimD2 = float(self.coreDimensions[1])                                                    
        self.DimD3 = float(self.coreDimensions[2])                                                    
        self.DimD4 = float(self.coreDimensions[3])                                                    
        self.DimD5 = float(self.coreDimensions[4])                                                    
        self.DimD6 = float(self.coreDimensions[5])                                                    
        self.DimD7 = float(self.coreDimensions[6])                                                    
        self.DimD8 = float(self.coreDimensions[7])                                                    
        
        self.SAng = int(step1.Properties["coreProperties/segAngle"].Value)
        self.NumSegs = int(360/self.SAng)

        self.AirGapS = 0
        self.AirGapC = 0
        self.TAirGap = 0
        
        DoAirgap = bool(step1.Properties["coreProperties/defAirgap"].Value)
        if DoAirgap == True:
            AGValue = float(step1.Properties["coreProperties/defAirgap/airgapValue"].Value)
            AirGapOn = step1.Properties["coreProperties/defAirgap/airgapOn"].Value
            if AirGapOn =='Center Leg':
                self.AirGapC = AGValue/2.0
            elif AirGapOn =='Side Leg':
                self.AirGapS = AGValue/2.0
            else:
                self.TAirGap = AGValue/2.0
                
        self.MNumLayers  = int(step2.Properties["windingProperties/drawWinding/numLayers"].Value)     
        self.MLSpacing   = float(step2.Properties["windingProperties/drawWinding/layerSpacing"].Value)
        self.MTopMargin  = float(step2.Properties["windingProperties/drawWinding/topMargin"].Value)
        self.MSideMargin = float(step2.Properties["windingProperties/drawWinding/sideMargin"].Value)
        self.MBobbinThk  = float(step2.Properties["windingProperties/drawWinding/bobThickness"].Value)
        self.MWdgType    = 1 if step2.Properties["windingProperties/drawWinding/layerType"].Value == 'TopDown' else 2
        self.BobStat     = 1 if bool(step2.Properties["windingProperties/drawWinding/includeBobbin"].Value) == True else 0
        self.WdgStatus = 1 if bool(step2.Properties["windingProperties/drawWinding"].Value) == True else 0
        self.WdgParDict  = {}
        
        if step2.Properties["windingProperties/drawWinding/conductorType"].Value == 'Rectangular':
            self.MCondType   = 1 
            table = step2.Properties["windingProperties/drawWinding/conductorType/tableLayers"]
            rowNum = table.RowCount
            for rowIndex in range(0, rowNum):
                self.WdgParDict[rowIndex+1] = []
                self.WdgParDict[rowIndex+1].append(float(table.Value["windingProperties/drawWinding/conductorType/tableLayers/conductorWidth"][rowIndex]))
                self.WdgParDict[rowIndex+1].append(float(table.Value["windingProperties/drawWinding/conductorType/tableLayers/conductorHeight"][rowIndex]))
                self.WdgParDict[rowIndex+1].append(int(table.Value["windingProperties/drawWinding/conductorType/tableLayers/turnsNumber"][rowIndex]))
                self.WdgParDict[rowIndex+1].append(float(table.Value["windingProperties/drawWinding/conductorType/tableLayers/insulationThick"][rowIndex]))
        else: 
            self.MCondType   = 2
            table = step2.Properties["windingProperties/drawWinding/conductorType/tableLayersCircles"]
            rowNum = table.RowCount
            for rowIndex in range(0, rowNum):
                self.WdgParDict[rowIndex+1] = []
                self.WdgParDict[rowIndex+1].append(float(table.Value["windingProperties/drawWinding/conductorType/tableLayersCircles/conductorDiameter"][rowIndex])  )
                self.WdgParDict[rowIndex+1].append(int(table.Value["windingProperties/drawWinding/conductorType/tableLayersCircles/layerSegNumber"][rowIndex])  )
                self.WdgParDict[rowIndex+1].append(int(table.Value["windingProperties/drawWinding/conductorType/tableLayersCircles/turnsNumber"][rowIndex]))
                self.WdgParDict[rowIndex+1].append(float(table.Value["windingProperties/drawWinding/conductorType/tableLayersCircles/insulationThick"][rowIndex]))
        filledDict= self.WdgParDict
         
    def createPolyline(self, points, segments, name, closed = True,color = '(165 42 42)'):
        self.oEditor.CreatePolyline(
            [
                "NAME:PolylineParameters",
                "IsPolylineCovered:="	, True,
                "IsPolylineClosed:="	, closed,
                points,
                segments,
                [
                    "NAME:PolylineXSection",
                    "XSectionType:="	, "None",
                    "XSectionOrient:="	, "Auto",
                    "XSectionWidth:="	, "0mm",
                    "XSectionTopWidth:="	, "0mm",
                    "XSectionHeight:="	, "0mm",
                    "XSectionNumSegments:="	, "0",
                    "XSectionBendType:="	, "Corner"
                ]
            ], 
            [
                "NAME:Attributes",
                "Name:="		, name,
                "Flags:="		, "",
                "Color:="		, color,
                "Transparency:="	, 0,
                "PartCoordinateSystem:=", Cores.CS,
                "UDMId:="		, "",
                "MaterialValue:="	, '"ferrite"',
                "SurfaceMaterialValue:=", "\"\"",
                "SolveInside:="		, True,
                "IsMaterialEditable:="	, True,
                "UseMaterialAppearance:=", False
            ])
        
    def move(self, selection, X, Y, Z):
        self.oEditor.Move(
            [
                "NAME:Selections",
                "Selections:="		, selection,
                "NewPartsModelFlag:="	, "Model"
            ], 
            [
                "NAME:TranslateParameters",
                "TranslateVectorX:="	, str(X) + 'mm',
                "TranslateVectorY:="	, str(Y) + 'mm',
                "TranslateVectorZ:="	, str(Z) + 'mm'
            ])
        
    def createBox(self,XPosition, YPosition, ZPosition,XSize,YSize,ZSize, Name, color = '(165 42 42)'):
        self.oEditor.CreateBox(
            [
                "NAME:BoxParameters",
                "XPosition:="		, str(XPosition) + 'mm',
                "YPosition:="		, str(YPosition) + 'mm',
                "ZPosition:="		, str(ZPosition) + 'mm',
                "XSize:="		    , str(XSize) + 'mm',
                "YSize:="		    , str(YSize) + 'mm',
                "ZSize:="		    , str(ZSize) + 'mm'
            ], 
            [
                "NAME:Attributes",
                "Name:="		, Name,
                "Flags:="		, "",
                "Color:="		, color,
                "Transparency:="	, 0,
                "PartCoordinateSystem:=", Cores.CS,
                "UDMId:="		, "",
                "MaterialValue:="	, '"ferrite"',
                "SurfaceMaterialValue:=", "\"\"",
                "SolveInside:="		, True,
                "IsMaterialEditable:="	, True,
                "UseMaterialAppearance:=", False
            ])
            
    def createRectangle(self,XPosition, YPosition, ZPosition,width,height, name, covered = True):
        self.oEditor.CreateRectangle(
                    [
                        "NAME:RectangleParameters",
                        "IsCovered:="		, covered,
                        "XStart:="		, str(XPosition) + 'mm',
                        "YStart:="		, str(YPosition) + 'mm',
                        "ZStart:="		, str(ZPosition) + 'mm',
                        "Width:="		, str(width)+'mm',
                        "Height:="		, str(height)+'mm',
                        "WhichAxis:="		, 'Y'
                    ], 
                    [
                        "NAME:Attributes",
                        "Name:="		, name,
                        "Flags:="		, "",
                        "Color:="		, "(143 175 143)",
                        "Transparency:="	, 0,
                        "PartCoordinateSystem:=", Cores.CS,
                        "UDMId:="		, "",
                        "MaterialValue:="	, "\"vacuum\"",
                        "SurfaceMaterialValue:=", "\"\"",
                        "SolveInside:="		, True,
                        "IsMaterialEditable:="	, True,
                        "UseMaterialAppearance:=", False
                    ])
            
    def createPolyhedron(self,XPosition, YPosition, ZPosition,XSize,YSize,ZSize,Height,segNum, Name, color = '(165 42 42)'):
        self.oEditor.CreateRegularPolyhedron(
            [
                "NAME:PolyhedronParameters",
                "XCenter:="		, str(XPosition) + 'mm',
                "YCenter:="		, str(YPosition) + 'mm',
                "ZCenter:="		, str(ZPosition) + 'mm',
                "XStart:="		, str(XSize) + 'mm',
                "YStart:="		, str(YSize) + 'mm',
                "ZStart:="		, str(ZSize) + 'mm',
                "Height:="		, str(Height) + 'mm',
                "NumSides:="		, str(segNum),
                "WhichAxis:="		, "Z"
            ], 
            [
                "NAME:Attributes",
                "Name:="		, Name,
                "Flags:="		, "",
                "Color:="		, color,
                "Transparency:="	, 0,
                "PartCoordinateSystem:=", Cores.CS,
                "UDMId:="		, "",
                "MaterialValue:="	, '"ferrite"',
                "SurfaceMaterialValue:=", "\"\"",
                "SolveInside:="		, True,
                "IsMaterialEditable:="	, True,
                "UseMaterialAppearance:=", False
            ])

    def createPolygon(self, XPosition, YPosition, ZPosition,XSize,YSize,ZSize,segNum, Name, axis, covered = True ):
        self.oEditor.CreateRegularPolygon(
                [
                    "NAME:RegularPolygonParameters",
                    "IsCovered:="		, covered,
                    "XCenter:="		, str(XPosition) + 'mm',
                    "YCenter:="		, str(YPosition) + 'mm',
                    "ZCenter:="		, str(ZPosition) + 'mm',
                    "XStart:="		, str(XSize) + 'mm',
                    "YStart:="		, str(YSize) + 'mm',
                    "ZStart:="		, str(ZSize) + 'mm',
                    "NumSides:="	, str(segNum),
                    "WhichAxis:="		, axis
                ], 
                [
                    "NAME:Attributes",
                    "Name:="		, Name,
                    "Flags:="		, "",
                    "Color:="		, "(143 175 143)",
                    "Transparency:="	, 0,
                    "PartCoordinateSystem:=", Cores.CS,
                    "UDMId:="		, "",
                    "MaterialValue:="	, "\"vacuum\"",
                    "SurfaceMaterialValue:=", "\"\"",
                    "SolveInside:="		, True,
                    "IsMaterialEditable:="	, True,
                    "UseMaterialAppearance:=", False
                ])
                
    def createCircle(self, XPosition, YPosition, ZPosition,Diameter,segNum, Name, axis, covered = True ):
        self.oEditor.CreateCircle(
            [
                "NAME:CircleParameters",
                "IsCovered:="		, covered,
                "XCenter:="		, str(XPosition) + 'mm',
                "YCenter:="		, str(YPosition) + 'mm',
                "ZCenter:="		, str(ZPosition) + 'mm',
                "Radius:="		, str(Diameter/2) + 'mm',
                "WhichAxis:="		, axis,
                "NumSegments:="		, str(segNum)
            ], 
            [
                "NAME:Attributes",
                "Name:="		, Name,
                "Flags:="		, "",
                "Color:="		, "(143 175 143)",
                "Transparency:="	, 0,
                "PartCoordinateSystem:=", Cores.CS,
                "UDMId:="		, "",
                "MaterialValue:="	, "\"vacuum\"",
                "SurfaceMaterialValue:=", "\"\"",
                "SolveInside:="		, True,
                "IsMaterialEditable:="	, True,
                "UseMaterialAppearance:=", False
            ])
    
    def fillet(self,name,Radius, X,Y,Z):
        ID = [self.oEditor.GetEdgeByPosition(["NAME:EdgeParameters",
        "BodyName:=", name, 
        "XPosition:=", str(X) + 'mm', 
        "YPosition:=", str(Y) + 'mm', 
        "ZPosition:=", str(Z) + 'mm'])]
        self.oEditor.Fillet(
        [
            "NAME:Selections",
            "Selections:="		, name,
            "NewPartsModelFlag:="	, "Model"
        ], 
        [
            "NAME:Parameters",
            [
                "NAME:FilletParameters",
                "Edges:="		, ID,
                "Vertices:="		, [],
                "Radius:="		, str(Radius) + 'mm',
                "Setback:="		, "0mm"
            ]
        ])

    def rename(self,oldName,newName):
        self.oEditor.ChangeProperty(
            ["NAME:AllTabs",
                ["NAME:Geometry3DAttributeTab",
                    ["NAME:PropServers", oldName],
                    ["NAME:ChangedProps",
                        ["NAME:Name","Value:=", newName]
                    ]
                ]
            ])  

    def unite(self,selection):
        self.oEditor.Unite(
            ["NAME:Selections","Selections:=", selection], 
            ["NAME:UniteParameters","KeepOriginals:=", False])   
      
    def subtract(self, selection, tool, keep = False ):
        self.oEditor.Subtract(
            [
                "NAME:Selections",
                "Blank Parts:="		, selection,
                "Tool Parts:="		, tool
            ], 
            ["NAME:SubtractParameters","KeepOriginals:=", keep])

    def sweepAlongVec(self, selection, VecZ):
        self.oEditor.SweepAlongVector(
            [
                "NAME:Selections",
                "Selections:="		, selection,
                "NewPartsModelFlag:="	, "Model"
            ], 
            [
                "NAME:VectorSweepParameters",
                "DraftAngle:="		, "0deg",
                "DraftType:="		, "Round",
                "CheckFaceFaceIntersection:=", False,
                "SweepVectorX:="	, "0mm",
                "SweepVectorY:="	, "0mm",
                "SweepVectorZ:="	, str(VecZ) + 'mm'
            ])
     
    def sweepAlongPath(self,selection):
        self.oEditor.SweepAlongPath(
            [
                "NAME:Selections",
                "Selections:="		, selection,
                "NewPartsModelFlag:="	, "Model"
            ], 
            [
                "NAME:PathSweepParameters",
                "DraftAngle:="		, "0deg",
                "DraftType:="		, "Round",
                "CheckFaceFaceIntersection:=", False,
                "TwistAngle:="		, "0deg"
            ])
     
    def duplicateMirror(self,selection, X, Y, Z):
        self.oEditor.DuplicateMirror(
            [
                "NAME:Selections",
                "Selections:="		, selection,
                "NewPartsModelFlag:="	, "Model"
            ], 
            [
                "NAME:DuplicateToMirrorParameters",
                "DuplicateMirrorBaseX:=", "0mm",
                "DuplicateMirrorBaseY:=", "0mm",
                "DuplicateMirrorBaseZ:=", "0mm",
                "DuplicateMirrorNormalX:=", str(X) + 'mm',
                "DuplicateMirrorNormalY:=", str(Y) + 'mm',
                "DuplicateMirrorNormalZ:=", str(Z) + 'mm'
            ], 
            [
                "NAME:Options",
                "DuplicateAssignments:=", False
            ], 
            [
                "CreateGroupsForNewObjects:=", False
            ])
    
    def pointsSegmentsGenerator(self,list):
        for i in range(list.count([])):
            list.remove([])

        pointsArray = ["NAME:PolylinePoints"]
        segmentsArray = ["NAME:PolylineSegments"]
        for i in range(len(list)):
            pointsArray.append([
                        "NAME:PLPoint",
                        "X:="			, str(list[i][0]) + 'mm',
                        "Y:="			, str(list[i][1]) + 'mm',
                        "Z:="			, str(list[i][2]) + 'mm'
                    ])
            if i != len(list)-1:        
                segmentsArray.append([
                            "NAME:PLSegment",
                            "SegmentType:="		, "Line",
                            "StartIndex:="		, i,
                            "NoOfPoints:="		, 2
                        ])
        return pointsArray, segmentsArray
             
class ECore(Cores):
    def initCore(self):
        self.MECoreLength    = self.DimD1
        self.MECoreWidth     = self.DimD6
        self.MECoreHeight    = self.DimD4
        self.MSLegWidth      = (self.DimD1-self.DimD2)/2
        self.MCLegWidth      = self.DimD3
        self.MSlotDepth      = self.DimD5

    def DrawWdg(self,DimD2,DimD3,DimD5,DimD6,SAng,ECoreDimD8 = 0):
        MNumLayers  = self.MNumLayers 
        MLSpacing   = self.MLSpacing  
        MTopMargin  = self.MTopMargin 
        MSideMargin = self.MSideMargin
        MBobbinThk  = self.MBobbinThk 
        MWdgType    = self.MWdgType   
        MCondType   = self.MCondType  
        BobStat     = self.BobStat    
        WdgParDict  = self.WdgParDict
        MSlotWidth  = (DimD2-DimD3)/2.0
        MSlotHeight = DimD5*2
        if BobStat > 0:
            self.DrawBobbin(MSlotHeight-2*MTopMargin,(DimD2/2.0)-MSideMargin,(DimD3/2.0)+MSideMargin+MBobbinThk,(DimD6/2.0)+MSideMargin+MBobbinThk,MBobbinThk,MSideMargin,SAng, ECoreDimD8)
        
        if MWdgType == 1:
            MTDx = MTopMargin+MBobbinThk
            for MAx in WdgParDict:
                for MBx in range(0,int(WdgParDict[MAx][2])):
                    MRecSzX = DimD3 + (2*(MSideMargin+MBobbinThk)) + ((2*MBx+1)*WdgParDict[MAx][0]) + (2*MBx*2*WdgParDict[MAx][3]) + 2*WdgParDict[MAx][3]
                    MRecSzY = DimD6 + (2*(MSideMargin+MBobbinThk)) + ((2*MBx+1)*WdgParDict[MAx][0]) + (2*MBx*2*WdgParDict[MAx][3]) + 2*WdgParDict[MAx][3]
                    if MCondType == 1:
                        MRecSzZ = -WdgParDict[MAx][1]/2.0
                        self.CreateSingleTurn2(MRecSzX,MRecSzY,MRecSzZ,MCondType,WdgParDict[MAx][0],WdgParDict[MAx][1],DimD5-MTDx-(WdgParDict[MAx][3]),MAx,MBx,(MRecSzX-DimD3)/2,SAng)
                    else:
                        MRecSzZ = -WdgParDict[MAx][0]/2.0
                        self.CreateSingleTurn(MRecSzX,MRecSzY,MRecSzZ,MCondType,WdgParDict[MAx][0],WdgParDict[MAx][1],DimD5-MTDx-(WdgParDict[MAx][3]),MAx,MBx,(MRecSzX-DimD3)/2,SAng)
                if MCondType == 1:
                    MTDx = MTDx + MLSpacing + WdgParDict[MAx][1] + 2*WdgParDict[MAx][3]
                else:
                    MTDx = MTDx + MLSpacing + WdgParDict[MAx][0] + 2*WdgParDict[MAx][3]

        else:
            MTDx = MSideMargin+MBobbinThk 
            for MAx in WdgParDict.keys():
                for MBx in range(0,int(WdgParDict[MAx][2])):
                    MRecSzX = DimD3 + (2*(MTDx+(WdgParDict[MAx][0]/2.0)))+ 2*WdgParDict[MAx][3]
                    MRecSzY = DimD6 + (2*(MTDx+(WdgParDict[MAx][0]/2.0)))+ 2*WdgParDict[MAx][3]
                    if MCondType == 1:
                        MRecSzZ = -WdgParDict[MAx][1]/2.0
                        self.CreateSingleTurn2(MRecSzX,MRecSzY,MRecSzZ,MCondType,WdgParDict[MAx][0],WdgParDict[MAx][1],DimD5-MTopMargin-MBobbinThk-(WdgParDict[MAx][3])-MBx*(2*WdgParDict[MAx][3]+WdgParDict[MAx][1]),MAx,MBx,(MRecSzX-DimD3)/2,SAng)
                    else:
                        MRecSzZ = -WdgParDict[MAx][0]/2.0
                        self.CreateSingleTurn(MRecSzX,MRecSzY,MRecSzZ,MCondType,WdgParDict[MAx][0],WdgParDict[MAx][1],DimD5-MTopMargin-MBobbinThk-(WdgParDict[MAx][3])-MBx*(2*WdgParDict[MAx][3]+WdgParDict[MAx][0]),MAx,MBx,(MRecSzX-DimD3)/2,SAng)
                MTDx = MTDx + MLSpacing + WdgParDict[MAx][0] + 2*WdgParDict[MAx][3]

    def CreateSingleTurn(self,PathX,PathY,PathZ,ProfTyp,ProfAX,ProfZ,ZPos,LayNum,TurnNum,FRad,SAng):
        NumSegs = int(90/SAng)
        SegAng = math.pi/(NumSegs*2)
        
        vertices = []
        for i in range(1000):
            vertices.append([])
        
        for cnt2 in range(0,NumSegs-1):
            vertices[2+cnt2] = [-(PathX/2.0) +FRad - FRad*math.sin(SegAng*(cnt2+1))]
            vertices[2+cnt2].append(-(PathY/2.0) +FRad - FRad*math.cos(SegAng*(cnt2+1)))
            vertices[2+cnt2].append(PathZ)
        
        for cnt3 in range(0,NumSegs-1):
            vertices[3+NumSegs+cnt3] = [-(PathX/2.0) +FRad - FRad*math.cos(SegAng*(cnt3+1))]
            vertices[3+NumSegs+cnt3].append((PathY/2.0) -FRad + FRad*math.sin(SegAng*(cnt3+1)))
            vertices[3+NumSegs+cnt3].append(PathZ)
            
        for cnt4 in range(0,NumSegs-1):
            vertices[4+2*NumSegs+cnt4] = [(PathX/2.0) -FRad + FRad*math.sin(SegAng*(cnt4+1))]
            vertices[4+2*NumSegs+cnt4].append((PathY/2.0) -FRad + FRad*math.cos(SegAng*(cnt4+1)))
            vertices[4+2*NumSegs+cnt4].append(PathZ)
            
        for cnt5 in range(0,NumSegs-1):
            vertices[5+3*NumSegs+cnt5] = [(PathX/2.0) -FRad + FRad*math.cos(SegAng*(cnt5+1))]
            vertices[5+3*NumSegs+cnt5].append(-(PathY/2.0) +FRad - FRad*math.sin(SegAng*(cnt5+1)))
            vertices[5+3*NumSegs+cnt5].append(PathZ)
            
        vertices[0] = [(PathX/2.0) -FRad]
        vertices[0].append(-(PathY/2.0))
        vertices[0].append(PathZ)
        vertices[1] = [-(PathX/2.0) +FRad]
        vertices[1].append(-(PathY/2.0))
        vertices[1].append(PathZ)
        vertices[1+NumSegs] = [-(PathX/2.0)]
        vertices[1+NumSegs].append(-(PathY/2.0) +FRad)
        vertices[1+NumSegs].append(PathZ)
        vertices[2+NumSegs] = [-(PathX/2.0)]
        vertices[2+NumSegs].append((PathY/2.0) -FRad)
        vertices[2+NumSegs].append(PathZ)
        vertices[2+2*NumSegs] = [-(PathX/2.0) + FRad]
        vertices[2+2*NumSegs].append((PathY/2.0))
        vertices[2+2*NumSegs].append(PathZ)
        vertices[3+2*NumSegs] = [(PathX/2.0)-FRad]
        vertices[3+2*NumSegs].append((PathY/2.0))
        vertices[3+2*NumSegs].append(PathZ)
        vertices[3+3*NumSegs] = [(PathX/2.0)]
        vertices[3+3*NumSegs].append((PathY/2.0)-FRad)
        vertices[3+3*NumSegs].append(PathZ)
        vertices[4+3*NumSegs] = [(PathX/2.0)]
        vertices[4+3*NumSegs].append(-(PathY/2.0)+FRad)
        vertices[4+3*NumSegs].append(PathZ)
        vertices[4+4*NumSegs] = [(PathX/2.0) -FRad]
        vertices[4+4*NumSegs].append(-(PathY/2.0))
        vertices[4+4*NumSegs].append(PathZ)
           
        pointsArray,segmentsArray = self.pointsSegmentsGenerator(vertices)       
        self.createPolyline( pointsArray, segmentsArray, "Polyline1", False)
               
        if ProfTyp == 1:    
            self.createRectangle(-((PathX/2)+(ProfAX/2)), 0, (PathZ)-(ProfZ/2),
                                 ProfAX,ProfZ, 'sweepProfile')
        else:
            self.createCircle(-(PathX/2), 0, PathZ,
                               ProfAX, ProfZ, 'sweepProfile', 'Y')
            
        self.rename( "Polyline1", 'Tool%s_%s'%(LayNum,TurnNum+1))
        
        self.sweepAlongPath('sweepProfile,Tool%s_%s'%(LayNum,TurnNum+1))

        self.move('sweepProfile',0,0,ZPos)
        self.rename( 'sweepProfile', 'Layer%s_%s'%(LayNum,TurnNum+1))
            
    def CreateSingleTurn2(self,PathX,PathY,PathZ,ProfTyp,ProfAX,ProfZ,ZPos,LayNum,TurnNum,FRad,SAng):

        NumSegs = int(90/SAng)
        SegAng = math.pi/(NumSegs*2)
         
        vertices1 = []
        for i in range(1000):
            vertices1.append([]) 
        
        PathX1 = PathX+ProfAX
        PathY1 = PathY+ProfAX
        FRad1 = FRad+(ProfAX/2.0)
             
        for cnt2 in range(0,NumSegs-1):
            vertices1[2+cnt2] =[ -(PathX1/2.0) +FRad1 - FRad1*math.sin(SegAng*(cnt2+1))]
            vertices1[2+cnt2].append(-(PathY1/2.0) +FRad1 - FRad1*math.cos(SegAng*(cnt2+1)))
            vertices1[2+cnt2].append(0)
        
        for cnt3 in range(0,NumSegs-1):
            vertices1[3+NumSegs+cnt3]= [-(PathX1/2.0) +FRad1 - FRad1*math.cos(SegAng*(cnt3+1))]
            vertices1[3+NumSegs+cnt3].append((PathY1/2.0) -FRad1 + FRad1*math.sin(SegAng*(cnt3+1)))
            vertices1[3+NumSegs+cnt3].append(0)
            
        for cnt4 in range(0,NumSegs-1):
            vertices1[4+2*NumSegs+cnt4]= [(PathX1/2.0) -FRad1 + FRad1*math.sin(SegAng*(cnt4+1))]
            vertices1[4+2*NumSegs+cnt4].append((PathY1/2.0) -FRad1 + FRad1*math.cos(SegAng*(cnt4+1)))
            vertices1[4+2*NumSegs+cnt4].append(0)
            
        for cnt5 in range(0,NumSegs-1):
            vertices1[5+3*NumSegs+cnt5]= [(PathX1/2.0) -FRad1 + FRad1*math.cos(SegAng*(cnt5+1))]
            vertices1[5+3*NumSegs+cnt5].append(-(PathY1/2.0) +FRad1 - FRad1*math.sin(SegAng*(cnt5+1)))
            vertices1[5+3*NumSegs+cnt5].append(0)
            
        vertices1[0]= [(PathX1/2.0) -FRad1]
        vertices1[0].append(-(PathY1/2.0))
        vertices1[0].append(0)
        vertices1[1]= [-(PathX1/2.0) +FRad1]
        vertices1[1].append(-(PathY1/2.0))
        vertices1[1].append(0)
        vertices1[1+NumSegs]= [-(PathX1/2.0)]
        vertices1[1+NumSegs].append(-(PathY1/2.0) +FRad1)
        vertices1[1+NumSegs].append(0)
        vertices1[2+NumSegs]= [-(PathX1/2.0)]
        vertices1[2+NumSegs].append((PathY1/2.0) -FRad1)
        vertices1[2+NumSegs].append(0)
        vertices1[2+2*NumSegs]= [-(PathX1/2.0) + FRad1]
        vertices1[2+2*NumSegs].append((PathY1/2.0))
        vertices1[2+2*NumSegs].append(0)
        vertices1[3+2*NumSegs]= [(PathX1/2.0)-FRad1]
        vertices1[3+2*NumSegs].append((PathY1/2.0))
        vertices1[3+2*NumSegs].append(0)
        vertices1[3+3*NumSegs]= [(PathX1/2.0)]
        vertices1[3+3*NumSegs].append((PathY1/2.0)-FRad1)
        vertices1[3+3*NumSegs].append(0)
        vertices1[4+3*NumSegs]= [(PathX1/2.0)]
        vertices1[4+3*NumSegs].append(-(PathY1/2.0)+FRad1)
        vertices1[4+3*NumSegs].append(0)
        vertices1[4+4*NumSegs]= [(PathX1/2.0) -FRad1]
        vertices1[4+4*NumSegs].append(-(PathY1/2.0))
        vertices1[4+4*NumSegs].append(0)                                                          
            
        vertices2 = []
        for i in range(1000):
            vertices2.append([]) 
            
        PathX2 = PathX-ProfAX
        PathY2 = PathY-ProfAX
        FRad2 = FRad-(ProfAX/2.0)
            
        for cnt2 in range(0,NumSegs-1):
            vertices2[2+cnt2] = [ -(PathX2/2.0) +FRad2 - FRad2*math.sin(SegAng*(cnt2+1))]
            vertices2[2+cnt2].append(-(PathY2/2.0) +FRad2 - FRad2*math.cos(SegAng*(cnt2+1)))
            vertices2[2+cnt2].append(0)
        
        for cnt3 in range(0,NumSegs-1):
            vertices2[3+NumSegs+cnt3] = [ -(PathX2/2.0) +FRad2 - FRad2*math.cos(SegAng*(cnt3+1))]
            vertices2[3+NumSegs+cnt3].append((PathY2/2.0) -FRad2 + FRad2*math.sin(SegAng*(cnt3+1)))
            vertices2[3+NumSegs+cnt3].append(0)
            
        for cnt4 in range(0,NumSegs-1):
            vertices2[4+2*NumSegs+cnt4] = [ (PathX2/2.0) -FRad2 + FRad2*math.sin(SegAng*(cnt4+1))]
            vertices2[4+2*NumSegs+cnt4].append((PathY2/2.0) -FRad2 + FRad2*math.cos(SegAng*(cnt4+1)))
            vertices2[4+2*NumSegs+cnt4].append(0)
            
        for cnt5 in range(0,NumSegs-1):
            vertices2[5+3*NumSegs+cnt5] = [ (PathX2/2.0) -FRad2 + FRad2*math.cos(SegAng*(cnt5+1))]
            vertices2[5+3*NumSegs+cnt5].append(-(PathY2/2.0) +FRad2 - FRad2*math.sin(SegAng*(cnt5+1)))
            vertices2[5+3*NumSegs+cnt5].append(0)
            
        vertices2[0] = [ (PathX2/2.0) -FRad2]
        vertices2[0].append(-(PathY2/2.0))
        vertices2[0].append(0)
        vertices2[1] = [ -(PathX2/2.0) +FRad2]
        vertices2[1].append(-(PathY2/2.0))
        vertices2[1].append(0)
        vertices2[1+NumSegs] = [ -(PathX2/2.0)]
        vertices2[1+NumSegs].append(-(PathY2/2.0) +FRad2)
        vertices2[1+NumSegs].append(0)
        vertices2[2+NumSegs] = [ -(PathX2/2.0)]
        vertices2[2+NumSegs].append((PathY2/2.0) -FRad2)
        vertices2[2+NumSegs].append(0)
        vertices2[2+2*NumSegs] = [ -(PathX2/2.0) + FRad2]
        vertices2[2+2*NumSegs].append((PathY2/2.0))
        vertices2[2+2*NumSegs].append(0)
        vertices2[3+2*NumSegs] = [ (PathX2/2.0)-FRad2]
        vertices2[3+2*NumSegs].append((PathY2/2.0))
        vertices2[3+2*NumSegs].append(0)
        vertices2[3+3*NumSegs] = [ (PathX2/2.0)]
        vertices2[3+3*NumSegs].append((PathY2/2.0)-FRad2)
        vertices2[3+3*NumSegs].append(0)
        vertices2[4+3*NumSegs] = [ (PathX2/2.0)]
        vertices2[4+3*NumSegs].append(-(PathY2/2.0)+FRad2)
        vertices2[4+3*NumSegs].append(0)
        vertices2[4+4*NumSegs] = [ (PathX2/2.0) -FRad2]
        vertices2[4+4*NumSegs].append(-(PathY2/2.0))
        vertices2[4+4*NumSegs].append(0)
            
        pointsArray,segmentsArray = self.pointsSegmentsGenerator(vertices1)
        pointsArray2,segmentsArray2 = self.pointsSegmentsGenerator(vertices2)
        
        self.createPolyline( pointsArray, segmentsArray, "Polyline1")
        self.createPolyline( pointsArray2, segmentsArray2, 'Tool%s_%s'%(LayNum,TurnNum+1))
        
        self.subtract( "Polyline1", 'Tool%s_%s'%(LayNum,TurnNum+1))
        self.sweepAlongVec("Polyline1", PathZ*2) 
            
        self.move( 'Polyline1',0,0, ZPos)
        self.rename( "Polyline1", 'Layer%s_%s'%(LayNum,TurnNum+1))
       
    def DrawBobbin(self,Hb,Db1,Db2,Db3,Tb,BRad,SAng,ECoreDimD8): 
        BRadE = (Db1-Db2)+BRad
        
        self.createBox(-Db1+ECoreDimD8,-(Db1-Db2+Db3),-Tb+(Hb/2.0),2*Db1-2*ECoreDimD8, 2*(Db1-Db2+Db3), Tb, 'Bobbin', '(255,248,157)')
        self.createBox(-Db1+ECoreDimD8,-(Db1-Db2+Db3),-(Hb/2.0), 2*Db1-2*ECoreDimD8, 2*(Db1-Db2+Db3), Tb, 'BobT2')
        self.createBox(-Db2,-Db3,Tb-(Hb/2.0),2*Db2, 2*Db3, Hb-2*Tb, 'BobT3')        
        self.unite('Bobbin,BobT2,BobT3')
   
        self.createBox(-Db2+Tb,-Db3+Tb,-Hb/2.0, 2*Db2-2*Tb, 2*Db3-2*Tb, Hb, 'BobSlot')
        self.subtract( 'Bobbin', 'BobSlot')
     
        self.fillet('Bobbin',BRad+Tb, -Db2, -Db3, 0)
        self.fillet('Bobbin',BRad+Tb, -Db2, Db3, 0)
        self.fillet('Bobbin',BRad+Tb, Db2, Db3, 0)
        self.fillet('Bobbin',BRad+Tb, Db2, -Db3, 0)
        
        self.fillet('Bobbin',BRad, -Db2+Tb, -Db3+Tb, 0)
        self.fillet('Bobbin',BRad, -Db2+Tb, Db3-Tb, 0)        
        self.fillet('Bobbin',BRad, Db2-Tb, Db3-Tb, 0)
        self.fillet('Bobbin',BRad, Db2-Tb, -Db3+Tb, 0)  
        
        self.fillet('Bobbin',BRadE, -Db1+ECoreDimD8, -(Db1-Db2+Db3), (-Hb+Tb)/2)
        self.fillet('Bobbin',BRadE, -Db1+ECoreDimD8, (Db1-Db2+Db3), (-Hb+Tb)/2)       
        self.fillet('Bobbin',BRadE, Db1-ECoreDimD8, -(Db1-Db2+Db3), (-Hb+Tb)/2)
        self.fillet('Bobbin',BRadE, Db1-ECoreDimD8, (Db1-Db2+Db3), (-Hb+Tb)/2)
        self.fillet('Bobbin',BRadE, -Db1+ECoreDimD8, -(Db1-Db2+Db3), (Hb-Tb)/2)
        self.fillet('Bobbin',BRadE, -Db1+ECoreDimD8, (Db1-Db2+Db3), (Hb-Tb)/2)
        self.fillet('Bobbin',BRadE, Db1-ECoreDimD8, -(Db1-Db2+Db3), (Hb-Tb)/2)
        self.fillet('Bobbin',BRadE, Db1-ECoreDimD8, (Db1-Db2+Db3), (Hb-Tb)/2)
            
    def drawGeometry(self):
        self.initCore()
        self.createBox(-(self.MECoreLength/2),-(self.MECoreWidth/2),-self.MECoreHeight-self.TAirGap,
                       self.MECoreLength, self.MECoreWidth, (self.MECoreHeight-self.MSlotDepth), 'E_Core_Bottom')
                       
        self.createBox(-(self.MECoreLength/2),-(self.MECoreWidth/2),-self.MECoreHeight-self.TAirGap,
                       self.MSLegWidth, self.MECoreWidth, self.MECoreHeight-self.AirGapS, 'Leg1')
                       
        self.createBox(-(self.MCLegWidth/2),-(self.MECoreWidth/2),-self.MECoreHeight-self.TAirGap,
                       self.MCLegWidth, self.MECoreWidth, self.MECoreHeight-self.AirGapC, 'Leg2')
                       
        self.createBox((self.MECoreLength/2)-self.MSLegWidth,-(self.MECoreWidth/2),-self.MECoreHeight-self.TAirGap,
                       self.MSLegWidth, self.MECoreWidth, self.MECoreHeight-self.AirGapS, 'Leg3')
 
        self.unite('E_Core_Bottom,Leg1,Leg2,Leg3')
                
        self.fillet('E_Core_Bottom',self.DimD7, -self.DimD1/2, 0, -self.DimD4-self.TAirGap) # outer edges D_7
        self.fillet('E_Core_Bottom',self.DimD7, self.DimD1/2, 0,  -self.DimD4-self.TAirGap)  # outer edges D_7
        self.fillet('E_Core_Bottom',self.DimD8, -self.DimD2/2, 0, -self.DimD5-self.TAirGap) # inner edges D_8
        self.fillet('E_Core_Bottom',self.DimD8, self.DimD2/2, 0,  -self.DimD5-self.TAirGap)  # inner edges D_8
         
        self.duplicateMirror('E_Core_Bottom',0,0,1) 
         
        self.rename('E_Core_Bottom_1','E_Core_Top')
        self.oEditor.FitAll()
        if self.WdgStatus == 1:
            self.DrawWdg(self.DimD2,self.DimD3,self.DimD5,self.DimD6,self.SAng, self.DimD8)

# EICore inherit from ECore functions DrawWdg, CreateSingleTurn, CreateSingleTurn2, DrawBobbin        
class EICore(ECore):           
    def drawGeometry(self):
        self.initCore()
        self.createBox(-(self.MECoreLength/2),-(self.MECoreWidth/2),-self.MECoreHeight-self.TAirGap,
                       self.MECoreLength, self.MECoreWidth, (self.MECoreHeight-self.MSlotDepth), 'E_Core')
                       
        self.createBox(-(self.MECoreLength/2),-(self.MECoreWidth/2),-self.MECoreHeight-self.TAirGap,
                       self.MSLegWidth, self.MECoreWidth, self.MECoreHeight-self.AirGapS, 'Leg1')
                       
        self.createBox(-(self.MCLegWidth/2),-(self.MECoreWidth/2),-self.MECoreHeight-self.TAirGap,
                       self.MCLegWidth, self.MECoreWidth, self.MECoreHeight-self.AirGapC, 'Leg2')
                       
        self.createBox((self.MECoreLength/2)-self.MSLegWidth,-(self.MECoreWidth/2),-self.MECoreHeight-self.TAirGap,
                       self.MSLegWidth, self.MECoreWidth, self.MECoreHeight-self.AirGapS, 'Leg3')
 
        self.unite('E_Core,Leg1,Leg2,Leg3')
                
        self.fillet('E_Core',self.DimD7, -self.DimD1/2, -self.DimD6/2, -(self.DimD4/2)-self.TAirGap)
        self.fillet('E_Core',self.DimD7, -self.DimD1/2, self.DimD6/2, -(self.DimD4/2)-self.TAirGap) 
        self.fillet('E_Core',self.DimD7, self.DimD1/2, -self.DimD6/2, -(self.DimD4/2)-self.TAirGap) 
        self.fillet('E_Core',self.DimD7, self.DimD1/2, self.DimD6/2, -(self.DimD4/2)-self.TAirGap) 
                 
        self.createBox(-self.DimD1/2, -self.DimD6/2, self.TAirGap,
                       self.DimD1,self.DimD6,self.DimD8, 'I_Core')
                       
        self.fillet('I_Core',self.DimD7, -self.DimD1/2, -self.DimD6/2, (self.DimD8/2)-self.TAirGap)
        self.fillet('I_Core',self.DimD7, -self.DimD1/2, self.DimD6/2, (self.DimD8/2)-self.TAirGap) 
        self.fillet('I_Core',self.DimD7, self.DimD1/2, -self.DimD6/2, (self.DimD8/2)-self.TAirGap) 
        self.fillet('I_Core',self.DimD7, self.DimD1/2, self.DimD6/2, (self.DimD8/2)-self.TAirGap)                
                       
        self.move('E_Core,I_Core',0,0,self.DimD5/2)
         
        self.oEditor.FitAll()
        if self.WdgStatus == 1:
			self.DrawWdg(self.DimD2,self.DimD3,(self.DimD5+2*self.TAirGap)/2,self.DimD6,self.SAng)  

# UCore inherit from ECore functions CreateSingleTurn, CreateSingleTurn2, DrawBobbin
class UCore(ECore):
    def initCore(self):
        self.MECoreLength    = self.DimD1
        self.MECoreWidth     = self.DimD5
        self.MECoreHeight    = self.DimD3
  
    def DrawWdg(self,DimD1,DimD2,DimD3,DimD4,DimD5,SAng):
        MNumLayers  = self.MNumLayers 
        MLSpacing   = self.MLSpacing  
        MTopMargin  = self.MTopMargin 
        MSideMargin = self.MSideMargin
        MBobbinThk  = self.MBobbinThk 
        MWdgType    = self.MWdgType   
        MCondType   = self.MCondType  
        BobStat     = self.BobStat    
        WdgParDict  = self.WdgParDict
        MSlotWidth  = DimD1-DimD2
        MSlotHeight = DimD4*2
        LegWidth = (DimD1-DimD2)/2
        
        if BobStat > 0:
            self.DrawBobbin(MSlotHeight-2*MTopMargin,((LegWidth/2)+DimD2)-MSideMargin,(LegWidth/2.0)+MSideMargin+MBobbinThk,(DimD5/2.0)+MSideMargin+MBobbinThk,MBobbinThk,MSideMargin,SAng)
        
        if MWdgType == 1:
            MTDx = MTopMargin+MBobbinThk
            for MAx in self.WdgParDict:
                for MBx in range(0,int(self.WdgParDict[MAx][2])):
                    MRecSzX = LegWidth + (2*(MSideMargin+MBobbinThk)) + ((2*MBx+1)*self.WdgParDict[MAx][0]) + (2*MBx*2*self.WdgParDict[MAx][3]) + 2*self.WdgParDict[MAx][3]
                    MRecSzY = DimD5 + (2*(MSideMargin+MBobbinThk)) + ((2*MBx+1)*self.WdgParDict[MAx][0]) + (2*MBx*2*self.WdgParDict[MAx][3]) + 2*self.WdgParDict[MAx][3]
                    if MCondType == 1:
                        MRecSzZ = -self.WdgParDict[MAx][1]/2.0
                        self.CreateSingleTurn2(MRecSzX,MRecSzY,MRecSzZ,MCondType,self.WdgParDict[MAx][0],self.WdgParDict[MAx][1],DimD4-MTDx-(self.WdgParDict[MAx][3]),MAx,MBx,(MRecSzX-LegWidth)/2,SAng)
                    else:
                        MRecSzZ = -self.WdgParDict[MAx][0]/2.0
                        self.CreateSingleTurn(MRecSzX,MRecSzY,MRecSzZ,MCondType,self.WdgParDict[MAx][0],self.WdgParDict[MAx][1],DimD4-MTDx-(self.WdgParDict[MAx][3]),MAx,MBx,(MRecSzX-LegWidth)/2,SAng)
                if MCondType == 1:
                    MTDx = MTDx + MLSpacing + self.WdgParDict[MAx][1] + 2*self.WdgParDict[MAx][3]
                else:
                    MTDx = MTDx + MLSpacing + self.WdgParDict[MAx][0] + 2*self.WdgParDict[MAx][3]
        else:
            MTDx = MSideMargin+MBobbinThk 
            for MAx in self.WdgParDict.keys():
                for MBx in range(0,int(self.WdgParDict[MAx][2])):
                    MRecSzX = LegWidth + (2*(MTDx+(self.WdgParDict[MAx][0]/2.0)))+ 2*self.WdgParDict[MAx][3]
                    MRecSzY = DimD5 + (2*(MTDx+(self.WdgParDict[MAx][0]/2.0)))+ 2*self.WdgParDict[MAx][3]
                    if MCondType == 1:
                        MRecSzZ = -self.WdgParDict[MAx][1]/2.0
                        self.CreateSingleTurn2(MRecSzX,MRecSzY,MRecSzZ,MCondType,self.WdgParDict[MAx][0],self.WdgParDict[MAx][1],DimD4-MTopMargin-MBobbinThk-(self.WdgParDict[MAx][3])-MBx*(2*self.WdgParDict[MAx][3]+self.WdgParDict[MAx][1]),MAx,MBx,(MRecSzX-LegWidth)/2,SAng)
                    else:
                        MRecSzZ = -self.WdgParDict[MAx][0]/2.0
                        self.CreateSingleTurn(MRecSzX,MRecSzY,MRecSzZ,MCondType,self.WdgParDict[MAx][0],self.WdgParDict[MAx][1],DimD4-MTopMargin-MBobbinThk-(self.WdgParDict[MAx][3])-MBx*(2*self.WdgParDict[MAx][3]+self.WdgParDict[MAx][0]),MAx,MBx,(MRecSzX-LegWidth)/2,SAng)
                MTDx = MTDx + MLSpacing + self.WdgParDict[MAx][0] + 2*self.WdgParDict[MAx][3]
    
    def drawGeometry(self):
        self.initCore()
        self.createBox(-(self.DimD1-self.DimD2)/4,-(self.MECoreWidth/2),-self.MECoreHeight-self.TAirGap,
                       self.MECoreLength, self.MECoreWidth, self.MECoreHeight, 'U_Core_Bottom')
                       
        self.createBox((self.DimD1-self.DimD2)/4, -(self.MECoreWidth)/2, -self.DimD4-self.TAirGap,
                       self.DimD2, self.MECoreWidth, self.DimD4, 'XSlot')
                       
        self.subtract('U_Core_Bottom','XSlot')      

        if self.AirGapC > 0:
            self.createBox((self.DimD1-self.DimD2)/4, -(self.MECoreWidth)/2, -self.AirGapC,
                            (self.DimD1-self.DimD2)/2, self.MECoreWidth, self.AirGapC, 'AgC')
            self.subtract('U_Core_Bottom','AgC') 
            
        if self.AirGapS > 0:
            self.createBox(self.DimD2+(self.DimD1-self.DimD2)/4, -(self.MECoreWidth)/2, -self.AirGapS,
                            (self.DimD1-self.DimD2)/2, self.MECoreWidth, self.AirGapS, 'AgS')
            self.subtract('U_Core_Bottom','AgS') 

        self.duplicateMirror('U_Core_Bottom',0,0,1)
        self.rename('U_Core_Bottom_1','U_Core_Top')
        self.oEditor.FitAll()
        if self.WdgStatus == 1:
			self.DrawWdg(self.DimD1,self.DimD2,self.DimD3,(self.DimD4+self.TAirGap),self.DimD5,self.SAng) 

class EFDCore(ECore):   
    def drawGeometry(self):
        self.initCore()
        self.createBox(-(self.MECoreLength/2),-(self.MECoreWidth/2),-self.MECoreHeight-self.TAirGap,
                       self.MECoreLength, self.MECoreWidth, (self.MECoreHeight-self.MSlotDepth), 'EFD_Core_Bottom')
                       
        self.createBox(-(self.MECoreLength/2),-(self.MECoreWidth/2),-self.MECoreHeight-self.TAirGap,
                       self.MSLegWidth, self.MECoreWidth, self.MECoreHeight-self.AirGapS, 'Leg1')
                       
        self.createBox(-(self.MCLegWidth/2),-(self.MECoreWidth/2)-self.DimD8,-self.MECoreHeight-self.TAirGap,
                       self.MCLegWidth, self.DimD7, self.MECoreHeight-self.AirGapC, 'Leg2')
                       
        self.createBox((self.MECoreLength/2)-self.MSLegWidth,-(self.MECoreWidth/2),-self.MECoreHeight-self.TAirGap,
                       self.MSLegWidth, self.MECoreWidth, self.MECoreHeight-self.AirGapS, 'Leg3')
 
        self.unite('EFD_Core_Bottom,Leg1,Leg2,Leg3')
                
        self.duplicateMirror('EFD_Core_Bottom',0,0,1) 
        self.rename('EFD_Core_Bottom_1','E_Core_Top')
        self.oEditor.FitAll()

 # difference between ECore and EFD is only offset  of central leg. 
 # We can compensate it by using relative CS for winding and bobbin

        self.oEditor.CreateRelativeCS(
            [
                "NAME:RelativeCSParameters",
                "Mode:="		, "Axis/Position",
                "OriginX:="		, "0mm",
                "OriginY:="		, str((self.DimD7-self.DimD6)/2-self.DimD8)+'mm',
                "OriginZ:="		, "0mm",
                "XAxisXvec:="		, "1mm",
                "XAxisYvec:="		, "0mm",
                "XAxisZvec:="		, "0mm",
                "YAxisXvec:="		, "0mm",
                "YAxisYvec:="		, "1mm",
                "YAxisZvec:="		, "0mm"
            ], 
            [
                "NAME:Attributes",
                "Name:="		, 'CentralLegCS'
            ])
        Cores.CS = 'CentralLegCS'  
                
        if self.WdgStatus == 1:
			self.DrawWdg(self.DimD2,self.DimD3,self.DimD5,self.DimD7,self.SAng)  
        self.oEditor.SetWCS(["NAME:SetWCS Parameter","Working Coordinate System:=", "Global","RegionDepCSOk:=", False]) 
        Cores.CS = 'Global'
    
    # need to redifine it since it does not work with relative CS
    def fillet(self,name,Radius, X,Y,Z):
            ID = [self.oEditor.GetEdgeByPosition(["NAME:EdgeParameters",
            "BodyName:=", name, 
            "XPosition:=", str(X) + 'mm', 
            "YPosition:=", str(Y+(self.DimD7-self.DimD6)/2-self.DimD8) + 'mm', 
            "ZPosition:=", str(Z) + 'mm'])]
            self.oEditor.Fillet(
            [
                "NAME:Selections",
                "Selections:="		, name,
                "NewPartsModelFlag:="	, "Model"
            ], 
            [
                "NAME:Parameters",
                [
                    "NAME:FilletParameters",
                    "Edges:="		, ID,
                    "Vertices:="		, [],
                    "Radius:="		, str(Radius) + 'mm',
                    "Setback:="		, "0mm"
                ]
            ])    
             
# UICore inherit from ECore functions CreateSingleTurn, CreateSingleTurn2, DrawBobbin        
# and from UCore inherit DrawWdg and __init__
class UICore(UCore):
    def drawGeometry(self):
        self.initCore()
        self.createBox(-(self.DimD1-self.DimD2)/4,-(self.MECoreWidth/2),(self.DimD4/2)-self.DimD3-self.TAirGap,
                       self.MECoreLength, self.MECoreWidth, self.MECoreHeight, 'U_Core')
                       
        self.createBox((self.DimD1-self.DimD2)/4, -(self.MECoreWidth)/2, -self.DimD4/2-self.TAirGap,
                       self.DimD2, self.MECoreWidth, self.DimD4, 'XSlot')
                       
        self.subtract('U_Core','XSlot')      

        if self.AirGapC > 0:
            self.createBox((self.DimD1-self.DimD2)/4, -(self.MECoreWidth)/2, -self.AirGapC,
                            (self.DimD1-self.DimD2)/2, self.MECoreWidth, self.AirGapC, 'AgC')
            self.subtract('U_Core','AgC') 
            
        if self.AirGapS > 0:
            self.createBox(self.DimD2+(self.DimD1-self.DimD2)/4, -(self.MECoreWidth)/2, -self.AirGapS,
                            (self.DimD1-self.DimD2)/2, self.MECoreWidth, self.AirGapS, 'AgS')
            self.subtract('U_Core','AgS') 

        self.createBox(-(self.DimD1-self.DimD2)/4+((self.DimD1-self.DimD6)/2),-self.DimD7/2,(self.DimD4/2)+self.TAirGap,
                       self.DimD6, self.DimD7, self.DimD8, 'I_Core')
            
        self.oEditor.FitAll()
        if self.WdgStatus == 1:
			self.DrawWdg(self.DimD1,self.DimD2,self.DimD3/2,(self.DimD4/2)+self.TAirGap,self.DimD5,self.SAng) 
    
class PQCore(Cores):
    def initCore(self):
        self.MECoreLength    = self.DimD1
        self.MECoreWidth     = self.DimD6
        self.MECoreHeight    = self.DimD4
        self.MSLegWidth      = (self.DimD1-self.DimD2)/2
        self.MCLegWidth      = self.DimD3
        self.MSlotDepth      = self.DimD5
           
    def DrawWdg(self,DimD2,DimD3,DimD5,DimD6,SAng):
        MNumLayers  = self.MNumLayers 
        MLSpacing   = self.MLSpacing  
        MTopMargin  = self.MTopMargin 
        MSideMargin = self.MSideMargin
        MBobbinThk  = self.MBobbinThk 
        MWdgType    = self.MWdgType   
        MCondType   = self.MCondType  
        BobStat     = self.BobStat    
        WdgParDict  = self.WdgParDict
        MSlotWidth  = (DimD2-DimD3)/2.0
        MSlotHeight = DimD5*2
        if BobStat > 0:
            self.DrawBobbin(MSlotHeight-2*MTopMargin,(DimD2/2.0)-MSideMargin,(DimD3/2.0)+MSideMargin+MBobbinThk,MBobbinThk,SAng)
        
        if MWdgType == 1:
            MTDx = MTopMargin+MBobbinThk
            for MAx in self.WdgParDict:
                for MBx in range(0,int(self.WdgParDict[MAx][2])):
                    MRecSzX = DimD3 + (2*(MSideMargin+MBobbinThk)) + ((2*MBx+1)*self.WdgParDict[MAx][0]) + (2*MBx*2*self.WdgParDict[MAx][3]) + 2*self.WdgParDict[MAx][3]
                    if MCondType == 1:
                        MRecSzZ = -self.WdgParDict[MAx][1]/2.0
                    else:
                        MRecSzZ = -self.WdgParDict[MAx][0]/2.0
                    self.CreateSingleTurn(MRecSzX,MRecSzZ,MCondType,self.WdgParDict[MAx][0],self.WdgParDict[MAx][1],DimD5-MTDx-(self.WdgParDict[MAx][3]),MAx,MBx,SAng)
                if MCondType == 1:
                    MTDx = MTDx + MLSpacing + self.WdgParDict[MAx][1] + 2*self.WdgParDict[MAx][3]
                else:
                    MTDx = MTDx + MLSpacing + self.WdgParDict[MAx][0] + 2*self.WdgParDict[MAx][3]
        else:
            MTDx = MSideMargin+MBobbinThk 
            for MAx in self.WdgParDict.keys():
                for MBx in range(0,int(self.WdgParDict[MAx][2])):
                    MRecSzX = DimD3 + (2*(MTDx+(self.WdgParDict[MAx][0]/2.0)))+ 2*self.WdgParDict[MAx][3]
                    if MCondType == 1:
                        MRecSzZ = -self.WdgParDict[MAx][1]/2.0
                        self.CreateSingleTurn(MRecSzX,MRecSzZ,MCondType,self.WdgParDict[MAx][0],self.WdgParDict[MAx][1],DimD5-MTopMargin-MBobbinThk-(self.WdgParDict[MAx][3])-MBx*(2*self.WdgParDict[MAx][3]+self.WdgParDict[MAx][1]),MAx,MBx,SAng)
                    else:
                        MRecSzZ = -self.WdgParDict[MAx][0]/2.0
                        self.CreateSingleTurn(MRecSzX,MRecSzZ,MCondType,self.WdgParDict[MAx][0],self.WdgParDict[MAx][1],DimD5-MTopMargin-MBobbinThk-(self.WdgParDict[MAx][3])-MBx*(2*self.WdgParDict[MAx][3]+self.WdgParDict[MAx][0]),MAx,MBx,SAng)
                MTDx = MTDx + MLSpacing + self.WdgParDict[MAx][0] + 2*self.WdgParDict[MAx][3]

    def CreateSingleTurn(self,PathX,PathZ,ProfTyp,ProfAX,ProfZ,ZPos,LayNum,TurnNum,SAng):
        NumSegs = int(360/SAng)
        self.createPolygon(0,0,PathZ,
                           -PathX/2,0,PathZ,NumSegs, 
                           'pathLine', 'Z', covered = False )
                       
        if ProfTyp == 1:   
            self.createRectangle(-((PathX/2)+(ProfAX/2)),0,((PathZ)-(ProfZ/2)),
                                 ProfAX,ProfZ, 'sweepProfile')
        else:
            self.createCircle(-(PathX/2),0,PathZ,
                           ProfAX,ProfZ,'sweepProfile', 'Y')
            
        self.rename( 'pathLine', 'Tool%s_%s'%(LayNum,TurnNum+1))
        self.sweepAlongPath('sweepProfile,Tool%s_%s'%(LayNum,TurnNum+1))

        self.move('sweepProfile',0,0,ZPos)
        self.rename( 'sweepProfile', 'Layer%s_%s'%(LayNum,TurnNum+1))

    def DrawBobbin(self,Hb,Db1,Db2,Tb,SAng): 
        self.createPolyhedron(0,0,-Tb+(Hb/2.0),
                              -Db1,0,-Tb+(Hb/2.0),
                              Tb,self.NumSegs,'Bobbin', '(255,248,157)')
                              
        self.createPolyhedron(0,0,-(Hb/2.0),
                              -Db1,0,-(Hb/2.0),
                              Tb,self.NumSegs,'BobT2')                      
                              
        self.createPolyhedron(0,0,Tb-(Hb/2.0),
                              -Db2,0,Tb-(Hb/2.0),
                              Hb-2*Tb,self.NumSegs,'BobT3')
                              
        self.unite('Bobbin,BobT2,BobT3')

        self.createPolyhedron(0,0,-Hb/2.0,
                              -(Db2-Tb),0,-Hb/2.0,
                              Hb,self.NumSegs,'BobT4')   

        self.subtract('Bobbin','BobT4')
            
    def drawGeometry(self):
        self.initCore()
        NumSegs = int(360/self.SAng)
        
        self.createBox(-self.DimD1/2,-self.DimD8/2,-(self.DimD4/2)-self.TAirGap,
                       (self.DimD1-self.DimD6)/2, self.DimD8, (self.DimD4/2)-self.AirGapS, 'PQ_Core_Bottom')
        
        IntL = math.sqrt(((self.DimD3/2)**2)-((self.DimD7/2)**2))
        vertices1 = [[-self.DimD6/2,-0.4*self.DimD8,-(self.DimD4/2)-self.TAirGap]]
        vertices1.append([-self.DimD6/2,0.4*self.DimD8,-(self.DimD4/2)-self.TAirGap])
        vertices1.append([-self.DimD7/2,+IntL,-(self.DimD4/2)-self.TAirGap])
        vertices1.append([-self.DimD7/2,-IntL,-(self.DimD4/2)-self.TAirGap])        
        vertices1.append(vertices1[0])
        
        pointsArray,segmentsArray = self.pointsSegmentsGenerator(vertices1)        
        self.createPolyline( pointsArray, segmentsArray, "Polyline1")
        self.sweepAlongVec("Polyline1", (self.DimD4/2)-self.AirGapS) 
       
        self.unite('PQ_Core_Bottom,Polyline1')
        self.createPolyhedron(0,0,-(self.DimD5/2)-self.TAirGap,
                              -(self.DimD2/2),0,-(self.DimD5/2)-self.TAirGap, 
                              self.DimD5/2, NumSegs, 'XCyl1')
                              
        self.subtract('PQ_Core_Bottom','XCyl1')
        self.duplicateMirror('PQ_Core_Bottom',1,0,0)
        
        self.createPolyhedron(0,0,-(self.DimD4/2)-self.TAirGap,
                              -(self.DimD3/2),0,-(self.DimD4/2)-self.TAirGap, 
                              self.DimD4/2 - self.AirGapC, NumSegs, 'XCyl2')
                              
        self.unite('PQ_Core_Bottom,PQ_Core_Bottom_1,XCyl2')
        self.duplicateMirror('PQ_Core_Bottom',0,0,1)
        
        self.rename('PQ_Core_Bottom_2','PQ_Core_Top')
        self.oEditor.FitAll()

        if self.WdgStatus == 1:
			self.DrawWdg(self.DimD2,self.DimD3,self.DimD5/2+self.TAirGap,self.DimD6,self.SAng)  

# ETDCore inherit from PQCore functions DrawWdg, CreateSingleTurn, DrawBobbin         
class ETDCore(PQCore):          
    def drawGeometry(self, coreName):
        self.initCore()
        self.createBox(-(self.MECoreLength/2),-(self.MECoreWidth/2),-self.MECoreHeight-self.TAirGap,
                       self.MECoreLength, self.MECoreWidth, self.MECoreHeight-self.AirGapS, coreName+'_Core_Bottom')
                       
        self.createPolyhedron(0,0,-self.MSlotDepth-self.TAirGap,
                              -self.DimD2/2,0,-self.MSlotDepth-self.TAirGap,
                              self.MSlotDepth,self.NumSegs,'XCyl1')
                              
        self.subtract(coreName+'_Core_Bottom','XCyl1')
        
        if coreName == 'ER' and self.DimD7>0:
            self.createBox(-(self.DimD7)/2,-(self.MECoreWidth)/2,-self.MSlotDepth-self.TAirGap,
                       self.DimD7,self.MECoreWidth,self.MSlotDepth, 'Tool')
            self.subtract(coreName+'_Core_Bottom','Tool')
        
        self.createPolyhedron(0,0,-self.MSlotDepth-self.TAirGap,
                              -self.DimD3/2,0,-self.MSlotDepth-self.TAirGap,
                              self.MSlotDepth,self.NumSegs,'XCyl2')
                              
        self.unite(coreName+'_Core_Bottom,XCyl2')
        self.duplicateMirror(coreName+'_Core_Bottom',0,0,1)
        self.rename(coreName+'_Core_Bottom_1',coreName+'_Core_Top')
                         
        self.oEditor.FitAll()
        if self.WdgStatus == 1:
			self.DrawWdg(self.DimD2,self.DimD3,(self.DimD5+self.TAirGap),self.DimD6,self.SAng) 

# RMCore inherit from PQCore functions DrawWdg, CreateSingleTurn, DrawBobbin        
class RMCore(PQCore):          
    def drawGeometry(self):
        self.initCore()
        NumSegs = int(360/self.SAng)     
        Dia = self.DimD7/math.sqrt(2)
                
        vertices1 = [[-self.DimD1/2,(Dia-(self.DimD1/2)),-self.AirGapS-self.TAirGap]]
        vertices1.append([-(Dia/2),(Dia/2),-self.AirGapS-self.TAirGap])
        vertices1.append([-(self.DimD8/2),(self.DimD8/2),-self.AirGapS-self.TAirGap])
        vertices1.append([(self.DimD8/2),(self.DimD8/2),-self.AirGapS-self.TAirGap])   
        vertices1.append([(Dia/2),(Dia/2),-self.AirGapS-self.TAirGap])
        vertices1.append([self.DimD1/2,(Dia-(self.DimD1/2)),-self.AirGapS-self.TAirGap])
        vertices1.append([self.DimD1/2,-(Dia-(self.DimD1/2)),-self.AirGapS-self.TAirGap])
        vertices1.append([(Dia/2),-(Dia/2),-self.AirGapS-self.TAirGap])
        vertices1.append([(self.DimD8/2),-(self.DimD8/2),-self.AirGapS-self.TAirGap])
        vertices1.append([-(self.DimD8/2),-(self.DimD8/2),-self.AirGapS-self.TAirGap])
        vertices1.append([-(Dia/2),-(Dia/2),-self.AirGapS-self.TAirGap])
        vertices1.append([-self.DimD1/2,-(Dia-(self.DimD1/2)),-self.AirGapS-self.TAirGap])       
        vertices1.append(vertices1[0])
        
        pointsArray,segmentsArray = self.pointsSegmentsGenerator(vertices1)        
        self.createPolyline( pointsArray, segmentsArray, 'RM_Core_Bottom')
        self.sweepAlongVec('RM_Core_Bottom', -(self.DimD4/2)+self.AirGapS) 
             
        self.createPolyhedron(0,0,-(self.DimD5/2)-self.TAirGap,
                              -(self.DimD2/2),0,-(self.DimD5/2)-self.TAirGap, 
                              self.DimD5/2, NumSegs, 'XCyl1')
                              
        self.subtract('RM_Core_Bottom','XCyl1')
        
        self.createPolyhedron(0,0,-(self.DimD5/2)-self.TAirGap,
                              -(self.DimD3/2),0,-(self.DimD5/2)-self.TAirGap, 
                              self.DimD5/2-self.AirGapC, NumSegs, 'XCyl2')
        
        self.unite('RM_Core_Bottom,XCyl2')
        if self.DimD6 !=0:
            self.createPolyhedron(0,0,-(self.DimD4/2)-self.TAirGap,
                                  -(self.DimD6/2),0,-(self.DimD4/2)-self.TAirGap, 
                                  self.DimD4/2, NumSegs, 'XCyl3')  
            self.subtract('RM_Core_Bottom','XCyl3')
            
        self.duplicateMirror('RM_Core_Bottom',0,0,1)
                
        self.rename('RM_Core_Bottom_1','RM_Core_Top')
        self.oEditor.FitAll()

        if self.WdgStatus == 1:
			self.DrawWdg(self.DimD2,self.DimD3,self.DimD5/2+self.TAirGap,self.DimD6,self.SAng)  
 
#depending on bobbin we can inherit bobbin, createturn,drawwdg from PQCore. speak with JMark
class EPCore(PQCore):
    def initCore(self):
        self.MECoreLength    = self.DimD1
        self.MECoreWidth     = self.DimD6
        self.MECoreHeight    = self.DimD4/2
        self.MSLegWidth      = (self.DimD1-self.DimD2)/2
        self.MCLegWidth      = self.DimD3
        self.MSlotDepth      = self.DimD5/2
  
    def drawGeometry(self, coreType = 'EP'):
        self.initCore()
        self.createBox(-(self.MECoreLength/2),-(self.MECoreWidth/2),-self.MECoreHeight-self.TAirGap,
                       self.MECoreLength, self.MECoreWidth, self.MECoreHeight-self.AirGapS, coreType+'_Core_Bottom')
                       
        self.createPolyhedron(0,(self.DimD6/2)-self.DimD7,-self.MSlotDepth-self.TAirGap,
                              -self.DimD2/2,(self.DimD6/2)-self.DimD7,-self.MSlotDepth-self.TAirGap,
                              self.MSlotDepth,self.NumSegs,'XCyl1')
        
        self.createBox(-self.DimD2/2,(self.DimD6/2)-self.DimD7,-self.MSlotDepth-self.TAirGap,
                       self.DimD2,self.DimD7,self.MSlotDepth, 'Box2')
        
        self.unite('Box2,XCyl1')
        self.subtract(coreType+'_Core_Bottom','Box2')
        
        self.createPolyhedron(0,(self.DimD6/2)-self.DimD7,-self.MSlotDepth-self.TAirGap,
                              -(self.DimD3/2),(self.DimD6/2)-self.DimD7,-self.MSlotDepth-self.TAirGap,
                              self.MSlotDepth-self.AirGapC,self.NumSegs,'XCyl2')                      
        
        self.unite(coreType+'_Core_Bottom,XCyl2')
        self.move(coreType+'_Core_Bottom',0,(-self.DimD6/2.0)+self.DimD7,0)
        
        self.duplicateMirror(coreType+'_Core_Bottom',0,0,1)
        self.rename(coreType+'_Core_Bottom_1',coreType+'_Core_Top')
                         
        self.oEditor.FitAll()
        if self.WdgStatus == 1:
			self.DrawWdg(self.DimD2,self.DimD3,(self.DimD5/2+self.TAirGap),self.DimD6,self.SAng) 
             
class PCore(PQCore):          
    def drawGeometry(self, coreName):
        self.initCore()
        if coreName == 'PH':
            self.DimD4*=2
            self.DimD5*=2
            
        self.createPolyhedron(0,0,-(self.DimD4/2)-self.TAirGap,
                              -self.DimD1/2, 0, -(self.DimD4/2)-self.TAirGap,
                              (self.DimD4/2)-self.AirGapS,self.NumSegs,coreName+'_Core_Bottom')
                              
        self.createPolyhedron(0,0,-(self.DimD5/2)-self.TAirGap,
                              -(self.DimD2/2),0,-(self.DimD5/2)-self.TAirGap,
                              self.DimD5/2,self.NumSegs,'XCyl1')
       
        self.subtract(coreName+'_Core_Bottom','XCyl1')
        
        self.createPolyhedron(0,0,-(self.DimD5/2)-self.TAirGap,
                              -(self.DimD3/2),0,-(self.DimD5/2)-self.TAirGap,
                              self.DimD5/2,self.NumSegs,'XCyl2')
                              
        self.unite(coreName+'_Core_Bottom,XCyl2')                      
                        
        if self.DimD6 != 0 :
            self.createPolyhedron(0,0,-(self.DimD4/2)-self.TAirGap,
                                  -(self.DimD6/2),0,-(self.DimD4/2)-self.TAirGap,
                                  self.DimD4/2,self.NumSegs,'Tool')
            self.subtract(coreName+'_Core_Bottom','Tool')
        
        if self.DimD7 != 0 :
            self.createBox(-self.DimD1/2, -self.DimD7/2, -(self.DimD4/2)-self.TAirGap,
                           (self.DimD1-self.DimD8)/2, self.DimD7,self.DimD4/2, 'Slot1')
             
            self.createBox(self.DimD1/2, -self.DimD7/2, -(self.DimD4/2)-self.TAirGap,
                           -(self.DimD1-self.DimD8)/2, self.DimD7,self.DimD4/2, 'Slot2')
                           
            self.subtract(coreName+'_Core_Bottom','Slot1,Slot2')               
        
        self.duplicateMirror(coreName+'_Core_Bottom',0,0,1)
        self.rename(coreName+'_Core_Bottom_1',coreName+'_Core_Top')
        
        if coreName == 'PT':
            self.createBox(-self.DimD1/2, -self.DimD1/2, (self.DimD4/2)+self.TAirGap,
                           (self.DimD1-self.DimD8)/2, self.DimD1,-self.DimD4/2, 'Slot3')
             
            self.createBox(self.DimD1/2, -self.DimD1/2, (self.DimD4/2)+self.TAirGap,
                           -(self.DimD1-self.DimD8)/2, self.DimD1,-self.DimD4/2, 'Slot4')
                           
            self.subtract(coreName+'_Core_Top','Slot3,Slot4')
        
        self.oEditor.FitAll()
        if self.WdgStatus == 1:
			self.DrawWdg(self.DimD2,self.DimD3,(self.DimD5/2+self.TAirGap),self.DimD6,self.SAng)        