# -*- coding: utf-8 -*-

#      Replacement of old ETK (Electronic Transformer Kit) for Maxwell
#
#      Originally written: Mark Christini (mark.christini@ansys.com)
#                          Tushar Sambharam (tushar.sambharam@ansys.com)
#
#      ACT Written by : Maksim Beliaev (maksim.beliaev@ansys.com)
#      Last updated : 19.06.2018

import datetime     # get the time library
import os         # import operations system module
import math
import re
from webbrowser import open as webopen

import clr        # import Common Language Runtime
clr.AddReferenceByPartialName("Microsoft.VisualBasic")      #import functions from Visual Basic
from Microsoft.VisualBasic.Constants import vbOKOnly,vbYesNo
from Microsoft.VisualBasic.Interaction import MsgBox
def M(string1='',string2='',string3='',string4='',string5='',string6=''):
  return MsgBox(str(string1)+'\n'+ str(string2)+'\n'+str(string3)+'\n'+str(string4)+'\n'+str(string5)+'\n'+str(string6))

#***************   ***************   ***************   ***************   ***************   --------#
def createModel(step):
  # invoke validation from valueChecker file
  checkWinding(step)
  checkBobbin(step)

  global designName, oProject
  flagAutoSave = -1
  # read custom material file and dataset points for it
  libPath = oDesktop.GetPersonalLibDirectory() + '\Materials\\'
  if os.path.isfile(libPath + 'matdata.tab'):
    with open(libPath + 'matdata.tab','r') as Inmat:
      next(Inmat)
      for line in Inmat:
        line = line.split()
        matDict[line[0]] = line[1:]
        try:
          datasheet = open(libPath+line[0]+'.tab','r')
        except:
          raise UserErrorMessageException('File {} in directory {} does not exist!'.format(line[0]+'.tab',libPath))
        bufferList=[]
        for lineData in datasheet:
          if not( """X" \t"Y""" in lineData):
            lineArray = lineData.rstrip().split()
            try:
              bufferList.append([float(str(a).replace(',','.')) for a in lineArray])
            except:
              raise UserErrorMessageException('Wrong values in file {}'.format(line[0]+'.tab'))
        datasheet.close()
        matKeyPoints[line[0]]=bufferList
  else:
    if not os.path.exists(libPath):
      os.makedirs(libPath)
    file = open(libPath + 'matdata.tab','w')
    file.write('Material Name\tConductivity\tCm\tx\ty\tdensity\n')
    file.close()

  getTime = datetime.datetime.now().strftime('%d_%m_%Y_%H_%M_%S')
  designName = 'Transformer_' + getTime
  oProject = oDesktop.GetActiveProject()
  if oProject == None:
    oProject = oDesktop.NewProject()
  oProject.InsertDesign("Maxwell 3D", designName , "EddyCurrent", "")

  oDesign = oProject.SetActiveDesign(designName)
  oEditor = oDesign.SetActiveEditor("3D Modeler")

  allCores={'E':ECore(), 'EI':EICore(), 'U':UCore(), 'UI':UICore(), 'PQ':PQCore(),
      'ETD':ETDCore(),'EQ':ETDCore(),'EC':ETDCore(), 'RM':RMCore(), 'EP':EPCore(),
      'EFD':EFDCore(), 'ER':ETDCore(), 'P':PCore(), 'PT':PCore(), 'PH':PCore()}

  coresArguments = {'EC':'EC', 'ETD':'ETD', 'EQ':'EQ', 'ER':'ER','P':'P','PT':'PT', 'PH':'PH'}

  if oDesktop.GetAutoSaveEnabled() == 1:
    flagAutoSave = 1
    oDesktop.EnableAutoSave(0)
    oDesktop.AddMessage(oProject.GetName(), '',1,'Autosave was disabled! It would be enabled back after model creation!')

  tryCore = str(step1.Properties["coreProperties/coreType"].Value)
  a=allCores[tryCore]
  if tryCore in coresArguments:
    a.drawGeometry(coresArguments[tryCore])
  else:
    a.drawGeometry()

  if flagAutoSave > 0:
    oDesktop.EnableAutoSave(1)
    oDesktop.AddMessage(oProject.GetName(), '',0,'Autosave was enabled!')


def setupAnalysis(step):
  oDesign = oProject.SetActiveDesign(designName)
  oEditor = oDesign.SetActiveEditor("3D Modeler")

  coreMaterial = step.Properties["defineSetup/coreMaterial"].Value
  coilMaterial = step.Properties["defineSetup/coilMaterial"].Value

  ObjList = oEditor.GetObjectsInGroup('Solids')
  LayList = []
  LaySecList = []
  LayDelList = []
  CoreList = []
  for EachObj in ObjList:
    if "Layer" in EachObj:
      LayList.append(EachObj)
      LaySecList.append(EachObj+"_Section1")
      LayDelList.append(EachObj+"_Section1_Separate1")
    if "Core" in EachObj:
      CoreList.append(EachObj)

  oDefinitionManager = oProject.GetDefinitionManager()
  if not oDefinitionManager.DoesMaterialExist("Material_"+coreMaterial):
    CordList = ["NAME:Coordinates"]
    for coordinatePair in range(len(matKeyPoints[coreMaterial])):
      CordList.append(["NAME:Coordinate","X:=", matKeyPoints[coreMaterial][coordinatePair][0],
                       "Y:=",matKeyPoints[coreMaterial][coordinatePair][1]])

    oProject.AddDataset(["NAME:$Mu_"+coreMaterial,CordList])

    oDefinitionManager.AddMaterial(
      [
        "NAME:Material_"+coreMaterial,
        "CoordinateSystemType:=", "Cartesian",
        ["NAME:AttachedData"],
        ["NAME:ModifierData"],
        "permeability:="  , "pwl($Mu_"+coreMaterial+",Freq)",
        "conductivity:="  , matDict[coreMaterial][0],
        [
          "NAME:core_loss_type",
          "property_type:="   , "ChoiceProperty",
          "Choice:="    , "Power Ferrite"
        ],
        "core_loss_cm:="  , matDict[coreMaterial][1],
        "core_loss_x:="   , matDict[coreMaterial][2],
        "core_loss_y:="   , matDict[coreMaterial][3],
        "core_loss_kdc:="   , "0",
        "mass_density:="  , matDict[coreMaterial][4]
      ])

  assignMaterial(oEditor, ','.join(CoreList), '"Material_'+coreMaterial+'"')
  assignMaterial(oEditor, ','.join(LayList), '"'+coilMaterial+'"')

  if coilMaterial == 'Copper':
    changeColor(oEditor,LayList,255,128,64)
  else:
    changeColor(oEditor,LayList,132,135,137)

  if (step2.Properties["windingProperties/drawWinding/includeBobbin"].Value == True
    and step2.Properties["windingProperties/drawWinding/layerType"].Value == 'Wound'):
    assignMaterial(oEditor, 'Bobbin', '"polyamide"')

  elif (step2.Properties["windingProperties/drawWinding/includeBobbin"].Value == True
    and step2.Properties["windingProperties/drawWinding/layerType"].Value == 'Planar'):
    boards = ','.join(oEditor.GetMatchedObjectName("Board*"))
    assignMaterial(oEditor, boards, '"polyamide"')

  # only for EFD core not to get an error due to section of winding
  if "CentralLegCS" in oEditor.GetCoordinateSystems():
    oEditor.SetWCS(
    	[
    		"NAME:SetWCS Parameter",
    		"Working Coordinate System:=", "CentralLegCS",
    		"RegionDepCSOk:="	, False
    	])


  oEditor.Section(
            [
              "NAME:Selections",
              "Selections:="    , ','.join(LayList),
              "NewPartsModelFlag:="   , "Model"
            ],
            [
              "NAME:SectionToParameters",
              "CreateNewObjects:="  , True,
              "SectionPlane:="  , "ZX",
              "SectionCrossObject:="  , False
            ])

  oEditor.SeparateBody(
            [
              "NAME:Selections",
              "Selections:="    , ','.join(LaySecList),
              "NewPartsModelFlag:="   , "Model"
            ])

  oEditor.Delete(["NAME:Selections","Selections:=",','.join(LayDelList)])

  oModule = oDesign.GetModule("BoundarySetup")
  ExcDict = {}
  for EachSec in LaySecList:
    SecList = EachSec.split("_")
    if SecList[0] in WdgSet.FinalPrimList:
      ExCName = SecList[0].replace("Layer","Primary")

      if ExCName not in ExcDict:
        ExcDict[ExCName] = []
      ExcDict[ExCName].append(EachSec.replace("Layer","Primary"))

      oModule.AssignCurrent(
                [
                  "NAME:"+EachSec.replace("Layer","Primary"),
                  "Objects:="   , [EachSec],
                  "Phase:="     , "0deg",
                  "Current:="   , "1A",
                  "IsSolid:="   , True,
                  "Point out of terminal:=", False
                ])
    if SecList[0] in WdgSet.FinalSecList:
      ExCName = SecList[0].replace("Layer","Secondary")

      if ExCName not in ExcDict:
        ExcDict[ExCName] = []
      ExcDict[ExCName].append(EachSec.replace("Layer","Secondary"))

      oModule.AssignCurrent(
                [
                  "NAME:"+EachSec.replace("Layer","Secondary"),
                  "Objects:="   , [EachSec],
                  "Phase:="     , "0deg",
                  "Current:="   , "1A",
                  "IsSolid:="   , True,
                  "Point out of terminal:=", True
                ])
  oModule.SetCoreLoss(CoreList, False)

  MatrixList = ["NAME:MatrixEntry"]
  MatrxGList = ["NAME:MatrixGroup"]
  for eachExG in ExcDict:
    for eachExIn in ExcDict[eachExG]:
      MatrixList.append(["NAME:MatrixEntry","Source:=", eachExIn,"NumberOfTurns:=", "1"])

  if ConnSet != None:
    # go through windings and set a group for matrix
    if len(ConnSet.FinalGroupDict.keys()) > 0:
      FGroupDict = {}
      FPDict = {}
      for Eachgr in ConnSet.FinalGroupDict:
        FGroupDict[Eachgr] = []
        FPDict[Eachgr] = ConnSet.FinalGroupDict[Eachgr][0]
        for EachInG in ConnSet.FinalGroupDict[Eachgr][1]:
          FGroupDict[Eachgr] = FGroupDict[Eachgr] + ExcDict[EachInG]
      for EachGxP in FGroupDict:
        MatrxGList.append(["NAME:MatrixGroup","GroupName:=", EachGxP,
                           "NumberOfBranches:=", FPDict[EachGxP],
                           "Sources:=",','.join(FGroupDict[EachGxP])])

  oModule = oDesign.GetModule("MaxwellParameterSetup")
  oModule.AssignMatrix(["NAME:Matrix1",MatrixList,MatrxGList])

  offset = int(step.Properties["defineSetup/offset"].Value)
  oEditor.CreateRegion(
    [
      "NAME:RegionParameters",
      "+XPaddingType:="   , "Percentage Offset",
      "+XPadding:="     , offset,
      "-XPaddingType:="   , "Percentage Offset",
      "-XPadding:="     , offset,
      "+YPaddingType:="   , "Percentage Offset",
      "+YPadding:="     , offset,
      "-YPaddingType:="   , "Percentage Offset",
      "-YPadding:="     , offset,
      "+ZPaddingType:="   , "Percentage Offset",
      "+ZPadding:="     , offset,
      "-ZPaddingType:="   , "Percentage Offset",
      "-ZPadding:="     , offset
    ],
    [
      "NAME:Attributes",
      "Name:="    , "Region",
      "Flags:="     , "Wireframe#",
      "Color:="     , "(255 0 0)",
      "Transparency:="  , 0,
      "PartCoordinateSystem:=", "Global",
      "UDMId:="     , "",
      "MaterialValue:="   , "\"vacuum\"",
      "SolveInside:="   , True
    ])
  oEditor.FitAll()

  maxNumPasses   = int(step.Properties["defineSetup/numPasses"].Value)
  percentError   = float(step.Properties["defineSetup/percError"].Value)
  frequency     = str(step.Properties["defineSetup/adaptFreq"].Value)+'Hz'

  startSweepFreq   = (str(step.Properties["defineSetup/freqSweep/startFreq"].Value)+
             str(step.Properties["defineSetup/freqSweep/freqSelect1"].Value))

  stopSweepFreq   = (str(step.Properties["defineSetup/freqSweep/stopFreq"].Value)+
             str(step.Properties["defineSetup/freqSweep/freqSelect2"].Value))

  samples     = int(step.Properties["defineSetup/freqSweep/samples"].Value)


  if step.Properties["defineSetup/freqSweep"].Value==True:
    if step.Properties["defineSetup/freqSweep/scale"].Value == 'Linear':
      insertSetup(oDesign, maxNumPasses, percentError, frequency, True,
                  'LinearCount', startSweepFreq, stopSweepFreq, samples)
    else:
      insertSetup(oDesign, maxNumPasses, percentError, frequency, True,
                  'LogScale', startSweepFreq, stopSweepFreq, samples)
  else:
    insertSetup(oDesign, maxNumPasses, percentError, frequency, False)

  EddyList = ["NAME:EddyEffectVector"]
  for EachLay in LayList:
    EddyList.append(
    [
      "NAME:Data",
      "Object Name:="   , EachLay,
      "Eddy Effect:="   , True,
      "Displacement Current:=", True
    ])

  oModule = oDesign.GetModule("BoundarySetup")
  oModule.SetEddyEffect(["NAME:Eddy Effect Setting",EddyList])
  MeshOpSz = max([float(Lx) for Lx in coreDimensionsSetup])/20.0
  oModule = oDesign.GetModule("MeshSetup")

  oModule.AssignLengthOp(
    [
      "NAME:Length_Coil",
      "RefineInside:="  , False,
      "Enabled:="   , True,
      "Objects:="   , LayList,
      "RestrictElem:="  , False,
      "NumMaxElem:="    , "1000",
      "RestrictLength:="  , True,
      "MaxLength:="     , str(MeshOpSz)+"mm"
    ])

  oModule.AssignLengthOp(
    [
      "NAME:Length_Core",
      "RefineInside:="  , False,
      "Enabled:="   , True,
      "Objects:="   , CoreList,
      "RestrictElem:="  , False,
      "NumMaxElem:="    , "1000",
      "RestrictLength:="  , True,
      "MaxLength:="     , str(MeshOpSz)+"mm"
    ])

  PrjPath=step.Properties["defineSetup/projPath"].Value

  if PrjPath == None:
    # that means that we are here due to auto run after reading a file
    return

  if PrjPath[-1:] != '/':
    PrjPath = PrjPath + '/'

  writeData(step)

  oProject.SaveAs(PrjPath+designName+'.aedt',True)
  step3.UserInterface.GetComponent("setupAnalysisButton").SetEnabledFlag("setupAnalysisButton", False)
  step3.UserInterface.GetComponent("defineConnectionsButton").SetEnabledFlag("defineConnectionsButton", False)
  step3.UserInterface.GetComponent("defineWindingsButton").SetEnabledFlag("defineWindingsButton", False)

  # mark that design was setup by user
  global flagAnalysisSet
  flagAnalysisSet = True


def assignMaterial(oEditor, selection, material):
  oEditor.AssignMaterial(
              [
                "NAME:Selections",
                "Selections:="    , selection
              ],
              [
                "NAME:Attributes",
                "MaterialValue:="   , material,
                "SolveInside:="   , True
              ])


def insertSetup(oDesign, maxNumPasses, percentError, frequency, hasSweep,
                sweepType = '', startSweepFreq = '', stopSweepFreq = '', samples = ''):
  oModule = oDesign.GetModule("AnalysisSetup")
  oModule.InsertSetup("EddyCurrent",
        [
          "NAME:Setup1",
          "Enabled:="   , True,
          "MaximumPasses:="   , maxNumPasses,
          "MinimumPasses:="   , 2,
          "MinimumConvergedPasses:=", 1,
          "PercentRefinement:="   , 30,
          "SolveFieldOnly:="  , False,
          "PercentError:="  , percentError,
          "SolveMatrixAtLast:="   , True,
          "PercentError:="  , percentError,
          "UseIterativeSolver:="  , False,
          "RelativeResidual:="  , 0.0001,
          "ComputeForceDensity:=" , False,
          "ComputePowerLoss:="  , False,
          "Frequency:="     , frequency,
          "HasSweepSetup:="   , hasSweep,
          "SweepSetupType:="  , sweepType,
          "StartValue:="    ,startSweepFreq ,
          "StopValue:="     , stopSweepFreq,
          "Samples:="     , samples,
          "SaveAllFields:="   , True,
          "UseHighOrderShapeFunc:=", False
        ])


def analyze(step):
  if flagAnalysisSet != True:
    setupAnalysis(step)
    oDesign = oProject.SetActiveDesign(designName)
    oDesign.Analyze("Setup1")

  else:
    try:
      oDesign = oProject.SetActiveDesign(designName)
      oDesign.Analyze("Setup1")
    except:
      pass


# write data to file
def writeData(step):
  global designName,filledDict
  WritePath = step3.Properties["defineSetup/projPath"].Value + '\\' + designName+ '_parameters.tab'

  txt= (str(step1.Properties["coreProperties/segAngle"].Value) +
        '\t\t\t%Segmentation Angle: should be between 0 to 20, 0 for True Surface\n'+
        "mm\t\t\t%Model Units: mm\n"+
        step1.Properties["coreProperties/supplier"].Value + '\t\t\t%Supplier Name\n'+
        step1.Properties["coreProperties/coreType"].Value + '\t\t\t%Core Type\n'+
        '\t'.join(coreDimensionsSetup)+'\t\t\t%CoreDimensions: D_1..D_8\n')

  if bool(step1.Properties["coreProperties/defAirgap"].Value) == True:
    AirGapOn = step1.Properties["coreProperties/defAirgap/airgapOn"].Value
    if AirGapOn =='Center Leg':
      txt+='1\t\t\t%Include Airgap: 0 to exclude, 1 for Airgap on central leg, 2 for Side leg, 3 for both\n'
    elif AirGapOn =='Side Leg':
      txt+='2\t\t\t%Include Airgap: 0 to exclude, 1 for Airgap on central leg, 2 for Side leg, 3 for both\n'
    else:
      txt+='3\t\t\t%Include Airgap: 0 to exclude, 1 for Airgap on central leg, 2 for Side leg, 3 for both\n'
    txt+=str(step1.Properties["coreProperties/defAirgap/airgapValue"].Value) + "\t\t\t% Airgap Value\n"
  else:
    txt+='0\t\t\t%Include Airgap: 0 to exclude, 1 for Airgap on central leg, 2 for Side leg, 3 for both\n'

  if bool(step2.Properties["windingProperties/drawWinding"].Value) == True:
    txt+=("1\t\t\t%Winding Status: 1 for Create Winding, 0 for exclude winding\n"+
        str(step2.Properties["windingProperties/drawWinding/numLayers"].Value)+'\t\t\t%Number of Layers\n'+
        str(step2.Properties["windingProperties/drawWinding/topMargin"].Value)  +'\t'+
        str(step2.Properties["windingProperties/drawWinding/sideMargin"].Value)   +'\t'+
        str(step2.Properties["windingProperties/drawWinding/layerSpacing"].Value) +'\t'+
        str(step2.Properties["windingProperties/drawWinding/bobThickness"].Value) +'\t'+
        '\t\t\t%Margin Dimensions (Top/Bottom Margin, Side Margin, Layer Spacing, Bobbin Thickness)\n')
    if float(step2.Properties["windingProperties/drawWinding/bobThickness"].Value)>0:
      if bool(step2.Properties["windingProperties/drawWinding/includeBobbin"].Value) == True:
        txt+='1\t\t\t%Bobbin Status 0:Exclude bobbin from Geometry 1:Include Bobbin in Geometry\n'
      else:
        txt+='0\t\t\t%Bobbin Status 0:Exclude bobbin from Geometry 1:Include Bobbin in Geometry\n'
    else:
      '0\t\t\t%Bobbin Status 0:Exclude bobbin from Geometry 1:Include Bobbin in Geometry\n'

    if step2.Properties["windingProperties/drawWinding/layerType"].Value == 'Planar':
      txt+='1\t\t\t%Winding Type 1:Planar 2:Wound\n'
    else:
      txt+='2\t\t\t%Winding Type 1:Planar 2:Wound\n'

    if step2.Properties["windingProperties/drawWinding/conductorType"].Value == 'Rectangular':
      txt+='1\t\t\t%Conductor Type 1:Rectangular 2:Circular\n'
      for i in range(1,int(step2.Properties["windingProperties/drawWinding/numLayers"].Value)+1):
        txt+=('\t'.join(str(x) for x in filledDict[i]) + '\t\t\t%Layer ' + str(i) +
              ' specifications :Conductor Width, Conductor Height, Number of Turns, Insulation Thickness\n' )
    else:
      txt+= '2\t\t\t%Conductor Type 1:Rectangular 2:Circular\n'
      for i in range(1,int(step2.Properties["windingProperties/drawWinding/numLayers"].Value)+1):
        txt+= ('\t'.join([str(filledDict[i][0]),str(filledDict[i][2]),str(filledDict[i][3]),str(filledDict[i][1])]) +
              '\t\t\t%Layer ' + str(i) +
              ' specifications :Conductor Diameter, Number of Turns, Insulation Thickness, Number of Segments\n')

  else:
    txt+='0\t\t\t%Winding Status: 1 for Create Winding, 0 for exclude winding\n'

  txt+=('1\t\t\t%Setup Defined: 0 for No, 1 for Define setup\n'+
      step.Properties["defineSetup/coreMaterial"].Value+'\t'+
      step.Properties["defineSetup/coilMaterial"].Value+'\t\t\t% Core Material and Coil Material\n')

  txt += '\t'.join([WrP.replace("Layer","") for WrP in WdgSet.FinalPrimList])+'\t\t\t%Layers Defined as Primary\n'
  txt += '\t'.join([WrS.replace("Layer","") for WrS in WdgSet.FinalSecList])+'\t\t\t%Layers Defined as Secondary\n'

  if ConnSet != None:
    if len(ConnSet.FinalGroupDict.keys()) == 0:
      txt+='0\t\t\t%No. of Winding Groups Defined: 0 for no definition\n'
    else:
      txt+=(str(len(ConnSet.FinalGroupDict.keys()))+'\t\t\t%No. of Winding Groups Defined: 0 for no definition\n')
      for WGrp in ConnSet.FinalGroupDict:
        WrTempG = []
        for WrG2 in ConnSet.FinalGroupDict[WGrp][1]:
          if "Primary" in WrG2:
            WrTempG.append(WrG2.replace("Primary",""))
          if "Secondary" in WrG2:
            WrTempG.append(WrG2.replace("Secondary",""))
        txt+=(WGrp+'\t'+'\t'.join(WrTempG)+'\t\t\t% Group Name followed by Layers in Group\n')
        txt+=(ConnSet.FinalGroupDict[WGrp][0]+'\t\t\t% No. of Parallel branches in Group\n')
  else:
    txt+=('0\t\t\t%No. of Winding Groups Defined: 0 for no definition\n')
  txt+=(str(step.Properties["defineSetup/adaptFreq"].Value) + 'Hz\t\t\t%Adaptive Frequency\n')

  if step.Properties["defineSetup/freqSweep"].Value==True:
    if step.Properties["defineSetup/freqSweep/scale"].Value == 'Linear':
      txt+=('1\t\t\t%Frequency Sweep Defined: 0 for No Sweep, 1 for Linear Sweep, 2 for Logarithmic Sweep\n')
    else:
      # Logarithmic
      txt+=('2\t\t\t%Frequency Sweep Defined: 0 for No Sweep, 1 for Linear Sweep, 2 for Logarithmic Sweep\n')

    txt+=((str(step.Properties["defineSetup/freqSweep/startFreq"].Value) +
         str(step.Properties["defineSetup/freqSweep/freqSelect1"].Value)) + '\t' +
         (str(step.Properties["defineSetup/freqSweep/stopFreq"].Value) +
         str(step.Properties["defineSetup/freqSweep/freqSelect2"].Value)) + '\t' +
         str(step.Properties["defineSetup/freqSweep/samples"].Value) +
         '\t\t\t%Sweep Definition: Start Frequency, Stop Frequency and Count/Samples\n')
  else:
    # no frequency sweep
    txt+=('0\t\t\t%Frequency Sweep Defined: 0 for No Sweep, 1 for Linear Sweep, 2 for Logarithmic Sweep\n')

  txt+= (str(step.Properties["defineSetup/numPasses"].Value) + '\t'+str(step.Properties["defineSetup/percError"].Value) +
         '\t\t\t%Solution Setup:Maximum No. of Passes and Percentage Error\n')

  txt+= (str(step.Properties["defineSetup/offset"].Value) +
          '\t\t\t%Region offset (can be skipped only with skipping of next step)\n')

  txt+= ('0' + '\t\t\t%1 for run model after read; 0 - read file, manual click to invoke next steps (can be skipped)\n')

  file = open(WritePath,"w")
  file.write(txt)
  file.close()


#read data from file
def readData():
  path = ExtAPI.UserInterface.UIRenderer.ShowFileOpenDialog('Text Files(*.txt;*.tab;)|*.txt;*.tab;')

  if path == None:
    return

  # array to contain all settings without comments from file
  Inparams = []
  with open(path, 'r') as settingsFile:
    for line in settingsFile:
      lineOut,null1,null2= line.partition('%')
      Inparams.append(lineOut)

  SegAngle = Inparams.pop(0).split()[0]
  try:
    float(SegAngle)
  except:
    return MsgBox("Incorrect Segment angle in text file", vbOKOnly, "Invalid Input")

  if not 0 <= float(SegAngle) < 20:
    return MsgBox("Incorrect Segment angle in text file", vbOKOnly, "Invalid Input")

  ModelUnits = (Inparams.pop(0).split()[0]).lower()
  if ModelUnits != "mm" and ModelUnits != "inches":
    return MsgBox("Incorrect model units in text file", vbOKOnly, "Invalid Input")

  SupName = Inparams.pop(0).split()[0]
  if SupName.lower() == 'ferroxcube':
    SupName='Ferroxcube'
  elif SupName.lower() == 'phillips':
    SupName = 'Phillips'
  else:
    return MsgBox("Incorrect Supplier name in text file", vbOKOnly, "Invalid Input")

  CoreType = Inparams.pop(0).split()[0]
  if CoreType not in Ferroxcube.keys():       #dangerous, if added core type in philips and not to feroxcube
    return MsgBox("Incorrect Core type in text file", vbOKOnly, "Invalid Input")

  #Read Core Dimensions
  Dim = Inparams.pop(0).split()


  # Read Airgap
  AgStat = Inparams.pop(0).split()[0]
  try:
    AgStat = int(AgStat)
  except:
    return MsgBox("Incorrect Airgap status in text file", vbOKOnly, "Invalid Input")

  if int(AgStat) > 0:
    AgVal = Inparams.pop(0).split()[0]
    try:
      AgVal = float(AgVal)
    except:
      return MsgBox("Incorrect Airgap value in text file", vbOKOnly, "Invalid Input")


  # Read Winding Status   (stepTwo)
  WdgStat = Inparams.pop(0).split()[0]
  try:
    WdgStat = int(WdgStat)
  except:
    return MsgBox("Incorrect Winding status in text file", vbOKOnly, "Invalid Input")

  if WdgStat > 0:
    if len(Inparams) == 0:
      return MsgBox("No winding parameters are added", vbOKOnly, "Invalid Input")

  #Read Number of Layers
    NumWdg = Inparams.pop(0).split()[0]
    try:
      NumWdg = int(NumWdg)
    except:
      return MsgBox("Incorrect input for number of layers in text file", vbOKOnly, "Invalid Input")

    if NumWdg <= 0:
      return MsgBox("Incorrect input for number of layers in text file", vbOKOnly, "Invalid Input")

    # Read Winding Margins
    margList = Inparams.pop(0).split()
    if len(margList) < 4:
      return MsgBox("Margin Parameters are incorrect", vbOKOnly, "Invalid Input")

    margList = [float(element) for element in margList[:4]]

    # Read Bobbin Status
    BobStat = Inparams.pop(0).split()[0]

    try:
      BobStat = int(BobStat)
    except:
      return MsgBox("Incorrect Bobbin status in text file", vbOKOnly, "Invalid Input")

    if BobStat > 0 and margList[3] == 0:
      return MsgBox("Include bobbin is checked but thickness is 0", vbOKOnly, "Invalid Input")

    # Read Winding Type
    WdgType = Inparams.pop(0).split()[0]
    try:
      WdgType = int(WdgType)
    except:
      return MsgBox("Incorrect input for winding type in text file", vbOKOnly, "Invalid Input")

    if WdgType != 1 and WdgType != 2:
      return MsgBox("Incorrect input for winding type in text file", vbOKOnly, "Invalid Input")

    # if planar and board thick == 0 and layer spacing == 0
    if WdgType == 2 and margList[3] == 0 and margList[2] == 0:
      return MsgBox("For planar transformer Board thickness and Layer spacing cannot be equal 0 at once", vbOKOnly, "Invalid Input")


    # Read Conductor Type
    CondType = Inparams.pop(0).split()[0]
    try:
      CondType = int(CondType)
    except:
      return MsgBox("Incorrect input for Conductor type in text file", vbOKOnly, "Invalid Input")

    if CondType != 1 and CondType != 2:
      return MsgBox("Incorrect input for Conductor type in text file", vbOKOnly, "Invalid Input")

    # Read winding Parameters
    LayerSpecDict = {}
    for EachLay in range(0,int(float(NumWdg))):
      TempSpecList = Inparams.pop(0).split()
      try:
        LayerSpecDict[EachLay+1] = [float(element) for element in TempSpecList[:4]]
      except:
        return MsgBox("Incorrect specifications for winding layers in text file", vbOKOnly, "Invalid Input")


  # Read Solution Setup Flag
  SetupDef = Inparams.pop(0).split()[0]
  try:
    SetupDef = int(SetupDef)
  except:
    return MsgBox("Incorrect Setup status in text file", vbOKOnly, "Invalid Input")

  if SetupDef > 0:
    Mat1List = Inparams.pop(0).split()
    MatList = Mat1List[:2]
    if len(MatList)< 2:
      return MsgBox("Incorrect Material definition in text file", vbOKOnly, "Invalid Input")

    if MatList[0] not in matDict.keys():
      return MsgBox("Defined core material not available in library", vbOKOnly, "Invalid Input")

    if MatList[1].lower() == 'copper':
      MatList[1]='Copper'
    elif MatList[1].lower() == 'aluminum':
      MatList[1] == 'Aluminum'
    else:
      return MsgBox("Defined coil material not available in library", vbOKOnly, "Invalid Input")

    Prim1List = filter(None,Inparams.pop(0).split())
    PrimList = []
    for EPrim in Prim1List:
      PrimList.append('Layer'+EPrim)
    Sec1List = filter(None,Inparams.pop(0).split())
    SecList = []
    for ESec in Sec1List:
      SecList.append('Layer'+ESec)
    if len(PrimList)+len(SecList) != int(NumWdg):
      return MsgBox("Incomplete primary and secondary definition", vbOKOnly, "Invalid Input")

    DefGrpDict = {}
    UndefList = []
    RDefList = []
    tempTDL = PrimList + SecList
    for Xund in tempTDL:
      if Xund in PrimList:
        UndefList.append(Xund.replace('Layer','Primary'))
      if Xund in SecList:
        UndefList.append(Xund.replace('Layer','Secondary'))
    NumGroups = Inparams.pop(0).split()[0]
    for RGrp in range(0, int(NumGroups)):
      TempGpL = filter(None,Inparams.pop(0).split())
      TempKy = TempGpL.pop(0)
      TempGpF = []
      for EGpL in TempGpL:
        if EGpL in Prim1List:
          TempGpF.append("Primary"+EGpL)
          UndefList.remove("Primary"+EGpL)
        elif EGpL in Sec1List:
          TempGpF.append("Secondary"+EGpL)
          UndefList.remove("Secondary"+EGpL)
        else:
          return MsgBox("Incorrect definition of groups", vbOKOnly, "Invalid Input")

      RDefList.append(TempKy+":"+(",".join(TempGpF)))
      TempPBr = Inparams.pop(0).split()[0]
      try:
        float(TempPBr)
      except:
        return MsgBox("Incorrect definition of groups", vbOKOnly, "Invalid Input")

      if not (float(BobStat).is_integer()):
        return MsgBox("Incorrect definition of groups", vbOKOnly, "Invalid Input")

      DefGrpDict[TempKy] = [TempPBr,TempGpF[:]]
    AdFr1 = Inparams.pop(0).split()[0]
    AdFrList = filter(None,re.split(r'([\d\.\d]+)',AdFr1))
    if not (AdFrList[1] in ["Hz", "kHz","MHz"]):
      return MsgBox("Incorrect frequency unit specified", vbOKOnly, "Invalid Input")

    if AdFrList[1] == 'kHz':
      AdFrVal = float(AdFrList[0]) * 1000
    elif AdFrList[1] == 'MHz':
      AdFrVal = float(AdFrList[0]) * 1000000
    elif AdFrList[1] == 'Hz':
      AdFrVal = float(AdFrList[0])

    FrsStat = Inparams.pop(0).split()[0]
    try:
      FrsStat = int(FrsStat)
    except:
      return MsgBox("Incorrect Frequency Sweep Status", vbOKOnly, "Invalid Input")

    if FrsStat> 0:
      FrsList1 = filter(None,Inparams.pop(0).split())[0:3]
      FrsList = [None]*6
      FrsList[3]= FrsStat-1
      TempFrList = []
      for EFr in FrsList1:
        TFrList = filter(None,re.split(r'([\d\.\d]+)',EFr))
        try:
          float(TFrList[0])
        except:
          return MsgBox("Incorrect Frequency Sweep Definition", vbOKOnly, "Invalid Input")

        if FrsList1.index(EFr) == 2:
          if not (float(TFrList[0]).is_integer()):
            return MsgBox("Incorrect Frequency Sweep Definition", vbOKOnly, "Invalid Input")

        else:
          if not (TFrList[1] in ["Hz", "kHz","MHz"]):
            return MsgBox("Incorrect frequency unit specified", vbOKOnly, "Invalid Input")

        TempFrList = TempFrList+TFrList
      FrsList[0] = TempFrList[0]
      FrsList[1] = TempFrList[2]
      FrsList[2] = TempFrList[4]
      FrsList[4] = TempFrList[1]
      FrsList[5] = TempFrList[3]

    SolSetL = filter(None,Inparams.pop(0).split())
    try:
      float(SolSetL[0])
      float(SolSetL[1])
    except:
      return MsgBox("Incorrect Solution settings specified", vbOKOnly, "Invalid Input")

    if not (float(SolSetL[0]).is_integer()):
      return MsgBox("Incorrect Solution settings specified", vbOKOnly, "Invalid Input")

    if len(Inparams) != 0 and Inparams[0] != '':
      offset = filter(None,Inparams.pop(0).split())[0]
      try:
        offset = float(offset)
      except:
        return MsgBox("Incorrect offset settings specified", vbOKOnly, "Invalid Input")
    else:
      offset = None

    if len(Inparams) != 0  and Inparams[0] != '':
      runSetSetup = filter(None,Inparams.pop(0).split())[0]
      try:
        runSetSetup = float(runSetSetup)
      except:
        return MsgBox("Incorrect value for running setup specified", vbOKOnly, "Invalid Input")
    else:
      runSetSetup = 0

  # Update GUI
  # Step One
  step1.Properties["coreProperties/segAngle"].Value = int(SegAngle)
  step1.Properties["coreProperties/supplier"].Value = SupName
  step1.Properties["coreProperties/coreType"].Value = CoreType

  try:
    for i in range(1,9):
      if Dim[i-1] != '' and ModelUnits == "inches":
        step1.Properties["coreProperties/coreType/D_" + str(i)].Value = float(Dim[i-1])/25.4

      elif Dim[i-1] != '' and ModelUnits == "mm":
        step1.Properties["coreProperties/coreType/D_" + str(i)].Value = float(Dim[i-1])

  except:
    return MsgBox("Incorrect Core dimensions in text file", vbOKOnly, "Invalid Input")

  if int(AgStat)>0:
    step1.Properties["coreProperties/defAirgap"].Value = True
    if int(AgStat) ==1:
      AirGapOn ='Center Leg'
    elif int(AgStat) ==2:
      AirGapOn ='Side Leg'
    elif int(AgStat) ==3:
      AirGapOn ='Both'
    step1.Properties["coreProperties/defAirgap/airgapOn"].Value = AirGapOn
    step1.Properties["coreProperties/defAirgap/airgapValue"].Value = AgVal if ModelUnits == 'mm' else AgVal/25.4
  else:
    step1.Properties["coreProperties/defAirgap"].Value = False

  if CoreType not in ['EP','ER','PQ','RM']:
    HTMLData = '<img width="300" height="200" src="' + str(ExtAPI.Extension.InstallDir) + '/images/'
  else:
    HTMLData = '<img width="275" height="360" src="' + str(ExtAPI.Extension.InstallDir) + '/images/'

  report = step1.UserInterface.GetComponent("coreImage")
  report.SetHtmlContent(HTMLData + CoreType + 'Core.png"/>')#set core names

  step1.UserInterface.GetComponent("Properties").UpdateData()
  step1.UserInterface.GetComponent("Properties").Refresh()

  # Step Two
  step2 = step1.Wizard.Steps["Winding"]
  if int(WdgStat)>0:
    step2.Properties["windingProperties/drawWinding"].Value = True

    step2.Properties["windingProperties/drawWinding/numLayers"].Value      = NumWdg
    step2.Properties["windingProperties/drawWinding/numLayers"].ReadOnly   = True
    step2.Properties["windingProperties/drawWinding/topMargin"].Value      = margList[0] if ModelUnits == 'mm' else margList[0]/25.4
    step2.Properties["windingProperties/drawWinding/sideMargin"].Value     = margList[1] if ModelUnits == 'mm' else margList[1]/25.4
    step2.Properties["windingProperties/drawWinding/layerSpacing"].Value   = margList[2] if ModelUnits == 'mm' else margList[2]/25.4
    step2.Properties["windingProperties/drawWinding/bobThickness"].Value   = margList[3] if ModelUnits == 'mm' else margList[3]/25.4
    step2.Properties["windingProperties/drawWinding/includeBobbin"].Value  = True if BobStat == 1 else False

    if CondType == 1:
    # Rectangular conductor
      step2.Properties["windingProperties/drawWinding/conductorType"].Value = 'Rectangular'
      table = step2.Properties["windingProperties/drawWinding/conductorType/tableLayers"]
      rowNum = table.RowCount
      for i in range(rowNum):
        table.DeleteRow(0)

      for i in range(1,int(NumWdg)+1):
        table.AddRow()
        table.Properties["conductorWidth"].Value  = LayerSpecDict[i][0] if ModelUnits == 'mm' else LayerSpecDict[i][0]/25.4
        table.Properties["conductorHeight"].Value = LayerSpecDict[i][1] if ModelUnits == 'mm' else LayerSpecDict[i][1]/25.4
        table.Properties["turnsNumber"].Value     = LayerSpecDict[i][2]
        table.Properties["insulationThick"].Value = LayerSpecDict[i][3] if ModelUnits == 'mm' else LayerSpecDict[i][3]/25.4
        table.Properties["layer"].Value       = 'Layer_' + str(i)
        table.SaveActiveRow()
    # Circular conductor
    else:
      step2.Properties["windingProperties/drawWinding/conductorType"].Value = 'Circular'
      table = step2.Properties["windingProperties/drawWinding/conductorType/tableLayersCircles"]
      rowNum = table.RowCount
      for i in range(rowNum):
        table.DeleteRow(0)

      for i in range(1,int(NumWdg)+1):
        table.AddRow()
        table.Properties["conductorDiameter"].Value = LayerSpecDict[i][0] if ModelUnits == 'mm' else LayerSpecDict[i][0]/25.4
        table.Properties["layerSegNumber"].Value    = LayerSpecDict[i][3]
        table.Properties["turnsNumber"].Value       = LayerSpecDict[i][1]
        table.Properties["insulationThick"].Value   = LayerSpecDict[i][2] if ModelUnits == 'mm' else LayerSpecDict[i][2]/25.4
        table.Properties["layer"].Value       = 'Layer_' + str(i)
        table.SaveActiveRow()

    # planar or wound
    if WdgType == 1:
      step2.Properties["windingProperties/drawWinding/layerType"].Value = 'Planar'
    else:
      step2.Properties["windingProperties/drawWinding/layerType"].Value = 'Wound'

    changeCaptions(step2, 'emptyArg', False)

  else:
    step2.Properties["windingProperties/drawWinding"].Value = False

  # Step Three
  if SetupDef > 0:
    step3 = step1.Wizard.Steps["setup"]

    global COREMATERIAL
    COREMATERIAL = MatList[0]
    step3.Properties["defineSetup/coreMaterial"].Value = MatList[0]
    step3.Properties["defineSetup/coilMaterial"].Value = MatList[1]
    step3.Properties["defineSetup/adaptFreq"].Value = AdFrVal
    step3.Properties["defineSetup/percError"].Value = float(SolSetL[1])
    step3.Properties["defineSetup/numPasses"].Value = float(SolSetL[0])

    if offset != None:
      step3.Properties["defineSetup/offset"].Value = int(offset)

    if FrsStat> 0:
      step3.Properties["defineSetup/freqSweep"].Value = True
      step3.Properties["defineSetup/freqSweep/startFreq"].Value   = float(FrsList[0])
      step3.Properties["defineSetup/freqSweep/freqSelect1"].Value = FrsList[4]
      step3.Properties["defineSetup/freqSweep/stopFreq"].Value    = float(FrsList[1])
      step3.Properties["defineSetup/freqSweep/freqSelect2"].Value = FrsList[5]
      step3.Properties["defineSetup/freqSweep/samples"].Value     = int(float(FrsList[2]))
      step3.Properties["defineSetup/freqSweep/scale"].Value = 'Logarithmic' if FrsStat== 2 else 'Linear'

    global WdgSet, ConnSet
    WdgSet = WdgForm()

    WdgSet.FinalPrimList = PrimList[:]
    WdgSet.FinalSecList = SecList[:]
    WdgSet.FinalWdgList = []

    WdgSet.FinalDefList = ([ExDef.replace("Layer","Primary") for ExDef in PrimList]+
                 [ExDef2.replace("Layer","Secondary") for ExDef2 in SecList])

    if len(DefGrpDict.keys()) > 0:
      if ConnSet == None:
        ConnSet = ConnForm()
      ConnSet.FinalInWdgList = UndefList
      ConnSet.FinalConnList = RDefList
      ConnSet.FinalGroupDict = DefGrpDict.copy()

    if int(runSetSetup) == 1:
      createModel(step2)
      setupAnalysis(step3)
      MsgBox("Analysis was set up successfully!", vbOKOnly, "Done!")


def changeColor(oEditor,selection,R,G,B):
  oEditor.ChangeProperty(
      ["NAME:AllTabs",["NAME:Geometry3DAttributeTab",
          ["NAME:PropServers"] + selection,
          ["NAME:ChangedProps",
            [
              "NAME:Color",
              "R:="			, R,
              "G:="			, G,
              "B:="			, B
            ]
          ]]])


# add rows in tables for second step
def addRows(step, prop):
  if prop.Name == 'numLayers' and float(prop.Value) <1:
      prop.Value = 1
      MsgBox(str(prop.Caption) + ' should be greater or eqaul then 1',vbOKOnly ,"Error")
      return False

  numLayersNew   = int(step.Properties["windingProperties/drawWinding/numLayers"].Value)  # number of slices

  if step.Properties["windingProperties/drawWinding/conductorType"].Value == 'Rectangular':
    XMLpathToTable = "windingProperties/drawWinding/conductorType/tableLayers"
    table = step.Properties[XMLpathToTable]
    flag = 'rect'
  else:
    XMLpathToTable = "windingProperties/drawWinding/conductorType/tableLayersCircles"
    table = step.Properties[XMLpathToTable]
    flag= 'circ'

  rowNum = table.RowCount
  if numLayersNew<rowNum:
    for i in range(0, rowNum-numLayersNew+1):
      table.DeleteRow(rowNum-i)
  elif numLayersNew > rowNum:
    if flag == 'rect':
      for i in range(rowNum+1,numLayersNew+1):
        table.AddRow()
        table.Properties["conductorWidth"].Value   = table.Value[XMLpathToTable + "/conductorWidth"][rowNum-1]
        table.Properties["conductorHeight"].Value  = table.Value[XMLpathToTable + "/conductorHeight"][rowNum-1]
        table.Properties["turnsNumber"].Value      = table.Value[XMLpathToTable + "/turnsNumber"][rowNum-1]
        table.Properties["insulationThick"].Value  = table.Value[XMLpathToTable + "/insulationThick"][rowNum-1]
        table.Properties["layer"].Value       = 'Layer_' + str(i)
        table.SaveActiveRow()
    elif flag == 'circ':
      for i in range(rowNum+1,numLayersNew+1):
        table.AddRow()
        table.Properties["conductorDiameter"].Value = table.Value[XMLpathToTable + "/conductorDiameter"][rowNum-1]
        table.Properties["layerSegNumber"].Value    = table.Value[XMLpathToTable + "/layerSegNumber"][rowNum-1]
        table.Properties["turnsNumber"].Value       = table.Value[XMLpathToTable + "/turnsNumber"][rowNum-1]
        table.Properties["insulationThick"].Value   = table.Value[XMLpathToTable + "/insulationThick"][rowNum-1]
        table.Properties["layer"].Value       = 'Layer_' + str(i)
        table.SaveActiveRow()


# initialize tables in step two
def InitTabularData(step):
  table = step.Properties["windingProperties/drawWinding/conductorType/tableLayers"]
  table.AddRow()
  table.Properties["conductorWidth"].Value = 0.2
  table.Properties["conductorHeight"].Value = 0.2
  table.Properties["turnsNumber"].Value = 2
  table.Properties["insulationThick"].Value = 0.05
  table.Properties["layer"].Value = 'Layer_1'
  table.SaveActiveRow()

  table = step.Properties["windingProperties/drawWinding/conductorType/tableLayersCircles"]
  table.AddRow()
  table.Properties["conductorDiameter"].Value = 0.2
  table.Properties["layerSegNumber"].Value = 8
  table.Properties["turnsNumber"].Value = 2
  table.Properties["insulationThick"].Value = 0.05
  table.Properties["layer"].Value = 'Layer_1'
  table.SaveActiveRow()


# invoked to change image and core dimensions when supplier or core type changed
def showCoreIMG(step,prop):
  step.Properties["coreProperties/coreType/coreModel"].Options.Clear()
  supl = step.Properties["coreProperties/supplier"].Value
  prop = step.Properties["coreProperties/coreType"].Value

  if prop not in ['EP','ER','PQ','RM']:
    HTMLData = '<img width="300" height="200" src="' + str(ExtAPI.Extension.InstallDir) + '/images/'
  else:
    HTMLData = '<img width="275" height="360" src="' + str(ExtAPI.Extension.InstallDir) + '/images/'

  report = step.UserInterface.GetComponent("coreImage")
  report.SetHtmlContent(HTMLData + prop + 'Core.png"/>')#set core names

  coreList = coreTypes(prop, supl)
  for i in range(len(coreList)):
    step.Properties["coreProperties/coreType/coreModel"].Options.Add(coreList[i][0])
  step.Properties["coreProperties/coreType/coreModel"].Value = coreList[0][0]

  #set core values for selected type
  for j in range(1,9):
    try:
      step.Properties["coreProperties/coreType/D_"+str(j)].Value = float(coreList[0][j])
      step.Properties["coreProperties/coreType/D_"+str(j)].Visible = True
    except:
      step.Properties["coreProperties/coreType/D_"+str(j)].Visible = False
  report.Refresh()
  step.UserInterface.GetComponent("Properties").UpdateData()
  step.UserInterface.GetComponent("Properties").Refresh()


# invoke when Core Model is changed
def insertDefaultValues(step,prop):
  supl = step.Properties["coreProperties/supplier"].Value
  prop = step.Properties["coreProperties/coreType"].Value

  #set core names
  coreList = coreTypes(prop, supl)
  for i in range(len(coreList)):
    if coreList[i][0] == step.Properties["coreProperties/coreType/coreModel"].Value:
      #set core values for selected type
      for j in range(1,9):
        try:
          step.Properties["coreProperties/coreType/D_"+str(j)].Value = float(coreList[i][j])
          step.Properties["coreProperties/coreType/D_"+str(j)].Visible = True
        except:
          step.Properties["coreProperties/coreType/D_"+str(j)].Value = -1
          step.Properties["coreProperties/coreType/D_"+str(j)].Visible = False
      step.UserInterface.GetComponent("Properties").UpdateData()
      step.UserInterface.GetComponent("Properties").Refresh()
      break


# invoke on step initialisation
def initializeData(step):
  step.Properties["coreProperties/coreType/coreModel"].Options.Clear()
  supl = step.Properties["coreProperties/supplier"].Value
  prop = step.Properties["coreProperties/coreType"].Value

  #set core names
  coreList = coreTypes(prop, supl)
  for i in range(len(coreList)):
    step.Properties["coreProperties/coreType/coreModel"].Options.Add(coreList[i][0])
  step.Properties["coreProperties/coreType/coreModel"].Value = coreList[0][0]

  #set core values for selected type
  for j in range(1,9):
    step.Properties["coreProperties/coreType/D_"+str(j)].Value = float(coreList[0][j])


# change captions depending on Wound or Planar transformer
def changeCaptions(step, prop, needRefresh = True):
  if step.Properties["windingProperties/drawWinding/layerType"].Value == 'Planar':
    step.Properties["windingProperties/drawWinding/bobThickness"].Caption = 'Board thickness:'
    step.Properties["windingProperties/drawWinding/includeBobbin"].Caption = 'Include board in geometry:'
    step.Properties["windingProperties/drawWinding/topMargin"].Caption = 'Bottom Margin'

    table = step.Properties["windingProperties/drawWinding/conductorType/tableLayers"]
    table.Properties["insulationThick"].Caption = 'Turn Spacing'

    step.Properties["windingProperties/drawWinding/conductorType"].Value = 'Rectangular'
    step.Properties["windingProperties/drawWinding/conductorType"].ReadOnly = True

  elif step.Properties["windingProperties/drawWinding/layerType"].Value == 'Wound':
    step.Properties["windingProperties/drawWinding/bobThickness"].Caption = 'Bobbin thickness:'
    step.Properties["windingProperties/drawWinding/includeBobbin"].Caption = 'Include bobbin in geometry:'
    step.Properties["windingProperties/drawWinding/topMargin"].Caption = 'Top Margin'

    table = step.Properties["windingProperties/drawWinding/conductorType/tableLayers"]
    table.Properties["insulationThick"].Caption = 'Insulation thickness'

    step.Properties["windingProperties/drawWinding/conductorType"].ReadOnly = False

  # do not need it if invoke read input from file
  if needRefresh == True:
    step.UserInterface.GetComponent("Properties").UpdateData()
    step.UserInterface.GetComponent("Properties").Refresh()


# create buttons and HTML data for first step
def CreateButtonsCore(step):
  global step1
  step1 = step

  updateBtnSession = step.UserInterface.GetComponent("readData")
  updateBtnSession.AddButton("readButton", "Read Settings File", ButtonPositionType.Left)
  updateBtnSession.ButtonClicked += readDataClick

  updateBtnSession = step.UserInterface.GetComponent("helpButton")
  updateBtnSession.AddButton("helpButton", "Help", ButtonPositionType.Center)
  updateBtnSession.ButtonClicked += helpClick
  updateBtnSession.AddCSSProperty("background-color", "#3383ff", "button")
  updateBtnSession.AddCSSProperty("color", "white", "button")

  # set core dimensions image according to core type
  prop = step.Properties["coreProperties/coreType"].Value
  if prop not in ['EP','ER','PQ','RM']:
    HTMLData = '<img width="300" height="200" src="' + str(ExtAPI.Extension.InstallDir) + '/images/'
  else:
    HTMLData = '<img width="275" height="360" src="' + str(ExtAPI.Extension.InstallDir) + '/images/'

  image = step.UserInterface.GetComponent("coreImage")
  image.SetHtmlContent(HTMLData + prop + 'Core.png"/>')
  image.Refresh()

  # set core names
  step.Properties["coreProperties/coreType/coreModel"].Options.Clear()
  supl = step.Properties["coreProperties/supplier"].Value
  coreList = coreTypes(prop, supl)

  for i in range(len(coreList)):
    step.Properties["coreProperties/coreType/coreModel"].Options.Add(coreList[i][0])
  step.Properties["coreProperties/coreType/coreModel"].Value = coreList[0][0]

  # set core values for selected type
  try:
    for j in range(1,9):
      step.Properties["coreProperties/coreType/D_"+str(j)].Value = float(coreList[0][j])
  except:
    pass

  global ConnSet,WdgSet,flagAnalysisSet
  ConnSet = WdgSet = flagAnalysisSet = None


def CreateButtonsWinding(step):
  global step2
  step2 = step

  updateBtnSession = step.UserInterface.GetComponent("helpButton")
  updateBtnSession.AddButton("helpButton", "Help", ButtonPositionType.Center)
  updateBtnSession.ButtonClicked += helpClick
  updateBtnSession.AddCSSProperty("background-color", "#3383ff", "button")
  updateBtnSession.AddCSSProperty("color", "white", "button")

  step.UserInterface.GetComponent("Properties").UpdateData()
  step.UserInterface.GetComponent("Properties").Refresh()


def CreateButtonsSetup(step):
  global step3
  step3 = step

  updateBtnSession = step.UserInterface.GetComponent("helpButton")
  updateBtnSession.AddButton("helpButton", "Help", ButtonPositionType.Center)
  updateBtnSession.ButtonClicked += helpClick
  updateBtnSession.AddCSSProperty("background-color", "#3383ff", "button")
  updateBtnSession.AddCSSProperty("color", "white", "button")

  updateBtnSession = step.UserInterface.GetComponent("analyzeButton")
  updateBtnSession.AddButton("analyzeButton", "Analyze", ButtonPositionType.Center, False)
  updateBtnSession.ButtonClicked += analyzeClick

  updateBtnSession = step.UserInterface.GetComponent("setupAnalysisButton")
  updateBtnSession.AddButton("setupAnalysisButton", "Setup Analysis", ButtonPositionType.Center, False)
  updateBtnSession.ButtonClicked += setupAnalysisClick

  updateBtnSession = step.UserInterface.GetComponent("defineWindingsButton")
  updateBtnSession.AddButton("defineWindingsButton", "Define Windings", ButtonPositionType.Center, False)
  updateBtnSession.ButtonClicked += defineWindingsClick

  updateBtnSession = step.UserInterface.GetComponent("defineConnectionsButton")
  updateBtnSession.AddButton("defineConnectionsButton", "Define Connections", ButtonPositionType.Center, False)
  updateBtnSession.ButtonClicked += defineConnectionsClick


  # insert path to the saved project as default path to label
  oProject = oDesktop.GetActiveProject()
  path=oProject.GetPath()
  if oProject== None:
    path=oDesktop.GetProjectPath()
  step.Properties["defineSetup/projPath"].Value = str(path)


  # add materials from vocabulary
  global COREMATERIAL # came if materials data was read from input file
  step.Properties["defineSetup/coreMaterial"].Options.Clear()
  for key in sorted(matDict):
    step.Properties["defineSetup/coreMaterial"].Options.Add(key)

  try:
    if COREMATERIAL != None:
      step.Properties["defineSetup/coreMaterial"].Value = COREMATERIAL
  except:
    step.Properties["defineSetup/coreMaterial"].Value = key


  if step2.Properties["windingProperties/drawWinding"].Value == True:
    step3.UserInterface.GetComponent("defineWindingsButton").SetEnabledFlag("defineWindingsButton", True)

  step.UserInterface.GetComponent("Properties").UpdateData()
  step.UserInterface.GetComponent("Properties").Refresh()

  global WdgSet,step2
  if WdgSet==None or step2.Properties["windingProperties/drawWinding/numLayers"].ReadOnly != True:
    WdgSet = WdgForm()
    WdgSet.FinalWdgList = []
    WdgSet.FinalDefList = []

    WdgSet.FinalPrimList = []
    WdgSet.FinalSecList = []

    for NumLay in range(0, int(step2.Properties["windingProperties/drawWinding/numLayers"].Value)):
      WdgSet.FinalWdgList.append('Layer'+str(NumLay+1))

  if (len(WdgSet._WgList.Items) == 0 and
    len(WdgSet.FinalDefList) !=0 and
    step2.Properties["windingProperties/drawWinding"].Value == True):
      step3.UserInterface.GetComponent("analyzeButton").SetEnabledFlag("analyzeButton", True)
      step3.UserInterface.GetComponent("setupAnalysisButton").SetEnabledFlag("setupAnalysisButton", True)
      step3.UserInterface.GetComponent("defineConnectionsButton").SetEnabledFlag("defineConnectionsButton", True)


def showConnectionDialog(allNamesList):
  global ConnSet
  if ConnSet == None:
    ConnSet = ConnForm()
    ConnSet.FinalInWdgList = []
    ConnSet.FinalConnList = []

    for EachConn in allNamesList:
      ConnSet.FinalInWdgList.append(EachConn)
    ConnSet.FinalGroupDict = {}
  ConnSet._InWgList.Items.Clear()
  for EachItem2 in ConnSet.FinalInWdgList:
    ConnSet._InWgList.Items.Add(EachItem2)
  ConnSet._ConnList.Items.Clear()
  for EachItem2 in ConnSet.FinalConnList:
    ConnSet._ConnList.Items.Add(EachItem2)
  ConnSet.GroupDict = ConnSet.FinalGroupDict.copy()
  ConnSet.ShowDialog()


def defineWindingsClick(sender, args):
  if ConnSet != None:
    if len(ConnSet.FinalConnList) > 0:
      if MsgBox('Changing Winding definition will remove defined connection\n'+
            'Do you wish to continue?', vbYesNo ,"Warning!") == 'No':
        return
      orDefList = WdgSet.FinalDefList [:]

  WdgSet._WgList.Items.Clear()
  WdgSet._DefList.Items.Clear()

  for EachItem in WdgSet.FinalWdgList:
    WdgSet._WgList.Items.Add(EachItem)
  WdgSet._DefList.Items.Clear()

  for EachItem in WdgSet.FinalDefList:
    WdgSet._DefList.Items.Add(EachItem)

  WdgSet.PrimList = WdgSet.FinalPrimList[:]
  WdgSet.SecList= WdgSet.FinalSecList[:]
  WdgSet.ShowDialog()

  if len(WdgSet.FinalWdgList) == 0:
    step3.UserInterface.GetComponent("analyzeButton").SetEnabledFlag("analyzeButton", True)
    step3.UserInterface.GetComponent("setupAnalysisButton").SetEnabledFlag("setupAnalysisButton", True)
    step3.UserInterface.GetComponent("defineConnectionsButton").SetEnabledFlag("defineConnectionsButton", True)

  if ConnSet != None:
    if len(ConnSet.FinalConnList) > 0:
      ModDList = list(set(orDefList).difference(set(WdgSet.FinalDefList)))
      if len(ModDList) > 0:
        for OrText in ModDList:
          KeyDict = ConnSet.FinalGroupDict.keys()[:]
          for EachKey in KeyDict:
            if OrText in ConnSet.FinalGroupDict[EachKey][1]:
              RemList = ConnSet.FinalGroupDict.pop(EachKey, None)

          OrConnList = ConnSet.FinalConnList[:]
          for EachConnx in OrConnList:
            if OrText in EachConnx:
              ConnSet.FinalConnList.remove(EachConnx)

          for EachRem in RemList[1]:
            if EachRem == OrText:
              if "Primary" in OrText:
                AddRem = EachRem.replace("Primary","Secondary")
              elif "Secondary" in OrText:
                AddRem = EachRem.replace("Secondary","Primary")

            else:
              AddRem= EachRem
            ConnSet.FinalInWdgList.append(AddRem)


def helpClick(sender, args):
  webopen(str(ExtAPI.Extension.InstallDir) + '/help/help.html')


def defineConnectionsClick(sender, args):
  listOfWindings= ([winding.replace("Layer","Secondary") for winding in WdgSet.FinalSecList] +
           [winding.replace("Layer","Primary") for winding in WdgSet.FinalPrimList])

  listOfWindings.sort(key=lambda x: '{0:0>20}'.format(x).lower())
  showConnectionDialog(listOfWindings)


def readDataClick(sender, args):
  try:
    readData()
  except:
    raise NameError('Unknown Error. Code 1.  Please contact technical support')


def setupAnalysisClick(sender, args):
  try:
    setupAnalysis(step3)
  except:
    raise NameError('Unknown Error. Code 2.  Please contact technical support')


def analyzeClick(sender, args):
  try:
    analyze(step3)
  except:
    raise NameError('Unknown Error. Code 3.  Please contact technical support')


def resetFunc(step):
  pass

def resetFunc2(step):
    global flagAnalysisSet
    flagAnalysisSet = False
