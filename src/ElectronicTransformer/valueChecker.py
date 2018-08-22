def segAngCheck(step, prop):
  if not (0<= float(prop.Value) <20):
    prop.Value = 0
    MsgBox(str(prop.Caption) + ' should be 0 for True Surface or less than 20',vbOKOnly ,"Error")
    return False
  return True


def GEThenZero(step, prop):
  if float(prop.Value) <0:
    prop.Value = 0
    MsgBox(str(prop.Caption) + ' cannot be negative',vbOKOnly ,"Error")
    return False
  return True


def GThenZero(step, prop):
  if float(prop.Value) <=0:
    prop.Value = 1
    MsgBox(str(prop.Caption) + ' should be greater than 0',vbOKOnly ,"Error")
    return False
  return True


def GEThenOne(step, prop):
  if float(prop.Value) <1:
    prop.Value = 1
    MsgBox(str(prop.Caption) + ' should be greater or eqaul than 1',vbOKOnly ,"Error")
    return False
  return True


def validityCheckTable(step, prop):
  if step.Properties["windingProperties/drawWinding/conductorType"].Value == 'Rectangular':
    table = step.Properties["windingProperties/drawWinding/conductorType/tableLayers"]
    flag = 'tableLayers'
  else:
    table = step.Properties["windingProperties/drawWinding/conductorType/tableLayersCircles"]
    flag= 'tableLayersCircles'

  for i in range (table.RowCount):
    if prop.Name != 'layerSegNumber':
      if float(table.Value["windingProperties/drawWinding/conductorType/" + flag + '/' + prop.Name][i]) <= 0:
        table.Value["windingProperties/drawWinding/conductorType/" + flag + '/' + prop.Name][i] = 1
        MsgBox(str(prop.Caption) +' must be greater then zero!',vbOKOnly ,"Error")
    else:
      if (float(table.Value["windingProperties/drawWinding/conductorType/" + flag + '/' + prop.Name][i]) < 8 and
        float(table.Value["windingProperties/drawWinding/conductorType/" + flag + '/' + prop.Name][i]) !=0):
        table.Value["windingProperties/drawWinding/conductorType/" + flag + '/' + prop.Name][i] = 8
        MsgBox(str(prop.Caption) +' must be Zero for True Surface or greater then 7!',vbOKOnly ,"Error")


def checkBobbin(step):
  bobbinThickness = float(step.Properties["windingProperties/drawWinding/bobThickness"].Value)
  includeBobbin = bool(step.Properties["windingProperties/drawWinding/includeBobbin"].Value)
  boardThickness  = float(step.Properties["windingProperties/drawWinding/bobThickness"].Value)
  layerSpacing    = float(step.Properties["windingProperties/drawWinding/layerSpacing"].Value)

  if bobbinThickness == 0 and includeBobbin == True:
    raise UserErrorMessageException("Include board/bobbin is checked but thickness is 0")

  if (step.Properties["windingProperties/drawWinding/layerType"].Value == 'Planar' and
      boardThickness == 0 and layerSpacing == 0):
    raise UserErrorMessageException("For planar transformer Board thickness and Layer spacing cannot be equal 0 at once")


def checkWinding(step):
  if bool(step.Properties["windingProperties/drawWinding/skipCheck"].Value) == True:
    return

  D_2 = float(STEP1.Properties["coreProperties/coreType/D_2"].Value)
  D_3 = float(STEP1.Properties["coreProperties/coreType/D_3"].Value)
  D_4 = float(STEP1.Properties["coreProperties/coreType/D_4"].Value)
  D_5 = float(STEP1.Properties["coreProperties/coreType/D_5"].Value)

  sideMargin      = float(step.Properties["windingProperties/drawWinding/sideMargin"].Value)
  bobbinThickness = float(step.Properties["windingProperties/drawWinding/bobThickness"].Value)
  boardThickness  = float(step.Properties["windingProperties/drawWinding/bobThickness"].Value)
  topMargin       = float(step.Properties["windingProperties/drawWinding/topMargin"].Value)
  layerSpacing    = float(step.Properties["windingProperties/drawWinding/layerSpacing"].Value)

  CoreType = STEP1.Properties["coreProperties/coreType"].Value
  conductorType = step.Properties["windingProperties/drawWinding/conductorType"].Value

  # ---- start checking for wound ---- #
  if step.Properties["windingProperties/drawWinding/layerType"].Value == 'Wound':
    # ---- Check possible width for wound---- #
    if conductorType == 'Rectangular':
      XMLpathToTable = 'windingProperties/drawWinding/conductorType/tableLayers'
      table = step.Properties[XMLpathToTable]
      # take sum of layer dimensions where one layer is: Width + 2 * Insulation
      maximumLayer = sum(
                            [
                              table.Value[XMLpathToTable + '/conductorWidth'][i] +
                               # do not forget that insulation is on both sides
                               2 * table.Value[XMLpathToTable + '/insulationThick'][i] +
                               layerSpacing # since number of layers - 1 for spacing)
                              for i in range(len(table.Value[XMLpathToTable + '/layer']))
                            ]
                          ) - layerSpacing

    elif conductorType == 'Circular':
      XMLpathToTable = 'windingProperties/drawWinding/conductorType/tableLayersCircles'
      table = step.Properties[XMLpathToTable]
      maximumLayer = sum(
                            [
                              (table.Value[XMLpathToTable + '/conductorDiameter'][i] +
                               2 * table.Value[XMLpathToTable + '/insulationThick'][i] + layerSpacing)
                              for i in range(len(table.Value[XMLpathToTable + '/layer']))
                            ]
                          ) - layerSpacing

    # do not forget that windings are laying on both sides of the core
    maximumPossibleWidth = 2 * (bobbinThickness + sideMargin + maximumLayer)

    # ---- Check possible height for wound---- #
    if conductorType == 'Rectangular':
      XMLpathToTable = 'windingProperties/drawWinding/conductorType/tableLayers'
      table = step.Properties[XMLpathToTable]
      # max value from each layer: (Height + 2 * Insulation) * number of layers
      maximumLayer = max(
                            [
                              ((table.Value[XMLpathToTable + '/conductorHeight'][i] +
                               # do not forget that insulation is on both sides
                               2 * table.Value[XMLpathToTable + '/insulationThick'][i])*
                               table.Value[XMLpathToTable + '/turnsNumber'][i])
                              for i in range(len(table.Value[XMLpathToTable + '/layer']))
                            ]
                          )

    elif conductorType == 'Circular':
      XMLpathToTable = 'windingProperties/drawWinding/conductorType/tableLayersCircles'
      table = step.Properties[XMLpathToTable]
      maximumLayer = max(
                            [
                              ((table.Value[XMLpathToTable + '/conductorDiameter'][i] +
                               2 * table.Value[XMLpathToTable + '/insulationThick'][i])*
                               table.Value[XMLpathToTable + '/turnsNumber'][i])
                              for i in range(len(table.Value[XMLpathToTable + '/layer']))
                            ]
                          )

    maximumPossibleHeight = (2 * bobbinThickness + topMargin + maximumLayer)
  # ---- Wound type limit found ---- #


  elif step.Properties["windingProperties/drawWinding/layerType"].Value == 'Planar':
    XMLpathToTable = 'windingProperties/drawWinding/conductorType/tableLayers'
    table = step.Properties[XMLpathToTable]
    # ---- Check width for planar---- #
    maximumLayer = max(
                          [
                            ((table.Value[XMLpathToTable + '/conductorWidth'][i] +
                             # in this case it is turn spacing (no need to x2)
                             table.Value[XMLpathToTable + '/insulationThick'][i])*
                             table.Value[XMLpathToTable + '/turnsNumber'][i])
                            for i in range(len(table.Value[XMLpathToTable + '/layer']))
                          ]
                        )
    maximumPossibleWidth = (2 * maximumLayer + 2 * sideMargin)

    # ---- Check Height for planar ---- #
    maximumLayer = sum(
                          [
                            (table.Value[XMLpathToTable + '/conductorHeight'][i] + boardThickness + layerSpacing)
                            for i in range(len(table.Value[XMLpathToTable + '/layer']))
                          ]
                        ) - layerSpacing

    maximumPossibleHeight = maximumLayer + topMargin
    # ---- Planar type limit found ---- #


  # ---- Check accomodation not depending on layer type ---- #
  # ---- Height ---- #
  if CoreType  in  ["E","EC","EFD","EQ","ER","ETD","PH"]:
    # D_5 is height of one half core
    if (maximumPossibleHeight > 2 * D_5):
      raise UserErrorMessageException("Cannot accommodate all windings, increase D_5")
  elif CoreType in ["EI","EP","P","PT","PQ","RM"]:
    # D_5 is height of core
    if (maximumPossibleHeight > D_5):
      raise UserErrorMessageException("Cannot accommodate all windings, increase D_5")
  elif CoreType == "UI":
     # D_4 is height of core
     if (maximumPossibleHeight > D_4):
       raise UserErrorMessageException("Cannot accommodate all windings, increase D_4")
  elif CoreType == "U":
     # D_4 is height of core
     if (maximumPossibleHeight > 2*D_4):
       raise UserErrorMessageException("Cannot accommodate all windings, increase D_4")

  # ---- Width ---- #
  if CoreType  not in  ["U","UI"]:
    # D_2 - D_3 is sum of dimesnsions of two slots for windngs (left + right)
    if (maximumPossibleWidth > (D_2 - D_3)):
      raise UserErrorMessageException("Cannot accommodate all windings, increase D_2")
  else:
    # D_2 is dimension of one side slot for winding
    if (maximumPossibleWidth/2 > D_2):
      raise UserErrorMessageException("Cannot accommodate all windings, increase D_2")


def CheckCoreDim(step):
  # stop invoking for unsupported versions
  if float(oDesktop.GetVersion()[:-2]) < 2018.1:
    raise UserErrorMessageException("Electronics Desktop Version is unsupported. Please use version r19.1 or higher")

  D_1 = float(step.Properties["coreProperties/coreType/D_1"].Value)
  D_2 = float(step.Properties["coreProperties/coreType/D_2"].Value)
  D_3 = float(step.Properties["coreProperties/coreType/D_3"].Value)
  D_4 = float(step.Properties["coreProperties/coreType/D_4"].Value)
  D_5 = float(step.Properties["coreProperties/coreType/D_5"].Value)
  D_6 = float(step.Properties["coreProperties/coreType/D_6"].Value)
  D_7 = float(step.Properties["coreProperties/coreType/D_7"].Value)
  D_8 = float(step.Properties["coreProperties/coreType/D_8"].Value)

  if (D_1 == 0):
  	raise UserErrorMessageException("Wrong core parameters!" + '\nD_1 cannot be zero')
  if (D_2 == 0):
  	raise UserErrorMessageException("Wrong core parameters!" + '\nD_2 cannot be zero')
  if (D_3 == 0):
  	raise UserErrorMessageException("Wrong core parameters!" + '\nD_3 cannot be zero')
  if (D_4 == 0):
  	raise UserErrorMessageException("Wrong core parameters!" + '\nD_4 cannot be zero')
  if (D_5 == 0):
  	raise UserErrorMessageException("Wrong core parameters!" + '\nD_5 cannot be zero')
  if (D_1 <= D_2):
  	raise UserErrorMessageException("Wrong core parameters!" + '\nD_1 must be greater then D_2')

  CoreType = step.Properties["coreProperties/coreType"].Value
  if CoreType  ==  "E":
    if (D_2 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_2 must be greater then D_3')
    if (D_4 <= D_5):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_4 must be greater then D_5')
    if (D_6 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_6 cannot be zero')

  elif CoreType  ==  "EC":
    if (D_2 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_2 must be greater then D_3')
    if (D_4 <= D_5):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_4 must be greater then D_5')
    if (D_6 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_6 cannot be zero')

  elif CoreType  ==  "EFD":
    if (D_2 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_2 must be greater then D_3')
    if (D_4 <= D_5):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_4 must be greater then D_5')
    if (D_6 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_6 cannot be zero')
    if (D_7 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_7 cannot be zero')

  elif CoreType  ==  "EI":
    if (D_2 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_2 must be greater then D_3')
    if (D_4 <= D_5):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_4 must be greater then D_5')
    if (D_6 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_6 cannot be zero')

  elif CoreType  ==  "EP":
    if (D_2 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_2 must be greater then D_3')
    if (D_4 <= D_5):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_4 must be greater then D_5')
    if (D_6<= D_7):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_6 must be greater then D_7')
    if (D_6 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_6 cannot be zero')
    if (D_7 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_7 cannot be zero')

  elif CoreType  ==  "EQ":
    if (D_2 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_2 must be greater then D_3')
    if (D_4 <= D_5):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_4 must be greater then D_5')
    if (D_6 < D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_6 must be greater then D_3')
    if (D_6 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_6 cannot be zero')

  elif CoreType  ==  "ER":
    if (D_2 <= D_7):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_2 must be greater then D_7')
    if (D_4 <= D_5):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_4 must be greater then D_5')
    if (D_7 < 2*math.sqrt((D_2/2)**2 -(D_6/2)**2) and D_7 != 0):
      raise UserErrorMessageException("Wrong core parameters!" + '\nPlease check D_7')
    if (D_6 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_6 cannot be zero')

  elif CoreType  ==  "ETD":
    if (D_2 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_2 must be greater then D_3')
    if (D_4 <= D_5):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_4 must be greater then D_5')
    if (D_6 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_6 cannot be zero')

  elif CoreType  ==  "P":
    if (D_2 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_2 must be greater then D_3')
    if (D_4 <= D_5):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_4 must be greater then D_5')
    if (D_3 <= D_6):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_3 must be greater then D_6')
    if (D_8 >= D_2):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_8 must be less then D_2')
    if (D_8 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_8 must be greater then D_3')
    if (D_7 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_7 cannot be zero')
    if (D_8 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_8 cannot be zero')

  elif CoreType  ==  "PH":
    if (D_2 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_2 must be greater then D_3')
    if (D_4 <= D_5):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_4 must be greater then D_5')
    if (D_3 <= D_6):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_3 must be greater then D_6')
    if (D_8 >= D_2):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_8 must be less then D_2')
    if (D_8 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_8 must be greater then D_3')
    if (D_7 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_7 cannot be zero')
    if (D_8 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_8 cannot be zero')

  elif CoreType  ==  "PQ":
    if (D_2 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_2 must be greater then D_3')
    if (D_4 <= D_5):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_4 must be greater then D_5')
    if (D_3 <= D_7):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_3 must be greater then D_7')
    if (D_6 >= D_1):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_6 must be less then D_1')
    if (D_6 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_6 must be greater then D_3')
    if (D_6 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_6 cannot be zero')
    if (D_7 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_7 cannot be zero')
    if (D_8 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_8 cannot be zero')

  elif CoreType  ==  "PT":
    if (D_2 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_2 must be greater then D_3')
    if (D_4 <= D_5):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_4 must be greater then D_5')
    if (D_3 <= D_6):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_3 must be greater then D_6')
    if (D_8 >= D_2):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_8 must be less then D_2')
    if (D_8 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_8 must be greater then D_3')
    if (D_6 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_6 cannot be zero')
    if (D_8 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_8 cannot be zero')

  elif CoreType  ==  "RM":
    if (D_2 <= D_3):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_2 must be greater then D_3')
    if (D_4 <= D_5):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_4 must be greater then D_5')
    if (D_3 <= D_6):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_3 must be greater then D_6')
    if (D_7 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_7 cannot be zero')
    if (D_8 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_8 cannot be zero')

  elif CoreType  ==  "U":
    if (D_3 <= D_4):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_3 must be greater then D_4')

  elif CoreType  ==  "UI":
    if (D_3 <= D_4):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_3 must be greater then D_4')
    if (D_6 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_6 cannot be zero')
    if (D_7 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_7 cannot be zero')
    if (D_8 == 0):
      raise UserErrorMessageException("Wrong core parameters!"+ '\nD_8 cannot be zero')
