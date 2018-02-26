def segAngCheck(step, prop):
    if not (0< float(prop.Value) <20):
        prop.Value = 15
        MsgBox(str(prop.Caption) + ' should be in range 0< Angle <20',vbOKOnly ,"Error")
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
        MsgBox(str(prop.Caption) + ' should be greater then 0',vbOKOnly ,"Error")
        return False
    return True
    
def GEThenOne(step, prop):
    if float(prop.Value) <1:
        prop.Value = 1
        MsgBox(str(prop.Caption) + ' should be greater or eqaul then 1',vbOKOnly ,"Error")
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
    
def CheckCoreDim(step):
    D_1 =step.Properties["coreProperties/coreType/D_1"].Value
    D_2 =step.Properties["coreProperties/coreType/D_2"].Value
    D_3 =step.Properties["coreProperties/coreType/D_3"].Value
    D_4 =step.Properties["coreProperties/coreType/D_4"].Value
    D_5 =step.Properties["coreProperties/coreType/D_5"].Value
    D_6 =step.Properties["coreProperties/coreType/D_6"].Value
    D_7 =step.Properties["coreProperties/coreType/D_7"].Value
    D_8 =step.Properties["coreProperties/coreType/D_8"].Value
    
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