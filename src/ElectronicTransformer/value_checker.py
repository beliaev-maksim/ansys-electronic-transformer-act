#
# copyright 2021, ANSYS Inc. Software is released under GNU license
#
def segmentation_angle_check(step, prop):
    if not (0 <= float(prop.Value) < 20):
        prop.Value = 0
        return add_error_message(str(prop.Caption) + ' should be 0 for True Surface or less than 20')
    return True


def validate_sides_number(step, prop):
    return transformer.check_sides()


def GEThanZero(step, prop):
    if float(prop.Value) < 0:
        prop.Value = 0
        return add_error_message(str(prop.Caption) + ' cannot be negative')
    return True


def GThanZero(step, prop):
    if float(prop.Value) <= 0:
        prop.Value = 1
        return add_error_message(str(prop.Caption) + ' should be greater than 0')
    return True


def GEThanOne(step, prop):
    if float(prop.Value) < 1:
        prop.Value = 1
        return add_error_message(str(prop.Caption) + ' should be greater or equal than 1')
    return True


def validityCheckTable(step, prop):
    xml_path = "winding_properties/conductor_type"
    if step.Properties[xml_path].Value == 'Rectangular':
        xml_path += "/" + 'table_layers'
        table = step.Properties[xml_path]
    else:
        xml_path += "/" + 'table_layers_circles'
        table = step.Properties[xml_path]

    for i in range(table.RowCount):
        if prop.Name != 'segments_number':
            if float(table.Value[xml_path + '/' + prop.Name][i]) <= 0:
                table.Value[xml_path + '/' + prop.Name][i] = 1
                return add_error_message(str(prop.Caption) + ' must be greater than zero!')
        else:
            if (float(table.Value[xml_path + '/' + prop.Name][i]) < 8 and
                    float(table.Value[xml_path + '/' + prop.Name][i]) != 0):
                table.Value[xml_path + '/' + prop.Name][i] = 8
                return add_error_message(str(prop.Caption) + ' must be Zero for True Surface or greater than 7!')


def validate_resistance(step, prop):
    xml_path = "define_setup/table_resistance"
    table = step.Properties[xml_path]

    for i in range(table.RowCount):
        if float(table.Value[xml_path + '/' + prop.Name][i]) <= 0:
            table.Value[xml_path + '/' + prop.Name][i] = 1e-6
            return add_error_message(str(prop.Caption) + ' must be greater than zero!')

def validate_aedt_version():
    """Raise error message for unsupported versions"""
    if float(oDesktop.GetVersion()[:-2]) < 2021.1:
        error = "Electronics Desktop Version is unsupported. Please use version 2021R1 or higher"
        add_error_message(error)
        raise UserErrorMessageException(error)


def check_core_dimensions(step):
    validate_aedt_version()

    D_1 = float(step.Properties["core_properties/core_type/D_1"].Value)
    D_2 = float(step.Properties["core_properties/core_type/D_2"].Value)
    D_3 = float(step.Properties["core_properties/core_type/D_3"].Value)
    D_4 = float(step.Properties["core_properties/core_type/D_4"].Value)
    D_5 = float(step.Properties["core_properties/core_type/D_5"].Value)
    D_6 = float(step.Properties["core_properties/core_type/D_6"].Value)
    D_7 = float(step.Properties["core_properties/core_type/D_7"].Value)
    D_8 = float(step.Properties["core_properties/core_type/D_8"].Value)

    if D_1 == 0:
        raise UserErrorMessageException("Wrong core parameters!" + '\nD_1 cannot be zero')
    if D_2 == 0:
        raise UserErrorMessageException("Wrong core parameters!" + '\nD_2 cannot be zero')
    if D_3 == 0:
        raise UserErrorMessageException("Wrong core parameters!" + '\nD_3 cannot be zero')
    if D_4 == 0:
        raise UserErrorMessageException("Wrong core parameters!" + '\nD_4 cannot be zero')
    if D_5 == 0:
        raise UserErrorMessageException("Wrong core parameters!" + '\nD_5 cannot be zero')
    if D_1 <= D_2:
        raise UserErrorMessageException("Wrong core parameters!" + '\nD_1 must be greater than D_2')

    core_type = step.Properties["core_properties/core_type"].Value
    if core_type == "E":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_2 must be greater than D_3')
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_4 must be greater than D_5')
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_6 cannot be zero')

    elif core_type == "EC":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_2 must be greater than D_3')
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_4 must be greater than D_5')
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_6 cannot be zero')

    elif core_type == "EFD":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_2 must be greater than D_3')
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_4 must be greater than D_5')
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_6 cannot be zero')
        if D_7 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_7 cannot be zero')

    elif core_type == "EI":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_2 must be greater than D_3')
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_4 must be greater than D_5')
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_6 cannot be zero')

    elif core_type == "EP":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_2 must be greater than D_3')
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_4 must be greater than D_5')
        if D_6 <= D_7:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_6 must be greater than D_7')
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_6 cannot be zero')
        if D_7 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_7 cannot be zero')

    elif core_type == "EQ":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_2 must be greater than D_3')
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_4 must be greater than D_5')
        if D_6 < D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_6 must be greater than D_3')
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_6 cannot be zero')

    elif core_type == "ER":
        if D_2 <= D_7:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_2 must be greater than D_7')
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_4 must be greater than D_5')
        if D_7 < 2 * math.sqrt((D_2 / 2) ** 2 - (D_6 / 2) ** 2) and D_7 != 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nPlease check D_7')
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_6 cannot be zero')

    elif core_type == "ETD":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_2 must be greater than D_3')
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_4 must be greater than D_5')
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_6 cannot be zero')

    elif core_type == "P":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_2 must be greater than D_3')
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_4 must be greater than D_5')
        if D_3 <= D_6:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_3 must be greater than D_6')
        if D_8 >= D_2:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_8 must be less than D_2')
        if D_8 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_8 must be greater than D_3')
        if D_7 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_7 cannot be zero')
        if D_8 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_8 cannot be zero')

    elif core_type == "PH":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_2 must be greater than D_3')
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_4 must be greater than D_5')
        if D_3 <= D_6:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_3 must be greater than D_6')
        if D_8 >= D_2:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_8 must be less than D_2')
        if D_8 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_8 must be greater than D_3')
        if D_7 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_7 cannot be zero')
        if D_8 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_8 cannot be zero')

    elif core_type == "PQ":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_2 must be greater than D_3')
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_4 must be greater than D_5')
        if D_3 <= D_7:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_3 must be greater than D_7')
        if D_6 >= D_1:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_6 must be less than D_1')
        if D_6 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_6 must be greater than D_3')
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_6 cannot be zero')
        if D_7 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_7 cannot be zero')
        if D_8 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_8 cannot be zero')

    elif core_type == "PT":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_2 must be greater than D_3')
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_4 must be greater than D_5')
        if D_3 <= D_6:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_3 must be greater than D_6')
        if D_8 >= D_2:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_8 must be less than D_2')
        if D_8 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_8 must be greater than D_3')
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_6 cannot be zero')
        if D_8 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_8 cannot be zero')

    elif core_type == "RM":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_2 must be greater than D_3')
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_4 must be greater than D_5')
        if D_3 <= D_6:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_3 must be greater than D_6')
        if D_7 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_7 cannot be zero')
        if D_8 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_8 cannot be zero')

    elif core_type == "U":
        if D_3 <= D_4:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_3 must be greater than D_4')

    elif core_type == "UI":
        if D_3 <= D_4:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_3 must be greater than D_4')
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_6 cannot be zero')
        if D_7 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_7 cannot be zero')
        if D_8 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + '\nD_8 cannot be zero')
