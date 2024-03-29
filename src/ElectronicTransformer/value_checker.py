import math


def segmentation_angle_check(step, prop):
    if not (0 <= float(prop.Value) < 20):
        prop.Value = 0
        return add_error_message("{} should be 0 for True Surface or less than 20".format(prop.Caption))
    return True


def validate_sides_number(step, prop):
    return transformer.check_sides()


def GEThanZero(step, prop):
    if float(prop.Value) < 0:
        prop.Value = 0
        return add_error_message("{} cannot be negative".format(prop.Caption))
    return True


def GThanZero(step, prop):
    if float(prop.Value) <= 0:
        prop.Value = 1
        return add_error_message("{} should be greater than 0".format(prop.Caption))
    return True


def GEThanOne(step, prop):
    if float(prop.Value) < 1:
        prop.Value = 1
        return add_error_message("{} should be greater or equal than 1".format(prop.Caption))
    return True


def validityCheckTable(step, prop):
    xml_path = "winding_properties/conductor_type"
    if step.Properties[xml_path].Value == "Rectangular":
        xml_path += "/" + "table_layers"
        table = step.Properties[xml_path]
    else:
        xml_path += "/" + "table_layers_circles"
        table = step.Properties[xml_path]

    for i in range(table.RowCount):
        if prop.Name != "segments_number":
            if float(table.Value[xml_path + "/" + prop.Name][i]) <= 0:
                table.Value[xml_path + "/" + prop.Name][i] = 1
                return add_error_message("{} must be greater than zero!".format(prop.Caption))
        else:
            if (
                float(table.Value[xml_path + "/" + prop.Name][i]) < 8
                and float(table.Value[xml_path + "/" + prop.Name][i]) != 0
            ):
                table.Value[xml_path + "/" + prop.Name][i] = 8
                return add_error_message("{} must be Zero for True Surface or greater than 7!".format(prop.Caption))


def validate_resistance(step, prop):
    xml_path = "define_setup/table_resistance"
    table = step.Properties[xml_path]

    for i in range(table.RowCount):
        if float(table.Value[xml_path + "/" + prop.Name][i]) <= 0:
            table.Value[xml_path + "/" + prop.Name][i] = 1e-6
            return add_error_message("{} must be greater than zero!".format(prop.Caption))


def validate_aedt_version():
    """Raise error message for unsupported versions"""
    if float(oDesktop.GetVersion()[:-2]) < 2023.1:
        error = "Electronics Desktop Version is unsupported. Please use version 2023R1 or higher"
        add_error_message(error)
        raise UserErrorMessageException(error)


def check_core_dimensions(transformer_definition):
    validate_aedt_version()

    D_1 = float(transformer_definition["core_dimensions"]["D_1"])
    D_2 = float(transformer_definition["core_dimensions"]["D_2"])
    D_3 = float(transformer_definition["core_dimensions"]["D_3"])
    D_4 = float(transformer_definition["core_dimensions"]["D_4"])
    D_5 = float(transformer_definition["core_dimensions"]["D_5"])
    D_6 = float(transformer_definition["core_dimensions"]["D_6"])
    D_7 = float(transformer_definition["core_dimensions"]["D_7"])
    D_8 = float(transformer_definition["core_dimensions"]["D_8"])

    if D_1 == 0:
        raise UserErrorMessageException("Wrong core parameters!" + "\nD_1 cannot be zero")
    if D_2 == 0:
        raise UserErrorMessageException("Wrong core parameters!" + "\nD_2 cannot be zero")
    if D_3 == 0:
        raise UserErrorMessageException("Wrong core parameters!" + "\nD_3 cannot be zero")
    if D_4 == 0:
        raise UserErrorMessageException("Wrong core parameters!" + "\nD_4 cannot be zero")
    if D_5 == 0:
        raise UserErrorMessageException("Wrong core parameters!" + "\nD_5 cannot be zero")
    if D_1 <= D_2:
        raise UserErrorMessageException("Wrong core parameters!" + "\nD_1 must be greater than D_2")

    core_type = transformer_definition["core_dimensions"]["core_type"]
    if core_type == "E":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_2 must be greater than D_3")
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_4 must be greater than D_5")
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_6 cannot be zero")

    elif core_type == "EC":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_2 must be greater than D_3")
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_4 must be greater than D_5")
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_6 cannot be zero")

    elif core_type == "EFD":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_2 must be greater than D_3")
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_4 must be greater than D_5")
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_6 cannot be zero")
        if D_7 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_7 cannot be zero")

    elif core_type == "EI":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_2 must be greater than D_3")
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_4 must be greater than D_5")
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_6 cannot be zero")

    elif core_type == "EP":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_2 must be greater than D_3")
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_4 must be greater than D_5")
        if D_6 <= D_7:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_6 must be greater than D_7")
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_6 cannot be zero")
        if D_7 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_7 cannot be zero")

    elif core_type == "EQ":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_2 must be greater than D_3")
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_4 must be greater than D_5")
        if D_6 < D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_6 must be greater than D_3")
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_6 cannot be zero")

    elif core_type == "ER":
        if D_2 <= D_7:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_2 must be greater than D_7")
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_4 must be greater than D_5")
        if D_7 < 2 * math.sqrt((D_2 / 2) ** 2 - (D_6 / 2) ** 2) and D_7 != 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nPlease check D_7")
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_6 cannot be zero")

    elif core_type == "ETD":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_2 must be greater than D_3")
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_4 must be greater than D_5")
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_6 cannot be zero")

    elif core_type == "P":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_2 must be greater than D_3")
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_4 must be greater than D_5")
        if D_3 <= D_6:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_3 must be greater than D_6")
        if D_8 >= D_2:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_8 must be less than D_2")
        if D_8 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_8 must be greater than D_3")
        if D_7 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_7 cannot be zero")
        if D_8 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_8 cannot be zero")

    elif core_type == "PH":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_2 must be greater than D_3")
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_4 must be greater than D_5")
        if D_3 <= D_6:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_3 must be greater than D_6")
        if D_8 >= D_2:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_8 must be less than D_2")
        if D_8 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_8 must be greater than D_3")
        if D_7 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_7 cannot be zero")
        if D_8 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_8 cannot be zero")

    elif core_type == "PQ":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_2 must be greater than D_3")
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_4 must be greater than D_5")
        if D_3 <= D_7:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_3 must be greater than D_7")
        if D_6 >= D_1:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_6 must be less than D_1")
        if D_6 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_6 must be greater than D_3")
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_6 cannot be zero")
        if D_7 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_7 cannot be zero")
        if D_8 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_8 cannot be zero")

    elif core_type == "PT":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_2 must be greater than D_3")
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_4 must be greater than D_5")
        if D_3 <= D_6:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_3 must be greater than D_6")
        if D_8 >= D_2:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_8 must be less than D_2")
        if D_8 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_8 must be greater than D_3")
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_6 cannot be zero")
        if D_8 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_8 cannot be zero")

    elif core_type == "RM":
        if D_2 <= D_3:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_2 must be greater than D_3")
        if D_4 <= D_5:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_4 must be greater than D_5")
        if D_3 <= D_6:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_3 must be greater than D_6")
        if D_7 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_7 cannot be zero")
        if D_8 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_8 cannot be zero")

    elif core_type == "U":
        if D_3 <= D_4:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_3 must be greater than D_4")

    elif core_type == "UI":
        if D_3 <= D_4:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_3 must be greater than D_4")
        if D_6 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_6 cannot be zero")
        if D_7 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_7 cannot be zero")
        if D_8 == 0:
            raise UserErrorMessageException("Wrong core parameters!" + "\nD_8 cannot be zero")
