# -*- coding: utf-8 -*-

#            Replacement of old ETK (Electronic Transformer Kit) for Maxwell
#
#            ACT Written by : Maksim Beliaev (maksim.beliaev@ansys.com)
#            Tested by: Mark Christini (mark.christini@ansys.com)
#            Last updated : 05.06.2018

import datetime    # get the time library
import os    # import operations system module
import re
from webbrowser import open as webopen
import math
from collections import OrderedDict
import json
import distutils.util


def atoi(letter):
    return int(letter) if letter.isdigit() else letter


def natural_keys(text):
    return [atoi(c) for c in re.split('(\d+)', text)]


class Step1:
    def __init__(self, step):
        self.step1 = step.Wizard.Steps["step1"]
        self.transformer_definition = OrderedDict()

        # create all objects for UI
        self.segmentation_angle = self.step1.Properties["core_properties/segmentation_angle"]
        self.supplier = self.step1.Properties["core_properties/supplier"]
        self.core_type = self.step1.Properties["core_properties/core_type"]
        self.core_model = self.step1.Properties["core_properties/core_type/core_model"]

        self.core_dimensions = {
            "D_1": self.step1.Properties["core_properties/core_type/D_1"],
            "D_2": self.step1.Properties["core_properties/core_type/D_2"],
            "D_3": self.step1.Properties["core_properties/core_type/D_3"],
            "D_4": self.step1.Properties["core_properties/core_type/D_4"],
            "D_5": self.step1.Properties["core_properties/core_type/D_5"],
            "D_6": self.step1.Properties["core_properties/core_type/D_6"],
            "D_7": self.step1.Properties["core_properties/core_type/D_7"],
            "D_8": self.step1.Properties["core_properties/core_type/D_8"]
        }

        self.define_airgap = self.step1.Properties["core_properties/define_airgap"]
        self.airgap_value = self.step1.Properties["core_properties/define_airgap/airgap_value"]
        self.airgap_on_leg = self.step1.Properties["core_properties/define_airgap/airgap_on_leg"]

    # read data from file
    def read_data(self, sender, args):
        path = ExtAPI.UserInterface.UIRenderer.ShowFileOpenDialog('Text Files(*.txt;*.json;)|*.txt;*.json;')

        if path is None:
            return

        with open(path, "r") as input_f:
            try:
                self.transformer_definition = json.load(input_f, object_pairs_hook=OrderedDict)
            except ValueError as e:
                match_line = re.match("Expecting object: line (.*) column", str(e))
                if match_line:
                    self.add_error_message(
                        'Please verify that all data in file is covered by double quotes "" ' +
                        '(integers and floats can be both covered or uncovered)')
                else:
                    match_line = re.match("Expecting property name: line (.*) column", str(e))
                    if match_line:
                        self.add_error_message("Please verify that there is no empty argument in the file. "
                                               "Cannot be two commas in a row without argument in between")
                    else:
                        # so smth unexpected
                        return self.add_error_message(e)

                return self.add_error_message("Please correct following line: {} in file: {}".format(
                                                                                            match_line.group(1), path))
        try:
            self.populate_ui_data()
        except ValueError:
            return self.add_error_message("Please verify that integer numbers in input file have proper format "
                                          "eg 1 and not 1.0")
        except KeyError as e:
            return self.add_error_message("Please specify parameter:{} in input file".format(e))

    def refresh_step1(self):
        """create buttons and HTML data for first step"""
        self.desktop = oDesktop

        self.setup_button(self.step1, "readData", "Read Settings File", ButtonPositionType.Left, self.read_data)
        self.setup_button(self.step1, "helpButton", "Help", ButtonPositionType.Center, self.help_button_clicked,
                          style="blue")

        self.show_core_img()

    def callback_step1(self):
        check_core_dimensions(self.step1)

    def insert_default_values(self):
        """invoke when Core Model is changed: set core names"""
        # set core values for selected type
        for j in range(1, 9):
            try:
                self.core_dimensions["D_" + str(j)].Value = float(self.core_models[self.core_model.Value][j - 1])
                self.core_dimensions["D_" + str(j)].Visible = True
            except ValueError:
                self.core_dimensions["D_" + str(j)].Visible = False

        self.update_ui(self.step1)

    def show_core_img(self):
        """invoked to change image and core dimensions when supplier or core type changed"""
        self.core_model.Options.Clear()

        if self.core_type.Value not in ['EP', 'ER', 'PQ', 'RM']:
            HTML_data = '<img width="300" height="200" src="' + str(ExtAPI.Extension.InstallDir) + '/images/'
        else:
            HTML_data = '<img width="275" height="360" src="' + str(ExtAPI.Extension.InstallDir) + '/images/'

        report = self.step1.UserInterface.GetComponent("coreImage")
        report.SetHtmlContent(HTML_data + self.core_type.Value + 'Core.png"/>')  # set core names

        self.core_models = cores_database[self.supplier.Value][self.core_type.Value]
        for model in sorted(self.core_models.keys(), key=natural_keys):
            self.core_model.Options.Add(model)
        self.core_model.Value = self.core_model.Options[0]

        report.Refresh()
        self.insert_default_values()


class Step2:
    def __init__(self, step):
        self.step2 = step.Wizard.Steps["step2"]

        self.draw_winding = self.step2.Properties["winding_properties/draw_winding"]
        self.layer_type = self.step2.Properties["winding_properties/draw_winding/layer_type"]
        self.number_of_layers = self.step2.Properties["winding_properties/draw_winding/number_of_layers"]
        self.layer_spacing = self.step2.Properties["winding_properties/draw_winding/layer_spacing"]
        self.bobbin_board_thickness = self.step2.Properties["winding_properties/draw_winding/bobbin_board_thickness"]
        self.top_margin = self.step2.Properties["winding_properties/draw_winding/top_margin"]
        self.side_margin = self.step2.Properties["winding_properties/draw_winding/side_margin"]
        self.include_bobbin = self.step2.Properties["winding_properties/draw_winding/include_bobbin"]
        self.conductor_type = self.step2.Properties["winding_properties/draw_winding/conductor_type"]

        self.table_layers = self.conductor_type.Properties["table_layers"]
        self.table_layers_circles = self.conductor_type.Properties["table_layers_circles"]

        self.skip_check = self.step2.Properties["winding_properties/draw_winding/skip_check"]

    def init_data_step2(self):
        # initialize tables
        self.table_layers.AddRow()
        self.table_layers.Properties["conductor_width"].Value = 0.2
        self.table_layers.Properties["conductor_height"].Value = 0.2
        self.table_layers.Properties["turns_number"].Value = 2
        self.table_layers.Properties["insulation_thickness"].Value = 0.05
        self.table_layers.Properties["layer"].Value = 'Layer_1'
        self.table_layers.SaveActiveRow()

        self.table_layers_circles.AddRow()
        self.table_layers_circles.Properties["conductor_diameter"].Value = 0.2
        self.table_layers_circles.Properties["segments_number"].Value = 8
        self.table_layers_circles.Properties["turns_number"].Value = 2
        self.table_layers_circles.Properties["insulation_thickness"].Value = 0.05
        self.table_layers_circles.Properties["layer"].Value = 'Layer_1'
        self.table_layers_circles.SaveActiveRow()

    def refresh_step2(self):
        self.setup_button(self.step2, "helpButton", "Help", ButtonPositionType.Right, self.help_button_clicked,
                          style="blue")

    def callback_step2(self):
        # invoke validation from valueChecker file
        self.check_winding()
        self.check_board_bobbin()

        # read custom material file and dataset points for it
        lib_path = oDesktop.GetPersonalLibDirectory() + '\Materials\\'
        if os.path.isfile(lib_path + 'matdata.tab'):
            with open(lib_path + 'matdata.tab', 'r') as Inmat:
                next(Inmat)
                for line in Inmat:
                    line = line.split()
                    matDict[line[0]] = line[1:]

                    if not os.path.isfile(lib_path + line[0] + '.tab'):
                        raise UserErrorMessageException(
                            'File {} in directory {} does not exist!'.format(line[0] + '.tab', lib_path))

                    with open(lib_path + line[0] + '.tab', 'r') as datasheet:
                        buffer_list = []
                        for lineData in datasheet:
                            if not ("""X" \t"Y""" in lineData):
                                line_array = lineData.rstrip().split()
                                try:
                                    buffer_list.append([float(str(a).replace(',', '.')) for a in line_array])
                                except:
                                    raise UserErrorMessageException('Wrong values in file {}'.format(line[0] + '.tab'))
                        matKeyPoints[line[0]] = buffer_list
        else:
            if not os.path.exists(lib_path):
                os.makedirs(lib_path)
            with open(lib_path + 'matdata.tab', 'w') as file:
                file.write('Material Name\tConductivity\tCm\tx\ty\tdensity\n')

        get_time = datetime.datetime.now().strftime('%d_%m_%Y_%H_%M_%S')
        self.design_name = 'Transformer_' + get_time
        self.oProject = oDesktop.GetActiveProject()
        if self.oProject is None:
            self.oProject = oDesktop.NewProject()
        self.oProject.InsertDesign("Maxwell 3D", self.design_name, "EddyCurrent", "")
        self.oDesign = self.oProject.SetActiveDesign(self.design_name)
        self.oEditor = self.oDesign.SetActiveEditor("3D Modeler")

        args = [self.step1, self.oProject, self.oDesign, self.oEditor]
        all_cores = {'E': ECore(args), 'EI': EICore(args), 'U': UCore(args), 'UI': UICore(args),
                     'PQ': PQCore(args), 'ETD': ETDCore(args), 'EQ': ETDCore(args),
                     'EC': ETDCore(args), 'RM': RMCore(args), 'EP': EPCore(args),
                     'EFD': EFDCore(args), 'ER': ETDCore(args), 'P': PCore(args),
                     'PT': PCore(args), 'PH': PCore(args)}

        cores_arguments = {'EC': 'EC', 'ETD': 'ETD', 'EQ': 'EQ', 'ER': 'ER', 'P': 'P', 'PT': 'PT', 'PH': 'PH'}

        flag_auto_save = oDesktop.GetAutoSaveEnabled()
        oDesktop.EnableAutoSave(False)

        self.draw_class = all_cores[self.core_type.Value]
        if self.core_type.Value in cores_arguments:
            self.draw_class.draw_geometry(cores_arguments[self.core_type.Value])
        else:
            self.draw_class.draw_geometry()

        oDesktop.EnableAutoSave(flag_auto_save)

    def check_winding(self):
        if self.skip_check.Value:
            return

        side_margin = self.side_margin.Value
        bobbin_board_thickness = self.bobbin_board_thickness.Value
        top_margin = self.top_margin.Value
        layer_spacing = self.layer_spacing.Value

        # ---- start checking for wound ---- #
        if self.layer_type.Value == 'Wound':
            # ---- Check possible width for wound---- #
            if self.conductor_type.Value == 'Rectangular':
                xml_path_to_table = 'winding_properties/draw_winding/conductor_type/table_layers'
                table = self.step2.Properties[xml_path_to_table]
                # take sum of layer dimensions where one layer is: Width + 2 * Insulation
                maximum_layer = sum(
                    [
                        table.Value[xml_path_to_table + '/conductor_width'][i] +
                        # do not forget that insulation is on both sides
                        2 * table.Value[xml_path_to_table + '/insulation_thickness'][i] +
                        layer_spacing  # since number of layers - 1 for spacing)
                        for i in range(len(table.Value[xml_path_to_table + '/layer']))
                    ]
                ) - layer_spacing

            elif self.conductor_type.Value == 'Circular':
                xml_path_to_table = 'winding_properties/draw_winding/conductor_type/table_layers_circles'
                table = self.step2.Properties[xml_path_to_table]
                maximum_layer = sum(
                    [
                        (table.Value[xml_path_to_table + '/conductor_diameter'][i] +
                         2 * table.Value[xml_path_to_table + '/insulation_thickness'][i] + layer_spacing)
                        for i in range(len(table.Value[xml_path_to_table + '/layer']))
                    ]
                ) - layer_spacing

            # do not forget that windings are laying on both sides of the core
            maximum_possible_width = 2 * (bobbin_board_thickness + side_margin + maximum_layer)

            # ---- Check possible height for wound---- #
            if self.conductor_type.Value == 'Rectangular':
                xml_path_to_table = 'winding_properties/draw_winding/conductor_type/table_layers'
                table = self.step2.Properties[xml_path_to_table]
                # max value from each layer: (Height + 2 * Insulation) * number of layers
                maximum_layer = max(
                    [
                        ((table.Value[xml_path_to_table + '/conductor_height'][i] +
                          # do not forget that insulation is on both sides
                          2 * table.Value[xml_path_to_table + '/insulation_thickness'][i]) *
                         table.Value[xml_path_to_table + '/turns_number'][i])
                        for i in range(len(table.Value[xml_path_to_table + '/layer']))
                    ]
                )

            elif self.conductor_type.Value == 'Circular':
                xml_path_to_table = 'winding_properties/draw_winding/conductor_type/table_layers_circles'
                table = self.step2.Properties[xml_path_to_table]
                maximum_layer = max(
                    [
                        ((table.Value[xml_path_to_table + '/conductor_diameter'][i] +
                          2 * table.Value[xml_path_to_table + '/insulation_thickness'][i]) *
                         table.Value[xml_path_to_table + '/turns_number'][i])
                        for i in range(len(table.Value[xml_path_to_table + '/layer']))
                    ]
                )

            maximum_possible_height = (2 * bobbin_board_thickness + top_margin + maximum_layer)
        # ---- Wound type limit found ---- #

        elif self.layer_type.Value == 'Planar':
            xml_path_to_table = 'winding_properties/draw_winding/conductor_type/table_layers'
            table = self.step2.Properties[xml_path_to_table]
            # ---- Check width for planar---- #
            maximum_layer = max(
                [
                    ((table.Value[xml_path_to_table + '/conductor_width'][i] +
                      # in this case it is turn spacing (no need to x2)
                      table.Value[xml_path_to_table + '/insulation_thickness'][i]) *
                     table.Value[xml_path_to_table + '/turns_number'][i])
                    for i in range(len(table.Value[xml_path_to_table + '/layer']))
                ]
            )
            maximum_possible_width = (2 * maximum_layer + 2 * side_margin)

            # ---- Check Height for planar ---- #
            maximum_layer = sum(
                [
                    (table.Value[xml_path_to_table + '/conductor_height'][i] + bobbin_board_thickness + layer_spacing)
                    for i in range(len(table.Value[xml_path_to_table + '/layer']))
                ]
            ) - layer_spacing

            maximum_possible_height = maximum_layer + top_margin
            # ---- Planar type limit found ---- #

        # ---- Check accomodation not depending on layer type ---- #
        # ---- Height ---- #
        if self.core_type.Value in ["E", "EC", "EFD", "EQ", "ER", "ETD", "PH"]:
            # D_5 is height of one half core
            if maximum_possible_height > 2 * self.core_dimensions["D_5"].Value:
                raise UserErrorMessageException("Cannot accommodate all windings, increase D_5")
        elif self.core_type.Value in ["EI", "EP", "P", "PT", "PQ", "RM"]:
            # D_5 is height of core
            if maximum_possible_height > self.core_dimensions["D_5"].Value:
                raise UserErrorMessageException("Cannot accommodate all windings, increase D_5")
        elif self.core_type.Value == "UI":
            # D_4 is height of core
            if maximum_possible_height > self.core_dimensions["D_4"].Value:
                raise UserErrorMessageException("Cannot accommodate all windings, increase D_4")
        elif self.core_type.Value == "U":
            # D_4 is height of core
            if maximum_possible_height > 2 * self.core_dimensions["D_4"].Value:
                raise UserErrorMessageException("Cannot accommodate all windings, increase D_4")

        # ---- Width ---- #
        if self.core_type.Value not in ["U", "UI"]:
            # D_2 - D_3 is sum of dimesnsions of two slots for windngs (left + right)
            if maximum_possible_width > (self.core_dimensions["D_2"].Value - self.core_dimensions["D_3"].Value):
                raise UserErrorMessageException("Cannot accommodate all windings, increase D_2")
        else:
            # D_2 is dimension of one side slot for winding
            if maximum_possible_width / 2 > self.core_dimensions["D_2"].Value:
                raise UserErrorMessageException("Cannot accommodate all windings, increase D_2")

    def check_board_bobbin(self):
        if self.bobbin_board_thickness == 0 and self.include_bobbin:
            raise UserErrorMessageException("Include board/bobbin is checked but thickness is 0")

        if self.layer_type == 'Planar' and self.bobbin_board_thickness == 0 and self.layer_spacing == 0:
            raise UserErrorMessageException(
                "For planar transformer Board thickness and Layer spacing cannot be equal 0 at once")

    def change_captions(self, need_refresh=True):
        """change captions depending on Wound or Planar transformer"""
        if self.layer_type.Value == 'Planar':
            self.bobbin_board_thickness.Caption = 'Board thickness:'
            self.include_bobbin.Caption = 'Include board in geometry:'
            self.top_margin.Caption = 'Bottom Margin'

            self.table_layers.Properties["insulation_thickness"].Caption = 'Turn Spacing'

            self.conductor_type.Value = 'Rectangular'
            self.conductor_type.ReadOnly = True

        elif self.layer_type.Value == 'Wound':
            self.bobbin_board_thickness.Caption = 'Bobbin thickness:'
            self.include_bobbin.Caption = 'Include bobbin in geometry:'
            self.top_margin.Caption = 'Top Margin'

            self.table_layers.Properties["insulation_thickness"].Caption = 'Insulation thickness'

            self.conductor_type.ReadOnly = False

        # do not need it if invoke read input from file
        if need_refresh:
            self.update_ui(self.step2)

    def update_rows(self):
        """when user changes number of layers we append/delete rows in tables"""
        if self.number_of_layers.Value < 1:
            self.number_of_layers.Value = 1
            self.add_error_message("Number of layers should be greater or equal than 1")
            return False

        if self.conductor_type.Value == 'Rectangular':
            xml_path_to_table = "winding_properties/draw_winding/conductor_type/table_layers"
            list_of_prop = ["conductor_width", "conductor_height", "turns_number", "insulation_thickness"]
        else:
            xml_path_to_table = "winding_properties/draw_winding/conductor_type/table_layers_circles"
            list_of_prop = ["conductor_diameter", "segments_number", "turns_number", "insulation_thickness"]

        table = self.step2.Properties[xml_path_to_table]
        num_layers_new = self.number_of_layers.Value
        row_num = table.RowCount
        if num_layers_new < row_num:
            for i in range(0, row_num - num_layers_new + 1):
                table.DeleteRow(row_num - i)
        elif num_layers_new > row_num:
            for i in range(row_num + 1, num_layers_new + 1):
                table.AddRow()
                for prop in list_of_prop:
                    table.Properties[prop].Value = table.Value[xml_path_to_table + "/" + prop][row_num - 1]

                table.Properties["layer"].Value = 'Layer_' + str(i)
                table.SaveActiveRow()


class Step3:
    def __init__(self, step):
        self.step3 = step.Wizard.Steps["step3"]

        self.core_material = self.step3.Properties["define_setup/core_material"]
        self.coil_material = self.step3.Properties["define_setup/coil_material"]

        self.adaptive_frequency = self.step3.Properties["define_setup/adaptive_frequency"]
        self.percentage_error = self.step3.Properties["define_setup/percentage_error"]
        self.number_passes = self.step3.Properties["define_setup/number_passes"]

        self.transformer_sides = self.step3.Properties["define_setup/transformer_sides"]
        self.excitation_strategy = self.step3.Properties["define_setup/excitation_strategy"]
        self.voltage = self.step3.Properties["define_setup/voltage"]
        self.resistance = self.step3.Properties["define_setup/resistance"]

        self.offset = self.step3.Properties["define_setup/offset"]
        self.project_path = self.step3.Properties["define_setup/project_path"]

        self.frequency_sweep = self.step3.Properties["define_setup/frequency_sweep"]
        self.start_frequency = self.step3.Properties["define_setup/frequency_sweep/start_frequency"]
        self.start_frequency_unit = self.step3.Properties["define_setup/frequency_sweep/start_frequency_unit"]
        self.stop_frequency = self.step3.Properties["define_setup/frequency_sweep/stop_frequency"]
        self.stop_frequency_unit = self.step3.Properties["define_setup/frequency_sweep/stop_frequency_unit"]
        self.samples = self.step3.Properties["define_setup/frequency_sweep/samples"]

    def refresh_step3(self):
        self.setup_button(self.step3, "helpButton", "Help", ButtonPositionType.Right, self.help_button_clicked,
                          style="blue")

        self.setup_button(self.step3, "analyzeButton", "Analyze", ButtonPositionType.Center, self.analyze_click,
                          active=False)

        self.setup_button(self.step3, "setupAnalysisButton", "Setup Analysis", ButtonPositionType.Center,
                          self.setup_analysis_click, active=False)

        self.setup_button(self.step3, "defineWindingsButton", "Define Windings", ButtonPositionType.Center,
                          self.define_windings_click, active=False)

        self.analysis_set = False

        # insert path to the saved project as default path to textbox
        path = self.oProject.GetPath()
        if self.oProject is None:
            path = oDesktop.GetProjectPath()
        self.project_path.Value = str(path)

        # add materials from vocabulary
        self.core_material.Options.Clear()
        for key in sorted(matDict):
            self.core_material.Options.Add(key)

        try:
            if core_material is not None:
                self.core_material.Value = core_material
        except:
            self.core_material.Value = key

        if self.draw_winding.Value:
            self.step3.UserInterface.GetComponent("defineWindingsButton").SetEnabledFlag("defineWindingsButton", True)

        self.update_ui(self.step3)

        if self.number_of_layers.ReadOnly:
            self.windings_definition = WindingForm(defined_layers_list=self.defined_layers_list)
            if self.transformer_sides.Value == 1:
                self.resistance.Visible = False
        else:
            self.transformer_sides.Value = 1
            self.resistance.Visible = False
            self.windings_definition = WindingForm(number_undefined_layers=int(self.number_of_layers.Value))

        if self.windings_definition.defined_layers_list:
            self.step3.UserInterface.GetComponent("analyzeButton").SetEnabledFlag("analyzeButton", True)
            self.step3.UserInterface.GetComponent("setupAnalysisButton").SetEnabledFlag("setupAnalysisButton", True)

    def reset_step3(self):
        """when back button on step3 is clicked we return one step back and we need to delete all winding definitions"""
        self.transformer_definition.pop("winding_definition", None)

    def check_sides(self):
        """if number of sides was changed just clear the data in winding dialogue and disable buttons"""
        self.windings_definition.defined_layers_list = []
        self.windings_definition.defined_layers_listbox.Items.Clear()
        self.step3.UserInterface.GetComponent("analyzeButton").SetEnabledFlag("analyzeButton", False)
        self.step3.UserInterface.GetComponent("setupAnalysisButton").SetEnabledFlag("setupAnalysisButton", False)

        if self.transformer_sides.Value < 1:
            self.transformer_sides.Value = 1
            return self.add_warning_message("Number of transformer sides cannot be less than 1")

        elif self.transformer_sides.Value > self.number_of_layers.Value:
            self.transformer_sides.Value = self.number_of_layers.Value
            return self.add_warning_message("Number of transformer sides cannot be less than number of layers")

        if self.transformer_sides.Value == 1:
            self.resistance.Visible = False
        else:
            self.resistance.Visible = True

        self.update_ui(self.step3)

    def analyze_click(self, sender, args):
        if not self.analysis_set:
            self.setup_analysis_click()

        try:
            self.oDesign.Analyze("Setup1")
        except:
            pass

    def setup_analysis_click(self, sender="", args=""):
        coil_material = self.coil_material.Value + "_temperature"

        layers_list = []
        layer_sections_list = []
        layers_sections_delete_list = []
        core_list = []

        obj_list = self.oEditor.GetObjectsInGroup('Solids')
        for each_obj in obj_list:
            if "Layer" in each_obj:
                layers_list.append(each_obj)
                layer_sections_list.append(each_obj + "_Section1")
                layers_sections_delete_list.append(each_obj + "_Section1_Separate1")
            elif "Core" in each_obj:
                core_list.append(each_obj)

        create_new_materials(self.oProject, self.core_material.Value, coil_material)
        assign_material(self.oEditor, ','.join(core_list), '"Material_' + self.core_material.Value + '"')
        assign_material(self.oEditor, ','.join(layers_list), '"' + coil_material + '"')

        if coil_material == 'Copper_temperature':
            change_color(self.oEditor, layers_list, 255, 128, 64)
        else:
            change_color(self.oEditor, layers_list, 132, 135, 137)

        if self.include_bobbin.Value:
            if self.layer_type.Value == 'Wound':
                assign_material(self.oEditor, 'Bobbin', '"polyamide"')
            else:
                boards = ','.join(self.oEditor.GetMatchedObjectName("Board*"))
                assign_material(self.oEditor, boards, '"polyamide"')

        adapt_freq = create_setup(self.step3, self.oDesign)
        enable_thermal(self.oDesign, obj_list)

        create_terminal_sections(self.oEditor, layers_list, layer_sections_list, layers_sections_delete_list)
        self.oModule_BoundarySetup = self.oDesign.GetModule("BoundarySetup")

        self.assign_winding_excitations(layer_sections_list)
        self.assign_matrix_winding()
        calculate_leakage(self.oDesign, self.transformer_sides.Value)

        self.oModule_BoundarySetup.SetCoreLoss(core_list, False)

        # turn on eddy effects
        eddy_list = ["NAME:EddyEffectVector"]
        for layer in layers_list:
            eddy_list.append(
                [
                    "NAME:Data",
                    "Object Name:=", layer,
                    "Eddy Effect:=", True,
                    "Displacement Current:=", True
                ])
        self.oModule_BoundarySetup.SetEddyEffect(["NAME:Eddy Effect Setting", eddy_list])

        create_region(self.step3, self.oEditor)
        self.create_skin_layers(adapt_freq, coil_material, core_list)
        self.oEditor.FitAll()

        if self.project_path.Value is None:
            # that means that we are here due to auto run after reading a file
            return

        self.write_json_data()

        self.oProject.SaveAs(os.path.join(self.project_path.Value, self.design_name + '.aedt'), True)
        self.analysis_set = True
        self.step3.UserInterface.GetComponent("setupAnalysisButton").SetEnabledFlag("setupAnalysisButton", False)
        self.step3.UserInterface.GetComponent("defineWindingsButton").SetEnabledFlag("defineWindingsButton", False)

    def define_windings_click(self, sender, args):
        self.windings_definition.number_of_sides = self.transformer_sides.Value

        # if layers read from file and after that number of sides was changed
        if not self.windings_definition.defined_layers_list:
            self.windings_definition.number_undefined_layers = int(self.number_of_layers.Value)

        self.windings_definition.refresh_ui_on_show()
        self.windings_definition.ShowDialog()

        if self.windings_definition.defined_layers_list:
            self.step3.UserInterface.GetComponent("analyzeButton").SetEnabledFlag("analyzeButton", True)
            self.step3.UserInterface.GetComponent("setupAnalysisButton").SetEnabledFlag("setupAnalysisButton", True)

    def assign_winding_excitations(self, layer_sections_list):
        if self.excitation_strategy.Value == "Voltage":
            self.create_winding("Side_1", "Voltage", current=0, resistance=1e-06, inductance=0,
                                voltage=self.voltage.Value)
        else:
            self.create_winding("Side_1", "Current", current=self.voltage.Value, resistance=0, inductance=0, voltage=0)

        for i in range(2, self.transformer_sides.Value+1):
            self.create_winding("Side_" + str(i), "Voltage", current=0, resistance=self.resistance.Value, inductance=0,
                                voltage=0)

        for section in layer_sections_list:
            for definition in self.windings_definition.defined_layers_list:
                if section.split("_")[0] == definition.split("_")[-1]:
                    side = definition.split("_")[1]

                    point_terminal = True if int(side) > 1 else False
                    self.oModule_BoundarySetup.AssignCoilTerminal(
                        [
                            "NAME:" + section,
                            "Objects:=", [section],
                            "ParentBndID:=", "Side_" + side,
                            "Conductor number:=", "1",
                            "Winding:="	, "Side_" + side,
                            "Point out of terminal:=", point_terminal
                        ])

    def create_winding(self, name, winding_type, current=0.0, resistance=0.0, inductance=0.0, voltage=0.0):
        self.oModule_BoundarySetup.AssignWindingGroup(
            [
                "NAME:" + name,
                "Type:=", winding_type,
                "IsSolid:=", True,
                "Current:=", str(current) + "A",
                "Resistance:="	, str(resistance) + "ohm",
                "Inductance:="		, str(inductance) + "nH",
                "Voltage:="		, str(voltage) + "V",
                "ParallelBranchesNum:="	, "1",
                "Phase:="		, "0deg"
            ])

    def assign_matrix_winding(self):
        oModule = self.oDesign.GetModule("MaxwellParameterSetup")

        matrix_entry = ["NAME:MatrixEntry"]
        for i in range(1, self.transformer_sides.Value + 1):
            matrix_entry.append([
                 "NAME:MatrixEntry",
                 "Source:=", "Side_" + str(i),
                 "NumberOfTurns:=", "1"
             ])

        oModule.AssignMatrix(
            ["NAME:Matrix1",
             matrix_entry,
             ["NAME:MatrixGroup"]
             ])

    def create_skin_layers(self, adapt_freq, coil_material, core_list):
        """Function to create 'manual' skin layers from face sheets and to assign mesh operations to
        core geometry and windings in which we do not require skin layers (thinner than 3 skin depths)"""

        mesh_op_sz = max([float(Lx) for Lx in self.draw_class.core_dims]) / 20.0

        layer_names = self.oEditor.GetMatchedObjectName("Layer*")
        layer_names = [name for name in layer_names if "Section" not in name]

        sigma = 58000000 if coil_material == 'Copper_temperature' else 38000000
        skin_depth = 503.292121 * math.sqrt(1 / (sigma * adapt_freq)) * 1000  # convert to mm

        oModule = self.oDesign.GetModule("MeshSetup")
        # first assign mesh operation for cores
        assign_length_op(oModule, core_list, mesh_op_sz)

        for layer in layer_names:
            layer_number = int(layer.split("_")[0][5:])
            face_ids = [int(ID) for ID in self.oEditor.GetFaceIDs(layer)]

            my_dict = {}
            list_of_round_faces = []
            for ID in face_ids:
                try:
                    # DE190482, fails on round objects, need to handle separate
                    my_dict[ID] = self.oEditor.GetFaceCenter(ID)
                except:
                    list_of_round_faces.append(ID)
            # sort by Z coordinate to get top and bottom face
            sorted_dict = sorted(my_dict.items(), key=lambda x: float(x[1][2]))

            bot_face_id = sorted_dict[0][0]
            top_face_id = sorted_dict[-1][0]

            if self.layer_type.Value == 'Planar':
                height = self.draw_class.WdgParDict[layer_number][1]
                if height < 3 * skin_depth:
                    # that means we do not need skin here
                    assign_length_op(oModule, [layer], mesh_op_sz)
                    continue

                for index in range(1, 3):
                    skinDepthLayer = skin_depth / index  # to get two layers of skin depth

                    create_object_from_face(self.oEditor, layer, [bot_face_id])
                    self.draw_class.rename(layer + "_ObjectFromFace1", layer + "_bot" + str(index))
                    self.draw_class.move(layer + "_bot" + str(index), 0, 0, skinDepthLayer)
                    create_object_from_face(self.oEditor, layer, [top_face_id])
                    self.draw_class.rename(layer + "_ObjectFromFace1", layer + "_top" + str(index))
                    self.draw_class.move(layer + "_top" + str(index), 0, 0, -skinDepthLayer)

            else:
                # make validation that skindepth is required
                # grab from table width values for all layers
                width = self.draw_class.WdgParDict[layer_number][0]

                if width < 3 * skin_depth:
                    # that means we do not need skin here
                    assign_length_op(oModule, [layer], mesh_op_sz)
                    continue

                for ID in list_of_round_faces:
                    my_dict[ID] = [0, 0, 0]  # just to fill with IDs which have failed due to round shape

                # pick all faces except top and bottom face
                if self.conductor_type.Value == 'Rectangular':
                    wound_faces = [int(key) for key in my_dict.keys() if int(key) not in [bot_face_id, top_face_id]]
                else:
                    wound_faces = [int(key) for key in my_dict.keys()]

                self.draw_class.rename(layer, layer + "_forskin")
                self.oEditor.Copy(
                    [
                        "NAME:Selections",
                        "Selections:=", layer + "_forskin"
                    ])
                self.oEditor.Paste()
                self.draw_class.rename(layer + "_forskin1", layer)

                for index in range(1, 3):
                    self.oEditor.MoveFaces(
                        [
                            "NAME:Selections",
                            "Selections:=", layer + "_forskin",
                            "NewPartsModelFlag:=", "Model"
                        ],
                        [
                            "NAME:Parameters",
                            [
                                "NAME:MoveFacesParameters",
                                "MoveAlongNormalFlag:=", True,
                                "OffsetDistance:=", str(-skin_depth / 2) + "mm",
                                "MoveVectorX:="	, "0mm",
                                "MoveVectorY:="		, "0mm",
                                "MoveVectorZ:="		, "0mm",
                                "FacesToMove:="		, wound_faces
                            ]
                        ])

                    create_object_from_face(self.oEditor, layer + "_forskin", wound_faces)

                # delete object only after loop since create object from faces twice
                self.oEditor.Delete(
                    [
                        "NAME:Selections",
                        "Selections:="	, layer + "_forskin"
                    ])

    def excitation_strategy_change(self):
        if self.excitation_strategy.Value == "Voltage":
            self.voltage.Caption = "Voltage [V]"
        else:
            self.voltage.Caption = "Current [A]"


class TransformerClass(Step1, Step2, Step3):
    """Main class which serves to manipulate Step classes. We initialize these classes from current class.
    Invoke all functions from each step here.
    Class TransformerClass will inherit all functions from classes Step1, Step2, Step3 including all callback actions"""
    def __init__(self, step):
        self.step1 = step

    def setup_button(self, step, comp_name, caption, position, callback, active=True, style=None):
        # Setup a button
        update_btn_session = step.UserInterface.GetComponent(comp_name)  # get component name from XML
        update_btn_session.AddButton(comp_name, caption, position, active)  # add button caption and position
        update_btn_session.ButtonClicked += callback  # connect to callback function
        if style == "blue":
            # change CSS properties of the button to change it's color
            update_btn_session.AddCSSProperty("background-color", "#3383ff", "button")
            update_btn_session.AddCSSProperty("color", "white", "button")

    def add_error_message(self, msg):
        self.desktop.AddMessage("", "", 2, "ACT:" + str(msg))

    def add_warning_message(self, msg):
        self.desktop.AddMessage("", "", 1, "ACT:" + str(msg))

    def add_info_message(self, msg):
        self.desktop.AddMessage("", "", 0, "ACT:" + str(msg))

    def update_ui(self, step):
        """Refresh UI data """
        step.UserInterface.GetComponent("Properties").UpdateData()
        step.UserInterface.GetComponent("Properties").Refresh()

    def help_button_clicked(self, sender, args):
        """when user clicks Help button HTML page will be opened in standard web browser"""
        webopen(str(ExtAPI.Extension.InstallDir) + '/help/help.html')

    def initialize_step1(self):
        """separate init is required because we do not have access to steps without XML callback"""
        Step1.__init__(self, self.step1)

    def initialize_step2(self):
        """separate init is required because we do not have access to steps without XML callback"""
        Step2.__init__(self, self.step1)
        Step2.init_data_step2(self)

    def initialize_step3(self):
        """separate init is required because we do not have access to steps without XML callback"""
        Step3.__init__(self, self.step1)

    def collect_ui_data(self):
        """collect data from all steps, write this data to dictionary"""
        # step1
        self.transformer_definition["core_dimensions"] = OrderedDict([
            ("segmentation_angle", self.segmentation_angle.Value),
            ("supplier", self.supplier.Value),
            ("core_type", self.core_type.Value),
            ("core_model", self.core_model.Value)
        ])

        for i in range(1, 9):
            if self.core_dimensions["D_" + str(i)].Visible:
                d_value = str(self.core_dimensions["D_" + str(i)].Value)
                self.transformer_definition["core_dimensions"]["D_" + str(i)] = d_value

        self.transformer_definition["core_dimensions"]["airgap"] = OrderedDict([
            ("define_airgap", str(bool(self.define_airgap.Value)))
        ])
        if self.define_airgap.Value:
            self.transformer_definition["core_dimensions"]["airgap"]["airgap_on_leg"] = self.airgap_on_leg.Value
            self.transformer_definition["core_dimensions"]["airgap"]["airgap_value"] = str(self.airgap_value.Value)

        # step 2
        winding_property = self.draw_winding

        self.transformer_definition["winding_definition"] = OrderedDict([
            ("draw_winding", str(bool(winding_property.Value)))
        ])

        for prop in ["layer_type", "number_of_layers", "layer_spacing", "bobbin_board_thickness", "top_margin",
                     "side_margin", "conductor_type"]:
            self.transformer_definition["winding_definition"][prop] = str(winding_property.Properties[prop].Value)

        self.transformer_definition["winding_definition"]["include_bobbin"] = str(
                                                            bool(winding_property.Properties["include_bobbin"].Value))

        self.transformer_definition["winding_definition"]["layers_definition"] = OrderedDict()
        if self.conductor_type.Value == "Circular":
            xml_path_to_table = 'winding_properties/draw_winding/conductor_type/table_layers_circles'
            list_of_prop = ["conductor_diameter", "segments_number", "insulation_thickness"]
        else:
            xml_path_to_table = 'winding_properties/draw_winding/conductor_type/table_layers'
            list_of_prop = ["conductor_width", "conductor_height", "insulation_thickness"]

        table = self.step2.Properties[xml_path_to_table]
        for i in range(1, int(self.transformer_definition["winding_definition"]["number_of_layers"]) + 1):
            self.transformer_definition["winding_definition"]["layers_definition"]["layer_" + str(i)] = OrderedDict()
            layer_dict = self.transformer_definition["winding_definition"]["layers_definition"]["layer_" + str(i)]
            for prop in list_of_prop:
                layer_dict[prop] = str(table.Value[xml_path_to_table + "/" + prop][i - 1])
            layer_dict["turns_number"] = str(int(table.Value[xml_path_to_table + "/turns_number"][i - 1]))

        # step 3
        self.transformer_definition["setup_definition"] = OrderedDict([
            ("core_material", self.core_material.Value),
            ("coil_material", self.coil_material.Value),
            ("adaptive_frequency", str(self.adaptive_frequency.Value)),
            ("percentage_error", str(self.percentage_error.Value)),
            ("number_passes", self.number_passes.Value),
            ("transformer_sides", self.transformer_sides.Value),
            ("excitation_strategy", self.excitation_strategy.Value),
            ("voltage", str(self.voltage.Value)),
            ("resistance", str(self.resistance.Value)),
            ("offset", str(self.offset.Value)),
            ("project_path", self.project_path.Value),
        ])

        freq_dict = OrderedDict([("frequency_sweep", str(bool(self.frequency_sweep.Value)))])
        if self.frequency_sweep.Value:
            freq_dict["start_frequency"] = str(self.start_frequency.Value)
            freq_dict["start_frequency_unit"] = self.start_frequency_unit.Value
            freq_dict["stop_frequency"] = str(self.stop_frequency.Value)
            freq_dict["stop_frequency_unit"] = self.stop_frequency_unit.Value
            freq_dict["samples"] = self.samples.Value

        self.transformer_definition["setup_definition"]["frequency_sweep_definition"] = freq_dict

        side_definition = []
        for i in range(1, self.transformer_sides.Value + 1):
            side_definition.append(("Side_" + str(i),
                                    [layer.replace("Side_{}_Layer".format(i), "") for layer in
                                     self.windings_definition.defined_layers_list if "Side_" + str(i) in layer]))

        self.transformer_definition["setup_definition"]["layer_side_definition"] = OrderedDict(side_definition)

    def write_json_data(self):
        self.collect_ui_data()

        write_path = os.path.join(self.project_path.Value, self.design_name + '_parameters.json')
        with open(write_path, "w") as output_f:
            json.dump(self.transformer_definition, output_f, indent=4)

    def populate_ui_data(self):
        # read data for step 1
        self.segmentation_angle.Value = int(self.transformer_definition["core_dimensions"]["segmentation_angle"])
        self.supplier.Value = self.transformer_definition["core_dimensions"]["supplier"]
        self.core_type.Value = self.transformer_definition["core_dimensions"]["core_type"]
        self.core_model.Value = self.transformer_definition["core_dimensions"]["core_model"]

        for i in range(1, 9):
            if "D_" + str(i) in self.transformer_definition["core_dimensions"].keys():
                d_value = self.transformer_definition["core_dimensions"]["D_" + str(i)]
                self.core_dimensions["D_" + str(i)].Value = float(d_value)
            else:
                self.core_dimensions["D_" + str(i)].Visible = False

        self.define_airgap.Value = bool(distutils.util.strtobool(
                                            self.transformer_definition["core_dimensions"]["airgap"]["define_airgap"]))
        if self.define_airgap.Value:
            self.airgap_on_leg.Value = self.transformer_definition["core_dimensions"]["airgap"]["airgap_on_leg"]
            self.airgap_value.Value = float(self.transformer_definition["core_dimensions"]["airgap"]["airgap_value"])

        self.update_ui(self.step1)

        # read data for step 2
        self.draw_winding.Value = bool(distutils.util.strtobool(
                                    self.transformer_definition["winding_definition"]["draw_winding"]))

        if not self.draw_winding.Value:
            return

        self.draw_winding.Properties["layer_type"].Value = self.transformer_definition[
            "winding_definition"]["layer_type"]

        self.draw_winding.Properties["number_of_layers"].Value = int(self.transformer_definition[
            "winding_definition"]["number_of_layers"])

        self.draw_winding.Properties["layer_spacing"].Value = float(self.transformer_definition[
            "winding_definition"]["layer_spacing"])

        self.draw_winding.Properties["bobbin_board_thickness"].Value = float(self.transformer_definition[
            "winding_definition"]["bobbin_board_thickness"])

        self.draw_winding.Properties["top_margin"].Value = float(self.transformer_definition[
            "winding_definition"]["top_margin"])

        self.draw_winding.Properties["side_margin"].Value = float(self.transformer_definition[
            "winding_definition"]["side_margin"])

        self.draw_winding.Properties["include_bobbin"].Value = bool(distutils.util.strtobool(
            self.transformer_definition["winding_definition"]["include_bobbin"]))

        self.draw_winding.Properties["conductor_type"].Value = self.transformer_definition[
            "winding_definition"]["conductor_type"]

        if self.conductor_type.Value == "Circular":
            xml_path_to_table = 'winding_properties/draw_winding/conductor_type/table_layers_circles'
            list_of_prop = ["conductor_diameter", "segments_number", "turns_number", "insulation_thickness"]
        else:
            xml_path_to_table = 'winding_properties/draw_winding/conductor_type/table_layers'
            list_of_prop = ["conductor_width", "conductor_height", "turns_number", "insulation_thickness"]

        table = self.step2.Properties[xml_path_to_table]
        row_num = table.RowCount
        for j in range(0, row_num):
            table.DeleteRow(0)

        for i in range(1, int(self.number_of_layers.Value) + 1):
            try:
                layer_dict = self.transformer_definition["winding_definition"]["layers_definition"]["layer_" + str(i)]
            except KeyError:
                return self.add_error_message("Number of layers does not correspond to defined parameters." +
                                              "Please specify parameters for each layer")

            table.AddRow()
            for prop in list_of_prop:
                if prop in ["segments_number", "turns_number"]:
                    table.Properties[prop].Value = int(layer_dict[prop])
                else:
                    table.Properties[prop].Value = float(layer_dict[prop])

            table.Properties["layer"].Value = 'Layer_' + str(i)
            table.SaveActiveRow()

        self.change_captions(need_refresh=False)

        # read data for step 3
        self.core_material.Value = self.transformer_definition["setup_definition"]["core_material"]
        self.coil_material.Value = self.transformer_definition["setup_definition"]["coil_material"]
        self.adaptive_frequency.Value = float(self.transformer_definition["setup_definition"]["adaptive_frequency"])
        self.percentage_error.Value = float(self.transformer_definition["setup_definition"]["percentage_error"])
        self.number_passes.Value = int(self.transformer_definition["setup_definition"]["number_passes"])

        self.transformer_sides.Value = int(self.transformer_definition["setup_definition"]["transformer_sides"])
        self.excitation_strategy.Value = self.transformer_definition["setup_definition"]["excitation_strategy"]
        self.voltage.Value = float(self.transformer_definition["setup_definition"]["voltage"])
        self.resistance.Value = float(self.transformer_definition["setup_definition"]["resistance"])
        self.offset.Value = float(self.transformer_definition["setup_definition"]["offset"])

        self.project_path.Value = self.transformer_definition["setup_definition"]["project_path"]

        freq_dict = self.transformer_definition["setup_definition"]["frequency_sweep_definition"]
        self.frequency_sweep.Value = bool(distutils.util.strtobool(freq_dict["frequency_sweep"]))

        if self.frequency_sweep.Value:
            self.start_frequency.Value = float(freq_dict["start_frequency"])
            self.start_frequency_unit.Value = freq_dict["start_frequency_unit"]
            self.stop_frequency.Value = float(freq_dict["stop_frequency"])
            self.stop_frequency_unit.Value = freq_dict["stop_frequency_unit"]
            self.samples.Value = int(freq_dict["samples"])

        # disable change of number of layers because Sides are already specified
        self.number_of_layers.ReadOnly = True

        self.defined_layers_list = []
        layer_dict = self.transformer_definition["setup_definition"]["layer_side_definition"]
        for key in layer_dict.keys():
            self.defined_layers_list.extend([key + "_" + "Layer" + str(lay_num) for lay_num in layer_dict[key]])


def calculate_leakage(oDesign, sides_number):
    """Create equations to calculate leakage inductance.
    For N winding transformer would be N*(N-1)/2 equations"""

    oModule = oDesign.GetModule("OutputVariable")
    list_x = list(range(1, sides_number + 1))[:]
    list_y = list(range(1, sides_number + 1))[:]

    all_leakages = []

    for x in list_x:
        for y in list_y:
            if x != y:
                equation = "Matrix1.L(Side_{0},Side_{0})*(1-sqr(Matrix1.CplCoef(Side_{0},Side_{1})))".format(x, y)
                all_leakages.append("LeakageInductance_{}{}".format(x, y))
                oModule.CreateOutputVariable("LeakageInductance_{}{}".format(x, y), equation,
                                             "Setup1 : LastAdaptive", "EddyCurrent", [])

        list_y.remove(x)

    if sides_number <= 1:
        all_leakages = ["Matrix1.L(Side_1,Side_1)"]

    oModule = oDesign.GetModule("ReportSetup")
    oModule.CreateReport("Leakage inductance", "EddyCurrent", "Data Table", "Setup1 : LastAdaptive", [],
                         [
                             "Freq:="	, ["All"]
                         ],
                         [
                             "X Component:="		, "Freq",
                             "Y Component:="		, all_leakages
                         ], [])


def enable_thermal(oDesign, solids):
    """function to turn on feedback for thermal coupling"""
    for i in range(len(solids)):
        solids.insert(i * 2 + 1, "22cel")

    oDesign.SetObjectTemperature(
        [
            "NAME:TemperatureSettings",
            "IncludeTemperatureDependence:=", True,
            "EnableFeedback:=", True,
            "Temperatures:=", solids
        ])


def create_setup(step, oDesign):
    """function which grabs parameters from UI and inserts Setup1 according to them"""
    max_num_passes = int(step.Properties["define_setup/number_passes"].Value)
    percent_error = float(step.Properties["define_setup/percentage_error"].Value)

    adapt_freq = step.Properties["define_setup/adaptive_frequency"].Value
    frequency = str(adapt_freq) + 'Hz'

    start_sweep_freq = (str(step.Properties["define_setup/frequency_sweep/start_frequency"].Value) +
                        str(step.Properties["define_setup/frequency_sweep/start_frequency_unit"].Value))

    stop_sweep_freq = (str(step.Properties["define_setup/frequency_sweep/stop_frequency"].Value) +
                       str(step.Properties["define_setup/frequency_sweep/stop_frequency_unit"].Value))

    samples = int(step.Properties["define_setup/frequency_sweep/samples"].Value)

    if step.Properties["define_setup/frequency_sweep"].Value:
        if step.Properties["define_setup/frequency_sweep/scale"].Value == 'Linear':
            insert_setup(oDesign, max_num_passes, percent_error, frequency, True,
                        'LinearCount', start_sweep_freq, stop_sweep_freq, samples)
        else:
            insert_setup(oDesign, max_num_passes, percent_error, frequency, True,
                        'LogScale', start_sweep_freq, stop_sweep_freq, samples)
    else:
        insert_setup(oDesign, max_num_passes, percent_error, frequency, False)

    return adapt_freq


def assign_length_op(oModule, objects, size):
    oModule.AssignLengthOp(
        [
            "NAME:Length_" + objects[0],
            "RefineInside:=", False,
            "Enabled:=", True,
            "Objects:=", objects,
            "RestrictElem:=", False,
            "NumMaxElem:=", "1000",
            "RestrictLength:=", True,
            "MaxLength:=", str(size) + "mm"
        ])


def create_object_from_face(oEditor, objects, listID):
    oEditor.CreateObjectFromFaces(
        [
            "NAME:Selections",
            "Selections:=", objects,
            "NewPartsModelFlag:=", "Model"
        ],
        [
            "NAME:Parameters",
            [
                "NAME:BodyFromFaceToParameters",
                "FacesToDetach:=", listID
            ]
        ],
        [
            "CreateGroupsForNewObjects:=", False
        ])


def create_region(step, oEditor):
    """Create vacuum region with offset specified by user"""
    offset = int(step.Properties["define_setup/offset"].Value)

    oEditor.CreateRegion(
        [
            "NAME:RegionParameters",
            "+XPaddingType:=", "Percentage Offset",
            "+XPadding:=", offset,
            "-XPaddingType:=", "Percentage Offset",
            "-XPadding:=", offset,
            "+YPaddingType:=", "Percentage Offset",
            "+YPadding:=", offset,
            "-YPaddingType:=", "Percentage Offset",
            "-YPadding:=", offset,
            "+ZPaddingType:=", "Percentage Offset",
            "+ZPadding:=", offset,
            "-ZPaddingType:=", "Percentage Offset",
            "-ZPadding:=", offset
        ],
        [
            "NAME:Attributes",
            "Name:=", "Region",
            "Flags:=", "Wireframe#",
            "Color:=", "(255 0 0)",
            "Transparency:=", 0,
            "PartCoordinateSystem:=", "Global",
            "UDMId:=", "",
            "MaterialValue:=", "\"vacuum\"",
            "SolveInside:=", True
        ])


def create_terminal_sections(oEditor, LayList, LaySecList, LayDelList):
    # only for EFD core not to get an error due to section of winding
    if "CentralLegCS" in oEditor.GetCoordinateSystems():
        oEditor.SetWCS(
            [
                "NAME:SetWCS Parameter",
                "Working Coordinate System:=", "CentralLegCS",
                "RegionDepCSOk:=", False
            ])

    oEditor.Section(
        [
            "NAME:Selections",
            "Selections:=", ','.join(LayList),
            "NewPartsModelFlag:=", "Model"
        ],
        [
            "NAME:SectionToParameters",
            "CreateNewObjects:=", True,
            "SectionPlane:=", "ZX",
            "SectionCrossObject:=", False
        ])

    oEditor.SeparateBody(
        [
            "NAME:Selections",
            "Selections:=", ','.join(LaySecList),
            "NewPartsModelFlag:=", "Model"
        ])

    oEditor.Delete(["NAME:Selections", "Selections:=", ','.join(LayDelList)])


def create_new_materials(oProject, core_material, coil_material):
    oDefinitionManager = oProject.GetDefinitionManager()

    # check if we are not having core material already
    if not oDefinitionManager.DoesMaterialExist("Material_" + core_material):
        cord_list = ["NAME:Coordinates"]
        for coordinatePair in range(len(matKeyPoints[core_material])):
            cord_list.append(["NAME:Coordinate", "X:=", matKeyPoints[core_material][coordinatePair][0],
                              "Y:=", matKeyPoints[core_material][coordinatePair][1]])

        oProject.AddDataset(["NAME:$Mu_" + core_material, cord_list])

        oDefinitionManager.AddMaterial(
            [
                "NAME:Material_" + core_material,
                "CoordinateSystemType:=", "Cartesian",
                "BulkOrSurfaceType:=", 1,
                [
                    "NAME:PhysicsTypes",
                    "set:=", ["Electromagnetic", "Thermal", "Structural"]
                ],
                ["NAME:AttachedData"],
                ["NAME:ModifierData"],
                "permeability:=", "pwl($Mu_" + core_material + ",Freq)",
                "conductivity:=", matDict[core_material][0],
                [
                    "NAME:core_loss_type",
                    "property_type:=", "ChoiceProperty",
                    "Choice:=", "Power Ferrite"
                ],
                "core_loss_cm:=", matDict[core_material][1],
                "core_loss_x:=", matDict[core_material][2],
                "core_loss_y:=", matDict[core_material][3],
                "core_loss_kdc:=", "0",
                "thermal_conductivity:=", "5",
                "mass_density:=", matDict[core_material][4],
                "specific_heat:=", "750",
                "thermal_expansion_coeffcient:=", "1e-05"
            ])

    # check if winding material exists
    if (coil_material == "Copper_temperature" and
            not oDefinitionManager.DoesMaterialExist(coil_material)):

        oDefinitionManager.AddMaterial(
            [
                "NAME:" + coil_material,
                "CoordinateSystemType:=", "Cartesian",
                "BulkOrSurfaceType:=", 1,
                [
                    "NAME:PhysicsTypes",
                    "set:=", ["Electromagnetic", "Thermal", "Structural"]
                ],
                [
                    "NAME:ModifierData",
                    [
                        "NAME:ThermalModifierData",
                        "modifier_data:=", "thermal_modifier_data",
                        [
                            "NAME:all_thermal_modifiers",
                            [
                                "NAME:one_thermal_modifier",
                                "Property::=", "conductivity",
                                "Index::=", 0,
                                "prop_modifier:=", "thermal_modifier",
                                "use_free_form:=", True,
                                "free_form_value:=", "1 / (1 + 0.0039 * (Temp - 22))"
                            ]
                        ]
                    ]
                ],
                "permeability:=", "0.999991",
                "conductivity:=", "58000000",
                "thermal_conductivity:=", "400",
                "mass_density:=", "8933",
                "specific_heat:=", "385",
                "youngs_modulus:=", "120000000000",
                "poissons_ratio:=", "0.38",
                "thermal_expansion_coeffcient:=", "1.77e-05"
            ])

    elif (coil_material == "Aluminum_temperature" and
          not oDefinitionManager.DoesMaterialExist(coil_material)):
        oDefinitionManager.AddMaterial(
            [
                "NAME:" + coil_material,
                "CoordinateSystemType:=", "Cartesian",
                "BulkOrSurfaceType:=", 1,
                [
                    "NAME:PhysicsTypes",
                    "set:=", ["Electromagnetic", "Thermal", "Structural"]
                ],
                [
                    "NAME:ModifierData",
                    [
                        "NAME:ThermalModifierData",
                        "modifier_data:=", "thermal_modifier_data",
                        [
                            "NAME:all_thermal_modifiers",
                            [
                                "NAME:one_thermal_modifier",
                                "Property::=", "conductivity",
                                "Index::=", 0,
                                "prop_modifier:=", "thermal_modifier",
                                "use_free_form:=", True,
                                "free_form_value:=", "1 / (1 + 0.0039 * (Temp - 22))"
                            ]
                        ]
                    ]
                ],
                "permeability:=", "1.000021",
                "conductivity:=", "38000000",
                "thermal_conductivity:=", "237.5",
                "mass_density:=", "2689",
                "specific_heat:=", "951",
                "youngs_modulus:=", "69000000000",
                "poissons_ratio:=", "0.31",
                "thermal_expansion_coeffcient:=", "2.33e-05"
            ])


def assign_material(oEditor, selection, material):
    oEditor.AssignMaterial(
        [
            "NAME:Selections",
            "Selections:=", selection
        ],
        [
            "NAME:Attributes",
            "MaterialValue:=", material,
            "SolveInside:=", True
        ])


def insert_setup(oDesign, max_num_passes, percent_error, frequency, has_sweep,
                 sweep_type='', start_sweep_freq='', stop_sweep_freq='', samples=''):
    oModule = oDesign.GetModule("AnalysisSetup")
    oModule.InsertSetup("EddyCurrent",
                                            [
                                                "NAME:Setup1",
                                                "Enabled:=", True,
                                                "MaximumPasses:=", max_num_passes,
                                                "MinimumPasses:=", 2,
                                                "MinimumConvergedPasses:=", 1,
                                                "PercentRefinement:=", 30,
                                                "SolveFieldOnly:=", False,
                                                "PercentError:=", percent_error,
                                                "SolveMatrixAtLast:=", True,
                                                "PercentError:=", percent_error,
                                                "UseIterativeSolver:=", False,
                                                "RelativeResidual:=", 0.0001,
                                                "ComputeForceDensity:=", False,
                                                "ComputePowerLoss:=", False,
                                                "Frequency:=", frequency,
                                                "HasSweepSetup:=", has_sweep,
                                                "SweepSetupType:=", sweep_type,
                                                "StartValue:=", start_sweep_freq,
                                                "StopValue:=", stop_sweep_freq,
                                                "Samples:=", samples,
                                                "SaveAllFields:=", True,
                                                "UseHighOrderShapeFunc:=", False
                                            ])


def change_color(oEditor, selection, R, G, B):
    oEditor.ChangeProperty(
        ["NAME:AllTabs", ["NAME:Geometry3DAttributeTab",
                                            ["NAME:PropServers"] + selection,
                                            ["NAME:ChangedProps",
                                             [
                                                 "NAME:Color",
                                                 "R:=", R,
                                                 "G:=", G,
                                                 "B:=", B
                                             ]
                                             ]]])


def on_init_step1(step):
    """invoke on step initialisation"""
    global transformer
    transformer = TransformerClass(step)
    transformer.initialize_step1()


def create_buttons_step1(step):
    """invoke on step1 refresh"""
    transformer.refresh_step1()


def callback_step1(step):
    """invoke on click 'Next' button in UI"""
    transformer.callback_step1()


def on_supplier_core_type_change(step, prop):
    transformer.show_core_img()


def on_core_model_change(step, prop):
    transformer.insert_default_values()


def on_init_step2(step):
    """invoke on step initialisation"""
    transformer.initialize_step2()


def create_buttons_step2(step):
    """invoke on step1 refresh"""
    transformer.refresh_step2()


def callback_step2(step):
    """invoke on click 'Next' button in UI"""
    transformer.callback_step2()


def on_layers_number_change(step, prop):
    transformer.update_rows()


def on_layer_type_change(step, prop):
    transformer.change_captions()


def on_excitation_strategy_change(step, prop):
    transformer.excitation_strategy_change()


def on_init_step3(step):
    """invoke on step initialisation"""
    transformer.initialize_step3()


def on_reset_step3(step):
    """invoke on step initialisation"""
    transformer.reset_step3()


def create_buttons_step3(step):
    """invoke on step1 refresh"""
    transformer.refresh_step3()


def on_step_back(step):
    pass
