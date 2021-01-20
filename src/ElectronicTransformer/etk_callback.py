# -*- coding: utf-8 -*-

#            Replacement of old ETK (Electronic Transformer Kit) for Maxwell
#
#            ACT Written by : Maksim Beliaev (maksim.beliaev@ansys.com)
#            Tested by: Mark Christini (mark.christini@ansys.com)
#            Last updated : 18.01.2021
import copy
import datetime
import os
import re
import shutil
from webbrowser import open as webopen
from collections import OrderedDict
import json
from abc import ABCMeta, abstractmethod

# all below imports are done by ACT, here they are listed for information purpose to increase transparency
# from .cores_geometry import ECore, EFDCore, EICore, EPCore, ETDCore, PCore, PQCore, UCore, UICore, RMCore
# from .value_checker import check_core_dimensions
# from .forms import WindingForm, ConnectionForm

# object to save UI settings that are cross shared between classes
transformer_definition = OrderedDict()


def atoi(letter):
    return int(letter) if letter.isdigit() else letter


def natural_keys(text):
    return [atoi(c) for c in re.split(r'(\d+)', text)]


def setup_button(step, comp_name, caption, position, callback, active=True, style=None):
    # Setup a button
    update_btn_session = step.UserInterface.GetComponent(comp_name)  # get component name from XML
    update_btn_session.AddButton(comp_name, caption, position, active)  # add button caption and position
    update_btn_session.ButtonClicked += callback  # connect to callback function
    if style == "blue":
        # change CSS properties of the button to change it's color
        update_btn_session.AddCSSProperty("background-color", "#3383ff", "button")
        update_btn_session.AddCSSProperty("color", "white", "button")


def help_button_clicked(_sender, _args):
    """when user clicks Help button HTML page will be opened in standard web browser"""
    webopen(str(ExtAPI.Extension.InstallDir) + '/help/help.html')


def update_ui(step):
    """Refresh UI data """
    step.UserInterface.GetComponent("Properties").UpdateData()
    step.UserInterface.GetComponent("Properties").Refresh()


def add_info_message(msg):
    oDesktop.AddMessage("", "", 0, "ACT:" + str(msg))


def add_warning_message(msg):
    oDesktop.AddMessage("", "", 1, "ACT:" + str(msg))


def add_error_message(msg):
    oDesktop.AddMessage("", "", 2, "ACT:" + str(msg))


def verify_input_data(function):
    """
    Check that all input data is present in JSON file for each step
    :param function: function to populate ui
    :return: None
    """
    def check(*args):
        try:
            function(*args)
        except ValueError:
            msg = "Please verify that integer numbers in input file have proper format eg 1 and not 1.0"
            return add_error_message(msg)
        except KeyError as e:
            return add_error_message("Please specify parameter:{} in input file".format(e))
    return check


class Step1(object):
    __metaclass__ = ABCMeta

    def __init__(self, step):
        self.step1 = step.Wizard.Steps["step1"]

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

        self.core_models = []
        self.cores_database = {}
        self.personal_lib_path = ""

    def github(self, _sender, _args):
        """
        opens GitHub page with project
        """
        webopen(r"https://github.com/beliaev-maksim/ansys-electronic-transformer-act")

    def open_custom_lib(self, _sender, _args):
        """
        Opens folder with custom library definitions
        """
        os.startfile(self.personal_lib_path)

    def open_examples(self, _sender, _args):
        """
        Opens read data method in folder with examples
        """

        examples_folder = os.path.join(ExtAPI.Extension.InstallDir, "examples").replace("/", "\\")
        self.read_data(_sender, _args, default_path=examples_folder)

    def read_data(self, _sender, _args, default_path=""):
        """
        Function called when click button Read Settings From File. Parse json file with settings and dump them to
        transformer definition object
        :param default_path: path to the folder that should be opened
        :param _sender: unused standard event
        :param _args: unused standard event
        :return: None
        """
        global transformer_definition
        path = ExtAPI.UserInterface.UIRenderer.ShowFileOpenDialog('Text Files(*.json;)|*.json;', default_path)

        if path is None:
            return

        with open(path, "r") as input_f:
            try:
                transformer_definition = json.load(input_f, object_pairs_hook=OrderedDict)
            except ValueError as e:
                match_line = re.match("Expecting object: line (.*) column", str(e))
                if match_line:
                    add_error_message(
                        'Please verify that all data in file is covered by double quotes "" ' +
                        '(integers and floats can be both covered or uncovered)')
                else:
                    match_line = re.match("Expecting property name: line (.*) column", str(e))
                    if match_line:
                        add_error_message("Please verify that there is no empty argument in the file. "
                                          "Cannot be two commas in a row without argument in between")
                    else:
                        # so smth unexpected
                        return add_error_message(e)

                return add_error_message("Please correct following line: {} in file: {}".format(
                                                                                            match_line.group(1), path))
        self.populate_ui_data_step1()

    def refresh_step1(self):
        """create buttons and HTML data for first step"""
        setup_button(self.step1, "read_data_button", "Read Settings File", ButtonPositionType.Left, self.read_data)
        setup_button(self.step1, "open_examples_button", "Open Examples", ButtonPositionType.Left, self.open_examples)
        setup_button(self.step1, "open_library_button", "Custom Library", ButtonPositionType.Center, self.open_custom_lib)
        setup_button(self.step1, "github_button", "Contribute on GitHub", ButtonPositionType.Center, self.github)
        setup_button(self.step1, "help_button", "Help", ButtonPositionType.Center, help_button_clicked, style="blue")
        self.personal_lib_path = os.path.join(oDesktop.GetPersonalLibDirectory(), 'ElectronicTransformer')

        if not transformer_definition:
            self.read_core_dimensions()
            self.prefill_supplier()
        else:
            self.show_core_img()

        validate_aedt_version()

    @verify_input_data
    def populate_ui_data_step1(self):
        """
        Set data to UI fields from transformer_definition dictionary
        """
        self.segmentation_angle.Value = int(transformer_definition["core_dimensions"]["segmentation_angle"])
        self.supplier.Value = transformer_definition["core_dimensions"]["supplier"]
        self.core_type.Value = transformer_definition["core_dimensions"]["core_type"]
        self.prefill_core_types(only_menu=True)  # populate drop down menu once we read core type
        self.core_model.Value = transformer_definition["core_dimensions"]["core_model"]

        for i in range(1, 9):
            if "D_" + str(i) in transformer_definition["core_dimensions"].keys():
                d_value = transformer_definition["core_dimensions"]["D_" + str(i)]
                self.core_dimensions["D_" + str(i)].Value = float(d_value)
            else:
                self.core_dimensions["D_" + str(i)].Visible = False

        self.define_airgap.Value = transformer_definition["core_dimensions"]["airgap"]["define_airgap"]
        if self.define_airgap.Value:
            self.airgap_on_leg.Value = transformer_definition["core_dimensions"]["airgap"]["airgap_on_leg"]
            self.airgap_value.Value = float(transformer_definition["core_dimensions"]["airgap"]["airgap_value"])

        update_ui(self.step1)

    def callback_step1(self):
        check_core_dimensions(self.step1)
        self.collect_ui_data_step1()

    def show_core_img(self):
        """invoked to change image and core dimensions when supplier or core type changed"""
        if self.core_type.Value not in ['EP', 'ER', 'PQ', 'RM']:
            width = 300
            height = 200
        else:
            width = 275
            height = 360

        html_data = '<img width="{}" height="{}" src="file:///{}/images/{}Core.png"/>'.format(width, height,
                                                                                          ExtAPI.Extension.InstallDir,
                                                                                          self.core_type.Value)

        report = self.step1.UserInterface.GetComponent("coreImage")
        report.SetHtmlContent(html_data)
        report.Refresh()

    def prefill_supplier(self):
        """
        Read supplier from input data
        :return:
        """
        self.supplier.Options.Clear()
        for key in sorted(self.cores_database.keys()):
            self.supplier.Options.Add(key)

        self.supplier.Value = self.supplier.Options[0]
        self.prefill_core_types()

    def prefill_core_types(self, only_menu=False):
        """
        Read core types from input data
        :return:
        """
        self.core_type.Options.Clear()
        for key in sorted(self.cores_database[self.supplier.Value].keys()):
            self.core_type.Options.Add(key)

        if not only_menu:
            self.core_type.Value = self.core_type.Options[0]
        self.prefill_core_models(only_menu)

    def prefill_core_models(self, only_menu=False):
        """
        Read core models from input data
        :return:
        """
        self.core_model.Options.Clear()
        self.core_models = self.cores_database[self.supplier.Value][self.core_type.Value]
        for model in sorted(self.core_models.keys(), key=natural_keys):
            self.core_model.Options.Add(model)
        self.core_model.Value = self.core_model.Options[0]

        self.show_core_img()

        if not only_menu:
            self.prefill_core_dimensions()

        update_ui(self.step1)

    def prefill_core_dimensions(self):
        """
        Set core dimensions from the predefined lists
        :return:
        """
        for j in range(1, 9):
            try:
                self.core_dimensions["D_" + str(j)].Value = float(self.core_models[self.core_model.Value][j - 1])
                self.core_dimensions["D_" + str(j)].Visible = True
            except ValueError:
                self.core_dimensions["D_" + str(j)].Visible = False

    def read_core_dimensions(self):
        """
        Read all possible core dimensions from input file
        :return:
        """
        core_dims_json = os.path.join(self.personal_lib_path, "core_dimensions.json")
        if not os.path.isfile(core_dims_json):
            # file does not exist, copy it from root location
            if not os.path.exists(self.personal_lib_path):
                os.makedirs(self.personal_lib_path)
            shutil.copy(os.path.join(ExtAPI.Extension.InstallDir, "core_dimensions.json"), core_dims_json)

        with open(core_dims_json) as file:
            self.cores_database = json.load(file)

    def collect_ui_data_step1(self):
        """collect data from all steps, write this data to dictionary"""
        # step1
        transformer_definition["core_dimensions"] = OrderedDict([
            ("segmentation_angle", self.segmentation_angle.Value),
            ("supplier", self.supplier.Value),
            ("core_type", self.core_type.Value),
            ("core_model", self.core_model.Value)
        ])

        for i in range(1, 9):
            if self.core_dimensions["D_" + str(i)].Visible:
                d_value = str(self.core_dimensions["D_" + str(i)].Value)
            else:
                d_value = 0  # fill with zeros for cores where no parameter

            transformer_definition["core_dimensions"]["D_" + str(i)] = d_value

        transformer_definition["core_dimensions"]["airgap"] = OrderedDict([
            ("define_airgap", bool(self.define_airgap.Value))  # need to specify boolean due to bug 324104
        ])
        if self.define_airgap.Value:
            transformer_definition["core_dimensions"]["airgap"]["airgap_on_leg"] = self.airgap_on_leg.Value
            transformer_definition["core_dimensions"]["airgap"]["airgap_value"] = str(self.airgap_value.Value)


class Step2(object):
    __metaclass__ = ABCMeta

    def __init__(self, step):
        self.step2 = step.Wizard.Steps["step2"]

        self.winding_prop = self.step2.Properties["winding_properties"]
        self.layer_type = self.step2.Properties["winding_properties/layer_type"]
        self.number_of_layers = self.step2.Properties["winding_properties/number_of_layers"]
        self.layer_spacing = self.step2.Properties["winding_properties/layer_spacing"]
        self.bobbin_board_thickness = self.step2.Properties["winding_properties/bobbin_board_thickness"]
        self.top_margin = self.step2.Properties["winding_properties/top_margin"]
        self.side_margin = self.step2.Properties["winding_properties/side_margin"]
        self.include_bobbin = self.step2.Properties["winding_properties/include_bobbin"]
        self.conductor_type = self.step2.Properties["winding_properties/conductor_type"]

        self.table_layers = self.conductor_type.Properties["table_layers"]
        self.table_layers_circles = self.conductor_type.Properties["table_layers_circles"]

        self.skip_check = self.step2.Properties["winding_properties/skip_check"]

        self.materials = {}

    def init_tables_step2(self):
        """initialize tables with some initial data"""
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
        """
        each time step layout is opening
        :return:
        """
        setup_button(self.step2, "help_button", "Help", ButtonPositionType.Right, help_button_clicked, style="blue")
        if "winding_definition" in transformer_definition:
            # that mean that we read settings file from user and need to populate UI
            self.populate_ui_data_step2()

        update_ui(self.step2)

    @verify_input_data
    def populate_ui_data_step2(self):
        # read data for step 2
        winding_def = transformer_definition["winding_definition"]
        self.winding_prop.Properties["layer_type"].Value = winding_def["layer_type"]

        self.winding_prop.Properties["number_of_layers"].Value = int(winding_def["number_of_layers"])

        self.winding_prop.Properties["layer_spacing"].Value = float(winding_def["layer_spacing"])

        self.winding_prop.Properties["bobbin_board_thickness"].Value = float(winding_def["bobbin_board_thickness"])

        self.winding_prop.Properties["top_margin"].Value = float(winding_def["top_margin"])

        self.winding_prop.Properties["side_margin"].Value = float(winding_def["side_margin"])

        self.winding_prop.Properties["include_bobbin"].Value = winding_def["include_bobbin"]

        self.winding_prop.Properties["conductor_type"].Value = winding_def["conductor_type"]

        if self.conductor_type.Value == "Circular":
            xml_path_to_table = 'winding_properties/conductor_type/table_layers_circles'
            list_of_prop = ["conductor_diameter", "segments_number", "turns_number", "insulation_thickness"]
        else:
            xml_path_to_table = 'winding_properties/conductor_type/table_layers'
            list_of_prop = ["conductor_width", "conductor_height", "turns_number", "insulation_thickness"]

        table = self.step2.Properties[xml_path_to_table]
        row_num = table.RowCount
        for j in range(0, row_num):
            table.DeleteRow(0)

        for i in range(1, int(self.number_of_layers.Value) + 1):
            try:
                layer_dict = winding_def["layers_definition"]["layer_" + str(i)]
            except KeyError:
                return add_error_message("Number of layers does not correspond to defined parameters." +
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

    def callback_step2(self):
        """
        Called when Next button is clicked on step 2
        :return:
        """
        # invoke validation from value_checker file
        self.check_winding()
        self.check_board_bobbin()
        self.read_material_data()

        self.collect_ui_data_step2()

    def read_material_data(self):
        """
        Read material definition dictionary from json file
        :return:
        """
        lib_path = os.path.join(oDesktop.GetPersonalLibDirectory(), 'ElectronicTransformer')
        materials_json = os.path.join(lib_path, "material_properties.json")
        if not os.path.isfile(materials_json):
            # file does not exist, copy it from root location
            if not os.path.exists(lib_path):
                os.makedirs(lib_path)
            shutil.copy(os.path.join(ExtAPI.Extension.InstallDir, "material_properties.json"), materials_json)

        with open(materials_json) as file:
            self.materials = json.load(file)

        if not self.materials:
            raise UserErrorMessageException("Materials are not defined in {}".format(materials_json))

        for mat, mat_definition in self.materials.items():
            for prop in ["conductivity", "cm", "x", "y", "density", "mu(freq)"]:
                if prop not in mat_definition:
                    msg = 'Property {} is not specified for material {}'.format(prop, mat)
                    raise UserErrorMessageException(msg)

    @staticmethod
    def reset_step2():
        """when back button on step3 is clicked we return one step back and we need to delete all winding definitions"""
        transformer_definition.pop("winding_definition", None)
        transformer_definition.pop("setup_definition", None)

    @abstractmethod
    def check_winding(self):
        """
        Validate that windings fit into core. For this we need to have access to core dimensions
        :return:
        """

    def collect_ui_data_step2(self):
        """collect data from all steps, write this data to dictionary"""
        # step 2
        winding_definition = OrderedDict()

        for prop in ["layer_type", "number_of_layers", "layer_spacing", "bobbin_board_thickness", "top_margin",
                     "side_margin", "conductor_type"]:
            winding_definition[prop] = str(self.winding_prop.Properties[prop].Value)

        winding_definition["include_bobbin"] = bool(self.include_bobbin.Value)

        winding_definition["layers_definition"] = OrderedDict()
        if self.conductor_type.Value == "Circular":
            xml_path_to_table = 'winding_properties/conductor_type/table_layers_circles'
            list_of_prop = ["conductor_diameter", "segments_number", "insulation_thickness"]
        else:
            xml_path_to_table = 'winding_properties/conductor_type/table_layers'
            list_of_prop = ["conductor_width", "conductor_height", "insulation_thickness"]

        table = self.step2.Properties[xml_path_to_table]
        for i in range(1, int(winding_definition["number_of_layers"]) + 1):
            layer_dict = OrderedDict()
            for prop in list_of_prop:
                layer_dict[prop] = table.Value[xml_path_to_table + "/" + prop][i - 1]
            layer_dict["turns_number"] = int(table.Value[xml_path_to_table + "/turns_number"][i - 1])
            winding_definition["layers_definition"]["layer_" + str(i)] = layer_dict

        transformer_definition["winding_definition"] = winding_definition

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
            update_ui(self.step2)

    @staticmethod
    def warn_about_winding_def_clean_up():
        if "setup_definition" in transformer_definition:
            msg = "Number of layers changed, this will cause clean up of winding and connection definitions"
            add_warning_message(msg)
            transformer_definition["setup_definition"]["layer_side_definition"] = {}
            transformer_definition["setup_definition"]["connections_definition"] = {}

    def update_rows(self):
        """when user changes number of layers we append/delete rows in tables"""
        if self.number_of_layers.Value < 1:
            self.number_of_layers.Value = 1
            add_error_message("Number of layers should be greater or equal than 1")
            return False

        if self.conductor_type.Value == 'Rectangular':
            xml_path_to_table = "winding_properties/conductor_type/table_layers"
            list_of_prop = ["conductor_width", "conductor_height", "turns_number", "insulation_thickness"]
        else:
            xml_path_to_table = "winding_properties/conductor_type/table_layers_circles"
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


class Step3(object):
    __metaclass__ = ABCMeta

    def __init__(self, step):
        self.step3 = step.Wizard.Steps["step3"]

        self.core_material = self.step3.Properties["define_setup/core_material"]
        self.coil_material = self.step3.Properties["define_setup/coil_material"]

        self.adaptive_frequency = self.step3.Properties["define_setup/adaptive_frequency"]
        self.draw_skin_layers = self.step3.Properties["define_setup/draw_skin_layers"]
        self.percentage_error = self.step3.Properties["define_setup/percentage_error"]
        self.number_passes = self.step3.Properties["define_setup/number_passes"]

        self.transformer_sides = self.step3.Properties["define_setup/transformer_sides"]
        self.excitation_strategy = self.step3.Properties["define_setup/excitation_strategy"]
        self.voltage = self.step3.Properties["define_setup/voltage"]
        self.resistance = self.step3.Properties["define_setup/resistance"]

        self.offset = self.step3.Properties["define_setup/offset"]
        self.full_model = self.step3.Properties["define_setup/full_model"]
        self.project_path = self.step3.Properties["define_setup/project_path"]

        self.frequency_sweep = self.step3.Properties["define_setup/frequency_sweep"]
        self.start_frequency = self.step3.Properties["define_setup/frequency_sweep/start_frequency"]
        self.start_frequency_unit = self.step3.Properties["define_setup/frequency_sweep/start_frequency_unit"]
        self.stop_frequency = self.step3.Properties["define_setup/frequency_sweep/stop_frequency"]
        self.stop_frequency_unit = self.step3.Properties["define_setup/frequency_sweep/stop_frequency_unit"]
        self.samples = self.step3.Properties["define_setup/frequency_sweep/samples"]
        self.scale = self.step3.Properties["define_setup/frequency_sweep/scale"]

        self.windings_def_form = None
        self.connection_def_form = ConnectionForm()

        self.analysis_set = False
        self.project = None

        self.defined_layers_dict = []
        self.defined_connections_dict = {}

        self.materials = {}  # should be redefined in child class (from class 2)

    def refresh_step3(self):
        """
        Method is called every time when step layout opens
        :return:
        """
        setup_button(self.step3, "help_button", "Help", ButtonPositionType.Right, help_button_clicked, style="blue")

        setup_button(self.step3, "analyze_button", "Analyze", ButtonPositionType.Center, self.analyze_click,
                     active=False)

        setup_button(self.step3, "setup_analysis_button", "Setup Analysis", ButtonPositionType.Center,
                     self.setup_analysis_click, active=False)

        setup_button(self.step3, "define_windings_button", "Define Windings", ButtonPositionType.Center,
                     self.define_windings_click)

        setup_button(self.step3, "define_connections_button", "Define Connections", ButtonPositionType.Center,
                     self.define_connection_click, active=False)

        self.analysis_set = False
        self.project = oDesktop.GetActiveProject()
        if self.project is None:
            self.project = oDesktop.NewProject()

        # insert path to the saved project as default path to textbox
        self.project_path.Value = self.project.GetPath()

        # add materials from dictionary
        self.core_material.Options.Clear()
        material = None
        for material in sorted(self.materials.keys()):
            self.core_material.Options.Add(material)

        self.core_material.Value = material

        if "setup_definition" in transformer_definition:
            # that means that we read settings file from user and need to populate UI
            self.populate_ui_data_step3()
            self.resistance.Visible = False if self.transformer_sides.Value == 1 else True
        else:
            # just refresh variables in case if user clicked back button and came back to the page or first run
            self.defined_layers_dict = {}
            self.defined_connections_dict = {}

            self.connection_def_form.connections_dict = {}

            self.transformer_sides.Value = 1
            self.resistance.Visible = False

        if transformer_definition["core_dimensions"]["core_type"] in ["U", "UI"]:
            # cannot split U and UI core due to the nature of the cores
            self.full_model.Value = True
            self.full_model.ReadOnly = True

        update_ui(self.step3)

        if self.defined_layers_dict:
            self.step3.UserInterface.GetComponent("define_connections_button").SetEnabledFlag(
                "define_connections_button", True
            )
            if self.defined_connections_dict:
                self.step3.UserInterface.GetComponent("analyze_button").SetEnabledFlag("analyze_button", True)
                self.step3.UserInterface.GetComponent("setup_analysis_button").SetEnabledFlag("setup_analysis_button",
                                                                                              True)

    @verify_input_data
    def populate_ui_data_step3(self):
        """
        Fill values in the user interface from the dictionary, only when start from text file
        :return:
        """
        setup_def_dict = transformer_definition["setup_definition"]
        self.core_material.Value = setup_def_dict["core_material"]
        self.coil_material.Value = setup_def_dict["coil_material"]
        self.adaptive_frequency.Value = float(setup_def_dict["adaptive_frequency"])
        self.draw_skin_layers.Value = setup_def_dict["draw_skin_layers"]
        self.percentage_error.Value = float(setup_def_dict["percentage_error"])
        self.number_passes.Value = int(setup_def_dict["number_passes"])

        self.transformer_sides.Value = int(setup_def_dict["transformer_sides"])
        self.check_sides(manual=True)

        self.excitation_strategy.Value = setup_def_dict["excitation_strategy"]
        self.voltage.Value = float(setup_def_dict["voltage"])
        self.resistance.Value = float(setup_def_dict["resistance"])
        self.offset.Value = float(setup_def_dict["offset"])
        self.full_model.Value = setup_def_dict["full_model"]

        if os.path.isdir(setup_def_dict["project_path"]):
            self.project_path.Value = setup_def_dict["project_path"]

        freq_dict = setup_def_dict["frequency_sweep_definition"]
        self.frequency_sweep.Value = freq_dict["frequency_sweep"]

        if self.frequency_sweep.Value:
            self.start_frequency.Value = float(freq_dict["start_frequency"])
            self.start_frequency_unit.Value = freq_dict["start_frequency_unit"]
            self.stop_frequency.Value = float(freq_dict["stop_frequency"])
            self.stop_frequency_unit.Value = freq_dict["stop_frequency_unit"]
            self.samples.Value = int(freq_dict["samples"])
            self.scale.Value = freq_dict["scale"]

        self.defined_layers_dict = copy.deepcopy(setup_def_dict["layer_side_definition"])
        self.defined_connections_dict = copy.deepcopy(setup_def_dict["connections_definition"])

    @abstractmethod
    def check_sides(self, manual=False):
        """if number of sides was changed this method will be called"""

    @abstractmethod
    def analyze_click(self, _sender, _args):
        pass

    @abstractmethod
    def setup_analysis_click(self, _sender="", _args=""):
        pass

    @abstractmethod
    def define_windings_click(self, _sender, _args):
        pass

    @abstractmethod
    def define_connection_click(self, _sender, _args):
        pass

    def collect_ui_data_step3(self):
        """collect data from UI, write this data to dictionary"""
        transformer_definition["setup_definition"] = OrderedDict([
            ("core_material", self.core_material.Value),
            ("coil_material", self.coil_material.Value),
            ("adaptive_frequency", str(self.adaptive_frequency.Value)),
            ("draw_skin_layers", bool(self.draw_skin_layers.Value)),
            ("percentage_error", str(self.percentage_error.Value)),
            ("number_passes", self.number_passes.Value),
            ("transformer_sides", self.transformer_sides.Value),
            ("excitation_strategy", self.excitation_strategy.Value),
            ("voltage", str(self.voltage.Value)),
            ("resistance", str(self.resistance.Value)),
            ("offset", str(self.offset.Value)),
            ("full_model", bool(self.full_model.Value)),
            ("project_path", self.project_path.Value),
        ])

        freq_dict = OrderedDict([("frequency_sweep", bool(self.frequency_sweep.Value))])
        if self.frequency_sweep.Value:
            freq_dict["start_frequency"] = str(self.start_frequency.Value)
            freq_dict["start_frequency_unit"] = self.start_frequency_unit.Value
            freq_dict["stop_frequency"] = str(self.stop_frequency.Value)
            freq_dict["stop_frequency_unit"] = self.stop_frequency_unit.Value
            freq_dict["samples"] = self.samples.Value
            freq_dict["scale"] = self.scale.Value

        transformer_definition["setup_definition"]["frequency_sweep_definition"] = freq_dict

        transformer_definition["setup_definition"]["layer_side_definition"] = self.defined_layers_dict
        transformer_definition["setup_definition"]["connections_definition"] = self.defined_connections_dict


class TransformerClass(Step1, Step2, Step3):
    """Main class which serves to manipulate Step classes. We initialize these classes from current class.
    Invoke all functions from each step here.
    Class TransformerClass will inherit all functions from classes Step1, Step2, Step3 including all callback actions"""
    def __init__(self, step):
        self.step1 = step

        self.design_name = ""

        self.project = None
        self.design = None
        self.editor = None

        self.module_analysis = None
        self.module_boundary_setup = None
        self.module_parameter_setup = None
        self.module_output_var = None
        self.module_report = None
        self.module_mesh = None
        self.module_fields_reporter = None

        self.flag_auto_save = True

        self.circuit = None

    def initialize_step1(self):
        """separate init is required because we do not have access to steps without XML callback"""
        Step1.__init__(self, self.step1)

    def initialize_step2(self):
        """separate init is required because we do not have access to steps without XML callback"""
        Step2.__init__(self, self.step1)
        Step2.init_tables_step2(self)

    def initialize_step3(self):
        """separate init is required because we do not have access to steps without XML callback"""
        Step3.__init__(self, self.step1)

    def write_json_data(self):
        """
        Save transformer definition to the file
        :return: None
        """
        write_path = os.path.join(self.project_path.Value, self.design_name + '_parameters.json')
        with open(write_path, "w") as output_f:
            json.dump(transformer_definition, output_f, indent=4)

    def mirror(self, objects):
        """
        Mirrors object around YZ plane
        :param objects: (str) objects, comma separated
        :return:
        """
        self.editor.Mirror(
            [
                "NAME:Selections",
                "Selections:="		, objects,
                "NewPartsModelFlag:="	, "Model"
            ],
            [
                "NAME:MirrorParameters",
                "MirrorBaseX:="		, "0mm",
                "MirrorBaseY:="		, "0mm",
                "MirrorBaseZ:="		, "0mm",
                "MirrorNormalX:="	, "1mm",
                "MirrorNormalY:="	, "0mm",
                "MirrorNormalZ:="	, "0mm"
            ])

    def create_field_plot(self, quantity, folder, adapt_freq, object_list):
        """Create filed overlay on the surface of the objects"""

        ids_list = [str(self.editor.GetObjectIDByName(obj)) for obj in object_list]
        self.module_fields_reporter.CreateFieldPlot(
            [
                "NAME:" + quantity,
                "SolutionName:=", "Setup1 : LastAdaptive",
                "UserSpecifyName:=", 0,
                "UserSpecifyFolder:=", 0,
                "QuantityName:=", quantity,
                "PlotFolder:="	, folder,
                "StreamlinePlot:="	, False,
                "AdjacentSidePlot:="	, False,
                "FullModelPlot:="	, False,
                "IntrinsicVar:="	, "Freq=\'{}kHz\' Phase=\'0deg\'".format(adapt_freq),
                "PlotGeomInfo:="	, [1, "Surface", "FacesList", len(object_list)] + ids_list,
                "FilterBoxes:="		, [0],
                [
                    "NAME:PlotOnSurfaceSettings",
                    "Filled:="		, False,
                    "IsoValType:="		, "Fringe",
                    "SmoothShade:="		, True,
                    "AddGrid:="		, False,
                    "MapTransparency:="	, True,
                    "Refinement:="		, 0,
                    "Transparency:="	, 0,
                    "SmoothingLevel:="	, 0,
                    [
                        "NAME:Arrow3DSpacingSettings",
                        "ArrowUniform:="	, True,
                        "ArrowSpacing:="	, 0,
                        "MinArrowSpacing:="	, 0,
                        "MaxArrowSpacing:="	, 0
                    ],
                    "GridColor:="		, [255, 255, 255]
                ],
                "EnableGaussianSmoothing:=", False
            ], "Field")

    def define_windings_click(self, _sender, _args):
        """
        When user clicks on the button to define winding side. Open UI for user input
        :param _sender:
        :param _args:
        :return:
        """
        # initialize classes each time since forms are unstable and may cause:
        # System.InvalidOperationException: Collection was modified; enumeration operation may not execute.
        if self.defined_layers_dict:
            self.windings_def_form = WindingForm(defined_layers_dict=self.defined_layers_dict)
        else:
            num_layers = transformer_definition["winding_definition"]["number_of_layers"]
            self.windings_def_form = WindingForm(number_undefined_layers=int(num_layers))

        self.windings_def_form.number_of_sides = self.transformer_sides.Value

        self.windings_def_form.refresh_ui_on_show()
        self.windings_def_form.ShowDialog()

        if self.windings_def_form.defined_layers_dict != self.defined_layers_dict:
            # user made changes
            self.defined_layers_dict = copy.deepcopy(self.windings_def_form.defined_layers_dict)
            self.defined_connections_dict = {}
            self.step3.UserInterface.GetComponent("analyze_button").SetEnabledFlag(
                "analyze_button", False
            )
            self.step3.UserInterface.GetComponent("setup_analysis_button").SetEnabledFlag(
                "setup_analysis_button", False
            )

        if self.defined_layers_dict:
            self.step3.UserInterface.GetComponent("define_connections_button").SetEnabledFlag(
                "define_connections_button", True
            )

    def define_connection_click(self, _sender, _args):
        """
        When click on Define connections button we prepopulate UI with data and show it
        :param _sender: unused
        :param _args: unused
        :return:
        """

        self.connection_def_form.connections_dict = self.defined_connections_dict
        if not self.defined_connections_dict:
            self.connection_def_form.winding_def_dict = self.defined_layers_dict

        self.connection_def_form.ShowDialog()
        self.defined_connections_dict = copy.deepcopy(self.connection_def_form.connections_dict)

        if self.defined_connections_dict:
            self.step3.UserInterface.GetComponent("analyze_button").SetEnabledFlag("analyze_button", True)
            self.step3.UserInterface.GetComponent("setup_analysis_button").SetEnabledFlag("setup_analysis_button", True)

    def assign_winding_excitations(self, layer_sections_list):
        """
        Create and assign coil terminals to windings
        :param layer_sections_list:
        :return:
        """

        # create winding for each layer for further assignment in circuit
        for i in range(self.number_of_layers.Value):
            self.create_winding("Layer_" + str(i+1), "External")

        for section in layer_sections_list:
            layer = section.split("_")[0]
            layer_num = layer[5:]

            point_terminal = True  # if int(side) > 1 else False  # todo does it matter in external circuit?
            self.module_boundary_setup.AssignCoilTerminal(
                [
                    "NAME:" + section,
                    "Objects:=", [section],
                    "ParentBndID:=", "Layer_" + layer_num,
                    "Conductor number:=", "1",
                    "Winding:="	, "Layer_" + layer_num,
                    "Point out of terminal:=", point_terminal
                ])

    def create_winding(self, name, winding_type, current=0.0, resistance=0.0, inductance=0.0, voltage=0.0):
        """
        Creates new winding
        :param name: name of the winding
        :param winding_type: type (Current|Voltage|External)
        :param current: current value in A
        :param resistance:  resistance in Ohm, only for Voltage type
        :param inductance: in uH, only for Voltage type
        :param voltage: voltage, in V
        :return:
        """
        self.module_boundary_setup.AssignWindingGroup(
            [
                "NAME:" + name,
                "Type:=", winding_type,
                "IsSolid:=", True,
                "Current:=", str(current) + "A",
                "Resistance:="	, str(resistance) + "ohm",
                "Inductance:="		, str(inductance) + "uH",
                "Voltage:="		, str(voltage) + "V",
                "ParallelBranchesNum:="	, "1",
                "Phase:="		, "0deg"
            ])

    def assign_matrix_winding(self):
        """Function to assign RL matrix to a winding group"""
        matrix_entry = ["NAME:MatrixEntry"]
        for i in range(1, self.number_of_layers.Value + 1):
            matrix_entry.append([
                 "NAME:MatrixEntry",
                 "Source:=", "Layer_" + str(i),
                 "NumberOfTurns:=", "1"
             ])

        self.module_parameter_setup.AssignMatrix(
            ["NAME:Matrix1",
             matrix_entry,
             ["NAME:MatrixGroup"]
             ])

        self.reduce_matrix()

    def reduce_matrix(self):
        """
        Run matrix reduction algorithm
        :return:
        """
        for i in range(self.number_of_layers.Value):
            # this loop is only required in 2021R1 and below. There is a bug in matrix reduction IDs, so we reserve some
            self.module_parameter_setup.AssignForce(
                [
                    "NAME:Force1",
                    "Reference CS:=", "Global",
                    "Is Virtual:=", True,
                    "Objects:="	, ["Layer1_1"]
                ])
            self.module_parameter_setup.DeleteParameters(["Force1"])

        def reduce(target_dict):
            """
            Run matrix reduction, join Serial/Parallel
            :param target_dict:
            :return:
            """
            reduction_list = []
            for key, val in target_dict.items():
                if isinstance(val, dict):
                    new_red_str = reduce(val)
                    reduction_type = "Series" if "S" in key[:1] else "Parallel"
                    name = key.split("_", maxsplit=1)[1] if "Side" in key else key

                    self.module_parameter_setup.AddReduceOp("Matrix1", "ReduceMatrix1",
                                                            [
                                                                "NAME:" + name,
                                                                "Type:=", "Join in " + reduction_type,
                                                                "Sources:=", new_red_str
                                                            ])
                    reduction_list.append(key)
                else:
                    reduction_list.append("Layer_" + key)

            reduction_str = ",".join(reduction_list)
            return reduction_str

        def rename(side_num, side_definition):
            """
            rename winding and circuit element to be Side_XXX for better UX
            :param side_num: transformer side number
            :param side_definition: dict with single layer in it
            :return:
            """
            layer = "Layer_" + list(side_definition.keys())[0]
            self.design.ChangeProperty(
                [
                    "NAME:AllTabs",
                    [
                        "NAME:Maxwell3D",
                        [
                            "NAME:PropServers",
                            "BoundarySetup:" + layer
                        ],
                        [
                            "NAME:ChangedProps",
                            [
                                "NAME:Name",
                                "Value:=", "Side_" + side_num
                            ]
                        ]
                    ]
                ])

            # also rename component in circuit
            comp = self.circuit.get_comp_by_name(layer)[0]
            self.circuit.change_prop(comp, "name", "Side_" + side_num)

        connections = copy.deepcopy(self.defined_connections_dict)
        for side_num, side_def in connections.items():
            if not any(isinstance(val, dict) for val in side_def.values()):
                rename(side_num, side_def)
            else:
                # replace key and append name of the side for main key, for better UX
                main_connection = list(side_def.keys())[0]
                side_def[main_connection + "_Side_" + side_num] = side_def.pop(main_connection)

                reduce(side_def)

    def assign_mesh(self, layers_list, core_list):
        """Function to create 'manual' skin layers from face sheets and to assign mesh operations to
        core geometry and windings in which we do not require skin layers (thinner than 3 skin depths)"""

        dimension_list = []
        for i in range(1, 9):
            dimension_list.append(float(transformer_definition["core_dimensions"]["D_" + str(i)]))

        mesh_op_sz = max(dimension_list)/20.0
        self.assign_length_op(core_list, mesh_op_sz, "core")
        self.assign_length_op(layers_list, mesh_op_sz/2, "layers")

    def excitation_strategy_change(self):
        if self.excitation_strategy.Value == "Voltage":
            self.voltage.Caption = "Voltage [V]"
        else:
            self.voltage.Caption = "Current [A]"

    def create_loss_report(self):
        self.module_report.CreateReport("Core and Solid Loss", "EddyCurrent", "Data Table", "Setup1 : LastAdaptive", [],
                                        ["Freq:=", ["All"]],
                                        [
                                          "X Component:=", "Freq",
                                          "Y Component:=", ["CoreLoss", "SolidLoss"]
                                        ])

    def calculate_leakage(self):
        """Create equations to calculate leakage inductance.
        For N winding transformer would be N*(N-1)/2 equations"""

        list_x = list(range(1, self.transformer_sides.Value + 1))[:]
        list_y = list(range(1, self.transformer_sides.Value + 1))[:]

        all_leakages = {}

        for x in list_x:
            for y in list_y:
                if x != y:
                    coupling_coef = "abs(L(Side_{0},Side_{1}))/sqrt(L(Side_{0},Side_{0})*L(Side_{1},Side_{1}))".format(
                                                                                                                x, y)
                    equation = "L(Side_{0},Side_{0})*{1}".format(x, coupling_coef)
                    all_leakages["Leakage_Inductance_{}{}".format(x, y)] = equation

            list_y.remove(x)

        if self.transformer_sides.Value <= 1:
            all_leakages["Leakage_Inductance_11"] = "L(Side_1,Side_1)"

        self.module_report.CreateReport("Leakage Inductance", "EddyCurrent", "Data Table", "Setup1 : LastAdaptive",
                                        [
                                            "Context:=", "Matrix1",
                                            "PointCount:="	, 1001,
                                            "Matrix:="		, "ReduceMatrix1"
                                        ],
                                        [
                                            "Freq:="		, ["All"]
                                        ],
                                        [
                                            "X Component:="		, "Freq",
                                            "Y Component:="		, list(all_leakages.values())
                                        ])

        for key, val in all_leakages.items():
            self.module_report.ChangeProperty(
                [
                    "NAME:AllTabs",
                    [
                        "NAME:Trace",
                        [
                            "NAME:PropServers",
                            "Leakage Inductance:" + val
                        ],
                        [
                            "NAME:ChangedProps",
                            [
                                "NAME:Specify Name",
                                "Value:="	, True
                            ]
                        ]
                    ]
                ])
            self.module_report.ChangeProperty(
                [
                    "NAME:AllTabs",
                    [
                        "NAME:Trace",
                        [
                            "NAME:PropServers",
                            "Leakage Inductance:" + val
                        ],
                        [
                            "NAME:ChangedProps",
                            [
                                "NAME:Name",
                                "Value:=", key
                            ]
                        ]
                    ]
                ])

            # change format to scientific
            self.module_report.ChangeProperty(
                [
                    "NAME:AllTabs",
                    [
                        "NAME:Data Filter",
                        [
                            "NAME:PropServers",
                            "Leakage Inductance:{}:Curve1".format(key)
                        ],
                        [
                            "NAME:ChangedProps",
                            [
                                "NAME:Number Format",
                                "Value:="	, "Scientific"
                            ]
                        ]
                    ]
                ])

    def enable_thermal(self, solids):
        """function to turn on feedback for thermal coupling"""
        for i in range(len(solids)):
            solids.insert(i * 2 + 1, "22cel")

        self.design.SetObjectTemperature(
            [
                "NAME:TemperatureSettings",
                "IncludeTemperatureDependence:=", True,
                "EnableFeedback:=", True,
                "Temperatures:=", solids
            ])

    def split_geom(self, list_to_split):
        self.editor.Split(
            [
                "NAME:Selections",
                "Selections:="	, ",".join(list_to_split),
                "NewPartsModelFlag:="	, "Model"
            ],
            [
                "NAME:SplitToParameters",
                "SplitPlane:="		, "YZ",
                "WhichSide:="		, "NegativeOnly",
                "ToolType:="		, "PlaneTool",
                "ToolEntityID:="	, -1,
                "SplitCrossingObjectsOnly:=", False,
                "DeleteInvalidObjects:=", True
            ])

    def create_setup(self):
        """function which grabs parameters from UI and inserts Setup1 according to them"""
        max_num_passes = int(self.number_passes.Value)
        percent_error = float(self.percentage_error.Value)

        adapt_freq = self.adaptive_frequency.Value
        frequency = str(adapt_freq) + 'kHz'

        start_sweep_freq = (str(self.start_frequency.Value) +
                            str(self.start_frequency_unit.Value))

        stop_sweep_freq = (str(self.stop_frequency.Value) +
                           str(self.stop_frequency_unit.Value))

        samples = int(self.samples.Value)

        if self.frequency_sweep.Value:
            if self.scale.Value == 'Linear':
                self.insert_setup(max_num_passes, percent_error, frequency, True,
                                  'LinearCount', start_sweep_freq, stop_sweep_freq, samples)
            else:
                self.insert_setup(max_num_passes, percent_error, frequency, True,
                                  'LogScale', start_sweep_freq, stop_sweep_freq, samples)
        else:
            self.insert_setup(max_num_passes, percent_error, frequency, False)

        return adapt_freq

    def assign_length_op(self, objects, size, name):
        """
        Assign in selection length mesh operation
        :param objects: selection
        :param size: element size
        :param name: name of the mesh
        :return:
        """
        self.module_mesh.AssignLengthOp(
            [
                "NAME:Length_" + name,
                "RefineInside:=", True,
                "Enabled:=", True,
                "Objects:=", objects,
                "RestrictElem:=", False,
                "NumMaxElem:=", "1000",
                "RestrictLength:=", True,
                "MaxLength:=", str(size) + "mm"
            ])

    def create_object_from_face(self, objects, listID):
        self.editor.CreateObjectFromFaces(
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

    def create_region(self, x_zero_region):
        """Create vacuum region with offset specified by user"""
        offset = int(self.step3.Properties["define_setup/offset"].Value)

        x_offset = 0 if x_zero_region else offset

        self.editor.CreateRegion(
            [
                "NAME:RegionParameters",
                "+XPaddingType:=", "Percentage Offset",
                "+XPadding:=", x_offset,
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

    def change_color(self, selection, rgb=(255, 255, 255)):
        """
        Change color of geometry object
        :param selection: (list) object selection
        :param rgb: (tuple) rgb color code
        :return:
        """
        self.editor.ChangeProperty(
            ["NAME:AllTabs", ["NAME:Geometry3DAttributeTab",
                              ["NAME:PropServers"] + selection,
                              ["NAME:ChangedProps",
                               [
                                   "NAME:Color",
                                   "R:=", rgb[0],
                                   "G:=", rgb[1],
                                   "B:=", rgb[2]
                               ]
                               ]]])

    def create_terminal_sections(self, layer_list, layer_sections_list, layers_sections_delete_list):
        """
        Create terminals to assign Coil Terminals
        :param layer_list: list of layers
        :param layer_sections_list: list of 2D sheets
        :param layers_sections_delete_list: list of 2D sheets duplicates to remove
        :return:
        """
        # only for EFD core not to get an error due to section of winding
        if "CentralLegCS" in self.editor.GetCoordinateSystems():
            self.editor.SetWCS(
                [
                    "NAME:SetWCS Parameter",
                    "Working Coordinate System:=", "CentralLegCS",
                    "RegionDepCSOk:=", False
                ])

        self.editor.Section(
            [
                "NAME:Selections",
                "Selections:=", ','.join(layer_list),
                "NewPartsModelFlag:=", "Model"
            ],
            [
                "NAME:SectionToParameters",
                "CreateNewObjects:=", True,
                "SectionPlane:=", "ZX",
                "SectionCrossObject:=", False
            ])

        self.editor.SeparateBody(
            [
                "NAME:Selections",
                "Selections:=", ','.join(layer_sections_list),
                "NewPartsModelFlag:=", "Model"
            ])

        self.editor.Delete(["NAME:Selections", "Selections:=", ','.join(layers_sections_delete_list)])

    def create_new_materials(self, coil_material):
        """
        Create new materials in AEDT if not yet exists
        :param coil_material:
        :return:
        """
        oDefinitionManager = self.project.GetDefinitionManager()
        core_material = self.core_material.Value
        # check if material not yet defined
        if not oDefinitionManager.DoesMaterialExist("Material_" + core_material):
            cord_list = ["NAME:Coordinates"]

            permeability = self.materials[core_material]["mu(freq)"]
            for coordinate_pair in permeability:
                cord_list.append(["NAME:Coordinate",
                                  "X:=", coordinate_pair[0],
                                  "Y:=", coordinate_pair[1]])

            self.project.AddDataset(["NAME:$Mu_" + core_material, cord_list])

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
                    "conductivity:=", self.materials[core_material]["conductivity"],
                    [
                        "NAME:core_loss_type",
                        "property_type:=", "ChoiceProperty",
                        "Choice:=", "Power Ferrite"
                    ],
                    "core_loss_cm:=", self.materials[core_material]["cm"],
                    "core_loss_x:=", self.materials[core_material]["x"],
                    "core_loss_y:=", self.materials[core_material]["y"],
                    "core_loss_kdc:=", "0",
                    "thermal_conductivity:=", "5",
                    "mass_density:=", self.materials[core_material]["density"],
                    "specific_heat:=", "750",
                    "thermal_expansion_coeffcient:=", "1e-05"
                ])

        # check if winding material exists
        if coil_material == "Copper_temperature" and not oDefinitionManager.DoesMaterialExist(coil_material):
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

    def assign_material(self, selection, material):
        self.editor.AssignMaterial(
            [
                "NAME:Selections",
                "Selections:=", selection
            ],
            [
                "NAME:Attributes",
                "MaterialValue:=", material,
                "SolveInside:=", True
            ])

    def insert_setup(self, max_num_passes, percent_error, frequency, has_sweep,
                     sweep_type='', start_sweep_freq='', stop_sweep_freq='', samples=''):
        """Function to create an analysis setup"""

        self.module_analysis.InsertSetup("EddyCurrent",
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

    def check_sides(self, manual=False):
        """
        if number of sides was changed just clear the data in winding dialogue and disable buttons
        :param manual: if want to call function from the class itself when populate UI data
        """
        if not manual:
            self.defined_layers_dict = {}
            self.step3.UserInterface.GetComponent("define_connections_button").SetEnabledFlag(
                "define_connections_button", False
            )
            self.step3.UserInterface.GetComponent("analyze_button").SetEnabledFlag(
                "analyze_button", False
            )
            self.step3.UserInterface.GetComponent("setup_analysis_button").SetEnabledFlag(
                "setup_analysis_button", False
            )

        if self.transformer_sides.Value < 1:
            self.transformer_sides.Value = 1
            return add_warning_message("Number of transformer sides cannot be less than 1")

        elif self.transformer_sides.Value > self.number_of_layers.Value:
            self.transformer_sides.Value = self.number_of_layers.Value
            return add_warning_message("Number of transformer sides cannot be less than number of layers")

        if self.transformer_sides.Value == 1:
            self.resistance.Visible = False
        else:
            self.resistance.Visible = True

        update_ui(self.step3)

    def analyze_click(self, _sender, _args):
        if not self.analysis_set:
            self.setup_analysis_click()

        try:
            self.design.Analyze("Setup1")
        except:
            pass

    def setup_analysis_click(self, _sender="", _args=""):
        """
        Run the process of model creation and setup
        :param _sender:
        :param _args:
        :return:
        """
        self.collect_ui_data_step3()
        self.create_model()
        coil_material = self.coil_material.Value + "_temperature"

        layers_list = []
        layer_sections_list = []
        layers_sections_delete_list = []
        core_list = []

        obj_list = self.editor.GetObjectsInGroup('Solids')  # get solids to avoid terminals and skin layers
        for each_obj in obj_list:
            if "Layer" in each_obj:
                layers_list.append(each_obj)
                layer_sections_list.append(each_obj + "_Section1")
                layers_sections_delete_list.append(each_obj + "_Section1_Separate1")
            elif "Core" in each_obj:
                core_list.append(each_obj)

        self.create_new_materials(coil_material)
        self.assign_material(','.join(core_list), '"Material_' + self.core_material.Value + '"')
        self.assign_material(','.join(layers_list), '"' + coil_material + '"')

        if coil_material == 'Copper_temperature':
            self.change_color(layers_list, rgb=(255, 128, 64))
        else:
            self.change_color(layers_list, rgb=(132, 135, 137))

        if self.include_bobbin.Value:
            if self.layer_type.Value == 'Wound':
                self.assign_material('Bobbin', '"polyamide"')
            else:
                boards = ','.join(self.editor.GetMatchedObjectName("Board*"))
                self.assign_material(boards, '"polyamide"')

        adapt_freq = self.create_setup()
        self.enable_thermal(layers_list[:])  # send a copy of the list

        self.create_terminal_sections(layers_list, layer_sections_list, layers_sections_delete_list)
        self.assign_winding_excitations(layer_sections_list)
        self.create_circuit()
        self.assign_matrix_winding()
        self.import_circuit()
        self.calculate_leakage()
        self.create_loss_report()

        self.module_boundary_setup.SetCoreLoss(core_list, False)

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
        self.module_boundary_setup.SetEddyEffect(["NAME:Eddy Effect Setting", eddy_list])

        self.assign_mesh(layers_list, core_list)

        sheet_objects = self.editor.GetMatchedObjectName("Layer*")
        layers_and_skins = []
        terminals = []
        for sheet in sheet_objects:
            if "Section" not in sheet:
                layers_and_skins.append(sheet)
            else:
                terminals.append(sheet)

        # if user selects to enable bobbin on boards we need to cut them, if not empty list
        bobbin_list = self.editor.GetMatchedObjectName("Bobbin")
        boards_list = self.editor.GetMatchedObjectName("Board*")

        x_zero_region = False
        if not self.full_model.Value:
            self.split_geom(core_list + layers_and_skins + bobbin_list + boards_list)

            # if terminals were created on the wrong side mirror them to be inside of the conductor
            first_vertex_id_first_terminal = int(self.editor.GetVertexIDsFromObject(terminals[0])[0])
            x_coord = float(self.editor.GetVertexPosition(first_vertex_id_first_terminal)[0])
            if x_coord >= 0:
                self.mirror(",".join(terminals))
            x_zero_region = True
            self.set_symmetry_multiplier()

        self.create_region(x_zero_region)

        self.create_field_plot("Mag_J", "J", adapt_freq, layers_list)
        self.create_field_plot("Ohmic_Loss", "Ohmic-Loss", adapt_freq, layers_list)
        self.create_field_plot("Mag_B", "B", adapt_freq, core_list)
        self.create_field_plot("Core_Loss", "Core-Loss", adapt_freq, core_list)

        self.project.SetActiveDesign(self.design_name)
        self.editor.FitAll()

        self.write_json_data()

        self.project.SaveAs(os.path.join(self.project_path.Value, self.design_name + '.aedt'), True)
        self.analysis_set = True
        self.step3.UserInterface.GetComponent("setup_analysis_button").SetEnabledFlag("setup_analysis_button", False)
        self.step3.UserInterface.GetComponent("define_windings_button").SetEnabledFlag("define_windings_button", False)
        self.step3.UserInterface.GetComponent("define_connections_button").SetEnabledFlag(
            "define_connections_button", False
        )
        oDesktop.EnableAutoSave(self.flag_auto_save)

    def create_model(self):
        """
        Generate Transformer geometry
        :return:
        """
        time_now = datetime.datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
        self.design_name = 'Transformer_' + time_now

        self.project.InsertDesign("Maxwell 3D", self.design_name, "EddyCurrent", "")
        self.design = self.project.SetActiveDesign(self.design_name)
        self.editor = self.design.SetActiveEditor("3D Modeler")

        # define all modules
        self.module_analysis = self.design.GetModule("AnalysisSetup")
        self.module_boundary_setup = self.design.GetModule("BoundarySetup")
        self.module_parameter_setup = self.design.GetModule("MaxwellParameterSetup")
        self.module_output_var = self.design.GetModule("OutputVariable")
        self.module_report = self.design.GetModule("ReportSetup")
        self.module_mesh = self.design.GetModule("MeshSetup")
        self.module_fields_reporter = self.design.GetModule("FieldsReporter")

        args = [transformer_definition, self.project, self.design, self.editor]
        all_cores = {'E': ECore, 'EI': EICore, 'U': UCore, 'UI': UICore,
                     'PQ': PQCore, 'ETD': ETDCore, 'EQ': ETDCore,
                     'EC': ETDCore, 'RM': RMCore, 'EP': EPCore,
                     'EFD': EFDCore, 'ER': ETDCore, 'P': PCore,
                     'PT': PCore, 'PH': PCore}

        cores_arguments = ['EC', 'ETD', 'EQ', 'ER', 'P', 'PT', 'PH']

        self.flag_auto_save = oDesktop.GetAutoSaveEnabled()
        oDesktop.EnableAutoSave(False)

        draw_class = all_cores[self.core_type.Value](args)
        if self.core_type.Value in cores_arguments:
            draw_class.draw_geometry(self.core_type.Value)
        else:
            draw_class.draw_geometry()

    def check_winding(self):
        """
        Validate that windings will fit into the transformer dimensions
        :return:
        """
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
                xml_path_to_table = 'winding_properties/conductor_type/table_layers'
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

            else:
                # conductor type: Circular
                xml_path_to_table = 'winding_properties/conductor_type/table_layers_circles'
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
                xml_path_to_table = 'winding_properties/conductor_type/table_layers'
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
                xml_path_to_table = 'winding_properties/conductor_type/table_layers_circles'
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

        else:
            # layer type: Planar
            xml_path_to_table = 'winding_properties/conductor_type/table_layers'
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

    def set_symmetry_multiplier(self):
        """
        Set symmetry multiplier equal to 2 due to cut of the geometry
        :return:
        """
        self.design.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:Maxwell3D",
                    [
                        "NAME:PropServers",
                        "Design Settings"
                    ],
                    [
                        "NAME:ChangedProps",
                        [
                            "NAME:Symmetry/Multiplier",
                            "Value:="	, "2"
                        ]
                    ]
                ]
            ])

    def create_circuit(self):
        """
        Create Maxwell Circuit
        :return:
        """
        if self.excitation_strategy.Value == "Voltage":
            current = 0
            voltage = self.voltage.Value
        else:
            voltage = 0
            current = self.voltage.Value

        self.circuit = Circuit(self.defined_connections_dict, self.project, self.design_name,
                               current=current, voltage=voltage,
                               resistance=self.resistance.Value, frequency=self.adaptive_frequency.Value)
        self.circuit.create()
        self.project.SetActiveDesign(self.design_name)

    def import_circuit(self):
        """
        Export circuit from Maxwell Circuit design and import into Maxwell Eddy Current
        :return:
        """
        circuit_path = os.path.join(self.project.GetPath(), self.design_name + "_circuit.sph")
        self.circuit.design.ExportNetlist("", circuit_path)

        self.module_boundary_setup.EditExternalCircuit(circuit_path, [], [], [], [])  # import circuit


def on_init_step1(step):
    """invoke on step initialisation, only once when you open the app"""
    global transformer

    # clean transformer definition since ACTs keep global variables till restart of AEDT
    for key in transformer_definition:
        transformer_definition.pop(key, None)

    transformer = TransformerClass(step)
    transformer.initialize_step1()


def create_buttons_step1(_step):
    """invoke on step1 refresh, Every time when step layout is opened."""
    transformer.refresh_step1()


def callback_step1(_step):
    """invoke on click 'Next' button in UI"""
    transformer.callback_step1()


def on_supplier_change(_step, _prop):
    """
    When supplier value is changed
    :param _step: unused
    :param _prop: unused
    :return:
    """
    transformer.prefill_core_types()


def on_core_type_change(_step, _prop):
    """
    When core type value is changed
    :param _step: unused
    :param _prop: unused
    :return:
    """
    transformer.prefill_core_models()


def on_core_model_change(_step, _prop):
    """
    When core model value is changed
    :param _step: unused
    :param _prop: unused
    :return:
    """
    transformer.prefill_core_dimensions()


def on_init_step2(_step):
    """invoke on step initialisation, only once when you open the app"""
    transformer.initialize_step2()


def on_reset_step2(_step):
    """Called when Back button on Step3 (not typo) is clicked"""
    transformer.reset_step2()


def create_buttons_step2(_step):
    """invoke on step2 refresh, Every time when step layout is opened."""
    transformer.refresh_step2()


def callback_step2(_step):
    """invoke on click 'Next' button in UI"""
    transformer.callback_step2()


def on_layers_number_change(_step, _prop):
    """
    When number of layers changed
    :param _step: unused
    :param _prop: unused
    :return:
    """
    transformer.update_rows()
    transformer.warn_about_winding_def_clean_up()


def on_layer_type_change(_step, _prop):
    """
    When layer type changed (Wound/Planar)
    :param _step: unused
    :param _prop: unused
    :return:
    """
    transformer.change_captions()


def on_excitation_strategy_change(_step, _prop):
    """
    When change from Voltage to Current or other way
    :param _step: unused
    :param _prop: unused
    :return:
    """
    transformer.excitation_strategy_change()


def on_init_step3(_step):
    """invoke on step initialisation, only once when you open the app"""
    transformer.initialize_step3()


def create_buttons_step3(_step):
    """invoke on step3 refresh, Every time when step layout is opened."""
    transformer.refresh_step3()


def on_step_back(_step):
    """Called when Back button is clicked"""
    pass
