import json
import os
import shutil
import unittest
from AEDTLib.Desktop import Desktop
from AEDTLib.Maxwell import Maxwell3D

import src.ElectronicTransformer.etk_callback as etk

AEDT_VERSION = "2021.2"


class BaseAEDT(unittest.TestCase):
    report_path = None
    transformer = None
    desktop = None
    project = None
    tests_dir = None
    root_dir = None
    m3d = None
    input_file = None

    @classmethod
    def setUpClass(cls):
        cls.desktop = Desktop(AEDT_VERSION)
        cls.project = cls.desktop._main.oDesktop.NewProject()
        cls.tests_dir = os.path.abspath(os.path.dirname(__file__))
        cls.root_dir = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
        cls.report_path = os.path.join(cls.tests_dir, "report.tab")
        etk.oDesktop = cls.desktop._main.oDesktop

        with open(os.path.join(cls.root_dir, cls.input_file)) as input_file:
            etk.transformer_definition = json.load(input_file)
        etk.transformer_definition["setup_definition"]["project_path"] = cls.tests_dir

        cls.transformer = etk.TransformerClass(None)
        cls.transformer.read_material_data()
        cls.transformer.project = cls.project
        cls.transformer.setup_analysis()
        cls.transformer.design.Analyze("Setup1")
        cls.m3d = Maxwell3D()

    @classmethod
    def tearDownClass(cls):
        cls.desktop.release_desktop(close_projects=True)

        project_file = os.path.join(cls.tests_dir, cls.transformer.design_name + '.aedt')
        if os.path.isfile(project_file):
            os.remove(project_file)

        project_results = project_file + "results"
        if os.path.isdir(project_results):
            shutil.rmtree(project_results)

        json_file = os.path.join(cls.tests_dir, cls.transformer.design_name + '_parameters.json')
        if os.path.isfile(json_file):
            os.remove(json_file)

        circuit_file = os.path.join(cls.tests_dir, cls.transformer.design_name + '_circuit.sph')
        if os.path.isfile(circuit_file):
            os.remove(circuit_file)

        if os.path.isfile(cls.report_path):
            os.remove(cls.report_path)

    def vertex_from_edge_coord(self, coord, name, sort_key=1):
        edge = self.m3d.modeler.primitives.get_edgeid_from_position(coord, obj_name=name)
        vertices = self.m3d.modeler.primitives.get_edge_vertices(edge)
        vertex_coord = [self.m3d.modeler.primitives.get_vertex_position(vertex) for vertex in vertices]
        vertex_coord.sort(key=lambda x: x[sort_key])
        return vertex_coord

    def set_freq_units(self, plot):
        """
        Set primary frequency units for report to kHz
        """
        self.transformer.module_report.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:Data Filter",
                    [
                        "NAME:PropServers",
                        plot + ":PrimarySweepDrawing"
                    ],
                    [
                        "NAME:ChangedProps",
                        [
                            "NAME:Number Format",
                            "Value:="	, "Scientific"
                        ],
                        [
                            "NAME:Units",
                            "Value:="		, "kHz"
                        ]
                    ]
                ]
            ])

    def compare_leakage(self, ref_name):
        """
        exports leakage report and compares all values for tolerance of 2%
        Args:
            ref_name: file of reference output under "reference_results" folder

        Returns: None
        """
        self.set_freq_units("Leakage Inductance")
        self.transformer.module_report.ExportToFile("Leakage Inductance", self.report_path, False)

        reference_path = os.path.join(self.tests_dir, "reference_results", ref_name)
        with open(reference_path) as ref_file, open(self.report_path) as actual_file:
            next(ref_file)
            next(actual_file)

            for line1, line2 in zip(ref_file, actual_file):
                ref_result = [float(val) for val in line1.split()]
                actual_result = [float(val) for val in line2.split()]

                for actual, ref in zip(actual_result, ref_result):
                    self.assertAlmostEqual(actual, ref, delta=ref*0.02,
                                           msg="Error at frequency {}kHz. Report is {}".format(ref_result[0],
                                                                                               actual_result))

    def compare_loss(self, loss_type, reference_list):
        """
        Compare Core/Solid loss data with reference
        Args:
            loss_type (str): SolidLoss/CoreLoss
            reference_list: list with reference results

        Returns: None
        """
        loss_data = self.m3d.post.get_report_data(expression=loss_type)
        loss_list = loss_data.data_magnitude(convert_to_SI=True)

        for actual, ref in zip(loss_list, reference_list):
            self.assertAlmostEqual(actual, ref, delta=ref*0.02)

    def compare_json(self):
        """
        Compare that generated JSON file is the same as example file
        Returns:None
        """
        json_file = os.path.join(self.tests_dir, self.transformer.design_name + '_parameters.json')
        with open(json_file) as actual, open(os.path.join(self.root_dir, self.input_file)) as ref:
            actual_data = json.load(actual)
            ref_data = json.load(ref)

        # project path might be different
        self.assertTrue(ref_data["setup_definition"].pop("project_path"))
        self.assertTrue(actual_data["setup_definition"].pop("project_path"))

        self.assertDictEqual(actual_data, ref_data)

    def assertListAlmostEqual(self, list1, list2, msg=""):
        for elem1, elem2 in zip(list1, list2):
            self.assertAlmostEqual(elem1, elem2, msg=msg, delta=abs(elem2*0.02))


class TestIEEE(BaseAEDT):
    @classmethod
    def setUpClass(cls):
        cls.input_file = r"src/ElectronicTransformer/examples/demo_IEEE.json"
        super(TestIEEE, cls).setUpClass()

    def test_01_layer_turns(self):
        """
        Check that only 7 turns are created on layer 8
        """
        # layers + terminals
        layers_list = self.m3d.modeler.get_matched_object_name("Layer8*")
        self.assertEqual(len(layers_list), 14)

        # terminals
        layers_list = self.m3d.modeler.get_matched_object_name("Layer8*Section*")
        self.assertEqual(len(layers_list), 7)

    def test_02_conductor_dimensions(self):
        """
        Test conductor cross section
        """
        # width
        vertex_coord = self.vertex_from_edge_coord((0.0, 5.0, -0.27), "Layer3_3")
        self.assertListEqual(vertex_coord[0], [0.0, 4.87, -0.27])
        self.assertListEqual(vertex_coord[1], [0.0, 5.68, -0.27])

        # height
        vertex_coord = self.vertex_from_edge_coord((0, 5.68, -0.3), "Layer3_3", sort_key=2)
        self.assertListEqual(vertex_coord[0], [0.0, 5.68, -0.34])
        self.assertListEqual(vertex_coord[1], [0.0, 5.68, -0.27])

    def test_03_board_dimensions(self):
        """
        Test board XYZ dimensions
        """
        # length X
        vertex_coord = self.vertex_from_edge_coord((-2, 6.5, 1.01), "Board_8", sort_key=0)
        self.assertListEqual(vertex_coord[0], [-5.5, 6.5, 1.01])
        self.assertListEqual(vertex_coord[1], [0.0, 6.5, 1.01])

        # height Z
        vertex_coord = self.vertex_from_edge_coord((0, 6.5, 0.9), "Board_8", sort_key=2)
        self.assertListEqual(vertex_coord[0], [0.0, 6.5, 0.81])
        self.assertListEqual(vertex_coord[1], [0.0, 6.5, 1.01])

        # width Y
        vertex_coord = self.vertex_from_edge_coord((0, 4, 1.01), "Board_8", sort_key=1)
        self.assertListEqual(vertex_coord[0], [0.0, 2.5, 1.01])
        self.assertListEqual(vertex_coord[1], [0.0, 6.5, 1.01])

    def test_04_solid_loss(self):
        """
        Validate that SolidLoss are in range of 2% compared to reference
        """
        reference_loss = [0.1035474408, 0.02305190848, 0.005264047956, 0.001412152157, 0.0005821054265, 0.0004035690909,
                          0.0003653214849, 0.0003579865491, 0.0003591455644, 0.000359511285, 0.0003648278448]
        self.compare_loss("SolidLoss", reference_loss)

    def test_05_core_loss(self):
        """
        Validate that CoreLoss are in range of 2% compared to reference
        """
        reference_loss = [0.00487814, 0.00317489, 0.00203019, 0.00129321, 0.000823085, 0.000523797,
                          0.000333353, 0.000212169, 0.000135052, 0.000129268, 8.60161e-05]
        self.compare_loss("CoreLoss", reference_loss)

    def test_06_leakage_inductance(self):
        """
        Validate that leakage is in range of 2% difference
        """
        self.compare_leakage("ieee_leakage.tab")

    def test_07_json(self):
        """
        Compare that generated JSON file is the same as example file
        """
        self.compare_json()


class TestWirewound(BaseAEDT):
    @classmethod
    def setUpClass(cls):
        cls.input_file = r"src/ElectronicTransformer/examples/demo_wirewound.json"
        super(TestWirewound, cls).setUpClass()

    def test_01_core(self):
        """
        Check core
        """
        # no airgap
        vertex_coord = self.vertex_from_edge_coord((0, 0, 0), "RM_Core_Bottom", sort_key=1)
        self.assertListAlmostEqual(vertex_coord[0], [0.0, -5.45, 0])
        self.assertListAlmostEqual(vertex_coord[1], [0.0, 5.45, 0])

        # RM shape test
        vertex_coord = self.vertex_from_edge_coord((-7.10840494, 7.10840494, -6.2), "RM_Core_Bottom")
        self.assertListAlmostEqual(vertex_coord[0], [-6.75, 6.75, -6.2])
        self.assertListAlmostEqual(vertex_coord[1], [-7.470992097, 7.511458427, -6.2])

    def test_02_layer_turns(self):
        """
        Check that only 7 turns are created on layer 3
        """
        # layers + terminals + two skin layers
        layers_list = self.m3d.modeler.get_matched_object_name("Layer3*")
        self.assertEqual(len(layers_list), 9*4)

        # terminals
        layers_list = self.m3d.modeler.get_matched_object_name("Layer3*Section*")
        self.assertEqual(len(layers_list), 9)

        # terminals
        layers_list = self.m3d.modeler.get_matched_object_name("Layer3*skin*")
        self.assertEqual(len(layers_list), 9*2)

    def test_03_conductor_dimensions(self):
        """
        Test conductor cross section
        """
        # skin layer
        vertex_coord = self.vertex_from_edge_coord((0, 9.602890868, -2.901599471), "Layer3_skin_1_skin_2_7", sort_key=1)
        self.assertListAlmostEqual(vertex_coord[0], [0, 9.5, -2.858980678])
        self.assertListAlmostEqual(vertex_coord[1], [0, 9.705781736, -2.944218264])

        # outer radius
        vertex_coord = self.vertex_from_edge_coord((0, 9.676776695, -2.723223305), "Layer3_7", sort_key=1)
        self.assertListAlmostEqual(vertex_coord[0], [0, 9.5, -2.65])
        self.assertListAlmostEqual(vertex_coord[1], [0, 9.853553391, -2.796446609])

    def test_04_bobbin_dimensions(self):
        """
        Test bobbin XYZ dimensions
        """
        # length Y
        vertex_coord = self.vertex_from_edge_coord((0.0, 8, 6.2), "Bobbin", sort_key=1)
        self.assertListAlmostEqual(vertex_coord[0], [0.0, 5.45, 6.2])
        self.assertListAlmostEqual(vertex_coord[1], [0.0, 10.6, 6.2])

        # thickness Z
        vertex_coord = self.vertex_from_edge_coord((0, 10.6, -5.5), "Bobbin", sort_key=2)
        self.assertListAlmostEqual(vertex_coord[0], [0.0, 10.6, -6.2])
        self.assertListAlmostEqual(vertex_coord[1], [0.0, 10.6, -5.2])

        # width Y
        vertex_coord = self.vertex_from_edge_coord((0, 8, -5.2), "Bobbin", sort_key=1)
        self.assertListAlmostEqual(vertex_coord[0], [0, 6.45, -5.2])
        self.assertListAlmostEqual(vertex_coord[1], [0.0, 10.6, -5.2])

    def test_05_solid_loss(self):
        """
        Validate that SolidLoss are in range of 2% compared to reference
        """
        reference_loss = [0.001791334027, 0.0011241232529999999, 0.00098726573489999999, 0.00098593841609999995,
                          0.0010940049879999999, 0.001410922602, 0.0020232734429999999]
        self.compare_loss("SolidLoss", reference_loss)

    def test_06_core_loss(self):
        """
        Validate that CoreLoss are in range of 2% compared to reference
        """
        reference_loss = [0.0014199200000000001, 0.00058693600000000003, 0.00024258099999999999, 0.000100207,
                          4.1315400000000001e-05, 1.6967199999999999e-05, 6.9389800000000003e-06]
        self.compare_loss("CoreLoss", reference_loss)

    def test_07_leakage_inductance(self):
        """
        Validate that leakage is in range of 2% difference
        """
        self.compare_leakage("wire_wound.tab")

    def test_08_json(self):
        """
        Compare that generated JSON file is the same as example file
        """
        self.compare_json()


class TestGapInfluence(BaseAEDT):
    @classmethod
    def setUpClass(cls):
        cls.input_file = r"src/ElectronicTransformer/examples/demo_gap_influence.json"
        super(TestGapInfluence, cls).setUpClass()

    def test_01_air_gap(self):
        """
        Check that air gap exists only on central leg
        """
        vertex_coord = self.vertex_from_edge_coord((0, 0, 1.2805), "RM_Core_Top")
        self.assertListEqual(vertex_coord[0], [0, -6.4, 1.2805])
        self.assertListEqual(vertex_coord[1], [0, 6.4, 1.2805])

        vertex_coord = self.vertex_from_edge_coord((0, 0, -1.2805), "RM_Core_Bottom")
        self.assertListEqual(vertex_coord[0], [0, -6.4, -1.2805])
        self.assertListEqual(vertex_coord[1], [0, 6.4, -1.2805])

        vertex_coord = self.vertex_from_edge_coord((-18.7, 0, 0), "RM_Core_Bottom")
        self.assertListAlmostEqual(vertex_coord[0], [-18.7, -2.371782079, 0])
        self.assertListAlmostEqual(vertex_coord[1], [-18.7, 2.371782079, 0])

    def test_02_solid_loss(self):
        """
        Validate that SolidLoss are in range of 2% compared to reference
        """
        reference_loss = [0.07441614777, 0.03841830668, 0.01793546889, 0.008098162492, 0.003709001461,
                          0.001719827518, 0.000783162344, 0.0004020461219]
        self.compare_loss("SolidLoss", reference_loss)

    def test_03_core_loss(self):
        """
        Validate that CoreLoss are in range of 2% compared to reference
        """
        reference_loss = [0.000774539, 0.000426522, 0.00023523, 0.000130118, 7.20314e-05,
                          3.98747e-05, 2.20806e-05, 1.37956e-05]
        self.compare_loss("CoreLoss", reference_loss)

    def test_04_json(self):
        """
        Compare that generated JSON file is the same as example file
        """
        self.compare_json()


class TestCoreEC(BaseAEDT):
    @classmethod
    def setUpClass(cls):
        cls.input_file = r"src/ElectronicTransformer/examples/demo_EC.json"
        super(TestCoreEC, cls).setUpClass()

    def test_01_layer_turns(self):
        """
        Check that all objects are created
        """

        solids_list = self.m3d.modeler.get_objects_in_group("Solids")
        self.assertListEqual(sorted(solids_list), ['Bobbin', 'EC_Core_Bottom', 'EC_Core_Top', 'Layer1', 'Layer1_1',
                                                   'Layer1_2', 'Layer1_3', 'Layer1_4', 'Layer1_5', 'Layer1_6',
                                                   'Layer1_7', 'Layer1_8', 'Layer1_9', 'Layer2', 'Layer2_1', 'Layer2_2',
                                                   'Layer2_3', 'Layer2_4', 'Layer2_5', 'Layer2_6', 'Layer2_7',
                                                   'Layer2_8', 'Layer2_9', 'Layer3', 'Layer3_1', 'Layer3_2', 'Layer3_3',
                                                   'Layer3_4', 'Layer3_5', 'Layer3_6', 'Layer3_7', 'Layer3_8',
                                                   'Layer3_9', 'Layer4', 'Layer4_1', 'Layer4_2', 'Layer4_3', 'Layer4_4',
                                                   'Layer4_5', 'Layer4_6', 'Layer4_7', 'Layer4_8', 'Layer4_9', 'Layer5',
                                                   'Layer5_1', 'Layer5_2', 'Layer5_3', 'Layer5_4', 'Layer5_5',
                                                   'Layer5_6', 'Layer5_7', 'Layer5_8', 'Layer5_9',
                                                   'Layer6', 'Layer6_1', 'Layer6_2', 'Layer6_3', 'Layer6_4', 'Layer6_5',
                                                   'Layer6_6', 'Layer6_7', 'Layer6_8', 'Layer6_9', 'Region'])

        sheets_list = self.m3d.modeler.get_objects_in_group("Sheets")
        self.assertListEqual(sorted(sheets_list), ['Layer1_1_Section1', 'Layer1_2_Section1', 'Layer1_3_Section1',
                                                   'Layer1_4_Section1', 'Layer1_5_Section1', 'Layer1_6_Section1',
                                                   'Layer1_7_Section1', 'Layer1_8_Section1', 'Layer1_9_Section1',
                                                   'Layer1_Section1', 'Layer2_1_Section1', 'Layer2_2_Section1',
                                                   'Layer2_3_Section1', 'Layer2_4_Section1', 'Layer2_5_Section1',
                                                   'Layer2_6_Section1', 'Layer2_7_Section1', 'Layer2_8_Section1',
                                                   'Layer2_9_Section1', 'Layer2_Section1', 'Layer3_1_Section1',
                                                   'Layer3_2_Section1', 'Layer3_3_Section1', 'Layer3_4_Section1',
                                                   'Layer3_5_Section1', 'Layer3_6_Section1', 'Layer3_7_Section1',
                                                   'Layer3_8_Section1', 'Layer3_9_Section1', 'Layer3_Section1',
                                                   'Layer4_1_Section1', 'Layer4_2_Section1', 'Layer4_3_Section1',
                                                   'Layer4_4_Section1', 'Layer4_5_Section1', 'Layer4_6_Section1',
                                                   'Layer4_7_Section1', 'Layer4_8_Section1', 'Layer4_9_Section1',
                                                   'Layer4_Section1', 'Layer5_1_Section1', 'Layer5_2_Section1',
                                                   'Layer5_3_Section1', 'Layer5_4_Section1', 'Layer5_5_Section1',
                                                   'Layer5_6_Section1', 'Layer5_7_Section1', 'Layer5_8_Section1',
                                                   'Layer5_9_Section1', 'Layer5_Section1', 'Layer6_1_Section1',
                                                   'Layer6_2_Section1', 'Layer6_3_Section1', 'Layer6_4_Section1',
                                                   'Layer6_5_Section1', 'Layer6_6_Section1', 'Layer6_7_Section1',
                                                   'Layer6_8_Section1', 'Layer6_9_Section1', 'Layer6_Section1'])

    def test_02_conductor_dimensions(self):
        """
        Test conductor cross section
        """
        # height
        vertex_coord = self.vertex_from_edge_coord((0, 6.7, 11.05), "Layer5", sort_key=2)
        self.assertListEqual(vertex_coord[0], [0, 6.7, 9.95])
        self.assertListEqual(vertex_coord[1], [0, 6.7, 12.15])

        # width
        vertex_coord = self.vertex_from_edge_coord((0, 6.6, 12.15), "Layer5")
        self.assertListEqual(vertex_coord[0], [0, 6.5, 12.15])
        self.assertListEqual(vertex_coord[1], [0, 6.7, 12.15])

        #
        vertex_coord = self.vertex_from_edge_coord((0, 6.6, 9.85), "Layer5_1", sort_key=2)
        self.assertListEqual(vertex_coord[0], [0, 6.5, 9.85])
        self.assertListEqual(vertex_coord[1], [0, 6.7, 9.85])

    def test_03_solid_loss(self):
        """
        Validate that SolidLoss are in range of 2% compared to reference
        """
        reference_loss = [0.48166285380000001, 0.12871921159999999, 0.051960901009999998,
                          0.034661488040000001, 0.029949885039999999, 0.028715741239999999, 0.02929752683]
        self.compare_loss("SolidLoss", reference_loss)

    def test_04_core_loss(self):
        """
        Validate that CoreLoss are in range of 2% compared to reference
        """
        reference_loss = [0.013777299999999999, 0.0069939700000000004, 0.0035386599999999999,
                          0.0017938800000000001, 0.00091125600000000005, 0.000461817, 0.000233544]
        self.compare_loss("CoreLoss", reference_loss)

    def test_05_leakage_inductance(self):
        """
        Validate that leakage is in range of 2% difference
        """
        self.compare_leakage("ec_leakage.tab")

    def test_06_json(self):
        """
        Compare that generated JSON file is the same as example file
        """
        self.compare_json()


class TestCoreU(BaseAEDT):
    @classmethod
    def setUpClass(cls):
        cls.input_file = r"src/ElectronicTransformer/examples/demo_U_skinlayers.json"
        super(TestCoreU, cls).setUpClass()

    def test_01_skins(self):
        """
        Check skin layers
        """
        # conductor
        vertex_coord = self.vertex_from_edge_coord((-5.232730116, 14.25, 10.15), "Layer2_1", sort_key=2)
        self.assertListAlmostEqual(vertex_coord[0], [-5.232730116, 14.25, 8.65])
        self.assertListAlmostEqual(vertex_coord[1], [-5.232730116, 14.25, 11.65])

        # outside skins
        vertex_coord = self.vertex_from_edge_coord((-5.230450129, 14.14550966, 10.15), "Layer2_skin_1_out_1",
                                                   sort_key=2)
        self.assertListAlmostEqual(vertex_coord[0], [-5.230450129, 14.14550966, 8.65])
        self.assertListAlmostEqual(vertex_coord[1], [-5.230450129, 14.14550966, 11.65])

        vertex_coord = self.vertex_from_edge_coord((-5.228170142, 14.04101932, 10.15), "Layer2_skin_1_skin_2_out_1",
                                                   sort_key=2)
        self.assertListAlmostEqual(vertex_coord[0], [-5.228170142, 14.04101932, 8.65])
        self.assertListAlmostEqual(vertex_coord[1], [-5.228170142, 14.04101932, 11.65])

        # inside skins
        vertex_coord = self.vertex_from_edge_coord((-5.171829858, 11.45898068, 10.15), "Layer2_skin_1_skin_2_in_1",
                                                   sort_key=2)
        self.assertListAlmostEqual(vertex_coord[0], [-5.171829858, 11.45898068, 8.65])
        self.assertListAlmostEqual(vertex_coord[1], [-5.171829858, 11.45898068, 11.65])

        vertex_coord = self.vertex_from_edge_coord((-5.169549871, 11.35449034, 10.15), "Layer2_skin_1_in_1",
                                                   sort_key=2)
        self.assertListAlmostEqual(vertex_coord[0], [-5.169549871, 11.35449034, 8.65])
        self.assertListAlmostEqual(vertex_coord[1], [-5.169549871, 11.35449034, 11.65])

    def test_02_solid_loss(self):
        """
        Validate that SolidLoss are in range of 2% compared to reference
        """
        reference_loss = [5.01460095]
        self.compare_loss("SolidLoss", reference_loss)

    def test_03_core_loss(self):
        """
        Validate that CoreLoss are in range of 2% compared to reference
        """
        reference_loss = [0.000101699]
        self.compare_loss("CoreLoss", reference_loss)

    def test_04_leakage_inductance(self):
        """
        Validate that leakage is in range of 2% difference
        """
        self.compare_leakage("u_leakage.tab")

    def test_05_json(self):
        """
        Compare that generated JSON file is the same as example file
        """
        self.compare_json()
