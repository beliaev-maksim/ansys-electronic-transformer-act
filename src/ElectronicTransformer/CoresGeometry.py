# already done: E, EI, ETD = EC = EQ = ER(with argument), EFD, EP, PQ, P=PH=PT(with    argument), RM, U, UI

# super class Cores
# all sub classes for core creation will inherit from here basic parameters from GUI and
# take all functions for primitive creation like createBox, createPolyhedron, move, rename and so on
import math


class Cores(object, Step1, Step2):
    CS = 'Global'

    def __init__(self, args_list):
        step = args_list[0]
        self.oProject = args_list[1]
        self.oDesign = args_list[2]
        self.oEditor = args_list[3]

        Step1.__init__(self, step)
        Step2.__init__(self, step)

        self.core_dims = []
        for i in range(1, 9):
            if self.core_dimensions["D_" + str(i)].Visible:
                self.core_dims.append(self.core_dimensions["D_" + str(i)].Value)
            else:
                self.core_dims.append(0)

        # Draw in Base Core
        self.dim_D1 = float(self.core_dims[0])
        self.dim_D2 = float(self.core_dims[1])
        self.dim_D3 = float(self.core_dims[2])
        self.dim_D4 = float(self.core_dims[3])
        self.dim_D5 = float(self.core_dims[4])
        self.dim_D6 = float(self.core_dims[5])
        self.dim_D7 = float(self.core_dims[6])
        self.dim_D8 = float(self.core_dims[7])

        self.segments_number = 0 if self.segmentation_angle.Value == 0 else int(360/self.segmentation_angle.Value)

        self.airgap_side = 0
        self.airgap_center = 0
        self.airgap_both = 0

        if self.define_airgap.Value:
            if self.airgap_on_leg.Value == 'Center':
                self.airgap_center = self.airgap_value.Value / 2.0
            elif self.airgap_on_leg.Value == 'Side':
                self.airgap_side = self.airgap_value.Value / 2.0
            else:
                self.airgap_both = self.airgap_value.Value / 2.0

        self.winding_parameters_dict = {}

        if self.conductor_type.Value == 'Rectangular':
            xml_path_to_table = 'winding_properties/conductor_type/table_layers'
            row_num = self.table_layers.RowCount

            for row_index in range(0, row_num):
                self.winding_parameters_dict[row_index + 1] = []
                self.winding_parameters_dict[row_index + 1].append(
                    float(self.table_layers.Value[xml_path_to_table + "/conductor_width"][row_index])
                )

                self.winding_parameters_dict[row_index + 1].append(
                    float(self.table_layers.Value[xml_path_to_table + "/conductor_height"][row_index])
                )

                self.winding_parameters_dict[row_index + 1].append(
                    int(self.table_layers.Value[xml_path_to_table + "/turns_number"][row_index])
                )

                self.winding_parameters_dict[row_index + 1].append(
                    float(self.table_layers.Value[xml_path_to_table + "/insulation_thickness"][row_index])
                )
        else:
            xml_path_to_table = 'winding_properties/conductor_type/table_layers_circles'
            row_num = self.table_layers_circles.RowCount
            for row_index in range(0, row_num):
                self.winding_parameters_dict[row_index + 1] = []
                self.winding_parameters_dict[row_index + 1].append(
                    float(self.table_layers_circles.Value[xml_path_to_table + "/conductor_diameter"][row_index])
                )
                self.winding_parameters_dict[row_index + 1].append(
                    int(self.table_layers_circles.Value[xml_path_to_table + "/segments_number"][row_index])
                )

                self.winding_parameters_dict[row_index + 1].append(
                    int(self.table_layers_circles.Value[xml_path_to_table + "/turns_number"][row_index])
                )

                self.winding_parameters_dict[row_index + 1].append(
                    float(self.table_layers_circles.Value[xml_path_to_table + "/insulation_thickness"][row_index])
                )

    def create_polyline(self, points, segments, name, covered=True, closed=True, color='(165 42 42)'):
        self.oEditor.CreatePolyline(
            [
                "NAME:PolylineParameters",
                "IsPolylineCovered:=", covered,
                "IsPolylineClosed:=", closed,
                points,
                segments,
                [
                    "NAME:PolylineXSection",
                    "XSectionType:=", "None",
                    "XSectionOrient:=", "Auto",
                    "XSectionWidth:=", "0mm",
                    "XSectionTopWidth:=", "0mm",
                    "XSectionHeight:=", "0mm",
                    "XSectionNumSegments:=", "0",
                    "XSectionBendType:=", "Corner"
                ]
            ],
            [
                "NAME:Attributes",
                "Name:=", name,
                "Flags:=", "",
                "Color:=", color,
                "Transparency:=", 0,
                "PartCoordinateSystem:=", Cores.CS,
                "UDMId:=", "",
                "MaterialValue:=", '"ferrite"',
                "SurfaceMaterialValue:=", "\"\"",
                "SolveInside:=", True,
                "IsMaterialEditable:=", True,
                "UseMaterialAppearance:=", False
            ])

    def move(self, selection, x, y, z):
        self.oEditor.Move(
            [
                "NAME:Selections",
                "Selections:=", selection,
                "NewPartsModelFlag:=", "Model"
            ],
            [
                "NAME:TranslateParameters",
                "TranslateVectorX:=", str(x) + 'mm',
                "TranslateVectorY:=", str(y) + 'mm',
                "TranslateVectorZ:=", str(z) + 'mm'
            ])

    def create_box(self, x_position, y_position, z_position, x_size, y_size, z_size, name, color='(165 42 42)'):
        self.oEditor.CreateBox(
            [
                "NAME:BoxParameters",
                "XPosition:=", str(x_position) + 'mm',
                "YPosition:=", str(y_position) + 'mm',
                "ZPosition:=", str(z_position) + 'mm',
                "XSize:=", str(x_size) + 'mm',
                "YSize:=", str(y_size) + 'mm',
                "ZSize:=", str(z_size) + 'mm'
            ],
            [
                "NAME:Attributes",
                "Name:=", name,
                "Flags:=", "",
                "Color:=", color,
                "Transparency:=", 0,
                "PartCoordinateSystem:=", Cores.CS,
                "UDMId:=", "",
                "MaterialValue:=", '"ferrite"',
                "SurfaceMaterialValue:=", "\"\"",
                "SolveInside:=", True,
                "IsMaterialEditable:=", True,
                "UseMaterialAppearance:=", False
            ])

    def create_rectangle(self, x_position, y_position, z_position, x_size, z_size, name, covered=True):
        self.oEditor.CreateRectangle(
                    [
                        "NAME:RectangleParameters",
                        "IsCovered:=", covered,
                        "XStart:=", str(x_position) + 'mm',
                        "YStart:=", str(y_position) + 'mm',
                        "ZStart:=", str(z_position) + 'mm',
                        "Width:=", str(z_size) + 'mm',    # height of conductor
                        "Height:=", str(x_size) + 'mm',    # width
                        "WhichAxis:=", 'Y'
                    ],
                    [
                        "NAME:Attributes",
                        "Name:=", name,
                        "Flags:=", "",
                        "Color:=", "(143 175 143)",
                        "Transparency:=", 0,
                        "PartCoordinateSystem:=", Cores.CS,
                        "UDMId:=", "",
                        "MaterialValue:=", "\"vacuum\"",
                        "SurfaceMaterialValue:=", "\"\"",
                        "SolveInside:=", True,
                        "IsMaterialEditable:=", True,
                        "UseMaterialAppearance:=", False
                    ])

    def create_cylinder(self, x_position, y_position, z_position, diameter, height, seg_num, name, color='(165 42 42)'):
        self.oEditor.CreateCylinder(
            [
                "NAME:CylinderParameters",
                "XCenter:=", str(x_position) + 'mm',
                "YCenter:=", str(y_position) + 'mm',
                "ZCenter:=", str(z_position) + 'mm',
                "Radius:=", str(diameter/2) + 'mm',
                "Height:=", str(height) + 'mm',
                "WhichAxis:=", "Z",
                "NumSides:=", str(seg_num)
            ],
            [
                "NAME:Attributes",
                "Name:=", name,
                "Flags:=", "",
                "Color:=", color,
                "Transparency:=", 0,
                "PartCoordinateSystem:=", Cores.CS,
                "UDMId:=", "",
                "MaterialValue:=", '"ferrite"',
                "SurfaceMaterialValue:=", "\"\"",
                "SolveInside:=", True,
                "IsMaterialEditable:=", True,
                "UseMaterialAppearance:=", False
            ])

    def create_circle(self, x_position, y_position, z_position, diameter, seg_num, name, axis, covered=True):
        self.oEditor.CreateCircle(
            [
                "NAME:CircleParameters",
                "IsCovered:=", covered,
                "XCenter:=", str(x_position) + 'mm',
                "YCenter:=", str(y_position) + 'mm',
                "ZCenter:=", str(z_position) + 'mm',
                "Radius:=", str(diameter/2) + 'mm',
                "WhichAxis:=", axis,
                "NumSegments:=", str(seg_num)
            ],
            [
                "NAME:Attributes",
                "Name:=", name,
                "Flags:=", "",
                "Color:=", "(143 175 143)",
                "Transparency:=", 0,
                "PartCoordinateSystem:=", Cores.CS,
                "UDMId:=", "",
                "MaterialValue:=", "\"vacuum\"",
                "SurfaceMaterialValue:=", "\"\"",
                "SolveInside:=", True,
                "IsMaterialEditable:=", True,
                "UseMaterialAppearance:=", False
            ])

    def fillet(self, name, radius, x, y, z):
        object_id = [self.oEditor.GetEdgeByPosition(
                                [
                                  "NAME:EdgeParameters",
                                  "BodyName:=", name,
                                  "XPosition:=", str(x) + 'mm',
                                  "YPosition:=", str(y) + 'mm',
                                  "ZPosition:=", str(z) + 'mm'
                                ])
        ]

        self.oEditor.Fillet(
                [
                    "NAME:Selections",
                    "Selections:=", name,
                    "NewPartsModelFlag:=", "Model"
                ],
                [
                    "NAME:Parameters",
                    [
                        "NAME:FilletParameters",
                        "Edges:=", object_id,
                        "Vertices:=", [],
                        "Radius:=", str(radius) + 'mm',
                        "Setback:=", "0mm"
                    ]
                ])

    def rename(self, old_name, new_name):
        self.oEditor.ChangeProperty(
            ["NAME:AllTabs",
                ["NAME:Geometry3DAttributeTab",
                    ["NAME:PropServers", old_name],
                    ["NAME:ChangedProps",
                        ["NAME:Name", "Value:=", new_name]
                    ]
                ]
            ])

    def unite(self, selection):
        self.oEditor.Unite(
            ["NAME:Selections", "Selections:=", selection],
            ["NAME:UniteParameters", "KeepOriginals:=", False])

    def subtract(self, selection, tool, keep=False):
        self.oEditor.Subtract(
            [
                "NAME:Selections",
                "Blank Parts:=", selection,
                "Tool Parts:=", tool
            ],
            ["NAME:SubtractParameters", "KeepOriginals:=", keep])

    def sweep_along_vector(self, selection, vec_z):
        self.oEditor.SweepAlongVector(
            [
                "NAME:Selections",
                "Selections:=", selection,
                "NewPartsModelFlag:=", "Model"
            ],
            [
                "NAME:VectorSweepParameters",
                "DraftAngle:=", "0deg",
                "DraftType:=", "Round",
                "CheckFaceFaceIntersection:=", False,
                "SweepVectorX:=", "0mm",
                "SweepVectorY:=", "0mm",
                "SweepVectorZ:=", str(vec_z) + 'mm'
            ])

    def sweep_along_path(self, selection):
        self.oEditor.SweepAlongPath(
            [
                "NAME:Selections",
                "Selections:=", selection,
                "NewPartsModelFlag:=", "Model"
            ],
            [
                "NAME:PathSweepParameters",
                "DraftAngle:=", "0deg",
                "DraftType:=", "Round",
                "CheckFaceFaceIntersection:=", False,
                "TwistAngle:=", "0deg"
            ])

    def duplicate_mirror(self, selection, X, Y, Z):
        self.oEditor.DuplicateMirror(
            [
                "NAME:Selections",
                "Selections:=", selection,
                "NewPartsModelFlag:=", "Model"
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

    def points_segments_generator(self, points_list):
        for i in range(points_list.count([])):
            points_list.remove([])

        points_array = ["NAME:PolylinePoints"]
        segments_array = ["NAME:PolylineSegments"]
        for i in range(len(points_list)):
            points_array.append(self.polyline_point(points_list[i][0], points_list[i][1], points_list[i][2]))
            if i != len(points_list) - 1:
                segments_array.append(self.line_segment(i))
        return points_array, segments_array
    
    @staticmethod
    def polyline_point(x=0, y=0, z=0):
        point = [
                        "NAME:PLPoint",
                        "X:=", str(x) + 'mm',
                        "Y:=", str(y) + 'mm',
                        "Z:=", str(z) + 'mm'
                    ]
        return point

    @staticmethod
    def line_segment(index):
        segment = [
            "NAME:PLSegment",
            "SegmentType:=", "Line",
            "StartIndex:="	, index,
            "NoOfPoints:="		, 2
        ]
        return segment

    @staticmethod
    def arc_segment(index, segments_number, center_x, center_y, center_z):
        segment = [
            "NAME:PLSegment",
            "SegmentType:="	, "AngularArc",
            "StartIndex:="		, index,
            "NoOfPoints:="		, 3,
            "NoOfSegments:="	, str(segments_number),
            "ArcAngle:="		, "-90deg",
            "ArcCenterX:="		, "{}mm".format(center_x),
            "ArcCenterY:="		, "{}mm".format(center_y),
            "ArcCenterZ:="		, "{}mm".format(center_z),
            "ArcPlane:="		, "XY"
        ]
        return segment


class ECore(Cores):
    def __init__(self, args_list):
        super(ECore, self).__init__(args_list)
        self.core_length = self.dim_D1
        self.core_width = self.dim_D6
        self.core_height = self.dim_D4
        self.side_leg_width = (self.dim_D1 - self.dim_D2) / 2
        self.center_leg_width = self.dim_D3
        self.slot_depth = self.dim_D5

    def draw_winding(self, dim_D2, dim_D3, dim_D5, dim_D6, segmentation_angle, ECoredim_D8=0.0):
        layer_spacing = self.layer_spacing.Value
        top_margin = self.top_margin.Value
        side_margin = self.side_margin.Value
        bobbin_thickness = self.bobbin_board_thickness.Value
        winding_parameters_dict = self.winding_parameters_dict
        slot_height = dim_D5*2

        if self.layer_type.Value == "Planar":
            margin = top_margin
            oDesktop.AddMessage("", "", 1, "ACT:" + str(winding_parameters_dict))
            for layer_num in winding_parameters_dict:
                conductor_width = winding_parameters_dict[layer_num][0]  # depends on Wound/Planar
                conductor_height = winding_parameters_dict[layer_num][1]
                num_of_turns = int(winding_parameters_dict[layer_num][2])
                insulation_thickness = winding_parameters_dict[layer_num][3]

                if self.include_bobbin.Value:
                    self.draw_board(slot_height - 2 * top_margin, (dim_D2 / 2.0), (dim_D3 / 2.0) + bobbin_thickness,
                                    (dim_D6 / 2.0) + bobbin_thickness, bobbin_thickness, -dim_D5 + margin, layer_num,
                                    ECoredim_D8)

                for turn_num in range(0, num_of_turns):
                    sweep_path_x = (
                        dim_D3 + 2 * side_margin + ((2 * turn_num + 1) * conductor_width) +
                        (2 * turn_num * insulation_thickness) + insulation_thickness)

                    sweep_path_y = (
                        dim_D6 + 2 * side_margin + ((2 * turn_num + 1) * conductor_width) +
                        (2 * turn_num * insulation_thickness) + insulation_thickness)

                    position_z_initial = -conductor_height / 2.0

                    self.create_single_turn(sweep_path_x, sweep_path_y, position_z_initial,
                                            conductor_width, conductor_height,
                                            -dim_D5 + margin + bobbin_thickness + conductor_height,
                                            layer_num, turn_num, (sweep_path_x - dim_D3) / 2, segmentation_angle)

                margin += layer_spacing + conductor_height + bobbin_thickness

        else:
            # ---- Wound transformer ---- #
            if self.include_bobbin.Value:
                self.draw_bobbin(slot_height - 2 * top_margin, (dim_D2 / 2.0), (dim_D3 / 2.0) + bobbin_thickness,
                                 (dim_D6 / 2.0) + bobbin_thickness, bobbin_thickness, ECoredim_D8)

            margin = side_margin + bobbin_thickness
            for layer_num in winding_parameters_dict:
                conductor_diameter = conductor_width = winding_parameters_dict[layer_num][0]  # for Circle = for Rect
                number_segments = conductor_height = winding_parameters_dict[layer_num][1]
                num_of_turns = int(winding_parameters_dict[layer_num][2])
                insulation_thickness = winding_parameters_dict[layer_num][3]

                conductor_z_position = dim_D5 - top_margin - bobbin_thickness - insulation_thickness

                # factor of 2 is applied due to existence of margin and insulation on both sides
                conductor_full_size = 2*margin + conductor_width + 2*insulation_thickness
                sweep_path_x = dim_D3 + conductor_full_size
                sweep_path_y = dim_D6 + conductor_full_size

                fillet_radius = sweep_path_x - dim_D3/2
                if self.conductor_type.Value == 'Rectangular':
                    position_z_initial = -conductor_height/2.0
                    move_vec = 2 * insulation_thickness + conductor_height
                    name = self.create_single_turn(sweep_path_x, sweep_path_y, position_z_initial,
                                                   conductor_z_position, layer_num, fillet_radius, segmentation_angle,
                                                   profile_width=conductor_width, profile_height=conductor_height)

                else:
                    position_z_initial = -conductor_diameter/2.0
                    move_vec = 2*insulation_thickness + conductor_diameter
                    name = self.create_single_turn(sweep_path_x, sweep_path_y, position_z_initial,
                                                   conductor_z_position, layer_num, fillet_radius, segmentation_angle,
                                                   profile_diameter=conductor_diameter,
                                                   profile_segments_num=number_segments)

                margin += layer_spacing + winding_parameters_dict[layer_num][0] + 2*insulation_thickness
                
                if num_of_turns > 1:
                    self.oEditor.DuplicateAlongLine(
                        [
                            "NAME:Selections",
                            "Selections:="	, name,
                            "NewPartsModelFlag:="	, "Model"
                        ],
                        [
                            "NAME:DuplicateToAlongLineParameters",
                            "CreateNewObjects:="	, True,
                            "XComponent:="		, "0mm",
                            "YComponent:="		, "0mm",
                            "ZComponent:="		, "-{}mm".format(move_vec),
                            "NumClones:="		, str(num_of_turns)
                        ],
                        [
                            "NAME:Options",
                            "DuplicateAssignments:=", False
                        ],
                        [
                            "CreateGroupsForNewObjects:=", False
                        ])

    def create_single_turn(self, sweep_path_x, sweep_path_y, position_z_initial, position_z_final,
                           layer_number, fillet_radius, segmentation_angle, profile_width=0,
                           profile_height=0, profile_diameter=0, profile_segments_num=8):
        
        name = self.create_sweep_path(sweep_path_x, sweep_path_y, position_z_initial, fillet_radius,
                                      layer_number, segmentation_angle)
        if self.conductor_type.Value == 'Rectangular':
            self.create_rectangle((sweep_path_x - profile_width) / 2, 0, (position_z_initial - profile_height / 2),
                                  profile_width, profile_height, 'sweep_profile')
        else:
            self.create_circle(sweep_path_x / 2, 0, position_z_initial,
                               profile_diameter, profile_segments_num, 'sweep_profile', 'Y')

        self.sweep_along_path('sweep_profile,' + name)

        self.move('sweep_profile', 0, 0, position_z_final)
        self.rename('sweep_profile', 'Layer'.format(layer_number))
        return 'Layer'.format(layer_number)

    def create_sweep_path(self, sweep_path_x, sweep_path_y, position_z, fillet_radius,
                          layer_number, segmentation_angle=0):
        """
        Function to create rectangular profile of the winding.
        Note: in this function we sacrifice code size for the explicitness. Generate points for each quadrant using
        polyline that consists of arcs and straight lines
        :param sweep_path_x: size of the coil in X direction
        :param sweep_path_y: size of the coil in Y direction
        :param position_z: position of the coil on Z axis
        :param fillet_radius: radius of the fillet, to make round corners of the coil
        :param layer_number: layer number for the name
        :param segmentation_angle: segmentation angle, true surface would be mesh overkill
        :return: name: name of the winding
        """

        segments_number = 12 if segmentation_angle == 0 else int(90 / segmentation_angle) * 2

        points = ["NAME:PolylinePoints"]
        points.append(self.polyline_point(x=-sweep_path_x/2 + fillet_radius,
                                          y=sweep_path_y/2,
                                          z=position_z))
        points.append(self.polyline_point(x=sweep_path_x/2 - fillet_radius,
                                          y=sweep_path_y/2,
                                          z=position_z))
        points.append(self.polyline_point(z=position_z))  # dummy point, aedt needs it
        points.append(self.polyline_point(x=sweep_path_x/2,
                                          y=sweep_path_y/2 - fillet_radius,
                                          z=position_z))
        points.append(self.polyline_point(x=sweep_path_x/2,
                                          y=-sweep_path_y/2 + fillet_radius,
                                          z=position_z))
        points.append(self.polyline_point())  # dummy point
        points.append(self.polyline_point(x=sweep_path_x/2 - fillet_radius,
                                          y=-sweep_path_y/2,
                                          z=position_z))
        points.append(self.polyline_point(x=-sweep_path_x/2 + fillet_radius,
                                          y=-sweep_path_y/2,
                                          z=position_z))
        points.append(self.polyline_point(z=position_z))  # dummy point
        points.append(self.polyline_point(x=-sweep_path_x/2,
                                          y=-sweep_path_y/2 + fillet_radius,
                                          z=position_z))
        points.append(self.polyline_point(x=-sweep_path_x/2,
                                          y=sweep_path_y/2 - fillet_radius,
                                          z=position_z))
        points.append(self.polyline_point(z=position_z))  # dummy point
        points.append(self.polyline_point(x=-sweep_path_x/2 + fillet_radius,
                                          y=sweep_path_y/2,
                                          z=position_z))

        segments = ["NAME:PolylineSegments"]
        segments.append(self.line_segment(0))
        segments.append(self.arc_segment(index=1,
                                         segments_number=segments_number,
                                         center_x=sweep_path_x/2 - fillet_radius,
                                         center_y=sweep_path_y/2 - fillet_radius,
                                         center_z=position_z))
        segments.append(self.line_segment(3))
        segments.append(self.arc_segment(index=4,
                                         segments_number=segments_number,
                                         center_x=sweep_path_x/2 - fillet_radius,
                                         center_y=-sweep_path_y/2 + fillet_radius,
                                         center_z=position_z))
        segments.append(self.line_segment(6))
        segments.append(self.arc_segment(index=7,
                                         segments_number=segments_number,
                                         center_x=-sweep_path_x/2 + fillet_radius,
                                         center_y=-sweep_path_y/2 + fillet_radius,
                                         center_z=position_z))
        segments.append(self.line_segment(9))
        segments.append(self.arc_segment(index=10,
                                         segments_number=segments_number,
                                         center_x=-sweep_path_x/2 + fillet_radius,
                                         center_y=sweep_path_y/2 - fillet_radius,
                                         center_z=position_z))

        name = "Tool{}".format(layer_number)
        self.create_polyline(points, segments, name, covered=False, closed=True)

        return name
    
    def draw_bobbin(self, Hb, Db1, Db2, Db3, Tb, e_coredim_d8=0.0):
        BRadE = (Db1 - Db2)

        self.create_box(-Db1 + e_coredim_d8, -(Db1 - Db2 + Db3), - Tb + (Hb/2.0),
                        2 * Db1 - 2 * e_coredim_d8, 2 * (Db1 - Db2 + Db3), Tb, 'Bobbin', '(255, 248, 157)')

        self.create_box(-Db1 + e_coredim_d8, -(Db1 - Db2 + Db3), -(Hb/2.0),
                        2 * Db1 - 2 * e_coredim_d8, 2 * (Db1 - Db2 + Db3), Tb, 'BobT2')

        self.create_box(-Db2, -Db3, Tb - (Hb/2.0),
                        2 * Db2, 2 * Db3, Hb - 2 * Tb, 'BobT3')

        self.unite('Bobbin,BobT2,BobT3')

        self.create_box(- Db2 + Tb, - Db3 + Tb, - Hb/2.0, 2 * Db2 - 2 * Tb, 2 * Db3 - 2 * Tb, Hb, 'BobSlot')
        self.subtract('Bobbin', 'BobSlot')

        # - - - fillet(NameOfObject, Radius, XCoord, YCoord, ZCoord)
        self.fillet('Bobbin', BRadE, -Db1 + e_coredim_d8, -(Db1 - Db2 + Db3), (- Hb + Tb)/2)
        self.fillet('Bobbin', BRadE, -Db1 + e_coredim_d8, (Db1 - Db2 + Db3), (- Hb + Tb)/2)
        self.fillet('Bobbin', BRadE, Db1 - e_coredim_d8, -(Db1 - Db2 + Db3), (- Hb + Tb)/2)
        self.fillet('Bobbin', BRadE, Db1 - e_coredim_d8, (Db1 - Db2 + Db3), (- Hb + Tb)/2)
        self.fillet('Bobbin', BRadE, -Db1 + e_coredim_d8, -(Db1 - Db2 + Db3), (Hb - Tb)/2)
        self.fillet('Bobbin', BRadE, -Db1 + e_coredim_d8, (Db1 - Db2 + Db3), (Hb - Tb)/2)
        self.fillet('Bobbin', BRadE, Db1 - e_coredim_d8, -(Db1 - Db2 + Db3), (Hb - Tb)/2)
        self.fillet('Bobbin', BRadE, Db1 - e_coredim_d8, (Db1 - Db2 + Db3), (Hb - Tb)/2)

    def draw_board(self, Hb, Db1, Db2, Db3, Tb, ZStart, layer_number, e_coredim_d8=0.0):
        self.create_box(-Db1 + e_coredim_d8, -(Db1 - Db2 + Db3), ZStart,
                        2 * Db1 - 2 * e_coredim_d8, 2 * (Db1 - Db2 + Db3), Tb, 'Board_{}'.format(layer_number), '(0, 128, 0)')

        self.create_box(-Db2 + Tb, -Db3 + Tb, ZStart,
                        2 * Db2 - 2 * Tb, 2 * Db3 - 2 * Tb, Hb, 'BoardSlot_{}'.format(layer_number))
        self.subtract('Board_{}'.format(layer_number), 'BoardSlot_{}'.format(layer_number))

    def draw_geometry(self):
        self.create_box(-(self.core_length / 2), -(self.core_width / 2), -self.core_height - self.airgap_both,
                        self.core_length, self.core_width, (self.core_height - self.slot_depth), 'E_Core_Bottom')

        self.create_box(-(self.core_length / 2), -(self.core_width / 2), -self.core_height - self.airgap_both,
                        self.side_leg_width, self.core_width, self.core_height - self.airgap_side, 'Leg1')

        self.create_box(-(self.center_leg_width / 2), -(self.core_width / 2), -self.core_height - self.airgap_both,
                        self.center_leg_width, self.core_width, self.core_height - self.airgap_center, 'Leg2')

        self.create_box((self.core_length / 2) - self.side_leg_width, -(self.core_width / 2),
                        -self.core_height - self.airgap_both,
                        self.side_leg_width, self.core_width, self.core_height - self.airgap_side, 'Leg3')

        self.unite('E_Core_Bottom,Leg1,Leg2,Leg3')

        self.fillet('E_Core_Bottom', self.dim_D7, -self.dim_D1 / 2, 0, -self.dim_D4 - self.airgap_both)    # outer edges D_7
        self.fillet('E_Core_Bottom', self.dim_D7, self.dim_D1 / 2, 0, -self.dim_D4 - self.airgap_both)    # outer edges D_7
        self.fillet('E_Core_Bottom', self.dim_D8, -self.dim_D2 / 2, 0, -self.dim_D5 - self.airgap_both)    # inner edges D_8
        self.fillet('E_Core_Bottom', self.dim_D8, self.dim_D2 / 2, 0, -self.dim_D5 - self.airgap_both)    # inner edges D_8

        self.duplicate_mirror('E_Core_Bottom', 0, 0, 1)

        self.rename('E_Core_Bottom_1', 'E_Core_Top')
        self.oEditor.FitAll()

        self.draw_winding(self.dim_D2, self.dim_D3, self.dim_D5, self.dim_D6, self.segmentation_angle.Value,
                          self.dim_D8)


# EICore inherit from ECore functions DrawWdg, CreateSingleTurn, draw_bobbin
class EICore(ECore):
    def draw_geometry(self):
        self.create_box(-(self.core_length / 2), -(self.core_width / 2), -self.core_height,
                        self.core_length, self.core_width, (self.core_height - self.slot_depth), 'E_Core')

        self.create_box(-(self.core_length / 2), -(self.core_width / 2), -self.core_height,
                        self.side_leg_width, self.core_width, self.core_height - 2 * self.airgap_side, 'Leg1')

        self.create_box(-(self.center_leg_width / 2), -(self.core_width / 2), -self.core_height,
                        self.center_leg_width, self.core_width, self.core_height - 2 * self.airgap_center, 'Leg2')

        self.create_box((self.core_length / 2) - self.side_leg_width, -(self.core_width / 2), -self.core_height,
                        self.side_leg_width, self.core_width, self.core_height - 2 * self.airgap_side, 'Leg3')

        self.unite('E_Core,Leg1,Leg2,Leg3')

        self.fillet('E_Core', self.dim_D7, -self.dim_D1 / 2, -self.dim_D6 / 2, -(self.dim_D4/2) - 2 * self.airgap_both)
        self.fillet('E_Core', self.dim_D7, -self.dim_D1 / 2, self.dim_D6 / 2, -(self.dim_D4/2) - 2 * self.airgap_both)
        self.fillet('E_Core', self.dim_D7, self.dim_D1 / 2, -self.dim_D6 / 2, -(self.dim_D4/2) - 2 * self.airgap_both)
        self.fillet('E_Core', self.dim_D7, self.dim_D1 / 2, self.dim_D6 / 2, -(self.dim_D4/2) - 2 * self.airgap_both)

        # two multiply air gap due to no symmetry for cores like in ECore
        self.create_box(-self.dim_D1 / 2, -self.dim_D6 / 2, 2 * self.airgap_both,
                        self.dim_D1, self.dim_D6, self.dim_D8, 'I_Core')

        self.fillet('I_Core', self.dim_D7, -self.dim_D1 / 2, -self.dim_D6 / 2, (self.dim_D8/2) - 2 * self.airgap_both)
        self.fillet('I_Core', self.dim_D7, -self.dim_D1 / 2, self.dim_D6 / 2, (self.dim_D8/2) - 2 * self.airgap_both)
        self.fillet('I_Core', self.dim_D7, self.dim_D1 / 2, -self.dim_D6 / 2, (self.dim_D8/2) - 2 * self.airgap_both)
        self.fillet('I_Core', self.dim_D7, self.dim_D1 / 2, self.dim_D6 / 2, (self.dim_D8/2) - 2 * self.airgap_both)

        self.move('E_Core,I_Core', 0, 0, self.dim_D5/2)

        self.oEditor.FitAll()

        self.draw_winding(self.dim_D2, self.dim_D3, self.dim_D5 / 2 + 2 * self.airgap_both, self.dim_D6,
                          self.segmentation_angle.Value)


class EFDCore(ECore):
    def draw_geometry(self):
        self.create_box(-(self.core_length / 2), -(self.core_width / 2), -self.core_height - self.airgap_both,
                        self.core_length, self.core_width, (self.core_height - self.slot_depth), 'EFD_Core_Bottom')

        self.create_box(-(self.core_length / 2), -(self.core_width / 2), -self.core_height - self.airgap_both,
                        self.side_leg_width, self.core_width, self.core_height - self.airgap_side, 'Leg1')

        self.create_box(-(self.center_leg_width / 2), -(self.core_width / 2) - self.dim_D8, -self.core_height - self.airgap_both,
                        self.center_leg_width, self.dim_D7, self.core_height - self.airgap_center, 'Leg2')

        self.create_box((self.core_length / 2) - self.side_leg_width, -(self.core_width / 2),
                        -self.core_height - self.airgap_both,
                        self.side_leg_width, self.core_width, self.core_height - self.airgap_side, 'Leg3')

        self.unite('EFD_Core_Bottom,Leg1,Leg2,Leg3')

        self.duplicate_mirror('EFD_Core_Bottom', 0, 0, 1)
        self.rename('EFD_Core_Bottom_1', 'E_Core_Top')
        self.oEditor.FitAll()

        # difference between ECore and EFD is only offset    of central leg.
        # We can compensate it by using relative CS for winding and bobbin

        self.oEditor.CreateRelativeCS(
            [
                "NAME:RelativeCSParameters",
                "Mode:=", "Axis/Position",
                "OriginX:=", "0mm",
                "OriginY:=", str((self.dim_D7 - self.dim_D6)/2 - self.dim_D8) + 'mm',
                "OriginZ:=", "0mm",
                "XAxisXvec:=", "1mm",
                "XAxisYvec:=", "0mm",
                "XAxisZvec:=", "0mm",
                "YAxisXvec:=", "0mm",
                "YAxisYvec:=", "1mm",
                "YAxisZvec:=", "0mm"
            ],
            [
                "NAME:Attributes",
                "Name:=", 'CentralLegCS'
            ])
        Cores.CS = 'CentralLegCS'

        self.draw_winding(self.dim_D2, self.dim_D3, self.dim_D5, self.dim_D7, self.segmentation_angle.Value)


# UCore inherit from ECore functions CreateSingleTurn, draw_bobbin, drawBoard
class UCore(ECore):
    def __init__(self, args_list):
        super(UCore, self).__init__(args_list)

        self.MECoreLength = self.dim_D1
        self.MECoreWidth = self.dim_D5
        self.MECoreHeight = self.dim_D3

    def draw_winding(self, dim_D1, dim_D2, dim_D3, dim_D4, dim_D5, segmentation_angle):
        MLSpacing = self.layer_spacing.Value
        Mtop_margin = self.top_margin.Value
        Mside_margin = self.side_margin.Value
        MBobbinThk = self.bobbin_board_thickness.Value
        WdgParDict = self.winding_parameters_dict
        MSlotHeight = dim_D4*2
        LegWidth = (dim_D1 - dim_D2)/2

        if self.include_bobbin.Value and self.layer_type.Value == "Wound":
            self.draw_bobbin(MSlotHeight - 2*Mtop_margin, dim_D2 + LegWidth/2, LegWidth/2.0 + MBobbinThk,
                            (dim_D5/2.0) + MBobbinThk, MBobbinThk)

        if self.layer_type.Value == "Planar":
            MTDx = Mtop_margin
            for MAx in self.winding_parameters_dict:
                if self.include_bobbin.Value:
                    self.draw_board(MSlotHeight - 2 * Mtop_margin, dim_D2 + LegWidth/2, LegWidth/2.0 + MBobbinThk,
                                    (dim_D5/2.0) + MBobbinThk, MBobbinThk, -dim_D4 + MTDx, MAx)

                for layerTurn in range(0, int(self.winding_parameters_dict[MAx][2])):
                    rectangle_size_x = (LegWidth + 2 * Mside_margin + ((2*layerTurn + 1) * self.winding_parameters_dict[MAx][0]) +
                                        (2 * layerTurn * 2 * self.winding_parameters_dict[MAx][3]) + 2 * self.winding_parameters_dict[MAx][3])

                    rectangle_size_y = (dim_D5 + 2 * Mside_margin + ((2*layerTurn + 1) * self.winding_parameters_dict[MAx][0]) +
                                        (2 * layerTurn * 2 * self.winding_parameters_dict[MAx][3]) + 2 * self.winding_parameters_dict[MAx][3])

                    rectangle_size_z = -self.winding_parameters_dict[MAx][1] / 2.0
                    self.create_single_turn(rectangle_size_x, rectangle_size_y, rectangle_size_z,
                                            self.winding_parameters_dict[MAx][0], self.winding_parameters_dict[MAx][1],
                                            (-dim_D4 + MTDx + MBobbinThk + WdgParDict[MAx][1]), MAx,
                                            layerTurn, (rectangle_size_x - LegWidth) / 2, segmentation_angle)

                MTDx += MLSpacing + self.winding_parameters_dict[MAx][1] + MBobbinThk

        else:
            # ---- Wound transformer ---- #
            MTDx = Mside_margin + MBobbinThk
            for MAx in self.winding_parameters_dict.keys():
                for MBx in range(0, int(self.winding_parameters_dict[MAx][2])):
                    rectangle_size_x = LegWidth + (2 * (MTDx + (self.winding_parameters_dict[MAx][0] / 2.0))) + 2 * self.winding_parameters_dict[MAx][3]
                    rectangle_size_y = dim_D5 + (2 * (MTDx + (self.winding_parameters_dict[MAx][0] / 2.0))) + 2 * self.winding_parameters_dict[MAx][3]
                    if self.conductor_type.Value == 'Rectangular':
                        rectangle_size_z = -self.winding_parameters_dict[MAx][1] / 2.0
                        ZProf = (dim_D4 - Mtop_margin - MBobbinThk - (self.winding_parameters_dict[MAx][3]) -
                                 MBx * (2 * self.winding_parameters_dict[MAx][3] + self.winding_parameters_dict[MAx][1]))

                        self.create_single_turn(rectangle_size_x, rectangle_size_y, rectangle_size_z,
                                                self.winding_parameters_dict[MAx][0], self.winding_parameters_dict[MAx][1], ZProf, MAx, MBx,
                                                (rectangle_size_x - LegWidth) / 2, segmentation_angle)
                    else:
                        rectangle_size_z = -self.winding_parameters_dict[MAx][0] / 2.0
                        ZProf = (dim_D4 - Mtop_margin - MBobbinThk - (self.winding_parameters_dict[MAx][3]) -
                                 MBx * (2 * self.winding_parameters_dict[MAx][3] + self.winding_parameters_dict[MAx][0]))

                        self.create_single_turn(rectangle_size_x, rectangle_size_y, rectangle_size_z,
                                                self.winding_parameters_dict[MAx][0], self.winding_parameters_dict[MAx][1], ZProf, MAx, MBx,
                                                (rectangle_size_x - LegWidth) / 2, segmentation_angle)

                MTDx = MTDx + MLSpacing + self.winding_parameters_dict[MAx][0] + 2 * self.winding_parameters_dict[MAx][3]

    def draw_geometry(self):
        self.create_box(-(self.dim_D1 - self.dim_D2) / 4, -(self.MECoreWidth/2), -self.MECoreHeight - self.airgap_both,
                        self.MECoreLength, self.MECoreWidth, self.MECoreHeight, 'U_Core_Bottom')

        self.create_box((self.dim_D1 - self.dim_D2) / 4, -self.MECoreWidth / 2, -self.dim_D4 - self.airgap_both,
                        self.dim_D2, self.MECoreWidth, self.dim_D4, 'XSlot')

        self.subtract('U_Core_Bottom', 'XSlot')

        if self.airgap_center > 0:
            self.create_box(-(self.dim_D1 - self.dim_D2) / 4, -self.MECoreWidth / 2, -self.airgap_center,
                            (self.dim_D1 - self.dim_D2) / 2, self.MECoreWidth, self.airgap_center, 'AgC')

            self.subtract('U_Core_Bottom', 'AgC')

        if self.airgap_side > 0:
            self.create_box(self.dim_D2 + (self.dim_D1 - self.dim_D2) / 4, -self.MECoreWidth / 2, -self.airgap_side,
                            (self.dim_D1 - self.dim_D2) / 2, self.MECoreWidth, self.airgap_side, 'AgS')

            self.subtract('U_Core_Bottom', 'AgS')

        self.duplicate_mirror('U_Core_Bottom', 0, 0, 1)
        self.rename('U_Core_Bottom_1', 'U_Core_Top')
        self.oEditor.FitAll()

        self.draw_winding(self.dim_D1, self.dim_D2, self.dim_D3, (self.dim_D4 + self.airgap_both), self.dim_D5,
                          self.segmentation_angle.Value)


# UICore inherit from ECore functions CreateSingleTurn, draw_bobbin, drawBoard
# and from UCore inherit DrawWdg
class UICore(UCore):
    def draw_geometry(self):
        self.create_box(-(self.dim_D1 - self.dim_D2) / 4, -(self.MECoreWidth/2), (self.dim_D4/2) - self.MECoreHeight,
                        self.MECoreLength, self.MECoreWidth, self.MECoreHeight, 'U_Core')

        self.create_box((self.dim_D1 - self.dim_D2) / 4, -self.MECoreWidth/2, -self.dim_D4/2,
                        self.dim_D2, self.MECoreWidth, self.dim_D4, 'XSlot')

        self.subtract('U_Core', 'XSlot')

        if self.airgap_center > 0:
            self.create_box((self.dim_D1 - self.dim_D2) / 4, -self.MECoreWidth / 2, (self.dim_D4/2) - 2 * self.airgap_center,
                            -(self.dim_D1 - self.dim_D2) / 2, self.MECoreWidth, 2 * self.airgap_center, 'AgC')
            self.subtract('U_Core', 'AgC')

        if self.airgap_side > 0:
            self.create_box(self.dim_D2 + (self.dim_D1 - self.dim_D2) / 4, -self.MECoreWidth / 2,
                            (self.dim_D4/2) - 2 * self.airgap_side,
                            (self.dim_D1 - self.dim_D2) / 2, self.MECoreWidth, 2 * self.airgap_side, 'AgS')
            self.subtract('U_Core', 'AgS')

        self.create_box(-(self.dim_D1 - self.dim_D2) / 4 + (self.dim_D1 - self.dim_D6) / 2, -self.dim_D7 / 2,
                        (self.dim_D4/2) + 2 * self.airgap_both,
                        self.dim_D6, self.dim_D7, self.dim_D8, 'I_Core')

        self.oEditor.FitAll()

        self.draw_winding(self.dim_D1, self.dim_D2, self.dim_D3 / 2, (self.dim_D4 / 2) + 2 * self.airgap_both,
                          self.dim_D5, self.segmentation_angle.Value)


class PQCore(Cores):
    def __init__(self, args_list):
        super(PQCore, self).__init__(args_list)
        self.MECoreLength = self.dim_D1
        self.MECoreWidth = self.dim_D6
        self.MECoreHeight = self.dim_D4
        self.MSLegWidth = (self.dim_D1 - self.dim_D2)/2
        self.MCLegWidth = self.dim_D3
        self.MSlotDepth = self.dim_D5

    def draw_wdg(self, dim_D2, dim_D3, dim_D5, segmentation_angle):
        MLSpacing = self.layer_spacing.Value
        Mtop_margin = self.top_margin.Value
        Mside_margin = self.side_margin.Value
        MBobbinThk = self.bobbin_board_thickness.Value
        WdgParDict = self.winding_parameters_dict
        MSlotHeight = dim_D5*2

        if self.include_bobbin.Value and self.layer_type.Value == "Wound":
            self.draw_bobbin(MSlotHeight - 2*Mtop_margin, (dim_D2/2.0), (dim_D3/2.0) + MBobbinThk, MBobbinThk)

        if self.layer_type.Value == "Planar":
            MTDx = Mtop_margin
            for MAx in self.winding_parameters_dict:
                if self.include_bobbin.Value:
                    self.drawBoard(MSlotHeight - 2*Mtop_margin, (dim_D2/2.0), (dim_D3/2.0) + MBobbinThk,
                                   MBobbinThk, -dim_D5 + MTDx, MAx)

                for layerTurn in range(0, int(self.winding_parameters_dict[MAx][2])):
                    rectangle_size_x = (dim_D3 + 2 * Mside_margin + ((2*layerTurn + 1) * self.winding_parameters_dict[MAx][0]) +
                                        (2 * layerTurn * 2 * self.winding_parameters_dict[MAx][3]) + 2 * self.winding_parameters_dict[MAx][3])

                    rectangle_size_z = -self.winding_parameters_dict[MAx][1] / 2.0

                    self.create_single_turn(rectangle_size_x, rectangle_size_z, self.winding_parameters_dict[MAx][0],
                                            self.winding_parameters_dict[MAx][1], -dim_D5 + MTDx + MBobbinThk + WdgParDict[MAx][1],
                                            MAx, layerTurn, segmentation_angle)

                MTDx += MLSpacing + self.winding_parameters_dict[MAx][1] + MBobbinThk
        else:
            # ---- Wound transformer ---- #
            MTDx = Mside_margin + MBobbinThk
            for MAx in self.winding_parameters_dict.keys():
                for MBx in range(0, int(self.winding_parameters_dict[MAx][2])):
                    rectangle_size_x = dim_D3 + (2 * (MTDx + (self.winding_parameters_dict[MAx][0] / 2.0))) + 2 * self.winding_parameters_dict[MAx][3]
                    if self.conductor_type.Value == 'Rectangular':
                        rectangle_size_z = -self.winding_parameters_dict[MAx][1] / 2.0
                        ZProf = (dim_D5 - Mtop_margin - MBobbinThk - self.winding_parameters_dict[MAx][3] -
                                 MBx * (2 * self.winding_parameters_dict[MAx][3] + self.winding_parameters_dict[MAx][1]))

                        self.create_single_turn(rectangle_size_x, rectangle_size_z, self.winding_parameters_dict[MAx][0],
                                                self.winding_parameters_dict[MAx][1], ZProf, MAx, MBx, segmentation_angle)
                    else:
                        rectangle_size_z = -self.winding_parameters_dict[MAx][0] / 2.0
                        ZProf = (dim_D5 - Mtop_margin - MBobbinThk - self.winding_parameters_dict[MAx][3] -
                                 MBx * (2 * self.winding_parameters_dict[MAx][3] + self.winding_parameters_dict[MAx][0]))

                        self.create_single_turn(rectangle_size_x, rectangle_size_z, self.winding_parameters_dict[MAx][0],
                                                self.winding_parameters_dict[MAx][1], ZProf, MAx, MBx, segmentation_angle)

                MTDx = MTDx + MLSpacing + self.winding_parameters_dict[MAx][0] + 2 * self.winding_parameters_dict[MAx][3]

    def create_single_turn(self, sweep_path_x, sweep_path_z, ProfAX, ProfZ, ZPos, layer_number, turn_number, segmentation_angle):
        segments_number = 0 if segmentation_angle == 0 else int(360/segmentation_angle)

        self.create_circle(0, 0, sweep_path_z,
                           sweep_path_x, segments_number, 'pathLine', 'Z', covered=False)

        if self.conductor_type.Value == 'Rectangular':
            self.create_rectangle((sweep_path_x/2 - ProfAX/2), 0, (sweep_path_z - ProfZ/2),
                                  ProfAX, ProfZ, 'sweep_profile')
        else:
            self.create_circle(sweep_path_x/2, 0, sweep_path_z,
                               ProfAX, ProfZ, 'sweep_profile', 'Y')

        self.rename('pathLine', 'Tool%s_%s' % (layer_number, turn_number + 1))
        self.sweep_along_path('sweep_profile,Tool%s_%s' % (layer_number, turn_number + 1))

        self.move('sweep_profile', 0, 0, ZPos)
        self.rename('sweep_profile', 'Layer%s_%s' % (layer_number, turn_number + 1))

    def draw_bobbin(self, Hb, Db1, Db2, Tb):
        self.create_cylinder(0, 0, - Tb + (Hb/2.0),
                             Db1 * 2, Tb, self.segments_number, 'Bobbin', '(255, 248, 157)')

        self.create_cylinder(0, 0, -(Hb/2.0),
                             Db1 * 2, Tb, self.segments_number, 'BobT2')

        self.create_cylinder(0, 0, Tb - (Hb/2.0),
                             Db2 * 2, Hb - 2 * Tb, self.segments_number, 'BobT3')

        self.unite('Bobbin,BobT2,BobT3')
        self.create_cylinder(0, 0, - Hb/2.0,
                             (Db2 - Tb) * 2, Hb, self.segments_number, 'BobT4')

        self.subtract('Bobbin', 'BobT4')

    def drawBoard(self, Hb, Db1, Db2, Tb, ZStart, layer_number):
        self.create_cylinder(0, 0, ZStart,
                             Db1 * 2, Tb, self.segments_number, 'Board_{}'.format(layer_number), '(0, 128, 0)')

        self.create_cylinder(0, 0, ZStart,
                             (Db2 - Tb) * 2, Hb, self.segments_number, 'BoardSlot_{}'.format(layer_number))

        self.subtract('Board_{}'.format(layer_number), 'BoardSlot_{}'.format(layer_number))

    def draw_geometry(self):
        self.create_box(-self.dim_D1 / 2, -self.dim_D8 / 2, -(self.dim_D4/2) - self.airgap_both,
                        (self.dim_D1 - self.dim_D6) / 2, self.dim_D8, (self.dim_D4/2) - self.airgap_side, 'PQ_Core_Bottom')

        IntL = math.sqrt(((self.dim_D3/2) ** 2) - ((self.dim_D7/2) ** 2))
        vertices1 = [[-self.dim_D6 / 2, - 0.4 * self.dim_D8, -(self.dim_D4/2) - self.airgap_both],
                     [-self.dim_D6 / 2, 0.4 * self.dim_D8, -(self.dim_D4/2) - self.airgap_both],
                     [-self.dim_D7 / 2, +IntL, -(self.dim_D4/2) - self.airgap_both],
                     [-self.dim_D7 / 2, - IntL, -(self.dim_D4/2) - self.airgap_both]]
        vertices1.append(vertices1[0])

        points_array, segments_array = self.points_segments_generator(vertices1)
        self.create_polyline(points_array, segments_array, "Polyline1")
        self.sweep_along_vector("Polyline1", (self.dim_D4/2) - self.airgap_side)

        self.unite('PQ_Core_Bottom,Polyline1')

        self.create_cylinder(0, 0, -(self.dim_D5/2) - self.airgap_both,
                             self.dim_D2, self.dim_D5 / 2, self.segments_number, 'XCyl1')

        self.subtract('PQ_Core_Bottom', 'XCyl1')
        self.duplicate_mirror('PQ_Core_Bottom', 1, 0, 0)

        self.create_cylinder(0, 0, -(self.dim_D4/2) - self.airgap_both,
                             self.dim_D3, self.dim_D4 / 2 - self.airgap_center, self.segments_number, 'XCyl2')

        self.unite('PQ_Core_Bottom,PQ_Core_Bottom_1,XCyl2')
        self.duplicate_mirror('PQ_Core_Bottom', 0, 0, 1)

        self.rename('PQ_Core_Bottom_2', 'PQ_Core_Top')
        self.oEditor.FitAll()

        self.draw_wdg(self.dim_D2, self.dim_D3, self.dim_D5 / 2 + self.airgap_both, self.segmentation_angle.Value)


# ETDCore inherit from PQCore functions DrawWdg, CreateSingleTurn, draw_bobbin
class ETDCore(PQCore):
    def draw_geometry(self, coreName):
        self.create_box(-(self.MECoreLength/2), -(self.MECoreWidth/2), -self.MECoreHeight - self.airgap_both,
                        self.MECoreLength, self.MECoreWidth, self.MECoreHeight - self.airgap_side,
                        coreName + '_Core_Bottom')

        self.create_cylinder(0, 0, -self.MSlotDepth - self.airgap_both,
                             self.dim_D2, self.MSlotDepth, self.segments_number, 'XCyl1')

        self.subtract(coreName + '_Core_Bottom', 'XCyl1')

        if coreName == 'ER' and self.dim_D7 > 0:
            self.create_box(-self.dim_D7 / 2, -self.MECoreWidth / 2, -self.MSlotDepth - self.airgap_both,
                            self.dim_D7, self.MECoreWidth, self.MSlotDepth, 'Tool')
            self.subtract(coreName + '_Core_Bottom', 'Tool')

        self.create_cylinder(0, 0, -self.MSlotDepth - self.airgap_both,
                             self.dim_D3, self.MSlotDepth - self.airgap_center, self.segments_number, 'XCyl2')

        self.unite(coreName + '_Core_Bottom,XCyl2')
        self.duplicate_mirror(coreName + '_Core_Bottom', 0, 0, 1)
        self.rename(coreName + '_Core_Bottom_1', coreName + '_Core_Top')

        self.oEditor.FitAll()

        self.draw_wdg(self.dim_D2, self.dim_D3, (self.dim_D5 + self.airgap_both), self.segmentation_angle.Value)


# RMCore inherit from PQCore functions DrawWdg, CreateSingleTurn, draw_bobbin
class RMCore(PQCore):
    def draw_geometry(self):
        Dia = self.dim_D7/math.sqrt(2)

        vertices1 = [[-self.dim_D1 / 2, (Dia - (self.dim_D1/2)), -self.airgap_side - self.airgap_both]]
        vertices1.append([-(Dia/2), (Dia/2), -self.airgap_side - self.airgap_both])
        vertices1.append([-(self.dim_D8/2), (self.dim_D8/2), -self.airgap_side - self.airgap_both])
        vertices1.append([(self.dim_D8/2), (self.dim_D8/2), -self.airgap_side - self.airgap_both])
        vertices1.append([(Dia/2), (Dia/2), -self.airgap_side - self.airgap_both])
        vertices1.append([self.dim_D1 / 2, (Dia - (self.dim_D1/2)), -self.airgap_side - self.airgap_both])
        vertices1.append([self.dim_D1 / 2, -(Dia - (self.dim_D1/2)), -self.airgap_side - self.airgap_both])
        vertices1.append([(Dia/2), -(Dia/2), -self.airgap_side - self.airgap_both])
        vertices1.append([(self.dim_D8/2), -(self.dim_D8/2), -self.airgap_side - self.airgap_both])
        vertices1.append([-(self.dim_D8/2), -(self.dim_D8/2), -self.airgap_side - self.airgap_both])
        vertices1.append([-(Dia/2), -(Dia/2), -self.airgap_side - self.airgap_both])
        vertices1.append([-self.dim_D1 / 2, -(Dia - (self.dim_D1/2)), -self.airgap_side - self.airgap_both])
        vertices1.append(vertices1[0])

        pointsArray, segmentsArray = self.points_segments_generator(vertices1)
        self.create_polyline(pointsArray, segmentsArray, 'RM_Core_Bottom')
        self.sweep_along_vector('RM_Core_Bottom', -(self.dim_D4/2) + self.airgap_side)

        self.create_cylinder(0, 0, -(self.dim_D5/2) - self.airgap_both,
                             self.dim_D2, self.dim_D5 / 2, self.segments_number, 'XCyl1')
        self.subtract('RM_Core_Bottom', 'XCyl1')

        self.create_cylinder(0, 0, -(self.dim_D5/2) - self.airgap_both,
                             self.dim_D3, self.dim_D5 / 2 - self.airgap_center, self.segments_number, 'XCyl2')

        self.unite('RM_Core_Bottom,XCyl2')
        if self.dim_D6 != 0:
            self.create_cylinder(0, 0, -(self.dim_D4/2) - self.airgap_both,
                                 self.dim_D6, self.dim_D4 / 2, self.segments_number, 'XCyl3')

            self.subtract('RM_Core_Bottom', 'XCyl3')

        self.duplicate_mirror('RM_Core_Bottom', 0, 0, 1)

        self.rename('RM_Core_Bottom_1', 'RM_Core_Top')
        self.oEditor.FitAll()

        self.draw_wdg(self.dim_D2, self.dim_D3, self.dim_D5 / 2 + self.airgap_both, self.segmentation_angle.Value)


# depending on bobbin we can inherit bobbin, createturn, drawwdg from PQCore. speak with JMark
class EPCore(PQCore):
    def __init__(self, args_list):
        super(EPCore, self).__init__(args_list)

        self.MECoreLength = self.dim_D1
        self.MECoreWidth = self.dim_D6
        self.MECoreHeight = self.dim_D4/2
        self.MSLegWidth = (self.dim_D1 - self.dim_D2)/2
        self.MCLegWidth = self.dim_D3
        self.MSlotDepth = self.dim_D5/2

    def draw_geometry(self, core_type='EP'):
        self.create_box(-(self.MECoreLength/2), -(self.MECoreWidth/2), -self.MECoreHeight - self.airgap_both,
                        self.MECoreLength, self.MECoreWidth, self.MECoreHeight - self.airgap_side,
                        core_type + '_Core_Bottom')

        self.create_cylinder(0, (self.dim_D6/2) - self.dim_D7, -self.MSlotDepth - self.airgap_both,
                             self.dim_D2, self.MSlotDepth, self.segments_number, 'XCyl1')

        self.create_box(-self.dim_D2 / 2, (self.dim_D6/2) - self.dim_D7, -self.MSlotDepth - self.airgap_both,
                        self.dim_D2, self.dim_D7, self.MSlotDepth, 'Box2')

        self.unite('Box2,XCyl1')
        self.subtract(core_type + '_Core_Bottom', 'Box2')

        self.create_cylinder(0, (self.dim_D6/2) - self.dim_D7, -self.MSlotDepth - self.airgap_both,
                             self.dim_D3, self.MSlotDepth, self.segments_number, 'XCyl2')

        self.unite(core_type + '_Core_Bottom,XCyl2')
        self.move(core_type + '_Core_Bottom', 0, (-self.dim_D6/2.0) + self.dim_D7, 0)

        self.duplicate_mirror(core_type + '_Core_Bottom', 0, 0, 1)
        self.rename(core_type + '_Core_Bottom_1', core_type + '_Core_Top')

        self.oEditor.FitAll()
        self.draw_wdg(self.dim_D2, self.dim_D3, (self.dim_D5 / 2 + self.airgap_both), self.segmentation_angle.Value)


class PCore(PQCore):
    def draw_geometry(self, core_name):
        if core_name == 'PH':
            self.dim_D4 *= 2
            self.dim_D5 *= 2

        self.create_cylinder(0, 0, -(self.dim_D4/2) - self.airgap_both, self.dim_D1,
                             (self.dim_D4/2) - self.airgap_side, self.segments_number, core_name + '_Core_Bottom')

        self.create_cylinder(0, 0, -(self.dim_D5/2) - self.airgap_both,
                             self.dim_D2, self.dim_D5 / 2, self.segments_number, 'XCyl1')
        self.subtract(core_name + '_Core_Bottom', 'XCyl1')

        self.create_cylinder(0, 0, -(self.dim_D5/2) - self.airgap_both,
                             self.dim_D3, self.dim_D5 / 2 - self.airgap_center, self.segments_number, 'XCyl2')
        self.unite(core_name + '_Core_Bottom,XCyl2')

        if self.dim_D6 != 0:
            self.create_cylinder(0, 0, -(self.dim_D4/2),
                                 self.dim_D6, self.dim_D4/2, self.segments_number, 'Tool')
            self.subtract(core_name + '_Core_Bottom', 'Tool')

        if self.dim_D7 != 0:
            self.create_box(-self.dim_D1 / 2, -self.dim_D7 / 2, -(self.dim_D4/2) - self.airgap_both,
                            (self.dim_D1 - self.dim_D8) / 2, self.dim_D7, self.dim_D4 / 2, 'Slot1')

            self.create_box(self.dim_D1 / 2, -self.dim_D7 / 2, -(self.dim_D4/2) - self.airgap_both,
                            -(self.dim_D1 - self.dim_D8) / 2, self.dim_D7, self.dim_D4 / 2, 'Slot2')

            self.subtract(core_name + '_Core_Bottom', 'Slot1,Slot2')

        self.duplicate_mirror(core_name + '_Core_Bottom', 0, 0, 1)
        self.rename(core_name + '_Core_Bottom_1', core_name + '_Core_Top')

        if core_name == 'PT':
            self.create_box(-self.dim_D1 / 2, -self.dim_D1 / 2, (self.dim_D4/2) + self.airgap_both,
                            (self.dim_D1 - self.dim_D8) / 2, self.dim_D1, -self.dim_D4 / 2, 'Slot3')

            self.create_box(self.dim_D1 / 2, -self.dim_D1 / 2, (self.dim_D4/2) + self.airgap_both,
                            -(self.dim_D1 - self.dim_D8) / 2, self.dim_D1, -self.dim_D4 / 2, 'Slot4')

            self.subtract(core_name + '_Core_Top', 'Slot3,Slot4')

        self.oEditor.FitAll()
        self.draw_wdg(self.dim_D2, self.dim_D3, (self.dim_D5 / 2 + self.airgap_both), self.segmentation_angle.Value)
