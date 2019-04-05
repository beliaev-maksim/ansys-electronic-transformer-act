# already done: E, EI, ETD = EC = EQ = ER(with argument), EFD, EP, PQ, P=PH=PT(with    argument), RM, U, UI

# super class Cores
# all sub classes for core creation will inherit from here basic parameters from GUI and
# take all functions for primitive creation like createBox, createPolyhedron, move, rename and so on
import math


class Cores(Step1, Step2):
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

        self.AirGapS = 0
        self.AirGapC = 0
        self.TAirGap = 0

        if self.define_airgap.Value:
            if self.airgap_on_leg.Value == 'Center':
                self.AirGapC = self.airgap_value.Value/2.0
            elif self.airgap_on_leg.Value == 'Side':
                self.AirGapS = self.airgap_value.Value/2.0
            else:
                self.TAirGap = self.airgap_value.Value/2.0

        self.WdgParDict = {}

        if self.conductor_type.Value == 'Rectangular':
            xml_path_to_table = 'winding_properties/draw_winding/conductor_type/table_layers'
            row_num = self.table_layers.RowCount

            for rowIndex in range(0, row_num):
                self.WdgParDict[rowIndex + 1] = []
                self.WdgParDict[rowIndex + 1].append(float(self.table_layers.Value[xml_path_to_table +
                                                           "/conductor_width"][rowIndex]))
                self.WdgParDict[rowIndex + 1].append(float(self.table_layers.Value[xml_path_to_table +
                                                           "/conductor_height"][rowIndex]))
                self.WdgParDict[rowIndex + 1].append(int(self.table_layers.Value[xml_path_to_table +
                                                         "/turns_number"][rowIndex]))
                self.WdgParDict[rowIndex + 1].append(float(self.table_layers.Value[xml_path_to_table +
                                                           "/insulation_thickness"][rowIndex]))
        else:
            xml_path_to_table = 'winding_properties/draw_winding/conductor_type/table_layers_circles'
            row_num = self.table_layers_circles.RowCount
            for rowIndex in range(0, row_num):
                self.WdgParDict[rowIndex + 1] = []
                self.WdgParDict[rowIndex + 1].append(float(self.table_layers_circles.Value[xml_path_to_table +
                                                           "/conductor_diameter"][rowIndex]))
                self.WdgParDict[rowIndex + 1].append(int(self.table_layers_circles.Value[xml_path_to_table +
                                                         "/segments_number"][rowIndex]))
                self.WdgParDict[rowIndex + 1].append(int(self.table_layers_circles.Value[xml_path_to_table +
                                                         "/turns_number"][rowIndex]))
                self.WdgParDict[rowIndex + 1].append(float(self.table_layers_circles.Value[xml_path_to_table +
                                                           "/insulation_thickness"][rowIndex]))

    def create_polyline(self, points, segments, name, closed=True, color='(165 42 42)'):
        self.oEditor.CreatePolyline(
            [
                "NAME:PolylineParameters",
                "IsPolylineCovered:=", True,
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
            points_array.append([
                        "NAME:PLPoint",
                        "X:=", str(points_list[i][0]) + 'mm',
                        "Y:=", str(points_list[i][1]) + 'mm',
                        "Z:=", str(points_list[i][2]) + 'mm'
                    ])
            if i != len(points_list) - 1:
                segments_array.append([
                            "NAME:PLSegment",
                            "SegmentType:=", "Line",
                            "StartIndex:=", i,
                            "NoOfPoints:=", 2
                        ])
        return points_array, segments_array


class ECore(Cores):
    def initCore(self):
        self.MECoreLength = self.dim_D1
        self.MECoreWidth = self.dim_D6
        self.MECoreHeight = self.dim_D4
        self.MSLegWidth = (self.dim_D1 - self.dim_D2)/2
        self.MCLegWidth = self.dim_D3
        self.MSlotDepth = self.dim_D5

    def draw_wdg(self, dim_D2, dim_D3, dim_D5, dim_D6, SAng, ECoredim_D8=0.0):
        MLSpacing = self.layer_spacing.Value
        Mtop_margin = self.top_margin.Value
        Mside_margin = self.side_margin.Value
        MBobbinThk = self.bobbin_board_thickness.Value
        WdgParDict = self.WdgParDict
        MSlotHeight = dim_D5*2

        if self.include_bobbin.Value and self.layer_type.Value == "Wound":
            self.draw_bobbin(MSlotHeight - 2*Mtop_margin, (dim_D2/2.0), (dim_D3/2.0) + MBobbinThk,
                            (dim_D6/2.0) + MBobbinThk, MBobbinThk, ECoredim_D8)

        if self.layer_type.Value == "Planar":
            MTDx = Mtop_margin
            for MAx in WdgParDict:
                if self.include_bobbin.Value:
                    self.draw_board(MSlotHeight - 2 * Mtop_margin, (dim_D2/2.0), (dim_D3/2.0) + MBobbinThk,
                                    (dim_D6/2.0) + MBobbinThk, MBobbinThk, -dim_D5 + MTDx, MAx, ECoredim_D8)

                for layerTurn in range(0, int(WdgParDict[MAx][2])):
                    rectangle_size_x = (dim_D3 + 2*Mside_margin + ((2*layerTurn + 1)*WdgParDict[MAx][0]) +
                                        (2*layerTurn*WdgParDict[MAx][3]) + WdgParDict[MAx][3])

                    rectangle_size_y = (dim_D6 + 2*Mside_margin + ((2*layerTurn + 1)*WdgParDict[MAx][0]) +
                                        (2*layerTurn*WdgParDict[MAx][3]) + WdgParDict[MAx][3])

                    rectangle_size_z = - WdgParDict[MAx][1]/2.0

                    self.create_single_turn(rectangle_size_x, rectangle_size_y, rectangle_size_z, WdgParDict[MAx][0],
                                            WdgParDict[MAx][1], (-dim_D5 + MTDx + MBobbinThk + WdgParDict[MAx][1]), MAx,
                                            layerTurn, (rectangle_size_x - dim_D3)/2, SAng)

                MTDx += MLSpacing + WdgParDict[MAx][1] + MBobbinThk

        else:
            # ---- Wound transformer ---- #
            MTDx = Mside_margin + MBobbinThk
            for MAx in WdgParDict.keys():
                for MBx in range(0, int(WdgParDict[MAx][2])):
                    rectangle_size_x = dim_D3 + (2*(MTDx + (WdgParDict[MAx][0]/2.0))) + 2*WdgParDict[MAx][3]
                    rectangle_size_y = dim_D6 + (2*(MTDx + (WdgParDict[MAx][0]/2.0))) + 2*WdgParDict[MAx][3]
                    if self.conductor_type.Value == 'Rectangular':
                        rectangle_size_z = -WdgParDict[MAx][1]/2.0
                        ZProf = (dim_D5 - Mtop_margin - MBobbinThk - (WdgParDict[MAx][3]) -
                                         MBx*(2*WdgParDict[MAx][3] + WdgParDict[MAx][1]))

                    else:
                        rectangle_size_z = -WdgParDict[MAx][0]/2.0
                        ZProf = (dim_D5 - Mtop_margin - MBobbinThk - WdgParDict[MAx][3] -
                                         MBx*(2*WdgParDict[MAx][3] + WdgParDict[MAx][0]))

                    self.create_single_turn(rectangle_size_x, rectangle_size_y, rectangle_size_z, WdgParDict[MAx][0],
                                            WdgParDict[MAx][1], ZProf, MAx, MBx, (rectangle_size_x - dim_D3)/2, SAng)

                MTDx += MLSpacing + WdgParDict[MAx][0] + 2*WdgParDict[MAx][3]

    def create_single_turn(self, PathX, PathY, PathZ, ProfAX, ProfZ, ZPos, LayNum, TurnNum, FRad, SAng):
        segments_number = 12 if SAng == 0 else int(90/SAng)*2
        SegAng = math.pi/(segments_number*2)

        vertices = []
        for i in range(1000):
            vertices.append([])

        for cnt2 in range(0, segments_number - 1):
            vertices[2 + cnt2] = [-(PathX/2.0) + FRad - FRad*math.sin(SegAng*(cnt2 + 1))]
            vertices[2 + cnt2].append(-(PathY/2.0) + FRad - FRad*math.cos(SegAng*(cnt2 + 1)))
            vertices[2 + cnt2].append(PathZ)

        for cnt3 in range(0, segments_number - 1):
            vertices[3 + segments_number + cnt3] = [-(PathX/2.0) + FRad - FRad*math.cos(SegAng*(cnt3 + 1))]
            vertices[3 + segments_number + cnt3].append((PathY/2.0) - FRad + FRad*math.sin(SegAng*(cnt3 + 1)))
            vertices[3 + segments_number + cnt3].append(PathZ)

        for cnt4 in range(0, segments_number - 1):
            vertices[4 + 2*segments_number + cnt4] = [(PathX/2.0) - FRad + FRad*math.sin(SegAng*(cnt4 + 1))]
            vertices[4 + 2*segments_number + cnt4].append((PathY/2.0) - FRad + FRad*math.cos(SegAng*(cnt4 + 1)))
            vertices[4 + 2*segments_number + cnt4].append(PathZ)

        for cnt5 in range(0, segments_number - 1):
            vertices[5 + 3*segments_number + cnt5] = [(PathX/2.0) - FRad + FRad*math.cos(SegAng*(cnt5 + 1))]
            vertices[5 + 3*segments_number + cnt5].append(-(PathY/2.0) + FRad - FRad*math.sin(SegAng*(cnt5 + 1)))
            vertices[5 + 3*segments_number + cnt5].append(PathZ)

        vertices[0] = [(PathX/2.0) - FRad]
        vertices[0].append(-(PathY/2.0))
        vertices[0].append(PathZ)
        vertices[1] = [-(PathX/2.0) + FRad]
        vertices[1].append(-(PathY/2.0))
        vertices[1].append(PathZ)
        vertices[1 + segments_number] = [-(PathX/2.0)]
        vertices[1 + segments_number].append(-(PathY/2.0) + FRad)
        vertices[1 + segments_number].append(PathZ)
        vertices[2 + segments_number] = [-(PathX/2.0)]
        vertices[2 + segments_number].append((PathY/2.0) - FRad)
        vertices[2 + segments_number].append(PathZ)
        vertices[2 + 2*segments_number] = [-(PathX/2.0) + FRad]
        vertices[2 + 2*segments_number].append((PathY/2.0))
        vertices[2 + 2*segments_number].append(PathZ)
        vertices[3 + 2*segments_number] = [(PathX/2.0) - FRad]
        vertices[3 + 2*segments_number].append((PathY/2.0))
        vertices[3 + 2*segments_number].append(PathZ)
        vertices[3 + 3*segments_number] = [(PathX/2.0)]
        vertices[3 + 3*segments_number].append((PathY/2.0) - FRad)
        vertices[3 + 3*segments_number].append(PathZ)
        vertices[4 + 3*segments_number] = [(PathX/2.0)]
        vertices[4 + 3*segments_number].append(-(PathY/2.0) + FRad)
        vertices[4 + 3*segments_number].append(PathZ)
        vertices[4 + 4*segments_number] = [(PathX/2.0) - FRad]
        vertices[4 + 4*segments_number].append(-(PathY/2.0))
        vertices[4 + 4*segments_number].append(PathZ)

        pointsArray, segmentsArray = self.points_segments_generator(vertices)
        self.create_polyline(pointsArray, segmentsArray, "Polyline1", False)

        if self.conductor_type.Value == 'Rectangular':
            self.create_rectangle((PathX/2 - ProfAX/2), 0, (PathZ - ProfZ/2),
                                  ProfAX, ProfZ, 'sweepProfile')
        else:
            self.create_circle(PathX/2, 0, PathZ,
                               ProfAX, ProfZ, 'sweepProfile', 'Y')

        self.rename("Polyline1", 'Tool%s_%s' % (LayNum, TurnNum + 1))

        self.sweep_along_path('sweepProfile,Tool%s_%s' % (LayNum, TurnNum + 1))

        self.move('sweepProfile', 0, 0, ZPos)
        self.rename('sweepProfile', 'Layer%s_%s' % (LayNum, TurnNum + 1))

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

    def draw_board(self, Hb, Db1, Db2, Db3, Tb, ZStart, LayNum, e_coredim_d8=0.0):
        self.create_box(-Db1 + e_coredim_d8, -(Db1 - Db2 + Db3), ZStart,
                        2 * Db1 - 2 * e_coredim_d8, 2 * (Db1 - Db2 + Db3), Tb, 'Board_{}'.format(LayNum), '(0, 128, 0)')

        self.create_box(-Db2 + Tb, -Db3 + Tb, ZStart,
                        2 * Db2 - 2 * Tb, 2 * Db3 - 2 * Tb, Hb, 'BoardSlot_{}'.format(LayNum))
        self.subtract('Board_{}'.format(LayNum), 'BoardSlot_{}'.format(LayNum))

    def draw_geometry(self):
        self.initCore()
        self.create_box(-(self.MECoreLength/2), -(self.MECoreWidth/2), -self.MECoreHeight - self.TAirGap,
                        self.MECoreLength, self.MECoreWidth, (self.MECoreHeight - self.MSlotDepth), 'E_Core_Bottom')

        self.create_box(-(self.MECoreLength/2), -(self.MECoreWidth/2), -self.MECoreHeight - self.TAirGap,
                        self.MSLegWidth, self.MECoreWidth, self.MECoreHeight - self.AirGapS, 'Leg1')

        self.create_box(-(self.MCLegWidth/2), -(self.MECoreWidth/2), -self.MECoreHeight - self.TAirGap,
                        self.MCLegWidth, self.MECoreWidth, self.MECoreHeight - self.AirGapC, 'Leg2')

        self.create_box((self.MECoreLength/2) - self.MSLegWidth, -(self.MECoreWidth/2),
                        -self.MECoreHeight - self.TAirGap,
                        self.MSLegWidth, self.MECoreWidth, self.MECoreHeight - self.AirGapS, 'Leg3')

        self.unite('E_Core_Bottom,Leg1,Leg2,Leg3')

        self.fillet('E_Core_Bottom', self.dim_D7, -self.dim_D1/2, 0, -self.dim_D4 - self.TAirGap)    # outer edges D_7
        self.fillet('E_Core_Bottom', self.dim_D7, self.dim_D1/2, 0, -self.dim_D4 - self.TAirGap)    # outer edges D_7
        self.fillet('E_Core_Bottom', self.dim_D8, -self.dim_D2/2, 0, -self.dim_D5 - self.TAirGap)    # inner edges D_8
        self.fillet('E_Core_Bottom', self.dim_D8, self.dim_D2/2, 0, -self.dim_D5 - self.TAirGap)    # inner edges D_8

        self.duplicate_mirror('E_Core_Bottom', 0, 0, 1)

        self.rename('E_Core_Bottom_1', 'E_Core_Top')
        self.oEditor.FitAll()
        if self.draw_winding.Value:
            self.draw_wdg(self.dim_D2, self.dim_D3, self.dim_D5, self.dim_D6, self.segmentation_angle.Value,
                          self.dim_D8)


# EICore inherit from ECore functions DrawWdg, CreateSingleTurn, draw_bobbin
class EICore(ECore):
    def draw_geometry(self):
        self.initCore()

        self.create_box(-(self.MECoreLength/2), -(self.MECoreWidth/2), -self.MECoreHeight,
                        self.MECoreLength, self.MECoreWidth, (self.MECoreHeight - self.MSlotDepth), 'E_Core')

        self.create_box(-(self.MECoreLength/2), -(self.MECoreWidth/2), -self.MECoreHeight,
                        self.MSLegWidth, self.MECoreWidth, self.MECoreHeight - 2 * self.AirGapS, 'Leg1')

        self.create_box(-(self.MCLegWidth/2), -(self.MECoreWidth/2), -self.MECoreHeight,
                        self.MCLegWidth, self.MECoreWidth, self.MECoreHeight - 2 * self.AirGapC, 'Leg2')

        self.create_box((self.MECoreLength/2) - self.MSLegWidth, -(self.MECoreWidth/2), -self.MECoreHeight,
                        self.MSLegWidth, self.MECoreWidth, self.MECoreHeight - 2 * self.AirGapS, 'Leg3')

        self.unite('E_Core,Leg1,Leg2,Leg3')

        self.fillet('E_Core', self.dim_D7, -self.dim_D1/2, -self.dim_D6/2, -(self.dim_D4/2) - 2*self.TAirGap)
        self.fillet('E_Core', self.dim_D7, -self.dim_D1/2, self.dim_D6/2, -(self.dim_D4/2) - 2*self.TAirGap)
        self.fillet('E_Core', self.dim_D7, self.dim_D1/2, -self.dim_D6/2, -(self.dim_D4/2) - 2*self.TAirGap)
        self.fillet('E_Core', self.dim_D7, self.dim_D1/2, self.dim_D6/2, -(self.dim_D4/2) - 2*self.TAirGap)

        # two multiply air gap due to no symmetry for cores like in ECore
        self.create_box(-self.dim_D1/2, -self.dim_D6/2, 2 * self.TAirGap,
                        self.dim_D1, self.dim_D6, self.dim_D8, 'I_Core')

        self.fillet('I_Core', self.dim_D7, -self.dim_D1/2, -self.dim_D6/2, (self.dim_D8/2) - 2*self.TAirGap)
        self.fillet('I_Core', self.dim_D7, -self.dim_D1/2, self.dim_D6/2, (self.dim_D8/2) - 2*self.TAirGap)
        self.fillet('I_Core', self.dim_D7, self.dim_D1/2, -self.dim_D6/2, (self.dim_D8/2) - 2*self.TAirGap)
        self.fillet('I_Core', self.dim_D7, self.dim_D1/2, self.dim_D6/2, (self.dim_D8/2) - 2*self.TAirGap)

        self.move('E_Core,I_Core', 0, 0, self.dim_D5/2)

        self.oEditor.FitAll()
        if self.draw_winding.Value:
            self.draw_wdg(self.dim_D2, self.dim_D3, self.dim_D5/2 + 2 * self.TAirGap, self.dim_D6,
                          self.segmentation_angle.Value)


class EFDCore(ECore):
    def draw_geometry(self):
        self.initCore()
        self.create_box(-(self.MECoreLength/2), -(self.MECoreWidth/2), -self.MECoreHeight - self.TAirGap,
                        self.MECoreLength, self.MECoreWidth, (self.MECoreHeight - self.MSlotDepth), 'EFD_Core_Bottom')

        self.create_box(-(self.MECoreLength/2), -(self.MECoreWidth/2), -self.MECoreHeight - self.TAirGap,
                        self.MSLegWidth, self.MECoreWidth, self.MECoreHeight - self.AirGapS, 'Leg1')

        self.create_box(-(self.MCLegWidth/2), -(self.MECoreWidth/2) - self.dim_D8, -self.MECoreHeight - self.TAirGap,
                        self.MCLegWidth, self.dim_D7, self.MECoreHeight - self.AirGapC, 'Leg2')

        self.create_box((self.MECoreLength/2) - self.MSLegWidth, -(self.MECoreWidth/2),
                        -self.MECoreHeight - self.TAirGap,
                        self.MSLegWidth, self.MECoreWidth, self.MECoreHeight - self.AirGapS, 'Leg3')

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

        if self.draw_winding.Value:
                        self.draw_wdg(self.dim_D2, self.dim_D3, self.dim_D5, self.dim_D7, self.segmentation_angle.Value)


# UCore inherit from ECore functions CreateSingleTurn, draw_bobbin, drawBoard
class UCore(ECore):
    def initCore(self):
        self.MECoreLength = self.dim_D1
        self.MECoreWidth = self.dim_D5
        self.MECoreHeight = self.dim_D3

    def draw_wdg(self, dim_D1, dim_D2, dim_D3, dim_D4, dim_D5, SAng):
        MLSpacing = self.layer_spacing.Value
        Mtop_margin = self.top_margin.Value
        Mside_margin = self.side_margin.Value
        MBobbinThk = self.bobbin_board_thickness.Value
        WdgParDict = self.WdgParDict
        MSlotHeight = dim_D4*2
        LegWidth = (dim_D1 - dim_D2)/2

        if self.include_bobbin.Value and self.layer_type.Value == "Wound":
            self.draw_bobbin(MSlotHeight - 2*Mtop_margin, dim_D2 + LegWidth/2, LegWidth/2.0 + MBobbinThk,
                            (dim_D5/2.0) + MBobbinThk, MBobbinThk)

        if self.layer_type.Value == "Planar":
            MTDx = Mtop_margin
            for MAx in self.WdgParDict:
                if self.include_bobbin.Value:
                    self.draw_board(MSlotHeight - 2 * Mtop_margin, dim_D2 + LegWidth/2, LegWidth/2.0 + MBobbinThk,
                                    (dim_D5/2.0) + MBobbinThk, MBobbinThk, -dim_D4 + MTDx, MAx)

                for layerTurn in range(0, int(self.WdgParDict[MAx][2])):
                    rectangle_size_x = (LegWidth + 2*Mside_margin + ((2*layerTurn + 1)*self.WdgParDict[MAx][0]) +
                                        (2*layerTurn*2*self.WdgParDict[MAx][3]) + 2*self.WdgParDict[MAx][3])

                    rectangle_size_y = (dim_D5 + 2*Mside_margin + ((2*layerTurn + 1)*self.WdgParDict[MAx][0]) +
                                        (2*layerTurn*2*self.WdgParDict[MAx][3]) + 2*self.WdgParDict[MAx][3])

                    rectangle_size_z = -self.WdgParDict[MAx][1]/2.0
                    self.create_single_turn(rectangle_size_x, rectangle_size_y, rectangle_size_z,
                                            self.WdgParDict[MAx][0], self.WdgParDict[MAx][1],
                                            (-dim_D4 + MTDx + MBobbinThk + WdgParDict[MAx][1]), MAx,
                                            layerTurn, (rectangle_size_x - LegWidth)/2, SAng)

                MTDx += MLSpacing + self.WdgParDict[MAx][1] + MBobbinThk

        else:
            # ---- Wound transformer ---- #
            MTDx = Mside_margin + MBobbinThk
            for MAx in self.WdgParDict.keys():
                for MBx in range(0, int(self.WdgParDict[MAx][2])):
                    rectangle_size_x = LegWidth + (2*(MTDx + (self.WdgParDict[MAx][0]/2.0))) + 2*self.WdgParDict[MAx][3]
                    rectangle_size_y = dim_D5 + (2*(MTDx + (self.WdgParDict[MAx][0]/2.0))) + 2*self.WdgParDict[MAx][3]
                    if self.conductor_type.Value == 'Rectangular':
                        rectangle_size_z = -self.WdgParDict[MAx][1]/2.0
                        ZProf = (dim_D4 - Mtop_margin - MBobbinThk - (self.WdgParDict[MAx][3]) -
                                         MBx*(2*self.WdgParDict[MAx][3] + self.WdgParDict[MAx][1]))

                        self.create_single_turn(rectangle_size_x, rectangle_size_y, rectangle_size_z,
                                                self.WdgParDict[MAx][0], self.WdgParDict[MAx][1], ZProf, MAx, MBx,
                                                (rectangle_size_x - LegWidth)/2, SAng)
                    else:
                        rectangle_size_z = -self.WdgParDict[MAx][0]/2.0
                        ZProf = (dim_D4 - Mtop_margin - MBobbinThk - (self.WdgParDict[MAx][3]) -
                                         MBx*(2*self.WdgParDict[MAx][3] + self.WdgParDict[MAx][0]))

                        self.create_single_turn(rectangle_size_x, rectangle_size_y, rectangle_size_z,
                                                self.WdgParDict[MAx][0], self.WdgParDict[MAx][1], ZProf, MAx, MBx,
                                                (rectangle_size_x - LegWidth)/2, SAng)

                MTDx = MTDx + MLSpacing + self.WdgParDict[MAx][0] + 2*self.WdgParDict[MAx][3]

    def draw_geometry(self):
        self.initCore()
        self.create_box(-(self.dim_D1 - self.dim_D2) / 4, -(self.MECoreWidth/2), -self.MECoreHeight - self.TAirGap,
                        self.MECoreLength, self.MECoreWidth, self.MECoreHeight, 'U_Core_Bottom')

        self.create_box((self.dim_D1 - self.dim_D2) / 4, -self.MECoreWidth/2, -self.dim_D4 - self.TAirGap,
                        self.dim_D2, self.MECoreWidth, self.dim_D4, 'XSlot')

        self.subtract('U_Core_Bottom', 'XSlot')

        if self.AirGapC > 0:
            self.create_box(-(self.dim_D1 - self.dim_D2) / 4, -self.MECoreWidth/2, -self.AirGapC,
                            (self.dim_D1 - self.dim_D2)/2, self.MECoreWidth, self.AirGapC, 'AgC')

            self.subtract('U_Core_Bottom', 'AgC')

        if self.AirGapS > 0:
            self.create_box(self.dim_D2 + (self.dim_D1 - self.dim_D2) / 4, -self.MECoreWidth/2, -self.AirGapS,
                            (self.dim_D1 - self.dim_D2)/2, self.MECoreWidth, self.AirGapS, 'AgS')

            self.subtract('U_Core_Bottom', 'AgS')

        self.duplicate_mirror('U_Core_Bottom', 0, 0, 1)
        self.rename('U_Core_Bottom_1', 'U_Core_Top')
        self.oEditor.FitAll()
        if self.draw_winding.Value:
            self.draw_wdg(self.dim_D1, self.dim_D2, self.dim_D3, (self.dim_D4 + self.TAirGap), self.dim_D5,
                          self.segmentation_angle.Value)


# UICore inherit from ECore functions CreateSingleTurn, draw_bobbin, drawBoard
# and from UCore inherit DrawWdg and initCore
class UICore(UCore):
    def draw_geometry(self):
        self.initCore()
        self.create_box(-(self.dim_D1 - self.dim_D2) / 4, -(self.MECoreWidth/2), (self.dim_D4/2) - self.MECoreHeight,
                        self.MECoreLength, self.MECoreWidth, self.MECoreHeight, 'U_Core')

        self.create_box((self.dim_D1 - self.dim_D2) / 4, -self.MECoreWidth/2, -self.dim_D4/2,
                        self.dim_D2, self.MECoreWidth, self.dim_D4, 'XSlot')

        self.subtract('U_Core', 'XSlot')

        if self.AirGapC > 0:
            self.create_box((self.dim_D1 - self.dim_D2) / 4, -self.MECoreWidth/2, (self.dim_D4/2) - 2 * self.AirGapC,
                            -(self.dim_D1 - self.dim_D2)/2, self.MECoreWidth, 2 * self.AirGapC, 'AgC')
            self.subtract('U_Core', 'AgC')

        if self.AirGapS > 0:
            self.create_box(self.dim_D2 + (self.dim_D1 - self.dim_D2) / 4, -self.MECoreWidth/2,
                            (self.dim_D4/2) - 2 * self.AirGapS,
                            (self.dim_D1 - self.dim_D2)/2, self.MECoreWidth, 2 * self.AirGapS, 'AgS')
            self.subtract('U_Core', 'AgS')

        self.create_box(-(self.dim_D1 - self.dim_D2) / 4 + (self.dim_D1 - self.dim_D6)/2, -self.dim_D7/2,
                        (self.dim_D4/2) + 2 * self.TAirGap,
                        self.dim_D6, self.dim_D7, self.dim_D8, 'I_Core')

        self.oEditor.FitAll()
        if self.draw_winding.Value:
            self.draw_wdg(self.dim_D1, self.dim_D2, self.dim_D3/2, (self.dim_D4/2) + 2 * self.TAirGap, 
                          self.dim_D5, self.segmentation_angle.Value)


class PQCore(Cores):
    def initCore(self):
        self.MECoreLength = self.dim_D1
        self.MECoreWidth = self.dim_D6
        self.MECoreHeight = self.dim_D4
        self.MSLegWidth = (self.dim_D1 - self.dim_D2)/2
        self.MCLegWidth = self.dim_D3
        self.MSlotDepth = self.dim_D5

    def draw_wdg(self, dim_D2, dim_D3, dim_D5, SAng):
        MLSpacing = self.layer_spacing.Value
        Mtop_margin = self.top_margin.Value
        Mside_margin = self.side_margin.Value
        MBobbinThk = self.bobbin_board_thickness.Value
        WdgParDict = self.WdgParDict
        MSlotHeight = dim_D5*2

        if self.include_bobbin.Value and self.layer_type.Value == "Wound":
            self.draw_bobbin(MSlotHeight - 2*Mtop_margin, (dim_D2/2.0), (dim_D3/2.0) + MBobbinThk, MBobbinThk)

        if self.layer_type.Value == "Planar":
            MTDx = Mtop_margin
            for MAx in self.WdgParDict:
                if self.include_bobbin.Value:
                    self.drawBoard(MSlotHeight - 2*Mtop_margin, (dim_D2/2.0), (dim_D3/2.0) + MBobbinThk,
                                   MBobbinThk, -dim_D5 + MTDx, MAx)

                for layerTurn in range(0, int(self.WdgParDict[MAx][2])):
                    rectangle_size_x = (dim_D3 + 2*Mside_margin + ((2*layerTurn + 1)*self.WdgParDict[MAx][0]) +
                                        (2*layerTurn*2*self.WdgParDict[MAx][3]) + 2*self.WdgParDict[MAx][3])

                    rectangle_size_z = -self.WdgParDict[MAx][1]/2.0

                    self.create_single_turn(rectangle_size_x, rectangle_size_z, self.WdgParDict[MAx][0], 
                                            self.WdgParDict[MAx][1], -dim_D5 + MTDx + MBobbinThk + WdgParDict[MAx][1], 
                                            MAx, layerTurn, SAng)

                MTDx += MLSpacing + self.WdgParDict[MAx][1] + MBobbinThk
        else:
            # ---- Wound transformer ---- #
            MTDx = Mside_margin + MBobbinThk
            for MAx in self.WdgParDict.keys():
                for MBx in range(0, int(self.WdgParDict[MAx][2])):
                    rectangle_size_x = dim_D3 + (2*(MTDx + (self.WdgParDict[MAx][0]/2.0))) + 2*self.WdgParDict[MAx][3]
                    if self.conductor_type.Value == 'Rectangular':
                        rectangle_size_z = -self.WdgParDict[MAx][1]/2.0
                        ZProf = (dim_D5 - Mtop_margin - MBobbinThk - self.WdgParDict[MAx][3] -
                                         MBx*(2*self.WdgParDict[MAx][3] + self.WdgParDict[MAx][1]))

                        self.create_single_turn(rectangle_size_x, rectangle_size_z, self.WdgParDict[MAx][0],
                                                self.WdgParDict[MAx][1], ZProf, MAx, MBx, SAng)
                    else:
                        rectangle_size_z = -self.WdgParDict[MAx][0]/2.0
                        ZProf = (dim_D5 - Mtop_margin - MBobbinThk - self.WdgParDict[MAx][3] -
                                         MBx*(2*self.WdgParDict[MAx][3] + self.WdgParDict[MAx][0]))

                        self.create_single_turn(rectangle_size_x, rectangle_size_z, self.WdgParDict[MAx][0], 
                                                self.WdgParDict[MAx][1], ZProf, MAx, MBx, SAng)

                MTDx = MTDx + MLSpacing + self.WdgParDict[MAx][0] + 2*self.WdgParDict[MAx][3]

    def create_single_turn(self, PathX, PathZ, ProfAX, ProfZ, ZPos, LayNum, TurnNum, SAng):
        segments_number = 0 if SAng == 0 else int(360/SAng)

        self.create_circle(0, 0, PathZ,
                           PathX, segments_number, 'pathLine', 'Z', covered=False)

        if self.conductor_type.Value == 'Rectangular':
            self.create_rectangle((PathX/2 - ProfAX/2), 0, (PathZ - ProfZ/2),
                                  ProfAX, ProfZ, 'sweepProfile')
        else:
            self.create_circle(PathX/2, 0, PathZ,
                               ProfAX, ProfZ, 'sweepProfile', 'Y')

        self.rename('pathLine', 'Tool%s_%s' % (LayNum, TurnNum + 1))
        self.sweep_along_path('sweepProfile,Tool%s_%s' % (LayNum, TurnNum + 1))

        self.move('sweepProfile', 0, 0, ZPos)
        self.rename('sweepProfile', 'Layer%s_%s' % (LayNum, TurnNum + 1))

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

    def drawBoard(self, Hb, Db1, Db2, Tb, ZStart, LayNum):
        self.create_cylinder(0, 0, ZStart,
                             Db1 * 2, Tb, self.segments_number, 'Board_{}'.format(LayNum), '(0, 128, 0)')

        self.create_cylinder(0, 0, ZStart,
                             (Db2 - Tb) * 2, Hb, self.segments_number, 'BoardSlot_{}'.format(LayNum))

        self.subtract('Board_{}'.format(LayNum), 'BoardSlot_{}'.format(LayNum))

    def draw_geometry(self):
        self.initCore()

        self.create_box(-self.dim_D1/2, -self.dim_D8/2, -(self.dim_D4/2) - self.TAirGap,
                        (self.dim_D1 - self.dim_D6)/2, self.dim_D8, (self.dim_D4/2) - self.AirGapS, 'PQ_Core_Bottom')

        IntL = math.sqrt(((self.dim_D3/2) ** 2) - ((self.dim_D7/2) ** 2))
        vertices1 = [[-self.dim_D6/2, - 0.4*self.dim_D8, -(self.dim_D4/2) - self.TAirGap],
                     [-self.dim_D6/2, 0.4*self.dim_D8, -(self.dim_D4/2) - self.TAirGap],
                     [-self.dim_D7/2, +IntL, -(self.dim_D4/2) - self.TAirGap],
                     [-self.dim_D7/2, - IntL, -(self.dim_D4/2) - self.TAirGap]]
        vertices1.append(vertices1[0])

        points_array, segments_array = self.points_segments_generator(vertices1)
        self.create_polyline(points_array, segments_array, "Polyline1")
        self.sweep_along_vector("Polyline1", (self.dim_D4/2) - self.AirGapS)

        self.unite('PQ_Core_Bottom,Polyline1')

        self.create_cylinder(0, 0, -(self.dim_D5/2) - self.TAirGap,
                             self.dim_D2, self.dim_D5/2, self.segments_number, 'XCyl1')

        self.subtract('PQ_Core_Bottom', 'XCyl1')
        self.duplicate_mirror('PQ_Core_Bottom', 1, 0, 0)

        self.create_cylinder(0, 0, -(self.dim_D4/2) - self.TAirGap,
                             self.dim_D3, self.dim_D4/2 - self.AirGapC, self.segments_number, 'XCyl2')

        self.unite('PQ_Core_Bottom,PQ_Core_Bottom_1,XCyl2')
        self.duplicate_mirror('PQ_Core_Bottom', 0, 0, 1)

        self.rename('PQ_Core_Bottom_2', 'PQ_Core_Top')
        self.oEditor.FitAll()

        if self.draw_winding.Value:
            self.draw_wdg(self.dim_D2, self.dim_D3, self.dim_D5/2 + self.TAirGap, self.segmentation_angle.Value)


# ETDCore inherit from PQCore functions DrawWdg, CreateSingleTurn, draw_bobbin
class ETDCore(PQCore):
    def draw_geometry(self, coreName):
        self.initCore()
        self.create_box(-(self.MECoreLength/2), -(self.MECoreWidth/2), -self.MECoreHeight - self.TAirGap,
                        self.MECoreLength, self.MECoreWidth, self.MECoreHeight - self.AirGapS,
                        coreName + '_Core_Bottom')

        self.create_cylinder(0, 0, -self.MSlotDepth - self.TAirGap,
                             self.dim_D2, self.MSlotDepth, self.segments_number, 'XCyl1')

        self.subtract(coreName + '_Core_Bottom', 'XCyl1')

        if coreName == 'ER' and self.dim_D7 > 0:
            self.create_box(-self.dim_D7/2, -self.MECoreWidth/2, -self.MSlotDepth - self.TAirGap,
                            self.dim_D7, self.MECoreWidth, self.MSlotDepth, 'Tool')
            self.subtract(coreName + '_Core_Bottom', 'Tool')

        self.create_cylinder(0, 0, -self.MSlotDepth - self.TAirGap,
                             self.dim_D3, self.MSlotDepth - self.AirGapC, self.segments_number, 'XCyl2')

        self.unite(coreName + '_Core_Bottom,XCyl2')
        self.duplicate_mirror(coreName + '_Core_Bottom', 0, 0, 1)
        self.rename(coreName + '_Core_Bottom_1', coreName + '_Core_Top')

        self.oEditor.FitAll()
        if self.draw_winding.Value:
            self.draw_wdg(self.dim_D2, self.dim_D3, (self.dim_D5 + self.TAirGap), self.segmentation_angle.Value)


# RMCore inherit from PQCore functions DrawWdg, CreateSingleTurn, draw_bobbin
class RMCore(PQCore):
    def draw_geometry(self):
        self.initCore()
        Dia = self.dim_D7/math.sqrt(2)

        vertices1 = [[-self.dim_D1/2, (Dia - (self.dim_D1/2)), -self.AirGapS - self.TAirGap]]
        vertices1.append([-(Dia/2), (Dia/2), -self.AirGapS - self.TAirGap])
        vertices1.append([-(self.dim_D8/2), (self.dim_D8/2), -self.AirGapS - self.TAirGap])
        vertices1.append([(self.dim_D8/2), (self.dim_D8/2), -self.AirGapS - self.TAirGap])
        vertices1.append([(Dia/2), (Dia/2), -self.AirGapS - self.TAirGap])
        vertices1.append([self.dim_D1/2, (Dia - (self.dim_D1/2)), -self.AirGapS - self.TAirGap])
        vertices1.append([self.dim_D1/2, -(Dia - (self.dim_D1/2)), -self.AirGapS - self.TAirGap])
        vertices1.append([(Dia/2), -(Dia/2), -self.AirGapS - self.TAirGap])
        vertices1.append([(self.dim_D8/2), -(self.dim_D8/2), -self.AirGapS - self.TAirGap])
        vertices1.append([-(self.dim_D8/2), -(self.dim_D8/2), -self.AirGapS - self.TAirGap])
        vertices1.append([-(Dia/2), -(Dia/2), -self.AirGapS - self.TAirGap])
        vertices1.append([-self.dim_D1/2, -(Dia - (self.dim_D1/2)), -self.AirGapS - self.TAirGap])
        vertices1.append(vertices1[0])

        pointsArray, segmentsArray = self.points_segments_generator(vertices1)
        self.create_polyline(pointsArray, segmentsArray, 'RM_Core_Bottom')
        self.sweep_along_vector('RM_Core_Bottom', -(self.dim_D4/2) + self.AirGapS)

        self.create_cylinder(0, 0, -(self.dim_D5/2) - self.TAirGap,
                             self.dim_D2, self.dim_D5/2, self.segments_number, 'XCyl1')
        self.subtract('RM_Core_Bottom', 'XCyl1')

        self.create_cylinder(0, 0, -(self.dim_D5/2) - self.TAirGap,
                             self.dim_D3, self.dim_D5/2 - self.AirGapC, self.segments_number, 'XCyl2')

        self.unite('RM_Core_Bottom,XCyl2')
        if self.dim_D6 != 0:
            self.create_cylinder(0, 0, -(self.dim_D4/2) - self.TAirGap,
                                 self.dim_D6, self.dim_D4/2, self.segments_number, 'XCyl3')

            self.subtract('RM_Core_Bottom', 'XCyl3')

        self.duplicate_mirror('RM_Core_Bottom', 0, 0, 1)

        self.rename('RM_Core_Bottom_1', 'RM_Core_Top')
        self.oEditor.FitAll()

        if self.draw_winding.Value:
            self.draw_wdg(self.dim_D2, self.dim_D3, self.dim_D5/2 + self.TAirGap, self.segmentation_angle.Value)


# depending on bobbin we can inherit bobbin, createturn, drawwdg from PQCore. speak with JMark
class EPCore(PQCore):
    def initCore(self):
        self.MECoreLength = self.dim_D1
        self.MECoreWidth = self.dim_D6
        self.MECoreHeight = self.dim_D4/2
        self.MSLegWidth = (self.dim_D1 - self.dim_D2)/2
        self.MCLegWidth = self.dim_D3
        self.MSlotDepth = self.dim_D5/2

    def draw_geometry(self, core_type='EP'):
        self.initCore()
        self.create_box(-(self.MECoreLength/2), -(self.MECoreWidth/2), -self.MECoreHeight - self.TAirGap,
                        self.MECoreLength, self.MECoreWidth, self.MECoreHeight - self.AirGapS,
                        core_type + '_Core_Bottom')

        self.create_cylinder(0, (self.dim_D6/2) - self.dim_D7, -self.MSlotDepth - self.TAirGap,
                             self.dim_D2, self.MSlotDepth, self.segments_number, 'XCyl1')

        self.create_box(-self.dim_D2/2, (self.dim_D6/2) - self.dim_D7, -self.MSlotDepth - self.TAirGap,
                        self.dim_D2, self.dim_D7, self.MSlotDepth, 'Box2')

        self.unite('Box2,XCyl1')
        self.subtract(core_type + '_Core_Bottom', 'Box2')

        self.create_cylinder(0, (self.dim_D6/2) - self.dim_D7, -self.MSlotDepth - self.TAirGap,
                             self.dim_D3, self.MSlotDepth, self.segments_number, 'XCyl2')

        self.unite(core_type + '_Core_Bottom,XCyl2')
        self.move(core_type + '_Core_Bottom', 0, (-self.dim_D6/2.0) + self.dim_D7, 0)

        self.duplicate_mirror(core_type + '_Core_Bottom', 0, 0, 1)
        self.rename(core_type + '_Core_Bottom_1', core_type + '_Core_Top')

        self.oEditor.FitAll()
        if self.draw_winding.Value:
            self.draw_wdg(self.dim_D2, self.dim_D3, (self.dim_D5/2 + self.TAirGap), self.segmentation_angle.Value)


class PCore(PQCore):
    def draw_geometry(self, core_name):
        self.initCore()
        if core_name == 'PH':
            self.dim_D4 *= 2
            self.dim_D5 *= 2

        self.create_cylinder(0, 0, -(self.dim_D4/2) - self.TAirGap, self.dim_D1,
                             (self.dim_D4/2) - self.AirGapS, self.segments_number, core_name + '_Core_Bottom')

        self.create_cylinder(0, 0, -(self.dim_D5/2) - self.TAirGap,
                             self.dim_D2, self.dim_D5/2, self.segments_number, 'XCyl1')
        self.subtract(core_name + '_Core_Bottom', 'XCyl1')

        self.create_cylinder(0, 0, -(self.dim_D5/2) - self.TAirGap,
                             self.dim_D3, self.dim_D5/2 - self.AirGapC, self.segments_number, 'XCyl2')
        self.unite(core_name + '_Core_Bottom,XCyl2')

        if self.dim_D6 != 0:
            self.create_cylinder(0, 0, -(self.dim_D4/2),
                                 self.dim_D6, self.dim_D4/2, self.segments_number, 'Tool')
            self.subtract(core_name + '_Core_Bottom', 'Tool')

        if self.dim_D7 != 0:
            self.create_box(-self.dim_D1/2, -self.dim_D7/2, -(self.dim_D4/2) - self.TAirGap,
                            (self.dim_D1 - self.dim_D8)/2, self.dim_D7, self.dim_D4/2, 'Slot1')

            self.create_box(self.dim_D1/2, -self.dim_D7/2, -(self.dim_D4/2) - self.TAirGap,
                            -(self.dim_D1 - self.dim_D8)/2, self.dim_D7, self.dim_D4/2, 'Slot2')

            self.subtract(core_name + '_Core_Bottom', 'Slot1,Slot2')

        self.duplicate_mirror(core_name + '_Core_Bottom', 0, 0, 1)
        self.rename(core_name + '_Core_Bottom_1', core_name + '_Core_Top')

        if core_name == 'PT':
            self.create_box(-self.dim_D1/2, -self.dim_D1/2, (self.dim_D4/2) + self.TAirGap,
                            (self.dim_D1 - self.dim_D8)/2, self.dim_D1, -self.dim_D4/2, 'Slot3')

            self.create_box(self.dim_D1/2, -self.dim_D1/2, (self.dim_D4/2) + self.TAirGap,
                            -(self.dim_D1 - self.dim_D8)/2, self.dim_D1, -self.dim_D4/2, 'Slot4')

            self.subtract(core_name + '_Core_Top', 'Slot3,Slot4')

        self.oEditor.FitAll()
        if self.draw_winding.Value:
            self.draw_wdg(self.dim_D2, self.dim_D3, (self.dim_D5/2 + self.TAirGap), self.segmentation_angle.Value)
