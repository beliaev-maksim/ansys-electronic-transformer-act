import copy


class Circuit:
    def __init__(self, winding_connection, oProject, design_name, voltage=None, current=None, resistance=None,
                 frequency=None):
        """
        :param winding_connection: dictionary with windings and corresponding connections definition
        :param oProject: aedt project object
        :param design_name: name of circuit design
        :param voltage: voltage to feed the primary side, in Volts
        :param current: current to feed the primary side, in Amps
        :param resistance: resistance for all non primary sides, in Ohm
        :param frequency: frequency for current/voltage source, in Hz
        """
        self.grid_cell_size = 0.00254
        self._id = 1000
        self.page = 1

        self.winding_connection = copy.deepcopy(winding_connection)
        self.project = oProject
        self.design = self.project.InsertDesign("Maxwell Circuit", "circuit_" + design_name, "", "")
        self.editor = self.design.SetActiveEditor("SchematicEditor")

        self.voltage = voltage
        self.current = current
        self.resistance = resistance
        self.frequency = frequency

    @property
    def new_id(self):
        """
            Function to generate new ID for wires and circuit components
        """
        self._id += 1
        return self._id

    @staticmethod
    def run_connection_reduction(connections):
        """
        Reduces connection definition dictionary by unpacking nested elements of the same type.
        For example serial connection of nested serial group is the same as just all elements in series
        :param connections: dictionary with connections definition
        :return:
        """
        def dict_walk(target_dict, conn_type=""):
            conn_type = conn_type[:1]
            for key, val in target_dict.items():
                if isinstance(val, dict):
                    if key[:3] == conn_type:
                        new_dict = target_dict.pop(key)
                        target_dict.update(new_dict)
                        return
                    else:
                        dict_walk(val, key)

        # loop until dictionaries after change are not the same
        dict2 = {}
        while dict2 != connections:
            dict2 = copy.deepcopy(connections)
            dict_walk(connections)

    @staticmethod
    def validate_dict(target_dict):
        """
        We allow users to have single layer per transformer side. To make current script compatible we assign
        this layer to serial connection
        :param target_dict: dictionary where this layer may exist
        :return:
        """
        if len(target_dict) == 1:
            for key, val in target_dict.items():
                if not isinstance(val, dict):
                    target_dict.pop(key)
                    target_dict["S99999"] = {key: "Layer"}

    def create_component(self, name, x, y, angle=0):
        x *= self.grid_cell_size * 4
        y *= self.grid_cell_size * 3
        component_name = self.editor.CreateComponent(
            [
                "NAME:ComponentProps",
                "Name:="	, name,
                "Id:="			, str(self.new_id)
            ],
            [
                "NAME:Attributes",
                "Page:="		, self.page,
                "X:="			, x,
                "Y:="			, y,
                "Angle:="		, angle,
                "Flip:="		, False
            ])

        return component_name

    def change_prop(self, component_name, prop_name, prop_value, netlist_unit=""):
        units = ["NetlistUnits:=", netlist_unit] if netlist_unit else []
        self.editor.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:PassedParameterTab",
                    [
                        "NAME:PropServers",
                        component_name
                    ],
                    [
                        "NAME:ChangedProps",
                        [
                            "NAME:" + prop_name,
                            "OverridingDef:=", True,
                            "Value:="	, str(prop_value) + netlist_unit
                        ] + units
                    ]
                ]
            ])

    def create_winding(self, winding_number, x, y):
        """
        Create a winding component and give it a name the same as in excitation
        :param winding_number: number of the winding to give component a name
        :param x: position in X (unitless)
        :param y: position in y (unitless)
        :return:
        """
        component_name = self.create_component("Maxwell Circuit Elements\\Dedicated Elements:Winding", x, y)
        self.change_prop(component_name, "name", "Layer_" + str(winding_number))

        return component_name

    def create_source_or_load(self, source_type, max_x):
        angle = 0
        if source_type == "Voltage":
            name = "Maxwell Circuit Elements\\Sources:VSin"
            angle = -90
        elif source_type == "Current":
            name = "Maxwell Circuit Elements\\Sources:ISin"
            angle = 90
        else:
            name = "Maxwell Circuit Elements\\Passive Elements:Res"

        component_name = self.create_component(name, x=0, y=-1, angle=angle)

        if source_type == "Voltage":
            self.change_prop(component_name, prop_name="Va", prop_value=self.voltage, netlist_unit="V")
            self.change_prop(component_name, prop_name="VFreq", prop_value=self.frequency, netlist_unit="")
        elif source_type == "Current":
            self.change_prop(component_name, prop_name="Ia", prop_value=self.current, netlist_unit="A")
            self.change_prop(component_name, prop_name="IFreq", prop_value=self.frequency, netlist_unit="")
        else:
            self.change_prop(component_name, prop_name="R", prop_value=self.resistance, netlist_unit="Ohm")

        self.wire(0, -1, 0, 0)
        self.wire(max_x, -1, max_x, 0)
        self.wire(1, -1, max_x, -1)

    def create_ground(self):
        x = -2
        y = -4
        self.editor.CreateGround(
            [
                "NAME:GroundProps",
                "Id:="		, self.new_id
            ],
            [
                "NAME:Attributes",
                "Page:="		, self.page,
                "X:="			, x * self.grid_cell_size,
                "Y:="			, y * self.grid_cell_size,
                "Angle:="		, 0,
                "Flip:="		, False
            ])

    def wire(self, x1, y1, x2, y2):
        """
        Create a wire between two coordinates
        :param x1:
        :param y1:
        :param x2:
        :param y2:
        :return:
        """
        x1 *= self.grid_cell_size * 4
        y1 *= self.grid_cell_size * 3
        x2 *= self.grid_cell_size * 4
        y2 *= self.grid_cell_size * 3

        x1 -= self.grid_cell_size * 2
        x2 -= self.grid_cell_size * 2

        self.editor.CreateWire(
            [
                "NAME:WireData",
                "Name:="	, "",
                "Id:="			, self.new_id,
                "Points:="		, ["({}, {})".format(x1, y1), "({}, {})".format(x2, y2)]
            ],
            [
                "NAME:Attributes",
                "Page:="		, self.page
            ])

    def calc_xy(self, target_dict, conn_type=""):
        """
        Calculate X and Y sizes of connections dictionary. Serial connection adds to X, parallel to Y
        So x_size is a number of serial elements, y_size - parallel elements
        :param target_dict: dictionary with reduced connections
        :param conn_type: used for recursion to pass parent connection type
        :return: x, y coordinates (unitless) calculated for nested circuit
        """
        x = y = 0
        conn_type = conn_type[:1]
        for key, val in target_dict.items():
            if isinstance(val, dict):
                new_x, new_y = self.calc_xy(val, key)

                if "S" in conn_type:
                    y = max(new_y, y)
                    x += new_x
                elif "P" in conn_type:
                    x = max(new_x, x)
                    y += new_y
            else:
                if "S" in conn_type:
                    x += 1
                elif "P" in conn_type:
                    y += 1

        # components take at least one cell, so add 1 if 0
        if x == 0:
            x += 1

        if y == 0:
            y += 1

        target_dict["x_size"] = x
        target_dict["y_size"] = y
        return x, y

    def _draw_circuit(self, target_dict, conn_type="", x=0, y=0, max_x=0):
        """
        Function to generate circuit of parallel and serial windings including nested circuits
        :param target_dict: connections dictionary
        :param conn_type: connection type for recursion cycle
        :param x: winding X position
        :param y: winding Y position
        :param max_x: maximum X coordinate reached
        :return:
        """
        conn_type = conn_type[:1]
        rel_x_pos = x  # save X coordinate as new relative origin

        def sort_func(arg):
            """
            sort function to sort elements by XY sizes (serial/parallel)
            :param arg: dict key
            :return:
            """

            if "size" in arg:
                # size keys go to the beginning
                return -1

            if not isinstance(target_dict[arg], dict):
                # if it is just layer goes to the end
                return 1e5

            # and dictionaries sort by XY size
            if "P" in conn_type:
                return target_dict[arg]["x_size"]
            elif "S" in conn_type:
                return target_dict[arg]["y_size"]
            else:
                return 0

        # sort all keys the way that first come windings, then dictionaries by descending size order.
        # all work of function is based that serial/parallel branches are in descending size order
        all_keys = sorted(target_dict.keys(), key=sort_func, reverse=True)

        if "P" in conn_type:
            # if parallel connection we need to draw a vertical wire at the beginning of nested circuit and at the end
            self.wire(x, y, x, y + target_dict["y_size"] - 2)
            self.wire(x + 2 * target_dict["x_size"], y, x + 2 * target_dict["x_size"], y + target_dict["y_size"] - 2)
            max_x = target_dict["x_size"]  # save max_x in case if there are nested serial circuits inside

        for i, key in enumerate(all_keys):
            val = target_dict[key]

            if isinstance(val, dict):
                if i > 0:
                    prev_elem = target_dict[all_keys[i - 1]]
                    if isinstance(prev_elem, dict):
                        # if previous item is dictionary then we need to make an offset according to nested circuit size
                        if "P" in conn_type:
                            y += prev_elem["y_size"]
                        else:
                            x += prev_elem["x_size"] * 2  # factor of 2 due to number of windings + wire for each
                self._draw_circuit(val, key, x, y, max_x)  # go into recursion

                if i == len(target_dict) - 3:
                    # last element in dict (last two are sizes)
                    if "S" in conn_type:
                        if x < 2 * (rel_x_pos + max_x - 1):
                            # draw horizontal wire (direction: right) till the end of nested circuit
                            self.wire(x + 1, y, rel_x_pos + 2 * max_x, y)

            elif isinstance(val, str) and val.lower() == "layer":
                # if it is layer value then just draw winding
                self.create_winding(key, x, y)

                if "P" in conn_type:
                    if i != len(target_dict) - 3:
                        # if not last element draw one vertical wire 1 unit height at the beginning of component
                        self.wire(x, y, x, y + 1)

                    if i < len(target_dict) - 3:
                        if target_dict["x_size"] > 1:
                            # if our function has nested serial circuits with size X we need to draw a horizontal wire
                            # Example:
                            # --~~~-------  # second parallel branch
                            # |          |  # connection between layers
                            # --~~~--~~~--  # first parallel branch with nested serial circuit
                            self.wire(x + 2 * target_dict["x_size"], y, x + 2 * target_dict["x_size"], y + 1)

                            # and vertical wire at the end of nested circuit
                            self.wire(x + 1, y, x + 2 * target_dict["x_size"], y)
                        else:
                            # if not, just draw one vertical wire at the end of component
                            self.wire(x + 1, y, x + 1, y + 1)

                    if i == 0:
                        # first element is always the longest, draw one horizontal wire to the right
                        self.wire(x + 1, y, x + 2, y)

                    y += 1  # increase Y since parallel connection
                elif "S" in conn_type:
                    x += 2
                    self.wire(x - 1, y, x, y)

        self.editor.ZoomToFit()

    def create(self):
        for winding, definition in self.winding_connection.items():
            self.page = int(winding)  # set new page for each side: Primary, Secondary, etc
            if self.page > 1:
                self.editor.CreatePage("Page{}".format(self.page))

            self.validate_dict(definition)
            self.run_connection_reduction(definition)
            self.calc_xy(definition)
            self._draw_circuit(definition)

            for key, val in definition.items():
                if "S" in key or "P" in key:
                    max_x = val["x_size"]  # get maximum size of main circuit
                    break

            if self.page == 1:
                source = "Voltage" if self.voltage else "Current"
                self.create_source_or_load(source, max_x * 2)
            else:
                self.create_source_or_load("Load", max_x * 2)

            self.create_ground()


if __name__ == '__main__':
    from AEDTLib.Desktop import Desktop
    import random

    windings = {
            "1": {
                "P6": {
                    "S3": {
                        "7": "Layer",
                        "P2": {
                            "1": "Layer",
                            "2": "Layer"
                        }
                    },
                    "S5": {
                        "4": "Layer",
                        "3": "Layer"
                    },
                    "S4": {
                        "6": "Layer",
                        "5": "Layer"
                    }
                }
            },
            "2": {
                "S7": {
                    "8": "Layer",
                    "9": "Layer"
                }
            },
            "3": {
                "10": "Layer"
            }
        }
    with Desktop("2021.1", release_on_exit=True):
        oDesktop.RestoreWindow()
        oProject = oDesktop.GetActiveProject()
        circuit = Circuit(windings, oProject, str(random.randint(0, 10000000)), current=120, resistance=10,
                          frequency=100000)
        circuit.create()
