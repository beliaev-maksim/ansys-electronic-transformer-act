import copy
import json
import os
import random
from unittest import TestCase

from pyaedt.desktop import Desktop

from src.ElectronicTransformer.circuit import Circuit

AEDT_VERSION = "2023.1"


class TestCircuit(TestCase):
    desktop = None
    project = None
    circuit_file = None

    @classmethod
    def setUpClass(cls):
        cls.desktop = Desktop(AEDT_VERSION)
        cls.project = cls.desktop._main.oDesktop.NewProject()
        tests_dir = os.path.abspath(os.path.dirname(__file__))
        cls.circuit_file = os.path.join(tests_dir, "circuit.sph")

    @classmethod
    def tearDownClass(cls):
        cls.desktop.release_desktop(close_projects=True)
        if os.path.isfile(cls.circuit_file):
            os.remove(cls.circuit_file)

    def setUp(self):
        with open("circuit_sample.json") as file:
            connections = json.load(file)

        conn2 = copy.deepcopy(connections)
        self.windings = {"2": connections, "1": conn2}
        self.circuit = Circuit(
            self.windings,
            self.project,
            str(random.randint(0, 10000000)),
            current=120,
            resistance_list=[1, 3],
            frequency=100000,
        )

    def test_create(self):
        self.circuit.create()

        self.circuit.design.ExportNetlist("", self.circuit_file)
        with open("reference_circuit.sph") as ref_file:
            with open(self.circuit_file) as circuit_file:
                diff = set(ref_file.readlines()[3:]).difference(circuit_file.readlines()[3:])

        # pass only if diff is empty
        self.assertFalse(diff)
