from abc import ABCMeta, abstractmethod
import clr
import ctypes  # lib to get screen resolution
import copy
from collections import OrderedDict

clr.AddReference("Ans.UI.Toolkit")
clr.AddReference("Ans.UI.Toolkit.Base")
clr.AddReference("Ans.Utilities")
clr.AddReference("Ansys.ACT.Core")

import Ansys
import Ansys.UI.Toolkit

from Ansys.ACT.Core.UI import *
from Ansys.ACT.Interfaces import *
from Ansys.ACT.Core.XmlDataModel import *
from System.Collections.Generic import List
from Ansys.UI.Toolkit import *
from Ansys.UI.Toolkit.Drawing import *
from Ansys.Utilities import *


def message_box(self, msg):
    MessageBox.Show(self, msg,
                    "Error", MessageBoxType.Error,
                    MessageBoxButtons.OK, MessageBoxDefaultButton.Button1)


class WindingFormUI(Ansys.UI.Toolkit.Dialog):
    """
        Abstract class that represents UI definition of the winding dialog
    """
    __metaclass__ = ABCMeta

    def __init__(self):
        self.screen_width = ctypes.windll.user32.GetSystemMetrics(0)
        self.screen_height = ctypes.windll.user32.GetSystemMetrics(1)

        # define components from UI library
        self.available_layers_label = Ansys.UI.Toolkit.Label()
        self.defined_layers_label = Ansys.UI.Toolkit.Label()
        self.move_to_side_label = Ansys.UI.Toolkit.Label()
        self.all_to_side_label = Ansys.UI.Toolkit.Label()

        self.available_layers_listbox = Ansys.UI.Toolkit.ListBox()
        self.defined_layers_listbox = Ansys.UI.Toolkit.ListBox()

        self.custom_sides_number = Ansys.UI.Toolkit.ComboBox()
        self.all_custom_sides_number = Ansys.UI.Toolkit.ComboBox()

        self.primary_button = Ansys.UI.Toolkit.Button()
        self.secondary_button = Ansys.UI.Toolkit.Button()
        self.tertiary_button = Ansys.UI.Toolkit.Button()
        self.custom_button = Ansys.UI.Toolkit.Button()

        self.remove_button = Ansys.UI.Toolkit.Button()

        self.all_primary_button = Ansys.UI.Toolkit.Button()
        self.all_secondary_button = Ansys.UI.Toolkit.Button()
        self.all_tertiary_button = Ansys.UI.Toolkit.Button()
        self.all_custom_button = Ansys.UI.Toolkit.Button()

        self.ok_button = Ansys.UI.Toolkit.Button()
        self.cancel_button = Ansys.UI.Toolkit.Button()

        # Describe all properties of components on init of class
        #
        # ExLabel
        #
        self.available_layers_label.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                         Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.available_layers_label.Location = Ansys.UI.Toolkit.Drawing.Point(20, 20)
        self.available_layers_label.Size = Ansys.UI.Toolkit.Drawing.Size(120, 23)
        self.available_layers_label.Text = "Available Layers"
        #
        # Defined Label
        #
        self.defined_layers_label.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                       Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.defined_layers_label.Location = Ansys.UI.Toolkit.Drawing.Point(360, 20)
        self.defined_layers_label.Size = Ansys.UI.Toolkit.Drawing.Size(161, 23)
        self.defined_layers_label.Text = "Defined Windings"
        #
        # to side label
        #
        self.move_to_side_label.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                     Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.move_to_side_label.Location = Ansys.UI.Toolkit.Drawing.Point(200, 20)
        self.move_to_side_label.Size = Ansys.UI.Toolkit.Drawing.Size(161, 23)
        self.move_to_side_label.Text = "Move to side >"
        #
        # ExcList
        #
        self.available_layers_listbox.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                           Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.available_layers_listbox.Location = Ansys.UI.Toolkit.Drawing.Point(20, 45)
        self.available_layers_listbox.IsMultiSelectable = True
        self.available_layers_listbox.Size = Ansys.UI.Toolkit.Drawing.Size(170, 250)
        self.available_layers_listbox.SelectionChanged += self.update_buttons
        #
        # ReduceList
        #
        self.defined_layers_listbox.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                         Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.defined_layers_listbox.Location = Ansys.UI.Toolkit.Drawing.Point(330, 45)
        self.defined_layers_listbox.IsMultiSelectable = True
        self.defined_layers_listbox.Size = Ansys.UI.Toolkit.Drawing.Size(170, 250)
        self.defined_layers_listbox.SelectionChanged += self.update_buttons
        #
        # primary
        #
        self.primary_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                 Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.primary_button.Location = Ansys.UI.Toolkit.Drawing.Point(200, 50)
        self.primary_button.Size = Ansys.UI.Toolkit.Drawing.Size(30, 27)
        self.primary_button.Text = "1"
        self.primary_button.Click += self.primary_add
        #
        # secondary
        #
        self.secondary_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                   Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.secondary_button.Location = Ansys.UI.Toolkit.Drawing.Point(245, 50)
        self.secondary_button.Size = Ansys.UI.Toolkit.Drawing.Size(30, 27)
        self.secondary_button.Text = "2"
        self.secondary_button.Click += self.secondary_add
        #
        # tertiary
        #
        self.tertiary_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                  Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.tertiary_button.Location = Ansys.UI.Toolkit.Drawing.Point(290, 50)
        self.tertiary_button.Size = Ansys.UI.Toolkit.Drawing.Size(30, 27)
        self.tertiary_button.Text = "3"
        self.tertiary_button.Click += self.tertiary_add
        #
        # custom
        #
        self.custom_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.custom_button.Location = Ansys.UI.Toolkit.Drawing.Point(200, 90)
        self.custom_button.Size = Ansys.UI.Toolkit.Drawing.Size(70, 27)
        self.custom_button.Text = "Custom:"
        self.custom_button.Click += self.custom_add
        #
        # Drop Down Box
        #
        self.custom_sides_number.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                      Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.custom_sides_number.Location = Ansys.UI.Toolkit.Drawing.Point(280, 91)
        self.custom_sides_number.Size = Ansys.UI.Toolkit.Drawing.Size(40, 27)
        #
        # RemoveButton
        #
        self.remove_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.remove_button.Location = Ansys.UI.Toolkit.Drawing.Point(200, 145)
        self.remove_button.Size = Ansys.UI.Toolkit.Drawing.Size(120, 27)
        self.remove_button.Text = "< Remove"
        self.remove_button.Click += self.remove
        #
        # all to side label
        #
        self.all_to_side_label.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                    Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.all_to_side_label.Location = Ansys.UI.Toolkit.Drawing.Point(200, 195)
        self.all_to_side_label.Size = Ansys.UI.Toolkit.Drawing.Size(161, 23)
        self.all_to_side_label.Text = "Move All to side >>"
        #
        # all to primary
        #
        self.all_primary_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                     Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.all_primary_button.Location = Ansys.UI.Toolkit.Drawing.Point(200, 220)
        self.all_primary_button.Size = Ansys.UI.Toolkit.Drawing.Size(30, 27)
        self.all_primary_button.Text = "1"
        self.all_primary_button.Click += self.all_primary
        #
        # all to secondary
        #
        self.all_secondary_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                       Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.all_secondary_button.Location = Ansys.UI.Toolkit.Drawing.Point(245, 220)
        self.all_secondary_button.Size = Ansys.UI.Toolkit.Drawing.Size(30, 27)
        self.all_secondary_button.Text = "2"
        self.all_secondary_button.Click += self.all_secondary
        #
        # all to tertiary
        #
        self.all_tertiary_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                      Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.all_tertiary_button.Location = Ansys.UI.Toolkit.Drawing.Point(290, 220)
        self.all_tertiary_button.Size = Ansys.UI.Toolkit.Drawing.Size(30, 27)
        self.all_tertiary_button.Text = "3"
        self.all_tertiary_button.Click += self.all_tertiary
        #
        # all custom
        #
        self.all_custom_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                    Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.all_custom_button.Location = Ansys.UI.Toolkit.Drawing.Point(200, 260)
        self.all_custom_button.Size = Ansys.UI.Toolkit.Drawing.Size(70, 27)
        self.all_custom_button.Text = "Custom:"
        self.all_custom_button.Click += self.all_custom
        #
        # Drop Down Box for all custom
        #
        self.all_custom_sides_number.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                          Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.all_custom_sides_number.Location = Ansys.UI.Toolkit.Drawing.Point(280, 261)
        self.all_custom_sides_number.Size = Ansys.UI.Toolkit.Drawing.Size(40, 27)
        #
        # OK
        #
        self.ok_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                            Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.ok_button.Location = Ansys.UI.Toolkit.Drawing.Point(330, 305)
        self.ok_button.Size = Ansys.UI.Toolkit.Drawing.Size(80, 27)
        self.ok_button.Text = "OK"
        self.ok_button.Click += self.ok_clicked
        #
        # Cancel
        #
        self.cancel_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.cancel_button.Location = Ansys.UI.Toolkit.Drawing.Point(420, 305)
        self.cancel_button.Size = Ansys.UI.Toolkit.Drawing.Size(80, 27)
        self.cancel_button.Text = "Cancel"
        self.cancel_button.Click += self.cancel_clicked

        #
        # Define the main form. Put here all listed above components
        #
        self.ClientSize = Ansys.UI.Toolkit.Drawing.Size(519, 340)
        self.Location = Ansys.UI.Toolkit.Drawing.Point(self.screen_width / 3, self.screen_height / 3)
        self.Controls.Add(self.cancel_button)
        self.Controls.Add(self.ok_button)
        self.Controls.Add(self.defined_layers_label)
        self.Controls.Add(self.move_to_side_label)
        self.Controls.Add(self.remove_button)
        self.Controls.Add(self.defined_layers_listbox)
        self.Controls.Add(self.primary_button)
        self.Controls.Add(self.secondary_button)
        self.Controls.Add(self.custom_button)
        self.Controls.Add(self.custom_sides_number)
        self.Controls.Add(self.all_to_side_label)
        self.Controls.Add(self.all_custom_button)
        self.Controls.Add(self.all_custom_sides_number)
        self.Controls.Add(self.tertiary_button)
        self.Controls.Add(self.all_primary_button)
        self.Controls.Add(self.all_secondary_button)
        self.Controls.Add(self.all_tertiary_button)
        self.Controls.Add(self.available_layers_label)
        self.Controls.Add(self.available_layers_listbox)
        self.MaximizeBox = False
        self.Name = 'WindingDefinition'
        self.Text = 'Define Windings'

        self.BeforeClose += Ansys.UI.Toolkit.WindowCloseEventDelegate(self.cancel_clicked)

    @abstractmethod
    def update_buttons(self, _sender, _e):
        """Function to enable or disable buttons when list of layers or definitions is changed"""

    @abstractmethod
    def primary_add(self, _sender, _e):
        """
        When click on Add Primary [1] button
        :param _sender: unused
        :param _e: unused
        :return: None
        """

    @abstractmethod
    def secondary_add(self, _sender, _e):
        """
        When click on Add Secondary [2] button
        :param _sender: unused
        :param _e: unused
        :return: None
        """

    @abstractmethod
    def tertiary_add(self, _sender, _e):
        """
        When click on Add Tertiary [3] button
        :param _sender: unused
        :param _e: unused
        :return: None
        """

    @abstractmethod
    def custom_add(self, _sender, _e):
        """
        On click of Custom button we grab number from Drop Down box and pass it as argument
        :param _sender: unused
        :param _e: unused
        :return: None
        """
        """"""

    @abstractmethod
    def all_primary(self, _sender, _e):
        """
        When click on All Primary [1] button
        :param _sender: unused
        :param _e: unused
        :return: None
        """

    @abstractmethod
    def all_secondary(self, _sender, _e):
        """
        When click on All Secondary [2] button
        :param _sender: unused
        :param _e: unused
        :return: None
        """

    @abstractmethod
    def all_tertiary(self, _sender, _e):
        """
        When click on All Tertiary [3] button
        :param _sender: unused
        :param _e: unused
        :return: None
        """

    @abstractmethod
    def all_custom(self, _sender, _e):
        """On click of Custom button we grab number from Drop Down box and pass it as argument"""

    @abstractmethod
    def remove(self, _sender, _e):
        """
        Function to remove item from defined list and move it to undefined layers.
        """

    @abstractmethod
    def ok_clicked(self, _sender, _e):
        """
        OK button clicked
        :param _sender: unused
        :param _e: unused
        :return: None
        """

    @abstractmethod
    def cancel_clicked(self, _sender, _e):
        """
        Cancel button clicked
        :param _sender: unused
        :param _e: unused
        :return: None
        """


class WindingForm(WindingFormUI):
    def __init__(self, number_undefined_layers=0, defined_layers_dict={}):
        super(WindingForm, self).__init__()

        self.number_of_sides = 1
        self.number_undefined_layers = number_undefined_layers
        self.defined_layers_dict = copy.deepcopy(defined_layers_dict)

    def refresh_ui_on_show(self):
        """function to call externally after fill the list of layers to update the UI"""
        if len(self.available_layers_listbox.Items) == 0 and len(self.defined_layers_listbox.Items) == 0:
            self.fill_lists()

        self.populate_custom_sides()
        self.update_buttons("", "")
        self.update_all_to_buttons()

    def fill_lists(self):
        """Fill both listboxes on show of window. Fill Layers list if nothing was specified,
        otherwise populate already defined layers from the list"""
        for i in range(1, self.number_undefined_layers+1):
            self.available_layers_listbox.Items.Add(('Layer' + str(i)))

        for side, val in self.defined_layers_dict.items():
            for layer in val:
                self.defined_layers_listbox.Items.Add("Side_{}_Layer{}".format(side, layer))

    def update_buttons(self, sender, e):
        """Function to enable or disable buttons when list of layers or definitions is changed"""
        if len(self.available_layers_listbox.SelectedItems) > 0:
            side_buttons_activated = True
        else:
            side_buttons_activated = False

        if len(self.defined_layers_listbox.SelectedItems) > 0:
            self.remove_button.Enabled = True
        else:
            self.remove_button.Enabled = False

        self.secondary_button.Enabled = side_buttons_activated
        self.primary_button.Enabled = side_buttons_activated
        self.tertiary_button.Enabled = side_buttons_activated
        self.custom_button.Enabled = side_buttons_activated
        self.custom_sides_number.Enabled = side_buttons_activated

        # disable secondary, tertiary and custom if number of sides less
        side_buttons_visible = False if self.number_of_sides <= 1 else True
        self.secondary_button.Visible = side_buttons_visible

        side_buttons_visible = False if self.number_of_sides <= 2 else True
        self.tertiary_button.Visible = side_buttons_visible

        side_buttons_visible = False if self.number_of_sides <= 3 else True
        self.custom_button.Visible = side_buttons_visible
        self.custom_sides_number.Visible = side_buttons_visible

    def update_all_to_buttons(self):
        """ disable/enable "All to" buttons if no more items available"""

        buttons_activated = False if len(self.available_layers_listbox.Items) == 0 else True

        self.all_primary_button.Enabled = buttons_activated
        self.all_secondary_button.Enabled = buttons_activated
        self.all_tertiary_button.Enabled = buttons_activated
        self.all_custom_button.Enabled = buttons_activated
        self.all_custom_sides_number.Enabled = buttons_activated

        # disable secondary, tertiary and custom if number of sides less
        buttons_visible = False if self.number_of_sides <= 1 else True
        self.all_secondary_button.Visible = buttons_visible

        buttons_visible = False if self.number_of_sides <= 2 else True
        self.all_tertiary_button.Visible = buttons_visible

        buttons_visible = False if self.number_of_sides <= 3 else True
        self.all_custom_button.Visible = buttons_visible
        self.all_custom_sides_number.Visible = buttons_visible

    def populate_custom_sides(self):
        """ Function to update number of sides for Drop Down box according to number
        of specified sides eg Primary, Secondary, etc"""

        if self.number_of_sides <= 3:
            return

        self.custom_sides_number.Clear()
        for i in range(4, self.number_of_sides + 1):
            self.custom_sides_number.AddItem(str(i))
            self.all_custom_sides_number.AddItem(str(i))

        self.custom_sides_number.SelectedIndex = 0
        self.all_custom_sides_number.SelectedIndex = 0

    def define_layer(self, side_number, all_available=False):
        all_available_layers_list = list(self.available_layers_listbox.Items)[:]
        selected_layers_list = sorted(list(self.available_layers_listbox.SelectedItems))[:]
        # need to remove selection before use Items.Remove
        self.available_layers_listbox.ClearSelectedItems()

        layers_list = all_available_layers_list if all_available else selected_layers_list

        for layer in layers_list:
            self.available_layers_listbox.Items.Remove(layer)
            new_item = layer.Text.replace("Layer", "Side_{}_Layer".format(side_number))
            self.defined_layers_listbox.Items.Add(new_item)

        self.update_all_to_buttons()

    def primary_add(self, sender, e):
        self.define_layer(1)

    def secondary_add(self, sender, e):
        self.define_layer(2)

    def tertiary_add(self, sender, e):
        self.define_layer(3)

    def custom_add(self, sender, e):
        """On click of Custom button we grab number from Drop Down box and pass it as argument"""
        self.define_layer(self.custom_sides_number.Text)

    def all_primary(self, sender, e):
        self.define_layer(1, True)

    def all_secondary(self, sender, e):
        self.define_layer(2, True)

    def all_tertiary(self, sender, e):
        self.define_layer(3, True)

    def all_custom(self, sender, e):
        """On click of Custom button we grab number from Drop Down box and pass it as argument"""
        self.define_layer(self.all_custom_sides_number.Text, True)

    def remove(self, sender, e):
        """Function to remove item from defined list and move it to undefined layers.
        We clean all not yet defined layers and then we append the list of "removed" items
        to get the whole list sorted according to the layer number"""
        selected_layers_list = self.defined_layers_listbox.SelectedItems
        all_available_layers_list = list(self.available_layers_listbox.Items)[:]

        self.available_layers_listbox.Items.Clear()
        self.defined_layers_listbox.ClearSelectedItems()

        layers_to_recover = []
        for layer in selected_layers_list:
            self.defined_layers_listbox.Items.Remove(layer)
            new_item = layer.Text.split("_")[-1]
            layers_to_recover.append(new_item)

        layers_to_recover.extend([item.Text for item in all_available_layers_list])

        # need custom sort to have 10, 11 after 9 and not after 1
        layers_to_recover.sort(key=lambda x: '{0:0>20}'.format(x).lower())
        for layer in layers_to_recover:
            self.available_layers_listbox.Items.Add(layer)

        self.update_all_to_buttons()

    def verify_sides_assignment(self, target_dict):
        """function to verify that layers were assigned to each of sides"""
        defined_sides = [int(key) for key in target_dict]
        result = sorted(defined_sides) == list(range(1, self.number_of_sides+1))
        return result

    def ok_clicked(self, sender, e):
        final_dict = {}
        if len(self.available_layers_listbox.Items) > 0:
            return message_box(self, "Windings are not Defined.\nPlease define all layers as windings")

        for item in self.defined_layers_listbox.Items:
            _, side, layer = item.Text.split("_")
            if side not in final_dict:
                final_dict[side] = []

            final_dict[side].append(layer[5:])

        completed = self.verify_sides_assignment(final_dict)

        if not completed:
            msg = "Not all sides are used. Please redefine windings or decrease number of Transformer sides"
            return message_box(self, msg)

        self.defined_layers_dict = copy.deepcopy(final_dict)
        self.Hide()

    def cancel_clicked(self, sender, e):
        self.Hide()


class ConnectionFormUI(Ansys.UI.Toolkit.Dialog):
    """
    Abstract class that represents UI definition of the connection dialog
    """
    __metaclass__ = ABCMeta

    def __init__(self):
        self.screen_width = ctypes.windll.user32.GetSystemMetrics(0)
        self.screen_height = ctypes.windll.user32.GetSystemMetrics(1)

        button_x_size = 120
        button_y_size = 35

        # define components from UI library
        self.side_label = Ansys.UI.Toolkit.Label()
        self.connections_label = Ansys.UI.Toolkit.Label()
        self.connect_label = Ansys.UI.Toolkit.Label()

        self.connections_listbox = Ansys.UI.Toolkit.ListBox()

        self.side_dropdown = Ansys.UI.Toolkit.ComboBox()

        self.serial_button = Ansys.UI.Toolkit.Button()
        self.parallel_button = Ansys.UI.Toolkit.Button()
        self.ungroup_button = Ansys.UI.Toolkit.Button()

        self.ok_button = Ansys.UI.Toolkit.Button()
        self.cancel_button = Ansys.UI.Toolkit.Button()

        # Describe all properties of components on init of class
        #
        # Side Label
        #
        self.side_label.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                             Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.side_label.Location = Ansys.UI.Toolkit.Drawing.Point(20, 23)
        self.side_label.Size = Ansys.UI.Toolkit.Drawing.Size(210, 23)
        self.side_label.Text = "Transformer side:"

        #
        # Side DropDown
        #
        self.side_dropdown.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.side_dropdown.Location = Ansys.UI.Toolkit.Drawing.Point(235, 20)
        self.side_dropdown.Size = Ansys.UI.Toolkit.Drawing.Size(50, 27)
        self.side_dropdown.SelectionChanged += self.side_change

        #
        # Connections Label
        #
        self.connections_label.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                    Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.connections_label.Location = Ansys.UI.Toolkit.Drawing.Point(20, 70)
        self.connections_label.Size = Ansys.UI.Toolkit.Drawing.Size(200, 23)
        self.connections_label.Text = "Connections:"
        #
        # Connections list
        #
        self.connections_listbox.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                      Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.connections_listbox.Location = Ansys.UI.Toolkit.Drawing.Point(20, 100)
        self.connections_listbox.IsMultiSelectable = True
        self.connections_listbox.Size = Ansys.UI.Toolkit.Drawing.Size(960, 300)

        #
        # Connect Label
        #
        self.connect_label.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.connect_label.Location = Ansys.UI.Toolkit.Drawing.Point(20, 413)
        self.connect_label.Size = Ansys.UI.Toolkit.Drawing.Size(110, 23)
        self.connect_label.Text = "Connect:"
        #
        # Serial Button
        #
        self.serial_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.serial_button.Location = Ansys.UI.Toolkit.Drawing.Point(130, 410)
        self.serial_button.Size = Ansys.UI.Toolkit.Drawing.Size(button_x_size, button_y_size)
        self.serial_button.Text = "Serial"
        self.serial_button.Click += self.serial_click
        #
        # Parallel Button
        #
        self.parallel_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                  Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.parallel_button.Location = Ansys.UI.Toolkit.Drawing.Point(270, 410)
        self.parallel_button.Size = Ansys.UI.Toolkit.Drawing.Size(button_x_size, button_y_size)
        self.parallel_button.Text = "Parallel"
        self.parallel_button.Click += self.parallel_click

        #
        # Ungroup Button
        #
        self.ungroup_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                 Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.ungroup_button.Location = Ansys.UI.Toolkit.Drawing.Point(410, 410)
        self.ungroup_button.Size = Ansys.UI.Toolkit.Drawing.Size(button_x_size, button_y_size)
        self.ungroup_button.Text = "Ungroup"
        self.ungroup_button.Click += self.ungroup_click

        #
        # OK Button
        #
        self.ok_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                            Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.ok_button.Location = Ansys.UI.Toolkit.Drawing.Point(650, 450)
        self.ok_button.Size = Ansys.UI.Toolkit.Drawing.Size(button_x_size, button_y_size)
        self.ok_button.Text = "OK"
        self.ok_button.Click += self.ok_clicked
        #
        # Cancel Button
        #
        self.cancel_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.cancel_button.Location = Ansys.UI.Toolkit.Drawing.Point(800, 450)
        self.cancel_button.Size = Ansys.UI.Toolkit.Drawing.Size(button_x_size, button_y_size)
        self.cancel_button.Text = "Cancel"
        self.cancel_button.Click += self.cancel_clicked

        #
        # Define the main form. Put here all listed above components
        #
        self.ClientSize = Ansys.UI.Toolkit.Drawing.Size(1000, 500)
        self.Location = Ansys.UI.Toolkit.Drawing.Point(self.screen_width / 3, self.screen_height / 3)

        self.Controls.Add(self.side_label)
        self.Controls.Add(self.side_dropdown)

        self.Controls.Add(self.connections_label)
        self.Controls.Add(self.connections_listbox)

        self.Controls.Add(self.connect_label)
        self.Controls.Add(self.serial_button)
        self.Controls.Add(self.parallel_button)
        self.Controls.Add(self.ungroup_button)

        self.Controls.Add(self.cancel_button)
        self.Controls.Add(self.ok_button)

        self.MaximizeBox = False
        self.Name = 'ConnectionDefinition'
        self.Text = 'Define Connections'

        self.VisibilityChanged += Ansys.UI.Toolkit.EventDelegate(self.on_show)

    @abstractmethod
    def on_show(self, _sender, _e):
        """
        Called on Window show
        :param _sender:
        :param _e:
        :return:
        """

    @abstractmethod
    def side_change(self, _sender, _e):
        """
        Called when drop down menu Side value is changed
        :param _sender: unused
        :param _e: unused
        :return:
        """

    @abstractmethod
    def serial_click(self, _sender, _e):
        """
        Serial button clicked
        :param _sender: unused
        :param _e: unused
        :return:
        """
        self.group_connection(conn_type="S")

    @abstractmethod
    def parallel_click(self, _sender, _e):
        """
        Parallel button clicked
        :param _sender: unused
        :param _e: unused
        :return:
        """

    @abstractmethod
    def ungroup_click(self, _sender, _e):
        """
        Ungroup button clicked. Ungroup connection and refill dictionary
        :param _sender: unused
        :param _e: unused
        :return:
        """

    @abstractmethod
    def ok_clicked(self, _sender, _e):
        """
        OK button clicked. just close, data is stored in connections_dict
        :param _sender: unused
        :param _e: unused
        :return:
        """

    @abstractmethod
    def cancel_clicked(self, _sender, _e):
        """
        Cancel button clicked. Revert data to backed up
        :param _sender: unused
        :param _e: unused
        :return:
        """


class ConnectionForm(ConnectionFormUI):
    def __init__(self):
        super(ConnectionForm, self).__init__()
        self.winding_def_dict = {}
        self.connections_dict = OrderedDict()
        self.temp_connections_dict = OrderedDict()
        self.active_side = "1"
        self._id = 0
        self.backup = None

    def on_show(self, _sender, _e):
        """
        Everytime UI visibility is changed, this method is called. This will prefill sides and lists in UI
        Note: due to bug in core, this method is fired always twice
        :param _sender: unused
        :param _e: unused
        :return:
        """
        if self.Visible:
            self.temp_connections_dict = OrderedDict()
            if not self.connections_dict:
                self._id = 1
                # no connections defined, make copy of mutable object
                temp_dict = copy.deepcopy(self.winding_def_dict)
                for key, val in temp_dict.items():
                    self.temp_connections_dict[key] = OrderedDict((item, "Layer") for item in val)
            else:
                self.temp_connections_dict = copy.deepcopy(self.connections_dict)
                self._id = self.find_max_id(self.temp_connections_dict) + 1

            self.active_side = "1"
            self.fill_sides()

    @property
    def new_id(self):
        """
            Function to generate new ID for parallel or serial connection dict keys
        """
        self._id += 1
        return self._id

    def find_max_id(self, target_dict, max_id=0):
        """
        Find maximum ID used for key in dictionary
        :param target_dict: dict to check
        :param max_id: maximum ID found, used for recursion
        :return:
        """
        for key, val in target_dict.items():
            if isinstance(val, dict):
                if "S" in key or "P" in key:
                    max_id = int(key[1:])  # cut S/P part

                new_id = self.find_max_id(val, max_id)
                max_id = max(new_id, max_id)

        return max_id

    def fill_sides(self):
        """
        Prefill transformers sides dropdown menu
        :return:
        """
        self.side_dropdown.Clear()
        for side in sorted(self.temp_connections_dict):
            self.side_dropdown.AddItem(side)

        self.side_dropdown.SelectedIndex = 0  # also triggers SelectionChanged event

    def fill_lists(self):
        """Fill both listboxes on show of window. Fill layers per side if nothing was specified,
        otherwise populate already defined connections from dictionary"""

        self.connections_listbox.Items.Clear()
        # natural_keys are defined in etk_callback, due to ACT architecture no explicit import
        for layer in sorted(self.temp_connections_dict[self.active_side].keys(), key=natural_keys):
            val = self.temp_connections_dict[self.active_side][layer]
            if isinstance(val, dict):
                self.connections_listbox.Items.Add((self.dict_to_str(val, conn_type=layer)))
            else:
                self.connections_listbox.Items.Add(('Layer' + layer))

    def dict_to_str(self, conn_dict, conn_type="", str_item=""):
        """
        Construct a string from dictionary. This string will represent element in UI
        :param conn_dict: dict from which we construct
        :param conn_type: S (Serial) or P (parallel)
        :param str_item: constructed str, need for recursion
        :return:
        """
        str_item += conn_type + "("
        layers = []
        for key, val in conn_dict.items():
            if isinstance(val, dict):
                nested = self.dict_to_str(val, conn_type=key)
                layers.append(nested)
            else:
                layers.append(key)

        str_item += ",".join(layers) + ")"

        return str_item

    def side_change(self, _sender, _e):
        """
        Called when drop down menu Side value is changed
        :param _sender:
        :param _e:
        :return:
        """
        self.active_side = self.side_dropdown.Text
        self.fill_lists()

    def serial_click(self, _sender, _e):
        """
        Serial button clicked
        :param _sender: unused
        :param _e: unused
        :return:
        """
        self.group_connection(conn_type="S")

    def parallel_click(self, _sender, _e):
        """
        Parallel button clicked
        :param _sender: unused
        :param _e: unused
        :return:
        """
        self.group_connection(conn_type="P")

    def group_connection(self, conn_type):
        """
        Group selection. Take multiple items and unite them under same key in dict
        :param conn_type: S or P for Serial and Parallel
        :return:
        """
        selected_connections_list = sorted(list(self.connections_listbox.SelectedItems))
        if len(selected_connections_list) < 2:
            return message_box(self, "You need at least 2 selections to create a connection")

        # need to remove selection before use Items.Remove
        self.connections_listbox.ClearSelectedItems()

        temp_dict = OrderedDict()
        new_id = str(self.new_id)
        self.temp_connections_dict[self.active_side][conn_type + new_id] = temp_dict
        for item in selected_connections_list:
            self.connections_listbox.Items.Remove(item)
            if "Layer" in item.Text:
                key = item.Text.replace("Layer", "")
            else:
                key = item.Text.split("(")[0]

            val = self.temp_connections_dict[self.active_side].pop(key)
            temp_dict[key] = val

        self.fill_lists()

    def ungroup_click(self, _sender, _e):
        """
        Ungroup button clicked. Ungroup connection and refill dictionary
        :param _sender: unused
        :param _e: unused
        :return:
        """
        selected_connections_list = sorted(list(self.connections_listbox.SelectedItems))
        if len(selected_connections_list) == 0:
            return message_box(self, "You need at least 1 selection to ungroup a connection")

        # need to remove selection before use Items.Remove
        self.connections_listbox.ClearSelectedItems()

        for item in selected_connections_list:
            self.connections_listbox.Items.Remove(item)
            if "Layer" not in item.Text:
                key = item.Text.split("(")[0]
                val = self.temp_connections_dict[self.active_side].pop(key)
                self.temp_connections_dict[self.active_side].update(val)

        self.fill_lists()

    def ok_clicked(self, _sender, _e):
        """
        OK button clicked. just close, data is stored in connections_dict
        :param _sender: unused
        :param _e: unused
        :return:
        """
        valid, side = self.validate()
        if not valid:
            msg = "You cannot have more than 1 item in the site list, please use connect Serial/Parallel button. "
            msg += "Undefined side: {}".format(side)
            return message_box(self, msg)

        self.connections_dict = copy.deepcopy(self.temp_connections_dict)
        self.Hide()

    def cancel_clicked(self, _sender, _e):
        """
        Cancel button clicked. Revert data to backed up
        :param _sender: unused
        :param _e: unused
        :return:
        """
        self.temp_connections_dict = copy.deepcopy(self.connections_dict)
        self.Hide()

    def validate(self):
        for side, val in self.temp_connections_dict.items():
            if len(val) > 1:
                return False, side
        return True, 0


class TabularDataEditor:
    DialogName = "TabularDataDialog"
    ComponentName = "TabularDataEditorDialog"

    def __init__(self, api, step, prop):
        self.dialog = None
        self.prop = prop
        self.propClone = None
        self.step = step

    def button_clicked(self, sender, args):
        tabular_data_comp = self.dialog.GetComponent(TabularDataEditor.ComponentName)

        # no "our" dialog
        if self.prop != tabular_data_comp.GetPropertyTable():
            return

        if args.ButtonName == "Cancel":
            self.prop.Cancel()
            self.prop.CopyFrom(self.propClone, True)
            # we change the property so we have to refresh the component with the "old" table
            tabular_data_comp.SetPropertyTable(self.prop)
            tabular_data_comp.Refresh()
        elif args.ButtonName == "Apply":
            # update properties and finish buttons states
            self.prop.Apply()
            self.step.NotifyChange()
        self.dialog.Hide()

    def createTabularDialog(self, panel, title, width, height):
        dialog = panel.CreateDialog(TabularDataEditor.DialogName, title, width, height)

        layout_def = LayoutDefinition()
        layout_def.Extension = None

        layout = Ansys.ACT.Core.XmlDataModel.UI.Layout()
        layout_def.Layout = layout
        layout.Name = "TabularDataDialogLayout"
        layout.Components = List[Ansys.ACT.Core.XmlDataModel.UI.Component]()

        component = Ansys.ACT.Core.XmlDataModel.UI.Component()
        component.Name = TabularDataEditor.ComponentName
        component.TopAttachment = ""
        component.TopOffset = 0
        component.BottomAttachment = "Buttons"
        component.BottomOffset = 0
        component.LeftAttachment = ""
        component.LeftOffset = 0
        component.RightAttachment = ""
        component.RightOffset = 0
        component.WidthType = Ansys.ACT.Interfaces.UserInterface.ComponentLengthType.Percentage
        component.WidthValue = 100
        component.HeightType = Ansys.ACT.Interfaces.UserInterface.ComponentLengthType.Percentage
        component.HeightValue = 100
        component.CustomJSFile = r'file:///' + str(ExtAPI.Extension.InstallDir) + r'\custom_table.js'
        component.ComponentType = "tabularDataComponent"
        layout.Components.Add(component)

        component = Ansys.ACT.Core.XmlDataModel.UI.Component()
        component.Name = "Buttons"
        component.TopAttachment = TabularDataEditor.ComponentName
        component.TopOffset = 0
        component.BottomAttachment = ""
        component.BottomOffset = 0
        component.LeftAttachment = ""
        component.LeftOffset = 0
        component.RightAttachment = ""
        component.RightOffset = 0
        component.WidthType = Ansys.ACT.Interfaces.UserInterface.ComponentLengthType.Percentage
        component.WidthValue = 100
        component.HeightType = Ansys.ACT.Interfaces.UserInterface.ComponentLengthType.FitToContent
        component.ComponentType = "buttonsComponent"
        layout.Components.Add(component)

        dialog.SetLayout(layout_def)

        button_comp = dialog.GetComponent("Buttons")
        button_comp.AddButton("Apply", "Apply", Ansys.ACT.Interfaces.UserInterface.ButtonPositionType.Center)
        button_comp.AddButton("Cancel", "Cancel", Ansys.ACT.Interfaces.UserInterface.ButtonPositionType.Center)

        return dialog

    def onshow(self, step, prop):
        self.dialog = step.UserInterface.GetComponent(TabularDataEditor.DialogName)
        if self.dialog is None:
            self.dialog = self.createTabularDialog(step.UserInterface.Panel, "TabularData", 100, 100)

        button_comp = self.dialog.GetComponent("Buttons")
        button_comp.ButtonClicked += self.button_clicked

    def onhide(self, step, prop):
        self.dialog = None

    def onactivate(self, step, prop):
        dialog_width = 532

        self.dialog.GetComponent(TabularDataEditor.ComponentName).SetPropertyTable(prop)
        self.dialog.GetComponent(TabularDataEditor.ComponentName).Refresh()
        self.dialog.SetDatas("TabularData", dialog_width, dialog_width)
        self.dialog.Show()

        # we save all value to take it back if we press the cancel button
        self.propClone = prop.Clone()

    def value2string(self, step, prop, value):
        return "Tabular Data"
