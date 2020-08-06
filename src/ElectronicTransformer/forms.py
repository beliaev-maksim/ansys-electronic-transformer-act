import clr
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

import ctypes  # lib to get screen resolution


class WindingForm(Ansys.UI.Toolkit.Dialog):
    def __init__(self, number_undefined_layers=0, defined_layers_list=[]):
        self.number_of_sides = 2
        self.number_undefined_layers = number_undefined_layers
        self.defined_layers_list = defined_layers_list

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
        self.available_layers_label.Name = "ExLabel"
        self.available_layers_label.Size = Ansys.UI.Toolkit.Drawing.Size(120, 23)
        self.available_layers_label.Text = "Available Layers"
        #
        # Defined Label
        #
        self.defined_layers_label.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                       Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.defined_layers_label.Location = Ansys.UI.Toolkit.Drawing.Point(360, 20)
        self.defined_layers_label.Name = "ReduceLabel"
        self.defined_layers_label.Size = Ansys.UI.Toolkit.Drawing.Size(161, 23)
        self.defined_layers_label.Text = "Defined Windings"
        #
        # to side label
        #
        self.move_to_side_label.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                     Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.move_to_side_label.Location = Ansys.UI.Toolkit.Drawing.Point(200, 20)
        self.move_to_side_label.Name = "to side label"
        self.move_to_side_label.Size = Ansys.UI.Toolkit.Drawing.Size(161, 23)
        self.move_to_side_label.Text = "Move to side >"
        #
        # ExcList
        #
        self.available_layers_listbox.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                           Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.available_layers_listbox.Location = Ansys.UI.Toolkit.Drawing.Point(20, 45)
        self.available_layers_listbox.Name = "ExcList"
        self.available_layers_listbox.IsMultiSelectable = True
        self.available_layers_listbox.Size = Ansys.UI.Toolkit.Drawing.Size(170, 250)
        self.available_layers_listbox.SelectionChanged += self.update_buttons
        #
        # ReduceList
        #
        self.defined_layers_listbox.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                         Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.defined_layers_listbox.Location = Ansys.UI.Toolkit.Drawing.Point(330, 45)
        self.defined_layers_listbox.Name = "ReduceList"
        self.defined_layers_listbox.IsMultiSelectable = True
        self.defined_layers_listbox.Size = Ansys.UI.Toolkit.Drawing.Size(170, 250)
        self.defined_layers_listbox.SelectionChanged += self.update_buttons
        #
        # primary
        #
        self.primary_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                 Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.primary_button.Location = Ansys.UI.Toolkit.Drawing.Point(200, 50)
        self.primary_button.Name = "PrimButton"
        self.primary_button.Size = Ansys.UI.Toolkit.Drawing.Size(30, 27)
        self.primary_button.Text = "1"
        self.primary_button.Click += self.primary_add
        #
        # secondary
        #
        self.secondary_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                   Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.secondary_button.Location = Ansys.UI.Toolkit.Drawing.Point(245, 50)
        self.secondary_button.Name = "SecButton"
        self.secondary_button.Size = Ansys.UI.Toolkit.Drawing.Size(30, 27)
        self.secondary_button.Text = "2"
        self.secondary_button.Click += self.secondary_add
        #
        # tertiary
        #
        self.tertiary_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                  Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.tertiary_button.Location = Ansys.UI.Toolkit.Drawing.Point(290, 50)
        self.tertiary_button.Name = "tertiary_button"
        self.tertiary_button.Size = Ansys.UI.Toolkit.Drawing.Size(30, 27)
        self.tertiary_button.Text = "3"
        self.tertiary_button.Click += self.tertiary_add
        #
        # custom
        #
        self.custom_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.custom_button.Location = Ansys.UI.Toolkit.Drawing.Point(200, 90)
        self.custom_button.Name = "custom_button"
        self.custom_button.Size = Ansys.UI.Toolkit.Drawing.Size(70, 27)
        self.custom_button.Text = "Custom:"
        self.custom_button.Click += self.custom_add
        #
        # Drop Down Box
        #
        self.custom_sides_number.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                      Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.custom_sides_number.Location = Ansys.UI.Toolkit.Drawing.Point(280, 91)
        self.custom_sides_number.Name = "custom_number_box"
        self.custom_sides_number.Size = Ansys.UI.Toolkit.Drawing.Size(40, 27)
        #
        # RemoveButton
        #
        self.remove_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.remove_button.Location = Ansys.UI.Toolkit.Drawing.Point(200, 145)
        self.remove_button.Name = "RemoveButton"
        self.remove_button.Size = Ansys.UI.Toolkit.Drawing.Size(120, 27)
        self.remove_button.Text = "< Remove"
        self.remove_button.Click += self.remove
        #
        # all to side label
        #
        self.all_to_side_label.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                    Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.all_to_side_label.Location = Ansys.UI.Toolkit.Drawing.Point(200, 195)
        self.all_to_side_label.Name = "all to side label"
        self.all_to_side_label.Size = Ansys.UI.Toolkit.Drawing.Size(161, 23)
        self.all_to_side_label.Text = "Move All to side >>"
        #
        # all to primary
        #
        self.all_primary_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                     Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.all_primary_button.Location = Ansys.UI.Toolkit.Drawing.Point(200, 220)
        self.all_primary_button.Name = "PrimButton"
        self.all_primary_button.Size = Ansys.UI.Toolkit.Drawing.Size(30, 27)
        self.all_primary_button.Text = "1"
        self.all_primary_button.Click += self.all_primary
        #
        # all to secondary
        #
        self.all_secondary_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                       Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.all_secondary_button.Location = Ansys.UI.Toolkit.Drawing.Point(245, 220)
        self.all_secondary_button.Name = "SecButton"
        self.all_secondary_button.Size = Ansys.UI.Toolkit.Drawing.Size(30, 27)
        self.all_secondary_button.Text = "2"
        self.all_secondary_button.Click += self.all_secondary
        #
        # all to tertiary
        #
        self.all_tertiary_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                      Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.all_tertiary_button.Location = Ansys.UI.Toolkit.Drawing.Point(290, 220)
        self.all_tertiary_button.Name = "SecButton"
        self.all_tertiary_button.Size = Ansys.UI.Toolkit.Drawing.Size(30, 27)
        self.all_tertiary_button.Text = "3"
        self.all_tertiary_button.Click += self.all_tertiary
        #
        # all custom
        #
        self.all_custom_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                    Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.all_custom_button.Location = Ansys.UI.Toolkit.Drawing.Point(200, 260)
        self.all_custom_button.Name = "all_custom_button"
        self.all_custom_button.Size = Ansys.UI.Toolkit.Drawing.Size(70, 27)
        self.all_custom_button.Text = "Custom:"
        self.all_custom_button.Click += self.all_custom
        #
        # Drop Down Box for all custom
        #
        self.all_custom_sides_number.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                          Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.all_custom_sides_number.Location = Ansys.UI.Toolkit.Drawing.Point(280, 261)
        self.all_custom_sides_number.Name = "all_custom_number_box"
        self.all_custom_sides_number.Size = Ansys.UI.Toolkit.Drawing.Size(40, 27)
        #
        # OK
        #
        self.ok_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                            Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.ok_button.Location = Ansys.UI.Toolkit.Drawing.Point(330, 305)
        self.ok_button.Name = "RedOK"
        self.ok_button.Size = Ansys.UI.Toolkit.Drawing.Size(80, 27)
        self.ok_button.Text = "OK"
        self.ok_button.Click += self.ok_clicked
        #
        # Cancel
        #
        self.cancel_button.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75,
                                                                Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
        self.cancel_button.Location = Ansys.UI.Toolkit.Drawing.Point(420, 305)
        self.cancel_button.Name = "RedCancel"
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
        # self.Controls.Add(
        self.Controls.Add(self.tertiary_button)
        self.Controls.Add(self.all_primary_button)
        self.Controls.Add(self.all_secondary_button)
        self.Controls.Add(self.all_tertiary_button)
        self.Controls.Add(self.available_layers_label)
        self.Controls.Add(self.available_layers_listbox)
        self.MaximizeBox = False
        self.Name = 'WindingDefinition'
        self.Text = 'Define Windings'

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

        for item in self.defined_layers_list:
            self.defined_layers_listbox.Items.Add(item)

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
        selected_layers_list = sorted(list(self.available_layers_listbox.SelectedItems))
        # need to remove selection before use Irems.Remove
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

    def reset(self):
        """Clear both lists. Defined list could be only in case if OK was clicked at least once.
        Otherwise we need to create only undefined list of layers."""
        self.defined_layers_listbox.Items.Clear()
        self.available_layers_listbox.Items.Clear()

        if not self.defined_layers_list:
            for i in range(1, self.number_undefined_layers+1):
                self.available_layers_listbox.Items.Add(('Layer' + str(i)))
        else:
            for item in self.defined_layers_list:
                self.defined_layers_listbox.Items.Add(item)

    def verify_sides_assignment(self):
        """function to verify that layers were assigned to each of sides"""
        check_list = []
        for layer in self.defined_layers_list:
            check_list.append(layer.split("_")[1])

        return all([str(item) in check_list for item in list(range(1, self.number_of_sides+1))])

    def ok_clicked(self, sender, e):
        self.defined_layers_list = []  # clear the list each time OK clicked not to append multiple times
        if len(self.available_layers_listbox.Items) > 0:
            return MessageBox.Show(self,
                                   "Windings are not Defined.\nPlease define all layers as windings",
                                   "Error",
                                   MessageBoxType.Error,
                                   MessageBoxButtons.OK, MessageBoxDefaultButton.Button1)

        for item in self.defined_layers_listbox.Items:
            self.defined_layers_list.append(item.Text)

        completed = self.verify_sides_assignment()

        if not completed:
            self.defined_layers_list = []  # clear the list in case if it was wrong (it is not listbox)
            return MessageBox.Show(self,
                            "Not all sides are used. Please redefine windings or decrease number of Transformer sides",
                            "Error",
                            MessageBoxType.Error,
                            MessageBoxButtons.OK, MessageBoxDefaultButton.Button1)
        self.Hide()

    def cancel_clicked(self, sender, e):
        self.reset()
        self.Hide()


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
            # we change the property so we have to refresh the compoenent with the "old" table
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
        component.CustomHTMLFile = str(ExtAPI.Extension.InstallDir) + r'\testComp.html'
        component.CustomCSSFile = str(ExtAPI.Extension.InstallDir) + r'\testComp.css'
        component.CustomJSFile = str(ExtAPI.Extension.InstallDir) + r'\testComp.js'
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
