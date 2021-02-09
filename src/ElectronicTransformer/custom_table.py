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
