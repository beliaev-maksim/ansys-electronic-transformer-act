import clr
clr.AddReference("Ans.UI.Toolkit")
clr.AddReference("Ans.UI.Toolkit.Base")
clr.AddReference("Ans.Utilities")
clr.AddReference("Ansys.ACT.Core")

import re
import Ansys
import Ansys.UI.Toolkit

from Ansys.ACT.Core.UI import *
from Ansys.ACT.Interfaces import *
from Ansys.ACT.Core.XmlDataModel import *
from System.Collections.Generic import List
from Ansys.UI.Toolkit import *
from Ansys.UI.Toolkit.Drawing import *
from Ansys.Utilities import *


class NameForm(Ansys.UI.Toolkit.Dialog):
  def __init__(self):
    self._label_Name = Ansys.UI.Toolkit.Label()
    self._label_Path = Ansys.UI.Toolkit.Label()
    self._textBox1 = Ansys.UI.Toolkit.TextBox()
    self._textBox2 = Ansys.UI.Toolkit.TextBox()
    self._button_OK = Ansys.UI.Toolkit.Button()
    #
    # label_Name
    #
    self._label_Name.Location = Ansys.UI.Toolkit.Drawing.Point(13, 13)
    self._label_Name.Name = "label_Name"
    self._label_Name.Size = Ansys.UI.Toolkit.Drawing.Size(205, 23)
    self._label_Name.Text = "Specify Name of Connection:"
    #
    # textBox1
    #
    self._textBox1.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._textBox1.Location = Ansys.UI.Toolkit.Drawing.Point(36, 39)
    self._textBox1.Name = "textBox1"
    self._textBox1.Size = Ansys.UI.Toolkit.Drawing.Size(205, 21)
    self._textBox1.Focus()
    #
    # label_Path
    #
    self._label_Path.Location = Ansys.UI.Toolkit.Drawing.Point(13, 75)
    self._label_Path.Name = "label_Path"
    self._label_Path.Size = Ansys.UI.Toolkit.Drawing.Size(205, 23)
    self._label_Path.Text = "No. of Parallel Paths:"
    #
    # textBox2
    #
    self._textBox2.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._textBox2.Location = Ansys.UI.Toolkit.Drawing.Point(36, 101)
    self._textBox2.Name = "textBox2"
    self._textBox2.Size = Ansys.UI.Toolkit.Drawing.Size(80, 21)
    self._textBox2.Text = "1"
    #
    # button_OK
    #
    self._button_OK.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._button_OK.Location = Ansys.UI.Toolkit.Drawing.Point(150, 125)
    self._button_OK.Name = "button_OK"
    self._button_OK.Size = Ansys.UI.Toolkit.Drawing.Size(73, 28)
    self._button_OK.Text = "OK"
    self._button_OK.Click += self.CheckName
    #
    # Form1
    #
    self.AcceptButton = self._button_OK
    self.BeforeClose += self.Cancel
    self.ClientSize = Ansys.UI.Toolkit.Drawing.Size(284, 170)
    self.Controls.Add(self._textBox1)
    self.Controls.Add(self._textBox2)
    self.Controls.Add(self._button_OK)
    self.Controls.Add(self._label_Name)
    self.Controls.Add(self._label_Path)
    self.MaximizeBox = False
    self.MinimizeBox = False
    self.Name = "Form1"
    self.Text = "Input"
    self.AcceptValue = False

  def CheckName(self, sender, e):
    if re.match('^[A-Za-z0-9-_]*$',self._textBox1.Text) :
      if self._textBox1.Text == "":
        MessageBox.Show(self,"Invalid Entry:\nPlease enter a valid name", "Error",
                        MessageBoxType.Error, MessageBoxButtons.OK, MessageBoxDefaultButton.Button1)

        self._textBox1.Focus()
        return
      if self._textBox1.Text in self.ExistingNames:
        MessageBox.Show(self,"Duplicate Name:\nPlease enter a valid name", "Error",
                        MessageBoxType.Error, MessageBoxButtons.OK, MessageBoxDefaultButton.Button1)
        return

      if self._textBox2.Text.isdigit():
        if not(float(self._textBox2.Text).is_integer()):
          MessageBox.Show(self,"Invalid Entry:\nPlease enter integer value for parallel paths", "Error",
                          MessageBoxType.Error, MessageBoxButtons.OK, MessageBoxDefaultButton.Button1)
          return

      else:
        MessageBox.Show(self,"Invalid Entry:\nExpected integer value for parallel paths", "Error",
                        MessageBoxType.Error, MessageBoxButtons.OK, MessageBoxDefaultButton.Button1)
        return

      self.DefName = self._textBox1.Text
      self.DefPath = self._textBox2.Text
      self.AcceptValue = True
      self.Close()
    else:
      MessageBox.Show(self,"Invalid Entry:\nPlease enter a valid name", "Error",
                      MessageBoxType.Error, MessageBoxButtons.OK, MessageBoxDefaultButton.Button1)
      self._textBox1.Focus()

  def Cancel(self, sender, e):
    self._textBox1.Text = ""
    self._textBox2.Text = "1"
    if not self.AcceptValue:
      self.DefName = ""

popForm = NameForm()
class ConnForm(Ansys.UI.Toolkit.Dialog):
  def __init__(self):
    self._InWgLabel = Ansys.UI.Toolkit.Label()
    self._ConnLabel = Ansys.UI.Toolkit.Label()
    self._InWgList = Ansys.UI.Toolkit.ListBox()
    self._ConnList = Ansys.UI.Toolkit.ListBox()
    self._GroupButton = Ansys.UI.Toolkit.Button()
    self._UngroupButton = Ansys.UI.Toolkit.Button()
    self._RedOK = Ansys.UI.Toolkit.Button()
    self._RedCancel = Ansys.UI.Toolkit.Button()
    #
    # InWgLabel
    #
    self._InWgLabel.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._InWgLabel.Location = Ansys.UI.Toolkit.Drawing.Point(20, 20)
    self._InWgLabel.Name = "WgLabel"
    self._InWgLabel.Size = Ansys.UI.Toolkit.Drawing.Size(120, 23)
    self._InWgLabel.Text = "Windings"
    #
    # ConnLabel
    #
    self._ConnLabel.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._ConnLabel.Location = Ansys.UI.Toolkit.Drawing.Point(320, 20)
    self._ConnLabel.Name = "ConnLabel"
    self._ConnLabel.Size = Ansys.UI.Toolkit.Drawing.Size(161, 23)
    self._ConnLabel.Text = "Defined Connections"
    #
    # InWgList
    #
    self._InWgList.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    #self._InWgList.ItemHeight = 16
    self._InWgList.Location = Ansys.UI.Toolkit.Drawing.Point(20, 45)
    self._InWgList.Name = "WdgList"
    self._InWgList.IsMultiSelectable = True
    self._InWgList.Size = Ansys.UI.Toolkit.Drawing.Size(170, 196)
    self._InWgList.SelectionChanged += self.UpdateGUI
    #
    # ConnList
    #
    self._ConnList.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    #self._ConnList.ItemHeight = 16
    self._ConnList.Location = Ansys.UI.Toolkit.Drawing.Point(330, 45)
    self._ConnList.Name = "ConnList"
    self._ConnList.IsMultiSelectable = True
    self._ConnList.Size = Ansys.UI.Toolkit.Drawing.Size(170, 196)
    self._ConnList.SelectionChanged += self.UpdateGUI
    #
    # GroupButton
    #
    self._GroupButton.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._GroupButton.Location = Ansys.UI.Toolkit.Drawing.Point(205, 75)
    self._GroupButton.Name = "GroupButton"
    self._GroupButton.Size = Ansys.UI.Toolkit.Drawing.Size(110, 27)
    self._GroupButton.Text = "Group >>"
    self._GroupButton.Click += self.GroupAdd
    #
    # UngroupButton
    #
    self._UngroupButton.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._UngroupButton.Location = Ansys.UI.Toolkit.Drawing.Point(205, 175)
    self._UngroupButton.Name = "UngroupButton"
    self._UngroupButton.Size = Ansys.UI.Toolkit.Drawing.Size(110, 27)
    self._UngroupButton.Text = "<< Ungroup"
    self._UngroupButton.Click += self.UngroupAny
    #
    # RedOK
    #
    self._RedOK.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._RedOK.Location = Ansys.UI.Toolkit.Drawing.Point(320, 275)
    self._RedOK.Name = "RedOK"
    self._RedOK.Size = Ansys.UI.Toolkit.Drawing.Size(80, 27)
    self._RedOK.Text = "OK"
    self._RedOK.Click += self.RedAssign
    #
    # RedCancel
    #
    self._RedCancel.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._RedCancel.Location = Ansys.UI.Toolkit.Drawing.Point(420, 275)
    self._RedCancel.Name = "RedCancel"
    self._RedCancel.Size = Ansys.UI.Toolkit.Drawing.Size(80, 27)
    self._RedCancel.Text = "Cancel"
    self._RedCancel.Click += self.RedHide
    #
    # ConnectionDefinition
    #
    self.ClientSize = Ansys.UI.Toolkit.Drawing.Size(519, 314)
    self.Controls.Add(self._RedCancel)
    self.Controls.Add(self._RedOK)
    self.Controls.Add(self._ConnLabel)
    self.Controls.Add(self._UngroupButton)
    self.Controls.Add(self._ConnList)
    self.Controls.Add(self._GroupButton)
    self.Controls.Add(self._InWgLabel)
    self.Controls.Add(self._InWgList)
    self.MaximizeBox = False
    self.Name = 'ConnectionDefinition'
    self.Text = 'Define Connections'


  def UpdateGUI(self, sender, e):
    if len(self._InWgList.SelectedItems) > 0:
      self._GroupButton.Enabled = True
    else:
      self._GroupButton.Enabled = False
    if len(self._ConnList.SelectedItems) > 0:
      self._UngroupButton.Enabled = True
    else:
      self._UngroupButton.Enabled = False

  def GroupAdd(self, sender, e):
    popForm.AcceptValue = False
    popForm._textBox1.Text == ""
    popForm._textBox2.Text == "1"
    popForm.ExistingNames = self.GroupDict.keys()
    popForm.ShowDialog()
    if popForm.DefName == "":
      return
    AllItemList2 = list(self._InWgList.Items)[:]
    GroupList = sorted(self._InWgList.SelectedItemTexts) [:]
    self._InWgList.Items.Clear()
    for AllRem2 in AllItemList2:
      if not(AllRem2.Text in GroupList):
        self._InWgList.Items.Add(AllRem2)
    self._ConnList.Items.Add(popForm.DefName+":"+(",".join(GroupList)))
    self.GroupDict[popForm.DefName] = []
    self.GroupDict[popForm.DefName].append(popForm.DefPath)
    self.GroupDict[popForm.DefName].append(GroupList[:])

  def UngroupAny(self, sender, e):
    SelRemList = list(self._ConnList.SelectedItemTexts)
    AllRemList = list(self._ConnList.Items)[:]
    self._ConnList.Items.Clear()
    for AllRem3 in AllRemList:
      if not(AllRem3.Text in SelRemList):
        self._ConnList.Items.Add(AllRem3)
    for EachItm3 in SelRemList:
      GroupName = EachItm3.split(":")[0]
      for EachName in self.GroupDict[GroupName][1]:
        self._InWgList.Items.Add(EachName)
      self.GroupDict.pop(GroupName, None)

  def RedAssign(self, sender, e):
    self.FinalGroupDict = self.GroupDict.copy()
    self.FinalInWdgList = []
    self.FinalConnList = []
    for EachFitm in self._InWgList.Items:
      self.FinalInWdgList.append(EachFitm.Text)
    for EachFitm2 in self._ConnList.Items:
      self.FinalConnList.append(EachFitm2.Text)
    self.Hide()

  def RedHide(self, sender, e):
    self.Hide()

class WdgForm(Ansys.UI.Toolkit.Dialog):

  def __init__(self):
    self._WgLabel = Ansys.UI.Toolkit.Label()
    self._DefLabel = Ansys.UI.Toolkit.Label()
    self._WgList = Ansys.UI.Toolkit.ListBox()
    self._DefList = Ansys.UI.Toolkit.ListBox()
    self._PrimButton = Ansys.UI.Toolkit.Button()
    self._AllPrimButton = Ansys.UI.Toolkit.Button()
    self._SecButton = Ansys.UI.Toolkit.Button()
    self._RemoveButton = Ansys.UI.Toolkit.Button()
    self._AllSecButton = Ansys.UI.Toolkit.Button()
    self._RedOK = Ansys.UI.Toolkit.Button()
    self._RedCancel = Ansys.UI.Toolkit.Button()
    #
    # ExLabel
    #
    self._WgLabel.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._WgLabel.Location = Ansys.UI.Toolkit.Drawing.Point(20, 20)
    self._WgLabel.Name = "ExLabel"
    self._WgLabel.Size = Ansys.UI.Toolkit.Drawing.Size(120, 23)
    self._WgLabel.Text = "Available Layers"
    #
    # ReduceLabel
    #
    self._DefLabel.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._DefLabel.Location = Ansys.UI.Toolkit.Drawing.Point(330, 20)
    self._DefLabel.Name = "ReduceLabel"
    self._DefLabel.Size = Ansys.UI.Toolkit.Drawing.Size(161, 23)
    self._DefLabel.Text = "Defined Windings"
    #
    # ExcList
    #
    self._WgList.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._WgList.Location = Ansys.UI.Toolkit.Drawing.Point(20, 45)
    self._WgList.Name = "ExcList"
    self._WgList.IsMultiSelectable = True
    self._WgList.Size = Ansys.UI.Toolkit.Drawing.Size(170, 250)
    self._WgList.SelectionChanged += self.UpdateGUI
    #
    # ReduceList
    #
    self._DefList.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._DefList.Location = Ansys.UI.Toolkit.Drawing.Point(330, 45)
    self._DefList.Name = "ReduceList"
    self._DefList.IsMultiSelectable = True
    self._DefList.Size = Ansys.UI.Toolkit.Drawing.Size(170, 250)
    self._DefList.SelectionChanged += self.UpdateGUI
    #
    # primary
    #
    self._PrimButton.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._PrimButton.Location = Ansys.UI.Toolkit.Drawing.Point(200, 50)
    self._PrimButton.Name = "PrimButton"
    self._PrimButton.Size = Ansys.UI.Toolkit.Drawing.Size(120, 27)
    self._PrimButton.Text = "Primary >"
    self._PrimButton.Click += self.PrimAdd
    #
    # secondary
    #
    self._SecButton.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._SecButton.Location = Ansys.UI.Toolkit.Drawing.Point(200, 100)
    self._SecButton.Name = "SecButton"
    self._SecButton.Size = Ansys.UI.Toolkit.Drawing.Size(120, 27)
    self._SecButton.Text = "Secondary >"
    self._SecButton.Click += self.SecAdd
    #
    # RemoveButton
    #
    self._RemoveButton.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._RemoveButton.Location = Ansys.UI.Toolkit.Drawing.Point(200, 150)
    self._RemoveButton.Name = "RemoveButton"
    self._RemoveButton.Size = Ansys.UI.Toolkit.Drawing.Size(120, 27)
    self._RemoveButton.Text = "< Remove"
    self._RemoveButton.Click += self.RemoveAny
    #
    # all to primary
    #
    self._AllPrimButton.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._AllPrimButton.Location = Ansys.UI.Toolkit.Drawing.Point(200, 200)
    self._AllPrimButton.Name = "PrimButton"
    self._AllPrimButton.Size = Ansys.UI.Toolkit.Drawing.Size(120, 27)
    self._AllPrimButton.Text = "All Primary >>"
    self._AllPrimButton.Click += self.allPrim
    #
    # all to secondary
    #
    self._AllSecButton.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._AllSecButton.Location = Ansys.UI.Toolkit.Drawing.Point(200, 250)
    self._AllSecButton.Name = "SecButton"
    self._AllSecButton.Size = Ansys.UI.Toolkit.Drawing.Size(120, 27)
    self._AllSecButton.Text = "All Secondary >>"
    self._AllSecButton.Click += self.allSec

    #
    # RedOK
    #
    self._RedOK.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._RedOK.Location = Ansys.UI.Toolkit.Drawing.Point(330, 325)
    self._RedOK.Name = "RedOK"
    self._RedOK.Size = Ansys.UI.Toolkit.Drawing.Size(80, 27)
    self._RedOK.Text = "OK"
    self._RedOK.Click += self.RedAssign
    #
    # RedCancel
    #
    self._RedCancel.Font = Ansys.UI.Toolkit.Drawing.Font("Microsoft Sans Serif", 9.75, Ansys.UI.Toolkit.Drawing.FontStyle.Normal)
    self._RedCancel.Location = Ansys.UI.Toolkit.Drawing.Point(420, 325)
    self._RedCancel.Name = "RedCancel"
    self._RedCancel.Size = Ansys.UI.Toolkit.Drawing.Size(80, 27)
    self._RedCancel.Text = "Cancel"
    self._RedCancel.Click += self.RedHide
    #
    # WindingDefinition
    #
    self.ClientSize = Ansys.UI.Toolkit.Drawing.Size(519, 375)
    self.Controls.Add(self._RedCancel)
    self.Controls.Add(self._RedOK)
    self.Controls.Add(self._DefLabel)
    self.Controls.Add(self._RemoveButton)
    self.Controls.Add(self._DefList)
    self.Controls.Add(self._PrimButton)
    self.Controls.Add(self._SecButton)
    self.Controls.Add(self._AllPrimButton)
    self.Controls.Add(self._AllSecButton)
    self.Controls.Add(self._WgLabel)
    self.Controls.Add(self._WgList)
    self.MaximizeBox = False
    self.Name = 'WindingDefinition'
    self.Text = 'Define Windings'


  def UpdateGUI(self, sender, e):
    if len(self._WgList.SelectedItems) > 0:
      self._SecButton.Enabled = True
      self._PrimButton.Enabled = True
    else:
      self._SecButton.Enabled = False
      self._PrimButton.Enabled = False
    if len(self._DefList.SelectedItems) > 0:
      self._RemoveButton.Enabled = True
    else:
      self._RemoveButton.Enabled = False

  def SecAdd(self, sender, e):
    AllItemList = list(self._WgList.Items) [:]
    SelItemList2 = sorted(list(self._WgList.SelectedItemTexts))
    self._WgList.Items.Clear()
    for AllRem in AllItemList:
      if not(AllRem.Text in SelItemList2):
        self._WgList.Items.Add(AllRem)
    for SelItem2 in SelItemList2:
      self.SecList.append(SelItem2)
      Indx1 = SelItem2.replace("Layer","Secondary")
      self._DefList.Items.Add(Indx1)

  def PrimAdd(self, sender, e):
    AllItemList2 = list(self._WgList.Items)[:]
    SelItemList = sorted(list(self._WgList.SelectedItemTexts))
    self._WgList.Items.Clear()
    for AllRem2 in AllItemList2:
      if not(AllRem2.Text in SelItemList):
        self._WgList.Items.Add(AllRem2)
    for SelItem in SelItemList:
      self.PrimList.append(SelItem)
      Indx2 = SelItem.replace("Layer","Primary")
      self._DefList.Items.Add(Indx2)

  def allPrim(self, sender, e):
    AllItemList2 = list(self._WgList.Items)[:]
    self._WgList.Items.Clear()
    for SelItem in AllItemList2:
      self.PrimList.append(SelItem.Text)
      Indx2 = SelItem.Text.replace("Layer","Primary")
      self._DefList.Items.Add(Indx2)

  def allSec(self, sender, e):
    AllItemList = list(self._WgList.Items) [:]
    self._WgList.Items.Clear()
    for SelItem2 in AllItemList:
      self.SecList.append(SelItem2.Text)
      Indx1 = SelItem2.Text.replace("Layer","Secondary")
      self._DefList.Items.Add(Indx1)

  def RemoveAny(self, sender, e):
    SelRemList = list(self._DefList.SelectedItemTexts)
    AllRemList = list(self._DefList.Items)[:]
    self._DefList.Items.Clear()
    for AllRem3 in AllRemList:
      if not(AllRem3.Text in SelRemList):
        self._DefList.Items.Add(AllRem3)
    for EachItm3 in SelRemList:
      if "Primary" in EachItm3:
        self._DefList.Items.Remove(EachItm3)
        Indx3 = EachItm3.replace("Primary","Layer")
        self._WgList.Items.Add(Indx3)
        self.PrimList.remove(Indx3)
      if "Secondary" in EachItm3:
        self._DefList.Items.Remove(EachItm3)
        Indx3 = EachItm3.replace("Secondary","Layer")
        self._WgList.Items.Add(Indx3)
        self.SecList.remove(Indx3)

  def RedAssign(self, sender, e):
    if len(self._WgList.Items) > 0:
      MessageBox.Show(self,
              "Windings are not Defined.\nPlease define all layers as windings",
              "Error",
              MessageBoxType.Error,
              MessageBoxButtons.OK, MessageBoxDefaultButton.Button1)
      return
    self.FinalPrimList = self.PrimList[:]
    self.FinalSecList = self.SecList[:]
    self.FinalWdgList = []
    self.FinalDefList = []
    for EachFitm in self._WgList.Items:
      self.FinalWdgList.append(EachFitm.Text)
    for EachFitm2 in self._DefList.Items:
      self.FinalDefList.append(EachFitm2.Text)
    self.Hide()

  def RedHide(self, sender, e):
    self.Hide()

class TabularDataEditor:
  DialogName = "TabularDataDialog"
  ComponentName = "TabularDataEditorDialog"

  def __init__(self, api, step, prop):
    self.dialog = None
    self.prop = prop
    self.propClone = None
    self.step = step

  def buttonClicked(self, sender, args):
    tabularDataComp = self.dialog.GetComponent(TabularDataEditor.ComponentName)

    # no "our" dialog
    if self.prop != tabularDataComp.GetPropertyTable():
      return

    if args.ButtonName == "Cancel":
      self.prop.Cancel()
      self.prop.CopyFrom(self.propClone, True)
      # we change the property so we have to refresh the compoenent with the "old" table
      tabularDataComp.SetPropertyTable(self.prop)
      tabularDataComp.Refresh()
    elif args.ButtonName == "Apply":
      # update properties and finish buttons states
      self.prop.Apply()
      self.step.NotifyChange()
    self.dialog.Hide()

  def createTabularDialog(self, panel, title, width, height):
    dialog = panel.CreateDialog(TabularDataEditor.DialogName, title, width, height)

    layoutDef = LayoutDefinition()
    layoutDef.Extension = None

    layout = Ansys.ACT.Core.XmlDataModel.UI.Layout()
    layoutDef.Layout = layout
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
    component.CustomCSSFile= str(ExtAPI.Extension.InstallDir) + r'\testComp.css'
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

    dialog.SetLayout(layoutDef)

    buttonComp = dialog.GetComponent("Buttons")
    buttonComp.AddButton("Apply", "Apply", Ansys.ACT.Interfaces.UserInterface.ButtonPositionType.Center)
    buttonComp.AddButton("Cancel", "Cancel", Ansys.ACT.Interfaces.UserInterface.ButtonPositionType.Center)

    return dialog

  def onshow(self, step, prop):
    self.dialog = step.UserInterface.GetComponent(TabularDataEditor.DialogName)
    if self.dialog is None:
      self.dialog = self.createTabularDialog(step.UserInterface.Panel, "TabularData", 100, 100)

    buttonComp = self.dialog.GetComponent("Buttons")
    buttonComp.ButtonClicked += self.buttonClicked

  def onhide(self, step, prop):
    self.dialog = None

  def onactivate(self, step, prop):
    dialogWidth = 532

    self.dialog.GetComponent(TabularDataEditor.ComponentName).SetPropertyTable(prop)
    self.dialog.GetComponent(TabularDataEditor.ComponentName).Refresh()
    self.dialog.SetDatas("TabularData", dialogWidth, dialogWidth)
    self.dialog.Show()

    # we save all value to take it back if we press the cancel button
    self.propClone = prop.Clone()

  def value2string(self, step, prop, value):
    return "Tabular Data"
