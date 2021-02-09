// copyright 2021, ANSYS Inc. Software is released under GNU license
define(
	[
		'scripts/knockout-3.3.0', 'scripts/ACTComponent.js',
		'jqx/jqxcore', 'jqx/jqxtooltip', 'jqx/jqxbuttons', 'jqx/jqxscrollbar',
		'jqx/jqxlistbox', 'jqx/jqxdropdownlist', 'jqx/jqxcheckbox',
		'jqx/jqxgrid', 'jqx/jqxgrid.edit.js', 'jqx/jqxinput',
		'jqx/jqxgrid.selection.js', 'jqx/jqxmenu','jqx/jqxdata',
	],
	function(ko, ACTComponent)
{
	function TabularDataComponentViewModel(params, element) {
		// when the context change we don't have variable "this" for reference the viewModel but we still have self
		// we could do : bind(this) but sometimes we need the variable "this"
		var self = this;

		self.base       = new ACTComponent($(element), params);
		self.params     = params;
		self.element    = element;
		self.rowCount   = 0;
		self.onlyOneSelect = true;
		self.properties = [];
		self.decimalSeparator = params.Model.DecimalSeparator;

		// for options selectedIndex
		self.currentEditor = null;

		// focus
		self.wantToGoNextRow     = false;
		self.cellIndexesOnEdit   = null;
		self.valueAtTheEndOfEdit = null;

		// jqx
		self.columnWidth        = 50;
		self.rows               = [];
		self.columns            = [];
		self.columnsIndexByName = {};


		self.base.onRefresh = function(obj) {
		    self.onlyOneSelect = obj.OnlyOneSelect;
		    self.columnWidth = obj.ColumnWidth

			

			// init and refresh
			if (!obj.PropertiesRows) {
				self.params.ready(true);
				return;
			}

			var properties = obj.PropertiesRows[0].RootProperty.Properties;
			var updateColumns = self.columns.length != properties.length;
			if (!updateColumns) {
				for (var i = 0; i < self.columns.length; i++) {
					var column = self.columns[i];
					var prop = properties[i];
					if (prop.Name != column.datafield || prop.Caption != column.text) {
						updateColumns = true;
						break;
					}
				}
			}

			var currentRowCount = self.rowCount;
			self.rowCount = obj.RowCount;

			if (!updateColumns) { // only update the modified rows
				if (self.rowCount > currentRowCount) {
					for (var i = currentRowCount; i < self.rowCount; i++) {
						self.rows.push({});
					}
				} else if (self.rowCount < currentRowCount) {
					for (var i = currentRowCount; i > self.rowCount; i--) {
						self.rows.pop();
						gridElement.jqxGrid('deleterow', i-1);
					}
				}

				self.properties = []
				var rows = [];
				var rowsIndexes = [];
				for (var i = 0; i < self.rowCount; i++) {
					var alreadyPushed = false;
					var props = obj.PropertiesRows[i].RootProperty;

					for (var key in props.Properties) {
						var prop = props.Properties[key];
						prop.ready = ko.observable(false);
						self.base.addCallbackX(prop, "onchange", "ValueChanged");
						self.base.addCallbackX(prop, "onclick", "ButtonClicked");

						if (!self.rows[i][prop.Name] || self.rows[i][prop.Name] != prop.DisplayValue) {
							self.rows[i][prop.Name] = prop.DisplayValue;
							if (!alreadyPushed) {
								rows.push(self.rows[i]);
								rowsIndexes.push(i);
							}
							alreadyPushed = true;
						}
					}

					self.properties.push(props.Properties);
				}

				// we do the addrow here because elsewhere he will ask for value in self.properties that haven't been add yet
				if (self.rowCount > currentRowCount) {
					for (var i = currentRowCount; i < self.rowCount; i++)
						gridElement.jqxGrid('addrow', null, {});
				}

				gridElement.jqxGrid('updaterow', rowsIndexes, rows);

				// if values are different that mean that the old value was incorect. So we keep the same focus.
				EditManager.BlockEdition(false);
				var editionOccur = self.cellIndexesOnEdit != null;
				if (editionOccur)
				{
					var oldValue = self.valueAtTheEndOfEdit;
					var rowIndex = self.cellIndexesOnEdit[0];
					var columnIndex = self.cellIndexesOnEdit[1];

					// we have to check the isn't been deleted
					if (rowIndex < self.properties.length && columnIndex < self.properties[rowIndex].length)
					{
						var newValue = self.properties[rowIndex][columnIndex].DisplayValue;
						var oldValueWasIncorrect = oldValue != "" && newValue != oldValue;
						if (self.wantToGoNextRow) {
							GoToNextFocus(oldValueWasIncorrect);
							self.wantToGoNextRow = false;
						} else if (oldValueWasIncorrect) {
							var prop = self.properties[rowIndex][columnIndex];
							if (!self.gotIsOwnEditor(prop)){
								gridElement.jqxGrid("setCellValue", rowIndex, prop.Name, newValue);
								gridElement.jqxGrid("begincelledit", rowIndex, prop.Name);
								gridElement.jqxGrid("endcelledit", rowIndex, prop.Name);
								self.cellIndexesOnEdit = null;
							}
						}
					}
				}
			} else {  // create a new jqxGrid
				// safe option cause the update of the properties will call callback on unknown properties on columns/rows
				self.properties = []
				for (var i = 0; i < self.rowCount; i++) {
					var props = obj.PropertiesRows[i].RootProperty;
					for (var key in props.Properties) {
						var prop = props.Properties[key];
						prop.ready = ko.observable(false);
						self.base.addCallbackX(prop, "onchange", "ValueChanged");
						self.base.addCallbackX(prop, "onclick", "ButtonClicked");
					}
					self.properties.push(props.Properties);
				}

				// columns
				self.columns = new Array();
				self.columnsIndexByName = {};

				var firstRowProperties = obj.PropertiesRows[0].RootProperty.Properties;
				for (var i = 0; i < firstRowProperties.length; i++) {
					var prop = firstRowProperties[i];
					self.columnsIndexByName[prop.Name] = i;

					var title = prop.Caption;
					if (title == "")
						title = prop.Name;
					if (prop.Unit)
						title += " [" + prop.Unit + "]";

					var columnType = "textbox";
					if (prop.Component == "property-select") {
						columnType = "dropdownlist";
					} else if (self.gotIsOwnEditor(prop)) {
						columnType = "template";
					}

					if (i==0) {columnWidth = 120;}
                    else if (i==1) {columnWidth = 130;}
                    else if (i==2) {columnWidth = 80;}
                    else if (i==3) {columnWidth = 120;}
                    else if (i==4) {columnWidth = 80;}

                   
                        var column = {
                            text: title,
                            datafield: prop.Name,
                            cellsrenderer : self.Cellsrenderer,
                            width: columnWidth,
                            columntype: columnType,
                            createeditor: self.CreateEditor,
                            initeditor: self.InitEditor,
                            geteditorvalue: columnType == "template" ? self.GetEditorValue : null,
                            cellclassname: self.CellClassname,
                    }; 
					self.columns.push(column);
				}

				// source
				self.rows = new Array();
				for (var i = 0; i < obj.RowCount; i++) {
					var row = {};
					for (var j = 0; j < obj.PropertiesRows[i].RootProperty.Properties.length; j++) {
						var prop = obj.PropertiesRows[i].RootProperty.Properties[j];
						row[prop.Name] = prop.DisplayValue;
					}
					self.rows.push(row);
				}

				gridElement.jqxGrid('clear');
				var dataAdapter = new $.jqx.dataAdapter( { localdata: self.rows, datatype: "array"}, {} );

				gridElement.jqxGrid({
					columns: self.columns,
					source: dataAdapter,
					cellhover: self.Cellhover,
					disabled: false,
					editable: true,
					editmode: 'click',
					height: '100%',
					width: '100%',
					selectionmode: obj.CanSelect ? "singlecell" : "True",  //multiplecellsadvanced 
				});
			}

		

			// valid / invalid
			var $headers = $(element).find(".jqx-grid-column-header.jqx-widget-header");
			var errorColor = "#deb7b1";
			if (obj.RowCount == 0) {
				$headers.css("background", errorColor);
			} else {
				$headers.css("background", "");
				if (obj.IsValid) {
					$($headers[0]).css("background", "");
				} else {
					$($headers[0]).css("background", errorColor);
				}

				for (var col = 0; col < obj.PropertiesRows[0].RootProperty.Properties.length; col++) {
					for (var row = 0; row < obj.PropertiesRows.length; row++) {
						var prop = obj.PropertiesRows[row].RootProperty.Properties[col];
						if (prop.IsVisible && prop.StateMessage)
						{
							$($headers[col + 1]).css("background", errorColor);
							break;
						}
					}
				}
			}

			self.params.ready(true);
		}.bind(self);

		self.gotIsOwnEditor = function(prop) {
			return prop.Component == "property-fileopen" || prop.Component == "property-applycancel" ||
				prop.Component == "property-custom" || prop.Component == "property-apply";
		}

		var contextMenuElement = $(element).find(".menu");
		contextMenuElement.find("#delete").text(self.transTextDeleteRow);

		var gridElement = $(element).find(".jqxgrid");
		var rowUnderContextMenu = 0;

		var contextMenu = contextMenuElement.jqxMenu({
			width: 200,
			autoOpenPopup: false,
			mode: 'popup',
			popupZIndex: 99999999   // to be on top of dialogboxes
		});

		contextMenuElement.on('itemclick', function(event) {
			var rowIndex = rowUnderContextMenu;
			params.baseModel.onTriggerAction("TriggerDeleteRowButtonClick", [rowIndex]);
			self.deleteRowClick([rowIndex]);
		});

		contextMenuElement.on("closed", function() {
			TooltipManager.ActivateTooltip();
		});

		// remove default context menu
		gridElement.on("contextmenu", function(e) { return false; });

		gridElement.on('rowclick', function(event) {
			// context menu
			if (event.args.rightclick) {
				// because the tooltip appear at the same position as the context menu
				TooltipManager.DesactivateTooltip();

				var rowIndex = event.args.rowindex;
				rowUnderContextMenu = rowIndex;

				var scrollTop = $(window).scrollTop();
				var scrollLeft = $(window).scrollLeft();
				contextMenu.jqxMenu(
					'open',
					parseInt(event.args.originalEvent.clientX) + 5 + scrollLeft,
					parseInt(event.args.originalEvent.clientY) + 5 + scrollTop
				);
			}
		});

		var GoToNextFocus = function(focusTheSame) {
			var rowIndex    = self.cellIndexesOnEdit[0];
			var columnIndex = self.cellIndexesOnEdit[1];

			var nextRow = rowIndex;
			var nextColumn = columnIndex;
			if (!focusTheSame) {
				do {
					nextRow = nextRow + 1;
					if (nextRow >= self.properties.length) {
						nextRow = 0;
						nextColumn = (nextColumn + 1) % self.properties[nextRow].length;
					}
				} while(!self.properties[nextRow][nextColumn].IsVisible || self.properties[nextRow][nextColumn].Component == "property-readonly");
			}
			gridElement.jqxGrid('begincelledit', nextRow, self.properties[nextRow][nextColumn].Name);
		}

		// to avoid start editing between messages
		var EditManager = new function() {
			this.editIsBlocked    = false;
			this.keepedCellInEdit = null;

			this.BlockEdition = function(block) {
				if (!block && this.keepedCellInEdit) {
					// timeout because updateRow take time and when it's happens we lost focus
					setTimeout(
						function() {
							gridElement.jqxGrid('begincelledit', EditManager.keepedCellInEdit[0], EditManager.keepedCellInEdit[1]);
							EditManager.keepedCellInEdit = null;
						}, 0
					);
				}
				this.editIsBlocked = block;
			}

			this.IsEditBlocked = function() {
				return this.editIsBlocked;
			}

			this.KeepCellInEdit = function(rowIndex, columnDataField) {
				this.keepedCellInEdit = [rowIndex, columnDataField];
			}
		}

		var TooltipManager = new function() {
			this.tooltipIsOpenOn = null;
			this.tooltipIsActive = true;
			this.initialized     = false;

			this.CloseTooltip = function() {
				// we disable it to be sure only us can open it
				gridElement.jqxTooltip({ disabled: true });
				gridElement.jqxTooltip('close');
				this.tooltipIsOpenOn = null;
			}

			this.OpenTooltip = function(strTooltip, element, posX, posY) {
				if (!this.initialized) {
					gridElement.jqxTooltip({opacity: 0.8, autoHide: false, animationShowDelay: 0, animationHideDelay: 0});
					this.initialized = true;
				}

				if (!this.tooltipIsActive)
					return;

				gridElement.jqxTooltip({ disabled: false });
				gridElement.jqxTooltip({ content: strTooltip });
				gridElement.jqxTooltip('open', posX + 10, posY + 10);

				if (element != this.tooltipIsOpenOn) {
					$(element).one("mouseleave", function(eventData) {
						TooltipManager.CloseTooltip();
					});
					this.tooltipIsOpenOn = element;
				}
			}

			this.DesactivateTooltip = function() {
				this.CloseTooltip();
				this.tooltipIsActive = false;
			}

			this.ActivateTooltip = function() {
				this.tooltipIsActive = true;
			}
		}

		gridElement.bind('cellbeginedit', function(event) {
			var columnDataField = args.datafield;
			var rowIndex        = args.rowindex;

			if (EditManager.IsEditBlocked()) {
				EditManager.KeepCellInEdit(rowIndex, columnDataField);
				gridElement.jqxGrid('endcelledit', rowIndex, columnDataField, true);
				return;
			}

			TooltipManager.DesactivateTooltip();

			var columnIndex = self.columnsIndexByName[columnDataField];
			self.cellIndexesOnEdit = [rowIndex, columnIndex];

			// user shouldn't be able to edit not visible and readonly properties, so we end it now
			var prop = self.properties[rowIndex][columnIndex];
			if (!prop.IsVisible || prop.Component == "property-readonly") {
				gridElement.jqxGrid('endcelledit', rowIndex, columnDataField, true);
			}
		});

		gridElement.on('cellendedit', function(event) {
			if (EditManager.IsEditBlocked())
				return;

			TooltipManager.ActivateTooltip();

			var args            = event.args;
			var columnDataField = args.datafield;
			var rowIndex        = args.rowindex;
			var cellValue       = args.value;

			var columnIndex = self.columnsIndexByName[columnDataField];
			var prop = self.properties[rowIndex][columnIndex];
			var cellOldValue = prop.DisplayValue;

			// cancel if click on other cell without clicking on apply
			if (prop.Component == "property-applycancel" && !prop.applyClick) {
				prop.onclick(rowIndex, "cancel");
			}

			// if not property-xxx because the editor already call this change
			if ( cellValue != cellOldValue && !self.gotIsOwnEditor(prop))
			{
				params.baseModel.onTriggerAction("TriggerPropertyValueChange", [rowIndex, prop.Name, cellValue]);

				// block editing until we receive the feedback
				EditManager.BlockEdition(true);
				self.valueAtTheEndOfEdit = cellValue;
				if (prop.Component == "property-select") {
					var selectedOptionIndex = self.currentEditor.jqxDropDownList("selectedIndex");
					prop.onchange(rowIndex, selectedOptionIndex);
				} else {
					prop.onchange(rowIndex, cellValue);
				}
			}
			else if (self.wantToGoNextRow) {
				// timeout because we are not after but before we close the editor.
				setTimeout(
					function() {
						GoToNextFocus(false);
						self.wantToGoNextRow = false;
					}, 0
				);
			}
		});

		gridElement.on('keydown', function(eventData) {
			// enter key
			if (eventData.keyCode == 13) {
				self.wantToGoNextRow = true;
			}
		});

		
		self.Cellsrenderer = function(row, columnfield, value, defaulthtml, columnproperties, rowdata) {
			// ! no html ko databind here ! (http://www.jqwidgets.com/community/topic/custom-cell-renderer-with-knockout-bindings/)
			var columnIndex = self.columnsIndexByName[columnproperties.datafield];
			var prop = self.properties[row][columnIndex];

			var element = $(defaulthtml);

			if (prop.IsVisible && prop.Component != "property-readonly") {
				var tooltipMessage = prop.StateMessage != "" && prop.StateMessage || prop.DisplayValue
				element.attr("tooltip", tooltipMessage);
				if (prop.Unit && value != "") {
					element.text(value + " [" + prop.Unit + "]");
				}
			} else if (!prop.IsVisible) {
			    element.text("");
			} else {
			    var tooltipMessage = prop.StateMessage != "" && prop.StateMessage || prop.DisplayValue
			    element.attr("tooltip", tooltipMessage);
			}

			return element[0].outerHTML;
		};

		// special editor
		self.CreateEditor = function(row, cellvalue, editor) {
			// ! don't use the row variable here it is useless : this function is called only once for all the column
			var datafield   = this.datafield;
			var columnIndex = self.columnsIndexByName[datafield];
			var prop        = self.properties[row][columnIndex];

			if (prop.Component == "property-select") {
				var list = prop.Options;
				editor.jqxDropDownList({ autoDropDownHeight: true, source: list });
			} else if (self.gotIsOwnEditor(prop))
			{
				var container = $("<center/>");
				editor.append(container);

				var input = $("<input class=\"textInput\"/>");
				input.jqxInput({searchMode:false, width:self.columnWidth*0.5});
				input.on("change", function() {
					var value     = $(this).val();
					var lastValue = $(this).attr("lastValue");
					var row       = parseInt($(this).attr("row"));
					var datafield = $(this).attr("datafield");
					var column    = self.columnsIndexByName[datafield];
					var prop      = self.properties[row][column];

					if (lastValue != value) {
						params.baseModel.onTriggerAction("TriggerPropertyValueChange", [row, prop.Name, value]);
						self.valueAtTheEndOfEdit = value;
						prop.onchange(row, value);
						$(this).attr("lastValue", value);
					}
				});
				container.append(input);

				var button = $("<input class=\"buttonInput\" type=\"button\"/>");
				button.jqxButton({width:self.columnWidth*0.4});
				button.on("click", function() {
					// because row, datafield and button change but "this" variable really correspond to the button
					var row       = parseInt($(this).attr("row"));
					var datafield = $(this).attr("datafield");
					var column    = self.columnsIndexByName[datafield];
					var prop      = self.properties[row][column];

					params.baseModel.onTriggerAction("TriggerPropertyButtonClick", [row, prop.Name, ""]);
					prop.onclick(row, prop.Component == "property-custom" ? "edit" : "apply");
					prop.applyClick = true;
					gridElement.jqxGrid("endcelledit", row, datafield);
				});
				container.append(button);

				if (prop.Component == "property-fileopen") {
					button.val("...");
				} else if (prop.Component == "property-custom") {
					var label = prop.Attributes["buttonlabel"] || "Edit";
					button.val(label);
				} else if (prop.Component == "property-apply" || prop.Component == "property-applycancel") {
					var label = prop.Attributes["buttonlabel"] || "Apply";
					input.jqxInput({disabled: true});
					button.val(label);
					prop.applyClick = false;
				}
			}
		};

		self.InitEditor = function(row, cellvalue, editor) {
			// update input and button attributes (row and column indexes)
			var datafield = this.datafield;
			var column    = self.columnsIndexByName[datafield];
			var prop      = self.properties[row][column];

			if (self.gotIsOwnEditor(prop))
			{
				var inputElement     = editor.find("input");
				var textInputElement = editor.find(".textInput");

				// We have to trigger for the input the change because when we reopen
				// an other property with that input the change event isn't called
				if (textInputElement.attr("row")) {  // If the attrs are defined
					textInputElement.trigger("change");
				}

				inputElement.attr("row", row);
				inputElement.attr("datafield", datafield);

				var value = prop.DisplayValue;
				inputElement.attr("lastValue", value);
				textInputElement.val(value);

				// focus the textInput
				// timeout to pass through the click on the cell (cellbeginedit happen before initEditor and they are no event after)
				setTimeout(function(textInputElement) {
					textInputElement.focus()
				}, 0, textInputElement);
			}
			if (prop.Component == "property-applycancel") {
				prop.onclick(row, "edit");
			}

			self.currentEditor = editor;
		}

		self.GetEditorValue = function(row, cellvalue, editor) {
			var columnIndex = self.columnsIndexByName[this.datafield];
			return self.properties[row][columnIndex].DisplayValue;
		}

		// for cell's background color (error, readonly, ...)
		self.CellClassname = function(rowIndex, columnDataField, value) {
			var columnIndex = self.columnsIndexByName[columnDataField];
			var prop        = self.properties[rowIndex][columnIndex];

			if (!prop.IsVisible) {
				return "cellNotVisible";
			} else if (prop.Component == "property-readonly") {
			    if (prop.StateMessage) {
			        return "cellReadOnly cellError";
			    }
			    else {
			        return "cellReadOnly";
			    }
			} else if (prop.StateMessage) {
				return "cellError";
			}
		}

		self.Cellhover = function(element, pageX, pageY) {
			var attrTooltip = $(element).children().attr("tooltip");
			if (attrTooltip) {	
				TooltipManager.OpenTooltip(attrTooltip, element, pageX, pageY);
			} else {
				TooltipManager.CloseTooltip();
			}
		}

		// selection
		self.ClearSelection = function() {
			var selectedRows = gridElement.jqxGrid('getselectedrowindexes');
			if (selectedRows != null && selectedRows.length != 0)
				self.selectRowClick([]);
		}

		self.ResetSelection = function() {
			self.ClearSelection();
			var selectedRows = gridElement.jqxGrid('getselectedrowindexes');
			self.selectRowClick(selectedRows);
		}

		

		gridElement.on('rowunselect', function (event) {
			if (self.justDelete)
				return;

			var args = event.args;
			var rowIndex = args.rowindex;  // can be an array
			if (Array.isArray(rowIndex)) {
				params.baseModel.onTriggerAction("TriggerSelectUnselectAllRowsClick");  // if it's an array that should be the all select checkbox
				self.unselectRowClick(rowIndex);
			} else {
				params.baseModel.onTriggerAction("TriggerSelectUnselectRowClick", [rowIndex]);
				self.unselectRowClick([rowIndex]);
			}
		});

		params.baseModel.resizeFunctions[params.baseModel.resizeFunctions.length] = function(height,width) {
			// if we don't call refresh here the grid doesn't update it size if it contain rows
			gridElement.jqxGrid("refresh");

			// ugly but seems to fix a bug where the horizontal scrollbar is correctly computed when the grid is visible (soon after this call)
			if (!self.notFirstResizeRefresh) {
				self.notFirstResizeRefresh = true;
				setTimeout(function() {
					gridElement.jqxGrid("refresh");
				}, 100);
			}
		}.bind(this);

		// replace numpad dot -> decimalSeparator
		$(element).on("keydown", function(e) {
			if (e.keyCode == 110) { // dot
				var $target = $(e.target);
				if ($target.prop("tagName").toLowerCase() == "input")
				{
					$target.val($target.val() + self.decimalSeparator);
					e.preventDefault();
					return false;
				}
			}
		});


		self.base.initialize();
	}

	return {
		createViewModel: function(params, componentInfo) {
			return new TabularDataComponentViewModel(params, componentInfo.element);
		}
	};

});

