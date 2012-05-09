/*
 *   copyright (c) -- chaitanya nalla and kechit goyal
 *
 */

function defineDionGrid() {

	dataObj = function() {
		this.type = "";
		this.cssClass = "";
		this.options = null;
		this.value = "";
		this.name = "";
	}

	changeHandler = function() {
		this.callBacks = [];
	};

	changeHandler.prototype.subscribe = function(fn) {
		this.callBacks.push(fn);
	};

	changeHandler.prototype.raiseCallBacks = function(e, args) {
		for ( var i = 0; i < this.callBacks.length; i++) {
			this.callBacks[i](e, args, this);
		}
	};

	// Constructor for the grid
	clsDionGrid = function() {
		this.gridName = "";
		this.parentDivId = "";
		this.parentDivElement = null;

		this.defaultCellHeight = 1;
		this.defaultCellWidth = 1;
		this.defaultRowGap = 0;
		this.defaultColGap = 0;

		this.curXLSRowNo = 1;
		this.curXLSColNo = 1;
		this.XLSRowCount = 0;
		this.XLSColCount = 0;

		this.top = 0;
		this.left = 0;
		this.height = 0;
		this.minWidth = 0;
		this.maxWidth = 0;
		this.width = 0;
		this.minHeight = 0;
		this.maxHeight = 0;

		this.displayCellArray = [];
		this.displayRowArray = [];
		this.displayColArray = [];
		this.XLSCellArray = [];
		this.XLSRowArray = [];
		this.XLSColArray = [];

		this.gridDiv = null;
		this.interpolationDiv = null;

		this.defaultEditableCellStyle = "gridCellBorder ";
		this.defaultOddEditableCellStyle = "oddGridCellBorder ";
		this.defaultEvenEditableCellStyle = "evenGridCellBorder ";

		this.defaultNonEditableCellStyle = "gridCell ";
		this.defaultOddNonEditableCellStyle = "oddGridCell ";
		this.defaultEvenNonEditableCellStyle = "evenGridCell ";

		this.defaultGridDivStyle = "gridDiv ";
		this.defaultTextAlign = " left";
		this.defaultErrorStyle = "borderError";

		this.mode = "select";
		this.dragItem = "";
		this.interpolatedRange = "";
		//selectedRange has been moved to workbook.
		this.showInterpolationDiv = true;
		this.showContextMenu = false;
		this.isCopyPasteEnabled = true;
		this.isCustomContextMenu = false;
		this.hasAutoWidth = true;
		this.hasAutoHeight = true;
		this.contextMenuFn = null;

		this.onCellChanged = new changeHandler();
		this.onCellClicked = null;
		
		this.nextGrid = {
				left : {
						grid : null,
						row_offset : 0,
						col_offset : 0
				},
				right : {
						grid : null,
						row_offset : 0,
						col_offset : 0
				},
				up :{
						grid : null,
						row_offset : 0,
						col_offset : 0
				},
				down:{
						grid : null,
						row_offset : 0,
						col_offset : 0
				}
		};

	};

	// Creates the full screen depending on the flags
	clsDionGrid.prototype.createFullScreen = function() {
		var start = new Date();
		this.createGridDiv();

		this.createScreenGrid();

		if (this.showInterpolationDiv)
			this.createInterpolationDiv();

		this.attachGridDiv();
		var end = new Date();
		// log("Grid Initialize Time is : " + (end-start));

	};

	// Create the Grid Div ie, the Div to contain the cells
	clsDionGrid.prototype.createGridDiv = function() {
		var THIS = this;

		this.gridDiv = document.createElement("div");
		this.gridDiv.id = this.gridName + ":" + "gridDiv";
		this.gridDiv.isAttached = false;

		this.gridDiv.style.left = this.left + "px";
		this.gridDiv.style.top = this.top + "px";
		this.gridDiv.style.width = (this.width) + "px";
		this.gridDiv.style.height = (this.height) + "px";

		this.gridDiv.className = this.defaultGridDivStyle;

		// On Moving mouse out of the grid, selected range must automatically
		// increase/decrease in case of cell Drag or interpolation div drag
		// TODO
		// Replace it by worker threads to automatically keep on scrolling till
		// the mouse comes inside
		this.gridDiv.onmouseout = function(e) {
			var parentDiv = THIS.parentDivElement;
			var rightPos = this.offsetLeft
					+ getLengthFromString(this.style.width);
			var leftPos = this.offsetLeft
					+ getLengthFromString(parentDiv.style.left);
			var bottomPos = this.offsetTop
					+ getLengthFromString(this.style.height)
					+ getLengthFromString(parentDiv.style.top);
			var topPos = this.offsetTop
					+ getLengthFromString(parentDiv.style.top);

			if (THIS.dragItem == "cell") {
				if (e == null)
					e = window.event;
				var x = 0;
				var y = 0;

				if (e.pageX) {
					x = e.pageX;
					y = e.pageY;
				} else {
					x = e.clientX + document.documentElement.scrollLeft;
					y = e.clientY + document.documentElement.scrollTop;
				}
				var rangeArray, newEndRangeX, newEndRangeY;
				if (y > bottomPos && x > rightPos) {
					rangeArray = workbook.selectedRange.split(":");
					newEndRangeX = ((parseInt(rangeArray[2]) + 1) > THIS.XLSRowCount) ? parseInt(rangeArray[2])
							: (parseInt(rangeArray[2]) + 1);
					newEndRangeY = ((parseInt(rangeArray[3]) + 1) > THIS.XLSColCount) ? parseInt(rangeArray[3])
							: (parseInt(rangeArray[3]) + 1);
				} else if (y > bottomPos) {
					rangeArray = workbook.selectedRange.split(":");
					newEndRangeX = ((parseInt(rangeArray[2]) + 1) > THIS.XLSRowCount) ? parseInt(rangeArray[2])
							: (parseInt(rangeArray[2]) + 1);
					newEndRangeY = rangeArray[3];

				} else if (x > rightPos) {
					rangeArray = workbook.selectedRange.split(":");
					newEndRangeX = rangeArray[2];
					newEndRangeY = ((parseInt(rangeArray[3]) + 1) > THIS.XLSColCount) ? parseInt(rangeArray[3])
							: (parseInt(rangeArray[3]) + 1);
				} else if (y < topPos && x < leftPos) {
					rangeArray = workbook.selectedRange.split(":");
					newEndRangeX = ((parseInt(rangeArray[2]) - 1) < 1) ? 1
							: (parseInt(rangeArray[2]) - 1);
					newEndRangeY = ((parseInt(rangeArray[3]) - 1) < 1) ? 1
							: (parseInt(rangeArray[3]) - 1);

				} else if (y < topPos) {
					rangeArray = workbook.selectedRange.split(":");
					newEndRangeX = ((parseInt(rangeArray[2]) - 1) < 1) ? 1
							: (parseInt(rangeArray[2]) - 1);
					newEndRangeY = rangeArray[3];

				} else if (x < leftPos) {
					rangeArray = workbook.selectedRange.split(":");
					newEndRangeX = rangeArray[2];
					newEndRangeY = ((parseInt(rangeArray[3]) - 1) < 1) ? 1
							: (parseInt(rangeArray[3]) - 1);

				}
				if (rangeArray != undefined || rangeArray != null) {
					var range = rangeArray[0] + ":" + rangeArray[1] + ":"
							+ newEndRangeX + ":" + newEndRangeY;
					if (workbook.selectedRange != range) {
						workbook.clearRange(workbook.selectedRange);
						workbook.selectedRange = range;
						THIS.applyRange();
					}
				}
				preventDefaultBehaviour(e);
			} else if (THIS.dragItem == "interpolationDiv") {
				if (e == null)
					e = window.event;
				var x = 0;
				var y = 0;
				if (e.pageX) {
					x = e.pageX;
					y = e.pageY;
				} else {
					x = e.clientX + document.documentElement.scrollLeft;
					y = e.clientY + document.documentElement.scrollTop;
				}
				var rangeArray, newEndRangeX, newEndRangeY;
				rangeArray = THIS.interpolatedRange.split(":");
				if (y > bottomPos && x > rightPos) {
					newEndRangeX = ((parseInt(rangeArray[2]) + 1) > THIS.XLSRowCount) ? parseInt(rangeArray[2])
							: (parseInt(rangeArray[2]) + 1);
					newEndRangeY = ((parseInt(rangeArray[3]) + 1) > THIS.XLSColCount) ? parseInt(rangeArray[3])
							: (parseInt(rangeArray[3]) + 1);
				} else if (y > bottomPos) {
					newEndRangeX = ((parseInt(rangeArray[2]) + 1) > THIS.XLSRowCount) ? parseInt(rangeArray[2])
							: (parseInt(rangeArray[2]) + 1);
					newEndRangeY = rangeArray[3];
				} else if (x > rightPos) {
					newEndRangeX = rangeArray[2];
					newEndRangeY = ((parseInt(rangeArray[3]) + 1) > THIS.XLSColCount) ? parseInt(rangeArray[3])
							: (parseInt(rangeArray[3]) + 1);
				} else if (y < topPos && x < leftPos) {
					newEndRangeX = ((parseInt(rangeArray[2]) - 1) < 1) ? 1
							: (parseInt(rangeArray[2]) - 1);
					newEndRangeY = ((parseInt(rangeArray[3]) - 1) < 1) ? 1
							: (parseInt(rangeArray[3]) - 1);
				} else if (y < topPos) {
					newEndRangeX = ((parseInt(rangeArray[2]) - 1) < 1) ? 1
							: (parseInt(rangeArray[2]) - 1);
					newEndRangeY = rangeArray[3];
				} else if (x < leftPos) {
					newEndRangeX = rangeArray[2];
					newEndRangeY = ((parseInt(rangeArray[3]) - 1) < 1) ? 1
							: (parseInt(rangeArray[3]) - 1);
				}

				var startXRange = getStartX(workbook.selectedRange);
				var startYRange = getStartY(workbook.selectedRange);
				var endXRange = getEndX(workbook.selectedRange);
				var endYRange = getEndY(workbook.selectedRange);
				_endCellX = newEndRangeX;
				_endCellY = newEndRangeY;

				if (Math.abs(endYRange - _endCellY) > Math.abs(endXRange
						- _endCellX)) {
					if (_endCellY > endYRange) {
						if (THIS.interpolatedRange != "")
							THIS.borderMapper(THIS.interpolatedRange,
									"noneDashed");
						THIS.interpolatedRange = startXRange + ":"
								+ startYRange + ":" + endXRange + ":"
								+ _endCellY;
						THIS.borderMapper(THIS.interpolatedRange, "allDashed");
					} else if (_endCellY < startYRange) {
						if (THIS.interpolatedRange != "")
							THIS.borderMapper(THIS.interpolatedRange,
									"noneDashed");
						THIS.interpolatedRange = startXRange + ":" + _endCellY
								+ ":" + endXRange + ":" + endYRange;
						THIS.borderMapper(THIS.interpolatedRange, "allDashed");
					} else {
						THIS.borderMapper(THIS.interpolatedRange, "noneDashed");
						THIS.interpolatedRange = "";
					}
				} else {
					if (_endCellX > endXRange) {
						if (THIS.interpolatedRange != "")
							THIS.borderMapper(THIS.interpolatedRange,
									"noneDashed");
						THIS.interpolatedRange = startXRange + ":"
								+ startYRange + ":" + _endCellX + ":"
								+ endYRange;
						THIS.borderMapper(THIS.interpolatedRange, "allDashed");
					} else if (_endCellX < startXRange) {
						if (THIS.interpolatedRange != "")
							THIS.borderMapper(THIS.interpolatedRange,
									"noneDashed");
						THIS.interpolatedRange = _endCellX + ":" + startYRange
								+ ":" + endXRange + ":" + endYRange;
						THIS.borderMapper(THIS.interpolatedRange, "allDashed");
					} else {
						THIS.borderMapper(THIS.interpolatedRange, "noneDashed");
						THIS.interpolatedRange = "";
					}
				}
				THIS.applyRange();
			}

		};

		this.gridDiv.onmouseup = function(e) {
			THIS.clearDrag();
			// preventDefaultBehaviour(e);
		};

		this.gridDiv.onmousedown = function(e) {
			THIS.hideContextMenu();
			// preventDefaultBehaviour(e);
		};
		this.gridDiv.onmousemove = function(e) {
			// preventDefaultBehaviour(e);
		};
	};

	clsDionGrid.prototype.attachGridDiv = function() {
		if (!this.gridDiv.isAttached) {
			if (this.parentDivElement && this.parentDivElement.childNodes
					&& this.parentDivElement.childNodes.length > 0) {
				this.parentDivElement.insertBefore(this.gridDiv,
						this.parentDivElement.childNodes[0]);
			} else {
				this.parentDivElement.appendChild(this.gridDiv);
			}
			this.gridDiv.isAttached = true;
		}

	}

	clsDionGrid.prototype.removeGridDiv = function() {
		if (this.gridDiv.isAttached) {
			this.parentDivElement.removeChild(this.gridDiv);
			this.gridDiv.isAttached = false;
		}
	}

	// Initializes Input Cells
	clsDionGrid.prototype.createScreenGrid = function() {
		this.displayCellArray[0] = null;
		for ( var i = 1; i <= this.XLSRowCount; i++) {
			this.displayCellArray[i] = [];
			this.displayCellArray[i][0] = null;

			for ( var j = 1; j <= this.XLSColCount; j++) {
				var cellDiv = this.createCellDiv(i, j);
				this.displayCellArray[i][j] = cellDiv;
				this.insertFormatter("input", i, j);
			}
		}
	};

	// Initialize the XLS Grid Cells
	clsDionGrid.prototype.createXLSGridInMemory = function() {
		this.XLSCellArray[0] = null;
		this.XLSRowArray[0] = null;
		this.XLSColArray[0] = null;

		this.XLSRowArray[0] = null;
		for ( var i = 1; i <= this.XLSRowCount; i++) {
			this.XLSRowArray[i] = new clsMemoryRow(this.defaultCellHeight);
			this.XLSRowArray[i].rowNo = i;
			this.XLSRowArray[i].gap = this.defaultRowGap;
		}

		this.XLSColArray[0] = null;
		for ( var j = 1; j <= this.XLSColCount; j++) {
			this.XLSColArray[j] = new clsMemoryCol(this.defaultCellWidth);
			this.XLSColArray[j].colNo = j;
			this.XLSColArray[j].gap = this.defaultColGap;
		}

		for ( var i = 1; i <= this.XLSRowCount; i++) {
			this.XLSCellArray[i] = [];
			this.XLSCellArray[i][0] = null;
			// Since all the cells are input by default so we update it by
			// string
			for ( var j = 1; j <= this.XLSColCount; j++) {
				this.initCellInMemory(i, j);
			}
		}
	};

	// Create the Interpolation Div
	clsDionGrid.prototype.createInterpolationDiv = function() {
		var THIS = this;
		this.interpolationDiv = document.createElement("div");
		this.interpolationDiv.className = "interpolationDivClass";
		this.gridDiv.appendChild(this.interpolationDiv);

		this.interpolationDiv.onmousedown = function(e) {
			THIS.saveOneScreenCellInMemory(THIS.curXLSRowNo, THIS.curXLSColNo);
			THIS.dragItem = "interpolationDiv";
			preventDefaultBehaviour(e);
			document.onmouseup = OnMouseUp;
		};

		// TODO
		// LOOK IF THIS FUNCTION IS NEEDED OR NOT
		this.interpolationDiv.onmouseup = function(e) {
			THIS.dragItem = "";
			if (THIS.interpolatedRange != "") {
				THIS.borderMapper(THIS.interpolatedRange, "noneDashed");
				workbook.clearRange(workbook.selectedRange);
				workbook.selectedRange = THIS.interpolatedRange;
				THIS.interpolatedRange = "";
				THIS.applyRange();
			}
		};
	};
	
	clsDionGrid.prototype.initCellInMemory = function(i, j) {
		this.initCellInXLSMemory(i, j);
		this.initCellInDisplay(i, j);
	};
	

	clsDionGrid.prototype.initCellInXLSMemory = function(i, j) {
		if (!this.XLSCellArray[i]) {
			this.XLSCellArray[i] = [];
		}
		this.XLSCellArray[i][j] = new clsMemoryCell();
		this.XLSCellArray[i][j].parentXLSArray = this.XLSCellArray;
		this.XLSCellArray[i][j].parentRow = this.XLSRowArray[i];
		this.XLSCellArray[i][j].parentCol = this.XLSColArray[j];

		var cellStyle = (i % 2) ? this.defaultEvenEditableCellStyle
				: this.defaultOddEditableCellStyle;

		this.XLSCellArray[i][j].className = cellStyle + " " + this.defaultTextAlign + "AlignCellClass ";
		this.XLSCellArray[i][j].grandParentDionGrid = this;
		this.XLSCellArray[i][j].updateCellByString("");
	};
	
	clsDionGrid.prototype.initCellInDisplay = function(i, j) {
		if (!this.displayCellArray[i]) {
			this.displayCellArray[i] = [];
		}
		var cellDiv = (!this.displayCellArray[i][j]) ? this.createCellDiv(i, j) : this.displayCellArray[i][j];
		this.displayCellArray[i][j] = cellDiv;
		this.insertFormatter("input", i, j);
	};

	clsDionGrid.prototype.updateXLSCount = function(lRowNo, lColNo) {
		for ( var i = this.XLSRowCount + 1; i <= lRowNo; i++) {
			this.XLSRowArray[i] = new clsMemoryRow(this.defaultCellHeight);
			this.XLSRowArray[i].rowNo = i;
			this.XLSRowArray[i].gap = this.defaultRowGap;
		}

		this.XLSColArray[0] = null;
		for ( var j = this.XLSColCount + 1; j <= lColNo; j++) {
			this.XLSColArray[j] = new clsMemoryCol(this.defaultCellWidth);
			this.XLSColArray[j].colNo = j;
			this.XLSColArray[j].gap = this.defaultColGap;
		}

		for ( var i = 1; i <= this.XLSRowCount; i++) {
			for ( var j = this.XLSColCount + 1; j <= lColNo; j++) {
				this.initCellInMemory(i, j);
			}
		}
		for ( var i = this.XLSRowCount + 1; i <= lRowNo; i++) {
			for ( var j = 1; j <= lColNo && j <= this.XLSColCount; j++) {
				this.initCellInMemory(i, j);
			}
		}

		for ( var i = this.XLSRowCount + 1; i <= lRowNo; i++) {
			for ( var j = this.XLSColCount + 1; j <= lColNo; j++) {
				this.initCellInMemory(i, j);
			}
		}

		this.XLSRowCount = lRowNo;
		this.XLSColCount = lColNo;
	}

	// This function inserts a row at the specified index
	clsDionGrid.prototype.insertRow = function(indexNo) {
		this.saveOneScreenCellInMemory(this.curXLSRowNo, this.curXLSColNo);
		// Increase the row count
		this.updateXLSCount(this.XLSRowCount + 1, this.XLSColCount);

		// If the index is out of bounds then return
		if (indexNo < 1 || (indexNo > this.XLSRowCount))
			return;

		for ( var i = this.XLSRowCount; i > indexNo; i--) {
			// Shift the rows accordingly
			this.XLSRowArray[i] = this.XLSRowArray[i - 1];
			this.XLSRowArray[i].rowNo = i;
			this.XLSRowArray[i].gap = this.defaultRowGap;
		}

		this.XLSRowArray[indexNo] = new clsMemoryRow(this.defaultCellHeight);
		this.XLSRowArray[indexNo].rowNo = indexNo;
		this.XLSRowArray[indexNo].gap = this.defaultRowGap;

		var tempList = [];
		for ( var i = 1; i <= this.XLSColCount; i++) {
			if (this.XLSCellArray[indexNo][i].spannedBy.row != 0
					&& this.XLSCellArray[indexNo][i].spannedBy.col != 0) {
				var spannedByCell = this.XLSCellArray[this.XLSCellArray[indexNo][i].spannedBy.row][this.XLSCellArray[indexNo][i].spannedBy.col];
				if (this.XLSCellArray[indexNo][i].spannedBy.row != indexNo) {
					if (indexOfElementInArray(tempList, spannedByCell) == -1) {
						tempList.push(spannedByCell)
					}
				}
			}
		}

		for ( var i = 0; i < tempList.length; i++) {
			tempList[i].rowSpan++;
		}

		for ( var i = this.XLSRowCount; i > indexNo; i--) {
			// Shift the cells accordingly
			for ( var j = 1; j <= this.XLSColCount; j++) {
				this.XLSCellArray[i][j] = this.XLSCellArray[i - 1][j];
				this.XLSCellArray[i][j].parentRow = this.XLSRowArray[i];
				this.XLSCellArray[i][j].parentCol = this.XLSColArray[j];
			}
		}

		//Finally initialize the cells
		for ( var j = 1; j <= this.XLSColCount; j++) {
			this.initCellInXLSMemory(indexNo, j);
		}

		var affectedRange = indexNo + ":1:" + this.XLSRowCount + ":"
				+ this.XLSColCount;
		this.adjustNamedRange(affectedRange, 1, 0);

		// and update the screen
		this.updateScreenFromXLS();
		this.refreshGridCells();
		this.refreshHeaders();
	};

	// This function inserts a row at the specified index
	clsDionGrid.prototype.insertRowAfter = function(indexNo) {
		this.saveOneScreenCellInMemory(this.curXLSRowNo, this.curXLSColNo);
		// Increase the row count
		this.updateXLSCount(this.XLSRowCount + 1, this.XLSColCount);

		// If the index is out of bounds then return
		if (indexNo + 1 < 1 || (indexNo + 1 > this.XLSRowCount))
			return;

		for ( var i = this.XLSRowCount; i > indexNo + 1; i--) {
			// Shift the rows accordingly
			this.XLSRowArray[i] = this.XLSRowArray[i - 1];
			this.XLSRowArray[i].rowNo = i;
			this.XLSRowArray[i].gap = this.defaultRowGap;
		}

		this.XLSRowArray[indexNo + 1] = new clsMemoryRow(this.defaultCellHeight);
		this.XLSRowArray[indexNo + 1].rowNo = indexNo + 1;
		this.XLSRowArray[indexNo + 1].gap = this.defaultRowGap;

		var tempList = [];
		for ( var i = 1; i <= this.XLSColCount; i++) {
			if (this.XLSCellArray[indexNo][i].spannedBy.row != 0
					&& this.XLSCellArray[indexNo][i].spannedBy.col != 0) {
				var spannedByCell = this.XLSCellArray[this.XLSCellArray[indexNo][i].spannedBy.row][this.XLSCellArray[indexNo][i].spannedBy.col];
				if (this.XLSCellArray[indexNo][i].spannedBy.row != indexNo) {
					if (indexOfElementInArray(tempList, spannedByCell) == -1) {
						tempList.push(spannedByCell)
					}
				}
			}
		}

		for ( var i = 0; i < tempList.length; i++) {
			tempList[i].rowSpan++;
		}

		for ( var i = this.XLSRowCount; i > indexNo + 1; i--) {
			// Shift the cells accordingly
			for ( var j = 1; j <= this.XLSColCount; j++) {
				this.XLSCellArray[i][j] = this.XLSCellArray[i - 1][j];
				this.XLSCellArray[i][j].parentRow = this.XLSRowArray[i];
				this.XLSCellArray[i][j].parentCol = this.XLSColArray[j];
			}
		}
		
		//Finally initialize the cells
		for ( var j = 1; j <= this.XLSColCount; j++) {
			this.initCellInXLSMemory(indexNo + 1, j);
		}

		var affectedRange = indexNo + ":1:" + this.XLSRowCount + ":"
				+ this.XLSColCount;
		this.adjustNamedRange(affectedRange, 1, 0);

		// and update the screen
		this.refreshSpan();
		this.updateScreenFromXLS();
		this.refreshGridCells();
		this.refreshHeaders();
	};

	// This function inserts multiple columns at the specified index
	clsDionGrid.prototype.insertMultipleRow = function(indexNo, count) {
		this.saveOneScreenCellInMemory(this.curXLSRowNo, this.curXLSColNo);
		// If the index is out of bounds then return
		if (count < 1 || indexNo < 1 || (indexNo > this.XLSRowCount + 1))
			return;
		
		var xlsRowCount = this.XLSRowCount;
		this.updateXLSCount(this.XLSRowCount + count, this.XLSColCount);

		for ( var i = xlsRowCount; i >= (indexNo + count); i--) {
			// Shift the rows accordingly
			this.XLSRowArray[i] = this.XLSRowArray[i - count];
			this.XLSRowArray[i].rowNo = i;
			this.XLSRowArray[i].gap = this.defaultRowGap;
		}
		// insert new row heights
		for ( var i = 0; i < count; i++) {
			this.XLSRowArray[indexNo + i].rowNo = indexNo + i;
			this.XLSRowArray[indexNo + i].gap = this.defaultRowGap;
		}

		for ( var i = this.XLSRowCount; i >= (indexNo + count); i--) {
			// Shift the cells accordingly
			for ( var j = 1; j <= this.XLSColCount; j++) {
				this.XLSCellArray[i][j] = this.XLSCellArray[i - count][j].copyCell();
				this.XLSCellArray[i - count][j].value = "";
				this.XLSCellArray[i - count][j].valueOptions.value = "";
				this.XLSCellArray[i - count][j].valueOptions.displayValue = "";
				this.XLSCellArray[i - count][j].valueOptions.stringValue = "";
				this.XLSCellArray[i][j].parentRow = this.XLSRowArray[i];
				this.XLSCellArray[i][j].parentCol = this.XLSColArray[j];
			};
		}
		
		var affectedRange = indexNo + ":1:" + xlsRowCount + ":"
				+ this.XLSColCount;
		this.adjustNamedRange(affectedRange, count, 0);

		// and update the screen
		this.refreshGridCells();
		this.updateScreenFromXLS();
		this.refreshHeaders();
	};

	// This function inserts a column at the specified index
	clsDionGrid.prototype.insertCol = function(indexNo) {
		this.saveOneScreenCellInMemory(this.curXLSRowNo, this.curXLSColNo);
		// Increase the column count
		this.updateXLSCount(this.XLSRowCount, this.XLSColCount + 1);

		// If the index is out of bounds then return
		if (indexNo < 1 || (indexNo > this.XLSColCount))
			return;

		for ( var j = this.XLSColCount; j > indexNo; j--) {
			this.XLSColArray[j] = this.XLSColArray[j - 1];
			this.XLSColArray[j].colNo = j;
			this.XLSColArray[j].gap = this.defaultColGap;
		}

		this.XLSColArray[indexNo] = new clsMemoryCol(this.defaultCellWidth);
		this.XLSColArray[indexNo].colNo = indexNo;
		this.XLSColArray[indexNo].gap = this.defaultColGap;

		var tempList = [];
		for ( var i = 1; i <= this.XLSRowCount; i++) {
			if (this.XLSCellArray[i][indexNo].spannedBy.row != 0
					&& this.XLSCellArray[i][indexNo].spannedBy.col != 0) {
				var spannedByCell = this.XLSCellArray[this.XLSCellArray[i][indexNo].spannedBy.row][this.XLSCellArray[i][indexNo].spannedBy.col];
				if (this.XLSCellArray[indexNo][i].spannedBy.col != indexNo) {
					if (indexOfElementInArray(tempList, spannedByCell) == -1) {
						tempList.push(spannedByCell)
					}
				}
			}
		}
		for ( var i = 0; i < tempList.length; i++) {
			tempList[i].colSpan++;
		}

		for ( var j = this.XLSColCount; j > indexNo; j--) {
			// Shift the cells accordingly
			for ( var i = 1; i <= this.XLSRowCount; i++) {
				this.XLSCellArray[i][j] = this.XLSCellArray[i][j - 1];
				this.XLSCellArray[i][j].parentRow = this.XLSRowArray[i];
				this.XLSCellArray[i][j].parentCol = this.XLSColArray[j];
			}
			// Shift the columns accordingly
		}

		// Finally initialize the cells
		for ( var i = 1; i <= this.XLSRowCount; i++) {
			this.initCellInXLSMemory(i, indexNo);
		}

		var affectedRange = "1:" + indexNo + ":" + this.XLSRowCount + ":"
				+ this.XLSColCount;
		//
		
		this.adjustNamedRange(affectedRange, 0, 1);

		// and update the screen
		this.refreshSpan();
		this.refreshGridCells();
		this.updateScreenFromXLS();
		this.refreshHeaders();
	};

	// This function inserts a column at the specified index
	clsDionGrid.prototype.insertColAfter = function(indexNo) {
		this.saveOneScreenCellInMemory(this.curXLSRowNo, this.curXLSColNo);
		// Increase the column count
		this.updateXLSCount(this.XLSRowCount, this.XLSColCount + 1);

		// If the index is out of bounds then return
		if (indexNo + 1 < 1 || (indexNo + 1 > this.XLSColCount))
			return;

		for ( var j = this.XLSColCount; j > indexNo + 1; j--) {
			this.XLSColArray[j] = this.XLSColArray[j - 1];
			this.XLSColArray[j].colNo = j;
			this.XLSColArray[j].gap = this.defaultColGap;
		}

		this.XLSColArray[indexNo + 1] = new clsMemoryCol(this.defaultCellWidth);
		this.XLSColArray[indexNo + 1].colNo = indexNo + 1;
		this.XLSColArray[indexNo + 1].gap = this.defaultColGap;

		var tempList = [];
		for ( var i = 1; i <= this.XLSRowCount; i++) {
			if (this.XLSCellArray[i][indexNo].spannedBy.row != 0
					&& this.XLSCellArray[i][indexNo].spannedBy.col != 0) {
				var spannedByCell = this.XLSCellArray[this.XLSCellArray[i][indexNo].spannedBy.row][this.XLSCellArray[i][indexNo].spannedBy.col];
				if (this.XLSCellArray[i][indexNo].spannedBy.col != indexNo) {
					if (indexOfElementInArray(tempList, spannedByCell) == -1) {
						tempList.push(spannedByCell)
					}
				}
			}
		}
		for ( var i = 0; i < tempList.length; i++) {
			tempList[i].colSpan++;
		}

		for ( var j = this.XLSColCount; j > indexNo + 1; j--) {
			// Shift the cells accordingly
			for ( var i = 1; i <= this.XLSRowCount; i++) {
				this.XLSCellArray[i][j] = this.XLSCellArray[i][j - 1];
				this.XLSCellArray[i][j].parentRow = this.XLSRowArray[i];
				this.XLSCellArray[i][j].parentCol = this.XLSColArray[j];
			}
			// Shift the columns accordingly
		}

		// Finally initialize the cells
		for ( var i = 1; i <= this.XLSRowCount; i++) {
			this.initCellInXLSMemory(i, indexNo + 1);
		}

		var affectedRange = "1:" + indexNo + ":" + this.XLSRowCount + ":"
				+ this.XLSColCount;
		this.adjustNamedRange(affectedRange, 0, 1);

		// and update the screen
		this.refreshSpan();
		this.updateScreenFromXLS();
		this.refreshGridCells();
		this.refreshHeaders();
		// this.afterInsertOrDelete(indexNo,indexNo, "col", "insert" );
	};

	// This function inserts multiple col at the specified index
	clsDionGrid.prototype.insertMultipleCol = function(indexNo, count) {
		this.saveOneScreenCellInMemory(this.curXLSRowNo, this.curXLSColNo);
		// If the index is out of bounds then return
		if (count < 1 || indexNo < 1 || (indexNo > this.XLSColCount + 1))
			return;

		for ( var j = this.XLSColCount; j >= (indexNo + count); j--) {
			// Shift the rows accordingly
			this.XLSColArray[j] = this.XLSColArray[j - count];
			this.XLSColArray[j].colNo = j;
			this.XLSColArray[j].gap = this.defaultColGap;
		}

		// insert new row heights
		for ( var i = 0; i < count; i++) {
			this.XLSColArray[indexNo + i] = new clsMemoryCol(
					this.defaultCellWidth);
			this.XLSColArray[indexNo + i].colNo = indexNo + i;
			this.XLSColArray[indexNo + i].gap = this.defaultColGap;
		}

		for ( var j = this.XLSColCount; j >= (indexNo + count); j--) {
			// Shift the cells accordingly
			for ( var i = 1; i <= this.XLSRowCount; i++) {
				this.XLSCellArray[i][j] = this.XLSCellArray[i][j - count];
				this.XLSCellArray[i][j].parentRow = this.XLSRowArray[i];
				this.XLSCellArray[i][j].parentCol = this.XLSColArray[j];
			}
		}

		// Finally initialize the cells
		for ( var i = 0; i < count; i++) {
			for ( var j = 1; j <= this.XLSRowCount; j++) {
				this.initCellInXLSMemory(j, indexNo + i);
			}
		}
		this.updateXLSCount(this.XLSRowCount, this.XLSColCount + count);
		var affectedRange = "1:" + indexNo + ":" + this.XLSRowCount + ":"
				+ this.XLSColCount;
		this.adjustNamedRange(affectedRange, 0, count);

		// and update the screen
		this.updateScreenFromXLS();
		this.refreshGridCells();
		this.refreshHeaders();
		// this.afterInsertOrDelete(indexNo,indexNo + count -1, "col", "insert"
		// );
	};

	// This function deletes the column given by the index number
	clsDionGrid.prototype.deleteCol = function(indexNo) {
		if (this.XLSColCount == 1)
			return;

		this.saveOneScreenCellInMemory(this.curXLSRowNo, this.curXLSColNo);
		// Return if indexNo is out of bounds
		if (indexNo < 1 || (indexNo > this.XLSColCount))
			return;

		for ( var j = indexNo; j < this.XLSColCount; j++) {
			// Delete the Column from XLS Col Array
			this.XLSColArray[j] = this.XLSColArray[j + 1];
			this.XLSColArray[j].colNo = j;
		}

		var tempList = [];
		for ( var i = 1; i <= this.XLSRowCount; i++) {
			var XLSCellArrayObj = this.XLSCellArray[i][indexNo];
			if (XLSCellArrayObj.spannedBy.row != 0
					&& XLSCellArrayObj.spannedBy.col != 0) {
				var spannedByCell = this.XLSCellArray[XLSCellArrayObj.spannedBy.row][XLSCellArrayObj.spannedBy.col];
				if (indexOfElementInArray(tempList, spannedByCell) == -1) {
					tempList.push(spannedByCell)
				}
			}
		}

		for ( var i = 0; i < tempList.length; i++) {
			tempList[i].colSpan--;
		}

		// Else shift the cells from XLS Cell Array
		for ( var j = indexNo; j < this.XLSColCount; j++) {
			for ( var i = 1; i <= this.XLSRowCount; i++) {
				this.XLSCellArray[i][j] = this.XLSCellArray[i][j + 1];
				this.XLSCellArray[i][j].parentRow = this.XLSRowArray[i];
				this.XLSCellArray[i][j].parentCol = this.XLSColArray[j];
			}
		}

		// Set the lasts cell to be null
		for ( var i = 1; i <= this.XLSRowCount; i++) {
			this.XLSCellArray[i][this.XLSColCount] = null;
			this.XLSCellArray[i].splice(this.XLSColCount, 1);
			this.displayCellArray[i].splice(this.XLSColCount, 1);
		}
		
		var affectedRange = "1:" + indexNo + ":" + this.XLSRowCount + ":"
		+ this.XLSColCount;
		this.adjustNamedRange(affectedRange, 0, -1);

		// If the current cell is last then move the current cell to left
		if (this.curXLSColNo == this.XLSColCount && this.XLSColCount != 1)
			this.curXLSColNo--;
		// Finally decrease the column count and update the screen
		this.XLSColCount = Math.max(1, this.XLSColCount - 1);

		if (!this.XLSCellArray[1] || !this.XLSCellArray[1][1]) {
			this.initCellInMemory(1, 1);
		}

		
		this.refreshSpan();
		this.updateScreenFromXLS();
		this.refreshGridCells();
		this.refreshHeaders();
		// this.afterInsertOrDelete(indexNo,indexNo, "col", "delete" );
	};

	// This function removes the row given by the index number
	clsDionGrid.prototype.deleteRow = function(indexNo) {
		if (this.XLSRowCount == 1)
			return;

		this.saveOneScreenCellInMemory(this.curXLSRowNo, this.curXLSColNo);
		// Return if indexNo is out of bounds
		if (indexNo < 1 || (indexNo > this.XLSRowCount))
			return;

		for ( var i = indexNo; i < this.XLSRowCount; i++) {
			// Delete the Row from XLS Row Array
			this.XLSRowArray[i] = this.XLSRowArray[i + 1];
			this.XLSRowArray[i].rowNo = i;
		}

		var tempList = [];
		for ( var i = 1; i <= this.XLSColCount; i++) {
			if (this.XLSCellArray[indexNo][i].spannedBy.row != 0
					&& this.XLSCellArray[indexNo][i].spannedBy.col != 0) {
				var spannedByCell = this.XLSCellArray[this.XLSCellArray[indexNo][i].spannedBy.row][this.XLSCellArray[indexNo][i].spannedBy.col];
				if (indexOfElementInArray(tempList, spannedByCell) == -1) {
					tempList.push(spannedByCell)
				}
			}
		}

		for ( var i = 0; i < tempList.length; i++) {
			tempList[i].rowSpan--;
		}

		// Else shift the cells from XLS Cell Array
		for ( var i = indexNo; i < this.XLSRowCount; i++) {
			for ( var j = 1; j <= this.XLSColCount; j++) {
				this.XLSCellArray[i][j] = this.XLSCellArray[i + 1][j];
				this.XLSCellArray[i][j].parentRow = this.XLSRowArray[i];
				this.XLSCellArray[i][j].parentCol = this.XLSColArray[j];
			}
		}

		this.displayCellArray.splice(this.XLSRowCount, 1);
		this.XLSCellArray.splice(this.XLSRowCount, 1);
		this.XLSRowArray.splice(this.XLSRowCount, 1);

		// If the current cell is last then move the current cell to left
		if (this.curXLSRowNo == this.XLSRowCount && this.XLSRowCount != 1)
			this.curXLSRowNo--;
		// Finally decrease the row count
		this.XLSRowCount = Math.max(1, this.XLSRowCount - 1);

		if (!this.XLSCellArray[1] || !this.XLSCellArray[1][1]) {
			this.initCellInMemory(1, 1);
		}

		var affectedRange = indexNo + ":1:" + this.XLSRowCount + ":"
				+ this.XLSColCount;
		this.adjustNamedRange(affectedRange, -1, 0);
		// and update the screen
		this.refreshSpan();
		this.updateScreenFromXLS();
		this.refreshGridCells();
		this.refreshHeaders();
		// this.afterInsertOrDelete(indexNo,indexNo, "row", "delete" );
	};

	// This function removes the row given by the index number
	clsDionGrid.prototype.deleteMultipleRow = function(indexNo, count) {
		if (this.XLSRowCount < count)
			count = this.XLSRowCount - 1;

		this.saveOneScreenCellInMemory(this.curXLSRowNo, this.curXLSColNo);
		// Return if indexNo is out of bounds
		if (indexNo < 1 || (indexNo > this.XLSRowCount) || count < 1)
			return;
		// adjust count if no more rows are left to delete
		if (indexNo + count > this.XLSRowCount)
			count = this.XLSRowCount - indexNo + 1;

		for ( var i = indexNo; (i + count) <= this.XLSRowCount; i++) {
			// Delete the Row from XLS Row Array
			this.XLSRowArray[i] = this.XLSRowArray[i + count];
			this.XLSRowArray[i].rowNo = i;
		}

		// Else shift the cells from XLS Cell Array
		for ( var i = indexNo; (i + count) <= this.XLSRowCount; i++) {
			for ( var j = 1; j <= this.XLSColCount; j++) {
				this.XLSCellArray[i][j] = this.XLSCellArray[i + count][j];
				this.XLSCellArray[i][j].parentRow = this.XLSRowArray[i];
				this.XLSCellArray[i][j].parentCol = this.XLSColArray[j];
			}
		}

		// Set the last cells to be null
		for ( var j = 0; j < count; j++) {
			this.displayCellArray.splice(this.XLSRowCount - j);
			this.XLSCellArray.splice(this.XLSRowCount - j);
		}
		// Decrease the row count
		this.XLSRowCount -= count;
		this.XLSRowCount -= count;
		this.XLSRowCount = Math.max(1, this.XLSRowCount - count);

		// If the current cell is last then move the current cell to left
		if (this.curXLSRowNo > this.XLSRowCount)
			this.curXLSRowNo = this.XLSRowCount;
		if (!this.XLSCellArray[1] || !this.XLSCellArray[1][1]) {
			this.initCellInMemory(1, 1);
		}
		// Then update the screen
		this.updateScreenFromXLS();
		this.refreshGridCells();
		this.refreshHeaders();
	};

	// This function removes the row given by the index number
	clsDionGrid.prototype.deleteMultipleCol = function(indexNo, count) {
		if (this.XLSColCount < count)
			count = this.XLSColCount - 1;

		this.saveOneScreenCellInMemory(this.curXLSRowNo, this.curXLSColNo);
		// Return if indexNo is out of bounds
		if (indexNo < 1 || (indexNo > this.XLSColCount) || count < 1)
			return;
		// adjust count if no more rows are left to delete
		if (indexNo + count > this.XLSColCount)
			count = this.XLSColCount - indexNo + 1;

		for ( var i = indexNo; (i + count) <= this.XLSColCount; i++) {
			// Delete the Row from XLS Row Array
			this.XLSColArray[i] = this.XLSColArray[i + count];
			this.XLSColArray[i].colNo = i;
		}

		// Else shift the cells from XLS Cell Array
		for ( var i = indexNo; (i + count) <= this.XLSColCount; i++) {
			for ( var j = 1; j <= this.XLSRowCount; j++) {
				this.XLSCellArray[j][i] = this.XLSCellArray[j][i + count];
				this.XLSCellArray[j][i].parentRow = this.XLSRowArray[j];
				this.XLSCellArray[j][i].parentCol = this.XLSColArray[i];
			}
		}

		for ( var i = 1; i <= this.XLSRowCount; i++) {
			this.XLSCellArray[i][this.XLSColCount - j] = null;
			this.displayCellArray[i].splice(this.XLSColCount - j);
			this.XLSCellArray[i].splice(this.XLSColCount - j);
		}

		this.XLSColCount = Math.max(1, this.XLSColCount - count);

		// If the current cell is last then move the current cell to left
		if (this.curXLSColNo > this.XLSColCount)
			this.curXLSColNo = this.XLSColCount;
		if (!this.XLSCellArray[1] || !this.XLSCellArray[1][1]) {
			this.initCellInMemory(1, 1);
		}
		// Then update the screen
		this.updateScreenFromXLS();
		this.refreshGridCells();
		this.refreshHeaders();
		// this.afterInsertOrDelete(indexNo,indexNo + count-1, "col", "delete"
		// );
	};

	// This function copies all the selected cells into copiedCellArray
	clsDionGrid.prototype.copyRange = function(range) {
		this.saveOneScreenCellInMemory(this.curXLSRowNo, this.curXLSColNo);
		// this.borderMapper(this.copiedRange,"noneDashed");

		copiedRange = "";
		workbook.copiedCellArray = [];
		// Get the required indices from the selected range
		var rangeArray = range.split(":");
		// If there is some problem with the selected range then break
		if (rangeArray.length != 4)
			return false;
		else {
			var startX = parseInt(rangeArray[0]);
			var startY = parseInt(rangeArray[1]);
			var endX = parseInt(rangeArray[2]);
			var endY = parseInt(rangeArray[3]);

			if (startX > endX) {
				var temp = endX;
				endX = startX;
				startX = temp;
			}
			if (startY > endY) {
				var temp = endY;
				endY = startY;
				startY = temp;
			}
			// Initialize the copied cell array
			for ( var i = 0; i < Math.abs(startX - endX) + 1; i++) {
				// There is a need to check the limits of i and j as someone
				// might first select the range and then delete those cells
				// (Will not happen with the current code....but better safe
				// than sorry)
				if (i + startX > this.XLSRowCount) {
					break;
				}
				workbook.copiedCellArray[i] = [];
				for ( var j = 0; j < Math.abs(startY - endY) + 1; j++) {
					if (j + startY > this.XLSColCount) {
						continue;
					}
					// Deep Copy

					workbook.copiedCellArray[i][j] = this.XLSCellArray[i
							+ startX][j + startY].copyCell();
				}
			}
		}
		copiedRange = range;
		workbook.setCopiedRange(this.gridName, copiedRange);
		// this.borderMapper(this.copiedRange,"allDashed");
	};

	// This function pastes all the cells in copiedCellArray to the cells
	// starting from rowNo and colNo
	clsDionGrid.prototype.pasteCopiedCells = function(rowNo, colNo) {
		var rowCount = workbook.copiedCellArray.length;
		var colCount = 0;
		if (workbook.copiedCellArray.length > 0) {
			var colCount = workbook.copiedCellArray[0].length;
		}

		// To check whether copy function has been called and there are cells
		// which have been copied
		var copiedCellArray = workbook.copiedCellArray;

		if (copiedCellArray.length > 0) {
			for ( var i = 0; i < rowCount; i++) {
				if (rowNo + i > this.XLSRowCount) {
					// If there are no more rows where it can be pasted then
					// break
					break;
				}
				for ( var j = 0; j < colCount; j++) {
					if (colNo + j > this.XLSColCount) {
						// If there are no more columns where it can be pasted
						// then continue to the next row
						continue;
					}
					// copy the cells (Deep copy)
					this.XLSCellArray[rowNo + i][colNo + j]
							.updateCellByString(copiedCellArray[i][j].value);
					this.XLSCellArray[rowNo + i][colNo + j].displayValue = copiedCellArray[i][j].displayValue;
					this.XLSCellArray[rowNo + i][colNo + j].displayType = copiedCellArray[i][j].displayType;
					this.XLSCellArray[rowNo + i][colNo + j].valueOptions.displayValue = copiedCellArray[i][j].valueOptions.displayValue;
					this.XLSCellArray[rowNo + i][colNo + j].valueOptions.value = copiedCellArray[i][j].valueOptions.value;
					this.XLSCellArray[rowNo + i][colNo + j].valueOptions.stringValue = copiedCellArray[i][j].valueOptions.stringValue;
					
					this.onCellChanged.raiseCallBacks(null, {
						row : (rowNo + i),
						cell : (colNo + j), grid: this
					});

				}
			}
			// this.borderMapper(this.copiedRange,"noneDashed");
		}
		this.updateScreenFromXLS();
	};
	
	clsDionGrid.prototype.fireOnChange = function(mRowNo, mColNo){
		this.onCellChanged.raiseCallBacks(null, {
			row : (mRowNo),
			cell : (mColNo), grid: this
		});
	}

	// This function pastes all the cells in copiedCellArray to the cells
	// starting from rowNo and colNo
	clsDionGrid.prototype.pasteSpecialCopiedCells = function(rowNo, colNo) {
		var rowCount = workbook.copiedCellArray.length;
		var colCount = 0;
		if (workbook.copiedCellArray.length > 0) {
			var colCount = workbook.copiedCellArray[0].length;
		}
		var copiedCellArray = workbook.copiedCellArray;

		// To check whether copy function has been called and there are cells
		// which have been copied
		if (workbook.copiedCellArray.length > 0) {
			for ( var i = 0; i < rowCount; i++) {
				if (rowNo + i > this.XLSRowCount) {
					// If there are no more rows where it can be pasted then
					// break
					break;
				}
				for ( var j = 0; j < colCount; j++) {
					if (colNo + j > this.XLSColCount) {
						// If there are no more columns where it can be pasted
						// then continue to the next row
						continue;
					}
					var XLSCellArrayObj = this.XLSCellArray[rowNo + i][colNo
							+ j];
					if (XLSCellArrayObj.readOnly) continue;
					this.moveCell(copiedCellArray[i][j], rowNo + i, colNo + j);
					this.onCellChanged.raiseCallBacks(null, {
						row : (rowNo + i),
						cell : (colNo + j), grid: this
					});
					if (copiedCellArray[i][j].cellType != XLSCellArrayObj.cellType) continue;
					if (copiedCellArray[i][j].cellType == "input" && XLSCellArrayObj.cellType == "input") {
						if (copiedCellArray[i][j].type != "formula") {
							XLSCellArrayObj
									.updateCellByString(copiedCellArray[i][j].stringValue);
						} else {
							XLSCellArrayObj.updateCellByString("="
									+ modifyFormula(
											copiedCellArray[i][j].stringValue
													.substr(1),
											(rowNo - getStartX(copiedRange)),
											colNo - getStartY(copiedRange)));
						}
					} else {
						if (XLSCellArrayObj.cellType == copiedCellArray[i][j].cellType){
							XLSCellArrayObj = copiedCellArray[i][j].copyCell();
							XLSCellArrayObj.updateCellByWidget();
						}
					}
				}
			}
			// this.borderMapper(this.copiedRange,"noneDashed");
		}
		this.updateScreenFromXLS();
	};

	// This function deletes the given cells from the XLSCellArray and shift
	// cells either left , up or noshift according to the argument
	clsDionGrid.prototype.deleteSelectedCells = function(range, shiftPlace) {
		// Get the indices from the range
		var rangeArray = range.split(":");
		if (rangeArray.length != 4)
			return false;

		var startX = parseInt(rangeArray[0]);
		var startY = parseInt(rangeArray[1]);
		var endX = parseInt(rangeArray[2]);
		var endY = parseInt(rangeArray[3]);

		if (startX > endX) {
			var temp = endX;
			endX = startX;
			startX = temp;
		}
		if (startY > endY) {
			var temp = endY;
			endY = startY;
			startY = temp;
		}

		var countX = endX - startX + 1;
		var countY = endY - startY + 1;
		// Left shift the cells from XLS Cell Array
		if (shiftPlace == "left") {
			for ( var i = startX; i <= endX; i++) {
				for ( var j = startY; j <= this.XLSColCount - countY; j++) {
					this.XLSCellArray[i][j] = this.XLSCellArray[i][j + countY];
					this.XLSCellArray[i][j].parentRow = this.XLSRowArray[i];
					this.XLSCellArray[i][j].parentCol = this.XLSColArray[j];
				}
			}

			for ( var i = startX; i <= endX; i++) {
				for ( var j = this.XLSColCount - countY + 1; j <= this.XLSColCount; j++) {
					this.XLSCellArray[i][j] = new clsMemoryCell();
					this.XLSCellArray[i][j].parentXLSArray = this.XLSCellArray;
					this.XLSCellArray[i][j].parentRow = this.XLSRowArray[i];
					this.XLSCellArray[i][j].parentCol = this.XLSColArray[j];

					var cellStyle = (i % 2) ? this.defaultEvenEditableCellStyle
							: this.defaultOddEditableCellStyle;

					this.XLSCellArray[i][j].className = cellStyle
							+ " " + this.defaultTextAlign + "AlignCellClass ";
					this.XLSCellArray[i][j].grandParentDionGrid = this;
					this.XLSCellArray[i][j].updateCellByString("");

				}
			}

			// Finally update the screen based on XLSCellArray
			this.updateScreenFromXLS();
		}

		// Shift the cells up from XLS Cell Array
		else if (shiftPlace == "up") {
			// Else shift the cells from XLS Cell Array
			for ( var i = startX; i <= this.XLSRowCount - countX; i++) {
				for ( var j = startY; j <= endY; j++) {
					this.XLSCellArray[i][j] = this.XLSCellArray[i + countX][j];
					this.XLSCellArray[i][j].parentRow = this.XLSRowArray[i];
					this.XLSCellArray[i][j].parentCol = this.XLSColArray[j];
				}
			}

			for ( var i = this.XLSRowCount - countX + 1; i <= this.XLSRowCount; i++) {
				for ( var j = startY; j <= endY; j++) {
					this.XLSCellArray[i][j] = new clsMemoryCell();
					this.XLSCellArray[i][j].parentXLSArray = this.XLSCellArray;
					this.XLSCellArray[i][j].parentRow = this.XLSRowArray[i];
					this.XLSCellArray[i][j].parentCol = this.XLSColArray[j];

					var cellStyle = (i % 2) ? this.defaultEvenEditableCellStyle
							: this.defaultOddEditableCellStyle;					
					this.XLSCellArray[i][j].className = cellStyle
							+ " " + this.defaultTextAlign + "AlignCellClass ";
					this.XLSCellArray[i][j].grandParentDionGrid = this;
					this.XLSCellArray[i][j].updateCellByString("");

				}
			}

			// Finally update the screen based on XLSCellArray
			this.updateScreenFromXLS();
		}

		// Do not shift he cells
		else if (shiftPlace == "noShift") {
			// Else shift the cells from XLS Cell Array
			for ( var i = startX; i <= endX; i++) {
				for ( var j = startY; j <= endY; j++) {
					if (this.XLSCellArray[i][j].readOnly == false) {
						if (this.XLSCellArray[i][j].cellType == "input") {
							this.XLSCellArray[i][j].updateCellByString("");
						} else {
							if (this.XLSCellArray[i][j].cellType == "notional"
									|| this.XLSCellArray[i][j].cellType == "date") {
								this.XLSCellArray[i][j].value = "";
								this.XLSCellArray[i][j].valueOptions.value = "";
							}
						}
						this.refreshCell(i, j);
					}
				}
			}
		}
	};

	clsDionGrid.prototype.isLabel = function(mRowNo, mColNo) {
		return (mRowNo > 0 && mRowNo <= this.XLSRowCount && mColNo > 0 && mColNo <= this.XLSColCount && this.XLSCellArray[mRowNo][mColNo].cellType == "label");
	}; 
	// Returns whether the given rowNumber or colNumber is valid row/col number
	clsDionGrid.prototype.isInXLSRange = function(mRowNo, mColNo) {
		return (mRowNo > 0 && mRowNo <= this.XLSRowCount && mColNo > 0 && mColNo <= this.XLSColCount);
	};

	clsDionGrid.prototype.isCurrentCell = function(mRowNo, mColNo) {
		return this.isInXLSRange(mRowNo, mColNo)
				&& (mRowNo == this.curXLSRowNo && mColNo == this.curXLSColNo);
	};

	// Returns whether the cell given by mRowNo and mColNo has a formula i.e.
	// starts with '=' or not.
	// If the first character is a space and second is '=' then it is not
	// treated as a formula.
	// We do not trim the string and then check for '='. This behaviour is same
	// as that of excel
	clsDionGrid.prototype.isFormula = function(mRowNo, mColNo) {
		if (this.isInXLSRange(mRowNo, mColNo)) {
			if (this.displayCellArray[mRowNo][mColNo].cellType == "input" || this.displayCellArray[mRowNo][mColNo].cellType == "notional") {
				var val = this.XLSCellArray[mRowNo][mColNo].stringValue;
				if (this.isCurrentCell(mRowNo, mColNo) && this.displayCellArray[mRowNo][mColNo].editor)
					val = this.displayCellArray[mRowNo][mColNo].editor.getValue();
				if (val.length > 0 && val.substr(0, 1) == "=") {
					return true;
				}
			}
		}
		return false;
	};

	clsDionGrid.prototype.interpolate = function(selectedRange,
			interpolatedRange) {
		var selectedStartX = getStartX(selectedRange);
		var selectedStartY = getStartY(selectedRange);
		var selectedEndX = getEndX(selectedRange);
		var selectedEndY = getEndY(selectedRange);

		var interpolatedStartX = getStartX(interpolatedRange);
		var interpolatedStartY = getStartY(interpolatedRange);
		var interpolatedEndX = getEndX(interpolatedRange);
		var interpolatedEndY = getEndY(interpolatedRange);

		var isIncreasing = true;
		var isColumn = true;

		if (interpolatedStartX == selectedStartX
				&& interpolatedStartY == interpolatedStartY
				&& (interpolatedEndX > selectedEndX || interpolatedEndY > selectedEndY)) {
			if (interpolatedEndX == selectedEndX) {
				interpolatedStartY = selectedEndY + 1;
				isColumn = false;
				isIncreasing = true;
			} else {
				interpolatedStartX = selectedEndX + 1;
				isColumn = true;
				isIncreasing = true;
			}
		} else {
			if (interpolatedStartX == selectedStartX) {
				interpolatedEndY = selectedStartY - 1;
				isColumn = false;
				isIncreasing = false;
			} else {
				interpolatedEndX = selectedStartX - 1;
				isColumn = true;
				isIncreasing = false;
			}
		}

		var matchX;
		var matchY;
		if (isIncreasing) {
			for ( var i = interpolatedStartX; i <= interpolatedEndX; i++) {
				for ( var j = interpolatedStartY; j <= interpolatedEndY; j++) {
					matchX = selectedStartX + (i - interpolatedStartX)
							% (selectedEndX - selectedStartX + 1);
					matchY = selectedStartY + (j - interpolatedStartY)
							% (selectedEndY - selectedStartY + 1);
					var XLSCellArrayObj = this.XLSCellArray[i][j];
					var matchCellArrayObj = this.XLSCellArray[matchX][matchY];
					this.moveCell(matchCellArrayObj, i, j);
					if (matchCellArrayObj.type == "formula") {
						XLSCellArrayObj.updateCellByString(modifyFormula("="
								+ matchCellArrayObj.stringValue.substr(1),
								(i - matchX), j - matchY));
					} else if (matchCellArrayObj.type == "integer"
							|| matchCellArrayObj.type == "float") {
						var newValue;
						var count = 0;
						var sum = 0;
						var lastVal = 0;

						if (isColumn) {
							for ( var k = selectedStartX; k <= selectedEndX; k++) {
								var tempObj = this.XLSCellArray[k][j];
								if (tempObj.type == "integer"
										|| tempObj.type == "float") {
									if (count > 0)
										sum += (tempObj.value - lastVal);
									count++;
									lastVal = tempObj.value;
								}
							}

							for ( var k = i - 1; k >= selectedStartX; k--) {
								var tempObj = this.XLSCellArray[k][j];
								if (tempObj.type == "integer"
										|| tempObj.type == "float") {
									lastVal = tempObj.value;
									break;
								}
							}
						} else {
							for ( var k = selectedStartY; k <= selectedEndY; k++) {
								var tempObj = this.XLSCellArray[i][k];
								if (tempObj.type == "integer"
										|| tempObj.type == "float") {
									if (count > 0)
										sum += (tempObj.value - lastVal);
									count++;
									lastVal = tempObj.value;
								}
							}

							for ( var k = j - 1; k >= selectedStartY; k--) {
								var tempObj = this.XLSCellArray[i][k];
								if (tempObj.type == "integer"
										|| tempObj.type == "float") {
									lastVal = tempObj.value;
									break;
								}
							}
						}
						if (count > 1)
							newValue = lastVal + sum / (count - 1);
						else
							newValue = lastVal;
						XLSCellArrayObj.updateCellByString("" + newValue);
					}
				}
			}
		}

		else {
			for ( var i = interpolatedEndX; i >= interpolatedStartX; i--) {
				for ( var j = interpolatedEndY; j >= interpolatedStartY; j--) {
					matchX = selectedEndX + (i - interpolatedEndX)
							% (selectedEndX - selectedStartX + 1);
					matchY = selectedEndY + (j - interpolatedEndY)
							% (selectedEndY - selectedStartY + 1);
					var XLSCellArrayObj = this.XLSCellArray[i][j];
					var matchCellArrayObj = this.XLSCellArray[matchX][matchY];
					this.moveCell(matchCellArrayObj, i, j);
					if (matchCellArrayObj.type == "formula") {
						XLSCellArrayObj.updateCellByString(modifyFormula("="
								+ matchCellArrayObj.stringValue.substr(1),
								(i - matchX), j - matchY));
					} else if (matchCellArrayObj.type == "integer"
							|| matchCellArrayObj.type == "float") {
						var newValue;
						var count = 0;
						var sum = 0;
						var lastVal = 0;

						if (isColumn) {
							for ( var k = selectedStartX; k <= selectedEndX; k++) {
								var tempObj = this.XLSCellArray[k][j];
								if (tempObj.type == "integer"
										|| tempObj.type == "float") {
									if (count > 0)
										sum += (tempObj.value - lastVal);
									count++;
									lastVal = tempObj.value;
								}
							}

							for ( var k = i + 1; k <= selectedEndX; k++) {
								var tempObj = this.XLSCellArray[k][j];
								if (tempObj.type == "integer"
										|| tempObj.type == "float") {
									lastVal = tempObj.value;
									break;
								}
							}
						} else {
							for ( var k = selectedStartY; k <= selectedEndY; k++) {
								var tempObj = this.XLSCellArray[i][k];
								if (tempObj.type == "integer"
										|| tempObj.type == "float") {
									if (count > 0)
										sum += (tempObj.value - lastVal);
									count++;
									lastVal = tempObj.value;
								}
							}

							for ( var k = j + 1; k <= selectedEndY; k++) {
								var tempObj = this.XLSCellArray[i][k];
								if (tempObj.type == "integer"
										|| tempObj.type == "float") {
									lastVal = tempObj.value;
									break;
								}
							}
						}
						if (count > 1)
							newValue = lastVal - sum / (count - 1);
						else
							newValue = lastVal;
						XLSCellArrayObj.updateCellByString("" + newValue);
					}
				}
			}
		}
	};

	// Creates a Container for the editors
	clsDionGrid.prototype.createCellDiv = function(i, j) {
		var THIS = this;
		var cellDiv = document.createElement("div");
		var cellStyle = (i % 2) ? this.defaultEvenEditableCellStyle
				: this.defaultOddEditableCellStyle;

		cellDiv.className = cellStyle;
		cellDiv.id = this.gridName + ":" + "CDiv_" + i + "_" + j;

		this.gridDiv.appendChild(cellDiv);

		// Do not remove this tabIndex otherwise cellDiv.onblur event will not
		// be called....
		cellDiv.tabIndex = 0;

		cellDiv.onmousedown = function(e) {
			THIS.hideContextMenu();
			THIS.clearInterpolation();

			if (e == null)
				e = window.event;
			// To check whether its a left click
			if ((e.button == 1 && window.event != null || e.button == 0)) {
				// get the row and col number
				var mStr = this.id.split(":")[1];
				document.onmouseup = OnMouseUp;
				mStr = mStr.substr(5);
				var mPos = mStr.indexOf("_");
				var mRowNo = parseInt(mStr.substr(0, mPos));
				var mColNo = parseInt(mStr.substr(mPos + 1));

				var value = "";
				var GRID = workbook.curDionGrid;
				if (!GRID) {
					workbook.curDionGrid = THIS;
					GRID = THIS;
				}
				if (GRID.displayCellArray[GRID.curXLSRowNo][GRID.curXLSColNo].editor)
					value = GRID.displayCellArray[GRID.curXLSRowNo][GRID.curXLSColNo].editor
							.getValue();

				// In case of shift key just modify the selected range and call
				// the apply range function
				if (e.shiftKey) {
					var range = THIS.curXLSRowNo + ":" + THIS.curXLSColNo + ":"
							+ mRowNo + ":" + mColNo;
					if (workbook.selectedRange != range) {
						workbook.clearRange(workbook.selectedRange);
						workbook.selectedRange = range;
						workbook.applyRange();
					}
					preventDefaultBehaviour(e);
				}

				else {
					var range = mRowNo + ":" + mColNo + ":" + mRowNo + ":"
							+ mColNo;
					var isCurGrid = workbook.isCurGrid(THIS);

					if (THIS.isLabel(mRowNo, mColNo)){
						return;
					}
					
					
					// If the user clicks on the current cell in edit mode then
					// do not do anything else go to the cell
					if (THIS.mode == "edit") {
						// IN case of a formula on clicking some other cell we
						// do
						// not do anything as the user might be selecting some
						// range
						if (value.length > 0 && value.substr(0, 1) == "=") {
							preventDefaultBehaviour(e);
						} else if (!(mRowNo == THIS.curXLSRowNo
								&& mColNo == THIS.curXLSColNo && isCurGrid)) {
							workbook.clearRange(workbook.selectedRange);
							workbook.selectedRange = range;
							THIS.goToCell(mRowNo, mColNo);
							preventDefaultBehaviour(e);
						} else {
							// preventDefaultBehaviour(e);
						}
					} else {
						if (!(mRowNo == THIS.curXLSRowNo
								&& mColNo == THIS.curXLSColNo && isCurGrid)) {
							workbook.clearRange(workbook.selectedRange);
							workbook.selectedRange = range;
							THIS.goToCell(mRowNo, mColNo);
							preventDefaultBehaviour(e);
						}
					}
				}
				if (!THIS.isLabel(mRowNo,mColNo))
					THIS.dragItem = "cell";
			}
		};

		cellDiv.onmouseup = function(e) {
			if (THIS.onCellClicked)
				THIS.onCellClicked(THIS.curXLSRowNo, THIS.curXLSColNo);
			if (e == null)
				e = window.event;
			if (THIS.dragItem == "cell") {
				if ((e.button == 1 && window.event != null || e.button == 0)) {
					var mStr = this.id.split(":")[1];

					mStr = mStr.substr(5);
					var mPos = mStr.indexOf("_");
					var mRowNo = parseInt(mStr.substr(0, mPos));
					var mColNo = parseInt(mStr.substr(mPos + 1));
					_endCellX = mRowNo;
					_endCellY = mColNo;

					// If the current cell is a formula in edit mode then append
					// the range to it and clear the selected Range
					// IT SHOULD APPEND THE RANGE ONLY IF IT IS A FORMULA AND
					// THE CELL ON WHICH MOUSE UP IS RAISED IS NOT THE CURRNT
					// CELL
					// There is no need to check the mode as if the mode was not
					// edit, then on mouse down it would go to that cell and
					// mode will be select again
					var GRID = workbook.curDionGrid;
					if (GRID.mode == "edit"
							&& GRID.isFormula(workbook.curXLSRowNo,
									workbook.curXLSColNo)
							&& (workbook.curXLSRowNo != _endCellX || workbook.curXLSColNo != _endCellY)) {
						var displayCellDiv = GRID.displayCellArray[GRID.curXLSRowNo][GRID.curXLSColNo]
						var cellObj = (displayCellDiv.cellType == "notional") ? displayCellDiv.childNodes[0].childNodes[0] : displayCellDiv.childNodes[0];
						var cursorIndex = cellObj.selectionEnd;
						var rangeString = "";
						if (!workbook.isCurGrid(THIS))
							rangeString = THIS.gridName + "." + rangeString;
						rangeString += THIS.XLSCellArray[mRowNo][mColNo].id;

						if (cursorIndex >= 0) {
							cellObj.value = cellObj.value
									.substr(0, cursorIndex)
									+ rangeString
									+ cellObj.value.substr(cursorIndex);
						} else
							cellObj.value += rangeString;

						// Reupdate the selected range to the original value
						workbook.clearRange(workbook.selectedRange);
						workbook.selectedRange = GRID.curXLSRowNo + ":"
								+ GRID.curXLSColNo + ":" + GRID.curXLSRowNo
								+ ":" + GRID.curXLSColNo;
						GRID.applyRange();

						// Update the formula bar Input
						if (this.showFormatDiv && this.formulaBarInput != null)
							GRID.formulaBarInput.value = cellObj.value;
					} else {
						if (GRID.isInXLSRange(_endCellX, _endCellY)
								&& !GRID.isLabel(_endCellX, _endCellY)) {
							var rangeArray = workbook.selectedRange.split(":");
							var range = rangeArray[0] + ":" + rangeArray[1]
									+ ":" + _endCellX + ":" + _endCellY;
							if (workbook.selectedRange != range) {
								workbook.clearRange(workbook.selectedRange);
								workbook.selectedRange = range;
								THIS.applyRange();
							}
						}
					}
				}
			}

			// if the mouse drag item is interpolation div, then fill the series
			// by calling the interpolate function and update from XLS
			else if (THIS.dragItem = "interpolationDiv") {
				if (THIS.interpolatedRange != "") {
					THIS.borderMapper(THIS.interpolatedRange, "noneDashed");
					THIS
							.interpolate(workbook.selectedRange,
									THIS.interpolatedRange);
					workbook.clearRange(workbook.selectedRange);
					workbook.selectedRange = THIS.interpolatedRange;
					THIS.interpolatedRange = "";
					THIS.applyRange();
					THIS.updateScreenFromXLS();
				}
			}
			// Clear the drag item
			THIS.dragItem = "";
			document.onmouseup = null;
			document.onmousemove = null;

			// preventDefaultBehaviour(e);
		};

		cellDiv.onmousemove = function(e) {
			if (e == null)
				e = window.event;
			var rangeArray = workbook.selectedRange.split(":");
			var mStr = this.id.split(":")[1];
			mStr = mStr.substr(5);
			var mPos = mStr.indexOf("_");
			var mRowNo = parseInt(mStr.substr(0, mPos));
			var mColNo = parseInt(mStr.substr(mPos + 1));
			document.mouseup = OnMouseUp;

			if (THIS.dragItem == "cell") {
				_endCellX = mRowNo;
				_endCellY = mColNo;
				var range = rangeArray[0] + ":" + rangeArray[1] + ":"
						+ _endCellX + ":" + _endCellY;

				if (workbook.selectedRange != range) {
					workbook.clearRange(workbook.selectedRange);
					workbook.selectedRange = range;
					THIS.applyRange();
				}

				preventDefaultBehaviour(e);
			}

			else if (THIS.dragItem == "row") {
				_endCellX = mRowNo;
				_endCellY = THIS.XLSRowCount;
				var range = rangeArray[0] + ":" + rangeArray[1] + ":"
						+ _endCellX + ":" + _endCellY;
				if (workbook.selectedRange != range) {
					workbook.clearRange(workbook.selectedRange);
					workbook.selectedRange = range;
					THIS.applyRange();
				}
				preventDefaultBehaviour(e);
			}

			else if (THIS.dragItem == "col") {
				_endCellX = THIS.XLSColCount;
				_endCellY = mColNo;
				var range = rangeArray[0] + ":" + rangeArray[1] + ":"
						+ _endCellX + ":" + _endCellY;
				if (workbook.selectedRange != range) {
					workbook.clearRange(workbook.selectedRange);
					workbook.selectedRange = range;
					THIS.applyRange();
				}
				preventDefaultBehaviour(e);
			}

			else if (THIS.dragItem == "interpolationDiv") {
				var startXRange = getStartX(workbook.selectedRange);
				var startYRange = getStartY(workbook.selectedRange);
				var endXRange = getEndX(workbook.selectedRange);
				var endYRange = getEndY(workbook.selectedRange);
				_endCellX = mRowNo;
				_endCellY = mColNo;

				if (Math.abs(endYRange - _endCellY) > Math.abs(endXRange
						- _endCellX)) {
					if (_endCellY > endYRange) {
						if (THIS.interpolatedRange != "")
							THIS.borderMapper(workbook.interpolatedRange,
									"noneDashed");
						THIS.interpolatedRange = startXRange + ":"
								+ startYRange + ":" + endXRange + ":"
								+ _endCellY;
						THIS.borderMapper(THIS.interpolatedRange, "allDashed");
					} else if (_endCellY < startYRange) {
						if (THIS.interpolatedRange != "")
							THIS.borderMapper(THIS.interpolatedRange,
									"noneDashed");
						THIS.interpolatedRange = startXRange + ":" + _endCellY
								+ ":" + endXRange + ":" + endYRange;
						THIS.borderMapper(THIS.interpolatedRange, "allDashed");
					} else {
						THIS.borderMapper(THIS.interpolatedRange, "noneDashed");
						THIS.interpolatedRange = "";
					}
				} else {
					if (_endCellX > endXRange) {
						if (THIS.interpolatedRange != "")
							THIS.borderMapper(THIS.interpolatedRange,
									"noneDashed");
						THIS.interpolatedRange = startXRange + ":"
								+ startYRange + ":" + _endCellX + ":"
								+ endYRange;
						THIS.borderMapper(THIS.interpolatedRange, "allDashed");
					} else if (_endCellX < startXRange) {
						if (THIS.interpolatedRange != "")
							THIS.borderMapper(THIS.interpolatedRange,
									"noneDashed");
						THIS.interpolatedRange = _endCellX + ":" + startYRange
								+ ":" + endXRange + ":" + endYRange;
						THIS.borderMapper(THIS.interpolatedRange, "allDashed");
					} else {
						THIS.borderMapper(THIS.interpolatedRange, "noneDashed");
						THIS.interpolatedRange = "";
					}
				}

				preventDefaultBehaviour(e);
			}

		};

		cellDiv.oncontextmenu = function(e) {
			if (THIS.showContextMenu == true) {
				var mStr = this.id.split(":")[1];
				if (e == null)
					e = window.event;
				mStr = mStr.substr(5);
				var mPos = mStr.indexOf("_");
				var displayRowNo = parseInt(mStr.substr(0, mPos));
				var displayColNo = parseInt(mStr.substr(mPos + 1));
				var mRowNo = parseInt(mStr.substr(0, mPos));
				var mColNo = parseInt(mStr.substr(mPos + 1));

				THIS.hideContextMenu();
				// Create an new context menu
				var contextMenuDiv = new contextMenu();
				contextMenuDiv.createContextMenu(this.gridName + ":"
						+ "cellContextMenu");
				var x = 0;
				var y = 0;
				if (e.pageX) {
					x = e.pageX;
					y = e.pageY;
				} else {
					x = e.clientX + document.documentElement.scrollLeft;
					y = e.clientY + document.documentElement.scrollTop;
				}

				contextMenuDiv.contextMenuDiv.style.left = x + "px";
				contextMenuDiv.contextMenuDiv.style.top = y + "px";

				workbook.contextMenuDiv = contextMenuDiv;

				if (THIS.isCustomContextMenu) {
					if (THIS.contextMenuFn)
						THIS
								.contextMenuFn(THIS, mRowNo, mColNo,
										contextMenuDiv);
					if (contextMenuDiv.contextMenuItemCount == 0) {
						THIS.hideContextMenu();
					}
				} else {
					// copyRow
					var copyRange = new contextMenuItem();
					copyRange.setParentDivId(contextMenuDiv.contextMenuId);
					copyRange.create("copy Range");
					copyRange.setLabel("Copy Range");
					copyRange.setAction(function() {
						THIS.hideContextMenu();
						THIS.copySelectedCells();
					});
					contextMenuDiv.addItem(copyRange);

					if (workbook.copiedCellArray.length != 0) {
						// Paste
						var pasteRange = new contextMenuItem();
						pasteRange.setParentDivId(contextMenuDiv.contextMenuId);
						pasteRange.create("pasteRange");
						pasteRange.setLabel("Paste");
						pasteRange.setAction(function() {
							THIS.hideContextMenu();
							THIS.pasteCopiedCells(mRowNo, mColNo);
						});
						contextMenuDiv.addItem(pasteRange);

						// Paste Special
						var pasteSpecialRange = new contextMenuItem();
						pasteSpecialRange
								.setParentDivId(contextMenuDiv.contextMenuId);
						pasteSpecialRange.create("pasteSpecialRange");
						pasteSpecialRange.setLabel("Paste Special");
						pasteSpecialRange.setAction(function() {
							THIS.hideContextMenu();
							THIS.pasteSpecialCopiedCells(mRowNo, mColNo);
						});
						contextMenuDiv.addItem(pasteSpecialRange);
					}
					// Delete
					var deleteRange = new contextMenuItem();
					deleteRange.setParentDivId(contextMenuDiv.contextMenuId);
					deleteRange.create("deleteRange");
					deleteRange.setLabel("delete shift left");
					deleteRange.setAction(function() {
						THIS.hideContextMenu();
						THIS.deleteSelectedCells(workbook.selectedRange, "left");
					});
					contextMenuDiv.addItem(deleteRange);

					// Delete
					var deleteRangeUp = new contextMenuItem();
					deleteRangeUp.setParentDivId(contextMenuDiv.contextMenuId);
					deleteRangeUp.create("deleteRangeUp");
					deleteRangeUp.setLabel("delete shift up");
					deleteRangeUp.setAction(function() {
						THIS.hideContextMenu();
						THIS.deleteSelectedCells(workbook.selectedRange, "up");
					});
					contextMenuDiv.addItem(deleteRangeUp);

					// Delete
					var deleteRangeNo = new contextMenuItem();
					deleteRangeNo.setParentDivId(contextMenuDiv.contextMenuId);
					deleteRangeNo.create("deleteRangeNo");
					deleteRangeNo.setLabel("delete no shift");
					deleteRangeNo
							.setAction(function() {
								THIS.hideContextMenu();
								THIS.deleteSelectedCells(workbook.selectedRange,
										"noShift");
							});
					contextMenuDiv.addItem(deleteRangeNo);

					var nameRangeItem = new contextMenuItem();
					nameRangeItem.setParentDivId(contextMenuDiv.contextMenuId);
					nameRangeItem.create("nameRange");
					nameRangeItem.setLabel("Name the Range");
					nameRangeItem
							.setAction(function() {
								THIS.hideContextMenu();
								var rangeArray = workbook.selectedRange.split(":");
								var rangeString = getColLabelFromNumber(parseInt(rangeArray[1]))
										+ rangeArray[0]
										+ ":"
										+ getColLabelFromNumber(parseInt(rangeArray[3]))
										+ rangeArray[2];
								workbook.nameRange("kechit", this.gridName,
										rangeString);
							});
					contextMenuDiv.addItem(nameRangeItem);

					if (THIS.displayCellArray[displayRowNo][displayColNo].cellType == "combo") {
						// Delete
						var changeValues = new contextMenuItem();
						changeValues
								.setParentDivId(contextMenuDiv.contextMenuId);
						changeValues.create("changeValues");
						changeValues.setLabel("Change Values");
						changeValues
								.setAction(function() {
									THIS.hideContextMenu();
									THIS.mode = "none";
									// document.getElementById("DebugInput6").value
									// =
									// "none";
									var comboEditValuesDiv = new comboEditValues();
									var submitAction = function() {
										var possibleValues = "";
										for ( var i = 0; i < comboEditValuesDiv.valuesInputArray.length; i++) {
											possibleValues += comboEditValuesDiv.valuesInputArray[i].value
													+ ",";
										}
										THIS
												.insertComboBox(
														mRowNo,
														mColNo,
														{
															valueOptions : {
																possibleValues : possibleValues,
																selectedIndex : 0
															}
														});
										document
												.getElementById(
														THIS.parentDivId)
												.removeChild(
														document
																.getElementById("ComboEditValuesDiv"));
										THIS.mode = "select";
										// document.getElementById("DebugInput6").value
										// = "select";

									};
									comboEditValuesDiv.createComboEditValues(
											THIS.parentDivId,
											"ComboEditValuesDiv", submitAction);
								});
						contextMenuDiv.addItem(changeValues);
					}
				}
				document.onclick = THIS.hideContextMenu;
				preventDefaultBehaviour(e);
			}

		};

		cellDiv.ondblclick = function(e) {
			var mStr = this.id.split(":")[1];
			mStr = mStr.substr(5);
			var mPos = mStr.indexOf("_");
			var mRowNo = parseInt(mStr.substr(0, mPos));
			var mColNo = parseInt(mStr.substr(mPos + 1));
			var isCurGrid = workbook.isCurGrid(THIS);

			if (THIS.mode == "edit") {
				if (!(mRowNo == THIS.curXLSRowNo && mColNo == THIS.curXLSColNo && isCurGrid)) {
				} else {
					return;
				}
			}
			
			if (!(THIS.isInXLSRange(mRowNo, mColNo)
					&& THIS.isLabel(mRowNo, mColNo))){
				THIS.goToCell(mRowNo, mColNo);
				var editor = THIS.displayCellArray[THIS.curXLSRowNo][THIS.curXLSColNo].editor;
				if (editor){
						editor.focus();
						THIS.mode = "edit";
				}
			}
			
			preventDefaultBehaviour(e);
		};

		cellDiv.onfocus = function(e) {
			THIS.mode = "select";
		};

		cellDiv.onblur = function(e) {
			hideErrorNotification();
			THIS.mode = "none";
			THIS.clearDrag();
		};
		return cellDiv;
	};

	// This function copies the cell value from a given cell to a cell given by
	// rowTo and colTo
	// TODO :- Modify this function to just load value in case the widget type
	// is same and check for display values for inputs
	// If we have the formatter then we will not be needing the display value
	clsDionGrid.prototype.moveCell = function(fromXLSCellArrayObj, rowTo, colTo) {
		if (!this.isInXLSRange(rowTo, colTo))
			return;

		var toXLSCellArrayObj = this.XLSCellArray[rowTo][colTo];
		if(fromXLSCellArrayObj.cellType != toXLSCellArrayObj.cellType) return;
		if (fromXLSCellArrayObj.cellType == "input") {
			if (toXLSCellArrayObj.cellType != "input") {
				this.insertWidget("input", rowTo, colTo, fromXLSCellArrayObj);
				this.displayCellArray[rowTo][colTo].editor.rowNo = rowTo;
				this.displayCellArray[rowTo][colTo].editor.colNo = colTo;
				toXLSCellArrayObj = this.XLSCellArray[rowTo][colTo];
			}
			toXLSCellArrayObj
					.updateCellByString(fromXLSCellArrayObj.stringValue);
			toXLSCellArrayObj.displayValue = convertTo(toXLSCellArrayObj.value,
					toXLSCellArrayObj.type, toXLSCellArrayObj.displayType);
			toXLSCellArrayObj.className = fromXLSCellArrayObj.className;
			toXLSCellArrayObj.backgroundColor = fromXLSCellArrayObj.backgroundColor;
			toXLSCellArrayObj.displayType = fromXLSCellArrayObj.displayType;
		} else {
			if (this.isCurrentCell(rowTo, colTo)) {
				this.insertWidget(fromXLSCellArrayObj.cellType, rowTo, colTo,
						fromXLSCellArrayObj);
				this.displayCellArray[rowTo][colTo].editor.rowNo = rowTo;
				this.displayCellArray[rowTo][colTo].editor.colNo = colTo;
			} else
				this.insertFormatter(fromXLSCellArrayObj.cellType, rowTo,
						colTo, fromXLSCellArrayObj);

			toXLSCellArrayObj.value = fromXLSCellArrayObj.value;
			toXLSCellArrayObj.valueOptions = fromXLSCellArrayObj.valueOptions;
			toXLSCellArrayObj.backgroundColor = fromXLSCellArrayObj.backgroundColor;
			toXLSCellArrayObj.displayType = fromXLSCellArrayObj.displayType;
			toXLSCellArrayObj.className = fromXLSCellArrayObj.className;
		}
		this.refreshCell(rowTo, colTo);
	};

	// This removes the dotted border from the interpolated Range and then sets
	// the seleted Range to interpolated Range
	// and also clears the interpolated Range
	clsDionGrid.prototype.clearInterpolation = function() {
		var THIS = this;
		if (THIS.interpolatedRange != "") {
			THIS.borderMapper(THIS.interpolatedRange, "noneDashed");
			var startX = getStartX(THIS.interpolatedRange);
			var startY = getStartY(THIS.interpolatedRange);
			var endX = getEndX(THIS.interpolatedRange);
			var endY = getEndY(THIS.interpolatedRange);
			THIS.interpolate(workbook.selectedRange, THIS.interpolatedRange);
			workbook.clearRange(workbook.selectedRange);
			workbook.selectedRange = THIS.interpolatedRange;
			THIS.interpolatedRange = "";
			THIS.applyRange();
			THIS.updateScreenFromXLS();
		}
	};

	clsDionGrid.prototype.clearDragValues = function() {
		this.dragItem = "";
	};

	clsDionGrid.prototype.clearDrag = function() {
		this.dragItem = "";
	};

	// This function adjust the named cell ranges in case of delete or insert
	// Row shift and colShift are negative in case of deletes and positive in
	// case of insert
	// Affected Range is decided by the operation
	clsDionGrid.prototype.adjustNamedRange = function(affectedRange, rowShift,colShift) {
		var valArray = workbook.nameMap.valSet();
		var keyArray = workbook.nameMap.keySet();
		var toBeRemovedIndex = [];
		for ( var i = 0; i < valArray.length; i++) {
			var gridName = valArray[i].split(".")[0];
			if (this.gridName == gridName) {
				var val = valArray[i].split(".")[1];
				if (isCell(val)) {
					var mArrColRow = getCellRowCol(val.toUpperCase());
					var rowNo = parseInt(mArrColRow[1]);
					var colNo = getColNoFromLabel(mArrColRow[0]);
					if (isInRange(rowNo, colNo, affectedRange)) {
						if (rowShift < 0 || colShift < 0) {
							if(isDeleteCell(rowNo,colNo,affectedRange,rowShift,colShift)){
								toBeRemovedIndex[toBeRemovedIndex.length] = i;
								
							} else
								{
								workbook.nameMap.put(keyArray[i], gridName + "." + modifyFormula(val, rowShift, colShift));
								}
							
						} else
						{
								
							valArray[i] = gridName + "."
									+ modifyFormula(val, rowShift, colShift);
							
						}
					}
				} else if (isRange(val)) {
					if (rowShift >= 0 && colShift >= 0) {
						var range = val.split(":");
						var newRangeStart, newRangeEnd;

						var mArrColRow = getCellRowCol(range[0].toUpperCase());
						var rowNo = parseInt(mArrColRow[1]);
						var colNo = getColNoFromLabel(mArrColRow[0]);
						if (isInRange(rowNo, colNo, affectedRange)) {
							newRangeStart = modifyFormula(range[0], rowShift,
									colShift);
						} else
							newRangeStart = range[0];
						
						mArrColRow = getCellRowCol(range[1].toUpperCase());
						rowNo = parseInt(mArrColRow[1]);
						colNo = getColNoFromLabel(mArrColRow[0]);

						if (isInRange(rowNo, colNo, affectedRange)) {
							newRangeEnd = modifyFormula(range[1], rowShift,
									colShift);
						} else
							newRangeEnd = range[1];
						valArray[i] = gridName + "." + newRangeStart + ":"
								+ newRangeEnd;
					} else if (rowShift <= 0 && colShift <= 0) {
						var range = val.split(":");
						var mArrColRow = getCellRowCol(range[0].toUpperCase());
						var startRowNo = parseInt(mArrColRow[1]);
						var startColNo = getColNoFromLabel(mArrColRow[0]);

						mArrColRow = getCellRowCol(range[1].toUpperCase());
						var endRowNo = parseInt(mArrColRow[1]);
						var endColNo = getColNoFromLabel(mArrColRow[0]);
						var newRangeStart, newRangeEnd;

						if (isInRange(startRowNo, startColNo, affectedRange)) {
							if (isInsideRange(startRowNo, startColNo,
									affectedRange)) {
								newRangeStart = modifyFormula(range[0],
										rowShift, colShift);
							} else
								newRangeStart = range[0];
						} else
							newRangeStart = range[0];

						if (isInRange(endRowNo, endColNo, affectedRange)) {
							newRangeEnd = modifyFormula(range[1], rowShift,
									colShift);
						} else
							newRangeEnd = range[1];

						if (newRangeStart != newRangeEnd) {
							valArray[i] = gridName + "." + newRangeStart + ":"
									+ newRangeEnd;
						} else
							valArray[i] = gridName + "." + newRangeStart;
					}
				}
			}
		}
		for (var i = 0; i < toBeRemovedIndex.length; i++){
			workbook.nameMap.remove(keyArray[toBeRemovedIndex[i]]);
		}
		
	};

}

defineDionGrid();
