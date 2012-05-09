/*
 *   copyright (c) -- chaitanya nalla and kechit goyal
 *
 */﻿

function defineDionDisplayFunctions() {

	//This function saves the cell in memory
	//This function is called on cell blur
    clsDionGrid.prototype.saveOneScreenCellInMemory = function(mRowNo, mColNo) {
    	var start = new Date();
		if (this.isCurrentCell(mRowNo, mColNo)){
			var displayCellArrayDiv = this.displayCellArray[mRowNo][mColNo];
			if (displayCellArrayDiv.editor && displayCellArrayDiv.editor.isValueChanged()){
				var XLSCellArrayObj = this.XLSCellArray[mRowNo][mColNo];
				if (displayCellArrayDiv.cellType == "input" || displayCellArrayDiv.cellType == "notional"){
					var editor = displayCellArrayDiv.editor;
					editor.applyValue(XLSCellArrayObj);
					var className = XLSCellArrayObj.className;
					if(XLSCellArrayObj.errorMesg != ""){
						className = addClassName(XLSCellArrayObj.className, " "+ this.defaultErrorStyle);
						showErrorNotification(XLSCellArrayObj.errorMesg, displayCellArrayDiv);
					}
					else {
						className = removeClassName(XLSCellArrayObj.className, this.defaultErrorStyle);
					}
					this.mode = "none";
					displayCellArrayDiv.className = className;
					this.onCellChanged.raiseCallBacks( null, {row: (mRowNo), cell: (mColNo), grid : this});
				}
				else {
					var editor = displayCellArrayDiv.editor;
					var strVal = editor.getValue();
					editor.applyValue(XLSCellArrayObj);
					var className = XLSCellArrayObj.className;
					if(XLSCellArrayObj.errorMesg != ""){
						className = addClassName(XLSCellArrayObj.className," " + this.defaultErrorStyle);
						showErrorNotification(XLSCellArrayObj.errorMesg, displayCellArrayDiv);
					}
					else {
						className = removeClassName(XLSCellArrayObj.className, this.defaultErrorStyle);
					}
					this.mode = "none";
					displayCellArrayDiv.className = className;
					this.onCellChanged.raiseCallBacks( null, {row: (mRowNo), cell: (mColNo), grid: this});
				}
			}
			
			//document.getElementById("DebugInput6").value = "none";
			this.dragItem = "";
			
		}
		
		var end = new Date();
		//log("Save One Screen Cell in Memory" + (end-start));
    };
	
	//TODO
	//Modify this function to include the formatter
	clsDionGrid.prototype.refreshCell = function (mRowNo, mColNo){
		var start = new Date();
		
		var screenRowNo = mRowNo;
		var screenColNo = mColNo;
		var XLSCellArrayObj = this.XLSCellArray[mRowNo][mColNo];
		if (this.isCurrentCell(mRowNo, mColNo) && FinsheetEditors.hasFormatter(XLSCellArrayObj.cellType)){
			var editor = this.displayCellArray[screenRowNo][screenColNo].editor;
			if (editor){
				editor.loadValue(XLSCellArrayObj);
			}
			else {
				this.insertFormatter(XLSCellArrayObj.cellType,mRowNo,mColNo, XLSCellArrayObj);
			}
			this.displayCellArray[screenRowNo][screenColNo].className = XLSCellArrayObj.className;
			if (this.showFormulaBar)this.formulaBarInput.value = XLSCellArrayObj.stringValue;
			
		}
		else if (FinsheetEditors.hasFormatter(XLSCellArrayObj.cellType)){
			this.insertFormatter(XLSCellArrayObj.cellType,mRowNo,mColNo, XLSCellArrayObj);
			this.displayCellArray[screenRowNo][screenColNo].className = XLSCellArrayObj.className;
		}
		else {
			/*var editor = this.displayCellArray[screenRowNo][screenColNo].editor;
			if (editor){
				editor.loadValue(XLSCellArrayObj);
			}
			else*/ if (XLSCellArrayObj.readOnly == false){
				this.insertWidget(XLSCellArrayObj.cellType,mRowNo,mColNo, XLSCellArrayObj);
			}
			this.displayCellArray[screenRowNo][screenColNo].className = XLSCellArrayObj.className;
		}
		this.refreshCellCSS(mRowNo, mColNo);

		var end = new Date();
		//log("Save One Screen Cell in Memory" + (end-start));
	};
	
	clsDionGrid.prototype.refreshCellCSS = function (mRowNo, mColNo){
	    var screenRowNo = mRowNo;
		var screenColNo = mColNo;
		if (this.isInXLSRange(mRowNo, mColNo)){
			var XLSCellArrayObj = this.XLSCellArray[mRowNo][mColNo];
		    var className = XLSCellArrayObj.className;
		    if(XLSCellArrayObj.errorMesg != ""){
				className = addClassName(XLSCellArrayObj.className, " " + this.defaultErrorStyle);
				//showErrorNotification(XLSCellArrayObj.errorMesg);
			}
			else {
				className = removeClassName(XLSCellArrayObj.className, this.defaultErrorStyle);
			}
			this.displayCellArray[screenRowNo][screenColNo].className = className;
			
			if (XLSCellArrayObj.isVisible == true) {
				this.displayCellArray[screenRowNo][screenColNo].style.visibility='visible';
				//displayCellArrayObjStyle.display='block';
			}
			else {
				//displayCellArrayObjStyle.display="none";
				this.displayCellArray[screenRowNo][screenColNo].style.visibility='hidden';
			}
		}
	}

    clsDionGrid.prototype.updateScreenFromXLS = function() {
    	var start = new Date();
		
		//This loop updates the row headers and whether it is active or not
		
		var displayCellArrayDiv;
		var XLSCellArrayObj;
		for (var i= 1; (i <= this.XLSRowCount) && (i <= this.XLSRowCount); i++) {
			for (var j= 1; (j <= this.XLSColCount) && (j <= this.XLSColCount); j++) {
				if (this.XLSCellArray[i] && this.XLSCellArray[i ][j]){
					XLSCellArrayObj = this.XLSCellArray[i][j];
				}
				else XLSCellArrayObj = {cellType : "input", value : "", displayValue : "", stringValue : ""};
				displayCellArrayDiv = this.displayCellArray[i][j];

				var cellStyle = (i % 2) ? this.defaultEvenEditableCellStyle
						: this.defaultOddEditableCellStyle;
				if (XLSCellArrayObj.className) cellStyle = XLSCellArrayObj.className;
				
				if (! this.isCurrentCell(i, j) && FinsheetEditors.hasFormatter(XLSCellArrayObj.cellType) ||  XLSCellArrayObj.readOnly){
					this.insertFormatter(XLSCellArrayObj.cellType,i,j, XLSCellArrayObj);
				}
				else if( XLSCellArrayObj.readOnly == false ){
					this.insertWidget(XLSCellArrayObj.cellType,i,j, XLSCellArrayObj);
				}
				
				if (this.XLSCellArray[i] && this.XLSCellArray[i][j]){
					if (this.XLSCellArray[i][j].errorMesg != ""){
						displayCellArrayDiv.className = cellStyle +  " " + this.defaultErrorStyle;
					}
					else displayCellArrayDiv.className = cellStyle;
				}

				else {
					displayCellArrayDiv.className = cellStyle + " " + this.defaultTextAlign +"AlignCellClass ";
				}
				
				
            }
        }
		this.applyRange();
		var end = new Date();
		//log("Save One Screen Cell in Memory" + (end-start));
    };
    
	clsDionGrid.prototype.applyRange = function(){
		var start = new Date();
		var displayCellArrayObj;
		for (var i= 1; (i <= this.XLSRowCount); i++) {
			for (var j= 1; j <= this.XLSColCount; j++) {
				displayCellArrayObj = this.displayCellArray[i][j];
				displayCellArrayObj.className = removeClassName(displayCellArrayObj.className, " " + "focusRange");
				//Set the type whether it is active cell or not
				if (isInRange(i ,j, workbook.selectedRange)){
					if (!this.isLabel(i,j)){
						displayCellArrayObj.className = addClassName(displayCellArrayObj.className, " " + "focusRange");
					}
				}
				else {
					if (this.XLSCellArray[i] && this.XLSCellArray[i][j]){
						displayCellArrayObj.style.background = this.XLSCellArray[i][j].backgroundColor;
					}
				}				
            }
        }
		if (this.showInterpolationDiv){
			var range = workbook.selectedRange;
			var endX = getEndX(range);
			var endY = getEndY(range);
			if (endX  > 0 && endX <= this.XLSRowCount && endY  > 0 && endY <= this.XLSColCount){
				this.interpolationDiv.style.top = (getLengthFromString(this.displayCellArray[endX][endY].style.top) + getLengthFromString(this.displayCellArray[endX][endY].style.height) - 4) + "px";
				this.interpolationDiv.style.left = (getLengthFromString(this.displayCellArray[endX][endY].style.left) + getLengthFromString(this.displayCellArray[endX][endY].style.width) - 4) + "px";
				this.interpolationDiv.style.display = "block";
			}
			else {
				this.interpolationDiv.style.display = "none";
			}
		}
		var end = new Date();
		//log("Refresh Cell : " + (end-start));
	};
	
	//This function sets the Dimension and Position of the row and column headers and all the cells
	clsDionGrid.prototype.refreshGridCells = function() {
		var start = new Date();
		
		var xPos = 0;
		var yPos = 0;
		var XLSRowArrayObj, XLSColArrayObj, displayCellArrayObj, displayCellArrayObjStyle;
		
		this.removeGridDiv();
		
		
		for (var i = 1;i <= this.XLSRowCount ; i++){
			XLSRowArrayObj = (this.XLSRowArray[i]) ? this.XLSRowArray[i] : {height : this.defaultCellHeight, gap:0};
			
			for (var j = 1;j <= this.XLSColCount; j++){
				//We can not break here as we did in updateScreen as we want to set the visibility of all the other cells to false
				XLSColArrayObj = (this.XLSColArray[j]) ? this.XLSColArray[j] : {width : this.defaultCellWidth, gap:0}  ;
				displayCellArrayObj = this.displayCellArray[i][j];	
				displayCellArrayObjStyle = displayCellArrayObj.style;
				var XLSRowNo = (i);
				var XLSColNo = (j);
				
				//If it is a valid cell then visibility is visible and all the required parameters are set else it is rendered invisible
				if ((XLSRowNo <= this.XLSRowCount) && (XLSColNo <= this.XLSColCount)){
					var rowSpan = 1; var colSpan = 1; var isVisible = true;
					var XLSCellArrayObj = this.XLSCellArray[XLSRowNo][XLSColNo];
					rowSpan = XLSCellArrayObj.rowSpan;
					colSpan = XLSCellArrayObj.colSpan;
					isVisible = XLSCellArrayObj.isVisible;
					
					//Shift the cells accordingly
					displayCellArrayObjStyle.top = yPos + "px";	
					displayCellArrayObjStyle.left = xPos + "px";
					displayCellArrayObjStyle.width = Math.max(this.getTotalWidth(XLSColNo, XLSColNo + colSpan-1)-1, 0) + "px";
					displayCellArrayObjStyle.height = Math.max(this.getTotalHeight(XLSRowNo, XLSRowNo + rowSpan-1)-1,0) + "px";
					
					if (isVisible == true) {
						displayCellArrayObjStyle.visibility='visible';
					}
					else {
						displayCellArrayObjStyle.visibility='hidden';
					}
					xPos = xPos + XLSColArrayObj.width + XLSColArrayObj.gap ;
				}
				else {
					displayCellArrayObjStyle.visibility='hidden';
				}				
			}
			
			if (i <= this.XLSRowCount){
				yPos = yPos + XLSRowArrayObj.height + XLSRowArrayObj.gap;
			}
			xPos = 0;
		}
		
		if (this.showInterpolationDiv){
			var range = workbook.selectedRange;
			var endX = getEndX(range);
			var endY = getEndY(range);
			if (endX  > 0 && endX < this.XLSRowCount && endY  > 0 && endY < this.XLSColCount){
				this.interpolationDiv.style.top = (getLengthFromString(this.displayCellArray[endX][endY].style.top) + getLengthFromString(this.displayCellArray[endX][endY].style.height) - 4) + "px";
				this.interpolationDiv.style.left = (getLengthFromString(this.displayCellArray[endX][endY].style.left) + getLengthFromString(this.displayCellArray[endX][endY].style.width) - 4) + "px";
				this.interpolationDiv.style.display = "block";
			}
			else {
				this.interpolationDiv.style.display = "none";
			}
		}
		this.attachGridDiv();
		var end = new Date();
		//log("Refresh Grid Cells : " + (end-start));
	};
	
	clsDionGrid.prototype.refreshSpan = function(){
		var start = new Date();
		
		for (var i = 1;i <= this.XLSRowCount; i++){
			for (var j = 1;j <= this.XLSColCount; j++){
				var XLSRowNo = i;
				var XLSColNo = j;
				var rowSpan = 1; var colSpan = 1; var isVisible = true;
				var XLSCellArrayObj = this.XLSCellArray[XLSRowNo][XLSColNo];
				rowSpan = XLSCellArrayObj.rowSpan;
				colSpan = XLSCellArrayObj.colSpan;
			
				if (rowSpan > 1 || colSpan > 1){
					for (var affectedRow = XLSRowNo; affectedRow < XLSRowNo + rowSpan; affectedRow++){
						for (var affectedCol = XLSColNo; affectedCol < XLSColNo + colSpan; affectedCol++){
							if (affectedRow == XLSRowNo && affectedCol == XLSColNo) continue;
							
							this.XLSCellArray[affectedRow][affectedCol].isVisible = false;
							this.XLSCellArray[affectedRow][affectedCol].spannedBy = {row: XLSRowNo, col: XLSColNo};
						}	
					}
				}
				if (this.XLSCellArray[XLSRowNo][XLSColNo].spannedBy.row != 0 && this.XLSCellArray[XLSRowNo][XLSColNo].spannedBy.col != 0){
					var spannedByCellRowNo = this.XLSCellArray[XLSRowNo][XLSColNo].spannedBy.row
					var spannedByCellColNo = this.XLSCellArray[XLSRowNo][XLSColNo].spannedBy.col
					var spannedByCell = this.XLSCellArray[spannedByCellRowNo][spannedByCellColNo];
					if (spannedByCell.rowSpan < (XLSRowNo-spannedByCellRowNo+1)){
						this.XLSCellArray[XLSRowNo][XLSColNo].spannedByCell = {row:0, col:0};
						this.XLSCellArray[XLSRowNo][XLSColNo].isVisible = true;
					}
					if (spannedByCell.colSpan < (XLSColNo-spannedByCellColNo+1)){
						this.XLSCellArray[XLSRowNo][XLSColNo].spannedByCell = {row:0, col:0};
						this.XLSCellArray[XLSRowNo][XLSColNo].isVisible = true;
					}
				}
			}
		}
		var end = new Date();
		//log("Refresh Span : " + (end-start));
	};
	
	clsDionGrid.prototype.refreshHeaders = function(){
		var start = new Date();
		
		var formatDivHeight = 0;
		var loneSquareHeight = 0;
		var loneSquareWidth = 0;
		
		this.gridDiv.style.left= this.left + loneSquareWidth + "px";
		this.gridDiv.style.top = this.top + formatDivHeight + loneSquareHeight + "px";
	
		var hasVerticalScroll = false;
		var hasHorizontalScroll = false;
		
		gridWidth = this.getTotalWidth(1,this.XLSColCount) ;
		gridHeight = this.getTotalHeight(1,this.XLSRowCount) ;
		
		this.gridDiv.style.width = (gridWidth  + 1) + "px";
		this.gridDiv.style.height = (gridHeight + 1) +"px";
		
		var end = new Date();
		//log("Refresh Headers : " + (end-start));
	};
	
	//This function checks whehter the given cell is already visible
	//if it is visible then do nothing
	//Else find out the required screen start row number and column number and update the screen grid from xls
    clsDionGrid.prototype.makeVisibleOnScreen = function(mRowNo, mColNo) {
    };
	
	//This function makes the cell active
    clsDionGrid.prototype.goToCell = function(mRowNo, mColNo) {
    	var start = new Date();
    	
    	// If it is an invalid cell then return
		if (!this.isInXLSRange(mRowNo, mColNo)) {
            return;
        }
        //If it is already the current cell then return
        var isCurGrid = workbook.isCurGrid(this);
        if (!isCurGrid) workbook.curDionGrid = this;
		if (this.curXLSRowNo == mRowNo && this.curXLSColNo == mColNo && isCurGrid && this.displayCellArray[mRowNo][mColNo].editor){
			this.displayCellArray[mRowNo][mColNo].editor.blur();
			this.applyRange();
			//this.displayCellArray[mRowNo][mColNo].className = removeClassName(this.displayCellArray[mRowNo][mColNo].className, "focusRange");
			this.mode = "select";
			return;
		}
		
		else {
			workbook.clearRange(workbook.selectedRange);
		}
		
		//Destroy the current cell's editor
		var curCellDiv = this.displayCellArray[this.curXLSRowNo][this.curXLSColNo];
		var cellType = this.XLSCellArray[this.curXLSRowNo][this.curXLSColNo].cellType;
		if (curCellDiv.editor){
			if (FinsheetEditors.hasFormatter(cellType)){
				if (curCellDiv.editor && curCellDiv.editor.blur)
					curCellDiv.editor.blur();
				curCellDiv.editor.destroy();
				curCellDiv.editor = null;
			}
			else {
				if (curCellDiv.editor && curCellDiv.editor.blur)
					curCellDiv.editor.blur();
				curCellDiv.editor.destroy();
			}
		}
		
		//Set Insert Formatter in the current cell
		var XLSCellArrayObj = this.XLSCellArray[this.curXLSRowNo][this.curXLSColNo];
		if (FinsheetEditors.hasFormatter(XLSCellArrayObj.cellType))
			this.insertFormatter(XLSCellArrayObj.cellType,this.curXLSRowNo,this.curXLSColNo, XLSCellArrayObj);
		this.mode  = "select";
		
		
		if (!this.isLabel(mRowNo,mColNo)){
			//Set the current cell of the grid
			this.curXLSRowNo = mRowNo;
			this.curXLSColNo = mColNo;
			workbook.curXLSRowNo = mRowNo;
			workbook.curXLSColNo = mColNo;
			workbook.clearRange(workbook.selectedRange);
			workbook.selectedRange = mRowNo + ":" + mColNo + ":" + mRowNo + ":" + mColNo;
		}
		else{
			workbook.clearRange(workbook.selectedRange);
			workbook.selectedRange = this.curXLSRowNo + ":" + this.curXLSColNo + ":" + this.curXLSRowNo + ":" + this.curXLSColNo;
		}
			
		if (this.isInXLSRange(this.curXLSRowNo, this.curXLSColNo)){
			if (this.XLSCellArray[this.curXLSRowNo] && this.XLSCellArray[this.curXLSRowNo][this.curXLSColNo])
				XLSCellArrayObj = this.XLSCellArray[this.curXLSRowNo][this.curXLSColNo];
			else XLSCellArrayObj = {cellType : "input", value: "", displayValue : "", stringValue: ""};
			
			if (XLSCellArrayObj.errorMesg){
				showErrorNotification(XLSCellArrayObj.errorMesg, this.displayCellArray[this.curXLSRowNo][this.curXLSColNo]);
			}
			
			if (FinsheetEditors.hasFormatter(XLSCellArrayObj.cellType) && XLSCellArrayObj.readOnly == false)
				this.insertWidget(XLSCellArrayObj.cellType,this.curXLSRowNo,this.curXLSColNo, XLSCellArrayObj);
				
			this.mode="select";
			
			var cellDiv = this.displayCellArray[this.curXLSRowNo][this.curXLSColNo];
			if (cellDiv.editor){
				//TODO : KECHIT Could be erroneous;
				cellDiv.focus();
				//this.mode = "select";
				if (cellDiv.editor.autofocus() == false){
					this.mode = "select";
				}
				else {
					cellDiv.editor.focus();
				}
			}
			
		}
		this.applyRange();
		var end = new Date();
		//log("Go To Cell" + (end-start));
    };
	
	//get ths total width including the start and end columns
	clsDionGrid.prototype.getTotalWidth= function(startCol, endCol){
		//Set the start and the end column number if the value is out of bounds
		if (startCol <= 0) startCol = 1;
		else if (startCol > this.XLSColCount) startCol = this.XLSColCount;
		
		if (endCol <= 0) endCol = 1;
		else if (endCol > this.XLSColCount) endCol = this.XLSColCount;
		
		if (endCol < startCol) endCol = startCol;
		
		//Add the width from XLS Col Array to get the total width
		var width = 0;
		for (var i = startCol; i <= endCol; i++){
			if (i != endCol)
				width += (this.XLSColArray[i]) ? (this.XLSColArray[i].width + this.XLSColArray[i].gap) : this.defaultCellWidth;
			else 
				width += (this.XLSColArray[i]) ? this.XLSColArray[i].width : this.defaultCellWidth;
		}
		return width;
	};
	
	//Get the total height including the start and the end rows
	clsDionGrid.prototype.getTotalHeight= function(startRow, endRow){
		//Set the start and the end row number if the value is out of bounds
		if (startRow <= 0) startRow = 1;
		else if (startRow > this.XLSRowCount) startRow = this.XLSRowCount;
		
		if (endRow <= 0) endRow = 1;
		else if (endRow > this.XLSRowCount) endRow = this.XLSRowCount;
		
		if (endRow < startRow) endRow = startRow;
		
		//Add the height from XLS Row Array to get the total height
		var height = 0;
		for (var i = startRow; i <= endRow; i++){
			if( i != endRow)
				height += (this.XLSRowArray[i]) ? this.XLSRowArray[i].height + this.XLSRowArray[i].gap: this.defaultCellHeight;
			else
				height += (this.XLSRowArray[i]) ? this.XLSRowArray[i].height: this.defaultCellHeight;
		}
		return height;
	};
	
	clsDionGrid.prototype.getLastRow = function(){
		return this.XLSRowCount;
	};

	clsDionGrid.prototype.getLastCol = function(){
		return this.XLSColCount;
	};

	//Hide the context menu and set the document.onmouse up listener to null
	clsDionGrid.prototype.hideContextMenu = function() {
		//Remove the already existing context menu div
		if (workbook.contextMenuDiv != null){
			workbook.contextMenuDiv.hideSubMenu();
			document.body.removeChild(document.getElementById(workbook.contextMenuDiv.contextMenuId));
			workbook.contextMenuDiv = null;
			//set the document.onclick listener to null
			document.onclick = null;	
		}
	};

	clsDionGrid.prototype.setNext = function (cell1, cell2){
		if (! isCell(cell1) || !isCell(cell2)) return;
		
		var mArrColRow1 = getCellRowCol(cell1);
		var rowNo1 = parseInt(mArrColRow1[1]);
		var colNo1 = getColNoFromLabel(mArrColRow1[0]);
		
		var mArrColRow2 = getCellRowCol(cell2);
		var rowNo2 = parseInt(mArrColRow2[1]);
		var colNo2 = getColNoFromLabel(mArrColRow2[0]);
		
		if (! this.isInXLSRange(rowNo1, colNo1) || !this.isInXLSRange(rowNo2, colNo2)) return;
		var cellObj1 = this.XLSCellArray[rowNo1][colNo1];
		var cellObj2 = this.XLSCellArray[rowNo2][colNo2];
		cellObj1.nextCell = cell2;
		cellObj2.previousCell = cell1;
	};
	
	clsDionGrid.prototype.multipleSetNext = function(mArray){
		if(mArray.length > 0){
			for (var i = 0; i+1 < mArray.length; i = i+2){
				this.setNext(mArray[i], mArray[i+1]);
			}
		}
	};

	clsDionGrid.prototype.rightEditableCell = function (){
		var curRow = this.curXLSRowNo;
		var curCol = this.curXLSColNo;
		
		for(var j = curCol+1;j <= this.XLSColCount;j++)
		{
			if(this.XLSCellArray[curRow][j].cellType != 'label' && this.XLSCellArray[curRow][j].isVisible==true)
				return {rowNo: curRow,colNo: j};
			
		}
		
		var prevRow = curRow;
		var prevCol = curCol;
		
		if(this.nextGrid.right.grid != null)
		{
			var nextGrid = this.nextGrid.right.grid;
			curRow = this.curXLSRowNo + this.nextGrid.right.row_offset;
			curCol = 1 + this.nextGrid.right.col_offset;
			
			for(var j = curCol;j <= nextGrid.XLSColCount;j++)
			{
				if(nextGrid.XLSCellArray[curRow][j].cellType != 'label' && nextGrid.XLSCellArray[curRow][j].isVisible==true)
				{
					workbook.clearRange(workbook.selectedRange);
					workbook.curDionGrid = nextGrid;
					nextGrid.curXLSRowNo = 1;
					nextGrid.curXLSColNo = 1;
					
					return {rowNo: curRow,colNo: j};
				}
			}
		}	
			
			
		return {rowNo: prevRow, colNo:prevCol};
	};
	
	clsDionGrid.prototype.leftEditableCell = function (){
		var curRow = this.curXLSRowNo;
		var curCol = this.curXLSColNo;
		
		for(var j = curCol-1;j > 0;j--)
		{
			if(this.XLSCellArray[curRow][j].cellType != 'label'  && this.XLSCellArray[curRow][j].isVisible==true)
				return {rowNo: curRow,colNo: j};
			
		}
		var prevRow = curRow;
		var prevCol = curCol;
		if(this.nextGrid.left.grid != null)
		{
			var nextGrid = this.nextGrid.left.grid;
			curRow = this.curXLSRowNo + this.nextGrid.left.row_offset;
			curCol = nextGrid.XLSColCount - this.nextGrid.left.col_offset;
		
			for(var j = curCol;j > 0;j--)
			{
				if(nextGrid.XLSCellArray[curRow][j].cellType != 'label'  && nextGrid.XLSCellArray[curRow][j].isVisible==true)
				{
					workbook.clearRange(workbook.selectedRange);
					workbook.curDionGrid = nextGrid;
					nextGrid.curXLSRowNo = 1;
					nextGrid.curXLSColNo = 1;
					
					return {rowNo: curRow,colNo: j};
				}
			}
		}
		return {rowNo: prevRow, colNo:prevCol};
};
	
	clsDionGrid.prototype.upEditableCell = function (){
		var curRow = this.curXLSRowNo;
		var curCol = this.curXLSColNo;

			for(var i = curRow-1;i > 0;i--)
			{
				if(this.XLSCellArray[i][curCol].cellType != 'label'  && this.XLSCellArray[i][curCol].isVisible==true)
					return {rowNo: i,colNo: curCol};
				
			}
			var prevRow = curRow;
			var prevCol = curCol;	
			if(this.nextGrid.up.grid != null)
			{
				var nextGrid = this.nextGrid.up.grid;
				curRow =  nextGrid.XLSRowCount- this.nextGrid.up.row_offset;
				curCol = this.curXLSColNo + this.nextGrid.up.col_offset;
				

				for(var i = curRow;i > 0;i--)
				{
					if(nextGrid.XLSCellArray[i][curCol].cellType != 'label'  && nextGrid.XLSCellArray[i][curCol].isVisible==true)
					{
						workbook.clearRange(workbook.selectedRange);
						workbook.curDionGrid = nextGrid;
						nextGrid.curXLSRowNo = 1;
						nextGrid.curXLSColNo = 1;
						
						return {rowNo: i,colNo: curCol};
					}
				}
			}
			return {rowNo: prevRow, colNo:prevCol};
	};
	
	clsDionGrid.prototype.downEditableCell = function (){
		var curRow = this.curXLSRowNo;
		var curCol = this.curXLSColNo;
		
		for(var i = curRow+1;i <= this.XLSRowCount;i++)
		{
			if(this.XLSCellArray[i][curCol].cellType != 'label'  && this.XLSCellArray[i][curCol].isVisible==true)
				return {rowNo: i,colNo: curCol};
			
		}
		var prevRow = curRow;
		var prevCol = curCol;
			if(this.nextGrid.down.grid != null)
			{
				var nextGrid = this.nextGrid.down.grid;
				curRow = 1 + this.nextGrid.down.row_offset;
				curCol = this.curXLSColNo + this.nextGrid.down.col_offset;
				for(var i = curRow;i <= nextGrid.XLSRowCount;i++)
				{
					if(nextGrid.XLSCellArray[i][curCol].cellType != 'label'  && nextGrid.XLSCellArray[i][curCol].isVisible==true)
					{
						workbook.clearRange(workbook.selectedRange);
						workbook.curDionGrid = nextGrid;
						nextGrid.curXLSRowNo = 1;
						nextGrid.curXLSColNo = 1;
						return {rowNo: i,colNo: curCol};
					}	
					
				}
			}
				return {rowNo: prevRow,colNo : prevCol};
	};

	
	
	
	clsDionGrid.prototype.nextEditableCell = function (){
		var curRow = this.curXLSRowNo;
		var curCol = this.curXLSColNo;
		
		if (curRow <= this.XLSRowCount){
			for (var j = curCol+1; j <= this.XLSColCount; j++){
				if ((this.XLSCellArray[curRow][j].cellType != "label" && this.XLSCellArray[curRow][j].isVisible==true)) return {rowNo: curRow, colNo:(j)};
			}
		}
		for (var i = curRow+1; i <= this.XLSRowCount; i++){
			for (var j = 1; j <= this.XLSColCount; j++){
				
				if (this.XLSCellArray[i][j].cellType != 'label' && this.XLSCellArray[i][j].isVisible==true) return {rowNo: i, colNo:j};
			}
		}
		
		var prevRow = curRow;
		var prevCol = curCol;
		
		//moving to next grid adjacent to the current one
		if(this.nextGrid.right.grid != null)
		{
			var nextGrid = this.nextGrid.right.grid;
			curRow = this.curXLSRowNo + this.nextGrid.right.row_offset;
			curCol = 1 + this.nextGrid.right.col_offset;

			if (curRow <= nextGrid.XLSRowCount){
				for (var j = curCol; j <= nextGrid.XLSColCount; j++){
				if ((nextGrid.XLSCellArray[curRow][j].cellType != "label" && nextGrid.XLSCellArray[curRow][j].isVisible==true)){
				workbook.clearRange(workbook.selectedRange);
				workbook.curDionGrid = nextGrid;
				nextGrid.curXLSRowNo = 1;
				nextGrid.curXLSColNo = 1;
				
				return {rowNo: curRow, colNo:(j)};
				}
				}
			}
			for (var i = curRow+1; i <= nextGrid.XLSRowCount; i++)
			{
				for (var j = 1; j <= nextGrid.XLSColCount; j++){
				
				if (nextGrid.XLSCellArray[i][j].cellType != 'label' && nextGrid.XLSCellArray[i][j].isVisible==true){
				
				workbook.clearRange(workbook.selectedRange);
				workbook.curDionGrid = nextGrid;
				return {rowNo: i, colNo:j};
				}
				}
			}
		}
		return {rowNo: prevRow, colNo:prevCol};
		
	};

	clsDionGrid.prototype.previousEditableCell = function (){
		var curRow = this.curXLSRowNo;
		var curCol = this.curXLSColNo;
		
		for (var j = curCol-1; j > 0; j--){
			if ((this.XLSCellArray[curRow][j].cellType != 'label' && this.XLSCellArray[curRow][j].isVisible==true)) return {rowNo: curRow, colNo:j};
		}
		
		for (var i = curRow-1; i > 0; i--){
			for (var j = this.XLSColCount; j > 0; j--){
				if ((this.XLSCellArray[i][j].cellType != 'label' && this.XLSCellArray[i][j].isVisible==true)) return {rowNo: i, colNo:j};
			}
		}
		
		var prevRow = curRow;
		var prevCol = curCol;
		if(this.nextGrid.left.grid != null)
		{
			workbook.clearRange(workbook.selectedRange);
			var nextGrid = this.nextGrid.left.grid;
			var curRow = this.XLSRowNo + this.nextGrid.left.row_offset;
			var curCol = nextGrid.XLSColCount - this.nextGrid.left.col_offset;
			for (var j = curCol; j > 0; j--){
				if ((nextGrid.XLSCellArray[curRow][j].cellType != 'label' && nextGrid.XLSCellArray[curRow][j].isVisible==true)){
				 workbook.clearRange(workbook.selectedRange);
				 workbook.curDionGrid = nextGrid;
				 nextGrid.curXLSRowNo = 1;
					nextGrid.curXLSColNo = 1;
					
				 return {rowNo: curRow, colNo:j};
			}
			}
			
			for (var i = curRow-1; i > 0; i--){
				for (var j = nextGrid.XLSColCount; j > 0; j--){
					if ((nextGrid.XLSCellArray[i][j].cellType != 'label' && nextGrid.XLSCellArray[i][j].isVisible==true)){
					 
					 workbook.clearRange(workbook.selectedRange);
					 workbook.curDionGrid = nextGrid;
					 return {rowNo: i, colNo:j};
					}
				}
			}
			
			
		}
		
		return {rowNo: prevRow, colNo:prevCol};
	};
}

defineDionDisplayFunctions();

