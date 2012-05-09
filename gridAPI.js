/*
 *   copyright (c) -- chaitanya nalla and kechit goyal
 *
 */

function defineGridAPI(){

//------------------------Get or Set Values -------------------------------------------------------------------

//--------------------Note Just for Set Data and getData-------------------------------------------
	//i and j start from 0,0 (because rowNo and colNo start from 0,0 in Harsh's Grid)
	clsDionGrid.prototype.setData = function(i,j, value, valueOptions, cellType){
		if (!this.isInXLSRange(i,j)) return;
		var XLSCellArrayObj = this.XLSCellArray[i][j];
		if (valueOptions){
			XLSCellArrayObj.valueOptions = valueOptions;
		}
		if (cellType) XLSCellArrayObj.cellType = cellType;
		if (XLSCellArrayObj.cellType == "none"){
			XLSCellArrayObj.readOnly = true;
			XLSCellArrayObj.cellType = "input";
		}
		if (XLSCellArrayObj.cellType == "input"  || XLSCellArrayObj.cellType == "notional"){
			XLSCellArrayObj.errorMesg = "";
			XLSCellArrayObj.updateCellByString(value);
		}
		else {
			XLSCellArrayObj.errorMesg = "";
			XLSCellArrayObj.value = value;
			if (valueOptions){
				XLSCellArrayObj.valueOptions = valueOptions;
			}
			else XLSCellArrayObj.valueOptions.value = value;
			XLSCellArrayObj.updateCellByWidget();
			
		}
		this.refreshCell(i,j);
		//this.fireOnChange(i,j);
	};
	
	clsDionGrid.prototype.modifyData = function(i,j, valueOptions){
		if (!this.isInXLSRange(i,j)) return;
		
		if (this.XLSCellArray[i][j].cellType == "input" ||this.XLSCellArray[i][j].cellType == "notional"){
			this.XLSCellArray[i][j].updateCellByString(value);
		}
		else {
			if (valueOptions){
				for (var opt in valueOptions){
					this.XLSCellArray[i][j].valueOptions[opt] = valueOptions[opt];
				}
			}
			this.XLSCellArray[i][j].updateCellByWidget();
			
		}
		this.refreshCell(i,j);
	};
	
	clsDionGrid.prototype.setCellId = function (i,j, cellId){
		if (!this.isInXLSRange(i,j)) return;
		workbook.setCellId(cellId, this, this.XLSCellArray[i][j]);
		this.XLSCellArray[i][j].id = cellId;
	}
	
	clsDionGrid.prototype.getCellId = function (cellId){
		return workbook.getCellById(cellId, this);
	}
	
	clsDionGrid.prototype.setHighlight = function(i,j, value){
		if (!this.isInXLSRange(i,j)) return;
		var XLSCellArrayObj = this.XLSCellArray[i][j];
		if (value && this.XLSCellArray[i][j].isVisible == true){
			XLSCellArrayObj.className = addClassName(XLSCellArrayObj.className, " highlightCell");
		}
		else {
			XLSCellArrayObj.className = removeClassName(XLSCellArrayObj.className, "highlightCell");
		}
		this.refreshCell(i,j);
	};
	
	clsDionGrid.prototype.setError = function(i,j,message){
		if (!this.isInXLSRange(i,j)) return;
		if (this.XLSCellArray[i][j].isVisible == true){
		    this.XLSCellArray[i][j].errorMesg = message;
		    this.refreshCellCSS(i, j);
		    if (message){
		    	showErrorNotification(message, this.displayCellArray[i][j]);
		    }
		}
	    else this.XLSCellArray[i][j].errorMesg = "";
	};
	

	clsDionGrid.prototype.hasError = function(i, j) {
		if (!this.isInXLSRange(i, j))
			return;
		var message = this.XLSCellArray[i][j].errorMesg;
		if (message) {
			return true;
		} else
			return false;

	};
	
	clsDionGrid.prototype.setVisibility = function(i,j,value){
		if (!this.isInXLSRange(i,j)) return;
		this.XLSCellArray[i][j].isVisible = value;
		this.refreshCell(i,j);
	};
	
	clsDionGrid.prototype.setTag = function(i,j,value){
		if (!this.isInXLSRange(i,j)) return;
		this.XLSCellArray[i][j].tag = value;
	};
	
	clsDionGrid.prototype.setType = function(i,j,value){
		if (!this.isInXLSRange(i,j)) return;
		this.XLSCellArray[i][j].cellType = value;
		this.refreshCell(i,j);
	};
	
	clsDionGrid.prototype.setCurrentCell = function (i, j){
		if (!this.isInXLSRange(i,j)) return;
		this.curXLSRowNo = i;
		this.curXLSColNo = j;
	};
	
	clsDionGrid.prototype.setOnCellClick = function(fn){
		if (fn)
			this.onCellClicked = fn;
	};
	

	//i and j start from 0,0 (because rowNo and colNo start from 0,0 in Harsh's Grid)
	clsDionGrid.prototype.getData = function (i,j, stringValue){
		if (!this.isInXLSRange(i,j)) return null;
		var val = (stringValue)? this.XLSCellArray[i][j].stringValue : this.XLSCellArray[i][j].value;
		return val;
	};
	
	clsDionGrid.prototype.getOptions = function (i,j){
		if (!this.isInXLSRange(i,j)) return null;
		var val = this.XLSCellArray[i][j].valueOptions;
		return val;
	};
	
	clsDionGrid.prototype.getTag = function (i,j){
		if (!this.isInXLSRange(i,j)) return null;
		var val = this.XLSCellArray[i][j].tag;
		return val;
	};
	
	clsDionGrid.prototype.setContextMenuFn = function (fn){
		this.contextMenuFn = fn;
	};
//---------------------------------------------------------------------------------------------	
	
//	Set parent Div Id is necessary as parent Div is the container for the grid.
//	Also it then adds parent div on mouse down handler that just sets the current grid to this grid.... 
	clsDionGrid.prototype.setParentDivId = function (parentDivId){
		var THIS = this;
		var parentDiv = document.getElementById(parentDivId);
		this.parentDivElement = parentDiv;
		removeAllChildren(parentDiv);
		parentDiv.onmousedown = function (){
			if (!workbook.curDionGrid) workbook.curDionGrid = THIS;
			else if (!(workbook.curDionGrid.isFormula(workbook.curXLSRowNo, workbook.curXLSColNo) && workbook.curDionGrid.mode == "edit"))
				workbook.curDionGrid = THIS;
		};
		THIS.parentDivId = parentDivId;
	};
	
	clsDionGrid.prototype.setParentDivElement = function (parentDivElement){
		var THIS = this;
		this.parentDivElement = parentDivElement;
		removeAllChildren(parentDivElement);
		parentDivElement.onmousedown = function (){
			if (!workbook.curDionGrid) workbook.curDionGrid = THIS;
			else if (!(workbook.curDionGrid.isFormula(workbook.curXLSRowNo, workbook.curXLSColNo) && workbook.curDionGrid.mode == "edit"))
				workbook.curDionGrid = THIS;
		};
		//THIS.parentDivId = parentDivElement.id;
	};

//Height, Width , rows, cols and id are the mandatory fields 
	clsDionGrid.prototype.init = function(id, options){
		if (!id) {
			return;
		}
		
		this.width = options.width;
		this.height = options.height;
		if (options.hasAutoWidth != undefined){
			this.hasAutoWidth = options.hasAutoWidth;
			this.maxWidth = options.maxWidth;
			this.minWidth = options.minWidth;
		}
		if(options.hasAutoHeight !=undefined)
		{
			this.hasAutoHeight = options.hasAutoHeight;
			this.maxHeight = options.maxHeight;
			this.minHeight = options.minHeight;
		}
		this.top = (options.top) ? options.top : 0;
		this.left = (options.left) ? options.left : 0;
		
		this.gridName = id;
		this.loneSquareWidth = 40;
        this.loneSquareHeight = 20;
        
		this.defaultEditableCellStyle = (options.defaultEditableCellStyle)? options.defaultEditableCellStyle : "gridCellBorder ";
		this.defaultOddEditableCellStyle = (options.defaultOddEditableCellStyle)? options.defaultOddEditableCellStyle : "gridCellBorder ";
		this.defaultEvenEditableCellStyle = (options.defaultEvenEditableCellStyle)? options.defaultEvenEditableCellStyle : "gridCellBorder ";
		
		this.defaultGridDivStyle = (options.defaultGridDivStyle)? options.defaultGridDivStyle : "gridDiv ";
		
		this.defaultNonEditableCellStyle = (options.defaultNonEditableCellStyle)? options.defaultNonEditableCellStyle : "gridCell ";
		this.defaultOddNonEditableCellStyle = (options.defaultOddNonEditableCellStyle)? options.defaultOddNonEditableCellStyle : "gridCell ";
		this.defaultEvenNonEditableCellStyle = (options.defaultEvenNonEditableCellStyle)? options.defaultEvenNonEditableCellStyle : "gridCell ";
		
        this.defaultCellHeight = (options.defaultCellHeight)? options.defaultCellHeight : 20;
        this.defaultCellWidth = (options.defaultCellWidth)? options.defaultCellWidth : 120;
        this.defaultRowGap = (options.defaultRowGap)? options.defaultRowGap : 0;
        this.defaultColGap = (options.defaultColGap)? options.defaultColGap : 0;
        this.isMouseWheel = (options.isMouseWheel)? options.isMouseWheel : false;
        this.defaultTextAlign = (options.defaultTextAlign)? options.defaultTextAlign : " left";

        
		if (options.hideTopPanel){
			if (options.hideTopPanel == true){
				this.showInterpolationDiv = false;
			}
			else {
				this.showInterpolationDiv = true;
			}
		}
		
		
		this.showContextMenu = (options.showContextMenu) ? options.showContextMenu : false;
		//
		
		this.showInterpolationDiv = options.showInterpolationDiv ? options.showInterpolationDiv : false;
		this.isCustomContextMenu = options.isCustomContextMenu ? options.isCustomContextMenu : false;

		this.createFullScreen();
		this.createXLSGridInMemory();
		this.updateXLSCount(options.rows, options.cols)
		
		if (options.data){
			this.setMatrix(options.data);
		}
		
		if (options.columnWidths){
			for (var i = 0; i < options.columnWidths.length && i < this.XLSColCount; i++){
				this.XLSColArray[i+1].width = options.columnWidths[i]; 
			}
		}
		if (options.rowHeights){
			for (var i = 0; i < options.rowHeights.length && i < this.XLSRowCount; i++){
				this.XLSRowArray[i+1].height = options.rowHeights[i]; 
			}
		}
		
		this.showFormatDiv = false;
		this.refreshSpan()
		this.refreshGridCells();
		this.updateScreenFromXLS();
		this.refreshHeaders();
		
		workbook.registerGrid(id, this);
	};

	clsDionGrid.prototype.setMatrix = function (data){
		this.XLSCellArray = [];
        
        for (var i= 1; i <= this.XLSRowCount; i++) {

            this.XLSCellArray[i] = [];
            this.XLSCellArray[i][0] = null;   

            for (var j= 1; j <= this.XLSColCount; j++) {
				this.initCellInMemory(i,j);
				
				
                if (data[i-1] == undefined || data[i-1][j-1] == undefined){
					continue;
				}
				
				if(data[i-1][j-1].type){
					if (data[i-1][j-1].type == "none") {
						this.XLSCellArray[i][j].cellType = "input";
						this.XLSCellArray[i][j].readOnly = true;
						var cellStyle = ((i-1)%2) ? this.defaultEvenNonEditableCellStyle : this.defaultOddNonEditableCellStyle;
						this.XLSCellArray[i][j].className = cellStyle + " " + this.defaultTextAlign +"AlignCellClass ";	
					}
					else if (data[i-1][j-1].type == "dropdown") this.XLSCellArray[i][j].cellType = "combo";
					else if (data[i-1][j-1].type == "classToggle") {
						this.XLSCellArray[i][j].cellType = "classToggle";
						this.XLSCellArray[i][j].readOnly = true;
					}
					else this.XLSCellArray[i][j].cellType = data[i-1][j-1].type;
				}
				else this.XLSCellArray[i][j].cellType = "input";
				
				if (data[i-1][j-1].readOnly == true){
					this.XLSCellArray[i][j].readOnly = true;
				}
				
				if(data[i-1][j-1].cssClass)
					this.XLSCellArray[i][j].className = data[i-1][j-1].cssClass + " ";
				
				if(data[i-1][j-1].rowSpan)
					this.XLSCellArray[i][j].rowSpan = data[i-1][j-1].rowSpan;
				
				if(data[i-1][j-1].colSpan)
					this.XLSCellArray[i][j].colSpan = data[i-1][j-1].colSpan;
				
				if(data[i-1][j-1].tag)
					this.XLSCellArray[i][j].tag = data[i-1][j-1].tag;
					
				if(data[i-1][j-1].isVisible != undefined)
					this.XLSCellArray[i][j].isVisible = data[i-1][j-1].isVisible;
				
				if(data[i-1][j-1].options){
					if (data[i-1][j-1].type == "dropdown"){
						for (var k = 0; k < data[i-1][j-1].options.length; k++){
							if (!data[i-1][j-1].options[k].displayLabel)
								data[i-1][j-1].options[k].displayLabel = data[i-1][j-1].options[k].label; 
						}
						this.XLSCellArray[i-1][j-1].valueOptions = {possibleValues : data[i-1][j-1].options, value : data[i-1][j-1].options[0].value, index: 0, subIndex: -1};
					}
					else if (data[i-1][j-1].type == "currencyPair"){
						this.XLSCellArray[i][j].valueOptions = data[i-1][j-1].options;
					}
					else if (data[i-1][j-1].type == "toggle"){
						this.XLSCellArray[i][j].valueOptions = data[i-1][j-1].options;
					}
					else if (data[i-1][j-1].type == "notional"){
						this.XLSCellArray[i][j].valueOptions = data[i-1][j-1].options;
					}
					else if (data[i-1][j-1].type == "delta"){
						this.XLSCellArray[i][j].valueOptions = data[i-1][j-1].options;
					}
					else if (data[i-1][j-1].type == "classToggle"){
						this.XLSCellArray[i][j].valueOptions = data[i-1][j-1].options;
					}
					else if (data[i-1][j-1].type == "radio"){
						this.XLSCellArray[i][j].valueOptions = data[i-1][j-1].options;
					}
					else {
						this.XLSCellArray[i][j].valueOptions = data[i-1][j-1].options;
					}
				}
				
				if(data[i-1][j-1].formatType)
					this.XLSCellArray[i][j].displayType = data[i-1][j-1].formatType.copy();
				
				if (data[i-1][j-1].value){
					this.XLSCellArray[i][j].valueOptions.value = data[i-1][j-1].value;
					this.XLSCellArray[i][j].updateCellByWidget();
				}
				
				
				  
				if (this.XLSCellArray[i][j].cellType == "input" || this.XLSCellArray[i][j].cellType == "notional")
					this.XLSCellArray[i][j].updateCellByString(data[i-1][j-1].value);
				else this.XLSCellArray[i][j].updateCellByWidget();
				
            }
            
        }
	};

	clsDionGrid.prototype.initialize = function(id, options){
		if (!id) {
			return;
		}
		
		this.width = options.width;
		this.height = options.height;
		if (options.hasAutoWidth != undefined){
			this.hasAutoWidth = options.hasAutoWidth;
			this.maxWidth = options.maxWidth;
			this.minWidth = options.minWidth;
		}
		if(options.hasAutoHeight !=undefined)
		{
			this.hasAutoHeight = options.hasAutoHeight;
			this.maxHeight = options.maxHeight;
			this.minHeight = options.minHeight;
		}
		this.top = (options.top) ? options.top : 0;
		this.left = (options.left) ? options.left : 0;
		this.lRowNo = this.XLSRowCount;
		this.lColNo = this.XLSColCount;
		
		
		this.gridName = id;
		this.loneSquareWidth = 40;
        this.loneSquareHeight = 20;
        
        this.defaultEditableCellStyle = (options.defaultEditableCellStyle)? options.defaultEditableCellStyle : "gridCellBorder ";
		this.defaultOddEditableCellStyle = (options.defaultOddEditableCellStyle)? options.defaultOddEditableCellStyle : "gridCellBorder ";
		this.defaultEvenEditableCellStyle = (options.defaultEvenEditableCellStyle)? options.defaultEvenEditableCellStyle : "gridCellBorder ";
		
		this.defaultGridDivStyle = (options.defaultGridDivStyle)? options.defaultGridDivStyle : "gridDiv ";
		
		this.defaultNonEditableCellStyle = (options.defaultNonEditableCellStyle)? options.defaultNonEditableCellStyle : "gridCell ";
		this.defaultOddNonEditableCellStyle = (options.defaultOddNonEditableCellStyle)? options.defaultOddNonEditableCellStyle : "gridCell ";
		this.defaultEvenNonEditableCellStyle = (options.defaultEvenNonEditableCellStyle)? options.defaultEvenNonEditableCellStyle : "gridCell ";
		
		this.defaultCellHeight = (options.defaultCellHeight)? options.defaultCellHeight : 20;
        this.defaultCellWidth = (options.defaultCellWidth)? options.defaultCellWidth : 120;
        this.defaultRowGap = (options.defaultRowGap)? options.defaultRowGap : 0;
        this.defaultColGap = (options.defaultColGap)? options.defaultColGap : 0;
        this.isMouseWheel = (options.isMouseWheel)? options.isMouseWheel : false;
        this.defaultTextAlign = (options.defaultTextAlign)? options.defaultTextAlign : " left";
        
        if (options.hideTopPanel){
			if (options.hideTopPanel == true){
				this.showInterpolationDiv = false;
			}
			else {
				this.showInterpolationDiv = true;
			}
		}
		
		
		this.showContextMenu = (options.showContextMenu) ? options.showContextMenu : false;
		this.showInterpolationDiv = options.showInterpolationDiv ? options.showInterpolationDiv : false;
		this.isCustomContextMenu = options.isCustomContextMenu ? options.isCustomContextMenu : false;

		this.createFullScreen();
		this.createXLSGridInMemory();
		if (options.data){
			var rowCount = options.data.length - 1;
			if (!rowCount || rowCount == 0){
				return;
			}
			var colCount = options.data[0].length - 1;
			if (!colCount || colCount == 0){
				return;
			}
			this.updateXLSCount(rowCount, colCount);
		}
		else {
			this.updateXLSCount(options.rows, options.cols);
		}
		
		
		if (options.data){
			this.setMatrixData(options.data);
		}
		
		if (options.columnWidths){
			for (var i = 0; i < options.columnWidths.length && i < this.XLSColCount; i++){
				this.XLSColArray[i+1].width = options.columnWidths[i]; 
			}
		}
		
		if (options.rowHeights){
			for (var i = 0; i < options.rowHeights.length && i < this.XLSRowCount; i++){
				this.XLSRowArray[i+1].height = options.rowHeights[i]; 
			}
		}
		
		this.refreshSpan()
		this.refreshGridCells();
		this.updateScreenFromXLS();
		this.refreshHeaders();
		workbook.registerGrid(id, this);
	};

	clsDionGrid.prototype.setMatrixData = function (data){
		this.XLSCellArray = [];
        
		for (var i= 1; i <= this.XLSRowCount; i++) {

            this.XLSCellArray[i] = [];
            this.XLSCellArray[i][0] = null;   

            for (var j= 1; j <= this.XLSColCount; j++) {
				this.initCellInMemory(i,j);
				
				
                if (data[i] == undefined || data[i][j] == undefined){
					continue;
				}
				
				if(data[i][j].type){
					if (data[i][j].type == "none") {
						this.XLSCellArray[i][j].cellType = "input";
						this.XLSCellArray[i][j].readOnly = true;
						var cellStyle = (i%2) ? this.defaultEvenNonEditableCellStyle : this.defaultOddNonEditableCellStyle;  
						this.XLSCellArray[i][j].className = cellStyle + " " + this.defaultTextAlign +"AlignCellClass ";	
					}
					else if (data[i][j].type == "dropdown") this.XLSCellArray[i][j].cellType = "combo";
					else if (data[i][j].type == "classToggle") {
						this.XLSCellArray[i][j].cellType = "classToggle";
						this.XLSCellArray[i][j].readOnly = true;
					}
					else this.XLSCellArray[i][j].cellType = data[i][j].type;
				}
				else this.XLSCellArray[i][j].cellType = "input";
				
				if (data[i][j].readOnly == true){
					this.XLSCellArray[i][j].readOnly = true;
				}
				
				if(data[i][j].cssClass)
					this.XLSCellArray[i][j].className = data[i][j].cssClass + " ";
				
				if(data[i][j].rowSpan)
					this.XLSCellArray[i][j].rowSpan = data[i][j].rowSpan;
				
				if(data[i][j].colSpan)
					this.XLSCellArray[i][j].colSpan = data[i][j].colSpan;
				
				if(data[i][j].tag)
					this.XLSCellArray[i][j].tag = data[i][j].tag;
				
				if(data[i][j].name)
					this.nameCell(i,j,data[i][j].name);
				
				if(data[i][j].isVisible != undefined)
					this.XLSCellArray[i][j].isVisible = data[i][j].isVisible;
				
				if(data[i][j].options){
					if (data[i][j].type == "dropdown"){
						for (var k = 0; k < data[i][j].options.length; k++){
							if (!data[i][j].options[k].displayLabel)data[i][j].options[k].displayLabel = data[i][j].options[k].label; 
						}
						this.XLSCellArray[i][j].valueOptions = {possibleValues : data[i][j].options, value : data[i][j].options[0].value, index: 0, subIndex: -1}
					}
					else if (data[i][j].type == "currencyPair"){
						this.XLSCellArray[i][j].valueOptions = data[i][j].options;
					}
					else if (data[i][j].type == "toggle"){
						this.XLSCellArray[i][j].valueOptions = data[i][j].options;
					}
					else if (data[i][j].type == "notional"){
						this.XLSCellArray[i][j].valueOptions = data[i][j].options;
					}
					else if (data[i][j].type == "delta"){
						this.XLSCellArray[i][j].valueOptions = data[i][j].options;
						this.XLSCellArray[i][j].readOnly = true;
					}
					else if (data[i][j].type == "classToggle"){
						this.XLSCellArray[i][j].valueOptions = data[i][j].options;
					}
					else if (data[i][j].type == "radio"){
						this.XLSCellArray[i][j].valueOptions = data[i][j].options;
					}
					else {
						this.XLSCellArray[i][j].valueOptions = data[i][j].options;
					}
				}
				
				if(data[i][j].formatType)
					this.XLSCellArray[i][j].displayType = data[i][j].formatType.copy();
				
				if (data[i][j].value){
					this.XLSCellArray[i][j].valueOptions.value = data[i][j].value;
					this.XLSCellArray[i][j].updateCellByWidget();
				}
				
				
				  
				if (this.XLSCellArray[i][j].cellType == "input" || this.XLSCellArray[i][j].cellType == "notional")
					this.XLSCellArray[i][j].updateCellByString(data[i][j].value);
				else this.XLSCellArray[i][j].updateCellByWidget();
				
            }
            
        }
	};
//---------------------------ReInitialize----------------------------------------------------------------------

	
	clsDionGrid.prototype.resetData = function (data, height, width){
		workbook.clearCellIdForGrid(this);
		
		if (height)this.height = height;
		if (width)this.width = width;
		
		var rowCount = data.length;
		if (!rowCount || rowCount == 0){
			return;
		}
		var colCount = data[1].length;
		if (!colCount || colCount == 0){
			return;
		}
		
		if (this.curXLSColNo > colCount -1) this.curXLSColNo = colCount -1;
		if (this.curXLSRowNo > rowCount -1) this.curXLSRowNo = rowCount -1;
		
		
		this.updateXLSCount(rowCount-1, colCount-1);
		
		for (var i= 1; i <= this.XLSRowCount; i++) {
            this.XLSCellArray[i] = [];
            this.XLSCellArray[i][0] = null;   

            for (var j= 1; j <= this.XLSColCount; j++) {
				this.initCellInMemory(i,j);
				
                if (data[i] == undefined || data[i][j] == undefined){
					continue;
				}
				
				if(data[i][j].type){
					if (data[i][j].type == "none") {
						this.XLSCellArray[i][j].cellType = "input";
						this.XLSCellArray[i][j].readOnly = true;
						var cellStyle = (i%2) ? this.defaultEvenNonEditableCellStyle : this.defaultOddNonEditableCellStyle;  
						this.XLSCellArray[i][j].className = cellStyle + " " + this.defaultTextAlign +"AlignCellClass ";	
					}
					else if (data[i][j].type == "dropdown") this.XLSCellArray[i][j].cellType = "combo";
					else if (data[i][j].type == "classToggle") {
						this.XLSCellArray[i][j].cellType = "classToggle";
						this.XLSCellArray[i][j].readOnly = true;
					}
					else this.XLSCellArray[i][j].cellType = data[i][j].type;
				}
				else this.XLSCellArray[i][j].cellType = "input";
				
				if (data[i][j].readOnly == true){
					this.XLSCellArray[i][j].readOnly = true;
				}
				
				if(data[i][j].cssClass)
					this.XLSCellArray[i][j].className = data[i][j].cssClass + " ";
				
				if(data[i][j].rowSpan)
					this.XLSCellArray[i][j].rowSpan = data[i][j].rowSpan;
				
				if(data[i][j].colSpan)
					this.XLSCellArray[i][j].colSpan = data[i][j].colSpan;
				
				if(data[i][j].tag)
					this.XLSCellArray[i][j].tag = data[i][j].tag;
				
				if(data[i][j].id)
					this.setCellId(i, j, data[i][j].id);
				
				if(data[i][j].name)
					this.nameCell(i,j,data[i][j].name);
				
				if(data[i][j].isVisible != undefined)
					this.XLSCellArray[i][j].isVisible = data[i][j].isVisible;
				
				
				if(data[i][j].options){
					if (data[i][j].type == "dropdown"){
						for (var k = 0; k < data[i][j].options.length; k++){
							if (!data[i][j].options[k].displayLabel)data[i][j].options[k].displayLabel = data[i][j].options[k].label; 
						}
						this.XLSCellArray[i][j].valueOptions = {possibleValues : data[i][j].options, value : data[i][j].options[0].value, index: 0, subIndex: -1}
					}
					else if (data[i][j].type == "currencyPair"){
						this.XLSCellArray[i][j].valueOptions = data[i][j].options;
					}
					else if (data[i][j].type == "toggle"){
						this.XLSCellArray[i][j].valueOptions = data[i][j].options;
					}
					else if (data[i][j].type == "notional"){
						this.XLSCellArray[i][j].valueOptions = data[i][j].options;
					}
					else if (data[i][j].type == "delta"){
						this.XLSCellArray[i][j].valueOptions = data[i][j].options;
						this.XLSCellArray[i][j].readOnly = true;
					}
					else if (data[i][j].type == "classToggle"){
						this.XLSCellArray[i][j].valueOptions = data[i][j].options;
					}
					else if (data[i][j].type == "radio"){
						this.XLSCellArray[i][j].valueOptions = data[i][j].options;
					}
					else {
						this.XLSCellArray[i][j].valueOptions = data[i][j].options;
					}
				}

				if(data[i][j].formatType)
					this.XLSCellArray[i][j].displayType = data[i][j].formatType.copy();
				
				if (data[i][j].value){
					this.XLSCellArray[i][j].valueOptions.value = data[i][j].value;
				}
            }
        }
        
		for (var i= 1; i <= this.XLSRowCount; i++) {
            for (var j= 1; j <= this.XLSColCount; j++) {

				if (this.XLSCellArray[i][j].cellType == "input" || this.XLSCellArray[i][j].cellType == "notional")
					this.XLSCellArray[i][j].updateCellByString(data[i][j].value);
				else this.XLSCellArray[i][j].updateCellByWidget();
            }
		}
		
        for (var j= 1; j <= this.XLSColCount; j++) {
	       	 if (data[1] == undefined || data[1][j] == undefined){
						continue;
				 }
	       	 if (data[1][j].width) {
	       		 this.XLSColArray[j].width = data[1][j].width;
	       	 }
	       	 
       }
	   this.refreshSpan();
	   this.refreshHeaders();
	   this.refreshGridCells();
	   this.updateScreenFromXLS();
	   this.goToCell(this.curXLSRowNo, this.curXLSColNo);
	};


//------------------------Insert/Delete Rows or Columns -------------------------------------------------------	
	
	clsDionGrid.prototype.insertRows = function (rowIndex, count){
		if (count < 0) return;
		if (rowIndex < 1 || rowIndex > this.XLSRowCount) return;
		if (count == 1) 
			this.insertRow (rowIndex);
		else
			this.insertMultipleRow(rowIndex, count);
	};
	
	clsDionGrid.prototype.insertCols = function (colIndex, count){
		if (count < 0) return;
		if (colIndex < 1 || colIndex > this.XLSColCount) return;
		if (count == 1) 
			this.insertCol (colIndex);
		else
			this.insertMultipleCol(colIndex, count);
	};
	
	clsDionGrid.prototype.deleteRows = function (rowIndex, count){
		if (count < 0) return;
		else if (count == 1) 
			this.deleteRow (rowIndex);
		else
			this.deleteMultipleRow(rowIndex, count);
	};
	
	clsDionGrid.prototype.deleteCols = function (colIndex, count){
		if (count < 0) return;
		else if (count == 1) 
			this.deleteCol (colIndex);
		else
			this.deleteMultipleCol(colIndex, count);
	};
//------------------------Make a Range Editable/Non Editable -----------------------------------------------------
	
	//make editable and make non editable takes range as A1 or A1:B10 or 1:1:2:9
	clsDionGrid.prototype.makeEditable = function (range){
		if (typeof(range) == "string"){
			if (isRange(range)){
				var rangeArray = range.split(":");
				var rangeStart = getCellRowCol(rangeArray[0].toUpperCase());	
				var rangeEnd = getCellRowCol(rangeArray[1].toUpperCase());
				var rangeStartY = getColNoFromLabel(rangeStart[0]);	
				var rangeEndY = getColNoFromLabel(rangeEnd[0]);
				var rangeStartX = parseInt(rangeStart[1]);
				var rangeEndX = parseInt(rangeEnd[1]);
				var numRange = "" + rangeStartX + ":" + rangeStartY + ":" + rangeEndX + ":" + rangeEndY;
				this.setReadonly(numRange, false);
			
			}
			else if (isNumRange(range)){
				this.setReadonly(range, false);
			}
			else if (isCell(range)){
				var cellNo = getCellRowCol(range.toUpperCase());
				var rangeX = getColNoFromLabel(cellNo[0]);
				var rangeY = parseInt(cellNo[1]);
				var numRange = "" + rangeX + ":" + rangeY + ":" + rangeX + ":" + rangeY;
				this.setReadonly(numRange, false);
					
			}
		}
	};
	
	clsDionGrid.prototype.makeNonEditable = function (range){
		if (typeof(range) == "string"){
			if (isRange(range)){
				var rangeArray = range.split(":");
				var rangeStart = getCellRowCol(rangeArray[0].toUpperCase());	
				var rangeEnd = getCellRowCol(rangeArray[1].toUpperCase());
				var rangeStartY = getColNoFromLabel(rangeStart[0]);	
				var rangeEndY = getColNoFromLabel(rangeEnd[0]);
				var rangeStartX = parseInt(rangeStart[1]);
				var rangeEndX = parseInt(rangeEnd[1]);
				var numRange = "" + rangeStartX + ":" + rangeStartY + ":" + rangeEndX + ":" + rangeEndY;
				this.setReadonly(numRange, true);
			}
			else if (isNumRange(range)){
				this.setReadonly(range, true);
			}
			else if (isCell(range)){
				var cellNo = getCellRowCol(range.toUpperCase());
				var rangeX = getColNoFromLabel(cellNo[0]);
				var rangeY = parseInt(cellNo[1]);
				var numRange = "" + rangeX + ":" + rangeY + ":" + rangeX + ":" + rangeY;
				this.setReadonly(numRange, true);
					
			}
		}
	};

	//SetReadonly takes numRange ie 1:2:7:2
	clsDionGrid.prototype.setReadonly = function (range, val){
		//Get the indices from the range
		var rangeArray = range.split(":");
		if (rangeArray.length != 4) return false;
		
		var startX = parseInt(rangeArray[0]);
		var startY = parseInt(rangeArray[1]);
		var endX = parseInt(rangeArray[2]);
		var endY = parseInt(rangeArray[3]);
		
		if (startX > endX){
			var temp = endX;
			endX = startX;
			startX = temp;
		}
		if (startY > endY){
			var temp = endY;
			endY = startY;
			startY = temp;
		}
		
		var countX = endX-startX+1;
		var countY = endY-startY+1;
		
		for (var i = startX; i <= endX && i <= this.XLSRowCount; i++){
			for (var j = startY; j <= endY && j <= this.XLSColCount; j++){
				this.XLSCellArray[i][j].readOnly = val;
				if (val == true){
					var cellStyle = (i%2) ? this.defaultEvenNonEditableCellStyle : this.defaultOddNonEditableCellStyle;
					this.XLSCellArray[i][j].className = cellStyle;
				}
				else {
					var cellStyle = (i%2) ? this.defaultEvenEditableCellStyle : this.defaultOddEditableCellStyle;
					this.XLSCellArray[i][j].className = cellStyle;
				}
				//this.refreshCell(i,j);
			}
		}
		this.updateScreenFromXLS();	
	};
	
//------------------------Set Row/Col position/Length------------------------------------------------------------	
	//Set the width of a column
	//It just has to set the column width in xls array and call refresh grid cells
	clsDionGrid.prototype.setColumnWidth = function(columnNumber, width) {
		this.XLSColArray[columnNumber].width = width;
		this.refreshGridCells();
		this.refreshHeaders();
	};
	
	//Set the Height of a row
	clsDionGrid.prototype.setRowHeight = function(rowNumber, height) {
		this.XLSRowArray[rowNumber].height = height;
		this.refreshGridCells();
		this.refreshHeaders();
	};
	
	clsDionGrid.prototype.setAllRowheight = function(height){
		for (var i = 0; i < this.XLSRowCount; i++){
			this.XLSRowArray[i].height = height;
		}
		this.refreshGridCells();
		this.refreshHeaders();
	};
	
	clsDionGrid.prototype.setAllColWidth = function(width){
		for (var i = 0; i < this.XLSColCount; i++){
			this.XLSColArray[i].width = width;
		}
		this.refreshGridCells();
		this.refreshHeaders();
	};
	
	clsDionGrid.prototype.setRowGap = function (rowNo, gap){
		this.XLSRowArray[rowNo].gap = gap;
		this.refreshGridCells();
		this.refreshHeaders();
	};
	
	clsDionGrid.prototype.setColGap = function (colNo, gap){
		this.XLSColArray[colNo].gap = gap;
		this.refreshGridCells();
		this.refreshHeaders();
	};
	
	clsDionGrid.prototype.setAllRowGap = function (gap){
		for (var i = 1; i <= this.XLSRowCount; i++){
			this.XLSRowArray[i].gap = gap;
		}
		this.refreshGridCells();
		this.refreshHeaders();
	};
	
	clsDionGrid.prototype.setAllColGap = function (gap){
		for (var i = 1; i <= this.XLSColCount; i++){
			this.XLSColArray[i].gap = gap;
		}
		this.refreshGridCells();
		this.refreshHeaders();
	};

	
//------------------------Insert Various Widgets -------------------------------------------------------------

	
	clsDionGrid.prototype.insertWidget = function (widgetName, i,j, item){
			var displayCellArrayDiv = this.displayCellArray[i][j];
			displayCellArrayDiv = removeAllChildren(displayCellArrayDiv);
			var editor = FinsheetEditors.getEditor(widgetName);
			if (!editor) return;
			
			var comboEditor = new editor({
				container: displayCellArrayDiv,
				rowNo: i,
				colNo: j,
				grid: this,
				item: item});
			
			displayCellArrayDiv.cellType = widgetName;
			displayCellArrayDiv.editor = comboEditor;
		if (this.isInXLSRange(i, j)){
			//this.XLSCellArray[i][j].cellType = widgetName;
			if (item != undefined) this.XLSCellArray[i][j].valueOptions = item.valueOptions;
		}
		
	};
	
	clsDionGrid.prototype.insertFormatter = function (widgetName, i,j, item){
		var displayCellArrayDiv = this.displayCellArray[i][j];
		displayCellArrayDiv = removeAllChildren(displayCellArrayDiv);
		var formatter = FinsheetEditors.getFormatter(widgetName);
		
		if (!formatter) return;
		
		var formatterObj = new formatter({
			container: displayCellArrayDiv,
			rowNo: i,
			colNo: j,
			grid: this,
			item: item});
		displayCellArrayDiv.cellType = widgetName;
		displayCellArrayDiv.editor = null;
		
		if (this.isInXLSRange(i, j)){
			//this.XLSCellArray[i][j].cellType = widgetName;
			if (item) this.XLSCellArray[i][j].valueOptions = item.valueOptions;
		}
		
	};
	
//------------------------Name Cell/Range --------------------------------------------------------------------
	
	clsDionGrid.prototype.nameCell = function (i,j, name){
		if (!this.isInXLSRange(i,j)) return;
		var rangeString = getColLabelFromNumber(j) + i;
		this.XLSCellArray[i][j].name = name;
		workbook.nameRange(name,this.gridName, rangeString);
	};
	
	clsDionGrid.prototype.nameRange = function (rangeString, name){
		if (isCell(rangeString) || isRange(rangeString))
			workbook.nameRange(name,this.gridName, rangeString);
		else if (isNumRange(rangeString)){
			var range = getRange(rangeString);
			workbook.nameRange(name,this.gridName, range);
		}
	};
	
	clsDionGrid.prototype.getName = function (i,j){
		if (!this.isInXLSRange(i,j)) return;
		return this.XLSCellArray[i][j].name;
	};
//------------------------Cut Copy Paste--------------------------------------------------------------------
	
	clsDionGrid.prototype.copyRow = function (rowNo){
		if (this.isInXLSRange(rowNo, 1)){
			var range = "" + rowNo + ":" + 1 + ":" + rowNo + ":" + this.XLSColCount; 
			this.copyRange(range);
		}
	};
	
	clsDionGrid.prototype.copyRows = function (startRowNo, endRowNo){
		var startRow = Math.min(startRowNo, endRowNo);
		var endRow = Math.max(startRowNo, endRowNo);
		if (this.isInXLSRange(startRow, 1) && this.isInXLSRange(endRow, 1)){
			var range = "" + startRow + ":" + 1 + ":" + endRow + ":" + this.XLSColCount; 
			this.copyRange(range);
		}
	};
	
	clsDionGrid.prototype.copyCol = function (colNo){
		if (this.isInXLSRange(1, colNo)){
			var range = "1:" + colNo + ":" + this.XLSRowCount + ":" + colNo;
			this.copyRange(range);
		}
	};
	
	clsDionGrid.prototype.copyCell = function (rowNo, colNo){
		if (this.isInXLSRange(rowNo, colNo)){
			var range = rowNo + ":" + colNo + ":" + rowNo + ":" + colNo;
			this.copyRange(range);
		}	
	};
	clsDionGrid.prototype.copySelectedCells = function (){
		this.copyRange(workbook.selectedRange);
	};
	
	clsDionGrid.prototype.paste = function (rowNo, colNo){
		if (this.isInXLSRange(rowNo, colNo)){
			this.pasteSpecialCopiedCells(rowNo, colNo);
		}
	};
	
	clsDionGrid.prototype.pasteValues = function (rowNo, colNo){
		if (this.isInXLSRange(rowNo, colNo)){
			this.pasteCopiedCells(rowNo, colNo);
		}
	};
	
		
//------------------------Format Functions--------------------------------------------------------------------
	clsDionGrid.prototype.setClassName = function (range, value){
		var numRange = getNumRange(range);
		if (numRange != null){
			var THIS = this;
			var cellFn = function (rowNo, colNo){
				THIS.XLSCellArray[rowNo][colNo].className = value;
			};
			this.mapFnToRange(numRange, cellFn);
		}
	};
	
	clsDionGrid.prototype.setCellClassName = function (i,j, value){
		this.XLSCellArray[i][j].className = value;
	};
	
	clsDionGrid.prototype.addCellClassName = function (i,j, value){
	    if (!this.isInXLSRange(i,j)) return;
	    var XLSCellArrayObj = this.XLSCellArray[i][j];
	    
	    if (value && this.XLSCellArray[i][j].isVisible == true){
		    XLSCellArrayObj.className = addClassName(XLSCellArrayObj.className, " " + value);
	    }
	    this.refreshCell(i,j);
	};
	
	clsDionGrid.prototype.setDisplayFormat = function (rowNo, colNo, formatType){
		if (this.isInXLSRange(rowNo, colNo)){
			var XLSCellArrayObj = this.XLSCellArray[rowNo][colNo];
			if (formatType){
				XLSCellArrayObj.displayType = formatType.copy();
				XLSCellArrayObj.displayValue = convertTo(XLSCellArrayObj.value, XLSCellArrayObj.type, XLSCellArrayObj.displayType);
				if (XLSCellArrayObj.valueOptions) XLSCellArrayObj.valueOptions.displayValue = XLSCellArrayObj.displayValue;
			}
			else {
				XLSCellArrayObj.displayType = null;
				XLSCellArrayObj.displayValue = XLSCellArrayObj.value;
				if (XLSCellArrayObj.valueOptions) XLSCellArrayObj.valueOptions.displayValue = XLSCellArrayObj.displayValue;
			}
			
			this.refreshCell(rowNo,colNo);
		}
	};
	
	clsDionGrid.prototype.changeDisplayFormat = function (range, formatType){
		if (isCell(range)){
			var mArrColRow = getCellRowCol(range.toUpperCase());
			var rowNo = parseInt(mArrColRow[1]);
			var colNo = getColNoFromLabel(mArrColRow[0]);
			if (this.isInXLSRange(rowNo, colNo)){
				var XLSCellArrayObj = this.XLSCellArray[rowNo][colNo];
				if (formatType){
					XLSCellArrayObj.displayType = formatType.copy();
					XLSCellArrayObj.displayValue = convertTo(XLSCellArrayObj.value, XLSCellArrayObj.type, XLSCellArrayObj.displayType);
				}
				else {
					XLSCellArrayObj.displayType = null;
					XLSCellArrayObj.displayValue = XLSCellArrayObj.value;
				}
				
				if (XLSCellArrayObj.cellType != "input") 
					return;
				
				this.refreshCell(rowNo,colNo);
			}
		}
		else if (isRange(range)){
			var rangeArray = range.split(":");
			var rangeStart = getCellRowCol(rangeArray[0].toUpperCase());	
			var rangeEnd = getCellRowCol(rangeArray[1].toUpperCase());
			var rangeStartY = getColNoFromLabel(rangeStart[0]);	
			var rangeEndY = getColNoFromLabel(rangeEnd[0]);
			var rangeStartX = parseInt(rangeStart[1]);
			var rangeEndX = parseInt(rangeEnd[1]);
			for (var i = rangeStartX; i <= rangeEndX; i++){
				for (var j = rangeStartY; j <= rangeEndY; j++){
					var XLSCellArrayObj = this.XLSCellArray[i][j];
					if (formatType){
						XLSCellArrayObj.displayType = formatType.copy();
						XLSCellArrayObj.displayValue = convertTo(XLSCellArrayObj.value, XLSCellArrayObj.type, XLSCellArrayObj.displayType);
					}
					else {
						XLSCellArrayObj.displayType = null;
						XLSCellArrayObj.displayValue = XLSCellArrayObj.value;
					}
					if (XLSCellArrayObj.cellType != "input") 
						return;
					this.refreshCell(i,j);
				}
			}
		}
		else if (isNumRange(range)){
			var rangeArray = range.split(":");
			var rangeStartY = parseInt(rangeArray[1]);
			var rangeEndY = parseInt(rangeArray[3]);
			var rangeStartX = parseInt(rangeArray[0]);
			var rangeEndX = parseInt(rangeArray[2]);
			for (var i = rangeStartX; i <= rangeEndX; i++){
				for (var j = rangeStartY; j <= rangeEndY; j++){
					var XLSCellArrayObj = this.XLSCellArray[i][j];
					if (formatType){
						XLSCellArrayObj.displayType = formatType.copy();
						XLSCellArrayObj.displayValue = convertTo(XLSCellArrayObj.value, XLSCellArrayObj.type, XLSCellArrayObj.displayType);
					}
					else {
						XLSCellArrayObj.displayType = null;
						XLSCellArrayObj.displayValue = XLSCellArrayObj.value;
					}
					if (XLSCellArrayObj.cellType != "input") 
						continue;
					this.refreshCell(i,j);
				}
			}
		}
	};

	clsDionGrid.prototype.leftAlign = function (range){
		var numRange = getNumRange(range);
		if (numRange != null)
			this.alignRangeFn(numRange, "left");
	};
	
	clsDionGrid.prototype.rightAlign = function (range){
		var numRange = getNumRange(range);
		if (numRange != null)
			this.alignRangeFn(numRange, "right");
	};
	
	clsDionGrid.prototype.centerAlign = function (range){
		var numRange = getNumRange(range);
		if (numRange != null)
			this.alignRangeFn(numRange, "center");
	};
	
	clsDionGrid.prototype.bold = function (range){
		var numRange = getNumRange(range);
		if (numRange != null){
			var THIS = this;
			var cellFn = function (rowNo, colNo){
					THIS.XLSCellArray[rowNo][colNo].className = addClassName(THIS.XLSCellArray[rowNo][colNo].className, "boldFontCellClass ");
			};
			this.mapFnToRange(numRange, cellFn);
		}
	};
	
	clsDionGrid.prototype.italics = function (range){
		var numRange = getNumRange(range);
		var THIS = this;
		if (numRange != null){
			var cellFn = function (rowNo, colNo){
				THIS.XLSCellArray[rowNo][colNo].className = addClassName(THIS.XLSCellArray[rowNo][colNo].className, "italicStyleCellClass ");
			};
			this.mapFnToRange(numRange, cellFn);
		}
	};
	
	clsDionGrid.prototype.underline = function (range){
		var numRange = getNumRange(range);
		var THIS = this;
		if (numRange != null){
			var cellFn = function (rowNo, colNo){
				THIS.XLSCellArray[rowNo][colNo].className = addClassName(THIS.XLSCellArray[rowNo][colNo].className, "underlineDecorationCellClass ");
			};
			this.mapFnToRange(numRange, cellFn);
		}
	};

	clsDionGrid.prototype.removeBold = function (range){
		var numRange = getNumRange(range);
		var THIS = this;
		if (numRange != null){
			var cellFn = function (rowNo, colNo){
				THIS.XLSCellArray[rowNo][colNo].className = removeClassName(THIS.XLSCellArray[rowNo][colNo].className, "boldFontCellClass ");
			};
			this.mapFnToRange(numRange, cellFn);
		}
	};
	
	clsDionGrid.prototype.removeItalics = function (range){
		var numRange = getNumRange(range);
		var THIS = this;
		if (numRange != null){
			var cellFn = function (rowNo, colNo){
				THIS.XLSCellArray[rowNo][colNo].className = removeClassName(THIS.XLSCellArray[rowNo][colNo].className, "italicStyleCellClass ");		
			};
			this.mapFnToRange(numRange, cellFn);
		}
	};
	
	clsDionGrid.prototype.removeUnderline = function (range){
		var numRange = getNumRange(range);
		var THIS = this;
		if (numRange != null){
			var cellFn = function (rowNo, colNo){
				THIS.XLSCellArray[rowNo][colNo].className = removeClassName(THIS.XLSCellArray[rowNo][colNo].className, "underlineDecorationCellClass ");
			};
			this.mapFnToRange(numRange, cellFn);
		}
	};
	
	//The values can be top, left, right, bottom, all, none, topDashed, rightDashed, leftDashed, rightDashed, noneDashed, allDashed
	clsDionGrid.prototype.border = function(range, value){
		var numRange = getNumRange(range);
		if (numRange != null){
			this.borderMapper (numRange, value);
		}
	};
	
	clsDionGrid.prototype.setRowSpan = function (row, col, value){
		
		if (this.isInXLSRange(row, col)){
			this.XLSCellArray[row][col].rowSpan = value;
		}
		this.refreshSpan();
		this.refreshGridCells();
	}
	
	clsDionGrid.prototype.setColSpan = function (row, col, value){
		
		if (this.isInXLSRange(row, col)){
			this.XLSCellArray[row][col].colSpan = value;
		}
		this.refreshSpan();
		this.refreshGridCells();
	}
	
	clsDionGrid.prototype.setRowDeletable = function (row, value){
		if (this.isInXLSRange(row, 1)){
				this.XLSRowArray[row].isDeletable = value;
		}
	}
	
	clsDionGrid.prototype.setColDeletable = function (row, value){
		if (this.isInXLSRange(1, col)){
				this.XLSColArray[col].isDeletable = value;
		}
	}
	
	clsDionGrid.prototype.setRightGrid = function (grid, row_offset, col_offset){
		if (row_offset < 0) row_offset = 0;
		if(col_offset < 0) col_offset = 0;
		this.nextGrid.right = {grid : grid, row_offset : row_offset, col_offset:col_offset};
	}
	clsDionGrid.prototype.setLeftGrid = function (grid, row_offset, col_offset){
		if (row_offset < 0) row_offset = 0;
		if(col_offset < 0) col_offset = 0;
		this.nextGrid.left = {grid : grid, row_offset : row_offset, col_offset:col_offset};
	}
	clsDionGrid.prototype.setUpGrid = function (grid, row_offset, col_offset){
		if (row_offset < 0) row_offset = 0;
		if(col_offset < 0) col_offset = 0;
		this.nextGrid.up = {grid : grid, row_offset : row_offset, col_offset:col_offset};
	}
	clsDionGrid.prototype.setDownGrid = function (grid, row_offset, col_offset){
		if (row_offset < 0) row_offset = 0;
		if(col_offset < 0) col_offset = 0;
		this.nextGrid.down = {grid : grid, row_offset : row_offset, col_offset:col_offset};
	}
	

}

defineGridAPI();
