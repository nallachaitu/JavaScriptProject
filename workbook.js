/*
 *   copyright (c) -- chaitanya nalla and kechit goyal
 *
 */

var workbook = {
	
	grids : new Map(),
	nameMap : new Map(),
	copiedCellArray : [],
	copiedRange : "",
	curDionGrid : null,
	curXLSRowNo : 0,
	curXLSColNo: 0,
	formulaRange : "",
	contextMenuDiv : null,
	selectedRange : "0:0:0:0",
	cellMap : {},
		
	registerGrid : function (key, object){
		workbook.grids.put (key, object);
	},
	
	getGridById : function (key){
		return workbook.grids.get(key);
	},
	
	nameRange : function (name, gridId, range ){
		workbook.nameMap.put (name, gridId + "." + range);
	},
	
	clearRange : function(range){
		
		if(range == "0:0:0:0" || range == undefined || range == null)
			return;
		
		var rangeArray = range.split(':');
		var startRow = parseInt(rangeArray[0]);
		var endRow = parseInt(rangeArray[2]);
		var startCol = parseInt(rangeArray[1]);
		var endCol = parseInt(rangeArray[3]);
			var GRID = workbook.curDionGrid;
			if(!GRID) return;
			
		var temp;	
		
		if(startRow > endRow ){
			temp = startRow;
			startRow = endRow;
			endRow = temp;
		}	
			
		if(startCol > endCol){
			temp = startCol;
			startCol = endCol;
			endCol = temp;
		
		}
			
		var displayCellArrayObj ;
		for(var i=startRow;i <= endRow && i <= GRID.XLSRowCount;i++){
			for(var j=startCol;j <= endCol && j <= GRID.XLSColCount;j++){
				displayCellArrayObj = GRID.displayCellArray[i][j]; 
				displayCellArrayObj.className = removeClassName(displayCellArrayObj.className, " " + "focusRange");
			}
		}
	},
	
	
	stopKeyCapture : function (){
	    if (workbook.curDionGrid) workbook.curDionGrid.mode = "none";
	},
	
	getCellByName : function (name){
		var cellRange = workbook.nameMap.get (name);
		if (!cellRange) return;
		var nameArray = cellRange.split(".");
		if (nameArray.length != 2) return;
		
		var grid = workbook.grids.get (nameArray[0]);
		if (!grid) return;
		
		if (isCell(nameArray[1])){
			var mArrColRow = getCellRowCol(nameArray[1].toUpperCase());
			var rowNo = parseInt(mArrColRow[1]);
			var colNo = getColNoFromLabel(mArrColRow[0]);
			if (grid.isInXLSRange(rowNo, colNo)){
				return grid.XLSCellArray[rowNo][colNo].value;
			}
		}
		return;
	},
	
	getCellRowColByName : function (name){
		var cellRange = workbook.nameMap.get (name);
		if (!cellRange) return;
		var nameArray = cellRange.split(".");
		if (nameArray.length != 2) return;
		
		var grid = workbook.grids.get (nameArray[0]);
		if (!grid) return;
		
		if (isCell(nameArray[1])){
			var mArrColRow = getCellRowCol(nameArray[1].toUpperCase());
			var rowNo = parseInt(mArrColRow[1]);
			var colNo = getColNoFromLabel(mArrColRow[0]);
			return {row : rowNo, col : colNo};
		}
		return;
	},
	
	getCellValueArrayByRange : function (name){
		var cellRange = workbook.nameMap.get (name);
		if (!cellRange) return;
		var nameArray = cellRange.split(".");
		if (nameArray.length != 2) return;
		
		var grid = workbook.grids.get (nameArray[0]);
		if (!grid) return;
		
		
		if (isRange(nameArray[1])){
			var temp = [];
			var mArrColRow = getCellRowCol(nameArray[1].split(":")[0].toUpperCase());
			var startRowNo = parseInt(mArrColRow[1]);
			var startColNo = getColNoFromLabel(mArrColRow[0]);
			
			mArrColRow = getCellRowCol(nameArray[1].split(":")[1].toUpperCase());
			var endRowNo = parseInt(mArrColRow[1]);
			var endColNo = getColNoFromLabel(mArrColRow[0]);
			
			for (var i = startRowNo; i <= endRowNo; i++){
				temp[i-startRowNo] = [];
				for (var j = startColNo; j <= endColNo; j++){
					if (grid.isInLRange(i,j)){
						temp[i-startRowNo][j-startColNo] = grid.XLSCellArray[i][j].value;
					}
				}
			}
			return temp;
		}
		else if (isCell(nameArray[1])){
			var mArrColRow = getCellRowCol(nameArray[1].toUpperCase());
			var rowNo = parseInt(mArrColRow[1]);
			var colNo = getColNoFromLabel(mArrColRow[0]);
			if (grid.isInXLSRange(rowNo, colNo)){
				var temp = []
				temp[0] = grid.XLSCellArray[rowNo][colNo].value;
				return temp;
			}
		}
		return;
	},
	
	getRangeNoByName : function (name){
		var cellRange = workbook.nameMap.get (name);
		if (!cellRange) return;
		var nameArray = cellRange.split(".");
		if (nameArray.length != 2) return;
		
		var grid = workbook.grids.get (nameArray[0]);
		if (!grid) return;
		
		
		if (isRange(nameArray[1])){
			var mArrColRow = getCellRowCol(nameArray[1].split(":")[0].toUpperCase());
			var startRowNo = parseInt(mArrColRow[1]);
			var startColNo = getColNoFromLabel(mArrColRow[0]);
			
			mArrColRow = getCellRowCol(nameArray[1].split(":")[1].toUpperCase());
			var endRowNo = parseInt(mArrColRow[1]);
			var endColNo = getColNoFromLabel(mArrColRow[0]);
			
			return {startRow : startRowNo, startCol : startColNo, endRow : endRowNo, endCol : endColNo} ;
		}
		else if (isCell(nameArray[1])){
			var mArrColRow = getCellRowCol(nameArray[1].toUpperCase());
			var rowNo = parseInt(mArrColRow[1]);
			var colNo = getColNoFromLabel(mArrColRow[0]);
			return {startRow : rowNo, startCol : colNo, endRow : rowNo, endCol : colNo};
		}
		return;
	},
	
	setValueByName : function (name, value){
		workbook.nameMap.get (name);
		var nameArray = name.split(".");
		if (nameArray.length != 2) return;
		
		var grid = workbook.getGridById(nameArray[0]);
		if (!grid) return;
		
		if (isCell(nameArray[1])){
			var mArrColRow = getCellRowCol(range.toUpperCase());
			var rowNo = parseInt(mArrColRow[1]);
			var colNo = getColNoFromLabel(mArrColRow[0]);
			if (grid.isInXLSRange(rowNo, colNo)){
				grid.setValue(rowNo,colNo,value);
			}
		}
		return null;
	},
	
	setCopiedRange : function (gridId, range){
		if (!isNumRange(range)) return;
		if (workbook.getGridById(gridId)){
			workbook.copiedRange = gridId + "." + range;
			//workbook.copiedCellArray = copiedCellArray;
		}
	},
	
	clearCopiedRange : function(){
		workbook.copiedRange = "";
		workbook.copiedCellArray = [];
	},
	
	getCellByString : function(myStr){
		if (!isGridCell(myStr)) return null;
		var mArray = myStr.split(".");
		var grid = workbook.getGridById(mArray[0]);
		if(!grid) return null;
		
		var mArrColRow = getCellRowCol(mArray[1].toUpperCase());
		var rowNo = parseInt(mArrColRow[1]);
		var colNo = getColNoFromLabel(mArrColRow[0]);
		
		return (grid.XLSCellArray[rowNo][colNo]);
	},
	
	isCurGrid : function (grid){
		var gridId = grid.gridName;
		if (workbook.curDionGrid != null)
			return (workbook.curDionGrid.gridName == gridId);
		return false;
	},
	
	setCellId : function (cellId, grid, cellRowCol){
		if (!cellId || !grid) return;
		
		var gridId = grid.gridName;
		if (!workbook.cellMap[gridId]) workbook.cellMap[gridId] = {};
		workbook.cellMap[gridId][cellId] = cellRowCol;
	},
	
	clearCellIdForGrid : function (grid){
		var gridId = grid.gridName;
		workbook.cellMap[gridId] = {};
	},
	
	getCellIdValueMapForGrid : function (grid){
		var gridId = grid.gridName;
		var idValMap = {};
		for (var i in workbook.cellMap[gridId]){
			if (typeof (workbook.cellMap[gridId][i].value) == typeof (0.0))
				idValMap[i] = workbook.cellMap[gridId][i].value;
		}
		return idValMap;
	},
	
	getCellById : function (cellId, grid){
		if (!cellId) return;
		
		cellId = "" + cellId;
		var mArray = cellId.split(".");
		
		var gridId = "";
		var cell = "";
		if (mArray.length == 1){
			gridId = grid.gridName;
			cell = cellId;
		}
		else if (mArray.length == 2){
			gridId = mArray[0];
			cellId = mArray[1];
		}
		else return null;
		
		if (!workbook.cellMap[gridId]) return null;
		return workbook.cellMap[gridId][cellId];
	}
};
