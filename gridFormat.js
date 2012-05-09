/*
 *   copyright (c) -- chaitanya nalla and kechit goyal
 *
 */

function defineGridFormat(){
	
	clsDionGrid.prototype.mapFnToRange = function(range, fn, arg){
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
		
		for (var i = startX; i <= endX; i++){
			for(var j = startY; j <= endY; j++){
				fn(i,j);
			}
		}
		this.updateScreenFromXLS();
	};

	clsDionGrid.prototype.boldRangeFn = function (range){
		var THIS = this;
		var boldCellFn = function (rowNo, colNo){
				THIS.XLSCellArray[rowNo][colNo].className = toggleClassName(THIS.XLSCellArray[rowNo][colNo].className, "boldFontCellClass ");
		};
		this.mapFnToRange(range, boldCellFn);
	};
	
	clsDionGrid.prototype.italicRangeFn = function (range){
		var THIS = this;
		var italicCellFn = function (rowNo, colNo){
				THIS.XLSCellArray[rowNo][colNo].className = toggleClassName(THIS.XLSCellArray[rowNo][colNo].className, "italicStyleCellClass ");
		};
		this.mapFnToRange(range, italicCellFn);
	};
	
	clsDionGrid.prototype.underlineRangeFn = function (range){
		var THIS = this;
		var underlineCellFn = function (rowNo, colNo){
				THIS.XLSCellArray[rowNo][colNo].className = toggleClassName(THIS.XLSCellArray[rowNo][colNo].className, "underlineDecorationCellClass ");
		};
		this.mapFnToRange(range, underlineCellFn);
	};
	
	clsDionGrid.prototype.alignRangeFn = function (range, value){
		var THIS = this;
		var helperFn = function (rowNo, colNo){
			THIS.XLSCellArray[rowNo][colNo].className = removeClassName(THIS.XLSCellArray[rowNo][colNo].className, THIS.XLSCellArray[rowNo][colNo].align + "AlignCellClass");
			THIS.XLSCellArray[rowNo][colNo].align= value;
			THIS.XLSCellArray[rowNo][colNo].className += (value + "AlignCellClass ");
		};
		this.mapFnToRange(range, helperFn);
	};
	
	clsDionGrid.prototype.borderMapper = function(range, value){
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
		
		if (value == 'top'){
			for (var j = startY; j <= endY; j++){
				this.XLSCellArray[startX][j].className = removeClassName (this.XLSCellArray[startX][j].className,"borderDashedTopCellClass " );
				this.XLSCellArray[startX][j].className = addClassName (this.XLSCellArray[startX][j].className,"borderTopCellClass " );
			}
		}

		else if (value == 'bottom'){
			for (var j = startY; j <= endY; j++){
				if (endX != this.XLSRowCount){
					this.XLSCellArray[endX+1][j].className = removeClassName(this.XLSCellArray[endX+1][j].className, "borderDashedTopCellClass");
					this.XLSCellArray[endX+1][j].className = addClassName (this.XLSCellArray[endX+1][j].className,"borderTopCellClass " );
				}
				else {
					this.XLSCellArray[endX][j].className = removeClassName(this.XLSCellArray[endX][j].className, "borderDashedBottomCellClass");
					this.XLSCellArray[endX][j].className = addClassName (this.XLSCellArray[endX][j].className,"borderBottomCellClass " );
				}
			}
		}
		
		else if (value == 'right'){
			for (var i = startX; i <= endX; i++){
				if (endY != this.XLSColCount){
					this.XLSCellArray[i][endY+1].className = removeClassName(this.XLSCellArray[i][endY+1].className, "borderDashedLeftCellClass");
					if (this.XLSCellArray[i][endY+1].className.indexOf("borderLeftCellClass ") == -1)this.XLSCellArray[i][endY+1].className += "borderLeftCellClass ";
				}
				else {
					this.XLSCellArray[i][endY].className = removeClassName(this.XLSCellArray[i][endY].className, "borderDashedRightCellClass");
					if (this.XLSCellArray[i][endY].className.indexOf("borderRightCellClass ") == -1)this.XLSCellArray[i][endY].className += "borderRightCellClass ";
				}
			}
		}
		
		else if (value == 'left'){
			for (var i = startX; i <= endX; i++){
				this.XLSCellArray[i][startY].className = removeClassName(this.XLSCellArray[i][startY].className, "borderDashedLeftCellClass ");
				if (this.XLSCellArray[i][startY].className.indexOf("borderLeftCellClass ") == -1)this.XLSCellArray[i][startY].className += "borderLeftCellClass ";
			}
		}
		
		else if (value == 'all'){
			this.borderMapper(range, "right");
			this.borderMapper(range, "left");
			this.borderMapper(range, "bottom");
			this.borderMapper(range, "top");
		}
		
		else if (value == 'topDashed'){
			for (var j = startY; j <= endY; j++){
				if (this.XLSCellArray[startX][j].className.indexOf("borderDashedTopCellClass ") == -1)this.XLSCellArray[startX][j].className += "borderDashedTopCellClass ";
			}
		}

		else if (value == 'bottomDashed'){
			for (var j = startY; j <= endY; j++){
				if (endX != this.XLSRowCount){
					if (this.XLSCellArray[endX+1][j].className.indexOf("borderDashedTopCellClass ") == -1)this.XLSCellArray[endX+1][j].className += "borderDashedTopCellClass ";
				}
				else {
					if (this.XLSCellArray[endX][j].className.indexOf("borderDashedBottomCellClass ") == -1)this.XLSCellArray[endX][j].className += "borderDashedBottomCellClass ";
				}
			}
		}
		
		else if (value == 'rightDashed'){
			for (var i = startX; i <= endX; i++){
				if (endY != this.XLSColCount){
					if (this.XLSCellArray[i][endY+1].className.indexOf("borderDashedLeftCellClass ") == -1)this.XLSCellArray[i][endY+1].className += "borderDashedLeftCellClass ";
				}
				else {
					if (this.XLSCellArray[i][endY].className.indexOf("borderDashedRightCellClass ") == -1)this.XLSCellArray[i][endY].className += "borderDashedRightCellClass ";
				}
			}
		}
		
		else if (value == 'leftDashed'){
			for (var i = startX; i <= endX; i++){
				if (this.XLSCellArray[i][startY].className.indexOf("borderDashedLeftCellClass ") == -1)this.XLSCellArray[i][startY].className += "borderDashedLeftCellClass ";
			}
		}
		
		else if (value == 'allDashed'){
			this.borderMapper(range, "rightDashed");
			this.borderMapper(range, "leftDashed");
			this.borderMapper(range, "bottomDashed");
			this.borderMapper(range, "topDashed");
		}

		else if (value == "none"){
			for(var i = startX; i <= endX; i++ ){
				for(var j = startY; j <= endY; j++ ){
					this.XLSCellArray[i][j].className = removeClassName(this.XLSCellArray[i][j].className, "borderLeftCellClass ");
					this.XLSCellArray[i][j].className = removeClassName(this.XLSCellArray[i][j].className, "borderDashedLeftCellClass ");
					this.XLSCellArray[i][j].className = removeClassName(this.XLSCellArray[i][j].className, "borderRightCellClass ");
					this.XLSCellArray[i][j].className = removeClassName(this.XLSCellArray[i][j].className, "borderDashedRightCellClass ");
					this.XLSCellArray[i][j].className = removeClassName(this.XLSCellArray[i][j].className, "borderTopCellClass ");
					this.XLSCellArray[i][j].className = removeClassName(this.XLSCellArray[i][j].className, "borderDashedTopCellClass ");
					this.XLSCellArray[i][j].className = removeClassName(this.XLSCellArray[i][j].className, "borderBottomCellClass ");
					this.XLSCellArray[i][j].className = removeClassName(this.XLSCellArray[i][j].className, "borderDashedBottomCellClass ");
				}
			}
			
			for (var i = startX; (i <= endX) && (endY+1 <= this.XLSColCount); i++){
					this.XLSCellArray[i][endY+1].className = removeClassName(this.XLSCellArray[i][endY+1].className, "borderLeftCellClass ");
					this.XLSCellArray[i][endY+1].className = removeClassName(this.XLSCellArray[i][endY+1].className, "borderDashedLeftCellClass ");
			}
			
			for (var j = startY; (j <= endY) && (endX+1 <= this.XLSRowCount); j++){
					this.XLSCellArray[endX+1][j].className = removeClassName(this.XLSCellArray[endX+1][j].className, "borderTopCellClass ");
					this.XLSCellArray[endX+1][j].className = removeClassName(this.XLSCellArray[endX+1][j].className, "borderDashedTopCellClass ");
			}
		}
		
		else if (value == "noneDashed"){
			for(var i = startX; i <= endX; i++ ){
				for(var j = startY; j <= endY; j++ ){
					this.XLSCellArray[i][j].className = removeClassName(this.XLSCellArray[i][j].className, "borderDashedLeftCellClass ");
					this.XLSCellArray[i][j].className = removeClassName(this.XLSCellArray[i][j].className, "borderDashedRightCellClass ");
					this.XLSCellArray[i][j].className = removeClassName(this.XLSCellArray[i][j].className, "borderDashedTopCellClass ");
					this.XLSCellArray[i][j].className = removeClassName(this.XLSCellArray[i][j].className, "borderDashedBottomCellClass ");
				}
			}
			
			for (var i = startX; (i <= endX) && (endY+1 <= this.XLSColCount); i++){
					this.XLSCellArray[i][endY+1].className = removeClassName(this.XLSCellArray[i][endY+1].className, "borderDashedLeftCellClass ");
			}
			
			for (var j = startY; (j <= endY) && (endX+1 <= this.XLSRowCount); j++){
					this.XLSCellArray[endX+1][j].className = removeClassName(this.XLSCellArray[endX+1][j].className, "borderDashedTopCellClass ");
			}
		}
		this.updateScreenFromXLS();
	};
}

defineGridFormat();
