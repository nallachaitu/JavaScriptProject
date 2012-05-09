/*
 *   copyright (c) -- chaitanya nalla and kechit goyal
 *
 */﻿

function defineClsXLSGrid() {

/******************    SINGLE ROW  &&  SINGLE COL *******************************************/
    clsMemoryRow = function(mDefaultHeight) {
        this.height = mDefaultHeight;
		this.rowNo = 0;
		this.gap = 0;
		this.isDeletable = true;
	};
	
    clsMemoryCol = function(mDefaultWidth) {
        this.width = mDefaultWidth;
		this.colNo = 0;
		this.gap = 0;
		this.isDeletable = true;
    };
    
/******************    SINGLE CELL  *******************************************/

    clsMemoryCell = function() {
        this.value = "";
        this.stringValue = "";
        this.type ="null";
        this.align = "left";
		this.backgroundColor = "";
		this.textColor = "";
		this.className = "";
		this.isDirty = false;
		this.isModified = false;
		this.cellType = "input";
		this.displayType = null;
		this.displayValue = "";
		this.nextCell = "";
		this.previousCell = "";
		this.readOnly = false;
		this.valueOptions = {value: ""};
		this.tag = "";
		this.id = "";
		
		this.rowSpan = 1;
		this.colSpan = 1;
		this.isVisible = true;
		this.spannedBy = {row: 0, col: 0};
		
		//  FOR FORMULAS  //
		this.parentRow = null;
		this.parentCol = null;
        this.parentXLSArray = null;
		this.grandParentDionGrid = null;
        this.formStr = "";
        this.outputQueue = [];
        this.operatorStack = [];
        this.calcStack = [];
		this.dependLinkArray = [];
		this.formulaParents = [];
		this.formulaCalcId = "";
		this.isCircular = false;
		this.errorMesg = "";
	};
	
	clsMemoryCell.prototype.updateCellByString = function(mString) {
		var mType;
		var mNo;
		var mArrRowCol;
		var mParent;
		var mCellRef;
		
		this.formulaParents = [];
		
		//mString = "" + mString;
		mType = getNumberType(mString);
		if (mType != this.type){
			//this.className = removeClassName (this.className, this.align + "AlignCellClass");
			//this.align = getAlignment(mType);
			this.type = mType;
			//this.className += (getAlignment(mType) + "AlignCellClass ");
		}

		if (mType == "formula") {
			this.stringValue = "" + mString;
			if (bracketMatch(mString) == 0)
				this.evalFormula();
			else {
				this.errorMesg += "\nBracket Mismatch";
				this.value = "#BRACKET";
			}		
		}
		else if (mType == "integer" ) {
			mString = removePreceedingZeros(mString);
			this.value = parseInt(mString);
			this.stringValue = "" + this.value;
			
		}
		else if (mType == "float") {
			mString = removePreceedingZeros(mString);
			this.value = parseFloat(mString);
			this.stringValue = "" + this.value;
		}
		else if (mType == "text") {
			this.value = mString;
			this.stringValue = this.value;
		}
		else if (mType == "boolean") {
			if(mString.toUpperCase() == "TRUE"){
				this.value = true;
			}
			else 
				this.value = false;
			this.stringValue = mString;
		}
		else if (mType == "null") {
			this.value = "";
			this.stringValue = "";
		}
		else {
		}
		
		var rString = createRandomString();
		if (this.type == "formula") {
			this.formulaCalcId = "";
			this.calculateRPN(rString);
		}
		else {
			this.updateParentValue(rString);
		}
		
		this.displayValue = convertTo(this.value, this.type, this.displayType);
		if (this.type == "text" && this.valueOptions && this.valueOptions.displayValue) this.displayValue = this.valueOptions.displayValue;
		this.valueOptions.value = this.value;
		this.valueOptions.displayValue = this.displayValue;
		this.valueOptions.stringValue = this.stringValue;
    };
	
	clsMemoryCell.prototype.updateParentValue = function(rString){
		for (mNo=0; mNo < this.dependLinkArray.length; mNo++) {
			mParent = workbook.getCellByString(this.dependLinkArray[mNo]);
			var rowNo = mParent.parentRow.rowNo;
			var colNo = mParent.parentCol.colNo;
			mParent.calculateRPN(rString);
			mParent.grandParentDionGrid.refreshCell(rowNo, colNo); //Kechit Goyal 18th may
		}	
	};
	
	clsMemoryCell.prototype.updateCellByWidget = function (){
		var valueOptions = this.valueOptions;
		this.type = "text";
		var mString  = valueOptions.value;
		this.stringValue = "" + mString;
		this.value = mString;
		var rString = createRandomString();
		this.updateParentValue(rString);
		this.displayValue = (valueOptions.displayValue) ? valueOptions.displayValue : convertTo(mString, this.type, this.displayType);
		if (!this.valueOptions.displayValue)this.valueOptions.displayValue = this.displayValue;
		
		if (this.cellType == "combo"){
			var values = this.valueOptions.possibleValues;
			for (var i = 0; i < values.length; i++){
				if (values[i].value == mString){
					this.valueOptions.index = i;
					this.valueOptions.subIndex = -1;
				}
				if (values[i].subMenu){
					for (var j =0 ; j < values[i].subMenu.length; j++){
						if (values[i].subMenu[j].value == mString){
							this.valueOptions.index = i;
							this.valueOptions.subIndex = j;
						}
					}
				}
			}
		}
	};
	
	clsMemoryCell.prototype.copyCell = function(){
		var newCell = new clsMemoryCell();
		newCell.value = this.value;
        newCell.stringValue = this.stringValue;
        newCell.type = this.type;
        newCell.align = this.align;
		newCell.backgroundColor = this.backgroundColor;
		newCell.textColor = this.textColor;
		newCell.className = this.className;
		newCell.isDirty = this.isDirty;
		newCell.isModified = this.isModified; 
		newCell.cellType = this.cellType;
		newCell.possibleValues = this.possibleValues;
		newCell.selectedIndex = this.selectedIndex;
		newCell.cellType = this.cellType;
		newCell.displayType = this.displayType;
		newCell.displayValue = this.displayValue;
		newCell.valueOptions = {};
		for (var i in this.valueOptions){
			newCell.valueOptions[i] = this.valueOptions[i];			
		}
		
                              //  FOR FORMULAS  //
        newCell.parentXLSArray = this.parentXLSArray; 
        newCell.formStr = this.formStr;
        newCell.outputQueue = this.outputQueue;
        newCell.operatorStack = this.operatorStack;
        newCell.calcStack = this.calcStack;
		newCell.dependLinkArray = this.dependLinkArray;
		newCell.parentRow = this.parentRow;
		newCell.parentCol = this.parentCol;
        newCell.grandParentDionGrid = this.grandParentDionGrid;
        newCell.formulaCalcId = this.formulaCalcId;
		newCell.isCircular = this.isCircular;
		return newCell;
	};

/**********************************   FORMULA  ***********************************/

	clsMemoryCell.prototype.evalFormula = function() {
        var mFormStr = this.stringValue.substr(1, this.stringValue.length-1);
        var mToken = "";
        var mChar = "";
		        
		while (this.outputQueue.length > 0) {
			this.outputQueue.pop();
		}
			
		while (this.operatorStack.length > 0) {
			this.operatorStack.pop();
		}

        for (var i = 0; i < mFormStr.length; i++) {
            mChar = mFormStr.substr(i, 1);
            if (mChar == " ") {
                continue;
            } 

            if (isOperator(mChar)) {
                this.handleToken(mToken);
                mToken = "";
                this.pushOperatorOnStack(mChar);
            }           

            else {
                mToken = mToken + mChar;
            }
        }
        this.handleToken(mToken);
        mToken = "";
        this.finalTransfer();	
    };

    clsMemoryCell.prototype.pushOperatorOnStack = function(mChar) {
        var mOp;
        
		if (mChar == "(") {
			this.outputQueue.push(mChar);
            this.operatorStack.push(mChar);
            return
        }

        if (mChar == ")") {
            this.handleCloseBracket();
            return
        }


        if (mChar == ",") {
            this.handleComma();
            return
        }
    
        if (operatorAssociativity(mChar) == "L") {
            while (true) {
                if (this.operatorStack.length == 0) {
                    this.operatorStack.push(mChar);
                    break;
                }
                else if (this.operatorStack[this.operatorStack.length-1] == ")" ) {
                    this.operatorStack.push(mChar);
                    break;
                }
                else {
                    if (operatorPrecedence(mChar) > operatorPrecedence(this.operatorStack[this.operatorStack.length-1])) {
                        this.operatorStack.push(mChar);
                        break;
                    }
                    else {
                        mOp = this.operatorStack.pop();
                        this.outputQueue.push(mOp);
                    }
                }
            }
        }
        else {
            while (true) {
                if (this.operatorStack.length == 0) {
                    this.operatorStack.push(mChar);
                    break;
                }
                else if (this.operatorStack[this.operatorStack.length-1] == ")" ) {
                    this.operatorStack.push(mChar);
                    break;
                }
                else {
                    if (operatorPrecedence(mChar) >= operatorPrecedence(this.operatorStack[this.operatorStack.length-1])) {
                        this.operatorStack.push(mChar);
                        break;
                    }
                    else {
                        mOp = this.operatorStack.pop();
                        this.outputQueue.push(mOp);
                    }
                }
            }
    
        }
    };

    clsMemoryCell.prototype.handleToken = function(mToken) {
        var mArrColRow;
		if ( (mToken == "") || (mToken == " ") ) {
            return
        }

        if (isFunction(mToken)) {
            this.operatorStack.push(mToken);
            return
        }
        else if ( getNumberType(mToken) == "integer") {
            this.outputQueue.push(parseInt(mToken));
        }            
        else if ( getNumberType(mToken) == "float") {
            this.outputQueue.push(parseFloat(mToken));
        }
		else if (getNumberType (mToken) == "boolean") {
			this.outputQueue.push(toBoolean(mToken));
		}
		else if (workbook.getCellById((mToken), this.grandParentDionGrid) != null){
			var cell = workbook.getCellById((mToken), this.grandParentDionGrid);
			this.formulaParents[this.formulaParents.length] = mToken;
			
			if (!cell) {
				this.value = "#CELL";
				this.errorMesg += "\nCell Not Found" + mToken;
			}
			this.outputQueue.push(mToken);

			if (mToken.toUpperCase == this.getCellRefAsString() ) {
				this.isCircular = true;
				this.errorMesg += "\nCircular Reference";
				this.displayValue = "#CIRC";
				this.value = "#CIRC";
			}
			this.updateParentDependLink(cell.parentXLSArray[cell.parentRow.rowNo][cell.parentCol.colNo]);
		}
		//Distinction Between a normal string and a cell name
		else if (mToken.substr(0,1) == "\"" && mToken.substr(mToken.length-1,1) == "\""){
			this.outputQueue.push(mToken);
		}
		else if (mToken.substr(0,1) == "'" && mToken.substr(mToken.length-1,1) == "'"){
			this.outputQueue.push(mToken);
		}
		else if (workbook.nameMap.get(mToken) != null){
			mToken = workbook.nameMap.get(mToken).split(".")[1];
			this.handleToken(mToken);
		}
		else {
			this.errorMesg += "\n" + mToken + " Not Found";
			this.displayValue = "#NAME";
			this.value = "#NAME";
		}
    };
	
    clsMemoryCell.prototype.handleCloseBracket = function() {
        var mOp;
		var mErr = true;
		var argArray = [];
		while(this.outputQueue.length > 0) {
			mOp = this.outputQueue.pop();
            if (mOp == "(" ) {
				break;
			}
			else {
				argArray.push(mOp);
			}
		}
		argArray = argArray.reverse();
        while (this.operatorStack.length > 0) {
            mOp = this.operatorStack.pop();
            if (mOp == "(" ) {
                if (this.operatorStack.length > 0) {
                    if (isFunction(this.operatorStack[this.operatorStack.length -1])) {
                        mOp = this.operatorStack.pop();
						argArray.push(mOp);
                        if (argArray.length > 0)
							this.outputQueue.push(argArray);
						else {
							this.errorMesg = "Empty Bracket Found";
							this.displayValue = "#EMPTY BRACKET";
							this.value = "#EMPTY BRACKET";
						}
                    }
					else {
						if (argArray.length > 0)
							this.outputQueue.push(argArray);
						else {
							this.errorMesg = "Empty Bracket Found";
							this.displayValue = "#EMPTY BRACKET";
							this.value = "#EMPTY BRACKET";
						}
					}
                }
				else {
					if (argArray.length > 0)
							this.outputQueue.push(argArray);
					else {
						this.errorMesg = "Empty Bracket Found";
						this.displayValue = "#EMPTY BRACKET";
						this.value = "#EMPTY BRACKET";
					}
				}
				mErr = false;
                break;
            }
            else {
				argArray.push(mOp);
            }
        }
		if (mErr) {
			this.errorMesg += "\nBracket Mismatch found";
			this.displayValue = "#BRACKET";
			this.value = "#BRACKET";
		}
    };
	
    clsMemoryCell.prototype.handleComma = function() {
        var mOp;
        while (this.operatorStack.length > 0) {
            if (this.operatorStack[this.operatorStack.length-1] == "(") {
                break;
            }
            mOp = this.operatorStack.pop();
            this.outputQueue.push(mOp);
        }
    };

   clsMemoryCell.prototype.updateParentDependLink = function(mParent) {	
		var i = 0;
		if (mParent == undefined){
			return;
		}
		var myDependLinkArray = mParent.dependLinkArray;
		var cellString = this.grandParentDionGrid.gridName + "." + this.getCellRefAsString();
		
		while (i < myDependLinkArray.length) {
			if (myDependLinkArray[i] == cellString ) {
				myDependLinkArray.splice(i,1);
			}
			else {
				i++;
			}
		}
		myDependLinkArray.push(cellString);	//Do we really need cur cell ref.... //Kechit Goyal 18th may	
	}	;

    clsMemoryCell.prototype.finalTransfer = function() {
        var mOp;
        while (this.operatorStack.length > 0) { 
            mOp = this.operatorStack.pop();
            this.outputQueue.push(mOp);
        }
    };

    clsMemoryCell.prototype.calculateRPN = function(mRandomString) {
        var mToken;
        var mPop1, mPop2, mPop3, mPop;
        var mNo;
		
		if (this.isCircular) {
			return
		}
		if (this.type == "formula") {
			while (this.calcStack.length > 0) {
				this.calcStack.pop();
			}
			if (this.errorMesg == "")
				var result = this.solveFormula(this.calcStack, this.outputQueue);
			
			
			if (this.formulaCalcId == mRandomString) {
				this.isCircular = true;
				this.errorMesg += "\nCircular Reference";
				this.displayValue = "#CIRC";
				this.value = "#CIRC";
			}
			else {
				if (this.formulaCalcId == "")
					this.formulaCalcId = mRandomString;
				this.isCircular = false;
			}
			
			if (this.errorMesg == "") {
				this.value = result;
				this.displayValue = this.value;
				this.valueOptions.value = result;
			}
		}
		
		var mNum, mCellRef, mParent, mArrRowCol;	
		for (mNum=0; mNum < this.dependLinkArray.length; mNum++) {
			mParent = workbook.getCellByString(this.dependLinkArray[mNum]);
			var rowNo = mParent.parentRow.rowNo;
			var colNo = mParent.parentCol.colNo;
			mParent.calculateRPN(mRandomString);
			mParent.grandParentDionGrid.refreshCell(rowNo, colNo);
			mParent.grandParentDionGrid.fireOnChange(rowNo, colNo);
		}
	};
	
	clsMemoryCell.prototype.popFromCalcStack = function(calcStack) {
		var mPop = calcStack.pop();
		if (workbook.getCellById(mPop, this.grandParentDionGrid) != null){
				var cell = workbook.getCellById(mPop, this.grandParentDionGrid);
				if (cell == undefined || cell == null){
					this.value = "#CELL";
					this.errorMesg = "Cell Not Found";
				}
				else if(!(cell.errorMesg == "")){
					this.errorMesg = cell.errorMesg;
					this.value = "#VALUE";
				}
				else {
					
					var val = cell.value;
					if (val == "") return 0;
					else return val;
				}
			}
		
		else if (typeof(mPop) == "string" && mPop.length > 0){
			if (toBoolean(mPop) == true) return true;
			else if (toBoolean(mPop) == false) return false;
			else return mPop;
		}
		else if (mPop == "") 
			return 0;
		else 
            return mPop;
    };

	clsMemoryCell.prototype.getCellRefAsString = function() {
		var mRowNo = this.parentRow.rowNo;
		var mColNo = this.parentCol.colNo;
		var mStr = "" + getColLabelFromNumber(mColNo) + "" + mRowNo;
		return mStr;
	};

	clsMemoryCell.prototype.clearAllBackLinks = function() {
		var thisRef = this.getCellRefAsString();
		var i, j;
		var mToken ;
        var mArrColRow;
		var mCellObj;
		var mcellObjArray;
		
		for (i=0; i < this.outputQueue.length; i++) {
			mToken = this.outputQueue[i];
			
			if (isCell(mToken)){
				mToken = this.grandParentDionGrid.gridName + "." + mToken;
				mArrColRow = getCellRowCol(mToken.toUpperCase());
				mCellObj = this.parentXLSArray[parseInt(mArrColRow[1])][getColNoFromLabel(mArrColRow[0])];
				mCellObjArray = mCellObj.dependLinkArray;
				
				j = 0;
				while ( j < mCellObjArray.length) {
					if (mToken  == mCellObjArray[j] ) {
						mCellObjArray.splice(j,1);
					}
					else {
						j++;
					}
				}
			}
			if (isGridCell(mToken)){
				var cell = workbook.getCellByString(mToken);
				mCellObjArray = cell.dependLinkArray;
				j = 0;
				while ( j < mCellObjArray.length) {
					if (mToken  == mCellObjArray[j] ) {
						mCellObjArray.splice(j,1);
					}
					else {
						j++;
					}
				}
			}
		}
	};
	
	clsMemoryCell.prototype.clearFormula = function() {	
		this.formStr = "";
        this.outputQueue = [];
        this.operatorStack = [];
        this.calcStack = [];
		this.formulaCalcId = "";
		this.isCircular = false;
		this.errorMesg = "";
	};

	clsMemoryCell.prototype.solveFormula = function (myCalcStack, outputQueue){
		var mToken;
		for (var i = 0; i < outputQueue.length; i++) {
			mToken = outputQueue[i];
			if (typeof (mToken) == "object"){
				if (mToken.length){
					var newCalcStack = [];
					myCalcStack.push(this.solveFormula(newCalcStack,mToken));
				}
			}
			else if (isOperator(mToken) ) {
				var mPop1, mPop2, mPop3, mPop;
				var argCount = operatorArgCount(mToken);
				var countMismatch = false;
				if(argCount != -1){
					if (argCount == 1){
						mPop1 = this.popFromCalcStack(myCalcStack);
						if (mPop1 == undefined || mPop1 == null){
							this.errorMesg += "\Operator " + mToken + " Argument Not Found";
							//this.displayValue = "#ARG";
							this.value = "#ARG";
							countMismatch = true;
						}
						
					}
					else if (argCount == 2){
						mPop1 = this.popFromCalcStack(myCalcStack);
						if (mPop1 == undefined || mPop1 == null){
							this.errorMesg += "\Operator " + mToken + " Argument Not Found";
							//this.displayValue = "#ARG";
							this.value = "#ARG";
							countMismatch = true;
						}
						mPop2 = this.popFromCalcStack(myCalcStack);
						if (mPop2 == undefined || mPop2 == null){
							this.errorMesg += "\Operator " + mToken + " Argument Not Found";
							//this.displayValue = "#ARG";
							this.value = "#ARG";
							countMismatch = true;
						}
					}
				}
				else {
					this.errorMesg += "\Operator " + mToken + " Argument Count Not Defined";
					//this.displayValue = "#ARG";
					this.value = "#ARG";
					countMismatch = true;
				}
				if (!countMismatch){
					switch (mToken) {
						case "+" :
							if (typeof(mPop2) != typeof(0) || typeof(mPop1) != typeof(0) )
								mPop = "(" +  mPop2 + "+" + mPop1 + ")";
							else mPop = mPop2 + mPop1;
							myCalcStack.push(mPop);
							break;
							
						case "-" :
							if (typeof(mPop2) != typeof(0) || typeof(mPop1) != typeof(0) )
								mPop = "(" +  mPop2 + "-" + mPop1 + ")";
							else mPop = mPop2 - mPop1;
							myCalcStack.push(mPop);
							break;
						case "*" :
							if (typeof(mPop2) != typeof(0) || typeof(mPop1) != typeof(0) )
								mPop = "(" + mPop2 + "*" + mPop1 + ")";
							else mPop = mPop2 * mPop1;
							myCalcStack.push(mPop);
							break;
						case "/" :
							if (typeof(mPop2) != typeof(0) || typeof(mPop1) != typeof(0) )
								mPop = "(" +  mPop2 + "/" + mPop1 + ")";
							else mPop = mPop2 / mPop1;
							
							myCalcStack.push(mPop);
							break;
					}         
				}  
			}				
			else {
				myCalcStack.push(mToken);
			}

		}
		var result = this.popFromCalcStack(myCalcStack);
		var extraArg = myCalcStack.pop();
			
		if (extraArg != undefined || extraArg != null){
			this.errorMesg += "\nExtra Arguments Found - (" + extraArg + ")";
			//this.displayValue = "#ARGEXTRA";
			this.value = "#ARGEXTRA";
		}
		if (this.errorMesg == "")return  result;
		else return this.value;
	};

}

defineClsXLSGrid();
