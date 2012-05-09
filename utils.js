/*
 *   copyright (c) -- chaitanya nalla and kechit goyal
 *
 */

var _startCellY = 1;
var _startCellX = 1;
var _endCellY = 1;
var _endCellX = 1;
var _isCellShiftDrag = false;
var _isCut = false;


function getColLabelFromNumber(mNo) {
    if (mNo <= 26) {
        return String.fromCharCode(mNo + 64);
    }
    var mNo1 = parseInt(mNo/26);
    var mNo2 = mNo - 26*mNo1;
    if (mNo2 == 0) {
        mNo2 = 26;
        mNo1--;
    }
    return String.fromCharCode(mNo1 + 64) + String.fromCharCode(mNo2 + 64); 
   
}

function getColNoFromLabel(mLabel) {
    var i;
    var mNo = 0;
    var mStr = mLabel.toUpperCase();
    
    for (var i=0; i<mStr.length;i++) {
        mNo = mNo*26;
        mNo = mNo + mStr.charCodeAt(i)-65;
    }
    return mNo + 1;
   
}

function getCellRowCol(mString) {
	var i;
	var mCode;
	var mArrTemp = [];
	var cellRefExpression = /^[\$]{0,1}[a-zA-Z]{1,3}[\$]{0,1}[1-9]{1,1}[0-9]{0,3}$/;
	if(!mString.match(cellRefExpression)){
		return null;
    }
    else {
		mString = mString.replace(/[\$]/g,"");	
        for (i=0; i < mString.length;i++) {
            mCode = mString.charCodeAt(i);
            if ((mCode >=48) && (mCode <= 57) ) {
                break;
            }
        }
        mArrTemp.push(mString.substr(0, i));
       // mArrTemp.push(mString.substr(i, mString.length,i));
        mArrTemp.push(mString.substr(i, mString.length ));
        return mArrTemp;
    }
}		

function getNumberType(mStr){
	if (mStr == null) {
        return "null";
    }
    mStr = "" + mStr;
    if (mStr == "") {
       return "null";
    }
	if ((""+ mStr).substr(0,1) == "="){
		return "formula";
	}
    
	var numericExpression = /^[0-9]+$/;
	var myStr;
	if ( (mStr.substr(0,1) == "+") || (mStr.substr(0,1) == "-") ) {
	    myStr = mStr.substr(1);
	}
	else {
	    myStr = mStr;
	}

	if(myStr.toUpperCase() == "TRUE" || myStr.toUpperCase() == "FALSE"){
		return "boolean";
	}
	if(myStr.match(numericExpression)){
		return "integer";
	}else{
	    var mPos = myStr.indexOf(".");
	    if (mPos < 0) {
	        return "text";
	    }
	    var myLeftStr = myStr.substr(0, mPos);
	    var myRightStr = myStr.substr(mPos+1);
	
	if(myStr.substr(0, 1) == "." && myRightStr.match(numericExpression)){
		return "float";
	}

    	if(!myLeftStr.match(numericExpression)){
	    	return "text";
		}
    	if(!myRightStr.match(numericExpression)){
	    	return "text";
		}
		return "float";
		
	}

}

function removePreceedingZeros(mString){
	mString = mString + "";
	while (mString.indexOf("0") === 0 && mString.length != 1){
		mString = mString.substr(1);
	}
	return mString;
}

//eg get 100 from "100px"
function getLengthFromString(width) {
	if (width.length > 2){
		return parseInt(width.substr(0,width.length-2));
	}
	else return null;
}

function removeAllChildren(cell){
	if ( cell.hasChildNodes() )
	{
		while ( cell.childNodes.length >= 1 )
		{
			cell.removeChild( cell.firstChild );       
		} 
	}
}

function isInRange(x,y,range){
	var rangeArray = range.split(":");
	if (rangeArray.length != 4) return false;
	else {
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
		
		if (x >=startX && x <= endX && y >=startY && y <= endY) return true;
		else return false;
	}
}

function isInsideRange(x,y,range){
	var rangeArray = range.split(":");
	if (rangeArray.length != 4) return false;
	else {
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
		
		if (x >startX && x < endX && y >startY && y < endY) return true;
		else return false;
	}
}

function isXInRange(x,range){
	var rangeArray = range.split(":");
	if (rangeArray.length != 4) return false;
	else {
		var startX = parseInt(rangeArray[0]);
		var endX = parseInt(rangeArray[2]);
		
		if (startX > endX){
			var temp = endX;
			endX = startX;
			startX = temp;
		}
	
		
		if (x >=startX && x <= endX) return true;
		else return false;
	}
}

function isYInRange(y,range){
	var rangeArray = range.split(":");
	if (rangeArray.length != 4) return false;
	else {
		var startY = parseInt(rangeArray[1]);
		var endY = parseInt(rangeArray[3]);
		
		if (startY > endY){
			var temp = endY;
			endY = startY;
			startY = temp;
		}
	
		if (y >=startY && y <= endY) return true;
		else return false;
	}
}

function removeClassName(classString, className){
	var temp = classString;
	while(temp.indexOf(className) != -1){
		temp = temp.replace(className, "");
	}
	return temp;
}

function toggleClassName (classString, className){
	if (classString.indexOf(className) != -1){
		return removeClassName(classString, className);
	}
	else return addClassName(classString, className);
}

function addClassName (classString, className){
	classString = removeClassName(classString, className);
	classString += className;
	return classString;
}

function colCount(range){
	var rangeArray = range.split(":");
	if (rangeArray.length != 4) return "";
	else {
		var startY = parseInt(rangeArray[1]);
		var endY = parseInt(rangeArray[3]);
		
		if (startY > endY){
			var temp = endY;
			endY = startY;
			startY = temp;
		}
		return endY-startY + 1;
	}
}

function rowCount(range){
	var rangeArray = range.split(":");
	if (rangeArray.length != 4) return "";
	else {
		var startX = parseInt(rangeArray[0]);
		var endX = parseInt(rangeArray[2]);
		
		if (startX > endX){
			var temp = endX;
			endX = startX;
			startX = temp;
		}
		return endX-startX + 1;
	}
}

function getStartX(range){
	var rangeArray = range.split(":");
	if (rangeArray.length != 4) return "";
	else {
		var startX = parseInt(rangeArray[0]);
		var endX = parseInt(rangeArray[2]);
		
		if (startX > endX){
			var temp = endX;
			endX = startX;
			startX = temp;
		}
		return startX;
	}
}

function getStartY(range){
	var rangeArray = range.split(":");
	if (rangeArray.length != 4) return "";
	else {
		var startY = parseInt(rangeArray[1]);
		var endY = parseInt(rangeArray[3]);
		
		if (startY > endY){
			var temp = endY;
			endY = startY;
			startY = temp;
		}
		return startY;
	}

}

function getEndX(range){
	var rangeArray = range.split(":");
	if (rangeArray.length != 4) return "";
	else {
		var startX = parseInt(rangeArray[0]);
		var endX = parseInt(rangeArray[2]);
		
		if (startX > endX){
			var temp = endX;
			endX = startX;
			startX = temp;
		}
		return endX;
	}
}

function getEndY(range){
	var rangeArray = range.split(":");
	if (rangeArray.length != 4) return "";
	else {
		var startY = parseInt(rangeArray[1]);
		var endY = parseInt(rangeArray[3]);
		
		if (startY > endY){
			var temp = endY;
			endY = startY;
			startY = temp;
		}
		return endY;
	}

}

function createRandomString() {
	var mString = "";
	mString = mString = mString + (parseInt(Math.random()*89999) + 10000 );
	mString = mString = mString + (parseInt(Math.random()*89999) + 10000 );
	mString = mString = mString + (parseInt(Math.random()*89999) + 10000 );
	return mString;
}

function preventDefaultBehaviour(event){
	if (event){
		if(event.preventDefault) {
			event.preventDefault();
			//event.cancelBubble = true;
			//event.returnValue = false;
			//event.stopPropagation();
		}
		else {
			event.cancelBubble = true;
			event.returnValue = false;
			return false;
		}
	}
}

function preventPropagation(event){
	if (event){
		if(event.preventDefault) {
			event.preventDefault();
			event.cancelBubble = true;
			event.returnValue = false;
			event.stopPropagation();
		}
		else {
			event.cancelBubble = true;
			event.returnValue = false;
			event.stopPropagation();
			return false;
		}
	}
}

function convertTo(myStr, type, formatType){
	if (!myStr && myStr != 0 && myStr != "0") return "";
	myStr = myStr.toString();
	if (formatType == undefined || formatType == null) {
		if (typeof(myStr)=="string")
			return removeDoubleQuotes(myStr);
		else 
			return myStr;
	}
	if (formatType.format == "percent"){
		return formatAsPercent(myStr, type, formatType);
	}
	else if (formatType.format  == "currency"){
		return formatAsCurrency(myStr, type, formatType);
	}
	else if (formatType.format  == "accounting"){
		return formatAsCurrency(myStr, type, formatType);
	}
	else if (formatType.format == "delta"){
		return removeDelta(myStr) + " &#9651";
	}
	return removeDoubleQuotes(myStr);
}

function removeDelta(myStr){
	while(myStr.indexOf(" &#9651") != -1)
		myStr = myStr.replace(" &#9651","");
	return myStr;
}

//Modify Formula handles double quotes and passes control to removeQuotes
function removeDoubleQuotes(myStr){
	if (myStr.split("\"").length % 2 == 0)return myStr;
	if (myStr.split("\'").length % 2 == 0)return myStr;
	
	var arr = myStr.split("\"");
	var result = "";
	for (var i = 0; i < arr.length; i= i+1){
		result += removeSingleQuotes(arr[i]);
	}
	return result;
}

//removeQuotes handles quotes and passes control to modify
function removeSingleQuotes(myStr){
	var arr = myStr.split("\'");
	var result = "";
	for (var i = 0; i < arr.length; i= i+2){
		result += arr[i];
	}
	return result;
}

function convertToNum(myStr, formatType){
	myStr = "" + myStr;
	
	if (formatType.symbol){
		while(myStr.indexOf(formatType.symbol) != -1)
			myStr = myStr.replace(formatType.symbol,"");
	}
	while(myStr.indexOf(",")  != -1){
		myStr = myStr.replace(",","");
	}
	while(myStr.indexOf(" ")  != -1){
		myStr = myStr.replace(" ","");
	}
	return myStr;
}

function formatAsCurrency(myStr, type, formatType){
	if (myStr) myStr = convertToNum(myStr, formatType);
	if (type == "formula"){type = getNumberType(""+myStr);};
	if (type == "integer"){
		myStr = "" + myStr;
		myStr = adjustDecimal(myStr, formatType.decimalCount);
	}
	else if (type == "float"){
		var factor = Math.pow(10, formatType.decimalCount);
		myStr = "" + (Math.round(parseFloat(myStr) * factor)/factor);
		myStr = adjustDecimal(myStr, formatType.decimalCount);
	}
	else return removeDoubleQuotes(myStr);
	
	var mPos = myStr.indexOf(".");
	var myLeftStr = myStr.substr(0, mPos);
	var myRightStr = myStr.substr(mPos+1);		
	if (mPos == -1){
		myLeftStr = myStr;
		myRightStr = "";
	}
	var hasNegative = false;
	if (myLeftStr.indexOf("-") != -1){
		myLeftStr = myLeftStr.substr(1);
		hasNegative = true;
	}
	var length = myLeftStr.length;
	if (length == 0) return formatType.symbol + "0";
	else if (length > 11) return formatAsCurSymbol(myStr);
	var index = length % 3;
	if (formatType.useThousandSeparator){
		var result = formatType.symbol + myLeftStr.substr(0,index);
		for (var i = index; i < length; i = i+3){
			if (i != 0)
				result += ("," + myLeftStr.substr(i, 3));
			else 
				result += myLeftStr.substr(i, 3);
		}
		if (hasNegative)
			result = "-" + result;
		
		if (mPos != -1)
			return result + "." + myRightStr;
		else return result;
	}
	else if (!(formatType.symbol + myStr)) return "";
	else return formatType.symbol + myStr;


}

function formatAsCurSymbol(myStr){
	//1,000,000,000,000
	myStr = "" + myStr;
	var strLength = myStr.length
	if (strLength < 11) return myStr;
	var number = parseInt(myStr)/ getNotional(strLength);
	var decimalAdjustedNum = adjustDecimal ("" + number, 2);
	return decimalAdjustedNum + getSymbol(strLength);
}


function getNotional(lengthNum){
	if (lengthNum > 9 && lengthNum <= 12) return Math.pow(10, 9);
	else if (lengthNum > 12) return Math.pow(10, 12);
}

function getSymbol(lengthNum){
	if (lengthNum > 9 && lengthNum <= 12) return "B";
	else if (lengthNum > 12) return "T";
}

function formatAsPercent(myStr, type, formatType){
	if (type == "formula"){type = getNumberType(""+myStr);};
	var value = "";
	if (type == "integer"){
		if (myStr != "0" && myStr != "")
			value =  myStr + "00";
		else myStr = "0";
	}
	else if (type == "float"){
		if (myStr != "0" && myStr != ""){
			var factor = Math.pow(10, formatType.decimalCount); 
			value = "" + Math.round(parseFloat(myStr) * 100 * factor)/factor;
		}
		else myStr = "0";
	}
	else return removeDoubleQuotes(myStr);;
	
	value = adjustDecimal(value, formatType.decimalCount);
	return value + "%";
}

function adjustDecimal(myStr, decimalCount){
	if (myStr == "0" || myStr == "" || !myStr) return "0";
	
	var mPos = myStr.indexOf(".");
	if (mPos < 0) {
		if (decimalCount > 0)
			return myStr + "." + zeroString(decimalCount);
		else return myStr;
	}
	var myLeftStr = myStr.substr(0, mPos);
	if (decimalCount <= 0){
		return myLeftStr;
	}
	var myRightStr = myStr.substr(mPos+1);
	var length = myRightStr.length;
	if (length < decimalCount ){
		return myLeftStr + "." + myRightStr + zeroString(decimalCount - length);
	}
	//rounding off needs to be done here
	else return myLeftStr + "." + myRightStr.substr(0,decimalCount);
}

function zeroString(count){
	var result = "";
	for (var i = 0; i < count; i++){
		result += "0";
	}
	return result;
}

function getAlignment (mType){
	switch(mType){
		case "integer": return "right";
		case "float": return "right";
		case "formula" : return "right";
		case "text" : return "left";
		case "boolean" : return "left";
		case "null" : return "left";
		default : return "left";
	}
}

function getNumRange(range){
	if (isCell(range)){
		var mArrColRow = getCellRowCol(range.toUpperCase());
		var rowNo = parseInt(mArrColRow[1]);
		var colNo = getColNoFromLabel(mArrColRow[0]);
		var cellRange = rowNo + ":" + colNo + ":" + rowNo + ":" + colNo;
		return cellRange;
		
	}
	else if (isRange(range)){
		var rangeArray = range.split(":");
		var rangeStart = getCellRowCol(rangeArray[0].toUpperCase());	
		var rangeEnd = getCellRowCol(rangeArray[1].toUpperCase());
		var rangeStartY = getColNoFromLabel(rangeStart[0]);	
		var rangeEndY = getColNoFromLabel(rangeEnd[0]);
		var rangeStartX = parseInt(rangeStart[1]);
		var rangeEndX = parseInt(rangeEnd[1]);
		var numRange = rangeStartX + ":" + rangeStartY + ":" + rangeEndX + ":" + rangeEndY;
		return numRange;
	}
	else if (isNumRange(range)) return range;
	else return null;
}

function getRange(range){
	if (!isNumRange(range)) return;
	var rangeArray = range.split(":");
	var startX = parseInt(rangeArray[0]);
	var startY = parseInt(rangeArray[1]);
	var endX = parseInt(rangeArray[2]);
	var endY = parseInt(rangeArray[3]);
	/*
	if (startX > endX){
		var temp = endX;
		endX = startX;
		startX = temp;
	}
	if (startY > endY){
		var temp = endY;
		endY = startY;
		startY = temp;
	}*/
	if (startX == endX && startY == endY) return getColLabelFromNumber(startY) + startX;
	else return getColLabelFromNumber(startY) + startX + ":" + getColLabelFromNumber(endY) + endX;
	
}

function stringToPercent(myStr){
	myStr = "" + myStr;
	myStr = myStr.replace(/^[ ]|[ ]$/g,"");
	myStr = myStr.replace(/%/,"");
	var type = getNumberType(myStr);
	if (type == "integer"){
		return (parseInt(myStr) / 100);
	}
	else if (type == "float"){
		return (parseFloat(myStr) / 100);
		
	}
	else return "0%";
}

function percentToVal(myStr){
	myStr = "" + myStr;
	var type = getNumberType(myStr);
	if (type == "integer"){
		return (parseInt(myStr) * 100) + "%";
	}
	else if (type == "float"){
		return (parseFloat(myStr) * 100) + "%";
		
	}
	else return myStr;
}

function removeAllChildren(element){
	if (element.childNodes){
		while(element.childNodes.length>0){
			element.removeChild(element.childNodes[0]);
		}
	}
	return element;
}



function trapUpKeyFn(event) {
    if (event == null) var event = window.event; 			
    switch (event.keyCode) {
		case 16: //Shift key
			_isCellShiftDrag = false;
			break;
	}
}

function OnMouseUp(e) {
	document.onmousemove = null;
	document.onmouseup = null;
    if (!workbook.curDionGrid) return;
	workbook.curDionGrid.clearInterpolation();
	workbook.curDionGrid.clearDrag();
	
}


function trapKeyFn(event) {
    if (!workbook.curDionGrid) return;
    if (event == null) var event = window.event; 
	var mmRowNo;
    var mmColNo;
    if (event.keyCode == 8){
    	if (workbook.curDionGrid.mode == "select"){
    		preventDefaultBehaviour(event);
    	}
    }
    if (workbook.curDionGrid.mode == "none") return;
	
	switch (event.keyCode) { 
	
		case 16: //Shift key
			if (! _isCellShiftDrag){
				var rangeArray = workbook.selectedRange.split(":");
				if (rangeArray.length != 4) return false;
				_isCellShiftDrag = true;
				_startCellX =  parseInt(rangeArray[0]);
				_startCellY = parseInt(rangeArray[1]);
				_endCellX = parseInt(rangeArray[2]);
				_endCellY =  parseInt(rangeArray[3]);
			}
			break;
			
		case 17:
			break;
			
        case 9:             // TAB Key
			var XLSCellObj = null;
			if (workbook.curDionGrid.isInXLSRange(workbook.curDionGrid.curXLSRowNo,workbook.curDionGrid.curXLSColNo))
				XLSCellObj = workbook.curDionGrid.XLSCellArray[workbook.curDionGrid.curXLSRowNo][workbook.curDionGrid.curXLSColNo];
            if (event.shiftKey) {
				if (XLSCellObj && XLSCellObj.previousCell != ""){
					var mArrColRow = getCellRowCol(XLSCellObj.previousCell);
					mmRowNo = parseInt(mArrColRow[1]);
					mmColNo = getColNoFromLabel(mArrColRow[0]);
				}
				else {
					var cell = workbook.curDionGrid.previousEditableCell();
					mmRowNo = cell.rowNo;
					mmColNo = cell.colNo;
				}
            }
            else {
				if (XLSCellObj && XLSCellObj.nextCell != ""){
					var mArrColRow = getCellRowCol(XLSCellObj.nextCell);
					mmRowNo = parseInt(mArrColRow[1]);
					mmColNo = getColNoFromLabel(mArrColRow[0]);
				}
				else {
					var cell = workbook.curDionGrid.nextEditableCell();
					mmRowNo = cell.rowNo;
					mmColNo = cell.colNo;
				}
            }
            
            workbook.curDionGrid.goToCell(mmRowNo, mmColNo);
			preventDefaultBehaviour(event);
            break;

        case 33:            //PGUP
            mmRowNo = workbook.curDionGrid.curXLSRowNo - 10;
            if (mmRowNo < 1) {
                mmRowNo = 1;
            }
            if (!event.shiftKey)
				workbook.curDionGrid.goToCell(mmRowNo, workbook.curDionGrid.curXLSColNo);
			else {
				_endCellX -= 10;
				if (_endCellX < 1) _endCellX = 1;
				workbook.clearRange(workbook.selectedRange);
				workbook.selectedRange = _startCellX + ":" + _startCellY + ":" + _endCellX + ":" + _endCellY; 
			}
			preventDefaultBehaviour(event);
            break;

        case 34:            //PGDOWN
            mmRowNo = workbook.curDionGrid.curXLSRowNo + 10;
            if (mmRowNo > workbook.curDionGrid.XLSRowCount) {
                mmRowNo = workbook.curDionGrid.XLSRowCount;
            }
            if (!event.shiftKey)
				workbook.curDionGrid.goToCell(mmRowNo, workbook.curDionGrid.curXLSColNo);
			else {
				_endCellX += 10;
				if (_endCellX > workbook.curDionGrid.XLSRowCount) _endCellX = workbook.curDionGrid.XLSRowCount;
				workbook.clearRange(workbook.selectedRange);
				workbook.selectedRange = _startCellX + ":" + _startCellY + ":" + _endCellX + ":" + _endCellY; 
			}
			break;
            
        case 37:            //LEFTARROW
			if (workbook.curDionGrid.mode == "select"){
				if (!event.shiftKey){
					var cell = workbook.curDionGrid.leftEditableCell();
					mmRowNo = cell.rowNo;
					mmColNo = cell.colNo;
					workbook.curDionGrid.goToCell(mmRowNo, mmColNo);
				}
				else {
					_endCellY -= 1;
					if (_endCellY < 1) _endCellY = 1;
					workbook.clearRange(workbook.selectedRange);
					workbook.selectedRange = _startCellX + ":" + _startCellY + ":" + _endCellX + ":" + _endCellY; 
				}	
				preventDefaultBehaviour(event);
			}
          break;

        case 38:            //UPARROW
            if (!event.shiftKey){
					var cell = workbook.curDionGrid.upEditableCell();
					mmRowNo = cell.rowNo;
					mmColNo = cell.colNo;
					workbook.curDionGrid.goToCell(mmRowNo, mmColNo);
			}
			else {
				_endCellX -= 1;
				if (_endCellX < 1) _endCellX = 1;
				workbook.clearRange(workbook.selectedRange);
				workbook.selectedRange = _startCellX + ":" + _startCellY + ":" + _endCellX + ":" + _endCellY; 
			}
			preventDefaultBehaviour(event);
            break;
        case 39:            //RIGHTARROW
			if (workbook.curDionGrid.mode == "select"){
				if (!event.shiftKey){
					var cell = workbook.curDionGrid.rightEditableCell();
					mmRowNo = cell.rowNo;
					mmColNo = cell.colNo;
					workbook.curDionGrid.goToCell(mmRowNo, mmColNo);
				}
				else {
					_endCellY += 1;
					if (_endCellY > workbook.curDionGrid.XLSColCount) _endCellY = workbook.curDionGrid.XLSColCount;
					workbook.clearRange(workbook.selectedRange);
					workbook.selectedRange = _startCellX + ":" + _startCellY + ":" + _endCellX + ":" + _endCellY; 
				}
				preventDefaultBehaviour(event);
			}
            break;
        case 40:            //DOWNARROW
            if (!event.shiftKey){
					var cell = workbook.curDionGrid.downEditableCell();
					mmRowNo = cell.rowNo;
					mmColNo = cell.colNo;
					workbook.curDionGrid.goToCell(mmRowNo, mmColNo);
			}
			else {
				_endCellX += 1;
				if (_endCellX > workbook.curDionGrid.XLSRowCount) _endCellX = workbook.curDionGrid.XLSRowCount;
				workbook.clearRange(workbook.selectedRange);
				workbook.selectedRange = _startCellX + ":" + _startCellY + ":" + _endCellX + ":" + _endCellY; 
			}
			preventDefaultBehaviour(event);
            break;  
		case 67:
			if (event.ctrlKey) // Ctrl-C
 			{
				workbook.curDionGrid.copySelectedCells();
				preventDefaultBehaviour(event);
			}
			break;
			
		case 86:
			if (event.ctrlKey) // Ctrl-V
 			{
				if (_isCut) {
					workbook.curDionGrid.deleteSelectedCells(workbook.curDionGrid.copiedRange,"noShift");
					_isCut = false;
				}
				workbook.curDionGrid.pasteCopiedCells(workbook.curDionGrid.curXLSRowNo, workbook.curDionGrid.curXLSColNo);
				preventDefaultBehaviour(event);
			}
			break;
		case 88:	//Ctrl-X
			if (event.ctrlKey)
			{
				workbook.curDionGrid.borderMapper(workbook.curDionGrid.copiedRange,"noneDashed");
				workbook.curDionGrid.copiedRange = "";
				workbook.curDionGrid.copiedCellArray = [];
				_isCut = true;
				workbook.curDionGrid.copySelectedCells();
				preventDefaultBehaviour(event);
			}
			break;
		case 46: //Delete
			if (workbook.curDionGrid.mode == "select"){
				workbook.curDionGrid.deleteSelectedCells(workbook.selectedRange, "noShift");
				preventDefaultBehaviour(event);
			}
			break;
			
		case 13: //ENTER
			if (!event.shiftKey){
				var cell = workbook.curDionGrid.downEditableCell();
				mmRowNo = cell.rowNo;
				mmColNo = cell.colNo;
				workbook.curDionGrid.goToCell(mmRowNo, mmColNo);
			}
			else {
				var cell = workbook.curDionGrid.upEditableCell();
				mmRowNo = cell.rowNo;
				mmColNo = cell.colNo;
				workbook.curDionGrid.goToCell(mmRowNo, mmColNo);
			}
			preventDefaultBehaviour(event);
            break; 
		case 113: //F2
			if (workbook.curDionGrid.XLSCellArray[workbook.curDionGrid.curXLSRowNo][workbook.curDionGrid.curXLSColNo].readOnly == false)
				var editor = workbook.curDionGrid.displayCellArray[workbook.curDionGrid.curXLSRowNo][workbook.curDionGrid.curXLSColNo].editor;
				if (editor)
					editor.focus();
				else if (workbook.curDionGrid.isInXLSRange(workbook.curDionGrid.curXLSRowNo, workbook.curDionGrid.curXLSColNo)) {
					var XLSCellArrayObj = workbook.curDionGrid.XLSCellArray[workbook.curDionGrid.curXLSRowNo][workbook.curDionGrid.curXLSColNo];
					workbook.curDionGrid.insertWidget(XLSCellArrayObj.cellType,workbook.curDionGrid.curXLSRowNo,workbook.curDionGrid.curXLSColNo, XLSCellArrayObj);
					if (workbook.curDionGrid.displayCellArray[workbook.curDionGrid.curXLSRowNo][workbook.curDionGrid.curXLSColNo].editor)
						workbook.curDionGrid.displayCellArray[workbook.curDionGrid.curXLSRowNo][workbook.curDionGrid.curXLSColNo].editor.focus();
				}
				
					
				
			break;
		case 27:
			trapEscKeyFn(event);
			preventDefaultBehaviour(event);
			break;
		default :
			if (workbook.curDionGrid.mode != "edit"){
				if ((workbook.curDionGrid.isInXLSRange(workbook.curDionGrid.curXLSRowNo,workbook.curDionGrid.curXLSColNo)) && workbook.curDionGrid.XLSCellArray[workbook.curDionGrid.curXLSRowNo][workbook.curDionGrid.curXLSColNo].readOnly == false){
					workbook.curDionGrid.displayCellArray[workbook.curDionGrid.curXLSRowNo][workbook.curDionGrid.curXLSColNo].blur();
					if (workbook.curDionGrid.displayCellArray[workbook.curDionGrid.curXLSRowNo][workbook.curDionGrid.curXLSColNo].editor)
						workbook.curDionGrid.displayCellArray[workbook.curDionGrid.curXLSRowNo][workbook.curDionGrid.curXLSColNo].editor.focus();
				
				}	
			}
			if (workbook.copiedCellArray.length > 0){
				workbook.clearCopiedRange();
				_isCut = false;
			}
    }
}


function addEvent(hoo,wot,fun){
    if(document.addEventListener) hoo.addEventListener(wot,fun,null);
    else if(document.attachEvent) hoo.attachEvent('on'+wot,fun); 
}

function startcatching(){
    addEvent(document,'keypress',trapEscKeyFn);
}

function trapEscKeyFn(event) {
    if (event == null) var event = window.event; 
	//if (workbook.curDionGrid.mode != "edit") return;
	if(event.keyCode==27){ 
		//TODO Formulas : is dirty is not set to false
		var curXLSRowNo = workbook.curDionGrid.curXLSRowNo;
		var curXLSColNo = workbook.curDionGrid.curXLSColNo;
		
		if (workbook.curDionGrid.isInXLSRange(curXLSRowNo,curXLSColNo)){
			workbook.curDionGrid.refreshCell(curXLSRowNo, curXLSColNo);
			if (workbook.curDionGrid.displayCellArray[curXLSRowNo][curXLSColNo].editor)
				workbook.curDionGrid.displayCellArray[curXLSRowNo][curXLSColNo].editor.blur();
			workbook.curDionGrid.refreshCell(curXLSRowNo, curXLSColNo);
		}
		workbook.curDionGrid.mode = "select";
		preventDefaultBehaviour(event);
	}
	return true;
}

function findPosX(obj)
  {
    var curleft = 0;
    if(obj.offsetParent)
        while(1) 
        {
          curleft += obj.offsetLeft;
          if(!obj.offsetParent)
            break;
          obj = obj.offsetParent;
        }
    else if(obj.x)
        curleft += obj.x;
    return curleft;
  }

function findPosY(obj){
	var curtop = 0;
	if(obj.offsetParent)
		while(1)
		{
		  curtop += obj.offsetTop;
		  if(!obj.offsetParent)
			break;
		  obj = obj.offsetParent;
		}
	else if(obj.y)
		curtop += obj.y;
	return curtop;
}

function showErrorNotification(message, div){
}
function hideErrorNotification(message, div){
}

function log(message, obj){
	/*if (console)
		console.log(message, obj);*/
}

function addInsertMultipleRows(contextMenuDiv, GRID, row, label){
	var insertRow = new contextMenuItem();
	insertRow.setParentDivId(contextMenuDiv.contextMenuId);
	insertRow.create("InsertRow");
	insertRow.setLabel(label);
	insertRow.setAction(function() {
		GRID.hideContextMenu();
		var range = workbook.selectedRange;
		rangeArray = range.split(":");
		var startRange = parseInt(rangeArray[0]);
		var endRange = parseInt(rangeArray[2]);
		GRID.insertRows(row, Math.abs(startRange-endRange) + 1);
	});
	contextMenuDiv.addItem(insertRow);
}

function addCopyRange(contextMenuDiv, GRID, label){
	var copyRow = new contextMenuItem();
	copyRow.setParentDivId(contextMenuDiv.contextMenuId);
	copyRow.create("CopyRow");
	copyRow.setLabel(label);
	copyRow.setAction(function() {
		workbook.curDionGrid.copySelectedCells();
	});
	contextMenuDiv.addItem(copyRow);
}

function addDeleteMultipleRows(contextMenuDiv, GRID, row, label){
	var deleteRow = new contextMenuItem();
	deleteRow.setParentDivId(contextMenuDiv.contextMenuId);
	deleteRow.create("DeleteRow");
	deleteRow.setLabel(label);
	deleteRow.setAction(function() {
		GRID.hideContextMenu();
		var range = workbook.selectedRange;
		rangeArray = range.split(":");
		var startRange = parseInt(rangeArray[0]);
		var endRange = parseInt(rangeArray[2]);
		GRID.deleteRows(row, Math.abs(startRange-endRange) + 1);
	});
	contextMenuDiv.addItem(deleteRow);
}

function addPasteRange(contextMenuDiv, GRID, row, col, label){
	var pasteRow = new contextMenuItem();
	pasteRow.setParentDivId(contextMenuDiv.contextMenuId);
	pasteRow.create("PasteRow");
	pasteRow.setLabel(label);
	pasteRow.setAction(function() {
		GRID.hideContextMenu();
		GRID.pasteValues(row, col);
	});
	contextMenuDiv.addItem(pasteRow);
}



//Context Menu ShortCuts
function addInsertRow(contextMenuDiv, GRID, row){
	var insertRow = new contextMenuItem();
	insertRow.setParentDivId(contextMenuDiv.contextMenuId);
	insertRow.create("InsertRow");
	insertRow.setLabel(gettext("Insert Tenor"));
	insertRow.setAction(function() {
		GRID.hideContextMenu();
		GRID.insertRows(row, 1);
	});
	contextMenuDiv.addItem(insertRow);
}


function addCopyRow(contextMenuDiv, GRID, row){
	var copyRow = new contextMenuItem();
	copyRow.setParentDivId(contextMenuDiv.contextMenuId);
	copyRow.create("CopyRow");
	copyRow.setLabel(gettext("Copy Row"));
	copyRow.setAction(function() {
		GRID.hideContextMenu();
		GRID.copyRow(row);
	});
	contextMenuDiv.addItem(copyRow);
}

function addDeleteRow(contextMenuDiv, GRID, row){
	var deleteRow = new contextMenuItem();
	deleteRow.setParentDivId(contextMenuDiv.contextMenuId);
	deleteRow.create("DeleteRow");
	deleteRow.setLabel(gettext("Delete Row"));
	deleteRow.setAction(function() {
		GRID.hideContextMenu();
		GRID.deleteRows(row, 1);
	});
	contextMenuDiv.addItem(deleteRow);
}

function addPasteRow(contextMenuDiv, GRID, row){
	var pasteRow = new contextMenuItem();
	pasteRow.setParentDivId(contextMenuDiv.contextMenuId);
	pasteRow.create("PasteRow");
	pasteRow.setLabel(gettext("Paste Row"));
	pasteRow.setAction(function() {
		GRID.hideContextMenu();
		GRID.paste(row, 1);
	});
	contextMenuDiv.addItem(pasteRow);
}

function indexOfElement(array, element){
	for (var i = 0; i < array.length; i++){
		if (array[i] == element) return i;
	}
	return -1;
};


function indexOfElementInArray(array, element){
	for (var i = 0; i < array.length; i++){
		if (array[i] == element) return i;
	}
	return -1;
};

function getNameOfMonth(monthStr){
	switch(monthStr) {
		case "1" : return "Jan";
		case "2" : return "Feb";
		case "3" : return "Mar";
		case "4" : return "Apr";
		case "5" : return "May";
		case "6" : return "Jun";
		case "7" : return "Jul";
		case "8" : return "Aug";
		case "9" : return "Sep";
		case "10" : return "Oct";
		case "11" : return "Nov";
		case "12" : return "Dec";
		case "01" : return "Jan";
		case "02" : return "Feb";
		case "03" : return "Mar";
		case "04" : return "Apr";
		case "05" : return "May";
		case "06" : return "Jun";
		case "07" : return "Jul";
		case "08" : return "Aug";
		case "09" : return "Sep";
	}
}

function getMonthFromName(monthStr){
	monthStr = "" + monthStr;
	monthStr = monthStr.toLowerCase();
	switch(monthStr) {
		case "jan" : return "01";
		case "feb" : return "02";
		case "mar" : return "03";
		case "apr" : return "04";
		case "may" : return "05";
		case "jun" : return "06";
		case "jul" : return "07";
		case "aug" : return "08";
		case "sep" : return "09";
		case "oct" : return "10";
		case "nov" : return "11";
		case "dec" : return "12";
		default : return monthStr; 
	}
}

function getDateFromText(dateString){
	var m = dateString.match(/^[1-9][0-9]*[dwmyDWMY]$/);
	if(m && m.length == 1) {
		return dateString;
	} 
	else {
		dateString = dateString.replace(/^[ ]|[ ]$/g,"");
		dateString = dateString.replace(" ", "-");
		dateString = dateString.replace(" ", "-");
		dateString = dateString.replace("/", "-");
		dateString = dateString.replace("/", "-");
		dateString = parseDate(dateString);
		var dateMs = Date.parse(dateString);
		if (dateMs){
			var dateObj = new Date(dateMs);
			var month = dateObj.getMonth() + 1;
			var day = dateObj.getDate();
			var year = dateObj.getFullYear();
			var formattedDate = (month + "/" + day + "/"+ year);
			return formattedDate;
		}
		else return NaN;
	}
}

function parseDate(dateStr){
	var dateArr = dateStr.split("-");
	if (dateArr.length != 3) return dateStr;
	
	if (dateArr[2].length == 2){
		dateArr[2] = "20" + dateArr[2];
	}

	if (dateArr[1].length == 3)
		return getMonthFromName(dateArr[1]) + "/" + dateArr[0] + "/" + dateArr[2];
	else if (dateArr[0].length == 3)
		return dateArr[1] + "/" + getMonthFromName(dateArr[0]) + "/" + dateArr[2];
	else {
		return dateArr[0] + "/" + dateArr[1] + "/" + dateArr[2];
	}
}

//utility functions
function date_str_to_time(str) {
    if ((typeof str) != (typeof "asd")) return str; 
	var dd = str.split("/");
	if (dd.length == 3){
		var dateObj = new Date(dd[2], dd[0]-1, dd[1]);
		if (dateObj)
			return dateObj.getTime();
    }
    return str;
}

function date_to_str(millis) {
    if ((typeof millis) != (typeof 1)) return millis;
    if (millis == 0) return "";
	var d = new Date(millis - 0);
	var st = (d.getMonth()+1)+"/"+d.getDate()+"/"+d.getFullYear();
	return st;
}

function validateDate(dateString){
	dateString = "" + dateString;
	var m = dateString.match(/^[1-9][0-9]*[dwmyDWMY]$/);
	if(m && m.length == 1) {
		return true;
	} 
	else {
		dateString = dateString.replace(/^[ ]|[ ]$/g,"");
		dateString = dateString.replace(" ", "-");
		dateString = dateString.replace(" ", "-");
		dateString = dateString.replace("/", "-");
		dateString = dateString.replace("/", "-");
		
		var dateArr = dateString.split("-");
		if (dateArr.length != 3) return false;

		if (getNumberType(dateArr[0]) != "integer" || (!(getNumberType(dateArr[1]) == "integer" || getNumberType(getMonthFromName(dateArr[1])) == "integer"))|| getNumberType(dateArr[2]) != "integer"){
			return false;
		}
		
		dateString = parseDate(dateString);
		var dateMs = Date.parse(dateString);
		if (dateMs){
			return true
		}
		else return false;
	}
}
