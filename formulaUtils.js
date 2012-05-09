/*
 *   copyright (c) -- chaitanya nalla and kechit goyal
 *
 */

﻿function isOperator(mChar) {
    switch (mChar) {
        case "(" :
        case ")" :
        case "^" :
        case "*":
        case "/":
        case "%":
        case "+" :
        case "-" :
        case "," :
		case ">" :
		case "<" :
		case "=" :
            return true;
        default :
            return false;
     }
}

function operatorPrecedence(mChar) {
    switch (mChar) {
        case "^":
            return 14;
        case "*":
        case "/":
        case "%":
            return 13;
        case "+" :
        case "-" :
            return 12;
		case ">" :
		case "<" :
		case "=" :
			return 11;
        default :
            return -1;
     }
}

function operatorAssociativity(mChar) {
    switch (mChar) {
        case "^":
            return "R";
        default :
            return "L";      
     }
}

function isFunction(mToken) {
    if (isNaN(mToken) == false) {
        return false;
    }
	if (typeof(mToken) != "string"){
		return false;
	}
    var mStr;
    mStr = mToken.toUpperCase();
    switch(mStr) {
        case "MATHMIN":    
        case "MATHMAX":    
        case "MATHSQRT":    
        case "MATHSQUARE":
		case "IF":
		case "SUM":
            return true;
        default:
            return false;           
    }

}

//This function returns argument count for an operator
function operatorArgCount(mToken) {
    var mStr = mToken.toUpperCase();
    switch(mStr) {
		case "+":    
        case "-":  
		case "*":    
        case "/":
		case "^":    
        case ">":
		case "<":    
        case "=": 	
            return 2;
        default:
            return -1;
    }
}


//This function returns argument count for a function
function functionArgCount(mToken) {
    var mStr = mToken.toUpperCase();
    switch(mStr) {
		case "IF":
			return 3;
        case "MATHSQRT":    
        case "MATHSQUARE":    
            return 1;
		case "MATHMIN":    
        case "MATHMAX":
		case "SUM":
            return 0; 
        default:
            return -1;
    }
}

function mathMin(mArray) {
    var result = mArray[0];
	for (var i = 0; i < mArray.length; i++){
		if (mArray[i] < result) result = mArray[i];
	}
	return result;
}


function mathMax(mArray) {
    var result = mArray[0];
	for (var i = 0; i < mArray.length; i++){
		if (mArray[i] > result) result = mArray[i];
	}
	return result; 
}

function sum(mArray) {
    var result = mArray[0];
	for (var i = 1; i < mArray.length; i++){
		result += mArray[i];
	}
	return result;
}

function mathSqrt(mNo) {
    return Math.pow(mNo, 1/2);
}

function mathSquare(mNo) {
    return mNo * mNo;
}

//This Function checks whether  opening and closing brackets are matched or not
function bracketMatch (mString){
	var count = 0;
	for (var i = 0; i < mString.length; i++) {
		mChar = mString.substr(i, 1);
		if (mChar == "(") count++;
		else if (mChar == ")") count --;
	}
	return count;
}

//This Function returns true if the string is true
//For any other string value it returns false
//For formula with IF, for any arbit string it will take the value as false (Same as Excel)
function toBoolean (mString){
	mString = "" + mString;
	if(mString.toUpperCase() == "TRUE"){
		return true;
	}
	else if (mString.toUpperCase() == "FALSE") 
		return false;
	else return null;
}

//Modify Formula handles double quotes and passes control to removeQuotes
function modifyFormula(myStr, rowDiff, colDiff){
	myStr = "" + myStr;
	if (myStr.split("\"").length % 2 == 0)return myStr;
	if (myStr.split("\'").length % 2 == 0)return myStr;
	
	var arr = myStr.split("\"");
	var result = "";
	for (var i = 0; i < arr.length; i= i+2){
		result += removeQuotes(arr[i], rowDiff, colDiff);
		if (i+1< arr.length) result += ("\"" + arr[i+1] + "\"");
	}
	return result;
}

function simpleModify(val, rowShift, colShift)
{
	var char =val.charAt(0);
	char --;
	val[0] = char;
	return val;  
	  
}

//removeQuotes handles quotes and passes control to modify
function removeQuotes(myStr, rowDiff, colDiff){
	myStr = "" + myStr;
	var arr = myStr.split("\'");
	var result = "";
	for (var i = 0; i < arr.length; i= i+2){
		result += modify(arr[i],rowDiff,colDiff);
		if (i+1< arr.length) result += ("\'" + arr[i+1] + "\'");
	}
	return result;
}

//Modify checks whether the string matches the cell regular expression
//For all the matches it calls a helper function which adjusts the values accordingly
function modify(myStr, i, j){
	myStr = "" + myStr;
	var cellRegEx = /[\$]{0,1}[a-zA-Z]{1,3}[\$]{0,1}[1-9]{1,1}[0-9]{0,3}/;
	var result = "";
	while (myStr.match(cellRegEx) != null){
		result += myStr.substr(0,myStr.match(cellRegEx).index) + "" + modifyFormulaHelper(myStr.match(cellRegEx)[0], i , j);
		myStr = myStr.substr (myStr.match(cellRegEx).index + myStr.match(cellRegEx)[0].length);
	}
	result += myStr;
	return result;
	
}

//Modify Helper changes a cell string and returns a modified string according to shift parameters i and j
//$ is handled by this function
function modifyFormulaHelper(myStr, i, j){
	myStr = "" + myStr;
	var cellRegExBoth = /[\$]{1,1}[a-zA-Z]{1,3}[\$]{1,1}[1-9]{1,1}[0-9]{0,3}/;
	var cellRegExCol = /[\$]{1,1}[a-zA-Z]{1,3}[1-9]{1,1}[0-9]{0,3}/;
	var cellRegExRow = /[a-zA-Z]{1,3}[\$]{1,1}[1-9]{1,1}[0-9]{0,3}/;
	var cellRegExNone = /[a-zA-Z]{1,3}[1-9]{1,1}[0-9]{0,3}/;
	var rowRegEx = /[1-9]{1,1}[0-9]{0,3}/;
	var colRegEx = /[a-zA-Z]{1,3}/;
	if (myStr.match(cellRegExBoth) != null) return myStr;
	else if (myStr.match(cellRegExCol) != null) {
		return  "$" + myStr.match(colRegEx)[0] + ((parseInt(myStr.match(rowRegEx)[0])) + i);
	}
	else if (myStr.match(cellRegExRow) != null) {
		return getColLabelFromNumber(getColNoFromLabel(myStr.match(colRegEx)[0])+j) +"$"+ myStr.match(rowRegEx)[0];
	}
	else return getColLabelFromNumber(getColNoFromLabel(myStr.match(colRegEx)[0])+j) + ((parseInt(myStr.match(rowRegEx)[0])) + i);
}

//This checks whether the string is a cell or not
function isCell(myStr){
	if (myStr == undefined || myStr == null) return false;
	var cellRegExBoth = /^[\$]{0,1}[a-zA-Z]{1,3}[\$]{0,1}[1-9]{1,1}[0-9]{0,3}$/;
	return ((""+myStr).match(cellRegExBoth) != null);
}

//This checks whether the string represents a range or not
function isRange(myStr){
	if (myStr == undefined || myStr == null) return false;
	myStr = "" + myStr;
	var cellRegExBoth = /^[\$]{0,1}[a-zA-Z]{1,3}[\$]{0,1}[1-9]{1,1}[0-9]{0,3}:[\$]{0,1}[a-zA-Z]{1,3}[\$]{0,1}[1-9]{1,1}[0-9]{0,3}$/;
	return ((""+myStr).match(cellRegExBoth) != null);
}
function isDeleteCell(rowNo,colNo,affectedRange,rowShift,colShift){
	
	var rangeArray = affectedRange.split(":");
	var startRow = parseInt(rangeArray[0]);
	var startCol =parseInt(rangeArray[1]);
	var endRow = parseInt(rangeArray[2]);
	var endCol = parseInt(rangeArray[3]);
	if(rowShift < 0 && colShift < 0 ) return false;
	if(rowShift < 0)
		return (rowNo === startRow);
	if(colShift < 0)
		return (colNo === startCol);
	
	return false;
	
}
//This checks whether the string represents a num range or not
function isNumRange(myStr){
	if (myStr == undefined || myStr == null) return false;
	myStr = "" + myStr;
	var cellRegExBoth = /^[1-9]{1,1}[0-9]{0,3}:[1-9]{1,1}[0-9]{0,3}:[1-9]{1,1}[0-9]{0,3}:[1-9]{1,1}[0-9]{0,3}$/;
	return ((""+myStr).match(cellRegExBoth) != null);
}

function isGridCell(myStr){
	if (myStr == undefined || myStr == null) return false;
	myStr = "" + myStr;
	var sheetCellArray = myStr.split("."); 
	if (sheetCellArray.length != 2) return false;
	var cellRegExBoth = /^[\$]{0,1}[a-zA-Z]{1,3}[\$]{0,1}[1-9]{1,1}[0-9]{0,3}$/;
	return ((""+sheetCellArray[1]).match(cellRegExBoth) != null);
}

function isGridRange(myStr){
	if (myStr == undefined || myStr == null) return false;
	myStr = "" + myStr;
	var sheetCellArray = myStr.split("."); 
	if (sheetCellArray.length != 2) return false;
	var cellRegExBoth = /^[\$]{0,1}[a-zA-Z]{1,3}[\$]{0,1}[1-9]{1,1}[0-9]{0,3}:[\$]{0,1}[a-zA-Z]{1,3}[\$]{0,1}[1-9]{1,1}[0-9]{0,3}$/;
	return ((""+sheetCellArray[1]).match(cellRegExBoth) != null);
}

function notionalValue(mString){
	mString = removePreceedingZeros(mString);
	if (!mString) return 0;
	mString = ""+mString;
	if (getNumberType(mString)=="integer"){
		return parseInt(mString);
	}
	else if (getNumberType(mString)=="float"){
		return parseFloat(mString);
	}
	else {
		var mString = mString.toUpperCase();
		var thousandRegex = /^[0-9]+[\.]{0,1}[0-9]*[K]$/;
		var millionRegex = /^[0-9]+[\.]{0,1}[0-9]*[M]$/;
		var billionRegex = /^[0-9]+[\.]{0,1}[0-9]*[B]$/;
		var trillionRegex = /^[0-9]+[\.]{0,1}[0-9]*[T]$/;
		if (mString.match(thousandRegex))
			return getNumVal(mString.substr(0,mString.length-1))*1000;
		else if (mString.match(millionRegex))
			return getNumVal(mString.substr(0,mString.length-1))*1000000;
		else if (mString.match(billionRegex))
			return getNumVal(mString.substr(0,mString.length-1))*1000000000;
		else if (mString.match(trillionRegex))
			return getNumVal(mString.substr(0,mString.length-1))*1000000000000;
		else return mString;
	}
}

function getNumVal (mString){
	if (!mString) return 0;
	mString = ""+mString;
	if (getNumberType(mString)=="integer"){
		return parseInt(mString);
	}
	else if (getNumberType(mString)=="float"){
		return parseFloat(mString);
	}
	else return "";
}



