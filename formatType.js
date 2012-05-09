/*
 *   copyright (c) -- chaitanya nalla and kechit goyal
 *
 */

function defineFormatType(){
	formatType = function (format){
		this.format = format;
		if (format == "integer"){
			this.decimalCount = 0;
			this.useThousandSeparator = false;
		}
		else if (format == "float"){
			this.decimalCount = 2;
			this.useThousandSeparator = false;
		}
		else if (format == "currency"){
			this.decimalCount = 2;
			this.useThousandSeparator = true;
			this.symbol = "$";
		}
		else if (format == "accounting"){
			this.decimalCount = 2;
			this.useThousandSeparator = true;
			this.symbol = "";
		}
		else if (format == "percent"){
			this.decimalCount = 0;
		}
	};
	
	formatType.prototype.copy = function(){
		var newFormat = new formatType(this.format);
		if (this.format == "integer"){
			newFormat.decimalCount = this.decimalCount;
			newFormat.useThousandSeparator = this.useThousandSeparator;
		}
		else if (this.format == "float"){
			newFormat.decimalCount = this.decimalCount;
			newFormat.useThousandSeparator = this.useThousandSeparator;
		}
		else if (this.format == "currency"){
			newFormat.decimalCount = this.decimalCount;
			newFormat.useThousandSeparator = this.useThousandSeparator;
			newFormat.symbol = this.symbol;
		}
		else if (this.format == "accounting"){
			newFormat.decimalCount = this.decimalCount;
			newFormat.useThousandSeparator = this.useThousandSeparator;
			newFormat.symbol = this.symbol;
		}
		else if (this.format == "percent"){
			newFormat.decimalCount = this.decimalCount;
		}
		return newFormat;
	
	};
}
defineFormatType();
