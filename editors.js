/*
 *   copyright (c) -- chaitanya nalla and kechit goyal
 *
 */

var FinsheetEditors = {

	InputCellFormatter : function(args) {
		if (args.item) {
			var className = "formatter_txt backgroundNone";
			var $obj = $("<div class='" + className + "'>"
					+ args.item.displayValue + "</div>");
			$obj.appendTo(args.container);
		}
	},

	textLabelFormatter : function(args) {
		if (args.item) {
			var className = "";
			var $label = $("<label class='" + className + "'>"
					+ args.item.valueOptions.value + "</label>");
			$label.appendTo(args.container);
		}

	},

	DeltaCellFormatter : function(args) {
		if (args.item) {
			var $input = $("<div class='cellStyle paddingTop allGrayBorder'>");
			var $text = $("<div class='formatter_txt'>");

			$input.html(args.item.valueOptions.value.toString());
			$text.html(args.item.valueOptions.text.toString());

			$input.width(30);
			$input.height(getLengthFromString(args.container.style.height) - 4);

			$text.width(getLengthFromString(args.container.style.width) - 32);
			$text.height(getLengthFromString(args.container.style.height));

			$text.css("left", 32);

			$input.appendTo(args.container);
			$text.appendTo(args.container);
		}
	},

	CheckBoxCellEditor : function(args) {
		var $checkboxdiv;
		var $input;
		var $textDiv;
		var defaultValue;

		this.init = function() {
			var THIS = this;
			$checkboxdiv = $("<div class='checkboxDiv'/>");
			$input = $("<input type='checkbox' class='checkbox'></input>");
			$textDiv = $("<div class='checkboxTextDiv'></div>");
			$checkboxdiv.append($input);
			$checkboxdiv.append($textDiv);
			// $checkboxdiv.appendTo(args.container);
			$checkboxdiv.appendTo(args.container);
			// $input.appendTo(args.container);

			$checkboxdiv.keydown(function(e) {
				var value;
				if (window.event)
					value = e.keyCode;
				else
					value = e.which;

				if (value == 13) {
					if ($input.is(':checked'))
						$input.attr('checked', false);
					else
						$input.attr("checked", true);
					preventDefaultBehaviour(e);
					preventPropagation(e);
				}
			});

			if (args.item)
				this.loadValue(args.item);

			$input.click(function() {
			
				
				if($input.is(':checked'))
					$input.attr('checked',false);
				else
					$input.attr("checked",true);
			
				args.grid.saveOneScreenCellInMemory(args.rowNo, args.colNo);
			});
			
			$checkboxdiv.click(function(){
				
				if($input.is(':checked'))
					$input.attr('checked',false);
				else
					$input.attr("checked",true);
					
				args.grid.saveOneScreenCellInMemory(args.rowNo,args.colNo);
				
			});
			$input.focus(function(){
				args.grid.mode = "select";
			
			});
		
		};

		this.destroy = function() {

		};

		this.loadValue = function(item) {
			var THIS = this;

			if (item.valueOptions == null)
				return;
			var valueOptions = item.valueOptions;

			var value = valueOptions.value;
			var text = valueOptions.text;
			this.defaultValue = value;
			if (value === true)
				$input.attr("checked", true);
			else
				$input.attr("checked", false);
			$textDiv.text(text);
		};

		this.applyValue = function(item) {
			item.valueOptions = {
				value : $input.is(':checked'),
				text : $checkboxdiv.html()
			};
			item.updateCellByWidget();
			this.defaultValue = item.valueOptions.value;
		};

		this.getValue = function() {
			return $input.is(':checked');
		};

		this.isValueChanged = function() {
			var isModified = (!(!$input.is(':checked') && this.defaultValue == null))
					&& ($input.is(':checked') != this.defaultValue);
			return isModified;
		};

		this.focus = function() {
			$input.focus();
			args.grid.mode = "select";
		};

		this.blur = function() {
			args.grid.mode = "none";
		};

		this.autofocus = function() {
			return true;
		};

		this.position = function(position) {
		};

		this.setBackground = function(value) {
		};

		this.setClass = function(className) {
		};

		this.init();
	},

	CurrencyPairEditor : function(args) {
		var $currency;
		var $toggle;
		this.defaultValue = "";
		this.rowNo = args.rowNo;
		this.colNo = args.colNo;
		var currency = [ "AUD / USD", "EUR / USD", "GBP / USD", "NZD / USD",
				"USD / CAD", "USD / CHF", "USD / JPY" ];

		this.init = function() {
			var GRID = args.grid;
			var THIS = this;
			$currency = $("<INPUT type='text' class='cellStyle' />");
			$toggle = $("<div class = 'currencyToggle'>&nbsp;</div>") // toggle

			$toggle.appendTo(args.container); // appending to container
			$currency.appendTo(args.container);
			$currency.autocomplete(currency, {
				minChars : 0,
				max : 110,
				autoFill : false,
				mustMatch : false,
				matchContains : false,
				scrollHeight : 220,
				width : getLengthFromString(args.container.style.width)
			});

			$currency
					.width(getLengthFromString(args.container.style.width) - 27);
			$toggle.width(21);
			$currency.height(getLengthFromString(args.container.style.height));

			var clickFn = function() {
				var str = $currency.val();
				var arr = str.split("/");
				if (arr.length !== 2)
					return;
				$currency.val($.trim(arr[1]) + " / " + $.trim(arr[0]));
				GRID.saveOneScreenCellInMemory(THIS.rowNo, THIS.colNo);
				preventPropagation(event);
			}
			$toggle.click(clickFn);
			$toggle.dblclick(clickFn);

			var currencyKeyDown = function(event) {
				if (event == null)
					event = window.event;
				switch (event.keyCode) {
				case 33:
				case 34:
				case 37:
				case 38:
				case 39:
				case 40:
					GRID.mode = "none";
					break;

				default:
					GRID.mode = "edit";
					break;
				}
			};

			$currency.keydown(currencyKeyDown);

			var changeFn = function() {
				GRID.saveOneScreenCellInMemory(THIS.rowNo, THIS.colNo);
			};
			var focusFn = function() {
				GRID.mode = "none";
			};
			var blurFn = function() {
				GRID.mode = "none";
				GRID.saveOneScreenCellInMemory(THIS.rowNo, THIS.colNo);
			};

			$currency.bind("change", changeFn);

			$currency.focus(focusFn);

			$currency.dblclick(function(e) {
				$currency.focus();
				$currency.select();
				preventPropagation(e);
			});
			$currency.blur(blurFn);
			if (args.item)
				this.loadValue(args.item);
		};

		this.destroy = function() {

		};
		this.focus = function() {
			$currency.focus();
			$currency.select();
		};

		this.blur = function() {
			$currency.blur();
		};

		this.autofocus = function() {
			return false;
		};

		this.loadValue = function(item) {
			if (item.valueOptions == null)
				return;

			var displayVal = args.item.valueOptions.value;
			if (this.validate(args.item.valueOptions.value)) {
				displayVal = args.item.valueOptions.value.toString().substr(0,
						3)
						+ " / "
						+ args.item.valueOptions.value.toString().substr(3, 3);
			}

			this.defaultValue = displayVal;
			$currency.val(displayVal);
			this.rowNo = item.parentRow.rowNo;
			this.colNo = item.parentCol.colNo;
		};

		this.serializeValue = function() {
			return $currency.val();
		};

		this.applyValue = function(item) {
			var value = $currency.val().toUpperCase().replace("/", "");

			while (value.indexOf(" ") >= 0) {
				value = value.replace(" ", "");
			}
			item.valueOptions = {
				value : value,
			};
			item.updateCellByWidget();
			this.defaultValue = item.valueOptions.value;

			if (this.validate($currency.val())) {
				item.errorMesg = "";
			} else {
				item.errorMesg = "Incorrect Input";
			}
		};

		this.getValue = function() {
			return $currency.val();
		};

		this.isValueChanged = function() {
			var isModified = (!($currency.val() == "" && this.defaultValue == null) && ($currency
					.val() != this.defaultValue));
			return isModified;
		};

		this.validate = function(strValue) {
			var value = strValue.toUpperCase().replace("/", "");
			while (value.indexOf(" ") >= 0) {
				value = value.replace(" ", "");
			}
			if (value.length != 6)
				return false;
			var rev_value = value.substr(3, 3) + " / " + value.substr(0, 3);
			value = value.substr(0, 3) + " / " + value.substr(3, 3);
			for ( var i = 0; i < currency.length; i++) {
				if (currency[i] == value || currency[i] == rev_value)
					return true;
			}
			return false;
		};
		this.setBackground = function(value) {
			$currency.css("background", value);

		};

		this.setClass = function(className) {
			// $input.attr("class", className);
		};

		this.init();

	},

	DeltaEditor : function(args) {
		var $input = null;
		var $text = null;

		this.init = function() {
			var GRID = args.grid;
			var THIS = this;
			$input = $("<INPUT type='text' class='gridCellBorder'>");
			$text = $("<div class='formatter_txt' />");
			$input.appendTo(args.container);
			$text.appendTo(args.container);

			$input.width(30);
			$text
					.width((getLengthFromString(args.container.style.width) - 2) - 32);
			$text.css("left", 32);
			$input.height(getLengthFromString(args.container.style.height) - 2);
			$text.height(getLengthFromString(args.container.style.height));

			var changeFn = function() {
				GRID.saveOneScreenCellInMemory(THIS.rowNo, THIS.colNo);
			};

			$input.bind("change", changeFn);
			if (args.item)
				this.loadValue(args.item);
		};

		this.focus = function() {
			args.grid.mode = "edit";
			$input.focus();
			$input.select();
		};

		this.blur = function() {
			args.grid.mode = "none";
			$input.blur();
		};

		this.autofocus = function() {
			return false;
		};

		this.destroy = function() {

		};
		this.loadValue = function(item) {
			if (item.valueOptions == null)
				return;
			$input.val(item.valueOptions.value);
			$text.html(item.valueOptions.text);
			this.defaultValue = item.valueOptions.value;
		};

		this.applyValue = function(item) {
			item.valueOptions = {
				value : $input.val(),
				text : $text.html()
			};
			item.updateCellByWidget();
			this.defaultValue = item.valueOptions.value;
		};

		this.getValue = function() {
			return $input.val() + " " + $text.html();
		};

		this.isValueChanged = function() {
			var isModified = ((!($input.val() == "" && this.defaultValue == null)) && ($input
					.val() != this.defaultValue));
			return isModified;
		};

		this.validate = function() {
			return true;
		};
		this.setBackground = function(value) {
			$input.css("background", value);
		};

		this.init();
	},

	ToggleEditor : function(args) {
		var $toggleDiv;
		var $input;
		var $img;
		this.defaultValue = "";
		this.rowNo = args.rowNo;
		this.colNo = args.colNo;
		this.possibleOptions = [];
		this.currentIndex = 0;

		this.init = function() {
			var GRID = args.grid;
			var THIS = this;
			$toggleDiv = $("<div class='toggle' toggleIndex='0'/>");
			$input = $("<div class='toggleText' toggleInde='0'x/>");
			$img = $("<div class='toggleImg' toggleIndex='0'>");

			$toggleDiv.appendTo(args.container);
			$input.appendTo($toggleDiv);
			$img.appendTo($toggleDiv);

			$toggleDiv.width((getLengthFromString(args.container.style.width)));
			$toggleDiv.height("100%");

			$input
					.width((getLengthFromString(args.container.style.width) - 16));
			$input.height("100%");

			$img.css("left",
					(getLengthFromString(args.container.style.width) - 16))
			$img.width(16);
			$img.height("100%");

			var clickFn = function() {
				if (THIS.possibleOptions.length > 0) {
					THIS.currentIndex = (THIS.currentIndex + 1)
							% THIS.possibleOptions.length;
					var text = THIS.possibleOptions[THIS.currentIndex].text;
					$input.html(text);
					GRID.saveOneScreenCellInMemory(THIS.rowNo, THIS.colNo);

				}
			};

			$toggleDiv.click(clickFn);

			var toggleDivKeyDown = function(event) {
				if (event == null)
					event = window.event;
				switch (event.keyCode) {
				case 13: // Enter key
					GRID.mode = "none";
					clickFn();
					break;
				default:
					GRID.mode = "select";
					break;
				}
			};

			$toggleDiv.keydown(toggleDivKeyDown);
			$input.keydown(toggleDivKeyDown);
			$toggleDiv.focus();
			if (args.item)
				this.loadValue(args.item);
		};

		this.destroy = function() {
			this.blur();
		};
		this.focus = function() {
			args.grid.mode = "edit";
			$img[0].className = "toggleImgFocus";
			$toggleDiv.focus();
		};

		this.blur = function() {
			args.grid.mode = "none";
			$img[0].className = "toggleImg";
			$toggleDiv.blur();
		};

		this.autofocus = function() {
			return true;
		};

		this.loadValue = function(item) {
			if (item.valueOptions == null)
				return;
			this.rowNo = item.parentRow.rowNo;
			this.colNo = item.parentCol.colNo;
			var valueOptions = item.valueOptions.possibleOptions;
			this.possibleOptions = [];
			var curIndex = 0;
			for ( var i in valueOptions) {
				if (i == item.valueOptions.value) {
					this.currentIndex = curIndex;
				}
				this.possibleOptions[this.possibleOptions.length] = {
					text : valueOptions[i],
					value : i
				};
				curIndex++;
			}
			var text = this.possibleOptions[this.currentIndex].text;
			$input.html(text);
			this.defaultValue = text;
		};

		this.serializeValue = function() {
			return $input.html();
		};

		this.applyValue = function(item) {
			var possibleOptions = this.possibleOptions;
			var valueOptions = {};
			for ( var i in possibleOptions) {
				valueOptions[possibleOptions[i].value] = possibleOptions[i].text;
			}
			item.valueOptions = {
				possibleOptions : valueOptions,
				value : possibleOptions[this.currentIndex].value,
				currentIndex : this.currentIndex
			};
			item.updateCellByWidget();
			this.defaultValue = $input.html();
		};

		this.getValue = function() {
			return $input.html();
		};

		this.isValueChanged = function() {
			var isModified = ((!($input.html() == "" && this.defaultValue == null)) && ($input
					.html() != this.defaultValue));
			return isModified;
		};

		this.validate = function() {
			return {
				valid : true,
				msg : null
			};
		};
		this.setBackground = function(value) {
			// $input.css("background", value);
		};

		this.setClass = function(className) {
			// $input.attr("class", className);
		};

		this.init();

	},

	ClassToggleEditor : function(args) {
		var $toggleDiv = null;
		var action = null;

		this.init = function() {
			$toggleDiv = $("<p class='curveTypeLink' />");
			$toggleDiv.appendTo(args.container);
			$toggleDiv
					.width((getLengthFromString(args.container.style.width)) - 2);
			$toggleDiv.height(getLengthFromString(args.container.style.height));
			if (args.item)
				this.loadValue(args.item);
		};

		this.destroy = function() {

		};
		this.focus = function() {
			$toggleDiv.focus();
		};

		this.blur = function() {
			$toggleDiv.blur();
		};

		this.autofocus = function() {
			return false;
		};

		this.loadValue = function(item) {
			if (item.valueOptions == null)
				return;
			this.rowNo = item.parentRow.rowNo;
			this.colNo = item.parentCol.colNo;
			$toggleDiv.html(item.value);
			this.defaultValue = item.value;
			$toggleDiv.click(function() {
				item.valueOptions.action(item.parentRow.rowNo,
						item.parentCol.colNo);
			});
			this.action = item.valueOptions.action;
		};

		this.applyValue = function(item) {
			item.valueOptions = {
				action : this.action
			};
			item.updateCellByWidget();
			this.defaultValue = $toggleDiv.html();
		};

		this.getValue = function() {
			return $toggleDiv.html();
		};

		this.isValueChanged = function() {
			var isModified = ((!($toggleDiv.html() == "" && this.defaultValue == null)) && ($toggleDiv
					.html() != this.defaultValue));
			return isModified;
		};

		this.validate = function() {
			return true;
		};
		this.setBackground = function(value) {
		};

		this.init();

	},

	NotionalEditor : function(args) {
		var $notionalDiv = null;
		var $input = null;
		var $toggleDiv = null;
		var $toggleText = null;
		this.defaultValue = "";
		this.defaultNotionalVal = "";

		this.rowNo = args.rowNo;
		this.colNo = args.colNo;
		this.possibleOptions = [];
		this.currentIndex = -1;

		this.displayType = new formatType("currency");
		this.displayMode = "format";

		this.init = function() {
			var GRID = args.grid;
			var THIS = this;

			// Removing the tab index will make the currency of notional not
			// toggle
			$notionalDiv = $("<div class='notional' tabindex='0'/>");
			$input = $("<input class='notionalInput' type='text'/>");
			$toggleDiv = $("<div class='toggle' tabindex='0'/>");
			$toggleText = $("<div class='currencyText' tabindex='1'/>");

			$input.appendTo($notionalDiv);
			$toggleDiv.appendTo($notionalDiv);
			$toggleText.appendTo($toggleDiv);
			$notionalDiv.appendTo(args.container);

			$toggleDiv.width(35);
			$toggleDiv.height(getLengthFromString(args.container.style.height));

			$toggleText.width(35);
			$toggleText
					.height(getLengthFromString(args.container.style.height));

			$toggleDiv.css("left",
					(getLengthFromString(args.container.style.width)) - 33);
			$input.width(getLengthFromString(args.container.style.width) - 33);

			var clickFn = function(event) {
				if (THIS.possibleOptions.length > 0) {
					THIS.currentIndex = (THIS.currentIndex + 1)
							% THIS.possibleOptions.length;
					$toggleText.html(THIS.possibleOptions[THIS.currentIndex]);
					GRID.saveOneScreenCellInMemory(THIS.rowNo, THIS.colNo);
				}
				THIS.blur();
				GRID.mode = "select";
				preventPropagation(event);
			};

			$toggleDiv.click(clickFn);
			$toggleDiv.dblclick(clickFn);

			var toggleDivKeyDown = function(event) {
				if (event == null)
					event = window.event;
				switch (event.keyCode) {
				case 13: // Enter key
					GRID.mode = "none";
					clickFn();
					break;
				case 9:
					if (event.shiftKey) {
						$input.focus();
						GRID.mode = "none";
						preventDefaultBehaviour(event);
					} else {
						GRID.mode = "edit";
					}
					break;
				default:
					GRID.mode = "select";
					break;
				}
			};

			$toggleDiv.keydown(toggleDivKeyDown);
			$toggleText.keydown(toggleDivKeyDown);

			$toggleDiv.focus();
			if (args.item)
				this.loadValue(args.item);

			$input.focus(function(event) {
				GRID.mode = "edit";
				if (THIS.defaultNotionalVal != 0)
					$input.val(THIS.defaultNotionalVal);
				THIS.displayMode = "value";
				preventDefaultBehaviour(event);
			});
			
			$input.blur(function(event) {
				GRID.saveOneScreenCellInMemory(THIS.rowNo, THIS.colNo);
				args.grid.mode = "none";
				var value = $input.val();
				if (value.length > 0 && value.substr(0,1) == "="){
					
				}
				else {
					var notionalVal = (notionalValue($input.val()));
					if (!notionalVal || notionalVal == 0 || notionalVal == ""
						|| notionalVal == "0")
						$input.val("");
					else
						$input.val(notionalVal);
				}
				GRID.saveOneScreenCellInMemory(THIS.rowNo, THIS.colNo);
				THIS.loadValue(args.item);
				THIS.displayMode = "format";
			});
			
			var inputKeyDown = function(event) {
				if (event == null)
					event = window.event;
				switch (event.keyCode) {
				case 9:
					if (!event.shiftKey) {
						THIS.blur();
						$toggleDiv.focus();
						GRID.mode = "none";
						preventDefaultBehaviour(event);
					} else {
						GRID.mode = "edit";
					}
					break;
				default:
					GRID.mode = "edit";
					break;
				}
			};

			$input.keydown(inputKeyDown);

		};

		this.destroy = function() {
			this.blur();
		};
		this.focus = function() {
			args.grid.mode = "edit";
			$input.focus();
			$input.select();
		};

		this.blur = function() {
			$input.blur();
			var THIS = this;
			var GRID = args.grid;
			
		};

		this.autofocus = function() {
			return true;
		};

		this.loadValue = function(item) {
			if (item.valueOptions == null)
				return;
			this.currentIndex = item.valueOptions.currentIndex;
			this.displayMode = "format";
			this.displayType = new formatType("currency");
			this.displayType.decimalCount = 0;
			this.displayType.symbol = "";
			this.rowNo = item.parentRow.rowNo;
			this.colNo = item.parentCol.colNo;
			this.possibleOptions = item.valueOptions.possibleOptions;
			$toggleText.html(this.possibleOptions[this.currentIndex]);
			this.defaultValue = this.possibleOptions[this.currentIndex];
			var value = "" + item.valueOptions.value;
			
			if (this.validate(value)) {
				var regEx = /^[-+]?[0-9]*\.?[0-9]*[mMkKbBtT]?$/;
				var m = value.match(regEx);
				if (m && m.length > 0 && value != "") {
					var notionalValue = convertTo(item.valueOptions.value,
							"integer", this.displayType);
					if (!notionalValue || notionalValue == "")
						$input.val("");
					else
						$input.val(notionalValue);
					this.defaultNotionalVal = item.valueOptions.stringValue;
					return;
				} else {
					$input.val("" + item.valueOptions.value);
					this.defaultNotionalVal = item.valueOptions.stringValue;
				}
			}
			else {
				$input.val(item.valueOptions.value);
				this.defaultNotionalVal = item.valueOptions.value;
			}
		};

		this.applyValue = function(item) {
			if (this.validate($input.val())) {
				var value = $input.val().toLowerCase();
				if (value.length > 0 && value.substr(0,1) == "="){
					item.errorMesg = "";
					item.valueOptions = {
						possibleOptions : this.possibleOptions,
						currentIndex : this.currentIndex,
						stringValue : value
					};
					item.updateCellByString(value);
					return;
				}
				var numVal = notionalValue(convertToNum($input.val(),
						this.displayType))
				if (numVal < 1000000000000000 && numVal >= 0
						&& $input.val() != "") {
					item.errorMesg = "";
					item.valueOptions = {
						possibleOptions : this.possibleOptions,
						value : notionalValue(numVal),
						stringValue : ""+notionalValue(numVal),
						currentIndex : this.currentIndex,
						displayValue : convertTo(notionalValue(numVal),
								"integer", this.displayType)
					};
					item.updateCellByString(notionalValue(numVal));
					this.defaultNotionalVal = item.valueOptions.value;
					this.defaultValue = item.valueOptions.possibleOptions[item.valueOptions.currentIndex];
					var newVal = convertTo(notionalValue(numVal), "integer",
							this.displayType)

					if (!newVal)
						$input.val("");
					else
						$input.val(newVal);
					return;
				}
				
			}
			this.defaultNotionalVal = $input.val();
			item.valueOptions = {
				possibleOptions : this.possibleOptions,
				value : $input.val(),
				currentIndex : this.currentIndex,
				displayValue : $input.val()
			};
			this.defaultValue = item.valueOptions.possibleOptions[item.valueOptions.currentIndex];
			item.errorMesg = "Incorrect Input";
			item.updateCellByString($input.val());
		};

		this.getValue = function() {
			return $input.val();
		};

		this.isValueChanged = function() {
			var isModified = ((!($toggleText.html() == "" && this.defaultValue == null)) && ($toggleText
					.html() != this.defaultValue))
					|| ((!($input.val() == "" && this.defaultNotionalVal == null))
							&& ($input.val() != this.defaultNotionalVal) && this.displayMode == "value");
			return isModified;
		};

		this.validate = function(strValue) {
			var regEx = /^[+]?[0-9]*\.?[0-9]*[mMkKbBtT]?$/;
			var value = strValue + "";
			var m = value.match(regEx);
			if (m && m.length > 0){
				return true;
			}
			if (value.length > 0 && value.substr(0,1) == "="){
				return true;
			}
			var notionalVal = notionalValue(convertToNum(strValue,
					this.displayType));
			if (parseInt(notionalVal) != NaN)
				return true;
			else
				return false;
		};
		this.setBackground = function(value) {
			$input.css("background", value);
			$notionalDiv.css("background", value);
		};

		this.init();

	},

	DateCellEditor : function(args) {
		var $input = null;
		var $img = null;
		var calendarOpen = false;

		this.init = function() {
			var GRID = args.grid;
			var THIS = this;
			$input = $("<INPUT type=text class='dateGridCell' />");
			$img = $("<div class='dateImage' />");
			$input.appendTo(args.container);
			$img.appendTo(args.container);

			this.isOpen = false;

			if (args.item)
				this.loadValue(args.item);

			$input.width(getLengthFromString(args.container.style.width));
			$img.css("top", 2);
			$img.height(15);
			$img.width(16);
			$img.css("left",
					getLengthFromString(args.container.style.width) - 21);
			$input.height(getLengthFromString(args.container.style.height));

			$input.focus(function(e) {
				GRID.mode = "edit";
				preventPropagation(e);
			});
			$input.blur(function(e) {
				GRID.mode = "none";
				GRID.saveOneScreenCellInMemory(THIS.rowNo, THIS.colNo);
			});
			var changeFn = function(e) {
				GRID.mode = "none";
				GRID.saveOneScreenCellInMemory(THIS.rowNo, THIS.colNo);
			};

			$input.bind("change", changeFn);

			var inputKeyDown = function(event) {
				if (event == null)
					event = window.event;
				switch (event.keyCode) {
				case 13: // Enter key
					dateToggle(event);
					GRID.mode = "select";
					break;
				default:
					break;
				}
			};

			var dateToggle = function(event) {
				var thisWidget = THIS;
				if (calendarOpen == false) {
					var currentTime = new Date();
					var month = currentTime.getMonth() + 1;
					var day = currentTime.getDate();
					var year = currentTime.getFullYear();
					var formattedDate = (day + " " + getNameOfMonth(month)
							+ " " + year);

					$input.DatePicker({
						flat : false,
						date : formattedDate,
						current : formattedDate,
						calendars : 1,
						starts : 1,
						onChange : function(formated, dates) {
							$input.val(formated);
							GRID.saveOneScreenCellInMemory(thisWidget.rowNo,
									thisWidget.colNo);
							$input.DatePickerHide();
							GRID.mode = "select";
						}
					});
					if ($input.val() && THIS.validate($input.val()))
						$input.DatePickerSetDate($input.val(), true);
					calendarOpen = true;
					$input.DatePickerShow();
				} else {
					calendarOpen = false;
					$input.DatePickerHide();
					$(".datepicker").remove();
				}
			};

			$input.keydown(inputKeyDown);

			$img.mousedown(function(e) {
				GRID.goToCell(THIS.rowNo, THIS.colNo);
				dateToggle(e);
				preventPropagation(e);
			});

			$input.mouseup(function() {
				$input.DatePickerHide();
			});
		};

		this.destroy = function() {
			$(".datepicker").remove();
		};

		this.position = function(position) {

		};

		this.autofocus = function() {
			return false;
		};

		this.focus = function() {
			$input.focus();
			$input.select();
			args.grid.mode = "edit";
		};

		this.blur = function() {
			$input.blur();
			args.grid.mode = "none";
		};

		this.loadValue = function(item) {
			this.rowNo = item.parentRow.rowNo;
			this.colNo = item.parentCol.colNo;
			
			if (item.valueOptions == null)
				return;
			var value = "";
			if (item.valueOptions.value) {
				var dateVal = getDateFromText(""+item.valueOptions.value);
				if (dateVal) {
					value = dateVal;
					var dateReturnValue = date_str_to_time(dateVal);
					if (dateVal) {
						item.errorMesg = "";
						item.valueOptions = {
							value : dateReturnValue
						};
						item.updateCellByWidget();
					} 
				}
				else value = date_to_str(item.valueOptions.value);
			}
			
			if (this.validate(value)){
				var dval = "";
				var dateArray = null;
				var m = value.match(/^[1-9][0-9]*[dwmyDWMY]$/);
				if (m && m.length == 1) {
					$input.val(value);
					return;
				} 
				value = date_to_str(value);
				dateArray = value.toString().split("/");
				dval = dateArray[1] + " " + getNameOfMonth(dateArray[0]) + " "
				+ dateArray[2];
				
				this.defaultValue = dval;
				$input.val(dval);
			}
			else {
				this.defaultValue = value;
				$input.val(value);
			}
		};

		this.serializeValue = function() {
			return $input.val();
		};

		this.applyValue = function(item) {
			this.defaultValue = $input.val();
			if (this.validate($input.val())){
				var dateVal = getDateFromText($input.val());
	
				var dateReturnValue = dateVal;
				var m = dateVal.match(/^[1-9][0-9]*[dwmyDWMY]$/);
				if (m && m.length == 1) {
				} else {
					dateReturnValue = date_str_to_time(dateVal);
				}
	
				if (dateVal) {
					item.errorMesg = "";
					item.valueOptions = {
						value : dateReturnValue
					};
					item.updateCellByWidget();
					this.loadValue(item);
					return;
				} 
			}
			item.errorMesg = "Incorrect Input";
			item.valueOptions = {
				value : $input.val()
			};

			item.updateCellByWidget();
			this.loadValue(item);
		};

		this.getValue = function() {
			return $input.val();
		};

		this.isValueChanged = function() {
			var dateVal = $input.val();
			var isModified = (dateVal
					&& (!(dateVal == "" && this.defaultValue == null)) && (dateVal != this.defaultValue));
			return isModified;
		};

		this.validate = function(dateString) {
			return validateDate(dateString);
		};
		this.setBackground = function(value) {
			$input.css("background", value);
			if (args.container)
				$(args.container).css("background", value);
		};

		this.setClass = function(className) {
			// $input.attr("class", className);
		};

		this.init();
	},

	InputCellEditor : function(args) {
		var $input = null;
		this.dataFormatType = "string";

		this.init = function() {
			var cellObj;
			var GRID = args.grid;
			var THIS = this;
			cellObj = document.createElement("input");
			cellObj.className = "inputCell ";

			cellObj.onblur = function() {
				GRID.saveOneScreenCellInMemory(THIS.rowNo, THIS.colNo);
				GRID.mode = "none";
			};

			cellObj.onfocus = function() {
				GRID.mode = "edit";
				$input.select();
				var displayCellArrayObj = $input;
				if (GRID.isInXLSRange(THIS.rowNo, THIS.colNo)) {
					var XLSCellArrayObj = GRID.XLSCellArray[THIS.rowNo][THIS.colNo];
					if (XLSCellArrayObj.displayType
							&& XLSCellArrayObj.displayType.format == "percent") {
						this.value = percentToVal(XLSCellArrayObj.stringValue);
					} else
						this.value = XLSCellArrayObj.stringValue;
					XLSCellArrayObj.isDirty = false;
				}
			};

			cellObj.onkeyup = function() {
				GRID.mode = "edit";
				var displayCellArrayObj = $input;
				if (GRID.formulaBarInput != null)
					GRID.formulaBarInput.value = displayCellArrayObj.val();
			};
			args.container.appendChild(cellObj);
			$input = $(cellObj);
			var isCurrent = GRID.isCurrentCell(THIS.rowNo, THIS.colNo);
			if (args.item)
				this.loadValue(args.item, isCurrent);
		};

		this.destroy = function() {
		};

		this.position = function(position) {
			$input.width(position.width);
			$input.height(position.height);
		};

		this.focus = function() {
			args.grid.mode = "edit";

		};

		this.blur = function() {
			args.grid.mode = "none";
		};

		this.autofocus = function() {
			return false;
		};

		this.loadValue = function(item, isCurrent) {
			if (item.valueOptions == null)
				return;
			this.rowNo = item.parentRow.rowNo;
			this.colNo = item.parentCol.colNo;
			this.defaultValue = item.valueOptions.value;

			if (isCurrent)
				$input.val(item.displayValue);
			else {
				if (item.displayType){
					this.displayType = item.displayType;
				}
				if (item.displayType && item.displayType.format == "percent") {
					this.value = percentToVal(item.stringValue);
				} else {
					this.value = item.stringValue;
				}
				$input.val(this.value);
			}
			
			if (item.valueOptions.dataFormatType)
				this.dataFormatType = item.valueOptions.dataFormatType;
		};

		this.serializeValue = function() {
			return $input.val();
		};

		this.applyValue = function(item) {
			var value = $input.val();
			if (this.validate()) {
				item.errorMesg = "";
				if (item.displayType && item.displayType.format == "percent") {
					value = stringToPercent(value);
				} else if (item.displayType
						&& item.displayType.format == "delta") {
					var regEx = /^[0-9]*$/;
					var match = value.match(regEx);
					if (match && match.length > 0) {
					} else
						value = "";
				}
			}
			else {
				item.errorMesg = "Incorrect Input";
			}
			item.valueOptions = {
					value : value,
					dataFormatType : this.dataFormatType
				};
			item.updateCellByString(value);
			item.isModified = false;
			this.defaultValue = item.valueOptions.value;
		};

		this.getValue = function() {
			return $input.val();
		};

		this.isValueChanged = function() {
			var isModified = (!($input.val() == "" && this.defaultValue == null))
					&& ($input.val() != this.defaultValue);
			return isModified;
		};

		this.validate = function() {
			if (this.displayType){
				var myStr = "" + $input.val();
				if (this.displayType.format == "percent"){
					myStr = myStr.replace(/^[ ]|[ ]$/g,"");
					myStr = myStr.replace(/%/,"");
				}
				else if (this.displayType.format == "delta"){
					myStr = removeDelta(myStr)
				}
				else if (this.displayType.format == "currency" && this.displayType.format == "accounting"){
					myStr = convertToNum(myStr, formatType);
				}
				var regEx = /^[-+]?[0-9]*\.?[0-9]*$/
				var match = myStr.match(regEx);
				if (match && match.length > 0) {
					return true;
				} else

					return false;
			} 
			else if (this.dataFormatType == "integer") {
				var regEx = /^[0-9]*$/;
				var match = $input.val().match(regEx);
				if (match && match.length > 0) {
					return true;
				} else
					return false;
			} else if (this.dataFormatType == "float") {
				var regEx = /^[-+]?[0-9]*\.?[0-9]*$/
				var match = $input.val().match(regEx);
				if (match && match.length > 0) {
					return true;
				} else
					return false;
			} else if (this.dataFormatType == "string")
				return true;
			else
				return true;
		};

		this.focus = function() {
			var GRID = args.grid;
			GRID.mode = "edit";
			$input.focus();
			$input.select();
		};

		this.blur = function() {
			$input.blur();
		};

		this.setBackground = function(value) {
			$input.css("background", value);
		};

		this.setClass = function(className) {
			$input.attr("class", className);
		};
		this.init();
	},

	DropDownCellEditor : function(args) {
		var input;
		this.rowNo = args.rowNo;
		this.colNo = args.colNo;
		this.dropdownMenu = null;
		this.selectedObj = null;
		this.shortcutMap = {};
		this.valueMap = {};
		this.isShortcutEnabled = false;
		this.isAutocompleteEnabled = true;
		this.highlightedItem = null;
		this.highlightedIndex = 0;
		var $div = null;

		this.init = function() {
			var THIS = this;

			input = document.createElement("input");
			input.type = "text";
			input.className = "dropdownText";
			var GRID = args.grid;

			$div = $("<div class='drop_down'/>");
			var $img = $("<div class='drop_down_img'>");

			$img
					.click(function(event) {

						if (THIS.dropdownMenu == null) {
							$(document).trigger("click");
							var values = THIS.possibleValues;
							if (values.length > 0) {
								THIS.dropdownMenu = new dropdown();
								THIS.dropdownMenu
										.createDropdown(args.container);
								THIS.dropdownMenu.dropdownDiv.style.minWidth = (args.container.clientWidth + "px");

								THIS.dropdownMenu
										.setPosition({
											left : ($(args.container).offset().left),
											top : ($(args.container).offset().top + args.container.clientHeight)
										});

								document.onclick = function() {
									THIS.destroy();
								}
								document.body.onclick = function() {
									THIS.destroy();
								}
							}
							for ( var i = 0; i < values.length; i++) {
								if (values[i].value != "") {
									var dropdownMenuItem = new dropdownItem();
									dropdownMenuItem.create();
									if (!values[i].displayLabel)
										values[i].displayLabel = values[i].label;
									dropdownMenuItem.setObj({
										name : values[i].label,
										displayName : values[i].displayLabel,
										shorcut : values[i].shortcut,
										value : values[i].value,
										index : i,
										subIndex : -1
									});
									THIS.dropdownMenu.addItem(dropdownMenuItem);

									if (values[i].subMenu) {
										for ( var j = 0; j < values[i].subMenu.length; j++) {
											var submenuItem = new dropdownItem();
											submenuItem.create();
											if (!values[i].subMenu[j].displayLabel)
												values[i].subMenu[j].displayLabel = values[i].subMenu[j].label;
											submenuItem
													.setObj({
														name : values[i].subMenu[j].label,
														displayName : values[i].subMenu[j].displayLabel,
														shorcut : values[i].subMenu[j].shortcut,
														value : values[i].subMenu[j].value,
														index : i,
														subIndex : j
													});
											submenuItem
													.setAction(function() {
														THIS.destroy();
														preventPropagation(event);
														if (this.parentObj) {
															THIS.selectedObj = this.parentObj.dropdownItemObj;
															input.value = this.parentObj.dropdownItemObj.displayName;
															GRID
																	.saveOneScreenCellInMemory(
																			args.rowNo,
																			args.colNo);
														}
													});
											dropdownMenuItem
													.addSubMenuItem(submenuItem);
										}
										var action = function(event) {
											preventPropagation(event);
										}
										dropdownMenuItem.setAction(action);
									} else {
										var action = function(event) {
											THIS.destroy();
											preventPropagation(event);
											THIS.selectedObj = this.parentObj.dropdownItemObj;
											input.value = this.parentObj.dropdownItemObj.displayName;
											GRID.saveOneScreenCellInMemory(
													args.rowNo, args.colNo);
										}
										dropdownMenuItem.setAction(action);
									}
								}
							}
							preventPropagation(event);
						} else {
							$(document).trigger("click");
							// THIS.destroy();
						}
					});
			input.onkeydown = function(event) {
				if (event == null)
					event = window.event;
				if (THIS.dropdownMenu) {
					if (THIS.isAutocompleteEnabled) {
						if (THIS.highlightedItem) {
							if (event.keyCode == 38) {
								if (THIS.highlightedIndex > 0) {
									THIS.dropdownMenu.dropdownItemArray[THIS.highlightedIndex]
											.setSelected(false);
									THIS.highlightedIndex--;
									THIS.highlightedItem = THIS.dropdownMenu.dropdownItemArray[THIS.highlightedIndex].dropdownItemObj;
									THIS.dropdownMenu.dropdownItemArray[THIS.highlightedIndex]
											.setSelected(true);
									preventPropagation(event);
								}
							} else if (event.keyCode == 40) {
								if (THIS.highlightedIndex >= 0) {
									if (THIS.highlightedIndex + 1 < THIS.dropdownMenu.dropdownItemArray.length) {
										THIS.dropdownMenu.dropdownItemArray[THIS.highlightedIndex]
												.setSelected(false);
										THIS.highlightedIndex++;
										THIS.highlightedItem = THIS.dropdownMenu.dropdownItemArray[THIS.highlightedIndex].dropdownItemObj;
										THIS.dropdownMenu.dropdownItemArray[THIS.highlightedIndex]
												.setSelected(true);
									}
									preventPropagation(event);
								}
							}
						}
					}
				}
			}

			input.onkeyup = function(event) {
				if (event == null)
					event = window.event;
				GRID.mode = "edit";
				if (event.keyCode == 38) {

				} else if (event.keyCode == 40) {

				}

				else if (THIS.isAutocompleteEnabled) {
					if (THIS.dropdownMenu != null)
						THIS.destroy();
					var values = THIS.possibleValues;
					if (values.length > 0) {
						THIS.dropdownMenu = new dropdown();
						THIS.dropdownMenu.createDropdown(args.container);
						THIS.dropdownMenu.dropdownDiv.style.minWidth = args.container.style.width;
						THIS.dropdownMenu
								.setPosition({
									left : (findPosX(args.container)),
									top : (findPosY(args.container) + getLengthFromString(args.container.style.height))
								})
						document.onclick = function() {
							THIS.destroy();
						}

					}
					var count = 0;
					for ( var i = 0; i < values.length; i++) {
						if (values[i].value != "") {
							if (!values[i].displayLabel)
								values[i].displayLabel = values[i].label;
							if (values[i].displayLabel.toLowerCase().match(
									input.value.toLowerCase())) {
								count++;
								var dropdownMenuItem = new dropdownItem();
								dropdownMenuItem.create();
								dropdownMenuItem.setObj({
									name : values[i].label,
									displayName : values[i].displayLabel,
									shorcut : values[i].shortcut,
									value : values[i].value,
									index : i,
									subIndex : -1
								});
								THIS.dropdownMenu.addItem(dropdownMenuItem);
								if (count == 1) {
									THIS.highlightedItem = {
										name : values[i].label,
										displayName : values[i].displayLabel,
										shorcut : values[i].shortcut,
										value : values[i].value,
										index : i,
										subIndex : -1
									};
									THIS.highlightedIndex = 0;
									dropdownMenuItem.setSelected(true);
								}
								var action = function(event) {
									THIS.destroy();
									preventPropagation(event);
									THIS.selectedObj = this.parentObj.dropdownItemObj;
									input.value = this.parentObj.dropdownItemObj.displayName;
									GRID.saveOneScreenCellInMemory(args.rowNo,
											args.colNo);
								}
								dropdownMenuItem.setAction(action);
							}
						}
					}
					if (count == 0) {
						var dropdownMenuItem = new dropdownItem();
						dropdownMenuItem.create();
						dropdownMenuItem.setObj({
							name : "Invalid Input",
							displayName : "Invalid Input",
							shorcut : "",
							value : "Invalid Input",
							index : -1,
							subIndex : -1
						});
						THIS.dropdownMenu.addItem(dropdownMenuItem);
						var action = function(event) {
							THIS.destroy();
							preventPropagation(event);
						}
						dropdownMenuItem.setAction(action);
					}

					preventPropagation(event);
				}
			}

			$(input).width("100%");
			$div.appendTo(args.container);
			$(input).appendTo($div);
			$img.appendTo($div);

			if (args.item)
				this.loadValue(args.item);
		};

		this.destroy = function() {
			this.highlightedItem = null;
			this.highlightedIndex = -1;
			if (this.dropdownMenu != null) {
				this.dropdownMenu.hideSubMenu();
				document.body.removeChild(this.dropdownMenu.dropdownDiv);
			}
			this.dropdownMenu = null;
			document.onclick = null;
		};

		this.loadValue = function(item) {
			if (item.valueOptions == null)
				return;
			var valueOptions = item.valueOptions;
			if (item.valueOptions.isAutocompleteEnabled == false) {
				this.isAutocompleteEnabled = false;
			}
			var i = item.valueOptions.index;
			var j = item.valueOptions.subIndex;
			if (j == -1) {
				var values = valueOptions.possibleValues;
				if (!values[i].displayLabel)
					values[i].displayLabel = values[i].label;
				this.selectedObj = {
					name : values[i].label,
					displayName : values[i].displayLabel,
					shorcut : values[i].shortcut,
					value : values[i].value,
					index : i,
					subIndex : -1
				}
			} else {
				var values = valueOptions.possibleValues;
				if (!values[i].subMenu[j].displayLabel)
					values[i].subMenu[j].displayLabel = values[i].subMenu[j].label;
				this.selectedObj = {
					name : values[i].subMenu[j].label,
					displayName : values[i].subMenu[j].displayLabel,
					shorcut : values[i].subMenu[j].shortcut,
					value : values[i].subMenu[j].value,
					index : i,
					subIndex : j
				}
			}
			input.value = this.selectedObj.displayName;
			this.defaultValue = this.selectedObj.displayName;
			this.possibleValues = valueOptions.possibleValues;
			this.rowNo = item.parentRow.rowNo;
			this.colNo = item.parentCol.colNo;

			this.valueMap = {};
			this.shortcutMap = {};

			for ( var i = 0; i < valueOptions.possibleValues.length; i++) {
				if (valueOptions.possibleValues[i].subMenu) {
					this.isAutocompleteEnabled = false;
					for ( var j = 0; j < valueOptions.possibleValues[i].subMenu.length; j++) {
						if (valueOptions.possibleValues[i].subMenu[j].shortcut) {
							this.shortcutMap[valueOptions.possibleValues[i].subMenu[j].shortcut
									.toUpperCase()] = {
								index : i,
								subIndex : j
							};
							this.valueMap[valueOptions.possibleValues[i].subMenu[j].value] = {
								index : i,
								subIndex : j
							};
						}
					}
				} else {
					if (valueOptions.possibleValues[i].shortcut) {
						this.shortcutMap[valueOptions.possibleValues[i].shortcut
								.toUpperCase()] = {
							index : i,
							subIndex : -1
						};
						this.valueMap[valueOptions.possibleValues[i].value] = {
							index : i,
							subIndex : -1
						};
					}
				}

			}

		};

		this.serializeValue = function() {
			return combo.value;
		};

		this.applyValue = function(item) {
			var inputValue = input.value.toUpperCase();
			var values = this.possibleValues;
			var shortcutMap = this.shortcutMap;
			var valueMap = this.valueMap;
			if (this.highlightedItem) {
				this.selectedObj = this.highlightedItem;
				this.highlightedItem = null;
				this.highlightedIndex = -1;
				input.value = this.selectedObj.displayName;
			} else {
				if (shortcutMap[inputValue]) {
					if (shortcutMap[inputValue].subIndex == -1) {
						var i = shortcutMap[inputValue].index;
						if (!values[i].displayLabel)
							values[i].displayLabel = values[i].label;
						this.selectedObj = {
							name : values[i].label,
							displayName : values[i].displayLabel,
							shorcut : values[i].shortcut,
							value : values[i].value,
							index : i,
							subIndex : -1
						}
						input.value = this.selectedObj.displayName;
					} else {
						var i = shortcutMap[inputValue].index;
						var j = shortcutMap[inputValue].subIndex;
						if (!values[i].subMenu[j].displayLabel)
							values[i].subMenu[j].displayLabel = values[i].subMenu[j].label;
						this.selectedObj = {
							name : values[i].subMenu[j].label,
							displayName : values[i].subMenu[j].displayLabel,
							shorcut : values[i].subMenu[j].shortcut,
							value : values[i].subMenu[j].value,
							index : i,
							subIndex : j
						};
						input.value = this.selectedObj.displayName;
					}
				} else if (valueMap[inputValue]) {
					if (valueMap[inputValue].subIndex == -1) {
						var i = valueMap[inputValue].index;
						if (!values[i].displayLabel)
							values[i].displayLabel = values[i].label;
						this.selectedObj = {
							name : values[i].label,
							displayName : values[i].displayLabel,
							shorcut : values[i].shortcut,
							value : values[i].value,
							index : i,
							subIndex : -1
						}
						input.value = this.selectedObj.displayName;
					} else {
						var i = valueMap[inputValue].index;
						var j = valueMap[inputValue].subIndex;
						if (!values[i].subMenu[j].displayLabel)
							values[i].subMenu[j].displayLabel = values[i].subMenu[j].label;
						this.selectedObj = {
							name : values[i].subMenu[j].label,
							displayName : values[i].subMenu[j].displayLabel,
							shorcut : values[i].subMenu[j].shortcut,
							value : values[i].subMenu[j].value,
							index : i,
							subIndex : j
						}
						input.value = this.selectedObj.displayName;
					}
				}
			}
			item.valueOptions = {
				value : this.selectedObj.value,
				possibleValues : this.possibleValues,
				index : this.selectedObj.index,
				subIndex : this.selectedObj.subIndex
			};
			input.value = this.selectedObj.displayName;
			item.updateCellByWidget();
			this.defaultValue = this.selectedObj.value;
		};

		this.getValue = function() {
			return this.selectedObj.value;
		};

		this.isValueChanged = function() {
			var isModified = (!(input.value == "" && this.defaultValue == null))
					&& (input.value != this.defaultValue);
			return isModified;
		};

		this.validate = function() {
			return {
				valid : true,
				msg : null
			};
		};

		this.focus = function() {
			input.focus();
			input.select();
			args.grid.mode = "select";
		};

		this.blur = function() {
			input.blur();
			var GRID = args.grid;
			GRID.saveOneScreenCellInMemory(this.rowNo, this.colNo);
			this.destroy();
			GRID.mode = "none";
		};

		this.autofocus = function() {
			return false;
		};

		this.position = function(position) {
			$(input).width(position.width);
			$(input).height(position.height);
		};

		this.setBackground = function(value) {
			$(input).css("background", value);
			$div.css("background", value);
		};

		this.setClass = function(className) {
			// combo.className = className;
		};
		this.init();
	},

	getEditor : function(widgetName) {
		switch (widgetName) {
		case "date":
			return FinsheetEditors.DateCellEditor;
		case "input":
			return FinsheetEditors.InputCellEditor;
		case "combo":
			return FinsheetEditors.DropDownCellEditor;
		case "dropdown":
			return FinsheetEditors.DropDownCellEditor;
		case "currencyPair":
			return FinsheetEditors.CurrencyPairEditor;
		case "toggle":
			return FinsheetEditors.ToggleEditor;
		case "notional":
			return FinsheetEditors.NotionalEditor;
		case "amount":
			return FinsheetEditors.NotionalEditor;
		case "delta":
			return FinsheetEditors.DeltaEditor;
		case "classToggle":
			return FinsheetEditors.ClassToggleEditor;
		case "radio":
			return FinsheetEditors.RadioCellEditor;
		case "label":
			return null;
		case "checkbox":
			return FinsheetEditors.CheckBoxCellEditor;
		default:
			return FinsheetEditors.InputCellEditor;
		}
	},

	getFormatter : function(widgetName) {
		switch (widgetName) {
		case "date":
			return FinsheetEditors.DateCellFormatter;
		case "input":
			return FinsheetEditors.InputCellFormatter;
		case "combo":
			return FinsheetEditors.DropDownCellEditor;
		case "dropdown":
			return FinsheetEditors.DropDownCellEditor;
		case "delta":
			return FinsheetEditors.DeltaCellFormatter;
		case "currencyPair":
			return FinsheetEditors.CurrencyPairEditor;
		case "toggle":
			return FinsheetEditors.ToggleEditor;
		case "notional":
			return FinsheetEditors.NotionalEditor;
		case "amount":
			return FinsheetEditors.NotionalEditor;
		case "classToggle":
			return FinsheetEditors.ClassToggleEditor;
		case "checkbox":
			return FinsheetEditors.CheckBoxCellEditor;
		case "radio":
			return FinsheetEditors.RadioCellEditor;
		case "label":
			return FinsheetEditors.textLabelFormatter;
		default:
			return FinsheetEditors.InputCellFormatter;
		}

	},

	hasFormatter : function(widgetName) {
		switch (widgetName) {
		case "date":
			return false;
		case "input":
			return true;
		case "combo":
			return false;
		case "dropdown":
			return false;
		case "delta":
			return true;
		case "currencyPair":
			return false;
		case "toggle":
			return false;
		case "notional":
			return false;
		case "amount":
			return false;
		case "classToggle":
			return true;
		case "radio":
			return false;
		case "checkbox":
			return false;
		case "label":
			return true;
		default:
			return true;
		}
	}

};
