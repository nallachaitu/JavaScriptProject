/*
 *   copyright (c) -- chaitanya nalla and kechit goyal
 *
 */

function defineDropdown() {
	dropdownItemObj = function(name, shortcut, value, displayName){
		this.name = name;
		this.shortcut = shortcut;
		this.value = value;
		this.displayName = displayName;
	};
	
	dropdown = function() {
		this.dropdownDiv = null;
		this.dropdownItemArray = [];
		this.dropdownItemCount = 0;
		this.subMenu = null;
	};
	
	dropdown.prototype.createDropdown = function(parentDivElement){
		this.dropdownDiv = document.createElement("div"); 
		this.dropdownDiv.className = "dropdownDiv";
		document.body.appendChild(this.dropdownDiv);
	};
	
	dropdown.prototype.addItem = function(item){
		this.dropdownItemArray[this.dropdownItemCount] = item;
		this.dropdownDiv.appendChild(item.dropdownItemDiv);
		this.dropdownItemCount++;
		item.parent = this;
		var THIS = this;
		item.dropdownItemDiv.onmouseover = function (){
			THIS.hideSubMenu();
		}
	};
	
	dropdown.prototype.hideSubMenu = function (){
		if (this.subMenu != null){
			this.subMenu.dropdownDiv.style.display = "none";
			document.body.removeChild(this.subMenu.dropdownDiv);
			this.subMenu = null;
		}
	};
	
	dropdown.prototype.showSubMenu = function (){
		if (this.subMenu != null){
			document.body.appendChild(this.subMenu.dropdownDiv);
			this.subMenu.dropdownDiv.style.display = "block";
		}
	};
	
	/*dropdown.prototype.showMenu = function (){
		$(this.dropdownDiv).slideDown("fast");
	};*/
	
	dropdown.prototype.setPosition = function (pos){
		this.dropdownDiv.style.left = pos.left + "px";
		this.dropdownDiv.style.top = pos.top + "px";
	};
	
	dropdown.prototype.setWidth = function(width){
		this.dropdownDiv.style.width = width + "px";
	}
	
	dropdown.prototype.destroy = function(){
		this.hideSubMenu();
		document.body.removeChild(this.dropdownDiv);
	}
	
	dropdownItem = function() {
		this.dropdownItemObj = null;
		this.action = null;
		this.subMenu = null;
		this.parent = null;
		this.dropdownItemDiv = null;
		this.parentDivElement = null;
	};
	
	dropdownItem.prototype.create = function(){
		this.dropdownItemDiv = document.createElement("div");
		this.dropdownItemDiv.parentObj = this;
		this.dropdownItemDiv.className = "dropdownCell";
	};
	
	dropdownItem.prototype.setObj = function (obj){
		this.dropdownItemObj = obj;
		this.dropdownItemDiv.innerHTML = obj.name;
	};
	
	dropdownItem.prototype.setSelected = function (value){
		if (value)
			this.dropdownItemDiv.className = "dropdownCellSelected";
		else this.dropdownItemDiv.className = "dropdownCell";
	};
	
	dropdownItem.prototype.setAction = function (action){
		this.action = action;
		this.dropdownItemDiv.onclick = action;
	};
	
	dropdownItem.prototype.setParentDivElement = function (parentDivElement){
		this.parentDivElement = parentDivElement;
	};
	
	dropdownItem.prototype.addSubMenuItem = function (subMenuItem){
		var THIS = this;
		if (this.subMenu == null){
			this.subMenu = new dropdown();
			this.subMenu.createDropdown(this.dropdownItemId + "subMenu");
			document.body.removeChild(this.subMenu.dropdownDiv);
			this.subMenu.addItem(subMenuItem);
			THIS.dropdownItemDiv.innerHTML = THIS.dropdownItemDiv.innerHTML + "<span class='fr' style='padding-left:5px'><img src='/static/newGrid/img/side-arrow.png'/></span>";
			this.dropdownItemDiv.onmouseover = function (){
				THIS.parent.hideSubMenu();
				THIS.parent.subMenu = THIS.subMenu;
				THIS.parent.showSubMenu();
			};
			this.subMenu.dropdownDiv.style.left = getLengthFromString(this.parent.dropdownDiv.style.left) + this.parent.dropdownDiv.offsetWidth  + "px";
			this.subMenu.dropdownDiv.style.top = findPosY(THIS.dropdownItemDiv) + "px";
		}
		else {
			this.subMenu.addItem(subMenuItem);
		}
	}
}

defineDropdown();
