/*
 *   copyright (c) -- chaitanya nalla and kechit goyal
 *
 */
function defineContextMenu() {
	contextMenu = function() {
		this.contextMenuId = "";
		this.contextMenuItemArray = [];
		this.contextMenuDiv = null;
		this.contextMenuItemCount = 0;
		this.subMenu = null;
		
	};
	
	contextMenu.prototype.createContextMenu = function(id, parentDivID){
		this.contextMenuDiv = document.createElement("div"); 
		this.contextMenuDiv.className = "contextMenuDiv";
		this.contextMenuId = id;
		this.contextMenuDiv.id = id;
		if (parentDivID)
			document.getElementById(parentDivID).appendChild(this.contextMenuDiv);
		else document.body.appendChild(this.contextMenuDiv);
	};
	
	contextMenu.prototype.addItem = function(item){
		this.contextMenuItemArray[this.contextMenuItemCount] = item;
		item.setParentDivId(this.contextMenuId);
		this.contextMenuDiv.appendChild(item.contextMenuItemDiv);
		item.setIndex(this.contextMenuItemCount);
		this.contextMenuItemCount++;
		item.parent = this;
		var THIS = this;
		item.contextMenuItemDiv.onmouseover = function (){
			THIS.hideSubMenu();
		};
	};
	
	contextMenu.prototype.hideSubMenu = function (){
		for (var i = 0;i < this.contextMenuItemArray.length; i++){
			if (this.contextMenuItemArray[i].subMenu){
				this.contextMenuItemArray[i].subMenu.hideSubMenu();
			}
		}
		if (this.subMenu != null){
			this.subMenu.contextMenuDiv.style.display = "none";
			if (document.body.contains(this.subMenu.contextMenuDiv))
				document.body.removeChild(this.subMenu.contextMenuDiv);
			this.subMenu = null;
		}
	};
	
	contextMenu.prototype.showSubMenu = function (){
		if (this.subMenu != null){
			document.body.appendChild(this.subMenu.contextMenuDiv);
			this.subMenu.contextMenuDiv.style.display = "block";
		}
	};
	
	contextMenuItem = function() {
		this.label= "";
		this.action = null;
		this.parentDivId = "";
		this.contextMenuItemDiv = null;
		this.contextMenuItemId = "";
		this.subMenu = null;
		this.parent = null;
		this.index = 0;
	};
	
	contextMenuItem.prototype.create = function(id){
		this.contextMenuItemDiv = document.createElement("div") ;
		this.contextMenuItemDiv.className = "contextMenuCell";
		this.contextMenuItemId = id;
		this.contextMenuItemDiv.id = id;
		document.getElementById(this.parentDivId).appendChild(this.contextMenuItemDiv);		
	};
	
	contextMenuItem.prototype.setIndex = function(index){
		this.index = index;		
	};
	
	contextMenuItem.prototype.setLabel = function (label){
		this.label = label;
		this.contextMenuItemDiv.innerHTML = label;
	};
	
	contextMenuItem.prototype.setAction = function (action){
		this.action = action;
		this.contextMenuItemDiv.onclick = action;
	};
	
	contextMenuItem.prototype.setParentDivId = function (parentDivId){
		this.parentDivId = parentDivId;
	};
	
	contextMenuItem.prototype.addSubMenuItem = function (subMenuItem){
		var THIS = this;
		if (this.subMenu == null){
			this.subMenu = new contextMenu();
			this.subMenu.createContextMenu(this.contextMenuItemId + "subMenu");
			document.body.removeChild(this.subMenu.contextMenuDiv);
			$(THIS.contextMenuItemDiv).append($("<span class='fr' style='padding-right:5px'><img src='resources/newGrid/img/side-arrow.png'/></span>"));
			
			this.subMenu.addItem(subMenuItem);
			this.contextMenuItemDiv.onmouseover = function (){
				THIS.parent.hideSubMenu();
				THIS.parent.subMenu = THIS.subMenu;
				THIS.parent.showSubMenu();
			};
			this.subMenu.contextMenuDiv.style.left = getLengthFromString(this.parent.contextMenuDiv.style.left) + 150 + "px";
			this.subMenu.contextMenuDiv.style.top = (25 * THIS.index + getLengthFromString(this.parent.contextMenuDiv.style.top)) + "px";
		}
		else {
			this.subMenu.addItem(subMenuItem);
		}
	}
}

defineContextMenu();
