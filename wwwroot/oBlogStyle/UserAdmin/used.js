function PIMMenu_Click(p_oEvent)
{
	var oEvent = p_oEvent ? p_oEvent : window.event;
	if(p_oEvent) oEvent.stopPropagation();
	else oEvent.cancelBubble = true;
	if(oEvent && oEvent.target && oEvent.target.parentNode && oEvent.target.parentNode.tagName == "a") window.location = oEvent.target.parentNode.href;
};

function Arrow_Click(p_oEvent)
{
	document.Selects = document.getElementsByTagName('select');

	if(document.Selects[0])
	{
		var nSelects = document.Selects.length-1;
		for(var i=nSelects;i>=0;i--) document.Selects[i].style.visibility = 'hidden';
	}

    var oAvatar = document.getElementById("swfcontainer");
    
    if(oAvatar) {
        oAvatar.style.visibility = "hidden";
    }

	var oEvent = p_oEvent || window.event;
	
	if(p_oEvent) oEvent.stopPropagation();
	else oEvent.cancelBubble = true;
	
	HideMenu();
	
	var oTab = this.parentNode.parentNode;
	var nTop = (oTab.offsetTop+oTab.parentNode.offsetHeight);
    
	g_oMenu = document.getElementById(this.href.split('#')[1]);
	g_oMenu.onclick = PIMMenu_Click;
	g_oMenu.style.visibility = "visible";

	document.onclick = Document_Click;	

	return false;
};

function Tabs_Init()
{
	var oAddressBookTab = document.getElementById('addressbooktab');
	var oNotepadTab = document.getElementById('notepadtab');		

	if(oAddressBookTab)
	{
		oAddressBookTab.getElementsByTagName("a")[1].onclick = Arrow_Click;
		oAddressBookTab.Selected = (oAddressBookTab.className == 'selected' || oAddressBookTab.className == 'first selected') ? true : false;
	}
	if(oNotepadTab)
	{
		oNotepadTab.getElementsByTagName("a")[1].onclick = Arrow_Click;
		oNotepadTab.Selected = (oNotepadTab.className == 'selected' || oNotepadTab.className == 'first selected') ? true : false;
	}

	return false;
};

function HideMenu()
{
	if(typeof g_oMenu != 'undefined' && g_oMenu)
	{
		if(g_oMenu.Hide) g_oMenu.Hide();
		else g_oMenu.style.visibility = 'hidden';

		var hideCB = g_oMenu.hideCB;
		if ( typeof hideCB == "function" ) {
			hideCB.call( g_oMenu );
		}
		
		g_oMenu = null;
		document.onclick = null;
		window.onresize = null;
	}
	else return;
};

function Document_Click()
{
	if(document.Selects)
	{
		var nSelects = document.Selects.length-1;
		for(var i=nSelects;i>=0;i--) document.Selects[i].style.visibility = 'visible';
	}

    var oAvatar = document.getElementById("swfcontainer");

    if(oAvatar) {
        oAvatar.style.visibility = "visible";
    }
    
	HideMenu();
};


function init() 
    { 
		if(document.getElementById)
		{
			Tabs_Init();
	    	document.onkeydown = function(evt) {  }
			if(typeof OnLoad != 'undefined') OnLoad();
		}
    }
onload=init