var version = navigator.appVersion;

function hookDiv(divid,imgid) {
	if (version.indexOf("Windows") != -1) {
		var thisDiv = document.getElementById(divid).style;
		if (imgid!="") {var thisImg = document.getElementById(imgid);}
		if (thisDiv.display == "") {
			thisDiv.display = "none";
			if (imgid!=""){thisImg.src = "../images/div_on.gif";}
		}
		else {
			thisDiv.display = "";
			if (imgid!=""){thisImg.src = "../images/div_off.gif";}
		}
	}
	return false;
}

function toggleProcedureOpen(currProcedure) {
	if (version.indexOf("Windows") != -1) {
		thisProcedure = document.getElementById("procedure"+currProcedure).style;
		thisExpander = document.getElementById("expander"+currProcedure);
		if (thisProcedure.display == "block" | thisProcedure.display == "") {
			thisProcedure.display = "none";
			thisExpander.src = "../../_sharedassets/expand.gif";
		}
		else {
			thisProcedure.display = "block";
			thisExpander.src = "../../_sharedassets/collapse.gif";
		}
	}
	return false
}