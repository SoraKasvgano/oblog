var objActive="oblog_edittext";
var sAgent=navigator.userAgent.toLowerCase();
var IsIE=sAgent.indexOf("msie")!=-1;
function GetActiveText(objHTML) {
	objActive=objHTML;
	if(document.selection){
		var obj=document.getElementById(objHTML);
		obj.currPos = document.selection.createRange().duplicate()
	}
}

function InsertText(objHTML,strText,bolReplace) {
	if(strText==""){return("")}
	var obj=document.getElementById(objHTML);
	if(document.selection){
		if (obj.currPos){
			if(bolReplace && (obj.value=="")){
				obj.currPos.text=strText
			}
			else{
				obj.currPos.text+=strText
			}
		}
		else{
			obj.value+=strText
		}
	}
	else{
		if(bolReplace){
			obj.value=obj.value.slice(0,obj.selectionStart) + strText + obj.value.slice(obj.selectionEnd,obj.value.length)
		}
		else{
			obj.value=obj.value.slice(0,obj.selectionStart) + strText + obj.value.slice(obj.selectionStart,obj.value.length)
		}
	}
	//obj.focus();
}

function ReplaceText(objHTML,strPrevious,strNext) {
	var obj=document.getElementById(objHTML);
	var strText;
	if(document.selection && document.selection.type == "Text"){
		if (obj.currPos || IsIE){
			var range = document.selection.createRange();
			range.text = strPrevious + range.text + strNext;
			return("");
		}
		else{
			strText=strPrevious + strNext;
			return(strText);
		}
	}
	else{
		if(obj.selectionStart || obj.selectionEnd){
			strText=strPrevious + obj.value.slice(obj.selectionStart,obj.selectionEnd) + strNext;
			return(strText);
		}
		else{
			strText=strPrevious + strNext;
			return(strText);
		}
	}
	
}

function UBB_smiley(){
  var smileyPos=new getPos('A_smiley');
  smileyPos.position="relative";
  smileyPos.Left=0;
  smileyPos.Top=0;
  smileyPanel=document.getElementById('oblog_ubbemot');
  document.getElementById("oblog_ubbemot").style.cssText="overflow-x: hidden;margin: 0 0 -127px 0;";
  smileyPanel.style.position=smileyPos.position;
  smileyPanel.style.left="110px";
  smileyPanel.style.top="0px";
  smileyPanel.style.visibility ="visible";
  smileyPanel.innerHTML=getemot();
  if (IsIE){
  	document.body.attachEvent("onclick",CloseSmileyPanel);
  }
  else{
  	document.body.addEventListener("click",CloseSmileyPanel,true);
  }
}

function CloseSmileyPanel(){
  smileyPanel=document.getElementById('oblog_ubbemot');
  smileyPanel.style.visibility ="hidden";
  if (IsIE){
  	document.body.detachEvent("onclick",CloseSmileyPanel);
  }else{
  document.body.removeEventListener("click",CloseSmileyPanel,true);
  }
}

function onClickEmot(str){
	var n=str.lastIndexOf("face");
	str=str.substring(n);
	str=str.replace("face","");
	str=str.replace(".gif","");
	InsertText(objActive,ReplaceText(objActive,'[emot]'+str,'[/emot]'),true);
    CloseSmileyPanel();
}

function getPos(obj){
  this.Left=0;
  this.Top=0;
  var tempObj=document.getElementById(obj);
  while (tempObj.tagName.toLowerCase()!="body"){
  	 this.Left+=tempObj.offsetLeft;
  	 this.Top+=tempObj.offsetTop;
  	 tempObj=tempObj.offsetParent;
  }
}

function getemot(){
	var s="<TBODY><TR>";
	for (i=1;i<=50;i++){
		s=s+"<TD align=\"center\"><img style=\"cursor:pointer;margin:3px;\"src='"+ubbimg+"editor\/images\/emot\/face"+i+".gif' onClick='onClickEmot(this.src)'></TD>";	
		if (i/10==parseInt(i/10) && i!=50){
			s=s+"</tr><tr>";
		}
	}
	s=s+"</tr></tbody>";
	return s;
}