/* admin_index_left.htm 展开隐藏 */
function menu(oblog){
if(oblog.style.display=='')oblog.style.display='none';
else oblog.style.display='';
}
/* admin_index_left.htm 点中后 */
function switchTab(tabpage,tabid){
var oItem = document.getElementById(tabpage).getElementsByTagName("a"); 
    for(var i=0; i<oItem.length; i++){
        var x = oItem[i];    
        x.className = "";
}
	document.getElementById(tabid).className = "Selected";
	var dvs=document.getElementById("cnt").getElementsByTagName("div");
	for (var i=0;i<dvs.length;i++){
	  if (dvs[i].id==('d'+tabid))
	    dvs[i].style.display='block';
	  else
  	  dvs[i].style.display='none';
	}
}
function border_left(tabpage,left_tabid){
var oItem = document.getElementById(tabpage).getElementsByTagName("li"); 
    for(var i=0; i<oItem.length; i++){
        var x = oItem[i];    
        x.className = "";
}
	document.getElementById(left_tabid).className = "Selected";
	var dvs=document.getElementById("left_tabid_cnt").getElementsByTagName("ul");
	for (var i=0;i<dvs.length;i++){
	  if (dvs[i].id==('d'+left_tabid))
	    dvs[i].style.display='block';
	  else
  	  dvs[i].style.display='none';
	}
}