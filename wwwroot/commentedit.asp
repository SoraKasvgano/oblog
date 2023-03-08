<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<%
dim edittype,c_uname,c_upass,c_uurl
Dim oblog
set oblog = new class_sys
oblog.autoupdate = False
oblog.start

'编辑器类型,1=纯文本,2=可视化,3=ubb
edittype=3
select case edittype
	case 1
%>
if (chkdiv('oblog_edit')) {
document.getElementById('oblog_edit').innerHTML='<textarea name="oblog_edittext"  rows="12"  cols="50" onfocus=\"addcode();\"></textarea>';
}
<%
	case 2
%>
if (chkdiv('oblog_edit')) {
	document.getElementById('oblog_edit').innerHTML='<textarea id="oblog_edittext" name="oblog_edittext" style="width:100%;height:320px;display:none"><\/textarea >';
	_editor_url  = '<%=C_Editor%>';
	_editor_lang = "ch";
	document.write('<script src="<%=blogurl%>editor/htmlarea.js"></script>');
	oblog_editors = null;
    oblog_init    = null;
    oblog_config  = null;
    oblog_plugins = null;
	oblog_editortype=2;
    oblog_init = oblog_init ? oblog_init : function()
    {
	oblog_editors = oblog_editors ? oblog_editors :
     ['oblog_edittext'];
	oblog_config = new HTMLArea.Config(oblog_editortype);
	oblog_config.width  = 360;
	oblog_config.height = 200;
     oblog_config = oblog_config ? oblog_config : new HTMLArea.Config(oblog_editortype);

      oblog_editors   = HTMLArea.makeEditors(oblog_editors, oblog_config, oblog_plugins);
      HTMLArea.startEditors(oblog_editors);
	  //HTMLArea.focusEditor();
      window.onload = null;
	 }
    window.onload   = oblog_init;

}
<%
case 3
%>
var ubbimg='<%=blogurl%>';
document.write('<script src="<%=blogurl%>editor/ubb.js"></script>');
document.write("<style type='text/css'>@import url('<%=blogurl%>editor/ubb.css');</style>");
ubbhtml="<div id=\"oblog_ubb\">";
ubbhtml+="<div class=\"oblog_ubbtoolbar\">";
ubbhtml+="	<a href=\"javascript:InsertText(objActive,ReplaceText(objActive,\'[b]\',\'[\/b]\'),true);void(0)\"><img src=\""+ubbimg+"images\/bold.gif\" alt=\"粗体\"  border=\"0\" align=\"absmiddle\"><\/a>";
ubbhtml+="	<a href=\"javascript:InsertText(objActive,ReplaceText(objActive,\'[i]\',\'[\/i]\'),true);void(0)\"><img src=\""+ubbimg+"images\/italic.gif\" alt=\"斜体\" border=\"0\" align=\"absmiddle\" ><\/a>";
ubbhtml+="	<a href=\"javascript:InsertText(objActive,ReplaceText(objActive,\'[u]\',\'[\/u]\'),true);void(0)\"><img src=\""+ubbimg+"images\/underline.gif\" alt=\"下划线\" border=\"0\" align=\"absmiddle\"><\/a>";
ubbhtml+="	<a href=\"javascript:InsertText(objActive,ReplaceText(objActive,\'[quote]\',\'[\/quote]\'),true);void(0)\"><img src=\""+ubbimg+"images\/quote.gif\" alt=\"插入引用\" border=\"0\" align=\"absmiddle\"><\/a>";
ubbhtml+="	<a href=\"javascript:UBB_smiley();void(0)\"><img src=\""+ubbimg+"images\/smiley.gif\" alt=\"插入表情\" border=\"0\" align=\"absmiddle\" id=\"A_smiley\"><\/a>";
ubbhtml+="	<\/div>";
ubbhtml+="	<div id=\"oblog_ubbemot\">";
ubbhtml+="	<\/div>";
ubbhtml+="	  <textarea name=\"oblog_edittext\" cols=\"92\" rows=\"10\" id=\"oblog_edittext\" class=\"oblog_ubbtext\" onfocus=\"addcode();\" ><\/textarea>";
ubbhtml+="<\/div>";
ubbhtml+="	<div id=\"oblog_vcode\">";
ubbhtml+="	<\/div>";
if (chkdiv('oblog_edit')) {
document.getElementById('oblog_edit').innerHTML=ubbhtml;
}
<%
end select

c_uname=Request.Cookies(cookies_name)("username")
c_upass=oblog.DecodeCookie(Request.Cookies(cookies_name)("Password"))
c_uurl=oblog.DecodeCookie(Request.Cookies(cookies_name)("userurl"))
If c_uname="" And c_upass="" And oblog.cacheConfig(90)=1 Then 
	Dim GuestTmpName
	GuestTmpName="访客"&RndPassword(6)
'	If true_domain =1 Then 
'	Response.Cookies(cookies_name).Path   = oblog.l_uDomain  
'	Else 
'	Response.Cookies(cookies_name).Path   =   blogdir
'	End If 
	Response.Cookies(cookies_name).Expires = Date + 999
	Response.Cookies(cookies_name)("username")=GuestTmpName
	c_uname=GuestTmpName
End If 
if left(c_uurl,1)<>"/" then
	c_uurl="http://"&c_uurl
end if
%>
if (chkdiv('UserName')) {
document.getElementById('UserName').value='<%=c_uname%>';
}
if (chkdiv('Password')) {
document.getElementById('Password').value='<%=c_upass%>';
}
if (chkdiv('homepage')) {
document.getElementById('homepage').value='<%=c_uurl%>';
}

function reply_quote(id)
{
	var etype='<%=edittype%>';
	if (etype=='1'){
		oblog_editors['oblog_edittext'].setHTML("<div class='quote'><strong>以下引用"+document.all["n_"+id].innerHTML+"在"+document.all["t_"+id].innerHTML+"发表的评论:</strong><br /><br />"+document.all["c_"+id].innerHTML+"</div><br />\n");
		//oblog_editors['oblog_edittext']._iframe.contentWindow.focus();
	}else{
		var ttext=document.all["c_"+id].innerHTML
		var simg;
		var simgs="";
		var simg1;
		ttext=ttext.replace(/<BR>/g,"[br]")
		ttext=ttext.replace(/(<STRONG>)(.[^\[]*)(<\/STRONG>)/,"[b]$2[/b]");
		ttext=ttext.replace(/(<U>)(.[^\[]*)(<\/U>)/,"[u]$2[/u]");
		ttext=ttext.replace(/(<EM>)(.[^\[]*)(<\/EM>)/,"[i]$2[/i]");
		ttext=ttext.replace(/<DIV class=quote>/g,"[quote]");
		ttext=ttext.replace(/<\/DIV>/g,"[/quote]");
		ttext=ttext.replace(/\.gif">/g,".gif\">##");
		simg=ttext.split("##");
		for(var i=0;i<simg.length;i++){
			simg1=simg[i].replace(/<IMG.[^\[]*face([^\.]*)\.gif">/,"[EMOT]$1[/EMOT]");
			simgs=simgs + simg1;
			}
		ttext=simgs;
		ttext=ttext.replace(/<IMG.[^\[]*face([^\.]*)\.gif">/,"[EMOT]$1[/EMOT]");
		document.getElementById('oblog_edittext').value+="[quote][b]以下引用"+document.all["n_"+id].innerHTML+"在"+document.all["t_"+id].innerHTML+"发表的评论:[/b]\n"+ttext+"[\/quote]\n";
		document.getElementById('oblog_edittext').focus();
	}
}

function Verifycomment()
{
	var oblog_edittext = document.getElementById("oblog_edittext");
	var commenttopic = document.getElementById("commenttopic");
	if(commenttopic.value==''){
		alert("请输入标题!");
		commenttopic.focus();
		return false;
	}
	if(oblog_edittext.value==''){
		alert("请输入评论内容!");
		oblog_edittext.focus();
		return false;
	}
	<%If oblog.CacheConfig(30) = "1" Then%>
	if(document.all("CodeStr").value==''){
		alert("请输入验证码");
		document.all("CodeStr").focus();
		return false;
	}
	<%else%>
	return true;
	<%End if%>
}
<% Set oblog = Nothing%>