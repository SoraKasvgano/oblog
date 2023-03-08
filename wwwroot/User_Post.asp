<!--#include file="user_top.asp"-->
<!--#include file="inc/class_Trackback.asp"-->
<script src="oBlogStyle/move.js" type="text/javascript"></script>
<script src="inc/function.js" type="text/javascript"></script>
<%
If oblog.l_uNewbie=1 Then
	Response.write("<script>parent.show_title('选择模版')</script>")
	oblog.ShowMsg "发布前请先选择一个喜欢的模版。","user_template.asp?action=showconfig"
	'jscmd="go_cmdurl('选择模版','tab3')"
end If
If oblog.Chkiplock() Then
	oblog.ShowMsg ("对不起！你的IP已被锁定，不允许操作！"),blogdir &"index.html"
	Set oblog = Nothing
End If
If Application(cache_name_user&"_systemenmod")<>"" Then
	Dim enStr
	enStr=Application(cache_name_user&"_systemenmod")
	enStr=Split(enStr,",")
	If enStr(2)="1" Then	Response.write("系统临时禁止操作日志与相册!"):Response.End()
End If
Dim rsGroup,sGroups
Set rsGroup=oblog.Execute("select groupid,g_name From oblog_groups Order By Groupid Desc")
Do While Not rsGroup.Eof
	sGroups=sGroups & "									<option value="&rsGroup(0)&">" & rsGroup(1) & "</option>" & vbcrlf
 	rsGroup.MoveNext
Loop
Dim tMode
tMode = Trim(Request("tMode"))
If tMode = "" Then tMode = "flash"
If t=1 or t=2 Then
	dim flashurl
	If tMode = "flash" Then
%>
<script type="text/javascript" src="inc/flash.js"></script>
<script type="text/vbscript" src="inc/flash_vb.js"></script>
<%
End if
	select case t
		case 1
			If tMode ="flash" Then
				dim tmpstr
				if Trim(Request("action"))="showphoto" then tmpstr="" else tmpstr="upload"
				flashurl="photo.swf?action="&tmpstr&"&blogurl=&userid="&oblog.l_uid&"&u_post=true"
			Else
				Call upload()
				Response.End
			End if
		case 2
			flashurl="cam3.swf"
	end select
%>
<%Sub upload()%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="#" onclick="purl('user_post.asp?t=1','FLASH')">ＦＬＡＳＨ</a></li>
					<li><a href="#" onclick="purl('user_photo.asp','相片管理')">相片管理</a></li>
					<li><a href="#" onclick="purl('user_Albumcomments.asp','相片评论')">相片评论</a></li>
					<li><a href="#" onclick="purl('user_subject.asp?t=1','相片分类')">相片分类</a></li>

				</ul>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
					<iframe id='d_file' frameborder='0' src='upload.asp?re=no&isphoto=1&tMode=2' width='100%' height='100%' scrolling='no'></iframe>
			</td>
		</tr>
	</tbody>
</table>
</body>
</html>
<%End sub%>
<table id="TableBody" cellpadding="0">
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<script language="JavaScript" type="text/javascript">
					var hasRightVersion = DetectFlashVer(requiredMajorVersion, requiredMinorVersion, requiredRevision);
					if(!hasRightVersion && <%=t%>==999) {
						document.write('\您的flash版本过低，<a href="http:\/\/www.adobe.com\/go\/getflash\/" target="_blank"\>请点击升级Flash Player插件来支持大头贴程序</a>');
					  }else{
						document.write("<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0' width='100%' height='500' ><param name='wmode' value='transparent' /><param name='movie' value='<%=flashurl%>' /><param name='quality' value='high' /><embed src='<%=flashurl%>' quality='high' pluginspage='http://www.macromedia.com/go/getflashplayer' type='application/x-shockwave-flash' width='100%' height='500'></embed></object>");
					  }
					</script>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/0.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
</table>
</body>
</html>
<%
	Response.End()
End If
Dim sDisable
action = Trim(Request("action"))
If action<>"savelog" Then
	sDisable=" "
Else
	sDisable=" disabled"
End if
%>
<table id="TableBody" class="UserPost" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="#" onClick="return doMenu('swin1');" <%=sDisable%>>高级选项</a></li>
					<li><a href="#" onClick="return doMenu('swin2');" <%=sDisable%>>引用通告</a></li>
					<li><a href="#" onClick="return doMenu('swin3');" <%=sDisable%>>上传文件</a></li>
					<li><a href="#" onClick="return doMenu('swin4');" <%=sDisable%>>文章摘要</a></li>
				</ul>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
<%
Dim action,rs
Set rs=Server.CreateObject("Adodb.Recordset")


If action="savelog" Then
		Call savelog
	Else
	    Call main
End If
%>
			</td>
		</tr>
	</tbody>
</table>
</body>
</html>
<%
Sub main()
    Dim  logid, log_specialid,photofile
    logid = Request.QueryString("logid")
    photofile = Trim(Request.QueryString("photofile"))
    Dim log_tags, tags, filename,  log_type,log_abstract,log_files
    Dim face, topic, classid, subjectid, logtext, istop, ishide, isencomment, showword, addtime, userid, ispassword, authorid,tburl, oldisdraft,isneedlogin,viewscores,ViewGroupId
    oldisdraft = 0
    If logid<>"" Then logid=CLng(logid)
    If logid > 0 Then
        Set rs=oblog.execute("select * From oblog_log Where logid="&logid&" And (authorid= "&oblog.l_uid&" or userid="&oblog.l_uid&")")
        If rs.EOF Then
            Set rs = Nothing
            oblog.adderrstr ("无此权限操作此" & tName & "！")
            oblog.showusererr
        End If
		if rs("isdel")=1 then
			set rs=nothing
			oblog.ShowMsg "已删除日志，请先恢复后再操作。",""
		end if
        topic = rs("topic")
        face = rs("face")
        classid = rs("classid")
        subjectid = rs("subjectid")
        logtext = Replace(rs("logtext"), "#isubb#", "")
        istop = rs("istop")
        ishide = rs("ishide")
        isencomment = rs("isencomment")
        showword = rs("showword")
        addtime = rs("addtime")
        userid = rs("userid")
		authorid = rs("authorid")
        ispassword = rs("ispassword")
        log_files=rs("logpics")
		If ispassword<>"" Then ispassword="已设密码，若不修改请不要操作"
        tburl = rs("tburl")
        oldisdraft = rs("isdraft")
        filename = rs("filename")
        log_type = rs("logtype")
		log_abstract=rs("abstract")
        If IsNull(rs("logtags")) Then
            tags = ""
        Else
            tags = rs("logtags")
        End If
        log_specialid=rs("specialid")
		isneedlogin = OB_IIF(rs("isneedlogin"),0)
		viewscores = OB_IIF(rs("viewscores"),0)
		ViewGroupId = OB_IIF(rs("ViewGroupId"),0)
        Set rs = Nothing
    Else
    	'检测发表限制
    	Dim sPostAccess
    	sPostAccess=oblog.CheckPostAccess
    	If sPostAccess<>"" Then
			oblog.AddErrstr sPostAccess
         	oblog.ShowUserErr
    	End If
		authorid=oblog.l_uid
    End If
    If isencomment = "" Then isencomment = 1
    If userid = "" Then userid = oblog.l_uId
    if oblog.CacheConfig(23)="1" then
		if filename="" then filename=Year(now) & Month(now) & Day(now)&hour(now())&minute(now())&second(now())
	else
		if filename="" then filename="自动编号"
	end if
	call getteam()
%>

<script language=javascript>
function checkValue(){
	var allnum = 0;
	if (document.oblogform.ishide[0].checked==true)allnum ++;
	if (document.oblogform.isneedlogin[0].checked==true)allnum ++;
	if (document.oblogform.viewscores.value>0)allnum ++;
	if (document.oblogform.ispassword.value!='')allnum ++;
	if (document.oblogform.viewgroupid.value>0)allnum ++;
	if (allnum>1){
		alert ('请勿选择多个条件限定日志访问');
		return false;
	}
	if (document.oblogform.viewscores.value>0&&document.oblogform.abstract.value==''){
		alert('启用积分浏览的情况下，建议您填写文章摘要！');
	}
	if (document.oblogform.isneedlogin[0].checked==true&&document.oblogform.abstract.value==''){
		alert('启用登录可见的情况下，建议您填写文章摘要！');
	}
	doMenu('swin1')
}
function ResetValue(tvalue)
{
	<%If logid > 0 Then%>
//	document.all('oblogform').reset();
	document.getElementById('oblogform').reset;
	return;
	<%else%>
	var showword='<%=oblog.l_uShowlogWord%>';
	var filename='<%=filename%>';
	if (tvalue==1){
//		document.getElementsByName('isencomment')[isencomment].checked="checked";
		document.oblogform.isencomment[0].checked=true;
		document.oblogform.ishide[1].checked=true;
		document.oblogform.istop[1].checked=true;
		document.oblogform.showword.value=showword;
		document.oblogform.filename.value=filename;
		document.oblogform.ispassword.value='';
		document.oblogform.blogteam.options[0].selected='selected';
	}
	else if (tvalue==2){
		document.oblogform.tb.value='';
	}
	else if (tvalue==3){
		document.oblogform.abstract.value='';
	}
	<%end if%>
}
parent.show_title("发布日志");
var in_ob_useradmin=true;
var issubmit=false;
function chkfilename()
{
	var filename=del_space(document.oblogform.filename.value);
	if (filename=="自动编号"){document.oblogform.filename.value=""}
	if (filename==""){document.oblogform.filename.value="自动编号"}
}
function checkerr(string)
{
var i=0;
for (i=0; i<string.length; i++)
{
	if((string.charAt(i) < '0' || string.charAt(i) > '9') && (string.charAt(i) < 'a' || string.charAt(i) > 'z')&& (string.charAt(i) < 'A' || string.charAt(i) > 'Z')&& (string.charAt(i)!='-')&& (string.charAt(i)!='_'))
	{return 1;}
	}
	return 0;//pass
}
//以下为自动保存草稿的操作函数
var issave = null;
var isauto = false;
function countDown(Secs) {
	<%
	Select Case C_Editor_Type
		Case 1
		%>
		var edit = del_space(oEdit1.getHTMLBody());
		<%
		Case 2
	%>
	document.oblogform.edit.value=IframeID.document.body.innerHTML;
	var edit = del_space(document.oblogform.edit.value);
	<%End Select%>
	var save_ing = document.getElementById("save_ing");
	save_ing.style.display="";
	save_ing.innerHTML='<font color=red>'+Secs +'</font>秒后保存到草稿箱';
	if(Secs>0) {
		if (edit.length > 0&&!issubmit&&issave){
				setTimeout('countDown('+Secs+'-1)',1000);
			}
		else {
				issave = false;
				save_ing.style.display="none";
			}
	}
	else {
		setdraft();
		issave = false;
	}
}
function autoSetDraft(){
	var logid = <%=OB_IIF(logid,0)%>;
	if (issave == true||logid>0){
		return false;
	}
	<%
	Select Case C_Editor_Type
		Case 1
		%>
		var edit = del_space(oEdit1.getHTMLBody());
		<%
		Case 2
	%>
	document.oblogform.edit.value=IframeID.document.body.innerHTML;
	var edit = del_space(document.oblogform.edit.value);
	<%End Select%>
	if (edit.length > 0&&!issubmit){
		issave = true;
		countDown(<%=oblog.CacheConfig(7)%>*60);
	}
}
function setdraft()
{
	document.oblogform.isdraft.value='1';
	isauto = true;
	savelog();
	issave = false;
}
//以上为自动保存草稿的函数
function savelog()
{
	<%If C_Editor_Type=2 Then%>submits();<%End If%>
	if (document.oblogform.isdraft.value !='1'){
		document.getElementById("ob_submit").disabled="disabled";
		document.getElementById("ob_submit_d").disabled="disabled";
		document.getElementById("ob_submit_p").disabled="disabled";
		isauto = false;
	}
	document.getElementById("save_ing").style.display="";
	document.getElementById("save_ing").innerHTML="<img src='images/loading.gif' align='absbottom'> 正在保存...";
	if (issubmit){
		var oDialog = new dialog("<%=blogurl%>");
		oDialog.init();
	 	oDialog.event("正在提交中，请稍候...",'');
		oDialog.button('dialogOk',"");
	}else{
		var errstr=""
		<%If C_Editor_Type=1 Then%>document.oblogform.edit.value=oEdit1.getHTMLBody();<%End if%>
		var topic = del_space(document.oblogform.topic.value);
		 if (topic.length == 0&&document.oblogform.isdraft.value!="1")
		 {
			errstr="您忘了填写题目。<br />";
		 }
		 var needclassid=<%=oblog.CacheConfig(25)%>;
		 if (needclassid==1 &&document.oblogform.isdraft.value!="1") {
		 if (document.oblogform.classid.value == 0)
		 {
			errstr=errstr+"请选择日志的类别。<br />";
		 }
		 }

		var filename=del_space(document.oblogform.filename.value);
		if ((checkerr(filename) == "1")&&(filename!="")&&(filename!="自动编号")){
			errstr=errstr+"文件名称请用0-9的数字和a-z的半角字母及下划线,不允许中文和怪字符（如!@#$%^等）。<br />";
		}

		if (document.oblogform.edit.value == "")
		 {
			errstr=errstr+"请输入日志的内容。<br />";
		 }
		var date=document.oblogform.selecty.value+"-"+document.oblogform.selectm.value+"-"+document.oblogform.selectd.value
		var datereg=/^(\d{4})-(\d{1,2})-(\d{1,2})$/
		var datareg=/^(\d){1,2}$/
		if (!datereg.test(date)){
		  errstr=errstr+"日志时间输入格式错误。<br />";
		 }
		var r=date.match(datereg)
		var d=new Date(r[1],r[2]-1,r[3])
		if (!(d.getFullYear()==r[1]&&d.getMonth()==r[2]-1&&d.getDate()==r[3])){
		  errstr=errstr+"日志时间输入格式错误。<br />";
		 }
		 <%if oblog.l_Group(31,0)=1 then%>
		if (document.oblogform.codestr.value== ""&&isauto==false)
		{
			errstr=errstr+"请输入验证码。<br />";
		}
		<%end if%>
		 if (errstr!=""){
			var oDialog = new dialog("<%=blogurl%>");
			oDialog.init();
			oDialog.event(errstr,'');
			oDialog.button('dialogOk',"document.getElementById('ob_submit').disabled=false;document.getElementById('ob_submit_d').disabled=false;document.getElementById('ob_submit_p').disabled=false;");
			issubmit=false;
			document.getElementById("save_ing").style.display="none";
		 }
		 else{
			var re=/\+/g;
			var topic=escape(document.oblogform.topic.value.replace(re,"<%=Chr(25)%>"));
			var classid=document.oblogform.classid.value;
			var logtags=escape(document.oblogform.logtags.value.replace(re,"<%=Chr(25)%>"));
			var edit=escape(document.oblogform.edit.value.replace(re,"<%=Chr(25)%>"));
			<%if oblog.l_Group(31,0)=1 then%>
			var codestr=document.oblogform.codestr.value;
			var ob_codename=document.oblogform.ob_codename.value
			<%else%>
			var codestr='';
			var ob_codename='';
			<%end if%>
			if (document.oblogform.teamid){
				var teamid=read_checkbox('teamid');}
			else{
				var teamid='';}
			var isencomment=read_radio("isencomment");
			var ishide=read_radio("ishide");
			var istop=read_radio("istop");
			var ispassword=document.oblogform.ispassword.value;
			var filename=document.oblogform.filename.value;
			var selecty=document.oblogform.selecty.value;
			var selectm=document.oblogform.selectm.value;
			var selectd=document.oblogform.selectd.value;
			var selecth=document.oblogform.selecth.value;
			var selectmi=document.oblogform.selectmi.value;
			var tb=document.oblogform.tb.value;
			var abstract=document.oblogform.abstract.value;
			var logid=document.oblogform.logid.value;
			var oldisdraft=document.oblogform.oldisdraft.value;
			var isdraft=document.oblogform.isdraft.value;
			var subjectid=document.oblogform.subjectid.value;
			var blogteam=document.oblogform.blogteam.value;
			var blogteamsubject=document.oblogform.blogteamsubject.value;
			var showword=document.oblogform.showword.value;
			var viewscores=document.oblogform.viewscores.value;
			var viewgroupid=document.oblogform.viewgroupid.value;
			var log_files=document.oblogform.log_files.value;
			var isneedlogin=read_radio("isneedlogin");
			var Ajax = new oAjax("ajaxServer.asp?action=savelog",show_returnsave);
			var arrKey = new Array("topic",
									"classid",
									"logtags",
									"edit",
									"codestr",
									"ob_codename",
									"teamid",
									"isencomment",
									"ishide",
									"istop",
									"ispassword",
									"filename",
									"selecty",
									"selectm",
									"selectd",
									"selecth",
									"selectmi",
									"tb",
									"abstract",
									"logid",
									"oldisdraft",
									"isdraft",
									"blogteam",
									"blogteamsubject",
									"showword",
									"isneedlogin",
									"viewscores",
									"viewgroupid",
									"subjectid",
									"log_files");
			var arrValue = new Array(topic,
									classid,
									logtags,
									edit,
									codestr,
									ob_codename,
									teamid,
									isencomment,
									ishide,
									istop,
									ispassword,
									filename,
									selecty,
									selectm,
									selectd,
									selecth,
									selectmi,
									tb,
									abstract,
									logid,
									oldisdraft,
									isdraft,
									blogteam,
									blogteamsubject,
									showword,
									isneedlogin,
									viewscores,
									viewgroupid,
									subjectid,
									log_files);
			Ajax.Post(arrKey,arrValue);
			issubmit=true;
		 }
	}
}


function show_returnsave(arrobj){
	if (arrobj){
		switch (arrobj[1]){
		case '0':
			var oDialog = new dialog("<%=blogurl%>");
			oDialog.init();
			oDialog.event(arrobj[0],'');
			oDialog.button('dialogOk',"document.getElementById('ob_submit').disabled=false;document.getElementById('ob_submit_d').disabled=false;document.getElementById('ob_submit_p').disabled=false;"); //已修改完毕。
			issubmit=false;
			document.getElementById("save_ing").style.display="none";
			if (chkdiv('ob_codeimg')){
				var ob_codeimg=document.getElementById("ob_codeimg");
				ob_codeimg.src=ob_codeimg.src+"&t="+Math.random();
			}
			break;
		case '1':
			parent.get_draft();
			var oDialog = new dialog("<%=blogurl%>");
			oDialog.init();
			oDialog.event(arrobj[0],'');
			oDialog.button('dialogOk',"window.location='"+window.location.href.substring(0,window.location.href.lastIndexOf("/"))+"/"+"user_blogmanage.asp'");
			document.getElementById("ob_submit").disabled="disabled";
			document.getElementById("ob_submit_d").disabled="disabled";
			document.getElementById("ob_submit_p").disabled="disabled";
			document.getElementById("save_ing").style.display="none";
			break;
		case '2':
			document.getElementById("save_ing").innerHTML=arrobj[0];
			document.getElementById("logid").value=arrobj[2];
			document.oblogform.isdraft.value='0';
			parent.get_draft();
			issubmit=false;
			break;
		case '3':
			var oDialog = new dialog("<%=blogurl%>");
			oDialog.init();
			oDialog.event(arrobj[0],'');
			oDialog.button('dialogOk',"top.location='"+window.location.href.substring(0,window.location.href.lastIndexOf("/"))+"/"+"index.asp'");
			issubmit=false;
			document.getElementById("save_ing").style.display="none";
			break;
		}
	}
}



function taghelp(){
	var str='<div style="height:200px;overflow:auto">一、什么是标签（TAG）？<br />　　简单的说,标签就是一篇文章的"关键词"。您可以将日志文章或者照片，选择一个或多个词语（标签）来标记，这样一来，凡是我们博客网站上使用该词语的文章自动成为一个列表显示。<br />二、使用标签的好处：<br />　　1、您添加标签的文章就会被直链接到网站相应标签的页面，这样浏览者在访问相关标签时，就有可能访问到您的文章，增加了您的文章被访问的机会。<br />　　2、您可以很方便地查找到与您使用了同样标签的文章，延伸您文章的视野；可以方便地查找到与您使用了同样标签的作者，作为志同道合的朋友，您可以将他们加为好友或友情博客，扩大您的朋友圈。 <br />　　3、增加标签的方式完全由您自主决定，不受任何的限制，不用受网站系统分类和自己原有日志分类的限制，便于信息的整理、记忆和查找。<br />三、如何使用标签?<br />　　例如：您写了一篇到北京旅游的文章，按照文章提到的内容，您可以给这篇文章加上：<br />　　北京旅游,天安门,长城,故宫<br />　　等几个标签，当浏览者想搜索关于长城的文章时，浏览者会点击标签：长城，从而看到所有关于长城的文章，方便了浏览者查找日志，同时您也可用此方法找到和您同样喜欢的人，以便一起相互交流等等。<br />四．如何添加“好”的Tag？ <br />　　1． Tag应该要能够体现出自己的特色，并且是大家经常采用的熟悉的词语。<br />　　2．用词尽量简单精炼,词语字数不要太长，两三个字的词语就可以了，尽量是有意义的词汇，不要使用一些只作为装饰的符号，如｛｝等。<br />　　3．不要使用一些语义比较弱的词汇，如“我的家”，“图片”等。<br /></div>'
	var oDialog = new dialog("<%=blogurl%>");
	oDialog.init();
	oDialog.event(str,'');
	oDialog.button('dialogOk',"");
	document.getElementById("ob_boxface").style.display="none";

}
function setSort(s1,s2)
{
param=s1.selectedIndex-1;
if(param>=0)
	{
	s2.options.length=0;
		for(i=0;i<p_array[param].length;i++)
		{
			s2.options.length++;
			s2.options[i].text=p_array[param][i];
			s2.options[i].value=p_array_id[param][i];
		}

	}else
	{
	s2.options.length=0;
	s2.options.length++;
	s2.options[0].text="我的专题";
	s2.options[0].value="0";
	}
}
function CheckTeamID(tValue){
	var TeamPost=true;
	<%if oblog.CacheConfig(72) = "0" then %>
	alert('系统禁止日志同时发布到<%=oblog.CacheConfig(69)%>');
	TeamPost = false;
	<%end if%>
	var j = 0;
	var TeamID = document.getElementsByName('teamid');
	for (var i=0;i<TeamID.length ; i++){
		if (TeamPost==false){
			TeamID[i].checked=false;
			break;
		}
		else{
			if (TeamID[i].value==tValue){
				var k = i;
			}
			if (TeamID[i].checked==true){
				j++;
			}
			if (j><%=oblog.CacheConfig(72)%>){
				TeamID[k].checked=false;
				alert('单篇日志同时发布到<%=oblog.CacheConfig(69)%>的上限是<%=oblog.CacheConfig(72)%>');
				break;
			}
		}
	}
}
</script>
				<div id="chk_idAll">
					<form action="user_post.asp?action=savelog&t=<%=t%>" method="post" name="oblogform" id="oblogform">
					<table id="UserPost" cellpadding="0">
						<tr>
							<td class="t1">

								<div class="d1">
									<span><label for="topic">标题</label></span>
									<input type="text" name="topic" id="topic" size="40" maxlength="50" value="<%=AnsiToUnicode(topic)%>" />
									<select name="classid" id="classid" <%If oblog.CacheConfig(25)="1" Then Response.Write "class=""blue"" title=""必须填写系统分类"" " %>>
										<%=oblog.show_Postclass(classid)%>
									</select>
									<select name="subjectid" id="subjectid">
										<option value="0">我的分类</option>
<%
Set rs = oblog.Execute("select subjectid,subjectname from oblog_subject where userid=" & userid & " And subjectType=" & t)
While Not rs.EOF
	If rs(0) = subjectid Then
		Response.Write ("										<option value=" & rs(0) & " selected>" & oblog.filt_html(rs(1)) & "</option>") & vbcrlf
	Else
		Response.Write ("										<option value=" & rs(0) & " >" & oblog.filt_html(rs(1)) & "</option>") & vbcrlf
	End If
	rs.movenext
Wend
%>									</select>
								</div>
								<div class="d2" id="addressbooktab">
									<span><label for="logtags">标签</label></span>
									<a><input name="logtags" id="logtags" type="text" size="40" maxlength="255" value="<%=tags%>" /></a><a id="usedtags" href="#logtags_used"><img src="oBlogStyle/UserAdmin/7/user_team_top.png" alt="曾经使用过的标签" /></a>
									以<font class="blue"><%
									select Case P_TAGS_SPLIT
										Case " "
											Response.Write "空格"
										Case ","
											Response.Write "逗号"
										Case Else
											Response.Write P_TAGS_SPLIT
									End select
									%></font>分隔
									（<a href="#" onclick="taghelp();">什么是标签？</a>）
									<div id="logtags_used">
<%=GetUserTags%>
									</div>
									<script language="JavaScript" src="oBlogStyle/UserAdmin/used.js" type="text/javascript"></script>
								</div>
								<div class="d3">
<%
'----------------------------------------------------------
'编辑器显示
Select Case C_Editor_Type
Case 1
	Dim EditorHeight
	if oblog.l_Group(31,0)=1 then
		EditorHeight=250
	Else
		EditorHeight=250
	End If
%>
									<span id="loadedit"><img src='images/loading.gif' align='absbottom'>正在载入编辑器……</span>
									<textarea id="edit" name="edit" style="width:100%;height:<%=EditorHeight%>px; display:none"><%=Server.HtmlEncode(logtext)%></textarea>
<%
Case 2
%>
									<input type="hidden" id="edit" name="edit" value="<%if logtext<>"" then response.Write(Server.HtmlEncode(logtext))%>" />
<%
	Server.Execute C_Editor & "/edit.asp"
	End Select
'----------------------------------------------------------
%>
								</div>

<%if oblog.l_Group(31,0)=1 then%>
								<div class="d4" id="codestr_div">
									<label>验证码：<input name="codestr" type="text" size="6" maxlength="20" /></label> <%=oblog.getcode%>
								</div>
<%end if%>
<%
Dim sHidden
If oblog.l_Group(11,0) = 0 or 1=1 Then
	sHidden=""
Else
	sHidden="style=""display:none;"""
End If
%>
								<div class="d5">
									<div class="left">
										<input type="hidden" name="isdraft" id="isdraft" value="0" />
										<input type="hidden" name="logid" id="logid" value="<%=logid%>" />
										<input type="hidden" name="oldisdraft" id="oldisdraft" value=<%=oldisdraft%> />
										<input type="button" id="ob_submit"  value="发布日志" title="发布日志" onclick="savelog();" />
										<input type="button" id="ob_submit_d" value="保存为草稿" title="保存为草稿" onClick="setdraft();" />
										<input type="button" id="ob_submit_p" onClick="<%If C_Editor_Type=1 Then%>oEdit1.insertHTML('#此前在首页部分显示#')<%ElseIf C_Editor_Type=2 Then%>part();<%End If%>" value="部分显示标记" />
									</div>
									<div class="right">
										<span id="save_ing"></span>
									</div>
								</div>
							</td>
							<td class="t2">
								<script language=JavaScript>
								function turnit(ss,ii)
								{
								if (ss.style.display=="none")
								  {ss.style.display="";
								   ii.src="oBlogStyle/UserAdmin/7/close.gif";
								}
								else
								  {ss.style.display="none";
								   ii.src="oBlogStyle/UserAdmin/7/open.gif";}
								}
								</script>
								<div id="close" onmouseup="turnit(content1,img1);"><img id="img1" src="oBlogStyle/UserAdmin/7/close.gif" /></div>
							</td>
							<td class="t3" id="content1">
								<div id="group" <%=sHidden%>>
									<div>同时发布到<%=oblog.CacheConfig(69)%></div>
									<ul>

<%
Dim rst
Set rs=oblog.Execute("select a.teamid,a.t_name From oblog_team a ,(select  teamid From oblog_teamusers Where userid=" & oblog.l_Uid & " and state>=3) b Where a.teamid=b.teamid And a.istate=3")
If rs.Eof Then
	'Response.Write "									<li class=""red"">没有加入<a href=""#"" onclick=""purl('user_team.asp','" &oblog.CacheConfig(69)& "')"">" &oblog.CacheConfig(69)& "</a>？</li>"
	Response.Write "									<li class=""green"">没有加入" &oblog.CacheConfig(69)& "？</li>"
Else
		Do While Not rs.EOF
%>
										<li><label for="teamid<%=rs(0)%>"><input type="checkbox" name="teamid" id="teamid<%=rs(0)%>" value="<%=rs(0)%>" onclick="CheckTeamID('<%=rs(0)%>')">&nbsp;<%=rs(1)%></label></li>
<%
rs.movenext
Loop
End If
Set rst=Nothing
%>
									</ul>
								</div>
							</td>
						</tr>
					</table>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/42.js" type="text/javascript"></script>
				<div id="swin1" style="display:none;position:absolute;top:34px;left:10px;z-index:100;">
					<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
						<tr>
							<td colspan='2' align='center' class='win_table_top'>高级选项(特殊日志仅可选择一项限定)</td>
						</tr>
						<tr>
							<td class='win_table_td'>允许评论：</td>
							<td><label><input value='1' name='isencomment' id="isencomment" <%If isencomment=1 Or isencomment="" Then Response.Write " checked='checked'" End If%> type='radio'>是</label>  <label><input value='0' name='isencomment' id="isencomment" type='radio' <%If isencomment=0 Then Response.Write " checked='checked'" End If%>> 否</label></td>
						</tr>
						<tr>
							<td class='win_table_td'>仅好友可见：</td>
							<td><label><input value='1' name='ishide' id="ishide" <%If ishide=1 Then Response.Write " checked='checked'" End If%> type='radio'>是</label>  <label><input value='0' id="ishide" name='ishide' type='radio' <%If ishide=0 Or ishide="" Then Response.Write " checked='checked'" End If%>> 否</label></td>
						</tr>
						<tr>
							<td class='win_table_td'>首页固顶：</td>
							<td><label><input value='1' name='istop' id="istop" <%If istop=1 Then Response.Write " checked='checked'" End If%> type='radio'>是</label>  <label><input value='0' name='istop' id="istop"  type='radio' <%If istop=0 Or istop="" Then Response.Write " checked='checked'" End If%>> 否</label></td>
						</tr>
						<tr>
							<td class='win_table_td'>登录可见：</td>
							<td><label><input value='1' name='isneedlogin' id="isneedlogin" <%If isneedlogin=1 Then Response.Write " checked='checked'" End If%> type='radio'>是</label>  <label><input value='0' name='isneedlogin' id="isneedlogin"  type='radio' <%If isneedlogin=0 Then Response.Write " checked='checked'" End If%>> 否</label><font color=red>(新功能)</font></td>
						</tr>
						<tr>
							<td class='win_table_td'><label for="viewscores">浏览积分：</label></td>
							<td><input name='viewscores' id='viewscores' size='30' value='<%=ob_IIF(viewscores,0)%>'><font color=red>(新功能)</font></td>
						</tr>
						<tr>
							<td class='win_table_td'><label for="viewgroupid">浏览用户组：</label></td>
							<td>
								<select size=1 name="viewgroupid">
									<option value="0">----不限制----</option>
<%=sGroups%>
								</select>
		<script>
	var jobObject = document.oblogform["viewgroupid"];
	for(var i = 0; i < jobObject.options.length; i++) {
		if (jobObject.options[i].value=="<%=OB_IIF(viewgroupid,0)%>")
		{
			jobObject.selectedIndex = i;
		}
	}
		</script><font color=red>(新功能)</font></td>
						</tr>
						<tr>
							<td class='win_table_td'><label for="ispassword">日志密码：</label></td>
							<td><input name='ispassword' id='ispassword' size='30' value='<%=ispassword%>' onfocus="this.value=''"></td>
						</tr>
						<tr>
							<td class='win_table_td'><label for="showword">部分显示字数：</label></td>
							<td><input name="showword" id="showword" size='30'value="<%if showword<>"" then Response.Write(showword) else Response.Write(oblog.l_uShowlogWord)%>"></td>
						</tr>
						<tr>
							<td class='win_table_td'><label for="filename">文件名称：</label></td>
							<td><input name='filename' id='filename' size='30' value="<%=filename%>" onClick="chkfilename();"  onBlur="chkfilename();">.<%=f_ext%><br/> (只能为英文、数字及下划线，最长30个字符。)</td>
						</tr>
						<tr>
							<td class='win_table_td'><label for="s1">共同撰写：</label></td>
							<td><select name="blogteam" id="s1" onChange="setSort(this,this.form.s2);">
							<option value="<%=authorid%>" >我的BLOG</option>
						<%
						set rs=oblog.execute("select a.mainuserid,b.blogname,a.id from oblog_blogteam a,[oblog_user] b where a.otheruserid="&oblog.l_uid&" and a.mainuserid=b.userid")
						while not rs.eof
							if CLng(rs(0))=CLng(userid) then
								Response.Write "<option value="&rs(0)&" selected>"&oblog.filt_html(rs(1))&"</option>"
							else
								Response.Write "<option value="&rs(0)&">"&oblog.filt_html(rs(1))&"</option>"
							end if
							rs.movenext
						wend
						%>
							</select>
							<label for="s2">专题：</label>
							<select name="blogteamsubject" id="s2">
						<%
						if oblog.l_uid<>userid and userid>0 then
							set rs=oblog.execute("select subjectid,subjectname from oblog_subject where userid="&userid)
							Response.Write("<option value=0>我的专题</option>")
							while not rs.eof
								if subjectid=rs(0) then
									Response.Write("<option value="&rs(0)&" selected>"&rs(1)&"</option>")
								else
									Response.Write("<option value="&rs(0)&" >"&rs(1)&"</option>")
								end if
								rs.movenext
							wend
						else
							Response.Write("<option value=0 selected>我的专题</option>")
						end if
						set rs=nothing
						%>
							</select>
							</td>
						</tr>
						<tr>
							<td class='win_table_td'>发表时间：</td>
							<td><%show_selectdate(addtime)%></td>
						</tr>
						<tr>
							<td colspan='2' class="win_table_end"><input type="button" onClick="return checkValue();" value=" 确 定 " title=" 确 定 " />&nbsp;&nbsp;&nbsp;<input type="button" onClick="ResetValue(1);setSort(this,this.form.s2);return doMenu('swin1');" value=" 取 消 " title=" 取 消 " /></td>
						</tr>
					</table>
				</div>
				<div id="swin2" style="display:none;position:absolute;top:34px;left:93px;z-index:100;">
					<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
						<tr>
							<td align='center' class='win_table_top'>引用通告</td>
						</tr>
						<tr>
							<td><p><label for="tb">引用通告(支持多个引用通告，每一行为一个)：</label></p>
								<textarea name="tb" type="text" id="tb" rows="3" cols="80"><%=oblog.filt_html(tburl)%></textarea></td>
						</tr>

						<tr>
							<td class="win_table_end"><input type="button" onClick="return doMenu('swin2');" value=" 确 定 " title=" 确 定 " />&nbsp;&nbsp;&nbsp;<input type="button" onClick="ResetValue(2);return doMenu('swin2');" value=" 取 消 " title=" 取 消 " /></td>
						</tr>
					</table>
				</div>
				<div id="swin3" style="display:none;position:absolute;top:34px;left:176px;z-index:100;">
					<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
						<tr>
							<td align='center' class='win_table_top'>上传文件</td>
						</tr>
						<tr>
							<td><iframe id='d_file' frameborder='0' src='upload.asp?tMode=<%=t%>&re=' width='100%' height='60' scrolling='no'></iframe></td>
						</tr>
						<tr>
							<td class="win_table_end"><input type="button" onClick="return doMenu('swin3');" value=" 确 定 " title=" 确 定 " /></td>
						</tr>
					</table>
				</div>
				<div id="swin4" style="display:none;position:absolute;top:34px;left:259px;z-index:100;">
					<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
						<tr>
							<td align='center' class='win_table_top'>文章摘要</td>
						</tr>
						<tr>
							<td><label for="abstract"><p>文章摘要(请填写您文章的简要内容)： </p>
								<p>内容摘要不支持Html格式,且应小于500字符</p></label>
							<textarea name="abstract" style="width: 100%; " type="text" id="abstract" rows="6" cols="50"><%=log_abstract%></textarea>
							<textarea style="display:none;" name="log_files" type="text" id="log_files" size="60" rows=0 cols=0><%=log_files%></textarea>
							</td>
						</tr>
						<tr>
							<td class="win_table_end"><input type="button" onClick="return doMenu('swin4');" value=" 确 定 " title=" 确 定 "/>&nbsp;&nbsp;&nbsp;<input type="button" onClick="ResetValue(3);return doMenu('swin4');" value=" 取 消 " title=" 取 消 " /></td>
						</tr>
					</table>
				</div>
				<div id="swin5"></div>
				<iframe id="DivShim" scrolling="no" frameborder="0" style="position:absolute;top:0px; left:0px;display:none"></iframe>
				</form>
<%
End Sub

Sub show_selectdate(addtime)
    Dim y, m, d, h, mi, s, ttime
    If addtime = "" Then ttime = oblog.ServerDate(Now()) Else ttime = addtime
    Response.Write ("<select name=selecty id=selecty>")
    For y = Year(Now())-10 To Year(Now())+10
        If Year(ttime) = y Then
            Response.Write "<option value="&y&" selected>"&y&"年</option>"
        Else
            Response.Write "<option value="&y&">"&y&"年</option>"
        End If
    Next
    Response.Write "</select><select name=selectm id=selectm >"
    For m = 1 To 12
        If Month(ttime) = m Then
            Response.Write "<option value="&m&" selected>"&m&"月</option>"
        Else
            Response.Write "<option value="&m&">"&m&"月</option>"
        End If
    Next
    Response.Write ("</select><select name=selectd id=selectd >")
    For d = 1 To 31
        If Day(ttime) = d Then
            Response.Write "<option value="&d&" selected>"&d&"日</option>"
        Else
            Response.Write "<option value="&d&">"&d&"日</option>"
        End If
    Next
    Response.Write ("</select><select name=selecth id=selecth>")
    For h = 0 To 23
        If Hour(ttime) = h Then
            Response.Write "<option value="&h&" selected>"&h&"时</option>"
        Else
            Response.Write "<option value="&h&">"&h&"时</option>"
        End If
    Next
    Response.Write ("</select><select name=selectmi id=selectmi>")
    For mi = 0 To 59
        If Minute(ttime) = mi Then
            Response.Write "<option value="&mi&" selected>"&mi&"分</option>"
        Else
            Response.Write "<option value="&mi&">"&mi&"分</option>"
        End If
    Next
    Response.Write ("</select>")
End Sub
sub getteam()
	dim s,i,s1,rs,rs1
	set rs=Server.CreateObject("adodb.recordset")
	rs.open "select oblog_blogteam.mainuserid,[oblog_user].blogname from oblog_blogteam,[oblog_user] where oblog_blogteam.otheruserid="&oblog.l_uid&" and oblog_blogteam.mainuserid=[oblog_user].userid",conn,1,1
	if not rs.eof then
		Response.write "<script language=""JavaScript"">"&vbcrlf
		s = "var p_array = new Array(" + cstr(rs.recordcount-1) + ");"&vbcrlf
		Response.write s
		s = "var p_array_id = new Array(" + cstr(rs.recordcount-1) + ");"&vbcrlf
		Response.write s
		i = 0
		while not rs.eof
			set rs1=Server.CreateObject("adodb.recordset")
			rs1.open "select subjectid,subjectname from oblog_subject where userid="&rs("mainuserid"),conn,1,1
			s = "var p"+cstr(rs("mainuserid"))+"_array = Array("
			s1 = "var p"+cstr(rs("mainuserid"))+"_array_id = Array("
			if rs1.recordcount > 0 then
			while not rs1.eof
				if Trim(rs1("subjectname"))<>"" then
					s = s + """" + oblog.filt_html(rs1("subjectname")) + """"
					s1 = s1 + """" + cstr(rs1("subjectid")) + """"
					s = s + ","
					s1 = s1 + ","
				end if
				rs1.movenext
			wend
			s = s + """" + "不选择专题" + """"
			s1 = s1 + """" + cstr(0) + """"
			else
				s = s + """" + "无可用专题" + """"
				s1 = s1 + """" + cstr(0) + """"
			end if
			s = s+ ");"&vbcrlf
			s1 = s1+ ");"&vbcrlf
			Response.write s
			Response.write s1
			Response.write "p_array["+cstr(i)+"] = p"+cstr(rs("mainuserid"))+"_array;"&vbcrlf
			Response.write "p_array_id["+cstr(i)+"] = p"+cstr(rs("mainuserid"))+"_array_id;"&vbcrlf
			i = i + 1
			rs.movenext
		wend
		Response.write  "</script>"&vbcrlf
		rs.close
		set rs=nothing
		rs1.close
		set rs1=nothing
	end if
end Sub

Private Function GetUserTags()
	Dim rs,strTemp
	Dim i
	i = 0
	Set rs = oblog.Execute ("SELECT TOP 10 name FROM oblog_tags a INNER JOIN (SELECT distinct tagid AS x FROM oblog_usertags  WHERE userid = "&oblog.l_uid&") b ON a.tagid = b.x ORDER BY inum DESC" )
	Do While Not rs.Eof
		strTemp = strTemp &"										<a href=""#"" id=""a"&i&""" onclick=""setTags('a"&i&"')"" title="""&rs(0)&""">"&rs(0)&"</a>" & vbcrlf
		i = i + 1
		rs.MoveNext
	Loop
	GetUserTags = strTemp
	strTemp = ""
End Function
%>
<%If C_Editor_Type=1 Then oblog.MakeEditorText "",1,"535","250"%>
<script>
setInterval(autoSetDraft,1000)
	function setTags(obj)
	{
		var tags=document.getElementById('logtags');
		var stags = document.getElementById(obj).innerHTML;
		if (del_space(tags.value).length==0)
		{
			tags.value=stags;
//			document.getElementById('logtags_used').style.display="none";
		}
		else
		{
			if (tags.value==stags)
			{
//					alert('请勿重复选择');
					return false;
			}
			else
			{
				if (tags.value.indexOf(stags+'<%=P_TAGS_SPLIT%>')>=0||tags.value.indexOf('<%=P_TAGS_SPLIT%>'+stags)>=0)
					{
//						alert('请勿重复选择');
						return false;
						}
					else
					{
						tags.value+='<%=P_TAGS_SPLIT%>'+stags;
//						document.getElementById('logtags_used').style.display="none";
					}
			}
		}
	}
</script>