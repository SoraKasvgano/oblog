<!--#include file="user_top.asp"-->
<!--#include file="inc/class_Trackback.asp"-->
<script src="oBlogStyle/move.js" type="text/javascript"></script>
<script src="inc/function.js" type="text/javascript"></script>
<%
If oblog.l_uNewbie=1 Then
	Response.write("<script>parent.show_title('ѡ��ģ��')</script>")
	oblog.ShowMsg "����ǰ����ѡ��һ��ϲ����ģ�档","user_template.asp?action=showconfig"
	'jscmd="go_cmdurl('ѡ��ģ��','tab3')"
end If
If oblog.Chkiplock() Then
	oblog.ShowMsg ("�Բ������IP�ѱ������������������"),blogdir &"index.html"
	Set oblog = Nothing
End If
If Application(cache_name_user&"_systemenmod")<>"" Then
	Dim enStr
	enStr=Application(cache_name_user&"_systemenmod")
	enStr=Split(enStr,",")
	If enStr(2)="1" Then	Response.write("ϵͳ��ʱ��ֹ������־�����!"):Response.End()
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
					<li><a href="#" onclick="purl('user_post.asp?t=1','FLASH')">�ƣ̣��ӣ�</a></li>
					<li><a href="#" onclick="purl('user_photo.asp','��Ƭ����')">��Ƭ����</a></li>
					<li><a href="#" onclick="purl('user_Albumcomments.asp','��Ƭ����')">��Ƭ����</a></li>
					<li><a href="#" onclick="purl('user_subject.asp?t=1','��Ƭ����')">��Ƭ����</a></li>

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
						document.write('\����flash�汾���ͣ�<a href="http:\/\/www.adobe.com\/go\/getflash\/" target="_blank"\>��������Flash Player�����֧�ִ�ͷ������</a>');
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
					<li><a href="#" onClick="return doMenu('swin1');" <%=sDisable%>>�߼�ѡ��</a></li>
					<li><a href="#" onClick="return doMenu('swin2');" <%=sDisable%>>����ͨ��</a></li>
					<li><a href="#" onClick="return doMenu('swin3');" <%=sDisable%>>�ϴ��ļ�</a></li>
					<li><a href="#" onClick="return doMenu('swin4');" <%=sDisable%>>����ժҪ</a></li>
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
            oblog.adderrstr ("�޴�Ȩ�޲�����" & tName & "��")
            oblog.showusererr
        End If
		if rs("isdel")=1 then
			set rs=nothing
			oblog.ShowMsg "��ɾ����־�����Ȼָ����ٲ�����",""
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
		If ispassword<>"" Then ispassword="�������룬�����޸��벻Ҫ����"
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
    	'��ⷢ������
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
		if filename="" then filename="�Զ����"
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
		alert ('����ѡ���������޶���־����');
		return false;
	}
	if (document.oblogform.viewscores.value>0&&document.oblogform.abstract.value==''){
		alert('���û������������£���������д����ժҪ��');
	}
	if (document.oblogform.isneedlogin[0].checked==true&&document.oblogform.abstract.value==''){
		alert('���õ�¼�ɼ�������£���������д����ժҪ��');
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
parent.show_title("������־");
var in_ob_useradmin=true;
var issubmit=false;
function chkfilename()
{
	var filename=del_space(document.oblogform.filename.value);
	if (filename=="�Զ����"){document.oblogform.filename.value=""}
	if (filename==""){document.oblogform.filename.value="�Զ����"}
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
//����Ϊ�Զ�����ݸ�Ĳ�������
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
	save_ing.innerHTML='<font color=red>'+Secs +'</font>��󱣴浽�ݸ���';
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
//����Ϊ�Զ�����ݸ�ĺ���
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
	document.getElementById("save_ing").innerHTML="<img src='images/loading.gif' align='absbottom'> ���ڱ���...";
	if (issubmit){
		var oDialog = new dialog("<%=blogurl%>");
		oDialog.init();
	 	oDialog.event("�����ύ�У����Ժ�...",'');
		oDialog.button('dialogOk',"");
	}else{
		var errstr=""
		<%If C_Editor_Type=1 Then%>document.oblogform.edit.value=oEdit1.getHTMLBody();<%End if%>
		var topic = del_space(document.oblogform.topic.value);
		 if (topic.length == 0&&document.oblogform.isdraft.value!="1")
		 {
			errstr="��������д��Ŀ��<br />";
		 }
		 var needclassid=<%=oblog.CacheConfig(25)%>;
		 if (needclassid==1 &&document.oblogform.isdraft.value!="1") {
		 if (document.oblogform.classid.value == 0)
		 {
			errstr=errstr+"��ѡ����־�����<br />";
		 }
		 }

		var filename=del_space(document.oblogform.filename.value);
		if ((checkerr(filename) == "1")&&(filename!="")&&(filename!="�Զ����")){
			errstr=errstr+"�ļ���������0-9�����ֺ�a-z�İ����ĸ���»���,���������ĺ͹��ַ�����!@#$%^�ȣ���<br />";
		}

		if (document.oblogform.edit.value == "")
		 {
			errstr=errstr+"��������־�����ݡ�<br />";
		 }
		var date=document.oblogform.selecty.value+"-"+document.oblogform.selectm.value+"-"+document.oblogform.selectd.value
		var datereg=/^(\d{4})-(\d{1,2})-(\d{1,2})$/
		var datareg=/^(\d){1,2}$/
		if (!datereg.test(date)){
		  errstr=errstr+"��־ʱ�������ʽ����<br />";
		 }
		var r=date.match(datereg)
		var d=new Date(r[1],r[2]-1,r[3])
		if (!(d.getFullYear()==r[1]&&d.getMonth()==r[2]-1&&d.getDate()==r[3])){
		  errstr=errstr+"��־ʱ�������ʽ����<br />";
		 }
		 <%if oblog.l_Group(31,0)=1 then%>
		if (document.oblogform.codestr.value== ""&&isauto==false)
		{
			errstr=errstr+"��������֤�롣<br />";
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
			oDialog.button('dialogOk',"document.getElementById('ob_submit').disabled=false;document.getElementById('ob_submit_d').disabled=false;document.getElementById('ob_submit_p').disabled=false;"); //���޸���ϡ�
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
	var str='<div style="height:200px;overflow:auto">һ��ʲô�Ǳ�ǩ��TAG����<br />�����򵥵�˵,��ǩ����һƪ���µ�"�ؼ���"�������Խ���־���»�����Ƭ��ѡ��һ�����������ǩ������ǣ�����һ�����������ǲ�����վ��ʹ�øô���������Զ���Ϊһ���б���ʾ��<br />����ʹ�ñ�ǩ�ĺô���<br />����1������ӱ�ǩ�����¾ͻᱻֱ���ӵ���վ��Ӧ��ǩ��ҳ�棬����������ڷ�����ر�ǩʱ�����п��ܷ��ʵ��������£��������������±����ʵĻ��ᡣ<br />����2�������Ժܷ���ز��ҵ�����ʹ����ͬ����ǩ�����£����������µ���Ұ�����Է���ز��ҵ�����ʹ����ͬ����ǩ�����ߣ���Ϊ־ͬ���ϵ����ѣ������Խ����Ǽ�Ϊ���ѻ����鲩�ͣ�������������Ȧ�� <br />����3�����ӱ�ǩ�ķ�ʽ��ȫ�������������������κε����ƣ���������վϵͳ������Լ�ԭ����־��������ƣ�������Ϣ����������Ͳ��ҡ�<br />�������ʹ�ñ�ǩ?<br />�������磺��д��һƪ���������ε����£����������ᵽ�����ݣ������Ը���ƪ���¼��ϣ�<br />������������,�찲��,����,�ʹ�<br />�����ȼ�����ǩ������������������ڳ��ǵ�����ʱ������߻�����ǩ�����ǣ��Ӷ��������й��ڳ��ǵ����£�����������߲�����־��ͬʱ��Ҳ���ô˷����ҵ�����ͬ��ϲ�����ˣ��Ա�һ���໥�����ȵȡ�<br />�ģ������ӡ��á���Tag�� <br />����1�� TagӦ��Ҫ�ܹ����ֳ��Լ�����ɫ�������Ǵ�Ҿ������õ���Ϥ�Ĵ��<br />����2���ôʾ����򵥾���,����������Ҫ̫�����������ֵĴ���Ϳ����ˣ�������������Ĵʻ㣬��Ҫʹ��һЩֻ��Ϊװ�εķ��ţ�������ȡ�<br />����3����Ҫʹ��һЩ����Ƚ����Ĵʻ㣬�硰�ҵļҡ�����ͼƬ���ȡ�<br /></div>'
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
	s2.options[0].text="�ҵ�ר��";
	s2.options[0].value="0";
	}
}
function CheckTeamID(tValue){
	var TeamPost=true;
	<%if oblog.CacheConfig(72) = "0" then %>
	alert('ϵͳ��ֹ��־ͬʱ������<%=oblog.CacheConfig(69)%>');
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
				alert('��ƪ��־ͬʱ������<%=oblog.CacheConfig(69)%>��������<%=oblog.CacheConfig(72)%>');
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
									<span><label for="topic">����</label></span>
									<input type="text" name="topic" id="topic" size="40" maxlength="50" value="<%=AnsiToUnicode(topic)%>" />
									<select name="classid" id="classid" <%If oblog.CacheConfig(25)="1" Then Response.Write "class=""blue"" title=""������дϵͳ����"" " %>>
										<%=oblog.show_Postclass(classid)%>
									</select>
									<select name="subjectid" id="subjectid">
										<option value="0">�ҵķ���</option>
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
									<span><label for="logtags">��ǩ</label></span>
									<a><input name="logtags" id="logtags" type="text" size="40" maxlength="255" value="<%=tags%>" /></a><a id="usedtags" href="#logtags_used"><img src="oBlogStyle/UserAdmin/7/user_team_top.png" alt="����ʹ�ù��ı�ǩ" /></a>
									��<font class="blue"><%
									select Case P_TAGS_SPLIT
										Case " "
											Response.Write "�ո�"
										Case ","
											Response.Write "����"
										Case Else
											Response.Write P_TAGS_SPLIT
									End select
									%></font>�ָ�
									��<a href="#" onclick="taghelp();">ʲô�Ǳ�ǩ��</a>��
									<div id="logtags_used">
<%=GetUserTags%>
									</div>
									<script language="JavaScript" src="oBlogStyle/UserAdmin/used.js" type="text/javascript"></script>
								</div>
								<div class="d3">
<%
'----------------------------------------------------------
'�༭����ʾ
Select Case C_Editor_Type
Case 1
	Dim EditorHeight
	if oblog.l_Group(31,0)=1 then
		EditorHeight=250
	Else
		EditorHeight=250
	End If
%>
									<span id="loadedit"><img src='images/loading.gif' align='absbottom'>��������༭������</span>
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
									<label>��֤�룺<input name="codestr" type="text" size="6" maxlength="20" /></label> <%=oblog.getcode%>
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
										<input type="button" id="ob_submit"  value="������־" title="������־" onclick="savelog();" />
										<input type="button" id="ob_submit_d" value="����Ϊ�ݸ�" title="����Ϊ�ݸ�" onClick="setdraft();" />
										<input type="button" id="ob_submit_p" onClick="<%If C_Editor_Type=1 Then%>oEdit1.insertHTML('#��ǰ����ҳ������ʾ#')<%ElseIf C_Editor_Type=2 Then%>part();<%End If%>" value="������ʾ���" />
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
									<div>ͬʱ������<%=oblog.CacheConfig(69)%></div>
									<ul>

<%
Dim rst
Set rs=oblog.Execute("select a.teamid,a.t_name From oblog_team a ,(select  teamid From oblog_teamusers Where userid=" & oblog.l_Uid & " and state>=3) b Where a.teamid=b.teamid And a.istate=3")
If rs.Eof Then
	'Response.Write "									<li class=""red"">û�м���<a href=""#"" onclick=""purl('user_team.asp','" &oblog.CacheConfig(69)& "')"">" &oblog.CacheConfig(69)& "</a>��</li>"
	Response.Write "									<li class=""green"">û�м���" &oblog.CacheConfig(69)& "��</li>"
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
							<td colspan='2' align='center' class='win_table_top'>�߼�ѡ��(������־����ѡ��һ���޶�)</td>
						</tr>
						<tr>
							<td class='win_table_td'>�������ۣ�</td>
							<td><label><input value='1' name='isencomment' id="isencomment" <%If isencomment=1 Or isencomment="" Then Response.Write " checked='checked'" End If%> type='radio'>��</label>  <label><input value='0' name='isencomment' id="isencomment" type='radio' <%If isencomment=0 Then Response.Write " checked='checked'" End If%>> ��</label></td>
						</tr>
						<tr>
							<td class='win_table_td'>�����ѿɼ���</td>
							<td><label><input value='1' name='ishide' id="ishide" <%If ishide=1 Then Response.Write " checked='checked'" End If%> type='radio'>��</label>  <label><input value='0' id="ishide" name='ishide' type='radio' <%If ishide=0 Or ishide="" Then Response.Write " checked='checked'" End If%>> ��</label></td>
						</tr>
						<tr>
							<td class='win_table_td'>��ҳ�̶���</td>
							<td><label><input value='1' name='istop' id="istop" <%If istop=1 Then Response.Write " checked='checked'" End If%> type='radio'>��</label>  <label><input value='0' name='istop' id="istop"  type='radio' <%If istop=0 Or istop="" Then Response.Write " checked='checked'" End If%>> ��</label></td>
						</tr>
						<tr>
							<td class='win_table_td'>��¼�ɼ���</td>
							<td><label><input value='1' name='isneedlogin' id="isneedlogin" <%If isneedlogin=1 Then Response.Write " checked='checked'" End If%> type='radio'>��</label>  <label><input value='0' name='isneedlogin' id="isneedlogin"  type='radio' <%If isneedlogin=0 Then Response.Write " checked='checked'" End If%>> ��</label><font color=red>(�¹���)</font></td>
						</tr>
						<tr>
							<td class='win_table_td'><label for="viewscores">������֣�</label></td>
							<td><input name='viewscores' id='viewscores' size='30' value='<%=ob_IIF(viewscores,0)%>'><font color=red>(�¹���)</font></td>
						</tr>
						<tr>
							<td class='win_table_td'><label for="viewgroupid">����û��飺</label></td>
							<td>
								<select size=1 name="viewgroupid">
									<option value="0">----������----</option>
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
		</script><font color=red>(�¹���)</font></td>
						</tr>
						<tr>
							<td class='win_table_td'><label for="ispassword">��־���룺</label></td>
							<td><input name='ispassword' id='ispassword' size='30' value='<%=ispassword%>' onfocus="this.value=''"></td>
						</tr>
						<tr>
							<td class='win_table_td'><label for="showword">������ʾ������</label></td>
							<td><input name="showword" id="showword" size='30'value="<%if showword<>"" then Response.Write(showword) else Response.Write(oblog.l_uShowlogWord)%>"></td>
						</tr>
						<tr>
							<td class='win_table_td'><label for="filename">�ļ����ƣ�</label></td>
							<td><input name='filename' id='filename' size='30' value="<%=filename%>" onClick="chkfilename();"  onBlur="chkfilename();">.<%=f_ext%><br/> (ֻ��ΪӢ�ġ����ּ��»��ߣ��30���ַ���)</td>
						</tr>
						<tr>
							<td class='win_table_td'><label for="s1">��ͬ׫д��</label></td>
							<td><select name="blogteam" id="s1" onChange="setSort(this,this.form.s2);">
							<option value="<%=authorid%>" >�ҵ�BLOG</option>
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
							<label for="s2">ר�⣺</label>
							<select name="blogteamsubject" id="s2">
						<%
						if oblog.l_uid<>userid and userid>0 then
							set rs=oblog.execute("select subjectid,subjectname from oblog_subject where userid="&userid)
							Response.Write("<option value=0>�ҵ�ר��</option>")
							while not rs.eof
								if subjectid=rs(0) then
									Response.Write("<option value="&rs(0)&" selected>"&rs(1)&"</option>")
								else
									Response.Write("<option value="&rs(0)&" >"&rs(1)&"</option>")
								end if
								rs.movenext
							wend
						else
							Response.Write("<option value=0 selected>�ҵ�ר��</option>")
						end if
						set rs=nothing
						%>
							</select>
							</td>
						</tr>
						<tr>
							<td class='win_table_td'>����ʱ�䣺</td>
							<td><%show_selectdate(addtime)%></td>
						</tr>
						<tr>
							<td colspan='2' class="win_table_end"><input type="button" onClick="return checkValue();" value=" ȷ �� " title=" ȷ �� " />&nbsp;&nbsp;&nbsp;<input type="button" onClick="ResetValue(1);setSort(this,this.form.s2);return doMenu('swin1');" value=" ȡ �� " title=" ȡ �� " /></td>
						</tr>
					</table>
				</div>
				<div id="swin2" style="display:none;position:absolute;top:34px;left:93px;z-index:100;">
					<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
						<tr>
							<td align='center' class='win_table_top'>����ͨ��</td>
						</tr>
						<tr>
							<td><p><label for="tb">����ͨ��(֧�ֶ������ͨ�棬ÿһ��Ϊһ��)��</label></p>
								<textarea name="tb" type="text" id="tb" rows="3" cols="80"><%=oblog.filt_html(tburl)%></textarea></td>
						</tr>

						<tr>
							<td class="win_table_end"><input type="button" onClick="return doMenu('swin2');" value=" ȷ �� " title=" ȷ �� " />&nbsp;&nbsp;&nbsp;<input type="button" onClick="ResetValue(2);return doMenu('swin2');" value=" ȡ �� " title=" ȡ �� " /></td>
						</tr>
					</table>
				</div>
				<div id="swin3" style="display:none;position:absolute;top:34px;left:176px;z-index:100;">
					<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
						<tr>
							<td align='center' class='win_table_top'>�ϴ��ļ�</td>
						</tr>
						<tr>
							<td><iframe id='d_file' frameborder='0' src='upload.asp?tMode=<%=t%>&re=' width='100%' height='60' scrolling='no'></iframe></td>
						</tr>
						<tr>
							<td class="win_table_end"><input type="button" onClick="return doMenu('swin3');" value=" ȷ �� " title=" ȷ �� " /></td>
						</tr>
					</table>
				</div>
				<div id="swin4" style="display:none;position:absolute;top:34px;left:259px;z-index:100;">
					<table class='win_table' align='center' border='0' cellpadding='0' cellspacing='1'>
						<tr>
							<td align='center' class='win_table_top'>����ժҪ</td>
						</tr>
						<tr>
							<td><label for="abstract"><p>����ժҪ(����д�����µļ�Ҫ����)�� </p>
								<p>����ժҪ��֧��Html��ʽ,��ӦС��500�ַ�</p></label>
							<textarea name="abstract" style="width: 100%; " type="text" id="abstract" rows="6" cols="50"><%=log_abstract%></textarea>
							<textarea style="display:none;" name="log_files" type="text" id="log_files" size="60" rows=0 cols=0><%=log_files%></textarea>
							</td>
						</tr>
						<tr>
							<td class="win_table_end"><input type="button" onClick="return doMenu('swin4');" value=" ȷ �� " title=" ȷ �� "/>&nbsp;&nbsp;&nbsp;<input type="button" onClick="ResetValue(3);return doMenu('swin4');" value=" ȡ �� " title=" ȡ �� " /></td>
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
            Response.Write "<option value="&y&" selected>"&y&"��</option>"
        Else
            Response.Write "<option value="&y&">"&y&"��</option>"
        End If
    Next
    Response.Write "</select><select name=selectm id=selectm >"
    For m = 1 To 12
        If Month(ttime) = m Then
            Response.Write "<option value="&m&" selected>"&m&"��</option>"
        Else
            Response.Write "<option value="&m&">"&m&"��</option>"
        End If
    Next
    Response.Write ("</select><select name=selectd id=selectd >")
    For d = 1 To 31
        If Day(ttime) = d Then
            Response.Write "<option value="&d&" selected>"&d&"��</option>"
        Else
            Response.Write "<option value="&d&">"&d&"��</option>"
        End If
    Next
    Response.Write ("</select><select name=selecth id=selecth>")
    For h = 0 To 23
        If Hour(ttime) = h Then
            Response.Write "<option value="&h&" selected>"&h&"ʱ</option>"
        Else
            Response.Write "<option value="&h&">"&h&"ʱ</option>"
        End If
    Next
    Response.Write ("</select><select name=selectmi id=selectmi>")
    For mi = 0 To 59
        If Minute(ttime) = mi Then
            Response.Write "<option value="&mi&" selected>"&mi&"��</option>"
        Else
            Response.Write "<option value="&mi&">"&mi&"��</option>"
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
			s = s + """" + "��ѡ��ר��" + """"
			s1 = s1 + """" + cstr(0) + """"
			else
				s = s + """" + "�޿���ר��" + """"
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
//					alert('�����ظ�ѡ��');
					return false;
			}
			else
			{
				if (tags.value.indexOf(stags+'<%=P_TAGS_SPLIT%>')>=0||tags.value.indexOf('<%=P_TAGS_SPLIT%>'+stags)>=0)
					{
//						alert('�����ظ�ѡ��');
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