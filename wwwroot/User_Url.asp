<!--#include file="user_top.asp"-->
<%
If oblog.l_Group(18,0)=0 Then
	oblog.AddErrStr ("��Ŀǰ�����ĵȼ�������ʹ�ö��Ĺ���")
    oblog.showUserErr
    Response.End
End if
Dim Action
Dim rs,rsSubject,UrlId,Id,Ids,targetSubjectid,Sql,OutRssDisplay,allsub
Dim sTitle,sUrl,sClassId,sSubjectId,sTags,sMemo,isPrivate,encodeing,mainuserid
Action = LCase(Trim(Request("action")))
Set rs=Server.CreateObject("Adodb.Recordset")
Id=Request("Id")
G_P_FileName="user_url.asp?page="
If Id<>"" And InStr(Id,",")<=0 Then Id=clng(Id)
if En_OutRss=0 then OutRssDisplay="style='display:none'"
'����һ��������Ϣ
Set rsSubject=oblog.Execute("Select * From oblog_subject Where userid=" & oblog.l_uid & " And subjecttype=3")
Select Case Action
	Case "save"
		Call Save
	Case "add","edit"
		Call EditForm
	Case "del"
		'ɾ��
		If Id<>"" Then
			Ids=FilterIds(Id)
			conn.Execute("Delete From oblog_myurl Where Id In(" & Ids & ") And userid=" & oblog.l_uid)
		End If
		Response.Write "<script>if(top.location.href!=self.location.href)parent.getfeedlist();</script>"
		response.Flush()
		oblog.ShowMsg "ɾ���ɹ�!","user_url.asp"
	Case "bdel"
		'����ɾ��
		Ids=FilterIds(UrlId)
		If Ids<>"" Then
			conn.Execute("Delete From oblog_myurl Where Id in(" & Ids & ") And userid=" & oblog.l_uid)
			'�����Ӽ������Ĵ���
		End If
		Response.Write "<script>if(top.location.href!=self.location.href)parent.getfeedlist();</script>"
		response.Flush()
		oblog.ShowMsg "ɾ���ɹ�!","user_url.asp"
	Case "bmove"
		Ids=FilterIds(UrlId)
		targetSubjectid=clng(Request("subject"))
		If Ids<>"" Then 		'������ת��
			conn.Execute("Update oblog_myurl Set subjectid=" &targetSubjectid  &" Where Id in(" & Ids & ") And userid=" & oblog.l_uid)
		End If
		Response.Write "<script>if(top.location.href!=self.location.href)parent.getfeedlist();</script>"
		response.Flush()
		oblog.ShowMsg "�ƶ��ɹ�!","user_url.asp"
	case "read"
		call readrss()
	Case Else
		Call List
End Select
Set rs=Nothing

%>
</table>
</body>
</html>
<%
sub readrss()
	dim feedurl,url
	url = LCase(trim(request("feedurl")))
	If InStr(url,"http://") = 0 Then
'		url = oblog.CacheConfig(3) & url
	End if
	if trim(request("mainuserid"))="" then
		mainuserid=0
	else
		mainuserid=clng(trim(request("mainuserid")))
	end if
	if mainuserid>0 and true_domain=0 then
		feedurl = url
	else
		if trim(request("encodeing"))="gb2312" then
			feedurl="readrss.asp?feedurl="&url
		else
			feedurl="readrss_utf8.asp?feedurl="&url
		end if
	end if
	if mainuserid>0 then oblog.execute("update oblog_myurl set isupdate=0 where mainuserid="&mainuserid&" and userid="&oblog.l_uid)
	oblog.execute("update oblog_user set sub_num=0")
%>
<style type="text/css">
<!--
#user_page_top {border:0;width:100%;position:absolute;top:0;left:0;}
#user_page_top li {margin:5px 0 0 10px;padding:3px 0 0 20px;background: url("oBlogStyle/li/ico.gif") no-repeat -900px -217px;font-size:14px;color:#115888;font-weight:600;}
.msg {position:absolute;top:50%;left:50%;margin-left:-100px;margin-top:-50px;padding:15px;background:#FFFFE0 url("oBlogStyle/UserAdmin/7/BoxOver_bd.png") repeat-x left bottom;border-top:1px solid #EBEBA9;border-left:1px solid #EBEBA9;border-right:1px solid #C3C370;border-bottom:1px solid #C3C370;text-align:center;color:#f00;}
#rssbody {background:#DAE5EF url("oBlogStyle/UserAdmin/7/user_post_top_bg.png") repeat-x left top;padding:0;height:100%;overflow:hidden;}
.rssTitleList{
	/*��ǩΪUL,���������б�*/
	list-style:none;
	margin:29px 0 0 8px;
	float:left;
	width:33%;
	height:93%;
	overflow-x:hidden;
	overflow-y:auto;
	background:#fff;
	border-top:1px #1A76B7 solid;
	border-right:1px #9CCCEF solid;
	border-bottom:1px #9CCCEF solid;
	border-left:1px #1A76B7 solid;
}

.rssTitleList li {
	/*��ǩΪli,���������б�*/
	list-style:none;
	padding:2px 2px 2px 18px;
	font-size:12px;
	border-bottom:1px #ECF0F9 solid;
	background: url(oBlogStyle/li/ico.gif) no-repeat -895px -56px;
}

.rssTitleList li a {
	/*��ǩΪa,���������б�*/
	text-decoration:none;
	color:#115888;
}
.rssTitleList li a:hover {
	text-decoration:underline;
	color:#000;
}

.rsslist {
	/*��ǩΪUL,ȫ��rssȫ�������ڴ�*/
	margin:29px 8px 0 0;
	list-style:none;
	float:right;
	width:64%;
	height:93%;
	overflow-x:hidden;
	overflow-y:auto;
	background:#fff;
	border-top:1px #1A76B7 solid;
	border-right:1px #9CCCEF solid;
	border-bottom:1px #9CCCEF solid;
	border-left:1px #1A76B7 solid;
}

.rssTitle {
	/*��ǩΪP,����*/
	font-size:14px;
	font-weight:600;
	padding:5px 0 0 20px;
	margin:5px;
	background: url(oBlogStyle/li/ico.gif) no-repeat -900px -216px;
}

.rssTitle a {
	color:#f60;
	text-decoration:none;
	border-bottom:1px #dedede solid;
}

.rssTitle a:hover {
	text-decoration:none;
	border-bottom:1px #f60 solid;
}

.rssMemo {
	/*��ǩΪUL,����ʵ�����壬�˴���ΪIEDOM���ݲ�����UL��ǩ������ΪDIV �˴�Ϊ����*/
	color:#555;
	margin:5px 5px 15px 5px;
	padding:0 5px 15px 15px;
	list-style:none;
	line-height:1.5;
	border-bottom:1px #ECF0F9 solid;
}
.rssMemo a {
	color:#1A76B7;
}
.rssMemo a:hover {
	color:#f60;
}

.floatTime {
	/*����ʱ��*/
	margin:0 0 0 20px;
	color:#999;
	font-size:12px;
}

-->
</style>
<ul id="user_page_top" >
	<li><%=request("title")%></li>
</ul>
<div id="rssbody"></div>
<script language="javascript" type="text/javascript">

var Class = {
  create: function() {
    return function() {
      this.initialize.apply(this, arguments);
    }
  }
}

var rssReader = Class.create();
rssReader.prototype = {
	initialize: function(url) {
		this.url = url;
		this.http_request = false;
		this.titlelist = document.createElement("ul");
		this.titlelist.setAttribute('className','rssTitleList');
		this.showload();
		this.getRss();
	},

	showload:function(){
		var loading = document.createElement('span');
		loading.setAttribute('id','loading');
		//var text = document.createTextNode("<img src='images/loading.gif'>");
		//loading.appendChild(text);
		loading.innerHTML="<div class='msg'><img src='images/loading.gif'> ���ڼ���...</div>"
		document.getElementById("rssbody").appendChild(loading);
	},

	getRss:function() {
		if(window.XMLHttpRequest) {
			this.http_request = new XMLHttpRequest();
			if (this.http_request.overrideMimeType) {
				this.http_request.overrideMimeType('text/xml');
			}
		}
		else if (window.ActiveXObject) {
			try {
				this.http_request = new ActiveXObject(" Msxml2.XMLHTTP");
			} catch (e) {
				try {
					this.http_request = new ActiveXObject("Microsoft.XMLHTTP");
				} catch (e) {}
			}
		}
		if (!this.http_request) {
			window.alert ("���ܴ���XMLHttpRequest����ʵ��.");
			return false;
		}
		var othis = this;
		this.http_request.onreadystatechange = function(){othis.callback();};
//		alert(this.url);
		this.http_request.open("get", this.url, true);
		this.http_request.send(null);
	},

	callback:function() {
		if (this.http_request.readyState == 4) {
			if (this.http_request.status == 200) {
//					alert(this.http_request.Responsetext);
//					alert(this.url);
					this.loadready(this.http_request.responseXML);
			} else {
//				alert(this.url);
				alert(this.http_request.Responsetext)
			//	alert("���������ҳ�����쳣��");
			}
		}
	},

	loadready:function(xml){
		var xol = document.createElement("ul");
		xol.setAttribute('className','rsslist');
		var allitems = xml.getElementsByTagName("item");
		for(var i = 0;i < allitems.length;i++){
			var xli = document.createElement("li");
			var div = document.createElement("div");
			var a = document.createElement("a");
			var p = document.createElement("p");
			p.setAttribute('className','rssTitle');
		   	var title = document.createTextNode(allitems[i].getElementsByTagName("title")[0].firstChild.data);
			a.appendChild(title);
			a.setAttribute('href',allitems[i].getElementsByTagName("link")[0].firstChild.data);
			a.setAttribute('target','_blank');
			p.setAttribute('id','tag'+i);
			p.appendChild(a);
			xli.appendChild(p);
				var title = document.createTextNode(allitems[i].getElementsByTagName("title")[0].firstChild.data);
				var li = document.createElement("li");
				var a = document.createElement("a");
				a.appendChild(title);
				a.setAttribute('href','#tag'+i);
				li.appendChild(a);
				this.titlelist.appendChild(li);
			var pubtime = allitems[i].getElementsByTagName("pubDate")[0].firstChild.data;
			var timespan = document.createElement("span");
			timespan.setAttribute('className','floatTime');
			var old = new Date(Date.parse(pubtime));
 			var now = new Date();
			var tmptime;
			var year;
			var month;
			var date;
			var hours;
			if ( old == "NaN" ){
				tmptime = pubtime.match(/(\d{4})-([0-1]?\d)-(\d?\d)\s(\d?\d):(\d?\d):(\d?\d)/i);
				if(tmptime){
					year = tmptime[1];
					month = tmptime[2]-1;
					date = tmptime[3];
					hours = tmptime[4];
					timespan.appendChild(document.createTextNode(datedif()));
				}else{
					timespan.appendChild(document.createTextNode(pubtime));
				}

			}else{
				year = old.getYear();
				month = old.getMonth();
				date = old.getDate();
				hours = old.getHours();
				try{
					timespan.appendChild(document.createTextNode(datedif()));
				}catch(e){}
			}



			li.appendChild(timespan);
			function datedif(){
				if (now.getYear()-year==0)	{
					if(now.getMonth()-month==0){
						if(now.getDate()-date==0){
							if(now.getHours()-hours==0){
								return ("�ո�");
							}else{
								return (now.getHours()-hours+"Сʱǰ");
							}
						}else{
							return (now.getDate()-date+"��ǰ");
						}
					}else{
						return (now.getMonth()-month+"��ǰ");
					}
				}else{
					return (now.getYear()-year+"��ǰ");
				}
			}

			try{
				description = allitems[i].getElementsByTagName("description")[0].firstChild.data;
			}
			catch(e)
			{
				description = "no";
			}
			var desul = document.createElement("ul");
			var deli = document.createElement("li");

			try{deli.innerHTML=description;}catch(e){deli.innerText=description;};

			desul.appendChild(deli);
			desul.setAttribute('className','rssMemo');
			xli.appendChild(desul);
			xol.appendChild(xli)
		}
		document.getElementById("rssbody").appendChild(this.titlelist);
		document.getElementById("rssbody").appendChild(xol);
		var loading = document.getElementById('loading');
		document.getElementById("rssbody").removeChild(loading);
	}

}

function strToDate(str)
{
  var val=Date.parse(str);
  var newDate=new Date(val);
  return newDate;
}
window.onload = function(){var Rss = new rssReader('<%=feedurl%>');}

</script>



<%end sub

Sub EditForm()
	Dim rst
	mainuserid=trim(request("mainuserid"))
	sTitle=trim(request("sTitle"))
	sUrl=trim(request("sUrl"))
	If Id<>"" Then
	Id=FilterIds(Id)
		Set rst=oblog.Execute("Select * From oblog_myurl Where userid=" & oblog.l_uid & " And  Id=" & Id)
		If rst.Eof Then
			oblog.ShowMsg "�������Ϣ���",""
		Else
			sUrl=rst("Url")
			sTitle=rst("Title")
			sSubjectId=rst("subjectid")
			sTags=rst("tags")
			sMemo=rst("memo")
			isPrivate=rst("isprivate")
			encodeing=rst("encodeing")
		End If
		set rst=nothing
	End If
%>
<script language=javascript>

function VerifySubmit()
{
    if (document.oblogform.url.value.length==0){
    	alert("���ĵ�ַ������д");
    	document.oblogform.url.focus();
    	return false;
    	}
    	return true;
}

function createXMLHttpRequest() {
    try { return new ActiveXObject("Msxml2.XMLHTTP");    } catch(e) {}
    try { return new ActiveXObject("Microsoft.XMLHTTP"); } catch(e) {}
    try { return new XMLHttpRequest();                   } catch(e) {}
    return null;
}
function SetTitle (){
	this.xmlhttp = createXMLHttpRequest()
	var url = document.getElementById("url");
	if (url.value.length == 0 )
	{
		alert("��������һ��URL");
		url.focus();
		return;
	}
	document.getElementById("settitle").disabled=true;
	xmlhttp.open("GET",url.value,true);
	xmlhttp.onreadystatechange = callback;
	xmlhttp.send(null);
}
function callback () {
  if (xmlhttp.readyState == 4 ){
	 if (xmlhttp.status == 200){
		var title = xmlhttp.responseXML.documentElement.selectSingleNode("//title").text
		alert("��ȡ����ɹ�");
		document.getElementById("title").value = title;
		document.getElementById("settitle").disabled=false;
		}
	 else{
//		alert(xmlhttp.status);
//		alert(xmlhttp.readyState);
		alert("������һ����Ч��URL");
		document.getElementById("settitle").disabled=false;
	 }
  }
}
</script>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="#" onclick="purl('user_url.asp?action=add','��Ӷ���')">��Ӷ���</a></li>
					<li><a href="#" onclick="purl('user_url.asp','������')">������</a></li>
					<li><a href="#" onclick="purl('user_subject.asp?t=3','����ά��')">����ά��</a></li>
				</ul>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<fieldset id="FeedAdd" class="FieldsetForm">
						<legend><%If action="add" Then%>��Ӷ��ģ�<%Else%>�޸Ķ��ģ�<%End If%></legend>
						<form action="user_url.asp?action=save&id=<%=id%>" method="post" name="oblogform" onSubmit="return VerifySubmit()">
							<ul>
								<li><label>��&nbsp;�⣺
								<input name="title" id = "title" type=text  size="20" maxlength="100" value="<%=sTitle%>" ></label><!-- &nbsp;<input type="button" id ="settitle" value="��RSS��ַ��ȡ" onclick="SetTitle();"/> --></li>
								<li <%=OutRssDisplay%>><label>Feed��
								<input name="url" id ="url" type=text size="60" maxlength="250" value="<%=sUrl%>" /></label></li>
								<li><label>��&nbsp;�ࣺ
								<%
								If rsSubject.Eof Then
									%>
									��Ŀǰ��û���趨���ķ��࣬�����Լ�����ӻ���<a href="user_subject.asp?t=3">�趨��������</a>
								<%
								Else
								%>
									<select name="subjectid">
									<option value="0">δ����</option>
									<%
									Do While Not rsSubject.Eof
										allsub=allsub&rsSubject(0)&"!!??(("&rsSubject(1)&"##))=="
										%>
										<option value="<%=rsSubject("subjectid")%>" <%If rsSubject("subjectid")=sSubjectId Then Response.Write "selected" End If%>><%=rsSubject("subjectname")%></option>
										<%
										rsSubject.MoveNext
									Loop
									%>
									</select>
								<%
								End If
								%>
								</label>
								</li>
								<li <%=OutRssDisplay%>><label>��&nbsp;�룺
									<select name="encodeing">
										<option value='auto'>�Զ����</option>
										<option value='utf-8' <%if encodeing="utf-8" then response.Write("selected")%> >utf-8</option>
										<option value='gb2312' <%if encodeing="gb2312" then response.Write("selected")%>>gb2312</option>
									</select>
								</label></li>


								<li><input type="hidden" value="<%=mainuserid%>" name="mainuserid" /><input type="submit" name="addsubmit" id="Submit" value="<%If action="add" Then%>��Ӷ���<%Else%>�޸Ķ���<%End If%>"  /></li>
							</ul>
						</form>
					</fieldset>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/42.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
<%
End Sub

Sub Save()
    If oblog.ChkPost() = False Then
        oblog.AddErrStr ("ϵͳ��������ⲿ�ύ��")
        oblog.showUserErr
        Exit Sub
    End If
    'Get
	dim encodeing
    sUrl=Request.Form("url")
	sTitle=Request.Form("title")
	sSubjectId=Request.Form("subjectid")
	sTags=Request.Form("tags")
	sMemo=Request.Form("Memo")
	IsPrivate=Request.Form("isPrivate")
	encodeing=Request.Form("encodeing")
	if request("mainuserid")<>"" then
		mainuserid=clng(request("mainuserid"))
	else
		mainuserid=0
	end if
	If IsPrivate<>"1" Then IsPrivate="0"
    'Check
    If Id="" Then
	    If sUrl = "" Or oblog.strLength(sUrl) > 200 Then oblog.AddErrStr ("���ĳ��Ȳ���Ϊ���Ҳ��ܴ���200���ַ�����")
	    If sTitle = "" Or oblog.strLength(sTitle) > 50 Then oblog.AddErrStr ("���ⲻ��Ϊ���Ҳ��ܴ���50���ַ�����")
    	If oblog.chk_badword(sTitle) > 0 Then oblog.AddErrStr ("�����к���ϵͳ���������Ĺؼ��֣�")
	End If
    'If oblog.chk_badword(sTags) > 0 Then oblog.AddErrStr ("��ǩ�к���ϵͳ���������Ĺؼ��֣�")
    'If oblog.chk_badword(sMemo) > 0 Then oblog.AddErrStr ("��ע��Ϣ�к���ϵͳ���������Ĺؼ��֣�")
	'if left(surl,7)<>"http://" and mainuserid=0 then oblog.AddErrStr ("���ĵ�ַ������""http://""��ͷ��")
    If oblog.ErrStr <> "" Then oblog.showUserErr
	if encodeing="auto" and mainuserid=0 then
		encodeing=test_encodeing(surl)
	end if
	if mainuserid>0 then encodeing="gb2312"
    If Trim(Id)<>"" Then
    	rs.Open "Select * From oblog_myurl Where Id=" &  CLng(Id) & " And userid=" & oblog.l_uid,conn,1,3
    	If rs.Eof Then
    		rs.Close
    		Set rs=Nothing
    		oblog.AddErrStr ("Ŀ�����ݲ����ڣ��뷵�����²�����")
        	oblog.showUserErr
    	End If
  	Else
 		'urlid=CheckMyUrl(sUrl,sTitle)
 		'If urlid="" Then Exit Sub
      	rs.Open "Select * From oblog_myurl Where userid="&oblog.l_uid&" and url='"&oblog.filt_badstr(surl)&"'",conn,1,3
		if not rs.eof then
			rs.close
			set rs=nothing
			oblog.AddErrStr ("���Ѿ����Ĺ��˲��͵ĸ��£�")
			oblog.showUserErr
			exit sub
		else
			rs.AddNew
			'rs("urlid") =  0
		end if
   	End If
    '��ʼд�����
    rs("classid") = 0
    If sSubjectId<>"" Then rs("subjectid") = sSubjectId else rs("subjectid")=0
    If sTags<>"" Then rs("tags") = sTags
	rs("url")=sUrl
    rs("userid")=oblog.l_uid
    rs("isprivate")=IsPrivate
    If sMemo<>"" Then rs("memo") = sMemo
    rs("addtime") = oblog.ServerDate(Now)
	rs("encodeing")=encodeing
	rs("title")=sTitle
	rs("mainuserid")=mainuserid
	if id="" and mainuserid>0 then rs("isupdate")=1
    rs.Update
    rs.Close
	if mainuserid>0 then oblog.execute("update oblog_user set sub_num=sub_num+1 where userid="&mainuserid)
	Response.Write "<script>if(top.location.href!=self.location.href)parent.getfeedlist();</script>"
	response.Flush()
	if id="" then
		oblog.ShowMsg "��ӳɹ�","user_url.asp"
	else
		oblog.ShowMsg "�޸ĳɹ�","user_url.asp"
	end if
End Sub

Sub List()
	Dim Sql,i,lPage,lAll,lPages,iPage,Subjectid,keyword,cmd,sGuide
	Subjectid=Request("Subjectid")
	keyword=Request("keyword")
	If Keyword <> "" Then Keyword = oblog.filt_badstr(Keyword)
	cmd=LCase(Request("cmd"))
	Select Case cmd
		Case "11"
			If keyword<>"" Then
				Sql="Select top 500 * From oblog_myurl Where userid=" & oblog.l_uid & " and Title like '%" & keyword&"%' Order By id Desc"
'			Else
'				If Subjectid<>"" Then
'					Subjectid=Int(Subjectid)
'					Sql="Select top 500 a.id,a.subjectid,b.* From oblog_myurl a,oblog_url b Where a.userid=" & oblog.l_uid & " And a.subjectid=" & subjectid &" And a.urlid=b.urlid Order By a.subjectid,a.addtime Desc"
'				End If
			End If
		Case Else
			Sql="Select top 500 * From oblog_myurl a Where a.userid=" & oblog.l_uid & " Order By id Desc"
	End Select
	rs.Open Sql,conn,1,3
	lAll=INT(rs.recordcount)
    If lAll=0 Then
    	rs.Close
    	Set rs=Nothing
    	%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="#" onclick="purl('user_url.asp?action=add','��������')"><%=OutRssDisplay%>��������</a></li>
				</ul>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<!-- û����ؼ�¼ -->
					<div class="msg"><%=sGuide & " û����ؼ�¼" %></div>
					<!-- û����ؼ�¼ end -->
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/42.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
    	<%
    	Exit Sub
    End If
    iPage=20
	'��ҳ
	If Request("page") = "" Or Request("page") ="0" then
		lPage = 1
	Else
		lPage = Int(Request("page"))
	End If

	'���û����С = ÿҳ����ʾ�ļ�¼��Ŀ
	rs.CacheSize = iPage
	rs.PageSize = iPage
	rs.movefirst
	lPages = rs.PageCount
	If lPage>lPages Then lPage=lPages
	rs.AbsolutePage = lPage
	i=0
	%>
<table id="TableBody" cellpadding="0">
	<thead>
		<tr class="thead_tr1">
			<th>
				<ul id="UserMenu">
					<li><a href="#" onclick="chk_idAll(myform,1)">ȫ��ѡ��</a></li>
					<li><a href="#" onclick="chk_idAll(myform,0)">ȫ��ȡ��</a></li>
					<li><a href="#" onclick="if (chk_idBatch(myform,'ɾ��ѡ�еĶ�����?')==true) { document.myform.submit();}">ɾ������</a></li>
					<li <%=OutRssDisplay%>><a href="user_url.asp?action=add">��Ӷ���</a></li>
					<li><a href="user_subject.asp?t=3">����ά��</a></li>
					<li><a href="user_logzip.asp?action=saversslist" target="_blank">��������</a></li>
					<li id="showpage">
						<%=MakeMiniPageBar(lAll,iPage,lPage,G_P_FileName)%>
					</li>
				</ul>
			</th>
		</tr>
		<tr class="thead_tr2">
			<th>
				<table id="FeedTop" class="ListTop" cellpadding="0">
					<tr>
						<td class="t1"></td>
						<td class="t2">����</td>
						<td class="t3">����</td>
						<td class="scroll"></td>
					</tr>
				</table>
			</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>
				<div id="chk_idAll">
					<form name="myform" method="Post" action="user_url.asp?action=del" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
					<table id="Feed" class="TableList" cellpadding="0">
						<%
						'Do while not rs.EOF
						Do While Not rs.Eof And i < rs.PageSize
						i = i + 1
						%>
						<tr id="u<%=rs("id")%>" onclick="chk_iddiv('<%=rs("id")%>')">
							<td class="t1" title="���ѡ��">
								<input name='id' type='checkbox'  id="c<%=rs("id")%>" value="<%=cstr(rs("id"))%>"  onclick="chk_iddiv('<%=rs("id")%>')" />
							</td>
							<td class="t2">
								<span class="green" title="<%=getsubname(rs("subjectid"),allsub)%>"><%=getsubname(rs("subjectid"),allsub)%></span>
								<a href="user_url.asp?action=read&feedurl=<%=rs("url")%>&encodeing=<%=rs("encodeing")%>&title=<%=rs("title")%>&mainuserid=<%=rs("mainuserid")%>" title="<%=rs("title")%>"><%=rs("title")%></a><br />
								<span class="message_user">

								</span>
								<!--ʱ��-->
								<div class="time">Feed:<%=rs("url")%></div>
							</td>
							<td class="t3">
								<a href="user_url.asp?action=edit&id=<%=rs("id")%>&mainuserid=<%=rs("mainuserid")%>"><span class="green">�޸�</span></a>
								<a href="user_url.asp?action=del&id=<%=rs("id")%>" onClick="return confirm('ȷ��Ҫɾ���˶�����Ϣ��');"><span class="red">ɾ��</span></a>
							</td>
						</tr>
						<%
							If i >= iPage Then Exit Do
							rs.Movenext
						Loop
						%>
					</table>
					</form>
				</div>
				<script language="JavaScript" src="oBlogStyle/UserAdmin/60.js" type="text/javascript"></script>
			</td>
		</tr>
	</tbody>
<%
End Sub


'ϵͳ�б����Url��󶼲���Ҫ��/
'������ڣ��򷵻�UrlId
'��������ڣ���д�����������UrlId
Function CheckMyUrl(byval sUrl,byval sTitle)
	Dim rst,urlId
	If sUrl="" Then Exit Function
	sUrl = ProtectSQL(sUrl)
	If oblog.chk_badword(sUrl) Then Exit Function
	If Right(sUrl,1)="/" Then sUrl=Left(sUrl,Len(sUrl)-1)
	sUrl=Lcase(Trim(sUrl))
	Set rst=Server.CreateObject("Adodb.RecordSet")
	rst.Open "Select * From Oblog_Url Where url='" & sUrl & "'",conn,1,3
	If rst.Eof Then
		rst.AddNew
		rst("url")=sUrl
		rst("title")=sTitle
		rst("iCount")=1
		rst("vCount")=0
		rst("lasttime")=oblog.ServerDate(Now)
		rst.Update
		rst.Close
		Set rst=oblog.Execute("Select urlid From oblog_url Where url='" & sUrl & "'")
		urlId=rst("urlid")
	Else
		rst("iCount")=rst("iCount")+1
		rst("lasttime")=oblog.ServerDate(Now)
		urlId=rst("urlid")
		rst.Update
	End If
	rst.Close
	Set rst=Nothing
	CheckMyUrl=urlId
End Function

function test_encodeing(sUrl)
	On Error Resume Next
	dim http,re,encodeing
	Set http=Server.CreateObject("Microsoft.XMLHTTP")
	http.Open "GET",sUrl,False
	http.send
	if http.status="200" then
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="encoding=\""gb2312"
		if re.test(http.responseText) then
			encodeing="gb2312"
		else
			encodeing="utf-8"
		end if
		set re=nothing
	end if
	If Err Then
		Err.Clear
		test_encodeing="utf-8"
	else
		test_encodeing=encodeing
    End If
	set http=nothing
end function
%>