<%@ LANGUAGE = VBScript CodePage = 936%>
<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<!--#include file="inc/md5.asp"-->
<script language="javascript" src="inc/main.js"></script>
<script language="javascript" src="inc/function.js"></script>
<%
dim oblog
set oblog=new class_sys
oblog.start
if not oblog.checkuserlogined() then
	Response.Redirect("login.asp?fromurl="&Replace(oblog.GetUrl,"&","$"))
end If

Dim groupName ,trs ,tsql
Set trs =oblog.execute ("select g_name FROM oblog_groups WHERE groupid = " & oblog.l_uGroupId)
groupName = trs (0)
trs.close
Set trs=Nothing
tsql = "or groups like '"&oblog.l_uGroupId&",%' or groups like '%,"&oblog.l_uGroupId&"' or groups like '%,"&oblog.l_uGroupId&",%' or groups ='"&oblog.l_uGroupId&"'"
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=oblog.l_uname%>-短消息管理</title>
<link href="OblogStyle/style.css" rel="stylesheet" type="text/css" />
</head>
<body style="overflow:hidden;background:#fff" scroll="no">
<div class="content">
    <div class="content_body">
<%
dim action
action=Request("action")
select case action
	case "send","batchsend"
		call send
	case "save"
		call save
	case "read0"
		call read(0)
	case "read1"
		call read(1)
	case "readteam"
		call read(2)
end select
%>
	</div>
  </div>
</body>
</html>
<%
sub send()
	dim rs,inceptIds,incepters
	inceptIds=Request("incept")
	If action="batchsend" Then
		'将ID转换为用户名
		inceptIds=FilterIds(inceptIds)
		If inceptIds<>"" Then
			Set rs=oblog.Execute("Select username From oblog_user Where userid In (" & inceptIds & ")")
			Do While Not rs.Eof
				incepters=incepters & "," & rs(0)
				rs.MoveNext
			Loop
			rs.Close
			If incepters<>"" Then incepters=Right(incepters,Len(incepters)-1)
		End If
	Else
		incepters=Request("incept")
	End If
%>
<SCRIPT language="javascript">
function changincept()
{
	var svalue=del_space(document.oblogform.incept.value);
	if (svalue!="")
	{
		if (svalue.indexOf(document.oblogform.selectincept.value+',')>=0||svalue.indexOf(','+document.oblogform.selectincept.value)>=0){alert('请勿重复选择');return false;}
		document.oblogform.incept.value = svalue+","+document.oblogform.selectincept.value;
	}

	else
	{
		document.oblogform.incept.value = document.oblogform.selectincept.value;
	}

}
</SCRIPT>
<form action="user_pm.asp?action=save" method="post" name="oblogform">
	<table class='win_pm_table' align='center' border='0' cellpadding='0' cellspacing='1'>
		<tr>
			<td colspan='2' class='win_pm_table_top'>发送站内短消息</td>
		</tr>
		<tr>
			<td colspan='2'></td>
		</tr>
		<tr>
			<td class='win_pm_table_td'>收件人：</td>
			<td>
<input type="text" name="incept" id="incept" width="20" value="<%=incepters%>" />
	<select size="1" id="selectincept" onChange='javascript:changincept()'>
	<option value="">选择好友</option>
<%
set rs=oblog.execute("select username from oblog_user,oblog_friend where oblog_user.userid=oblog_friend.friendid and oblog_friend.isblack=0 and oblog_friend.userid="&oblog.l_uid)
while not rs.eof
	Response.Write("<option value='"&oblog.filt_html(rs(0))&"'>"&oblog.filt_html(rs(0))&"</option>")
	rs.movenext
wend
set rs=nothing
%>
	</select></td>
		</tr>
		<tr>
			<td class='win_pm_table_td'>标题：</td>
			<td><input type="text" name="topic" size="35" maxlength="50" id = "topic" value="<%=oblog.filt_html(Trim(Request("topic")))%>" /></td>
		</tr>
		<tr>
			<td class='win_pm_table_td'>内容(限制为240字)：</td>
			<td><textarea name="content" id ="content" cols="35" rows="8"  maxlength="240"></textarea></td>
		</tr>
		<tr>
			<td colspan='2' align="center"><INPUT type="hidden" name="id" value=""><input type="button"  value=" 发送 " onclick="sendpm();"> <input type="button" onClick="window.close();" value=" 关闭 "></td>
		</tr>
	</table>
</form>
<%
end sub

sub save()
	dim incept,content,sql,rs,rs1,inceptid,topic,Err1,inceptname,restr
	Dim ajax
	set ajax=new AjaxXml
	incept=oblog.filt_badstr(Trim(Request("incept")))
	content=Trim(Request("content"))
	topic=Trim(Request("topic"))
	if incept="" then oblog.adderrstr("收件人不能为空")
	if content="" then oblog.adderrstr("短消息内容不能为空")
	if topic="" then oblog.adderrstr("短消息标题不能为空")
	if oblog.errstr<>"" then
'		oblog.ShowMsg Replace(oblog.errstr,"_","\n"),""
		restr = Split(Replace(oblog.errstr,"_","\n")&"$$$1","$$$")
		ajax.re(restr)
		Response.End
	end if
	'incept
	'需要处理掉重复的数据
'	incept=Join(FilterArr(Split(incept,",")),"','")
	'Response.Write incept
	'Response.End
	sql="select userid,username from [oblog_user] where username in ('"&incept&"')"
	set rs=oblog.execute(sql)
	if rs.eof then
'		oblog.ShowMsg "无此用户,请检查用户名！",""
		restr = Split("无此用户,请检查用户名！$$$1","$$$")
		ajax.re(restr)
		Response.End
	end if
	Do While Not rs.Eof
		inceptid=CLng(rs(0))
		inceptname=rs(1)
		set rs1=oblog.execute("select id from oblog_friend where isblack=1 and userid="&inceptid&" and friendid="&oblog.l_uid)
		If Not rs1.eof then
			Err1= "你在收件人的黑名单中，无法发送短消息！"
		Else
			rs1.Close
			sql="select top 1 * from oblog_pm Where 1=0"
			set rs1=Server.CreateObject("adodb.recordset")
			rs1.open sql,conn,1,3
			rs1.addnew
			rs1("incept")=oblog.Interceptstr(inceptname,50)
			rs1("topic")=oblog.Interceptstr(oblog.filt_badword(topic),100)
		'	rs1("content")=oblog.Interceptstr(oblog.filt_badword(content),250)
			rs1("content")=Left(oblog.filt_badword(content),250)
			rs1("sender")=oblog.l_uname
			rs1.update
		End If
		rs1.close
		rs.MoveNext
	Loop
	restr ="发送成功$$$2"
	set rs=Nothing
	ajax.re(Split(restr,"$$$"))
	Response.End
'	Oblog.ShowMsg "短消息发送成功!","close"
end sub

sub read(sAction)
	dim id,rs
	id=CLng(Trim(Request("id")))
	if sAction="0" Then
		set rs=oblog.execute("select * from oblog_pm where (id="&id & " And Incept='"&oblog.l_uname&"' And delR=0) or ((groups like '"&oblog.l_uGroupId&",%' or groups like '%,"&oblog.l_uGroupId&"' or groups like '%,"&oblog.l_uGroupId&",%' or groups ='"&oblog.l_uGroupId&"') and id="&id & ")")
	Elseif sAction="1" Then
		set rs=oblog.execute("select * from oblog_pm where id="&id & " And sender='"&oblog.l_uname&"' And delS=0")
	Elseif sAction="2" Then
		'需判断查看者是否为组之管理员
		set rs=oblog.execute("select teamid,addtime,'申请加入群组说明' as topic,info as content,'0' as issys,'-' as sender,'" &oblog.l_uname &"' as incept from oblog_teamusers where id="&id )
		Dim rst
		If Not rs.Eof Then
			Set rst=oblog.execute("select teamid From oblog_team Where createrid=" & oblog.l_uid & " And teamid=" & rs("teamid"))
			If rst.Eof Then
				Set rst=Nothing
				Set rs=Nothing
				Response.Write("<ul><li>您无权察看该信息，<a href='javascript:window.close();'>关闭窗口</a></li></ul>")
				exit sub
			End If
		End If
	Else
		Response.Write("<ul><li>错误的参数，<a href='javascript:window.close();'>关闭窗口</a></li></ul>")
		exit sub
	End If

	if rs.eof then
		set rs=nothing
		Response.Write("<ul><li>无此纪录，<a href='javascript:window.close();'>关闭窗口</a></li></ul>")
		exit sub
	end if
%>
	<table class='win_pm_table' align='center' border='0' cellpadding='0' cellspacing='1'>
		<tr>
			<td colspan='2' class='win_pm_table_top'>查看短消息</td>
		</tr>
		<tr>
			<td class='win_pm_table_td'>发送时间：</td>
			<td><%=rs("addtime")%></td>
		</tr>
		<tr>
			<td class='win_pm_table_td'>短消息标题：</td>
			<td><%=oblog.filt_html(rs("topic"))%></td>
		</tr>
		<tr>
			<td class='win_pm_table_td'>短消息内容：</td>
			<td><TEXTAREA NAME="con" ROWS="6" COLS="35"><%=oblog.filt_html(rs("content"))%></TEXTAREA></td>
		</tr>
<%If sAction<>"2" Then%>
		<tr>
			<td class='win_pm_table_td'>发件人：</td>
			<td><%
Dim UserUrl
'If oblog.CacheConfig(5)=1 Then
'	If Left(oblog.l_udomain,8)="http://." Or Trim(oblog.l_udomain)="." Then
'		UserUrl="<a href="""&oblog.l_udir&"/"&oblog.l_ufolder&"/index."&f_ext&""" target=""_blank"">"&oblog.filt_html(rs("sender"))&"</a>"
'	Else
'		UserUrl="<a href=""http://"&oblog.l_udomain&""" target=""_blank"">"&oblog.l_udomain&"</a>"
'	End If
'Else
'	UserUrl="<a href="""&oblog.l_udir&"/"&oblog.l_ufolder&"/index."&f_ext&""" target=""_blank"">"&oblog.filt_html(rs("sender"))&"</a>"
'End If
'If true_domain=1 and oblog.l_ucustomdomain<>"" then
'	UserUrl="<a href=""http://"&oblog.l_ucustomdomain&""" target=""_blank"">"&oblog.filt_html(rs("sender"))&"</a>"
'End If
UserUrl = "<a href=go.asp?user="&rs("sender")&" target=_blank>"&oblog.filt_html(rs("sender"))&"</a>"
	If rs("issys")= 1 Then
		Response.Write "<font color=red style=font-weight:600>" &rs("sender") &"</font>"
	Else
		Response.Write  UserUrl & "&nbsp;&nbsp;<input type=""button"" onclick=""openScript('user_friends.asp?action=add&close=true&friendname="&rs("sender")&"',450,400)"" onmouseup=""window.close();"" value="" 点击添加对方为好友 "" >"
	End If
	%></td>
		</tr>
		<tr>
			<td class='win_pm_table_td'>收件人：</td>
			<td><%
	If rs("incept")="0" Then
		Response.Write "<font color=green style=font-weight:600>" &groupName& "</font>"
	Else
		Response.Write oblog.filt_html(rs("incept"))
	End if%></td>
		</tr>
		<%End If%>
		<tr>
			<td colspan='2' align="center">
		<%If sAction="0" And rs("sender")<>"系统管理员" Then%>
			<input type="button" onclick="openScript('user_pm.asp?action=send&incept=<%=rs("sender")%>&topic=<%="回复:"&oblog.filt_html(rs("topic"))%>',450,400)" onmouseup="self.close();" value=" 回复 " <%If rs("issys")= 1 Then%>disabled<%End if%>>
		<%End If%>
		<input type="button" onClick="window.close();" value=" 关闭 ">
			</td>
		</tr>
	</table>
<%
	if oblog.l_uname=rs("incept") And rs("sender")<>"系统管理员" then
		oblog.execute("update oblog_pm set isreaded=1 where id="&id&" and incept='"&oblog.l_uname&"'")
	ElseIf rs("sender")="系统管理员" Then
	oblog.execute("update oblog_pm set dels=1,delr=1,isreaded=1 where id="&id&" and incept='"&oblog.l_uname&"'")
	end if
	set rs=nothing
end sub
%>
<script>
function sendpm(logid){
	var content = document.getElementById("content").value;
	var topic = document.getElementById("topic").value;
	var incept = document.getElementById("incept").value;
	if (del_space(incept)=='')	{
		alert('请输入接受站内短信的用户名！');
		return false;
	}
	if (del_space(topic)=='')	{
		alert('请输入标题！');
		return false;
	}
	if (del_space(content)=='')	{
		alert('请输入内容！');
		return false;
	}
	var Ajax = new oAjax("user_pm.asp?action=save",show_returnsave);
	var arrKey = new Array("incept","topic","content");
	var arrValue = new Array(incept,topic,content);
	Ajax.Post(arrKey,arrValue);
}
function show_returnsave(arrobj){
	if (arrobj){
			alert(arrobj[0]);
			self.close();
		}
	}
</script>