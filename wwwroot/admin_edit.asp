<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/class_blog.asp"-->
<!-- <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"> -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>վ������</title>
<%
'�����������,�ضϳ���ִ��,��ʡ��Դ. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
If oblog.CheckAdmin(0) = False Then
	Response.Write("��Ȩ��")
	Response.end
end if
dim action,skintype,skinorder,t,ActionText,ActionField,actionid,logid
t=0
Action=Trim(Request("Action"))
skintype=Trim(Request("skintype"))
skinorder=Trim(Request("skinorder"))
actionid=Trim(Request("do"))
logid=Trim(Request("logid"))
%>
<%If Session("adminname")<>"" And logid="" Then%>
<link rel="stylesheet" href="<%=SYSFOLDER_ADMIN%>/images/admin/style_edit.css" type="text/css" />
<%
Else
%>
<link rel="stylesheet" href="<%=SYSFOLDER_MANAGER%>/images/admin/style_edit.css" type="text/css" />
<%End If%>
<script src="<%=SYSFOLDER_ADMIN%>/images/menu.js" type="text/javascript"></script>
</head>
<span id ="TableBody" style="diplay:none"></span>
<span id ="chk_idAll" style="diplay:none"></span>
<%
select case actionid
	Case "1"
		ActionText="�޸���������(htm����)"
		ActionField="site_friends"
	Case "2"
		ActionText="�޸���վ����(htm����)"
		ActionField="site_placard"
	Case "3"
		ActionText="�޸�ע������(htm����)"
		ActionField="reg_text"
	Case "4"
		'ǰ̨����Ա���ɲ�����ѡ��
		ActionText="�޸��û������̨֪ͨ(htm����)"
		ActionField="user_placard"
	Case "5"
end select
select case Request("Actionsave")
	case "saveskin"
		call saveskin()
	case "savemodi"
		call savemodi()
	Case "savemodilog"
		Call savemodilog()
end select

select case Action
	case "modiskin"
		call modiskin()
	Case "modilog"
		Call modilog()
	Case Else
		call modi()
end select

sub savemodi()
	dim rs,strNote
	strNote=Request("edit")
	if not IsObject(conn) then link_database
	set rs=Server.CreateObject("adodb.recordset")
	rs.open "select "& ActionField &" from oblog_setup",conn,1,3
	rs(0)=strNote
	rs.update
	rs.close
	set rs=nothing
	oblog.reloadsetup
	WriteSysLog "�������޸��û���̨֪ͨ����վ���桢�������ӣ�ע������Ĳ���",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "����ɹ�",""
end sub

sub saveskin()
	dim rs,sql,table
	if Trim(Request("skinname"))="" then Response.Write("ģ��������Ϊ��"):Response.End()
	if Trim(Request("edit"))="" then Response.Write("ģ�����ݲ���Ϊ��"):Response.End()
	if skintype="user" then
		table="oblog_userskin"
	elseif skintype="sys" then
		table="oblog_sysskin"
	ElseIf skintype="team" Then
		table = "oblog_teamskin"
	else
		Response.Write("��������")
		Response.end
	end if
	set rs=Server.CreateObject("adodb.recordset")
	sql="select * from "&table&" where id="&CLng(Request.QueryString("id"))
	if not IsObject(conn) then link_database
	rs.open sql,conn,1,3
	if skintype="sys" then
		rs("sysskinname")=Trim(Request("skinname"))
	else
		rs("userskinname")=Trim(Request("skinname"))
		rs("skinpic")=Trim(Request("skinpic"))
		rs("skinauthorurl")=Trim(Request("skinauthorurl"))
	end if
	rs("skinauthor")=Trim(Request("skinauthor"))
	if skinorder="0" then
		rs("skinmain")=Request("edit")
	else
		rs("skinshowlog")=Request("edit")
	end if
	rs.update
	rs.close
	set rs=nothing
	oblog.reloadsetup
	WriteSysLog "������ģ��༭������Ŀ��ģ��ID��"&Request.QueryString("id")&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "����ɹ�",""
end Sub
sub modiskin()
	dim rs,table
	if skintype="user" then
		table="oblog_userskin"
	elseif skintype="sys" then
		table="oblog_sysskin"
	ElseIf skintype="team" Then
		table = "oblog_teamskin"
	else
		Response.Write("��������")
		Response.end
	end if
	set rs=oblog.execute("select * from "&table&" where id="&CLng(Request.QueryString("id")))
	if rs.eof then
		Response.write("�޼�¼")
		Response.End()
	end If
	If C_Editor_Type = 2 Then
%>
<body>
<%
Else
	Response.Write "<body>"
End if%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">��վ����</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="POST" action="" id="oblogform" name="oblogform"   <%If C_Editor_Type=2 Then%>onsubmit="submits();"<%End If%>>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
    <tr>
      <td height="25">ģ�����ƣ�
        <input name="skinname" type="text" id="skinname" value="<%if skintype="sys" then Response.Write rs("sysskinname") else Response.Write rs("userskinname")%>">
        �������ߣ�
        <input name="skinauthor" type="text" id="skinauthor" value="<%=rs("skinauthor")%>">
<%if skintype="user" Or skintype = "team" then%>
<br>
        ģ���������ӣ�
        <input name="skinauthorurl" type="text" id="skinauthorurl" size="50" value="<%=rs("skinauthorurl")%>">
         <br>
        ģ��Ԥ��ͼƬ��
        <input name="skinpic" type="text" id="skinpic" size="50" value="<%=rs("skinpic")%>">
<%end if%>
		</td>
    </tr>
    <tr>
      <td>
		<%
      	Dim sValue
      	if skinorder="0" then
			if rs("skinmain")<>"" then sValue = Server.HtmlEncode(filtskinpath(rs("skinmain")))
		else
			if rs("skinshowlog")<>"" then sValue = Server.HtmlEncode(filtskinpath(rs("skinshowlog")))
		end if
      	%>
		<div id="textarea">
		<span id="loadedit" style="font-size:12px;display:<%=C_Editor_LoadIcon%>;"><img src='images/loading.gif' align='absbottom'> ��������༭��...</span>
		<textarea id="edit" name="edit" style="width:100%;height:320px;display:none"><%=sValue%></textarea>
		<%If C_Editor_Type=2 Then Server.Execute C_Editor & "/edit.asp"%>
		</div>
		<% sValue=""%>
</td>
    </tr>
    <tr>
      <td> <div align="center">
        <input name="Actionsave" type="hidden" id="Action" value="saveskin">
        <input name="cmdSave" type="submit" id="cmdSave" value=" �����޸� " >
      </div></td>
    </tr>
  </table>
</form>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<%If C_Editor_Type=1 Then oblog.MakeEditorText "edit",0,"930","405"%>
<%
set rs=nothing
end sub

sub modi()
	dim rs
	if ActionField="" then
		Response.write("����Ĳ���")
		Response.end
	end if
	set rs=oblog.execute("select "& ActionField &" from oblog_setup")
	If C_Editor_Type = 2 Then
%>
<body>
<%
Else
Response.Write "<body>"
End if%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left"><%=ActionText%></li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg"  >
	<form name="oblogform" method="post" action="" <%If C_Editor_Type=2 Then%>onsubmit="submits();"<%End If%>>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="border">
    <tr>
      <td>
		<span id="loadedit" style="font-size:12px;display:<%=C_Editor_LoadIcon%>;"><img src='images/loading.gif' align='absbottom'> ��������༭��...</span>
    	<textarea id="edit" name="edit" style="width:100%;height:320px; display:none"><%=Server.HtmlEncode(filtskinpath(OB_IIF(rs(0),"")))%></textarea >
			<%If C_Editor_Type=2 Then Server.Execute C_Editor & "/edit.asp"%>
		</td>
    </tr>
    <tr>
      <td>
				<div align="center">
				<br />
				<input name="Actionsave" type="hidden" id="Action" value="savemodi">
                <input type="submit" name="Submit" id="Submit" value="�ύ�޸�">
				</div>
		</td>
    </tr>
  </table>
      </form>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<%If C_Editor_Type=1 Then oblog.MakeEditorText "edit",0,"780","260"%>
<%
set rs=nothing
end Sub
sub modilog()
	dim rs
	Dim showword
	If logid<>"" Then
		logid = CLng (logid)
	Else
		Response.Write "��������"
		Response.End
	End if
	set rs=oblog.execute("select showword,logid,author,topic,classid,abstract,logtext FROM oblog_log WHERE logid="&logid)
	if rs.eof then
		Response.write("�޼�¼")
		Response.End()
	end If
'-------------------------------------------����Ȩ���ж�---------------------------
	Dim Z_ClassID,Z_LogRoles
	If Not oblog.CheckAdmin(1) Then
		Z_ClassID=","&rs("classid")&","
		Z_LogRoles=session("r_classes1")
		If Z_LogRoles<>""  Then
			Z_LogRoles=","&Z_LogRoles&","
			If  Not  InStr(Z_LogRoles,Z_ClassID) > 0 Then
				Response.write("��û�д˷���Ĺ���Ȩ��!")
				Response.End()
			End If
		End If
	End If

'-----------------------------------------------------------
	showword = rs(0)
%>
<style>
html {overflow-x:hidden;}
body {background:#fff;}
</style><%
If C_Editor_Type = 2 Then
%>
<body>
<%
Else
	Response.Write "<body>"
End if%>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�޸���־</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="POST" action="" id="oblogform" name="oblogform"   <%If C_Editor_Type=2 Then%>onsubmit="submits();"<%End If%>>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
	<tr>
		<td height="25" class="tdbg">
			<strong>�༭IDΪ��<%=rs(1)%>������Ϊ��<%=rs(2)%>����־</strong><br><br>
			<strong>���⣺</strong><input name="topic" type="text" id="topic" value="<%=rs(3)%>" size="40">
			<select name="classid" id="classid">
			<%=oblog.show_class("log",rs(4),0)%>
			</select>
			<strong>������ʾ������</strong><input name="showword" type="text" id="showword" value="<%if showword<>"" then Response.Write(showword) else Response.Write(500)%>" size="10"><br>
			<br><strong>ժҪ��������û���ҳ������ң�����дժҪ��</strong><br>
			<textarea name="abstract" type="text" id="abstract" rows="5" cols="82"><%=rs(5)%></textarea><br>
			<strong>�������ݣ�</strong><br>
			<div id="textarea">
			<span id="loadedit" style="font-size:12px;display:<%=C_Editor_LoadIcon%>;"><img src='images/loading.gif' align='absbottom'> ��������༭��...</span>
			<textarea id="edit" name="edit" style="width:100%;height:320px; display:none"><%= Server.HtmlEncode(rs(6))%></textarea >
			<%If C_Editor_Type=2 Then Server.Execute C_Editor & "/edit.asp"%>
			</div>

		</td>
    </tr>
    <tr>
      <td class="tdbg">
		<input name="logid" type="hidden" id="logid" value="<%=logid%>">
		<input name="Actionsave" type="hidden" id="Action" value="savemodilog">
		<input name="cmdSave" type="submit" id="cmdSave" style="height:30px;" value=" �����޸� " >
	  </td>
    </tr>
  </table>
</form>
		</div>
	</div>
	<ul class="main_end">
		<li class="main_end_left left"></li>
		<li class="main_end_right right"></li>
	</ul>
</div>
<%If C_Editor_Type=1 Then oblog.MakeEditorText "edit",0,"535","260"%>
<%
set rs=nothing
end Sub

sub savemodilog()
	dim rs,sql,blog
	Dim topic,showword,abstract,logtext,log_classid,authorid,log_blogteam
	logid = CLng (Request("logid"))
	topic = Trim(Request("topic"))
	showword = Trim(Request("showword"))
	abstract = Trim(Request("abstract"))
	logtext = Trim(Request("edit"))
	log_classid = Trim(Request("classid"))
	if topic = "" then Response.Write("���ⲻ��Ϊ��"):Response.End()
	if logtext = "" then Response.Write("��־���ݲ���Ϊ��"):Response.End()
	if logid = 0 Then
		Response.Write("��������")
		Response.end
	end if
	set rs=Server.CreateObject("adodb.recordset")
	sql="select * from oblog_log where logid="&logid
	if not IsObject(conn) then link_database
	rs.open sql,conn,1,3
	rs("topic") = topic
	rs("showword") = showword
	rs("abstract") = abstract
	rs("logtext") = logtext
	rs("classid") = log_classid
	log_blogteam = rs("userid")
	authorid = rs("authorid")
	rs.update
	rs.close
	set rs=Nothing
	Set blog = new class_blog
	blog.userid = authorid
	blog.CreateFunctionPage
	blog.Update_log logid, 0
	set rs=oblog.execute("select top 1 logid from oblog_log where logid<"&logid&" and userid="&log_blogteam&" and logtype=0 order by logid desc")
	If Not rs.EOF Then blog.Update_log rs(0), 0
	blog.Update_calendar (logid)
	blog.Update_newblog (authorid)
	blog.Update_Subject (authorid)
	blog.Update_index 0
	blog.Update_info authorid
	If log_blogteam<>authorid Then
		blog.userid=log_blogteam
		blog.CreateFunctionPage
		blog.update_calendar(logid)
		blog.update_newblog(log_blogteam)
		blog.update_subject(log_blogteam)
		blog.update_index 0
		blog.update_info log_blogteam
	End If
	Set blog=Nothing
	WriteSysLog "��������־�޸Ĳ�����Ŀ����־ID��"&logid&"",oblog.NowUrl&"?"&Request.QueryString
	oblog.ShowMsg "�޸ĳɹ�",""
end Sub
Sub WriteSysLog(ByVal sContents,ByVal Strings)
	Dim sIP,rs
	sIP=oblog.userIp
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.Open "select * From oblog_syslog Where 1=0",conn,1,3
	rs.AddNew
	rs("username")=OB_IIF(session("m_name"),session("adminname"))
	rs("addtime")=oblog.ServerDate(Now)
	rs("addip")=sIP
	rs("desc")=OB_IIF(session("m_name"),session("adminname")) & " �� " & oblog.ServerDate(Now()) & " �� " & sIP  & " " & sContents
	rs("QueryStrings") = Strings
	rs("itype") = 3		'���ݹ���Ա������¼
	rs.Update
	rs.Close
	Set rs=Nothing
End Sub
%>
</body>
</html>