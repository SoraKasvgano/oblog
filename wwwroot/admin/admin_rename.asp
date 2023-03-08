<!--#include file="inc/inc_sys.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>站点配置</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<%
dim rs,newname,old,username,sql,action
action=oblog.filt_badstr(Request("action"))
newname=oblog.filt_badstr(Trim(Request("newname")))
old=oblog.filt_badstr(Trim(Request("old")))
	if old<>"" then
		set rs=oblog.execute("select userid from oblog_user where username='"&oblog.filt_badstr(old) & "'")
		if rs.eof then
			oblog.ShowMsg "原用户名不存在！",""
		end if
	end if
	if newname<>"" then
		set rs=oblog.execute("select userid from oblog_user where username='"&oblog.filt_badstr(newname) & "'")
		if not rs.eof then
			oblog.ShowMsg "新用户名已经存在！",""
		end if
	end if
	if Instr(newname,"=")>0 or Instr(newname,"%")>0 or Instr(newname,chr(32))>0 or Instr(newname,"?")>0 or Instr(newname,"&")>0 or Instr(newname,";")>0 or Instr(newname,",")>0 or Instr(newname,"'")>0 or Instr(newname,",")>0 or Instr(newname,chr(34))>0 or Instr(newname,chr(9))>0 or Instr(newname,"")>0 or Instr(newname,"$")>0 then
		oblog.ShowMsg "用户名含有非法字符！",""
	end if
if action="DoUpdate" Then
	If Request.QueryString<>"" Then Response.End
	If newname="" Then
		oblog.ShowMsg "新用户名不能为空！",""
	End If
	If old="" Then
		oblog.ShowMsg "原用户名不能为空！",""
	End If
	If old=newname Then
		oblog.ShowMsg "原用户名与新用户名相同，无需更改！",""
	End If
	oblog.execute("update [oblog_user] set username='"&newname&"' where username='"&old&"'")
	oblog.execute("update [oblog_log] set author='"&newname&"' where author='"&old&"'")
	oblog.execute("update [oblog_comment] set comment_user='"&newname&"' where isguest=0 and comment_user='"&old&"'")
	oblog.execute("update [oblog_Albumcomment] set comment_user='"&newname&"' where isguest=0 and comment_user='"&old&"'")
	oblog.execute("update [oblog_message] set message_user='"&newname&"' where isguest=0 and  message_user='"&old&"'")
	oblog.execute("update [oblog_pm] set sender='"&newname&"' where sender='"&old&"'")
	oblog.execute("update [oblog_pm] set incept='"&newname&"' where incept='"&old&"'")
	oblog.Execute ("UPDATE oblog_arguelist SET author='"&newname&"' WHERE author='"&old&"'")
	oblog.Execute ("UPDATE oblog_team SET managername='"&newname&"' WHERE managername='"&old&"'")
	oblog.Execute ("UPDATE oblog_team SET creatername='"&newname&"' WHERE creatername='"&old&"'")
	oblog.Execute ("UPDATE oblog_teampost SET author='"&newname&"' WHERE author='"&old&"'")
	EventLog "进行了用户名修改操作，新用户名："&newname&"，原用户名："&old&"",""
	Response.Write("<br>已经成功将用户名进行了更改！")
	else
	%>
<form name="form1" method="post" action="">
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">更 改 用 户 名</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr class="tdbg">
    <td colspan="2"><p>说明：
        本操作将更改用户名，请慎重操作。</p>
      <p>原用户名：
        <input name="old" type="text" id="old">
         <br>
         新用户名：
         <input name="newname" type="text" id="newname">
         用户名禁止特殊符号<br>
        </p></td>
  </tr>
  <tr class="tdbg">
    <td height="25" colspan="2"><input name="Submit" type="submit" id="Submit" value="改名">
    <input name="Action" type="hidden" id="Action" value="DoUpdate"></td>
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
<%
end if
%>