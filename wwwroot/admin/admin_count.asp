<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����ϵͳ����</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>

<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">�� �� ϵ ͳ �� ��</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">

<form name="form1" method="post" action="">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr class="tdbg">
    <td colspan="2">
<%
On Error Resume Next 
dim action,rs,sql,trs
action=Request("action")
if action="DoUpdate" And Request.QueryString = "" Then
    Dim blog,p
    Set blog = New class_blog
    blog.progress_init
    p = 6
	Set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select top 1 * from oblog_setup"
	rs.Open sql,Conn,1,3
	blog.progress Int(1 / p * 100), "����ͳ����־��"
	set trs=oblog.execute("select count(logID) from [oblog_log] WHERE isdel = 0 ")
	if isNull(trs(0)) then
		rs("log_count")=0
	else
		rs("log_count")=trs(0)
	end If
		blog.progress Int(2 / p * 100), "����ͳ��������"
	set trs=oblog.execute("select count(commentID) from oblog_comment WHERE isdel = 0")
	if isNull(trs(0)) then
		rs("comment_count")=0
	else
		rs("comment_count")=trs(0)
	end If
	blog.progress Int(3 / p * 100), "����ͳ��������"
	set trs=oblog.execute("select count(messageID) from oblog_message WHERE isdel = 0")
	if isNull(trs(0)) then
		rs("message_count")=0
	else
		rs("message_count")=trs(0)
	end If
	blog.progress Int(4 / p * 100), "����ͳ���û���"
	set trs=oblog.execute("select count(userID) from [oblog_user]")
	if isNull(trs(0)) then
		rs("user_count")=0
	else
		rs("user_count")=trs(0)
	end if
	rs.update
	rs.close
	set rs=nothing
	set trs=Nothing
	blog.progress Int(5 / p * 100), "����ϵͳ����"
	oblog.ReloadSetup()
	oblog.ReLoadCache()
	blog.progress Int(6 / p * 100), "����ϵͳ�������"
	Set blog = Nothing
	EventLog "�����˸���ϵͳ���ݲ���",""
	Response.Write	"<script src="""&blogdir&"index.asp?re=0""></script>"
	%>
	<br /><a href="javascript:history.back(-1)">������һҳ</a>
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

	<%else%>
	<p>˵����<br>
        <b>1�������������¼���ϵͳ����־�����ۣ������������¸�����ҳ��</b><br />
		<b>2������������������ϵͳ���档</b><br />
        <b>3�������������������ʱ�������ķ�������Դ������ϸȷ��ÿһ��������ִ�С�</b></p></td>
  </tr>
  <tr class="tdbg">
    <td height="25"><input name="Submit" type="submit" id="Submit" value="����ϵͳ����">
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