<!--#include file="inc/inc_sys.asp"-->
<%
dim Action,FoundErr,ErrMsg
dim rs,sql
Action=Trim(Request("Action"))
sql="select * from oblog_Admin where UserName='" & Admin_Name & "'"
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open sql,conn,1,3
if rs.Bof and rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>不存在此用户！</li>"
else
	if Action="Modify" then
		call ModifyPwd()
	else
		call main()
	end if
end if
rs.close
set rs=nothing
if FoundErr=True then
	call WriteErrMsg()
end if
Set oblog = Nothing
sub ModifyPwd()
	dim password,PwdConfirm,Pwd
	password=Trim(Request("Password"))
	PwdConfirm=Trim(Request("PwdConfirm"))
	Pwd=Trim(Request("Pwd"))
	If MD5(pwd) <> rs("password") Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>原密码输入错误！</li>"
		exit sub
	End if
	if password="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>新密码不能为空！</li>"
	end if
	if PwdConfirm<>Password then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>确认密码必须与新密码相同！</li>"
		exit sub
	end if
	if Password<>"" then
		rs("password")=md5(password)
	end if
   	rs.update
	EventLog "进行了密码修改操作",oblog.NowUrl&"?"&Request.QueryString
	Response.Write"<script language=JavaScript>"
	Response.Write"alert(""密码修改成功，请重新登录！"");"
	Response.Write"window.history.go(-1);"
	Response.Write"</script>"
end sub

sub main()
%>
<script language=javascript>
function check()
{
  if(document.form1.Pwd.value=="")
    {
      alert("原密码不能为空！");
	  document.form1.Pwd.focus();
      return false;
    }
  if(document.form1.Password.value=="")
    {
      alert("密码不能为空！");
	  document.form1.Password.focus();
      return false;
    }

  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("初始密码与确认密码不同！");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();
      return false;
    }
}
</script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>oBlog--修改管理员密码</title>
<link rel="stylesheet" href="images/style.css" type="text/css" />
<script src="images/menu.js" type="text/javascript"></script>
</head>
<body>
<div id="main_body">
	<ul class="main_top">
		<li class="main_top_left left">修改管理员密码</li>
		<li class="main_top_right right"> </li>
	</ul>
	<div class="main_content_rightbg">
		<div class="main_content_leftbg">
<form method="post" action="admin_adminmodifypwd.asp" name="form1" onsubmit="javascript:return check();">
  <table width="300" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border" >
    <tr class="title">
      <td height="25" colspan="2" class="topbg"><strong>修改管理员密码</strong></td>
    </tr>
    <tr>
      <td width="100" align="right" class="tdbg">用 户 名：</td>
      <td class="tdbg"><%=rs("UserName")%></td>
    </tr>
    <tr>
      <td width="100" align="right" class="tdbg">原 密 码：</td>
      <td class="tdbg"><input type="password" name="Pwd"> </td>
    </tr>
    <tr>
      <td width="100" align="right" class="tdbg">新 密 码：</td>
      <td class="tdbg"><input type="password" name="Password"> </td>
    </tr>
    <tr>
      <td width="100" align="right" class="tdbg">确认密码：</td>
      <td class="tdbg"><input type="password" name="PwdConfirm"> </td>
    </tr>
    <tr>
      <td height="40" colspan="2" class="tdbg"><input name="Action" type="hidden" id="Action" value="Modify">
        <input  type="submit" name="Submit" value=" 确 定 " style="cursor:hand;">
        &nbsp; <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="reset()" style="cursor:hand;"></td>
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
</body>
</html>
<%
end sub

sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbcrlf
	strErr=strErr & "<head>" & vbcrlf
	strErr=strErr & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />" & vbcrlf
	strErr=strErr & "<title>oBlog--后台管理</title>" & vbcrlf
	strErr=strErr & "<link rel=""stylesheet"" href=""images/style.css"" type=""text/css"" />" & vbcrlf
	strErr=strErr & "</head>" & vbcrlf
	strErr=strErr & "<body>" & vbcrlf
	strErr=strErr & "<div id=""main_body"">" & vbcrlf
	strErr=strErr & "	<ul class=""main_top"">" & vbcrlf
	strErr=strErr & "		<li class=""main_top_left left"">错误信息</li>" & vbcrlf
	strErr=strErr & "		<li class=""main_top_right right""> </li>" & vbcrlf
	strErr=strErr & "	</ul>" & vbcrlf
	strErr=strErr & "	<div class=""main_content_rightbg"">" & vbcrlf
	strErr=strErr & "		<div class=""main_content_leftbg"">" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>错误信息</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>产生错误的可能原因：</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "		</div>" & vbcrlf
	strErr=strErr & "	</div>" & vbcrlf
	strErr=strErr & "	<ul class=""main_end"">" & vbcrlf
	strErr=strErr & "		<li class=""main_end_left left""></li>" & vbcrlf
	strErr=strErr & "		<li class=""main_end_right right""></li>" & vbcrlf
	strErr=strErr & "	</ul>" & vbcrlf
	strErr=strErr & "</div>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	Response.write strErr
end sub
%>