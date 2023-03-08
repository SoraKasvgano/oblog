<!--#include file="../conn.asp"-->
<!--#include file="../inc/class_sys.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.cachecontrol = "no-cache"
If Request("action")="logout" Then
	session("AdminName")=""
	session("adminpassword")=""
	Session("roleid") = ""
	Response.Redirect "../index.asp"
End If
dim oblog
set oblog=new class_sys
oblog.start
'------------------------------------------------
'检测搜索引擎,截断程序执行,节省资源. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
if Request("action")<>"login" then
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>OBlog后台管理员登录</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="images/style.css">
<script language=javascript>
<!--
function SetFocus()
{
if (document.Login.username.value=="")
	document.Login.username.focus();
else
	document.Login.username.select();
}
function CheckForm()
{
	if(document.Login.username.value=="")
	{
		alert("请输入用户名！");
		document.Login.username.focus();
		return false;
	}
	if(document.Login.password.value == "")
	{
		alert("请输入密码！");
		document.Login.password.focus();
		return false;
	}
	if (document.Login.codestr.value==""){
       alert ("请输入您的验证码！");
       document.Login.codestr.focus();
       return false;
    }
}

function CheckBrowser()
{
  var app=navigator.appName;
  var verStr=navigator.appVersion;
  if (app.indexOf('Netscape') != -1) {
    alert("提示：\n    你使用的是Netscape浏览器，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。");
  }
  else if (app.indexOf('Microsoft') != -1) {
    if (verStr.indexOf("MSIE 3.0")!=-1 || verStr.indexOf("MSIE 4.0") != -1 || verStr.indexOf("MSIE 5.0") != -1 || verStr.indexOf("MSIE 5.1") != -1)
      alert("提示：\n    您的浏览器版本太低，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。");
  }
}
//-->
</script>
</head>
<body>
<div id="Login">
	<div id="ver"><strong>Version&nbsp;</strong><%=ver%></div>
	<form name="Login" action="admin_login.asp?action=login" method="post" target="_parent" onSubmit="return CheckForm();">
		<fieldset>
			<legend>OBlog后台管理员登录</legend>
				<ul>
				<span id="bizupdateMsg"></span>
					<li><label for="username">用户名称：
					<input name="username"  type="text"  id="username" maxlength="20" onFocus="this.select();this.style.background='#ffC';" onBlur="this.style.background='#FFF'" /></label></li>
					<li><label for="password">用户密码：
					<input name="password"  type="password" id="password" onFocus="this.select();this.style.background='#ffC';" onBlur="this.style.background='#FFF'" maxlength="20" /></label></li>
					<li>登录身份：
						<label for="logintype0"><input name="logintype"  type="radio"  id="logintype0" value="0" checked />系统管理员</label>
						<label for="logintype1"><input name="logintype"  type="radio"  id="logintype1" value="1" />内容管理员</label>
					</li>
					<li><label for="codestr">验 证 码：
						<input name="codestr" id="codestr" onFocus="this.select();this.style.background='#ffC';" onBlur="this.style.background='#FFF'" size="6" maxlength="20" /></label>
						<%=oblog.getcode%>
					</li>
					<li><input type="submit" id="Submit" value=" 登 录 " /></li>
				</ul>
		</fieldset>
	</form>
</div>
<script language=javascript src="http://www.oblog.cn/oblog4update.asp?ver0=<%=ver0%>&ver1=<%=ver1%>&ver2=<%=ver2%>&ver3=<%=ver3%>"></script>
<script language="JavaScript" type="text/JavaScript">
SetFocus();
</script>
</body>
</html>
<%
else
	dim sql,rs
	dim username,password
	dim founderr,errmsg
	Dim logintype,strlogin
	Dim WriteErrLog
	Dim sIP
	sIP=oblog.userIp

	WriteErrLog = True
	if not oblog.codepass Then
		WriteErrLog = False
		FoundErr=True
		errmsg=errmsg & "<br><li>验证码错误！</li>"
	end if
	username=oblog.filt_badstr(Trim(Request("username")))
	password=Trim(Request("password"))
	logintype=Trim(Request("logintype"))
	If logintype<>"" Then
		logintype = CLng (logintype)
		If logintype > 1 Then logintype = 0
	Else
		logintype = 0
	End If
	If logintype = 0 Then
		strlogin = "系统"
	Else
		strlogin = "内容"
	End if
	if username="" Then
		WriteErrLog = False
		FoundErr=True
		errmsg=errmsg & "<br><li>用户名不能为空！</li>"
	end if
	if password="" Then
		WriteErrLog = False
		FoundErr=True
		errmsg=errmsg & "<br><li>密码不能为空！</li>"
	end if
	if FoundErr<>True then
		password=md5(password)
		set rs=Server.CreateObject("adodb.recordset")
		sql="select * from oblog_admin where username='"&username&"'"
		if not IsObject(conn) then link_database
		rs.open sql,conn,1,3
		if rs.bof and rs.eof then
			FoundErr=True
			errmsg=errmsg & "<br><li>用户名、密码错误或者权限不足！</li>"
		else
			if password<>rs("password") then
				FoundErr=True
				errmsg=errmsg & "<br><li>用户名、密码错误或者权限不足！</li>"
			Else
				If logintype = 0 Then
					If rs("roleid") <> 0 Then
						FoundErr=True
						errmsg=errmsg & "<br><li>用户名、密码错误或者权限不足！</li>"
						oblog.sys_err(errmsg)
						Response.End
					End If
				Else
					If rs("roleid") = -1 Then
						FoundErr=True
						errmsg=errmsg & "<br><li>用户名、密码错误或者权限不足！</li>"
						oblog.sys_err(errmsg)
						Response.End
					End If
				End if
				rs("LastLoginIP")=Request.ServerVariables("REMOTE_ADDR")
				rs("LastLoginTime")=oblog.ServerDate(Now())
				rs("LoginTimes")=rs("LoginTimes")+1
				rs.update
				session.Timeout=60
				If logintype = 0 Then
					session("adminname")=rs("username")
					session("adminpassword")=rs("password")
				Else
					Session("m_name")=rs("username")
					session("m_pwd")=rs("password")
				End if
				Session("roleid")=rs("roleid")
				rs.close
				'清理邀请码
				Call oblog.ClearOldOBCodes
				'清理回收站的日志
				'Call oblog.ClearOldUserRLog
				'---------------------------------------
				'写日志
				rs.Open "select * From oblog_syslog Where 1=0",conn,1,3
				rs.AddNew
				If logintype = 0 Then
					rs("username")=session("adminname")
				Else
					rs("username")=Session("m_name")
				End if
				rs("addtime")=oblog.ServerDate(Now())
				rs("addip")=sIP
				rs("desc")=session("adminname") & " 于 " & oblog.ServerDate(Now()) & " 自 " & sIP & " (admin/admin_login.asp) 登入"&strlogin&"管理员界面"
				rs("itype")=1 '2系统自动记录类/1:管理员操作类
				rs.Update
				rs.Close
				'---------------------------------------
				set rs=Nothing
				If logintype = 0 Then
					Response.redirect "admin_index.asp"
				Else
					Response.redirect "../manager/m_index.asp"
				End if
			end if
		end if
		rs.close
		set rs=nothing
	end if
	if founderr=True Then
		if WriteErrLog then
			'---------------------------------------
			'写日志
			set rs=Server.CreateObject("adodb.recordset")
			rs.Open "select * From oblog_syslog Where 1=0",conn,1,3
			rs.AddNew
			rs("username")=username
			rs("addtime")=oblog.ServerDate(Now())
			rs("addip")=sIP
			rs("desc")=username & " 于 " & oblog.ServerDate(Now()) & " 自 " & sIP & " (admin/admin_login.asp) 尝试登入"&strlogin&"管理员界面失败"
			rs("itype")=0	'错误登录日志
			rs.Update
			rs.Close
			'---------------------------------------
		End if
		oblog.sys_err(errmsg)
	end if
end if
%>