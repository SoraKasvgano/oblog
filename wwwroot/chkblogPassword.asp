<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Dim uid,blog_password,rs,udir,uname,fromurl,ufolder
Dim groupID,gUrl,CookieName
Dim ShowTitle
uid=Trim(Request.QueryString("userid"))
blog_password=Trim(Request.Form("blog_password"))
fromurl=Trim(Request("fromurl"))
groupID = Trim(Request.QueryString("groupid"))
If uid<>"" Then uid = CLng(uid)
If groupID<>"" Then groupID = CLng(groupID)
If uid <> "" Then
	gUrl = "pwblog.asp?action=blog&userid="&uid
	CookieName = "blog_pwd_"&uid
	ShowTitle = "blog"
	set rs=oblog.execute("select blog_password,user_dir,username,user_folder from [oblog_user] where userid="&uid)
	if not rs.eof then
		udir=rs(1)
		uname=rs(2)
		ufolder=rs(3)
		if rs(0)="" Or IsNull(rs(0))then
			set rs=nothing
			Response.Redirect blogurl&udir&"/"&ufolder&"/index."&f_ext
		end if
	else
		set rs=nothing
		Response.Write("无此用户")
		Response.End()
	end if
End If
If groupID <> "" Then
	gUrl = blogdir&"group.asp?gid="&groupID
	CookieName = "group_pwd_"&groupID
	ShowTitle = oblog.CacheConfig(69)
	Set rs = oblog.Execute ("select ViewPassWord FROM oblog_team WHERE teamid = "&groupID)
	If Not rs.EOF Then
		If rs(0) = "" Or IsNull(rs(0)) Then
			rs.Close
			Set rs = Nothing
			Response.Redirect gUrl
			Response.End
		End if
	Else
		rs.Close
		Set rs = Nothing
		Response.Write "无此" &oblog.CacheConfig(69)
		Response.End
	End if
End if
if blog_password<>"" then
	blog_password=MD5(blog_password)
	if rs(0)=blog_password then
		set rs=Nothing
		Response.Cookies(cookies_name).Path   =   blogdir
		Response.Cookies(cookies_name)(CookieName) =blog_password
		if fromurl<>"" then
			Response.Redirect(Replace(fromurl,"$","&"))
		else
			Response.Redirect gUrl
		end if
	Else
		set rs=nothing
		oblog.ShowMsg "密码错误，请重新输入。",""
	end if
end if
set rs=nothing
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>加密<%=ShowTitle%>访问验证页面</title>
<style type="text/css">
<!--
body {
color: #333;
background: #fff;
text-align: left;
margin:0;
font-family: 'Century Gothic', Arial, Helvetica, sans-serif;
font-size: 12px;
line-height: 150%;
}
.content {
width:412px;
height:232px;
margin: 80px 0px 0px 0px;
background: url("images/passwordbg.png") no-repeat top center;
}
#list form {
padding:110px 0px 0px 60px;
}
#list form #password{
border: 0px #694659 dotted;
width:240px;
height:30px;
font-size:30px;
color:#099;
background: url("images/none.png");
}
#list form #submit {
margin:28px 0px 0px 165px;
}
-->
</style>
</head>
<body>
<center>
<div class="content">

	<div id="list">
	<ul>
	<form method="post">
	<input type="password" size="18" maxlength="20" name="blog_password" id="password"/>
	<input type="image" value="提交" src="images/passwordbt.png" alt="Login" id="submit"/>
	</form>
	</ul>
	</div>
  </div>
</center>
</body>
</html>