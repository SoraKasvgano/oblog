<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/Cls_XmlDoc.asp"-->
<!--#include file="API/Class_API.asp" -->
<%
Response.expires = 0
Response.expiresabsolute = now() - 1
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.addHeader "P3P","CP=CAO PSA OUR"
Response.cachecontrol = "no-cache"
'------------------------------------------------
'检测搜索引擎,截断程序执行,节省资源. *#0801Spider
oblog.ChkSpider(1)
'------------------------------------------------
Dim username,password,show_login,CookieDate,fromurl,action
action=Request("action")
if action<>"showindexlogin" and action<>"showjs" then
	if  oblog.checkuserlogined() then Response.Redirect("user_index.asp")
end if

username=oblog.filt_badstr(Trim(Request.form("username")))
password=Trim(Request.form("password"))
CookieDate=Trim(Request.form("CookieDate"))
fromurl=Trim(Request.form("fromurl"))
if username<>"" or Request("chk")="1" then
	call sub_chklogin
else
	if action="showindexlogin" then
		call sub_showindexlogin()
	elseif action="showjs" then
		blogurl=oblog.CacheConfig(3)
		call sub_showindexlogin()
	else
		call sub_showlogin()
	end if
end If
Set oblog = Nothing
sub sub_showlogin()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>用户管理登录</title>
<link rel="stylesheet" href="oBlogStyle/login/css.css" type="text/css" />
</head>
<body>
<!-- header //-->
<div id="Header">
	<div id="logo" title="用户管理登录">用户管理登录</div>
	<ul id="menu">
		<li>
			<a href="index.asp">首页</a>&nbsp;|&nbsp;
			<a href="reg.asp">注册</a>
		</li>
	</ul>
</div>
<!-- header end //-->
<!-- Container //-->
<div id="Container">
	<div id="PageBody">
		<div class="Sidebar">
			<form name="UserLogin" method="post" action="login.asp?chk=1&fromurl=<%=fromurl%>">
				<ul>
					<li><label>用户名：<input type="text" name="UserName" id="UserName" onFocus="this.className='input_onFocus'" onBlur="this.className='input_onBlur'" value="<%=Request.Cookies(cookies_name)("username")%>" /></label></li>
					<li><label>密　码：<input type="password" name="Password" id="Password" onFocus="this.className='input_onFocus'" onBlur="this.className='input_onBlur'" /></label></li>
					<%if oblog.CacheConfig(29)=1 then%>
					<li><label>验证码：<input name="codestr" id="codestr" type="text" class="put2" size="6" maxlength="20" onFocus="this.className='input_onFocus'" onBlur="this.className='input_onBlur'" /><%Response.Write(oblog.getcode)%></label></li>
					<%end if%>
					<li class="CookieDate"><label for="CookieDate"><input type="checkbox" name="CookieDate" id="CookieDate" value="3" />保存我的登录信息</label></li>
					<li><input type="hidden" name="fromurl" value="<%=fromurl%>"><input name="Submit" id="Submit" type="submit" value="登　录" /><a href="lostpassword.asp">忘记密码？</a></li>
					<li class="hr"></li>
					<li>如果你不是本站会员，请——</li>
					<li class="regbt"><a href="reg.asp"><img src="oBlogStyle/login/reg.jpg" /></a></li>
				</ul>
			</form>
			<ul class="help">
				<li>如果你密码丢失或原有用户名登录不了，请试试<a href="lostpassword.asp">找回密码</a>。</li>
				<%if oblog.CacheConfig(29)=1 then%>
				<li>当你看不清验证码时请点验证码图片刷新。</li>
				<%end if%>
			</ul>
		</div>
		<div class="MainBody">
			<div class="ad">稳定的平台，完善的功能，满意的服务，和谐的环境。</div>
			<dl class="d1">
				<dt>发布网络文章</dt>
				<dd>在网络中用文字记录您的日常生活</dd>
			</dl>
			<dl class="d2">
				<dt>共享您的照片</dt>
				<dd>保存和共享您的照片，用光和影展现您的生活</dd>
			</dl>
			<dl class="d3">
				<dt>展示个性的您</dt>
				<dd>您可自由设置空间，展示一个独一无二的自我</dd>
			</dl>
		</div>
		<div class="clear"></div>
	</div>
	<div class="clear"></div>
</div>
<!-- Container end //-->
<!-- Footer //-->
<div id="Footer"><%=oblog.site_bottom%></div>
<!-- Footer end //-->
</body>
</html>
<%
end sub

sub sub_chklogin()
	dim ajax,rearr,gohref
	'set ajax=new AjaxXml
	if oblog.CacheConfig(29)=1 then
		if not oblog.codepass then oblog.adderrstr("验证码错误！")
	end If
	If oblog.Chkiplock() Then
		oblog.ShowMsg ("对不起！你的IP已被锁定，不允许操作！"),blogdir &"index.html"
		Set oblog = Nothing
	End If
	if UserName="" then oblog.adderrstr("登录用户名不能为空！")
	if Password="" then oblog.adderrstr("登录密码不能为空！")
	if oblog.errstr<>"" then
		rearr=split(Replace(oblog.errstr,"_","\n")&"$$1","$$")
		Response.Write "<script language=JavaScript>alert(""" & rearr(0) & """);history.go(-1)</script>"
		Response.End()
		'ajax.re(rearr)
		'Response.end
	end if
	if CookieDate="" then CookieDate=0	else CookieDate=CLng(CookieDate)
'	password=md5(password)
	if Is_ot_User=1 then
		call ot_chklogin()
	Else
		oblog.ob_chklogin UserName,MD5(password),CookieDate

		If API_Enable Then
				Dim blogAPI
				Set blogAPI = New DPO_API_OBLOG
				blogAPI.LoadXmlFile True
				blogAPI.UserName=username
				blogAPI.PassWord=password
				blogAPI.CookieDate=CookieDate
				blogAPI.userip=oblog.userip
				Call blogAPI.ProcessMultiPing("login")
				Set blogAPI=Nothing
				Dim strUrl,i,turl
				For i=0 To UBound(aUrls)
					strUrl=aUrls(i)
					if CookieDate=0 then CookieDate=3
					If Left(strUrl,7)="http://" Then
						turl=strUrl&"?syskey="&MD5(UserName&oblog_Key)&"&username="&UserName&"&password="&MD5(PassWord)&"&savecookie="&CookieDate & "@@@"& turl
					End If
				Next
				session("turl")=turl
				Dim trearr
				trearr="$$"&MD5(username & oblog_Key )&"$$"&username&"$$"&MD5(password)
		End If
	End If
	if oblog.errstr<>"" then
		rearr=split(Replace(oblog.errstr,"_","\n")&"$$1","$$")
		Response.Write "<script language=JavaScript>alert(""" & rearr(0) & """);history.go(-1)</script>"
		Response.End()
		'ajax.re(rearr)
		'Response.end
	end if
	if fromurl<>"" then
		gohref=Replace(fromurl,"&","$")
		rearr=split("登录成功!$$2$$"&gohref & trearr,"$$")
	else
		if action="showindexlogin" then
			gohref=oblog.comeurl
		else
			gohref="user_index.asp"
		end if
	end if
	rearr=split("登录成功!$$2$$"&gohref & trearr,"$$")
	if rearr(1)=2 Then
		If InStr (rearr(2),"user_index.asp")>0 Then
			Response.Redirect(rearr(2))
		Else
			Response.Redirect(Replace(rearr(2),"$","&"))
		End if
	else
		Response.Write "<script language=JavaScript>alert(""" & rearr(0) & """);history.go(-1)</script>"
	end if
	'ajax.re(rearr)
	'Response.End
end Sub

sub ot_chklogin()
	dim sql,rs,rsreg
	Dim ajax,rearr
	set ajax=new AjaxXml
	Dim TruePassWord
	TruePassWord = RndPassword(16)
	if not IsObject(ot_conn) then link_database
	sql="select * from "&ot_usertable&" where "&ot_username&"='"& username & "' and "&ot_password&" ='" & md5(password) &"'"
	set rs=ot_conn.execute(sql)
	if rs.bof and rs.eof then
		set rs=nothing
		if isobject(ot_conn) then ot_conn.close:set ot_conn=nothing
		oblog.adderrstr("用户名或密码错误，请重新输入！！")
		exit sub
	else
		set rsreg=Server.CreateObject("adodb.recordset")
		rsreg.open "select * from [oblog_user] where username='"& username &"'",conn,1,3
		if rsreg.eof then
			dim reguserlevel
			If oblog.CacheConfig(18) = 1 Then reguserlevel = 6 Else reguserlevel = 7
			set rsreg=Server.CreateObject("adodb.recordset")
			rsreg.open "select top 1 * from [oblog_user]",conn,1,3
			rsreg.addnew
			rsreg("username")=username
			rsreg("password")=MD5(password)
			rsreg("TruePassWord") = TruePassWord
			rsreg("user_dir")=oblog.setup(8,0)
			rsreg("user_level")=reguserlevel
			rsreg("lockuser")=0
			rsreg("en_blogteam")=1
			rsreg("adddate")=oblog.ServerDate(Now())
			rsreg("regip")=oblog.userip
			rsreg("lastloginip")=oblog.userip
			rsreg("lastlogintime")=oblog.ServerDate(now())
			rsreg("user_group") = oblog.defaultGroup
			rsreg("scores") = oblog.cacheScores(1)
			rsreg("newbie") = 1
			if oblog.CacheConfig(40)=1 then rsreg("comment_isasc")=1
			If oblog.chkdomain(UserName)=False Then
				rsreg("Nickname")=UserName
			End If
			rsreg.update
			oblog.execute("update oblog_user set user_folder=userid where username='"&username&"'")
			oblog.execute("update oblog_setup set user_count=user_count+1")
			rsreg.close
			set rsreg=nothing
			oblog.SaveCookie username,TruePassWord,0
			oblog.CreateUserDir username,1
			set rs=Nothing
			'rearr=split("您是第一次激活blog系统，请完善blog资料!$$2$$user_index.asp","$$")
			oblog.ShowMsg "您是第一次激活blog系统，请完善blog资料!","user_index.asp"
			'ajax.re(rearr)
			Response.End
		Else
			If rsreg("lockuser") = 1 Then
				rsreg.Close: Set rsreg = Nothing
				oblog.ShowMsg ("对不起！你的ID已被锁定，不能登录！")
				Exit Sub
			Else
				If rsreg("password")<>MD5(password) Then rsreg("password")=MD5(password)
				rsreg("LastLoginIP")=oblog.userip
				rsreg("LastLoginTime")=oblog.ServerDate(Now())
				rsreg("LoginTimes")=rsreg("LoginTimes")+1
				rsreg("TruePassWord") = TruePassWord
				rsreg.update
			End if
		end if
		rsreg.close
		set rsreg=nothing
		set rs=nothing
		if isobject(ot_conn) then ot_conn.close:set ot_conn=nothing
		oblog.SaveCookie username,TruePassWord,CookieDate
	end if
end sub


sub sub_showindexlogin()
	dim show_userlogin
	if oblog.CheckUserLogined()=False then
		if Request("n")="1" then '横向登录口
			show_userlogin="<form action="""&blogurl&"login.asp?action=showindexlogin&chk=1"" method=""post"" name=""UserLogin"">" & vbcrlf
			show_userlogin=show_userlogin & "	<table class=""Before"" align=""center"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" style=""font-size:12px"">" & vbcrlf
			show_userlogin=show_userlogin & "		<tr>" & vbcrlf
			show_userlogin=show_userlogin & "			<td>" & vbcrlf
			show_userlogin=show_userlogin & "				<label for=""UserName"">用户名：<input name=""UserName"" type=""text"" id=""UserName"" size=""12"" maxlength=""20"" value="""&Request.Cookies(cookies_name)("username")&""" /></label>" & vbcrlf
			show_userlogin=show_userlogin & "				<label for=""Password"">　密码：<input name=""Password"" type=""password"" id=""Password"" size=""12"" maxlength=""20"" /></label>" & vbcrlf
			if oblog.CacheConfig(29)=1 then
				show_userlogin=show_userlogin&"				　<label for=""codestr"">验证码：<input name=""codestr"" type=""text"" id=""codestr"" size=""4"" maxlength=""20"" /></label>" & oblog.getcode & vbcrlf
			end if
			show_userlogin=show_userlogin & "				<label for=""CookieDate"">　<input type=""checkbox"" name=""CookieDate"" id=""CookieDate"" value=""3"">记住密码</label>" & vbcrlf
			show_userlogin=show_userlogin & "				　<input name=""fromurl"" type=""hidden""><input name=""Login"" type=""submit"" id=""submit"" value=""登录"" >" & vbcrlf
			show_userlogin=show_userlogin & "				<a href="""&blogurl&"reg.asp"">用户注册</a>&nbsp;<a href="""&blogurl&"lostpassword.asp"">忘记密码</a>" & vbcrlf
			show_userlogin=show_userlogin & "			</td>" & vbcrlf
			show_userlogin=show_userlogin & "		</tr>" & vbcrlf
			show_userlogin=show_userlogin & "	</table>" & vbcrlf
			show_userlogin=show_userlogin & "</form>" & vbcrlf
		Else '竖向登录口
			show_userlogin="<form action="""&blogurl&"login.asp?action=showindexlogin&chk=1"" method=""post"" name=""UserLogin"">" & vbcrlf
			show_userlogin=show_userlogin & "	<table class=""Before"" align=""center"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
			show_userlogin=show_userlogin & "		<tr class=""t1"">" & vbcrlf
			show_userlogin=show_userlogin & "			<td height=""25"">" & vbcrlf
			show_userlogin=show_userlogin & "				<label for=""UserName"">用户名：<input name=""UserName"" type=""text"" id=""UserName"" size=""12"" maxlength=""20""  value="""&Request.Cookies(cookies_name)("username")&""" /></label>" & vbcrlf
			show_userlogin=show_userlogin & "			</td>" & vbcrlf
			show_userlogin=show_userlogin & "		</tr>" & vbcrlf
			show_userlogin=show_userlogin & "		<tr class=""t2"">" & vbcrlf
			show_userlogin=show_userlogin & "			<td height=""25"">" & vbcrlf
			show_userlogin=show_userlogin & "				<label for=""Password"">密　码：<input name=""Password"" type=""password"" id=""Password"" size=""12"" maxlength=""20"" /></label>" & vbcrlf
			show_userlogin=show_userlogin & "			</td>" & vbcrlf
			show_userlogin=show_userlogin & "		</tr>" & vbcrlf
			if oblog.CacheConfig(29)=1 Then
				show_userlogin=show_userlogin & "		<tr class=""t3"">" & vbcrlf
				show_userlogin=show_userlogin & "			<td height=""25"">" & vbcrlf
				show_userlogin=show_userlogin & "				<label for=""codestr"">验证码：<input name=""codestr"" id=""codestr"" type=""text"" size=""4"" maxlength=""20"" /></label>" & oblog.getcode & vbcrlf
				show_userlogin=show_userlogin & "			</td>" & vbcrlf
				show_userlogin=show_userlogin & "		</tr>" & vbcrlf
			end If
			show_userlogin=show_userlogin & "		<tr class=""t4"">" & vbcrlf
			show_userlogin=show_userlogin & "			<td height=""25"">" & vbcrlf
			show_userlogin=show_userlogin & "				　　　　<label for=""CookieDate""><input type=""checkbox"" name=""CookieDate"" id=""CookieDate"" value=""3"">记住密码</label>" & vbcrlf
			show_userlogin=show_userlogin & "			</td>" & vbcrlf
			show_userlogin=show_userlogin & "		</tr>" & vbcrlf
			show_userlogin=show_userlogin & "		<tr class=""t5"">" & vbcrlf
			show_userlogin=show_userlogin & "			<td height=""25"">" & vbcrlf
			show_userlogin=show_userlogin & "				<input name=""fromurl"" type=""hidden""><input name=""Login"" type=""submit"" id=""Login"" value=""登录"" />&nbsp;<a href="""&blogurl&"reg.asp"">用户注册</a>&nbsp;<a href="""&blogurl&"lostpassword.asp"">忘记密码</a>" & vbcrlf
			show_userlogin=show_userlogin & "			</td>" & vbcrlf
			show_userlogin=show_userlogin & "		</tr>" & vbcrlf
			show_userlogin=show_userlogin & "	</table>" & vbcrlf
			show_userlogin=show_userlogin & "</form>" & vbcrlf
		end If
	Else
		if Request("n")="1" then '横向登录后状态
			show_userlogin="<table class=""After"" align=""center"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
			show_userlogin=show_userlogin & "	<tr>" & vbcrlf
			show_userlogin=show_userlogin & "		<td>" & vbcrlf
			show_userlogin=show_userlogin & "			欢迎您," & oblog.l_uname & "&nbsp;&nbsp;" & vbcrlf
			show_userlogin=show_userlogin & "			您的身份：" & oblog.l_Group(1,0) & vbcrlf
			show_userlogin=show_userlogin & "			&nbsp;&nbsp;<a href="""&blogurl&"go.asp?user="&oblog.l_uname&""" target=""_blank"">我的首页</a>" & vbcrlf
			show_userlogin=show_userlogin & "			&nbsp;&nbsp;<a href=""" & blogurl & "user_index.asp"" target=""_blank"">管理中心</a>" & vbcrlf
			show_userlogin=show_userlogin & "			&nbsp;&nbsp;<a href="""&blogurl&"user_index.asp?url=user_post.asp"" target=""_blank"">发表日志</a>" & vbcrlf
			show_userlogin=show_userlogin & "			&nbsp;&nbsp;<a href="""&blogurl&"user_index.asp?t=logout&re=1"">注销登录</a>" & vbcrlf
			show_userlogin=show_userlogin & "		</td>" & vbcrlf
			show_userlogin=show_userlogin & "	</tr>" & vbcrlf
			show_userlogin=show_userlogin & "</table>" & vbcrlf
		Else '竖向登录后状态
			show_userlogin= "<table class=""After"" align=""center"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
			show_userlogin=show_userlogin & "	<tr class=""t1"">" & vbcrlf
			show_userlogin=show_userlogin & "		<td height=""25"">" & vbcrlf
			show_userlogin=show_userlogin & "			--欢迎您," & oblog.l_uname & "--" & vbcrlf
			show_userlogin=show_userlogin & "		</td>" & vbcrlf
			show_userlogin=show_userlogin & "	</tr>" & vbcrlf
			show_userlogin=show_userlogin & "	<tr class=""t2"">" & vbcrlf
			show_userlogin=show_userlogin & "		<td height=""25"">" & vbcrlf
			show_userlogin=show_userlogin & "			您的身份：" & oblog.l_Group(1,0) & vbcrlf
			show_userlogin=show_userlogin & "		</td>" & vbcrlf
			show_userlogin=show_userlogin & "	<tr class=""t3"">" & vbcrlf
			show_userlogin=show_userlogin & "	<tr>" & vbcrlf
			show_userlogin=show_userlogin & "		<td height=""25"">" & vbcrlf
			show_userlogin=show_userlogin & "			<a href="""&blogurl&"go.asp?user="&oblog.l_uname&""" target=""_blank"">我的首页</a>&nbsp;&nbsp;<a href=""" & blogurl & "user_index.asp"" target=""_blank"">管理中心</a>" & vbcrlf
			show_userlogin=show_userlogin & "		</td>" & vbcrlf
			show_userlogin=show_userlogin & "	<tr class=""t4"">" & vbcrlf
			show_userlogin=show_userlogin & "	<tr>" & vbcrlf
			show_userlogin=show_userlogin & "		<td height=""25"">" & vbcrlf
			show_userlogin=show_userlogin & "			<a href="""&blogurl&"user_index.asp?url=user_post.asp"" target=""_blank"">发表日志</a>&nbsp;&nbsp;<a href="""&blogurl&"user_index.asp?t=logout&re=1"">注销登录</a>" & vbcrlf
			show_userlogin=show_userlogin & "		</td>" & vbcrlf
			show_userlogin=show_userlogin & "	</tr>" & vbcrlf
			show_userlogin=show_userlogin & "</table>" & vbcrlf
		end If
		If API_Enable Then
			If session("turl")<>"" Then
				Dim arrturl,i,turl,scrurl
				turl=Replace(session("turl"),"$","&")
				arrturl=Split(turl,"@@@")
				For i=0 To UBound(arrturl)
					If arrturl(i)="" Then Exit For
					scrurl= scrurl& "<iframe src="""&arrturl(i)&""" frameborder=""0"" scrolling=""no"" height=""0"" width=""0""></iframe>" & vbcrlf
					'scrurl= scrurl& "<script type=""text/javascript"" language=""JavaScript"" src="""&arrturl(i)&""" charset=""gb2312""></script>" & vbcrlf
				Next
				response.Write("document.write('"&Replace(Replace(Replace(Replace(scrurl,"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"');")
				Response.Flush
				session("turl")=""
			End if
		End if
	end if
	Response.Write oblog.htm2js_div(show_userlogin,"ob_login")
end sub
%>
